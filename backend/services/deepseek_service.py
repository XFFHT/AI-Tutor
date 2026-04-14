import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from openai import OpenAI
from pydantic import BaseModel, Field

from services.task_manager import TaskStatus, task_manager
from utils.config import API_KEY, API_URL, PROCESSED_JSON_DIR, PROMPT_PATH, TEMP_DIR
from utils.storage import storage


class Question(BaseModel):
    title: str = Field(..., description="题目编号或所属大题标题")
    source: str = Field(..., description="题目来源，如试卷名称、书籍名称等")
    content: str = Field(..., description="修复后的完整题干，包含条件、子题或选项")
    answer: str = Field(..., description="该题的参考答案或计算结果")
    analysis: str = Field(..., description="该题的核心知识点解析与解题思路")


@dataclass
class OCRObject:
    page_idx: int
    raw: str
    preview: str


@dataclass
class Chunk:
    label: str
    text: str
    preview: str


@dataclass
class StreamAttempt:
    content: str
    finish_reason: str
    parser_questions: List[Question]
    full_questions: List[Question]
    full_parse_ok: bool
    parse_error: Optional[str]


class IncrementalQuestionsParser:
    def __init__(self):
        self.buffer = ""
        self.questions: List[Question] = []
        self.errors: List[str] = []
        self.array_started = False
        self.array_ended = False
        self.scan_pos = 0
        self.current_object_start: Optional[int] = None
        self.depth = 0
        self.in_string = False
        self.escaped = False

    def feed(self, chunk: str) -> None:
        if not chunk:
            return
        self.buffer += chunk
        self._consume()

    def _locate_questions_array(self) -> None:
        if self.array_started:
            return

        match = re.search(r'"questions"\s*:\s*\[', self.buffer)
        if match:
            self.array_started = True
            self.scan_pos = match.end()
            return

        stripped = self.buffer.lstrip()
        if stripped.startswith("["):
            self.array_started = True
            self.scan_pos = len(self.buffer) - len(stripped) + 1

    def _consume(self) -> None:
        self._locate_questions_array()
        if not self.array_started:
            return

        i = self.scan_pos
        while i < len(self.buffer):
            ch = self.buffer[i]

            if self.current_object_start is None:
                if ch == "{":
                    self.current_object_start = i
                    self.depth = 1
                    self.in_string = False
                    self.escaped = False
                elif ch == "]":
                    self.array_ended = True
                    i += 1
                    self.scan_pos = i
                    return
                i += 1
                continue

            if self.escaped:
                self.escaped = False
                i += 1
                continue

            if ch == "\\":
                self.escaped = True
                i += 1
                continue

            if ch == '"':
                self.in_string = not self.in_string
                i += 1
                continue

            if not self.in_string:
                if ch == "{":
                    self.depth += 1
                elif ch == "}":
                    self.depth -= 1
                    if self.depth == 0 and self.current_object_start is not None:
                        raw_object = self.buffer[self.current_object_start : i + 1]
                        self._try_append_question(raw_object)
                        self.current_object_start = None

            i += 1

        self.scan_pos = i

    def _try_append_question(self, raw_object: str) -> None:
        try:
            item = json.loads(raw_object)
            self.questions.append(Question(**item))
        except Exception as exc:
            self.errors.append(str(exc))


class DeepSeekService:
    PAGE_PASS_MAX_TOKENS = 4096
    QUESTION_PASS_MAX_TOKENS = 1800

    def __init__(self):
        self.client = OpenAI(api_key=API_KEY, base_url=API_URL)
        self.prompt_path = PROMPT_PATH

    def _load_system_prompt(self) -> str:
        base_prompt = self.prompt_path.read_text(encoding="utf-8")
        contract = """

额外输出约束：
1. 你只能输出单个 JSON 对象，严格格式为 {"questions":[...]}。
2. 不能输出 Markdown、解释文字、代码块或前后缀。
3. questions 必须是数组，每道题都必须包含 title、source、content、answer、analysis。
4. analysis 要简洁，优先控制在 120 字以内。
5. 如果当前输入片段没有完整题目，返回 {"questions": []}。
"""
        return f"{base_prompt.rstrip()}\n{contract.strip()}"

    def _safe_label(self, label: str) -> str:
        return re.sub(r"[^0-9a-zA-Z_-]+", "_", label)

    def _debug_response_path(self, task_id: str, label: str) -> Path:
        return TEMP_DIR / f"{task_id}_{self._safe_label(label)}.txt"

    def _report_path(self, task_id: str) -> Path:
        return TEMP_DIR / f"{task_id}_deepseek_report.json"

    def _save_debug_text(self, path: Path, content: str) -> None:
        try:
            path.write_text(content or "", encoding="utf-8")
        except Exception:
            pass

    def _save_report(self, task_id: str, report: Dict[str, Any]) -> Optional[str]:
        try:
            path = self._report_path(task_id)
            path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
            return str(path)
        except Exception:
            return None

    def _extract_json_candidate(self, content: str) -> str:
        stripped = content.strip()
        if not stripped:
            raise ValueError("DeepSeek returned empty content")

        fence_match = re.search(r"```(?:json)?\s*(.*?)\s*```", stripped, flags=re.DOTALL)
        if fence_match:
            stripped = fence_match.group(1).strip()

        if stripped.startswith("{") or stripped.startswith("["):
            return stripped

        balanced = self._extract_balanced_json(stripped)
        if balanced:
            return balanced

        raise ValueError("Failed to extract a complete JSON payload from DeepSeek output")

    def _extract_balanced_json(self, text: str) -> Optional[str]:
        for start_char, end_char in (("{", "}"), ("[", "]")):
            start = text.find(start_char)
            if start == -1:
                continue

            depth = 0
            in_string = False
            escaped = False

            for idx in range(start, len(text)):
                ch = text[idx]
                if escaped:
                    escaped = False
                    continue
                if ch == "\\":
                    escaped = True
                    continue
                if ch == '"':
                    in_string = not in_string
                    continue
                if in_string:
                    continue
                if ch == start_char:
                    depth += 1
                elif ch == end_char:
                    depth -= 1
                    if depth == 0:
                        return text[start : idx + 1]
        return None

    def _normalize_items(self, data: Any) -> List[Dict[str, Any]]:
        if isinstance(data, dict) and isinstance(data.get("questions"), list):
            return data["questions"]
        if isinstance(data, list):
            return data
        if isinstance(data, dict):
            return [data]
        return []

    def _parse_questions_from_content(self, content: str) -> List[Question]:
        candidate = self._extract_json_candidate(content)
        data = json.loads(candidate)
        items = self._normalize_items(data)
        return [Question(**item) for item in items]

    def _extract_preview(self, raw_object: str) -> str:
        patterns = [
            r'"text"\s*:\s*"((?:[^"\\]|\\.)*)"',
            r'"list_items"\s*:\s*\[\s*"((?:[^"\\]|\\.)*)"',
        ]
        for pattern in patterns:
            match = re.search(pattern, raw_object, flags=re.DOTALL)
            if match:
                preview = match.group(1)
                preview = preview.replace("\\n", " ").replace('\\"', '"')
                return preview.strip()
        return ""

    def _extract_ocr_objects(self, ocr_data: str) -> List[OCRObject]:
        pattern = re.compile(r"\{.*?\"page_idx\"\s*:\s*(\d+)\s*\}", re.DOTALL)
        objects: List[OCRObject] = []
        for match in pattern.finditer(ocr_data):
            raw = match.group(0)
            page_idx = int(match.group(1))
            objects.append(OCRObject(page_idx=page_idx, raw=raw, preview=self._extract_preview(raw)))
        return objects

    def _objects_to_chunk_text(self, objects: List[OCRObject]) -> str:
        return "[\n" + ",\n".join(obj.raw for obj in objects) + "\n]"

    def _split_ocr_by_page(self, ocr_data: str) -> List[Chunk]:
        grouped: Dict[int, List[OCRObject]] = {}
        for obj in self._extract_ocr_objects(ocr_data):
            grouped.setdefault(obj.page_idx, []).append(obj)

        chunks: List[Chunk] = []
        for page_idx in sorted(grouped):
            objects = grouped[page_idx]
            preview = next((obj.preview for obj in objects if obj.preview), f"page_{page_idx + 1}")
            chunks.append(
                Chunk(
                    label=f"page_{page_idx + 1}",
                    text=self._objects_to_chunk_text(objects),
                    preview=preview[:120],
                )
            )
        return chunks

    def _looks_like_question_start(self, preview: str) -> bool:
        text = preview.strip()
        patterns = [
            r"^\(?\d{1,3}\s*[\.．、\)]",
            r"^第\s*\d+\s*题",
        ]
        return any(re.match(pattern, text) for pattern in patterns)

    def _split_page_into_questions(self, page_chunk: Chunk) -> List[Chunk]:
        page_objects = self._extract_ocr_objects(page_chunk.text)
        if not page_objects:
            return []

        prefix: List[OCRObject] = []
        groups: List[List[OCRObject]] = []
        current_group: List[OCRObject] = []
        seen_first_question = False

        for obj in page_objects:
            if self._looks_like_question_start(obj.preview):
                seen_first_question = True
                if current_group:
                    groups.append(current_group)
                current_group = [obj]
                continue

            if not seen_first_question:
                prefix.append(obj)
            elif current_group:
                current_group.append(obj)

        if current_group:
            groups.append(current_group)

        if not groups:
            return []

        shared_prefix = prefix[:8]
        chunks: List[Chunk] = []
        for index, group in enumerate(groups, start=1):
            preview = next((obj.preview for obj in group if obj.preview), f"{page_chunk.label}_question_{index}")
            chunk_objects = shared_prefix + group
            chunks.append(
                Chunk(
                    label=f"{page_chunk.label}_question_{index}",
                    text=self._objects_to_chunk_text(chunk_objects),
                    preview=preview[:160],
                )
            )
        return chunks

    def _stream_request(self, system_prompt: str, chunk: Chunk, max_tokens: int) -> StreamAttempt:
        parser = IncrementalQuestionsParser()
        parts: List[str] = []
        finish_reason = ""

        stream = self.client.chat.completions.create(
            model="deepseek-chat",
            messages=[
                {"role": "system", "content": system_prompt},
                {
                    "role": "user",
                    "content": (
                        f"以下是需要处理的 OCR JSON 数据片段（{chunk.label}）。\n"
                        f"请只提取这一片段里能确认的完整题目。\n\n{chunk.text}"
                    ),
                },
            ],
            response_format={"type": "json_object"},
            max_tokens=max_tokens,
            temperature=0.2,
            stream=True,
        )

        for event in stream:
            if not event.choices:
                continue
            choice = event.choices[0]
            delta = getattr(choice.delta, "content", None) or ""
            if delta:
                parts.append(delta)
                parser.feed(delta)
            if choice.finish_reason:
                finish_reason = choice.finish_reason

        content = "".join(parts)
        try:
            full_questions = self._parse_questions_from_content(content)
            return StreamAttempt(
                content=content,
                finish_reason=finish_reason,
                parser_questions=parser.questions,
                full_questions=full_questions,
                full_parse_ok=True,
                parse_error=None,
            )
        except Exception as exc:
            return StreamAttempt(
                content=content,
                finish_reason=finish_reason,
                parser_questions=parser.questions,
                full_questions=[],
                full_parse_ok=False,
                parse_error=str(exc),
            )

    def _process_question_chunks(
        self,
        system_prompt: str,
        page_chunk: Chunk,
        task_id: str,
        progress_base: int,
        progress_span: int,
    ) -> Tuple[List[Question], List[Dict[str, Any]]]:
        task = task_manager.get_task(task_id)
        question_chunks = self._split_page_into_questions(page_chunk)
        if not question_chunks:
            return [], [
                {
                    "scope": "page",
                    "label": page_chunk.label,
                    "preview": page_chunk.preview,
                    "reason": "Page-level retry failed and no question boundaries could be detected",
                }
            ]

        kept_questions: List[Question] = []
        skipped_items: List[Dict[str, Any]] = []

        for index, question_chunk in enumerate(question_chunks, start=1):
            if task:
                progress = progress_base + int(progress_span * index / max(len(question_chunks), 1))
                task.update_status(
                    TaskStatus.PROCESSING,
                    progress=progress,
                    step=f"Page retry failed, retrying single question: {question_chunk.label}",
                )

            attempt = self._stream_request(system_prompt, question_chunk, self.QUESTION_PASS_MAX_TOKENS)
            self._save_debug_text(
                self._debug_response_path(task_id, question_chunk.label),
                attempt.content,
            )

            if attempt.finish_reason == "length" or not attempt.full_parse_ok:
                skipped_items.append(
                    {
                        "scope": "question",
                        "label": question_chunk.label,
                        "preview": question_chunk.preview,
                        "reason": (
                            "Single-question retry still failed"
                            if not attempt.parse_error
                            else f"Single-question retry failed: {attempt.parse_error}"
                        ),
                        "finish_reason": attempt.finish_reason or None,
                    }
                )
                continue

            if not attempt.full_questions:
                skipped_items.append(
                    {
                        "scope": "question",
                        "label": question_chunk.label,
                        "preview": question_chunk.preview,
                        "reason": "Single-question retry returned no question",
                        "finish_reason": attempt.finish_reason or None,
                    }
                )
                continue

            kept_questions.extend(attempt.full_questions)

        return kept_questions, skipped_items

    async def fix_json(self, json_file_path: Path, task_id: str) -> Optional[Path]:
        task = task_manager.get_task(task_id)
        if not task:
            return None

        try:
            task.update_status(TaskStatus.PROCESSING, progress=10, step="Reading OCR JSON...")
            ocr_data = json_file_path.read_text(encoding="utf-8")

            task.update_status(TaskStatus.PROCESSING, progress=20, step="Loading DeepSeek prompt...")
            system_prompt = self._load_system_prompt()

            page_chunks = self._split_ocr_by_page(ocr_data)
            if not page_chunks:
                raise RuntimeError("No OCR page chunks could be extracted from the MinerU output")

            all_questions: List[Question] = []
            skipped_items: List[Dict[str, Any]] = []

            for index, page_chunk in enumerate(page_chunks, start=1):
                task.update_status(
                    TaskStatus.PROCESSING,
                    progress=20 + int(50 * index / max(len(page_chunks), 1)),
                    step=f"Streaming DeepSeek on {page_chunk.label} ({index}/{len(page_chunks)})...",
                )

                attempt = self._stream_request(system_prompt, page_chunk, self.PAGE_PASS_MAX_TOKENS)
                self._save_debug_text(
                    self._debug_response_path(task_id, page_chunk.label),
                    attempt.content,
                )

                if attempt.finish_reason != "length" and attempt.full_parse_ok:
                    all_questions.extend(attempt.full_questions)
                    continue

                fallback_questions, fallback_skips = self._process_question_chunks(
                    system_prompt,
                    page_chunk,
                    task_id,
                    progress_base=70,
                    progress_span=15,
                )

                if fallback_questions:
                    all_questions.extend(fallback_questions)
                else:
                    skipped_items.append(
                        {
                            "scope": "page",
                            "label": page_chunk.label,
                            "preview": page_chunk.preview,
                            "reason": (
                                "Page-level streaming output was truncated"
                                if attempt.finish_reason == "length"
                                else f"Page-level parse failed: {attempt.parse_error}"
                            ),
                            "finish_reason": attempt.finish_reason or None,
                        }
                    )

                skipped_items.extend(fallback_skips)

            if not all_questions:
                raise RuntimeError(
                    "DeepSeek could not extract any complete questions. "
                    "All page/question retries failed; see skipped_items in task data."
                )

            task.update_status(TaskStatus.PROCESSING, progress=90, step="Saving cleaned questions...")
            output_file = PROCESSED_JSON_DIR / f"{json_file_path.stem}_cleaned.json"
            result_json = [question.model_dump() for question in all_questions]
            output_file.write_text(
                json.dumps(result_json, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )

            report = {
                "strategy": "streaming_incremental_page_then_question",
                "questions_count": len(result_json),
                "skipped_items": skipped_items,
            }
            report_file = self._save_report(task_id, report)

            task.update_status(TaskStatus.COMPLETED, progress=100, step="Completed")
            task.set_result(
                str(output_file),
                {
                    "questions": result_json,
                    "count": len(result_json),
                    "skipped_items": skipped_items,
                    "report_file": report_file,
                    "strategy": report["strategy"],
                },
            )
            return output_file

        except Exception as exc:
            task.set_error(str(exc))
            return None

    def get_questions_from_task(self, task_id: str) -> Optional[List[Dict[str, Any]]]:
        task = task_manager.get_task(task_id)
        if not task or task.status != TaskStatus.COMPLETED:
            return None

        if task.data and "questions" in task.data:
            return task.data["questions"]

        if task.result_file:
            return storage.read_json(Path(task.result_file))

        return None


deepseek_service = DeepSeekService()
