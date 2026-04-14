from pydantic import BaseModel, Field
from typing import Optional, List, Literal, Any
from datetime import datetime

class FileUploadResponse(BaseModel):
    file_id: str
    filename: str
    file_type: str
    file_size: int
    upload_time: datetime
    status: str

class FileListResponse(BaseModel):
    total: int
    page: int
    limit: int
    files: List[dict]

class MinerUExtractRequest(BaseModel):
    file_id: str

class MinerUExtractResponse(BaseModel):
    task_id: str
    file_id: str
    status: str
    created_at: datetime

class MinerUResultResponse(BaseModel):
    task_id: str
    status: str
    result_file: Optional[str]
    questions_count: int
    data: Optional[dict]

class DeepSeekFixRequest(BaseModel):
    minerU_task_id: Optional[str] = None
    json_file_path: Optional[str] = None
    use_existing_file: bool = False

class DeepSeekFixResponse(BaseModel):
    task_id: str
    status: str
    source: str
    created_at: datetime

class DeepSeekResultResponse(BaseModel):
    task_id: str
    status: str
    result_file: Optional[str]
    questions_count: int
    json_content: Optional[List[dict]]
    data: Optional[Any]

class PPTGenerateRequest(BaseModel):
    json_file_path: str
    use_animation: bool = True
    title: Optional[str] = "试卷评讲"

class PPTGenerateResponse(BaseModel):
    task_id: str
    status: str
    input_json: str
    created_at: datetime

class PPTResultResponse(BaseModel):
    task_id: str
    status: str
    pptx_file: Optional[str]
    slide_count: Optional[int]
    questions_count: Optional[int]

class GenerateSimilarRequest(BaseModel):
    count: int = Field(default=1, ge=1, le=5)
    difficulty_adjustment: Literal["easier", "same", "harder"] = "same"
    preserve_knowledge_points: bool = True
    exam_id: str
    user_id: str = "default_user"

class GenerateSimilarResponse(BaseModel):
    task_id: str
    status: str
    source_question: dict
    generation_config: dict
    created_at: datetime

class SimilarQuestionResultResponse(BaseModel):
    task_id: str
    status: str
    source_question: dict
    similar_questions: List[dict]
    generated_count: int
    created_at: datetime
    completed_at: Optional[datetime]

class BatchGenerateSimilarRequest(BaseModel):
    question_indices: List[int]
    count_per_question: int = Field(default=1, ge=1, le=5)
    difficulty_adjustment: Literal["easier", "same", "harder"] = "same"
    exam_id: str
    user_id: str = "default_user"

class BatchGenerateSimilarResponse(BaseModel):
    batch_task_id: str
    tasks: List[dict]
    total_questions: int
    created_at: datetime

class PipelineFullRequest(BaseModel):
    use_animation: bool = True
    title: Optional[str] = "试卷评讲"

class PipelineFullResponse(BaseModel):
    pipeline_id: str
    stages: dict
    created_at: datetime

class PipelineStatusResponse(BaseModel):
    pipeline_id: str
    stages: dict
    all_completed: bool
    exam_id: Optional[str] = None

class TaskStatusResponse(BaseModel):
    task_id: str
    task_type: str
    status: str
    progress: int
    current_step: str
    error: Optional[str]
    created_at: datetime
    updated_at: datetime

class HealthResponse(BaseModel):
    status: str
    timestamp: datetime


# ==================== 题目收藏相关模型 ====================

class CreateFolderRequest(BaseModel):
    teacher_id: str
    name: str
    parent_id: Optional[str] = None
    color: Optional[str] = "#1890ff"


class FolderResponse(BaseModel):
    folder_id: str
    name: str
    parent_id: Optional[str]
    color: str
    children: Optional[List["FolderResponse"]] = None
    created_at: datetime


class CollectQuestionRequest(BaseModel):
    teacher_id: str
    question_id: str
    folder_id: Optional[str] = None
    tags: Optional[List[str]] = []
    difficulty_note: Optional[str] = None
    teach_note: Optional[str] = None
    common_errors: Optional[str] = None
    teach_points: Optional[str] = None


class CollectionResponse(BaseModel):
    collection_id: str
    question_id: str
    folder_id: Optional[str]
    folder_name: Optional[str]
    title: str
    content: str
    answer: Optional[str]
    analysis: Optional[str]
    tags: List[str]
    difficulty_note: Optional[str]
    teach_note: Optional[str]
    common_errors: Optional[str]
    teach_points: Optional[str]
    use_count: int
    last_used_at: Optional[datetime]
    created_at: datetime


class CollectionListResponse(BaseModel):
    total: int
    page: int
    limit: int
    collections: List[CollectionResponse]


class UpdateCollectionRequest(BaseModel):
    folder_id: Optional[str] = None
    tags: Optional[List[str]] = None
    difficulty_note: Optional[str] = None
    teach_note: Optional[str] = None
    common_errors: Optional[str] = None
    teach_points: Optional[str] = None


class CreateTagRequest(BaseModel):
    teacher_id: str
    tag_name: str
    color: Optional[str] = "#52c41a"


class TagResponse(BaseModel):
    tag_id: str
    tag_name: str
    color: str
    use_count: int
    created_at: datetime


class CollectionStatsResponse(BaseModel):
    total_count: int
    folder_count: int
    tag_count: int
    this_week_count: int
    most_used_tags: List[dict]
    recent_collections: List[dict]


class BatchMoveRequest(BaseModel):
    collection_ids: List[str]
    folder_id: str
    teacher_id: str


class BatchDeleteRequest(BaseModel):
    collection_ids: List[str]
    teacher_id: str


class BatchOperationResponse(BaseModel):
    success: bool
    affected_count: int
    error: Optional[str] = None


class ChatMessage(BaseModel):
    role: Literal["system", "user", "assistant"]
    content: str


class AIChatRequest(BaseModel):
    messages: List[ChatMessage]
    model: Optional[str] = "deepseek-chat"


class AIChatResponse(BaseModel):
    answer: str
    usage: Optional[dict] = None
