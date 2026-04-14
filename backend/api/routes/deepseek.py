from fastapi import APIRouter, BackgroundTasks, HTTPException
from datetime import datetime
from pathlib import Path

from api.models.schemas import DeepSeekFixRequest, DeepSeekFixResponse, DeepSeekResultResponse
from services.task_manager import task_manager, TaskType, TaskStatus
from services.deepseek_service import deepseek_service
from utils.storage import storage

router = APIRouter(prefix="/deepseek", tags=["DeepSeek修复"])

@router.post("/fix", response_model=DeepSeekFixResponse)
async def start_fix(request: DeepSeekFixRequest, background_tasks: BackgroundTasks):
    json_file_path = None
    source = ""

    if request.use_existing_file and request.json_file_path:
        json_file_path = Path(request.json_file_path)
        if not json_file_path.exists():
            raise HTTPException(status_code=404, detail="JSON文件不存在")
        source = request.json_file_path
    elif request.minerU_task_id:
        mineru_task = task_manager.get_task(request.minerU_task_id)
        if not mineru_task or mineru_task.status != TaskStatus.COMPLETED:
            raise HTTPException(status_code=400, detail="MinerU任务未完成或不存在")

        if mineru_task.result_file:
            json_file_path = Path(mineru_task.result_file)
        source = request.minerU_task_id
    else:
        raise HTTPException(status_code=400, detail="请提供minerU_task_id或json_file_path")

    task = task_manager.create_task(TaskType.DEEPSEEK, source_id=source)
    background_tasks.add_task(deepseek_service.fix_json, json_file_path, task.task_id)

    return DeepSeekFixResponse(
        task_id=task.task_id,
        status=task.status.value,
        source=source,
        created_at=task.created_at
    )

@router.get("/tasks/{task_id}")
async def get_task_status(task_id: str):
    task = task_manager.get_task(task_id)
    if not task:
        raise HTTPException(status_code=404, detail="任务不存在")

    return task.to_dict()

@router.get("/tasks/{task_id}/result", response_model=DeepSeekResultResponse)
async def get_fix_result(task_id: str):
    task = task_manager.get_task(task_id)
    if not task:
        raise HTTPException(status_code=404, detail="任务不存在")

    if task.status != TaskStatus.COMPLETED:
        raise HTTPException(status_code=400, detail="任务未完成")

    json_content = None
    if task.result_file:
        json_content = storage.load_json(Path(task.result_file))

    payload = task.data if task.data else json_content

    return DeepSeekResultResponse(
        task_id=task_id,
        status=task.status.value,
        result_file=task.result_file,
        questions_count=len(json_content) if isinstance(json_content, list) else 0,
        json_content=json_content,
        data=payload
    )

@router.post("/batch-fix")
async def batch_fix(request: dict, background_tasks: BackgroundTasks):
    task_ids = request.get("task_ids", [])
    tasks = []

    for task_id in task_ids:
        mineru_task = task_manager.get_task(task_id)
        if mineru_task and mineru_task.status == TaskStatus.COMPLETED and mineru_task.result_file:
            task = task_manager.create_task(TaskType.DEEPSEEK, source_id=task_id)
            background_tasks.add_task(deepseek_service.fix_json, Path(mineru_task.result_file), task.task_id)
            tasks.append({"mineru_task_id": task_id, "deepseek_task_id": task.task_id})

    return {
        "total": len(tasks),
        "tasks": tasks
    }
