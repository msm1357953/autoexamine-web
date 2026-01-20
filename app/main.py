"""
FastAPI 웹 서버 메인 모듈 - SSE 실시간 진행상황 지원
"""
from fastapi import FastAPI, Request, HTTPException
from fastapi.responses import StreamingResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from pathlib import Path
import traceback
import json
import asyncio
from typing import List
from queue import Queue
from threading import Thread

from .ppt_generator import PPTGenerator
from .dropbox_client import get_dropbox_client

# FastAPI 앱 생성
app = FastAPI(
    title="심의자료 자동화",
    description="삼성증권 광고 소재의 심의자료 PPT를 자동 생성합니다.",
    version="2.0.0"
)

# 정적 파일 및 템플릿 설정
BASE_DIR = Path(__file__).parent
app.mount("/static", StaticFiles(directory=BASE_DIR / "static"), name="static")
templates = Jinja2Templates(directory=BASE_DIR / "templates")

# PPT 결과 저장용 딕셔너리
ppt_results = {}


@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    """메인 페이지"""
    return templates.TemplateResponse("index.html", {"request": request})


@app.get("/api/all-materials")
async def get_all_materials():
    """전체 소재 목록 조회 API"""
    try:
        client = get_dropbox_client()
        materials = client.get_materials_list(None)
        
        return {
            "success": True,
            "count": len(materials),
            "materials": list(materials.keys()),
            "details": {name: sizes for name, sizes in materials.items()}
        }
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/api/materials")
async def get_materials(keyword: str = ""):
    """소재 목록 조회 API (키워드 필터)"""
    try:
        client = get_dropbox_client()
        materials = client.get_materials_list(keyword if keyword else None)
        
        return {
            "success": True,
            "keyword": keyword,
            "count": len(materials),
            "materials": list(materials.keys()),
            "details": {name: sizes for name, sizes in materials.items()}
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/api/generate-sse")
async def generate_sse(materials: str = ""):
    """
    PPT 생성 with SSE 실시간 진행상황
    """
    if not materials:
        raise HTTPException(status_code=400, detail="소재를 선택해주세요.")
    
    selected_materials = [m.strip() for m in materials.split(",") if m.strip()]
    
    async def event_generator():
        progress_queue = Queue()
        result_holder = {"ppt": None, "error": None}
        
        def send_progress(step: str, current: int, total: int, detail: str = ""):
            progress_queue.put({
                "type": "progress",
                "step": step,
                "current": current,
                "total": total,
                "detail": detail,
                "percent": int((current / total) * 100) if total > 0 else 0
            })
        
        def generate_in_thread():
            try:
                generator = PPTGenerator()
                ppt_buffer = generator.generate_with_progress(
                    selected_materials, 
                    progress_callback=send_progress
                )
                result_holder["ppt"] = ppt_buffer
            except Exception as e:
                result_holder["error"] = str(e)
                traceback.print_exc()
            finally:
                progress_queue.put({"type": "done"})
        
        # 백그라운드 스레드에서 PPT 생성
        thread = Thread(target=generate_in_thread)
        thread.start()
        
        # SSE 이벤트 스트리밍
        while True:
            await asyncio.sleep(0.1)
            
            while not progress_queue.empty():
                msg = progress_queue.get()
                
                if msg["type"] == "done":
                    if result_holder["error"]:
                        yield f"data: {json.dumps({'type': 'error', 'message': result_holder['error']}, ensure_ascii=False)}\n\n"
                    else:
                        # PPT를 임시 저장하고 다운로드 토큰 생성
                        import uuid
                        token = str(uuid.uuid4())
                        ppt_results[token] = result_holder["ppt"]
                        yield f"data: {json.dumps({'type': 'complete', 'token': token, 'filename': f'심의자료_{len(selected_materials)}개소재.pptx'}, ensure_ascii=False)}\n\n"
                    return
                else:
                    yield f"data: {json.dumps(msg, ensure_ascii=False)}\n\n"
    
    return StreamingResponse(
        event_generator(),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "Connection": "keep-alive",
            "X-Accel-Buffering": "no"
        }
    )


@app.get("/api/download/{token}")
async def download_ppt(token: str):
    """생성된 PPT 다운로드"""
    if token not in ppt_results:
        raise HTTPException(status_code=404, detail="PPT를 찾을 수 없습니다. 다시 생성해주세요.")
    
    ppt_buffer = ppt_results.pop(token)  # 다운로드 후 삭제
    ppt_buffer.seek(0)
    
    # 파일명 인코딩 (한글 지원)
    from urllib.parse import quote
    filename = "autoexamine_result.pptx"
    encoded_filename = quote(filename)
    
    return StreamingResponse(
        ppt_buffer,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={
            "Content-Disposition": f"attachment; filename*=UTF-8''{encoded_filename}"
        }
    )


@app.post("/api/generate")
async def generate(keyword: str = "", materials: str = ""):
    """PPT 생성 API (기존 호환용)"""
    if not keyword and not materials:
        raise HTTPException(status_code=400, detail="키워드 또는 소재를 선택해주세요.")
    
    try:
        generator = PPTGenerator()
        
        if materials:
            selected_materials = [m.strip() for m in materials.split(",") if m.strip()]
            ppt_buffer = generator.generate_with_materials(selected_materials)
            filename = f"selected_{len(selected_materials)}_materials.pptx"
        else:
            ppt_buffer = generator.generate(keyword)
            filename = f"{keyword}.pptx"
        
        return StreamingResponse(
            ppt_buffer,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            headers={
                "Content-Disposition": f'attachment; filename="{filename}"'
            }
        )
    except ValueError as e:
        raise HTTPException(status_code=404, detail=str(e))
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"PPT 생성 중 오류: {str(e)}")


@app.get("/api/health")
async def health_check():
    """헬스체크 API"""
    return {"status": "healthy", "service": "autoexamine-web"}
