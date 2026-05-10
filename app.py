import sys
sys.stdout.reconfigure(encoding='utf-8')
from pathlib import Path
import os
import shutil
import uuid
import requests
from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from audit_core import run_ar_audit

app = FastAPI(title="AR Audit API", version="1.0.0-cloud")

app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "https://wiseaudit-frontend-48af4f6q9-lxt157522-afks-projects.vercel.app",
        "https://*.vercel.app",
        "http://localhost:3000",
        "http://localhost:5173",
    ],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = Path("/tmp/uploads")
OUTPUT_DIR = Path("/tmp/outputs")

UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


def save_upload(file: UploadFile) -> str:
    file_path = UPLOAD_DIR / f"{uuid.uuid4().hex}_{file.filename}"
    with file_path.open("wb") as f:
        shutil.copyfileobj(file.file, f)
    return str(file_path)


@app.get("/")
def root():
    return {
        "message": "AR Audit API is running",
        "version": "1.0.0-cloud",
        "environment": os.getenv("NODE_ENV", "production")
    }


def build_ai_conclusion(summary: dict) -> str:
    prompt = f"""
你是一名具有丰富经验的注册会计师，请基于以下应收账款审计数据，
生成一段审计说明与结论。

要求：
1. 使用专业审计语言
2. 风格类似审计底稿结论
3. 包含风险分析和合理判断
4. 控制在200到300字
5. 直接输出正文，不要加标题，不要加引号

数据如下：
应收账款余额：{summary.get("ar_total")}
坏账准备：{summary.get("bad_debt_total")}
账面价值：{summary.get("book_value")}
函证客户数：{summary.get("confirm_customer_count")}
收入：{summary.get("revenue")}
坏账测算金额：{summary.get("bad_debt_calc_total")}
期末余额：{summary.get("related_end_bal")}
"""

    api_key = os.getenv("DEEPSEEK_API_KEY")
    if not api_key:
        return "AI结论生成失败：未配置 DEEPSEEK_API_KEY"
    
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json"
    }
    data = {
        "model": "deepseek-chat",
        "messages": [
            {"role": "system", "content": "你是专业审计师。"},
            {"role": "user", "content": prompt}
        ],
        "temperature": 0.7
    }
    
    try:
        response = requests.post(
            "https://api.deepseek.com/chat/completions",
            headers=headers,
            json=data,
            timeout=30
        )
        result = response.json()
        return result["choices"][0]["message"]["content"]
    except Exception as e:
        return f"AI结论生成失败：{str(e)}"


@app.post("/audit/ar/run")
async def audit_ar_run(
    balance_file: UploadFile = File(...),
    ar_support_file: UploadFile = File(...),
    related_file: UploadFile = File(...),
    contract_liab_file: UploadFile = File(...),
    journal_file: UploadFile = File(...),
    template_file: UploadFile = File(...),
):
    try:
        balance_path = save_upload(balance_file)
        ar_support_path = save_upload(ar_support_file)
        related_path = save_upload(related_file)
        contract_liab_path = save_upload(contract_liab_file)
        journal_path = save_upload(journal_file)
        template_path = save_upload(template_file)

        output_file_name = f"新-应收账款底稿（已填充）_{uuid.uuid4().hex}.xlsx"
        output_path = str(OUTPUT_DIR / output_file_name)

        result = run_ar_audit(
            balance_path=balance_path,
            ar_support_path=ar_support_path,
            related_path=related_path,
            contract_liab_path=contract_liab_path,
            journal_path=journal_path,
            template_path=template_path,
            output_path=output_path,
        )

        if result["status"] != "success":
            raise HTTPException(status_code=500, detail=result["message"])

        summary = result.get("summary", {})
        ai_conclusion = build_ai_conclusion(summary)

        filename = Path(output_path).name
        
        result["download_url"] = f"/audit/ar/download/{filename}"
        result["download_full_url"] = None
        result["ai_conclusion"] = ai_conclusion

        return JSONResponse(
            content=result,
            ensure_ascii=False
        )

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/audit/ar/download/{filename}")
def download_output(filename: str):
    file_path = OUTPUT_DIR / filename
    if not file_path.exists():
        raise HTTPException(status_code=404, detail="文件不存在")
    return FileResponse(
        path=str(file_path),
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.on_event("startup")
async def startup_event():
    import asyncio
    async def cleanup():
        while True:
            await asyncio.sleep(3600)
            cutoff = asyncio.get_event_loop().time() - 86400
            for f in UPLOAD_DIR.glob("*"):
                if f.stat().st_mtime < cutoff:
                    f.unlink()
            for f in OUTPUT_DIR.glob("*"):
                if f.stat().st_mtime < cutoff:
                    f.unlink()
    asyncio.create_task(cleanup())
