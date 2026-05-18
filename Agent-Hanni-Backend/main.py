from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from typing import List
import openai
import os
import json
from analyzer import analyze_files
 
app = FastAPI()
 
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)
 
OPENAI_KEY = os.getenv("OPENAI_API_KEY")
 
@app.get("/")
def root():
    return {"status": "Agent Hanni Backend running"}
 
@app.get("/health")
def health():
    return {"status": "ok"}
 
@app.post("/analyze")
async def analyze(files: List[UploadFile] = File(...)):
    if not files:
        raise HTTPException(status_code=400, detail="Không có file nào được gửi lên")
 
    # Read all files
    files_data = []
    for f in files:
        content = await f.read()
        print(f"Received file: '{f.filename}' size={len(content)} bytes")
        files_data.append({'filename': f.filename, 'content': content})
 
    # Run Python analysis (no AI tokens used here)
    try:
        analysis_result = analyze_files(files_data)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Lỗi phân tích: {str(e)}")
 
    # Only call OpenAI if we have anomalies to explain
    if OPENAI_KEY and analysis_result['alerts']:
        try:
            analysis_result = await enrich_with_ai(analysis_result)
        except Exception as e:
            # If OpenAI fails, return Python analysis anyway
            analysis_result['globalSummary'] += f" (AI không khả dụng: {str(e)})"
 
    return analysis_result
 
async def enrich_with_ai(result):
    """Send only the anomaly summary to OpenAI — NOT the raw data"""
    client = openai.AsyncOpenAI(api_key=OPENAI_KEY)
 
    critical = [a for a in result['alerts'] if a['level'] == 'CRITICAL']
    risk = [a for a in result['alerts'] if a['level'] == 'RISK']
 
    # Build compact summary for AI (max ~800 tokens)
    summary_lines = []
    for a in (critical + risk)[:20]:
        summary_lines.append(
            f"- [{a['level']}] {a['customer']} / {a['style']} / {a['color']} "
            f"(xuất: {a['shipDate']}, còn {a['daysToShip']} ngày): {a['issue']}"
        )
 
    prompt = f"""Bạn là chuyên gia chuỗi cung ứng ngành may mặc tại Việt Nam.
 
Dưới đây là các vấn đề được phát hiện hôm nay:
{chr(10).join(summary_lines)}
 
Tóm tắt tình hình trong 2-3 câu ngắn gọn bằng tiếng Việt, nêu rõ mức độ nghiêm trọng và bộ phận cần hành động ngay.
Chỉ trả lời bằng đoạn văn ngắn, không dùng bullet points."""
 
    response = await client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[{"role": "user", "content": prompt}],
        max_tokens=300,
        temperature=0.2
    )
 
    result['globalSummary'] = response.choices[0].message.content.strip()
    return result