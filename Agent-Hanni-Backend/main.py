from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import List
import openai
import os
import base64
from analyzer import analyze_files

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

OPENAI_KEY = os.getenv("OPENAI_API_KEY")

class FileData(BaseModel):
    filename: str
    content_b64: str

class AnalyzeRequest(BaseModel):
    files: List[FileData]

@app.get("/")
def root():
    return {"status": "Agent Hanni Backend running"}

@app.get("/health")
def health():
    return {"status": "ok"}

@app.post("/analyze")
async def analyze(request: AnalyzeRequest):
    if not request.files:
        raise HTTPException(status_code=400, detail="Aucun fichier reçu")

    files_data = []
    for f in request.files:
        try:
            content = base64.b64decode(f.content_b64)
            print(f"Received: '{f.filename}' size={len(content)} bytes")
            files_data.append({'filename': f.filename, 'content': content})
        except Exception as e:
            print(f"Error decoding {f.filename}: {e}")

    if not files_data:
        raise HTTPException(status_code=400, detail="Fichiers invalides")

    try:
        result = analyze_files(files_data)
    except Exception as e:
        print(f"ANALYZE ERROR: {str(e)}")
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Erreur analyse: {str(e)}")

    if OPENAI_KEY and result['alerts']:
        try:
            result = enrich_with_ai(result)
        except Exception as e:
            print(f"OpenAI error: {e}")
            result['globalSummary'] += " (AI indisponible)"

    return result

def enrich_with_ai(result):
    client = openai.OpenAI(api_key=OPENAI_KEY)
    critical = [a for a in result['alerts'] if a['level'] == 'CRITICAL']
    risk = [a for a in result['alerts'] if a['level'] == 'RISK']
    summary_lines = []
    for a in (critical + risk)[:20]:
        summary_lines.append(
            f"- [{a['level']}] {a['customer']} / {a['style']} / {a['color']} "
            f"(xuất: {a['shipDate']}, còn {a['daysToShip']} ngày): {a['issue']}"
        )
    prompt = f"""Bạn là chuyên gia chuỗi cung ứng ngành may mặc tại Việt Nam.
Các vấn đề phát hiện hôm nay:
{chr(10).join(summary_lines)}
Tóm tắt trong 2-3 câu ngắn gọn bằng tiếng Việt, nêu rõ mức độ nghiêm trọng và bộ phận cần hành động ngay."""

    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[{"role": "user", "content": prompt}],
        max_tokens=300,
        temperature=0.2
    )
    result['globalSummary'] = response.choices[0].message.content.strip()
    return result