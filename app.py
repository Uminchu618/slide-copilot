import os
from dotenv import load_dotenv
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
import openai

# 環境変数ロード
load_dotenv(dotenv_path=".env")
openai.api_key = os.getenv("OPENAI_API_KEY")


# ─── リクエストボディ定義 ─────────────────────────────
class SuggestRequest(BaseModel):
    text: str
    image_base64: str


app = FastAPI()

# ─── CORS 設定（開発用。必要に応じて allow_origins を限定） ────
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["GET", "POST", "OPTIONS"],
    allow_headers=["*"],
)


# ─── OpenAI 呼び出し用エンドポイント ───────────────────────
@app.post("/api/suggest")
async def suggest(req: SuggestRequest):
    try:
        # messages = [
        #     {
        #         "role": "system",
        #         "content": "あなたはパワーポイント編集のアシスタントです。スライドのテキストと画像をもとに改善案を提案してください。",
        #     },
        #     {
        #         "role": "user",
        #         "content": f"テキスト:\n{req.text}",
        #     },
        # ]
        # resp = openai.ChatCompletion.create(
        #     model="gpt-4o",
        #     messages=messages,
        #     max_tokens=256,
        #     temperature=0.7,
        # )
        # suggestion = resp.choices[0].message.content.strip()
        # return {"suggestion": suggestion}
        return {"suggestion": "OK"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


# ─── 静的ファイル（frontend ディレクトリ）を配信 ────────────
# __file__ = /path/to/backend/app.py
frontend_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "frontend"))
app.mount("/", StaticFiles(directory=frontend_dir, html=True), name="frontend")
