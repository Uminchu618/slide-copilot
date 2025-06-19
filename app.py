import os
from dotenv import load_dotenv
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
from openai import OpenAI
import traceback

# 環境変数ロード
# load_dotenv(dotenv_path=".env")
client = OpenAI(api_key=OPENAI_API_KEY)


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
        text_block = {
            "type": "text",
            "text": req.text,
        }
        image_block = {
            "type": "image_url",
            "image_url": {"url": f"data:image/png;base64,{req.image_base64}"},
        }
        messages = [
            {
                "role": "system",
                "content": """あなたは、PowerPointスライドの図表引用チェッカーです。ユーザーから[text_block, image_block]形式でメッセージが渡されます。以下の手順で応答してください。
1. text_blockにはスライド内の文字データが含まれており、image_blockにはスライド全体の画像が含まれています。  
2. image_blockを解析し、図や表が外部資料から引用されている可能性を検出します。  
3. text_block内に既に引用元が記載されている図表は除外します。  
4. 検出した未引用の図表について、著者名・出版年・タイトル・出典（出版社やURL等）をもとに、APAスタイルの引用文献を生成します。  
5. 出力は箇条書きの参考文献リストのみとし、余分な説明は不要です。  
6. 新たに引用元が検出されない場合は「新たな引用元は検出されませんでした。」とだけ返してください。
7. 出力は日本語で行ってください。""",
            },
            {
                "role": "user",
                "content": [text_block, image_block],
            },
        ]
        resp = client.chat.completions.create(
            model="gpt-4.1-2025-04-14",
            messages=messages,
            max_tokens=1024 * 8,
            temperature=0.7,
        )
        suggestion = resp.choices[0].message.content.strip()
        return {"suggestion": suggestion}
    except Exception as e:
        tb = traceback.format_exc()
        logging.error("Suggest API failed:\n%s", tb)
        raise HTTPException(status_code=500, detail=str(e))


# ─── 静的ファイル（frontend ディレクトリ）を配信 ────────────
# __file__ = /path/to/backend/app.py
frontend_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "frontend"))
app.mount("/", StaticFiles(directory=frontend_dir, html=True), name="frontend")
