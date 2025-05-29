import os
import io
import json
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload
from google.oauth2 import service_account

# サービスアカウントファイルとスコープ
SERVICE_ACCOUNT_FILE = "service_account.json"
SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/presentations.readonly",
]

# アップロード&画像保存先
output_dir = "gs_slides_images"
os.makedirs(output_dir, exist_ok=True)


def upload_pptx_and_convert_to_slides(pptx_path):
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES
    )
    drive_service = build("drive", "v3", credentials=creds)

    file_metadata = {
        "name": os.path.splitext(os.path.basename(pptx_path))[0],
        "mimeType": "application/vnd.google-apps.presentation",
    }
    media = MediaFileUpload(
        pptx_path,
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        resumable=True,
    )

    print("PPTXファイルをアップロードしてGoogle Slidesに変換中…")
    file = (
        drive_service.files()
        .create(body=file_metadata, media_body=media, fields="id")
        .execute()
    )
    slides_file_id = file.get("id")
    print(f"Google SlidesのID: {slides_file_id}")
    return slides_file_id


def export_each_slide_as_image(slides_file_id):
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES
    )
    slides_service = build("slides", "v1", credentials=creds)

    # プレゼンテーション情報取得
    presentation = (
        slides_service.presentations().get(presentationId=slides_file_id).execute()
    )
    slides = presentation.get("slides")
    image_paths = []

    for idx, slide in enumerate(slides):
        slide_object_id = slide.get("objectId")
        # PNG画像のエクスポートURL
        export_url = f"https://slides.googleapis.com/v1/presentations/{slides_file_id}/pages/{slide_object_id}/thumbnail?thumbnailProperties.mimeType=PNG&thumbnailProperties.thumbnailSize=LARGE"

        # 認証トークン取得
        creds = service_account.Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE, scopes=SCOPES
        )
        authed_session = creds.authorize(google.auth.transport.requests.Request())
        import requests

        headers = {"Authorization": f"Bearer {creds.token}"}
        r = requests.get(export_url, headers=headers)
        if r.status_code != 200:
            print(f"スライド{idx+1}のサムネイル取得失敗: {r.text}")
            continue
        thumbnail_url = r.json()["contentUrl"]
        # 画像データ取得
        img_response = requests.get(thumbnail_url)
        img_path = os.path.join(output_dir, f"slide_{idx+1}.png")
        with open(img_path, "wb") as f:
            f.write(img_response.content)
        print(f"保存: {img_path}")
        image_paths.append(img_path)
    return image_paths


# --- 使用例 ---
pptx_path = "example.pptx"

# 1. PPTXをアップロードしてGoogle Slidesに変換
slides_file_id = upload_pptx_and_convert_to_slides(pptx_path)

# 2. スライドを画像としてダウンロード
slide_images = export_each_slide_as_image(slides_file_id)

# 3. 画像リストをjson保存
with open(os.path.join(output_dir, "slide_images.json"), "w", encoding="utf-8") as f:
    json.dump(slide_images, f, ensure_ascii=False, indent=2)
