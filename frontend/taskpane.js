// --- バックエンドのベース URL をここで指定 ---
const BACKEND_URL = "https://localhost:3000"; // 例： https://api.example.com

Office.onReady(({ host }) => {
  if (host !== Office.HostType.PowerPoint) return;
  const btn = document.getElementById("suggestBtn");
  const statusDiv = document.getElementById("status");
  const suggestionDiv = document.getElementById("suggestion");

  btn.addEventListener("click", async () => {
    statusDiv.textContent = "スライド内容を取得中…";
    btn.disabled = true;

    try {
      const [imageBase64] = await Promise.all([fetchSlideImageBase64()]);
      //const text = await fetchSelectedText();
      const text = "画像";
      if (!text) {
        statusDiv.textContent = "テキストが選択されていません。";
        return;
      }

      statusDiv.textContent = "AI サジェスト中…";
      const suggestion = await getAISuggestion(text, imageBase64);
      suggestionDiv.textContent = suggestion;
      statusDiv.textContent = "AI サジェスト完了。";
    } catch (err) {
      console.error(err);
      statusDiv.textContent = `エラー: ${err.message}`;
    } finally {
      btn.disabled = false;
    }
  });
});

async function fetchSelectedText() {
  return new Promise((resolve, reject) => {
    Office.context.document.getSelectedDataAsync(
      Office.CoercionType.Text,
      (res) => {
        if (res.status === Office.AsyncResultStatus.Succeeded) {
          resolve((res.value || "").trim());
        } else {
          reject(res.error);
        }
      }
    );
  });
}

function truncate(str, maxLen) {
  return str.length > maxLen ? str.substring(0, maxLen) + "…" : str;
}

async function fetchSlideImageBase64() {
  return PowerPoint.run(async (context) => {
    const slides = context.presentation.getSelectedSlides();
    const slide = slides.getItemAt(0);

    if (typeof slide.exportAsBase64 !== "function") {
      throw new Error("exportAsBase64() がサポートされていません。");
    }

    // 1) ClientResult<string> を受け取る
    const result = slide.exportAsBase64({ format: "png" });
    // 2) ここで同期
    await context.sync();
    // 3) 実際の文字列を返す
    return result.value;
  });
}

async function getAISuggestion(text, imageBase64) {
  const resp = await fetch(`${BACKEND_URL}/api/suggest`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ text, image_base64: imageBase64 }),
  });
  if (!resp.ok) {
    throw new Error(`バックエンドエラー: ${resp.statusText}`);
  }
  const json = await resp.json();
  return json.suggestion;
}
