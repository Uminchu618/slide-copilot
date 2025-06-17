// --- バックエンドのベース URL をここで指定 ---
const BACKEND_URL = "https://localhost:3000"; // 例： https://api.example.com

Office.onReady(({ host }) => {
  if (host !== Office.HostType.PowerPoint) return;

  const btn = document.getElementById("actionBtn");
  const statusDiv = document.getElementById("status");
  const suggestionDiv = document.getElementById("suggestion");
  const textsContainer = document.getElementById("texts-container");
  const imagesContainer = document.getElementById("images-container");

  btn.addEventListener("click", async () => {
    statusDiv.textContent = "スライド内容を取得中…";
    // 表示エリアをクリア
    textsContainer.innerHTML = "";
    imagesContainer.innerHTML = "";
    suggestionDiv.textContent = "AI サジェスト結果がここに表示されます。";
    btn.disabled = true;

    try {
      // テキストデータと画像データを並行して取得
      const [allText, allImages] = await Promise.all([
        fetchAllTextFromSlide(),
        fetchAllImagesFromSlide(),
      ]);

      // 取得したテキストをHTMLに列挙
      if (allText.trim()) {
        const textArray = allText.split("\n");
        textArray.forEach((text) => {
          const p = document.createElement("p");
          p.textContent = text;
          textsContainer.appendChild(p);
        });
      } else {
        textsContainer.textContent =
          "スライドにテキストが見つかりませんでした。";
      }

      // 取得した画像をHTMLに列挙
      if (allImages.length > 0) {
        allImages.forEach((base64Image) => {
          const img = document.createElement("img");
          img.src = `data:image/png;base64,${base64Image}`;
          imagesContainer.appendChild(img);
        });
      } else {
        imagesContainer.textContent = "スライドに画像が見つかりませんでした。";
      }

      statusDiv.textContent = "データの取得が完了しました。";

      if (allText.trim() || allImages.length > 0) {
        statusDiv.textContent = "AI サジェスト中…";
        // 複数の画像を送信する場合、getAISuggestion関数の修正が必要です
        const suggestion = await getAISuggestion(allText, allImages);
        suggestionDiv.textContent = suggestion;
        statusDiv.textContent = "AI サジェスト完了。";
      }
    } catch (err) {
      console.error(err);
      statusDiv.textContent = `エラー: ${err.message}`;
    } finally {
      btn.disabled = false;
    }
  });
});

/**
 * 現在選択されているスライドからすべてのテキストを取得します。（堅牢性向上版）
 * @returns {Promise<string>} スライド内のすべてのテキストを改行で連結した文字列。
 */
async function fetchAllTextFromSlide() {
  return PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const shapes = slide.shapes;
    shapes.load("items/id,items/type,items/textFrame/hasText");
    await context.sync();
    console.log(shapes);
    const textShapes = shapes.items.filter(
      (shape) => shape.textFrame.hasText === true
    );
    console.log(`スライド内のテキストシェイプ数: ${textShapes.length}`);
    if (textShapes.length === 0) {
      return "";
    }
    textShapes.forEach((shape) => shape.textFrame.textRange.load("text"));
    await context.sync();
    const allText = textShapes
      .map((shape) => shape.textFrame.textRange.text.trim())
      .filter((t) => t.length > 0);
    console.log(`スライド内のテキスト: ${allText}`);

    return allText.join("\n");
  });
}

async function fetchAllImagesFromSlide() {
  return PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    // height/width は任意。指定しなければ本来のサイズで取得
    const imageResult = slide.getImageAsBase64({ height: 300 });
    await context.sync();
    return [imageResult.value]; // Base64文字列
  });
}
/**
 * バックエンドにテキストと画像を送信し、AIによるサジェストを取得します。
 * @param {string} text 提案の基になるテキスト。
 * @param {string[]} imagesBase64 スライドのBase64エンコードされた画像の配列。
 * @returns {Promise<string>} AIからのサジェスト文字列。
 */
async function getAISuggestion(text, imageBase64) {
  const resp = await fetch(`${BACKEND_URL}/api/suggest`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    // バックエンドのAPI仕様に応じて送信するデータの形式を調整してください
    body: JSON.stringify({ text, image_base64: imageBase64 }),
  });
  if (!resp.ok) {
    const errorBody = await resp.text();
    throw new Error(
      `バックエンドエラー: ${resp.status} ${resp.statusText} - ${errorBody}`
    );
  }
  const json = await resp.json();
  return json.suggestion;
}
