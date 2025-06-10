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

      // --- AIへの問い合わせ（コメントアウト） ---
      // if (allText.trim() || allImages.length > 0) {
      //   statusDiv.textContent = "AI サジェスト中…";
      //   // 複数の画像を送信する場合、getAISuggestion関数の修正が必要です
      //   const suggestion = await getAISuggestion(allText, allImages);
      //   suggestionDiv.textContent = suggestion;
      //   statusDiv.textContent = "AI サジェスト完了。";
      // }
      // --- ここまで ---
    } catch (err) {
      console.error(err);
      statusDiv.textContent = `エラー: ${err.message}`;
    } finally {
      btn.disabled = false;
    }
  });
});

/**
 * 現在選択されているスライドからすべてのテキストを取得します。
 * @returns {Promise<string>} スライド内のすべてのテキストを改行で連結した文字列。
 */
async function fetchAllTextFromSlide() {
  return PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const shapes = slide.shapes;
    shapes.load("items/textFrame/textRange/text");

    await context.sync();

    const allText = [];
    shapes.items.forEach((shape) => {
      try {
        const text = shape.textFrame.textRange.text.trim();
        if (text) {
          allText.push(text);
        }
      } catch (error) {
        // テキストフレームを持たない図形の場合はエラーを無視
      }
    });

    return allText.join("\n");
  });
}

/**
 * 現在選択されているスライドからすべての画像データをBase64形式の配列として取得します。
 * @returns {Promise<string[]>} Base64形式の画像データ配列。
 */
async function fetchAllImagesFromSlide() {
  return PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const shapes = slide.shapes;
    // 図形の種類とPictureオブジェクトをロード
    shapes.load("items/type, items/picture");

    await context.sync();

    // "Image"タイプの図形のみをフィルタリング
    const imageShapes = shapes.items.filter((shape) => shape.type === "Image");

    if (imageShapes.length === 0) {
      return [];
    }

    // 各画像図形からBase64データを取得するリクエストを作成
    const imageBase64Results = imageShapes.map((shape) =>
      shape.picture.getImageAsBase64(PowerPoint.PictureFormat.Png)
    );

    // すべての画像データ取得リクエストを同期
    await context.sync();

    // ClientResultオブジェクトから実際のBase64文字列を抽出して返す
    return imageBase64Results.map((result) => result.value);
  });
}

/**
 * バックエンドにテキストと画像を送信し、AIによるサジェストを取得します。
 * @param {string} text 提案の基になるテキスト。
 * @param {string[]} imagesBase64 スライドのBase64エンコードされた画像の配列。
 * @returns {Promise<string>} AIからのサジェスト文字列。
 */
async function getAISuggestion(text, imagesBase64) {
  const resp = await fetch(`${BACKEND_URL}/api/suggest`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    // バックエンドのAPI仕様に応じて送信するデータの形式を調整してください
    body: JSON.stringify({ text, images_base64: imagesBase64 }),
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

// 以下の関数は直接使用していませんが、参考のために残しています。
/**
 * スライドのスクリーンショットをBase64エンコードされた文字列として取得します。
 * @returns {Promise<string>} Base64形式の画像データ。
 */
async function fetchSlideImageBase64() {
  return PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const result = slide.exportAsBase64({ format: "png" });
    await context.sync();
    return result.value;
  });
}
