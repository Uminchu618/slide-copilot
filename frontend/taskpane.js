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
 * 現在選択されているスライドからすべてのテキストを取得します。（堅牢性向上版）
 * @returns {Promise<string>} スライド内のすべてのテキストを改行で連結した文字列。
 */
async function fetchAllTextFromSlide() {
  return PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const shapes = slide.shapes;

    // ステップ1: テキストフレームを持つシェイプのIDを安全に特定する
    // 'hasTextFrame' は、テキストの有無を安全に確認できるプロパティです。
    shapes.load("items/id, items/hasTextFrame");
    await context.sync();

    const textShapeIds = [];
    shapes.items.forEach((shape) => {
      // 'hasTextFrame'がtrueのシェイプのIDを収集
      if (shape.hasTextFrame) {
        textShapeIds.push(shape.id);
      }
    });

    // テキストを持つシェイプがなければ、ここで処理を終了
    if (textShapeIds.length === 0) {
      return "";
    }

    // ステップ2: 特定したIDのシェイプからのみ、テキストを一括で読み込む
    const allText = [];
    const textRangesToLoad = [];

    for (const id of textShapeIds) {
      // IDを使ってシェイプを取得し、そのテキスト範囲をロード対象にする
      const shape = shapes.getItem(id);
      const textRange = shape.textFrame.textRange;
      textRange.load("text");
      textRangesToLoad.push(textRange);
    }

    // ロード対象としたすべてのテキスト範囲を一度に同期
    await context.sync();

    // 同期完了後、各テキスト範囲からテキストを抽出
    textRangesToLoad.forEach((textRange) => {
      const text = textRange.text.trim();
      if (text) {
        allText.push(text);
      }
    });

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
 * シェイプコレクションを再帰的に探索し、画像シェイプのIDを収集するヘルパー関数。
 * @param {PowerPoint.RequestContext} context - 現在のコンテキスト。
 * @param {PowerPoint.ShapeCollection} shapes - 探索対象のシェイプコレクション。
 * @param {string[]} imageShapeIds - 見つかった画像IDを格納する配列。
 */
async function findImagesRecursively(context, shapes, imageShapeIds) {
  // 処理に必要なプロパティをロード
  // group.shapes をロードするために 'group' が必要
  shapes.load("items/id, items/type, items/group/shapes");
  await context.sync();

  for (const shape of shapes.items) {
    if (shape.type === "Image") {
      imageShapeIds.push(shape.id);
    } else if (shape.type === "Group") {
      // グループシェイプが見つかったら、その中のシェイプに対してこの関数を再度呼び出す（再帰）
      await findImagesRecursively(context, shape.group.shapes, imageShapeIds);
    }
  }
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
