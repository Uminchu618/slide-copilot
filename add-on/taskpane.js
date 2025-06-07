const OPENAI_API_KEY = "XXX";

Office.onReady(({ host }) => {
  if (host !== Office.HostType.PowerPoint) return;
  console.log("PowerPoint アドインが読み込まれました");

  const btn = document.getElementById("suggestBtn");
  const statusDiv = document.getElementById("status");
  const suggestionDiv = document.getElementById("suggestion");

  btn.addEventListener("click", async () => {
    statusDiv.textContent = "スライド内容を取得中…";
    btn.disabled = true;

    try {
      // テキストと画像を同時に取得
      const [imageBase64] = await Promise.all([fetchSlideImageBase64()]);
      const text = "AAA";
      if (!text) {
        statusDiv.textContent =
          "テキストが選択されていません。スライド上のテキストボックスを選択してください。";
        return;
      }

      statusDiv.textContent = `検出：「${truncate(
        text,
        30
      )}」 → AI サジェスト中…`;

      // AI に問い合わせ
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

/**
 * 選択テキストを取得する
 * @returns {Promise<string>}
 */
function fetchSelectedText() {
  return new Promise((resolve, reject) => {
    Office.context.document.getSelectedDataAsync(
      Office.CoercionType.Text,
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve((result.value || "").trim());
        } else {
          reject(result.error);
        }
      }
    );
  });
}

/**
 * 文字列を指定長に切り詰める
 * @param {string} str
 * @param {number} maxLen
 * @returns {string}
 */
function truncate(str, maxLen) {
  return str.length > maxLen ? str.substring(0, maxLen) + "…" : str;
}

/**
 * プレビュー API のみを用いてスライドを Base64 PNG として取得
 * サポートされていない場合はエラーを投げる
 * @returns {Promise<string>} Base64 エンコード文字列
 */
async function fetchSlideImageBase64() {
  return PowerPoint.run(async (context) => {
    const slides = context.presentation.getSelectedSlides();
    const slide = slides.getItemAt(0);

    // API 存在チェック
    if (typeof slide.exportAsBase64 !== "function") {
      throw new Error(
        "exportAsBase64() API はこの環境でサポートされていません。Insiders または Web プレビュー版をご利用ください。"
      );
    }

    // PNG フォーマットでエクスポート
    const base64 = await slide.exportAsBase64({ format: "png" });
    return base64;
  });
}

/**
 * OpenAI チャット API にテキストと画像を送信し、提案を取得する
 * @param {string} text
 * @param {string} imageBase64
 * @returns {Promise<string>}
 */
async function getAISuggestion(text, imageBase64) {
  const suggestionDiv = document.getElementById("suggestion");
  suggestionDiv.textContent = "AI からの応答を取得中…";

  const messages = [
    {
      role: "system",
      content:
        "あなたはパワーポイント編集のアシスタントです。スライドのテキストと画像をもとに改善案を提案してください。",
    },
    {
      role: "user",
      content: [
        {
          type: "text",
          text: `以下のスライド内容を改善してください:\n\n${text}`,
        },
        {
          type: "image_url",
          image_url: {
            url: `data:image/png;base64,${imageBase64}`,
            detail: "high",
          },
        },
      ],
    },
  ];

  const response = await fetch("https://api.openai.com/v1/chat/completions", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${OPENAI_API_KEY}`,
    },
    body: JSON.stringify({
      model: "gpt-4o", // vision-enabled モデルを指定
      messages,
      max_tokens: 256,
      temperature: 0.7,
    }),
  });

  if (!response.ok) {
    throw new Error(`OpenAI API エラー: ${response.statusText}`);
  }

  const data = await response.json();
  return data.choices[0].message.content.trim();
}
