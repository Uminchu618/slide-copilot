const OPENAI_API_KEY = "XXX";

Office.onReady(({ host }) => {
  if (host !== Office.HostType.PowerPoint) return;
  console.log("PowerPoint アドインが読み込まれました");
  const btn = document.getElementById("suggestBtn");
  const statusDiv = document.getElementById("status");
  const suggestionDiv = document.getElementById("suggestion");

  btn.addEventListener("click", async () => {
    statusDiv.textContent = "選択テキストを取得中…";
    try {
      const text = await fetchSelectedText();
      if (!text) {
        statusDiv.textContent =
          "テキストが選択されていません。スライド上のテキストボックスを選択してください。";
        return;
      }

      statusDiv.textContent = `検出：「${truncate(
        text,
        30
      )}」 → AI サジェスト中…`;
      btn.disabled = true;

      await getAISuggestion(text);
      statusDiv.textContent = "AI サジェスト完了。";
    } catch (err) {
      console.error(err);
      statusDiv.textContent =
        "エラーが発生しました。コンソールを確認してください。";
    } finally {
      btn.disabled = false;
    }
  });
});

/**
 * 選択テキスト取得
 */
function fetchSelectedText() {
  return new Promise((resolve, reject) => {
    Office.context.document.getSelectedDataAsync(
      Office.CoercionType.Text,
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const txt = result.value?.trim() || "";
          resolve(txt);
        } else {
          reject(result.error);
        }
      }
    );
  });
}

function truncate(str, maxLen) {
  return str.length > maxLen ? str.substring(0, maxLen) + "…" : str;
}

/**
 * OpenAI API 呼び出し
 */
async function getAISuggestion(text) {
  const suggestionDiv = document.getElementById("suggestion");
  suggestionDiv.textContent = "AI からの応答を取得中…";

  const response = await fetch("https://api.openai.com/v1/chat/completions", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${OPENAI_API_KEY}`,
    },
    body: JSON.stringify({
      model: "gpt-4.1-nano-2025-04-14",
      messages: [
        {
          role: "system",
          content:
            "あなたはパワーポイント編集のアシスタントです。選択されたテキストを改善する提案を行ってください。",
        },
        {
          role: "user",
          content: `以下のテキストをより分かりやすく、伝わりやすくするための提案をください:\n\n"${text}"`,
        },
      ],
      max_tokens: 256,
      temperature: 0.7,
    }),
  });

  if (!response.ok) {
    throw new Error(`OpenAI API エラー: ${response.statusText}`);
  }
  const data = await response.json();
  suggestionDiv.textContent = data.choices[0].message.content.trim();
}
