// ...existing code...

/**
 * 企業情報の言語化のみを行うテスト関数（スカウト文面は生成しない）
 * - A,B列の企業名・企業情報を読み取り
 * - Geminiに「ペルソナ・Growth Factor・会社概要200字」をJSONで返させ、G列に保存
 */
function generateCompanyAnalyses() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error(`シート「${SHEET_NAME}」が見つかりません。`);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('データ行がありません。');
    return;
  }

  const values = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  let processed = 0;

  for (let i = 0; i < values.length; i++) {
    const rowIndex = i + 2;
    const [companyName, companyInfo, , , status, , analysisJson] = values[i];

    if (!companyName) continue;
    if (analysisJson) continue; // 既に分析済みならスキップ

    const prompt = buildAnalysisPrompt(companyName, companyInfo);
    const responseText = callGemini(prompt);
    if (!responseText) {
      Logger.log(`Row ${rowIndex}: Geminiからレスポンスが得られませんでした。`);
      continue;
    }

    const json = parseResultJson(responseText);
    if (!json) {
      Logger.log(`Row ${rowIndex}: JSON抽出/パースに失敗しました: ${responseText}`);
      continue;
    }

    const payload = JSON.stringify(json, null, 2);
    sheet.getRange(rowIndex, 7).setValue(payload); // G列: 分析結果JSON
    Logger.log(`Row ${rowIndex}: 分析のみ保存しました。`);
    processed++;
  }

  Logger.log(`分析処理完了: ${processed} 件を更新しました。`);
}

// ...existing code...
// ...existing code...

function buildAnalysisPrompt(companyName, companyInfo) {
  const infoText = companyInfo && companyInfo.trim()
    ? companyInfo.trim()
    : '企業URLや求人要約などの詳細情報は与えられていません。一般論に留めてください。';

  return `
あなたは、人材紹介会社「株式会社BOX」のスカウト文面作成パートナーです。

以下の企業情報を読み込み、【STEP1：情報分析】のみを実施してください。スカウト文面や件名は出力してはいけません。

【出力すべきJSONフィールド】
{
  "company_name": "入力企業名をそのまま記載",
  "persona": [
    "この企業が公に求めている人材像を3項目以内で箇条書き（経験・スキル・価値観など）"
  ],
  "growth_factor": "この企業の最大のウリを1段落（150字以内）で説明",
  "summary_200chars": "会社概要を200文字程度で要約"
}

【厳守事項】
- 入力に無い数値や実績を捏造しない。
- 一般論を書く場合は「〜といったケースが多い」のように表現を弱める。
- 文字装飾は禁止。JSON以外のテキストを返さない。
- company_name は必ず ${companyName} を返す。

【入力された企業情報】
- 企業名: ${companyName}
- 企業情報・URL・求人要約など:
${infoText}
`.trim();
}
// ...existing code...