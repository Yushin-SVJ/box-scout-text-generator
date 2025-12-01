/***** 設定 *****/
const SHEET_NAME = 'シート1'; // シート名
const MAX_PER_RUN = 10;         // 1回の実行で処理する最大件数

// 使用するGeminiモデル
const GEMINI_MODEL = 'gemini-2.5-flash';
const GEMINI_ENDPOINT =
  `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent`;

/**
 * メイン関数：
 * - シートのA,B列を読み取り
 * - C,D列が空 & ステータスが未/空の行を対象に
 * - 最大 MAX_PER_RUN 件まで順番にGeminiを叩いて件名/本文を埋める
 */
function generateScoutMails() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    throw new Error(`シート「${SHEET_NAME}」が見つかりません。`);
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('データ行がありません。');
    return;
  }

  // A〜E列をまとめて取得
  const values = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  // [A:企業名, B:企業情報, C:件名, D:本文, E:ステータス]

  let processed = 0;

  for (let i = 0; i < values.length; i++) {
    if (processed >= MAX_PER_RUN) break;

    const rowIndex = i + 2; // シート上の行番号
    const [companyName, companyInfo, subject, body, status] = values[i];

    // 企業名なし or 既にdone or 件名/本文が埋まっている → スキップ
    if (!companyName) continue;
    if (status === 'done') continue;
    if (subject && body) continue;

    const prompt = buildPromptForCompany(companyName, companyInfo);
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

    const outCompany = json.company_name || companyName;
    const outSubject = json.subject || '';
    const outBody = json.body || '';

    // シートに書き込み
    sheet.getRange(rowIndex, 1).setValue(outCompany); // A:企業名（整形されれば上書き）
    sheet.getRange(rowIndex, 3).setValue(outSubject); // C:件名
    sheet.getRange(rowIndex, 4).setValue(outBody);    // D:本文
    sheet.getRange(rowIndex, 5).setValue('done');     // E:ステータス

    processed++;
  }

  Logger.log(`処理完了：${processed} 件を更新しました。`);
}

/**
 * 企業ごとのプロンプトを生成
 * - ここに「STEP1〜3をAIの内部でやらせる」設計を埋め込んでいる
 */
function buildPromptForCompany(companyName, companyInfo) {
  const infoText = companyInfo && companyInfo.trim()
    ? companyInfo.trim()
    : '企業URLや求人要約などの詳細情報は与えられていません。一般的な情報に基づき、過度な推測は避けてください。';

  const FIXED_FOOTER = `
  ▼ 私が提供できる価値

コロナ禍明けから選考ハードルが高い状態ですが、直近3ヶ月で書類通過率84％（※転職平均30％）を実現しており、72％の方が内定獲得しています。

その要因としては、下記の2点がございます。

【1】企業別の面接対策を行い、希望者に合計10回以上の徹底的な言語化のサポート
　┗※過去の面接データや非公開情報を元に対策し、年収アップの転職を実現しています。

【2】会計事務所として関わり、経営陣との距離感がかなり近いため口添えできる
　┗※本来は書類見送りの方も、弊社の紹介であれば選考を通して頂いております。

■■■■■■■■■■■■■■■■■■■
▽その他ご提案先の一部をご紹介します▽
マネーフォワード / kubell / ユーザベース
LegalForce / SATORI / オープンエイト
カミナシ / Sales Marker / プレイド
アンドパッド / LayerX / ヤプリ / スタディスト / freee
ビットキー / BASE / ROXX / プレックス
dely / ポジウィル / フェズ
スマートニュース / ビズリーチ
Speee / メルカリ / レバレジーズ

◆直近の私の支援実績 ※一部のみ抜粋◆
------------------------------
(1) 30歳 / 男性 / Sier（年収590万円）
　┗▶︎ 大手SaaS / FS（年収700万円）
(2) 32歳 / 女性 / 未上場SaaS（年収540万円）
　┗▶︎ 上場SaaS / CS（年収640万円）
(3) 33歳 / 女性 / 大手百貨店マネージャー（年収480万円）
　┗▶︎ 大手人材系 / 新規事業部 / 法人営業（年収500万円）
(4) 35歳 / 女性 / 大手メディア / 営業マネージャー（年収900万円）
　┗▶︎ 上場SaaS / 営業マネージャー（年収1,000万円）

◆面談について◆
・面談手法：全てWEB完結です
・所要時間：30分程度
週によっては土日祝もご対応可能です。

//////////////////////////////////////////////
株式会社BOX
採用支援事業部マネージャー
{担当者名}
〒150-0031 東京都渋谷区桜丘町9－8 ＫＮ渋谷3ビル 2F
//////////////////////////////////////////////
`.trim();

  const prompt = `
あなたは、人材紹介会社「株式会社BOX」のスカウト文面作成パートナーです。

求職者に「刺さる」スカウトを作るため、以下のプロセスを「内部で」自動的に実行してください。
ユーザーへの質問や確認は不要です（すべて自動処理してください）。

-------------------------------------
【内部プロセス（出力しない）】
1. 企業情報を読み取り、次を整理する（推論は控えめに）。
   - 企業が求めている人材像（ペルソナ）
   - 企業の「最大のウリ（Growth Factor）」
   - 会社概要（200文字程度）

2. 元プロンプトに基づき、最適と思われる
   - 構造（A/B/C）
   - モード（1/2）
   の組み合わせを1つ選択する。
   ※ユーザーへの確認なしで自動で選ぶ。

3. 選んだ組み合わせに従い、件名1つと本文1通を生成する。

-------------------------------------
【構造ルール】

■ A：網羅型（リクルート式）
- 挨拶 → 企業紹介 → オススメポイント（番号つき）→ クロージング
- 見出しに「◆」「【】」を使用可
- ポイントは (1)(2) の形式

■ B：手紙型（ナラティブ式）
- セクション見出しなし
- 手紙調で自然な文章構造
- 段落で読みやすく

■ C：要点直球型（箇条書き）
- 挨拶のあとに「今回ご連絡した理由」を箇条書きで3点
- 全角「・」を使う
- その後簡潔に締める

-------------------------------------
【モードルール】

■ 1：情熱キャリアモード
- 丁寧語を保ちつつ前向き
- 強い断定は禁止（代わりに「可能性が高い」「希少な環境です」）
- 口語体禁止

■ 2：市場分析モード
- ロジカル・客観的に表現
- 感情表現を減らす
- 企業フェーズや市場特性を整理して伝える

-------------------------------------
【共通禁止事項（厳守）】
- 名前のプレースホルダー禁止（〇〇様、{Name} など）
- 太文字（**）禁止
- 年収などの変動要素禁止（例：「800万」）
- 具体的職種・部署名禁止（配属リスク回避）
- HR用語（ビジネス職・総合職など）禁止
- 推測での断言禁止（デカコーンに成長中！ など）

-------------------------------------
【件名の作成ルール】
構造の異なる以下の3パターンを比較し、今回の企業に最適と判断した“1パターンだけ”を採用し件名を1つ作成する。

- ミッション/進化型（企業の挑戦テーマ重視）
- タグ/ブランド型（【 】で並べる形式）
- インパクト/機会型（フェーズの希少性）

※禁止表現は上記ルールに従う。

-------------------------------------
【本文の最後】
以下の固定フッターを必ず「そのまま」挿入する（改変禁止・{担当者名}もそのまま出力）：

${FIXED_FOOTER}

-------------------------------------
【出力形式（厳守）】
以下の JSON のみを返してください。

{
  "company_name": "企業名",
  "subject": "件名1つ",
  "body": "本文全文（フッター含む）"
}

【入力情報】
企業名: ${{companyName}}
企業情報:
${{infoText}}


`.trim();

  return prompt;
}

/**
 * Gemini APIを呼び出す
 * @param {string} prompt
 * @returns {string|null} モデルの生テキストレスポンス
 */
function callGemini(prompt) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    throw new Error('GEMINI_API_KEY がスクリプトプロパティに設定されていません。');
  }

  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    // 必要に応じて temperature などを追加
    // generationConfig: { temperature: 0.3 }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-goog-api-key': apiKey
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  const res = UrlFetchApp.fetch(GEMINI_ENDPOINT, options);
  const code = res.getResponseCode();
  const text = res.getContentText();

  if (code !== 200) {
    Logger.log('Gemini API error: ' + code + ' ' + text);
    return null;
  }

  const data = JSON.parse(text);
  const candidates = data.candidates;
  if (!candidates || !candidates.length) {
    Logger.log('Gemini: candidates が空です。');
    return null;
  }

  const content = candidates[0].content;
  if (!content || !content.parts || !content.parts.length) {
    Logger.log('Gemini: content.parts が空です。');
    return null;
  }

  // 通常は最初のpartにテキストが入る想定
  return content.parts[0].text.trim();
}

/**
 * モデルの出力テキストからJSON部分を抽出してパースする
 * - ```json ... ``` で囲まれていてもOK
 * - テキストの中から最初の { 〜 最後の } を抜き出してJSON.parse
 */
function parseResultJson(text) {
  if (!text) return null;

  let cleaned = text.trim();

  // コードブロックを除去
  if (cleaned.startsWith('```')) {
    // ```xxx の1行目を削除
    cleaned = cleaned.replace(/^```[\s\S]*?\n/, '');
    // 末尾の ``` を削除
    cleaned = cleaned.replace(/```[\s\S]*$/, '');
    cleaned = cleaned.trim();
  }

  // テキスト中の最初の { 〜 最後の } を抜き出す
  const firstBrace = cleaned.indexOf('{');
  const lastBrace = cleaned.lastIndexOf('}');
  if (firstBrace === -1 || lastBrace === -1 || lastBrace <= firstBrace) {
    try {
      // そのままJSONとしてパースできるか一応試す
      return JSON.parse(cleaned);
    } catch (e) {
      Logger.log('JSON抽出失敗: ' + cleaned);
      return null;
    }
  }

  const jsonString = cleaned.substring(firstBrace, lastBrace + 1);
  try {
    return JSON.parse(jsonString);
  } catch (e) {
    Logger.log('JSONパース失敗: ' + jsonString);
    return null;
  }
}
