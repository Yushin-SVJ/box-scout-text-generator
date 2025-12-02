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

以下の企業情報を読み込み、あなたの内部で次のプロセスを行ってください。

【内部プロセス（出力しない）】
1. 企業情報を分析し、
   - 公に求めている人材像（ペルソナ）
   - この企業の最大のウリ（Growth Factor）
   - 会社概要の要約（200文字程度）
   を頭の中で整理する。

2. 構造（A/B/C）とモード（1/2）の組み合わせを1つ選ぶ。
   - 構造：
     - A：網羅型（リクルート式）…見出しや箇条書きを用いた情報整理型
     - B：手紙型（ナラティブ式）…セクション区切りを使わない手紙調
     - C：要点直球型（箇条書き式）…冒頭から「ご連絡した理由」を端的に列挙
   - モード：
     - 1：情熱キャリアモード（未来への期待と共感）
     - 2：市場分析モード（冷静なロジック）

3. 選んだパターンに従って、
   - 候補者向けのメール件名（1つ）
   - スカウト本文（1通）
   を作成する。

【事実・表現に関する厳守ルール】
- 企業の業績、社員数、年収水準、具体的な数字などは、
  「入力された企業情報に明記されている内容」以外は推測で書かないこと。
- 「日本発デカコーンを目指す」「非公開求人」「残り2名枠」「年収800万保障」など、
  事実として確認できないインパクトの強い表現は使用禁止。
- 不明な点は、
  「〜な環境であることが多いです」「〜といったキャリアを目指す方に選ばれることが多いポジションです」
  のように、一般論として留めること。

【言葉遣い・トーンのルール】
- 名前プレースホルダー禁止：
  「〇〇様」「{Name}」などの変数は一切使わない。
  代わりに「これまでのご経歴」「貴殿のご実績」「そうした知見をお持ちの方」などの表現を使う。
- 「〜しちゃいましょう！」「〜なんです！」のような馴れ馴れしい口語体は禁止。
- 「絶対成功します」「間違いありません」といった強い断定は禁止。
  代わりに「〜を実現できる希少な環境です」「〜なキャリアを描ける可能性が高いです」と表現する。
- どのモードでも、信頼できるキャリアパートナーとしての品格（丁寧語・謙譲語）を保つこと。

【構成に関する詳細ルール】
- A：網羅型
  - 挨拶 → 企業紹介 → オススメポイント（箇条書き）→ クロージング → 固定フッター
  1. 挨拶・導入
    - 挨拶：初めまして。（改行）株式会社BOXの{担当者名}と申します。
    - 導入：「ご紹介する《企業名》は…」から始め、候補者が得られるメリットを提示する。
  2. 会社概要
    - 見出し：{企業名}について（前後に空白行を入れる）
    - 内容：事業内容やフェーズを要約。
  3. オススメポイント
    - 見出し：オススメポイント（前後に空白行を入れる）
    - (1) (2) のリスト形式。間に必ず空白行を入れる。
    - フォーマット：
      - (1){見出し}：{説明}
    - 情熱モード時の注意: 「すごいですよ！」ではなく「〜という稀有なチャンスです」「〜を実現できる環境です」と表現する。

- B：手紙型
  1. フック（個別化）
    - 名前は呼ばず、「ご経歴を拝見し、〜という点に魅力を感じました」から入る。
  2. 課題と再定義
     見出し記号は使わない。
   - 「今、{企業名}は〜という面白いフェーズにあります」と文章で繋ぐ。
  3. 接続と提案
    - 名前は呼ばず、条件接続を行う。
    - OK例: 「だからこそ、{定義したスキル}をお持ちの方の力が不可欠なのです」
    - 箇条書きは使わず、段落分け（空白行）で読みやすくする。

- C：要点直球型
  1.単刀直入な導入
    - 挨拶の後、すぐに本題に入る。
    - OK例: 「今回ご連絡したのは、以下の3つの理由から、これまでのご実績と{企業名}の相性が最高だと確信したためです。」

  2.3つの理由リスト
    - 見出し記号は使わない。
    - "・"（全角ナカグロ）を使った3点のリストにする。間に空白行を入れる。
    - 内容: スキルマッチ、企業の成長性、キャリアメリットを端的に述べる。

  3.結び
    - 「詳細な資料もございますので…」と簡潔に締める。

【固定フッター】
本文の最後には、次のフッターをそのまま挿入してください（改変禁止・{担当者名}はそのまま出力）:

${FIXED_FOOTER}

【出力フォーマット（厳守）】
- 以下のJSON形式「のみ」を返してください。
- 説明文やコメント、コードブロック記号（\`\`\`）などは一切出力しないでください。
- 太字()

{
  "company_name": "企業名をここに",
  "subject": "ここにメール件名を1つ",
  "body": "ここに本文全体"
}

【入力された企業情報】
- 企業名: ${companyName}
- 企業情報・URL・求人要約など:
${infoText}
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
