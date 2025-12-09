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
  try {
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      throw new Error(`シート「${SHEET_NAME}」が見つかりません。`);
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      Logger.log('データ行がありません。');
      return;
    }

    const values = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
    let processed = 0;
    let recentPatterns = []; // 直近のパターンを記録（多様性確保用）

    for (let i = 0; i < values.length; i++) {
      if (processed >= MAX_PER_RUN) break;

      const rowIndex = i + 2;
      const [companyName, , subject, body, status] = values[i];

      if (!companyName) continue;
      if (status === 'done') continue;
      if (subject && body) continue;

      const companyInfo = fetchCompanyProfile(String(companyName).trim());
      if (!companyInfo) {
        Logger.log(`Row ${rowIndex}: 企業情報を取得できませんでした（${companyName}）。`);
        continue;
      }

      // 直近2件のパターンを渡して多様性を確保
      const prompt = buildPromptForCompany(companyName, companyInfo, recentPatterns);
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

      let patternReason = '';
      if (json.pattern_reason && String(json.pattern_reason).trim()) {
        patternReason = String(json.pattern_reason).trim();
      } else if (json.reason && String(json.reason).trim()) {
        patternReason = String(json.reason).trim();
      } else {
        const m = responseText.match(/理由[:：]\s*([\s\S]*?)(?:\n\s*\n|$)/);
        if (m && m[1]) patternReason = m[1].trim();
      }
      if (patternReason) {
        Logger.log(`Row ${rowIndex}: モデルの選択理由: ${patternReason}`);
      } else {
        Logger.log(`Row ${rowIndex}: モデルの選択理由は出力されていませんでした。`);
      }

      const outCompany = json.company_name || companyName;
      const outSubject = json.subject || '';
      const outBody = json.body || '';
      const outPattern = resolvePatternIdentifier(json.pattern, responseText, rowIndex);

      // 直近パターンを更新（多様性確保用）
      if (outPattern) {
        const structureOnly = outPattern.charAt(0); // A, B, C のみ抽出
        recentPatterns.push(structureOnly);
        if (recentPatterns.length > 2) recentPatterns.shift(); // 直近2件のみ保持
      }

      sheet.getRange(rowIndex, 1).setValue(outCompany);
      sheet.getRange(rowIndex, 3).setValue(outSubject);
      sheet.getRange(rowIndex, 4).setValue(outBody);
      sheet.getRange(rowIndex, 5).setValue('done');
      sheet.getRange(rowIndex, 6).setValue(outPattern);

      processed++;
    }

    Logger.log(`処理完了：${processed} 件を更新しました。`);
  } catch (e) {
    Logger.log(e.stack);
    SpreadsheetApp.getUi().alert(`エラーが発生しました: ${e.message}`);
  }
}

/**
 * 企業ごとのプロンプトを生成
 * - ここに「STEP1〜3をAIの内部でやらせる」設計を埋め込んでいる
 * @param {string} companyName
 * @param {string} companyInfo
 * @param {string[]} recentPatterns - 直近のパターン配列（多様性確保用）
 */
function buildPromptForCompany(companyName, companyInfo, recentPatterns = []) {
  const infoText = companyInfo && companyInfo.trim()
    ? companyInfo.trim()
    : '企業URLや求人要約などの詳細情報は与えられていません。一般的な情報に基づき、過度な推測は避けてください。';

  // 多様性強制ルールの文言を生成
  let diversityRule = '';
  if (recentPatterns.length >= 2 && recentPatterns[0] === recentPatterns[1]) {
    const avoidPattern = recentPatterns[0];
    diversityRule = `
【多様性強制ルール（最優先）】
直前2件が「${avoidPattern}」構造で連続しています。
今回は必ず「${avoidPattern}」以外の構造（${avoidPattern === 'A' ? 'B または C' : avoidPattern === 'B' ? 'A または C' : 'A または B'}）を選択してください。
pattern_reason に「多様性確保のため」と明記すること。
`;
  }

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

2. 構造（A/B/C）とモード（1/2）を「企業特性に合わせて」1つ選ぶ。
${diversityRule}

  【構造（A/B/C）の選定基準】
  企業の「知名度」と「情報の複雑性」で判断します。

  ■ A（網羅型 / リクルート式）
  - 対象: 複雑なビジネスモデルの企業、または制度・環境が整っているメガベンチャー
  - 理由: 「何をしている会社か」を丁寧に説明しないと魅力が伝わらない、または福利厚生・環境などの「安心材料」が強い武器になる場合
  - 適用シグナル: 事業が多角化している、BtoB×BtoCの両面がある、複数プロダクトを持つ

  ■ B（手紙型 / ナラティブ式）
  - 対象: 創業初期（シード〜アーリー）、または大きな変革期（第二創業期）にある企業
  - 理由: 条件面や知名度では競合に劣る可能性があるため、「なぜ今やるのか」「どんな世界を作りたいか」というストーリー（物語）で共感を呼ぶ必要がある場合
  - 適用シグナル: 「0→1」「事業立ち上げ」「リーダー候補」「シリーズA〜B」「第二創業」などのキーワード

  ■ C（要点直球型 / 箇条書き式）
  - 対象: 圧倒的な知名度がある企業、またはハイレイヤー・エンジニアなど「忙しい層」がターゲットの場合
  - 理由: 長い前置きが不要、または逆効果になる層に対し、「あなたが必要な理由」だけを端的に突き刺す方が好まれる場合
  - 適用シグナル: 東証プライム上場、業界の代名詞的存在（例: 名刺管理=Sansan）、累計調達100億円以上、エンジニア専門職向け

  【知名度の判定シグナル】
  - 高: 上場企業（東証プライム/グロース/スタンダード）、累計調達100億円以上、業界内の「代名詞」的存在
  - 中: 業界内で有名、メディア露出あり、社員数500名以上
  - 低: 一般的には無名、シード〜アーリー期
  ※ 知名度が「高」で、かつターゲット層が「ハイレイヤー/エンジニア」なら C を最優先

  【構造選択の優先ルール】
  1. まず「ターゲット層」を確認 → ハイレイヤー/エンジニア向けで知名度高 → C を最優先
  2. 次に「企業フェーズ」を確認 → シード〜アーリー/変革期 → B
  3. それ以外で「ビジネスモデルが複雑 or 安心材料が強い」→ A
  4. 上記に該当しない場合 → デフォルトは C（短く刺す）

  【モード（1/2）の選定基準】
  企業の「勝ち筋」と、ターゲットが求める「キャリアの価値観」で判断します。

  ■ 1（情熱キャリアモード / 未来への期待）
  - キーワード: 社会課題解決、組織づくり、カオス、泥臭さ、ミッションドリブン
  - 理由: まだ整っていない環境で「自分たちが正解を作っていく」ことに喜びを感じる層がターゲットの場合
  - 適用: BtoCサービス、ミッションドリブンな組織、スタートアップ

  ■ 2（市場分析モード / 冷静なロジック）
  - キーワード: SaaS、プラットフォーム、勝ち馬、市場シェア、生産性、ARR
  - 理由: 「この会社に入れば市場価値が確実に上がる」という合理的な勝算を求める層がターゲットの場合
  - 適用: Horizontal SaaS、Fintech、コンサル出身者狙い

【重要：出力ルール（必須）】
- 出力は「JSON のみ」で返してください（他テキストや説明は一切禁止）。
- JSON に必ず以下のフィールドを含めること:
  - "pattern": 選択した組み合わせ（例: "B・1"）
  - "pattern_reason": 選択理由を日本語で「短い一文」（※必須）
  - pattern_reason には「知名度: 高/中/低」「ターゲット層」「フェーズ」など判断根拠を含めること。

3. 選んだパターンに従って、
  - 候補者向けのメール件名（1つ）
  - スカウト本文（1通）
   を作成する。

-------------------------------------
【構成に関する詳細ルール】

【全体共通ルール】
- 太字（**太字**）や装飾記号（###、■■など）は一切使用禁止。
- 見出しは必ず【見出し】の形式で統一すること。

- A：網羅型
  - 挨拶 → 企業紹介 → オススメポイント（箇条書き）→ クロージング → 固定フッター
  1. 挨拶・導入
    - 挨拶：初めまして。（改行）株式会社BOXの{担当者名}と申します。
    - 導入：「ご紹介する《企業名》は…」から始め、候補者が得られるメリットを提示する。
  2. 会社概要
    - 見出し：【{企業名}について】（前後に空白行を入れる）
    - 内容：事業内容やフェーズを要約。
  3. オススメポイント
    - 見出し：【オススメポイント】（前後に空白行を入れる）
    - (1) (2) のリスト形式。間に必ず空白行を入れる。
    - フォーマット：
      - (1){見出し}：{説明}
    - 情熱モード時の注意: 「すごいですよ！」ではなく「〜という稀有なチャンスです」「〜を実現できる環境です」と表現する。

- B：手紙型
  1. フック（個別化）
    - 名前は呼ばず、「ご経歴を拝見し、〜という点に魅力を感じました」から入る。
  2. 課題と再定義
    - 見出し記号は使わない。
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

-------------------------------------
【共通禁止事項（厳守）】
- 名前のプレースホルダー禁止（〇〇様、{Name} など）
- 太文字（**）禁止
- 年収などの変動要素禁止（例：「800万」）
- 具体的職種・部署名禁止（配属リスク回避）
- HR用語（ビジネス職・総合職など）禁止
- 推測での断言禁止（デカコーンに成長中！ など）
-------------------------------------
【文章表現に関する禁止事項と代替案】
企業情報に載っている場合を除いて、以下の表現は使用せず、右側の代替案や、より自然な文脈に書き換えてください。
- 貴殿　　　→個人を特定しない言い方にする
- 極めて　　→ 非常に / とても
- まさに　　→ （削除する） / ～のような / ～そのもの
- 確信した　→ 感じております / 考えております
- 最適　　　→ マッチする / 親和性が高い
- 強く惹かれ → 大変興味を持ち / 魅力を感じ
- 不可欠　　→ 重要 / カギとなる
- 稀有な　　→ 貴重な / ユニークな
- 最高　　　→ 素晴らしい / 非常に魅力的な

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
【出力フォーマット（厳守）】
- 以下のJSON形式「のみ」を返してください。
- 説明文やコメント、コードブロック記号（\`\`\`）などは一切出力しないでください。

{
  "company_name": "企業名をここに", 
  "subject": "ここにメール件名を1つ",
  "body": "ここに本文全体",
  "pattern": "例: B・1",
  "pattern_reason": "選択した理由を簡潔に1文で（知名度・ターゲット層・フェーズなどの根拠を含める）"
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
    generationConfig: {
      temperature: 1.0,
      topP: 0.95,
      topK: 40,
    }
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

  return content.parts[0].text.trim();
}

/**
 * モデルの出力テキストからJSON部分を抽出してパースする
 */
function parseResultJson(text) {
  if (!text) return null;

  let cleaned = text.trim();

  if (cleaned.startsWith('```')) {
    cleaned = cleaned.replace(/^```[\s\S]*?\n/, '');
    cleaned = cleaned.replace(/```[\s\S]*$/, '');
    cleaned = cleaned.trim();
  }

  const firstBrace = cleaned.indexOf('{');
  const lastBrace = cleaned.lastIndexOf('}');
  if (firstBrace === -1 || lastBrace === -1 || lastBrace <= firstBrace) {
    try {
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

/**
 * patternの文字列表記を正規化する（例: "a-2" → "A・2"）
 */
function normalizePattern(rawPattern) {
  if (!rawPattern) return '';

  const normalized = String(rawPattern)
    .replace(/[\s\u3000]/g, '')
    .replace(/[･・\.\/-]/g, '・')
    .toUpperCase();

  let match = normalized.match(/^([ABC])・?([12])$/);
  if (match) {
    return `${match[1]}・${match[2]}`;
  }

  match = normalized.match(/([ABC])・?([12])/);
  if (match) {
    return `${match[1]}・${match[2]}`;
  }

  return '';
}

/**
 * JSONにpatternが無い場合、レスポンステキスト全体から推測する
 */
function resolvePatternIdentifier(rawPattern, responseText, rowIndex) {
  const normalizedFromJson = normalizePattern(rawPattern);
  if (normalizedFromJson) return normalizedFromJson;

  const fallback = normalizePattern(responseText);
  if (fallback) {
    Logger.log(`Row ${rowIndex}: JSONにpatternが無かったため、テキストから「${fallback}」を補完しました。`);
    return fallback;
  }

  Logger.log(`Row ${rowIndex}: パターン識別子を抽出できませんでした。`);
  return '';
}

/**
 * 企業情報を Gemini から取得して整形する
 */
function fetchCompanyProfile(companyName) {
  if (!companyName) return '';

  const profilePrompt = `
あなたは企業リサーチャーです。以下の企業について公開情報をもとに要約してください。

【企業名】
${companyName}

【出力（JSONのみ）】
{
  "business": "事業概要（150字以内）",
  "phase": "事業フェーズ・資金調達状況など（100字以内）",
  "roles": "募集している/相性が良いと想定される職種・ロール（100字以内）",
  "skills": "求めるスキル・価値観・カルチャーフィット（120字以内）",
  "target_layer": "想定ターゲット層（例: 若手ポテンシャル / ミドルマネジメント / ハイレイヤー / エンジニア専門職）",
  "recognition": "知名度レベル（高: 上場or累計調達100億以上 / 中: 業界内で有名 / 低: 一般的には無名）",
  "proof_points": [
    "補強できる定量・エピソード（任意、1文）",
    "..."
  ]
}

【厳守】
- 公開情報のみ。推測は「〜と推測されます」と明記。
- JSON以外のテキストを出力しない。
`.trim();

  const responseText = callGemini(profilePrompt);
  if (!responseText) return '';

  const json = parseResultJson(responseText);
  if (!json) return '';

  const segments = [];
  if (json.business) segments.push(`【事業】${json.business}`);
  if (json.phase) segments.push(`【フェーズ】${json.phase}`);
  if (json.roles) segments.push(`【想定ポジション】${json.roles}`);
  if (json.skills) segments.push(`【求める人物像】${json.skills}`);
  if (json.target_layer) segments.push(`【ターゲット層】${json.target_layer}`);
  if (json.recognition) segments.push(`【知名度】${json.recognition}`);

  if (Array.isArray(json.proof_points) && json.proof_points.length) {
    const proof = json.proof_points.filter(Boolean).map((p, idx) => `・${idx + 1}. ${p}`).join('\n');
    if (proof) segments.push(`【補足エピソード】\n${proof}`);
  }

  return segments.join('\n\n').trim();
}
