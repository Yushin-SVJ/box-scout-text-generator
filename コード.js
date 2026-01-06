/***** 設定 *****/
const SHEET_NAME = 'シート1'; // シート名
const MAX_PER_RUN = 10;         // 1回の実行で処理する最大件数

// 使用するGeminiモデル
const GEMINI_MODEL = 'gemini-3-flash-preview';
const GEMINI_ENDPOINT =
  `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent`;

// 固定フッター（本文末尾にコード側で付与する）
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

      // プロンプト生成 (制御ロジックを廃止し、純粋に企業情報のみを渡す)
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

      // デバッグログ: 生レスポンスとパース結果のキーを確認
      Logger.log(`Row ${rowIndex} [Raw Response]: ${responseText}`);
      Logger.log(`Row ${rowIndex} [JSON Keys]: ${JSON.stringify(Object.keys(json))}`);
      Logger.log(`Row ${rowIndex} [Pattern Value]: ${json.pattern}`);

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
      const bodyMain = json.body_main || json.body || '';
      const outBody = bodyMain ? `${bodyMain}\n\n${FIXED_FOOTER}` : FIXED_FOOTER;
      const outPattern = resolvePatternIdentifier(json.pattern, responseText, rowIndex);

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
    // UIが使用できないコンテキスト（トリガー実行など）を考慮し、alertではなくログ出力に留める
    console.error(`エラーが発生しました: ${e.message}`);
    Logger.log(`エラーが発生しました: ${e.message}`);
  }
}

/**
 * 企業ごとのプロンプトを生成
 * - 属性ベースのルールを撤廃し、戦略ベースの思考プロセスを導入
 * @param {string} companyName
 * @param {string} companyInfo
 */
function buildPromptForCompany(companyName, companyInfo) {
  const infoText = companyInfo && companyInfo.trim()
    ? companyInfo.trim()
    : '企業URLや求人要約などの詳細情報は与えられていません。一般的な情報に基づき、過度な推測は避けてください。';

  const prompt = `
あなたは、人材紹介会社「株式会社BOX」のプロのヘッドハンターです。
与えられた企業情報に基づき、求職者の心理に最も響く「戦略的なスカウト文面」を作成してください。

【制約事項】
これまでの「知名度が高いからC」「スタートアップだからB」といった、属性に基づく機械的な判断ルールは**完全に忘れてください**。
代わりに、「この企業の魅力を目の前の候補者に届けるには、どの《心理効果》を使うべきか？」という戦略的観点のみで判断してください。

【重要：パターン選択の多様性（必読）】
A/B/Cは「企業の属性」ではなく「誰に刺すか（ターゲットペルソナの心理）」で決まります。
同じスタートアップでも、以下のように使い分けてください：
- 慎重派・比較検討派に刺したい → A（網羅・安心材料で信頼を勝ち取る）
- ワクワク・挑戦志向に刺したい → B（ビジョン・熱量で共感を生む）
- 多忙・即決派に刺したい → C（端的・インパクトで即座に価値を伝える）

※ 単に「スタートアップだからB」「有名企業だからC」という短絡的判断は禁止です。
※ 企業情報から「どのタイプの候補者が最もフィットするか」を推論し、そのペルソナに最適な戦略を選んでください。

【警告：Cパターンへの偏り禁止】
C（効率インパクト戦略）は便利ですが、**すべての企業にCを使うのは禁止**です。
- スタートアップでも「安心感を求める候補者」には A が有効
- スタートアップでも「ビジョンに共感する候補者」には B が有効
- Cは「本当に多忙なハイレイヤー」または「説明不要な超有名企業」に限定

A/B/Cの選択理由は「ペルソナの心理」に言及し、「勢いがあるからC」のような企業属性だけの理由は避けてください。

---
【タスクの全体像】
以下の4ステップを思考プロセスとして実行し、最終的な成果物（JSON）を作成してください。

### 【STEP 1: 企業と求人情報の分析】
まず、提供された企業情報を深く読み込み、以下の要素を特定してください。
1. **真の魅力（Growth Factor）**: 単なる事業内容ではなく、「なぜ今この会社に入ると面白いのか？」「他社にはない熱量は何か？」
2. **隠れた課題**: 成長に伴う組織の歪みや、まだ未完成な部分（ここが逆に「介入余地」として魅力になる）。

### 【STEP 2: ターゲット・ペルソナの定義】
この企業で最も活躍し、かつ幸せになれる人物像（ペルソナ）を具体的に定義してください。
- どんな経験を持ち、今のキャリアにどんなモヤモヤ（悩み）を抱えている人か？
- 安定志向か、挑戦志向か？
- 「年収」よりも「裁量」を求めているか、あるいは「社会的意義」を求めているか？
**※重要：ここでのペルソナはあくまで「誰に刺さる文章を書くか」の戦略用です。実際の送信相手の経歴は不明なため、メール本文内で「〇〇のご経験」などと架空の経歴を断定してはいけません。**

### 【STEP 3: アプローチ戦略の策定（最重要）】
定義したペルソナの心に最も深く刺さる「スカウトの型（A/B/C）」を、以下の**戦略的定義**に基づいて選択してください。

- **A：情報信頼戦略（Information Trust）**
  - **心理効果**: 「納得感」「安心感」「網羅性」
  - **狙い**: 複数の魅力や事実を整理して提示することで、「キャリアの選択肢として間違いない」という確信を与える。
  - **有効なケース**: ビジネスモデルが複雑で解説が必要な場合や、福利厚生・環境面も強い武器になる場合。

- **B：情緒共感戦略（Emotional Empathy）**
  - **心理効果**: 「高揚感」「同志感」「ストーリーへの没入」
  - **狙い**: 創業の背景や変革の熱量を,"あなたへの手紙"として感情的に語りかけ、理屈を超えた共感を生む。
  - **【重要】以下のいずれかに該当する場合はBを積極的に選んでください**:
    1. 社会課題解決（脱炭素、働き方改革、医療、教育など）がミッションの中心にある企業
    2. 「なぜこの会社が存在するのか」というビジョンや創業ストーリーが最大の武器である企業
    3. 条件面（給与、福利厚生）よりも「何を成し遂げるか」で候補者を惹きつけたい企業
    4. アーリー期で実績は少ないが、ビジョンへの共感で仲間を集めたい企業
  - ※ 資金調達額や成長率があっても、ミッションドリブンな企業はBが有効です。

- **C：効率インパクト戦略（Efficiency Impact）**
  - **心理効果**: 「希少性」「自信」「スピード感」
  - **狙い**: あえて情報を絞り込み、「あなたが必要な簡潔な理由」だけを突き刺すことで、企業の自信や勢いを演出する。
  - **【重要】Cを選ぶ前に以下の3条件をすべて確認してください**:
    1. ターゲットが本当に「読む時間がない」レベルで多忙か？（単に優秀＝多忙ではない）
    2. 企業名だけで「あ、聞いたことある」と言えるレベルの知名度があるか？（リクルート、マイナビ、サイバーエージェント級）
    3. 「資金調達額」「成長率」「IPO準備」だけがCを選ぶ理由になっていないか？
  - **3条件すべてを満たさない場合、AまたはBを選んでください。**
  - ※ 資金調達額や成長率は「勢い」の証拠にはなりますが、それだけではCを選ぶ理由として不十分です。

### 【STEP 4: スカウト文面の執筆】
選択した戦略に基づき、スカウト文を作成してください。
- 文体：プロフェッショナルでありながら、"一人の人間"としての体温を感じさせること。
- **禁止事項（重要）**:
  - 架空の経歴の断定（NG:「SaaSでの5年のご経験」「リーダーとしてのご活躍」など、入力されていない情報の捏造）。
  - テンプレ的な挨拶、過度なへりくだり、プレースホルダー（〇〇様など）。

---

【出力フォーマット（JSONのみ）】
思考プロセスは出力せず、以下のJSONのみを出力してください。

{
  "company_name": "企業名", 
  "subject": "選択した戦略に基づく、開封したくなる件名（1つ）",
  "body_main": "本文（挨拶から結びまで。固定フッターは含めない）",
  "pattern": "選択した型（例: A・1）",
  "reason": "【戦略的選択理由】なぜこのペルソナを設定し、なぜこの型を選んだのか？（例：『複雑な事業だが、ペルソナは多忙なため詳細な説明よりもCの「効率インパクト」で勢いを伝える戦略を選択』など）"
}

【入力された企業情報】
- 企業名: ${companyName}
- 企業情報:
${infoText}

-------------------------------------
【構成に関する詳細ルール（厳守）】

【全体共通ルール】
- 太字（**太字**）や装飾記号（###、■■など）は一切使用禁止。
- 見出しは必ず【見出し】の形式で統一すること。

- A：網羅型（情報信頼戦略）
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

- B：手紙型（情緒共感戦略）
  1. フック（個別化）
    - 名前は呼ばず、「ご経歴を拝見し、〜という点に魅力を感じました」から入る。
    - **注意**: 具体的な社名や職種には触れず、「これまでのキャリア」や「これからの挑戦」といった抽象度で留めること。
  2. 課題と再定義
    - 見出し記号は使わない。
    - 「今、{企業名}は〜という面白いフェーズにあります」と文章で繋ぐ。
  3. 接続と提案
    - 名前は呼ばず、条件接続を行う。
    - OK例: 「だからこそ、{定義したスキル}をお持ちの方の力が不可欠なのです」
    - 箇条書きは使わず、段落分け（空白行）で読みやすくする。

- C：要点直球型（効率インパクト戦略）
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
- 具体的な経歴の捏造禁止（「〇〇でのご経験」「××プロジェクトでの実績」など、入力データにない事象の断定）
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
本文にはフッターを含めないでください（コード側で固定フッターを後付けします）。本文だけを返してください。
`.trim();

  return prompt;
}

/**
 * Gemini APIを呼び出す
 * @param {string} prompt
 * @param {string} [modelName] - 使用するモデル名（省略時は GEMINI_MODEL 定数を使用）
 * @returns {string|null} モデルの生テキストレスポンス
 */
function callGemini(prompt, modelName) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    throw new Error('GEMINI_API_KEY がスクリプトプロパティに設定されていません。');
  }

  // モデル名が指定されていればそれを、なければ定数(GEMINI_MODEL)を使用
  const useModel = modelName || GEMINI_MODEL;
  const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/${useModel}:generateContent`;

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

  // リトライロジック (最大3回 / 指数バックオフ)
  const maxRetries = 3;
  let attempt = 0;

  while (attempt < maxRetries) {
    try {
      const res = UrlFetchApp.fetch(endpoint, options);
      const code = res.getResponseCode();
      const text = res.getContentText();

      if (code === 200) {
        // 成功
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

      // 一時的なエラーの場合はリトライ (429: Too Many Requests, 503: Service Unavailable)
      if (code === 429 || code === 503) {
        attempt++;
        if (attempt < maxRetries) {
          const waitTime = Math.pow(2, attempt) * 1000; // 2s, 4s, 8s...
          Logger.log(`Gemini API Error (${code}). Retrying in ${waitTime}ms... (Model: ${useModel})`);
          Utilities.sleep(waitTime);
          continue;
        }
      }

      // リトライ対象外、または回数切れ
      Logger.log(`Gemini API error (Model: ${useModel}): ` + code + ' ' + text);
      return null;

    } catch (e) {
      Logger.log(`UrlFetchAuth Error: ${e.message}`);
      return null;
    }
  }

  return null;
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

  // フルフォーマット: A・1, B・2 など
  let match = normalized.match(/^([ABC])・?([12])$/);
  if (match) {
    return `${match[1]}・${match[2]}`;
  }

  // 部分マッチ: テキスト内に A・1 が含まれる場合
  match = normalized.match(/([ABC])・?([12])/);
  if (match) {
    return `${match[1]}・${match[2]}`;
  }

  // モード番号なしのパターン (A, B, C のみ) を許容
  match = normalized.match(/^([ABC])$/);
  if (match) {
    return match[1];  // モード番号なしで返す
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
