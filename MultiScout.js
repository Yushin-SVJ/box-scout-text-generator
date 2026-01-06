
/**
 * 複数社スカウト生成機能 (MultiScout.js)
 * 既存の `コード.js` とは独立して動作するが、`callGemini` などの共通関数は利用する。
 * 
 * Update: ヘッダー位置変更に対応するため、列指定を動的に変更
 * Update: プロンプトの大幅改修とGems形式（テキスト形式）出力への対応
 */

// 定数定義
const MULTI_SHEET_NAME = '複数社検討シート';
const SOURCE_SHEET_NAME = 'シート1';
const MULTI_MAX_PER_RUN = 5; // 1回の実行で処理する最大件数

// コード.js にある FIXED_FOOTER を複製
const FIXED_FOOTER_MULTI = `
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


// 新プロンプト定義（ユーザー指定の内容 + データ注入部）
const MULTI_SCOUT_PROMPT = `
あなたは、人材紹介会社「株式会社BOX」のスカウト文面作成パートナーです。
今回は、2〜3社の企業をまとめてご紹介する「複数社スカウトメール」を作成します。

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

------------------------------------------------
【ターゲット像（固定前提）】
- 想定読者は「営業経験者」です。
- 成長やキャリアアップ、将来のマネジメント層へのステップに関心がある人を想定してください。
- ただし、転職後のポジションは営業に限らず多岐にわたる前提とし、「SaaS営業」「IT営業」など、特定の業種・職種に限定する表現は避けてください。
------------------------------------------------
【業種・職種に関する禁止事項】
- 「SaaS」「IT企業」「Web業界」「広告業界」など、入力に書かれていない業種ラベルを推測で付与することは禁止です。
- 入力文に明記されていない業界名・職種名を勝手に補うことは禁止です。
- SaaS特有の用語（ARR、MRR など）や、特定業界の常識を前提にした表現は、入力文に明記されている場合を除き使用しないでください。
- 業界ラベルではなく、「どのような事業フェーズか」「どのようなスキルが伸びる環境か」など、キャリア軸・フェーズ軸で表現してください。
------------------------------------------------
【文章表現に関する禁止事項と代替案】
企業情報に載っている場合を除いて、以下の表現は使用せず、右側の代替案や、より自然な文脈に書き換えてください。
- 極めて　　→ 非常に / とても
- まさに　　→ （削除する） / ～のような / ～そのもの
- 確信した　→ 感じております / 考えております
- 最適　　　→ マッチする / 親和性が高い
- 強く惹かれ → 大変興味を持ち / 魅力を感じ
- 不可欠　　→ 重要 / カギとなる
- 稀有な　　→ 貴重な / ユニークな
- 最高　　　→ 素晴らしい / 非常に魅力的な
- 具体的職種・部署名禁止（配属リスク回避）
- 推測での断言禁止（デカコーンに成長中！ など）
------------------------------------------------
【入力として与えられるもの】
ユーザーから、2社分についてそれぞれ以下が与えられます。
- 企業名
- その企業向けに過去に作成した「単社スカウト本文」
単社スカウト本文は、「その企業の特徴・フェーズ・ウリを候補者向けに説明した文章」として扱ってください。
------------------------------------------------
【内部プロセス（あなたの頭の中で行い、出力はしない）】
1. それぞれの単社スカウト本文を読み、各社について以下を頭の中で整理してください（推測は控えめに、本文に書かれている内容をベースとする）：
   - その企業が求めていそうな人材像（ペルソナ）
   - その企業の「最大のウリ（Growth Factor）」
   - 会社概要（2〜3文程度の要約）
2. 2〜3社を横並びで見比べ、共通するキャリア軸・魅力や、逆に際立つ個性を特定してください。
   例：
   - 「営業経験を土台に、事業づくりに関われる」
   - 「成長フェーズの事業で裁量を持てる」
   - 「将来のマネジメント候補として育っていける」

3. メール全体として採用する「戦略（A/B/C）」と「モード（1/2）」を1つ選んでください。

【アプローチ戦略の策定（A/B/C）】（※属性ではなく心理効果で選ぶ）
- **A：情報信頼戦略（Information Trust）**
  - **心理効果**: 「納得感」「安心感」「網羅性」
  - **狙い**: 複数の企業の魅力や事実を整理して提示することで、「キャリアの選択肢として間違いない」という確信を与える。
  - **有効なケース**: 複数社の共通点や相乗効果をロジカルに説明したい場合。

- **B：情緒共感戦略（Emotional Empathy）**
  - **心理効果**: 「高揚感」「同志感」「ストーリーへの没入」
  - **狙い**: 複数社に共通する熱量やビジョンを、"あなたへの手紙"として感情的に語りかけ、理屈を超えた共感を生む。
  - **【重要】以下のいずれかに該当する場合はBを積極的に選んでください**:
    1. 社会課題解決（脱炭素、働き方改革、医療、教育など）がミッションの中心にある企業
    2. 「なぜこの会社が存在するのか」というビジョンや創業ストーリーが最大の武器である企業
    3. 条件面（給与、福利厚生）よりも「何を成し遂げるか」で候補者を惹きつけたい企業
    4. アーリー期で実績は少ないが、ビジョンへの共感で仲間を集めたい企業
  - ※ 資金調達額や成長率があっても、ミッションドリブンな企業はBが有効です。

- **C：効率インパクト戦略（Efficiency Impact）**
  - **心理効果**: 「希少性」「自信」「スピード感」
  - **狙い**: あえて情報を絞り込み、「あなたが必要な簡潔な理由」だけを突き刺すことで、勢いを演出する。
  - **【重要】Cを選ぶ前に以下の3条件をすべて確認してください**:
    1. ターゲットが本当に「読む時間がない」レベルで多忙か？（単に優秀＝多忙ではない）
    2. 企業名だけで「あ、聞いたことある」と言えるレベルの知名度があるか？（リクルート、マイナビ、サイバーエージェント級）
    3. 「資金調達額」「成長率」「IPO準備」だけがCを選ぶ理由になっていないか？
  - **3条件すべてを満たさない場合、AまたはBを選んでください。**

【モード（1/2）の選定基準】
- **1：情熱キャリアモード（未来への期待）**
  - キーワード：社会課題解決、組織づくり、カオス、ミッションドリブン
  - 適用例：スタートアップ、ミッション重視のベンチャー、成長段階の事業
- **2：市場分析モード（合理的ロジック）**
  - キーワード：プラットフォーム、成長率、マーケットシェア、プロダクト価値
  - 適用例：成熟したSaaS、Fintech、コンサル出身者向けの実績重視案件

【重要：出力ルール（必須）】
- 出力は「JSON のみ」で返してください（他テキストや説明は一切禁止）。
- JSON に必ず以下のフィールドを含めること:
    - "subject": 件名（1つ / 戦略に基づいたもの）
    - "body_main": フッターを含めない本文（挨拶〜クロージングまで）
  - "pattern": 選択した組み合わせ（例: "B・1"）
  - "pattern_reason": 選択理由を日本語で「短い一文」（例: "多忙な層と想定し、C（効率インパクト）で勢いを重視"）

4. 選んだパターンに従って、
   - 候補者向けのメール件名（1つ）
     - スカウト本文（1通、フッターを含めない）
   を作成する。
-------------------------------------
【構成に関する詳細ルール】
【全体共通ルール】
- 太字（**太字**）や装飾記号（###、■■など）は一切使用禁止。
- 見出しは必ず【見出し】の形式で統一すること。
- A：網羅型
  - 挨拶 → 企業紹介 → オススメポイント → クロージング → 固定フッター
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
  - フック → 課題提示 → 提案 → クロージング → 固定フッター
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
  - 挨拶 → 要点リスト → クロージング → 固定フッター
    1.単刀直入な導入
        - 挨拶の後、すぐに本題に入る。
        - OK例: 「今回ご連絡したのは、以下の3つの理由から、これまでのご実績と{企業名}の相性が最高だと確信したためです。」

    2.3つの理由リスト
        - 見出し記号は使わない。
        - "・"（全角ナカグロ）を使った3点のリストにする。間に空白行を入れる。
        - 内容: スキルマッチ、企業の成長性、キャリアメリットを端的に述べる。

    3.結び
        - 「詳細な資料もございますので…」と簡潔に締める。
【件名の作成ルール】
- 構造の異なる以下の3パターンを比較し、今回の企業に最適と判断した“1パターンだけ”を採用し件名を1つ作成する。
  - ミッション/進化型（企業の挑戦テーマ重視）
  - タグ/ブランド型（【】で並べる形式）
  - インパクト/機会型（フェーズの希少性）

【固定フッター】
本文には固定フッターを含めず、本文のみを生成してください（フッターはコード側で後付けします）。

【対象企業情報】
■1社目：{companyA_Name}
{companyA_Body}

■2社目：{companyB_Name}
{companyB_Body}
`.trim();


/**
 * ヘッダー名から列番号（1-based）を取得するヘルパー関数
 * @param {Sheet} sheet
 * @param {string} headerName
 * @returns {number} 1-based index (見つからない場合は -1)
 */
function getColumnIndex(sheet, headerName) {
    const lastCol = sheet.getLastColumn();
    if (lastCol < 1) return -1;
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const index = headers.indexOf(headerName);
    return index === -1 ? -1 : index + 1;
}

/**
 * 列インデックスが見つからない場合にエラーを投げるラップ関数
 */
function getRequiredColumnIndex(sheet, headerName) {
    const idx = getColumnIndex(sheet, headerName);
    if (idx === -1) {
        throw new Error(`シート「${sheet.getName()}」にヘッダー「${headerName}」が見つかりません。`);
    }
    return idx;
}


/**
 * 1. transferSingleScoutData()
 * 「シート1」から最新の単社データを取得し、「複数社検討シート」の参照列を更新する。
 */
function transferSingleScoutData() {
    const ss = SpreadsheetApp.getActive();

    // ------------------------------
    // 1. ソースデータの読み込み
    // ------------------------------
    const sourceSheet = ss.getSheetByName(SOURCE_SHEET_NAME);
    if (!sourceSheet) {
        Logger.log(`シート「${SOURCE_SHEET_NAME}」が見つかりません。`);
        return;
    }
    // ソース側の列特定 (企業名, 本文, パターン)
    // ※もしシート1のヘッダー名が固定でないならここも修正が必要だが、
    //   一般的にコード.jsと合わせるため、既存カラム名「企業名」「本文」「パターン」を想定
    //   見つからない場合は従来の固定列(A=1, D=4, F=6)にフォールバックする安全策をとる
    let srcIdxName = getColumnIndex(sourceSheet, '企業名');
    let srcIdxBody = getColumnIndex(sourceSheet, '本文'); // または 'スカウト本文'
    let srcIdxPattern = getColumnIndex(sourceSheet, 'パターン');

    // フォールバック
    if (srcIdxName === -1) srcIdxName = 1; // A
    if (srcIdxBody === -1) srcIdxBody = 4; // D
    if (srcIdxPattern === -1) srcIdxPattern = 6; // F

    Logger.log(`Source Columns - Name:${srcIdxName}, Body:${srcIdxBody}, Pattern:${srcIdxPattern}`);

    const sourceLastRow = sourceSheet.getLastRow();
    if (sourceLastRow < 2) {
        Logger.log('参照元のデータがありません。');
        return;
    }
    // 全データ取得 (行ごとに処理)
    const sourceC = sourceSheet.getLastColumn();
    const sourceValues = sourceSheet.getRange(2, 1, sourceLastRow - 1, sourceC).getValues();

    // Map化 (Key: 企業名 -> Value: { body, pattern })
    const companyMap = new Map();
    for (const row of sourceValues) {
        const name = String(row[srcIdxName - 1]).trim();
        if (!name) continue;

        companyMap.set(name, {
            body: row[srcIdxBody - 1],
            pattern: row[srcIdxPattern - 1]
        });
    }

    // ------------------------------
    // 2. ターゲットシートの読み込み
    // ------------------------------
    const targetSheet = ss.getSheetByName(MULTI_SHEET_NAME);
    if (!targetSheet) {
        Logger.log(`シート「${MULTI_SHEET_NAME}」が見つかりません。`);
        return;
    }

    const colName1 = getRequiredColumnIndex(targetSheet, '企業名1');
    const colName2 = getRequiredColumnIndex(targetSheet, '企業名2');
    const colRefBody1 = getRequiredColumnIndex(targetSheet, '[参照] 企業名1本文');
    const colRefBody2 = getRequiredColumnIndex(targetSheet, '[参照] 企業名2本文');
    const colRefPat1 = getRequiredColumnIndex(targetSheet, '[参照] 企業名1パターン');
    const colRefPat2 = getRequiredColumnIndex(targetSheet, '[参照] 企業名2パターン');

    const targetLastRow = targetSheet.getLastRow();
    if (targetLastRow < 2) {
        Logger.log('処理対象のデータがありません。');
        return;
    }

    const targetMaxCol = targetSheet.getLastColumn();
    const targetRange = targetSheet.getRange(2, 1, targetLastRow - 1, targetMaxCol);
    const targetValues = targetRange.getValues();

    const updatesRefBody1 = [];
    const updatesRefBody2 = [];
    const updatesRefPat1 = [];
    const updatesRefPat2 = [];

    for (let i = 0; i < targetValues.length; i++) {
        const row = targetValues[i];
        const name1 = String(row[colName1 - 1]).trim();
        const name2 = String(row[colName2 - 1]).trim();

        const info1 = companyMap.get(name1) || { body: '', pattern: '' };
        const info2 = companyMap.get(name2) || { body: '', pattern: '' };

        updatesRefBody1.push([info1.body]);
        updatesRefBody2.push([info2.body]);
        updatesRefPat1.push([info1.pattern]);
        updatesRefPat2.push([info2.pattern]);
    }

    if (targetValues.length > 0) {
        targetSheet.getRange(2, colRefBody1, targetValues.length, 1).setValues(updatesRefBody1);
        targetSheet.getRange(2, colRefBody2, targetValues.length, 1).setValues(updatesRefBody2);
        targetSheet.getRange(2, colRefPat1, targetValues.length, 1).setValues(updatesRefPat1);
        targetSheet.getRange(2, colRefPat2, targetValues.length, 1).setValues(updatesRefPat2);

        Logger.log(`参照データ転記完了: ${targetValues.length} 件`);
    }
}

/**
 * Geminiのレスポンスをパースする関数
 * JSONフォーマット または "Gems用簡易形式" (Key-Value テキスト) の両方に対応
 * @param {string} text
 * @returns {object|null} { subject, body_main, pattern, pattern_reason }
 */
function parseGeminiResponse(text) {
    if (!text) return null;
    let cleaned = text.trim();

    // 1. JSONパースを試みる
    let jsonResult = null;
    // コードブロック除去
    if (cleaned.startsWith('```')) {
        const blockRemoved = cleaned.replace(/^```[a-z]*\n/i, '').replace(/\n```$/, '');
        try {
            jsonResult = JSON.parse(blockRemoved);
        } catch (e) { }
    } else {
        // 最初の { から 最後の } までを切り出し
        const firstBrace = cleaned.indexOf('{');
        const lastBrace = cleaned.lastIndexOf('}');
        if (firstBrace !== -1 && lastBrace > firstBrace) {
            const sub = cleaned.substring(firstBrace, lastBrace + 1);
            try {
                jsonResult = JSON.parse(sub);
            } catch (e) { }
        }
    }

    if (jsonResult) {
        // 正規化：body_main がなければ body を流用
        if (jsonResult.body && !jsonResult.body_main) {
            jsonResult.body_main = jsonResult.body;
            delete jsonResult.body;
        }
        return jsonResult;
    }

    // 2. Gems簡易形式 (テキスト) のパース
    // フォーマット:
    // 件名：
    // ...
    // 本文：
    // ...
    // pattern：
    // ...
    // pattern_reason：
    // ...

    const result = {};

    // 正規表現で各セクションを抽出
    // (?=...) 先読みを使って次のセクション見出しの手前までを取得する
    const subjectMatch = cleaned.match(/件名[：:]\s*([\s\S]*?)(?=\n\s*(本文|pattern|パターン)[：:])/i);
    const bodyMatch = cleaned.match(/本文[：:]\s*([\s\S]*?)(?=\n\s*(pattern|パターン)[：:])/i);

    // パターンと理由は後ろにあることが多いので、残りを柔軟に取る
    const patternMatch = cleaned.match(/pattern[：:]\s*([A-C]・?[12])?/i);
    // pattern_reason は最後に来ることが多い
    const reasonMatch = cleaned.match(/pattern_reason[：:]\s*([\s\S]*)$/i);

    if (subjectMatch) result.subject = subjectMatch[1].trim();
    if (bodyMatch) result.body_main = bodyMatch[1].trim();
    if (patternMatch && patternMatch[1]) result.pattern = patternMatch[1].trim();
    if (reasonMatch) result.pattern_reason = reasonMatch[1].trim();

    // 必須項目が取れていれば採用
    if (result.subject && result.body_main) {
        return result;
    }

    Logger.log('パース失敗: JSONでもGems形式でも解釈できませんでした。');
    return null;
}


/**
 * 2. generateMultiScoutMails()
 * 複合スカウト文面を生成する（本番実行用）。
 */
function generateMultiScoutMails() {
    const ss = SpreadsheetApp.getActive();

    // 転記実行
    transferSingleScoutData();

    const sheet = ss.getSheetByName(MULTI_SHEET_NAME);
    if (!sheet) return;

    const colName1 = getRequiredColumnIndex(sheet, '企業名1');
    const colName2 = getRequiredColumnIndex(sheet, '企業名2');
    const colStatus = getRequiredColumnIndex(sheet, 'ステータス');
    const colErr = getRequiredColumnIndex(sheet, '[エラー] エラー理由');

    const colRefBody1 = getRequiredColumnIndex(sheet, '[参照] 企業名1本文');
    const colRefBody2 = getRequiredColumnIndex(sheet, '[参照] 企業名2本文');

    const colOutSubject = getRequiredColumnIndex(sheet, '【生成】複合版件名');
    const colOutBody = getRequiredColumnIndex(sheet, '【生成】複合版本文');
    const colOutPattern = getRequiredColumnIndex(sheet, '【生成】複合版パターン');
    const colOutReason = getRequiredColumnIndex(sheet, '【生成】選定理由');

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    const maxCol = sheet.getLastColumn();
    const values = sheet.getRange(2, 1, lastRow - 1, maxCol).getValues();

    let processedCount = 0;

    for (let i = 0; i < values.length; i++) {
        if (processedCount >= MULTI_MAX_PER_RUN) break;

        const row = values[i];
        const rowIndex = i + 2;

        const status = String(row[colStatus - 1]);
        if (status === 'done') continue;

        const company1 = row[colName1 - 1];
        const company2 = row[colName2 - 1];

        if (!company1 || !company2) continue;

        const body1 = row[colRefBody1 - 1];
        const body2 = row[colRefBody2 - 1];

        if (!body1 || !body2) {
            const msg = `単社情報不足のためスキップ: ${company1}, ${company2}`;
            Logger.log(`Row ${rowIndex}: ${msg}`);
            sheet.getRange(rowIndex, colErr).setValue(msg);
            continue;
        }

        // プロンプト作成
        let prompt = MULTI_SCOUT_PROMPT
            .replace('{companyA_Name}', company1)
            .replace('{companyB_Name}', company2)
            .replace('{companyA_Body}', body1)
            .replace('{companyB_Body}', body2);

        Logger.log(`Row ${rowIndex}: Gemini 生成開始...`);
        // コード.jsの callGemini は変更せずそのまま使う
        const responseText = callGemini(prompt);

        if (!responseText) {
            Logger.log(`Row ${rowIndex}: Geminiレスポンスなし`);
            sheet.getRange(rowIndex, colErr).setValue('Gemini API Error');
            continue;
        }

        // 新しいパース関数を利用
        const json = parseGeminiResponse(responseText);
        if (!json) {
            Logger.log(`Row ${rowIndex}: パース失敗`);
            sheet.getRange(rowIndex, colErr).setValue('Parse Error');
            continue;
        }

        const bodyMain = json.body_main || json.body || '';
        const finalBody = bodyMain ? `${bodyMain}\n\n${FIXED_FOOTER_MULTI}` : FIXED_FOOTER_MULTI;

        sheet.getRange(rowIndex, colOutSubject).setValue(json.subject || '');
        sheet.getRange(rowIndex, colOutBody).setValue(finalBody);
        sheet.getRange(rowIndex, colOutPattern).setValue(json.pattern || '');
        sheet.getRange(rowIndex, colOutReason).setValue(json.pattern_reason || '');
        sheet.getRange(rowIndex, colStatus).setValue('done');
        sheet.getRange(rowIndex, colErr).clearContent();

        processedCount++;
        Logger.log(`Row ${rowIndex}: 生成完了`);
    }

    if (processedCount > 0) {
        Logger.log(`完了: ${processedCount} 件`);
    } else {
        Logger.log('処理対象がありませんでした。');
    }
}
