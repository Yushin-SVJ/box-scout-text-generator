/**
 * 複数社スカウト生成機能 (MultiScout.js)
 * 既存の `コード.js` とは独立して動作するが、`callGemini` などの共通関数は利用する。
 * 
 * Update: 推論精度の高い Gemini Pro モデルを使用し、思考プロセス（analysis）を経てから文面生成を行う「Chain of Thought」構成に変更
 */

// 定数定義
const MULTI_SHEET_NAME = '複数社検討シート';
const SOURCE_SHEET_NAME = 'シート1';
const MULTI_MAX_PER_RUN = 5; // 1回の実行で処理する最大件数

// ユーザー指定の「推論用高精度モデル」
// ※ 一般的に Pro = "gemini-1.5-pro" です。ユーザー指定の "2.5" があればそれに書き換えてください。
const GEMINI_HIGH_QUALITY_MODEL = 'gemini-2.5-pro';

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
これより2つの企業情報を入力します。その内容を深く分析し、比較・統合した上で、候補者に「刺さる」スカウトメールを作成してください。

【プロセス指示】
モデルとして以下の思考プロセスを経て、最終的なJSONを出力してください。
1. **分析（Analysis）**: 各企業の特徴（DNA）や「なぜ伸びているか（Growth Factor）」を言語化し、2社の関係性を定義する。
2. **スコアリング（Scoring）**: 後述する3つの構造パターン（A/B/C）に対し、今回の2社との適合度を0〜10点で採点する。
   - A（対比）: 違いが明確か？
   - B（共鳴）: 共通のテーマは深いか？
   - C（資産）: キャリア資産性は高いか？
3. **戦略決定（Strategy）**: スコアが「7点以上」のパターンの中から、最も効果的な1つを選択する。
   - ※重要: 特定のパターンに偏らないよう、有効な選択肢が複数ある場合は柔軟に選ぶこと。
4. **生成（Generation）**: 決定した戦略に従って文面を作成する。

------------------------------------------------
【ターゲット像の動的定義】
ターゲットは固定ではありません。今回の「2社の組み合わせ」から最も惹きつけられる人物像を逆算して定義してください。
例：Startup × Startup なら「カオスを楽しめる人」、Startup × Mega なら「環境を選びたい人」、HR Tech × HR Tech なら「業界を変えたい人」。

------------------------------------------------
【構造（A/B/C）の定義】

■ **A：環境選択型 / The Choice（対比）**
- **適合度基準**: カルチャーやフェーズ、成長メカニズムの違いが「ハッキリしている」ほど高得点。
- **ロジック**: 「あなたはどちらの環境で輝きますか？」という問いかけ。
- **構成**: 挨拶 → 各社の違いを際立たせた紹介 → 選択の提案 → クロージング

■ **B：テーマ・ミッション型 / The Theme（共鳴）**
- **適合度基準**: 業界トレンド、社会課題、ミッションの共通性が「深い」ほど高得点。
- **ロジック**: 「この巨大な波（トレンド）に乗りませんか？」という招待。
- **構成**: フック（業界の波）→ 共通課題 → その先駆者としての2社紹介 → クロージング

■ **C：キャリア資産・ROI型 / The Asset（実利・未来）**
- **適合度基準**: 「出身者ブランド（Alumni）」や「希少スキル（Deep Skill）」の獲得価値が「高い」ほど高得点。
- **ロジック**: 「この会社での経験は、あなたのキャリアにおける"資産"になります」という投資対効果。
- **構成**: 単刀直入な導入 → 得られる3つの資産/メリット → クロージング
- **注意**: 以前の「ポジション提案」ではなく、「得られるスキル/経験/実績」に焦点を当てること。

------------------------------------------------
【モード（1/2）の定義：訴求の「トーン」で選ぶ】

■ **1：Willモード（情熱・想い）**
- **キーワード**: ミッション、カオス、組織づくり、人、熱狂。
- **トーン**: エモーショナル、熱い、主観的。
- **対象**: 「ワクワクしたい」「誰と働くか」を重視する人。

■ **2：Logicモード（合理的・勝算）**
- **キーワード**: 勝ち筋、市場シェア、プロダクト優位性、キャリアハック。
- **トーン**: クール、戦略的、客観的。
- **対象**: 「負け戦はしたくない」「確実にスキルをつけたい」と考える合理的な人。

------------------------------------------------
【禁止事項・制約】
- 「SaaS」「IT業界」など、入力にない情報を勝手に補完しない。
- 以下のワードは禁止（代替表現を使うこと）：
  - 貴殿 → 
  - 極めて → 非常に
  - まさに → （削除 or 言い換え）
  - 確信した → 感じております
  - 最適 → マッチする
  - 不可欠 → 重要
  - 稀有な → 貴重な / ユニークな
  - 最高 → 素晴らしい

------------------------------------------------
【入力データ】
■1社目：{companyA_Name}
{companyA_Body}

■2社目：{companyB_Name}
{companyB_Body}

------------------------------------------------
【出力フォーマット（JSON Only）】
必ず以下のJSON形式のみを出力してください。Markdown記法や前置きは不要です。

{
  "analysis": {
    "company_A_dna": "1社目の特徴（50文字以内）",
    "company_B_dna": "2社目の特徴（50文字以内）",
    "relationship": "2社の関係性（例: 成長フェーズの対比構造 / ◯◯業界の共闘構造）",
    "scores": {
        "A_Contrast": 8,
        "B_Theme": 6,
        "C_Asset": 9
    },
    "selected_pattern_reason": "スコアに基づき、なぜそのパターンを選んだかの理由"
  },
  "subject": "件名（パターンに合わせて作成）",
  "body_main": "本文（挨拶からクロージングまで。フッターは含めない）",
  "pattern": "A・1 などの記号",
  "pattern_reason": "選択理由（短く）"
}
`.trim();


/**
 * ヘッダー名から列番号（1-based）を取得するヘルパー関数
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

    let srcIdxName = getColumnIndex(sourceSheet, '企業名');
    let srcIdxBody = getColumnIndex(sourceSheet, '本文');
    let srcIdxPattern = getColumnIndex(sourceSheet, 'パターン');

    // フォールバック
    if (srcIdxName === -1) srcIdxName = 1; // A
    if (srcIdxBody === -1) srcIdxBody = 4; // D
    if (srcIdxPattern === -1) srcIdxPattern = 6; // F

    const sourceLastRow = sourceSheet.getLastRow();
    if (sourceLastRow < 2) {
        Logger.log('参照元のデータがありません。');
        return;
    }
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
 * JSONフォーマット (analysisブロック付き) に対応
 */
function parseGeminiResponse(text) {
    if (!text) return null;
    let cleaned = text.trim();

    // コードブロック除去
    if (cleaned.startsWith('```')) {
        cleaned = cleaned.replace(/^```[a-z]*\n/i, '').replace(/\n```$/, '');
    }

    // JSON部分抽出
    const firstBrace = cleaned.indexOf('{');
    const lastBrace = cleaned.lastIndexOf('}');

    if (firstBrace !== -1 && lastBrace > firstBrace) {
        const sub = cleaned.substring(firstBrace, lastBrace + 1);
        try {
            const json = JSON.parse(sub);
            // 正規化
            if (json.body && !json.body_main) {
                json.body_main = json.body;
            }
            return json;
        } catch (e) {
            Logger.log('JSON Parse Error: ' + e.message);
        }
    }

    Logger.log('パース失敗: JSON構造が見つかりません。');
    return null;
}


/**
 * 2. generateMultiScoutMails()
 * 複合スカウト文面を生成する（本番実行用）。
 * 高精度モデル (Pro) を呼び出す。
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
        // if (processedCount >= MULTI_MAX_PER_RUN) break; 
        // ユーザー要望のProモデル使用に伴い処理時間を考慮して件数制限は維持、または減らすことも検討

        const row = values[i];
        const rowIndex = i + 2;

        const status = String(row[colStatus - 1]);
        if (status === 'done') continue; // 完了済みはスキップ

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

        Logger.log(`Row ${rowIndex}: Gemini (${GEMINI_HIGH_QUALITY_MODEL}) 生成開始...`);

        // ★ここで高精度モデルを指定して呼び出す
        const responseText = callGemini(prompt, GEMINI_HIGH_QUALITY_MODEL);

        if (!responseText) {
            Logger.log(`Row ${rowIndex}: Geminiレスポンスなし`);
            sheet.getRange(rowIndex, colErr).setValue('Gemini API Error');
            continue;
        }

        // パース
        const json = parseGeminiResponse(responseText);
        if (!json) {
            Logger.log(`Row ${rowIndex}: パース失敗`);
            sheet.getRange(rowIndex, colErr).setValue('Parse Error');
            continue;
        }

        // 推論結果(analysis)があればログに出すか、reasonに入れる
        let reason = json.pattern_reason || '';
        if (json.analysis && json.analysis.strategy_reason) {
            reason = `【戦略】${json.analysis.strategy_reason} (Pattern: ${reason})`;
        }

        sheet.getRange(rowIndex, colOutSubject).setValue(json.subject || '');
        sheet.getRange(rowIndex, colOutBody).setValue(json.body_main || json.body || '');
        sheet.getRange(rowIndex, colOutPattern).setValue(json.pattern || '');
        sheet.getRange(rowIndex, colOutReason).setValue(reason);
        sheet.getRange(rowIndex, colStatus).setValue('done');
        sheet.getRange(rowIndex, colErr).clearContent();

        processedCount++;
        Logger.log(`Row ${rowIndex}: 生成完了 (Pattern: ${json.pattern})`);

        // ProモデルはRate Limitに引っかかりやすいので少しWaitを入れる
        Utilities.sleep(1000);
    }

    if (processedCount > 0) {
        Logger.log(`完了: ${processedCount} 件`);
    } else {
        Logger.log('処理対象がありませんでした。');
    }
}
