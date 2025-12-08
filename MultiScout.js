
/**
 * 複数社スカウト生成機能 (MultiScout.js)
 * 既存の `コード.js` とは独立して動作するが、`callGemini` などの共通関数は利用する。
 * 
 * Update: ヘッダー位置変更に対応するため、列指定を動的に変更
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


// プロンプト定義
const MULTI_SCOUT_PROMPT = `
あなたは、人材紹介会社「株式会社BOX」のスカウト文面作成パートナーです。
今回は、2〜3社の企業をまとめてご紹介する「複数社スカウトメール」を作成します。

【入力として与えられるもの】
ユーザーから、2社分についてそれぞれ以下が与えられます。
- 企業名
- その企業向けに過去に作成した「単社スカウト本文」

単社スカウト本文は、
「その企業の特徴・フェーズ・ウリを候補者向けに説明した文章」
として扱い、そこから情報を抽出して構成してください。
（※Web検索や新たな推測は行わず、与えられた本文情報をベースにすること）

【内部プロセス】
1. それぞれの単社スカウト本文を読み、各社の「求めている人材像」「最大のウリ」「会社概要」を把握する。
2. 2社を横並びで見比べ、共通するキャリア軸や、逆に際立つ個性を特定する。
3. メール全体として採用する「構造（A/B/C）」と「モード（1/2）」を1つ選ぶ。

【構造（A/B/C）の選定基準】
A：網羅型（比較整理が必要な場合）
- 企業ごとの “方向性の違い” を整理して見せる。「A社は組織作り、B社は事業開発」のように訴求を分けたい場合。
B：手紙型（ストーリーで惹きつける場合）
- 2社の共通項（ミッション、市場の波）を一つの物語として語れる場合。
C：要点直球型（短く刺す方が有効な場合）
- 共通の魅力がシンプルで、忙しいターゲットに要点だけ届けたい場合。

【モード（1/2）の選定基準】
1（情熱キャリアモード）：社会課題解決、組織づくり、カオス、ミッションドリブン
2（市場分析モード）：SaaS、勝ち馬、市場シェア、合理的な勝算

【出力フォーマット（JSONのみ）】
以下のJSON形式のみを返してください。Markdown記号は不要です。
{
  "subject": "候補者向けのメール件名（1つ）",
  "body": "スカウト本文全体（挨拶〜固定フッターまで）",
  "pattern": "選択した組み合わせ（例: A・1）",
  "pattern_reason": "選択理由を日本語で1文"
}

【固定フッター】
本文の最後には必ず以下のフッターを入れてください：
▼ 私が提供できる価値
（※ここには既存の固定フッターが入ります。コード内で補完してください）
...
（省略：既存のコード.jsにあるFIXED_FOOTERと同じものを使用）
...

【対象企業情報】
■1社目：{companyA_Name}
\${companyA_Body}

■2社目：{companyB_Name}
\${companyB_Body}
`;


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
        const name = String(row[srcIdxName - 1]).trim(); // 0-based
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

    // 必要な列インデックスを取得
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

    // 全列読み込み
    const targetMaxCol = targetSheet.getLastColumn();
    const targetRange = targetSheet.getRange(2, 1, targetLastRow - 1, targetMaxCol);
    const targetValues = targetRange.getValues();

    // 書き込み内容を保持する配列 (行番号 -> { col: val, ... }) の代わりに
    // 効率化のため getValuesで取った配列を直接書き換えて setValues するか、
    // update用の配列を用意して setValues する。
    // ここでは更新対象のカラムが飛び飛びになる可能性があるため、API呼び出し回数を減らす工夫が必要。
    // しかし GAS setValues は矩形範囲が必要。
    // 「参照列」は H〜K のように連続していると想定されるが、動的カラム対応なので連続とは限らない。
    // 安全のため、行ごとに更新データを準備して、最後にまとめて...は難しい（列が散らばる）。
    // したがって、書き込みは「参照データブロック」が連続していると期待しつつ、
    // ここでは実装の単純さと堅牢性を優先し、
    // setValues用配列を「全行 × 全列」用意して、必要な箇所だけ上書きし、最後にドカンと書き込むのがベストだが、
    // シートの他のデータ（生成済み本文など）を上書きして消してしまうリスクがある。
    // → 結論: update用の配列を作成し、列ごとにまとめて書き込む。

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

    // 列ごとに一括書き込み
    if (targetValues.length > 0) {
        targetSheet.getRange(2, colRefBody1, targetValues.length, 1).setValues(updatesRefBody1);
        targetSheet.getRange(2, colRefBody2, targetValues.length, 1).setValues(updatesRefBody2);
        targetSheet.getRange(2, colRefPat1, targetValues.length, 1).setValues(updatesRefPat1);
        targetSheet.getRange(2, colRefPat2, targetValues.length, 1).setValues(updatesRefPat2);

        Logger.log(`参照データ転記完了: ${targetValues.length} 件`);
    }
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

    // 列位置特定
    const colName1 = getRequiredColumnIndex(sheet, '企業名1');
    const colName2 = getRequiredColumnIndex(sheet, '企業名2');
    const colStatus = getRequiredColumnIndex(sheet, 'ステータス');
    const colErr = getRequiredColumnIndex(sheet, '[エラー] エラー理由');

    // 参照データ列
    const colRefBody1 = getRequiredColumnIndex(sheet, '[参照] 企業名1本文');
    const colRefBody2 = getRequiredColumnIndex(sheet, '[参照] 企業名2本文');

    // 出力先列
    const colOutSubject = getRequiredColumnIndex(sheet, '【生成】複合版件名');
    const colOutBody = getRequiredColumnIndex(sheet, '【生成】複合版本文');
    const colOutPattern = getRequiredColumnIndex(sheet, '【生成】複合版パターン');
    const colOutReason = getRequiredColumnIndex(sheet, '【生成】選定理由');

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    // 全データ取得
    const maxCol = sheet.getLastColumn();
    const values = sheet.getRange(2, 1, lastRow - 1, maxCol).getValues();

    let processedCount = 0;

    for (let i = 0; i < values.length; i++) {
        if (processedCount >= MULTI_MAX_PER_RUN) break;

        const row = values[i];
        const rowIndex = i + 2;

        const status = String(row[colStatus - 1]); // 0-based
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
            .replace('${companyA_Body}', body1)
            .replace('${companyB_Body}', body2)
            .replace('${FIXED_FOOTER}', FIXED_FOOTER_MULTI);

        Logger.log(`Row ${rowIndex}: Gemini 生成開始...`);
        const responseText = callGemini(prompt);

        if (!responseText) {
            Logger.log(`Row ${rowIndex}: Geminiレスポンスなし`);
            sheet.getRange(rowIndex, colErr).setValue('Gemini API Error');
            continue;
        }

        const json = parseResultJson(responseText);
        if (!json) {
            Logger.log(`Row ${rowIndex}: JSONパース失敗`);
            sheet.getRange(rowIndex, colErr).setValue('JSON Parse Error');
            continue;
        }

        // 結果書き込み (セル単位で書き込む)
        sheet.getRange(rowIndex, colOutSubject).setValue(json.subject || '');
        sheet.getRange(rowIndex, colOutBody).setValue(json.body || '');
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
