
/**
 * 複数社スカウト生成機能 (MultiScout.js)
 * 既存の `コード.js` とは独立して動作するが、`callGemini` などの共通関数は利用する。
 */

// 定数定義
const MULTI_SHEET_NAME = '複数社検討シート';
const SOURCE_SHEET_NAME = 'シート1';
const MULTI_MAX_PER_RUN = 5; // 1回の実行で処理する最大件数

// コード.js にある FIXED_FOOTER を複製（既存ファイルを変更しないため）
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
 * 1. transferSingleScoutData()
 * 「シート1」から最新の単社データを取得し、「複数社検討シート」の参照列（H〜K列）を更新する。
 */
function transferSingleScoutData() {
    const ss = SpreadsheetApp.getActive();

    // 1. ソースデータの読み込み
    const sourceSheet = ss.getSheetByName(SOURCE_SHEET_NAME);
    if (!sourceSheet) {
        Logger.log(`シート「${SOURCE_SHEET_NAME}」が見つかりません。`);
        return;
    }
    const sourceLastRow = sourceSheet.getLastRow();
    if (sourceLastRow < 2) {
        Logger.log('参照元のデータがありません。');
        return;
    }

    // A列:企業名, B列:URL(未使用), C列:件名(未使用), D列:本文, E列:ステータス(未使用), F列:パターン
    // getRange(row, col, numRows, numCols) -> A2:F(lastRow)
    const sourceValues = sourceSheet.getRange(2, 1, sourceLastRow - 1, 6).getValues();

    // Map化 (Key: 企業名 -> Value: { body, pattern })
    const companyMap = new Map();
    for (const row of sourceValues) {
        const name = String(row[0]).trim();
        if (!name) continue;

        // 最新の行を優先するか、最初の行を優先するか。ここでは「上書き」＝下の行（より新しい行と仮定）が残るようにする
        // 必要であれば逆順ループにするなどの調整が可能
        companyMap.set(name, {
            body: row[3],      // D列
            pattern: row[5]    // F列
        });
    }

    // 2. ターゲットシートの読み込み
    const targetSheet = ss.getSheetByName(MULTI_SHEET_NAME);
    if (!targetSheet) {
        Logger.log(`シート「${MULTI_SHEET_NAME}」が見つかりません。作成してください。`);
        return;
    }

    const targetLastRow = targetSheet.getLastRow();
    if (targetLastRow < 2) {
        Logger.log('処理対象のデータがありません。');
        return;
    }

    // A列(企業名1), B列(企業名2) を取得
    const targetRange = targetSheet.getRange(2, 1, targetLastRow - 1, 2);
    const targetValues = targetRange.getValues();

    // 書き込み用配列の準備 (H, I, J, K 列用) -> 4カラム
    const updates = [];

    for (let i = 0; i < targetValues.length; i++) {
        const name1 = String(targetValues[i][0]).trim();
        const name2 = String(targetValues[i][1]).trim();

        const info1 = companyMap.get(name1) || { body: '', pattern: '' };
        const info2 = companyMap.get(name2) || { body: '', pattern: '' };

        updates.push([
            info1.body,    // H: 企業名1本文
            info2.body,    // I: 企業名2本文
            info1.pattern, // J: 企業名1パターン
            info2.pattern  // K: 企業名2パターン
        ]);
    }

    // 3. 一括書き込み (H2:K(lastRow))
    if (updates.length > 0) {
        targetSheet.getRange(2, 8, updates.length, 4).setValues(updates);
        Logger.log(`参照データ転記完了: ${updates.length} 件`);
    }
}


/**
 * 2. generateMultiScoutMails()
 * 複合スカウト文面を生成する（本番実行用）。
 */
function generateMultiScoutMails() {
    const ss = SpreadsheetApp.getActive();

    // まず参照データを最新化
    transferSingleScoutData();

    const sheet = ss.getSheetByName(MULTI_SHEET_NAME);
    if (!sheet) return; // transferSingleScoutDataでログ出力済み

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    // 読み込み範囲: A列〜K列 (11カラム)
    // A:企業1, B:企業2, C:件名, D:本文, E:パターン, F:理由, G:ステータス, H:本1, I:本2, J:パ1, K:パ2
    const dataRange = sheet.getRange(2, 1, lastRow - 1, 11);
    const values = dataRange.getValues();

    let processedCount = 0;

    for (let i = 0; i < values.length; i++) {
        if (processedCount >= MULTI_MAX_PER_RUN) break;

        const row = values[i];
        const rowIndex = i + 2;

        const company1 = row[0];
        const company2 = row[1];
        const status = row[6];
        const body1 = row[7]; // H列
        const body2 = row[8]; // I列

        // 処理条件: ステータスが 'done' でない、かつ 企業名が両方ある
        if (status === 'done') continue;
        if (!company1 || !company2) continue;

        // 参照データチェック
        if (!body1 || !body2) {
            const msg = `単社情報不足のためスキップ: ${company1}, ${company2}`;
            Logger.log(`Row ${rowIndex}: ${msg}`);
            sheet.getRange(rowIndex, 12).setValue(msg); // L列にエラー出力
            continue;
        }

        // プロンプト作成
        let prompt = MULTI_SCOUT_PROMPT
            .replace('{companyA_Name}', company1)
            .replace('{companyB_Name}', company2)
            .replace('${companyA_Body}', body1)
            .replace('${companyB_Body}', body2)
            .replace('${FIXED_FOOTER}', FIXED_FOOTER_MULTI); // フッター埋め込み

        // 既存コードの callGemini を呼び出し
        // ※コード.js が同じプロジェクトにある前提
        Logger.log(`Row ${rowIndex}: Gemini 生成開始...`);
        const responseText = callGemini(prompt);

        if (!responseText) {
            Logger.log(`Row ${rowIndex}: Geminiレスポンスなし`);
            sheet.getRange(rowIndex, 12).setValue('Gemini API Error');
            continue;
        }

        // JSONパース (既存コードの parseResultJson を利用)
        const json = parseResultJson(responseText);
        if (!json) {
            Logger.log(`Row ${rowIndex}: JSONパース失敗`);
            sheet.getRange(rowIndex, 12).setValue('JSON Parse Error');
            continue;
        }

        // 書き込み
        // C:件名, D:本文, E:パターン, F:選定理由, G:ステータス
        sheet.getRange(rowIndex, 3).setValue(json.subject || '');
        sheet.getRange(rowIndex, 4).setValue(json.body || '');
        sheet.getRange(rowIndex, 5).setValue(json.pattern || '');
        sheet.getRange(rowIndex, 6).setValue(json.pattern_reason || '');
        sheet.getRange(rowIndex, 7).setValue('done');
        sheet.getRange(rowIndex, 12).clearContent(); // エラー列クリア

        processedCount++;
        Logger.log(`Row ${rowIndex}: 生成完了`);
    }

    if (processedCount === 0) {
        Logger.log('処理対象がありませんでした（すべて完了済み or データ不足）。');
    } else {
        Logger.log(`完了: ${processedCount} 件`);
    }
}
