// csvExport.gs
// 指定された日付（yyyy-MM-dd形式）のデータをCSV形式で出力する機能

// すべての値をダブルクォートで囲み、先頭ゼロ落ち・カンマ・改行を防ぐ
function q_(val) {
  const s = (val === null || val === undefined) ? '' : String(val);
  return '"' + s.replace(/"/g, '""') + '"';
}

function markCsvExported_(orderIds) {
  if (!orderIds || orderIds.length === 0) return;
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID_ORDERS);
  const sheet = ss.getSheetByName(SHEET_ORDER_PARENT);
  const range = sheet.getDataRange();
  const data = range.getValues();
  const headers = data[0];

  let csvFlagIdx = headers.indexOf('CSV出力フラグ');
  if (csvFlagIdx === -1) {
    const newCol = headers.length + 1;
    sheet.getRange(1, newCol).setValue('CSV出力フラグ');
    csvFlagIdx = newCol - 1;
  }

  const idIdx = headers.indexOf('注文ID');
  if (idIdx === -1) return;
  const targetSet = new Set(orderIds.map(String));
  const now = Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd HH:mm:ss');

  for (let i = 1; i < data.length; i++) {
    if (targetSet.has(String(data[i][idIdx]))) {
      sheet.getRange(i + 1, csvFlagIdx + 1).setValue(now);
    }
  }
}

function exportDailyCsv(targetDateStr) {
  // HTMLからの入力(yyyy-MM-dd)をDB検索用(yyyyMMdd)に変換
  const targetDateKey = targetDateStr.replace(/-/g, '');

  const parents = getSheetDataAsObjects_(SHEET_ORDER_PARENT, SPREADSHEET_ID_ORDERS);
  const children = getSheetDataAsObjects_(SHEET_ORDER_CHILD, SPREADSHEET_ID_ORDERS);

  // 1. 親データを日付でフィルタリング
  const allDateParents = parents.filter(p => {
    let pDate = String(p.注文日);
    if (p.注文日 instanceof Date) {
      pDate = Utilities.formatDate(p.注文日, 'JST', 'yyyyMMdd');
    }
    return pDate === targetDateKey;
  });
  if (allDateParents.length === 0) return null;

  // CSV出力済みを除外（二重出力防止）
  const targetParents = allDateParents.filter(p => !p['CSV出力フラグ']);
  if (targetParents.length === 0) return 'ALREADY_EXPORTED';

  // 親IDをキーにしたマップ作成
  const parentMap = new Map(targetParents.map(p => [String(p.注文ID), p]));

  // 2. CSVヘッダー作成
  const header = [
    "売上日", "伝票番号", "得意先コード", "納入先コード", "出荷日", "納品日", "配送会社コード", "伝票摘要",
    "行番号", "品目コード", "単位コード", "入数", "合数", "箱数", "倉庫コード", "ロット番号",
    "数量", "単価", "金額", "消費税率", "課税区分", "摘要", "備考", "規格"
  ];

  // 3. データ行の作成
  const csvRows = [];

  children.forEach(c => {
    const parent = parentMap.get(String(c.注文ID));

    // 親が存在する場合のみ出力
    if (parent) {
      // 数値項目の空欄対応
      const qty      = (c.数量     === "" || c.数量     === undefined) ? "0" : c.数量;
      const unitPrice = (c.単価    === "" || c.単価     === undefined) ? "0" : c.単価;
      const amount   = (c.金額     === "" || c.金額     === undefined) ? "0" : c.金額;
      const taxRate  = (c.消費税率 === "" || c.消費税率 === undefined) ? "0" : c.消費税率;

      // 日付フォーマットの調整（DBがyyyyMMddならそのまま、Dateなら変換）
      const salesDate    = (parent.注文日 instanceof Date) ? Utilities.formatDate(parent.注文日, 'JST', 'yyyyMMdd') : parent.注文日;
      const shipDate     = (c.出荷日     instanceof Date) ? Utilities.formatDate(c.出荷日,     'JST', 'yyyyMMdd') : (c.出荷日     || '');
      const deliveryDate = (c.納品日     instanceof Date) ? Utilities.formatDate(c.納品日,     'JST', 'yyyyMMdd') : (c.納品日     || '');

      // 伝票番号の先頭にあるテキスト強制用「'」を除去してから出力
      const slipNo = String(c.伝票番号 || '').replace(/^'/, '');

      // 全フィールドを q_() でテキスト化（先頭ゼロ落ち防止）
      const row = [
        q_(salesDate),
        q_(slipNo),
        q_(parent.得意先コード),
        q_(parent.納入先),
        q_(shipDate),
        q_(deliveryDate),
        q_(parent.運送会社),
        q_(""),
        q_(c.行番号),
        q_(c.品目コード),
        q_(c.単位コード),
        q_(c.入数),
        q_(c.合数),
        q_(c.箱数),
        q_(c.倉庫コード || ''),
        q_(c.ロット番号 || ''),
        q_(qty),
        q_(unitPrice),
        q_(amount),
        q_(taxRate),
        q_(c.課税区分),
        q_(c.摘要  || ''),
        q_(c.備考  || ''),
        q_(c.規格  || '')
      ];

      csvRows.push(row.join(','));
    }
  });

  if (csvRows.length === 0) return null;

  // CSV出力済みフラグを書き込む
  markCsvExported_(targetParents.map(p => p.注文ID));

  // BOM付きCSV文字列を返す（ヘッダーも全てテキスト化）
  return header.map(q_).join(',') + "\n" + csvRows.join('\n');
}
