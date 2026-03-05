// csvExport.gs
// 指定された日付（yyyy-MM-dd形式）のデータをCSV形式で出力する機能

function exportDailyCsv(targetDateStr) {
  // HTMLからの入力(yyyy-MM-dd)をDB検索用(yyyyMMdd)に変換
  const targetDateKey = targetDateStr.replace(/-/g, '');

  const parents = getSheetDataAsObjects_(SHEET_ORDER_PARENT);
  const children = getSheetDataAsObjects_(SHEET_ORDER_CHILD);
  
  // 1. 親データを日付でフィルタリング
  const targetParents = parents.filter(p => {
    // DB上の日付が yyyyMMdd 文字列になっている前提
    // 古いデータ(Date型)が混在している場合は変換して比較
    let pDate = String(p.注文日); 
    if (p.注文日 instanceof Date) {
      pDate = Utilities.formatDate(p.注文日, 'JST', 'yyyyMMdd');
    }
    return pDate === targetDateKey;
  });
  if (targetParents.length === 0) return null;

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
      const qty = (c.数量 === "" || c.数量 === undefined) ? "0" : c.数量;
      const unitPrice = (c.単価 === "" || c.単価 === undefined) ? "0" : c.単価;
      const amount = (c.金額 === "" || c.金額 === undefined) ? "0" : c.金額;
    
      const taxRate = (c.消費税率 === "" || c.消費税率 === undefined) ? "0" : c.消費税率;

      // 日付フォーマットの調整（DBがyyyyMMddならそのまま、Dateなら変換）
      const salesDate = (parent.注文日 instanceof Date) ? Utilities.formatDate(parent.注文日, 'JST', 'yyyyMMdd') : parent.注文日;
      const shipDate = (c.出荷日 instanceof Date) ? Utilities.formatDate(c.出荷日, 'JST', 'yyyyMMdd') : (c.出荷日 || '');
      const deliveryDate = (c.納品日 instanceof Date) ? Utilities.formatDate(c.納品日, 'JST', 'yyyyMMdd') : (c.納品日 || '');

      // ★追加：伝票番号の先頭にあるかもしれないゼロ落ち防止用の「'」を除去する
      const slipNo = String(c.伝票番号 || '').replace(/^'/, '');

    
      const row = [
        salesDate,            // 売上日
        slipNo,               // 伝票番号（9桁連番）
        parent.得意先コード,   // 得意先コード
        parent.納入先,        // 納入先コード
        shipDate,             
        // 出荷日
        deliveryDate,         // 納品日
        parent.運送会社,      // 配送会社コード
        "",                  // 伝票摘要 (空白)
        c.行番号,             // 行番号
        c.品目コード,        
        // 品目コード
        c.単位コード,         // 単位コード
        c.入数,               // 入数
        c.合数,               // 合数
        c.箱数,               // 箱数
        
        c.倉庫コード || '',   // 倉庫コード
        c.ロット番号 ||
        '',   // ロット番号
        qty,                 // 数量
        unitPrice,           // 単価
        amount,              // 金額
        taxRate,             // 消費税率
     
        c.課税区分,           // 課税区分
        `"${(c.摘要||'').replace(/"/g, '""')}"`, // 摘要
        `"${(c.備考||'').replace(/"/g, '""')}"`, // 備考
        c.規格 ||
        ''          // 規格
      ];
      
      csvRows.push(row.join(','));
    }
  });

  if (csvRows.length === 0) return null;

  // BOM付きCSV文字列を返す
  return header.join(',') + "\n" + csvRows.join('\n');
}
