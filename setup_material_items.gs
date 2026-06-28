function setupMaterialItems() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── material_items シート作成 ──
  let s = ss.getSheetByName('material_items') || ss.insertSheet('material_items');
  s.clear();
  s.getRange(1, 1, 1, 7).setValues([[
    'id', 'material_id', 'order', 'category', 'name', 'review_auto', 'memo'
  ]]);
  s.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#E0F7FA');

  // ITパスポート 10日間完成 のサンプルデータ
  const items = [
    ['item_001', 'mat_001', 1, 'Day1', '企業活動', 'ON', ''],
    ['item_002', 'mat_001', 2, 'Day1', '法務', 'ON', ''],
    ['item_003', 'mat_001', 3, 'Day1', '経営戦略マネジメント', 'ON', ''],
    ['item_004', 'mat_001', 4, 'Day2', 'システム戦略', 'ON', ''],
    ['item_005', 'mat_001', 5, 'Day2', '開発技術', 'ON', ''],
    ['item_006', 'mat_001', 6, 'Day3', 'プロジェクトマネジメント', 'ON', ''],
    ['item_007', 'mat_001', 7, 'Day3', 'サービスマネジメント', 'ON', ''],
    ['item_008', 'mat_001', 8, 'Day4', '基礎理論', 'ON', ''],
    ['item_009', 'mat_001', 9, 'Day4', 'アルゴリズムとプログラミング', 'ON', ''],
    ['item_010', 'mat_001', 10, 'Day5', 'コンピュータ構成要素', 'ON', ''],
    ['item_011', 'mat_001', 11, 'Day5', 'システム構成要素', 'ON', ''],
    ['item_012', 'mat_001', 12, 'Day6', 'ソフトウェア', 'ON', ''],
    ['item_013', 'mat_001', 13, 'Day6', 'ハードウェア', 'ON', ''],
    ['item_014', 'mat_001', 14, 'Day7', 'ヒューマンインタフェース', 'ON', ''],
    ['item_015', 'mat_001', 15, 'Day7', 'マルチメディア', 'ON', ''],
    ['item_016', 'mat_001', 16, 'Day7', 'データベース', 'ON', ''],
    ['item_017', 'mat_001', 17, 'Day8', 'ネットワーク', 'ON', ''],
    ['item_018', 'mat_001', 18, 'Day8', 'セキュリティ', 'ON', ''],
    ['item_019', 'mat_001', 19, 'Day9', '企業と法務（実践）', 'ON', ''],
    ['item_020', 'mat_001', 20, 'Day10', '総合演習', 'ON', ''],
  ];
  s.getRange(2, 1, items.length, 7).setValues(items);
  s.setFrozenRows(1);
  s.autoResizeColumns(1, 7);

  // materials シートに review_auto 列を追加（なければ）
  const ms = ss.getSheetByName('materials');
  const headers = ms.getRange(1, 1, 1, ms.getLastColumn()).getValues()[0];
  if (!headers.includes('review_auto')) {
    const nextCol = ms.getLastColumn() + 1;
    ms.getRange(1, nextCol).setValue('review_auto').setFontWeight('bold');
    // 既存行にデフォルト ON を設定
    const lastRow = ms.getLastRow();
    if (lastRow > 1) {
      const vals = Array.from({length: lastRow - 1}, () => ['ON']);
      ms.getRange(2, nextCol, lastRow - 1, 1).setValues(vals);
    }
  }

  SpreadsheetApp.getUi().alert('material_items シート作成完了！\n\nITパスポート 20項目を登録しました。\nmaterials に review_auto 列を追加しました。');
}
