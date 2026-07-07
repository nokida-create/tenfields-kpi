// =====================================================
//  営業KPIダッシュボード - サーバーサイド
//  設定: SPREADSHEET_ID を変更してください
// =====================================================

const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE'; // ← ここを変更
const EXCLUDE_COLUMNS = ['平均', '合計', '全体平均', '平均値'];

// =====================================================
//  Web App エントリーポイント
// =====================================================
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('営業KPIダッシュボード')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// =====================================================
//  シート一覧を返す
// =====================================================
function getSheetList() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheets = ss.getSheets().map(s => s.getName());
    return { ok: true, sheets };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// =====================================================
//  指定シートのKPIデータを返す
//  想定フォーマット:
//    A1: (空) / B1〜: メンバー名 / 最終列: 平均
//    A2〜: KPI項目名 / B2〜: 数値データ
// =====================================================
function getSheetData(sheetName) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { ok: false, error: `シート "${sheetName}" が見つかりません` };

    const values = sheet.getDataRange().getValues();
    if (values.length < 2) return { ok: false, error: 'データが空です' };

    // ── メンバー名を1行目から取得（除外列を除く）──
    const memberColumns = []; // [{ name, col }]
    const headerRow = values[0];
    for (let i = 1; i < headerRow.length; i++) {
      const name = String(headerRow[i]).trim();
      if (name && !EXCLUDE_COLUMNS.includes(name)) {
        memberColumns.push({ name, col: i });
      }
    }

    // ── KPI項目名をA列から取得（2行目以降）──
    const kpiRows = []; // [{ name, row }]
    for (let i = 1; i < values.length; i++) {
      const name = String(values[i][0]).trim();
      if (name) kpiRows.push({ name, row: i });
    }

    // ── メンバーごとのKPIデータを構造化 ──
    const memberData = {};
    for (const member of memberColumns) {
      memberData[member.name] = {};
      for (const kpi of kpiRows) {
        const raw = values[kpi.row][member.col];
        memberData[member.name][kpi.name] =
          typeof raw === 'number' ? raw : (parseFloat(String(raw).replace(/,/g, '')) || 0);
      }
    }

    return {
      ok: true,
      members: memberColumns.map(m => m.name),
      kpiItems: kpiRows.map(k => k.name),
      memberData,
      sheetName
    };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}
