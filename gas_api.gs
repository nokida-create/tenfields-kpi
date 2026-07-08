// ============================================================
//  営業部 KPI ダッシュボード - GAS APIサーバー
//  doGet() でJSON返却 → HTMLからfetchして使う
//
//  【自動スキャン方式】
//  KPI_FOLDER_ID のフォルダを毎回スキャンし、
//  "YYYY/M営業KPI管理" という名前のスプレッドシートを自動検出。
//  翌月シートがフォルダに追加されるだけで自動的に反映される。
// ============================================================

// KPI月次スプレッドシートが入っているフォルダID
const KPI_FOLDER_ID = '1lzAuqoQfAlcUPbTfxVrIatwimcONQpcH';

// ============================================================
//  フォルダをスキャンしてSHEET_MAPを動的生成
// ============================================================
function buildSheetMap() {
  const folder = DriveApp.getFolderById(KPI_FOLDER_ID);
  const files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  const map = [];

  while (files.hasNext()) {
    const file = files.next();
    const name = file.getName().trim();
    // "2025/7営業KPI管理" や "2026/10営業KPI管理" にマッチ
    const m = name.match(/^(\d{4})\/(\d{1,2})営業KPI管理/);
    if (m) {
      const month = `${m[1]}/${m[2]}`;
      map.push({ month, id: file.getId() });
    }
  }

  // 年月順にソート
  map.sort((a, b) => {
    const [ay, am] = a.month.split('/').map(Number);
    const [by, bm] = b.month.split('/').map(Number);
    return ay !== by ? ay - by : am - bm;
  });

  return map;
}

// KPIサマリーシート名（各スプレッドシートの1枚目）
const SUMMARY_SHEET_INDEX = 0;

// 読み取るKPI行のキーワード（部分一致）
const KPI_KEYS = {
  tel:    "電話",
  mtg:    "総商談数",
  action: "総アクション",
};

// ============================================================
//  メインAPI
// ============================================================
function doGet(e) {
  const result = getAllKpi();
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
//  全月データ取得
// ============================================================
function getAllKpi() {
  const months = [];
  const SHEET_MAP = buildSheetMap();

  for (const entry of SHEET_MAP) {
    try {
      const monthData = getMonthKpi(entry.month, entry.id);
      months.push(monthData);
    } catch (err) {
      months.push({ month: entry.month, error: err.toString(), members: {}, team_total: {} });
    }
  }

  return {
    updated_at: new Date().toISOString(),
    months: months,
  };
}

// ============================================================
//  1ヶ月分のKPI取得
// ============================================================
function getMonthKpi(month, ssId) {
  const ss = SpreadsheetApp.openById(ssId);
  const sheet = ss.getSheets()[SUMMARY_SHEET_INDEX];
  const data  = sheet.getDataRange().getValues();

  // ヘッダー行を探す（メンバー名が並ぶ行）
  let headerRowIdx = -1;
  let headerRow    = [];

  for (let i = 0; i < Math.min(data.length, 30); i++) {
    const row = data[i];
    // 空でない文字列セルが5個以上 → ヘッダー候補
    const strCells = row.filter(c => typeof c === "string" && c.trim() !== "" && c.trim() !== ":-:");
    if (strCells.length >= 5) {
      // 「電話」や数字が含まれていない → メンバー名行
      const hasKpi = strCells.some(c => c.includes("電話") || c.includes("商談") || c.includes("アクション"));
      if (!hasKpi) {
        headerRowIdx = i;
        headerRow    = row;
        break;
      }
    }
  }

  if (headerRowIdx === -1) {
    return { month, members: {}, team_total: {}, error: "header_not_found" };
  }

  // メンバー名リスト（空・「平均」を除外）
  const members = [];
  const memberIdx = {}; // name -> colIndex

  for (let c = 0; c < headerRow.length; c++) {
    const name = String(headerRow[c]).trim();
    if (name !== "" && name !== "平均" && name !== ":-:") {
      members.push(name);
      memberIdx[name] = c;
    }
  }

  // KPI行を探してパース
  // 注意: シート下部に列配置の異なる重複集計表があるため、
  //  - 各KPIは最初にマッチした行のみ採用（先勝ち）
  //  - ヘッダー直後の最初の表が終わる（全セル空行）時点で走査終了
  //  - "電話" は「電話に対してのアポ率」を除外するため完全一致のみ
  const kpiData = { tel: {}, mtg: {}, action: {} };
  let teamTotal = { tel: null, mtg: null, action: null };
  const captured = new Set();

  for (let i = headerRowIdx + 1; i < data.length; i++) {
    const row     = data[i];
    const rowLabel = String(row[0]).trim();

    if (captured.size === Object.keys(KPI_KEYS).length) break;
    if (row.every(c => String(c).trim() === "")) break;

    for (const [key, keyword] of Object.entries(KPI_KEYS)) {
      if (rowLabel.includes(keyword)) {
        if (captured.has(key)) break;
        if (key === "tel" && rowLabel !== "電話") break;
        captured.add(key);
        // 個人別
        for (const [name, ci] of Object.entries(memberIdx)) {
          const val = toNum(row[ci]);
          kpiData[key][name] = val;
        }
        // チーム合計（最後の数値セル）
        for (let c = row.length - 1; c >= 0; c--) {
          const v = toNum(row[c]);
          if (v !== null && v > 0) {
            teamTotal[key] = v;
            break;
          }
        }
        break;
      }
    }
  }

  // メンバー別にまとめる
  const membersResult = {};
  for (const name of members) {
    const tel    = kpiData.tel[name]    ?? null;
    const mtg    = kpiData.mtg[name]    ?? null;
    const action = kpiData.action[name] ?? null;
    if (tel !== null || mtg !== null || action !== null) {
      membersResult[name] = { tel, mtg, action };
    }
  }

  return {
    month,
    members:    membersResult,
    team_total: teamTotal,
  };
}

// ============================================================
//  数値変換
// ============================================================
function toNum(val) {
  if (val === null || val === undefined || val === "") return null;
  if (typeof val === "number") return isNaN(val) ? null : Math.round(val);
  const s = String(val).replace(/,/g, "").trim();
  if (s === "" || s.startsWith("#")) return null;
  const n = parseFloat(s);
  return isNaN(n) ? null : Math.round(n);
}

// ============================================================
//  デバッグ用（スクリプトエディタから実行して確認）
// ============================================================
function testRun() {
  const result = getAllKpi();
  Logger.log(JSON.stringify(result, null, 2));
}
