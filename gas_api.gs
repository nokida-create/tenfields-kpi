// ============================================================
//  営業部 KPI ダッシュボード - GAS APIサーバー
//  doGet() でJSON返却 → HTMLからfetchして使う
// ============================================================

// 月別スプレッドシートID
const SHEET_MAP = [
  { month: "2025/7",  id: "1QtyPCx9SMnaVziF27iivfM6xbvHxphaTtoFXhBGttMU" },
  { month: "2025/8",  id: "1pDWvhD6A_kSF7ityfsBxPmE0B4_73vxHdtlbVIxhKBI" },
  { month: "2025/9",  id: "1GEjPfngTbhDwgzF8-PHqlwj9zkV2ePOvKhOeCCDncxA" },
  { month: "2025/10", id: "1-CxUfcnDJPGln3wqABmwANSoU27-s5ra0Apph9jgPZg" },
  { month: "2025/11", id: "1FbEbt9SXjZ4S9d-bxz3LsGKsT2q1FXBaTdCBk0vS7SY" },
  { month: "2025/12", id: "1b4VwYcmciVp0zbAQUfEgdtXjG2qVE-RkWDzKCSSvqwQ" },
  { month: "2026/1",  id: "1KeV6YHmmY3AsfPNohLAeXyCA1nBnmYQhPmnA9rdL4eY" },
  { month: "2026/2",  id: "1iK-y2RoPzAfF4idZ5B11_6WcfVjCw7kBqBwhPHP6UZQ" },
  { month: "2026/3",  id: "1mUjwX0S-D82FhxJpEQi-l7R3299l4WwtzyOldMfOIh8" },
  { month: "2026/4",  id: "1vvZc4ibSm-Fna0--HeKSu1N8TgxA5Um7pr00FylBLu8" },
  { month: "2026/5",  id: "1z9_dsgX9c51gTzfr3cOxfmrcm8mx2NcApc-_PXFxCwY" },
  { month: "2026/6",  id: "1nX_tVBkr6OACnDeChX7Bp-VJGV0foWL9ssZMKme7mlA" },
];

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
  const kpiData = { tel: {}, mtg: {}, action: {} };
  let teamTotal = { tel: null, mtg: null, action: null };

  for (let i = headerRowIdx + 1; i < data.length; i++) {
    const row     = data[i];
    const rowLabel = String(row[0]).trim();

    for (const [key, keyword] of Object.entries(KPI_KEYS)) {
      if (rowLabel.includes(keyword)) {
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
