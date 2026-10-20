/**
 * OTC在庫台帳 — 店舗間備蓄共有サーバー v3
 * v3: 品目が多い薬局でも保存できるよう、品目データを複数のセルに分けて記録します
 *     (Googleスプレッドシートは 1つのセルに 50,000文字までしか入らないため)。
 *     v2で保存したデータもそのまま読めます。
 * script.google.com に貼り付けて「ウェブアプリ」として公開するだけで動きます。
 * 初回実行時に「OTC備蓄共有」という名前のスプレッドシートが
 * あなたのGoogleドライブに自動作成され、共有データはそこに記録されます。
 */

const SHEET_NAME = "備蓄共有";

function getSheet_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty("SSID");
  let ss = null;
  if (id) {
    try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; }
  }
  if (!ss) {
    ss = SpreadsheetApp.create("OTC備蓄共有");
    props.setProperty("SSID", ss.getId());
  }
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(["id", "薬局名", "地区", "更新日時", "品目データ(JSON・続きは右の列)"]);
  }
  return sh;
}

/** 一覧取得: GET {URL}?action=list */
function doGet(e) {
  // Yahoo!ショッピングAPIの中継(ブラウザから直接呼べないため、ここで代理取得する)
  if (e && e.parameter && e.parameter.action === "yahoo") {
    return yahooLookup_(e.parameter.appid, e.parameter.jan);
  }
  const sh = getSheet_();
  const rows = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0]) continue;
    let items = [];
    // 5列目から右に 分けて入れた JSON を つなげて読む
    let json = "";
    for (let c = 4; c < r.length; c++) json += String(r[c] || "");
    try { items = JSON.parse(json || "[]"); } catch (err) {}
    out.push({
      id: String(r[0]),
      pharmacy: { name: String(r[1] || ""), area: String(r[2] || "") },
      updatedAt: String(r[3] || ""),
      items: items,
    });
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

/** 登録・更新・削除: POST {URL} に JSON を送信 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const sh = getSheet_();
    const rows = sh.getDataRange().getValues();
    let rowIndex = -1;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(d.id)) { rowIndex = i + 1; break; }
    }

    if (d.action === "upload") {
      const json = JSON.stringify(d.items || []);
      const CHUNK = 45000;               // 1セルの上限(50,000文字)より少なめに区切る
      const parts = [];
      for (let p = 0; p < json.length; p += CHUNK) parts.push(json.slice(p, p + CHUNK));
      if (!parts.length) parts.push("[]");
      const values = [
        String(d.id),
        String((d.pharmacy && d.pharmacy.name) || ""),
        String((d.pharmacy && d.pharmacy.area) || ""),
        String(d.updatedAt || new Date().toISOString()),
      ].concat(parts);
      const row = rowIndex > 0 ? rowIndex : sh.getLastRow() + 1;
      // 前回の 続きの列が 残らないよう 先に 消してから 書く
      const lastCol = Math.max(sh.getLastColumn(), values.length);
      sh.getRange(row, 1, 1, lastCol).clearContent();
      sh.getRange(row, 1, 1, values.length).setNumberFormat("@").setValues([values]);
    } else if (d.action === "delete") {
      if (rowIndex > 0) sh.deleteRow(rowIndex);
    }

    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}


/** Yahoo!ショッピング商品検索APIの中継 */
function yahooLookup_(appid, jan) {
  var empty = ContentService.createTextOutput(JSON.stringify({ hits: [] }))
    .setMimeType(ContentService.MimeType.JSON);
  if (!appid || !jan) return empty;
  try {
    var url = "https://shopping.yahooapis.jp/ShoppingWebService/V3/itemSearch"
      + "?appid=" + encodeURIComponent(appid)
      + "&jan_code=" + encodeURIComponent(jan)
      + "&results=10";
    var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return empty;
    return ContentService.createTextOutput(res.getContentText())
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return empty;
  }
}


