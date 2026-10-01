/**
 * SCTV工事費判定ツール 設定保存用 Google Apps Script
 *
 * スプレッドシートに「工事費単価」「機器」「追加工事項目」「更新履歴」の各シートを作り、
 * 工事費判定ツールの管理者ページから保存された設定を読み書きする。
 * 管理者パスワードはコードに書かず、メニュー「工事費ツール ＞ 管理者パスワードを設定」から
 * 登録した値の SHA-256 ハッシュをスクリプトプロパティに保存して照合する。
 *
 * セットアップ手順は同じフォルダの README.md を参照。
 */

const SHEET_NAMES = {
  rates: "工事費単価",
  devices: "機器",
  extras: "追加工事項目",
  log: "更新履歴"
};

const GENRES = { stb: "STB関連", remote: "リモコン", net: "ONU・その他", camera: "防犯カメラ" };

const PASSWORD_PROPERTY = "ADMIN_PASSWORD_SHA256";
const MAX_FAILURES = 10;          // この回数パスワードを間違えると
const LOCK_SECONDS = 10 * 60;     // この秒数、保存を受け付けない
const MAX_ROWS = 500;
const MAX_TEXT = 80;

// ===== スプレッドシートのメニュー =====

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("工事費ツール")
    .addItem("管理者パスワードを設定", "setAdminPassword")
    .addToUi();
}

function setAdminPassword() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt("管理者パスワードを設定",
    "工事費判定ツールの管理者ページと同じパスワードを入力してください。",
    ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const pw = res.getResponseText().trim();
  if (!pw) { ui.alert("パスワードが空のため設定しませんでした。"); return; }
  PropertiesService.getScriptProperties().setProperty(PASSWORD_PROPERTY, sha256Hex_(pw));
  ui.alert("管理者パスワードを設定しました。");
}

// ===== Web API =====

// 設定の読み込み（パスワード不要）。
function doGet() {
  return json_(Object.assign({ ok: true }, readConfig_()));
}

// 設定の保存（管理者パスワードが必要）。
// ブラウザからは CORS のプリフライトを避けるため Content-Type: text/plain で JSON を送る。
function doPost(e) {
  let req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: "リクエストの形式が正しくありません。" });
  }

  const stored = PropertiesService.getScriptProperties().getProperty(PASSWORD_PROPERTY);
  if (!stored) {
    return json_({ ok: false, error: "スプレッドシート側で管理者パスワードが未設定です。メニュー「工事費ツール ＞ 管理者パスワードを設定」から設定してください。" });
  }
  const cache = CacheService.getScriptCache();
  const failures = Number(cache.get("failures") || 0);
  if (failures >= MAX_FAILURES) {
    return json_({ ok: false, error: "パスワードの誤りが続いたため、しばらく保存を停止しています。10分ほど待ってから再度お試しください。" });
  }
  if (typeof req.password !== "string" || sha256Hex_(req.password) !== stored) {
    cache.put("failures", String(failures + 1), LOCK_SECONDS);
    return json_({ ok: false, error: "パスワードが違います。" });
  }
  cache.remove("failures");

  if (req.action !== "save") {
    return json_({ ok: false, error: "不明な操作です。" });
  }

  let config;
  try {
    config = validateConfig_(req.config);
  } catch (err) {
    return json_({ ok: false, error: err.message });
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) {
    return json_({ ok: false, error: "他の保存処理と重なりました。少し待ってから再度お試しください。" });
  }
  try {
    writeConfig_(config);
  } finally {
    lock.releaseLock();
  }
  return json_(Object.assign({ ok: true }, readConfig_()));
}

// ===== 読み書き =====

function readConfig_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rateRows = dataRows_(ss.getSheetByName(SHEET_NAMES.rates));
  if (!rateRows.length) return { empty: true, config: null, updatedAt: null };

  const rates = {};
  rateRows.forEach(r => {
    const key = String(r[0]).trim();
    const fee = toFee_(r[2]);
    if (key && fee !== null) rates[key] = fee;
  });

  const genreKey = v => {
    const s = String(v).trim();
    if (GENRES[s]) return s;
    return Object.keys(GENRES).find(k => GENRES[k] === s) || "";
  };

  const devicesSheet = ss.getSheetByName(SHEET_NAMES.devices);
  const devices = devicesSheet ? dataRows_(devicesSheet).map(r => ({
    value: String(r[0]).trim(),
    label: String(r[1]).trim(),
    genre: genreKey(r[2]),
    rate: toFee_(r[3])
  })) : null;

  const extras = dataRows_(ss.getSheetByName(SHEET_NAMES.extras)).map(r => ({
    id: String(r[0]).trim(),
    label: String(r[1]).trim(),
    unit: String(r[2]).trim(),
    rate: toFee_(r[3])
  }));

  const updatedAt = PropertiesService.getScriptProperties().getProperty("UPDATED_AT");
  return { empty: false, config: { rates: rates, devices: devices, extras: extras }, updatedAt: updatedAt };
}

function writeConfig_(config) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  writeSheet_(ss, SHEET_NAMES.rates, ["キー（変更しないでください）", "項目名", "金額（税込）"],
    config.rates.map(r => [r.key, r.label, r.rate]), [1], 3);
  writeSheet_(ss, SHEET_NAMES.devices, ["ID（変更しないでください）", "機器名", "ジャンル", "交換費（税込）"],
    config.devices.map(d => [d.value, d.label, GENRES[d.genre], d.rate]), [1, 2], 4);
  writeSheet_(ss, SHEET_NAMES.extras, ["ID（変更しないでください）", "項目名", "単位", "単価（税込）"],
    config.extras.map(x => [x.id, x.label, x.unit, x.rate]), [1, 2, 3], 4);

  const now = new Date();
  PropertiesService.getScriptProperties().setProperty("UPDATED_AT", now.toISOString());
  let log = ss.getSheetByName(SHEET_NAMES.log);
  if (!log) {
    log = ss.insertSheet(SHEET_NAMES.log);
    log.appendRow(["日時", "内容"]);
    log.getRange(1, 1, 1, 2).setFontWeight("bold");
    log.setFrozenRows(1);
  }
  log.appendRow([now, "管理者ページから保存（単価" + config.rates.length + "件・機器" +
    config.devices.length + "件・追加工事項目" + config.extras.length + "件）"]);
}

// シートの中身を丸ごと書き換える。textCols は文字列として扱う列（1始まり）、feeCol は金額列。
function writeSheet_(ss, name, header, rows, textCols, feeCol) {
  const sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  sheet.clearContents();
  const values = [header].concat(rows);
  textCols.forEach(c => sheet.getRange(1, c, values.length, 1).setNumberFormat("@"));
  sheet.getRange(1, 1, values.length, header.length).setValues(values);
  sheet.getRange(1, 1, 1, header.length).setFontWeight("bold").setBackground("#e7f1fb");
  if (rows.length) sheet.getRange(2, feeCol, rows.length, 1).setNumberFormat("#,##0");
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, header.length);
}

// ===== 入力チェック =====

function validateConfig_(c) {
  if (!c || typeof c !== "object") throw new Error("設定データがありません。");
  const text = (v, what) => {
    const s = typeof v === "string" ? v.trim() : "";
    if (!s) throw new Error(what + "が空です。");
    if (s.length > MAX_TEXT) throw new Error(what + "が長すぎます（" + MAX_TEXT + "文字まで）。");
    return s;
  };
  const fee = (v, what) => {
    if (typeof v !== "number" || !Number.isInteger(v) || v < 0) throw new Error(what + "の金額が正しくありません。");
    return v;
  };
  const list = (v, what) => {
    if (!Array.isArray(v)) throw new Error(what + "の形式が正しくありません。");
    if (v.length > MAX_ROWS) throw new Error(what + "の件数が多すぎます。");
    return v;
  };
  const unique = (arr, what) => {
    const seen = {};
    arr.forEach(id => { if (seen[id]) throw new Error(what + "のIDが重複しています。"); seen[id] = true; });
  };

  const rates = list(c.rates, "工事費単価").map(r => ({
    key: text(r && r.key, "単価のキー"), label: text(r && r.label, "単価の項目名"), rate: fee(r && r.rate, "単価")
  }));
  const devices = list(c.devices, "機器").map(d => {
    if (!d || !GENRES[d.genre]) throw new Error("機器のジャンルが正しくありません。");
    return { value: text(d.value, "機器のID"), label: text(d.label, "機器名"), genre: d.genre, rate: fee(d.rate, "機器") };
  });
  const extras = list(c.extras, "追加工事項目").map(x => ({
    id: text(x && x.id, "追加工事項目のID"), label: text(x && x.label, "追加工事項目の名称"),
    unit: text(x && x.unit, "追加工事項目の単位"), rate: fee(x && x.rate, "追加工事項目")
  }));
  unique(rates.map(r => r.key), "工事費単価");
  unique(devices.map(d => d.value), "機器");
  unique(extras.map(x => x.id), "追加工事項目");
  return { rates: rates, devices: devices, extras: extras };
}

// ===== ユーティリティ =====

function dataRows_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues()
    .filter(r => r.some(v => String(v).trim() !== ""));
}

// シートを直接編集した場合に「8,800円」のような入力も受け付ける。
function toFee_(v) {
  const n = typeof v === "number" ? v : Number(String(v).replace(/[,，円\s]/g, ""));
  return Number.isInteger(n) && n >= 0 ? n : null;
}

function sha256Hex_(text) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8)
    .map(b => ((b + 256) % 256).toString(16).padStart(2, "0")).join("");
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
