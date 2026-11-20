function doGet() {
  const docName = SpreadsheetApp.getActiveSpreadsheet().getName();
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle(docName)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 試算表上方新增功能選單（已更新為最新 0~5 步驟）
// 試算表選單 (新增：一鍵產生說明分頁)
// 更新選單（新增步驟 6）
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🚀 自動化工具')
    .addItem('📖 建立/更新『使用說明』分頁', 'createHelpSheet')
    .addSeparator()
    .addItem('0. 一鍵清空，開始新的報名', 'clearAllData')
    .addItem('1. 貼上「原始名單」與「導師名單」(說明)', 'showStep1Instruction')
    .addItem('2. 設定題目與活動時間', 'openAdminDialog')
    .addItem('3. 依原始名單建立班級分頁', 'initClassSheets')
    .addItem('4. 進入 Apps Script 部署網頁與複製連結(說明)', 'showStep4Instruction')
    .addSeparator()
    .addItem('5. 匯出班級試算表並寄信給導師', 'exportAndEmailSheets')
    .addItem('🔄 手動同步更新已寄出的導師試算表', 'syncAllExportedSheets')
    .addItem('6. 🧾 班級收據批次列印', 'openReceiptDialog')
    .addToUi();
}

// 步驟 1 提示說明視窗
function showStep1Instruction() {
  SpreadsheetApp.getUi().alert(
    '📌 步驟 1：資料準備說明',
    '請依序至對應分頁貼上資料：\n\n' +
    '1. 『原始名單』分頁：\n' +
    '   - 第 1 列為表頭（班級、座號、學號、姓名、生日、身分證字號...）\n' +
    '   - 第 2 列起請貼上學生資料。\n\n' +
    '2. 『導師名單』分頁：\n' +
    '   - 第 1 列為表頭（班級、導師、Email）\n' +
    '   - 第 2 列起請貼上各班導師聯絡資訊。',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// 開啟管理端彈跳視窗
function openAdminDialog() {
  const html = HtmlService.createHtmlOutputFromFile('Admin')
    .setWidth(850)
    .setHeight(700)
    .setTitle('⚙️ 設定題目與活動時間');
  SpreadsheetApp.getUi().showModalDialog(html, '⚙️ 設定題目與活動時間');
}

// 取得管理端設定資料
function getFormAdminConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let respSheet = ss.getSheetByName("表單回復");

  if (!respSheet) return null;

  const description = respSheet.getRange("A1").getValue() || "";
  const startTime = respSheet.getRange("A2").getValue();
  const endTime = respSheet.getRange("A3").getValue();

  const fmtStart = (startTime instanceof Date) ? Utilities.formatDate(startTime, "GMT+8", "yyyy-MM-dd'T'HH:mm") : "";
  const fmtEnd = (endTime instanceof Date) ? Utilities.formatDate(endTime, "GMT+8", "yyyy-MM-dd'T'HH:mm") : "";

  const lastCol = respSheet.getLastColumn();
  const questions = [];

  if (lastCol >= 6) {
    const rawQuestions = respSheet.getRange(1, 6, 3, lastCol - 5).getValues();
    for (let c = 0; c < rawQuestions[0].length; c++) {
      const qTitle = rawQuestions[0][c];
      const qType = String(rawQuestions[1][c]).trim();
      const qOptions = rawQuestions[2][c];
      
      if (!qTitle && !qType) continue;

      questions.push({
        title: qTitle,
        type: qType,
        options: String(qOptions || "")
      });
    }
  }

  return {
    description: description,
    startTime: fmtStart,
    endTime: fmtEnd,
    questions: questions
  };
}

// 儲存管理端設定資料並同步班級表頭
function saveFormAdminConfig(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let respSheet = ss.getSheetByName("表單回復");

  if (!respSheet) {
    respSheet = ss.insertSheet("表單回復");
  }

  // 1. 寫入 A1:A3 (活動說明、開放時間、截止時間)
  respSheet.getRange("A1").setValue(payload.description || "");

  if (payload.startTime) {
    respSheet.getRange("A2").setValue(new Date(payload.startTime));
  } else {
    respSheet.getRange("A2").clearContent();
  }

  if (payload.endTime) {
    respSheet.getRange("A3").setValue(new Date(payload.endTime));
  } else {
    respSheet.getRange("A3").clearContent();
  }

  // 2. 清空舊的題目欄位 (F欄起至最右邊)
  const lastCol = Math.max(respSheet.getLastColumn(), 6);
  if (lastCol >= 6) {
    respSheet.getRange(1, 6, 3, lastCol - 5 + 1).clearContent();
  }

  // 3. 寫入新的題目資料 (寫入第 1~3 列)
  if (payload.questions && payload.questions.length > 0) {
    const titles = [];
    const types = [];
    const options = [];

    payload.questions.forEach(q => {
      titles.push(q.title || "");
      types.push(q.type || "簡答");
      options.push(q.options || "");
    });

    respSheet.getRange(1, 6, 3, payload.questions.length).setValues([titles, types, options]);
  }

  // 4. 自動同步更新班級分頁表頭
  initClassSheets();

  return { success: true };
}

// 判斷是否為總計欄位
function isTotalCol(q) {
  if (!q) return false;
  const typeStr = String(q.type || "").trim();
  const titleStr = String(q.title || "").trim();
  return typeStr.indexOf("總計") !== -1 || titleStr.indexOf("總計") !== -1;
}

// A1-Z26 欄位英文字母轉換輔助函式 (用於建立公式)
function getColLetter(colIndex) {
  let temp, letter = '';
  while (colIndex > 0) {
    temp = (colIndex - 1) % 26;
    letter = String.fromCharCode(65 + temp) + letter;
    colIndex = Math.floor((colIndex - temp - 1) / 26);
  }
  return letter;
}

// 根據原始名單自動建立各班分頁 (已加入 A1 班級標題與底部 SUM 加總公式)
function initClassSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName("原始名單");
  const respSheet = ss.getSheetByName("表單回復");

  if (!rawSheet || !respSheet) {
    SpreadsheetApp.getUi().alert("❌ 找不到『原始名單』或『表單回復』分頁！");
    return;
  }

  const rawData = rawSheet.getDataRange().getValues();

  // 檢查第二列起是否有學生資料
  if (rawData.length <= 1) {
    SpreadsheetApp.getUi().alert("⚠️『原始名單』第 2 列起沒有學生資料！請先貼上資料後再執行本功能。");
    return;
  }

  // 1. 動態抓取「原始名單」基本欄位標題 (座號, 學號, 姓名, 生日...)
  const profileHeaders = rawData[0].slice(1);

  // 2. 抓取「表單回復」題目名稱
  let lastCol = respSheet.getLastColumn();
  if (lastCol < 5) lastCol = 5;

  let questionHeaders = [];
  if (lastCol >= 6) {
    questionHeaders = respSheet.getRange(1, 6, 1, lastCol - 5).getValues()[0];
  }

  // 3. 同步寫回「表單回復」第 4 列表頭
  const fullRespHeaders = ["班級", "座號", "學號", "姓名", "時間"].concat(questionHeaders);
  respSheet.getRange(4, 1, 1, fullRespHeaders.length).setValues([fullRespHeaders]).setFontWeight("bold");

  // 4. 建立班級分頁表頭
  const classHeaders = profileHeaders.concat(["時間"]).concat(questionHeaders);

  const classMap = {};

  for (let i = 1; i < rawData.length; i++) {
    const className = String(rawData[i][0]).trim();
    if (!className) continue;

    const studentProfile = rawData[i].slice(1);
    if (!classMap[className]) classMap[className] = [];
    classMap[className].push(studentProfile);
  }

  // 建立/更新各班分頁
  for (let className in classMap) {
    let sheet = ss.getSheetByName(className);
    if (!sheet) {
      sheet = ss.insertSheet(className);
    } else {
      sheet.clear();
    }

    // 格式調整 1：A1 填上班級名稱（加粗、加大字體方便列印）
    const titleText = className.indexOf("班") !== -1 ? `${className} 報名統計表` : `${className}班 報名統計表`;
    sheet.getRange("A1").setValue(titleText).setFontWeight("bold").setFontSize(14);

    // 格式調整 2：表頭移至第 2 列
    sheet.getRange(2, 1, 1, classHeaders.length).setValues([classHeaders]);
    sheet.getRange(2, 1, 1, classHeaders.length).setFontWeight("bold").setBackground("#e9ecef");

    // 格式調整 3：學生資料從第 3 列開始填入
    const students = classMap[className];
    const studentCount = students.length;
    if (studentCount > 0) {
      sheet.getRange(3, 1, studentCount, profileHeaders.length).setValues(students);
    }

    // 格式調整 4：學生資料最下方建立「合計」與 SUM 公式列
    const totalRowIndex = 3 + studentCount;
    sheet.getRange(totalRowIndex, 1).setValue("合計").setFontWeight("bold");

    // 檢查每一個欄位，若表頭包含「總計」或「金額」，自動寫入 =SUM(X3:XN)
    for (let c = 0; c < classHeaders.length; c++) {
      const headerTitle = String(classHeaders[c] || "");
      if (headerTitle.indexOf("總計") !== -1 || headerTitle.indexOf("金額") !== -1) {
        const colLetter = getColLetter(c + 1);
        const sumFormula = `=SUM(${colLetter}3:${colLetter}${totalRowIndex - 1})`;
        sheet.getRange(totalRowIndex, c + 1).setFormula(sumFormula).setFontWeight("bold");
      }
    }

    // 設定合計列背景樣式
    sheet.getRange(totalRowIndex, 1, 1, classHeaders.length).setBackground("#f8f9fa").setFontWeight("bold");
  }

  try {
    syncAllExportedClassSheets_(false);
  } catch (e) {}

  SpreadsheetApp.getUi().alert("✅ 班級分頁已建立完成！\n已更新 A1 班級標題，並於底部自動加入金額加總公式。");
}

// 依班級名稱尋找對應的班級分頁（容許前後空白與「班」字尾差異）
function findClassSheet_(ss, className) {
  const normalized = String(className || "").trim();
  if (!normalized) return null;

  const exact = ss.getSheetByName(normalized);
  if (exact) return exact;

  const altName = normalized.endsWith("班")
    ? normalized.slice(0, -1).trim()
    : normalized + "班";

  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    const sName = sheets[i].getName().trim();
    if (sName === normalized || (altName && sName === altName)) {
      return sheets[i];
    }
  }
  return null;
}

// 偵測班級分頁的表頭列位置（僅比對文字儲存格，避免日期物件轉字串「台北標準時間」誤判）
function detectClassHeaderRow_(classSheet, lastRow, lastCol) {
  const maxCheckRow = Math.min(3, lastRow);
  const candidates = [];
  if (maxCheckRow >= 2) candidates.push(2);
  candidates.push(1);
  if (maxCheckRow >= 3) candidates.push(3);

  const headerKeywords = ["班級", "座號", "學號", "姓名", "時間", "報名時間", "填報時間", "填寫時間"];
  let bestRow = null;
  let bestScore = 0;

  for (let i = 0; i < candidates.length; i++) {
    const r = candidates[i];
    const rowValues = classSheet.getRange(r, 1, 1, lastCol).getValues()[0];
    let score = 0;
    for (let c = 0; c < rowValues.length; c++) {
      const cell = rowValues[c];
      if (typeof cell !== "string") continue; // 排除 Date 與數字儲存格
      const text = cell.trim();
      if (!text) continue;
      if (headerKeywords.indexOf(text) !== -1 || (text.length <= 10 && text.indexOf("時間") !== -1)) {
        score++;
      }
    }
    if (score > bestScore) {
      bestScore = score;
      bestRow = { rowIndex: r, headers: rowValues };
    }
  }

  if (bestRow) return bestRow;

  const fallbackRow = lastRow >= 2 ? 2 : 1;
  return {
    rowIndex: fallbackRow,
    headers: classSheet.getRange(fallbackRow, 1, 1, lastCol).getValues()[0]
  };
}

// 在已修剪空白的表頭陣列中尋找「時間」欄位索引 (0-based，找不到回傳 -1)
function findTimeColIndex_(trimmedHeaders) {
  const exactIdx = trimmedHeaders.indexOf("時間");
  if (exactIdx !== -1) return exactIdx;
  return trimmedHeaders.findIndex(h => h.indexOf("時間") !== -1);
}

// 讀取「班級分頁名稱 -> 已匯出試算表 ID」對照表
function getExportedSheetsMap_() {
  const raw = PropertiesService.getDocumentProperties().getProperty("EXPORTED_CLASS_SHEETS_MAP");
  if (!raw) return {};
  try {
    const parsed = JSON.parse(raw);
    return (parsed && typeof parsed === "object") ? parsed : {};
  } catch (e) {
    return {};
  }
}

// 儲存「班級分頁名稱 -> 已匯出試算表 ID」對照表
function saveExportedSheetsMap_(map) {
  PropertiesService.getDocumentProperties().setProperty("EXPORTED_CLASS_SHEETS_MAP", JSON.stringify(map || {}));
}

// 將來源班級分頁的最新內容同步至導師獨立試算表（維度相符時原地更新，否則整頁替換以保留相同網址）
function syncClassSheetContent_(sourceSheet, targetSs) {
  const sheetName = sourceSheet.getName();
  const lastRow = sourceSheet.getLastRow();
  const lastCol = sourceSheet.getLastColumn();
  const destSheet = findClassSheet_(targetSs, sheetName) || targetSs.getSheets()[0];

  if (destSheet && lastRow > 0 && lastCol > 0 &&
      destSheet.getLastRow() === lastRow && destSheet.getLastColumn() === lastCol) {
    const srcRange = sourceSheet.getRange(1, 1, lastRow, lastCol);
    const values = srcRange.getValues();
    const formulas = srcRange.getFormulas();
    for (let r = 0; r < values.length; r++) {
      for (let c = 0; c < values[r].length; c++) {
        if (formulas[r][c]) {
          values[r][c] = formulas[r][c];
        }
      }
    }
    destSheet.getRange(1, 1, lastRow, lastCol).setValues(values);
    if (destSheet.getName() !== sheetName) {
      destSheet.setName(sheetName);
    }
    return;
  }

  const copied = sourceSheet.copyTo(targetSs);
  targetSs.getSheets().forEach(s => {
    if (s.getSheetId() !== copied.getSheetId()) {
      targetSs.deleteSheet(s);
    }
  });
  copied.setName(sheetName);
}

// 依據主表單班級分頁，即時同步更新對應的已匯出導師試算表
function syncExportedClassSheet_(classSheet) {
  if (!classSheet) return false;
  const map = getExportedSheetsMap_();
  const sheetName = classSheet.getName().trim();
  const altName = sheetName.endsWith("班") ? sheetName.slice(0, -1).trim() : sheetName + "班";
  const exportedId = map[sheetName] || map[altName];
  if (!exportedId) return false;

  try {
    const targetSs = SpreadsheetApp.openById(exportedId);
    syncClassSheetContent_(classSheet, targetSs);
    return true;
  } catch (e) {
    return false;
  }
}

// 批次同步所有已建立的導師獨立試算表
function syncAllExportedClassSheets_(showUi) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const map = getExportedSheetsMap_();
  const entries = Object.keys(map);
  if (entries.length === 0) {
    if (showUi) {
      SpreadsheetApp.getUi().alert("ℹ️ 目前尚未匯出任何導師獨立試算表，請先執行『5. 匯出班級試算表並寄信給導師』。");
    }
    return { updated: 0, skipped: 0 };
  }

  let updated = 0;
  let skipped = 0;
  entries.forEach(clsName => {
    const classSheet = findClassSheet_(ss, clsName);
    if (!classSheet) {
      skipped++;
      return;
    }
    try {
      const targetSs = SpreadsheetApp.openById(map[clsName]);
      syncClassSheetContent_(classSheet, targetSs);
      updated++;
    } catch (e) {
      skipped++;
    }
  });

  if (showUi) {
    SpreadsheetApp.getUi().alert(`✅ 導師試算表同步完成！\n已更新：${updated} 班${skipped > 0 ? `\n未更新（檔案已刪除或無對應分頁）：${skipped} 班` : ""}`);
  }
  return { updated: updated, skipped: skipped };
}

// 選單入口：手動同步更新已寄出的導師試算表
function syncAllExportedSheets() {
  syncAllExportedClassSheets_(true);
}

// 可安裝的編輯觸發器：承辦人手動修改主表班級分頁時自動同步至導師試算表
function onMasterSheetEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  const ignoreSheets = ["表單回復", "原始名單", "導師名單", "📜 使用說明"];
  if (ignoreSheets.indexOf(sheet.getName()) !== -1) return;
  syncExportedClassSheet_(sheet);
}

// 確保已安裝主表編輯自動同步觸發器
function ensureEditSyncTrigger_(ss) {
  const triggers = ScriptApp.getProjectTriggers();
  const exists = triggers.some(t => t.getHandlerFunction() === "onMasterSheetEdit");
  if (!exists) {
    ScriptApp.newTrigger("onMasterSheetEdit").forSpreadsheet(ss).onEdit().create();
  }
}

// 批次匯出各班獨立試算表並 Email 寄給導師（含無人報名班級，且固定檔案與連結以支援動態同步更新）
function exportAndEmailSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const docName = ss.getName().trim();
  const teacherSheet = ss.getSheetByName("導師名單");

  if (!teacherSheet) {
    SpreadsheetApp.getUi().alert("❌ 找不到『導師名單』分頁！請先建立分頁並填寫『班級』、『導師』、『Email』。");
    return;
  }

  const teacherData = teacherSheet.getDataRange().getValues();
  if (teacherData.length <= 1) {
    SpreadsheetApp.getUi().alert("❌『導師名單』分頁沒有填寫資料！");
    return;
  }

  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert(
    '📧 批次匯出與寄信確認（動態更新模式）',
    '系統將會為『導師名單』中的所有有效班級建立（或沿用既有）獨立試算表並寄信給導師：\n' +
    '1. 尚無學生報名的班級也會先建立統計表並寄出連結。\n' +
    '2. 之後學生送出／修改報名或承辦人編輯名單，會自動同步更新至同一份試算表連結。',
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) return;

  const teacherHeaders = teacherData[0].map(h => String(h || "").trim());
  const tClassIdx = teacherHeaders.indexOf("班級") !== -1 ? teacherHeaders.indexOf("班級") : 0;
  const tNameIdx = teacherHeaders.indexOf("導師") !== -1 ? teacherHeaders.indexOf("導師") : 1;
  let tEmailIdx = teacherHeaders.findIndex(h => h.toLowerCase().indexOf("email") !== -1 || h.indexOf("信箱") !== -1 || h.indexOf("郵件") !== -1);
  if (tEmailIdx === -1) tEmailIdx = 2;

  const exportedMap = getExportedSheetsMap_();
  let createdCount = 0;
  let updatedCount = 0;
  let skipCount = 0;
  let failCount = 0;
  let logs = [];

  for (let i = 1; i < teacherData.length; i++) {
    const className = String(teacherData[i][tClassIdx] || "").trim();
    const teacherName = String(teacherData[i][tNameIdx] || "").trim();
    const email = String(teacherData[i][tEmailIdx] || "").trim();

    if (!className) continue;

    const classLabel = className.indexOf("班") !== -1 ? className : `${className}班`;

    if (!email) {
      logs.push(`⏭️ 跳過 ${classLabel} (${teacherName || "未填導師"})：未填寫 Email`);
      skipCount++;
      continue;
    }

    const classSheet = findClassSheet_(ss, className);

    if (!classSheet || classSheet.getLastRow() <= 1 || classSheet.getLastColumn() === 0) {
      logs.push(`⏭️ 跳過 ${classLabel}：分頁不存在或無學生名單`);
      skipCount++;
      continue;
    }

    const headerInfo = detectClassHeaderRow_(classSheet, classSheet.getLastRow(), classSheet.getLastColumn());
    if (classSheet.getLastRow() <= headerInfo.rowIndex) {
      logs.push(`⏭️ 跳過 ${classLabel}：無學生名單資料`);
      skipCount++;
      continue;
    }

    try {
      const actualSheetName = classSheet.getName().trim();
      const newFileName = `${docName} - ${classLabel} 報名統計`;
      let targetSs = null;
      let isReused = false;
      const existingId = exportedMap[actualSheetName] || exportedMap[className];

      if (existingId) {
        try {
          targetSs = SpreadsheetApp.openById(existingId);
          if (targetSs.getName() !== newFileName) {
            targetSs.setName(newFileName);
          }
          syncClassSheetContent_(classSheet, targetSs);
          isReused = true;
        } catch (openErr) {
          targetSs = null;
        }
      }

      if (!targetSs) {
        targetSs = SpreadsheetApp.create(newFileName);
        syncClassSheetContent_(classSheet, targetSs);
      }

      exportedMap[actualSheetName] = targetSs.getId();
      targetSs.addViewer(email);
      const fileUrl = targetSs.getUrl();

      const subject = `【報名統計通知】${docName} - ${classLabel} 學生報名結果`;
      const body = `${teacherName} 老師您好：\n\n` +
                   `這是貴班（${classLabel}）在「${docName}」活動的學生報名統計表：\n\n` +
                   `🔗 統計表連結（將隨報名狀況動態更新）：\n${fileUrl}\n\n` +
                   `此信件為系統自動發送，之後若有學生新增或修改報名，開啟同一連結重新整理即可即時查看最新名單。`;

      MailApp.sendEmail(email, subject, body);
      if (isReused) {
        updatedCount++;
      } else {
        createdCount++;
      }
    } catch (e) {
      logs.push(`❌ ${classLabel} (${teacherName}) 處理失敗：${e.message}`);
      failCount++;
    }
  }

  saveExportedSheetsMap_(exportedMap);

  try {
    ensureEditSyncTrigger_(ss);
  } catch (triggerErr) {
    logs.push(`⚠️ 自動同步編輯觸發器建立失敗（請改用選單手動同步）：${triggerErr.message}`);
  }

  let resultMsg = `✅ 處理完成！\n新建並寄出：${createdCount} 封\n更新並寄出：${updatedCount} 封\n跳過（缺 Email / 無名單）：${skipCount} 班`;
  if (failCount > 0) resultMsg += `\n執行失敗：${failCount} 班`;
  if (logs.length > 0) resultMsg += `\n\n詳細紀錄：\n` + logs.join("\n");

  ui.alert(resultMsg);
}

// 一鍵清空功能
function clearAllData() {
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert(
    '⚠️ 警告：確定要清空嗎？',
    '此操作將會：\n1. 刪除所有「班級分頁」\n2. 清空「表單回復」的所有學生報名紀錄（保留前 4 列設定）\n\n請問是否確定繼續？',
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const keepSheets = ["原始名單", "表單回復", "導師名單"];
  const sheets = ss.getSheets();
  for (let sheet of sheets) {
    if (!keepSheets.includes(sheet.getName())) {
      ss.deleteSheet(sheet);
    }
  }

  const respSheet = ss.getSheetByName("表單回復");
  if (respSheet) {
    const lastRow = respSheet.getLastRow();
    if (lastRow >= 5) {
      respSheet.getRange(5, 1, lastRow - 4, respSheet.getLastColumn()).clearContent();
    }
  }

  PropertiesService.getDocumentProperties().deleteProperty("EXPORTED_CLASS_SHEETS_MAP");

  ui.alert("✅ 已成功清空所有報名資料與班級分頁！");
}

// 取得前台網頁表單設定
function getFormConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const respSheet = ss.getSheetByName("表單回復");

  const docName = ss.getName();
  const description = respSheet.getRange("A1").getValue() || "";
  const startTime = respSheet.getRange("A2").getValue();
  const endTime = respSheet.getRange("A3").getValue();

  const now = new Date();
  let status = "active";
  let message = "";

  if (startTime && new Date(startTime) > now) {
    status = "not_started";
    message = "報名尚未開放，開放時間為：" + Utilities.formatDate(new Date(startTime), "GMT+8", "yyyy/MM/dd HH:mm");
  } else if (endTime && new Date(endTime) < now) {
    status = "ended";
    message = "報名已經截止，截止時間為：" + Utilities.formatDate(new Date(endTime), "GMT+8", "yyyy/MM/dd HH:mm");
  }

  const lastCol = respSheet.getLastColumn();
  const questions = [];
  
  if (lastCol >= 6) {
    const rawQuestions = respSheet.getRange(1, 6, 3, lastCol - 5).getValues();
    for (let c = 0; c < rawQuestions[0].length; c++) {
      const qTitle = rawQuestions[0][c];
      const qType = String(rawQuestions[1][c]).trim();
      const qOptionsRaw = rawQuestions[2][c];

      if (!qTitle) continue;

      let options = [];
      if (qOptionsRaw) {
        options = String(qOptionsRaw).split(',').map(s => s.trim());
      }

      questions.push({
        colIndex: c + 6,
        title: qTitle,
        type: qType,
        options: options
      });
    }
  }

  return {
    docName: docName,
    status: status,
    message: message,
    description: description,
    questions: questions
  };
}

// 1. 搜尋學生資料（已整合重複報名檢視與舊資料載入）
function searchStudent(studentId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetId = String(studentId).trim();
  const respSheet = ss.getSheetByName("表單回復");
  const rawSheet = ss.getSheetByName("原始名單");

  if (!rawSheet) return { found: false, message: "找不到『原始名單』分頁！" };

  // 先尋找學生的基本資料
  const rawData = rawSheet.getDataRange().getValues();
  let studentInfo = null;
  for (let i = 1; i < rawData.length; i++) {
    if (String(rawData[i][2]).trim() === targetId) { // C欄位為學號
      studentInfo = {
        className: String(rawData[i][0]).trim(),
        seatNo: rawData[i][1],
        studentId: String(rawData[i][2]).trim(),
        name: rawData[i][3]
      };
      break;
    }
  }

  if (!studentInfo) return { found: false, message: "查無此學號，請確認輸入是否正確！" };

  // 檢查「表單回復」是否已報名過
  if (respSheet && respSheet.getLastRow() >= 5) {
    const lastCol = respSheet.getLastColumn();
    const headers = respSheet.getRange(4, 1, 1, lastCol).getValues()[0];
    const respData = respSheet.getRange(5, 1, respSheet.getLastRow() - 4, lastCol).getValues();

    for (let i = 0; i < respData.length; i++) {
      if (String(respData[i][2]).trim() === targetId) { // 第 3 欄為學號
        // 包裝舊的答題紀錄 (從第 6 欄起為題目回答)
        const previousAnswers = {};
        for (let c = 5; c < lastCol; c++) {
          const qTitle = headers[c];
          previousAnswers[qTitle] = respData[i][c];
        }

        return {
          found: true,
          alreadyRegistered: true,
          registerTime: respData[i][4] ? Utilities.formatDate(new Date(respData[i][4]), "GMT+8", "yyyy/MM/dd HH:mm") : "已完成",
          studentInfo: studentInfo,
          previousAnswers: previousAnswers,
          rowIndexInResp: i + 5 // 記錄在表單回復的列號
        };
      }
    }
  }

  return {
    found: true,
    alreadyRegistered: false,
    studentInfo: studentInfo
  };
}
// 2. 儲存/更新報名資料 (避免重複寫入，並即時同步至導師試算表)
function saveRegistration(formData) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return { success: false, message: "系統目前較忙碌，請稍候幾秒再重新送出！" };
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const respSheet = ss.getSheetByName("表單回復");
    const classSheet = findClassSheet_(ss, formData.className);
    const nowStr = Utilities.formatDate(new Date(), "GMT+8", "yyyy/MM/dd HH:mm:ss");

    if (!respSheet || !classSheet) return { success: false, message: "資料表結構異常！" };

    // A. 寫入或更新「表單回復」分頁
    const respLastCol = respSheet.getLastColumn();
    const respHeaders = respSheet.getRange(4, 1, 1, respLastCol).getValues()[0];
    const respTypes = respSheet.getRange(2, 1, 1, respLastCol).getValues()[0];
    const rowData = [formData.className, formData.seatNo, formData.studentId, formData.name, nowStr];
    const totalTitleSet = new Set();

    // 依照題目順序填入答案（含總計欄位）
    for (let c = 5; c < respHeaders.length; c++) {
      const rawTitle = respHeaders[c];
      const qTitle = String(rawTitle || "").trim();
      const qType = String(respTypes[c] || "").trim();
      if (isTotalCol({ title: qTitle, type: qType })) {
        totalTitleSet.add(qTitle);
        rowData.push(formData.totalAmount !== undefined ? formData.totalAmount : "");
      } else {
        const ans = formData.answers[rawTitle] !== undefined
          ? formData.answers[rawTitle]
          : (formData.answers[qTitle] !== undefined ? formData.answers[qTitle] : "");
        rowData.push(ans);
      }
    }

    // 尋找是否已存在「表單回復」中
    let existingRespRow = -1;
    if (respSheet.getLastRow() >= 5) {
      const ids = respSheet.getRange(5, 3, respSheet.getLastRow() - 4, 1).getValues();
      for (let r = 0; r < ids.length; r++) {
        if (String(ids[r][0]).trim() === String(formData.studentId).trim()) {
          existingRespRow = r + 5;
          break;
        }
      }
    }

    if (existingRespRow !== -1) {
      // 更新舊紀錄
      respSheet.getRange(existingRespRow, 1, 1, rowData.length).setValues([rowData]);
    } else {
      // 新增紀錄
      respSheet.appendRow(rowData);
    }

    // B. 同步更新「班級分頁」
    const classLastRow = classSheet.getLastRow();
    const classLastCol = classSheet.getLastColumn();
    if (classLastRow > 1 && classLastCol > 0) {
      const headerInfo = detectClassHeaderRow_(classSheet, classLastRow, classLastCol);
      const trimmedHeaders = headerInfo.headers.map(h => String(h || "").trim());
      const dataRowCount = classLastRow - headerInfo.rowIndex;

      if (dataRowCount > 0) {
        const classData = classSheet.getRange(headerInfo.rowIndex + 1, 1, dataRowCount, classLastCol).getValues();
        let targetClassRow = -1;
        const idColIdx = trimmedHeaders.indexOf("學號");

        if (idColIdx !== -1) {
          for (let r = 0; r < classData.length; r++) {
            if (String(classData[r][idColIdx]).trim() === String(formData.studentId).trim()) {
              targetClassRow = r + headerInfo.rowIndex + 1;
              break;
            }
          }
        }

        if (targetClassRow !== -1) {
          const timeColIdx = findTimeColIndex_(trimmedHeaders) + 1;
          if (timeColIdx > 0) classSheet.getRange(targetClassRow, timeColIdx).setValue(nowStr);

          for (let qTitle in formData.answers) {
            const colIdx = trimmedHeaders.indexOf(String(qTitle).trim()) + 1;
            if (colIdx > 0) {
              classSheet.getRange(targetClassRow, colIdx).setValue(formData.answers[qTitle]);
            }
          }

          if (formData.totalAmount !== undefined) {
            for (let c = 0; c < trimmedHeaders.length; c++) {
              const h = trimmedHeaders[c];
              if (h && (totalTitleSet.has(h) || isTotalCol({ title: h }))) {
                classSheet.getRange(targetClassRow, c + 1).setValue(formData.totalAmount);
              }
            }
          }
        }
      }
    }

    // C. 即時同步更新已匯出的導師獨立試算表
    try {
      syncExportedClassSheet_(classSheet);
    } catch (syncErr) {}

    return { success: true, isUpdate: existingRespRow !== -1 };
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

// 提交報名資料
function submitForm(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const respSheet = ss.getSheetByName("表單回復");
  const rawSheet = ss.getSheetByName("原始名單");
  const timestamp = Utilities.formatDate(new Date(), "GMT+8", "yyyy/MM/dd HH:mm:ss");

  const config = getFormConfig();
  const answers = payload.answers;

  let totalAmount = 0;
  for (let i = 0; i < config.questions.length; i++) {
    const q = config.questions[i];
    const userAns = answers[q.colIndex] || "";

    if (q.type.includes("單選(金額)") || q.type.includes("單選:金額")) {
      let basePrice = parseNumber(userAns);
      if (i + 1 < config.questions.length && (config.questions[i + 1].type.includes("單選(折扣)") || config.questions[i + 1].type.includes("單選:折扣"))) {
        const discountAns = answers[config.questions[i + 1].colIndex] || "";
        const discountRate = parseDiscount(discountAns);
        basePrice = basePrice * discountRate;
      }
      totalAmount += basePrice;
    } else if (q.type.includes("多選(金額)") || q.type.includes("多選:金額")) {
      if (Array.isArray(userAns)) {
        userAns.forEach(item => { totalAmount += parseNumber(item); });
      } else if (userAns) {
        totalAmount += parseNumber(userAns);
      }
    }
  }

  const formattedRow = [
    payload.className,
    payload.seatNo,
    payload.studentId,
    payload.name,
    timestamp
  ];

  for (let i = 0; i < config.questions.length; i++) {
    const q = config.questions[i];
    const userAns = answers[q.colIndex] || "";

    if (isTotalCol(q)) {
      formattedRow.push(totalAmount);
    } else {
      formattedRow.push(Array.isArray(userAns) ? userAns.join(", ") : userAns);
    }
  }

  respSheet.appendRow(formattedRow);

  const classSheet = ss.getSheetByName(payload.className);
  if (classSheet && rawSheet && payload.rowIndex > 0) {
    const startCol = rawSheet.getLastColumn();
    const classUpdateValues = [timestamp];

    for (let i = 0; i < config.questions.length; i++) {
      const q = config.questions[i];
      const userAns = answers[q.colIndex] || "";

      if (isTotalCol(q)) {
        classUpdateValues.push(totalAmount);
      } else {
        classUpdateValues.push(Array.isArray(userAns) ? userAns.join(", ") : userAns);
      }
    }

    classSheet.getRange(payload.rowIndex, startCol, 1, classUpdateValues.length).setValues([classUpdateValues]);
    try {
      syncExportedClassSheet_(classSheet);
    } catch (syncErr) {}
  }

  return { success: true, totalAmount: totalAmount };
}

function parseNumber(text) {
  if (!text) return 0;
  const match = String(text).match(/\((\d+)\)/);
  return match ? parseFloat(match[1]) : 0;
}

function parseDiscount(text) {
  if (!text) return 1;
  const match = String(text).match(/\((\d+)%\)/);
  return match ? parseFloat(match[1]) / 100 : 1;
}
// 步驟 4：網頁部署說明視窗
function showStep4Instruction() {
  SpreadsheetApp.getUi().alert(
    '📌 步驟 4：部署網頁與取得表單連結說明',
    '請依以下步驟發布學生報名網頁：\n\n' +
    '1. 點擊上方選單『擴充功能』➔『Apps Script』開啟編輯器。\n' +
    '2. 點選右上角『部署』➔『新增部署』。\n' +
    '3. 點擊齒輪圖示，選擇『Web 應用程式』：\n' +
    '   - 執行身份：選擇『我的帳戶 (您自己)』\n' +
    '   - 誰有存取權限：選擇『任何人』或『您網域內的任何人』\n' +
    '4. 點擊『部署』按鈕並完成帳號授權。\n' +
    '5. 複製畫面產生的『Web 應用程式 URL』，即可發送給學生或貼至校園公告！',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
// 自動建立/更新「📜 使用說明」分頁
function createHelpSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("📜 使用說明");

  if (!sheet) {
    sheet = ss.insertSheet("📜 使用說明", 0); // 放在最前頁
  } else {
    sheet.clear();
  }

  // 設定整體欄寬
  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(2, 280);
  sheet.setColumnWidth(3, 450);

  const docsData = [
    ["🚀 自動化報名系統 — 各分頁格式與讀取邏輯說明", "", ""],
    ["", "", ""],
    
    ["📄 1.『原始名單』分頁", "", ""],
    ["欄位名稱", "格式要求", "系統讀取與運作邏輯"],
    ["第 1 列 (表頭)", "A欄：班級\nB欄起：座號、學號、姓名、生日...", "B欄起的名稱會自動當作各班分頁的個人基本資料欄位（可自由增加/修改欄位名稱）。"],
    ["第 2 列起 (資料)", "填寫學生資料", "1. 系統執行『3. 建立班級分頁』時，若第 2 列無資料會停止執行。\n2. 系統依 A 欄(班級)進行分班分類。\n3. 學生線上報名時，系統會比對 C 欄(學號)驗證身分。"],
    ["", "", ""],

    ["📄 2.『導師名單』分頁", "", ""],
    ["欄位名稱", "格式要求", "系統讀取與運作邏輯"],
    ["第 1 列 (表頭)", "A欄：班級 | B欄：導師 | C欄：Email", "表頭欄位名稱固定，請勿更動。"],
    ["第 2 列起 (資料)", "填寫各班導師資訊", "執行『5. 匯出並寄信』時，系統會比對班級報名狀態：\n- 有學生報名：自動建立獨立副本並寄信給該班導師。\n- 無人報名：自動跳過不寄信。"],
    ["", "", ""],

    ["📄 3.『表單回復』分頁", "", ""],
    ["欄位名稱", "格式要求", "系統讀取與運作邏輯"],
    ["第 1 列 (題目備份)", "F欄起為題目名稱", "「2. 設定題目」時系統自動填入，作為全校題目備份。"],
    ["第 4 列 (總表頭)", "系統自動建立", "整合：班級、座號、學號、姓名、時間、題目1, 題目2..."],
    ["第 5 列起 (報名紀錄)", "系統自動寫入", "1. 學生從網頁報名成功後，資料即時寫入此處。\n2. 系統會比對 C 欄(學號)判斷是否重複報名。"],
    ["", "", ""],

    ["📄 4.『各班級分頁』（如：101、102...）", "", ""],
    ["欄位名稱", "格式要求", "系統讀取與運作邏輯"],
    ["第 1 列 (A1)", "班級名稱標題", "自動顯示大字體標題（如：101班 報名統計表），方便列印。"],
    ["第 2 列 (表頭)", "系統自動生成", "自動組合：[基本資料欄位] + 時間 + [題目1, 題目2...]"],
    ["第 3 列起 (學生列)", "預載學生資料", "學生線上報名後，對應列的「時間」與「選擇答案/金額」會即時填入。"],
    ["最下方 (合計列)", "自動產生 =SUM 公式", "若題目包含「總計」或「金額」，底部會自動建立 SUM 公式即時加總。"]
  ];

  sheet.getRange(1, 1, docsData.length, 3).setValues(docsData);

  // 格式美化
  sheet.getRange("A1").setFontSize(14).setFontWeight("bold").setFontColor("#1a73e8");
  
  const sectionRows = [3, 10, 17, 24];
  sectionRows.forEach(row => {
    sheet.getRange(row, 1, 1, 3).merge()
      .setFontWeight("bold")
      .setFontSize(11)
      .setBackground("#e8f0fe")
      .setFontColor("#1967d2");
  });

  const headerRows = [4, 11, 18, 25];
  headerRows.forEach(row => {
    sheet.getRange(row, 1, 1, 3)
      .setFontWeight("bold")
      .setBackground("#f1f3f4")
      .setBorder(true, true, true, true, true, true);
  });

  // 設定框線與自動換行
  const allDataRange = sheet.getRange(1, 1, docsData.length, 3);
  allDataRange.setWrap(true).setVerticalAlignment("middle");

  SpreadsheetApp.getUi().alert("✅『📜 使用說明』分頁已順利建立完成！已放置於試算表第一頁。");
}
// 開啟收據列印設定視窗
function openReceiptDialog() {
  const html = HtmlService.createHtmlOutputFromFile('ReceiptDialog')
    .setWidth(900)
    .setHeight(700)
    .setTitle('🧾 班級收據批次列印');
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

// 3. 取得班級選單（包含「全部班級」選項）
function getClassListForReceipt() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const ignoreSheets = ["表單回復", "原始名單", "導師名單", "📜 使用說明"];
  
  const classList = ["ALL"]; // ALL 代表全校班級
  sheets.forEach(sheet => {
    const name = sheet.getName();
    if (!ignoreSheets.includes(name)) {
      classList.push(name);
    }
  });
  return classList;
}

// 從儲存格內容提取金額（支援純數字與「項目(金額)」格式）
function extractAmountValue_(val) {
  if (val === "" || val === null || val === undefined) return 0;
  if (typeof val === "number") return isNaN(val) ? 0 : val;
  const str = String(val).trim();
  if (!str) return 0;
  const num = Number(str);
  if (!isNaN(num)) return num;
  return parseNumber(str);
}

// 從「項目(金額)」擷取價格前的名稱，若無則使用欄位標題
function parseReceiptItemName_(val, fallbackTitle) {
  const str = String(val === null || val === undefined ? "" : val).trim();
  if (!str) return fallbackTitle || "報名項目";
  if (!isNaN(Number(str))) return fallbackTitle || "報名項目";
  const cleaned = str.replace(/\(\s*\d+\s*\)/g, "").trim();
  return cleaned || fallbackTitle || "報名項目";
}

// 計算單筆報名紀錄的應繳金額與細項清單
function calculateReceiptDetail_(row, headers, types) {
  let totalColIdx = -1;
  let hasTypedAmountCol = false;

  for (let i = 5; i < headers.length; i++) {
    const qTitle = String(headers[i] || "").trim();
    const qType = String((types && types[i]) || "").trim();
    if (totalColIdx === -1 && isTotalCol({ title: qTitle, type: qType })) {
      totalColIdx = i;
    }
    if (qType.indexOf("金額") !== -1) {
      hasTypedAmountCol = true;
    }
  }

  const items = [];

  if (hasTypedAmountCol) {
    let totalAmount = 0;
    for (let i = 5; i < headers.length; i++) {
      const qTitle = String(headers[i] || "").trim();
      const qType = String((types && types[i]) || "").trim();
      const cellVal = row[i];

      if (qType.indexOf("單選(金額)") !== -1 || qType.indexOf("單選:金額") !== -1) {
        const basePrice = extractAmountValue_(cellVal);
        let finalPrice = basePrice;
        let itemName = parseReceiptItemName_(cellVal, qTitle);

        if (i + 1 < headers.length) {
          const nextType = String((types && types[i + 1]) || "").trim();
          if (nextType.indexOf("單選(折扣)") !== -1 || nextType.indexOf("單選:折扣") !== -1) {
            const discountRaw = row[i + 1];
            const discountRate = parseDiscount(discountRaw);
            finalPrice = basePrice * discountRate;
            const discountLabel = String(discountRaw || "").replace(/\(\s*\d+\s*%\s*\)/g, "").trim();
            if (discountLabel && discountRate !== 1) {
              itemName = `${itemName}（${discountLabel}）`;
            }
          }
        }

        if (finalPrice > 0) {
          items.push({ name: itemName, amount: finalPrice });
          totalAmount += finalPrice;
        }
      } else if (qType.indexOf("多選(金額)") !== -1 || qType.indexOf("多選:金額") !== -1) {
        String(cellVal || "").split(",").forEach(part => {
          const price = extractAmountValue_(part);
          if (price > 0) {
            items.push({ name: parseReceiptItemName_(part, qTitle), amount: price });
            totalAmount += price;
          }
        });
      }
    }

    if (totalColIdx !== -1) {
      const rawTotal = row[totalColIdx];
      if (rawTotal !== "" && rawTotal !== null && rawTotal !== undefined && !isNaN(Number(rawTotal))) {
        const storedTotal = Number(rawTotal);
        if (items.length === 0 && storedTotal > 0) {
          items.push({ name: String(headers[totalColIdx] || "報名費").trim(), amount: storedTotal });
        }
        return { amount: storedTotal, items: items };
      }
    }

    return { amount: totalAmount, items: items };
  }

  if (totalColIdx !== -1) {
    const rawTotal = row[totalColIdx];
    const totalVal = extractAmountValue_(rawTotal);
    if (totalVal > 0) {
      items.push({ name: parseReceiptItemName_(rawTotal, String(headers[totalColIdx] || "報名費").trim()), amount: totalVal });
    }
    return { amount: totalVal, items: items };
  }

  // 備援：依常見費用/金額關鍵字搜尋欄位並支援「項目(金額)」解析
  const fallbackKeywords = ["總計", "金額", "費用", "報名費", "小計", "費"];
  for (let k = 0; k < fallbackKeywords.length; k++) {
    const kw = fallbackKeywords[k];
    for (let i = 0; i < headers.length; i++) {
      const h = String(headers[i] || "").trim();
      if (h.indexOf(kw) !== -1) {
        const val = extractAmountValue_(row[i]);
        if (val > 0) {
          items.push({ name: parseReceiptItemName_(row[i], h), amount: val });
        }
        return { amount: val, items: items };
      }
    }
  }

  return { amount: 0, items: items };
}

// 取得已儲存的單位印章（從 DocumentProperties 重組分段資料）
function getReceiptStamp_() {
  const props = PropertiesService.getDocumentProperties();
  const count = parseInt(props.getProperty("RECEIPT_STAMP_CHUNKS") || "0", 10);
  if (!count || count <= 0) return "";
  let dataUrl = "";
  for (let i = 0; i < count; i++) {
    const chunk = props.getProperty("RECEIPT_STAMP_" + i);
    if (chunk === null) return "";
    dataUrl += chunk;
  }
  return dataUrl;
}

// 清除已儲存的單位印章
function clearReceiptStamp() {
  const props = PropertiesService.getDocumentProperties();
  const count = parseInt(props.getProperty("RECEIPT_STAMP_CHUNKS") || "0", 10);
  for (let i = 0; i < count; i++) {
    props.deleteProperty("RECEIPT_STAMP_" + i);
  }
  props.deleteProperty("RECEIPT_STAMP_CHUNKS");
  return { success: true };
}

// 儲存單位印章（依 8000 字元分段寫入 DocumentProperties，避開單一屬性 9KB 上限）
function saveReceiptStamp(dataUrl) {
  if (typeof dataUrl !== "string" || dataUrl.indexOf("data:image/") !== 0) {
    return { success: false, message: "請上傳有效的圖片格式（PNG / JPG）！" };
  }
  if (dataUrl.length > 250000) {
    return { success: false, message: "印章圖片檔案過大，請選擇較小的圖片！" };
  }

  clearReceiptStamp();
  const props = PropertiesService.getDocumentProperties();
  const chunkSize = 8000;
  const count = Math.ceil(dataUrl.length / chunkSize);
  for (let i = 0; i < count; i++) {
    props.setProperty("RECEIPT_STAMP_" + i, dataUrl.slice(i * chunkSize, (i + 1) * chunkSize));
  }
  props.setProperty("RECEIPT_STAMP_CHUNKS", String(count));
  return { success: true };
}

// 收據視窗初始化：一次取得班級清單、試算表檔名與已儲存之印章
function getReceiptInitData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return {
    classes: getClassListForReceipt(),
    docName: ss.getName().trim(),
    stampDataUrl: getReceiptStamp_()
  };
}

// 4. 收據資料讀取邏輯 (直接讀取「表單回復」、依學號去重保留最新一筆，並按班級、座號排序)
function getReceiptDataAll(selectedClass) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const docName = ss.getName().trim();
  const respSheet = ss.getSheetByName("表單回復");

  if (!respSheet || respSheet.getLastRow() < 5) {
    return { error: "目前『表單回復』無任何報名資料！" };
  }

  const lastCol = respSheet.getLastColumn();
  const types = respSheet.getRange(2, 1, 1, lastCol).getValues()[0];
  const headers = respSheet.getRange(4, 1, 1, lastCol).getValues()[0];
  const rawData = respSheet.getRange(5, 1, respSheet.getLastRow() - 4, lastCol).getValues();

  // 依學號去重，保留時間最新的一筆（若時間無法比較則以後列為準）
  const dedupMap = new Map();

  rawData.forEach((row, idx) => {
    const className = String(row[0]).trim();
    const seatNo = parseInt(row[1], 10) || 0;
    const studentId = String(row[2]).trim();
    const name = String(row[3]).trim();
    const time = row[4];

    if (!className || !name || !time) return;
    if (selectedClass !== "ALL" && selectedClass !== className) return;

    const ts = (time instanceof Date) ? time.getTime() : new Date(String(time)).getTime();
    const validTs = isNaN(ts) ? 0 : ts;
    const key = studentId || `${className}__${seatNo}__${name}`;
    const prev = dedupMap.get(key);

    if (!prev || validTs > prev.ts || (validTs === prev.ts && idx > prev.idx)) {
      const detail = calculateReceiptDetail_(row, headers, types);
      dedupMap.set(key, {
        ts: validTs,
        idx: idx,
        receipt: {
          className: className,
          seatNo: seatNo,
          studentId: studentId,
          name: name,
          items: detail.items,
          amount: detail.amount
        }
      });
    }
  });

  const receipts = Array.from(dedupMap.values()).map(entry => entry.receipt);

  if (receipts.length === 0) {
    return { error: "找不到符合條件的報名收據資料！" };
  }

  // 依照「班級」排序，再依「座號」排序
  receipts.sort((a, b) => {
    if (a.className !== b.className) {
      return a.className.localeCompare(b.className, 'zh-Hant', { numeric: true });
    }
    return a.seatNo - b.seatNo;
  });

  return { docName: docName, receipts: receipts };
}