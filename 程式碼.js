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

// 檢查班級分頁是否有學生報名／填寫資料（排除合計列）
function hasClassSubmission_(classSheet, headerInfo, rawProfileHeaders) {
  const lastRow = classSheet.getLastRow();
  const lastCol = classSheet.getLastColumn();
  if (lastRow <= headerInfo.rowIndex || lastCol === 0) return false;

  const trimmedHeaders = headerInfo.headers.map(h => String(h || "").trim());
  const timeColIdx = findTimeColIndex_(trimmedHeaders);

  const checkColIndices = [];
  if (timeColIdx !== -1) {
    // 檢查「時間」欄及其後的題目欄位（若任一有填寫即視為有報名資料）
    for (let c = timeColIdx; c < trimmedHeaders.length; c++) {
      if (trimmedHeaders[c] !== "" || c === timeColIdx) {
        checkColIndices.push(c);
      }
    }
  } else {
    // 若找不到「時間」欄，則檢查基本資料欄位以外的所有欄位
    const baseSet = new Set(["班級", "座號", "學號", "姓名", "生日", "性別", "身分證字號"]);
    (rawProfileHeaders || []).forEach(h => {
      const t = String(h || "").trim();
      if (t) baseSet.add(t);
    });
    for (let c = 0; c < trimmedHeaders.length; c++) {
      if (trimmedHeaders[c] && !baseSet.has(trimmedHeaders[c])) {
        checkColIndices.push(c);
      }
    }
  }

  if (checkColIndices.length === 0) return false;

  const dataRowCount = lastRow - headerInfo.rowIndex;
  const rows = classSheet.getRange(headerInfo.rowIndex + 1, 1, dataRowCount, lastCol).getValues();

  for (let r = 0; r < rows.length; r++) {
    const row = rows[r];
    // 排除底部的「合計」列，避免 =SUM 公式結果誤判
    const leadCells = row.slice(0, Math.min(3, row.length)).map(v => String(v || "").trim());
    if (leadCells.indexOf("合計") !== -1) continue;

    for (let k = 0; k < checkColIndices.length; k++) {
      const val = row[checkColIndices[k]];
      if (val instanceof Date) return true;
      if (String(val === null || val === undefined ? "" : val).trim() !== "") {
        return true;
      }
    }
  }

  return false;
}

// 批次匯出各班獨立試算表並 Email 寄給導師 (適應新版第 2 列表頭結構，並相容班級名稱與表頭位置差異)
function exportAndEmailSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const docName = ss.getName();
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
    '📧 批次匯出與寄信確認',
    '系統將會檢查各班報名狀況：\n1. 有學生報名：自動建立獨立試算表並寄信給導師。\n2. 尚無學生報名：自動跳過不寄信。\n\n請問是否確定開始執行？',
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) return;

  const teacherHeaders = teacherData[0].map(h => String(h || "").trim());
  const tClassIdx = teacherHeaders.indexOf("班級") !== -1 ? teacherHeaders.indexOf("班級") : 0;
  const tNameIdx = teacherHeaders.indexOf("導師") !== -1 ? teacherHeaders.indexOf("導師") : 1;
  let tEmailIdx = teacherHeaders.findIndex(h => h.toLowerCase().indexOf("email") !== -1 || h.indexOf("信箱") !== -1 || h.indexOf("郵件") !== -1);
  if (tEmailIdx === -1) tEmailIdx = 2;

  const rawSheet = ss.getSheetByName("原始名單");
  const rawProfileHeaders = (rawSheet && rawSheet.getLastRow() >= 1 && rawSheet.getLastColumn() >= 2)
    ? rawSheet.getRange(1, 2, 1, rawSheet.getLastColumn() - 1).getValues()[0]
    : [];

  let successCount = 0;
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

    const hasSubmission = hasClassSubmission_(classSheet, headerInfo, rawProfileHeaders);
    if (!hasSubmission) {
      logs.push(`⏭️ 跳過 ${classLabel} (${teacherName})：目前尚無學生完成報名`);
      skipCount++;
      continue;
    }

    try {
      const newFileName = `${docName} - ${classLabel} 報名統計`;
      const newSs = SpreadsheetApp.create(newFileName);
      const defaultSheet = newSs.getSheets()[0];

      const copiedSheet = classSheet.copyTo(newSs);
      if (defaultSheet && defaultSheet.getSheetId() !== copiedSheet.getSheetId()) {
        newSs.deleteSheet(defaultSheet);
      }
      copiedSheet.setName(classSheet.getName());

      newSs.addViewer(email);
      const fileUrl = newSs.getUrl();

      const subject = `【報名統計通知】${docName} - ${classLabel} 學生報名結果`;
      const body = `${teacherName} 老師您好：\n\n` +
                   `這是貴班（${classLabel}）在「${docName}」活動的學生最新報名統計表：\n\n` +
                   `🔗 統計表連結：\n${fileUrl}\n\n` +
                   `此信件為系統自動發送，點擊連結即可即時查看學生報名狀況。`;

      MailApp.sendEmail(email, subject, body);
      successCount++;
    } catch (e) {
      logs.push(`❌ ${classLabel} (${teacherName}) 處理失敗：${e.message}`);
      failCount++;
    }
  }

  let resultMsg = `✅ 處理完成！\n成功寄出：${successCount} 封\n自動跳過（無人報名/無名單）：${skipCount} 班`;
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
// 2. 儲存/更新報名資料 (避免重複寫入)
function saveRegistration(formData) {
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

  return { success: true, isUpdate: existingRespRow !== -1 };
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

// 4. 收據資料讀取邏輯 (直接讀取「表單回復」並按班級、座號排序)
function getReceiptDataAll(selectedClass, customItemTitle) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const respSheet = ss.getSheetByName("表單回復");

  if (!respSheet || respSheet.getLastRow() < 5) {
    return { error: "目前『表單回復』無任何報名資料！" };
  }

  const lastCol = respSheet.getLastColumn();
  const headers = respSheet.getRange(4, 1, 1, lastCol).getValues()[0];
  const rawData = respSheet.getRange(5, 1, respSheet.getLastRow() - 4, lastCol).getValues();

  // 自動尋找「金額」、「總計」、「費用」欄位
  let amountColIdx = -1;
  for (let i = 0; i < headers.length; i++) {
    const h = String(headers[i] || "");
    if (h.includes("金額") || h.includes("總計") || h.includes("費用") || h.includes("小計")) {
      amountColIdx = i;
      break;
    }
  }

  let receipts = [];

  rawData.forEach(row => {
    const className = String(row[0]).trim();
    const seatNo = parseInt(row[1], 10) || 0;
    const studentId = String(row[2]).trim();
    const name = String(row[3]).trim();
    const time = row[4];
    
    // 金額提取
    let amount = 0;
    if (amountColIdx !== -1) {
      amount = Number(row[amountColIdx]) || 0;
    }

    if (className && name && time) {
      // 依選取班級過濾
      if (selectedClass === "ALL" || selectedClass === className) {
        receipts.push({
          className: className,
          seatNo: seatNo,
          studentId: studentId,
          name: name,
          itemTitle: customItemTitle || "活動報名費用",
          amount: amount
        });
      }
    }
  });

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

  return { receipts: receipts };
}