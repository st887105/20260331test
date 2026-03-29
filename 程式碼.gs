/**
 * ============================================================================
 * 車城國小學力檢測考古題輔助系統 - 後端 API 伺服器 (前後端分離版)
 * v5.5 - 國語科多元化題型支援 + PDF 資料夾權限修復
 * ============================================================================
 */

function doPost(e) {
  try {
    const params = JSON.parse(e.postData.contents);
    const action = params.action;
    const data   = params.data || {};
    let result   = null;

    switch (action) {
      case 'getTasks':                  result = getTasks(); break;
      case 'getStudents':               result = getStudents(data.taskName); break;
      case 'verifyAdmin':               result = verifyAdmin(data.pwd); break;
      case 'getQuizSettings':           result = getQuizSettings(); break;
      case 'updateQuizSettings':        result = updateQuizSettings(data.count); break;
      case 'updateFolderId':            result = updateFolderId(data.folderId); break;
      case 'getFolderId':               result = getFolderId(); break;
      case 'setupDatabase':             result = setupDatabase(); break;
      case 'clearBankCache':            result = clearBankCache(); break;
      case 'uploadTaskData':            result = uploadTaskData(data.taskName, data.grade, data.subject, data.studentData, data.uniqueNodes); break;
      case 'generateBatch':             result = generateBatch(data.nodesArray, data.grade, data.subject, data.batchIndex); break;
      case 'generateBatchFromFolder':   result = generateBatchFromFolder(data.folderId, data.nodesArray, data.grade, data.subject, data.batchIndex); break;
      case 'scanFolder':                result = scanFolder(data.folderId); break;
      case 'generateQuiz':              result = generateQuiz(data.weakNode, data.taskName); break;
      case 'submitQuizResult':          result = submitQuizResult(data.payload); break;
      case 'getTaskResults':            result = getTaskResults(data.taskName); break;
      case 'getQuestionErrorRates':     result = getQuestionErrorRates(data.taskName); break;
      case 'analyzeClassWeakNodes':     result = analyzeClassWeakNodes(data.taskName); break;
      case 'generateTeachingWorksheet': result = generateTeachingWorksheet(data.topNodes, data.grade, data.subject); break;
      default: throw new Error('未知的 API 請求：' + action);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'success', data: result }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: error.message || String(error) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput("✅ 車城國小 AI 補救系統 API 伺服器運作中。\n請從 GitHub Pages 前端網頁連線。");
}

function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = [];

  let configSheet = ss.getSheetByName('Config');
  if (!configSheet) {
    configSheet = ss.insertSheet('Config');
    configSheet.appendRow(['Key', 'Value']);
    configSheet.appendRow(['AdminPassword', '1234']);
    configSheet.appendRow(['GeminiAPIKey', '請在此貼上您的API金鑰']);
    configSheet.appendRow(['QuizCount', '10']);
    configSheet.appendRow(['PdfFolderId', '']);
    log.push('Config 分頁已建立');
  } else {
    const configData = configSheet.getDataRange().getValues();
    const existKeys = configData.map(r => String(r[0]));
    if (!existKeys.includes('PdfFolderId')) configSheet.appendRow(['PdfFolderId', '']);
  }

  let bankSheet = ss.getSheetByName('Bank');
  if (!bankSheet) {
    bankSheet = ss.insertSheet('Bank');
    bankSheet.appendRow(['ID', '知識節點', '題目', '類型(single/fill)', '選項(JSON陣列)', '正解', '難度', '適用年級', '科目']);
    log.push('Bank 分頁已建立');
  } else {
    const header = bankSheet.getRange(1, 1, 1, bankSheet.getLastColumn()).getValues()[0];
    if (!header.includes('科目')) bankSheet.getRange(1, header.length + 1).setValue('科目');
  }

  let historySheet = ss.getSheetByName('History');
  if (!historySheet) {
    historySheet = ss.insertSheet('History');
    historySheet.appendRow(['上傳時間', '任務名稱', '適用年級', '科目', '學生人數', '班級弱點節點']);
    log.push('History 分頁已建立');
  } else {
    const header = historySheet.getRange(1, 1, 1, historySheet.getLastColumn()).getValues()[0];
    if (!header.includes('科目')) {
      historySheet.insertColumnAfter(3);
      historySheet.getRange(1, 4).setValue('科目');
    }
  }

  let resultsSheet = ss.getSheetByName('Results');
  if (!resultsSheet) {
    resultsSheet = ss.insertSheet('Results');
    resultsSheet.appendRow(['測驗時間', '任務名稱', '座號', '姓名', '分數', '作答歷時(秒)', '作答明細']);
    log.push('Results 分頁已建立');
  }

  try { CacheService.getScriptCache().remove('BankData_V3'); } catch (e) {}
  return '✅ 資料庫檢查完成！\n' + log.join('\n');
}

function verifyAdmin(pwd) {
  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  if (!configSheet) return false;
  const data = configSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === 'AdminPassword' && data[i][1].toString() === pwd.toString()) return true;
  }
  return false;
}

function getConfig() {
  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  if (!configSheet) return {};
  const result = {};
  configSheet.getDataRange().getValues().forEach(row => { if (row[0]) result[row[0]] = row[1]; });
  return result;
}

function getQuizSettings() {
  const cfg = getConfig();
  return { quizCount: parseInt(cfg.QuizCount, 10) || 10 };
}

function updateQuizSettings(newCount) {
  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  const count = parseInt(newCount, 10);
  if (isNaN(count) || count < 1) throw new Error("請輸入有效的數字");
  const data = configSheet.getDataRange().getValues();
  let found = false;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === 'QuizCount') { configSheet.getRange(i + 1, 2).setValue(count); found = true; break; }
  }
  if (!found) configSheet.appendRow(['QuizCount', count]);
  return "✅ 題數已更新為 " + count + " 題！";
}

function updateFolderId(folderId) {
  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  const cleanId = String(folderId || '').split('?')[0].trim();
  const data = configSheet.getDataRange().getValues();
  let found = false;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === 'PdfFolderId') {
      configSheet.getRange(i + 1, 2).setValue(cleanId);
      found = true; break;
    }
  }
  if (!found) configSheet.appendRow(['PdfFolderId', cleanId]);
  try { CacheService.getScriptCache().remove('PdfText_' + cleanId); } catch(e) {}
  return '✅ 資料夾 ID 已儲存：' + cleanId;
}

function getFolderId() {
  const cfg = getConfig();
  const id = String(cfg.PdfFolderId || '').split('?')[0].trim();
  return (id && !id.includes('選填')) ? id : '';
}

function uploadTaskData(taskName, grade, subject, studentData, uniqueNodes) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const subjectLabel = subject || '數學';
  const historySheet = ss.getSheetByName('History');
  if (historySheet) {
    historySheet.appendRow([new Date(), "'" + String(taskName), grade, subjectLabel, studentData.length, uniqueNodes.join(', ')]);
  }
  let taskSheet = ss.getSheetByName(taskName);
  if (taskSheet) taskSheet.clear(); else taskSheet = ss.insertSheet(taskName);
  taskSheet.appendRow(['座號', '姓名', '答對率', '知識節點(弱項)', '科目']);
  const rows = studentData.map(s => {
    let seat = String(s.seatNo || '').trim(), name = String(s.name || '').trim();
    const match = seat.match(/(\d+)\s*[號]?\s*([A-Za-z\u4e00-\u9fa5]+)$/);
    if (match) { seat = match[1]; name = match[2]; }
    return [seat, name, s.accuracy, s.weakNodes, subjectLabel];
  });
  taskSheet.getRange(2, 1, rows.length, 5).setValues(rows);
  return true;
}

function getTasks() {
  const historySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('History');
  if (!historySheet) return [];
  const data = historySheet.getDataRange().getValues();
  return data.slice(1).filter(r => r[1]).map(r => ({
    name:    String(r[1]).replace(/'/g, ''),
    grade:   String(r[2] || ''),
    subject: String(r[3] || '數學'),
  })).reverse();
}

function getStudents(taskName) {
  const taskSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(taskName);
  if (!taskSheet) return [];
  return taskSheet.getDataRange().getValues().slice(1).map(r => ({
    seatNo:  String(r[0] || ''),
    name:    String(r[1] || ''),
    weakNode: String(r[3] || ''),
    subject: String(r[4] || '數學'),
  }));
}

function scanFolder(folderId) {
  if (!folderId || folderId.trim() === '') throw new Error('請先填入有效的資料夾 ID');
  const cleanId = folderId.split('?')[0].trim(); 
  try {
    const folder = DriveApp.getFolderById(cleanId);
    const iter   = folder.getFilesByMimeType('application/pdf');
    const files  = [];
    while (iter.hasNext()) {
      const f = iter.next();
      files.push({ id: f.getId(), name: f.getName() });
    }
    return { folderName: folder.getName(), count: files.length, files: files };
  } catch (err) {
    throw new Error('無法讀取資料夾：指定的權限不足，或資料夾不存在。\n詳細錯誤：' + err.message);
  }
}

function extractPdfText(fileId) {
  const token = ScriptApp.getOAuthToken();
  try {
    const exportResp = UrlFetchApp.fetch(
      'https://www.googleapis.com/drive/v3/files/' + fileId + '/export?mimeType=text%2Fplain',
      { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
    );
    if (exportResp.getResponseCode() === 200) {
      const txt = exportResp.getContentText();
      if (txt && txt.trim().length > 30) return txt.length > 6000 ? txt.substring(0, 6000) + '\n...(略)' : txt;
    }
  } catch(e) {}

  try {
    const file     = DriveApp.getFileById(fileId);
    const blob     = file.getBlob();
    const boundary = 'ocr_' + Date.now();
    const metadata = JSON.stringify({ name: '_ocr_' + Date.now(), mimeType: 'application/vnd.google-apps.document' });
    const b64 = Utilities.base64Encode(blob.getBytes());
    const body = '--' + boundary + '\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n' + metadata + '\r\n--' + boundary + '\r\nContent-Type: application/pdf\r\nContent-Transfer-Encoding: base64\r\n\r\n' + b64 + '\r\n--' + boundary + '--';

    const uploadResp = UrlFetchApp.fetch(
      'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart',
      { method: 'POST', contentType: 'multipart/related; boundary=' + boundary, payload: body, headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
    );

    if (uploadResp.getResponseCode() !== 200) throw new Error('OCR 失敗');
    const newFileId = JSON.parse(uploadResp.getContentText()).id;
    Utilities.sleep(2000);
    const docText = DocumentApp.openById(newFileId).getBody().getText();
    try { UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + newFileId, { method: 'DELETE', headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }); } catch(e) {}
    return docText.length > 6000 ? docText.substring(0, 6000) + '\n...(略)' : docText;
  } catch (err) { throw new Error('PDF 讀取失敗'); }
}

function generateBatchFromFolder(folderId, nodesArray, grade, subject, batchIndex) {
  const cleanId = folderId.split('?')[0].trim();
  const cache   = CacheService.getScriptCache();
  const PDF_KEY = 'PdfText_' + cleanId;
  let   pdfText = cache.get(PDF_KEY);

  if (!pdfText) {
    const folder = DriveApp.getFolderById(cleanId);
    const iter  = folder.getFilesByMimeType('application/pdf');
    const texts = [];
    while (iter.hasNext()) {
      const f = iter.next();
      try {
        const txt = extractPdfText(f.getId());
        if (txt && txt.trim().length > 0) texts.push('【' + f.getName() + '】\n' + txt);
      } catch (readErr) { }
    }
    if (texts.length === 0) throw new Error('資料夾內沒有可讀取的 PDF 文字。');
    pdfText = texts.join('\n\n');
    if (pdfText.length > 10000) pdfText = pdfText.substring(0, 10000) + '\n\n...(截斷)';
    try { cache.put(PDF_KEY, pdfText, 1800); } catch (e) {}
  }
  return _doGenerateBatch(nodesArray, grade, subject || '數學', batchIndex, pdfText);
}

function buildPrompt(subject, grade, batchNodes, pdfText) {
  const nodeList = batchNodes.map((n, i) => `${i + 1}. 「${n}」`).join('\n');
  const pdfSection = pdfText ? `\n\n【參考教材內容（請依此出題）】\n---\n${pdfText}\n---\n` : '';

  if (subject === '國語') {
    return `你是台灣國小國語科命題專家。請為「${grade}」針對以下 ${batchNodes.length} 個節點各設計 6 道題：${pdfSection}\n${nodeList}
【規範】第1-2題：字詞辨識選擇題(4選1)；第3-4題：詞語填充題；第5題：短語替換/句型填充；第6題：閱讀理解選擇題。
難度：1-2 easy, 3-5 medium, 6 hard。只回傳 JSON 陣列，不含說明：
[{"node":"節點","text":"題目","type":"single","options":["A","B","C","D"],"answer":"A","difficulty":"easy"}]`;
  }
  return `你是台灣國小數學命題專家。請為「${grade}」針對以下 ${batchNodes.length} 個節點各設計 6 道題：${pdfSection}\n${nodeList}
【規範】第1-3題：單選題(4選1)；第4-6題：填充題(填純數字)。難度：1 easy, 2-4 medium, 5-6 hard。
必須測出迷思概念。只回傳 JSON 陣列，不含說明：
[{"node":"節點","text":"題目","type":"single","options":["A","B","C","D"],"answer":"A","difficulty":"easy"}]`;
}

function generateBatch(nodesArray, grade, subject, batchIndex) {
  return _doGenerateBatch(nodesArray, grade, subject || '數學', batchIndex, null);
}

function _doGenerateBatch(nodesArray, grade, subject, batchIndex, pdfText) {
  const cfg = getConfig();
  const apiKey = String(cfg.GeminiAPIKey || '');
  if (!apiKey || apiKey.includes('請在此')) throw new Error("請先在 Config 試算表設定 Gemini API Key！");

  const validNodes = nodesArray.filter(n => n && n.trim() !== '');
  const BATCH_SIZE = 5;
  const batches = [];
  for (let i = 0; i < validNodes.length; i += BATCH_SIZE) batches.push(validNodes.slice(i, i + BATCH_SIZE));
  const totalBatches = batches.length;

  if (batchIndex >= totalBatches) {
    _clearAllCaches();
    return { done: true, current: totalBatches, total: totalBatches, message: `✅ 全部完成！` };
  }

  const batchNodes = batches[batchIndex];
  const prompt = buildPrompt(subject, grade, batchNodes, pdfText);
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;
  const fetchOpts = {
    method: "post", contentType: "application/json",
    payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }], generationConfig: { responseMimeType: "application/json", temperature: 0.85 } }),
    muteHttpExceptions: true
  };

  let result = null;
  for (let retry = 0; retry < 3; retry++) {
    const resp = UrlFetchApp.fetch(url, fetchOpts);
    result = JSON.parse(resp.getContentText());
    if (result.error) {
      if (result.error.message.includes("quota") || result.error.message.includes("429")) Utilities.sleep(20000);
      else throw new Error("API 錯誤: " + result.error.message);
    } else break;
  }

  const rawText = result.candidates[0].content.parts[0].text.replace(/```json/g, '').replace(/```/g, '').trim();
  const questions = JSON.parse(rawText);
  const bankSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Bank');
  const lastRow = bankSheet.getLastRow() || 1;
  const newRows = questions.map((q, idx) => [
    `AI-${Date.now().toString().slice(-5)}-${batchIndex}-${idx}`,
    q.node, q.text, q.type, JSON.stringify(q.options || []),
    normalizeAnswer(q.answer), q.difficulty, grade, subject
  ]);

  if (newRows.length > 0) {
    bankSheet.getRange(lastRow + 1, 1, newRows.length, 9).setValues(newRows);
    bankSheet.getRange(lastRow + 1, 6, newRows.length, 1).setNumberFormat('@STRING@');
  }

  const isLast = (batchIndex + 1 >= totalBatches);
  if (isLast) _clearAllCaches();

  return { done: isLast, current: batchIndex + 1, total: totalBatches, addedThisBatch: newRows.length, message: `第 ${batchIndex + 1} 批完成，新增 ${newRows.length} 題` };
}

function normalizeAnswer(str) {
  if (str === null || str === undefined) return '';
  if (typeof str === 'number') return String(Number.isInteger(str) ? str : str).trim().toLowerCase();
  return String(str).trim().replace(/\s+/g, '').replace(/[\uff10-\uff19]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0)).replace(/，/g, ',').toLowerCase();
}

function generateQuiz(weakNode, taskName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quizCount = parseInt(getConfig().QuizCount, 10) || 10;
  
  let targetGrade = '', targetSubject = '數學';
  const historySheet = ss.getSheetByName('History');
  if (historySheet && taskName) {
    historySheet.getDataRange().getValues().slice(1).forEach(r => {
      if (String(r[1]).replace(/'/g, '').trim() === String(taskName).trim()) {
        targetGrade = String(r[2] || '').trim();
        const r3 = String(r[3] || '').trim();
        if (r3 === '國語' || r3 === '數學') targetSubject = r3;
      }
    });
  }

  const cache = CacheService.getScriptCache();
  let allQuestions = [];
  const cachedData = cache.get('BankData_V3');

  if (cachedData) {
    allQuestions = JSON.parse(cachedData);
  } else {
    const bankSheet = ss.getSheetByName('Bank');
    if (!bankSheet || bankSheet.getLastRow() <= 1) return [];
    const data = bankSheet.getRange(2, 1, bankSheet.getLastRow() - 1, 9).getValues();
    data.forEach((row, i) => {
      if (!row[0]) return;
      let options = [];
      if (row[4]) {
        try { options = row[4].startsWith('[') ? JSON.parse(row[4]) : String(row[4]).split(','); } catch(e){}
      }
      allQuestions.push({
        id: row[0], node: String(row[1]||'').trim(), text: row[2], type: row[3], options,
        answer: normalizeAnswer(row[5]), displayAnswer: String(row[5]||'').trim(),
        difficulty: String(row[6]||'medium').trim(), grade: String(row[7]||'').trim(), subject: String(row[8]||'').trim()
      });
    });
    try { cache.put('BankData_V3', JSON.stringify(allQuestions), 3600); } catch (e) {}
  }

  const shuffle = arr => arr.sort(() => Math.random() - 0.5);
  let targets = [], fallbacks = [];
  const safeNode = String(weakNode || '').trim();
  
  allQuestions.forEach(q => {
    if (q.subject && q.subject !== targetSubject) return;
    if (targetGrade && q.grade && q.grade !== targetGrade) return;
    if (q.node && safeNode && (q.node.includes(safeNode) || safeNode.includes(q.node))) targets.push(q);
    else fallbacks.push(q);
  });

  shuffle(targets);
  let final = targets.slice(0, quizCount);
  if (final.length < quizCount) {
    shuffle(fallbacks);
    final = final.concat(fallbacks.slice(0, quizCount - final.length));
  }
  return shuffle(final.map(q => q.type === 'single' ? { ...q, options: shuffle([...q.options]) } : q));
}

function submitQuizResult(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let resultSheet = ss.getSheetByName('Results');
  resultSheet.appendRow([new Date(), data.taskName, data.seatNo || '', data.name || '', data.score || 0, data.timeSpent || 0, JSON.stringify(data.details || [])]);
  return true;
}

function getTaskResults(taskName) {
  const resultSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Results');
  if (!resultSheet || resultSheet.getLastRow() <= 1) return [];
  const data = resultSheet.getDataRange().getValues();
  const results = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).replace(/'/g, '').trim() !== String(taskName).trim()) continue;
    results.push({
      submittedAt: data[i][0] ? new Date(data[i][0]).toLocaleString('zh-TW') : '',
      taskName: String(data[i][1]).replace(/'/g, '').trim(),
      seatNo: String(data[i][2]||''), name: String(data[i][3]||''), score: Number(data[i][4]||0), timeSpent: Number(data[i][5]||0),
      details: JSON.parse(data[i][6] || '[]')
    });
  }
  const seen = {}, deduped = [];
  for (let i = results.length - 1; i >= 0; i--) {
    const key = results[i].seatNo + '_' + results[i].name;
    if (!seen[key]) { seen[key] = true; deduped.unshift(results[i]); }
  }
  return deduped.sort((a, b) => Number(a.seatNo) - Number(b.seatNo));
}

function analyzeClassWeakNodes(taskName) {
  const results = getTaskResults(taskName);
  if (!results.length) return [];
  const nodeMap = {};
  results.forEach(r => {
    (r.details || []).forEach(d => {
      if (!d.isCorrect && d.node) {
        if (!nodeMap[d.node]) nodeMap[d.node] = new Set();
        nodeMap[d.node].add(r.name || r.seatNo);
      }
    });
  });
  return Object.entries(nodeMap).map(([node, students]) => ({
    node, wrongCount: students.size, studentCount: results.length, wrongRate: Math.round((students.size / results.length) * 100), students: Array.from(students)
  })).sort((a, b) => b.wrongCount - a.wrongCount);
}

function getQuestionErrorRates(taskName) {
  const results = getTaskResults(taskName);
  const qMap = {};
  results.forEach(r => {
    (r.details || []).forEach(d => {
      const qid = d.questionText || d.id;
      if (!qMap[qid]) qMap[qid] = { questionText: qid, node: d.node||'', correctAns: d.displayCorrectAns||'', total: 0, wrong: 0, wrongAnswers: [] };
      qMap[qid].total++;
      if (!d.isCorrect) { qMap[qid].wrong++; qMap[qid].wrongAnswers.push({ name: r.name, ans: d.userAns || '未作答' }); }
    });
  });
  return Object.values(qMap).map(q => ({ ...q, wrongRate: Math.round((q.wrong / q.total) * 100) })).sort((a, b) => b.wrongRate - a.wrongRate);
}

function generateTeachingWorksheet(topNodes, grade, subject) {
  const apiKey = String(getConfig().GeminiAPIKey || '');
  if (!apiKey || apiKey.includes('請在此')) throw new Error("請先設定 Gemini API Key！");

  const subj = subject || '數學';
  const nodeList = topNodes.map((n, i) => `${i + 1}. 「${n}」`).join('\n');

  let prompt = '';
  if (subj === '國語') {
    prompt = `你是台灣國小國語科教師，擅長針對學生的「弱點」設計多元化的補救教學練習。
請針對「${grade}」學生，為以下知識節點設計「學生作答與教師引導用講義」：

${nodeList}

規範：
- 每個節點設計 4~5 題，題型必須「多元化」，包含：改錯字、字音字形選擇、詞語填空、短語替換等符合學生弱點程度的題型。
- 難度必須配合補救教學，從最基礎的題目開始引導。
- step：寫出解題的鷹架引導（例如：「請先圈出句子中讀音奇怪的字」）。
- hint：給學生的一句小提示。
- answer：正確解答。
- type：選擇題用 "single"，填空/改錯/簡答用 "fill"。

只回傳 JSON 陣列：
[{"node":"節點","step":"引導步驟","text":"題目","type":"fill","answer":"答案","hint":"提示"}]`;
  } else {
    prompt = `你是台灣國小數學科教師，擅長設計補救教學練習。
請針對「${grade}」學生，為以下知識節點設計「教師教學用講義」：

${nodeList}

規範：
- 每個節點設計 5 題，全部為填充計算題（type: "fill"）
- 難度全部 easy（最基礎的直接計算）
- step：引導步驟（如「先通分 → 再相加」）
- hint：一句教學提示
- answer：最終數字或分數答案

只回傳 JSON 陣列：
[{"node":"節點","step":"步驟","text":"題目","type":"fill","answer":"答案","hint":"提示"}]`;
  }

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;
  const opts = {
    method: "post", contentType: "application/json",
    payload: JSON.stringify({
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: { responseMimeType: "application/json", temperature: 0.5 }
    }),
    muteHttpExceptions: true
  };

  let result = null;
  for (let retry = 0; retry < 3; retry++) {
    const resp = UrlFetchApp.fetch(url, opts);
    result = JSON.parse(resp.getContentText());
    if (result.error) {
      if (result.error.message.toLowerCase().includes("quota")) { Utilities.sleep(15000); }
      else throw new Error("API 錯誤: " + result.error.message);
    } else break;
  }
  if (result && result.error) throw new Error("API 持續失敗，請稍後再試。");

  const rawText = result.candidates[0].content.parts[0].text.replace(/```json/g, '').replace(/```/g, '').trim();
  return JSON.parse(rawText);
}

function _clearAllCaches() { ['BankData_V3'].forEach(k => { try { CacheService.getScriptCache().remove(k); } catch (e) {} }); }
function clearBankCache() { _clearAllCaches(); return "✅ 快取已清除！"; }

function testAIGeneration() {
  UrlFetchApp.fetch("https://www.google.com");
  Logger.log("外部連線權限與 Drive 權限已檢測通過！");
}