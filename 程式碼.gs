// ============================================================================
// 車城國小學力檢測AI補救系統 - Google Apps Script 後端
// v6.1 最終版：多科目 + PDF + 成績明細 + 差異化分析修正
// ============================================================================

// ==========================================
// API 入口
// ==========================================
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
      case 'generateBatchFromFolder':   result = generateBatchFromFolder(data.folderId, data.nodesArray, data.grade, data.subject, data.batchIndex, data.mode); break;
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
  return ContentService.createTextOutput(
    '✅ 車城國小 AI 補救系統 API v6.1 運作中\n請從前端網頁連線。'
  );
}

// ==========================================
// 1. 資料庫初始化
// ==========================================
function setupDatabase() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const log = [];

  // ── Config ──
  let cfg = ss.getSheetByName('Config');
  if (!cfg) {
    cfg = ss.insertSheet('Config');
    cfg.appendRow(['Key', 'Value']);
    cfg.appendRow(['AdminPassword', '1234']);
    cfg.appendRow(['GeminiAPIKey',  '請在此貼上您的 Gemini API 金鑰']);
    cfg.appendRow(['QuizCount',     '10']);
    cfg.appendRow(['PdfFolderId',   '']);
    cfg.appendRow(['WebAppUrl',     '（部署後貼上 Web App 網址）']);
    log.push('Config 分頁已建立（含 WebAppUrl）');
  } else {
    const rows      = cfg.getDataRange().getValues();
    const existKeys = rows.map(r => String(r[0]));
    if (!existKeys.includes('PdfFolderId')) { cfg.appendRow(['PdfFolderId', '']);           log.push('補充 PdfFolderId'); }
    if (!existKeys.includes('WebAppUrl'))   { cfg.appendRow(['WebAppUrl',  '（待填入）']); log.push('補充 WebAppUrl');  }
  }

  // ── Bank ──
  let bank = ss.getSheetByName('Bank');
  if (!bank) {
    bank = ss.insertSheet('Bank');
    bank.appendRow(['ID','知識節點','題目','類型(single/fill)','選項(JSON陣列)','正解','難度','適用年級','科目']);
    log.push('Bank 分頁已建立');
  } else {
    const h = bank.getRange(1,1,1,bank.getLastColumn()).getValues()[0];
    if (!h.includes('科目')) { bank.getRange(1, h.length+1).setValue('科目'); log.push('Bank 補充「科目」欄'); }
  }

  // ── History ──
  let hist = ss.getSheetByName('History');
  if (!hist) {
    hist = ss.insertSheet('History');
    hist.appendRow(['上傳時間','任務名稱','適用年級','科目','學生人數','班級弱點節點']);
    log.push('History 分頁已建立');
  } else {
    const h = hist.getRange(1,1,1,hist.getLastColumn()).getValues()[0];
    if (!h.includes('科目')) { hist.insertColumnAfter(3); hist.getRange(1,4).setValue('科目'); log.push('History 補充「科目」欄'); }
  }

  // ── Results ──
  let res = ss.getSheetByName('Results');
  if (!res) {
    res = ss.insertSheet('Results');
    res.appendRow(['測驗時間','任務名稱','座號','姓名','分數','作答歷時(秒)','作答明細']);
    log.push('Results 分頁已建立');
  }

  try { CacheService.getScriptCache().remove('BankData_V3'); } catch(e) {}
  return '✅ 資料庫初始化完成！\n' + log.join('\n');
}

// ==========================================
// 2. 認證與設定
// ==========================================
function verifyAdmin(pwd) {
  const cfg = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  if (!cfg) return false;
  const rows = cfg.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === 'AdminPassword' && String(rows[i][1]) === String(pwd)) return true;
  }
  return false;
}

function getConfig() {
  const cfg = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  if (!cfg) return {};
  const out = {};
  cfg.getDataRange().getValues().forEach(r => { if (r[0]) out[r[0]] = r[1]; });
  return out;
}

function getQuizSettings() {
  const cfg = getConfig();
  return { quizCount: parseInt(cfg.QuizCount, 10) || 10 };
}

function updateQuizSettings(newCount) {
  const cfg   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  if (!cfg) throw new Error('找不到 Config');
  const count = parseInt(newCount, 10);
  if (isNaN(count) || count < 1) throw new Error('請輸入有效的數字');
  const rows  = cfg.getDataRange().getValues();
  let found   = false;
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === 'QuizCount') { cfg.getRange(i+1,2).setValue(count); found = true; break; }
  }
  if (!found) cfg.appendRow(['QuizCount', count]);
  return '✅ 題數已更新為 ' + count + ' 題';
}

function updateFolderId(folderId) {
  const cfg     = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  if (!cfg) throw new Error('找不到 Config');
  const cleanId = String(folderId || '').split('?')[0].trim();
  const rows    = cfg.getDataRange().getValues();
  let found     = false;
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === 'PdfFolderId') { cfg.getRange(i+1,2).setValue(cleanId); found = true; break; }
  }
  if (!found) cfg.appendRow(['PdfFolderId', cleanId]);
  try { CacheService.getScriptCache().remove('PdfText_' + cleanId); } catch(e) {}
  return '✅ 資料夾 ID 已儲存：' + cleanId;
}

function getFolderId() {
  const cfg = getConfig();
  const id  = String(cfg.PdfFolderId || '').split('?')[0].trim();
  return (id && !id.startsWith('（') && id.length > 5) ? id : '';
}

// ==========================================
// 3. 任務與學生管理
// ==========================================
function uploadTaskData(taskName, grade, subject, studentData, uniqueNodes) {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const subj = subject || '數學';
  const hist = ss.getSheetByName('History');
  if (hist) hist.appendRow([new Date(), "'" + String(taskName), grade, subj, studentData.length, uniqueNodes.join(', ')]);

  let sheet = ss.getSheetByName(taskName);
  if (sheet) sheet.clear(); else sheet = ss.insertSheet(taskName);
  sheet.appendRow(['座號','姓名','答對率','知識節點(弱項)','科目']);

  const rows = studentData.map(s => {
    let seat = String(s.seatNo || '').trim(), name = String(s.name || '').trim();
    const m  = seat.match(/(\d+)\s*[號]?\s*([A-Za-z\u4e00-\u9fa5]+)$/);
    if (m) { seat = m[1]; name = m[2]; }
    return [seat, name, s.accuracy, s.weakNodes, subj];
  });
  if (rows.length > 0) sheet.getRange(2, 1, rows.length, 5).setValues(rows);
  return true;
}

function getTasks() {
  const hist = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('History');
  if (!hist) return [];
  return hist.getDataRange().getValues().slice(1)
    .filter(r => r[1])
    .map(r => ({ name: String(r[1]).replace(/'/g,''), grade: String(r[2]||''), subject: String(r[3]||'數學') }))
    .reverse();
}

function getStudents(taskName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(taskName);
  if (!sheet) return [];
  return sheet.getDataRange().getValues().slice(1)
    .filter(r => r[1])
    .map(r => ({ seatNo: String(r[0]||''), name: String(r[1]||''), weakNode: String(r[3]||''), subject: String(r[4]||'數學') }));
}

// ==========================================
// 4. PDF 工具（修正版：支援共用雲端硬碟）
// ==========================================
function scanFolder(folderId) {
  if (!folderId || folderId.trim() === '') throw new Error('請填寫有效的資料夾 ID');
  const cleanId = folderId.split('?')[0].trim();
  const token   = ScriptApp.getOAuthToken();
  const hdrs    = { Authorization: 'Bearer ' + token };
  const muteOpt = { muteHttpExceptions: true };

  let folderName    = cleanId;
  let isSharedDrive = false;

  // 嘗試1：一般資料夾
  const metaResp = UrlFetchApp.fetch(
    `https://www.googleapis.com/drive/v3/files/${cleanId}?fields=name&supportsAllDrives=true&includeItemsFromAllDrives=true`,
    { headers: hdrs, ...muteOpt }
  );

  if (metaResp.getResponseCode() === 200) {
    folderName = JSON.parse(metaResp.getContentText()).name || cleanId;
  } else {
    // 嘗試2：Shared Drive
    const driveResp = UrlFetchApp.fetch(
      `https://www.googleapis.com/drive/v3/drives/${cleanId}`,
      { headers: hdrs, ...muteOpt }
    );
    if (driveResp.getResponseCode() === 200) {
      folderName    = JSON.parse(driveResp.getContentText()).name || cleanId;
      isSharedDrive = true;
    } else {
      throw new Error(
        `無法存取資料夾（HTTP ${metaResp.getResponseCode()}）\n` +
        `請確認：① 資料夾 ID 是否正確 ② 已共用給腳本執行帳號 ③ 若為共用雲端硬碟請確認已加入成員`
      );
    }
  }

  // 列出 PDF
  let listUrl;
  if (isSharedDrive) {
    const q = encodeURIComponent("trashed=false and mimeType='application/pdf'");
    listUrl = `https://www.googleapis.com/drive/v3/files?q=${q}&driveId=${cleanId}&corpora=drive&includeItemsFromAllDrives=true&supportsAllDrives=true&fields=files(id,name)&pageSize=100`;
  } else {
    const q = encodeURIComponent(`'${cleanId}' in parents and mimeType='application/pdf' and trashed=false`);
    listUrl = `https://www.googleapis.com/drive/v3/files?q=${q}&fields=files(id,name)&pageSize=100&supportsAllDrives=true&includeItemsFromAllDrives=true`;
  }

  const listResp = UrlFetchApp.fetch(listUrl, { headers: hdrs, ...muteOpt });
  if (listResp.getResponseCode() !== 200) throw new Error('列出 PDF 失敗：' + listResp.getContentText());
  const files = JSON.parse(listResp.getContentText()).files || [];
  return { folderName, count: files.length, files };
}

function extractPdfText(fileId) {
  const token = ScriptApp.getOAuthToken();
  const hdrs  = { Authorization: 'Bearer ' + token };

  // 方法1：直接 export 純文字（有文字層的 PDF）
  try {
    const exportResp = UrlFetchApp.fetch(
      `https://www.googleapis.com/drive/v3/files/${fileId}/export?mimeType=text%2Fplain&supportsAllDrives=true`,
      { headers: hdrs, muteHttpExceptions: true }
    );
    if (exportResp.getResponseCode() === 200) {
      const txt = exportResp.getContentText();
      if (txt && txt.trim().length > 30) {
        return txt.length > 6000 ? txt.substring(0, 6000) + '\n...(略)' : txt;
      }
    }
  } catch(e) { Logger.log('Export 失敗: ' + e.message); }

  // 方法2：REST API 下載後 OCR（掃描版 PDF）
  try {
    const dlResp = UrlFetchApp.fetch(
      `https://www.googleapis.com/drive/v3/files/${fileId}?alt=media&supportsAllDrives=true`,
      { headers: hdrs, muteHttpExceptions: true }
    );
    if (dlResp.getResponseCode() !== 200) throw new Error(`下載失敗 (HTTP ${dlResp.getResponseCode()})`);

    const pdfBytes = dlResp.getContent();
    if (pdfBytes.length > 10 * 1024 * 1024) throw new Error('PDF 超過 10MB，跳過 OCR');

    const b64      = Utilities.base64Encode(pdfBytes);
    const boundary = 'ocr_' + Date.now();
    const metadata = JSON.stringify({ name: '_ocr_tmp_' + Date.now(), mimeType: 'application/vnd.google-apps.document' });
    const body     =
      '--' + boundary + '\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n' + metadata + '\r\n' +
      '--' + boundary + '\r\nContent-Type: application/pdf\r\nContent-Transfer-Encoding: base64\r\n\r\n' + b64 +
      '\r\n--' + boundary + '--';

    const uploadResp = UrlFetchApp.fetch(
      'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart',
      { method:'POST', contentType:'multipart/related; boundary='+boundary, payload:body, headers:hdrs, muteHttpExceptions:true }
    );
    if (uploadResp.getResponseCode() !== 200) throw new Error('OCR 上傳失敗: ' + uploadResp.getContentText());

    const newFileId = JSON.parse(uploadResp.getContentText()).id;
    Utilities.sleep(3000);

    const docResp = UrlFetchApp.fetch(
      `https://www.googleapis.com/drive/v3/files/${newFileId}/export?mimeType=text%2Fplain`,
      { headers: hdrs, muteHttpExceptions: true }
    );
    const docText = docResp.getContentText();

    try { UrlFetchApp.fetch(`https://www.googleapis.com/drive/v3/files/${newFileId}`, { method:'DELETE', headers:hdrs, muteHttpExceptions:true }); } catch(e) {}

    return docText.length > 6000 ? docText.substring(0, 6000) + '\n...(略)' : docText;
  } catch(err) {
    throw new Error('PDF 讀取失敗：' + err.message);
  }
}

function generateBatchFromFolder(folderId, nodesArray, grade, subject, batchIndex, mode) {
  if (!folderId || folderId.trim() === '') throw new Error('請先設定 Google Drive 資料夾 ID');
  const cleanId    = folderId.split('?')[0].trim();
  const actualMode = mode || 'pdf';
  const token      = ScriptApp.getOAuthToken();
  const hdrs       = { Authorization: 'Bearer ' + token };
  const cache      = CacheService.getScriptCache();
  const PDF_KEY    = 'PdfText_' + cleanId;
  let   pdfText    = cache.get(PDF_KEY);

  if (!pdfText) {
    const q        = encodeURIComponent(`'${cleanId}' in parents and mimeType='application/pdf' and trashed=false`);
    const listResp = UrlFetchApp.fetch(
      `https://www.googleapis.com/drive/v3/files?q=${q}&fields=files(id,name)&pageSize=100&supportsAllDrives=true&includeItemsFromAllDrives=true`,
      { headers: hdrs, muteHttpExceptions: true }
    );
    if (listResp.getResponseCode() !== 200) throw new Error('無法列出 PDF：' + listResp.getContentText());

    const pdfFiles = JSON.parse(listResp.getContentText()).files || [];
    if (pdfFiles.length === 0) throw new Error('資料夾內沒有 PDF 檔案，請先上傳教材。');

    const texts = [];
    pdfFiles.forEach(f => {
      try {
        const txt = extractPdfText(f.id);
        if (txt && txt.trim().length > 0) { texts.push('【' + f.name + '】\n' + txt); Logger.log('已讀取：' + f.name); }
      } catch(e) { Logger.log('跳過：' + f.name + ' → ' + e.message); }
    });

    if (texts.length === 0) throw new Error('無法讀取任何 PDF。請確認：① 檔案格式 ② 腳本帳號有存取權限。');
    pdfText = texts.join('\n\n');
    if (pdfText.length > 10000) pdfText = pdfText.substring(0, 10000) + '\n\n...(截斷)';
    try { cache.put(PDF_KEY, pdfText, 1800); } catch(e) {}
  }

  return _doGenerateBatch(nodesArray, grade, subject || '數學', batchIndex, pdfText, actualMode);
}

// ==========================================
// 5. Prompt 工廠
// ==========================================
function buildPrompt(subject, grade, batchNodes, pdfText, mode) {
  const nodeList   = batchNodes.map((n, i) => `${i+1}. 「${n}」`).join('\n');
  let   pdfSection = '';
  if (pdfText) {
    pdfSection = mode === 'mixed'
      ? `\n\n【參考教材內容（優先參考，不足可補充專業知識）】\n---\n${pdfText}\n---`
      : `\n\n【教材內容（所有題目必須有教材依據，不可超出範圍）】\n---\n${pdfText}\n---`;
  }

  // ★ 降低門檻 Prompt：針對學習扶助（後10%）學生，加入 hint 鷹架欄位
  if (subject === '國語') {
    return `你是台灣資深國小國語科學習扶助命題專家，熟悉108課綱與低成就學生學習特性。
請為「${grade}」學力落在後10%的學生，針對以下 ${batchNodes.length} 個知識節點，各設計 6 道題目：${pdfSection}
${nodeList}

【題型規範（每個節點，針對低成就學生）】
- 第1-3題（very_easy）：字音字形辨識選擇題（4選1，最基礎認字，選項誘答性低）
- 第4題（easy）：詞語應用選擇題（4選1，題幹不超過15字，極度生活化情境）
- 第5題（easy）：簡單詞語填充題（答案限2字以內的高頻常用詞彙）
- 第6題（medium）：極短文閱讀選擇題（4選1，附20-30字短文，只問可直接在文中「找到」的明確事實）

【嚴格限制】
- 題幹必須使用學生能懂的基礎詞彙
- 填充題答案必須唯一、明確、高頻常用字
- 選擇題 answer 必須與 options 之一完全相同（字元完全一致）
- 每題必須提供 hint：一句話引導（如「請在文章第一句找答案喔！」）

只回傳 JSON 陣列，不含任何說明或 Markdown：
[{"node":"節點名稱","text":"題目","type":"single","options":["A","B","C","D"],"answer":"A","difficulty":"easy","hint":"引導提示"}]`;
  }

  return `你是台灣資深國小數學學習扶助命題專家，熟悉108課綱，擅長將抽象數學概念具象化，精準診斷低成就學生迷思概念。
請為「${grade}」學力落在後10%的學生，針對以下 ${batchNodes.length} 個知識節點，各設計 6 道題目：${pdfSection}
${nodeList}

【題型規範（針對低成就學生）】
- 第1-2題（very_easy）：基礎概念單選題（4選1，數字極小，單一基本概念如數的對應、圖形辨識）
- 第3題（easy）：極短生活情境單選題（4選1，題幹低於20字，單一步驟解題）
- 第4-5題（easy）：迷思概念單選題（4選1，針對該節點最常見的「單一」錯誤想法設計誘答）
- 第6題（medium）：基礎單步計算填充題（填入純數字，不進位或數字極小，建立運算信心）

【嚴格限制】
- 絕對不可出現多步驟計算或複雜文字敘述
- 選擇題 answer 必須與 options 之一完全相同（字元完全一致）
- 填充題 answer 只能是純數字或分數（如 "12"、"3/4"），不含單位或文字
- 不超出該年級範圍，鎖定最核心的基礎次技能
- 每題必須提供 hint：白話具象引導（如「想一想，3個十是多少？」）

只回傳 JSON 陣列，不含任何說明或 Markdown：
[{"node":"節點名稱","text":"題目","type":"single","options":["A","B","C","D"],"answer":"A","difficulty":"easy","hint":"引導提示"}]`;
}

// ==========================================
// 6. 批次出題引擎
// ==========================================
function generateBatch(nodesArray, grade, subject, batchIndex) {
  return _doGenerateBatch(nodesArray, grade, subject || '數學', batchIndex, null, 'ai');
}

function _doGenerateBatch(nodesArray, grade, subject, batchIndex, pdfText, mode) {
  const cfg    = getConfig();
  const apiKey = String(cfg.GeminiAPIKey || '');
  if (!apiKey || apiKey.includes('請在此')) throw new Error('請先在 Config 設定 Gemini API Key！');

  const validNodes = nodesArray.filter(n => n && n.trim() !== '');
  const BATCH_SIZE = 5;
  const batches    = [];
  for (let i = 0; i < validNodes.length; i += BATCH_SIZE) batches.push(validNodes.slice(i, i+BATCH_SIZE));
  const total      = batches.length;

  if (batchIndex >= total) {
    _clearAllCaches();
    return { done: true, current: total, total, message: `✅ 全部完成！共 ${total} 批次。` };
  }

  const batchNodes = batches[batchIndex];
  const prompt     = buildPrompt(subject, grade, batchNodes, pdfText, mode || 'ai');
  const url        = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;
  const fetchOpts  = {
    method: 'post', contentType: 'application/json',
    payload: JSON.stringify({
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: { responseMimeType: 'application/json', temperature: 0.85 }
    }),
    muteHttpExceptions: true
  };

  let result = null;
  for (let retry = 0; retry < 3; retry++) {
    const resp = UrlFetchApp.fetch(url, fetchOpts);
    result     = JSON.parse(resp.getContentText());
    if (result.error) {
      if (result.error.code === 429 || String(result.error.message).toLowerCase().includes('quota')) {
        Utilities.sleep(20000);
      } else {
        throw new Error('Gemini API 錯誤: ' + result.error.message);
      }
    } else break;
  }
  if (result && result.error) throw new Error('API 持續失敗，請等待後重試。');

  const rawText   = result.candidates[0].content.parts[0].text.replace(/```json/g,'').replace(/```/g,'').trim();
  const questions = JSON.parse(rawText);

  const bank    = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Bank');
  const lastRow = bank.getLastRow() || 1;
  const newRows = questions.map((q, idx) => [
    `AI-${subject}-${Date.now().toString().slice(-6)}-${batchIndex}-${idx}`,
    q.node, q.text, q.type, JSON.stringify(q.options || []),
    normalizeAnswer(q.answer), q.difficulty, grade, subject
  ]);

  if (newRows.length > 0) {
    bank.getRange(lastRow+1, 1, newRows.length, 9).setValues(newRows);
    bank.getRange(lastRow+1, 6, newRows.length, 1).setNumberFormat('@STRING@');
  }

  const isLast = (batchIndex + 1 >= total);
  if (isLast) _clearAllCaches();

  return {
    done: isLast, current: batchIndex+1, total,
    addedThisBatch: newRows.length, nodes: batchNodes,
    message: `[${subject}] 第 ${batchIndex+1}/${total} 批完成（${batchNodes.join('、')}）— 新增 ${newRows.length} 題${pdfText?' (PDF出題)':''}`
  };
}

// ==========================================
// 7. 派題引擎
// ==========================================
function normalizeAnswer(str) {
  if (str === null || str === undefined) return '';
  if (str instanceof Date) return '';
  if (typeof str === 'number') return String(str).trim().toLowerCase();
  return String(str).trim()
    .replace(/\s+/g,'')
    .replace(/[\uff10-\uff19]/g, c => String.fromCharCode(c.charCodeAt(0)-0xFEE0))
    .replace(/，/g,',')
    .toLowerCase();
}

function generateQuiz(weakNode, taskName) {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const cfg       = getConfig();
  const quizCount = parseInt(cfg.QuizCount, 10) || 10;
  let targetGrade = '', targetSubject = '數學';

  const taskSheet = ss.getSheetByName(taskName);
  if (taskSheet && taskSheet.getLastRow() > 1) {
    try {
      const sv = String(taskSheet.getRange(2,5,1,1).getValue()||'').trim();
      if (sv === '國語' || sv === '數學') targetSubject = sv;
    } catch(e) {}
  }

  const hist = ss.getSheetByName('History');
  if (hist && taskName) {
    hist.getDataRange().getValues().slice(1).forEach(r => {
      if (String(r[1]).replace(/'/g,'').trim() === String(taskName).trim()) {
        targetGrade = String(r[2]||'').trim();
        const r3    = String(r[3]||'').trim();
        if (r3 === '國語' || r3 === '數學') targetSubject = r3;
      }
    });
  }

  const cache     = CacheService.getScriptCache();
  const CACHE_KEY = 'BankData_V3';
  let   allQ      = [];
  const cached    = cache.get(CACHE_KEY);

  if (cached) {
    allQ = JSON.parse(cached);
  } else {
    const bank = ss.getSheetByName('Bank');
    if (!bank || bank.getLastRow() <= 1) return [];
    const lastRow = bank.getLastRow();
    bank.getRange(2, 6, lastRow-1, 1).setNumberFormat('@STRING@');
    const data     = bank.getRange(2, 1, lastRow-1, 9).getValues();
    const dispData = bank.getRange(2, 1, lastRow-1, 9).getDisplayValues();

    data.forEach((row, i) => {
      if (!row[0] && !row[2]) return;
      let options = [];
      const rawOpts = row[4] ? String(row[4]).trim() : '';
      if (rawOpts.startsWith('[')) { try { options = JSON.parse(rawOpts); } catch(e) {} }
      else if (rawOpts) { options = rawOpts.split(',').map(o => o.trim()).filter(o => o); }
      const rawAns = (dispData[i] && dispData[i][5]) ? dispData[i][5] : String(row[5]||'');
      allQ.push({
        id: row[0]||`Q${i+2}`, node: row[1]?String(row[1]).trim():'',
        text: row[2], type: row[3], options,
        answer: normalizeAnswer(rawAns), displayAnswer: String(rawAns).trim(),
        difficulty: String(row[6]||'medium').trim(),
        grade: String(row[7]||'').trim(),
        subject: (row[8] && (String(row[8]).trim()==='國語'||String(row[8]).trim()==='數學')) ? String(row[8]).trim() : ''
      });
    });
    try { cache.put(CACHE_KEY, JSON.stringify(allQ), 3600); } catch(e) {}
  }

  const shuffle = arr => {
    for (let i = arr.length-1; i > 0; i--) {
      const j = Math.floor(Math.random()*(i+1));
      [arr[i],arr[j]] = [arr[j],arr[i]];
    }
    return arr;
  };

  let targets = [], fallbacks = [];
  const safeNode = String(weakNode||'').trim();
  allQ.forEach(q => {
    if (q.subject && q.subject !== targetSubject) return;
    if (targetGrade && q.grade && q.grade !== targetGrade) return;
    if (q.node && safeNode && (q.node.includes(safeNode) || safeNode.includes(q.node))) targets.push(q);
    else fallbacks.push(q);
  });

  const pickByDiff = (arr, n) => {
    const e  = shuffle(arr.filter(q => q.difficulty==='easy'));
    const m  = shuffle(arr.filter(q => q.difficulty==='medium'));
    const h  = shuffle(arr.filter(q => q.difficulty==='hard'));
    const ec = Math.round(n*0.3), hc = Math.round(n*0.2), mc = n-ec-hc;
    let   res = [...e.slice(0,ec), ...m.slice(0,mc), ...h.slice(0,hc)];
    if (res.length < n) {
      const rest = shuffle([...e.slice(ec),...m.slice(mc),...h.slice(hc)]);
      res = res.concat(rest.slice(0, n-res.length));
    }
    return res;
  };

  shuffle(targets);
  let final = pickByDiff(targets, Math.min(quizCount, targets.length));
  if (final.length < quizCount) {
    shuffle(fallbacks);
    final = final.concat(pickByDiff(fallbacks, Math.min(quizCount-final.length, fallbacks.length)));
  }
  final = final.map(q => q.type==='single' && q.options.length>1 ? {...q, options: shuffle([...q.options])} : q);
  return shuffle(final);
}

function submitQuizResult(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let res  = ss.getSheetByName('Results');
  if (!res) {
    res = ss.insertSheet('Results');
    res.appendRow(['測驗時間','任務名稱','座號','姓名','分數','作答歷時(秒)','作答明細']);
  }
  if (!data || !data.taskName) throw new Error('作答資料格式錯誤');
  res.appendRow([new Date(), data.taskName, data.seatNo||'', data.name||'', data.score||0, data.timeSpent||0, JSON.stringify(data.details||[])]);
  return true;
}

// ==========================================
// 8. 教學講義生成
// ==========================================
function generateTeachingWorksheet(topNodes, grade, subject) {
  const cfg    = getConfig();
  const apiKey = String(cfg.GeminiAPIKey||'');
  if (!apiKey || apiKey.includes('請在此')) throw new Error('請先設定 Gemini API Key！');

  const subj     = subject || '數學';
  const nodeList = topNodes.map((n,i) => `${i+1}. 「${n}」`).join('\n');

  const prompt = subj === '國語'
    ? `你是台灣資深國小國語科補救教學專家。
請針對「${grade}」學生，為以下知識節點設計補救教學練習：
${nodeList}

每個節點設計 4 題，從最基礎開始，包含多元題型（字音字形選擇、詞語填空、句型練習）。
每題必須包含：step（引導步驟）、hint（思考提示）、type、answer（唯一正確答案）。

只回傳 JSON 陣列：
[{"node":"節點","step":"引導步驟","text":"題目","type":"fill","answer":"答案","hint":"提示"}]`
    : `你是台灣資深國小數學科補救教學專家。
請針對「${grade}」學生，為以下知識節點設計補救教學練習：
${nodeList}

每個節點設計 5 題（填充計算題），從最基礎開始。
每題必須包含：step（解題引導步驟）、hint（關鍵觀念提示）、type（統一"fill"）、answer（純數字或分數）。

只回傳 JSON 陣列：
[{"node":"節點","step":"引導步驟","text":"題目","type":"fill","answer":"答案","hint":"提示"}]`;

  const url  = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;
  const opts = {
    method: 'post', contentType: 'application/json',
    payload: JSON.stringify({ contents:[{parts:[{text:prompt}]}], generationConfig:{responseMimeType:'application/json', temperature:0.5} }),
    muteHttpExceptions: true
  };

  let result = null;
  for (let retry = 0; retry < 3; retry++) {
    const resp = UrlFetchApp.fetch(url, opts);
    result     = JSON.parse(resp.getContentText());
    if (result.error) {
      if (String(result.error.message).toLowerCase().includes('quota')) Utilities.sleep(15000);
      else throw new Error('API 錯誤: ' + result.error.message);
    } else break;
  }
  if (result && result.error) throw new Error('API 持續失敗，請稍後再試。');
  const rawText = result.candidates[0].content.parts[0].text.replace(/```json/g,'').replace(/```/g,'').trim();
  return JSON.parse(rawText);
}

// ==========================================
// 9. 成績分析
// ==========================================
function getTaskResults(taskName) {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const rSht = ss.getSheetByName('Results');
  if (!rSht || rSht.getLastRow() <= 1) return [];

  const data    = rSht.getDataRange().getValues();
  const results = [];
  for (let i = 1; i < data.length; i++) {
    const row     = data[i];
    const rowTask = String(row[1]||'').replace(/'/g,'').trim();
    const clean   = String(taskName||'').replace(/'/g,'').trim();
    if (rowTask !== clean) continue;
    let details = [];
    try { details = JSON.parse(row[6]||'[]'); } catch(e) {}
    results.push({
      submittedAt: row[0] ? new Date(row[0]).toLocaleString('zh-TW') : '',
      taskName: rowTask, seatNo: String(row[2]||''), name: String(row[3]||''),
      score: Number(row[4]||0), timeSpent: Number(row[5]||0), details
    });
  }

  // 每位學生只保留最新一筆
  const seen = {}, deduped = [];
  for (let i = results.length-1; i >= 0; i--) {
    const key = results[i].seatNo + '_' + results[i].name;
    if (!seen[key]) { seen[key] = true; deduped.unshift(results[i]); }
  }
  return deduped.sort((a,b) => Number(a.seatNo)-Number(b.seatNo));
}

function analyzeClassWeakNodes(taskName) {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const results = getTaskResults(taskName);
  if (!results.length) return [];

  // 建立 題目ID → 知識節點 對照表（修正：補查缺少 node 的舊資料）
  const idToNode = {};
  const bank     = ss.getSheetByName('Bank');
  if (bank && bank.getLastRow() > 1) {
    bank.getRange(2, 1, bank.getLastRow()-1, 2).getValues().forEach(r => {
      if (r[0]) idToNode[String(r[0])] = String(r[1]||'');
    });
  }

  const nodeMap = {};
  results.forEach(r => {
    (r.details||[]).forEach(d => {
      if (!d.isCorrect) {
        const node = (d.node && d.node.trim()) || idToNode[String(d.id||'')] || '';
        if (!node) return;
        if (!nodeMap[node]) nodeMap[node] = new Set();
        nodeMap[node].add(r.name || ('座號' + r.seatNo));
      }
    });
  });

  const total = results.length;
  return Object.entries(nodeMap).map(([node, students]) => ({
    node,
    wrongCount:   students.size,
    studentCount: total,
    wrongRate:    Math.round((students.size / total) * 100),
    students:     Array.from(students)
  })).sort((a,b) => b.wrongCount - a.wrongCount);
}

function getQuestionErrorRates(taskName) {
  const results = getTaskResults(taskName);
  if (!results.length) return [];

  const qMap = {};
  results.forEach(r => {
    (r.details||[]).forEach(d => {
      const qid = d.id || d.questionText || 'unknown';
      if (!qMap[qid]) {
        qMap[qid] = {
          questionText:  d.questionText || qid,
          node:          d.node || '',
          correctAns:    d.displayCorrectAns || d.correctAns || '',
          total: 0, wrong: 0,
          wrongStudents: []
        };
      }
      qMap[qid].total++;
      if (!d.isCorrect) {
        qMap[qid].wrong++;
        qMap[qid].wrongStudents.push({
          name:   r.name   || '',
          seatNo: r.seatNo || '',
          ans:    d.userAns || '未作答'
        });
      }
    });
  });

  return Object.values(qMap)
    .map(q => ({ ...q, wrongRate: Math.round((q.wrong / q.total) * 100) }))
    .sort((a,b) => b.wrongRate - a.wrongRate);
}

// ==========================================
// 10. 快取管理
// ==========================================
function _clearAllCaches() {
  ['BankData_V3','BankData_V2','BankData_V1'].forEach(k => {
    try { CacheService.getScriptCache().remove(k); } catch(e) {}
  });
}

function clearBankCache() {
  _clearAllCaches();
  const bank = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Bank');
  if (bank && bank.getLastRow() > 1) bank.getRange(2, 6, bank.getLastRow()-1, 1).setNumberFormat('@STRING@');
  return '✅ 題庫快取已清除！';
}