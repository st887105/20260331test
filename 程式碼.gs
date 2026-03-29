/**
 * ============================================================================
 * 車城國小學力檢測考古題輔助系統 - 後端邏輯 (API 版本)
 * v4.0 - 前後端分離架構
 * ============================================================================
 */

// ==========================================
// API 入口 (供 GitHub Pages 呼叫)
// ==========================================
function doPost(e) {
  try {
    const params = JSON.parse(e.postData.contents);
    const action = params.action;
    const data = params.data || {};
    let result = null;

    switch (action) {
      case 'getTasks': result = getTasks(); break;
      case 'getStudents': result = getStudents(data.taskName); break;
      case 'verifyAdmin': result = verifyAdmin(data.pwd); break;
      case 'getQuizSettings': result = getQuizSettings(); break;
      case 'updateQuizSettings': result = updateQuizSettings(data.count); break;
      case 'setupDatabase': result = setupDatabase(); break;
      case 'clearBankCache': result = clearBankCache(); break;
      case 'uploadTaskData': result = uploadTaskData(data.taskName, data.grade, data.studentData, data.uniqueNodes); break;
      case 'generateBatch': result = generateBatch(data.nodesArray, data.grade, data.batchIndex); break;
      case 'generateQuiz': result = generateQuiz(data.weakNode, data.taskName); break;
      case 'submitQuizResult': result = submitQuizResult(data.payload); break;
      case 'getTaskResults': result = getTaskResults(data.taskName); break;
      case 'getQuestionErrorRates': result = getQuestionErrorRates(data.taskName); break;
      case 'analyzeClassWeakNodes': result = analyzeClassWeakNodes(data.taskName); break;
      case 'generateTeachingWorksheet': result = generateTeachingWorksheet(data.topNodes, data.grade); break;
      default: throw new Error('未知的 API 請求');
    }

    return ContentService.createTextOutput(JSON.stringify({ status: 'success', data: result }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: error.message || String(error) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doOptions(e) {
  return ContentService.createTextOutput("")
    .setMimeType(ContentService.MimeType.TEXT)
    .setHeaders({
      "Access-Control-Allow-Origin": "*",
      "Access-Control-Allow-Methods": "POST, OPTIONS",
      "Access-Control-Allow-Headers": "Content-Type",
      "Access-Control-Max-Age": "86400"
    });
}

// 為了讓您也能從原本的網址預覽，保留簡單的 doGet
function doGet(e) {
  return ContentService.createTextOutput("API 伺服器運作中！請從 GitHub Pages 前端網頁連線。");
}


// ==========================================
// 1. 資料庫初始化
// ==========================================
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ['Config','Bank','History','Results'].forEach(sheetName => {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      if (sheetName === 'Config') {
        sheet.appendRow(['Key','Value']);
        sheet.appendRow(['AdminPassword','1234']);
        sheet.appendRow(['GeminiAPIKey','請在此貼上您的API金鑰']);
        sheet.appendRow(['QuizCount','10']);
      } else if (sheetName === 'Bank') {
        sheet.appendRow(['ID','知識節點','題目','類型(single/fill)','選項(JSON陣列)','正解','難度','適用年級']);
      } else if (sheetName === 'History') {
        sheet.appendRow(['上傳時間','任務名稱','適用年級','學生人數','班級弱點節點']);
      } else if (sheetName === 'Results') {
        sheet.appendRow(['測驗時間','任務名稱','座號','姓名','分數','作答歷時(秒)','作答明細']);
      }
    }
  });
  try { CacheService.getScriptCache().remove('BankData_V2'); } catch(e) {}
  return "✅ 資料庫初始化完成！請至 Config 分頁設定您的 Gemini API Key。";
}

function verifyAdmin(pwd) {
  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  if (!configSheet) return false;
  const data = configSheet.getDataRange().getValues();
  for(let i = 1; i < data.length; i++) {
    if(data[i][0] === 'AdminPassword' && data[i][1].toString() === pwd.toString()) return true;
  }
  return false;
}

function getQuizSettings() {
  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  if (!configSheet) return { quizCount: 10 };
  let quizCount = 10;
  configSheet.getDataRange().getValues().forEach(row => {
    if(row[0] === 'QuizCount') quizCount = parseInt(row[1], 10) || 10;
  });
  return { quizCount };
}

function updateQuizSettings(newCount) {
  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  if (!configSheet) throw new Error("找不到 Config 設定頁");
  const count = parseInt(newCount, 10);
  if (isNaN(count) || count < 1) throw new Error("請輸入有效的數字");
  const data = configSheet.getDataRange().getValues();
  let found = false;
  for(let i = 1; i < data.length; i++) {
    if(data[i][0] === 'QuizCount') { configSheet.getRange(i+1,2).setValue(count); found = true; break; }
  }
  if (!found) configSheet.appendRow(['QuizCount', count]);
  return "✅ 題數已更新為 " + count + " 題！";
}

// ==========================================
// 2. 任務與學生管理
// ==========================================
function uploadTaskData(taskName, grade, studentData, uniqueNodes) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const historySheet = ss.getSheetByName('History');
  if (historySheet) {
    historySheet.appendRow([new Date(), "'" + String(taskName), grade, studentData.length, uniqueNodes.join(', ')]);
  }
  let taskSheet = ss.getSheetByName(taskName);
  if(taskSheet) taskSheet.clear(); else taskSheet = ss.insertSheet(taskName);
  taskSheet.appendRow(['座號','姓名','答對率','知識節點(弱項)']);
  const rows = studentData.map(s => {
    let seat = String(s.seatNo || '').trim(), name = String(s.name || '').trim();
    const match = seat.match(/(\d+)\s*[號]?\s*([A-Za-z\u4e00-\u9fa5]+)$/);
    if (match) { seat = match[1]; name = match[2]; }
    return [seat, name, s.accuracy, s.weakNodes];
  });
  taskSheet.getRange(2, 1, rows.length, 4).setValues(rows);
  return true;
}

function getTasks() {
  const historySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('History');
  if(!historySheet) return [];
  const data = historySheet.getDataRange().getValues();
  return data.slice(1).filter(r => r[1]).map(r => String(r[1])).reverse();
}

function getStudents(taskName) {
  const taskSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(taskName);
  if(!taskSheet) return [];
  return taskSheet.getDataRange().getValues().slice(1).map(r => ({
    seatNo: String(r[0] || ''), name: String(r[1] || ''), weakNode: String(r[3] || '')
  }));
}

// ==========================================
// 3. 派題引擎
// ==========================================
function normalizeAnswer(str) {
  if (str === null || str === undefined) return '';
  if (str instanceof Date) return '';
  if (typeof str === 'number') {
    return String(Number.isInteger(str) ? str : str).trim().toLowerCase();
  }
  return String(str).trim()
    .replace(/\s+/g, '')
    .replace(/[\uff10-\uff19]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0))
    .replace(/，/g, ',')
    .toLowerCase();
}

function generateQuiz(weakNode, taskName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let quizCount = 10;
  const configData = (ss.getSheetByName('Config') || {getDataRange:()=>({getValues:()=>[]})}).getDataRange().getValues();
  configData.forEach(r => { if(r[0]==='QuizCount') quizCount = parseInt(r[1],10)||10; });

  let targetGrade = '';
  const historySheet = ss.getSheetByName('History');
  if (historySheet && taskName) {
    historySheet.getDataRange().getValues().slice(1).forEach(r => {
      if (String(r[1]).replace(/'/g,'') === String(taskName)) targetGrade = String(r[2]).trim();
    });
  }

  const cache = CacheService.getScriptCache();
  const CACHE_KEY = 'BankData_V2';
  let allQuestions = [];
  const cachedData = cache.get(CACHE_KEY);

  if (cachedData) {
    allQuestions = JSON.parse(cachedData);
  } else {
    const bankSheet = ss.getSheetByName('Bank');
    if (!bankSheet || bankSheet.getLastRow() <= 1) return [];
    const lastRow = bankSheet.getLastRow();
    bankSheet.getRange(2, 6, lastRow-1, 1).setNumberFormat('@STRING@');
    const data        = bankSheet.getRange(2, 1, lastRow-1, 8).getValues();
    const displayData = bankSheet.getRange(2, 1, lastRow-1, 8).getDisplayValues();
    data.forEach((row, i) => {
      if (!row[0] && !row[2]) return;
      let options = [];
      const rawOpts = row[4] ? String(row[4]).trim() : '';
      if (rawOpts) {
        if (rawOpts.startsWith('[')) { try { options = JSON.parse(rawOpts); } catch(e) {} }
        else { options = rawOpts.split(',').map(o=>o.trim()).filter(o=>o); }
      }
      const rawAnswer = (displayData[i] && displayData[i][5]) ? displayData[i][5] : String(row[5]||'');
      allQuestions.push({
        id: row[0]||`Q${i+2}`, node: row[1]?String(row[1]).trim():'',
        text: row[2], type: row[3], options,
        answer: normalizeAnswer(rawAnswer),
        displayAnswer: String(rawAnswer).trim(),
        difficulty: String(row[6]||'medium').trim(),
        grade: String(row[7]||'').trim()
      });
    });
    try { cache.put(CACHE_KEY, JSON.stringify(allQuestions), 3600); } catch(e) {}
  }

  const shuffle = arr => {
    for(let i=arr.length-1;i>0;i--){const j=Math.floor(Math.random()*(i+1));[arr[i],arr[j]]=[arr[j],arr[i]];}
    return arr;
  };

  let targets = [], fallbacks = [];
  const safeNode = String(weakNode||'').trim();
  allQuestions.forEach(q => {
    if (targetGrade && q.grade && q.grade !== targetGrade) return;
    if(q.node&&safeNode&&(q.node.includes(safeNode)||safeNode.includes(q.node))) targets.push(q);
    else if(!targetGrade||!q.grade||q.grade===targetGrade) fallbacks.push(q);
  });

  const pickByDiff = (arr, n) => {
    const e=shuffle(arr.filter(q=>q.difficulty==='easy'));
    const m=shuffle(arr.filter(q=>q.difficulty==='medium'));
    const h=shuffle(arr.filter(q=>q.difficulty==='hard'));
    const ec=Math.round(n*0.3),hc=Math.round(n*0.2),mc=n-ec-hc;
    let res=[...e.slice(0,ec),...m.slice(0,mc),...h.slice(0,hc)];
    if(res.length<n){const rest=shuffle([...e.slice(ec),...m.slice(mc),...h.slice(hc)]);res=res.concat(rest.slice(0,n-res.length));}
    return res;
  };

  shuffle(targets);
  let final = pickByDiff(targets, Math.min(quizCount, targets.length));
  if(final.length<quizCount){shuffle(fallbacks);final=final.concat(pickByDiff(fallbacks,Math.min(quizCount-final.length,fallbacks.length)));}

  final = final.map(q => q.type==='single'&&q.options.length>1 ? {...q,options:shuffle([...q.options])} : q);
  return shuffle(final);
}

function submitQuizResult(data) {
  const resultSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Results');
  if (resultSheet) {
    resultSheet.appendRow([new Date(), data.taskName, data.seatNo, data.name, data.score, data.timeSpent, JSON.stringify(data.details)]);
  }
  return true;
}

// ==========================================
// 4. 🚀 批次出題引擎
// ==========================================
function generateBatch(nodesArray, grade, batchIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configData = ss.getSheetByName('Config').getDataRange().getValues();
  let apiKey = '';
  configData.forEach(r => { if(r[0]==='GeminiAPIKey') apiKey = String(r[1]); });
  if(!apiKey || apiKey.includes('請在此')) throw new Error("請先設定 Gemini API Key！");

  const validNodes = nodesArray.filter(n => n && n.trim() !== '');
  const BATCH_SIZE = 5;
  const batches = [];
  for(let i=0;i<validNodes.length;i+=BATCH_SIZE) batches.push(validNodes.slice(i,i+BATCH_SIZE));

  const totalBatches = batches.length;

  if (batchIndex >= totalBatches) {
    try { CacheService.getScriptCache().remove('BankData_V2'); } catch(e) {}
    try { CacheService.getScriptCache().remove('BankData_V1'); } catch(e) {}
    return { done: true, current: totalBatches, total: totalBatches, message: `✅ 全部完成！共 ${totalBatches} 批次。` };
  }

  const batchNodes = batches[batchIndex];
  const nodeList = batchNodes.map((n,i) => `${i+1}. 「${n}」`).join('\n');

  const prompt = `你是台灣資深國小數學命題專家。
請為「${grade}」學生，針對以下 ${batchNodes.length} 個知識節點，各設計 6 道題：

${nodeList}

規範：
- 第1-3題：單選題（type:"single"），4個選項，選項不含雙引號
- 第4-6題：填充題（type:"fill"），options為[]
- 難度：第1題 easy，第2-4題 medium，第5-6題 hard  
- 情境多樣：純計算、生活情境、錯誤辨析各至少出現一次
- answer 必須與 options 某一項完全相同（選擇題），或為純數字/分數（填充題）

只回傳 JSON 陣列，不含任何說明文字：
[{"node":"節點","text":"題目","type":"single","options":["A","B","C","D"],"answer":"A","difficulty":"easy"}]`;

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;
  const fetchOpts = {
    method: "post", contentType: "application/json",
    payload: JSON.stringify({ contents:[{parts:[{text:prompt}]}], generationConfig:{responseMimeType:"application/json",temperature:0.85} }),
    muteHttpExceptions: true
  };

  let result = null;
  for(let retry=0; retry<3; retry++) {
    const resp = UrlFetchApp.fetch(url, fetchOpts);
    result = JSON.parse(resp.getContentText());
    if(result.error) {
      if(result.error.message.toLowerCase().includes("quota")||result.error.message.includes("429")) { Utilities.sleep(20000); }
      else throw new Error("API 錯誤: " + result.error.message);
    } else break;
  }
  if(result && result.error) throw new Error("API 持續失敗，請等待 1 分鐘後重試。");

  let text = result.candidates[0].content.parts[0].text.replace(/```json/g,'').replace(/```/g,'').trim();
  const questions = JSON.parse(text);
  const bankSheet = ss.getSheetByName('Bank');
  const lastRow = bankSheet.getLastRow() || 1;
  const newRows = questions.map((q,idx) => [
    `AI-${Date.now().toString().slice(-6)}-${batchIndex}-${idx}`,
    q.node, q.text, q.type,
    JSON.stringify(q.options||[]),
    normalizeAnswer(q.answer),
    q.difficulty, grade
  ]);
  if(newRows.length > 0) {
    const writeRange = bankSheet.getRange(lastRow+1, 1, newRows.length, 8);
    writeRange.setValues(newRows);
    bankSheet.getRange(lastRow+1, 6, newRows.length, 1).setNumberFormat('@STRING@');
  }

  const isLast = (batchIndex+1 >= totalBatches);
  if(isLast) {
    try { CacheService.getScriptCache().remove('BankData_V2'); } catch(e) {}
    try { CacheService.getScriptCache().remove('BankData_V1'); } catch(e) {}
  }

  return {
    done: isLast,
    current: batchIndex+1,
    total: totalBatches,
    addedThisBatch: newRows.length,
    nodes: batchNodes,
    message: `第 ${batchIndex+1} / ${totalBatches} 批完成（${batchNodes.join('、')}）— 新增 ${newRows.length} 題`
  };
}

// ==========================================
// 5. 清除題庫快取
// ==========================================
function clearBankCache() {
  try { CacheService.getScriptCache().remove('BankData_V2'); } catch(e) {}
  try { CacheService.getScriptCache().remove('BankData_V1'); } catch(e) {}
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bankSheet = ss.getSheetByName('Bank');
  if (bankSheet && bankSheet.getLastRow() > 1) {
    bankSheet.getRange(2, 6, bankSheet.getLastRow()-1, 1).setNumberFormat('@STRING@');
  }
  return "✅ 快取已清除，正解欄格式已修正！下次學生進入將重新讀取最新題庫。";
}

function testAIGeneration() {
  UrlFetchApp.fetch("https://www.google.com");
  Logger.log("權限檢測通過！");
}

// ==========================================
// 7. 教師儀表板：讀取任務作答記錄
// ==========================================
function getTaskResults(taskName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resultSheet = ss.getSheetByName('Results');
  if (!resultSheet || resultSheet.getLastRow() <= 1) return [];

  const data = resultSheet.getDataRange().getValues();
  const results = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowTask = String(row[1] || '').replace(/'/g, '').trim();
    const cleanTask = String(taskName || '').replace(/'/g, '').trim();
    if (rowTask !== cleanTask) continue;

    let details = [];
    try { details = JSON.parse(row[6] || '[]'); } catch(e) {}

    results.push({
      submittedAt : row[0] ? new Date(row[0]).toLocaleString('zh-TW') : '',
      taskName    : rowTask,
      seatNo      : String(row[2] || ''),
      name        : String(row[3] || ''),
      score       : Number(row[4] || 0),
      timeSpent   : Number(row[5] || 0),
      details     : details
    });
  }

  const seen = {};
  const deduped = [];
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
        const node = String(d.node).trim();
        if (!nodeMap[node]) nodeMap[node] = new Set();
        nodeMap[node].add(r.name || r.seatNo);
      }
    });
  });

  const totalStudents = results.length;
  const analysis = Object.entries(nodeMap).map(([node, students]) => ({
    node,
    wrongCount   : students.size,
    studentCount : totalStudents,
    wrongRate    : Math.round((students.size / totalStudents) * 100),
    students     : Array.from(students)
  }));

  return analysis.sort((a, b) => b.wrongCount - a.wrongCount);
}

function generateTeachingWorksheet(topNodes, grade) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configData = ss.getSheetByName('Config').getDataRange().getValues();
  let apiKey = '';
  configData.forEach(r => { if(r[0]==='GeminiAPIKey') apiKey = String(r[1]); });
  if(!apiKey || apiKey.includes('請在此')) throw new Error("請先設定 Gemini API Key！");

  const nodeList = topNodes.map((n, i) => `${i+1}. 「${n}」`).join('\n');

  const prompt = `你是台灣國小數學科教師，擅長設計簡單易懂的補救教學練習。
請針對「${grade}」學生，為以下知識節點設計一份「教師教學用講義」：

${nodeList}

## 講義設計規範
- 每個節點設計 5 題，全部為「填充計算題」（type: "fill"）
- 難度全部設為 easy（最基礎的直接計算，不要應用題情境）
- 題目要有引導步驟（例如：先通分 → 再計算）
- 讓學生能看到計算過程，而不只是填答案
- answer 只寫最終數字答案（如 "3/4" 或 "2"）
- hint 欄位寫一句教學提示（如「先找公分母！」）

只回傳 JSON 陣列，不含說明：
[
  {
    "node": "節點名稱",
    "step": "引導步驟說明（如：第一步通分，第二步相加）",
    "text": "題目（最簡單的直接計算）",
    "type": "fill",
    "answer": "答案",
    "hint": "教學提示一句話"
  }
]`;

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;
  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { responseMimeType: "application/json", temperature: 0.5 }
  };
  const opts = { method:"post", contentType:"application/json", payload:JSON.stringify(payload), muteHttpExceptions:true };

  let result = null;
  for(let retry=0; retry<3; retry++) {
    const resp = UrlFetchApp.fetch(url, opts);
    result = JSON.parse(resp.getContentText());
    if(result.error) {
      if(result.error.message.toLowerCase().includes("quota")) { Utilities.sleep(15000); }
      else throw new Error("API 錯誤: " + result.error.message);
    } else break;
  }
  if(result && result.error) throw new Error("API 持續失敗，請稍後再試。");

  let text = result.candidates[0].content.parts[0].text.replace(/```json/g,'').replace(/```/g,'').trim();
  return JSON.parse(text);
}

// ==========================================
// 8. 全班各題錯誤率統計
// ==========================================
function getQuestionErrorRates(taskName) {
  const results = getTaskResults(taskName);
  if (!results.length) return [];

  const qMap = {};

  results.forEach(r => {
    (r.details || []).forEach(d => {
      const qid = d.id || d.questionText || 'unknown';
      if (!qMap[qid]) {
        qMap[qid] = {
          questionText : d.questionText || qid,
          node         : d.node || '',
          correctAns   : d.displayCorrectAns || d.correctAns || '',
          total        : 0,
          wrong        : 0,
          wrongAnswers : []
        };
      }
      qMap[qid].total++;
      if (!d.isCorrect) {
        qMap[qid].wrong++;
        qMap[qid].wrongAnswers.push({ name: r.name, seatNo: r.seatNo, ans: d.userAns || '未作答' });
      }
    });
  });

  return Object.values(qMap)
    .map(q => ({ ...q, wrongRate: Math.round((q.wrong / q.total) * 100) }))
    .sort((a, b) => b.wrongRate - a.wrongRate);
}