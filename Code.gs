var SHEET_URL = 'https://docs.google.com/spreadsheets/d/1nqU0CZd9ggJQoMDagoLCqQYdmV994zenVjU7qZiNf1k/edit';
var TEST_SHEET_URL = "https://docs.google.com/spreadsheets/d/1-0Lvo_SSe0BATsJ33Rm0-DP__ZcM-9TAlEngnOqtT_M/edit";

function doGet(e) {
  var page = (e && e.parameter && e.parameter.page) ? e.parameter.page : '';
  if (page === 'survey') {
    return HtmlService.createTemplateFromFile('SimpleTest').evaluate()
      .setTitle('더심온 상담플랫폼 시모니(beta ver 0.5)')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('더심온 상담플랫폼 시모니(beta ver 0.5)')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getAppUrl() { return ScriptApp.getService().getUrl(); }

function formatBirth(val) {
  if (!val) return "-";
  if (val instanceof Date) {
    var y = val.getFullYear().toString();
    var m = (val.getMonth() + 1).toString().padStart(2, '0');
    var d = val.getDate().toString().padStart(2, '0');
    return (y + m + d).substring(2);
  }
  var str = String(val).replace(/[^0-9]/g, "");
  if (str.length === 8) return str.substring(2);
  return str.length === 6 ? str : (str || "-");
}

function getCleanRefNum(val) {
  if (val === null || val === undefined || val === "") return "";
  if (val instanceof Date) {
    return val.getFullYear() + "_" + parseInt(val.getMonth() + 1, 10);
  }
  return String(val).replace(/[^0-9_]/g, "").trim();
}

function getCleanName(name) { return String(name || "").replace(/\s/g, ""); }

function verifyLogin(region, userid, password) {
try {
var ss = SpreadsheetApp.openByUrl(SHEET_URL);
var sheet = ss.getSheetByName('상담사계정');
var data = sheet.getDataRange().getValues();
for (var i = 1; i < data.length; i++) {
if (String(data[i][1]).trim() === region && String(data[i][2]).trim() === userid && String(data[i][3]).trim() === String(password)) {
return { success: true, message: data[i][4] + ' 상담사님 환영합니다.', userName: data[i][4], region: region };
}
}
return { success: false, message: '로그인 정보 불일치' };
} catch(e) { return { success: false }; }
}

function loadDashboard(region, userName) {
  var template = HtmlService.createTemplateFromFile('Dashboard');
  template.region = region || "";
  template.userName = userName || "";
  return template.evaluate().getContent();
}

function searchClient(keyword, region) {
  try {
    var ss = SpreadsheetApp.openByUrl(SHEET_URL);
    var sheet = ss.getSheetByName(region + "DB");
    var data = sheet.getDataRange().getValues();
    var results = [];
    var searchStr = getCleanName(keyword);
    for (var i = 1; i < data.length; i++) {
      var name = getCleanName(data[i][3]);
      var empId = getCleanName(data[i][9]);
      if (name.includes(searchStr) || empId.includes(searchStr)) {
        results.push({
          name: String(data[i][3] || '').trim(),
          empId: String(data[i][9] || '').trim(),
          internalId: String(data[i][8] || '').trim(),
          birth: formatBirth(data[i][5]),
          rank: String(data[i][2]),
          dept: String(data[i][1]),
          station: String(data[i][0] || ''),
          gender: String(data[i][6] || '')
        });
      }
    }
    return { success: true, data: results };
  } catch (e) { return { success: false }; }
}

function getAllClientsForGroup(region) {
  try {
    var ss = SpreadsheetApp.openByUrl(SHEET_URL);
    var sheet = ss.getSheetByName(region + "DB");
    if (!sheet) return { success: false };
    var data = sheet.getDataRange().getValues();
    var results = [];
    
    for (var i = 1; i < data.length; i++) {
      var name = String(data[i][3] || '').trim();
      if (!name) continue;
      
      var fullDept = String(data[i][1] || '').trim();
      var cleanStr = fullDept.replace(/^(경상남도|경남|창원시|창원|중앙119구조본부)\s*/, '').trim();
      var parts = cleanStr.split(/\s+/);
      var st = "", dp = "";
      
      if (parts.length === 1) { 
        st = parts[0]; dp = parts[0]; 
      } else if (parts.length >= 2) { 
        st = parts[0]; dp = parts.slice(1).join(' '); 
      }

      results.push({
        name: name,
        empId: String(data[i][9] || '').trim(),
        internalId: String(data[i][8] || '').trim(),
        birth: formatBirth(data[i][5]),
        rank: String(data[i][2] || ''),
        station: st,
        dept: dp,
        gender: String(data[i][6] || '')
      });
    }
    return { success: true, data: results };
  } catch (e) { return { success: false }; }
}

function loadClientDetailPage(internalId, region, userName) {
  var template = HtmlService.createTemplateFromFile('ClientDetail');
  template.internalId = internalId;
  template.region = region || "";
  template.userName = userName || ""; 
  return template.evaluate().getContent();
}

function getClientDetailData(internalId, region) {
  try {
    var ss = SpreadsheetApp.openByUrl(SHEET_URL);
    var dbSheet = ss.getSheetByName(region + "DB");
    if (dbSheet) {
      var dbData = dbSheet.getDataRange().getValues();
      for (var i = 1; i < dbData.length; i++) {
        if (String(dbData[i][8]).trim() === String(internalId).trim()) {
          return { success: true, info: { dept: String(dbData[i][1]), rank: String(dbData[i][2]), name: String(dbData[i][3]), birth: formatBirth(dbData[i][5]), gender: String(dbData[i][6]), empId: String(dbData[i][9] || "").trim() } };
        }
      }
    }
    var dataSheet = ss.getSheetByName(region + "DATA");
    if (dataSheet) {
      var recordData = dataSheet.getDataRange().getValues();
      for (var j = recordData.length - 1; j >= 1; j--) {
        if (String(recordData[j][33]).trim() === String(internalId).trim()) {
          return { success: true, info: { dept: String(recordData[j][4]), rank: String(recordData[j][6]), name: String(recordData[j][5]), birth: String(recordData[j][7]), gender: String(recordData[j][8]), empId: String(recordData[j][34] || "").trim() } };
        }
      }
    }
    return { success: false };
  } catch(e) { return { success: false }; }
}

function getMyClients(region, userName) {
  try {
    var ss = SpreadsheetApp.openByUrl(SHEET_URL);
    var recordSheet = ss.getSheetByName(region + "DATA");
    if (!recordSheet) return { success: true, data: [] }; 
    var recordData = recordSheet.getDataRange().getValues();
    var clientMap = {}; 
    var searchUserName = getCleanName(userName);
    for (var i = recordData.length - 1; i >= 1; i--) {
      var sheetCounselorName = getCleanName(recordData[i][1]);
      if (sheetCounselorName === searchUserName && searchUserName !== "") { 
        var internalId = String(recordData[i][33]).trim(); 
        if (!clientMap[internalId]) {
          clientMap[internalId] = { name: String(recordData[i][5]), empId: String(recordData[i][34] || "").trim(), internalId: internalId, birth: String(recordData[i][7]), rank: String(recordData[i][6]), dept: String(recordData[i][4]) };
        }
      }
    }
    return { success: true, data: Object.values(clientMap) };
  } catch (e) { return { success: false, message: "목록 로드 오류: " + e.message }; }
}

function loadConsultationForm(internalId, region, userName, editRowIndex) {
  var template = HtmlService.createTemplateFromFile('ConsultationForm');
  template.internalId = internalId; template.region = region || ""; template.userName = userName || ""; template.editRowIndex = editRowIndex || "";
  return template.evaluate().getContent();
}

function saveConsultation(formData) {
  var lock = LockService.getScriptLock();
  if (lock.tryLock(30000)) {
    try {
      var ss = SpreadsheetApp.openByUrl(SHEET_URL);
      var sheetName = formData.region + "DATA"; 
      var dataSheet = ss.getSheetByName(sheetName);
      if (!dataSheet) return { success: false, message: sheetName + " 시트가 존재하지 않습니다." };
      var data = dataSheet.getDataRange().getValues();
      var internalId = formData.internalId;
      var currentYear = new Date().getFullYear();

      if (formData.editRowIndex && formData.editRowIndex !== "") {
        var updateRowIndex = parseInt(formData.editRowIndex, 10);
        var existingRefNum = String(data[updateRowIndex - 1][0]); 
        var updatedRow = [
          existingRefNum, formData.counselor, formData.sido, formData.station, formData.dept, formData.clientName, formData.rank, formData.birth, formData.gender, formData.ageGroup,
          formData.job, formData.mainCat, formData.subCat, formData.disasterName, formData.prevCounseling, formData.path, formData.medExp, formData.month, formData.day, 
          formData.timeRange, formData.duration, formData.method, formData.place, formData.stress, formData.problemType, formData.action, formData.summary, formData.testType, formData.testResult,
          formData.emergencyDate, formData.emergencyMethod, formData.emergencyContent, formData.emergencyAction, formData.internalId, formData.empId, formData.digitalId,
          formData.majorSituation || "", formData.situationDate || "", formData.inDepthComplaint || "", formData.inDepthContent || "", formData.counselorEval || "", formData.inDepthAction || "", new Date() 
        ];
        dataSheet.getRange(updateRowIndex, 1, 1, updatedRow.length).setNumberFormat('@');
        dataSheet.getRange(updateRowIndex, 1, 1, updatedRow.length).setValues([updatedRow]);
        return { success: true, message: "성공적으로 수정되었습니다." };
      }
      if (internalId.startsWith("NEW_")) {
        try {
          var testSs = SpreadsheetApp.openByUrl(TEST_SHEET_URL); var testSheet = testSs.getSheetByName(formData.region + "_검사DATA");
          if (testSheet) {
            var testData = testSheet.getDataRange().getValues();
            for (var t = testData.length - 1; t >= 1; t--) {
              if (getCleanName(testData[t][1]) === getCleanName(formData.clientName) && String(testData[t][14]).startsWith("TEMP_")) {
                internalId = String(testData[t][14]).trim(); formData.internalId = internalId; break; 
              }
            }
          }
        } catch(err) {}
      }
      var firstCaseNum = 0, sessionCount = 0, maxFirstNum = 0;
      for (var i = 1; i < data.length; i++) {
        var rowId = String(data[i][33]); 
        var refStr = getCleanRefNum(data[i][0]);
        if (refStr) {
          var refParts = refStr.split('_');
          if (refParts.length >= 2) {
            var num = parseInt(refParts[1], 10);
            if (!isNaN(num) && num > maxFirstNum) maxFirstNum = num;
          }
        }
        if (rowId === internalId) {
          if (firstCaseNum === 0 && refStr) {
            var pNum = parseInt(refStr.split('_')[1], 10);
            if(!isNaN(pNum)) firstCaseNum = pNum;
          }
          sessionCount++;
        }
      }
      var finalRefNum = (firstCaseNum === 0) ? (currentYear + "_" + (maxFirstNum + 1)) : (currentYear + "_" + firstCaseNum + "_" + (sessionCount + 1));
      var newRowData = [
        finalRefNum, formData.counselor, formData.sido, formData.station, formData.dept, formData.clientName, formData.rank, formData.birth, formData.gender, formData.ageGroup,
        formData.job, formData.mainCat, formData.subCat, formData.disasterName, formData.prevCounseling, formData.path, formData.medExp, formData.month, formData.day, 
        formData.timeRange, formData.duration, formData.method, formData.place, formData.stress, formData.problemType, formData.action, formData.summary, formData.testType, formData.testResult,
        formData.emergencyDate, formData.emergencyMethod, formData.emergencyContent, formData.emergencyAction, formData.internalId, formData.empId, formData.digitalId,
        formData.majorSituation || "", formData.situationDate || "", formData.inDepthComplaint || "", formData.inDepthContent || "", formData.counselorEval || "", formData.inDepthAction || "", new Date()
      ];
      var newRowIndex = dataSheet.getLastRow() + 1;
      dataSheet.getRange(newRowIndex, 1, 1, newRowData.length).setNumberFormat('@');
      dataSheet.getRange(newRowIndex, 1, 1, newRowData.length).setValues([newRowData]);
      return { success: true, message: "성공적으로 저장되었습니다. 관리번호: " + finalRefNum };
    } catch (e) { return { success: false, message: "저장 오류: " + e.message }; } finally { lock.releaseLock(); }
  }
}

function loadGroupForm(region, userName) {
  var template = HtmlService.createTemplateFromFile('GroupForm');
  template.region = region || "";
  template.userName = userName || "";
  return template.evaluate().getContent();
}

function saveGroupConsultation(commonData, participants) {
  var lock = LockService.getScriptLock();
  if (lock.tryLock(30000)) { // 30초 안정 락
    try {
      var ss = SpreadsheetApp.openByUrl(SHEET_URL);
      var sheetName = commonData.region + "DATA";
      var dataSheet = ss.getSheetByName(sheetName);
      if (!dataSheet) return { success: false, message: sheetName + " 시트가 존재하지 않습니다." };

      var data = dataSheet.getDataRange().getValues();
      var currentYear = new Date().getFullYear();
      var groupId = "GRP_" + new Date().getTime(); 

      // 1. 데이터 시트에 있는 최고 번호 찾기 (NaN 방지 철저)
      var maxFirstNum = 0;
      for (var i = 1; i < data.length; i++) {
        var refStr = getCleanRefNum(data[i][0]);
        if (refStr) {
          var refParts = refStr.split('_');
          if (refParts.length >= 2) {
            var num = parseInt(refParts[1], 10);
            if (!isNaN(num) && num > maxFirstNum) {
              maxFirstNum = num;
            }
          }
        }
      }
      var nextNewNum = maxFirstNum + 1;
      var newRows = [];

      // 2. 참석자별로 고유 번호 생성 및 배열에 밀어넣기
      for (var p = 0; p < participants.length; p++) {
        var pt = participants[p];
        var internalId = pt.internalId;
        if (!internalId || internalId === "") { internalId = "TEMP_" + new Date().getTime() + "_" + p; }

        var firstCaseNum = 0, sessionCount = 0;
        // 기존 상담 기록 스캔
        for (var j = 1; j < data.length; j++) {
          var rowId = String(data[j][33]).trim();
          if (rowId !== "" && rowId === internalId) {
            if (firstCaseNum === 0) {
              var pastRefStr = getCleanRefNum(data[j][0]);
              if (pastRefStr) {
                var pastRefParts = pastRefStr.split('_');
                if (pastRefParts.length >= 2) {
                  var pastNum = parseInt(pastRefParts[1], 10);
                  if (!isNaN(pastNum)) firstCaseNum = pastNum;
                }
              }
            }
            sessionCount++;
          }
        }
        
        var finalRefNum = "";
        if (firstCaseNum === 0) {
          // 신규 내담자
          finalRefNum = currentYear + "_" + nextNewNum;
          nextNewNum++; // 다음 신규 사람을 위해 번호 1 증가
        } else {
          // 기존 내담자
          finalRefNum = currentYear + "_" + firstCaseNum + "_" + (sessionCount + 1);
        }

        var ageGrp = "";
        var birthStr = String(pt.birth).replace(/[^0-9]/g, '');
        if (birthStr.length === 6) {
          var yy = parseInt(birthStr.substring(0, 2), 10);
          var age = new Date().getFullYear() - ((yy > 30) ? 1900 + yy : 2000 + yy);
          if (age >= 60) ageGrp = "60대 이상"; else if (age >= 50) ageGrp = "50대"; else if (age >= 40) ageGrp = "40대"; else if (age >= 30) ageGrp = "30대"; else if (age >= 20) ageGrp = "20대";
        }

        // 불필요한 항목(14~16, 23, 24, 26 등) 완벽하게 공백 처리
        var newRowData = [
          finalRefNum, commonData.counselor, commonData.region, pt.station || "", pt.dept || "",
          pt.name || "", pt.rank || "", pt.birth || "", pt.gender || "", ageGrp,
          "", "집단상담", commonData.subCat, "", "", "", "", // 10:직무, 11:대분류, 12:중분류, 13:재난, 14:기존상담, 15:경로, 16:정신과
          commonData.month, commonData.day, commonData.timeRange, commonData.duration, commonData.method, commonData.place,
          "", "", commonData.action, "", "", "", // 23:스트레스, 24:문제유형, 25:내용, 26:조치, 27:검사, 28:검사결과
          "", "", "", "", internalId, pt.empId || "", groupId, // 29~32:긴급콜, 33:internalId, 34:empId, 35:groupId
          "", "", "", "", commonData.eval, "", new Date() // 36~39:심층상담, 40:평가, 41:조치, 42:시간
        ];
        newRows.push(newRowData);
      }

      if (newRows.length > 0) {
        var startRow = dataSheet.getLastRow() + 1;
        dataSheet.getRange(startRow, 1, newRows.length, newRows[0].length).setNumberFormat('@');
        dataSheet.getRange(startRow, 1, newRows.length, newRows[0].length).setValues(newRows);
      }
      return { success: true, message: participants.length + "명의 집단상담 기록이 성공적으로 일괄 저장되었습니다." };
    } catch (e) { 
      return { success: false, message: "저장 오류: " + e.message }; 
    } finally { 
      lock.releaseLock(); 
    }
  }
  return { success: false, message: "시스템이 바쁩니다. 잠시 후 다시 시도해주세요." };
}

function getPastConsultations(internalId, region) {
  try {
    var ss = SpreadsheetApp.openByUrl(SHEET_URL);
    var sheet = ss.getSheetByName(region + "DATA");
    if (!sheet) return { success: false };
    var data = sheet.getDataRange().getValues();
    var history = [];
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][33]).trim() === String(internalId).trim()) {
        history.push({ rowIndex: i + 1, refNum: String(data[i][0]), counselor: data[i][1], month: data[i][17], day: data[i][18], stress: data[i][23] });
      }
    }
    history.reverse(); return { success: true, data: history };
  } catch(e) { return { success: false, message: e.message }; }
}

function loadConsultationView(rowIndex, region, userName) {
  var template = HtmlService.createTemplateFromFile('ConsultationView');
  template.rowIndex = rowIndex; template.region = region || ""; template.userName = userName || ""; 
  return template.evaluate().getContent();
}

function getSingleConsultationByIndex(rowIndex, region) {
  try {
    var ss = SpreadsheetApp.openByUrl(SHEET_URL);
    var sheet = ss.getSheetByName(region + "DATA");
    if (!sheet) return { success: false, message: "DATA 시트를 찾을 수 없습니다." };
    var data = sheet.getDataRange().getDisplayValues();
    var index = parseInt(rowIndex, 10);
    if (index > 0 && index <= data.length) { return { success: true, record: data[index - 1] }; }
    return { success: false, message: "해당 데이터를 찾을 수 없습니다." };
  } catch(e) { return { success: false, message: e.message }; }
}

function getTestRecords(internalId, region) {
  try {
    var ss = SpreadsheetApp.openByUrl(SHEET_URL);
    var sheet = ss.getSheetByName(region + "DATA");
    if (!sheet) return { success: false };
    var data = sheet.getDataRange().getValues();
    var testHistory = [];
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][33]).trim() === String(internalId).trim() && (data[i][27] || data[i][28])) {
        testHistory.push({ date: (data[i][17] ? data[i][17]+"월 " : "") + (data[i][18] ? data[i][18]+"일" : "-"), type: data[i][27], result: data[i][28], action: "-" });
      }
    }
    testHistory.reverse(); return { success: true, data: testHistory };
  } catch(e) { return { success: false }; }
}

function findInternalId(name, birth, region) {
  try {
    var ss = SpreadsheetApp.openByUrl(SHEET_URL);
    var searchBirth = String(birth).replace(/[^0-9]/g, ""); 
    var dbSheet = ss.getSheetByName(region + "DB");
    if (dbSheet) {
      var dbData = dbSheet.getDataRange().getValues();
      for (var i = 1; i < dbData.length; i++) {
        if (getCleanName(dbData[i][3]) === getCleanName(name) && formatBirth(dbData[i][5]) === searchBirth) { return { success: true, internalId: String(dbData[i][8]).trim() }; }
      }
    }
    var dataSheet = ss.getSheetByName(region + "DATA");
    if (dataSheet) {
      var recordData = dataSheet.getDataRange().getValues();
      for (var j = recordData.length - 1; j >= 1; j--) {
        if (getCleanName(recordData[j][5]) === getCleanName(name) && String(recordData[j][7]).replace(/[^0-9]/g, "") === searchBirth) { return { success: true, internalId: String(recordData[j][33]).trim() }; }
      }
    }
    return { success: true, internalId: "TEMP_" + new Date().getTime() };
  } catch(e) { return { success: false, message: e.message }; }
}

function saveSimpleTest(testData) {
  try {
    var ss = SpreadsheetApp.openByUrl(TEST_SHEET_URL);
    var sheet = ss.getSheetByName(testData.region + "_검사DATA");
    sheet.appendRow([new Date(), testData.name, testData.birth, testData.region, testData.ptsdScore, testData.depressScore, testData.suicideScore, testData.q1, testData.q2, testData.q3, testData.q4, testData.q5, testData.q6, testData.q7, testData.internalId]);
    return { success: true, message: "검사가 완료되었습니다. 감사합니다." };
  } catch(e) { return { success: false }; }
}

function getSimpleTestRecords(internalId, region) {
  try {
    var ss = SpreadsheetApp.openByUrl(TEST_SHEET_URL);
    var sheet = ss.getSheetByName(region + "_검사DATA");
    if (!sheet) return { success: false };
    var data = sheet.getDataRange().getValues();
    var testHistory = [];
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][14]).trim() === String(internalId).trim()) {
        var d = data[i][0];
        testHistory.push({ date: (d instanceof Date ? d.getFullYear() + "-" + (d.getMonth() + 1) + "-" + d.getDate() : String(d).substring(0, 10)), ptsd: data[i][4], depress: data[i][5], suicide: data[i][6] });
      }
    }
    testHistory.reverse(); return { success: true, data: testHistory };
  } catch(e) { return { success: false }; }
}

function loadLoginPage() { return HtmlService.createTemplateFromFile('Index').evaluate().getContent(); }

function loadAdminStatsPage(region, userName) {
  var template = HtmlService.createTemplateFromFile('AdminStats');
  template.region = region || ""; template.userName = userName || "";
  return template.evaluate().getContent();
}

function calculateAdminStats(region, targetYear) {
  try {
    var ss = SpreadsheetApp.openByUrl(SHEET_URL);
    var sheet = ss.getSheetByName(region + "DATA");
    if (!sheet) return { success: false, message: region + "DATA 시트를 찾을 수 없습니다." };
    
    var data = sheet.getDataRange().getValues();
    var year = targetYear ? parseInt(targetYear, 10) : new Date().getFullYear();

    var stats = {};
    var quarters = ['q1', 'q2', 'q3', 'q4'];
    
    quarters.forEach(function(q) {
      stats[q] = {
        pSet: {}, count: 0,
        screen_pSet: {}, screen_c: 0,
        indepth_pSet: {}, indepth_c: 0,
        life_pSet: {}, life_c: 0, life_test_c: 0,
        emerg_pSet: {}, emerg_c: 0,
        group_pSet: {}, group_sessionSet: {},
        call_c: 0,
        gender: {}, age: {}, rank: {}, job: {},
        life_gender: {}, life_rank: {},
        method: {}, place: {}, path: {}, prevExp: {},
        stress: {}, problem: {}, action: {}
      };
    });

    var personSessionCount = {};
    var stationMap = {};  // { "소방서명": { c: 횟수, pSet: {} } }

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var refNum = String(row[0] || "");
      var rowYear = refNum.split('_')[0];
      if (rowYear && parseInt(rowYear, 10) !== year) continue; 

      var month = parseInt(row[17], 10);
      if (isNaN(month)) continue;

      var qTarget = 'q1';
      if (month >= 4 && month <= 6) qTarget = 'q2';
      else if (month >= 7 && month <= 9) qTarget = 'q3';
      else if (month >= 10 && month <= 12) qTarget = 'q4';

      var empId = String(row[34] || "").trim();
      var name = String(row[5] || "").trim();
      var birth = String(row[7] || "").trim();
      var uniqueId = empId ? empId : (name + "_" + birth);
      if (uniqueId === "_") uniqueId = "UNKNOWN_" + i;

      personSessionCount[uniqueId] = (personSessionCount[uniqueId] || 0) + 1;

      var subCat = String(row[12] || "");
      var mainCat = String(row[11] || "");
      var isGroup = mainCat.indexOf("집단상담") > -1;
      var groupId = String(row[35] || "");
      
      var gender = String(row[8] || "").trim();
      var ageGrp = String(row[9] || "").trim();
      var rankRaw = String(row[6] || "").trim();
      var jobRaw = String(row[10] || "").trim();
      
      var methodRaw = String(row[21] || "");
      var placeRaw = String(row[22] || "");
      var pathRaw = String(row[15] || "");
      var prevExpRaw = String(row[14] || "");
      var stressRaw = String(row[23] || "");
      var probRaw = String(row[24] || "");
      var actionRaw = String(row[26] || "");
      var testRaw = String(row[27] || "");

      var st = stats[qTarget];
      st.pSet[uniqueId] = true;
      st.count++;
      // 소방서/센터별 횟수·인원 집계
      var stationName = String(row[3] || '').trim();
      if (stationName) {
        if (!stationMap[stationName]) stationMap[stationName] = { c: 0, pSet: {} };
        stationMap[stationName].c++;
        stationMap[stationName].pSet[uniqueId] = true;
      }

      if(subCat.indexOf("스크리닝") > -1) { st.screen_pSet[uniqueId] = true; st.screen_c++; }
      if(subCat.indexOf("심층상담") > -1) { st.indepth_pSet[uniqueId] = true; st.indepth_c++; }
      if(subCat.indexOf("긴급심리") > -1) { st.emerg_pSet[uniqueId] = true; st.emerg_c++; }
      if(subCat.indexOf("긴급콜") > -1) { st.call_c++; }
      if(isGroup) {
        st.group_pSet[uniqueId] = true;
        if(groupId) st.group_sessionSet[groupId] = true;
      }

      if(mainCat.indexOf("생애주기") > -1) { 
        st.life_pSet[uniqueId] = true; 
        st.life_c++; 
        if(testRaw !== "") st.life_test_c++;
        if(gender) st.life_gender[gender] = (st.life_gender[gender] || 0) + 1;
        if(rankRaw) st.life_rank[rankRaw] = (st.life_rank[rankRaw] || 0) + 1;
      }

      if(gender) st.gender[gender] = (st.gender[gender] || 0) + 1;
      if(ageGrp) st.age[ageGrp] = (st.age[ageGrp] || 0) + 1;
      if(rankRaw) st.rank[rankRaw] = (st.rank[rankRaw] || 0) + 1;
      
      var jMatch = jobRaw.match(/^(\d)/);
      if(jMatch) st.job[jMatch[1]] = (st.job[jMatch[1]] || 0) + 1;

      var mMatch = methodRaw.match(/^(\d)/);
      if(mMatch) st.method[mMatch[1]] = (st.method[mMatch[1]] || 0) + 1;

      var plMatch = placeRaw.match(/^(\d)/);
      if(plMatch) st.place[plMatch[1]] = (st.place[plMatch[1]] || 0) + 1;

      var paMatch = pathRaw.match(/^(\d)/);
      if(paMatch) st.path[paMatch[1]] = (st.path[paMatch[1]] || 0) + 1;

      var prMatch = prevExpRaw.match(/^(\d(-\d)?)/);
      if(prMatch) st.prevExp[prMatch[1]] = (st.prevExp[prMatch[1]] || 0) + 1;

      var aMatch = actionRaw.match(/^(\d)/);
      if(aMatch) st.action[aMatch[1]] = (st.action[aMatch[1]] || 0) + 1;

      if(stressRaw) {
         var sMatch = stressRaw.match(/^(\d-\d+)/);
         var sKey = sMatch ? sMatch[1] : '6'; 
         st.stress[sKey] = (st.stress[sKey] || 0) + 1;
      }

      if(probRaw) {
        var pArr = probRaw.split(",");
        for(var p=0; p<pArr.length; p++) {
          var numMatch = pArr[p].trim().match(/^(\d+)/);
          if(numMatch) st.problem[numMatch[1]] = (st.problem[numMatch[1]] || 0) + 1;
        }
      }
    }

    var result = {};
    quarters.forEach(function(q) {
       var st = stats[q];
       result[q] = {
         totalP: Object.keys(st.pSet).length, totalC: st.count,
         screenP: Object.keys(st.screen_pSet).length, screenC: st.screen_c,
         indepthP: Object.keys(st.indepth_pSet).length, indepthC: st.indepth_c,
         lifeP: Object.keys(st.life_pSet).length, lifeC: st.life_c, lifeTestC: st.life_test_c,
         emergP: Object.keys(st.emerg_pSet).length, emergC: st.emerg_c,
         groupP: Object.keys(st.group_pSet).length, groupC: Object.keys(st.group_sessionSet).length,
         callC: st.call_c,
         gender: st.gender, age: st.age, rank: st.rank, job: st.job,
         life_gender: st.life_gender, life_rank: st.life_rank,
         method: st.method, place: st.place, path: st.path, prevExp: st.prevExp,
         stress: st.stress, problem: st.problem, action: st.action
       };
    });

    var over4Count = 0;
    for(var k in personSessionCount) {
      if(personSessionCount[k] >= 4) over4Count++;
    }
    result.over4Total = over4Count;

    // 소방서/센터별 완료 데이터 생성
    var stationDone = {};
    for (var sn in stationMap) {
      stationDone[sn] = {
        c: stationMap[sn].c,
        p: Object.keys(stationMap[sn].pSet).length
      };
    }
    result.stationDone = stationDone;
    return { success: true, data: result, year: year };
  } catch (error) { 
    return { success: false, message: error.message }; 
  }
}

function getUserRole(userName, region) {
try {
var ss = SpreadsheetApp.openByUrl(SHEET_URL);
var sheet = ss.getSheetByName("상담사계정");
if (!sheet) return "일반";
var data = sheet.getDataRange().getValues();
for (var i = 1; i < data.length; i++) {
var row = data[i];
if (row[1] === region && row[4] === userName) { return String(row[5]).trim(); }
}
return "일반";
} catch (error) { return "일반"; }
}

// --- 기존 코드들 아래쪽에 추가해 주세요 ---

function loadGroupView(groupId, region, userName) {
  var template = HtmlService.createTemplateFromFile('GroupView');
  template.groupId = groupId;
  template.region = region || "";
  template.userName = userName || "";
  return template.evaluate().getContent();
}

function getGroupConsultationData(groupId, region) {
  try {
    var ss = SpreadsheetApp.openByUrl(SHEET_URL);
    var sheet = ss.getSheetByName(region + "DATA");
    if (!sheet) return { success: false, message: "시트를 찾을 수 없습니다." };

    var data = sheet.getDataRange().getDisplayValues(); 
    var participants = [];
    var common = null;

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][35]).trim() === groupId) { // 35번 열이 groupId
        if (!common) {
          var refParts = String(data[i][0]).split('_');
          var year = refParts[0] || new Date().getFullYear();
          var month = String(data[i][17]).padStart(2, '0');
          var day = String(data[i][18]).padStart(2, '0');
          
          common = {
            date: year + ". " + month + ". " + day + ".",
            timeRange: data[i][19],
            place: data[i][22],
            counselor: data[i][1],
            subCat: data[i][12],
            action: data[i][25],
            eval: data[i][40]
          };
        }
        participants.push({
          name: data[i][5],
          station: data[i][3],
          dept: data[i][4],
          rank: data[i][6]
        });
      }
    }

    if (participants.length > 0) {
      return { success: true, common: common, participants: participants };
    }
    return { success: false, message: "해당 그룹 데이터를 찾을 수 없습니다." };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function loadGroupList(region, userName) {
var template = HtmlService.createTemplateFromFile('GroupList');
template.region = region || "";
template.userName = userName || "";
return template.evaluate().getContent();
}

function getGroupList(region) {
try {
var ss = SpreadsheetApp.openByUrl(SHEET_URL);
var sheet = ss.getSheetByName(region + "DATA");
if (!sheet) return { success: false, message: "시트를 찾을 수 없습니다." };

var data = sheet.getDataRange().getDisplayValues();
var groupMap = {}; 

for (var i = data.length - 1; i >= 1; i--) {
  var groupId = String(data[i][35]).trim(); 
  if (groupId && groupId.startsWith("GRP_")) {
    if (!groupMap[groupId]) {
      var refParts = String(data[i][0]).split('_');
      var year = refParts[0] || new Date().getFullYear();
      var month = String(data[i][17]).padStart(2, '0');
      var day = String(data[i][18]).padStart(2, '0');

      groupMap[groupId] = {
        groupId: groupId,
        date: year + "-" + month + "-" + day,
        subCat: data[i][12], 
        place: data[i][22],  
        counselor: data[i][1], 
        count: 1
      };
    } else {
      groupMap[groupId].count++;
    }
  }
}

var results = Object.values(groupMap);
return { success: true, data: results };
} catch(e) {
return { success: false, message: e.message };
}
}

function loadGroupView(groupId, region, userName) {
var template = HtmlService.createTemplateFromFile('GroupView');
template.groupId = groupId;
template.region = region || "";
template.userName = userName || "";
return template.evaluate().getContent();
}

function getGroupConsultationData(groupId, region) {
try {
var ss = SpreadsheetApp.openByUrl(SHEET_URL);
var sheet = ss.getSheetByName(region + "DATA");
if (!sheet) return { success: false, message: "시트를 찾을 수 없습니다." };

var data = sheet.getDataRange().getDisplayValues(); 
var participants = [];
var common = null;

for (var i = 1; i < data.length; i++) {
  if (String(data[i][35]).trim() === groupId) { 
    if (!common) {
      var refParts = String(data[i][0]).split('_');
      var year = refParts[0] || new Date().getFullYear();
      var month = String(data[i][17]).padStart(2, '0');
      var day = String(data[i][18]).padStart(2, '0');
      
      common = {
        date: year + ". " + month + ". " + day + ".",
        timeRange: data[i][19],
        place: data[i][22],
        counselor: data[i][1],
        subCat: data[i][12],
        action: data[i][25],
        eval: data[i][40]
      };
    }
    participants.push({
      name: data[i][5],
      station: data[i][3],
      dept: data[i][4],
      rank: data[i][6]
    });
  }
}

if (participants.length > 0) {
  return { success: true, common: common, participants: participants };
}
return { success: false, message: "해당 그룹 데이터를 찾을 수 없습니다." };
} catch(e) {
return { success: false, message: e.message };
}
}