// GAS 後端程式碼.gs - 完整修正版本
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // 確保高併發時資料不會寫爛
    
    var requestData = JSON.parse(e.postData.contents);
    var action = requestData.action;
    var adminLineUid = requestData.lineUid ? String(requestData.lineUid).trim() : "";
    
    // 安全驗證：確保操作者具有管理員權限
    if (action === "reviewForm" || action === "saveHRProfile" || action === "deleteHRProfile") {
      if (!isAdmin(adminLineUid)) {
        return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "權限不足" })).setMimeType(ContentService.MimeType.JSON);
      }
    }
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 功能一：審核表單
    if (action === "reviewForm") {
      var formSheet = ss.getSheetByName("Form_Applications");
      var data = formSheet.getDataRange().getValues();
      var formId = requestData.formId;
      var statusResult = requestData.statusResult;
      
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim() === String(formId).trim()) {
          formSheet.getCell(i + 1, 9).setValue(statusResult); // 第9欄為審核狀態
          break;
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ status: "success" })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // 功能二：儲存/修改 員工檔案 (解決開頭0被吃掉的問題)
    if (action === "saveHRProfile") {
      var hrSheet = ss.getSheetByName("員工檔案");
      var emp = requestData.empData;
      
      // 確保關鍵字串不帶前導單引號進來干擾比對，我們統一格式
      var rawEmpId = String(emp.empId).replace(/^'/, "").trim();
      var rawLineUid = String(emp.lineUid).replace(/^'/, "").trim();
      var rawPhone = String(emp.phone).replace(/^'/, "").trim();
      
      var data = hrSheet.getDataRange().getValues();
      var foundRow = -1;
      
      // 尋找是否已有該員工
      for (var i = 1; i < data.length; i++) {
        var checkId = String(data[i][0]).replace(/^'/, "").trim(); // 去除可能存在的試算表隱形單引號
        if (checkId === rawEmpId) {
          foundRow = i + 1;
          break;
        }
      }
      
      // 寫入資料時，強制在儲存格寫入加上隱形單引號 "'" 的字串，防止 Google 試算表自動轉數字
      var rowData = [
        "'" + rawEmpId,
        emp.name,
        emp.dept,
        emp.title,
        emp.arrivalDate,
        rawPhone === "無" ? "無" : "'" + rawPhone,
        "'" + rawLineUid,
        emp.status,
        emp.role
      ];
      
      if (foundRow !== -1) {
        // 更新現有員工
        hrSheet.getRange(foundRow, 1, 1, 9).setValues([rowData]);
      } else {
        // 新增全新員工
        hrSheet.appendRow(rowData);
      }
      return ContentService.createTextOutput(JSON.stringify({ status: "success" })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // 功能三：刪除員工檔案 (全新修正：補上之前漏掉的 deleteHRProfile 行為)
    if (action === "deleteHRProfile") {
      var hrSheet = ss.getSheetByName("員工檔案");
      var targetEmpId = String(requestData.empId).replace(/^'/, "").trim();
      var data = hrSheet.getDataRange().getValues();
      
      for (var i = 1; i < data.length; i++) {
        var checkId = String(data[i][0]).replace(/^'/, "").trim();
        if (checkId === targetEmpId) {
          hrSheet.deleteRow(i + 1);
          return ContentService.createTextOutput(JSON.stringify({ status: "success" })).setMimeType(ContentService.MimeType.JSON);
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "找不到該工號的員工資料" })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // 功能四：員工打卡 (上班 / 下班)
    if (action === "clockInOut") {
      var logSheet = ss.getSheetByName("打卡紀錄");
      if (!logSheet) {
        return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "找不到「打卡紀錄」工作表" })).setMimeType(ContentService.MimeType.JSON);
      }

      var now = new Date();
      var dateStr = Utilities.formatDate(now, "GMT+8", "yyyy/MM/dd");
      var timeStr = Utilities.formatDate(now, "GMT+8", "HH:mm:ss");

      var empId  = String(requestData.empId  || "").replace(/^'/, "").trim();
      var name   = String(requestData.name   || "");
      var type   = String(requestData.type   || "");   // 上班 or 下班
      var memo   = String(requestData.memo   || "");

      // 欄位順序對應試算表標題：員工編號 | 姓名 | 日期 | 時間 | 打卡類型 | 備註
      logSheet.appendRow(["'" + empId, name, dateStr, timeStr, type, memo]);

      return ContentService.createTextOutput(JSON.stringify({ status: "success" })).setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "未定義的POST動作" })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  var action = e.parameter.action;
  var lineUid = e.parameter.lineUid ? String(e.parameter.lineUid).trim() : "";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. 驗證登入與權限
  if (action === "verifyLogin") {
    var hrSheet = ss.getSheetByName("員工檔案");
    var data = hrSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var checkUid = String(data[i][6]).replace(/^'/, "").trim();
      if (checkUid === lineUid && data[i][7] === "在職") {
        return ContentService.createTextOutput(JSON.stringify({
          status: "success",
          data: { name: data[i][1], role: data[i][8] }
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "找不到該使用者" })).setMimeType(ContentService.MimeType.JSON);
  }
  
  // 2. 獲取員工清單 (解決帶入資料全字串化)
  if (action === "getHRProfiles") {
    var hrSheet = ss.getSheetByName("員工檔案");
    var data = hrSheet.getDataRange().getValues();
    var list = [];
    for (var i = 1; i < data.length; i++) {
      // 全數強制轉為文字，確保 0 開頭不會消失
      list.push({
        empId: String(data[i][0]).replace(/^'/, "").trim(),
        name: String(data[i][1]),
        dept: String(data[i][2]),
        title: String(data[i][3]),
        arrivalDate: data[i][4] ? Utilities.formatDate(new Date(data[i][4]), "GMT+8", "yyyy-MM-dd") : "",
        phone: String(data[i][5]).replace(/^'/, "").trim(),
        lineUid: String(data[i][6]).replace(/^'/, "").trim(),
        status: String(data[i][7]),
        role: String(data[i][8])
      });
    }
    return ContentService.createTextOutput(JSON.stringify({ status: "success", data: list })).setMimeType(ContentService.MimeType.JSON);
  }
  
  // 3. 獲取待審核名單
  if (action === "getPendingForms") {
    var formSheet = ss.getSheetByName("Form_Applications");
    var data = formSheet.getDataRange().getValues();
    var list = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][8] === "待審核") {
        list.push({
          formId: data[i][0],
          empId: String(data[i][1]).replace(/^'/, "").trim(),
          name: data[i][2],
          formType: data[i][3],
          subType: data[i][4],
          startTime: data[i][5],
          endTime: data[i][6],
          totalHours: data[i][7],
          reason: data[i][9]
        });
      }
    }
    return ContentService.createTextOutput(JSON.stringify({ status: "success", data: list })).setMimeType(ContentService.MimeType.JSON);
  }
  
  // 4. 獲取歷史打卡紀錄
  if (action === "getAttendanceReports") {
    var logSheet = ss.getSheetByName("打卡紀錄");  // ← 修正：對應實際工作表名稱
    var data = logSheet.getDataRange().getValues();
    var list = [];
    // 倒序排列，讓最新的打卡紀錄在最上面
    // 欄位順序：員工編號[0] | 姓名[1] | 日期[2] | 時間[3] | 打卡類型[4] | 備註[5]
    for (var i = data.length - 1; i >= 1; i--) {
      if(data[i][0]) {
        list.push({
          empId: String(data[i][0]).replace(/^'/, "").trim(),
          name:  String(data[i][1]),
          date:  String(data[i][2]),
          time:  String(data[i][3]),
          type:  String(data[i][4]),
          memo:  String(data[i][5])
        });
      }
    }
    return ContentService.createTextOutput(JSON.stringify({ status: "success", data: list })).setMimeType(ContentService.MimeType.JSON);
  }
}

// 內部輔助函數：確認是否有管理員權限
function isAdmin(lineUid) {
  if(!lineUid) return false;
  var hrSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("員工檔案");
  var data = hrSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var checkUid = String(data[i][6]).replace(/^'/, "").trim();
    if (checkUid === lineUid && data[i][7] === "在職" && data[i][8] === "Admin") {
      return true;
    }
  }
  return false;
}