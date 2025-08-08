const FORM_SPREADSHEET_ID = '1RKJoKUZuP-ByBJXyypgNfBKqoZ82Rk3_BMGnuwfOHwo';
const SHEET_NAME = '用戶資料';

// 防止電話號碼開頭0消失，這裡先用字串格式儲存
function doPost(e) {
  var sheet = SpreadsheetApp.openById(FORM_SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  
  // 確保電話號碼為字串
  var phone = e.parameter.phone;
  if(phone && !phone.startsWith('0')) {
    phone = '0' + phone;  // 如果沒開頭0，強制加回（依需求調整）
  }
  
  var data = [
    e.parameter.parentName,
    phone,
    e.parameter.email,
    e.parameter.childAge,
    e.parameter.ipRegion,
    '', // 訂閱起始時間欄位，初始為空
    '', // 最新訂單號碼欄位，初始為空
    '', // 當月扣款狀態欄位，初始為空
    '', // 最後更新時間欄位，初始為空
    '', // 最後通知時間欄位，初始為空
    generateRandomKey() // 動態生成金鑰欄位
  ];
  
  sheet.appendRow(data);
  return ContentService.createTextOutput("OK");
}

function generateRandomKey() {
  var chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  var result = '';
  for (var i = 0; i < 5; i++) {
    result += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return result;
}
// 主函數：處理未讀郵件並更新表單資料
function processUnreadEmailsAndUpdateForm() {
  var today = new Date();
  today.setHours(0, 0, 0, 0);
  
  var yesterday = new Date(today);
  yesterday.setDate(yesterday.getDate() - 1);
  
  var formattedYesterday = Utilities.formatDate(yesterday, Session.getScriptTimeZone(), 'yyyy/MM/dd');
  
  // 搜尋未讀的收款成功郵件
  var threads = GmailApp.search('is:unread after:' + formattedYesterday + ' subject:"收款成功"');
  
  Logger.log('找到 ' + threads.length + ' 個收款成功未讀郵件串');
  
  threads.forEach(function(thread) {
    var messages = thread.getMessages();
    
    messages.forEach(function(message) {
      if (message.isUnread()) {
        var subject = message.getSubject();
        var body = message.getPlainBody();
        Logger.log('處理郵件，主旨：' + subject);
        
        // 檢查郵件是否符合處理條件
        if (subject.includes("收款成功") && 
            body.includes("收款成功通知") && 
            body.includes("訂購明細") && 
            body.includes("AI英語伴讀每月499方案")) {
          
          processEmailAndUpdateForm(body, message.getDate());
        } else {
          Logger.log('郵件不符合處理條件');
        }
        
        // 無論郵件是否符合處理條件，都將其標記為已讀
        message.markRead();
      }
    });
  });
  
  // 搜尋未讀的信用卡授權成功郵件並自動標記已讀
  var authThreads = GmailApp.search('is:unread after:' + formattedYesterday + ' subject:"信用卡授權成功"');
  
  Logger.log('找到 ' + authThreads.length + ' 個信用卡授權成功未讀郵件串');
  
  authThreads.forEach(function(thread) {
    var messages = thread.getMessages();
    
    messages.forEach(function(message) {
      if (message.isUnread()) {
        var subject = message.getSubject();
        var body = message.getPlainBody();
        Logger.log('處理信用卡授權郵件，主旨：' + subject);
        
        // 檢查是否包含AI英語伴讀方案，如果是則標記為已讀（因為收款成功郵件已處理過）
        if (body.includes("AI英語伴讀每月499方案")) {
          Logger.log('發現重複的信用卡授權郵件，標記為已讀');
        } else {
          Logger.log('信用卡授權郵件不符合處理條件');
        }
        
        // 無論如何都標記為已讀，避免累積
        message.markRead();
      }
    });
  });
  
  // 搜尋未讀的收款失敗郵件
  var failureThreads = GmailApp.search('is:unread after:' + formattedYesterday + ' (subject:"收款失敗" OR subject:"扣款失敗" OR subject:"付款失敗")');
  
  Logger.log('找到 ' + failureThreads.length + ' 個收款失敗未讀郵件串');
  
  failureThreads.forEach(function(thread) {
    var messages = thread.getMessages();
    
    messages.forEach(function(message) {
      if (message.isUnread()) {
        var subject = message.getSubject();
        var body = message.getPlainBody();
        Logger.log('處理失敗郵件，主旨：' + subject);
        
        // 檢查郵件是否符合處理條件
        if ((subject.includes("收款失敗") || subject.includes("扣款失敗") || subject.includes("付款失敗")) && 
            body.includes("AI英語伴讀每月499方案")) {
          
          processFailureEmailAndUpdateForm(body, message.getDate());
        } else {
          Logger.log('失敗郵件不符合處理條件');
        }
        
        // 無論郵件是否符合處理條件，都將其標記為已讀
        message.markRead();
      }
    });
  });
  
  // 檢查逾期未扣款成功的用戶並發送提醒給管理員
  checkOverdueUsers();
  
  // 發送成功繳費用戶清單給管理員
function sendSuccessfulPaymentsListToAdmin(successfulUsers, paymentDate) {
  try {
    var subject = 'AI英語伴讀 ' + paymentDate + ' 成功繳費用戶清單（開發票用）';
    
    // 建立表格格式的用戶清單
    var tableHeader = '姓名\t電話\t電子信箱\t訂單號碼';
    var userRows = successfulUsers.map(function(user) {
      return user.name + '\t' + user.phone + '\t' + user.email + '\t' + user.orderNumber;
    });
    
    var userTable = tableHeader + '\n' + userRows.join('\n');
    
    // 建立HTML格式的表格
    var htmlTable = `
    <table border="1" cellpadding="8" cellspacing="0" style="border-collapse: collapse; width: 100%; font-family: Arial, sans-serif;">
      <thead>
        <tr style="background-color: #4CAF50; color: white;">
          <th>序號</th>
          <th>姓名</th>
          <th>電話</th>
          <th>電子信箱</th>
          <th>訂單號碼</th>
        </tr>
      </thead>
      <tbody>
        ${successfulUsers.map(function(user, index) {
          return `
          <tr style="background-color: ${index % 2 === 0 ? '#f2f2f2' : 'white'};">
            <td style="text-align: center;">${index + 1}</td>
            <td>${user.name}</td>
            <td>${user.phone}</td>
            <td>${user.email}</td>
            <td>${user.orderNumber}</td>
          </tr>
          `;
        }).join('')}
      </tbody>
    </table>
    `;
    
    var htmlBody = `
    <html>
      <head>
        <style>
          body {
            font-family: Arial, sans-serif;
            line-height: 1.6;
            color: #333;
            margin: 20px;
          }
          .summary {
            background-color: #e8f5e8;
            padding: 15px;
            border-radius: 5px;
            margin-bottom: 20px;
          }
          .note {
            background-color: #fff3cd;
            padding: 10px;
            border-radius: 5px;
            margin-top: 20px;
            font-size: 14px;
          }
        </style>
      </head>
      <body>
        <h2>AI英語伴讀 ${paymentDate} 成功繳費用戶清單</h2>
        
        <div class="summary">
          <strong>📊 統計摘要：</strong><br>
          • 繳費日期：${paymentDate}<br>
          • 成功繳費人數：${successfulUsers.length} 人<br>
          • 每人金額：NT$ 499<br>
          • 總金額：NT$ ${successfulUsers.length * 499}<br>
        </div>
        
        <h3>📋 用戶明細清單</h3>
        ${htmlTable}
        
        <div class="note">
          <strong>💡 使用說明：</strong><br>
          • 此清單為當日成功扣款的用戶<br>
          • 可直接複製表格內容到Excel或其他系統<br>
          • 建議核對訂單號碼與金流系統資料<br>
          • 系統已自動記錄通知時間，避免重複發送<br>
          • 如有疑問請查看系統日誌或聯繫技術人員
        </div>
      </body>
    </html>
    `;
    
    var plainTextBody = `
管理員您好：

以下是 ${paymentDate} 成功繳費的用戶清單，共 ${successfulUsers.length} 位用戶：

${userTable}

統計摘要：
• 繳費日期：${paymentDate}
• 成功繳費人數：${successfulUsers.length} 人
• 每人金額：NT$ 499
• 總金額：NT$ ${successfulUsers.length * 499}

請依此清單開立發票。

系統自動通知
萬里遊科技股份有限公司
    `;
    
    MailApp.sendEmail({
      to: "info.aifunschool@gmail.com",
      subject: subject,
      body: plainTextBody,
      htmlBody: htmlBody
    });
    
    Logger.log('成功繳費清單已發送給管理員');
  } catch (error) {
    Logger.log('發送成功繳費清單郵件時發生錯誤: ' + error.toString());
  }
}

function sendOverdueNotificationToAdmin(overdueUsers) {
  try {
    var subject = 'AI英語伴讀逾期未扣款用戶清單';
    var userList = overdueUsers.map(function(user) {
      return '• ' + user.name + ' (' + user.email + ') - 逾期 ' + user.daysSinceLastPayment + ' 天，最後成功扣款：' + 
             Utilities.formatDate(user.lastPaymentDate, Session.getScriptTimeZone(), 'yyyy/MM/dd');
    }).join('\n');
    
    var body = `
管理員您好：

以下是逾期超過一個月未成功扣款的用戶清單：

${userList}

共計 ${overdueUsers.length} 位用戶。

這些用戶可能需要人工確認扣款狀態或聯繫客戶。

系統自動通知
萬里遊科技股份有限公司
    `;
    
    MailApp.sendEmail({
      to: "info.aifunschool@gmail.com",
      subject: subject,
      body: body
    });
    
    Logger.log('逾期用戶清單已發送給管理員');
  } catch (error) {
    Logger.log('發送管理員通知時發生錯誤: ' + error.toString());
  }
}

// 發送錯誤通知給管理員
function sendErrorNotification(subject, errorDetails) {
  try {
    MailApp.sendEmail({
      to: "info.aifunschool@gmail.com",
      subject: "訂閱管理系統錯誤: " + subject,
      body: "錯誤詳情:\n" + errorDetails + "\n\n時間: " + new Date().toString()
    });
  } catch (error) {
    Logger.log('發送錯誤通知失敗: ' + error.toString());
  }
}

// 設置觸發器 - 每5分鐘執行一次
function setTrigger() {
  // 刪除現有的觸發器
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'processUnreadEmailsAndUpdateForm') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // 創建新的觸發器
  ScriptApp.newTrigger('processUnreadEmailsAndUpdateForm')
    .timeBased()
    .everyMinutes(5)
    .create();
    
  Logger.log('觸發器已設置完成');
}

// 設置每日檢查觸發器 - 每天早上9點檢查逾期用戶
function setDailyCheckTrigger() {
  // 刪除現有的每日檢查觸發器
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'checkOverdueUsers') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // 創建每日檢查觸發器
  ScriptApp.newTrigger('checkOverdueUsers')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();
    
  Logger.log('每日逾期檢查觸發器已設置完成');
}

// 手動測試函數
function testSystem() {
  Logger.log('開始測試系統...');
  processUnreadEmailsAndUpdateForm();
  Logger.log('測試完成');
}

// 手動檢查逾期用戶
function manualCheckOverdue() {
  Logger.log('開始手動檢查逾期用戶...');
  checkOverdueUsers();
  Logger.log('檢查完成');
}

// 手動發送當日成功繳費清單
function manualSendTodaySuccessfulPayments() {
  Logger.log('開始手動發送當日成功繳費清單...');
  sendMonthlySuccessfulPaymentsToAdmin();
  Logger.log('發送完成');
}

// 查看表單資料狀態
function checkFormDataStatus() {
  try {
    var sheet = SpreadsheetApp.openById(FORM_SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();
    var currentDate = new Date();
    
    Logger.log('表單總記錄數：' + (values.length - 1));
    
    var subscribedCount = 0;
    var recentSuccessCount = 0;
    var recentFailureCount = 0;
    var overdueCount = 0;
    
    for (var i = 1; i < values.length; i++) {
      var subscriptionStartTime = values[i][5];
      var paymentStatus = values[i][7];
      var dynamicKey = values[i][10];
      
      if (subscriptionStartTime && subscriptionStartTime !== '') {
        subscribedCount++;
        
        if (paymentStatus) {
          if (paymentStatus.includes('成功') && isRecentMonth(paymentStatus, currentDate)) {
            recentSuccessCount++;
          } else if (paymentStatus.includes('失敗')) {
            recentFailureCount++;
          } else if (paymentStatus.includes('逾期未成功')) {
            overdueCount++;
          }
        }
        
        // 檢查是否有動態金鑰
        if (!dynamicKey || dynamicKey === '') {
          Logger.log('用戶（行 ' + (i + 1) + '）' + values[i][0] + ' 缺少動態金鑰');
        }
      }
    }
    
    Logger.log('已訂閱用戶數：' + subscribedCount);
    Logger.log('近期成功扣款：' + recentSuccessCount);
    Logger.log('扣款失敗：' + recentFailureCount);
    Logger.log('逾期未成功：' + overdueCount);
    
  } catch (error) {
    Logger.log('檢查表單資料時發生錯誤: ' + error.toString());
  }
}

// 為現有用戶補充動態金鑰
function addMissingDynamicKeys() {
  try {
    var sheet = SpreadsheetApp.openById(FORM_SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();
    var updatedCount = 0;
    
    for (var i = 1; i < values.length; i++) {
      var dynamicKey = values[i][10]; // K欄
      
      if (!dynamicKey || dynamicKey === '') {
        var newKey = generateRandomKey();
        sheet.getRange(i + 1, 11).setValue(newKey); // K欄
        updatedCount++;
        Logger.log('為用戶（行 ' + (i + 1) + '）' + values[i][0] + ' 生成動態金鑰：' + newKey);
      }
    }
    
    Logger.log('共為 ' + updatedCount + ' 位用戶補充了動態金鑰');
    
  } catch (error) {
    Logger.log('補充動態金鑰時發生錯誤: ' + error.toString());
  }
}

// 清理重複處理記錄（維護用）
function cleanupDuplicateRecords() {
  try {
    var sheet = SpreadsheetApp.openById(FORM_SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();
    var cleanedCount = 0;
    
    Logger.log('開始清理重複處理記錄...');
    
    for (var i = 1; i < values.length; i++) {
      var paymentStatus = values[i][7] ? values[i][7].toString() : '';
      
      if (paymentStatus) {
        // 檢查是否有重複的同月份記錄
        var statusParts = paymentStatus.split(',');
        var uniqueStatuses = [];
        var monthsProcessed = [];
        
        statusParts.forEach(function(status) {
          var trimmedStatus = status.trim();
          var dateMatch = trimmedStatus.match(/\((\d{4}\/\d{2})(?:\/(\d{2}))?\)/);
          if (dateMatch) {
            // 若只有 yyyy/MM，補齊為 yyyy/MM/01
            var fullDate = dateMatch[2] ? dateMatch[1] + '/' + dateMatch[2] : dateMatch[1] + '/01';
            var normalizedStatus = trimmedStatus.replace(/\((\d{4}\/\d{2})(?:\/\d{2})?\)/, '(' + fullDate + ')');
            var month = fullDate.substring(0, 7); // 取年月部分
            if (monthsProcessed.indexOf(month) === -1) {
              monthsProcessed.push(month);
              uniqueStatuses.push(normalizedStatus);
            }
          } else {
            uniqueStatuses.push(trimmedStatus);
          }
        });

        var cleanedStatus = uniqueStatuses.join(', ');
        if (cleanedStatus !== paymentStatus) {
          sheet.getRange(i + 1, 8).setValue(cleanedStatus);
          cleanedCount++;
          Logger.log('清理用戶（行 ' + (i + 1) + '）重複記錄（已統一日期格式）');
        }
      }
    }
    
    Logger.log('共清理了 ' + cleanedCount + ' 筆重複記錄');
    
  } catch (error) {
    Logger.log('清理重複記錄時發生錯誤: ' + error.toString());
  }
}

// 初始化所有觸發器
function initializeSystem() {
  setTrigger();
  setDailyCheckTrigger();
  Logger.log('系統初始化完成，已設置所有觸發器');
  // 當月成功繳費用戶清單給管理員（用於開發票）
  sendMonthlySuccessfulPaymentsToAdmin();
}

// 處理郵件內容並更新表單資料
function processEmailAndUpdateForm(body, emailDate) {
  try {
    var orderInfo = extractOrderInfo(body);
    
    if (orderInfo) {
      Logger.log('提取到的訂單資訊：' + JSON.stringify(orderInfo));
      
      // 在表單資料中查找匹配的記錄
      var matchingRows = findMatchingFormRecords(orderInfo);
      
      if (matchingRows.length > 0) {
        // 檢查是否同月份重複處理
        var isDuplicate = checkDuplicateProcessing(matchingRows, emailDate);
        
        if (!isDuplicate) {
          updateFormRecords(matchingRows, emailDate, orderInfo.ecpayOrderNumber, '成功');
          Logger.log('成功更新 ' + matchingRows.length + ' 筆表單記錄為扣款成功');
        } else {
          Logger.log('檢測到同月份重複處理，跳過更新');
        }
      } else {
        Logger.log('未找到匹配的表單記錄');
        sendErrorNotification('未找到匹配記錄', JSON.stringify(orderInfo));
      }
    } else {
      Logger.log('無法從郵件中提取所需資料');
      sendErrorNotification('無法提取郵件資料', body);
    }
  } catch (error) {
    Logger.log('處理郵件時發生錯誤: ' + error.toString());
    sendErrorNotification('處理郵件錯誤', error.toString());
  }
}

// 處理失敗郵件內容並更新表單資料
function processFailureEmailAndUpdateForm(body, emailDate) {
  try {
    var orderInfo = extractOrderInfo(body);
    
    if (orderInfo) {
      Logger.log('提取到的失敗訂單資訊：' + JSON.stringify(orderInfo));
      
      // 在表單資料中查找匹配的記錄
      var matchingRows = findMatchingFormRecords(orderInfo);
      
      if (matchingRows.length > 0) {
        // 檢查是否同月份重複處理
        var isDuplicate = checkDuplicateProcessing(matchingRows, emailDate);
        
        if (!isDuplicate) {
          updateFormRecords(matchingRows, emailDate, orderInfo.ecpayOrderNumber || 'N/A', '失敗');
          Logger.log('成功更新 ' + matchingRows.length + ' 筆表單記錄為扣款失敗');
        } else {
          Logger.log('檢測到同月份重複處理失敗郵件，跳過更新');
        }
      } else {
        Logger.log('未找到匹配的失敗記錄');
        sendErrorNotification('未找到失敗記錄匹配', JSON.stringify(orderInfo));
      }
    } else {
      Logger.log('無法從失敗郵件中提取所需資料');
      sendErrorNotification('無法提取失敗郵件資料', body);
    }
  } catch (error) {
    Logger.log('處理失敗郵件時發生錯誤: ' + error.toString());
    sendErrorNotification('處理失敗郵件錯誤', error.toString());
  }
}

// 檢查是否同月份重複處理
function checkDuplicateProcessing(matchingRows, emailDate) {
  try {
    var emailMonth = Utilities.formatDate(emailDate, Session.getScriptTimeZone(), 'yyyy/MM');
    
    for (var i = 0; i < matchingRows.length; i++) {
      var paymentStatus = matchingRows[i].data[7] ? matchingRows[i].data[7].toString() : ''; // 當月扣款狀態
      
      // 檢查是否已經有相同月份的處理記錄
      if (paymentStatus.includes('(' + emailMonth + ')')) {
        Logger.log('發現同月份重複記錄，用戶：' + matchingRows[i].data[0] + '，月份：' + emailMonth);
        return true;
      }
    }
    
    return false;
  } catch (error) {
    Logger.log('檢查重複處理時發生錯誤: ' + error.toString());
    return false;
  }
}

// 從郵件內容中提取訂單信息
function extractOrderInfo(body) {
  var lines = body.split('\n');
  var orderInfo = {};
  var inShippingInfo = false;
  
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].trim();
    
    if (line.startsWith("*綠界訂單編號：*")) {
      orderInfo.ecpayOrderNumber = line.split("*")[2].trim();
    } else if (line.startsWith("*E-mail：*")) {
      orderInfo.payerEmail = line.split("*")[2].trim();
    } else if (line === "*出貨資訊*") {
      inShippingInfo = true;
    } else if (inShippingInfo) {
      if (line.startsWith("*姓名：*")) {
        orderInfo.recipientName = line.split("*")[2].split(" ")[0].trim();
      } else if (line.startsWith("*電話：*")) {
        orderInfo.recipientPhone = line.split("*")[2].trim();
        break;
      }
    }
  }
  
  // 驗證是否提取到必要資訊（失敗郵件可能沒有訂單編號）
  if (orderInfo.payerEmail && orderInfo.recipientName && orderInfo.recipientPhone) {
    Logger.log('成功提取基本所需資訊');
    return orderInfo;
  } else {
    Logger.log('無法提取基本所需資訊。已提取的資訊：' + JSON.stringify(orderInfo));
    return null;
  }
}

// 在表單資料中尋找匹配的記錄
function findMatchingFormRecords(orderInfo) {
  try {
    var sheet = SpreadsheetApp.openById(FORM_SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();
    var matchingRows = [];
    
    // 資料結構：A欄-家長姓名, B欄-聯絡電話, C欄-Email, D欄-兒童年齡, E欄-IP地區, 
    // F欄-訂閱起始時間, G欄-最新訂單號碼, H欄-當月扣款狀態, I欄-最後更新時間, J欄-最後通知時間, K欄-動態金鑰
    for (var i = 1; i < values.length; i++) { // 跳過標題行
      var row = values[i];
      var parentName = row[0] ? row[0].toString().trim() : '';
      var phone = row[1] ? row[1].toString().trim() : '';
      var email = row[2] ? row[2].toString().trim() : '';
      
      // 比對家長姓名、聯絡電話、Email
      var nameMatch = parentName === orderInfo.recipientName;
      var phoneMatch = phone === orderInfo.recipientPhone;
      var emailMatch = email.toLowerCase() === orderInfo.payerEmail.toLowerCase();
      
      if (nameMatch && phoneMatch && emailMatch) {
        matchingRows.push({
          rowIndex: i + 1, // Google Sheets 的行號從1開始
          data: row
        });
        Logger.log('找到匹配記錄，行號：' + (i + 1) + '，資料：' + JSON.stringify({
          name: parentName,
          phone: phone,
          email: email
        }));
      }
    }
    
    return matchingRows;
  } catch (error) {
    Logger.log('查找匹配記錄時發生錯誤: ' + error.toString());
    throw error;
  }
}

// 更新表單記錄
function updateFormRecords(matchingRows, paymentDate, orderNumber, status) {
  try {
    var sheet = SpreadsheetApp.openById(FORM_SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    var currentDate = new Date();
    var formattedUpdateTime = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
    var currentMonthDay = Utilities.formatDate(paymentDate, Session.getScriptTimeZone(), 'yyyy/MM/dd');
    
    matchingRows.forEach(function(matchingRow) {
      var subscriptionStartTime = matchingRow.data[5]; // F欄-訂閱起始時間
      
      // 如果是首次訂閱（沒有訂閱起始時間），設置起始時間
      if (!subscriptionStartTime || subscriptionStartTime === '') {
        var formattedStartTime = Utilities.formatDate(paymentDate, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
        sheet.getRange(matchingRow.rowIndex, 6).setValue(formattedStartTime); // F欄
        Logger.log('設置訂閱起始時間為：' + formattedStartTime);
      }
      
      // 更新最新訂單號碼（G欄）
      sheet.getRange(matchingRow.rowIndex, 7).setValue(orderNumber);
      
      // 更新當月扣款狀態（H欄）- 精確到年月日
      sheet.getRange(matchingRow.rowIndex, 8).setValue(status + ' (' + currentMonthDay + ')');
      
      // 更新最後更新時間（I欄）
      sheet.getRange(matchingRow.rowIndex, 9).setValue(formattedUpdateTime);
      
      // 如果沒有動態金鑰，生成一個
      if (!matchingRow.data[10] || matchingRow.data[10] === '') {
        sheet.getRange(matchingRow.rowIndex, 11).setValue(generateRandomKey());
        Logger.log('為用戶生成動態金鑰');
      }
      
      Logger.log('已更新行號 ' + matchingRow.rowIndex + ' 的扣款狀態為' + status + '，訂單號：' + orderNumber);
      
      // 只有扣款成功時才發送確認郵件
      if (status === '成功') {
        var userEmail = matchingRow.data[2]; // Email 欄位
        var userName = matchingRow.data[0];  // 家長姓名
        sendPaymentSuccessEmail(userEmail, userName, currentMonthDay, orderNumber);
      }
    });
    
  } catch (error) {
    Logger.log('更新表單記錄時發生錯誤: ' + error.toString());
    throw error;
  }
}

// 檢查逾期未扣款成功的用戶（距離上次扣款成功超過一個月）
function checkOverdueUsers() {
  try {
    var sheet = SpreadsheetApp.openById(FORM_SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();
    var currentDate = new Date();
    var overdueUsers = [];
    
    // 檢查每個用戶的扣款狀態
    for (var i = 1; i < values.length; i++) { // 跳過標題行
      var row = values[i];
      var parentName = row[0] ? row[0].toString().trim() : '';
      var email = row[2] ? row[2].toString().trim() : '';
      var subscriptionStartTime = row[5]; // 訂閱起始時間
      var paymentStatus = row[7] ? row[7].toString() : ''; // 當月扣款狀態
      var lastUpdateTime = row[8]; // 最後更新時間
      
      // 只檢查已訂閱的用戶（有訂閱起始時間）
      if (subscriptionStartTime && subscriptionStartTime !== '') {
        var lastSuccessfulPayment = getLastSuccessfulPaymentDate(paymentStatus, subscriptionStartTime);
        
        if (lastSuccessfulPayment) {
          var daysSinceLastPayment = Math.floor((currentDate - lastSuccessfulPayment) / (1000 * 60 * 60 * 24));
          
          // 如果距離上次成功扣款超過35天（約一個月多5天緩衝）
          if (daysSinceLastPayment > 35) {
            // 檢查最近是否有失敗記錄或者完全沒有記錄
            var hasRecentFailure = paymentStatus.includes('失敗');
            var noRecentSuccess = !paymentStatus.includes('成功') || 
                                  !isRecentMonth(paymentStatus, currentDate);
            
            if (hasRecentFailure || noRecentSuccess) {
              overdueUsers.push({
                rowIndex: i + 1,
                name: parentName,
                email: email,
                daysSinceLastPayment: daysSinceLastPayment,
                lastPaymentDate: lastSuccessfulPayment,
                data: row
              });
              
              // 更新扣款狀態為逾期未成功
              var currentMonth = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), 'yyyy/MM/dd');
              sheet.getRange(i + 1, 8).setValue('逾期未成功 (' + currentMonth + ')');
              
              Logger.log('發現逾期未扣款用戶：' + parentName + ' (' + email + ')，距離上次成功扣款 ' + daysSinceLastPayment + ' 天');
            }
          }
        }
      }
    }
    
    // 只發送通知給管理員
    if (overdueUsers.length > 0) {
      sendOverdueNotificationToAdmin(overdueUsers);
      Logger.log('共發現 ' + overdueUsers.length + ' 位用戶逾期未成功扣款');
    } else {
      Logger.log('目前沒有用戶逾期未扣款');
    }
    
  } catch (error) {
    Logger.log('檢查逾期用戶時發生錯誤: ' + error.toString());
    sendErrorNotification('檢查逾期用戶錯誤', error.toString());
  }
}

// 獲取最後一次成功扣款日期
function getLastSuccessfulPaymentDate(paymentStatus, subscriptionStartTime) {
  try {
    if (!paymentStatus || paymentStatus === '') {
      return new Date(subscriptionStartTime);
    }
    
    // 從扣款狀態中提取最後成功的日期
    var successMatch = paymentStatus.match(/成功 \((\d{4}\/\d{2}\/\d{2})\)/);
    if (successMatch) {
      return new Date(successMatch[1]);
    }
    
    // 舊格式兼容：提取月份格式
    var monthMatch = paymentStatus.match(/成功 \((\d{4}\/\d{2})\)/);
    if (monthMatch) {
      var monthStr = monthMatch[1] + '/01';
      return new Date(monthStr);
    }
    
    // 如果沒有成功記錄，返回訂閱開始時間
    return new Date(subscriptionStartTime);
  } catch (error) {
    Logger.log('解析最後扣款日期時發生錯誤: ' + error.toString());
    return new Date(subscriptionStartTime);
  }
}

// 檢查是否為近期月份的成功扣款
function isRecentMonth(paymentStatus, currentDate) {
  try {
    // 取得本月第一天與上月第一天
    var thisMonth = new Date(currentDate.getFullYear(), currentDate.getMonth(), 1);
    var lastMonth = new Date(currentDate.getFullYear(), currentDate.getMonth() - 1, 1);

    // 取出所有成功 (yyyy/MM/dd) 的日期
    var matches = paymentStatus.match(/成功 \((\d{4}\/\d{2}\/\d{2})\)/g);
    if (!matches) return false;

    for (var i = 0; i < matches.length; i++) {
      var dateStr = matches[i].match(/\((\d{4}\/\d{2}\/\d{2})\)/)[1];
      var payDate = new Date(dateStr);
      if (payDate >= lastMonth) {
        return true;
      }
    }
    return false;
  } catch (error) {
    return false;
  }
}

// 發送扣款成功確認郵件給用戶
function sendPaymentSuccessEmail(email, name, paymentDate, orderNumber) {
  try {
    var subject = 'AI英語伴讀 ' + paymentDate + ' 扣款成功通知 - 萬里遊科技';
    var htmlBody = `
    <html>
      <head>
        <style>
          body {
            font-family: 'Arial', sans-serif;
            line-height: 1.6;
            color: #333;
            background-color: #f4f4f4;
            margin: 0;
            padding: 0;
          }
          .container {
            max-width: 600px;
            margin: 50px auto;
            padding: 20px;
            background-color: #fff;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
          }
          .header {
            background-color: #4CAF50;
            color: white;
            padding: 20px;
            text-align: center;
            border-top-left-radius: 8px;
            border-top-right-radius: 8px;
          }
          .content {
            padding: 20px;
          }
          .success-icon {
            color: #4CAF50;
            font-size: 24px;
            text-align: center;
            margin: 20px 0;
          }
          .footer {
            text-align: center;
            margin-top: 20px;
            font-size: 12px;
            color: #777;
          }
        </style>
      </head>
      <body>
        <div class="container">
          <div class="header">
            <h1>扣款成功通知</h1>
          </div>
          <div class="content">
            <div class="success-icon">✅</div>
            <p>親愛的 ${name} 您好：</p>
            <p>您的 AI英語伴讀每月499方案 ${paymentDate} 扣款已成功完成！</p>
            <p><strong>訂單號碼：</strong> ${orderNumber}</p>
            <p><strong>扣款日期：</strong> ${paymentDate}</p>
            <p>服務將持續為您提供，如有任何問題請隨時聯繫我們。</p>
            <p>感謝您的支持！</p>
          </div>
          <div class="footer">
            <p>萬里遊科技股份有限公司 敬上</p>
          </div>
        </div>
      </body>
    </html>
    `;
    
    MailApp.sendEmail({
      to: email,
      subject: subject,
      htmlBody: htmlBody
    });
    
    Logger.log('扣款成功通知郵件已發送給：' + email);
  } catch (error) {
    Logger.log('發送扣款成功郵件時發生錯誤: ' + error.toString());
  }
}

// 發送當月成功繳費用戶清單給管理員（用於開發票）
function sendMonthlySuccessfulPaymentsToAdmin() {
  try {
    var sheet = SpreadsheetApp.openById(FORM_SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();
    var currentDate = new Date();
    var currentDateStr = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), 'yyyy/MM/dd');
    var successfulUsers = [];
    
    // 檢查每個用戶的扣款狀態
    for (var i = 1; i < values.length; i++) { // 跳過標題行
      var row = values[i];
      var parentName = row[0] ? row[0].toString().trim() : '';
      var phone = row[1] ? row[1].toString().trim() : '';
      var email = row[2] ? row[2].toString().trim() : '';
      var subscriptionStartTime = row[5]; // 訂閱起始時間
      var paymentStatus = row[7] ? row[7].toString() : ''; // 當月扣款狀態
      var lastOrderNumber = row[6] ? row[6].toString() : ''; // 最新訂單號碼
      var lastNotifyTime = row[9] ? row[9].toString() : ''; // 最後通知時間
      var lastUpdateTime = row[8] ? row[8].toString() : ''; // 最後更新時間
      
      // 只檢查已訂閱且當天成功扣款的用戶
      if (subscriptionStartTime && subscriptionStartTime !== '') {
        if (paymentStatus.includes('成功 (' + currentDateStr + ')')) {
          // 若最後通知時間為空，直接寫入當前時間並允許本次通知
          if (!lastNotifyTime) {
            sheet.getRange(i + 1, 10).setValue(currentDateStr);
            lastNotifyTime = currentDateStr;
          }
          // 僅當「最後更新時間」比「最後通知時間」新，才發送通知
          if (lastNotifyTime && lastUpdateTime && lastNotifyTime >= lastUpdateTime) {
            Logger.log('用戶 ' + parentName + ' 今日已通知過管理員或無新異動，跳過');
            continue;
          }
          
          successfulUsers.push({
            rowIndex: i + 1,
            name: parentName,
            phone: phone,
            email: email,
            orderNumber: lastOrderNumber,
            paymentDate: currentDateStr
          });
          
          // 更新最後通知時間
          sheet.getRange(i + 1, 10).setValue(currentDateStr);
        }
      }
    }
    
    // 發送成功繳費清單給管理員
    if (successfulUsers.length > 0) {
      sendSuccessfulPaymentsListToAdmin(successfulUsers, currentDateStr);
      Logger.log('當日成功繳費用戶清單已發送給管理員，共 ' + successfulUsers.length + ' 位用戶');
    } else {
      Logger.log('今日目前沒有新的成功繳費用戶需要通知');
    }
    
  } catch (error) {
    Logger.log('發送成功繳費清單時發生錯誤: ' + error.toString());
    sendErrorNotification('發送成功繳費清單錯誤', error.toString());
  }
}

// 發送成功繳費用戶清單給管理員
function sendSuccessfulPaymentsListToAdmin(successfulUsers, month) {
  try {
    var subject = 'AI英語伴讀 ' + month + ' 月成功繳費用戶清單（開發票用）';
    
    // 建立表格格式的用戶清單
    var tableHeader = '姓名\t電話\t電子信箱\t訂單號碼';
    var userRows = successfulUsers.map(function(user) {
      return user.name + '\t' + user.phone + '\t' + user.email + '\t' + user.orderNumber;
    });
    
    var userTable = tableHeader + '\n' + userRows.join('\n');
    
    // 建立HTML格式的表格
    var htmlTable = `
    <table border="1" cellpadding="8" cellspacing="0" style="border-collapse: collapse; width: 100%; font-family: Arial, sans-serif;">
      <thead>
        <tr style="background-color: #4CAF50; color: white;">
          <th>序號</th>
          <th>姓名</th>
          <th>電話</th>
          <th>電子信箱</th>
          <th>訂單號碼</th>
        </tr>
      </thead>
      <tbody>
        ${successfulUsers.map(function(user, index) {
          return `
          <tr style="background-color: ${index % 2 === 0 ? '#f2f2f2' : 'white'};">
            <td style="text-align: center;">${index + 1}</td>
            <td>${user.name}</td>
            <td>${user.phone}</td>
            <td>${user.email}</td>
            <td>${user.orderNumber}</td>
          </tr>
          `;
        }).join('')}
      </tbody>
    </table>
    `;
    
    var htmlBody = `
    <html>
      <head>
        <style>
          body {
            font-family: Arial, sans-serif;
            line-height: 1.6;
            color: #333;
            margin: 20px;
          }
          .summary {
            background-color: #e8f5e8;
            padding: 15px;
            border-radius: 5px;
            margin-bottom: 20px;
          }
          .note {
            background-color: #fff3cd;
            padding: 10px;
            border-radius: 5px;
            margin-top: 20px;
            font-size: 14px;
          }
        </style>
      </head>
      <body>
        <h2>AI英語伴讀 ${month} 月成功繳費用戶清單</h2>
        
        <div class="summary">
          <strong>📊 統計摘要：</strong><br>
          • 繳費月份：${month}<br>
          • 成功繳費人數：${successfulUsers.length} 人<br>
          • 每人金額：NT$ 499<br>
          • 總金額：NT$ ${successfulUsers.length * 499}<br>
        </div>
        
        <h3>📋 用戶明細清單</h3>
        ${htmlTable}
        
        <div class="note">
          <strong>💡 使用說明：</strong><br>
          • 此清單為本月成功扣款的所有用戶<br>
          • 可直接複製表格內容到Excel或其他系統<br>
          • 建議核對訂單號碼與金流系統資料<br>
          • 如有疑問請查看系統日誌或聯繫技術人員
        </div>
      </body>
    </html>
    `;
    
    var plainTextBody = `
管理員您好：

以下是 ${month} 月份成功繳費的用戶清單，共 ${successfulUsers.length} 位用戶：

${userTable}

統計摘要：
• 繳費月份：${month}
• 成功繳費人數：${successfulUsers.length} 人
• 每人金額：NT$ 499
• 總金額：NT$ ${successfulUsers.length * 499}

請依此清單開立發票。

系統自動通知
萬里遊科技股份有限公司
    `;
    
    MailApp.sendEmail({
      to: "info.aifunschool@gmail.com",
      subject: subject,
      body: plainTextBody,
      htmlBody: htmlBody
    });
    
    Logger.log('成功繳費清單已發送給管理員');
  } catch (error) {
    Logger.log('發送成功繳費清單郵件時發生錯誤: ' + error.toString());
  }
}
function sendOverdueNotificationToAdmin(overdueUsers) {
  try {
    var subject = 'AI英語伴讀逾期未扣款用戶清單';
    var userList = overdueUsers.map(function(user) {
      return '• ' + user.name + ' (' + user.email + ') - 逾期 ' + user.daysSinceLastPayment + ' 天，最後成功扣款：' + 
             Utilities.formatDate(user.lastPaymentDate, Session.getScriptTimeZone(), 'yyyy/MM/dd');
    }).join('\n');
    
    var body = `
管理員您好：

以下是逾期超過一個月未成功扣款的用戶清單：

${userList}

共計 ${overdueUsers.length} 位用戶。

這些用戶可能需要人工確認扣款狀態或聯繫客戶。

系統自動通知
萬里遊科技股份有限公司
    `;
    
    MailApp.sendEmail({
      to: "info.aifunschool@gmail.com",
      subject: subject,
      body: body
    });
    
    Logger.log('逾期用戶清單已發送給管理員');
  } catch (error) {
    Logger.log('發送管理員通知時發生錯誤: ' + error.toString());
  }
}

// 發送錯誤通知給管理員
function sendErrorNotification(subject, errorDetails) {
  try {
    MailApp.sendEmail({
      to: "info.aifunschool@gmail.com",
      subject: "訂閱管理系統錯誤: " + subject,
      body: "錯誤詳情:\n" + errorDetails + "\n\n時間: " + new Date().toString()
    });
  } catch (error) {
    Logger.log('發送錯誤通知失敗: ' + error.toString());
  }
}

// 設置觸發器 - 每5分鐘執行一次
function setTrigger() {
  // 刪除現有的觸發器
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'processUnreadEmailsAndUpdateForm') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // 創建新的觸發器
  ScriptApp.newTrigger('processUnreadEmailsAndUpdateForm')
    .timeBased()
    .everyMinutes(5)
    .create();
    
  Logger.log('觸發器已設置完成');
}

// 設置每日檢查觸發器 - 每天早上9點檢查逾期用戶
function setDailyCheckTrigger() {
  // 刪除現有的每日檢查觸發器
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'checkOverdueUsers') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // 創建每日檢查觸發器
  ScriptApp.newTrigger('checkOverdueUsers')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();
    
  Logger.log('每日逾期檢查觸發器已設置完成');
}

// 手動測試函數
function testSystem() {
  Logger.log('開始測試系統...');
  processUnreadEmailsAndUpdateForm();
  Logger.log('測試完成');
}

// 手動檢查逾期用戶
function manualCheckOverdue() {
  Logger.log('開始手動檢查逾期用戶...');
  checkOverdueUsers();
  Logger.log('檢查完成');
}

// 手動發送當月成功繳費清單
function manualSendMonthlySuccessfulPayments() {
  Logger.log('開始手動發送當月成功繳費清單...');
  sendMonthlySuccessfulPaymentsToAdmin();
  Logger.log('發送完成');
}

// 查看表單資料狀態
function checkFormDataStatus() {
  try {
    var sheet = SpreadsheetApp.openById(FORM_SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();
    var currentDate = new Date();
    
    Logger.log('表單總記錄數：' + (values.length - 1));
    
    var subscribedCount = 0;
    var recentSuccessCount = 0;
    var recentFailureCount = 0;
    var overdueCount = 0;
    
    for (var i = 1; i < values.length; i++) {
      var subscriptionStartTime = values[i][5];
      var paymentStatus = values[i][7];
      var dynamicKey = values[i][9];
      
      if (subscriptionStartTime && subscriptionStartTime !== '') {
        subscribedCount++;
        
        if (paymentStatus) {
          if (paymentStatus.includes('成功') && isRecentMonth(paymentStatus, currentDate)) {
            recentSuccessCount++;
          } else if (paymentStatus.includes('失敗')) {
            recentFailureCount++;
          } else if (paymentStatus.includes('逾期未成功')) {
            overdueCount++;
          }
        }
        
        // 檢查是否有動態金鑰
        if (!dynamicKey || dynamicKey === '') {
          Logger.log('用戶（行 ' + (i + 1) + '）' + values[i][0] + ' 缺少動態金鑰');
        }
      }
    }
    
    Logger.log('已訂閱用戶數：' + subscribedCount);
    Logger.log('近期成功扣款：' + recentSuccessCount);
    Logger.log('扣款失敗：' + recentFailureCount);
    Logger.log('逾期未成功：' + overdueCount);
    
  } catch (error) {
    Logger.log('檢查表單資料時發生錯誤: ' + error.toString());
  }
}

// 為現有用戶補充動態金鑰
function addMissingDynamicKeys() {
  try {
    var sheet = SpreadsheetApp.openById(FORM_SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();
    var updatedCount = 0;
    
    for (var i = 1; i < values.length; i++) {
      var dynamicKey = values[i][9]; // J欄
      
      if (!dynamicKey || dynamicKey === '') {
        var newKey = generateRandomKey();
        sheet.getRange(i + 1, 10).setValue(newKey); // J欄
        updatedCount++;
        Logger.log('為用戶（行 ' + (i + 1) + '）' + values[i][0] + ' 生成動態金鑰：' + newKey);
      }
    }
    
    Logger.log('共為 ' + updatedCount + ' 位用戶補充了動態金鑰');
    
  } catch (error) {
    Logger.log('補充動態金鑰時發生錯誤: ' + error.toString());
  }
}

// 初始化所有觸發器
function initializeSystem() {
  setTrigger();
  setDailyCheckTrigger();
  Logger.log('系統初始化完成，已設置所有觸發器');
}
}
