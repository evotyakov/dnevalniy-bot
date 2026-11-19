function setWebhook() {
  var webAppUrl = "https://script.google.com/macros/s/AKfycbw2__xMDg7VzRDKoK_aNSFeZdw8x2_iAC7zg2aHcUaXK_XY2RVqNEGDQ7OHAkd6rWTx/exec"; 
  var url = "https://api.telegram.org/bot" + TOKEN + "/setWebhook?url=" + webAppUrl;
  var response = UrlFetchApp.fetch(url);
  Logger.log(response.getContentText());
}

function isAdmin(chatId, userId) {
  var url = "https://api.telegram.org/bot" + TOKEN + "/getChatMember?chat_id=" + chatId + "&user_id=" + userId;
  try {
    var response = UrlFetchApp.fetch(url, { "muteHttpExceptions": true });
    var data = JSON.parse(response.getContentText());
    
    if (data.ok && data.result) {
      var status = data.result.status;
      if (status === "administrator" || status === "creator") {
        return true;
      }
    }
  } catch (err) {
    Logger.log("Ошибка проверки админа: " + err.toString());
  }
  return false;
}

function doPost(e) {
  try {
    var contents = JSON.parse(e.postData.contents);
    
    // Приветственное сообщение при добавлении бота или нового пользователя в чат
    if (contents.message && contents.message.new_chat_members) {
      var chatId = contents.message.chat.id;
      var welcomeText = "Здравия желаю, товарищи студенты! Студент Мбапенко в чат взвода прибыл.\n\n" +
        "**Доступные приказы:**\n" + 
        "- Мбапенко, доложи обстановку\n" +
        "- Мбапенко, вызови рандомного студента\n" +
        "- Мбапенко, напомни (что-то сделать)\n" +
        "- Студент (фамилия) респект\n" + 
        "- Студент (фамилия) косяк\n\n" +
        "- Мбапенко, напомни ... \n" +
        "- Мбапенко, забудь про ...\n\n" +
        "Каждый четверг буду назначать дежурного и нарядного. Доклад окончил";
      sendMessage(chatId, welcomeText);
      return;
    }

    if (contents.message && contents.message.text) {
      var chatId = contents.message.chat.id;
      var text = contents.message.text.trim();
      var lowerText = text.toLowerCase();

      // Команда получения актуального распределения
      if (lowerText === "/status" || lowerText.indexOf("мбапенко, доложи обстановку") !== -1) {
        sendCurrentStatus(chatId);
        return;
      }

      // Команда вызова случайного студента
      if (lowerText.indexOf("мбапенко, вызови рандомного студента") !== -1) {
        sendRandomStudent(chatId);
        return;
      }

      // Команда добавления напоминания
      if (lowerText.startsWith("мбапенко, напомни")) {
        var reminderText = text.substring("мбапенко, напомни ".length).trim();
        if (reminderText !== "") {
          addReminder(chatId, reminderText);
        } else {
          sendMessage(chatId, "Не могу знать, что именно нужно напомнить");
        }
        return;
      }

      if (lowerText.startsWith("мбапенко, забудь про")){
        var reminderText = text.substring("мбапенко, забудь про ".length).trim();
        if (reminderText !== ""){
          removeReminder(chatId, reminderText);
        }
        else {
          sendMessage(chatId, "Не могу знать что забыть");
        }
        return;
      }

      // Обработка команд типа "студент <Фамилия> косяк" или "студент <Фамилия> респект"
      if (lowerText.startsWith("студент ")) {
        var parts = text.split(/\s+/);
        
        if (parts.length >= 3) {
          var surname = parts[1];
          var action = parts[parts.length - 1].toLowerCase();

          if (action === "косяк" || action === "респект") {
            rateStudent(chatId, surname, action);
          }
        }
      }
    }
  } catch (err) {
    Logger.log("Ошибка: " + err.toString());
  }
}

function sendWeeklyReport(chatId) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var intro = "Здравия желаю, товарищи студенты! Завтра пятница, день военной подготовки. Не забудьте: \n\n";
  
  var lastRow = sheet.getLastRow();
  var listItems = [];
  
  if (lastRow >= 19) {
    var values = sheet.getRange(19, 1, lastRow - 18, 1).getValues();
    for (var i = 0; i < values.length; i++) {
      var val = String(values[i][0]).trim();
      if (val !== "") {
        listItems.push("- " + val);
      }
    }
  }
  
  var msg = intro + (listItems.length > 0 ? listItems.join("\n") : "Список пуст");
  
  sendMessage(chatId, msg);
}

// Функция добавления напоминания в таблицу, начиная с ячейки A19 и вниз
function addReminder(chatId, reminderText) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  
  var row = 19;
  while (sheet.getRange(row, 1).getValue() !== "") {
    row++;
  }
  
  sheet.getRange(row, 1).setValue(reminderText);

  sendMessage(chatId, "Есть! В четверг вечером напомню " + reminderText);
}

function removeReminder(chatId, reminderText) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var lastRow = sheet.getLastRow();
  var needle = reminderText.toLowerCase();

  for (var row = 19; row <= lastRow; row++) {
    var val = String(sheet.getRange(row, 1).getValue()).trim();
    if (val !== "" && val.toLowerCase().indexOf(needle) !== -1) {
      sheet.getRange(row, 1).setValue("");
      sendMessage(chatId, "Есть! Отменил напоминание: " + val);
      return;
    }
  }
  sendMessage(chatId, "Не могу знать такого напоминания");
}


function rateStudent(chatId, surname, action) {
  var targetCol = (action === "косяк") ? 5 : 6; // Столбец E (5) или F (6)
  var result = addPointBySurname(surname, targetCol);
  
  if (result.success) {
    var emoji = (action === "косяк") ? "Позор!" : "Ура!";
    var verb = (action === "косяк") ? "Вы допустили" : "Вам объявляется";
    sendMessage(chatId, "Студент **" + result.only_surname + "! " + verb + " " + action + ". " + emoji + "\nВаш рейтинг: " + result.newValue + ".");
  } else {
    sendMessage(chatId, "Студент с фамилией **" + surname + "** не найден в таблице.");
  }
}

function sendRandomStudent(chatId) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var range = sheet.getRange("A4:A12");
  var values = range.getValues();
  
  var students = [];
  for (var i = 0; i < values.length; i++) {
    var surname = String(values[i][0]).trim();
    if (surname !== "") {
      students.push(surname);
    }
  }
  
  if (students.length === 0) {
    sendMessage(chatId, "В диапазоне A4:A12 нет фамилий студентов.");
    return;
  }
  
  var randomIndex = Math.floor(Math.random() * students.length);
  var randomStudent = students[randomIndex];
  
  var msg = "Есть. Случайный студент: **" + randomStudent + "**";
  sendMessage(chatId, msg);
}

function addPointBySurname(surname, colIndex) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var lastRow = sheet.getLastRow();
  
  if (lastRow < 3) return { success: false };
  
  var surnames = sheet.getRange(3, 1, lastRow - 2, 2).getValues();
  
  for (var i = 0; i < surnames.length; i++) {
    var dbSurname = String(surnames[i][0]).trim();
    var dbName = String(surnames[i][1]).trim();
    
    if (dbSurname.toLowerCase() === surname.toLowerCase()) {
      var rowIndex = i + 3;
      
      var cell = sheet.getRange(rowIndex, colIndex);
      var currentValue = Number(cell.getValue()) || 0;
      var newValue = currentValue + 1;
      cell.setValue(newValue);
      
      var totalCell = sheet.getRange(rowIndex, 7); // Столбец G
      var totalValue = totalCell.getValue();
      
      return {
        success: true,
        fullName: dbSurname + " " + dbName,
        only_surname: dbSurname,
        newValue: totalValue
      };
    }
  }
  
  return { success: false };
}

function sendCurrentStatus(chatId) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  
  var dutyStudent = sheet.getRange("B15").getValue();
  var paradeStudent = sheet.getRange("B16").getValue();
  
  var msg = "Есть. **Актуальное назначение:**\n\n" +
            "Дежурный: студент " + dutyStudent + "\n" +
            "Нарядный: студент " + paradeStudent;
            
  sendMessage(chatId, msg);
}

function sendMessage(chatId, text) {
  var url = "https://api.telegram.org/bot" + TOKEN + "/sendMessage";
  var payload = {
    "chat_id": chatId,
    "text": text,
    "parse_mode": "Markdown"
  };
  var options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload)
  };
  UrlFetchApp.fetch(url, options);
}