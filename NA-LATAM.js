BOT_USERNAME = "CESAR CBS NA-LATAM Audit"

SHEET_ID = "1k3yAaMQkrZRbEGa9hJTywXu8pPYEE8EETvahekaKBYE";
DEVICES_SHEET = "CBS";
USERS_SHEET = "team";
LOG_SHEET = "log"
LITE_SEARCH_SHEET = "SearchAuditBot";

DEVICE_NAME_COLUMN = 1;
DEVICE_SERIAL_NUMBER_COLUMN = 3; 
DEVICE_PROJECT_OWNER_COLUMN = 8;
DEVICE_ID_COLUMN = 12;
DEVICE_CURRENT_OWNER_COLUMN = 11;

BOX_NAME = "Cofre";
MISSING_NAME = "Perdido";

HEADER_ICON_URL = 'https://i.imgur.com/fUeNuyj.png';
PHONE_ICON_URL = "https://i.imgur.com/ZC9LZpF.png"
BORROW_ICON_URL = "https://i.imgur.com/CXtYODs.png"
PHONE_BOX_ICON_URL = "https://i.imgur.com/DvgkAJq.png"
PHONE_MISSING_ICON_URL = "https://i.imgur.com/bfA8fZB.png"

HEADER_TITLE = "Auditoria de Devices";
HELP_HEADER = {
  header: {
    title : HEADER_TITLE,
    subtitle : "Help",
    imageUrl : HEADER_ICON_URL
  }
};


function onMessage(event) {
  try {
    var email = event.message.sender.email;
    var coreId = email.split("@")[0];
    
    var usersSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(USERS_SHEET);
    var matches = usersSheet.getRange(2,1,usersSheet.getLastRow(),2).createTextFinder(coreId).findAll();    
    if (matches.length == 0) {
      return {text: "ERROR: CoreID não encontrado: \"" + coreId + "\". Confirme na aba " + USERS_SHEET + " da planilha."};
    }
    if (matches.length > 1) {
      return {text: "ERROR: CoreID duplicado na aba " + USERS_SHEET + " da planilha."};
    }
    
    var matchRow = matches[0].getRow();
    var callerName = usersSheet.getRange(matchRow, 1).getValue();
    
    var userMessage = event.message.text;
    userMessage = userMessage.replace("@" + BOT_USERNAME,"");
    var fun = userMessage.trim().split(" ");
    
    var letAuditFootprint = false;
    var generateHoReport = false;
    
    for (i in fun) {
      if (fun[i] == "--report" || fun[i] == "-r") {
        generateHoReport = true;
        fun.pop(i)
        i -= 1;
      }
    }
    
    if (fun[0] == "audit") {
      letAuditFootprint = true;
      
      return prepareAndInvokeResponse(coreId, letAuditFootprint, generateHoReport, callerName);
    }
    
    if (fun[0] == "list" && fun.length > 1) {
      if (fun[i] == coreId)
        letAuditFootprint = true;
      else {
        coreId = fun[1];
        letAuditFootprint = false;
      }
      
      return prepareAndInvokeResponse(coreId, letAuditFootprint, generateHoReport, callerName);
    }
    
    if (fun[0] == "find" && fun.length > 1) {
      var search = fun.slice(1).join(" ");
      return findAndListDevices(coreId, callerName, search);
    }
    
    if (fun[0] == "coreid" && fun.length > 1) {
      searchText = fun[1];
      
      var usersSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(USERS_SHEET);
      var matches = usersSheet.getRange(2,1,usersSheet.getLastRow(),2).createTextFinder(searchText).findAll();
      var ret = "";
      
      if (matches.length == 0) {
        ret = "Nada encontrado para \"" + searchText + "\"";
      } else {
        for (i = 0 ; i < matches.length ; i++){
          var row = matches[i].getRow();
          ret = ret + usersSheet.getRange(row, 3).getValue() + " (" +
            usersSheet.getRange(row, 2).getValue() + ")\n";
        }
      }
      
      return {text: ret};
    }
    
    return createHelpResponse();
  }
  catch(e) {
     return {text: e.message};
  }   
}

function prepareAndInvokeResponse(coreId, letAuditFootprint, generateHoReport, callerName) {
  if (coreId == null)
    return {text: "ERROR: CoreID nulo."};
  
  // Sheet inside spreadsheet
  var usersSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(USERS_SHEET);
  
  var cellFinder = usersSheet.getRange(2,2,usersSheet.getLastRow()).createTextFinder(coreId);
  if (cellFinder.findAll().length == 0) {
    return {text: "ERROR: CoreID não encontrado: \"" + coreId + "\". Confirme na aba " + USERS_SHEET + " da planilha."};
  } else if (cellFinder.findAll().length > 1) {
    return {text: "ERROR: CoreID duplicado na aba " + USERS_SHEET + " da planilha."};
  }
  
  var userName = null;
  var userTeam = null;
  var row = cellFinder.findNext().getRow();
  try {
    userName = usersSheet.getRange(row, 1).getValue();
    coreId = usersSheet.getRange(row,2).getValue();
    userTeam = usersSheet.getRange(row,3).getValue();
    var header = {
      header: {
        title : HEADER_TITLE,
        subtitle : userName + " (" + coreId +")",
        imageUrl : HEADER_ICON_URL
      }};
  } catch (err) {
    return {text: "ERROR: Não encontrado: " + userName + ", err: " + err};
  }
  
  if (userName != null) {
    var today = new Date();
    var date = today.getFullYear()+'-'+(today.getMonth()+1)+'-'+today.getDate();
    var time = today.getHours() + ":" + today.getMinutes() + ":" + today.getSeconds();
    var dateTime = date+' '+time;
    
    // devices
    var devices = SpreadsheetApp.openById(SHEET_ID).getSheetByName(DEVICES_SHEET);
    var cell_finder = devices.createTextFinder(userName);
    var matches = cell_finder.findAll();
    var widgets = [];
    var report = "[Assunto]\nComunicado de Saída de Protótipo para Home Office\n\n";
    report = report + "[Corpo]\nSegue lista atualizada dos protótipos que estão comigo em Home Office: \n\n";
    for (i = 0 ; i < matches.length ; i++){
      if (matches[i].getColumn() == DEVICE_CURRENT_OWNER_COLUMN) {
        var row = matches[i].getRow();
        
        if (letAuditFootprint) {
          // mark cell with date time
          devices.getRange(row, DEVICE_CURRENT_OWNER_COLUMN).setNote("Última auditoria executada às " + dateTime + " por " + callerName);
        }
        
        var deviceName = devices.getRange(row,DEVICE_NAME_COLUMN).getValue();
        var deviceSerial = devices.getRange(row,DEVICE_SERIAL_NUMBER_COLUMN).getValue();
        var deviceProjectOwner = devices.getRange(row,DEVICE_PROJECT_OWNER_COLUMN).getValue();
        var deviceID = devices.getRange(row,DEVICE_ID_COLUMN).getValue();

        var device = {
          keyValue: {
            iconUrl: PHONE_ICON_URL,
            topLabel: deviceProjectOwner,
            content: deviceID,
            contentMultiline: true,
            bottomLabel: "" + deviceSerial + " - " + deviceName,
            button: {
              textButton: {
               text: "Gerenciar",
                 onClick: {
                   action: {
                     actionMethodName: "manageDevice",
                     parameters: [
                       {key: "row",      value: "" + row},
                       {key: "callerName",   value: callerName}
                     ]
                   }
                 }
              }
            }
          }
        };
        widgets.push(device);

        report = report +
          "Modelo do telefone Celular: " + deviceName + "\n" +
            "Serial number: " + deviceSerial + "\n" +
              "ID na planilha: " + deviceID + "\n" +
                "Device está no meu nome na planilha de protótipos: SIM\n" +
                  "Device é do meu grupo: " + (userTeam == deviceProjectOwner? "SIM": "NAO") + "\n\n";
      }
    }
    
    if (widgets.length == 0) {
      return createCardResponse(header, [{textParagraph: {text: "Nenhum device com<br><b>" + userName + " (" +  coreId + ")<br>" }}]);
    } else {
      if (generateHoReport)
        return {text: "```" + report + '```'};
      else
        return createCardResponse(header, widgets);
    }
  }
}

function findAndListDevices(coreId, callerName, search) {
    if (search == null)
    return {text: "ERROR: Termo de busca nulo."};
  
  var header = {
    header: {
      title : HEADER_TITLE,
      subtitle : "Busca por: " + search,
      imageUrl : HEADER_ICON_URL
    }
  };
  
  var lite = SpreadsheetApp.openById(SHEET_ID).getSheetByName(LITE_SEARCH_SHEET);
  var devices = SpreadsheetApp.openById(SHEET_ID).getSheetByName(DEVICES_SHEET);

  // The Lite sheet only has the criteria used on the search
  var cellFinder = lite.getRange(2, 1, lite.getLastRow()-1, lite.getLastColumn()).createTextFinder(search);
  
  var matches = cellFinder.findAll();
  
  // Matches can be duplicated due to searching on Name and ID
  var deviceAlreadyAdded = {}  
  var widgets = [];
  
  for (i = 0 ; i < matches.length ; i++) {
    var row = matches[i].getRow();
    var deviceID = devices.getRange(row,DEVICE_ID_COLUMN).getValue();
    
    if (deviceAlreadyAdded[deviceID]) continue;
    else deviceAlreadyAdded[deviceID] = true;
    
    var deviceProjectOwner = devices.getRange(row,DEVICE_PROJECT_OWNER_COLUMN).getValue();
    var deviceCurrentOwner = devices.getRange(row,DEVICE_CURRENT_OWNER_COLUMN).getValue();
    
              
    var device = {
      keyValue: {
        iconUrl: PHONE_ICON_URL,
        topLabel: "" + deviceProjectOwner,
        content: deviceID,
        contentMultiline: true,
        bottomLabel: deviceCurrentOwner + " ",
        button: {
        textButton: {
           text: "Gerenciar",
           onClick: {
              action: {
                actionMethodName: "manageDevice",
                  parameters: [
                   {key: "row",      value: "" + row},
                   {key: "callerName",   value: callerName}
                 ]
               }
             }
           }
         }
       }
    };
    widgets.push(device);
  }

  if (widgets.length == 0) {
    return createCardResponse(header, [{textParagraph: {text: "Nenhum device com nome <b>" + search + "</b> encontrado. Tente pesquisar pelo <b>Nome</b>, <b>Serial Number</b> ou <b>ID</b> do device." }}]);
  } else {
    return createCardResponse(header, widgets);
  }
}

function createHelpResponse() {
  var helpResponse = [];

  var audit = {
    textParagraph: {
      text: "<b>audit</b><br>Lista devices que estão com você e deixa um registro na planilha de controle."
    }
  };

  var listSomeoneElsesBorrowing = {
    textParagraph: {
      text: "<b>list &lt;coreid&gt;</b><br>Lista devices que estão com um usuário. <br> <i>list missing</i> - Mostra devices perdidos. <br> <i>list box</i> - Mostra devices no cofre. "
    }
  };
  
  var findDevice = {
    textParagraph: {
      text: "<b>find &lt;string&gt;</b><br>Busca por devices pelo Nome, Serial Number ou ID."
    }
  };
  
  var findCoreid = {
    textParagraph: {
      text: "<b>coreid &lt;string&gt;</b><br>Busca o CoreID por parte do nome."
    }
  };

  helpResponse.push(audit);
  helpResponse.push(listSomeoneElsesBorrowing);
  helpResponse.push(findDevice);
  helpResponse.push(findCoreid);
  
  return createCardResponse(HELP_HEADER, helpResponse);  
}

function createCardResponse(header, widgets = []) {
  return {
    cards: [header, {
      sections: [{
        widgets: widgets
      }]
    }]
  };
}

function createManageDeviceCard(row, callerName) {
  var devices = SpreadsheetApp.openById(SHEET_ID).getSheetByName(DEVICES_SHEET);
  var deviceName = devices.getRange(row,DEVICE_NAME_COLUMN).getValue();
  var deviceSerial = devices.getRange(row,DEVICE_SERIAL_NUMBER_COLUMN).getValue();
  var deviceOwner = devices.getRange(row,DEVICE_CURRENT_OWNER_COLUMN).getValue();
  var deviceProjectOwner = devices.getRange(row,DEVICE_PROJECT_OWNER_COLUMN).getValue();
  
  var header = {
      header: {
        title : "" + deviceName + " " + deviceSerial,
        subtitle : deviceOwner
      }
  };
  
  var widgets = [];
  
  var box = {
    keyValue: {
      iconUrl: PHONE_BOX_ICON_URL,
      content: "Devolver ao Cofre",      
      onClick: {
        action: {
          actionMethodName: "putBackInBox",
          parameters: [
            {key: "row",          value: "" + row},
            {key: "callerName",   value: callerName}
          ]
        } 
      }
    }
  }
  if (deviceOwner != BOX_NAME)
    widgets.push(box);
  
  var get = {
    keyValue: {
      iconUrl: BORROW_ICON_URL,
      content: "Pegar Device",
      onClick: {
        action: {
          actionMethodName: "borrowDevice",
          parameters: [
            {key: "row",          value: "" + row},
            {key: "callerName",   value: callerName}
          ]
        }          
      }
    }
  }
  if (deviceOwner != callerName)
    widgets.push(get);
  
  var missing = {
    keyValue: {
      iconUrl: PHONE_MISSING_ICON_URL,
      content: "Device Perdido",
      onClick: {
        action: {
          actionMethodName: "doNotHaveDevice",
          parameters: [
            {key: "row",          value: "" + row},
            {key: "callerName",   value: callerName}
          ]
        }
      }
    }
  }
  widgets.push(missing);

  return createCardResponse(header, widgets);
}

// Mark device as missing
function doNotHaveDevice(row, callerName) {
  var devices = SpreadsheetApp.openById(SHEET_ID).getSheetByName(DEVICES_SHEET);
  var deviceName = devices.getRange(row,DEVICE_NAME_COLUMN).getValue();
  var deviceSerial = devices.getRange(row,DEVICE_SERIAL_NUMBER_COLUMN).getValue();
  var deviceOwner = devices.getRange(row,DEVICE_CURRENT_OWNER_COLUMN).getValue();
  
  devices.getRange(row, DEVICE_CURRENT_OWNER_COLUMN).setValue(MISSING_NAME);
  addChangeToLog(deviceOwner, callerName, deviceName, deviceSerial); 
  return { text: "Device marcado como " + MISSING_NAME + "." };
}

function putBackInBox(row, callerName) {
  var devices = SpreadsheetApp.openById(SHEET_ID).getSheetByName(DEVICES_SHEET);
  var deviceName = devices.getRange(row,DEVICE_NAME_COLUMN).getValue();
  var deviceSerial = devices.getRange(row,DEVICE_SERIAL_NUMBER_COLUMN).getValue();
  var deviceOwner = devices.getRange(row,DEVICE_CURRENT_OWNER_COLUMN).getValue();
  
  // If the device is not on the caller's name, add a log of this change
  if (deviceOwner != callerName)
    addChangeToLog(deviceOwner, callerName, deviceName, deviceSerial);
  devices.getRange(row, DEVICE_CURRENT_OWNER_COLUMN).setValue(BOX_NAME);
  addChangeToLog(callerName, BOX_NAME, deviceName, deviceSerial);
  return { text: "Device colocado no " + BOX_NAME + "." };
}

function borrowDevice(row, callerName) {
  var devices = SpreadsheetApp.openById(SHEET_ID).getSheetByName(DEVICES_SHEET);
  var deviceName = devices.getRange(row,DEVICE_NAME_COLUMN).getValue();
  var deviceSerial = devices.getRange(row,DEVICE_SERIAL_NUMBER_COLUMN).getValue();
  var deviceOwner = devices.getRange(row,DEVICE_CURRENT_OWNER_COLUMN).getValue();
  
  devices.getRange(row, DEVICE_CURRENT_OWNER_COLUMN).setValue(callerName);
  addChangeToLog(deviceOwner, callerName, deviceName, deviceSerial); 
  return { text: "Device estava com " + deviceOwner + ", colocado no nome de " + callerName + "." };
}

function onCardClick(event) {
  console.info(event);

  var row = event.action.parameters[0].value;
  var callerName = event.action.parameters[1].value;
  
  if (event.action.actionMethodName == "manageDevice") {
    return createManageDeviceCard(row, callerName)
  }
  
  if (event.action.actionMethodName == "doNotHaveDevice") {
    return doNotHaveDevice(row, callerName);
  }
  
  if (event.action.actionMethodName == "putBackInBox") {
    return putBackInBox(row, callerName);
  }
  
  if (event.action.actionMethodName == "borrowDevice") {
    return borrowDevice(row, callerName);
  }
  
  return { text: "Houve um problema nesta ação." };
}

function addChangeToLog(from, to, deviceName, deviceSerial) {
  var logSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(LOG_SHEET) 
  var insertRow = 2
  logSheet.insertRows(insertRow);
  logSheet.getRange(insertRow,1).setValue(from)
  logSheet.getRange(insertRow,2).setValue(to)
  logSheet.getRange(insertRow,3).setValue(deviceName)
  logSheet.getRange(insertRow,4).setValue(deviceSerial)
  logSheet.getRange(insertRow,5).setValue(new Date())
}


function onAddToSpace(event) {
  var message = "";

  if (event.space.type == "DM") {
    message = "𝙱𝚎𝚎𝚙 𝚋𝚘𝚘𝚙. Obrigado por me adicionar, " + event.user.displayName + "!";
  } else {
    message = "𝙱𝚎𝚎𝚙 𝚋𝚘𝚘𝚙. Obrigado por me adicionar em " + event.space.displayName;
  }

  return { text: message };
}

function onRemoveFromSpace(event) {
  console.info("Bot removed from ", event.space.name);
}
