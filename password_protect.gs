function displayPrompt(text) {
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt(text);

  var button = result.getSelectedButton();
  
  if (button !== ui.Button.OK) {
    return null
  }

  return result.getResponseText()
}

function encrypt(pw, cell) {
  const module = { exports: {} };
  eval(
    UrlFetchApp.fetch(
      "https://cdnjs.cloudflare.com/ajax/libs/sjcl/1.0.8/sjcl.min.js"
    ).getContentText()
  );
  sjcl = module.exports;
  //Logger.log('value ' + cell.getValue())
  // The encrypted message is in JSON format
  var json = sjcl.encrypt(pw, cell.getValue());
  //Logger.log('json ' + json)
  //Logger.log('stringify ' + Utilities.jsonStringify(json))
  // Convert the JSON to base64 (easier to copy-paste)
  var msg = Utilities.base64Encode(Utilities.jsonStringify(json));
  //Logger.log('msg ' + msg)
  //Logger.log('val ' + Utilities.base64Decode(msg))

  cell.setValue(msg)
}

function decrypt(pw, cell) {
  const module = { exports: {} };
  eval(
    UrlFetchApp.fetch(
      "https://cdnjs.cloudflare.com/ajax/libs/sjcl/1.0.8/sjcl.min.js"
    ).getContentText()
  );
  sjcl = module.exports;
  //Logger.log('value decrypt ' + cell.getValue())
  // Convert the JSON to base64 (easier to copy-paste)
  //Logger.log('base decoded ' + Utilities.base64Decode(cell.getValue()))
  var json = Utilities.jsonParse(Utilities.newBlob(Utilities.base64Decode(cell.getValue())).getDataAsString());
  //Logger.log('json ' + json)
  // The encrypted message is in JSON format
  var value = sjcl.decrypt(pw, json);
  //Logger.log('value ' + value)
  cell.setValue(value)
}

function loop_all(pw, is_decrypt) {
  for(var i = 0; i<SpreadsheetApp.getActive().getNumSheets();i++){
    //Logger.log('sheet ' + i)
    var sheet = SpreadsheetApp.getActive().getSheets()[i]
    var tf = sheet.createTextFinder('.+').useRegularExpression(true)
    var all = tf.findAll()
    for (var j = 0; j < all.length; j++) {
      cell = sheet.getRange(all[j].getA1Notation())
      //Logger.log('cell ' + cell)
      if (is_decrypt === true) {
        //Logger.log('decrypt')
        decrypt(pw, cell)
      } else {
        //Logger.log('encrypt')
        encrypt(pw, cell)
      }
    }
  }
}

function test() {
  Logger.log(SpreadsheetApp.getActive().getNumSheets())
  for(var i = 0; i<SpreadsheetApp.getActive().getNumSheets();i++){
    var sheet = SpreadsheetApp.getActive().getSheets()[i]
    var tf = sheet.createTextFinder('.+').useRegularExpression(true)
    var all = tf.findAll()
    for (var i = 0; i < all.length; i++) {
      cell = sheet.getRange(all[i].getA1Notation())
      Logger.log(cell.getValue())
    }
  }
}


function check_encrypt() {
  test_value = SpreadsheetApp.getActive().getSheetByName('Encrypt-Validation').getRange(1, 1).getValue()
  Logger.log(test_value)
  if (test_value === 'VALIDATE') {
    pw = displayPrompt("Please enter encryption password")
    if (pw === null) {
      return
    }
    loop_all(pw, false)
  }
  else {
    pw = displayPrompt("Please enter decryption password")
    if (pw === null) {
      return
    }
    loop_all(pw, true)
  }
}


function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Encryption')
      .addItem('Encrypt/Decrypt', 'check_encrypt')
      .addToUi();
}


