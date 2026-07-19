/**
 * Northern Run - Server-side Code (Code.gs)
 * Optimized for speed, reliability, and to avoid shared state bugs in Web Apps.
 * [รุ่นรวมสมบูรณ์ - คำนวณเงิน 450/550 อัตโนมัติและแก้ดัชนีคอลัมน์สลิป AH/AI ตรงกัน 100%]
 */

const LINE_ACCESS_TOKEN = "i+uAHn+oyCM7tDDqWLVFRDV0s/vw44+He8cqwDK+uakj7gvJBHgPrUlFZ306u4Wk3h5Wk6XF7RcB9CSVAJItH29scPP4ZiUtyCu20AMRMpJCtbXsvUhr+JpDeW3AoUPRadnvM3s/xvisp50lxmkqXgdB04t89/1O/w1cDnyilFU=";
const LINE_GROUP_ID = "C6be883b5bf0e48f3391a15c88bdd1859";

function doGet(e) {
  if (e && e.parameter && e.parameter.fileId) {
    try {
      var file = DriveApp.getFileById(e.parameter.fileId);
      var blob = file.getBlob();
      return HtmlService.createHtmlOutput('<html><body style="margin:0; background:#0e0e0e; display:flex; justify-content:center; align-items:center; height:100vh;"><img src="data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes()) + '" style="max-width:100%; max-height:100vh;"></body></html>');
    } catch(err) {
      return HtmlService.createHtmlOutput("<b>ไม่สามารถเข้าถึงไฟล์ภาพได้:</b> " + err.message);
    }
  }
  return HtmlService.createTemplateFromFile("index").evaluate()
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/** ดึงข้อมูลการตั้งค่าทั้งหมดพร้อมแปลงลิงก์รูปภาพ Google Drive */
function getSettings() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('setting');
  var rangeC = sheet.getRange('C3:C25').getValues();
  var priceSend = sheet.getRange('D13').getValue();
  
  var rawBanner = rangeC[5][0] || "";
  var bannerUrl = rawBanner;
  if (rawBanner.indexOf("drive.google.com") !== -1) {
    var fileId = "";
    if (rawBanner.indexOf("/d/") !== -1) { fileId = rawBanner.split("/d/")[1].split("/")[0]; } 
    else if (rawBanner.indexOf("id=") !== -1) { fileId = rawBanner.split("id=")[1].split("&")[0]; }
    if (fileId) { bannerUrl = "https://lh3.googleusercontent.com/d/" + fileId; }
  }
  
  return {
    title: rangeC[0][0] || "",
    date: rangeC[1][0] || "",
    time: rangeC[2][0] || "",
    director: rangeC[3][0] || "",
    place: rangeC[4][0] || "",
    banner: bannerUrl,
    id_map: rangeC[6][0] || "",
    id_promptpay: rangeC[7][0] || "",
    idfolder: rangeC[8][0] || "",
    email: rangeC[9][0] || "",
    send: rangeC[10][0] || "",
    promptpay: rangeC[11][0] || "",
    bank_number: rangeC[12][0] || "",
    bank_name: rangeC[13][0] || "",
    bank: rangeC[14][0] || "",
    guarantee: rangeC[15][0] || "",
    facebook: rangeC[16][0] || "",
    shirt1: rangeC[19][0] || "", 
    shirt2: rangeC[20][0] || "", 
    shirt3: rangeC[21][0] || "", 
    shirt4: rangeC[22][0] || "", 
    price_send: priceSend || 0
  };
}

function getAdminData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("register");
  if (!ss) return [];
  return ss.getDataRange().getDisplayValues().slice(1);
}

function updateApplicantStatus(idCard, newStatus, htmlText) {
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("register");
  var data = ss.getDataRange().getDisplayValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][4] == idCard) {
      ss.getRange(i + 1, 29).setValue(newStatus); // คอลัมน์ AC
      ss.getRange(i + 1, 30).setValue(htmlText);  // คอลัมน์ AD
      return true;
    }
  }
  return false;
}

/** บันทึกฟอร์มลงทะเบียนรายใหม่ + คำนวณยอดเงินส่งลงชีต AG (คอลัมน์ที่ 33) อัตโนมัติ */
function processForm(obj) {
  var ss = SpreadsheetApp.getActive().getSheetByName("register");
  var output = {};
  
  var autoPrice = 0;
  if (obj.type_run && obj.type_run.indexOf("10.5") !== -1) {
    autoPrice = 550;
  } else if (obj.type_run && obj.type_run.indexOf("5.5") !== -1) {
    autoPrice = 450;
  }

  ss.appendRow([
    "", "", obj.fname, obj.lname, "'" + obj.idcard, obj.gender,
    obj.d_birth, obj.m_birth, obj.y_birth, obj.email, "'" + obj.tel,
    '"' + obj.address1, obj.address2, obj.address3, obj.address4, obj.address5, obj.address6,
    obj.illness, obj.contact, "'" + obj.tel_contact, obj.team, obj.type_run,
    obj.shirt, obj.size, obj.sendshirt, "'" + obj.send_address,
    new Date(), '<span class="text-primary"><i class="fa-regular fa-clock text-primary"></i> รอชำระเงิน</span>', "รอชำระเงิน",
    "", "", "", autoPrice, "" 
  ]);
  var data = ss.getRange(ss.getLastRow(), 1, 1, ss.getLastColumn()).getDisplayValues()[0];
  output.data = data;
  
  try {
    let msg = "🏃‍♂️ มีผู้สมัครวิ่งรายใหม่! 🏃‍♂️" +
              "\n👤 ชื่อ-นามสกุล: " + obj.fname + " " + obj.lname +
              "\n🆔 เลขบัตร ปชช: " + obj.idcard +
              "\n🏁 ประเภท: " + obj.type_run +
              "\n👕 ไซส์เสื้อ: " + obj.size + " (" + obj.sendshirt + ")";
    sendLineMessagingAPI(msg);
  } catch(e) {
    Logger.log("ไม่สามารถส่ง LINE สมัครใหม่ได้: " + e.message);
  }
  
  return output;
}

function filterData(keySearch) {
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("register");
  if (!ss) return [];
  var dataAll = ss.getDataRange().getDisplayValues();
  return dataAll.filter(f => f[4] == keySearch);
}

function getDropdownSize() { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('setting').getRange('setting!G26:G36').getValues(); }
function getDropdownType() { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('setting').getRange('setting!G40:G49').getValues(); }

function checkLogin(username, pwd) {
  var userlog = "null";
  const ws = SpreadsheetApp.getActive().getSheetByName('Admin');
  if (!ws) return userlog;
  const data = ws.getDataRange().getDisplayValues().slice(1).filter(r => r[0] !== "");
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] == username && data[i][1] == pwd) { userlog = data[i][2]; break; }
  }
  return userlog;
}

function getDirectImageUrl(fileId) {
  var webAppUrl = ScriptApp.getService().getUrl();
  return webAppUrl.indexOf('?') !== -1 ? webAppUrl + "&fileId=" + fileId : webAppUrl + "?fileId=" + fileId;
}

/** อัปโหลดสลิปใบเสร็จเงิน + บันทึกลงคอลัมน์ AI ช่องที่ 35 และยิงแจ้งเตือนกลุ่ม LINE */
function uploadSlip(formObj) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("register");
    var sheetSetting = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('setting');
    
    var folderId = sheetSetting.getRange('C11').getValue();
    if (!folderId) { throw new Error("ไม่พบข้อมูลไอดีโฟลเดอร์เก็บสลิปในหน้าตั้งค่า (ช่อง C11)"); }
    
    var folder = DriveApp.getFolderById(folderId);
    var fileBlob = formObj.myFile; 
    var idCardSearch = formObj.slipID;
    
    if (!idCardSearch || idCardSearch.trim() === "") { throw new Error("ไม่พบข้อมูลเลขบัตรประชาชน กรุณาลองใหม่อีกครั้งจากเมนูค้นหา"); }
    if (!fileBlob || fileBlob.getName() == "") { throw new Error("ไม่พบไฟล์รูปภาพสลิป กรุณาเลือกไฟล์ใหม่อีกครั้ง"); }
    
    fileBlob.setName("Slip_" + idCardSearch + "_" + new Date().getTime());
    var file = folder.createFile(fileBlob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    var fileUrl = getDirectImageUrl(file.getId()); 
    
    var dataAll = ss.getDataRange().getDisplayValues();
    var targetRow = -1;
    for (var i = 0; i < dataAll.length; i++) {
      if (dataAll[i][4] == idCardSearch) { targetRow = i + 1; break; }
    }
    
    if (targetRow != -1) {
      ss.getRange(targetRow, 29).setValue("รอตรวจสอบ"); 
      ss.getRange(targetRow, 30).setValue('<span class="text-success fw-bold"><i class="fa-solid fa-circle-check me-1"></i> แจ้งแล้ว รอตรวจสอบ</span>'); 
      ss.getRange(targetRow, 35).setValue(fileUrl); // บันทึกลิงกลงช่องคอลัมน์ AI ให้หน้าบ้านดักจับได้ถูกต้อง
      
      var applicantData = ss.getRange(targetRow, 1, 1, ss.getLastColumn()).getDisplayValues()[0];
      var firstName = applicantData[2] || "-";  
      var lastName = applicantData[3] || "-";   
      var typeRun = applicantData[21] || "-";   
      
      var amount = applicantData[32];
      if(!amount || amount === "0" || amount === "0.00" || amount === "") {
        amount = typeRun.indexOf("10.5") !== -1 ? "550" : "450";
      }
      
      try {
        let msg = "💰 มีคนแจ้งสลิปโอนเงินเข้ามา! 💰" +
                  "\n👤 ชื่อ-นามสกุล: " + firstName + " " + lastName +
                  "\n🆔 เลขบัตร ปชช: " + idCardSearch + 
                  "\n🏁 รุ่นที่สมัคร: " + typeRun +
                  "\n💵 จำนวนเงิน: " + amount + " บาท" +
                  "\n📝 สถานะปัจจุบัน: แจ้งแล้ว รอตรวจสอบ" +
                  "\n📂 โปรดเข้าหน้าต่างระบบ Admin เพื่อทำการตรวจสอบ";
        sendLineMessagingAPI(msg);
        
        if (file.getId()) {
          var lineImageUrl = "https://lh3.googleusercontent.com/d/" + file.getId();
          sendLineImageAPI(lineImageUrl);
        }
      } catch(errLine) {
        Logger.log("ไม่สามารถส่ง LINE สลิปได้: " + errLine.message);
      }
      
      return { success: true, message: "อัปโหลดหลักฐานแจ้งชำระเงินเรียบร้อยแล้ว!" };
    } else {
      throw new Error("ไม่พบข้อมูลผู้สมัครรายนี้ในระบบ");
    }
  } catch (error) {
    throw new Error(error.message);
  }
}

function sendLineMessagingAPI(textMessage) {
  const url = "https://api.line.me/v2/bot/message/push";
  const postData = { "to": LINE_GROUP_ID, "messages": [{ "type": "text", "text": textMessage }] };
  const options = {
    "method": "post",
    "headers": { "Content-Type": "application/json", "Authorization": "Bearer " + LINE_ACCESS_TOKEN },
    "payload": JSON.stringify(postData), "muteHttpExceptions": true
  };
  try { UrlFetchApp.fetch(url, options); } catch(e) {}
}

function sendLineImageAPI(imageUrl) {
  const url = "https://api.line.me/v2/bot/message/push";
  const postData = { "to": LINE_GROUP_ID, "messages": [{ "type": "image", "originalContentUrl": imageUrl, "previewImageUrl": imageUrl }] };
  const options = {
    "method": "post",
    "headers": { "Content-Type": "application/json", "Authorization": "Bearer " + LINE_ACCESS_TOKEN },
    "payload": JSON.stringify(postData), "muteHttpExceptions": true
  };
  try { UrlFetchApp.fetch(url, options); } catch(e) {}
}
