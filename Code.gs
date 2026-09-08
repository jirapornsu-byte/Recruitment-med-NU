// 📂 กำหนด Folder ID ปลายทางสำหรับเก็บไฟล์แนบ (ใช้เฉพาะ ID ตัวอักษรล้วน)
var FOLDER_ID = '1voNG8_kUVCGy_2zeFFPBt1AiXC_UguD4';

// 📊 กำหนด Spreadsheet ID ที่ใช้เก็บข้อมูล
var SHEET_ID = '1wKZZNiwxtC9yQSWLAjlN7q3CzWDTIP7daB8orWLxMeU';

// สถานะมาตรฐานที่ HR เลือกใช้ได้ (แสดงในหน้าเจ้าหน้าที่)
var STATUS_OPTIONS = [
  '⏳ ได้รับใบสมัครแล้ว (รอตรวจสอบเอกสาร)',
  '📄 เอกสารไม่ครบ ต้องส่งเพิ่มเติม',
  '🔍 อยู่ระหว่างพิจารณาคุณสมบัติ',
  '🗓️ นัดหมายสอบ/สัมภาษณ์',
  '✅ ผ่านการคัดเลือก',
  '❌ ไม่ผ่านการคัดเลือก',
  '🚫 สละสิทธิ์ / ยกเลิกการสมัคร'
];

function doGet(e) {
  var t = HtmlService.createTemplateFromFile('index');
  t.shareToken = (e && e.parameter && e.parameter.share) ? e.parameter.share.toString() : '';
  return t.evaluate()
    .setTitle('ระบบสรรหาบุคลากร คณะแพทยศาสตร์ มน.')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 🗂️ ชีต "Applicants" คือบอร์ดสถานะที่ HR ใช้ติดตาม/อัปเดต — ระบบ auto-sync ให้อัตโนมัติทุกครั้งที่มีผู้สมัครส่งใบสมัครใหม่
// คอลัมน์: A=เลขบัตร, B=ชื่อ, C=ตำแหน่ง, D=สถานะ, E=กำหนดการสอบ, F=ลิงก์ประกาศ,
//          G=เอกสารที่ขาด(คั่นด้วยจุลภาค), H=กำหนดส่งเอกสารเพิ่ม, I=หน่วยงาน,
//          J=วันที่ส่งใบสมัคร, K=อัปเดตล่าสุด, L=Share Token
function ensureApplicantsSheet(ss) {
  var sheet = ss.getSheetByName('Applicants');
  if (!sheet) {
    sheet = ss.insertSheet('Applicants');
    sheet.appendRow(['เลขประจำตัวประชาชน','ชื่อ-นามสกุล','ตำแหน่ง','สถานะ','กำหนดการสอบ','ลิงก์ประกาศ','เอกสารที่ขาด','กำหนดส่งเอกสารเพิ่ม','หน่วยงาน','วันที่ส่งใบสมัคร','อัปเดตล่าสุด','Share Token']);
  }
  return sheet;
}

function findApplicantRow(sheet, idClean, position) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var rid = data[i][0] ? data[i][0].toString().replace(/[^0-9]/g, '').trim() : '';
    var rpos = data[i][2] ? data[i][2].toString().trim() : '';
    if (rid === idClean && rpos === position) return { rowIndex: i + 1, values: data[i] };
  }
  return null;
}

// 📝 ฟังก์ชันบันทึกการส่งใบสมัครใหม่ลงในชีต "ส่งใบสมัคร"
// คอลัมน์ในชีตเรียงตามลำดับ array ที่ appendRow ด้านล่าง (A, B, C, ... ตามลำดับ)
function submitApplication(appData) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName('ส่งใบสมัคร');

    if (!sheet) {
      return { success: false, message: 'ไม่พบแผ่นงานชื่อ "ส่งใบสมัคร" กรุณาตรวจสอบการตั้งชื่อแท็บ' };
    }
    if (!appData) {
      return { success: false, message: 'ไม่พบข้อมูลที่ส่งมาจากหน้าเว็บ' };
    }
    if (!appData.citizenId) {
      return { success: false, message: 'กรุณากรอกเลขประจำตัวประชาชน' };
    }

    var data = sheet.getDataRange().getValues();
    var idClean = appData.citizenId.toString().replace(/[^0-9]/g, '').trim();

    if (idClean.length !== 13) {
      return { success: false, message: 'เลขประจำตัวประชาชนไม่ถูกต้อง ต้องมี 13 หลัก' };
    }

    var newPosition = (appData.position || '').toString().trim();

    // ตรวจสอบว่าเคยสมัคร "ตำแหน่งเดียวกัน" ไปหรือยัง (เช็คคอลัมน์ A=เลขบัตร, B=ตำแหน่ง)
    // เลขบัตรเดียวกันสามารถสมัครต่างตำแหน่งซ้ำได้ ห้ามเฉพาะสมัครตำแหน่งเดิมซ้ำเท่านั้น
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      var sheetId = data[i][0].toString().replace(/[^0-9]/g, '').trim();
      var sheetPosition = data[i][1] ? data[i][1].toString().trim() : '';
      if (sheetId === idClean && sheetPosition === newPosition) {
        return { success: false, message: 'เลขประจำตัวประชาชนนี้เคยส่งใบสมัครในตำแหน่ง "' + newPosition + '" แล้ว ไม่สามารถส่งซ้ำตำแหน่งเดิมได้' };
      }
    }

    // 🗂️ เรียกใช้งานโฟลเดอร์ Google Drive
    var folder = DriveApp.getFolderById(FOLDER_ID);

    function uploadFile(fileData, prefixName) {
      if (!fileData || !fileData.data || fileData.data === "") return "ไม่ได้แนบเอกสาร";
      try {
        var contentType = fileData.data.substring(5, fileData.data.indexOf(';'));
        var bytes = Utilities.base64Decode(fileData.data.substr(fileData.data.indexOf('base64,') + 7));
        var safeName = appData.name ? appData.name.toString().trim() : "ไม่ระบุชื่อ";
        var blob = Utilities.newBlob(bytes, contentType, prefixName + "_" + safeName + "_" + fileData.name);
        var file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        return file.getUrl();
      } catch(e) {
        return "อัปโหลดไม่สำเร็จ";
      }
    }

    // ทำการอัปโหลดเอกสารตามรายการในข้อ 6 ของแบบฟอร์ม
    var licenseUrl = uploadFile(appData.license, "1_ใบอนุญาตประกอบวิชาชีพ");
    var transcriptUrl = uploadFile(appData.transcript, "2_Transcript");
    var degreeUrl = uploadFile(appData.degree, "3_ใบปริญญา");
    var workCertUrl = uploadFile(appData.workCert, "4_ใบผ่านงาน");
    var idCardUrl = uploadFile(appData.idCard, "5_บัตรปชช");
    var houseRegUrl = uploadFile(appData.houseReg, "6_ทะเบียนบ้าน");
    var nameChangeUrl = uploadFile(appData.nameChange, "7_เปลี่ยนชื่อ");
    var medicalUrl = uploadFile(appData.medical, "8_ใบรับรองแพทย์");
    var militaryFileUrl = uploadFile(appData.militaryFile, "9_ผ่านเกณฑ์ทหาร");
    var photoUrl = uploadFile(appData.photo, "10_รูปถ่าย");
    var otherDocUrl = uploadFile(appData.otherDoc, "11_เอกสารอื่นๆ");

    var safePhone = appData.phone ? "'" + appData.phone.toString().replace(/[^0-9]/g, '').trim() : "";
    var safePhoneHome = appData.phoneHome ? "'" + appData.phoneHome.toString().replace(/[^0-9]/g, '').trim() : "";
    var safePhone2 = appData.phone2 ? "'" + appData.phone2.toString().replace(/[^0-9]/g, '').trim() : "";

    // บันทึกข้อมูลลงแถวใหม่ — เรียงตามหัวข้อ 1-8 ของแบบฟอร์ม
    sheet.appendRow([
      "'" + idClean,                                          // เลขประจำตัวประชาชน
      newPosition,                                             // ตำแหน่งที่สมัคร
      appData.department || "",                                // หน่วยงาน
      // ----- 1. ประวัติส่วนตัว -----
      appData.title || "",
      appData.name ? appData.name.toString().trim() : "",
      appData.nickname || "",
      appData.nameEN || "",
      appData.nationality || "",
      appData.race || "",
      appData.religion || "",
      safePhoneHome,
      safePhone,
      safePhone2,
      appData.fax || "",
      appData.email ? appData.email.toString().trim() : "",
      appData.dob || "",
      appData.age || "",
      appData.birthPlace || "",
      appData.weight || "",
      appData.height || "",
      appData.bloodGroup || "",
      appData.disease || "",
      appData.smoking || "",
      appData.alcohol || "",
      appData.militaryStatus || "",
      appData.militaryYear || "",
      appData.ordainStatus || "",
      appData.fatherName || "",
      appData.fatherSurname || "",
      appData.fatherStatus || "",
      appData.fatherOccupation || "",
      appData.motherName || "",
      appData.motherSurname || "",
      appData.motherStatus || "",
      appData.motherOccupation || "",
      appData.maritalStatus || "",
      appData.spouseName || "",
      appData.spouseOccupation || "",
      appData.spouseWorkplace || "",
      appData.spousePhone || "",
      appData.hasChildren || "",
      appData.childrenText || "",
      appData.addressNo || "",
      appData.addressMoo || "",
      appData.addressSoi || "",
      appData.addressRoad || "",
      appData.addressTambon || "",
      appData.addressAmphoe || "",
      appData.addressProvince || "",
      appData.addressZip || "",
      appData.emgName || "",
      appData.emgRelation || "",
      appData.emgAddress || "",
      appData.emgPhoneHome || "",
      appData.emgPhoneMobile || "",
      appData.emgEmail || "",
      // ----- 2. ประวัติการศึกษา -----
      appData.educationText || "",
      appData.studying || "",
      appData.studyingText || "",
      appData.studyLeave || "",
      appData.studyFund || "",
      appData.studyFundDetail || "",
      appData.training || "",
      appData.trainingText || "",
      // ----- 3. ความรู้ความสามารถพิเศษ -----
      appData.langText || "",
      appData.localLang || "",
      appData.localLangLevel || "",
      appData.typeThai || "",
      appData.typeEng || "",
      appData.computer || "",
      appData.computerPrograms || "",
      appData.otherSkills || "",
      appData.hobbies || "",
      // ----- 4. ประวัติการทำงาน -----
      appData.experienceText || "",
      appData.decoration || "",
      appData.decorationText || "",
      appData.refName || "",
      appData.refRelation || "",
      appData.refPhone || "",
      appData.discPast || "",
      appData.discInvestigate || "",
      appData.discCommittee || "",
      appData.lawsuit || "",
      appData.lawsuitStage || "",
      appData.discDetail || "",
      // ----- 5. ผลงานความรู้ความเชี่ยวชาญ -----
      appData.specialWorks || "",
      // ----- เงินเดือนที่ต้องการ -----
      appData.expectedSalary || "",
      // ----- 6. เอกสารแนบ (ลิงก์ไฟล์) -----
      licenseUrl,
      transcriptUrl,
      degreeUrl,
      workCertUrl,
      idCardUrl,
      houseRegUrl,
      nameChangeUrl,
      medicalUrl,
      militaryFileUrl,
      photoUrl,
      otherDocUrl,
      appData.otherDocDetail || "",
      // ----- 7. ประวัติการสมัครก่อนหน้า -----
      appData.prevApplied || "",
      appData.prevAppliedTimes || "",
      appData.prevAppliedText || "",
      // ----- 8. ความพร้อมในการทำงาน -----
      appData.readiness || "",
      appData.readinessDate || "",
      appData.readinessOther || "",
      // ----- เวลาที่ส่ง -----
      new Date()
    ]);

    // 🔄 Auto-sync: สร้างแถวในบอร์ดสถานะ "Applicants" ให้อัตโนมัติ
    // เพื่อไม่ให้ HR ต้องคีย์ข้อมูลผู้สมัครซ้ำสองรอบ (ในชีตสมัคร + บอร์ดสถานะ)
    try {
      var appSheet = ensureApplicantsSheet(ss);
      var now = new Date();
      appSheet.appendRow([
        "'" + idClean,
        appData.name ? appData.name.toString().trim() : "",
        newPosition,
        STATUS_OPTIONS[0],
        "", "", "", "",
        appData.department || "",
        now, now, ""
      ]);
    } catch (syncErr) {
      // ไม่ให้การ sync ล้มเหลวไปกระทบการส่งใบสมัครหลัก
    }

    return { success: true };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// 🔍 ฟังก์ชันค้นหาสถานะผู้สมัคร — คืนค่าเป็นรายการ "ทุกตำแหน่ง" ที่เลขบัตรนี้เคยสมัคร
function checkStatus(citizenId) {
  if (!citizenId) {
    return { success: false, message: '❌ ไม่พบข้อมูลการค้นหา' };
  }
  var idClean = citizenId.toString().replace(/[^0-9]/g, '').trim();
  var ss = SpreadsheetApp.openById(SHEET_ID);

  if (idClean.length !== 13) {
    return { success: false, message: '❌ รูปแบบเลขประจำตัวประชาชนไม่ถูกต้อง' };
  }

  // ชั้นที่ 1: ดึงสถานะจากบอร์ดอัปเดตของ HR (Applicants) มาเก็บตามตำแหน่ง
  var statusByPosition = {};
  var appSheet = ss.getSheetByName('Applicants');
  if (appSheet) {
    var appData = appSheet.getDataRange().getValues();
    for (var i = 1; i < appData.length; i++) {
      if (!appData[i][0]) continue;
      var sheetId = appData[i][0].toString().replace(/[^0-9]/g, '').trim();
      if (sheetId !== idClean) continue;
      var pos = appData[i][2] ? appData[i][2].toString().trim() : '';

      var missingDocsRaw = appData[i][6] ? appData[i][6].toString().trim() : '';
      var missingDocsList = missingDocsRaw ? missingDocsRaw.split(',').map(function(s){ return s.trim(); }).filter(function(s){ return s; }) : [];
      var deadlineRaw = appData[i][7];
      var deadlineText = deadlineRaw ? formatThaiDate(deadlineRaw) : '';

      statusByPosition[pos] = {
        status: appData[i][3] ? appData[i][3].toString() : STATUS_OPTIONS[0],
        examDate: formatThaiDate(appData[i][4]),
        announcementUrl: appData[i][5] ? appData[i][5].toString().trim() : '',
        missingDocs: missingDocsList,
        documentDeadline: deadlineText
      };
    }
  }

  // ชั้นที่ 2: ไล่ดูใบสมัครทั้งหมดของเลขบัตรนี้จากไฟล์ดิบ "ส่งใบสมัคร" (อาจมีหลายตำแหน่ง)
  var applications = [];
  var regSheet = ss.getSheetByName('ส่งใบสมัคร');
  if (regSheet) {
    var regData = regSheet.getDataRange().getValues();
    for (var j = 1; j < regData.length; j++) {
      if (!regData[j][0]) continue;
      var regId = regData[j][0].toString().replace(/[^0-9]/g, '').trim();
      if (regId !== idClean) continue;

      var position = regData[j][1] ? regData[j][1].toString().trim() : 'ไม่ระบุตำแหน่ง';
      var name = regData[j][4] ? regData[j][4].toString() : 'ไม่ระบุชื่อ';
      var known = statusByPosition[position];

      if (known) {
        applications.push({
          name: name, position: position,
          status: known.status, examDate: known.examDate,
          announcementUrl: known.announcementUrl,
          missingDocs: known.missingDocs, documentDeadline: known.documentDeadline
        });
      } else {
        // ตรวจสอบเอกสารบังคับที่อัปโหลดไม่สำเร็จหรือไม่ได้แนบ (ตรวจอัตโนมัติจากข้อมูลตอนสมัคร)
        var docChecks = [
          { idx: 88, label: 'สำเนาระเบียนแสดงผลการศึกษา (Transcript)' },
          { idx: 89, label: 'สำเนาใบปริญญาบัตร/หนังสือรับรองวุฒิการศึกษา' },
          { idx: 91, label: 'สำเนาบัตรประจำตัวประชาชน' },
          { idx: 92, label: 'สำเนาทะเบียนบ้าน' },
          { idx: 94, label: 'ใบรับรองแพทย์แผนปัจจุบันสาขาเวชกรรม' },
          { idx: 96, label: 'รูปถ่าย 1 นิ้ว' }
        ];
        var autoMissing = [];
        docChecks.forEach(function(d) {
          var val = regData[j][d.idx] ? regData[j][d.idx].toString() : '';
          if (val === 'ไม่ได้แนบเอกสาร' || val === 'อัปโหลดไม่สำเร็จ' || val === '') {
            autoMissing.push(d.label);
          }
        });

        applications.push({
          name: name, position: position,
          status: STATUS_OPTIONS[0],
          examDate: 'จะประกาศกำหนดการสอบให้ทราบในขั้นตอนถัดไป',
          announcementUrl: '',
          missingDocs: autoMissing, documentDeadline: ''
        });
      }
    }
  }

  if (!applications.length) {
    return { success: false, message: '❌ ไม่พบข้อมูลการสมัครงานในระบบ กรุณาไปที่แท็บ "ส่งใบสมัคร" เพื่อลงทะเบียนก่อนใช้งานระบบติดตามสถานะ' };
  }
  return { success: true, applications: applications };
}

// 🛠️ ฟังก์ชันเสริม (Helper Function)
function formatThaiDate(dateVal) {
  if (!dateVal) return 'จะประกาศกำหนดการสอบให้ทราบในขั้นตอนถัดไป';
  if (!(dateVal instanceof Date)) return dateVal.toString();
  try {
    var days = ['วันอาทิตย์', 'วันจันทร์', 'วันอังคาร', 'วันพุธ', 'วันพฤหัสบดี', 'วันศุกร์', 'วันเสาร์'];
    var months = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];
    var dayName = days[dateVal.getDay()];
    var dayNum = dateVal.getDate();
    var monthName = months[dateVal.getMonth()];
    var yearTH = dateVal.getFullYear() + 543;
    return dayName + 'ที่ ' + dayNum + ' ' + monthName + ' ' + yearTH;
  } catch(err) {
    return dateVal.toString();
  }
}

// 🧠 ระบบ AI Sentiment Analysis Engine
function analyzeSentimentAI(text) {
  if (!text) return 'ทั่วไป/คำแนะนำ';
  var positiveWords = ['ดี', 'เร็ว', 'ง่าย', 'สะดวก', 'ประทับใจ', 'ยอดเยี่ยม', 'ชอบ', 'ชัดเจน'];
  var negativeWords = ['ช้า', 'งง', 'ยาก', 'ปรับปรุง', 'ไม่เข้าใจ', 'รอนาน', 'ค้าง', 'สับสน'];
  var posCount = 0, negCount = 0;
  var lowerText = text.toLowerCase();

  positiveWords.forEach(function(word) { if (lowerText.indexOf(word) !== -1) posCount++; });
  negativeWords.forEach(function(word) { if (lowerText.indexOf(word) !== -1) negCount++; });

  if (negCount > posCount) return '🔴 เชิงลบ (ต้องการการปรับปรุงด่วน)';
  if (posCount > negCount) return '🟢 เชิงบวก (พึงพอใจสูง)';
  return '🟡 เชิงทั่วไป/แนะนำ';
}

// 📝 ฟังก์ชันบันทึกแบบประเมิน
function saveSurvey(citizenId, rating, feedback) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName('Surveys') || ss.insertSheet('Surveys');
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['เลขประจำตัวประชาชน', 'คะแนนความพึงพอใจ', 'ข้อเสนอแนะเพิ่มเติม', 'ผลวิเคราะห์อารมณ์โดย AI', 'วัน-เวลาที่ประเมิน']);
    }
    var aiResult = analyzeSentimentAI(feedback);
    if (!citizenId) citizenId = "";
    sheet.appendRow(["'" + citizenId.toString().replace(/[^0-9]/g, ''), rating, feedback, aiResult, new Date()]);
    return { success: true };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// 🔴 ฟังก์ชันกระตุ้นสิทธิ์การเข้าถึง (รันเองครั้งแรกใน Apps Script Editor)
function forceAuth() {
  DriveApp.getFolderById(FOLDER_ID);
  SpreadsheetApp.openById(SHEET_ID);
}

// 👨‍💼 ฟังก์ชันดึง "รายการตำแหน่งที่เคยสมัคร" ของเลขบัตรหนึ่งใบ ใช้เมื่อเลขบัตรเดียวสมัครหลายตำแหน่ง
// เพื่อให้เจ้าหน้าที่เลือกใบสมัครที่ต้องการดูได้ถูกต้อง
function getApplicantPositions(citizenId) {
  if (!citizenId) return { success: false, message: '❌ ไม่พบข้อมูลการค้นหา' };
  var idClean = citizenId.toString().replace(/[^0-9]/g, '').trim();
  if (idClean.length !== 13) return { success: false, message: '❌ รูปแบบเลขประจำตัวประชาชนไม่ถูกต้อง' };

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('ส่งใบสมัคร');
  if (!sheet) return { success: false, message: 'ไม่พบชีต "ส่งใบสมัคร"' };

  var data = sheet.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    var rowId = data[i][0].toString().replace(/[^0-9]/g, '').trim();
    if (rowId !== idClean) continue;
    list.push({
      position: data[i][1] || '',
      department: data[i][2] || '',
      name: data[i][4] || '',
      submittedAt: data[i][105] instanceof Date ? Utilities.formatDate(data[i][105], 'Asia/Bangkok', 'dd/MM/yyyy HH:mm') : ''
    });
  }
  if (!list.length) return { success: false, message: '❌ ไม่พบข้อมูลผู้สมัครที่มีเลขประจำตัวประชาชนนี้ในระบบ' };
  return { success: true, list: list };
}

// 👨‍💼 ฟังก์ชันดึงข้อมูลผู้สมัครสำหรับเจ้าหน้าที่ (Staff Panel)
// คืนข้อมูลทุกฟิลด์จากชีต "ส่งใบสมัคร" ตามเลขประจำตัวประชาชน + ตำแหน่งที่สมัคร (ระบุ position เพื่อเจาะจงใบสมัครที่ต้องการเมื่อสมัครหลายตำแหน่ง)
function getApplicantData(citizenId, position) {
  if (!citizenId) return { success: false, message: '❌ ไม่พบข้อมูลการค้นหา' };
  var idClean = citizenId.toString().replace(/[^0-9]/g, '').trim();
  if (idClean.length !== 13) return { success: false, message: '❌ รูปแบบเลขประจำตัวประชาชนไม่ถูกต้อง' };
  var posFilter = position ? position.toString().trim() : '';

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('ส่งใบสมัคร');
  if (!sheet) return { success: false, message: 'ไม่พบชีต "ส่งใบสมัคร"' };

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    var rowId = data[i][0].toString().replace(/[^0-9]/g, '').trim();
    if (rowId !== idClean) continue;
    if (posFilter && (data[i][1] ? data[i][1].toString().trim() : '') !== posFilter) continue;

    var r = data[i];
    // คอลัมน์เรียงตาม appendRow ใน submitApplication:
    // 0=citizenId, 1=position, 2=department,
    // 3=title, 4=name, 5=nickname, 6=nameEN,
    // 7=nationality, 8=race, 9=religion, 10=phoneHome, 11=phone, 12=phone2, 13=fax, 14=email,
    // 15=dob, 16=age, 17=birthPlace,
    // 18=weight, 19=height, 20=bloodGroup, 21=disease, 22=smoking, 23=alcohol,
    // 24=militaryStatus, 25=militaryYear, 26=ordainStatus,
    // 27=fatherName, 28=fatherSurname, 29=fatherStatus, 30=fatherOccupation,
    // 31=motherName, 32=motherSurname, 33=motherStatus, 34=motherOccupation,
    // 35=maritalStatus, 36=spouseName, 37=spouseOccupation, 38=spouseWorkplace, 39=spousePhone,
    // 40=hasChildren, 41=childrenText,
    // 42=addressNo, 43=addressMoo, 44=addressSoi, 45=addressRoad,
    // 46=addressTambon, 47=addressAmphoe, 48=addressProvince, 49=addressZip,
    // 50=emgName, 51=emgRelation, 52=emgAddress, 53=emgPhoneHome, 54=emgPhoneMobile, 55=emgEmail,
    // 56=educationText, 57=studying, 58=studyingText, 59=studyLeave, 60=studyFund, 61=studyFundDetail,
    // 62=training, 63=trainingText,
    // 64=langText, 65=localLang, 66=localLangLevel, 67=typeThai, 68=typeEng,
    // 69=computer, 70=computerPrograms, 71=otherSkills, 72=hobbies,
    // 73=experienceText,
    // 74=decoration, 75=decorationText,
    // 76=refName, 77=refRelation, 78=refPhone,
    // 79=discPast, 80=discInvestigate, 81=discCommittee, 82=lawsuit, 83=lawsuitStage, 84=discDetail,
    // 85=specialWorks,
    // 86=expectedSalary,
    // 87=licenseUrl, 88=transcriptUrl, 89=degreeUrl, 90=workCertUrl,
    // 91=idCardUrl, 92=houseRegUrl, 93=nameChangeUrl, 94=medicalUrl, 95=militaryFileUrl, 96=photoUrl, 97=otherDocUrl, 98=otherDocDetail,
    // 99=prevApplied, 100=prevAppliedTimes, 101=prevAppliedText,
    // 102=readiness, 103=readinessDate, 104=readinessOther,
    // 105=timestamp

    return {
      success: true,
      citizenId:       r[0]  ? r[0].toString().replace(/[^0-9]/g,'') : '',
      position:        r[1]  || '', department:    r[2]  || '',
      title:           r[3]  || '', name:          r[4]  || '', nickname:      r[5]  || '', nameEN:        r[6]  || '',
      nationality:     r[7]  || '', race:          r[8]  || '', religion:      r[9]  || '',
      phoneHome:       r[10] || '', phone:         r[11] || '', phone2:        r[12] || '', email: r[14] || '',
      dob:             r[15] ? (r[15] instanceof Date ? Utilities.formatDate(r[15], 'Asia/Bangkok', 'dd/MM/yyyy') : r[15].toString()) : '',
      birthPlace:      r[17] || '',
      weight:          r[18] || '', height:        r[19] || '', bloodGroup:    r[20] || '', disease:       r[21] || '',
      smoking:         r[22] || '', alcohol:       r[23] || '',
      militaryStatus:  r[24] || '', militaryYear: r[25] || '', ordainStatus:  r[26] || '',
      fatherName:      r[27] || '', fatherSurname: r[28] || '', fatherStatus:  r[29] || '', fatherOccupation: r[30] || '',
      motherName:      r[31] || '', motherSurname: r[32] || '', motherStatus:  r[33] || '', motherOccupation: r[34] || '',
      maritalStatus:   r[35] || '', spouseName:    r[36] || '', spouseOccupation: r[37] || '', spousePhone: r[39] || '',
      hasChildren:     r[40] || '', childrenText:  r[41] || '',
      addressNo:       r[42] || '', addressMoo:    r[43] || '', addressSoi:    r[44] || '', addressRoad:   r[45] || '',
      addressTambon:   r[46] || '', addressAmphoe: r[47] || '', addressProvince: r[48] || '', addressZip:  r[49] || '',
      emgName:         r[50] || '', emgRelation:   r[51] || '', emgPhoneMobile: r[54] || '',
      educationText:   r[56] || '', studying:      r[57] || '', studyingText:  r[58] || '',
      training:        r[62] || '', trainingText:  r[63] || '',
      langText:        r[64] || '',
      typeThai:        r[67] || '', typeEng:       r[68] || '',
      computer:        r[69] || '', computerPrograms: r[70] || '',
      otherSkills:     r[71] || '', hobbies:       r[72] || '',
      experienceText:  r[73] || '',
      decoration:      r[74] || '', decorationText: r[75] || '',
      refName:         r[76] || '', refRelation:   r[77] || '', refPhone:      r[78] || '',
      discPast:        r[79] || '', discInvestigate: r[80] || '', discCommittee: r[81] || '',
      lawsuit:         r[82] || '', lawsuitStage:  r[83] || '', discDetail:    r[84] || '',
      specialWorks:    r[85] || '',
      expectedSalary:  r[86] || '',
      licenseUrl:      r[87] || '', transcriptUrl: r[88] || '', degreeUrl:     r[89] || '',
      workCertUrl:     r[90] || '', idCardUrl:     r[91] || '', houseRegUrl:   r[92] || '',
      nameChangeUrl:   r[93] || '', medicalUrl:    r[94] || '', militaryFileUrl: r[95] || '', photoUrl: r[96] || '',
      otherDocUrl:     r[97] || '', otherDocDetail: r[98] || '',
      prevApplied:     r[99] || '', prevAppliedTimes: r[100] || '', prevAppliedText: r[101] || '',
      readiness:       r[102] || '', readinessDate: r[103] || '', readinessOther: r[104] || ''
    };
  }
  return { success: false, message: '❌ ไม่พบข้อมูลผู้สมัครที่มีเลขประจำตัวประชาชนนี้ในระบบ' };
}

// 📋 ฟังก์ชันดึงสถานะปัจจุบันของผู้สมัคร (จากบอร์ด Applicants) เพื่อเติมฟอร์มแก้ไขสถานะฝั่งเจ้าหน้าที่
function getApplicantStatus(citizenId, position) {
  var idClean = citizenId.toString().replace(/[^0-9]/g, '').trim();
  var pos = (position || '').toString().trim();
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Applicants');
  if (!sheet) return { success: true, found: false };
  var row = findApplicantRow(sheet, idClean, pos);
  if (!row) return { success: true, found: false };
  var r = row.values;
  var missing = r[6] ? r[6].toString().split(',').map(function(s){ return s.trim(); }).filter(function(s){ return s; }) : [];
  return {
    success: true, found: true,
    status: r[3] ? r[3].toString() : '',
    examDate: r[4] instanceof Date ? Utilities.formatDate(r[4], 'Asia/Bangkok', 'yyyy-MM-dd') : (r[4] ? r[4].toString() : ''),
    announcementUrl: r[5] ? r[5].toString() : '',
    missingDocs: missing,
    documentDeadline: r[7] instanceof Date ? Utilities.formatDate(r[7], 'Asia/Bangkok', 'yyyy-MM-dd') : (r[7] ? r[7].toString() : ''),
    shareToken: r[11] ? r[11].toString() : ''
  };
}

// ✏️ ฟังก์ชันให้เจ้าหน้าที่เปลี่ยนสถานะผู้สมัคร (เอกสารไม่ครบ / อยู่ระหว่างดำเนินการ / นัดสอบ ฯลฯ)
// สร้างแถวใหม่ในบอร์ด Applicants อัตโนมัติถ้ายังไม่มี (เผื่อกรณีข้อมูลเก่าก่อนมีระบบ auto-sync)
function updateApplicantStatus(payload) {
  try {
    if (!payload || !payload.citizenId || !payload.position) {
      return { success: false, message: 'ข้อมูลไม่ครบถ้วนสำหรับการอัปเดตสถานะ' };
    }
    var idClean = payload.citizenId.toString().replace(/[^0-9]/g, '').trim();
    var position = payload.position.toString().trim();
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ensureApplicantsSheet(ss);
    var row = findApplicantRow(sheet, idClean, position);
    var missingDocsStr = Array.isArray(payload.missingDocs) ? payload.missingDocs.join(', ') : (payload.missingDocs || '');
    var rowValues = [
      "'" + idClean,
      payload.name || (row ? row.values[1] : '') || '',
      position,
      payload.status || STATUS_OPTIONS[0],
      payload.examDate || '',
      payload.announcementUrl || '',
      missingDocsStr,
      payload.documentDeadline || '',
      payload.department || (row ? row.values[8] : '') || ''
    ];
    var now = new Date();
    if (!row) {
      sheet.appendRow(rowValues.concat([now, now, '']));
    } else {
      sheet.getRange(row.rowIndex, 1, 1, 9).setValues([rowValues]);
      sheet.getRange(row.rowIndex, 11).setValue(now);
    }
    return { success: true };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// 🔗 สร้างลิงก์แชร์แบบอ่านอย่างเดียวสำหรับผู้สมัครรายหนึ่ง เพื่อให้กรรมการสัมภาษณ์/ฝ่ายอื่นเปิดดูได้โดยไม่ต้องมีรหัสผ่านเจ้าหน้าที่
function createShareLink(citizenId, position) {
  try {
    var idClean = citizenId.toString().replace(/[^0-9]/g, '').trim();
    var pos = (position || '').toString().trim();
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ensureApplicantsSheet(ss);
    var row = findApplicantRow(sheet, idClean, pos);
    if (!row) {
      // ยังไม่มีแถวสถานะ (ไม่ควรเกิดขึ้นเพราะ auto-sync ตอนสมัคร) — สร้างให้ก่อน
      updateApplicantStatus({ citizenId: idClean, position: pos, status: STATUS_OPTIONS[0] });
      row = findApplicantRow(sheet, idClean, pos);
      if (!row) return { success: false, message: 'ไม่สามารถสร้างลิงก์แชร์ได้ กรุณาลองใหม่' };
    }
    var token = row.values[11] ? row.values[11].toString() : '';
    if (!token) {
      token = Utilities.getUuid().replace(/-/g, '');
      sheet.getRange(row.rowIndex, 12).setValue(token);
    }
    var url = ScriptApp.getService().getUrl() + '?share=' + token;
    return { success: true, url: url };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// 👁️ ดึงข้อมูลผู้สมัครแบบอ่านอย่างเดียวผ่านลิงก์แชร์ (ไม่ต้องใส่รหัสผ่านเจ้าหน้าที่)
function getApplicantByShareToken(token) {
  try {
    if (!token) return { success: false, message: 'ลิงก์ไม่ถูกต้อง' };
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName('Applicants');
    if (!sheet) return { success: false, message: 'ไม่พบข้อมูล' };
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][11] && data[i][11].toString() === token) {
        var citizenId = data[i][0].toString().replace(/[^0-9]/g, '');
        var position = data[i][2] ? data[i][2].toString() : '';
        var result = getApplicantData(citizenId, position);
        if (result.success) {
          result.shareStatus = data[i][3] || '';
        }
        return result;
      }
    }
    return { success: false, message: 'ลิงก์นี้ไม่ถูกต้องหรือถูกยกเลิกแล้ว' };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// 📊 ฟังก์ชันดึงรายการผู้สมัครทั้งหมด (สำหรับแดชบอร์ดคิวงานฝั่งเจ้าหน้าที่)
// ช่วยลดการค้นหาทีละเลขบัตร — เจ้าหน้าที่เห็นภาพรวมและคลิกเปิดได้ทันที
// filterStatus: ถ้าระบุ จะกรองเฉพาะสถานะนั้น (ส่ง '' หรือไม่ระบุ = ทั้งหมด)
function getApplicantsQueue(filterStatus) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName('Applicants');
    if (!sheet) return { success: true, list: [], statusOptions: STATUS_OPTIONS };
    var data = sheet.getDataRange().getValues();
    var list = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      var status = data[i][3] ? data[i][3].toString() : STATUS_OPTIONS[0];
      if (filterStatus && status !== filterStatus) continue;
      var submittedRaw = data[i][9] instanceof Date ? data[i][9].getTime() : 0;
      list.push({
        citizenId: data[i][0].toString().replace(/[^0-9]/g, ''),
        name: data[i][1] || '', position: data[i][2] || '', status: status,
        department: data[i][8] || '',
        submittedAt: data[i][9] instanceof Date ? Utilities.formatDate(data[i][9], 'Asia/Bangkok', 'dd/MM/yyyy HH:mm') : '',
        lastUpdatedAt: data[i][10] instanceof Date ? Utilities.formatDate(data[i][10], 'Asia/Bangkok', 'dd/MM/yyyy HH:mm') : '',
        _sortKey: submittedRaw
      });
    }
    // เรียงล่าสุดขึ้นก่อน
    list.sort(function(a, b) { return b._sortKey - a._sortKey; });
    list.forEach(function(item) { delete item._sortKey; });
    return { success: true, list: list, statusOptions: STATUS_OPTIONS };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}
