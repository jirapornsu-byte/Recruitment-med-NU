// 📂 กำหนด Folder ID ปลายทางสำหรับเก็บไฟล์แนบ (ใช้เฉพาะ ID ตัวอักษรล้วน)
var FOLDER_ID = '1voNG8_kUVCGy_2zeFFPBt1AiXC_UguD4';

// 📊 กำหนด Spreadsheet ID ที่ใช้เก็บข้อมูล
var SHEET_ID = '1wKZZNiwxtC9yQSWLAjlN7q3CzWDTIP7daB8orWLxMeU';

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('ระบบสรรหาบุคลากร คณะแพทยศาสตร์ มน.')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
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

    // ตรวจสอบว่าเคยสมัครไปหรือยัง (เช็คคอลัมน์ A)
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      var sheetId = data[i][0].toString().replace(/[^0-9]/g, '').trim();
      if (sheetId === idClean) {
        return { success: false, message: 'เลขประจำตัวประชาชนนี้เคยส่งใบสมัครในระบบแล้ว ไม่สามารถส่งซ้ำได้' };
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
      appData.position || "",                                 // ตำแหน่งที่สมัคร
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

    return { success: true };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// 🔍 ฟังก์ชันค้นหาสถานะผู้สมัคร
function checkStatus(citizenId) {
  if (!citizenId) {
    return { success: false, message: '❌ ไม่พบข้อมูลการค้นหา' };
  }
  var idClean = citizenId.toString().replace(/[^0-9]/g, '').trim();
  var ss = SpreadsheetApp.openById(SHEET_ID);

  if (idClean.length !== 13) {
    return { success: false, message: '❌ รูปแบบเลขประจำตัวประชาชนไม่ถูกต้อง' };
  }

  // ชั้นที่ 1: ค้นหาจากบอร์ดอัปเดตสถานะของ HR (Applicants)
  // คอลัมน์: A=เลขบัตร, B=ชื่อ, C=ตำแหน่ง, D=สถานะ, E=กำหนดการสอบ, F=ลิงก์ประกาศ,
  //          G=เอกสารที่ขาด (กรอกโดย HR, คั่นด้วยจุลภาค), H=กำหนดส่งเอกสารเพิ่มภายในวันที่
  var appSheet = ss.getSheetByName('Applicants');
  if (appSheet) {
    var appData = appSheet.getDataRange().getValues();
    for (var i = 1; i < appData.length; i++) {
      if (!appData[i][0]) continue;
      var sheetId = appData[i][0].toString().replace(/[^0-9]/g, '').trim();

      if (sheetId === idClean) {
        var formattedExamDate = formatThaiDate(appData[i][4]);

        var missingDocsRaw = appData[i][6] ? appData[i][6].toString().trim() : '';
        var missingDocsList = missingDocsRaw ? missingDocsRaw.split(',').map(function(s){ return s.trim(); }).filter(function(s){ return s; }) : [];
        var deadlineRaw = appData[i][7];
        var deadlineText = deadlineRaw ? formatThaiDate(deadlineRaw) : '';

        return {
          success: true,
          name: appData[i][1] ? appData[i][1].toString() : 'ไม่ระบุชื่อ',
          position: appData[i][2] ? appData[i][2].toString() : 'ไม่ระบุตำแหน่ง',
          status: appData[i][3] ? appData[i][3].toString() : 'กำลังประมวลผล',
          examDate: formattedExamDate,
          announcementUrl: appData[i][5] ? appData[i][5].toString().trim() : '',
          missingDocs: missingDocsList,
          documentDeadline: deadlineText
        };
      }
    }
  }

  // ชั้นที่ 2: ถ้าในบอร์ด HR ยังไม่มี ให้ไปค้นหาในไฟล์ดิบ "ส่งใบสมัคร"
  var regSheet = ss.getSheetByName('ส่งใบสมัคร');
  if (regSheet) {
    var regData = regSheet.getDataRange().getValues();
    for (var j = 1; j < regData.length; j++) {
      if (!regData[j][0]) continue;
      var regId = regData[j][0].toString().replace(/[^0-9]/g, '').trim();

      if (regId === idClean) {
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

        return {
          success: true,
          name: regData[j][4] ? regData[j][4].toString() : 'ไม่ระบุชื่อ',
          position: regData[j][1] ? regData[j][1].toString() : 'ไม่ระบุตำแหน่ง',
          status: '⏳ ได้รับข้อมูลใบสมัครแล้ว (อยู่ระหว่างตรวจสอบเอกสาร)',
          examDate: 'จะประกาศกำหนดการสอบให้ทราบในขั้นตอนถัดไป',
          announcementUrl: '',
          missingDocs: autoMissing,
          documentDeadline: ''
        };
      }
    }
  }

  return { success: false, message: '❌ ไม่พบข้อมูลการสมัครงานในระบบ กรุณาไปที่แท็บ "ส่งใบสมัคร" เพื่อลงทะเบียนก่อนใช้งานระบบติดตามสถานะ' };
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

// 👨‍💼 ฟังก์ชันดึงข้อมูลผู้สมัครสำหรับเจ้าหน้าที่ (Staff Panel)
// คืนข้อมูลทุกฟิลด์จากชีต "ส่งใบสมัคร" ตามเลขประจำตัวประชาชน
function getApplicantData(citizenId) {
  if (!citizenId) return { success: false, message: '❌ ไม่พบข้อมูลการค้นหา' };
  var idClean = citizenId.toString().replace(/[^0-9]/g, '').trim();
  if (idClean.length !== 13) return { success: false, message: '❌ รูปแบบเลขประจำตัวประชาชนไม่ถูกต้อง' };

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('ส่งใบสมัคร');
  if (!sheet) return { success: false, message: 'ไม่พบชีต "ส่งใบสมัคร"' };

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    var rowId = data[i][0].toString().replace(/[^0-9]/g, '').trim();
    if (rowId !== idClean) continue;

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
      militaryStatus:  r[24] || '', ordainStatus:  r[26] || '',
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
      lawsuit:         r[82] || '', discDetail:    r[84] || '',
      specialWorks:    r[85] || '',
      expectedSalary:  r[86] || '',
      licenseUrl:      r[87] || '', transcriptUrl: r[88] || '', degreeUrl:     r[89] || '',
      workCertUrl:     r[90] || '', idCardUrl:     r[91] || '', houseRegUrl:   r[92] || '',
      nameChangeUrl:   r[93] || '', medicalUrl:    r[94] || '', photoUrl:      r[96] || '',
      prevApplied:     r[99] || '',
      readiness:       r[102] || '', readinessDate: r[103] || ''
    };
  }
  return { success: false, message: '❌ ไม่พบข้อมูลผู้สมัครที่มีเลขประจำตัวประชาชนนี้ในระบบ' };
}
