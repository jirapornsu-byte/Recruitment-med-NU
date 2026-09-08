// ============================================================
//  Auto-deploy: ดึงโค้ดล่าสุดจาก GitHub (branch main) มาอัปเดตเว็บแอปจริงให้อัตโนมัติ
//  ไฟล์นี้ไม่ได้เป็นส่วนหนึ่งของแอปหลัก — ให้สร้างเป็น Apps Script "โปรเจกต์แยกต่างหาก"
//  แล้ววาง code นี้ทับ แก้ค่า 2 ตัวด้านล่างให้ตรงกับโปรเจกต์จริงก่อนใช้งาน
// ============================================================

// 🔧 ตั้งค่า 2 ค่านี้ก่อนใช้งาน (หาได้จาก Apps Script Editor ของแอปหลัก)
var TARGET_SCRIPT_ID = 'ใส่ SCRIPT ID ของโปรเจกต์ Recruitment-med-NU ตรงนี้';   // Project Settings (⚙️) > Script ID
var DEPLOYMENT_ID    = 'ใส่ DEPLOYMENT ID ของเว็บแอปที่ใช้งานอยู่ตรงนี้';        // Deploy > Manage deployments

var GITHUB_OWNER  = 'jirapornsu-byte';
var GITHUB_REPO   = 'Recruitment-med-NU';
var GITHUB_BRANCH = 'main';

var FILES_TO_SYNC = [
  { name: 'Code',       type: 'SERVER_JS', path: 'Code.gs' },
  { name: 'index',      type: 'HTML',      path: 'index.html' },
  { name: 'appsscript', type: 'JSON',      path: 'appsscript.json' }
];

// ▶️ ฟังก์ชันหลัก — กด Run เอง หรือให้ Trigger เรียกอัตโนมัติ
function syncFromGithub() {
  if (TARGET_SCRIPT_ID.indexOf('ใส่') === 0 || DEPLOYMENT_ID.indexOf('ใส่') === 0) {
    throw new Error('กรุณาแก้ค่า TARGET_SCRIPT_ID และ DEPLOYMENT_ID ที่ด้านบนของไฟล์นี้ก่อนใช้งาน');
  }

  var token = ScriptApp.getOAuthToken();

  // 1) ดึงเนื้อหาไฟล์ล่าสุดจาก GitHub
  var files = FILES_TO_SYNC.map(function (f) {
    var url = 'https://raw.githubusercontent.com/' + GITHUB_OWNER + '/' + GITHUB_REPO + '/' + GITHUB_BRANCH + '/' + f.path;
    var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) {
      throw new Error('ดึงไฟล์ ' + f.path + ' จาก GitHub ไม่สำเร็จ (HTTP ' + res.getResponseCode() + ')');
    }
    return { name: f.name, type: f.type, source: res.getContentText() };
  });

  // 2) อัปเดตเนื้อหาไฟล์ในโปรเจกต์เป้าหมาย
  var contentUrl = 'https://script.googleapis.com/v1/projects/' + TARGET_SCRIPT_ID + '/content';
  var updateResp = UrlFetchApp.fetch(contentUrl, {
    method: 'put',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({ files: files }),
    muteHttpExceptions: true
  });
  if (updateResp.getResponseCode() !== 200) {
    throw new Error('อัปเดตโค้ดไม่สำเร็จ: ' + updateResp.getContentText());
  }

  // 3) สร้างเวอร์ชันใหม่จากโค้ดที่เพิ่งอัปเดต
  var versionUrl = 'https://script.googleapis.com/v1/projects/' + TARGET_SCRIPT_ID + '/versions';
  var versionResp = UrlFetchApp.fetch(versionUrl, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({ description: 'Auto-sync from GitHub ' + new Date().toISOString() }),
    muteHttpExceptions: true
  });
  if (versionResp.getResponseCode() !== 200) {
    throw new Error('สร้างเวอร์ชันใหม่ไม่สำเร็จ: ' + versionResp.getContentText());
  }
  var versionNumber = JSON.parse(versionResp.getContentText()).versionNumber;

  // 4) อัปเดต deployment เดิม (เว็บแอป URL เดิม) ให้ชี้ไปเวอร์ชันใหม่
  var deployUrl = 'https://script.googleapis.com/v1/projects/' + TARGET_SCRIPT_ID + '/deployments/' + DEPLOYMENT_ID;
  var deployResp = UrlFetchApp.fetch(deployUrl, {
    method: 'put',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({
      deploymentConfig: {
        scriptId: TARGET_SCRIPT_ID,
        versionNumber: versionNumber,
        manifestFileName: 'appsscript',
        description: 'Auto-sync from GitHub'
      }
    }),
    muteHttpExceptions: true
  });
  if (deployResp.getResponseCode() !== 200) {
    throw new Error('อัปเดต deployment ไม่สำเร็จ: ' + deployResp.getContentText());
  }

  Logger.log('✅ ซิงก์สำเร็จ — เว็บแอปอัปเดตเป็นเวอร์ชัน ' + versionNumber + ' แล้ว');
}
