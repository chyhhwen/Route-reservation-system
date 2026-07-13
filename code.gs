// 全域設定
const CALENDAR_ID = "";
const SPREADSHEET_ID = "";
const SHEET_NAME = "user";
const ADMIN_SHEET = "administrator";
const MAX_SELF = 2;
const PREFIX_SELF = "[飯店]";
const PREFIX_CH = "[中華]";
const PREFIX_PH = "[電話]";
const PREFIX_WF = "[WF]";

function getUserPhoneMap() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) { Logger.log(`找不到工作表：${SHEET_NAME}`); return {}; }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};
  const data = sheet.getRange(1, 1, lastRow, 2).getValues();
  const map = {};
  data.forEach(row => {
    let email = (row[0] || "").toString().trim().toLowerCase();
    let phone = (row[1] || "").toString().trim();
    if (email && phone) map[email] = phone;
  });
  Logger.log(map);
  return map;
}

function getAdmin() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(ADMIN_SHEET);
  if (!sheet) { Logger.log(`找不到工作表：${ADMIN_SHEET}`); return []; }
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return [];
  return sheet.getRange(1, 1, lastRow).getValues();
}

function checkIsAdmin(userEmail) {
  const adminData = getAdmin();
  return adminData.some(row =>
    row[0] && String(row[0]).toLowerCase().trim() === userEmail.toLowerCase().trim()
  );
}

function getOrCreateFolder() {
  const folders = DriveApp.getFoldersByName("線路預約圖檔");
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder("線路預約圖檔");
}

function doGet(e) {
  var userEmail = Session.getActiveUser().getEmail();
  if (!userEmail) {
    return HtmlService.createHtmlOutput(`
      <h1 style="text-align:center; margin-top:100px; color:#d32f2f;">請先使用公司 Google 帳號登入</h1>
      <p style="text-align:center;">關閉此頁面後重新開啟，或確認已登入正確帳號。</p>
    `);
  }
  var template = HtmlService.createTemplateFromFile('index');
  template.userEmail = userEmail;
  return template.evaluate()
    .setTitle('臨時線路預約系統')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getUserEmail() {
  return Session.getActiveUser().getEmail();
}

function getAutoPhone() {
  const email = Session.getActiveUser().getEmail();
  if (!email) return '';
  const key = email.trim().toLowerCase();
  const phoneMap = getUserPhoneMap();
  return phoneMap[key] || '';
}

function processBooking(data, images) {
  var userEmail = Session.getActiveUser().getEmail() || '未知使用者';

  if (!data.startDate || !data.endDate || !data.type || !data.activityName ||
      !data.venue || !data.name) {
    return "資料不完整，請檢查所有欄位";
  }

  const typeMap = {
    'self': { prefix: PREFIX_SELF, max: MAX_SELF, label: '飯店專用線路', color: CalendarApp.EventColor.PALE_RED },
    'ch':   { prefix: PREFIX_CH,   label: '中華電信臨時線路',           color: CalendarApp.EventColor.CYAN },
    'ph':   { prefix: PREFIX_PH,   label: '臨時電話線路',               color: CalendarApp.EventColor.YELLOW },
    'wf':   { prefix: PREFIX_WF,   label: 'WIFI名稱更改',               color: CalendarApp.EventColor.ORANGE }
  };

  const typeInfo = typeMap[data.type];
  if (!typeInfo) return "無效線路類型";

  const start = new Date(data.startDate);
  const end   = new Date(data.endDate);
  const cal   = CalendarApp.getCalendarById(CALENDAR_ID);
  if (!cal) return "日曆失效，請檢查 CALENDAR_ID";

  const endForQuery = new Date(data.endDate);
  endForQuery.setDate(endForQuery.getDate() + 1);

  if (typeInfo.prefix === PREFIX_SELF) {
    const allEvents = cal.getEvents(start, endForQuery);
    const countByDate = {};
    allEvents.forEach(e => {
      if (e.getTitle().startsWith(typeInfo.prefix)) {
        onst evStart = e.getAllDayStartDate();
        const evEnd = e.getAllDayEndDate(); // exclusive end date
        let d = new Date(evStart);
        while (d < evEnd) {
          const dStr = Utilities.formatDate(d, "Asia/Taipei", "yyyy-MM-dd");
          countByDate[dStr] = (countByDate[dStr] || 0) + 1;
          d.setDate(d.getDate() + 1);
        }
      }
    });

    let hasFullDay = false;
    let fullDate = '';
    let current = new Date(start);
    while (current <= end) {
      const dateStr = Utilities.formatDate(current, "Asia/Taipei", "yyyy-MM-dd");
      const count = countByDate[dateStr] || 0;
      const lineCount = parseInt(data.line) || 1;
      if ((lineCount === 2 && count >= 1) || count >= typeInfo.max) {
        hasFullDay = true;
        fullDate = dateStr;
        break;
      }
      current.setDate(current.getDate() + 1);
    }

    if (hasFullDay) {
      return `${fullDate} 的${typeInfo.label}已滿 (上限 ${typeInfo.max} 條)，無法預約整個區間`;
    }
  }

  let imageLinks = [];
  if (images && images.length > 0) {
    const folder = getOrCreateFolder();
    images.forEach((file, idx) => {
      try 
      {
        const bytes = Utilities.base64Decode(file.base64);
        const filenameLower = file.name.toLowerCase();
        let mimeType = 'application/octet-stream';
        let ext = '';
        if (filenameLower.endsWith('.pdf')) { mimeType = 'application/pdf'; ext = '.pdf'; }
        else if (filenameLower.endsWith('.png')){ mimeType = 'image/png'; ext = '.png'; }
        else if (filenameLower.endsWith('.jpg') || filenameLower.endsWith('.jpeg')){ mimeType = 'image/jpeg'; ext = filenameLower.endsWith('.jpeg') ? '.jpeg' : '.jpg'; }
        const endDateStr = Utilities.formatDate(new Date(data.endDate), Session.getScriptTimeZone(), 'yyyy-MM-dd');
        const safeName = endDateStr.replace(/[\\/:*?"<>|]/g, '_');
        const newFileName = `${safeName}_${idx + 1}${ext}`;
        const blob = Utilities.newBlob(bytes, mimeType, newFileName);
        const driveFile = folder.createFile(blob);
        driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        const fileTypeText = filenameLower.endsWith('.pdf') ? 'PDF檔' : '圖檔';
        imageLinks.push(`${fileTypeText}${idx + 1}: ${driveFile.getUrl()}`);
      } 
      catch (e) 
      {
        Logger.log(`檔案${idx + 1} (${file.name || '未知'}) 上傳失敗: ${e}`);
      }
    });
  }

  try {
    const evtStart = new Date(data.startDate + 'T00:00:00');
    const evtEnd   = new Date(data.endDate   + 'T00:00:00');
    evtEnd.setDate(evtEnd.getDate() + 1);

    const dateRange = data.startDate === data.endDate
      ? data.startDate
      : `${data.startDate} → ${data.endDate}`;

    const title = `${typeInfo.prefix}-${data.activityName}-${data.venue}`;
    let description =
`郵件：${userEmail}
姓名：${data.name}
電話：${data.phone}
期間：${dateRange}
類型：${typeInfo.label}
活動：${data.activityName}
場地：${data.venue}
備註：${data.note || ''}`;

    if (imageLinks && imageLinks.length) {
      description += `\n\n圖片連結：\n${imageLinks.join('\n')}`;
    }

    const event1 = cal.createAllDayEvent(title, evtStart, evtEnd);
    event1.setDescription(description);
    event1.setColor(typeInfo.color || CalendarApp.EventColor.LAVENDER);

    if (data.type === 'self' && parseInt(data.line) === 2) {
      const event2 = cal.createAllDayEvent(title, evtStart, evtEnd);
      event2.setDescription(description);
      event2.setColor(typeInfo.color || CalendarApp.EventColor.LAVENDER);
    }

    Logger.log(`新增成功：${title} (${data.startDate} ~ ${data.endDate})`);
  } catch (e) {
    Logger.log(e.stack);
    return "建立事件失敗：" + e.message;
  }

  var adminEmail = "itd@gmail.com.tw";
  var subject = "線路預約通知 - " + data.activityName + " " + data.startDate +
    (data.startDate !== data.endDate ? " ~ " + data.endDate : "");

  var body = `
線路預約成功！

業務負責人：${data.name}@gmail.com.tw
電話：${data.phone}

活動名稱：${data.activityName}
場地：${data.venue}

開始日期：${data.startDate}
結束日期：${data.endDate}

類型：${typeInfo.label}

備註：${data.note || '無'}

圖檔數量：${imageLinks.length} 張
${imageLinks.length > 0 ? '圖檔連結：\n' + imageLinks.join('\n') : ''}

本郵件由臨時線路預約系統自動通知 `;

  var guest_body = `
線路預約成功！

預約者：${userEmail}
業務負責人：${data.name}@gmail.com.tw
電話：${data.phone}

活動名稱：${data.activityName}
場地：${data.venue}

開始日期：${data.startDate}
結束日期：${data.endDate}

類型：${typeInfo.label}

備註：${data.note || '無'}

圖檔數量：${imageLinks.length} 張
${imageLinks.length > 0 ? '圖檔連結：\n' + imageLinks.join('\n') : ''}

本郵件由臨時線路預約系統自動通知`;

  const guestEmail = data.name + "@gmail.com.tw";
  if (userEmail.split('@')[0].toLowerCase() !== data.name.toLowerCase()) {
    GmailApp.sendEmail(userEmail, subject, guest_body, { cc: [guestEmail, adminEmail].join(",") });
  } else {
    GmailApp.sendEmail(userEmail, subject, body, { cc: adminEmail });
  }

  return `預約成功！${data.startDate} ~ ${data.endDate} 已安排\n活動：${data.activityName}\n類型：${typeInfo.label}\n場地：${data.venue}`;
}

function getMyBookings(startStr, endStr) {
  const userEmail = Session.getActiveUser().getEmail().toLowerCase().trim();
  const isAdmin = checkIsAdmin(userEmail);
  const cal = CalendarApp.getCalendarById(CALENDAR_ID);

  const start = new Date(startStr + 'T00:00:00');
  const end   = new Date(endStr   + 'T23:59:59');
  const events = cal.getEvents(start, end);

  return events
    .map(e => ({
      id:      e.getId(),
      title:   e.getTitle(),
      start:   Utilities.formatDate(e.getAllDayStartDate(), "Asia/Taipei", "yyyy-MM-dd"),
      end:     Utilities.formatDate(e.getAllDayEndDate(),   "Asia/Taipei", "yyyy-MM-dd"),
      creator: (e.getDescription() || '').match(/郵件：(.+)/)?.[1]?.trim() || ''
    }))
    .filter(evt => isAdmin || evt.creator.toLowerCase() === userEmail);
}

function deleteBooking(eventId) {
  const userEmail = Session.getActiveUser().getEmail().toLowerCase().trim();
  const isAdmin = checkIsAdmin(userEmail);
  const cal = CalendarApp.getCalendarById(CALENDAR_ID);

  let event;
  try {
    event = cal.getEventById(eventId);
  } catch(e) {
    return "找不到此預約，可能已被刪除";
  }
  if (!event) return "找不到此預約，可能已被刪除";

  const desc = event.getDescription() || '';
  const creator = (desc.match(/郵件：(.+)/)?.[1] || '').toLowerCase().trim();

  if (!isAdmin && creator !== userEmail) {
    return "權限不足：只能刪除自己的預約";
  }

  const title     = event.getTitle();
  const startDate = Utilities.formatDate(event.getAllDayStartDate(), "Asia/Taipei", "yyyy-MM-dd");
  const endDate   = Utilities.formatDate(event.getAllDayEndDate(),   "Asia/Taipei", "yyyy-MM-dd");

  const nameMatch  = desc.match(/姓名：(.+)/);
  const phoneMatch = desc.match(/電話：(.+)/);
  const typeMatch  = desc.match(/類型：(.+)/);
  const actMatch   = desc.match(/活動：(.+)/);
  const venueMatch = desc.match(/場地：(.+)/);
  const noteMatch  = desc.match(/備註：(.+)/);

  const name  = nameMatch  ? nameMatch[1].trim()  : '';
  const phone = phoneMatch ? phoneMatch[1].trim()  : '';
  const type  = typeMatch  ? typeMatch[1].trim()   : '';
  const act   = actMatch   ? actMatch[1].trim()    : '';
  const venue = venueMatch ? venueMatch[1].trim()  : '';
  const note  = noteMatch  ? noteMatch[1].trim()   : '';

  event.deleteEvent();

  const adminEmail = "itd@gmail.com.tw";
  const dateRange  = startDate === endDate ? startDate : `${startDate} → ${endDate}`;
  const subject    = "線路預約【已刪除】- " + act + " " + startDate +
    (startDate !== endDate ? " ~ " + endDate : "");

  const body = `
線路預約已刪除！

業務負責人：${name}@gmail.com.tw
電話：${phone}

活動名稱：${act}
場地：${venue}

開始日期：${startDate}
結束日期：${endDate}

類型：${type}

備註：${note || '無'}

操作者：${userEmail}
本郵件由臨時線路預約系統自動通知`;

  const guest_body = `
線路預約已刪除！

預約者：${creator}
業務負責人：${name}@gmail.com.tw
電話：${phone}

活動名稱：${act}
場地：${venue}

開始日期：${startDate}
結束日期：${endDate}

類型：${type}

備註：${note || '無'}

操作者：${userEmail}
本郵件由臨時線路預約系統自動通知`;

  const guestEmail = name + "@gmail.com.tw";
  if (userEmail.split('@')[0].toLowerCase() !== name.toLowerCase()) {
    GmailApp.sendEmail(userEmail, subject, guest_body, { cc: [guestEmail, adminEmail].join(",") });
  } else {
    GmailApp.sendEmail(userEmail, subject, body, { cc: adminEmail });
  }

  return `已成功刪除：${act}（${dateRange}）`;
}
