/**
 * Google Apps Script — מחשבון אומדן מענק ימי מחלה בפרישה
 * אלפי ביטוח ופיננסים
 *
 * שומר את הפנייה ב-Google Sheet ושולח מייל אוטומטי לאלפי.
 *
 * הוראות התקנה — ראה README.md
 */

// ═══════════════════════════════════════════════════════════════════
// הגדרות — שנה אותן בהתאם
// ═══════════════════════════════════════════════════════════════════

const RECIPIENT_EMAIL = 'office@alafi-ins.co.il';   // המייל שאליו יישלחו ההתראות
const SHEET_NAME      = 'הגשות';                    // שם הגיליון לשמירת הנתונים
const COMPANY_NAME    = 'אלפי ביטוח ופיננסים';

// ═══════════════════════════════════════════════════════════════════
// קוד ראשי — אין צורך לערוך מתחת לשורה הזו
// ═══════════════════════════════════════════════════════════════════

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    saveToSheet(data);
    sendNotificationEmail(data);

    return jsonResponse({ success: true });
  } catch (err) {
    console.error('Error in doPost:', err);
    // עדיין נחזיר תשובה לדפדפן כדי שלא יציג שגיאה במצב CORS
    return jsonResponse({ success: false, error: err.message });
  }
}

function doGet(e) {
  // נקודת קצה לבדיקה שהסקריפט פועל
  return jsonResponse({
    status: 'ok',
    service: 'Lemoney Sick Days Calculator Webhook',
    company: COMPANY_NAME,
  });
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * שמירה ב-Google Sheet
 * יוצר אוטומטית את הגיליון ושורת הכותרות בפעם הראשונה
 */
function saveToSheet(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  // יצירת גיליון אם לא קיים
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  const headers = [
    'תאריך הגשה',
    'שם מלא', 'ת.ז.', 'טלפון', 'תאריך לידה', 'משרד ממשלתי',
    'תחילת העסקה', 'תאריך פרישה',
    'יתרת ימי מחלה', 'שכר חודשי ברוטו', 'שכר יומי',
    'שנים', 'חודשים', 'שנים עשרוניות',
    'סה״כ ימי מחלה שנצברו', 'ימי מחלה שנוצלו', 'אחוז ניצול',
    'מדרגת זכאות', 'זכאי?', 'סכום מענק משוער (₪)',
    'אישור תקנון', 'אישור מסירת פרטים'
  ];

  // הוספת כותרות אם הגיליון ריק
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#1e5288');
    headerRange.setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }

  const row = [
    new Date(),
    data.fullName || '',
    data.idNumber || '',
    data.phone || '',
    data.birthDate || '',
    data.ministry || '',
    data.startDate || '',
    data.retireDate || '',
    data.sickBalance || 0,
    data.monthlySalary || 0,
    data.dailyWage || 0,
    data.years || 0,
    data.months || 0,
    data.decimalYears || 0,
    data.totalAccrued || 0,
    data.used || 0,
    data.usagePct || 0,
    data.tier || '',
    data.eligible ? 'כן' : 'לא',
    data.grant || 0,
    data.termsApproved ? 'כן' : 'לא',
    data.consentApproved ? 'כן' : 'לא',
  ];

  sheet.appendRow(row);

  // התאמת רוחב עמודות אוטומטית מדי פעם
  if (sheet.getLastRow() % 10 === 1) {
    sheet.autoResizeColumns(1, headers.length);
  }
}

/**
 * שליחת מייל התראה לאלפי
 */
function sendNotificationEmail(data) {
  const subject = `🔔 פנייה חדשה במחשבון מענק ימי מחלה — ${data.fullName}`;

  const fmtMoney = n => Number(n || 0).toLocaleString('he-IL') + ' ₪';
  const fmtPct   = n => Number(n || 0).toFixed(1) + '%';

  const eligibleBadge = data.eligible
    ? '<span style="background:#dcf3e6;color:#2d7a5a;padding:4px 12px;border-radius:100px;font-weight:700;font-size:13px;">זכאי</span>'
    : '<span style="background:#fce6df;color:#c75a3a;padding:4px 12px;border-radius:100px;font-weight:700;font-size:13px;">לא זכאי</span>';

  const htmlBody = `
<div style="font-family: Arial, sans-serif; direction: rtl; max-width: 640px; margin: 0 auto; color: #0f2742;">

  <div style="background: linear-gradient(135deg, #154169 0%, #1e5288 50%, #2d6ba8 100%); padding: 28px 24px; border-radius: 6px 6px 0 0; text-align: right;">
    <div style="color:#7fb8de; font-size:11px; letter-spacing:0.18em; text-transform:uppercase; font-weight:700;">פנייה חדשה</div>
    <h1 style="color:#fff; margin:8px 0 0 0; font-size:22px; font-weight:500;">מחשבון אומדן מענק ימי מחלה</h1>
    <div style="color:rgba(255,255,255,0.7); font-size:13px; margin-top:4px;">${COMPANY_NAME}</div>
  </div>

  <div style="background:#fff; padding: 24px; border: 1px solid #c5d6e8; border-top:none;">

    <h2 style="color:#154169; font-size:16px; margin:0 0 12px 0; padding-bottom:6px; border-bottom:2px solid #4a9bd1;">פרטי העובד</h2>
    <table style="width:100%; border-collapse:collapse; font-size:14px;">
      <tr><td style="padding:6px 0; color:#4a5a68; width:40%;">שם מלא</td><td style="font-weight:700; padding:6px 0;">${escapeHtml(data.fullName)}</td></tr>
      <tr><td style="padding:6px 0; color:#4a5a68;">תעודת זהות</td><td style="font-weight:700; padding:6px 0;">${escapeHtml(data.idNumber)}</td></tr>
      <tr><td style="padding:6px 0; color:#4a5a68;">טלפון</td><td style="font-weight:700; padding:6px 0;"><a href="tel:${escapeHtml(data.phone)}" style="color:#1e5288; text-decoration:none;">${escapeHtml(data.phone)}</a></td></tr>
      <tr><td style="padding:6px 0; color:#4a5a68;">תאריך לידה</td><td style="font-weight:700; padding:6px 0;">${escapeHtml(data.birthDate)}</td></tr>
      <tr><td style="padding:6px 0; color:#4a5a68;">משרד ממשלתי</td><td style="font-weight:700; padding:6px 0;">${escapeHtml(data.ministry)}</td></tr>
    </table>

    <h2 style="color:#154169; font-size:16px; margin:24px 0 12px 0; padding-bottom:6px; border-bottom:2px solid #4a9bd1;">פרטי העסקה</h2>
    <table style="width:100%; border-collapse:collapse; font-size:14px;">
      <tr><td style="padding:6px 0; color:#4a5a68; width:40%;">תאריך תחילת העסקה</td><td style="font-weight:700; padding:6px 0;">${escapeHtml(data.startDate)}</td></tr>
      <tr><td style="padding:6px 0; color:#4a5a68;">תאריך פרישה</td><td style="font-weight:700; padding:6px 0;">${escapeHtml(data.retireDate)}</td></tr>
      <tr><td style="padding:6px 0; color:#4a5a68;">תקופת עבודה</td><td style="font-weight:700; padding:6px 0;">${data.years} שנים, ${data.months} חודשים (${data.decimalYears} שנים עשרוניות)</td></tr>
      <tr><td style="padding:6px 0; color:#4a5a68;">שכר חודשי ברוטו</td><td style="font-weight:700; padding:6px 0;">${data.monthlySalary > 0 ? fmtMoney(data.monthlySalary) : '—'}</td></tr>
      <tr><td style="padding:6px 0; color:#4a5a68;">שכר יומי</td><td style="font-weight:700; padding:6px 0;">${data.dailyWage > 0 ? fmtMoney(data.dailyWage) : 'לא הוזן'}</td></tr>
    </table>

    <h2 style="color:#154169; font-size:16px; margin:24px 0 12px 0; padding-bottom:6px; border-bottom:2px solid #4a9bd1;">חישוב ימי מחלה</h2>
    <table style="width:100%; border-collapse:collapse; font-size:14px;">
      <tr><td style="padding:6px 0; color:#4a5a68; width:40%;">סה״כ ימי מחלה שנצברו</td><td style="font-weight:700; padding:6px 0;">${data.totalAccrued} ימים</td></tr>
      <tr><td style="padding:6px 0; color:#4a5a68;">יתרת ימי מחלה</td><td style="font-weight:700; padding:6px 0;">${data.sickBalance} ימים</td></tr>
      <tr><td style="padding:6px 0; color:#4a5a68;">ימי מחלה שנוצלו</td><td style="font-weight:700; padding:6px 0;">${data.used} ימים</td></tr>
      <tr><td style="padding:6px 0; color:#4a5a68;">אחוז ניצול</td><td style="font-weight:700; padding:6px 0;">${fmtPct(data.usagePct)}</td></tr>
      <tr><td style="padding:6px 0; color:#4a5a68;">מדרגת זכאות</td><td style="font-weight:700; padding:6px 0;">${escapeHtml(data.tier)} ${eligibleBadge}</td></tr>
    </table>

    <div style="background: linear-gradient(135deg, #154169 0%, #1e5288 50%, #2d6ba8 100%); margin-top:24px; padding:24px; border-radius:4px; display:flex; justify-content:space-between; align-items:center;">
      <div style="color:rgba(255,255,255,0.85); font-size:14px;">סכום המענק המשוער</div>
      <div style="color:#fff; font-size:28px; font-weight:700;">${data.eligible ? fmtMoney(data.grant) : 'אין זכאות'}</div>
    </div>

    <div style="background:#f0f5fa; border-right:3px solid #1e5288; margin-top:20px; padding:14px 18px; font-size:12px; color:#4a5a68; line-height:1.6;">
      <strong style="color:#0f2742;">הסתייגות:</strong> החישוב המוצג הוא אומדן בלבד ואינו מהווה ייעוץ מס או אישור זכאות רשמי. הלקוח אישר את התקנון ומסר את הפרטים מרצונו דרך מחשבון Lemoney.
    </div>

    <div style="margin-top:20px; padding-top:16px; border-top:1px solid #c5d6e8; font-size:12px; color:#7a8da3; text-align:center;">
      <strong style="color:#154169;">${COMPANY_NAME}</strong> · מומחי פרישה וטופס 161
    </div>

  </div>
</div>
`;

  const plainBody = `
פנייה חדשה במחשבון מענק ימי מחלה
═══════════════════════════════

פרטי העובד:
שם: ${data.fullName}
ת.ז.: ${data.idNumber}
טלפון: ${data.phone}
תאריך לידה: ${data.birthDate}
משרד: ${data.ministry}

פרטי העסקה:
תחילת העסקה: ${data.startDate}
תאריך פרישה: ${data.retireDate}
תקופה: ${data.years} שנים, ${data.months} חודשים

שכר:
חודשי: ${fmtMoney(data.monthlySalary)}
יומי: ${data.dailyWage > 0 ? fmtMoney(data.dailyWage) : 'לא הוזן'}

ימי מחלה:
סה״כ נצברו: ${data.totalAccrued}
יתרה: ${data.sickBalance}
נוצלו: ${data.used}
אחוז ניצול: ${fmtPct(data.usagePct)}

תוצאה:
${data.tier}
סכום מענק: ${data.eligible ? fmtMoney(data.grant) : 'אין זכאות'}

═══════════════════════════════
הלקוח אישר את התקנון ומסר את הפרטים מרצון.
${COMPANY_NAME} · מחשבון Lemoney
`;

  MailApp.sendEmail({
    to: RECIPIENT_EMAIL,
    subject: subject,
    body: plainBody,
    htmlBody: htmlBody,
    name: COMPANY_NAME,
  });
}

/**
 * Escape HTML to prevent injection in the email body
 */
function escapeHtml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/**
 * פונקציית בדיקה — הרץ ידנית פעם אחת מתוך עורך הסקריפט
 * תיצור הגשת דוגמה כדי לוודא שהכל עובד
 */
function testSubmission() {
  const sampleData = {
    submittedAt: new Date().toISOString(),
    fullName: 'ישראל ישראלי (בדיקה)',
    idNumber: '123456789',
    phone: '050-1234567',
    birthDate: '1965-05-15',
    ministry: 'משרד החינוך',
    startDate: '1990-09-01',
    retireDate: '2026-08-31',
    sickBalance: 200,
    monthlySalary: 18000,
    dailyWage: 0,
    years: 35,
    months: 11,
    decimalYears: 35.99,
    totalAccrued: 1079.7,
    used: 879.7,
    usagePct: 81.5,
    tier: 'מעל 65% — אין זכאות למענק',
    eligible: false,
    grant: 0,
    consentApproved: true,
    termsApproved: true,
  };

  saveToSheet(sampleData);
  sendNotificationEmail(sampleData);
  Logger.log('בדיקה הסתיימה. בדוק את הגיליון ואת תיבת המייל.');
}
