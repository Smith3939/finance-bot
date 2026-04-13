require('dotenv').config();
const TelegramBot = require('node-telegram-bot-api');
const { GoogleGenerativeAI } = require('@google/generative-ai');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

const bot = new TelegramBot(process.env.TELEGRAM_TOKEN, { polling: true });
const genAI = new GoogleGenerativeAI(process.env.GEMINI_API_KEY);
const model = genAI.getGenerativeModel({ model: 'gemini-2.5-flash' });

const EXCEL_FILE = path.join(__dirname, 'expenses.xlsx');
const USERS_FILE = path.join(__dirname, 'users.json');

const CATEGORIES = {
  'מזון וסופר': ['סופר','רמי לוי','שופרסל','מגה','יוחננוף','אושר עד','ויקטורי','חצי חינם','פירות','ירקות','בשר','לחם','חלב','מכולת','קניות','מזון','אוכל','גבינה','ביצים'],
  'מסעדות ואוכל בחוץ': ['מסעדה','פיצה','המבורגר','שווארמה','פלאפל','קפה','בית קפה','סושי','משלוח','וולט','תן ביס','חומוס','מאפייה','סנדוויץ'],
  'דיור': ['שכר דירה','שכירות','ארנונה','ועד בית','משכנתא'],
  'חשבונות בית': ['חשמל','מים','גז','אינטרנט','טלפון','סלולר','כבלים','טלוויזיה','נטפליקס','ספוטיפיי'],
  'תחבורה': ['דלק','בנזין','רכבת','אוטובוס','מונית','גט','חניה','כביש אגרה','כביש 6','טסט','ביטוח רכב','טיפול רכב'],
  'בריאות': ['רופא','מרפאה','תרופות','בית מרקחת','סופר פארם','רופא שיניים','אופטיקה','משקפיים','ביטוח בריאות','קופת חולים'],
  'חינוך': ['גן','בית ספר','חוגים','שיעור','קורס','לימודים','ספרים','מחברות','ציוד לימודי'],
  'ביגוד והנעלה': ['בגדים','נעליים','חולצה','מכנסיים','שמלה','זארה','H&M','קסטרו','פוקס','גולף'],
  'בידור ופנאי': ['סרט','קולנוע','הצגה','קונצרט','בריכה','חדר כושר','ספורט','משחק','טיול','חופשה','מלון','צימר','סאונה'],
  'ביטוחים': ['ביטוח','ביטוח חיים','ביטוח דירה','ביטוח רכב','ביטוח בריאות'],
  'מוצרי בית': ['ריהוט','מכשיר חשמלי','ניקיון','כביסה','איקאה','כלי מטבח','מצעים'],
  'טיפוח': ['מספרה','תספורת','קוסמטיקה','איפור','קרם','שמפו'],
  'ילדים': ['צעצועים','חיתולים','מזון תינוקות','עגלה','בייביסיטר','צהרון'],
  'חיסכון והשקעות': ['חיסכון','השקעה','קרן','פנסיה','קופת גמל'],
  'הכנסה': ['משכורת','שכר','בונוס','מענק','החזר','העברה','הכנסה'],
  'אחר': []
};

function loadUsers() {
  if (fs.existsSync(USERS_FILE)) return JSON.parse(fs.readFileSync(USERS_FILE, 'utf8'));
  return {};
}
function saveUsers(users) {
  fs.writeFileSync(USERS_FILE, JSON.stringify(users, null, 2), 'utf8');
}

async function parseExpenseWithAI(text) {
  const prompt = `אתה מערכת לניתוח הוצאות והכנסות ביתיות. נתח את ההודעה הבאה וחלץ ממנה עסקאות.

הודעה: "${text}"

רשימת הקטגוריות האפשריות:
${Object.keys(CATEGORIES).join(', ')}

החזר תשובה בפורמט JSON בלבד (בלי markdown, בלי backticks), כמערך של אובייקטים:
[{"description": "תיאור קצר", "amount": מספר, "category": "קטגוריה", "type": "expense או income"}]

כללים:
- מספר במילים ("ארבעה אלף מאה") = המר למספר (4100)
- כמה הוצאות במשפט = פצל לשורות
- אם לא ברור הסכום = החזר []
- החזר JSON תקין בלבד`;

  try {
    const result = await model.generateContent(prompt);
    let cleanJson = result.response.text().trim().replace(/```json\s*/g, '').replace(/```\s*/g, '').trim();
    const parsed = JSON.parse(cleanJson);
    if (!Array.isArray(parsed)) return [];
    return parsed.filter(item => item.description && typeof item.amount === 'number' && item.amount > 0 && item.category && Object.keys(CATEGORIES).includes(item.category)).map(item => ({ ...item, type: item.type || 'expense' }));
  } catch (error) {
    console.error('Gemini error:', error.message);
    return null;
  }
}

async function initExcel() {
  if (fs.existsSync(EXCEL_FILE)) return;
  const workbook = new ExcelJS.Workbook();
  const txSheet = workbook.addWorksheet('עסקאות');
  txSheet.columns = [
    { header: 'תאריך', key: 'date', width: 14 },
    { header: 'תיאור', key: 'description', width: 30 },
    { header: 'סכום', key: 'amount', width: 12 },
    { header: 'קטגוריה', key: 'category', width: 18 },
    { header: 'סוג', key: 'type', width: 10 },
    { header: 'מקור', key: 'source', width: 18 },
    { header: 'משתמש', key: 'user', width: 20 },
    { header: 'כרטיס', key: 'card', width: 18 }
  ];
  txSheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 12 };
  txSheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } };
  const dashSheet = workbook.addWorksheet('דשבורד');
  dashSheet.columns = [
    { header: 'קטגוריה', key: 'category', width: 20 },
    { header: 'סה"כ החודש', key: 'monthly_total', width: 15 },
    { header: 'סה"כ השבוע', key: 'weekly_total', width: 15 },
    { header: 'מספר עסקאות', key: 'count', width: 14 }
  ];
  dashSheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 12 };
  dashSheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF70AD47' } };
  const catSheet = workbook.addWorksheet('קטגוריות');
  catSheet.columns = [
    { header: 'קטגוריה', key: 'category', width: 20 },
    { header: 'מילות מפתח', key: 'keywords', width: 60 }
  ];
  catSheet.getRow(1).font = { bold: true, size: 12 };
  for (const [cat, keywords] of Object.entries(CATEGORIES)) {
    catSheet.addRow({ category: cat, keywords: keywords.join(', ') });
  }
  await workbook.xlsx.writeFile(EXCEL_FILE);
  console.log('Excel file created');
}

async function addTransaction(transaction) {
  await initExcel();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(EXCEL_FILE);
  const sheet = workbook.getWorksheet('עסקאות');
  const now = new Date();
  const dateStr = `${now.getDate().toString().padStart(2,'0')}/${(now.getMonth()+1).toString().padStart(2,'0')}/${now.getFullYear()}`;
  sheet.addRow({ date: dateStr, description: transaction.description, amount: transaction.amount, category: transaction.category, type: transaction.type === 'income' ? 'הכנסה' : 'הוצאה', source: transaction.source || 'טלגרם', user: transaction.user || '', card: transaction.card || '' });
  await workbook.xlsx.writeFile(EXCEL_FILE);
}

async function getSummary(period) {
  if (!fs.existsSync(EXCEL_FILE)) return null;
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(EXCEL_FILE);
  const sheet = workbook.getWorksheet('עסקאות');
  const now = new Date();
  const summary = {};
  let totalExpenses = 0, totalIncome = 0;
  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    const dateStr = row.getCell(1).value;
    const amount = parseFloat(row.getCell(3).value) || 0;
    const category = row.getCell(4).value;
    const type = row.getCell(5).value;
    if (!dateStr) return;
    const parts = dateStr.toString().split('/');
    if (parts.length !== 3) return;
    const rowDate = new Date(parts[2], parts[1]-1, parts[0]);
    let include = false;
    if (period === 'weekly') { const weekAgo = new Date(now); weekAgo.setDate(weekAgo.getDate()-7); include = rowDate >= weekAgo; }
    else if (period === 'monthly') { include = rowDate.getMonth() === now.getMonth() && rowDate.getFullYear() === now.getFullYear(); }
    if (include) {
      if (type === 'הכנסה') { totalIncome += amount; }
      else { totalExpenses += amount; if (!summary[category]) summary[category] = 0; summary[category] += amount; }
    }
  });
  return { summary, totalExpenses, totalIncome };
}

bot.onText(/\/start/, async (msg) => {
  const chatId = msg.chat.id;
  const users = loadUsers();
  const userName = msg.from.first_name + (msg.from.last_name ? ' ' + msg.from.last_name : '');
  users[chatId] = { name: userName, chatId };
  saveUsers(users);
  await bot.sendMessage(chatId, `שלום ${userName}! אני הבוט לניהול הוצאות משק הבית.\n\nפשוט כתבו בעברית רגילה, למשל:\n- "קניות ברמי לוי 850 שקל"\n- "שכר דירה פלוס חשמל ומים ארבעת אלפים מאה"\n- "דלק שלוש מאות חמישים"\n\nפקודות:\n/summary - סיכום שבועי\n/monthly - סיכום חודשי\n/categories - קטגוריות\n/help - עזרה`);
});

bot.onText(/\/help/, async (msg) => {
  await bot.sendMessage(msg.chat.id, 'שלחו הודעה עם הוצאה בעברית רגילה.\nמספרים בספרות: "דלק 350"\nמספרים במילים: "סופר שלוש מאות"\nכמה הוצאות: "שכירות 3100, חשמל 400"\n\n/summary - סיכום שבועי\n/monthly - סיכום חודשי\n/categories - קטגוריות');
});

bot.onText(/\/categories/, async (msg) => {
  await bot.sendMessage(msg.chat.id, 'קטגוריות:\n' + Object.keys(CATEGORIES).map(c => '- ' + c).join('\n'));
});

bot.onText(/\/summary/, async (msg) => {
  const data = await getSummary('weekly');
  if (!data || Object.keys(data.summary).length === 0) return bot.sendMessage(msg.chat.id, 'אין עסקאות בשבוע האחרון.');
  let text = 'סיכום שבועי:\n\n';
  for (const [cat, total] of Object.entries(data.summary).sort((a,b) => b[1]-a[1])) text += `${cat}: ${total.toLocaleString()} ש"ח\n`;
  text += `\nסה"כ הוצאות: ${data.totalExpenses.toLocaleString()} ש"ח`;
  if (data.totalIncome > 0) { text += `\nסה"כ הכנסות: ${data.totalIncome.toLocaleString()} ש"ח`; text += `\nמאזן: ${(data.totalIncome - data.totalExpenses).toLocaleString()} ש"ח`; }
  await bot.sendMessage(msg.chat.id, text);
});

bot.onText(/\/monthly/, async (msg) => {
  const data = await getSummary('monthly');
  if (!data || Object.keys(data.summary).length === 0) return bot.sendMessage(msg.chat.id, 'אין עסקאות החודש.');
  const monthNames = ['ינואר','פברואר','מרץ','אפריל','מאי','יוני','יולי','אוגוסט','ספטמבר','אוקטובר','נובמבר','דצמבר'];
  let text = `סיכום חודשי - ${monthNames[new Date().getMonth()]}:\n\n`;
  for (const [cat, total] of Object.entries(data.summary).sort((a,b) => b[1]-a[1])) text += `${cat}: ${total.toLocaleString()} ש"ח\n`;
  text += `\nסה"כ הוצאות: ${data.totalExpenses.toLocaleString()} ש"ח`;
  if (data.totalIncome > 0) { text += `\nסה"כ הכנסות: ${data.totalIncome.toLocaleString()} ש"ח`; text += `\nמאזן: ${(data.totalIncome - data.totalExpenses).toLocaleString()} ש"ח`; }
  await bot.sendMessage(msg.chat.id, text);
});

bot.on('message', async (msg) => {
  const chatId = msg.chat.id;
  const text = msg.text;
  if (!text || text.startsWith('/')) return;
  const users = loadUsers();
  const userName = users[chatId]?.name || msg.from.first_name || 'לא ידוע';
  if (!users[chatId]) { users[chatId] = { name: userName, chatId }; saveUsers(users); }
  const processingMsg = await bot.sendMessage(chatId, 'מעבד...');
  try {
    const transactions = await parseExpenseWithAI(text);
    if (!transactions || transactions.length === 0) {
      await bot.editMessageText('לא הצלחתי לזהות הוצאה או הכנסה.\nנסה: "קניות ברמי לוי 500" או "שכר דירה 3100"', { chat_id: chatId, message_id: processingMsg.message_id });
      return;
    }
    let confirmText = 'נרשם בהצלחה!\n\n';
    for (const tx of transactions) {
      await addTransaction({ ...tx, user: userName, source: 'טלגרם' });
      confirmText += `${tx.type === 'income' ? 'הכנסה' : 'הוצאה'}: ${tx.description}\n${tx.amount.toLocaleString()} ש"ח | ${tx.category} | ${userName}\n\n`;
    }
    await bot.editMessageText(confirmText, { chat_id: chatId, message_id: processingMsg.message_id });
  } catch (error) {
    console.error('Error:', error);
    await bot.editMessageText('שגיאה בעיבוד ההודעה. נסה שוב.', { chat_id: chatId, message_id: processingMsg.message_id });
  }
});

async function start() {
  await initExcel();
  console.log('הבוט עלה בהצלחה! SMITH משק בית פעיל.');
}
start();