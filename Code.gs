const SS = SpreadsheetApp.getActiveSpreadsheet();
const TZ = 'Africa/Cairo';
const ROOT_NAME = 'حساب العمارة';
const MONTHS = ['01 - يناير', '02 - فبراير', '03 - مارس', '04 - أبريل', '05 - مايو', '06 - يونيو',
  '07 - يوليو', '08 - أغسطس', '09 - سبتمبر', '10 - أكتوبر', '11 - نوفمبر', '12 - ديسمبر'];
let CTX = { did: '', act: 'view' };

// ---------- الأسماء العربية (الكود يقبل العربي والإنجليزي) ----------
const SK = {
  'اسم أمين الصندوق': 'treasurer_name', 'موبايل أمين الصندوق': 'treasurer_phone', 'بيانات الدفع': 'pay_info',
  'نوع وسيلة الدفع 1': 'pay1_type', 'رقم وسيلة الدفع 1': 'pay1_number', 'نوع وسيلة الدفع 2': 'pay2_type', 'رقم وسيلة الدفع 2': 'pay2_number',
  'إعلان أعلى الصفحة': 'announcement', 'إيميل التقارير': 'report_email', 'إظهار رصيد الصندوق': 'show_fund',
  'إظهار تفاصيل المصروفات': 'show_expenses', 'رابط مجلد الإيصالات': 'receipts_folder_url', 'رابط مجلد الفواتير': 'invoices_folder_url',
  'اسم مسؤول الصيانة': 'maint_name', 'رمز مسؤول الصيانة': 'maint_code',
  'رمز الإدارة': 'admin_code', 'إظهار جملة اليوم': 'show_quote', 'إظهار تحصيل المشروع': 'show_project'
};
const SK_REV = {};
Object.keys(SK).forEach(a => { SK_REV[SK[a]] = a; });
const ST = { 'قيد المراجعة': 'pending', 'pending': 'pending', 'تم التأكيد': 'confirmed', 'confirmed': 'confirmed', 'مرفوض': 'rejected', 'rejected': 'rejected' };
const ST_AR = { pending: 'قيد المراجعة', confirmed: 'تم التأكيد', rejected: 'مرفوض' };
const RES_AR = { ok: 'نجح', bad_pin: 'رقم سري خاطئ', bad_apt: 'شقة غير موجودة', locked: 'مقفول' };
const ACT_AR = { view: 'عرض', upload: 'رفع', m_view: 'عرض مهام', m_update: 'تحديث مهمة', a_view: 'عرض إدارة', a_write: 'تعديل إدارة' };

function st(v) { return ST[String(v || '').trim()] || ''; }
function skey(label) { const l = String(label || '').trim(); return SK[l] || l; }
function yes(v) { return ['yes', 'نعم'].indexOf(String(v || '').trim().toLowerCase()) >= 0; }
// رقم دولي لواتساب: يعالج ضياع الصفر ويضيف 2 للأرقام المصرية
function intl(v) {
  let d = nd(v).replace(/\D/g, '');
  if (d.length === 10 && d[0] === '1') d = '0' + d;
  if (d.length === 11 && d[0] === '0') d = '2' + d;
  return d;
}
// رقم سري: لو ضاع الصفر من أوله (123 بدل 0123) نكمله لـ 4 أرقام
// رقم وسيلة دفع: لو أرقام فقط نشيل المسافات والشرطات، ولو موبايل ضاع صفره (10 أرقام تبدأ بـ 1) نرجّعه. غير ذلك (مثل عنوان إنستا باي) يظهر كما هو
function payNum(v) {
  let x = nd(v);
  if (/^[\d\s-]+$/.test(x)) x = x.replace(/\D/g, '');
  if (/^1\d{9}$/.test(x)) x = '0' + x;
  return x;
}
function pinNorm(v) {
  const x = nd(v);
  return /^\d{1,3}$/.test(x) ? ('0000' + x).slice(-4) : x;
}

// ---------- أدوات عامة ----------
function toDate(v) {
  if (v instanceof Date) return v;
  const m = String(v || '').match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  return m ? new Date(+m[1], +m[2] - 1, +m[3]) : null;
}
function fmt(v) {
  const d = toDate(v);
  return d ? Utilities.formatDate(d, TZ, 'yyyy-MM-dd') : String(v || '');
}
function settings() {
  const o = {};
  const sh = SS.getSheetByName('Settings');
  if (sh) sh.getDataRange().getValues().forEach(r => {
    if (r[0]) o[skey(r[0])] = String(r[1] == null ? '' : r[1]).trim();
  });
  return o;
}
function setSetting(k, v) {
  const sh = SS.getSheetByName('Settings');
  if (!sh) return;
  const i = sh.getDataRange().getValues().findIndex(r => skey(r[0]) === k);
  if (i >= 0) sh.getRange(i + 1, 2).setValue(v); else sh.appendRow([SK_REV[k] || k, v]);
}
function adminEmail() {
  return settings().report_email || Session.getEffectiveUser().getEmail();
}
// يحوّل الأرقام العربية (٠-٩) والفارسية للإنجليزية، فلا يفرق الكيبورد
function nd(s) {
  return String(s == null ? '' : s)
    .replace(/[٠-٩]/g, c => c.charCodeAt(0) - 1632)
    .replace(/[۰-۹]/g, c => c.charCodeAt(0) - 1776)
    .trim();
}
// مقارنة أكواد البنود بدون حساسية لحالة الحروف أو المسافات
function ck(v) { return String(v == null ? '' : v).trim().toLowerCase(); }
function clean(s) {
  return String(s).replace(/[\\/:*?"<>|]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 100);
}
function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// ---------- مجلدات Drive ----------
function getOrCreate(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}
function rootFolder() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('root_id');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const f = DriveApp.createFolder(ROOT_NAME);
  props.setProperty('root_id', f.getId());
  return f;
}
function yearFolder(kind, year) {
  return getOrCreate(getOrCreate(rootFolder(), kind), String(year));
}
// الإيصالات/السنة/12 شهر + الفواتير/السنة
function ensureYear(year) {
  const y = yearFolder('الإيصالات', year);
  MONTHS.forEach(m => getOrCreate(y, m));
  yearFolder('الفواتير', year);
  PropertiesService.getScriptProperties().setProperty('ens_' + year, '1');
}
function monthlyFolders() {
  const now = new Date();
  ensureYear(now.getFullYear());
  if (now.getMonth() === 11) ensureYear(now.getFullYear() + 1);
}
function setupFolders() {
  ensureYear(new Date().getFullYear());
  setSetting('receipts_folder_url', getOrCreate(rootFolder(), 'الإيصالات').getUrl());
  setSetting('invoices_folder_url', getOrCreate(rootFolder(), 'الفواتير').getUrl());
}

// ---------- سجل الدخول ----------
function logAccess(apt, result) {
  try {
    let sh = SS.getSheetByName('AccessLog');
    if (!sh) {
      sh = SS.insertSheet('AccessLog');
      sh.appendRow(['الوقت', 'الشقة', 'النتيجة (نجح / رقم سري خاطئ / شقة غير موجودة / مقفول)', 'معرّف الجهاز', 'العملية']);
    }
    sh.appendRow([new Date(), String(apt).slice(0, 10), RES_AR[result] || result, CTX.did, ACT_AR[CTX.act] || CTX.act]);
    if (result !== 'ok' && CTX.did) detectProbe();
  } catch (e) {}
}
// جهاز فشل 3 مرات (أو أكثر) في شقق غير شقته خلال 24 ساعة ⟵ تنبيه بالإيميل (مرة كل 24 ساعة لكل جهاز).
// "شقته" = الشقق التي دخلها هذا الجهاز بنجاح من قبل (تظهر في التنبيه للتعرف عليه).
function detectProbe() {
  if (!CTX.did || CTX.did === 'nostore') return;
  const sh = SS.getSheetByName('AccessLog');
  const n = sh.getLastRow();
  const from = Math.max(2, n - 500);
  const rows = sh.getRange(from, 1, n - from + 1, 5).getValues().filter(r => String(r[3]) === CTX.did && ['صيانة', 'إدارة'].indexOf(String(r[1])) < 0);
  const OK = RES_AR.ok;
  const homes = [];
  rows.forEach(r => { if (r[2] === OK && homes.indexOf(String(r[1])) < 0) homes.push(String(r[1])); });
  const since = Date.now() - 24 * 3600 * 1000;
  const fails = rows.filter(r => r[2] !== OK && r[0] instanceof Date && r[0].getTime() >= since && homes.indexOf(String(r[1])) < 0);
  if (fails.length < 3) return;
  const props = PropertiesService.getScriptProperties();
  const key = 'alert_' + CTX.did;
  if (Date.now() - Number(props.getProperty(key) || 0) < 24 * 3600 * 1000) return;
  props.setProperty(key, String(Date.now()));
  const tried = {};
  fails.forEach(r => { tried[String(r[1])] = (tried[String(r[1])] || 0) + 1; });
  MailApp.sendEmail(adminEmail(), 'تنبيه: جهاز يحاول فتح شقق غير شقته',
    'الجهاز (' + CTX.did + ') فشل ' + fails.length + ' مرات خلال 24 ساعة في شقق غير شقته.\n' +
    'الشقة/الشقق التي دخلها بنجاح من قبل: ' + (homes.length ? homes.join('، ') : 'لا يوجد (لم يدخل بنجاح قبل كده)') + '\n' +
    'الشقق التي حاول فتحها: ' + Object.keys(tried).map(k => k + ' (' + tried[k] + ')').join('، ') + '\n' +
    'راجع تاب AccessLog.');
}

// ---------- نقاط الدخول ----------
// ---------- جملة اليوم ----------
// 149 جملة: آيات وأحاديث (بلا سند) وأقوال السلف والعلماء والأمثال. [النص، المصدر، النوع v=آية h=حديث s=قول]
// ترتيبها مخلوط مسبقاً حتى لا يتكرر النوع، وتتغير الجملة يومياً بحسب رقم اليوم. لتعديل جملة عدّل هنا ثم New version.
const QUOTES = [["ما زال جبريل يوصيني بالجار حتى ظننت أنه سيورّثه","حديث شريف","h"],["وَاعْبُدُوا اللَّهَ وَلَا تُشْرِكُوا بِهِ شَيْئًا ۖ وَبِالْوَالِدَيْنِ إِحْسَانًا وَبِذِي الْقُرْبَىٰ وَالْيَتَامَىٰ وَالْمَسَاكِينِ وَالْجَارِ ذِي الْقُرْبَىٰ وَالْجَارِ الْجُنُبِ وَالصَّاحِبِ بِالْجَنبِ","سورة النساء، آية 36","v"],["من كان يؤمن بالله واليوم الآخر فليكرم جاره","حديث شريف","h"],["وَقُولُوا لِلنَّاسِ حُسْنًا","سورة البقرة، آية 83","v"],["ليس حسن الجوار كفّ الأذى، ولكن حسن الجوار الصبر على الأذى","يُروى عن الحسن البصري","s"],["من كان يؤمن بالله واليوم الآخر فليقل خيراً أو ليصمت","حديث شريف","h"],["وَتَعَاوَنُوا عَلَى الْبِرِّ وَالتَّقْوَىٰ ۖ وَلَا تَعَاوَنُوا عَلَى الْإِثْمِ وَالْعُدْوَانِ","سورة المائدة، آية 2","v"],["إذا طبخت مرقةً فأكثر ماءها، وتعاهد جيرانك","حديث شريف","h"],["إِنَّ اللَّهَ يَأْمُرُ بِالْعَدْلِ وَالْإِحْسَانِ وَإِيتَاءِ ذِي الْقُرْبَىٰ","سورة النحل، آية 90","v"],["خير الأصحاب عند الله خيرهم لصاحبه، وخير الجيران عند الله خيرهم لجاره","حديث شريف","h"],["وَأَحْسِنُوا ۛ إِنَّ اللَّهَ يُحِبُّ الْمُحْسِنِينَ","سورة البقرة، آية 195","v"],["حسن الخلق: بسط الوجه، وبذل المعروف، وكفّ الأذى","يُروى عن عبد الله بن المبارك","s"],["لا ضرر ولا ضرار","حديث شريف","h"],["ادْفَعْ بِالَّتِي هِيَ أَحْسَنُ فَإِذَا الَّذِي بَيْنَكَ وَبَيْنَهُ عَدَاوَةٌ كَأَنَّهُ وَلِيٌّ حَمِيمٌ","سورة فصلت، آية 34","v"],["تبسّمك في وجه أخيك لك صدقة","حديث شريف","h"],["إِنَّمَا الْمُؤْمِنُونَ إِخْوَةٌ فَأَصْلِحُوا بَيْنَ أَخَوَيْكُمْ","سورة الحجرات، آية 10","v"],["لا تحقرنّ من المعروف شيئاً ولو أن تلقى أخاك بوجه طلق","حديث شريف","h"],["وَاعْتَصِمُوا بِحَبْلِ اللَّهِ جَمِيعًا وَلَا تَفَرَّقُوا","سورة آل عمران، آية 103","v"],["الكلمة الطيبة صدقة","حديث شريف","h"],["الدين كله خلق، فمن زاد عليك في الخلق زاد عليك في الدين","ابن القيم، مدارج السالكين","s"],["وَالْكَاظِمِينَ الْغَيْظَ وَالْعَافِينَ عَنِ النَّاسِ ۗ وَاللَّهُ يُحِبُّ الْمُحْسِنِينَ","سورة آل عمران، آية 134","v"],["اتق الله حيثما كنت، وأتبع السيئة الحسنة تمحها، وخالق الناس بخلق حسن","حديث شريف","h"],["خُذِ الْعَفْوَ وَأْمُرْ بِالْعُرْفِ وَأَعْرِضْ عَنِ الْجَاهِلِينَ","سورة الأعراف، آية 199","v"],["إن من أحبكم إليّ وأقربكم مني مجلساً يوم القيامة أحاسنكم أخلاقاً","حديث شريف","h"],["وَلْيَعْفُوا وَلْيَصْفَحُوا ۗ أَلَا تُحِبُّونَ أَن يَغْفِرَ اللَّهُ لَكُمْ","سورة النور، آية 22","v"],["إنما بُعثت لأتمم مكارم الأخلاق","حديث شريف","h"],["وَاقْصِدْ فِي مَشْيِكَ وَاغْضُضْ مِن صَوْتِكَ","سورة لقمان، آية 19","v"],["حاسبوا أنفسكم قبل أن تُحاسبوا","يُروى عن عمر بن الخطاب","s"],["الطهور شطر الإيمان","حديث شريف","h"],["وَلَا تُصَعِّرْ خَدَّكَ لِلنَّاسِ وَلَا تَمْشِ فِي الْأَرْضِ مَرَحًا","سورة لقمان، آية 18","v"],["اتقوا اللاعنَين: الذي يتخلى في طريق الناس أو في ظلهم","حديث شريف","h"],["اجْتَنِبُوا كَثِيرًا مِّنَ الظَّنِّ إِنَّ بَعْضَ الظَّنِّ إِثْمٌ ۖ وَلَا تَجَسَّسُوا","سورة الحجرات، آية 12","v"],["إن الله كتب الإحسان على كل شيء","حديث شريف","h"],["يَا أَيُّهَا الَّذِينَ آمَنُوا لَا تَدْخُلُوا بُيُوتًا غَيْرَ بُيُوتِكُمْ حَتَّىٰ تَسْتَأْنِسُوا وَتُسَلِّمُوا عَلَىٰ أَهْلِهَا","سورة النور، آية 27","v"],["الراحمون يرحمهم الرحمن، ارحموا من في الأرض يرحمكم من في السماء","حديث شريف","h"],["من عدّ كلامه من عمله قلّ كلامه إلا فيما يعنيه","عمر بن عبد العزيز","s"],["وَإِذَا حُيِّيتُم بِتَحِيَّةٍ فَحَيُّوا بِأَحْسَنَ مِنْهَا أَوْ رُدُّوهَا","سورة النساء، آية 86","v"],["من لا يرحم لا يُرحم","حديث شريف","h"],["إن الرفق لا يكون في شيء إلا زانه، ولا يُنزع من شيء إلا شانه","حديث شريف","h"],["وَأَوْفُوا بِالْعَهْدِ ۖ إِنَّ الْعَهْدَ كَانَ مَسْئُولًا","سورة الإسراء، آية 34","v"],["إن الله رفيق يحب الرفق في الأمر كله","حديث شريف","h"],["وَآتِ ذَا الْقُرْبَىٰ حَقَّهُ وَالْمِسْكِينَ وَابْنَ السَّبِيلِ وَلَا تُبَذِّرْ تَبْذِيرًا","سورة الإسراء، آية 26","v"],["إذا قدرت على عدوك فاجعل العفو عنه شكراً للقدرة عليه","يُنسب إلى علي بن أبي طالب","s"],["ليس الشديد بالصرعة، إنما الشديد الذي يملك نفسه عند الغضب","حديث شريف","h"],["وَقُل لِّعِبَادِي يَقُولُوا الَّتِي هِيَ أَحْسَنُ","سورة الإسراء، آية 53","v"],["لا تغضب","حديث شريف","h"],["لَن تَنَالُوا الْبِرَّ حَتَّىٰ تُنفِقُوا مِمَّا تُحِبُّونَ","سورة آل عمران، آية 92","v"],["الدين النصيحة","حديث شريف","h"],["مَّن ذَا الَّذِي يُقْرِضُ اللَّهَ قَرْضًا حَسَنًا فَيُضَاعِفَهُ لَهُ أَضْعَافًا كَثِيرَةً","سورة البقرة، آية 245","v"],["المستشار مؤتمن","حديث شريف","h"],["وَمَا أَنفَقْتُم مِّن شَيْءٍ فَهُوَ يُخْلِفُهُ ۖ وَهُوَ خَيْرُ الرَّازِقِينَ","سورة سبأ، آية 39","v"],["قيمة كل امرئ ما يحسنه","يُنسب إلى علي بن أبي طالب","s"],["أولا أدلكم على شيء إذا فعلتموه تحاببتم؟ أفشوا السلام بينكم","حديث شريف","h"],["هَلْ جَزَاءُ الْإِحْسَانِ إِلَّا الْإِحْسَانُ","سورة الرحمن، آية 60","v"],["تهادوا تحابوا","حديث شريف","h"],["فَمَن يَعْمَلْ مِثْقَالَ ذَرَّةٍ خَيْرًا يَرَهُ","سورة الزلزلة، آية 7","v"],["إذا استأذن أحدكم ثلاثاً فلم يؤذن له فليرجع","حديث شريف","h"],["وَتَوَاصَوْا بِالْحَقِّ وَتَوَاصَوْا بِالصَّبْرِ","سورة العصر، آية 3","v"],["الناس أعداء ما جهلوا","يُنسب إلى علي بن أبي طالب","s"],["إنما جُعل الاستئذان من أجل البصر","حديث شريف","h"],["وَالصُّلْحُ خَيْرٌ","سورة النساء، آية 128","v"],["والله في عون العبد ما كان العبد في عون أخيه","حديث شريف","h"],["وَيُؤْثِرُونَ عَلَىٰ أَنفُسِهِمْ وَلَوْ كَانَ بِهِمْ خَصَاصَةٌ","سورة الحشر، آية 9","v"],["أحب الناس إلى الله أنفعهم للناس","حديث شريف","h"],["وَلَا تَنسَوُا الْفَضْلَ بَيْنَكُمْ","سورة البقرة، آية 237","v"],["خيركم خيركم لأهله، وأنا خيركم لأهلي","حديث شريف","h"],["ابن آدم، إنما أنت أيام، كلما ذهب يوم ذهب بعضك","يُروى عن الحسن البصري","s"],["إِنَّ اللَّهَ مَعَ الصَّابِرِينَ","سورة البقرة، آية 153","v"],["ما نقصت صدقة من مال","حديث شريف","h"],["فَاصْبِرْ صَبْرًا جَمِيلًا","سورة المعارج، آية 5","v"],["اتقوا النار ولو بشق تمرة","حديث شريف","h"],["وَمَن يَتَوَكَّلْ عَلَى اللَّهِ فَهُوَ حَسْبُهُ","سورة الطلاق، آية 3","v"],["الصدقة تطفئ الخطيئة كما يطفئ الماء النار","حديث شريف","h"],["إِنَّ مَعَ الْعُسْرِ يُسْرًا","سورة الشرح، آية 6","v"],["وتعين الرجل في دابته فتحمله عليها أو ترفع له عليها متاعه صدقة","حديث شريف","h"],["رأيي صواب يحتمل الخطأ، ورأي غيري خطأ يحتمل الصواب","يُنسب إلى الإمام الشافعي","s"],["اليد العليا خير من اليد السفلى","حديث شريف","h"],["وَقُل رَّبِّ زِدْنِي عِلْمًا","سورة طه، آية 114","v"],["إن الله يحب إذا عمل أحدكم عملاً أن يتقنه","حديث شريف","h"],["إِنَّ اللَّهَ لَا يُغَيِّرُ مَا بِقَوْمٍ حَتَّىٰ يُغَيِّرُوا مَا بِأَنفُسِهِمْ","سورة الرعد، آية 11","v"],["كلكم راعٍ وكلكم مسؤول عن رعيته","حديث شريف","h"],["وَأَن لَّيْسَ لِلْإِنسَانِ إِلَّا مَا سَعَىٰ","سورة النجم، آية 39","v"],["الظلم مؤذن بخراب العمران","ابن خلدون، المقدمة","s"],["عليكم بالصدق، فإن الصدق يهدي إلى البر","حديث شريف","h"],["وَقُلِ اعْمَلُوا فَسَيَرَى اللَّهُ عَمَلَكُمْ وَرَسُولُهُ وَالْمُؤْمِنُونَ","سورة التوبة، آية 105","v"],["دع ما يريبك إلى ما لا يريبك","حديث شريف","h"],["وَأَمْرُهُمْ شُورَىٰ بَيْنَهُمْ","سورة الشورى، آية 38","v"],["البر حسن الخلق","حديث شريف","h"],["وَشَاوِرْهُمْ فِي الْأَمْرِ","سورة آل عمران، آية 159","v"],["إن الحياء لا يأتي إلا بخير","حديث شريف","h"],["وإنما الأمم الأخلاق ما بقيت، فإن همُ ذهبت أخلاقهم ذهبوا","أحمد شوقي","s"],["فَبِمَا رَحْمَةٍ مِّنَ اللَّهِ لِنتَ لَهُمْ ۖ وَلَوْ كُنتَ فَظًّا غَلِيظَ الْقَلْبِ لَانفَضُّوا مِنْ حَوْلِكَ","سورة آل عمران، آية 159","v"],["إياكم والظن، فإن الظن أكذب الحديث","حديث شريف","h"],["فَمَنْ عَفَا وَأَصْلَحَ فَأَجْرُهُ عَلَى اللَّهِ","سورة الشورى، آية 40","v"],["المرء على دين خليله، فلينظر أحدكم من يخالل","حديث شريف","h"],["وَلَمَن صَبَرَ وَغَفَرَ إِنَّ ذَٰلِكَ لَمِنْ عَزْمِ الْأُمُورِ","سورة الشورى، آية 43","v"],["من لا يشكر الناس لا يشكر الله","حديث شريف","h"],["على قدر أهل العزم تأتي العزائم","المتنبي","s"],["إِنَّ اللَّهَ يُحِبُّ التَّوَّابِينَ وَيُحِبُّ الْمُتَطَهِّرِينَ","سورة البقرة، آية 222","v"],["من صنع إليكم معروفاً فكافئوه، فإن لم تجدوا ما تكافئونه فادعوا له","حديث شريف","h"],["وَاللَّهُ يُحِبُّ الْمُطَّهِّرِينَ","سورة التوبة، آية 108","v"],["احرص على ما ينفعك","حديث شريف","h"],["وَلَا تُفْسِدُوا فِي الْأَرْضِ بَعْدَ إِصْلَاحِهَا","سورة الأعراف، آية 56","v"],["من دلّ على خير فله مثل أجر فاعله","حديث شريف","h"],["وَلَا تَبْخَسُوا النَّاسَ أَشْيَاءَهُمْ","سورة الأعراف، آية 85","v"],["من غشّنا فليس منا","حديث شريف","h"],["إن للحسنة ضياءً في الوجه، ونوراً في القلب، وسعةً في الرزق، ومحبةً في قلوب الخلق","يُروى عن ابن عباس","s"],["إِنَّ اللَّهَ يَأْمُرُكُمْ أَن تُؤَدُّوا الْأَمَانَاتِ إِلَىٰ أَهْلِهَا","سورة النساء، آية 58","v"],["وما تواضع أحد لله إلا رفعه الله","حديث شريف","h"],["وَالَّذِينَ هُمْ لِأَمَانَاتِهِمْ وَعَهْدِهِمْ رَاعُونَ","سورة المؤمنون، آية 8","v"],["ليس الغنى عن كثرة العرض، ولكن الغنى غنى النفس","حديث شريف","h"],["إنما الصبر عند الصدمة الأولى","حديث شريف","h"],["يَا أَيُّهَا الَّذِينَ آمَنُوا أَوْفُوا بِالْعُقُودِ","سورة المائدة، آية 1","v"],["الجار قبل الدار، والرفيق قبل الطريق","مثل عربي مأثور","s"],["ومن يتصبّر يصبّره الله","حديث شريف","h"],["وَإِن جَنَحُوا لِلسَّلْمِ فَاجْنَحْ لَهَا","سورة الأنفال، آية 61","v"],["إن الله لا ينظر إلى صوركم وأموالكم، ولكن ينظر إلى قلوبكم وأعمالكم","حديث شريف","h"],["يَا أَيُّهَا الَّذِينَ آمَنُوا اتَّقُوا اللَّهَ وَقُولُوا قَوْلًا سَدِيدًا","سورة الأحزاب، آية 70","v"],["يد الله مع الجماعة","حديث شريف","h"],["لَّا خَيْرَ فِي كَثِيرٍ مِّن نَّجْوَاهُمْ إِلَّا مَنْ أَمَرَ بِصَدَقَةٍ أَوْ مَعْرُوفٍ أَوْ إِصْلَاحٍ بَيْنَ النَّاسِ","سورة النساء، آية 114","v"],["لا تُنزع الرحمة إلا من شقي","حديث شريف","h"],["من جاور السعيد يسعد","مثل عربي","s"],["وَإِنَّكَ لَعَلَىٰ خُلُقٍ عَظِيمٍ","سورة القلم، آية 4","v"],["أدِّ الأمانة إلى من ائتمنك","حديث شريف","h"],["فَاعْفُ عَنْهُمْ وَاصْفَحْ ۚ إِنَّ اللَّهَ يُحِبُّ الْمُحْسِنِينَ","سورة المائدة، آية 13","v"],["إذا كنتم ثلاثة فلا يتناجى اثنان دون الثالث","حديث شريف","h"],["فَاسْتَبِقُوا الْخَيْرَاتِ","سورة البقرة، آية 148","v"],["لا يقيم الرجل الرجل من مجلسه ثم يجلس فيه","حديث شريف","h"],["رَبَّنَا آتِنَا فِي الدُّنْيَا حَسَنَةً وَفِي الْآخِرَةِ حَسَنَةً وَقِنَا عَذَابَ النَّارِ","سورة البقرة، آية 201","v"],["يد واحدة لا تصفّق","مثل عربي","s"],["رحم الله رجلاً سمحاً إذا باع، وإذا اشترى، وإذا اقتضى","حديث شريف","h"],["وَلَا تَنَابَزُوا بِالْأَلْقَابِ","سورة الحجرات، آية 11","v"],["من أنظر معسراً أو وضع عنه أظله الله في ظله","حديث شريف","h"],["وَاسْتَعِينُوا بِالصَّبْرِ وَالصَّلَاةِ","سورة البقرة، آية 45","v"],["من ضارّ أضرّ الله به، ومن شاقّ شقّ الله عليه","حديث شريف","h"],["يَا أَيُّهَا النَّاسُ إِنَّا خَلَقْنَاكُم مِّن ذَكَرٍ وَأُنثَىٰ وَجَعَلْنَاكُمْ شُعُوبًا وَقَبَائِلَ لِتَعَارَفُوا","سورة الحجرات، آية 13","v"],["الحياء من الإيمان","حديث شريف","h"],["من أصلح سريرته أصلح الله علانيته","يُنسب إلى علي بن أبي طالب","s"],["ثُمَّ كَانَ مِنَ الَّذِينَ آمَنُوا وَتَوَاصَوْا بِالصَّبْرِ وَتَوَاصَوْا بِالْمَرْحَمَةِ","سورة البلد، آية 17","v"],["لا يمنعنّ جارٌ جارَه أن يغرز خشبة في جداره","حديث شريف","h"],["وَإِن كَانَ ذُو عُسْرَةٍ فَنَظِرَةٌ إِلَىٰ مَيْسَرَةٍ ۚ وَأَن تَصَدَّقُوا خَيْرٌ لَّكُمْ","سورة البقرة، آية 280","v"],["من كان معه فضل ظهر فليعد به على من لا ظهر له، ومن كان له فضل زاد فليعد به على من لا زاد له","حديث شريف","h"],["الَّذِينَ يُنفِقُونَ أَمْوَالَهُم بِاللَّيْلِ وَالنَّهَارِ سِرًّا وَعَلَانِيَةً فَلَهُمْ أَجْرُهُمْ عِندَ رَبِّهِمْ","سورة البقرة، آية 274","v"],["اتقوا الظلم، فإن الظلم ظلمات يوم القيامة","حديث شريف","h"],["أحسن إلى من شئت تكن أميره، واستغن عمّن شئت تكن نظيره","يُنسب إلى علي بن أبي طالب","s"],["وَمَن يُؤْتَ الْحِكْمَةَ فَقَدْ أُوتِيَ خَيْرًا كَثِيرًا","سورة البقرة، آية 269","v"],["أحب الأعمال إلى الله أدومها وإن قلّ","حديث شريف","h"],["وَلَا تَقْفُ مَا لَيْسَ لَكَ بِهِ عِلْمٌ","سورة الإسراء، آية 36","v"],["نعمتان مغبون فيهما كثير من الناس: الصحة والفراغ","حديث شريف","h"]];
function todayQuote() {
  const p = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd').split('-').map(Number);
  const n = Math.floor(Date.UTC(p[0], p[1] - 1, p[2]) / 86400000);
  const q = QUOTES[n % QUOTES.length];
  return { t: q[0], s: q[1], v: q[2] === 'v' ? 1 : 0 };
}

function doGet(e) {
  const p = e.parameter || {};
  if (p.t) return decisionPage(String(p.t));
  CTX = { did: String(p.did || '').slice(0, 40), act: 'view' };
  return json(check(nd(p.apt), nd(p.pin)));
}
function doPost(e) {
  let b = {};
  try { b = JSON.parse(e.postData.contents); } catch (x) {}
  CTX = { did: String(b.did || '').slice(0, 40), act: 'upload' };
  if (b.action === 'upload') return json(upload(b));
  if (b.action === 'm_list') { CTX.act = 'm_view'; return json(mList(b)); }
  if (b.action === 'm_update') { CTX.act = 'm_update'; return json(mUpdate(b)); }
  if (b.action === 'a_list') { CTX.act = 'a_view'; return json(aList(b)); }
  const AW = { a_toggle: aToggle, a_ann: aAnn, a_project: aProject, a_decide: aDecide };
  if (AW[b.action]) { CTX.act = 'a_write'; return json(AW[b.action](b)); }
  return json({ ok: false, error: 'bad' });
}

// ---------- التحقق (شقة + رقم سري + قفل) ----------
function auth(apt, pin) {
  const props = PropertiesService.getScriptProperties();
  const apts = SS.getSheetByName('Apartments').getDataRange().getValues().slice(1);
  const row = apts.find(r => nd(r[0]) === apt);
  if (!row) { logAccess(apt, 'bad_apt'); return { error: 'bad' }; }

  const key = 'f_' + apt;
  const now = Date.now();
  let f = JSON.parse(props.getProperty(key) || '{"n":0,"t":0}');
  if (f.n >= 5 && now - f.t < 3600000) { logAccess(apt, 'locked'); return { error: 'locked' }; }
  if (f.n >= 5) f.n = 0;

  if (pinNorm(row[3]) !== pinNorm(pin)) {
    props.setProperty(key, JSON.stringify({ n: f.n + 1, t: now }));
    logAccess(apt, 'bad_pin');
    return { error: 'bad' };
  }
  props.deleteProperty(key);
  logAccess(apt, 'ok');
  return { row: row };
}

// ---------- بيانات الشقة ----------
function check(apt, pin) {
  const a = auth(apt, pin);
  if (a.error) return { ok: false, error: a.error };

  const items = SS.getSheetByName('Items').getDataRange().getValues().slice(1).filter(r => r[0]);
  const pays = SS.getSheetByName('Payments').getDataRange().getValues().slice(1)
    .filter(r => nd(r[1]) === apt);

  // المطلوب = البنود غير المكتملة فقط (الحالية). البند المكتمل ينتقل إلى done (المشاريع المدفوعة) ولا يدخل في المطلوب
  const all = items.map(r => {
    const mine = pays.filter(x => ck(x[2]) === ck(r[0]));
    const p = mine.reduce((s, x) => s + Number(x[3]), 0);
    const amt = Number(r[3]) || 0;
    return { code: r[0], desc: r[1], date: fmt(r[2]), due: fmt(r[5] || r[2]), amount: amt, paid: p, remaining: Math.max(amt - p, 0),
      paidOn: mine.map(x => fmt(x[0])).sort().pop() || '' };
  });
  const list = all.filter(i => i.remaining > 0);
  const done = all.filter(i => i.remaining <= 0 && i.amount > 0).sort((x, y) => y.paidOn.localeCompare(x.paidOn));
  let required = 0, paid = 0;
  list.forEach(i => { required += i.amount; paid += i.paid; });

  const cs = SS.getSheetByName('Claims');
  const mine = cs ? cs.getDataRange().getValues().slice(1).filter(r => nd(r[1]) === apt && st(r[4])) : [];
  const names = {};
  all.forEach(i => { names[ck(i.code)] = i.desc; });
  list.forEach(i => {
    const cl = mine.filter(r => ck(r[2]) === ck(i.code));
    const last = cl.length ? cl[cl.length - 1] : null;
    const pen = cl.filter(r => st(r[4]) === 'pending').pop();
    i.pending = i.remaining > 0 && !!pen;
    i.pendingAmount = i.pending ? Number(pen[3]) || 0 : 0;
    i.rejected = i.remaining > 0 && !i.pending && !!last && st(last[4]) === 'rejected';
    i.rejectNote = i.rejected ? String(last[7] || '').trim() : '';
  });
  const claims = mine.slice(-10).reverse().map(r => ({
    date: fmt(r[0]), desc: names[ck(r[2])] || String(r[2]), amount: Number(r[3]) || 0,
    status: st(r[4]), note: st(r[4]) === 'rejected' ? String(r[7] || '').trim() : ''
  }));

  const s = settings();

  let expenses = null;
  if (yes(s.show_expenses)) {
    expenses = SS.getSheetByName('Expenses').getDataRange().getValues().slice(1)
      .filter(r => r[1] || r[2])
      .map(r => ({ date: fmt(r[0]), desc: String(r[1] || ''), amount: Number(r[2]) || 0, details: String(r[3] || '') }))
      .sort((x, y) => y.date.localeCompare(x.date));
  }

  return {
    ok: true,
    apt: apt,
    required: required,
    paid: paid,
    remaining: Math.max(required - paid, 0),
    items: list,
    done: done,
    payments: pays.map(x => ({ date: fmt(x[0]), code: x[2], amount: Number(x[3]) })),
    claims: claims,
    fund: yes(s.show_fund)
      ? (Number(SS.getSheetByName('Summary').getRange('H2').getValue()) || 0)
      : null,
    expenses: expenses,
    project: yes(s.show_project) ? openProjects() : null,
    quote: yes(s.show_quote) ? todayQuote() : null,
    settings: {
      treasurer_name: s.treasurer_name || '',
      treasurer_phone: intl(s.treasurer_phone),
      pay_info: s.pay_info || '',
      pay: [1, 2].map(n => ({ type: s['pay' + n + '_type'] || '', number: payNum(s['pay' + n + '_number']) })).filter(p => p.number),
      announcement: s.announcement || ''
    }
  };
}

// ---------- رفع الإيصال + تسجيل "في انتظار التأكيد" ----------
function upload(b) {
  const apt = nd(b.apt), pin = nd(b.pin), code = String(b.code || '').trim();
  const a = auth(apt, pin);
  if (a.error) return { ok: false, error: a.error };

  const mime = String(b.mime || '');
  const ext = { 'image/jpeg': 'jpg', 'image/png': 'png', 'image/webp': 'webp' }[mime];
  if (!ext || !b.data || String(b.data).length > 3000000) return { ok: false, error: 'file' };

  const item = SS.getSheetByName('Items').getDataRange().getValues().slice(1).find(r => ck(r[0]) === ck(code));
  if (!item) return { ok: false, error: 'item' };

  const paid0 = SS.getSheetByName('Payments').getDataRange().getValues().slice(1)
    .filter(r => nd(r[1]) === apt && ck(r[2]) === ck(code))
    .reduce((t, x) => t + Number(x[3]), 0);
  const rest0 = Math.max(Number(item[3]) - paid0, 0);
  if (rest0 <= 0) return { ok: false, error: 'done' };
  const amount = (b.amount === undefined || b.amount === null || b.amount === '') ? rest0 : Math.round(Number(nd(b.amount)) * 100) / 100;
  if (!isFinite(amount) || amount <= 0 || amount > rest0) return { ok: false, error: 'amount' };

  const now = new Date();
  const year = Number(Utilities.formatDate(now, TZ, 'yyyy'));
  const mi = Number(Utilities.formatDate(now, TZ, 'M')) - 1;
  if (!PropertiesService.getScriptProperties().getProperty('ens_' + year)) ensureYear(year);
  const month = getOrCreate(yearFolder('الإيصالات', year), MONTHS[mi]);

  const base = clean('شقة ' + apt + ' - ' + Utilities.formatDate(now, TZ, 'yyyy-MM-dd') + ' - ' + item[1]);
  let name = base + '.' + ext;
  if (month.getFilesByName(name).hasNext()) name = base + ' - ' + Utilities.formatDate(now, TZ, 'HHmmss') + '.' + ext;
  const file = month.createFile(Utilities.newBlob(Utilities.base64Decode(b.data), mime, name));

  return addClaim(apt, a.row, code, item, file.getUrl(), amount);
}

function addClaim(apt, row, code, item, fileUrl, amount) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) {}
  try {
    let cs = SS.getSheetByName('Claims');
    if (!cs) {
      cs = SS.insertSheet('Claims');
      initClaims(cs);
    }
    const rows = cs.getDataRange().getValues().slice(1);
    const idx = rows.findIndex(r => nd(r[1]) === apt && ck(r[2]) === ck(code) && st(r[4]) === 'pending');
    let id;
    if (idx >= 0) {
      const cell = cs.getRange(idx + 2, 6);
      cell.setValue(cell.getValue() ? cell.getValue() + '\n' + fileUrl : fileUrl);
      cs.getRange(idx + 2, 4).setValue(amount);
      id = String(rows[idx][8] || '');
      if (!id) { id = newId(); cs.getRange(idx + 2, 9).setValue(id); }
    } else {
      id = newId();
      cs.appendRow([new Date(), apt, code, amount, ST_AR.pending, fileUrl, '', '', id]);
    }

    try {
      const base = ScriptApp.getService().getUrl();
      const cUrl = base ? base + '?t=' + makeToken(id, 'c') : '';
      const rUrl = base ? base + '?t=' + makeToken(id, 'r') : '';
      const head = 'شقة ' + apt + ' (' + row[1] + ') رفعت إيصال مساهمة بقيمة ' + amount + ' جنيه في بند "' + item[1] + '".';
      const plain = head + '\nالإيصال: ' + fileUrl +
        (cUrl ? '\n\nتأكيد: ' + cUrl + '\nرفض: ' + rUrl + '\n(الرابط صالح 14 يوماً ولمرة واحدة، ويفتح صفحة تطلب منك الضغط للتنفيذ)' : '\nغيّر الحالة في Claims.');
      const btn = (u, t, c) => '<a href="' + u + '" style="display:inline-block;padding:11px 22px;margin:4px;border-radius:10px;background:' + c + ';color:#fff;text-decoration:none;font-weight:bold">' + t + '</a>';
      const html = '<div dir="rtl" style="font-family:Tahoma,Arial,sans-serif;font-size:15px;line-height:1.8">' + escH(head) +
        '<br><a href="' + escH(fileUrl.split('\n')[0]) + '">فتح صورة الإيصال</a>' +
        (cUrl ? '<br><br>' + btn(cUrl, '✅ تأكيد', '#12a150') + btn(rUrl, '❌ رفض', '#d93a3a') +
          '<br><small style="color:#777">الرابط صالح 14 يوماً ولمرة واحدة، ويفتح صفحة تطلب منك الضغط للتنفيذ.</small>' : '') + '</div>';
      MailApp.sendEmail({ to: adminEmail(), subject: 'إيصال مساهمة جديد - شقة ' + apt, body: plain, htmlBody: html });
    } catch (err) {}
    return { ok: true };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// ---------- تهيئة وتحديث الشيت ----------
function initClaims(cs) {
  if (cs.getLastRow() === 0)
    cs.appendRow(['التاريخ', 'الشقة', 'كود البند', 'المبلغ', 'الحالة (قيد المراجعة / تم التأكيد / مرفوض)', 'رابط الإيصال', 'تاريخ التسجيل في Payments', 'ملاحظة الرفض (اختياري)', 'معرّف (لا تغيّره)']);
  try { cs.hideColumns(9); } catch (e) {}
  cs.getRange('E2:E500').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['قيد المراجعة', 'تم التأكيد', 'مرفوض'], true).setAllowInvalid(false).build());
}

// شغّلها مرة واحدة على الشيت القديم: تحوّل الأسماء والحالات للعربي وتضيف القوائم المنسدلة وتحمي الأرقام من ضياع الصفر
function migrate() {
  const YN = SpreadsheetApp.newDataValidation().requireValueInList(['نعم', 'لا'], true).setAllowInvalid(false).build();

  const se = SS.getSheetByName('Settings');
  if (se) {
    let vals = se.getDataRange().getValues();
    const have = {};
    vals.slice(1).forEach(r => { have[skey(r[0])] = true; });
    Object.keys(SK_REV).forEach(k => { if (!have[k]) se.appendRow([SK_REV[k], k.indexOf('show_') === 0 ? 'لا' : '']); });
    vals = se.getDataRange().getValues();
    for (let i = 1; i < vals.length; i++) {
      const k = skey(vals[i][0]);
      if (!k) continue;
      if (SK_REV[k]) se.getRange(i + 1, 1).setValue(SK_REV[k]);
      if (k === 'pay1_type' || k === 'pay2_type')
        se.getRange(i + 1, 2).setDataValidation(SpreadsheetApp.newDataValidation()
          .requireValueInList(['إنستا باي', 'محفظة إلكترونية'], true).setAllowInvalid(true).build());
      if (/^show_/.test(k))
        se.getRange(i + 1, 2).setValue(yes(vals[i][1]) ? 'نعم' : 'لا').setDataValidation(YN);
    }
    se.getRange(2, 2, Math.max(se.getLastRow() - 1, 1), 1).setNumberFormat('@');
  }

  const cs = SS.getSheetByName('Claims');
  if (cs) {
    const n = cs.getLastRow();
    if (n > 1) {
      const rg = cs.getRange(2, 5, n - 1, 1);
      rg.setValues(rg.getValues().map(x => [ST_AR[st(x[0])] || x[0]]));
    }
    cs.getRange(1, 5).setValue('الحالة (قيد المراجعة / تم التأكيد / مرفوض)');
    cs.getRange(1, 7).setValue('تاريخ التسجيل في Payments');
    cs.getRange(1, 8).setValue('ملاحظة الرفض (اختياري)');
    cs.getRange(1, 9).setValue('معرّف (لا تغيّره)');
    initClaims(cs);
  }

  tasksSheet();

  const al = SS.getSheetByName('AccessLog');
  if (al) {
    const n = al.getLastRow();
    if (n > 1) {
      const c = al.getRange(2, 3, n - 1, 1); c.setValues(c.getValues().map(x => [RES_AR[x[0]] || x[0]]));
      const a = al.getRange(2, 5, n - 1, 1); a.setValues(a.getValues().map(x => [ACT_AR[x[0]] || x[0]]));
    }
    al.getRange(1, 3).setValue('النتيجة (نجح / رقم سري خاطئ / شقة غير موجودة / مقفول)');
  }

  const ap = SS.getSheetByName('Apartments');
  if (ap) ap.getRange('C2:D200').setNumberFormat('@');

  // معادلات تنبيه واتساب بالصياغة الجديدة
  const su = SS.getSheetByName('Summary');
  if (su) {
    const f = []; for (let r = 2; r <= 31; r++) f.push(summaryCols(r));
    su.getRange(2, 1, 30, 6).setFormulas(f);
  }

  const s2 = settings();
  const old = s2.pay_info && !s2.pay1_number && !s2.pay2_number;
  SS.toast('تم تحديث الشيت (الإعدادات، Claims، Tasks، معادلات Summary وواتساب).' + (old ? ' انقل بيانات الدفع القديمة إلى وسيلة الدفع 1 و2 في Settings.' : ''), 'اكتمل');
}

// ---------- تأكيد الإيصال ⟵ تسجيل الدفعة تلقائياً ----------
// Simple trigger: يعمل تلقائياً عند تعديل الشيت، ولا يحتاج Deploy ولا installTrigger.
// لما تغيّر الحالة في Claims (عمود E) إلى confirmed تُضاف دفعة في Payments بنفس الشقة والبند والمبلغ (عمود D).
function onEdit(e) {
  try {
    const r = e.range, sh = r.getSheet();
    if (sh.getName() !== 'Claims' || r.getColumn() !== 5 || r.getRow() < 2) return;
    if (st(e.value) !== 'confirmed') return;
    const res = registerClaimPayment(sh, r.getRow());
    if (res.ok) SS.toast(res.msg, 'تم');
    else if (res.warn) SS.toast(res.msg, 'لم تُسجَّل الدفعة');
  } catch (err) {}
}

// تسجّل دفعة الصف في Payments (مرة واحدة فقط بعمود G). تُستخدم من onEdit ومن زر التأكيد في الإيميل
function registerClaimPayment(sh, rowIdx) {
  const row = sh.getRange(rowIdx, 1, 1, 7).getValues()[0];
  if (row[6]) return { ok: false, msg: 'سُجّلت قبل كده' }; // لا تكرار
  const apt = nd(row[1]), code = String(row[2]).trim(), amt = Number(row[3]);
  if (!apt || !code || !(amt > 0)) return { ok: false, warn: true, msg: 'تأكد من الشقة والبند والمبلغ في Claims' };

  const pay = SS.getSheetByName('Payments');
  const colA = pay.getRange(1, 1, pay.getMaxRows(), 1).getValues();
  let i = 1;
  while (i < colA.length && String(colA[i][0]) !== '') i++;
  const rowNum = i + 1;
  pay.getRange(rowNum, 1, 1, 4).setValues([[new Date(), apt, code, amt]]);
  pay.getRange(rowNum, 1).setNumberFormat('yyyy-mm-dd');
  pay.getRange(rowNum, 5).setFormula('=IFERROR(VLOOKUP(C' + rowNum + ',Items!A:E,5,0),"")');
  sh.getRange(rowIdx, 7).setValue(new Date()).setNumberFormat('yyyy-mm-dd hh:mm');
  return { ok: true, msg: 'تم تسجيل ' + amt + ' جنيه لشقة ' + apt + ' في Payments' };
}

// ---------- أزرار الإيميل (تأكيد / رفض) برمز موقّع ----------
// الرابط GET يفتح صفحة فقط ولا ينفّذ شيئاً (حتى لا تنفّذه برامج فحص الروابط)، والتنفيذ يتم بزر داخل الصفحة.
// الاستخدام لمرة واحدة: بعد أي قرار تتغير حالة الإيصال فيصبح الرابط (والآخر) غير صالح.
function newId() { return Utilities.getUuid().replace(/-/g, '').slice(0, 12); }
function escH(v) { return String(v == null ? '' : v).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }
function secret() {
  const props = PropertiesService.getScriptProperties();
  let k = props.getProperty('hmac_secret');
  if (!k) { k = Utilities.getUuid() + Utilities.getUuid(); props.setProperty('hmac_secret', k); }
  return k;
}
function sign(msg) {
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(msg, secret())).replace(/=+$/, '');
}
function makeToken(id, act) {
  const exp = Date.now() + 14 * 24 * 3600 * 1000;
  const msg = id + '.' + act + '.' + exp;
  return msg + '.' + sign(msg);
}
function readToken(t) {
  const m = String(t || '').match(/^([a-f0-9]{12})\.([cr])\.(\d{10,15})\.([A-Za-z0-9_-]+)$/);
  if (!m) return null;
  const good = sign(m[1] + '.' + m[2] + '.' + m[3]);
  let diff = good.length ^ m[4].length;
  for (let i = 0; i < good.length; i++) diff |= good.charCodeAt(i) ^ (m[4].charCodeAt(i) || 0);
  if (diff !== 0 || Date.now() > Number(m[3])) return null;
  return { id: m[1], act: m[2] };
}
function findClaim(id) {
  const sh = SS.getSheetByName('Claims');
  if (!sh) return null;
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) if (String(rows[i][8]) === id) return { sh: sh, rowIdx: i + 1, r: rows[i] };
  return null;
}
function pageHtml(body, script) {
  return HtmlService.createHtmlOutput('<!DOCTYPE html><html dir="rtl" lang="ar"><head><meta charset="utf-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1"><base target="_top"></head>' +
    '<body style="font-family:Tahoma,Arial,sans-serif;max-width:460px;margin:30px auto;padding:0 16px;line-height:1.8;color:#1a1f33">' +
    body + (script ? '<script>' + script + '</script>' : '') + '</body></html>').setTitle('عمارة 220');
}
function decisionPage(t) {
  const tk = readToken(t);
  const c = tk && findClaim(tk.id);
  if (!c) return pageHtml('<h2>الرابط غير صالح أو منتهي</h2><p>افتح تاب Claims في الشيت.</p>');
  if (st(c.r[4]) !== 'pending')
    return pageHtml('<h2>تمت معالجة هذا الإيصال</h2><p>الحالة الحالية: ' + escH(ST_AR[st(c.r[4])] || c.r[4]) + '</p>');
  const items = SS.getSheetByName('Items').getDataRange().getValues().slice(1);
  const it = items.find(r => ck(r[0]) === ck(c.r[2]));
  const conf = tk.act === 'c';
  const info = '<p>شقة <b>' + escH(nd(c.r[1])) + '</b><br>البند: <b>' + escH(it ? it[1] : c.r[2]) + '</b><br>المبلغ: <b>' + escH(c.r[3]) + ' جنيه</b></p>' +
    '<p>' + String(c.r[5] || '').split('\n').map(u => '<a href="' + escH(u) + '" target="_blank">فتح صورة الإيصال</a>').join('<br>') + '</p>';
  const body = '<h2 id="h">' + (conf ? 'تأكيد الإيصال' : 'رفض الإيصال') + '</h2>' + info +
    (conf ? '' : '<p><input id="n" placeholder="سبب الرفض (اختياري)" style="width:100%;padding:10px;font-size:16px;box-sizing:border-box"></p>') +
    '<button id="b" style="width:100%;padding:14px;font-size:17px;border:0;border-radius:12px;color:#fff;background:' + (conf ? '#12a150' : '#d93a3a') + '">' +
    (conf ? '✅ تأكيد وتسجيل الدفعة' : '❌ رفض الإيصال') + '</button><p id="m"></p>';
  const script = 'var T=' + JSON.stringify(t) + ';document.getElementById("b").onclick=function(){var b=this;b.disabled=true;' +
    'var n=document.getElementById("n");' +
    'google.script.run.withSuccessHandler(function(r){document.getElementById("m").textContent=r;b.style.display="none"})' +
    '.withFailureHandler(function(){b.disabled=false;document.getElementById("m").textContent="حدث خطأ، حاول مرة أخرى"})' +
    '.decide(T,n?n.value:"")};';
  return pageHtml(body, script);
}
// تُستدعى من الصفحة عبر google.script.run
function decide(t, note) {
  const tk = readToken(t);
  if (!tk) return 'الرابط غير صالح أو منتهي';
  return decideCore(tk.id, tk.act, note);
}
// المنطق المشترك لتأكيد/رفض إيصال (من رابط الإيميل أو من صفحة الإدارة). act: 'c' تأكيد، 'r' رفض
function decideCore(id, act, note) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) {}
  try {
    const c = findClaim(id);
    if (!c) return 'الإيصال غير موجود';
    if (st(c.r[4]) !== 'pending') return 'تمت معالجة هذا الإيصال من قبل';
    if (act === 'r') {
      c.sh.getRange(c.rowIdx, 5).setValue(ST_AR.rejected);
      c.sh.getRange(c.rowIdx, 8).setValue(String(note || '').slice(0, 200));
      return 'تم رفض الإيصال ✓';
    }
    const apt = nd(c.r[1]), amt = Number(c.r[3]);
    const item = SS.getSheetByName('Items').getDataRange().getValues().slice(1).find(r => ck(r[0]) === ck(c.r[2]));
    const paid = SS.getSheetByName('Payments').getDataRange().getValues().slice(1)
      .filter(r => nd(r[1]) === apt && ck(r[2]) === ck(c.r[2])).reduce((x, y) => x + Number(y[3]), 0);
    if (item && amt > Number(item[3]) - paid + 0.001) return 'المبلغ أكبر من المتبقي على الشقة. عدّله من تاب Claims ثم أكّد من هناك.';
    const res = registerClaimPayment(c.sh, c.rowIdx);
    if (!res.ok) return res.msg;
    c.sh.getRange(c.rowIdx, 5).setValue(ST_AR.confirmed);
    return 'تم التأكيد وتسجيل الدفعة ✓';
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// ---------- صفحة مسؤول الصيانة (maintenance.html) ----------
// دخول برمز خاص (Settings: رمز مسؤول الصيانة، 6 خانات على الأقل) بلا أي بيانات عن السكان أو الصندوق.
// المهام تُضاف من تاب Tasks. المسؤول يحدّث الحالة والتكلفة والملاحظة ويرفع صورة/فاتورة (تُحفظ في الفواتير/السنة، والروابط للأمين فقط).
const T_STATUS = ['جديدة', 'جارية', 'منتهية'];
const T_ALL = T_STATUS.concat(['ملغاة']);
const T_HEAD = ['تاريخ الإضافة', 'المهمة', 'تفاصيل (اختياري)', 'الموعد', 'الحالة (جديدة / جارية / منتهية / ملغاة)', 'التكلفة', 'ملاحظة المسؤول', 'روابط الصور/الفواتير', 'آخر تحديث', 'معرّف (لا تغيّره)'];

function initTasks(sh) {
  sh.getRange(1, 1, 1, T_HEAD.length).setValues([T_HEAD]);
  sh.getRange('E2:E500').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(T_ALL, true).setAllowInvalid(false).build());
  sh.getRange('A2:A500').setNumberFormat('yyyy-mm-dd');
  sh.getRange('D2:D500').setNumberFormat('yyyy-mm-dd');
  sh.getRange('I2:I500').setNumberFormat('yyyy-mm-dd hh:mm');
  try { sh.hideColumns(10); } catch (e) {}
}
function tasksSheet() {
  let sh = SS.getSheetByName('Tasks');
  if (!sh) { sh = SS.insertSheet('Tasks'); initTasks(sh); }
  return sh;
}
function tStatus(v) { const x = String(v || '').trim(); return T_ALL.indexOf(x) >= 0 ? x : 'جديدة'; }
function ensureTaskIds(sh) {
  const n = sh.getLastRow();
  if (n < 2) return;
  const rows = sh.getRange(2, 1, n - 1, 10).getValues();
  let changed = false;
  const ids = rows.map(r => {
    if (String(r[1]).trim() && !String(r[9]).trim()) { changed = true; return [newId()]; }
    return [r[9]];
  });
  if (changed) sh.getRange(2, 10, n - 1, 1).setValues(ids);
}
function mAuth(code) {
  const s = settings();
  const real = nd(s.maint_code);
  if (real.length < 6) return { error: 'off' };
  const props = PropertiesService.getScriptProperties();
  const key = 'f_maint', now = Date.now();
  const f = JSON.parse(props.getProperty(key) || '{"n":0,"t":0}');
  if (f.n >= 5 && now - f.t < 3600000) { logAccess('صيانة', 'locked'); return { error: 'locked' }; }
  if (f.n >= 5) f.n = 0;
  if (nd(code) !== real) {
    props.setProperty(key, JSON.stringify({ n: f.n + 1, t: now }));
    logAccess('صيانة', 'bad_pin');
    return { error: 'bad' };
  }
  props.deleteProperty(key);
  logAccess('صيانة', 'ok');
  return { name: s.maint_name || '' };
}
function mList(b) {
  const a = mAuth(b.code);
  if (a.error) return { ok: false, error: a.error };
  const sh = tasksSheet();
  ensureTaskIds(sh);
  const n = sh.getLastRow();
  const rows = n >= 2 ? sh.getRange(2, 1, n - 1, 10).getValues() : [];
  const all = rows.filter(r => String(r[1]).trim() && tStatus(r[4]) !== 'ملغاة').map(r => {
    const c = r[5] === '' || r[5] == null ? '' : (Number(nd(r[5]).replace(/[٫,]/g, '.')) || 0);
    return {
      id: String(r[9]), title: String(r[1]).trim(), details: String(r[2] || '').trim(), due: fmt(r[3]),
      status: tStatus(r[4]), cost: c, note: String(r[6] || '').trim(),
      files: String(r[7] || '').split('\n').filter(x => x.trim()).length,
      updated: r[8] instanceof Date ? Utilities.formatDate(r[8], TZ, 'yyyy-MM-dd') : ''
    };
  });
  const active = all.filter(t => t.status !== 'منتهية').sort((x, y) => (x.due || '9999').localeCompare(y.due || '9999'));
  const done = all.filter(t => t.status === 'منتهية').sort((x, y) => y.updated.localeCompare(x.updated)).slice(0, 10);
  return { ok: true, name: a.name, tasks: active.concat(done) };
}
function mUpdate(b) {
  const a = mAuth(b.code);
  if (a.error) return { ok: false, error: a.error };
  const status = String(b.status || '').trim();
  if (T_STATUS.indexOf(status) < 0) return { ok: false, error: 'status' };
  let cost = null;
  if (b.cost !== undefined && b.cost !== null && String(b.cost).trim() !== '') {
    cost = Math.round(Number(nd(b.cost).replace(/[٫,]/g, '.')) * 100) / 100;
    if (!isFinite(cost) || cost < 0 || cost > 10000000) return { ok: false, error: 'cost' };
  }
  const note = String(b.note || '').trim().slice(0, 300);
  const mime = String(b.mime || '');
  const ext = { 'image/jpeg': 'jpg', 'image/png': 'png', 'image/webp': 'webp' }[mime];
  if (b.data && (!ext || String(b.data).length > 3000000)) return { ok: false, error: 'file' };

  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) {}
  try {
    const sh = tasksSheet();
    const n = sh.getLastRow();
    const rows = n >= 2 ? sh.getRange(2, 1, n - 1, 10).getValues() : [];
    const i = rows.findIndex(r => String(r[9]) === String(b.id || '') && String(r[9]) !== '');
    if (i < 0 || tStatus(rows[i][4]) === 'ملغاة') return { ok: false, error: 'task' };
    const r = rows[i], rowIdx = i + 2, old = tStatus(r[4]), title = String(r[1]).trim();

    let url = '';
    if (b.data) {
      const now = new Date();
      const year = Number(Utilities.formatDate(now, TZ, 'yyyy'));
      if (!PropertiesService.getScriptProperties().getProperty('ens_' + year)) ensureYear(year);
      const folder = yearFolder('الفواتير', year);
      const base = clean('صيانة - ' + title + ' - ' + Utilities.formatDate(now, TZ, 'yyyy-MM-dd'));
      let name = base + '.' + ext;
      if (folder.getFilesByName(name).hasNext()) name = base + ' - ' + Utilities.formatDate(now, TZ, 'HHmmss') + '.' + ext;
      url = folder.createFile(Utilities.newBlob(Utilities.base64Decode(b.data), mime, name)).getUrl();
    }
    sh.getRange(rowIdx, 5).setValue(status);
    if (cost !== null) sh.getRange(rowIdx, 6).setValue(cost);
    sh.getRange(rowIdx, 7).setValue(note);
    if (url) {
      const cell = sh.getRange(rowIdx, 8);
      cell.setValue(cell.getValue() ? cell.getValue() + '\n' + url : url);
    }
    sh.getRange(rowIdx, 9).setValue(new Date()).setNumberFormat('yyyy-mm-dd hh:mm');

    try {
      MailApp.sendEmail(adminEmail(), 'تحديث مهمة صيانة: ' + title + (status === 'منتهية' ? ' ✅' : ''),
        (a.name || 'مسؤول الصيانة') + ' حدّث المهمة "' + title + '"\n' +
        'الحالة: ' + old + ' ⟵ ' + status + '\n' +
        (cost !== null ? 'التكلفة: ' + cost + ' جنيه\n' : '') +
        (note ? 'ملاحظته: ' + note + '\n' : '') +
        (url ? 'المرفق: ' + url + '\n' : '') +
        (status === 'منتهية' && cost ? '\nلو دُفعت التكلفة من الصندوق سجّلها في تاب Expenses.' : '') +
        '\nالتفاصيل في تاب Tasks.');
    } catch (err) {}
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
  return { ok: true };
}

// ---------- صفحة الإدارة (admin.html) لأمين الصندوق ----------
// كل شيء يُقرأ من الشيت مباشرة ويُكتب فيه، فأي تعديل في أحدهما يظهر في الآخر (الصفحة تتحدّث كل دقيقة وبزر تحديث).
// الدخول برمز Settings: رمز الإدارة (8 خانات على الأقل)، بقفل 5 محاولات/ساعة.
function aAuth(code) {
  const s = settings();
  const real = nd(s.admin_code);
  if (real.length < 8) return { error: 'off' };
  const props = PropertiesService.getScriptProperties();
  const key = 'f_admin', now = Date.now();
  const f = JSON.parse(props.getProperty(key) || '{"n":0,"t":0}');
  if (f.n >= 5 && now - f.t < 3600000) { logAccess('إدارة', 'locked'); return { error: 'locked' }; }
  if (f.n >= 5) f.n = 0;
  if (nd(code) !== real) {
    props.setProperty(key, JSON.stringify({ n: f.n + 1, t: now }));
    logAccess('إدارة', 'bad_pin');
    return { error: 'bad' };
  }
  props.deleteProperty(key);
  logAccess('إدارة', 'ok');
  return {};
}
function aptRows() {
  return SS.getSheetByName('Apartments').getDataRange().getValues().slice(1).filter(r => nd(r[0]));
}
// كل مشاريع (بنود النوع \"مشروع\") مع المطلوب والمحصّل على مستوى العمارة
function projectsInfo() {
  const n = aptRows().length;
  const items = SS.getSheetByName('Items').getDataRange().getValues().slice(1).filter(r => r[0] && String(r[4]).trim() === 'مشروع');
  const pays = SS.getSheetByName('Payments').getDataRange().getValues().slice(1).filter(r => nd(r[1]));
  return items.map(r => {
    const per = Number(r[3]) || 0, req = per * n;
    const col = pays.filter(p => ck(p[2]) === ck(r[0])).reduce((t, p) => t + (Number(p[3]) || 0), 0);
    return { code: String(r[0]), desc: String(r[1] || ''), per: per, required: req, collected: col,
      due: fmt(r[5] || r[2]), open: col < req - 0.001 };
  });
}
function openProjects() {
  const open = projectsInfo().filter(p => p.open);
  return open.length ? open.slice(0, 3).map(p => ({ desc: p.desc, required: p.required, collected: p.collected })) : null;
}
function aList(b) {
  const a = aAuth(b.code);
  if (a.error) return { ok: false, error: a.error };
  const s = settings();
  const apts = aptRows();
  const items = SS.getSheetByName('Items').getDataRange().getValues().slice(1).filter(r => r[0]);
  const pays = SS.getSheetByName('Payments').getDataRange().getValues().slice(1).filter(r => nd(r[1]));
  const byApt = {};
  pays.forEach(p => { const k = nd(p[1]); (byApt[k] = byApt[k] || {})[ck(p[2])] = ((byApt[k] || {})[ck(p[2])] || 0) + (Number(p[3]) || 0); });
  // لكل شقة: المطلوب = البنود غير المكتملة فقط
  const late = apts.map(r => {
    const apt = nd(r[0]), m = byApt[apt] || {};
    let req = 0, paid = 0;
    items.forEach(it => {
      const amt = Number(it[3]) || 0, p = m[ck(it[0])] || 0;
      if (amt > 0 && p < amt) { req += amt; paid += p; }
    });
    return { apt: apt, name: String(r[1] || ''), phone: intl(r[2]), required: req, paid: paid, remaining: Math.round((req - paid) * 100) / 100 };
  }).filter(x => x.remaining > 0.001).sort((x, y) => x.apt.localeCompare(y.apt, 'en', { numeric: true }));

  const pj = projectsInfo();
  const open = pj.filter(p => p.open);
  const projects = (open.length ? open : pj.slice(-1)).map(p => Object.assign({}, p, { remaining: Math.max(p.required - p.collected, 0) }));

  const names = {}; items.forEach(r => { names[ck(r[0])] = String(r[1] || ''); });
  const aptName = {}; apts.forEach(r => { aptName[nd(r[0])] = String(r[1] || ''); });
  const claims = [];
  const cs = SS.getSheetByName('Claims');
  if (cs) {
    const rows = cs.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (st(r[4]) !== 'pending') continue;
      let id = String(r[8] || '');
      if (!id) { id = newId(); cs.getRange(i + 1, 9).setValue(id); }
      claims.push({ id: id, apt: nd(r[1]), name: aptName[nd(r[1])] || '', desc: names[ck(r[2])] || String(r[2]),
        amount: Number(r[3]) || 0, date: fmt(r[0]), files: String(r[5] || '').split('\n').filter(x => x.trim()) });
    }
  }
  return {
    ok: true,
    fund: Number(SS.getSheetByName('Summary').getRange('H2').getValue()) || 0,
    show: { fund: yes(s.show_fund), project: yes(s.show_project), quote: yes(s.show_quote) },
    announcement: s.announcement || '',
    pageUrl: String(SS.getSheetByName('Summary').getRange('J2').getValue() || '').trim(),
    aptCount: apts.length,
    projects: projects,
    late: late,
    claims: claims
  };
}
function aToggle(b) {
  const a = aAuth(b.code);
  if (a.error) return { ok: false, error: a.error };
  const key = String(b.key || '');
  if (['show_fund', 'show_project', 'show_quote'].indexOf(key) < 0) return { ok: false, error: 'key' };
  setSetting(key, b.on ? 'نعم' : 'لا');
  return { ok: true };
}
function aAnn(b) {
  const a = aAuth(b.code);
  if (a.error) return { ok: false, error: a.error };
  setSetting('announcement', String(b.text == null ? '' : b.text).trim().slice(0, 500));
  return { ok: true };
}
function aProject(b) {
  const a = aAuth(b.code);
  if (a.error) return { ok: false, error: a.error };
  const desc = clean(String(b.desc || ''));
  if (!desc) return { ok: false, error: 'desc' };
  const per = Math.round(Number(nd(b.amount).replace(/[٫,]/g, '.')) * 100) / 100;
  if (!isFinite(per) || per <= 0 || per > 10000000) return { ok: false, error: 'amount' };
  let due = '';
  if (String(b.due || '').trim()) {
    const m = String(b.due).trim().match(/^(\d{4})-(\d{2})-(\d{2})$/);
    due = m ? new Date(+m[1], +m[2] - 1, +m[3]) : null;
    if (!due || due.getFullYear() !== +m[1] || due.getMonth() !== +m[2] - 1 || due.getDate() !== +m[3]) return { ok: false, error: 'date' };
  }
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) {}
  try {
    const sh = SS.getSheetByName('Items');
    let max = 0;
    sh.getDataRange().getValues().slice(1).forEach(r => { const m = String(r[0]).trim().match(/^P(\d+)$/i); if (m) max = Math.max(max, +m[1]); });
    const code = 'P' + (max + 1);
    sh.appendRow([code, desc, new Date(), per, 'مشروع', due]);
    const row = sh.getLastRow();
    sh.getRange(row, 3).setNumberFormat('yyyy-mm-dd');
    sh.getRange(row, 6).setNumberFormat('yyyy-mm-dd');
    return { ok: true, code: code };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}
function aDecide(b) {
  const a = aAuth(b.code);
  if (a.error) return { ok: false, error: a.error };
  if (['c', 'r'].indexOf(b.act) < 0 || !/^[a-f0-9]{12}$/.test(String(b.id || ''))) return { ok: false, error: 'bad' };
  return { ok: true, msg: decideCore(String(b.id), b.act, b.note) };
}

// ---------- الأمان والاسترجاع ----------
// اختبار أمني آلي: شغّلها يدوياً من المحرر وشوف النتيجة في Execution log. لا تغيّر بيانات (تمسح سطور AccessLog الخاصة بها).
function securityTest() {
  const out = [];
  const t = (name, ok) => out.push((ok ? '✅ ' : '❌ ') + name);
  CTX = { did: 'selftest', act: 'view' };
  const props = PropertiesService.getScriptProperties();
  const aptRow = SS.getSheetByName('Apartments').getDataRange().getValues().slice(1).find(r => r[0] && r[3]);
  const item = SS.getSheetByName('Items').getDataRange().getValues().slice(1).find(r => r[0]);

  t('شقة غير موجودة ⟵ رفض', check('zz-none', '0000').error === 'bad');
  t('رمز تأكيد مزوّر ⟵ مرفوض', readToken('aaaaaaaaaaaa.c.9999999999999.xxxx') === null && decide('garbage', '') === 'الرابط غير صالح أو منتهي');
  const good = makeToken('abcdef012345', 'c');
  const parts = good.split('.');
  t('رمز صحيح ⟵ مقبول', !!readToken(good));
  t('تغيير الإجراء من تأكيد لرفض ⟵ مرفوض', readToken(parts[0] + '.r.' + parts[2] + '.' + parts[3]) === null);
  const oldExp = Date.now() - 1000, msg = 'abcdef012345.c.' + oldExp;
  t('رمز منتهي ⟵ مرفوض', readToken(msg + '.' + sign(msg)) === null);

  t('جملة اليوم: القائمة كاملة والجملة غير فارغة', QUOTES.length >= 100 && QUOTES.every(q => q[0] && q[1]) && !!todayQuote().t);

  const mc = nd(settings().maint_code);
  if (mc.length < 6) {
    t('صيانة بلا رمز مفعّل ⟵ مغلقة', mList({ code: 'x' }).error === 'off');
  } else {
    const mf = props.getProperty('f_maint');
    t('رمز صيانة خاطئ ⟵ رفض', mList({ code: mc + 'x' }).error === 'bad');
    t('تحديث بحالة غير مسموحة ⟵ رفض', mUpdate({ code: mc, id: 'x', status: 'ملغاة' }).error === 'status');
    t('تحديث مهمة غير موجودة ⟵ رفض', mUpdate({ code: mc, id: 'no-such-id', status: 'جارية' }).error === 'task');
    t('تكلفة سالبة ⟵ رفض', mUpdate({ code: mc, id: 'x', status: 'جارية', cost: -1 }).error === 'cost');
    t('نوع ملف غير مسموح في الصيانة ⟵ رفض', mUpdate({ code: mc, id: 'x', status: 'جارية', mime: 'application/pdf', data: 'AAAA' }).error === 'file');
    t('رمز صيانة صحيح ⟵ نجاح', mList({ code: mc }).ok === true);
    props.deleteProperty('f_maint');
  }

  const ac = nd(settings().admin_code);
  if (ac.length < 8) {
    t('صفحة الإدارة بلا رمز مفعّل ⟵ مقفولة', aList({ code: 'x' }).error === 'off');
  } else {
    t('رمز إدارة خاطئ ⟵ رفض', aList({ code: ac + 'x' }).error === 'bad');
    t('مفتاح إعداد غير مسموح ⟵ رفض', aToggle({ code: ac, key: 'admin_code', on: true }).error === 'key');
    t('مشروع بلا وصف ⟵ رفض', aProject({ code: ac, desc: '', amount: 5 }).error === 'desc');
    t('مشروع بمبلغ سالب ⟵ رفض', aProject({ code: ac, desc: 'x', amount: -5 }).error === 'amount');
    t('مشروع بتاريخ غير صحيح ⟵ رفض', aProject({ code: ac, desc: 'x', amount: 5, due: '2026-13-45' }).error === 'date');
    t('قرار على معرّف مزوّر ⟵ رفض', aDecide({ code: ac, id: 'zz', act: 'c' }).error === 'bad');
    t('رمز إدارة صحيح ⟵ نجاح', aList({ code: ac }).ok === true);
    props.deleteProperty('f_admin');
  }

  if (aptRow && item) {
    const apt = nd(aptRow[0]), pin = pinNorm(aptRow[3]);
    const wrong = pin === '0000' ? '1111' : '0000';
    const f = JSON.parse(props.getProperty('f_' + apt) || '{"n":0,"t":0}');
    if (f.n >= 5 && Date.now() - f.t < 3600000) {
      out.push('⚠️ شقة ' + apt + ' مقفولة حالياً، تخطّيت اختبارات الشقة');
    } else {
      t('رقم سري خاطئ ⟵ رفض', check(apt, wrong).error === 'bad');
      t('رقم سري صحيح ⟵ نجاح', check(apt, pin).ok === true);
      const base = { apt: apt, pin: pin, code: item[0], mime: 'image/jpeg', data: 'AAAA' };
      t('رفع بدون رقم سري صحيح ⟵ رفض', upload(Object.assign({}, base, { pin: wrong })).ok === false);
      t('نوع ملف غير مسموح ⟵ رفض', upload(Object.assign({}, base, { mime: 'application/pdf' })).error === 'file');
      t('ملف ضخم ⟵ رفض', upload(Object.assign({}, base, { data: new Array(3000002).join('A') })).error === 'file');
      t('بند غير موجود ⟵ رفض', upload(Object.assign({}, base, { code: 'NO-SUCH-ITEM' })).error === 'item');
      t('مبلغ أكبر من المتبقي ⟵ رفض', ['amount', 'done'].indexOf(upload(Object.assign({}, base, { amount: 99999999 })).error) >= 0);
      t('مبلغ سالب أو صفر ⟵ رفض', ['amount', 'done'].indexOf(upload(Object.assign({}, base, { amount: -5 })).error) >= 0);
    }
    props.deleteProperty('f_' + apt);
  }

  // تنظيف سطور الاختبار من AccessLog
  const al = SS.getSheetByName('AccessLog');
  if (al) for (let i = al.getLastRow(); i >= 2; i--) if (String(al.getRange(i, 4).getValue()) === 'selftest') al.deleteRow(i);
  props.deleteProperty('alert_selftest');

  Logger.log(out.join('\n'));
  SS.toast(out.filter(x => x.indexOf('❌') === 0).length ? 'فيه اختبارات فشلت، راجع Execution log' : 'كل الاختبارات نجحت ✅', 'نتيجة الاختبار الأمني');
}

// حماية الخلايا الحساسة (تحذير عند التعديل، ولا تمنع السكريبت). تُشغَّل مرة، وتُعاد بأمان.
function protectSheets() {
  SS.getSheets().forEach(sh => sh.getProtections(SpreadsheetApp.ProtectionType.RANGE)
    .concat(sh.getProtections(SpreadsheetApp.ProtectionType.SHEET))
    .forEach(p => { if (String(p.getDescription()).indexOf('bf:') === 0) p.remove(); }));
  const lock = (name, a1, unprot) => {
    const sh = SS.getSheetByName(name);
    if (!sh) return;
    const p = (a1 ? sh.getRange(a1) : sh.getDataRange()).protect().setDescription('bf:' + name + ':' + (a1 || 'all')).setWarningOnly(true);
    if (unprot) p.setUnprotectedRanges(unprot.map(r => sh.getRange(r)));
  };
  const whole = name => { const sh = SS.getSheetByName(name); if (sh) sh.protect().setDescription('bf:' + name).setWarningOnly(true); };
  const wholeExcept = (name, unprot) => { const sh = SS.getSheetByName(name); if (sh) sh.protect().setDescription('bf:' + name).setWarningOnly(true).setUnprotectedRanges(unprot.map(r => sh.getRange(r))); };
  wholeExcept('Summary', ['J2']);
  whole('AccessLog');
  lock('Payments', 'A1:E1'); lock('Payments', 'E2:E999');
  lock('Apartments', 'A1:D1'); lock('Items', 'A1:F1'); lock('Expenses', 'A1:E1');
  lock('Claims', 'A1:I1'); lock('Claims', 'A2:C500'); lock('Claims', 'F2:G500'); lock('Claims', 'I2:I500');
  lock('Tasks', 'A1:J1'); lock('Tasks', 'I2:J500');
  lock('Settings', 'A1:C500');
  SS.toast('تمت حماية الخلايا الحساسة (تحذير عند تعديلها)', 'اكتمل');
}

function summaryF(r) {
  return '=IF(A' + r + '="","",IF(D' + r + '>0,HYPERLINK("https://wa.me/"&Apartments!C' + r + '&"?text="&ENCODEURL("السلام عليكم، شقة "&A' + r + '&". نشكر لكم مساهمتكم في صندوق عمارة 220 حي القرنفل، فهي استثمار في راحة الجميع وسلامة المبنى. يتبقى من إسهامكم "&D' + r + '&" جنيه، ويسعدنا متابعتكم للتفاصيل والمساهمة عبر: "&$J$2),"ارسل تنبيه"),"")';
}
// أعمدة Summary (A..F) لصف r. المطلوب/المدفوع/المتبقي تحسب البنود غير المكتملة فقط؛ البند المكتمل لا يدخل فيها
function summaryCols(r) {
  const P = 'SUMIFS(Payments!$D$2:$D$5000,Payments!$B$2:$B$5000,$A' + r + ',Payments!$C$2:$C$5000,Items!$A$2:$A$500)';
  const M = '(Items!$A$2:$A$500<>"")*(' + P + '<Items!$D$2:$D$500)';
  return [
    '=IF(Apartments!A' + r + '="","",Apartments!A' + r + ')',
    '=IF(A' + r + '="","",SUMPRODUCT(' + M + '*Items!$D$2:$D$500))',
    '=IF(A' + r + '="","",B' + r + '-D' + r + ')',
    '=IF(A' + r + '="","",SUMPRODUCT(' + M + '*(Items!$D$2:$D$500-' + P + ')))',
    '=IF(A' + r + '="","",IF(D' + r + '<=0,"مدفوع ✅","عليه "&D' + r + '))',
    summaryF(r)];
}
// تعيد العناوين والمعادلات والقوائم المنسدلة لو اتخربت. لا تمس بيانات السكان أو المدفوعات.
function resetSheets() {
  const H = {
    Apartments: ['رقم الشقة', 'الاسم', 'الموبايل (بصيغة 201012345678)', 'الرقم السري'],
    Items: ['كود', 'الوصف', 'التاريخ', 'المبلغ لكل شقة', 'النوع (مشروع / صندوق)', 'تاريخ الاستحقاق'],
    Payments: ['التاريخ', 'رقم الشقة', 'كود البند', 'المبلغ', 'النوع (تلقائي)'],
    Expenses: ['التاريخ', 'الوصف', 'المبلغ', 'التفاصيل (تظهر للسكان لو إظهار تفاصيل المصروفات = نعم)', 'رابط الفاتورة (خاص بك)'],
    Summary: ['رقم الشقة', 'المطلوب', 'المدفوع', 'المتبقي', 'الحالة', 'تنبيه واتساب'],
    Settings: ['الإعداد (لا تغيّر الاسم)', 'القيمة', 'ملاحظة'],
    Claims: ['التاريخ', 'الشقة', 'كود البند', 'المبلغ', 'الحالة (قيد المراجعة / تم التأكيد / مرفوض)', 'رابط الإيصال', 'تاريخ التسجيل في Payments', 'ملاحظة الرفض (اختياري)', 'معرّف (لا تغيّره)'],
    Tasks: T_HEAD,
    AccessLog: ['الوقت', 'الشقة', 'النتيجة (نجح / رقم سري خاطئ / شقة غير موجودة / مقفول)', 'معرّف الجهاز', 'العملية']
  };
  Object.keys(H).forEach(n => {
    const sh = SS.getSheetByName(n) || SS.insertSheet(n);
    sh.getRange(1, 1, 1, H[n].length).setValues([H[n]]);
  });

  const ap = SS.getSheetByName('Apartments'); ap.getRange('C2:D200').setNumberFormat('@');
  const it = SS.getSheetByName('Items');
  it.getRange('E2:E200').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['مشروع', 'صندوق'], true).setAllowInvalid(true).build());

  const pay = SS.getSheetByName('Payments');
  const pf = []; for (let r = 2; r <= 999; r++) pf.push(['=IFERROR(VLOOKUP(C' + r + ',Items!A:E,5,0),"")']);
  pay.getRange(2, 5, 998, 1).setFormulas(pf);

  const su = SS.getSheetByName('Summary');
  const sf = []; for (let r = 2; r <= 31; r++) sf.push(summaryCols(r));
  su.getRange(2, 1, 30, 6).setFormulas(sf);
  su.getRange('H1').setValue('رصيد صندوق الصيانة');
  su.getRange('H2').setFormula('=SUMIF(Payments!E:E,"صندوق",Payments!D:D)-SUM(Expenses!C:C)');
  su.getRange('J1').setValue('رابط الصفحة');

  const YN = SpreadsheetApp.newDataValidation().requireValueInList(['نعم', 'لا'], true).setAllowInvalid(false).build();
  const PT = SpreadsheetApp.newDataValidation().requireValueInList(['إنستا باي', 'محفظة إلكترونية'], true).setAllowInvalid(true).build();
  const se = SS.getSheetByName('Settings');
  se.getDataRange().getValues().forEach((r, i) => {
    const k = skey(r[0]);
    if (/^show_/.test(k)) se.getRange(i + 1, 2).setDataValidation(YN);
    if (k === 'pay1_type' || k === 'pay2_type') se.getRange(i + 1, 2).setDataValidation(PT);
  });
  se.getRange(2, 2, Math.max(se.getLastRow() - 1, 1), 1).setNumberFormat('@');

  initClaims(SS.getSheetByName('Claims'));
  initTasks(SS.getSheetByName('Tasks'));
  SS.toast('تمت إعادة العناوين والمعادلات والقوائم. (الألوان الشرطية لا تُعاد)', 'اكتمل');
}

// نسخة احتياطية من الشيت إلى Drive (حساب العمارة/نسخ احتياطية) ويحتفظ بآخر 8 نسخ.
// ملاحظة: النسخة تحوي الأرقام السرية، والمجلد خاص بك ولا يُشارك.
function backupSheet() {
  const folder = getOrCreate(rootFolder(), 'نسخ احتياطية');
  DriveApp.getFileById(SS.getId()).makeCopy('نسخة احتياطية - ' + Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'), folder);
  const files = [];
  const it = folder.getFiles();
  while (it.hasNext()) files.push(it.next());
  files.sort((x, y) => y.getDateCreated() - x.getDateCreated());
  files.slice(8).forEach(f => f.setTrashed(true));
}

// ---------- التقارير والجدولة ----------
function overdueReport() {
  const apts = SS.getSheetByName('Apartments').getDataRange().getValues().slice(1).filter(r => r[0]);
  const items = SS.getSheetByName('Items').getDataRange().getValues().slice(1).filter(r => r[0]);
  const pays = SS.getSheetByName('Payments').getDataRange().getValues().slice(1);
  const today = new Date(); today.setHours(0, 0, 0, 0);

  const lines = [];
  let total = 0;
  apts.forEach(a => {
    let owed = 0;
    const late = [];
    items.forEach(it => {
      const due = toDate(it[5] || it[2]);
      if (!due || due >= today) return;
      const paid = pays.filter(x => nd(x[1]) === nd(a[0]) && ck(x[2]) === ck(it[0]))
        .reduce((t, x) => t + Number(x[3]), 0);
      const rest = Number(it[3]) - paid;
      if (rest > 0) { owed += rest; late.push(it[1] + ' (' + rest + ')'); }
    });
    if (owed > 0) { total += owed; lines.push('شقة ' + a[0] + ' - ' + a[1] + ': إسهام متبقي ' + owed + ' جنيه - ' + late.join('، ')); }
  });

  if (!lines.length) return;
  MailApp.sendEmail(adminEmail(), 'إسهامات متأخرة (' + lines.length + ' شقة)',
    lines.join('\n') + '\n\nإجمالي الإسهامات المتأخرة: ' + total + ' جنيه');
}

// شغّلها مرة واحدة يدوياً: تقرير يوم 5 كل شهر + مجلدات أول كل شهر + نسخة احتياطية أسبوعية (الأحد) + إنشاء مجلدات Drive
function installTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => ['overdueReport', 'monthlyFolders', 'backupSheet'].indexOf(t.getHandlerFunction()) >= 0)
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('overdueReport').timeBased().onMonthDay(5).atHour(9).create();
  ScriptApp.newTrigger('monthlyFolders').timeBased().onMonthDay(1).atHour(1).create();
  ScriptApp.newTrigger('backupSheet').timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(3).create();
  setupFolders();
}
