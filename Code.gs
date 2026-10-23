// v3.8: + حقل «حارس ليلي» للموظف (عمود جديد بالافتراضي لا) مع أوقات الحارس الليلي ونظام البديل في إعدادات الدوام (work.night*). شغّل setup مرة واحدة بعد اللصق (آمن) ثم أعد النشر.
// v3.7: + «ربط البصمات» (ورقة FpMap): رقم البصمة لكل موظف حسب الجهاز (أجهزة متعددة بأرقام مختلفة) مع اتجاه الجهاز (دخول/خروج/كلاهما) ونافذة وقت اختيارية (للحراس) + مفتاح الحركة صار (الرقم+الوقت+الجهاز) فلا تتصادم أرقام الأجهزة المختلفة + ربط بصمة غير مطابقة يتم لجهاز محدد + إلغاء انتهاء الجلسة بالخمول كليًا (الجلسات تُحفظ أيضًا في خصائص السكربت فلا تضيع بتفريغ الذاكرة المؤقتة) + عمود «اتجاه البصمة» للجهاز. شغّل setup مرة واحدة بعد اللصق (آمن) ثم أعد النشر.
// v3.6: + نقل نوافذ وصلاحيات الموارد البشرية لمسؤول البصمة (تبقى لمسؤول الموارد أيضًا) + حقل «نوع العقد» للموظف (عمود جديد: شغّل setup مرة واحدة ثم أعد النشر)
// ===== نظام البصمة - Code.gs v3.5 : جوجل شيت هو المصدر الوحيد للبيانات =====
// v3.5: + النسخ الاحتياطي ملف Excel (.xlsx) يوميًا عند منتصف الليل (12:00 ص) مع تنزيله من الواجهة (إجراء bkdl) + إلغاء الملخص البريدي كليًا (dailyDigest/MailApp) + إلغاء فحص السلامة والأرشفة (dbstats/archive/unarchive) + رقم بصمة ثانٍ للموظف (fpId2) يُطابَق ويُمنع تكراره كالأول + عمود «الساعات» للحالات (الإجازة الزمنية: كل 6 ساعات = يوم) + نوع حركة رصيد «خصم زمنيات». شغّل setup مرة واحدة بعد اللصق (آمن: يضيف أعمدة ولا يمسّ البيانات) ثم أعد النشر. ملاحظة: التصدير إلى Excel يستخدم UrlFetchApp فستُطلب صلاحية جديدة عند أول تشغيل.
// v3.4 (المحور 3 - الحضور وحركات البصمة + المحور 4 - الموظفون والبيانات الوظيفية): + سجل مركزي للحركات (طريقة الاستيراد/وقت الاستلام) + سجل مستبعد (PunchLog) للمكرر والمحذوف + ربط بصمة غير مطابقة بموظف بإجراء موثّق (linkfp) + تصحيحات حضور لا تمسّ الحركة الأصلية وتحتاج اعتمادًا (Corrections) + فصل الطلب عن الاعتماد + سجل انتقالات الموظف (Movements) + حالة الخدمة وتواريخها + منع تكرار رقم الموظف/البصمة على الخادم + رقم الأمر الإداري للحالات. أعد تشغيل setup مرة واحدة (آمن: يضيف أوراقًا وأعمدة ولا يمسّ البيانات).
// v3.3 (المحور 2 - الحسابات والصلاحيات، مكتمل): + انتهاء صلاحية كلمة السر (Settings: set.pwDays) + منع إعادة استخدام آخر 5 كلمات سر + تخزين hash مغلّف بمفتاح الخادم (نسخ Drive لا تصلح للدخول) + تعطيل مؤقت بمدة + إنهاء جلسات مستخدم. السابق: صلاحيات تفصيلية تُفرض على الخادم عند الحفظ (لا يكفي إخفاء الأزرار)، منع تسريب hash كلمات السر لغير صاحبها، إبطال الجلسات عند تغيير كلمة السر/التعطيل/تغيير الصفة، انتهاء الجلسة بالخمول، إجراء logout حقيقي، قفل تدريجي، حماية آخر أدمن، وسجل خادم موثوق للعمليات الحساسة. أعد تشغيل setup مرة واحدة (يضيف 4 أعمدة لورقة Users ولا يمسّ البيانات).
// v3.2 (المحور 1 - لوحة المعلومات): تسجيل حالة النسخ الاحتياطي (نجاح/فشل) وإرسالها للوحة، وعمود «وقت القرار» للإجازات، وعمود «تاريخ الحل» للأخطاء. أعد تشغيل setup مرة واحدة لتحديث رؤوس الأوراق (آمن، لا يمسّ البيانات).
// v3.1: الجسر (fp_bridge.py v4) يقرأ قائمة الأجهزة المفعّلة من هنا ويرسل حالة كل سحب (تظهر في تبويب الأجهزة).
// 1) الصق الكود مكان القديم  2) شغّل setup مرة واحدة (ينقل بياناتك القديمة تلقائيًا ويطبع مفتاح الجسر)
// 3) نشر > إدارة عمليات النشر > تعديل > إصدار جديد (نفس الرابط). الوصول: أي شخص
// تسجيل الدخول وكلمات السر والصلاحيات تُفحص هنا من ورقة Users، لا من المتصفح.

var SHEET_ID  = '15xpNSB6Gj44a60Ab2CB9kq5tdF4mco4FmYhhSb8wJDM';
var FOLDER_ID = '1maRDd_35ls3emglzFQlPBGMFJK0AJYIb';
var TZ = 'Asia/Baghdad', P_ = PropertiesService.getScriptProperties(), C_ = CacheService.getScriptCache();
function ss_() { return SpreadsheetApp.openById(SHEET_ID); }
function folder_(n) {
  var root = DriveApp.getFolderById(FOLDER_ID), it = root.getFoldersByName(n);
  return it.hasNext() ? it.next() : root.createFolder(n);
}

// رؤوس الأوراق + ترميز الحقول (n=رقم، a=فعال/معطل، y=نعم/لا)
var HD = {
  Users:     ['المعرف','الصفة','الاسم','رقم الموظف','الحالة','تاريخ الإنشاء','salt (لا تعدّله)','hash (لا تعدّله)','تغيير إلزامي','آخر دخول','آخر تغيير كلمة سر','أنشأه','صلاحيات إضافية','تعطيل حتى'],
  Employees: ['رقم الموظف','الاسم','القسم','العنوان الوظيفي','رصيد الإجازات','رقم البصمة','حالة الخدمة','تاريخ المباشرة','تاريخ انتهاء الخدمة','id القسم','رقم البصمة الثانية','نوع العقد','حارس ليلي','رقم البصمة الثالثة'],
  Holidays:  ['التاريخ','النوع','المناسبة'],
  Cases:     ['id','رقم الموظف','الحالة','من','إلى','ملاحظة','رقم الأمر/المستند','أنشأه','حالة الاعتماد','الساعات'],
  Leaves:    ['id','رقم الموظف','من','إلى','السبب','الحالة','وقت الطلب','وقت القرار','id النوع','الأيام المحتسبة','ملاحظات','المستند','المرحلة','المراجع','وقت المراجعة','ملاحظة المراجعة','المعتمد','ملاحظة القرار'],
  Overtime:  ['id','رقم الموظف','التاريخ','الساعات','السبب','أمر التكليف/المستند','سجّله','الحالة','صاحب القرار','وقت التسجيل','وقت القرار','ملاحظة القرار','الساعات المعتمدة'],
  Devices:   ['id','الاسم','IP','المنفذ','الموديل','الطريقة','يعمل','الموقع','القسم','الرقم التسلسلي','إصدار البرنامج','حد الصمت (ساعة)','مؤرشف','ملاحظات','اتجاه البصمة'],
  FpMap:     ['id','رقم الموظف','الجهاز','رقم البصمة','الاتجاه','من وقت','إلى وقت','ملاحظة'],
  Settings:  ['المفتاح','القيمة'],
  Logs:      ['الوقت','المستخدم','الصفة','الحركة'],
  Errors:    ['الوقت','المستخدم','الخطأ','تاريخ الحل'],
  Punches:   ['fpId','datetime','device','طريقة الاستيراد','وقت الاستلام'],
  Movements: ['id','رقم الموظف','نوع التغيير','من','إلى','تاريخ السريان','رقم الأمر','السبب','سجّله','وقت التسجيل'],
  Corrections: ['id','رقم الموظف','التاريخ','دخول معتمد','خروج معتمد','السبب','المستند','طالب التصحيح','الحالة','صاحب القرار','وقت الطلب','وقت القرار','ملاحظة القرار'],
  PunchLog:  ['الوقت','fpId','datetime','device','السبب','المصدر'],
  Schedules: ['id','اسم الجدول','بداية الدوام','نهاية الدوام','سماح التأخير (د)','سماح الخروج (د)','أيام الدوام','فعّال'],
  SchedAssign: ['id','رقم الموظف','id الجدول','يسري من','سجّله'],
  Audit:     ['الوقت','المستخدم','الصفة','النوع','الجدول','المعرّف','الحقل','قبل','بعد','ملاحظة','البصمة المتسلسلة'],
  LogArchive: ['الوقت','المستخدم','الصفة','الحركة'],
  BackupLog: ['الوقت','النوع','اسم النسخة','id الملف','الصفوف','الحالة','ملاحظة','المنفّذ','id ملف Excel'],
  ReportLog: ['id','المرجع','النوع','العنوان','المعايير','الفترة','عدد الصفوف','بصمة النتيجة','أصدره','وقت الإصدار'],
  Departments: ['id','اسم القسم','فعّال'],
  DevLog:    ['الوقت','الجهاز','الحدث','النتيجة','التفاصيل','المنفّذ','المدة (ms)','المقروء','الجديد','المكرر','المرفوض'],
  PunchesArchive: ['fpId','datetime','device','طريقة الاستيراد','وقت الاستلام','تاريخ الأرشفة'],
  LeaveTypes: ['id','النوع','نوع الحالة المقابل','يخصم من الرصيد','احتساب الأيام','يتطلب مستندًا','مراحل الاعتماد','يسمح بأثر رجعي','فعّال'],
  LeaveLedger: ['id','رقم الموظف','التاريخ','نوع الحركة','الكمية (±)','id الإجازة','البيان','سجّله','الوقت'],
  MonthClose: ['الشهر','رقم الموظف','حضور','تأخير','دقائق التأخير','خروج مبكر','دقائق الخروج المبكر','غياب','حالات','بصمة ناقصة','أغلقه','وقت الإغلاق']
};
var SC = {
  users:     ['Users', 'id,role,name,empNo,active:a,created,salt,hash,must:y,lastLogin,pwAt,by,perms,until'],
  employees: ['Employees', 'empNo,name,dept,title,balance:n,fpId,svc,start,end,did:n,fpId2,contract,night:y,fpId3'],
  holidays:  ['Holidays', 'date,type,name'],
  cases:     ['Cases', 'id:n,empNo,type,from,to,note,ref,by,ast,hrs:n'],
  leaves:    ['Leaves', 'id:n,empNo,from,to,reason,status,ts,dts,ltype,days:n,notes,doc,stage,rev,rdts,rnote,appr,anote'],
  ltypes:    ['LeaveTypes', 'id:n,name,ctype,deduct:y,count,needDoc:y,stages:n,retro:y,on:y'],
  ledger:    ['LeaveLedger', 'id:n,empNo,date,kind,delta:n,leaveId,note,by,at'],
  overtime:  ['Overtime', 'id:n,empNo,date,hours:n,reason,ref,by,status,appr,ts,dts,note,ahours:n'],
  devices:   ['Devices', 'id:n,name,ip,port,model,method,on:y,loc,dept,serial,fw,silent:n,arch:y,note,dir'],
  fpmap:     ['FpMap', 'id:n,empNo,dev,fpId,dir,from,to,note'],
  depts:     ['Departments', 'id:n,name,on:y'],
  rlog:      ['ReportLog', 'id:n,ref,type,title,params,period,rows:n,hash,by,at'],
  movements: ['Movements', 'id:n,empNo,kind,from,to,eff,ref,reason,by,at'],
  corrections: ['Corrections', 'id:n,empNo,date,in,out,reason,doc,by,status,appr,ts,dts,note'],
  schedules: ['Schedules', 'id:n,name,start,end,grace:n,egrace:n,days,on:y'],
  sassign:   ['SchedAssign', 'id:n,empNo,sid:n,from,by'],
  closes:    ['MonthClose', 'month,empNo,present:n,late:n,lateMin:n,early:n,earlyMin:n,absent:n,cases:n,incomplete:n,by,at']
};

function setup() {
  try { installTriggers(); } catch (x) { /* يُثبَّت يدويًا عبر installTriggers إن لم تُمنح الصلاحية */ }
  var ss = ss_(), old = null, dbs = ss.getSheetByName('DB');
  if (dbs && dbs.getLastRow() > 1) { // نسخة النظام القديمة (v2) تُنقل إلى الأوراق الحقيقية
    try { old = JSON.parse(dbs.getRange(2, 1, dbs.getLastRow() - 1, 1).getValues().map(function (r) { return r[0]; }).join('')); } catch (x) {}
  }
  Object.keys(HD).forEach(function (n) {
    var sh = ss.getSheetByName(n) || ss.insertSheet(n), h = HD[n];
    if (old && n !== 'Punches') sh.clear();
    sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground('#0f766e').setFontColor('#ffffff');
    sh.setFrozenRows(1); sh.setRightToLeft(true);
    P_.deleteProperty('H_' + n);
  });
  ss.getSheetByName('Punches').getRange('A:B').setNumberFormat('@'); ss.getSheetByName('Punches').getRange('E:E').setNumberFormat('@');
  ss.getSheetByName('PunchLog').getRange('A:F').setNumberFormat('@');
  if (old) { write_(old); dbs.setName('DB_old'); }
  var def = ss.getSheetByName('Sheet1') || ss.getSheetByName('Sheet 1') || ss.getSheetByName('ورقة1');
  if (def && ss.getSheets().length > 1) ss.deleteSheet(def);
  ['Backups', 'Reports', 'Uploads'].forEach(folder_);
  seedDevices_();
  P_.setProperty('INIT', '1');
  bump_(); Logger.log('تم الإعداد' + (old ? ' ونُقلت البيانات القديمة' : ''));
  showKey();
}

function seedDevices_() { // أجهزة افتراضية عند فراغ التبويب (تطابق fp_config.json): عدّلها أو احذفها من الواجهة
  var d = ss_().getSheetByName('Devices');
  if (!d || d.getLastRow() > 1) return;
  var rows = [1, 2, 3].map(function (n) { return [n, 'FaceID-' + n, '192.168.1.20' + n, 4370, 'ZKTeco', 'وسيط بايثون', 'نعم', '', '', '', '', 0, 'نعم', '', '']; });
  setTxt_(d, 2, rows);
}

function ensure_() { // تهيئة تلقائية عند أول طلب على جدول جديد فارغ
  if (P_.getProperty('INIT')) return;
  if (!ss_().getSheetByName('Users')) setup();
  P_.setProperty('INIT', '1');
}

function key_() {
  var k = P_.getProperty('API_KEY');
  if (!k) { k = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, ''); P_.setProperty('API_KEY', k); }
  return k;
}
function showKey() { Logger.log('مفتاح الجسر (لـ fp_bridge.py فقط، ضعه في fp_key.txt): ' + key_()); }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function sh_(n) { var s = ss_(); return s.getSheetByName(n) || s.insertSheet(n); }
function nid_(s) { return String(s == null ? '' : s).trim().replace(/^0+(?=\d)/, ''); }
function num_(k) { return Number(P_.getProperty(k)) || 0; }
function bkSt_() { return { ok: P_.getProperty('BK_OK') || '', err: P_.getProperty('BK_ERR') || '' }; } // حالة النسخ الاحتياطي للوحة المعلومات
function brSt_() { try { return JSON.parse(P_.getProperty('BR_STATUS') || '{}'); } catch (x) { return {}; } } // حالة آخر سحب لكل جهاز
function bump_() { var v = num_('REV') + 1; P_.setProperty('REV', String(v)); return v; }
function nowS_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'); }
function hex_(b) { return b.map(function (x) { return ('0' + (x & 255).toString(16)).slice(-2); }).join(''); }
function hash_(x) { return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify(x))); }
function setTxt_(sh, r, rows) { // نص صريح: يمنع حقن المعادلات ويحفظ الأصفار البادئة
  if (!rows.length) return;
  var rg = sh.getRange(r, 1, rows.length, rows[0].length);
  rg.setNumberFormat('@'); rg.setValues(rows);
}

// ---------- قراءة/كتابة الجداول ----------
function rd_(n) {
  var s = sh_(n), m = s.getLastRow();
  if (m < 2) return [];
  return s.getRange(2, 1, m - 1, HD[n].length).getValues().map(function (r) {
    return r.map(function (c) { return c instanceof Date ? Utilities.formatDate(c, TZ, 'yyyy-MM-dd') : c == null ? '' : String(c); });
  }).filter(function (r) { return r.join('') !== ''; });
}
function wrT_(n, rows) { // يكتب فقط إن تغيّر الجدول (أسرع، ولا يدوس تعديلاً يدويًا في الشيت)
  var h = hash_(rows);
  if (P_.getProperty('H_' + n) === h) return;
  var s = sh_(n), m = s.getLastRow();
  if (m > 1) s.getRange(2, 1, m - 1, HD[n].length).clearContent();
  setTxt_(s, 2, rows.map(function (r) { return r.map(function (c) { return c == null ? '' : c; }); }));
  P_.setProperty('H_' + n, h);
}
function enc_(f, v) { var t = f.split(':')[1]; return t == 'a' ? (v === false ? 'معطل' : 'فعال') : t == 'y' ? (v ? 'نعم' : 'لا') : (v == null ? '' : v); }
function dec_(f, v) { var t = f.split(':')[1]; return t == 'a' ? v !== 'معطل' : t == 'y' ? v === 'نعم' : t == 'n' ? Number(v) || 0 : v; }
function obj_(k, r) { var o = {}; SC[k][1].split(',').forEach(function (c, i) { o[c.split(':')[0]] = dec_(c, r[i]); }); return o; }

function setRows_(db) { // كل الإعدادات (الألوان، الخط، الاسم، المصمم، أوقات الدوام، الشعار...) في ورقة Settings
  var o = [], s = db.set || {}, w = db.work || {}, l = s.logo || '';
  Object.keys(s).forEach(function (k) { if (k !== 'url' && k !== 'logo') o.push(['set.' + k, String(s[k])]); });
  Object.keys(w).forEach(function (k) { o.push(['work.' + k, Array.isArray(w[k]) ? w[k].join(',') : String(w[k])]); });
  for (var i = 0, j = 0; i < l.length; i += 40000, j++) o.push(['logo#' + j, l.substr(i, 40000)]);
  return o;
}
function setObj_(rows) {
  var set = {}, work = {}, logo = [];
  rows.forEach(function (r) {
    var k = r[0], v = r[1];
    if (k.indexOf('set.') == 0) { k = k.slice(4); set[k] = k == 'size' ? Number(v) : k == 'dark' ? v === 'true' : v; }
    else if (k.indexOf('work.') == 0) { k = k.slice(5); work[k] = k == 'days' ? (v === '' ? [] : v.split(',').map(Number)) : (k == 'grace' || k == 'dedupe') ? Number(v) : v; }
    else if (k.indexOf('logo#') == 0) logo[+k.slice(5)] = v;
  });
  if (logo.length) set.logo = logo.join('');
  return { set: set, work: work };
}
function all_() {
  var db = {};
  Object.keys(SC).forEach(function (k) { db[k] = rd_(SC[k][0]).map(function (r) { return obj_(k, r); }); });
  var s = setObj_(rd_('Settings')); db.set = s.set; db.work = s.work;
  db.logs = rd_('Logs').map(function (r) { return { ts: r[0], user: r[1], role: r[2], action: r[3] }; });
  db.errors = rd_('Errors').map(function (r) { return { ts: r[0], user: r[1], msg: r[2], fixed: r[3] || '' }; });
  return db;
}
function write_(db) {
  Object.keys(SC).forEach(function (k) {
    var f = SC[k][1].split(',');
    wrT_(SC[k][0], (db[k] || []).map(function (o) { return f.map(function (c) { return enc_(c, o[c.split(':')[0]]); }); }));
  });
  wrT_('Settings', setRows_(db));
  wrT_('Logs', (db.logs || []).map(function (l) { return [l.ts, l.user, l.role, l.action]; }));
  wrT_('Errors', (db.errors || []).map(function (e) { return [e.ts, e.user, e.msg, e.fixed || '']; }));
}
function addLog_(u, role, a) { var s = sh_('Logs'); setTxt_(s, s.getLastRow() + 1, [[nowS_(), u, role, a]]); P_.deleteProperty('H_Logs'); }
function mergeLogs_(db, u) { // السجل للإضافة فقط: لا حذف ولا تعديل لسجلات الخادم، ونسب أي سجل جديد للمستخدم الحقيقي (الجلسة) لا لما يدّعيه المتصفح
  var srv = rd_('Logs'), seen = {}, out = [], add = 0;
  srv.forEach(function (r) { seen[r[0] + '|' + r[1] + '|' + r[3]] = 1; out.push({ ts: r[0], user: r[1], role: r[2], action: r[3] }); });
  (db.logs || []).forEach(function (l) {
    if (!l || add >= 200 || !/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/.test(String(l.ts))) return;
    var act = String(l.action || '').slice(0, 300), k = l.ts + '|' + u.id + '|' + act;
    if (seen[k]) return;
    seen[k] = 1; add++;
    out.push({ ts: l.ts, user: u.id, role: u.role, action: act });
  });
  out.sort(function (a, b) { return a.ts < b.ts ? -1 : a.ts > b.ts ? 1 : 0; });
  var drop = out.length > 1000 ? out.slice(0, out.length - 1000) : [];
  if (drop.length) { var la = sh_('LogArchive'); setTxt_(la, la.getLastRow() + 1, drop.map(function (l) { return [l.ts, l.user, l.role, l.action]; })); } // لا يُمحى سجل قديم: يُؤرشف
  db.logs = out.slice(-1000);
}
function mergeErrors_(db, u) { // غير الأدمن يضيف أخطاء جديدة فقط ولا يعدّل «تاريخ الحل» ولا يحذف
  if (u.role === 'admin') return;
  var srv = rd_('Errors').map(function (r) { return { ts: r[0], user: r[1], msg: r[2], fixed: r[3] || '' }; }), seen = {}, add = 0;
  srv.forEach(function (e) { seen[e.ts + '|' + e.msg] = 1; });
  (db.errors || []).forEach(function (e) {
    if (!e || add >= 50 || seen[e.ts + '|' + e.msg]) return;
    seen[e.ts + '|' + e.msg] = 1; add++;
    srv.push({ ts: String(e.ts || '').slice(0, 19), user: u.id, msg: String(e.msg || '').slice(0, 300), fixed: '' });
  });
  db.errors = srv.slice(-500);
}
function rowsKey_(k, a) { // مفتاح مقارنة موحّد لجدول (يطبّع الأنواع كما تُكتب في الشيت)
  var f = SC[k][1].split(',');
  return JSON.stringify(a.map(function (o) { return f.map(function (c) { var v = enc_(c, o[c.split(':')[0]]); return String(v == null ? '' : v); }); }));
}
function secure_(db, u, den) { // يفرض الصلاحيات على ما يرسله المتصفح قبل الكتابة في الشيت
  var tables = { employees: 'employees.write', holidays: 'holidays.write', cases: 'cases.write', leaves: 'leaves.write', overtime: 'overtime.write', devices: 'devices.write', movements: 'employees.write', corrections: 'corrections.write', schedules: 'work.write', sassign: 'work.write', closes: 'month.close', ltypes: 'work.write', ledger: 'leaves.write', depts: 'employees.write', rlog: 'reports.issue', fpmap: 'employees.write' };
  var subEmp = db.employees;
  Object.keys(tables).forEach(function (k) {
    if (can_(u, tables[k]) || ((k === 'employees' || k === 'ledger') && can_(u, 'cases.write'))) return; // مسؤول الحالات يخصم الرصيد: يُفحص أدناه بدقة
    var oldO = rd_(SC[k][0]).map(function (r) { return obj_(k, r); });
    if (k === 'devices') { db[k] = oldO; return; } // الأجهزة يعدّلها الأدمن فقط؛ غيره يرى نسخة بلا IP فلا نعدّها محاولة اختراق
    if (rowsKey_(k, oldO) !== rowsKey_(k, db[k] || [])) { db[k] = oldO; den.push(k); }
  });
  chkClosed_(db, u, den); mergeCases_(db, u, den); mergeOT_(db, u, den); mergeRLog_(db, u, den); mergeLedger_(db, u, den); mergeLeaves_(db, u, den);
  if (!can_(u, 'employees.write') && can_(u, 'cases.write')) balOnly_(db, u, subEmp, den);
  chkEmp_(db, u, den); chkFpMap_(db, u, den); mergeCorr_(db, u, den);
  var cur = setObj_(rd_('Settings'));
  if (!can_(u, 'settings.write')) db.set = cur.set;  // تُهمل بصمت: لا يغيّرها إلا من يملك الصلاحية
  if (!can_(u, 'work.write')) db.work = cur.work;
  mergeUsers_(db, u, den);
  if (den.length) addLog_(u.id, u.role, 'نظام: رُفضت تعديلات غير مصرح بها: ' + den.slice(0, 10).join(' ، '));
  mergeErrors_(db, u);
  mergeLogs_(db, u);
}
function mergeUsers_(db, u, den) {
  var oldRows = rd_('Users'), old = {}, out = [], seen = {}, ev = [], bumps = [], now = nowS_();
  oldRows.forEach(function (r) { old[r[0]] = obj_('users', r); });
  var cap = function (role) { return role === 'emp' ? can_(u, 'accounts.emp') : can_(u, 'accounts.staff'); };
  var okRole = function (r) { return ROLE_IDS.indexOf(r) >= 0; };
  (db.users || []).forEach(function (x) {
    if (!x || x.id == null) return;
    var id = String(x.id).trim(), o = old[id], n;
    if (!id || seen[id]) return; seen[id] = 1;
    if (o) {
      if (!cap(o.role)) { // لا يملك صلاحية على هذا الحساب: يُتجاهل أي تغيير فيه
        if ((x.active !== false) !== o.active || (okRole(x.role) && x.role !== o.role) || (x.name != null && String(x.name) !== o.name) || (id !== u.id && x.hash && !pwMatch_(o.hash, x.hash))) den.push('users:' + id);
        out.push(o); return;
      }
      n = JSON.parse(JSON.stringify(o));
      if (okRole(x.role) && x.role !== o.role) {
        if (cap(x.role)) { n.role = x.role; ev.push('تغيير صفة الحساب ' + id + ': ' + o.role + ' ← ' + x.role); bumps.push(id); } else den.push('users:role:' + id);
      }
      var act = x.active !== false;
      if (act !== o.active) { n.active = act; ev.push((act ? 'تفعيل حساب ' : 'تعطيل حساب ') + id); if (!act) bumps.push(id); }
      var un = x.until != null ? cleanUntil_(x.until) : (o.until || '');
      n.until = n.active ? '' : un;
      if (n.until && n.until !== (o.until || '')) ev.push('تعطيل مؤقت للحساب ' + id + ' حتى ' + n.until);
      if (x.name != null) n.name = String(x.name).slice(0, 80);
      if (x.empNo != null) n.empNo = String(x.empNo).slice(0, 40);
      if (id !== u.id && x.salt && x.hash && !pwMatch_(o.hash, x.hash)) {
        if (okSalt_(x.salt) && okHash_(x.hash)) { phAdd_(id, o.salt, o.hash); n.salt = x.salt; n.hash = wrap_(x.hash); phAdd_(id, n.salt, n.hash); n.pwAt = now; n.must = true; ev.push('إعادة تعيين كلمة سر ' + id); bumps.push(id); } else den.push('users:cred:' + id);
      }
      if (u.role === 'admin' && x.perms != null && cleanPerms_(x.perms) !== o.perms) { n.perms = cleanPerms_(x.perms); ev.push('تغيير صلاحيات ' + id + ': ' + (n.perms || 'الافتراضية')); bumps.push(id); }
      if (id === u.id && (!n.active || n.role !== o.role)) { n = o; den.push('users:self'); } // حماية الذات (كلمة السر عبر savepw فقط ويُتجاهل أي hash يرسله المتصفح لنفس الحساب)
      out.push(n);
    } else {
      var rl = okRole(x.role) ? x.role : '';
      if (!rl || !cap(rl)) { den.push('users:+' + id); return; }
      if (!/^[^\s'"\/\\,]{2,40}$/.test(id) || !okSalt_(x.salt) || !okHash_(x.hash) || (rl === 'emp' && !String(x.empNo || '').trim())) { den.push('users:bad:' + id); return; }
      out.push({ id: id, role: rl, name: String(x.name || id).slice(0, 80), empNo: String(x.empNo || '').slice(0, 40), active: x.active !== false, created: now, salt: x.salt, hash: wrap_(x.hash), must: true, lastLogin: '', pwAt: now, by: u.id, perms: u.role === 'admin' ? cleanPerms_(x.perms) : '', until: '' });
      ev.push('إنشاء حساب ' + id + ' (' + rl + ')');
    }
  });
  oldRows.forEach(function (r) { // حذف الحساب: فقط لمن يملك صلاحية على صفته، وليس نفسه
    var o = old[r[0]];
    if (seen[o.id]) return;
    if (cap(o.role) && o.id !== u.id) { ev.push('حذف حساب ' + o.id); bumps.push(o.id); } else out.push(o);
  });
  if (!out.some(function (x) { return x.role === 'admin' && x.active; })) { // حماية آخر أدمن فعّال
    den.push('users:last-admin');
    db.users = oldRows.map(function (r) { return obj_('users', r); });
    return;
  }
  db.users = out;
  bumps.filter(function (v, i, a) { return a.indexOf(v) === i; }).forEach(bumpUv_);
  ev.forEach(function (t) { addLog_(u.id, u.role, 'نظام: ' + t); });
}

// ---------- خصم/إرجاع رصيد الإجازة الاعتيادية بواسطة من يملك cases.write فقط ----------
function wdays_(a, b, wk, hol) { var n = 0, d = a; while (d <= b && n < 400) { var w = new Date(d + 'T00:00:00Z').getUTCDay(); if (wk.indexOf(w) >= 0 && !hol[d]) n++; var t = new Date(d + 'T00:00:00Z'); t.setUTCDate(t.getUTCDate() + 1); d = t.toISOString().slice(0, 10); } return n; }
function balOnly_(db, u, sub, den) { // يُقبل تغيير «الرصيد» فقط، وبمقدار ما تفرضه الحالات الاعتيادية المضافة/المحذوفة؛ أي شيء آخر يُهمل
  var oldE = rd_('Employees').map(function (r) { return obj_('employees', r); }), oldC = {}, newC = {}, delta = {}, hol = {}, wk = (setObj_(rd_('Settings')).work.days || [0, 1, 2, 3, 4]), bad = 0;
  rd_('Holidays').forEach(function (r) { hol[r[0]] = 1; });
  rd_('Cases').forEach(function (r) { oldC[r[0]] = obj_('cases', r); });
  (db.cases || []).forEach(function (c) { newC[String(c.id)] = c; });
  Object.keys(newC).forEach(function (id) { var c = newC[id]; if (!oldC[id] && c.type === 'إجازة اعتيادية') delta[c.empNo] = (delta[c.empNo] || 0) - wdays_(c.from, c.to, wk, hol); });
  Object.keys(oldC).forEach(function (id) { var c = oldC[id]; if (!newC[id] && c.type === 'إجازة اعتيادية') delta[c.empNo] = (delta[c.empNo] || 0) + wdays_(c.from, c.to, wk, hol); });
  var subM = {}; (sub || []).forEach(function (e) { subM[String(e.empNo)] = e; });
  db.employees = oldE.map(function (o) {
    var x = subM[o.empNo], exp = o.balance + (delta[o.empNo] || 0);
    if (!x) { bad++; return o; }
    if (Number(x.balance) !== o.balance) { if (Number(x.balance) === exp && exp >= 0) { o.balance = exp; } else bad++; }
    else if (delta[o.empNo] && exp >= 0) { /* العميل لم يعدّل الرصيد: نطبّق الفرق الصحيح */ o.balance = exp; }
    ['name', 'dept', 'title', 'fpId', 'fpId2', 'fpId3', 'svc', 'start', 'end', 'contract', 'night'].forEach(function (k) { if (String(x[k] == null ? '' : x[k]) !== String(o[k] == null ? '' : o[k])) bad++; });
    return o;
  });
  if ((sub || []).length !== oldE.length) bad++;
  if (bad) den.push('employees: لا يُسمح لمسؤول الحالات إلا بخصم الرصيد الناتج عن الإجازات الاعتيادية');
}

// ---------- المحور 15: سجل التدقيق المتسلسل (كل سطر يحمل بصمة تعتمد على ما قبله) ----------
function sha_(x) { return hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(x))); }
function auditAdd_(u, rows) { // rows: [kind, table, id, field, before, after, note]
  if (!rows.length) return;
  var sh = sh_('Audit'), head = P_.getProperty('AUD_HEAD') || '', now = nowS_(), out = [];
  rows.slice(0, 400).forEach(function (r) {
    var cells = [now, u.id, u.role, r[0], r[1], String(r[2] == null ? '' : r[2]), String(r[3] == null ? '' : r[3]), String(r[4] == null ? '' : r[4]).slice(0, 200), String(r[5] == null ? '' : r[5]).slice(0, 200), String(r[6] || '').slice(0, 200)];
    head = sha_(head + '|' + cells.join('|'));
    out.push(cells.concat([head]));
  });
  setTxt_(sh, sh.getLastRow() + 1, out);
  P_.setProperty('AUD_HEAD', head);
}
var AUD_SKIP_ = { salt: 1, hash: 1, pwAt: 1, lastLogin: 1, created: 1 }; // لا تُدوَّن أبدًا قيم أسرار كلمات المرور
function auditDiff_(db, u) { // فرق «قبل/بعد» يولّده الخادم نفسه للجداول الحساسة: لا يعتمد على ما يدّعيه المتصفح
  var rows = [], T = [['employees', 'empNo', 'الموظفون'], ['users', 'id', 'المستخدمون'], ['devices', 'id', 'الأجهزة'], ['holidays', 'date', 'العطل'], ['schedules', 'id', 'جداول الدوام'], ['sassign', 'id', 'تعيينات الجداول'], ['fpmap', 'id', 'ربط البصمات'], ['ltypes', 'id', 'أنواع الإجازات'], ['depts', 'id', 'الأقسام']];
  T.forEach(function (t) {
    var k = t[0], key = t[1], old = {}, neu = {}, nm = SC[k] ? SC[k][0] : '';
    if (!nm) return;
    var fl = SC[k][1].split(',').map(function (c) { return c.split(':')[0]; });
    rd_(nm).forEach(function (r) { var o = obj_(k, r); old[String(o[key])] = o; });
    (db[k] || []).forEach(function (o) { neu[String(o[key])] = o; });
    Object.keys(neu).forEach(function (id) {
      var n = neu[id], o = old[id];
      if (!o) { rows.push(['إضافة', t[2], id, '', '', k === 'users' ? (n.role + ' / ' + n.name) : k === 'fpmap' ? (n.empNo + ' / ' + n.dev + ' / ' + n.fpId) : (n.name || n.title || ''), '']); return; }
      fl.forEach(function (f) { if (AUD_SKIP_[f] || (k === 'employees' && f === 'balance')) return; var a = Array.isArray(o[f]) ? o[f].join(',') : o[f], b = Array.isArray(n[f]) ? n[f].join(',') : n[f]; if ((a || b) && String(a == null ? '' : a) !== String(b == null ? '' : b)) rows.push(['تعديل', t[2], id, f, a, b, '']); });
    });
    Object.keys(old).forEach(function (id) { if (!neu[id]) rows.push(['حذف', t[2], id, '', k === 'fpmap' ? (old[id].empNo + ' / ' + old[id].dev + ' / ' + old[id].fpId) : (old[id].name || old[id].title || ''), '', '']); });
  });
  var cur = setObj_(rd_('Settings')), nw = { set: db.set || {}, work: db.work || {} };
  ['set', 'work'].forEach(function (g) {
    var a = cur[g] || {}, b = nw[g];
    Object.keys(b).forEach(function (f) { if (f === 'logo' || f === 'url') return; var x = Array.isArray(a[f]) ? a[f].join(',') : a[f], y = Array.isArray(b[f]) ? b[f].join(',') : b[f]; if (String(x == null ? '' : x) !== String(y == null ? '' : y)) rows.push(['تعديل إعداد', g === 'work' ? 'إعدادات الدوام' : 'إعدادات النظام', f, f, x, y, 'يسري على الأيام المفتوحة؛ الأشهر المغلقة تحتفظ بلقطتها']); });
    if (g === 'set' && String(cur.set.logo || '') !== String(b.logo || '')) rows.push(['تعديل إعداد', 'إعدادات النظام', 'logo', 'logo', cur.set.logo ? 'شعار' : '', b.logo ? 'شعار جديد' : '', '']);
  });
  return rows;
}
function auditVerify_() { // يعيد حساب السلسلة كاملة ويكشف أول سطر عُدّل أو حُذف أو أُدرج يدويًا
  var sh = sh_('Audit'), m = sh.getLastRow(), prev = '', n = 0, bad = 0;
  if (m < 2) return { ok: true, checked: 0, brokenAt: 0 };
  var v = sh.getRange(2, 1, m - 1, 11).getValues();
  for (var i = 0; i < v.length; i++) {
    var r = v[i], exp = sha_(prev + '|' + r.slice(0, 10).map(String).join('|'));
    if (exp !== String(r[10])) { bad = i + 2; break; }
    prev = exp; n++;
  }
  var head = P_.getProperty('AUD_HEAD') || '';
  return { ok: !bad && prev === head, checked: n, brokenAt: bad, headMatch: prev === head };
}

// ---------- المحور 16: نسخ احتياطي كامل للملف + فحص واستعادة ----------
function bkFolder_(sub) { var f = folder_('Backups'), it = f.getFoldersByName(sub); return it.hasNext() ? it.next() : f.createFolder(sub); }
function bkRows_() { var n = 0; Object.keys(HD).forEach(function (k) { var sx = ss_().getSheetByName(k); if (sx) n += Math.max(sx.getLastRow() - 1, 0); }); return n; }
function bkLog_(type, name, id, status, note, by, xid) { var sh = sh_('BackupLog'); setTxt_(sh, sh.getLastRow() + 1, [[nowS_(), type, name, id, bkRows_(), status, String(note || '').slice(0, 200), by || 'تلقائي', xid || '']]); if (sh.getLastRow() > 800) sh.deleteRows(2, 200); }
function bkRetain_(sub, keep) { var it = bkFolder_(sub).getFiles(), a = []; while (it.hasNext()) a.push(it.next()); a.sort(function (x, y) { return y.getDateCreated() - x.getDateCreated(); }).slice(keep).forEach(function (x) { x.setTrashed(true); }); }
function bkXlsx_(ssId, name, folderObj) { // يصدّر الملف كاملًا إلى Excel (.xlsx) ويحفظه في مجلد Drive
  var r = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/' + ssId + '/export?format=xlsx', { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) throw new Error('تصدير Excel رفضه الخادم (' + r.getResponseCode() + ')');
  return folderObj.createFile(r.getBlob().setName(name + '.xlsx'));
}
function bkFull_(type, by) { // type: daily | manual | pre_restore — كل نسخة = ملف Excel (.xlsx) للتنزيل + نسخة شيت داخلية للفحص والاستعادة
  var name = 'SS_' + type + '_' + Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd_HHmm'), src = DriveApp.getFileById(ss_().getId()), keep = { daily: 30, manual: 20, pre_restore: 10 }[type] || 10;
  try {
    var cp = src.makeCopy(name, bkFolder_(type)), note = '', xid = '';
    try { var xf = bkXlsx_(ss_().getId(), name, bkFolder_('Excel_' + type)); xid = xf.getId(); bkRetain_('Excel_' + type, keep); } catch (x) { note = 'تعذّر تصدير Excel: ' + String(x).slice(0, 90); P_.setProperty('BK_ERR', nowS_() + ' | ' + note); }
    var ext = String(setObj_(rd_('Settings')).set.bkExt || '').trim();
    if (ext) { try { var ef = DriveApp.getFolderById(ext); if (xid) DriveApp.getFileById(xid).makeCopy(name + '.xlsx', ef); else src.makeCopy(name, ef); note += (note ? ' · ' : '') + 'ونسخة في موقع منفصل'; } catch (x) { note += (note ? ' · ' : '') + 'فشلت النسخة الخارجية: ' + String(x).slice(0, 80); } }
    bkRetain_(type, keep);
    if (xid) { P_.setProperty('BK_OK', nowS_()); P_.deleteProperty('BK_ERR'); }
    bkLog_(type, name, cp.getId(), 'نجاح', note, by, xid);
    return { ok: true, name: name, id: cp.getId(), xid: xid };
  } catch (x) { P_.setProperty('BK_ERR', nowS_() + ' | ' + String(x).slice(0, 150)); bkLog_(type, name, '', 'فشل', String(x), by, ''); return { ok: false, err: String(x).slice(0, 150) }; }
}
function dailyBackup() { bkFull_('daily'); } // مشغّل زمني يعمل كل يوم عند منتصف الليل (بين 12:00 و1:00 ص): نسخة Excel يومية
function installTriggers() { // شغّلها مرة واحدة من محرر Apps Script (تُستدعى أيضًا من setup): تعيد بناء المشغّلات
  var made = [];
  ScriptApp.getProjectTriggers().forEach(function (t) { var h = t.getHandlerFunction(); if (h === 'dailyBackup' || h === 'dailyDigest') ScriptApp.deleteTrigger(t); }); // يحذف القديم (02:00) وملخص البريد الملغى
  ScriptApp.newTrigger('dailyBackup').timeBased().everyDays(1).atHour(0).inTimezone(TZ).create(); made.push('dailyBackup (00:00)');
  return 'أُنشئ: ' + made.join('، ');
}
function bkVerify_(id) { // «استعادة تجريبية» غير مُتلِفة: تفتح النسخة وتقارن بنية الأوراق وأعداد الصفوف بالأصل
  var b = SpreadsheetApp.openById(id), rep = [], bad = 0;
  Object.keys(HD).forEach(function (k) {
    var sx = b.getSheetByName(k), live = ss_().getSheetByName(k);
    if (!sx) { bad++; rep.push([k, 'مفقودة', 0, live ? Math.max(live.getLastRow() - 1, 0) : 0]); return; }
    var hdr = sx.getRange(1, 1, 1, HD[k].length).getValues()[0].map(String).join('|') === HD[k].join('|');
    if (!hdr) bad++;
    rep.push([k, hdr ? 'سليمة' : 'ترويسة مختلفة', Math.max(sx.getLastRow() - 1, 0), live ? Math.max(live.getLastRow() - 1, 0) : 0]);
  });
  return { ok: bad === 0, sheets: rep, bad: bad };
}
function bkRestore_(id, tables, by) { // يستبدل محتوى الأوراق المختارة من النسخة (بعد نسخة أمان تلقائية)، ولا يمسّ الأوراق الأخرى
  var pre = bkFull_('pre_restore', by); if (!pre.ok) throw new Error('تعذّرت نسخة الأمان قبل الاستعادة؛ لم يُنفَّذ شيء');
  var b = SpreadsheetApp.openById(id), done = [];
  tables.forEach(function (k) {
    if (!HD[k] || k === 'Audit' || k === 'BackupLog') return; // سلسلة التدقيق وسجل النسخ لا تُستبدل أبدًا
    var sx = b.getSheetByName(k), live = sh_(k); if (!sx) return;
    var m = sx.getLastRow(), lm = live.getLastRow(), rows = m > 1 ? sx.getRange(2, 1, m - 1, HD[k].length).getValues() : [];
    if (lm > 1) live.getRange(2, 1, lm - 1, HD[k].length).clearContent();
    if (rows.length) setTxt_(live, 2, rows);
    P_.deleteProperty('H_' + k); done.push(k + ' (' + rows.length + ')');
  });
  P_.setProperty('EPOCH', String(num_('EPOCH') + 1)); bump_();
  return done;
}

// ---------- المحور 11: سجل التقارير الرسمية (إضافة فقط) ----------
function mergeRLog_(db, u, den) {
  if (!can_(u, 'reports.issue')) return;
  var oldL = rd_('ReportLog').map(function (r) { return obj_('rlog', r); }), idx = {}, ex = {}, adds = [], fl = 0, now = nowS_();
  oldL.forEach(function (x) { idx[String(x.id)] = 1; ex[x.ref] = 1; });
  (db.rlog || []).forEach(function (x) {
    if (!x || x.id == null || idx[String(x.id)]) return;
    var ref = String(x.ref || '');
    if (!/^RP-\d{8}-[A-Za-z0-9_.]{1,20}-\d{3}$/.test(ref) || ex[ref] || ref.indexOf('-' + u.id + '-') < 0) { fl++; return; }
    ex[ref] = 1;
    adds.push({ id: Number(x.id) || Date.now(), ref: ref, type: String(x.type || '').slice(0, 20), title: String(x.title || '').slice(0, 100), params: String(x.params || '').slice(0, 500), period: String(x.period || '').slice(0, 60), rows: Number(x.rows) || 0, hash: String(x.hash || '').slice(0, 20), by: u.id, at: now });
  });
  db.rlog = oldL.concat(adds); // لا حذف ولا تعديل لسجل التقارير الصادرة
  if (fl) den.push('rlog: ' + fl + ' سجل تقرير غير صالح');
}

// ---------- المحور 7: العمل الإضافي والحالات الوظيفية ----------
var NEEDAPPR_ = ['إيفاد', 'ورقة عمل', 'إجازة بدون راتب', 'إجازة تعويضية', 'إجازة زمنية', 'إجازة مرضية', 'إجازة دراسية'];
function mergeCases_(db, u, den) { // حالات تحتاج اعتمادًا ولا تُحتسب قبله؛ من يسجّلها لا يعتمدها (إلا الأدمن)
  if (!can_(u, 'cases.write')) return;
  var ap = can_(u, 'corrections.approve'), old = {}, fl = 0, cl = closedSet_();
  rd_('Cases').forEach(function (r) { var o = obj_('cases', r); old[String(o.id)] = o; });
  (db.cases || []).forEach(function (x) {
    var o = old[String(x.id)], leaveOrig = String(x.ref || '').indexOf('إجازة #') === 0 && can_(u, 'leaves.write');
    if (!o) { x.hrs = x.type === 'إجازة زمنية' ? Math.max(0, Math.min(744, Number(x.hrs) || 0)) : 0; x.ast = (NEEDAPPR_.indexOf(x.type) < 0 || leaveOrig) ? '' : (ap && u.role === 'admin' ? 'معتمدة' : 'بانتظار الاعتماد'); if (x.ast === 'بانتظار الاعتماد') x.by = u.id; return; }
    if (String(x.ast || '') === String(o.ast || '')) return;
    var ns = x.ast, os = o.ast || 'معتمدة', ok = ap && ((os === 'بانتظار الاعتماد' && (ns === 'معتمدة' || ns === 'مرفوضة')) || (os === 'معتمدة' && ns === 'ملغاة'));
    if (ok && ns === 'معتمدة' && o.by === u.id && u.role !== 'admin') ok = false; // لا اعتماد ذاتي
    if (ok && ns === 'معتمدة' && (cl[String(o.from).slice(0, 7)] || cl[String(o.to).slice(0, 7)])) ok = false;
    if (!ok) { x.ast = o.ast; fl++; } else addLog_(u.id, u.role, 'نظام: ' + (ns === 'معتمدة' ? 'اعتماد' : ns === 'مرفوضة' ? 'رفض' : 'إلغاء') + ' حالة ' + o.type + ' للموظف ' + o.empNo + ' ' + o.from);
  });
  if (fl) den.push('cases: ' + fl + ' تغيير اعتماد غير مسموح');
}
function mergeOT_(db, u, den) { // ساعات إضافية: تسجيل ثم اعتماد من شخص آخر؛ البصمة مؤشر لا دليل
  if (!can_(u, 'overtime.write')) return;
  var ap = can_(u, 'overtime.approve'), cl = closedSet_(), now = nowS_(), dm = /^\d{4}-\d{2}-\d{2}$/;
  var oldL = rd_('Overtime').map(function (r) { return obj_('overtime', r); }), idx = {}, adds = [], seen = {}, fl = 0, ev = [];
  oldL.forEach(function (o) { idx[String(o.id)] = o; });
  (db.overtime || []).forEach(function (x) {
    if (!x || x.id == null || seen[String(x.id)]) return;
    var id = String(x.id), o = idx[id]; seen[id] = 1;
    if (o) {
      if (o.status) ['empNo', 'date', 'hours', 'reason', 'ref', 'by', 'ts'].forEach(function (k) { if (String(x[k] == null ? '' : x[k]) !== String(o[k] == null ? '' : o[k])) fl++; });
      if (String(x.status || '') === String(o.status || '')) return;
      var ns = x.status, os = o.status || 'معتمدة', ok = false, mo = String(o.date).slice(0, 7);
      if (ap && os === 'بانتظار الاعتماد' && (ns === 'معتمدة' || ns === 'مرفوضة')) ok = !(ns === 'معتمدة' && (cl[mo] || (o.by === u.id && u.role !== 'admin')));
      else if (ap && ns === 'ملغاة' && os !== 'ملغاة') ok = !cl[mo];
      else if (o.by === u.id && os === 'بانتظار الاعتماد' && ns === 'ملغاة') ok = true;
      if ((ns === 'مرفوضة' || ns === 'ملغاة') && !String(x.note || '').trim()) ok = false;
      if (!ok) { fl++; return; }
      o.status = ns; o.appr = u.id; o.dts = now; o.note = String(x.note || '').slice(0, 200);
      if (ns === 'معتمدة') { var ah = Number(x.ahours); o.ahours = (isFinite(ah) && ah > 0 && ah <= o.hours) ? ah : o.hours; }
      ev.push((ns === 'معتمدة' ? 'اعتماد' : ns === 'مرفوضة' ? 'رفض' : 'إلغاء') + ' عمل إضافي ' + o.empNo + ' ' + o.date + ' (' + (o.ahours || o.hours) + ' س)');
    } else {
      var emp = String(x.empNo || '').trim(), dt = String(x.date || ''), h = Number(x.hours), rs = String(x.reason || '').trim().slice(0, 300);
      var okE = (db.employees || []).some(function (e) { return String(e.empNo) === emp; });
      if (!okE || !dm.test(dt) || cl[dt.slice(0, 7)] || !isFinite(h) || h <= 0 || h > 16 || !rs) { fl++; return; }
      adds.push({ id: Number(x.id) || Date.now(), empNo: emp, date: dt, hours: h, reason: rs, ref: String(x.ref || '').slice(0, 100), by: u.id, status: 'بانتظار الاعتماد', appr: '', ts: now, dts: '', note: '', ahours: 0 });
      ev.push('تسجيل عمل إضافي ' + emp + ' ' + dt + ' (' + h + ' س)');
    }
  });
  db.overtime = oldL.concat(adds); // لا حذف: الإلغاء حالة
  if (fl) den.push('overtime: ' + fl + ' تعديل غير مسموح');
  ev.forEach(function (t) { addLog_(u.id, u.role, 'نظام: ' + t); });
}

// ---------- المحور 6: الإجازات وسجل الأرصدة ----------
var LT_DEF_ = [
  { id: 1, name: 'اعتيادية', ctype: 'إجازة اعتيادية', deduct: true, count: 'work', needDoc: false, stages: 1, retro: false, on: true },
  { id: 2, name: 'مرضية', ctype: 'إجازة مرضية', deduct: false, count: 'cal', needDoc: true, stages: 1, retro: true, on: true },
  { id: 3, name: 'بدون راتب', ctype: 'إجازة بدون راتب', deduct: false, count: 'work', needDoc: false, stages: 2, retro: false, on: true },
  { id: 4, name: 'دراسية', ctype: 'إجازة دراسية', deduct: false, count: 'work', needDoc: true, stages: 2, retro: false, on: true }
];
function ltypes_() { var r = rd_('LeaveTypes').map(function (x) { return obj_('ltypes', x); }); return r.length ? r : LT_DEF_; }
function leaveDays_(t, a, b) { // أيام تقويمية أو أيام دوام فقط (بحسب نوع الإجازة)
  var st = setObj_(rd_('Settings')), wk = (st.work && st.work.days) || [0, 1, 2, 3, 4], hol = {}, n = 0, d = a, tm;
  rd_('Holidays').forEach(function (r) { hol[r[0]] = 1; });
  while (d <= b && n < 400) { var w = new Date(d + 'T00:00:00Z').getUTCDay(); if (t.count === 'cal' || (wk.indexOf(w) >= 0 && !hol[d])) n++; tm = new Date(d + 'T00:00:00Z'); tm.setUTCDate(tm.getUTCDate() + 1); d = tm.toISOString().slice(0, 10); }
  return n;
}
function mergeLedger_(db, u, den) { // سجل الأرصدة إضافة فقط: لا تعديل ولا حذف
  if (!can_(u, 'leaves.write') && !can_(u, 'cases.write')) return;
  var oldL = rd_('LeaveLedger').map(function (r) { return obj_('ledger', r); }), idx = {}, adds = [], seen = {}, fl = 0, now = nowS_();
  var full = can_(u, 'leaves.write') || can_(u, 'employees.write'), KINDS = ['رصيد افتتاحي', 'استهلاك إجازة', 'إرجاع', 'تسوية إدارية', 'تسوية سنوية', 'خصم زمنيات'];
  oldL.forEach(function (o) { idx[String(o.id)] = o; });
  (db.ledger || []).forEach(function (x) {
    if (!x || x.id == null || seen[String(x.id)]) return;
    var id = String(x.id), o = idx[id]; seen[id] = 1;
    if (o) { if (['empNo', 'date', 'kind', 'delta', 'leaveId', 'note'].some(function (k) { return String(x[k] == null ? '' : x[k]) !== String(o[k] == null ? '' : o[k]); })) fl++; return; }
    var kd = String(x.kind || ''), dl = Number(x.delta), nt = String(x.note || '').trim();
    if (KINDS.indexOf(kd) < 0 || !isFinite(dl) || dl === 0 || !String(x.empNo || '') || !/^\d{4}-\d{2}-\d{2}$/.test(String(x.date || ''))) { fl++; return; }
    if (!full && !((kd === 'استهلاك إجازة' && dl < 0) || (kd === 'إرجاع' && dl > 0))) { fl++; return; }
    if ((kd === 'تسوية إدارية' || kd === 'تسوية سنوية' || kd === 'خصم زمنيات') && !nt) { fl++; return; }
    adds.push({ id: Number(x.id) || Date.now(), empNo: String(x.empNo), date: String(x.date), kind: kd, delta: dl, leaveId: String(x.leaveId || ''), note: nt.slice(0, 200), by: u.id, at: now });
  });
  db.ledger = oldL.concat(adds);
  if (fl) den.push('ledger: ' + fl + ' حركة رصيد غير مسموحة (السجل لا يُعدَّل ولا يُحذف)');
}
function mergeLeaves_(db, u, den) { // طلبات الإجازة تُنشأ من الموظف فقط عبر reqleave؛ هنا انتقالات الحالة المسموحة فقط
  if (!can_(u, 'leaves.write')) return;
  var cl = closedSet_(), now = nowS_(), oldL = rd_('Leaves').map(function (r) { return obj_('leaves', r); }), idx = {}, seen = {}, fl = 0, ev = [];
  oldL.forEach(function (o) { idx[String(o.id)] = o; });
  (db.leaves || []).forEach(function (x) {
    if (!x || x.id == null || seen[String(x.id)]) return;
    var id = String(x.id), o = idx[id]; seen[id] = 1;
    if (!o) { fl++; return; }
    ['empNo', 'from', 'to', 'reason', 'ltype', 'days', 'notes', 'doc', 'ts'].forEach(function (k) { if (String(x[k] == null ? '' : x[k]) !== String(o[k] == null ? '' : o[k])) fl++; });
    if (o.status === 'معلق' && o.stage === 'مراجعة' && x.stage === 'اعتماد') { o.stage = 'اعتماد'; o.rev = u.id; o.rdts = now; o.rnote = String(x.rnote || '').slice(0, 200); ev.push('مراجعة طلب إجازة ' + o.empNo + ' ' + o.from); }
    if (x.status !== o.status) {
      var ns = x.status, ok = (o.status === 'معلق' && (ns === 'موافق' || ns === 'مرفوض' || ns === 'ملغاة')) || (o.status === 'موافق' && ns === 'ملغاة');
      if (ns === 'موافق' && o.stage === 'مراجعة') ok = false; // لا موافقة قبل المراجعة
      if (ns === 'موافق' && (cl[String(o.from).slice(0, 7)] || cl[String(o.to).slice(0, 7)])) ok = false;
      if ((ns === 'مرفوض' || ns === 'ملغاة') && !String(x.anote || '').trim()) ok = false; // السبب إلزامي
      if (!ok) { fl++; return; }
      o.status = ns; o.dts = now; o.appr = u.id; o.anote = String(x.anote || '').slice(0, 200);
      ev.push((ns === 'موافق' ? 'موافقة على' : ns === 'مرفوض' ? 'رفض' : 'إلغاء') + ' طلب إجازة ' + o.empNo + ' ' + o.from + ' → ' + o.to + (o.anote ? ' — ' + o.anote : ''));
    }
  });
  db.leaves = oldL; // لا حذف لطلب؛ الإلغاء حالة
  if (fl) den.push('leaves: ' + fl + ' تعديل غير مسموح (محتوى الطلب ثابت، والانتقالات: معلق→موافق/مرفوض/ملغاة وموافق→ملغاة بسبب)');
  ev.forEach(function (t) { addLog_(u.id, u.role, 'نظام: ' + t); });
}

// ---------- المحور 4: سلامة بيانات الموظفين ----------
function dupsOf_(rows) { // القيم المكررة في (رقم الموظف / أرقام البصمة الأولى والثانية والثالثة معًا)
  var n = {}, f = {}, dn = {}, df = {};
  rows.forEach(function (e) {
    var a = String(e.empNo == null ? '' : e.empNo).trim(), b = nid_(e.fpId), b2 = nid_(e.fpId2), b3 = nid_(e.fpId3);
    if (a) { if (n[a]) dn[a] = 1; n[a] = 1; }
    if ((b && b === b2) || (b && b === b3)) df[b] = 1; else if (b2 && b2 === b3) df[b2] = 1; // أرقام الموظف الواحد متطابقة
    [b, b2, b3].forEach(function (x) { if (x) { if (f[x]) df[x] = 1; f[x] = 1; } });
  });
  return { n: dn, f: df };
}
function chkEmp_(db, u, den) { // لا رقم موظف ولا رقم بصمة مكرر جديد؛ عند المخالفة يعود جدول الموظفين كما على الخادم
  if (!can_(u, 'employees.write')) return;
  var oldRows = rd_('Employees').map(function (r) { return obj_('employees', r); }), od = dupsOf_(oldRows), nd = dupsOf_(db.employees || []), bad = '', k;
  for (k in nd.n) if (!od.n[k]) bad = 'رقم الموظف مكرر: ' + k;
  for (k in nd.f) if (!od.f[k] && !bad) bad = 'رقم البصمة ' + k + ' مرتبط بأكثر من موظف';
  (db.employees || []).forEach(function (e) {
    var s = String(e.start || ''), z = String(e.end || ''), dm = /^\d{4}-\d{2}-\d{2}$/;
    if (!String(e.empNo || '').trim() && !bad) bad = 'موظف بلا رقم';
    if (((s && !dm.test(s)) || (z && !dm.test(z)) || (s && z && z < s)) && !bad) bad = 'تواريخ الخدمة غير صالحة للموظف ' + e.empNo;
  });
  if (bad) { db.employees = oldRows; den.push('employees: ' + bad); }
}

function chkFpMap_(db, u, den) { // ربط البصمات حسب الجهاز: لا يُربط (جهاز + رقم) بأكثر من موظف، والنافذة الزمنية كاملة أو فارغة
  if (!can_(u, 'employees.write')) return;
  var emps = {}, tm = /^([01]\d|2[0-3]):[0-5]\d$/, DIRS = ['', 'دخول', 'خروج', 'كلاهما'], seen = {}, ids = {}, out = [], bad = '';
  (db.employees || []).forEach(function (e) { emps[String(e.empNo)] = 1; });
  (db.fpmap || []).forEach(function (x, i) {
    if (!x) return;
    var o = { id: Number(x.id) || (Date.now() + i), empNo: String(x.empNo == null ? '' : x.empNo).trim(), dev: String(x.dev == null ? '' : x.dev).trim().slice(0, 60), fpId: nid_(x.fpId), dir: DIRS.indexOf(String(x.dir || '')) >= 0 ? String(x.dir || '') : '', from: String(x.from || '').trim(), to: String(x.to || '').trim(), note: String(x.note || '').trim().slice(0, 100) };
    if (!emps[o.empNo]) return; // صف لموظف محذوف يُهمل
    if (!o.dev || !o.fpId) { if (!bad) bad = 'ربط بصمة ناقص (الجهاز أو الرقم) للموظف ' + o.empNo; return; }
    if ((o.from && !tm.test(o.from)) || (o.to && !tm.test(o.to)) || (!!o.from !== !!o.to)) { if (!bad) bad = 'نافذة الوقت غير صالحة للموظف ' + o.empNo + ' (املأ «من» و«إلى» معًا بصيغة HH:MM أو اتركهما فارغين)'; return; }
    var k = o.dev + '|' + o.fpId;
    if (ids[k] && ids[k] !== o.empNo) { if (!bad) bad = 'رقم البصمة ' + o.fpId + ' في الجهاز «' + o.dev + '» مرتبط بأكثر من موظف'; return; }
    ids[k] = o.empNo;
    if (seen[o.empNo + '|' + k]) return; seen[o.empNo + '|' + k] = 1;
    out.push(o);
  });
  if (bad) { db.fpmap = rd_('FpMap').map(function (r) { return obj_('fpmap', r); }); den.push('fpmap: ' + bad); return; }
  db.fpmap = out;
}
function owner_() { // يعيد دالة (جهاز، رقم بصمة) ← رقم الموظف: الربط الدقيق بالجهاز أولًا ثم أرقام الموظف العامة (القديمة)
  var fm = {}, leg = {}, en = {};
  rd_('FpMap').forEach(function (r) { fm[r[2] + '|' + nid_(r[3])] = r[1]; });
  rd_('Employees').forEach(function (r) { en[r[0]] = 1; [nid_(r[5]), nid_(r[10])].forEach(function (k) { if (k && !leg[k]) leg[k] = r[0]; }); });
  return function (dev, fp) { var f = nid_(fp); dev = String(dev || ''); return fm[dev + '|' + f] || (dev === 'يدوي' && en[String(fp)] ? String(fp) : '') || leg[f] || ''; };
}

// ---------- المحور 3: تصحيحات الحضور (الحركة الأصلية لا تُعدَّل أبدًا) ----------
var CORR_ST = ['بانتظار الاعتماد', 'معتمدة', 'مرفوضة', 'ملغاة'];
function closedSet_() { var o = {}; rd_('MonthClose').forEach(function (r) { o[String(r[0])] = 1; }); return o; }
function chkClosed_(db, u, den) { // لا حالات جديدة تمسّ شهرًا مغلقًا؛ المسار: إعادة فتح الشهر بصلاحية month.close ثم إعادة الإغلاق
  if (!can_(u, 'cases.write')) return;
  var cl = closedSet_(), old = {}, rm = 0;
  if (!Object.keys(cl).length) return;
  rd_('Cases').forEach(function (r) { old[String(r[0])] = 1; });
  db.cases = (db.cases || []).filter(function (c) { if (old[String(c.id)]) return true; if (cl[String(c.from || '').slice(0, 7)] || cl[String(c.to || '').slice(0, 7)]) { rm++; return false; } return true; });
  if (rm) den.push('cases: ' + rm + ' حالة ضمن شهر مغلق');
}
function mergeCorr_(db, u, den) {
  if (!can_(u, 'corrections.write')) return; // سبق ردّ الجدول كما هو
  var cl = closedSet_();
  var appr = can_(u, 'corrections.approve'), now = nowS_(), td = now.slice(0, 10), hm = /^([01]\d|2[0-3]):[0-5]\d$/, dm = /^\d{4}-\d{2}-\d{2}$/;
  var oldL = rd_('Corrections').map(function (r) { return obj_('corrections', r); }), idx = {}, adds = [], seen = {}, fl = 0, ev = [];
  oldL.forEach(function (o) { idx[String(o.id)] = o; });
  (db.corrections || []).forEach(function (x) {
    if (!x || x.id == null || seen[String(x.id)]) return;
    var id = String(x.id), o = idx[id]; seen[id] = 1;
    if (o) {
      ['empNo', 'date', 'in', 'out', 'reason'].forEach(function (k) { if (String(x[k] == null ? '' : x[k]) !== String(o[k] == null ? '' : o[k])) fl++; }); // محتوى الطلب ثابت
      if (x.status === o.status || CORR_ST.indexOf(x.status) < 0) return;
      var ok = false;
      if (appr && o.status === 'بانتظار الاعتماد' && (x.status === 'معتمدة' || x.status === 'مرفوضة')) ok = !(x.status === 'معتمدة' && (o.by === u.id && u.role !== 'admin' || cl[String(o.date).slice(0, 7)])); // لا اعتماد ذاتي ولا اعتماد داخل شهر مغلق
      else if (appr && x.status === 'ملغاة' && o.status !== 'ملغاة') ok = true;
      else if (o.by === u.id && o.status === 'بانتظار الاعتماد' && x.status === 'ملغاة') ok = true;
      if (!ok) { fl++; return; }
      o.status = x.status; o.appr = u.id; o.dts = now; o.note = String(x.note || '').slice(0, 200);
      ev.push((x.status === 'معتمدة' ? 'اعتماد' : x.status === 'مرفوضة' ? 'رفض' : 'إلغاء') + ' تصحيح حضور ' + o.empNo + ' ' + o.date + (o.note ? ' — ' + o.note : ''));
    } else {
      var emp = String(x.empNo || '').trim(), dt = String(x.date || ''), tin = String(x.in || ''), tout = String(x.out || ''), rs = String(x.reason || '').trim().slice(0, 300);
      var okE = (db.employees || []).some(function (e) { return String(e.empNo) === emp; });
      var dup = oldL.concat(adds).some(function (c) { return c.empNo === emp && c.date === dt && (c.status === 'بانتظار الاعتماد' || c.status === 'معتمدة'); });
      if (!okE || !dm.test(dt) || dt > td || (!tin && !tout) || (tin && !hm.test(tin)) || (tout && !hm.test(tout)) || (tin && tout && tout <= tin) || !rs || dup || cl[dt.slice(0, 7)]) { fl++; return; }
      adds.push({ id: Number(x.id) || Date.now(), empNo: emp, date: dt, 'in': tin, out: tout, reason: rs, doc: String(x.doc || '').slice(0, 100), by: u.id, status: 'بانتظار الاعتماد', appr: '', ts: now, dts: '', note: '' });
      ev.push('طلب تصحيح حضور ' + emp + ' ' + dt + ' — ' + rs.slice(0, 80));
    }
  });
  db.corrections = oldL.concat(adds); // لا حذف: الإلغاء حالة وليس محوًا
  if (fl) den.push('corrections: ' + fl + ' تعديل غير مسموح (محتوى الطلب ثابت، والاعتماد لصاحب الصلاحية من غير مقدّم الطلب)');
  ev.forEach(function (t) { addLog_(u.id, u.role, 'نظام: ' + t); });
}

// ---------- الصلاحيات (المحور 2) ----------
// لكل صفة صلاحيات افتراضية. يستطيع الأدمن سحب/إضافة صلاحية لمستخدم معيّن من عمود «صلاحيات إضافية» بصيغة: +devices.write,-overtime.write
var PERMS = {
  'accounts.staff': 'إدارة حسابات الأدمن والموارد البشرية والبصمة',
  'accounts.emp': 'إدارة حسابات الموظفين',
  'employees.write': 'إضافة/تعديل/حذف الموظفين والأرصدة',
  'leaves.write': 'معالجة طلبات الإجازة',
  'cases.write': 'تسجيل الحالات الوظيفية',
  'holidays.write': 'العطل الرسمية',
  'overtime.write': 'الساعات الإضافية',
  'work.write': 'أوقات وأيام الدوام',
  'devices.write': 'إدارة أجهزة البصمة',
  'settings.write': 'تنسيقات النظام وإعداداته',
  'punches.write': 'رفع وإضافة البصمات',
  'punches.delete': 'حذف وتعديل البصمات',
  'punches.link': 'ربط رقم بصمة غير مطابق بموظف',
  'corrections.write': 'تقديم طلبات تصحيح الحضور',
  'corrections.approve': 'اعتماد/رفض/إلغاء تصحيح الحضور',
  'month.close': 'إغلاق/إعادة فتح الشهر الإداري',
  'overtime.approve': 'اعتماد/رفض العمل الإضافي',
  'reports.issue': 'إصدار التقارير الرسمية وتسجيلها',
  'audit.view': 'عرض سجل التدقيق والتحقق من سلامته',
  'backup.manage': 'إدارة النسخ الاحتياطي والاستعادة'
};
var ROLE_PERMS = {
  admin: Object.keys(PERMS),
  hr: ['accounts.emp', 'employees.write', 'leaves.write', 'cases.write', 'holidays.write', 'overtime.write', 'work.write', 'punches.link', 'corrections.write', 'corrections.approve', 'month.close', 'overtime.approve', 'reports.issue'],
  fp: ['punches.write', 'punches.delete', 'accounts.emp', 'employees.write', 'leaves.write', 'cases.write', 'holidays.write', 'overtime.write', 'work.write', 'punches.link', 'corrections.write', 'corrections.approve', 'month.close', 'overtime.approve', 'reports.issue'],
  emp: []
};
var NO_GRANT = ['accounts.staff', 'settings.write']; // لا تُمنح لغير الأدمن حتى بالتخصيص (منعًا لتصعيد الصلاحيات)
function cfg_(k, def, lo, hi) { // إعداد أمني من ورقة Settings (مخبّأ دقيقة) بحدود دنيا/عليا آمنة
  var c = C_.get('cfg'), o; try { o = c ? JSON.parse(c) : null; } catch (x) { o = null; }
  if (!o) { o = setObj_(rd_('Settings')).set; var t = {}; ['maxFail', 'lockMin', 'pwMin'].forEach(function (z) { t[z] = o[z]; }); o = t; C_.put('cfg', JSON.stringify(o), 60); }
  var v = Number(o[k]); if (!isFinite(v) || o[k] === undefined || o[k] === '') v = def;
  return Math.max(lo, Math.min(hi, v));
}
var ROLE_IDS = ['admin', 'hr', 'fp', 'emp'];

function perms_(role, extra) { // الفعلية = افتراضي الصفة + تخصيص المستخدم
  var p = (ROLE_PERMS[role] || []).slice();
  if (role === 'admin') return p;
  String(extra || '').split(',').forEach(function (t) {
    t = t.trim(); var sg = t.charAt(0), k = t.slice(1);
    if (!PERMS[k]) return;
    if (sg === '+' && NO_GRANT.indexOf(k) < 0 && p.indexOf(k) < 0) p.push(k);
    if (sg === '-') p = p.filter(function (x) { return x !== k; });
  });
  return p;
}
function cleanPerms_(s) { // يُبقي التوكنات الصحيحة فقط
  return String(s || '').split(',').map(function (t) { return t.trim(); }).filter(function (t) {
    var sg = t.charAt(0), k = t.slice(1);
    return (sg === '+' || sg === '-') && PERMS[k] && !(sg === '+' && NO_GRANT.indexOf(k) >= 0);
  }).join(',');
}
function can_(u, k) { return !!u && (u.perms || []).indexOf(k) >= 0; }
function okHash_(h) { return /^[0-9a-f]{64}$/.test(String(h || '')); }
function okSalt_(x) { return /^[0-9a-f]{32}$/.test(String(x || '')); }
function uvMap_() { try { return JSON.parse(P_.getProperty('UV') || '{}'); } catch (x) { return {}; } }
function uv_(id) { return uvMap_()[id] | 0; }
function wrap_(h) { return 'w:' + hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(h) + '|' + key_())); } // ما يُخزَّن في الشيت: لا يصلح وحده للدخول
function pwMatch_(stored, h) { stored = String(stored || ''); return !!h && (stored.indexOf('w:') === 0 ? stored === wrap_(h) : stored === h); } // يقبل القديم (hash خام) ويرقّيه عند أول دخول
function ph_(id) { try { return JSON.parse(P_.getProperty('PH_' + id) || '[]'); } catch (x) { return []; } } // آخر 5 أزواج [salt, hash مخزَّن]
function phAdd_(id, salt, hash) { if (!salt || !hash) return; var a = ph_(id).filter(function (p) { return p[1] !== hash; }); a.push([salt, hash]); P_.setProperty('PH_' + id, JSON.stringify(a.slice(-5))); }
function cleanUntil_(v) { v = String(v || '').slice(0, 19); return /^\d{4}-\d{2}-\d{2}( \d{2}:\d{2}(:\d{2})?)?$/.test(v) ? v : ''; }
function bumpUv_(id) { var m = uvMap_(); m[id] = (m[id] | 0) + 1; P_.setProperty('UV', JSON.stringify(m)); } // يبطل كل جلسات هذا المستخدم

// ---------- الجلسات وتسجيل الدخول (من ورقة Users) ----------
// الجلسة لا تنتهي بالخمول: تبطل فقط بتسجيل الخروج أو تغيير كلمة السر/التعطيل/الصفة (uv). الذاكرة المؤقتة قد تُفرَّغ في أي وقت، لذا يُحفظ نسخة من جلسات الكادر في خصائص السكربت أيضًا
var SESS_MAX = 60, SESS_DAYS = 30;
function sessIdx_() { try { return JSON.parse(P_.getProperty('SS_IDX') || '[]'); } catch (x) { return []; } }
function sessP_(t) { try { var s = JSON.parse(P_.getProperty('SS_' + t) || 'null'); return s && Date.now() - (s.la || 0) < SESS_DAYS * 864e5 ? s : null; } catch (x) { return null; } }
function sessSave_(t, c) {
  var ix = sessIdx_(), hit = null, sig = (c.must ? 1 : 0) + ':' + (c.uv | 0), now = Date.now();
  ix.forEach(function (r) { if (r[0] === t) hit = r; });
  if (hit && now - hit[1] < 600000 && hit[2] === sig) return; // تُكتب كل 10 دقائق كحد أقصى ما لم يتغير وضعها
  var keep = [];
  ix.forEach(function (r) { if (r[0] === t) return; if (now - r[1] >= SESS_DAYS * 864e5) P_.deleteProperty('SS_' + r[0]); else keep.push(r); });
  keep.push([t, now, sig]);
  while (keep.length > SESS_MAX) P_.deleteProperty('SS_' + keep.shift()[0]);
  P_.setProperty('SS_' + t, JSON.stringify(c)); P_.setProperty('SS_IDX', JSON.stringify(keep));
}
function sessDel_(t) { C_.remove('s:' + t); P_.deleteProperty('SS_' + t); P_.setProperty('SS_IDX', JSON.stringify(sessIdx_().filter(function (r) { return r[0] !== t; }))); }
function sess_(t) {
  var s = C_.get('s:' + t), fromP = false;
  if (s) { try { s = JSON.parse(s); } catch (x) { s = null; } } else s = null;
  if (!s) { s = sessP_(t); fromP = !!s; }
  if (!s) return null;
  if ((s.uv | 0) !== uv_(s.id)) { sessDel_(t); return null; }
  if (fromP) C_.put('s:' + t, JSON.stringify(s), 21600);
  return s;
}
function putSess_(t, u) { var c = {}; Object.keys(u).forEach(function (k) { if (k.charAt(0) !== '_') c[k] = u[k]; }); C_.put('s:' + t, JSON.stringify(c), 21600); if (c.role !== 'emp') sessSave_(t, c); }
function urow_(id) { // رقم صف المستخدم الفعلي في الورقة
  var s = sh_('Users'), m = s.getLastRow();
  if (m < 2) return 0;
  var v = s.getRange(2, 1, m - 1, 1).getValues();
  for (var i = 0; i < v.length; i++) if (String(v[i][0]) === id) return i + 2;
  return 0;
}
function start_(u, extra) {
  var t = Utilities.getUuid().replace(/-/g, ''),
      s = { id: u.id, role: u.role, name: u.name, empNo: u.empNo, must: !!u.must, uv: uv_(u.id), perms: perms_(u.role, extra), la: Date.now() };
  putSess_(t, s);
  P_.setProperty('ST_' + u.role, String(num_('ST_' + u.role) + 1));
  var rw = urow_(u.id);
  if (rw) { sh_('Users').getRange(rw, 10).setNumberFormat('@').setValue(nowS_()); P_.deleteProperty('H_Users'); }
  addLog_(u.id, u.role, 'تسجيل دخول');
  var p = pull_(s);
  return { ok: true, tok: t, user: s, rev: p.rev, db: p.db };
}
function info_() { // عام (قبل الدخول): الشعار والألوان وهل يوجد أدمن
  return { ok: true, v: 3, set: setObj_(rd_('Settings')).set, hasAdmin: rd_('Users').some(function (r) { return r[1] === 'admin'; }) };
}
function salt_(d) {
  var id = String(d.id || '').trim(), u = rd_('Users').filter(function (r) { return r[0] === id; })[0];
  return { ok: true, salt: u && u[6] ? u[6] : hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, id + key_())).slice(0, 32) };
}
function fail_(id, role) { // قفل تدريجي: 5 محاولات → 5 دقائق، ثم 15 دقيقة، ثم ساعة
  var n = (Number(C_.get('f:' + id)) || 0) + 1, g = Number(C_.get('g:' + id)) || 0, fl = Number(C_.get('flg')) || 0;
  if (n >= cfg_('maxFail', 5, 3, 10)) {
    C_.put('f:' + id, String(n), [cfg_('lockMin', 5, 1, 60) * 60, cfg_('lockMin', 5, 1, 60) * 180, 3600][Math.min(g, 2)]); C_.put('g:' + id, String(g + 1), 21600);
    addLog_(id, role || '-', 'قفل الحساب مؤقتًا بعد محاولات دخول فاشلة');
  } else C_.put('f:' + id, String(n), 300);
  if (fl < 30) { C_.put('flg', String(fl + 1), 300); addLog_(id, role || '-', 'محاولة دخول فاشلة'); } // سقف لمنع إغراق السجل
}
function login_(d) {
  var id = String(d.id || '').trim(), n = Number(C_.get('f:' + id)) || 0, nowT = nowS_();
  if (n >= cfg_('maxFail', 5, 3, 10)) return { ok: false, err: 'locked' };
  var u = rd_('Users').filter(function (r) { return r[0] === id && r[1] === d.role && (r[4] !== 'معطل' || (r[13] && String(r[13]) <= nowT)); })[0]; // التعطيل المؤقت ينتهي تلقائيًا
  if (!u || !pwMatch_(u[7], d.hash)) { fail_(id, d.role); return { ok: false, err: 'bad' }; }
  C_.remove('f:' + id); C_.remove('g:' + id);
  var rw = urow_(id), sU = sh_('Users'), must = u[8] === 'نعم', pd = Number(setObj_(rd_('Settings')).set.pwDays) || 0, base = String(u[10] || u[5] || '');
  if (rw) {
    if (u[4] === 'معطل') { sU.getRange(rw, 5).setNumberFormat('@').setValue('فعال'); sU.getRange(rw, 14).setNumberFormat('@').setValue(''); addLog_(id, u[1], 'نظام: انتهى التعطيل المؤقت وأُعيد تفعيل الحساب'); }
    if (String(u[7]).indexOf('w:') !== 0) sU.getRange(rw, 8).setNumberFormat('@').setValue(wrap_(d.hash)); // ترقية التخزين
    if (!must && pd > 0 && base && (Date.parse(nowT.replace(' ', 'T')) - Date.parse(base.replace(' ', 'T'))) / 864e5 >= pd) {
      must = true; sU.getRange(rw, 9).setNumberFormat('@').setValue('نعم'); addLog_(id, u[1], 'نظام: انتهت صلاحية كلمة السر (' + pd + ' يوم) — تغيير إلزامي');
    }
    P_.deleteProperty('H_Users');
  }
  return start_({ id: u[0], role: u[1], name: u[2], empNo: u[3], must: must }, u[12]);
}
function setadmin_(d) { // أول أدمن فقط، ومرة واحدة، يُتحقق من الشيت نفسه
  var us = rd_('Users'), id = String(d.id || '').trim();
  if (us.some(function (r) { return r[1] === 'admin'; })) return { ok: false, err: 'يوجد أدمن مسبقًا' };
  if (!/^[A-Za-z0-9_.-]{4,}$/.test(id) || !okSalt_(d.salt) || !okHash_(d.hash)) return { ok: false, err: 'بيانات غير صالحة (يلزم اتصال آمن HTTPS لتشفير كلمة السر)' };
  var nm = String(d.name || 'الأدمن').slice(0, 60);
  us.push([id, 'admin', nm, '', 'فعال', nowS_(), d.salt, wrap_(d.hash), 'لا', '', nowS_(), 'setup', '', '']);
  wrT_('Users', us); bump_();
  return start_({ id: id, role: 'admin', name: nm, empNo: '', must: false }, '');
}
function pull_(u) {
  var db = all_(), n = u.empNo;
  // الأمان: hash لا يُرسل لأحد (حتى صاحبه)، والـ salt لصاحبه فقط
  db.users = db.users.map(function (x) { if (x.id === u.id) return x; var c = {}; Object.keys(x).forEach(function (k) { c[k] = x[k]; }); c.hash = ''; if (x.id !== u.id) c.salt = ''; return c; });
  db.stats = { logins: { admin: num_('ST_admin'), hr: num_('ST_hr'), fp: num_('ST_fp'), emp: num_('ST_emp') } };
  if (u.role !== 'emp') { db.stats.bridge = brSt_(); db.stats.backup = bkSt_(); try { db.stats.hb = JSON.parse(P_.getProperty('BR_HB') || '{}'); } catch (x) { db.stats.hb = {}; } }
  if (!can_(u, 'devices.write')) db.devices = db.devices.map(function (x) { var c = {}; Object.keys(x).forEach(function (k) { c[k] = x[k]; }); c.ip = ''; c.port = ''; return c; }); // عناوين الاتصال حساسة
  if (u.role === 'emp') { // الموظف يرى بياناته فقط
    db.users = db.users.filter(function (x) { return x.id === u.id; });
    db.employees = db.employees.filter(function (x) { return x.empNo === n; });
    ['cases', 'leaves', 'overtime', 'corrections', 'movements', 'sassign', 'closes', 'ledger', 'fpmap'].forEach(function (k) { db[k] = db[k].filter(function (x) { return x.empNo === n; }); }); db.rlog = [];
    db.devices = []; db.logs = []; db.errors = [];
  }
  return { ok: true, db: db, rev: num_('REV') };
}

function doGet() { return out_({ ok: true, v: 3, msg: 'FP system ready v3.8' }); }

function doPost(e) {
  var d;
  try { d = JSON.parse(e.postData.contents); } catch (x) { return out_({ ok: false, err: 'bad request' }); }
  var lock = LockService.getScriptLock();
  try { lock.waitLock(25000); } catch (x) { return out_({ ok: false, err: 'الخادم مشغول، أعد المحاولة' }); }
  try {
    ensure_();
    var a = d.action;
    if (a === 'info') return out_(info_());
    if (a === 'salt') return out_(salt_(d));
    if (a === 'login') return out_(login_(d));
    if (a === 'setadmin') return out_(setadmin_(d));
    var u = null;
    if (d.t) { u = sess_(d.t); if (u) u._t = d.t; }
    else if (d.k && d.k === key_()) u = { role: 'bridge', perms: ['punches.write'] }; // المفتاح للجسر فقط
    if (!u) return out_({ ok: false, err: 'unauthorized' });
    if (u.must && ['savepw', 'pwhist', 'check', 'pull', 'rev', 'logout'].indexOf(a) < 0) return out_({ ok: false, err: 'forbidden' }); // لا عمل قبل تغيير كلمة السر الإلزامي
    var res = route_(d, u);
    if (u._t && !u._gone && ['rev', 'check', 'bstat'].indexOf(a) < 0) { u.la = Date.now(); putSess_(u._t, u); } // تمديد الجلسة بالنشاط الفعلي فقط (الاستعلام الدوري لا يمدّدها)
    return out_(res);
  } catch (err) { return out_({ ok: false, err: String(err && err.message ? err.message : err) }); }
  finally { lock.releaseLock(); }
}

function route_(d, u) {
  var r = u.role, a = d.action, staff = r === 'admin' || r === 'hr' || r === 'fp', p, m, rows;
  if (r === 'bridge' && ['punches', 'devices', 'bridge'].indexOf(a) < 0) return { ok: false, err: 'unauthorized' };
  if (r === 'emp' && ['check', 'rev', 'pull', 'since', 'savepw', 'pwhist', 'reqleave', 'cancelleave', 'reqcorr', 'logout'].indexOf(a) < 0) return { ok: false, err: 'forbidden' };
  if (a === 'punches' && !can_(u, 'punches.write')) return { ok: false, err: 'forbidden' };
  if (a === 'delpunch' && !can_(u, 'punches.delete')) return { ok: false, err: 'forbidden' };
  switch (a) {
    case 'check': return { ok: true, role: r };
    case 'rev': return { ok: true, rev: num_('REV') };
    case 'pull': return pull_(u);

    case 'pwhist': // أملاح كلمات السر السابقة لصاحب الحساب، ليحسب المتصفح hash المرشّح ويتحقق الخادم من عدم التكرار
      var rwh = urow_(u.id), sl = [String(sh_('Users').getRange(rwh, 7).getValue())];
      ph_(u.id).forEach(function (p) { if (sl.indexOf(p[0]) < 0) sl.push(p[0]); });
      return { ok: true, salts: sl };

    case 'savepw': // تغيير كلمة سر المستخدم نفسه فقط
      if (!okSalt_(d.salt) || !okHash_(d.hash)) throw new Error('تعذّر تشفير كلمة السر في هذا المتصفح (يلزم اتصال آمن HTTPS)');
      if (!Array.isArray(d.chk)) throw new Error('حدّث الصفحة (Ctrl+F5) ثم أعد المحاولة');
      var rw = urow_(u.id), sU = sh_('Users');
      if (!rw) throw new Error('الحساب غير موجود');
      var cr = sU.getRange(rw, 1, 1, 8).getValues()[0], hist = ph_(u.id).concat([[String(cr[6]), String(cr[7])]]);
      if (hist.some(function (p) { return d.chk.slice(0, 8).some(function (c) { return okHash_(c) && pwMatch_(p[1], c); }); })) throw new Error('لا يجوز إعادة استخدام كلمة سر سابقة (آخر 5 كلمات)');
      phAdd_(u.id, String(cr[6]), String(cr[7]));
      var nh = wrap_(d.hash);
      sU.getRange(rw, 7, 1, 2).setNumberFormat('@').setValues([[d.salt, nh]]);
      sU.getRange(rw, 9).setNumberFormat('@').setValue('لا');
      sU.getRange(rw, 11).setNumberFormat('@').setValue(nowS_());
      phAdd_(u.id, d.salt, nh);
      P_.deleteProperty('H_Users');
      bumpUv_(u.id); u.uv = uv_(u.id); u.must = false; // تُبطل جلسات هذا المستخدم الأخرى وتبقى الحالية
      addLog_(u.id, u.role, 'نظام: تغيير كلمة السر');
      return { ok: true, rev: bump_() };

    case 'killsess': // إنهاء كل جلسات مستخدم (يملك الصلاحية على صفته)
      var tg = rd_('Users').filter(function (x) { return x[0] === String(d.id || ''); })[0];
      if (!tg) return { ok: false, err: 'الحساب غير موجود' };
      if (!(tg[1] === 'emp' ? can_(u, 'accounts.emp') : can_(u, 'accounts.staff'))) return { ok: false, err: 'forbidden' };
      bumpUv_(tg[0]); if (tg[0] === u.id) u.uv = uv_(u.id); // لا تُنهي جلستك الحالية
      addLog_(u.id, u.role, 'نظام: إنهاء كل جلسات ' + tg[0]);
      return { ok: true };

    case 'logout': // إنهاء الجلسة على الخادم فعليًا
      if (u._t) { sessDel_(u._t); u._gone = true; if (u.role === 'emp') addLog_(u.id, u.role, 'تسجيل خروج'); }
      return { ok: true };

    case 'reqleave': // طلب إجازة من الموظف: يتحقق الخادم من النوع والأيام والتعارض والرصيد والمستند
      if (r !== 'emp' || !u.empNo) return { ok: false, err: 'forbidden' };
      if (!/^\d{4}-\d{2}-\d{2}$/.test(d.from) || !/^\d{4}-\d{2}-\d{2}$/.test(d.to) || d.to < d.from) throw new Error('تواريخ غير صالحة');
      var lt = ltypes_().filter(function (x) { return String(x.id) === String(d.type || 1); })[0];
      if (!lt || lt.on === false) throw new Error('نوع الإجازة غير متاح');
      var td6 = nowS_().slice(0, 10), cl6 = closedSet_();
      if (d.from < td6 && !lt.retro) throw new Error('لا يجوز طلب هذا النوع بتاريخ سابق');
      if (cl6[d.from.slice(0, 7)] || cl6[d.to.slice(0, 7)]) throw new Error('الفترة ضمن شهر مغلق');
      var nd = leaveDays_(lt, d.from, d.to); if (!nd) throw new Error('الفترة لا تحتوي أيامًا تُحتسب لهذا النوع');
      var doc6 = String(d.doc || '').trim().slice(0, 100); if (lt.needDoc && !doc6) throw new Error('هذا النوع يتطلب رقم/وصف المستند المؤيد');
      var allL = rd_('Leaves').map(function (x) { return obj_('leaves', x); }), pendD = 0;
      allL.forEach(function (x) {
        if (x.empNo !== u.empNo || (x.status !== 'معلق' && x.status !== 'موافق')) return;
        if (x.from <= d.to && x.to >= d.from) throw new Error('يتعارض مع طلب إجازة آخر (' + x.from + ' إلى ' + x.to + ')');
        var xt = ltypes_().filter(function (y) { return String(y.id) === String(x.ltype || 1); })[0] || LT_DEF_[0];
        if (x.status === 'معلق' && xt.deduct) pendD += x.days || leaveDays_(xt, x.from, x.to);
      });
      if (lt.deduct) { var er = rd_('Employees').filter(function (x) { return x[0] === u.empNo; })[0], bal = er ? Number(er[4]) || 0 : 0; if (nd + pendD > bal) throw new Error('الرصيد (' + bal + ') لا يكفي، علمًا بوجود ' + pendD + ' يوم قيد الموافقة'); }
      var lv = { id: Date.now(), empNo: u.empNo, from: d.from, to: d.to, reason: String(d.reason || '').slice(0, 200), status: 'معلق', ts: nowS_(), dts: '', ltype: String(lt.id), days: nd, notes: String(d.notes || '').slice(0, 300), doc: doc6, stage: lt.stages === 2 ? 'مراجعة' : 'اعتماد', rev: '', rdts: '', rnote: '', appr: '', anote: '' };
      p = sh_('Leaves'); setTxt_(p, p.getLastRow() + 1, [SC.leaves[1].split(',').map(function (c) { return enc_(c, lv[c.split(':')[0]]); })]);
      P_.deleteProperty('H_Leaves');
      return { ok: true, rev: bump_(), leave: lv };

    case 'reqcorr': // الموظف يطلب تصحيح بصمته أو يعترض على نتيجة يوم (يمرّ بالاعتماد ككل التصحيحات)
      if (r !== 'emp' || !u.empNo) return { ok: false, err: 'forbidden' };
      var dc = String(d.date || ''), ci = String(d.in || ''), co = String(d.out || ''), rc = String(d.reason || '').trim().slice(0, 300), tdc = nowS_().slice(0, 10);
      var hmc = /^([01]\d|2[0-3]):[0-5]\d$/;
      if (!/^\d{4}-\d{2}-\d{2}$/.test(dc) || dc > tdc) throw new Error('التاريخ غير صالح');
      if (dc < new Date(new Date(tdc + 'T00:00:00Z').getTime() - 45 * 864e5).toISOString().slice(0, 10)) throw new Error('لا يُقبل طلب لتاريخ أقدم من 45 يومًا');
      if (closedSet_()[dc.slice(0, 7)]) throw new Error('هذا اليوم ضمن شهر مغلق؛ راجع الموارد البشرية');
      if (!ci && !co) throw new Error('أدخل وقت الدخول و/أو الخروج الصحيح');
      if ((ci && !hmc.test(ci)) || (co && !hmc.test(co)) || (ci && co && co <= ci)) throw new Error('الأوقات غير صالحة');
      if (!rc) throw new Error('السبب مطلوب');
      var allc = rd_('Corrections').map(function (x) { return obj_('corrections', x); }), pc = 0;
      allc.forEach(function (x) { if (x.empNo === u.empNo && x.status === 'بانتظار الاعتماد') pc++; if (x.empNo === u.empNo && x.date === dc && (x.status === 'بانتظار الاعتماد' || x.status === 'معتمدة')) throw new Error('يوجد طلب قائم أو معتمد لهذا اليوم'); });
      if (pc >= 5) throw new Error('لديك 5 طلبات معلّقة؛ انتظر البتّ فيها');
      var cr = { id: Date.now(), empNo: u.empNo, date: dc, 'in': ci, out: co, reason: (d.kind === 'اعتراض' ? 'اعتراض: ' : 'طلب موظف: ') + rc, doc: String(d.doc || '').slice(0, 100), by: u.id, status: 'بانتظار الاعتماد', appr: '', ts: nowS_(), dts: '', note: '' };
      p = sh_('Corrections'); setTxt_(p, p.getLastRow() + 1, [SC.corrections[1].split(',').map(function (c) { return enc_(c, cr[c.split(':')[0]]); })]);
      P_.deleteProperty('H_Corrections');
      addLog_(u.id, r, 'نظام: ' + (d.kind === 'اعتراض' ? 'اعتراض' : 'طلب تصحيح') + ' من الموظف ' + u.empNo + ' ليوم ' + dc);
      return { ok: true, rev: bump_() };

    case 'cancelleave': // إلغاء الموظف لطلبه المعلّق فقط
      if (r !== 'emp' || !u.empNo) return { ok: false, err: 'forbidden' };
      var sL = sh_('Leaves'), cA = sL.getRange(2, 1, Math.max(sL.getLastRow() - 1, 1), 6).getValues(), rw6 = 0;
      for (var q6 = 0; q6 < cA.length; q6++) if (String(cA[q6][0]) === String(d.id) && cA[q6][1] === u.empNo && cA[q6][5] === 'معلق') { rw6 = q6 + 2; break; }
      if (!rw6) throw new Error('لا يوجد طلب معلّق بهذا الرقم');
      sL.getRange(rw6, 6).setValue('ملغاة'); sL.getRange(rw6, 8).setValue(nowS_()); sL.getRange(rw6, 18).setValue('ألغاه الموظف');
      P_.deleteProperty('H_Leaves');
      addLog_(u.id, r, 'نظام: إلغاء الموظف لطلب إجازة ' + d.id);
      return { ok: true, rev: bump_() };

    case 'push':
      if (!staff) return { ok: false, err: 'forbidden' };
      var db = d.db, rev = num_('REV'), den = [];
      if (!db || !Array.isArray(db.users) || !Array.isArray(db.employees)) throw new Error('بنية البيانات غير صالحة');
      if (!Array.isArray(db.fpmap)) db.fpmap = rd_('FpMap').map(function (r) { return obj_('fpmap', r); }); // صفحة قديمة لا تعرف الربط: لا تمحُه
      if (!d.force && (d.base | 0) !== rev) return { ok: false, conflict: true, rev: rev, err: 'conflict' };
      secure_(db, u, den); // فرض الصلاحيات على الخادم: ما لا يحق له يُتجاهل ويُسجَّل
      try { auditAdd_(u, auditDiff_(db, u)); } catch (x) { addLog_(u.id, r, 'نظام: تعذّر تدوين سجل التدقيق: ' + String(x).slice(0, 100)); }
      // النسخ الاحتياطي: ملف Excel يومي عند منتصف الليل (dailyBackup) — لا نسخ JSON ساعي بعد الآن
      write_(db);
      return { ok: true, rev: bump_(), denied: den };

    case 'punches': // مفتاح الحركة: (الرقم | الوقت | الجهاز) — رقم بصمة واحد قد يخص موظفين مختلفين في جهازين مختلفين
      p = sh_('Punches'); m = p.getLastRow(); var ex = {};
      if (m > 1) p.getRange(2, 1, m - 1, 3).getValues().forEach(function (x) { ex[nid_(x[0]) + '|' + x[1] + '|' + String(x[2] || '')] = 1; });
      var add = [], bad = 0, dups = [], resent = 0, inq = {}, nw = nowS_(), via = r === 'bridge' ? String(d.method || 'جسر Python').slice(0, 30) : String(d.method || '').slice(0, 30), defDev = String(d.dev || '').slice(0, 60);
      (d.rows || []).slice(0, 5000).forEach(function (x) { // حد أقصى 5000 صف للطلب الواحد
        var f = nid_(x[0]), dv = String(x[2] || defDev).slice(0, 60), k = f + '|' + x[1] + '|' + dv;
        if (!f || !/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$/.test(x[1])) { bad++; return; }
        if (inq[k]) { dups.push([nw, f, String(x[1]), dv, 'مكررة داخل نفس الملف/الدفعة', via]); return; }
        inq[k] = 1;
        if (ex.hasOwnProperty(k)) { resent++; return; } // إعادة إرسال نفس الحركة من نفس الجهاز أمر عادي (الجسر يعيد قراءة آخر أيام)
        ex[k] = 1; add.push([f, String(x[1]), dv, via, nw]);
      });
      setTxt_(p, m + 1, add);
      if (dups.length) { var pl = sh_('PunchLog'); setTxt_(pl, pl.getLastRow() + 1, dups.slice(0, 500)); if (pl.getLastRow() > 6000) pl.deleteRows(2, 2000); }
      return { ok: true, added: add.length, skipped: bad, dups: dups.length, resent: resent };

    case 'devices': // الأجهزة المفعّلة: الجسر يقرأها من هنا بدل تعديل ملف بايثون
      return { ok: true, devices: rd_('Devices').map(function (x) { return obj_('devices', x); }).filter(function (x) { return x.on && x.ip; }) };

    case 'bridge': // حالة السحب يرسلها الجسر (بمفتاح الجسر فقط): آخر حالة لكل جهاز + سجل عمليات + نبضة الخدمة
      if (r !== 'bridge') return { ok: false, err: 'forbidden' };
      var st = brSt_(), ks, nb = nowS_(), dl = [];
      (d.items || []).slice(0, 20).forEach(function (i) {
        var nm = String(i.name || '?').slice(0, 60), pv = st[nm] || {}, ok8 = !!i.ok;
        st[nm] = { ts: nb, ok: ok8, n: i.n | 0, err: String(i.err || '').slice(0, 200), last: String(i.last || '').slice(0, 20), kind: String(i.kind || '').slice(0, 20), ms: i.ms | 0, read: i.read | 0, dups: i.dups | 0, bad: i.bad | 0, okts: ok8 ? nb : (pv.okts || ''), fails: ok8 ? 0 : ((pv.fails | 0) + 1) };
        dl.push([nb, nm, 'سحب', ok8 ? 'نجاح' : 'فشل (' + String(i.kind || 'غير مصنّف').slice(0, 20) + ')', ok8 ? '' : String(i.err || '').slice(0, 200), 'bridge', i.ms | 0, i.read | 0, i.n | 0, i.dups | 0, i.bad | 0]);
      });
      ks = Object.keys(st);
      if (ks.length > 20) ks.sort(function (x, y) { return st[x].ts < st[y].ts ? -1 : 1; }).slice(0, ks.length - 20).forEach(function (k) { delete st[k]; });
      P_.setProperty('BR_STATUS', JSON.stringify(st));
      if (d.hb) P_.setProperty('BR_HB', JSON.stringify({ ts: nb, v: String(d.hb.v || '').slice(0, 10), host: String(d.hb.host || '').slice(0, 40), mode: String(d.hb.mode || '').slice(0, 20), err: String(d.hb.err || '').slice(0, 200), okts: String(d.hb.okts || '').slice(0, 20), fails: d.hb.fails | 0 }));
      if (dl.length) { var dsh = sh_('DevLog'); setTxt_(dsh, dsh.getLastRow() + 1, dl); if (dsh.getLastRow() > 6000) dsh.deleteRows(2, 1500); }
      return { ok: true };

    case 'bstat': return { ok: true, bridge: brSt_(), hb: (function () { try { return JSON.parse(P_.getProperty('BR_HB') || '{}'); } catch (x) { return {}; } })() };

    case 'devlog': // سجل عمليات الأجهزة (آخر 300، ويمكن تصفيته باسم جهاز)
      if (!staff) return { ok: false, err: 'forbidden' };
      p = sh_('DevLog'); m = p.getLastRow();
      var dvr = m > 1 ? p.getRange(Math.max(2, m - 1499), 1, Math.min(1500, m - 1), 11).getValues() : [];
      if (d.name) dvr = dvr.filter(function (x) { return x[1] === d.name; });
      return { ok: true, rows: dvr.slice(-300) };

    case 'auditlog': // سجل التدقيق (آخر 400 مع تصفية) — لمن يملك audit.view فقط
      if (!can_(u, 'audit.view')) return { ok: false, err: 'forbidden' };
      p = sh_('Audit'); m = p.getLastRow();
      var ar = m > 1 ? p.getRange(Math.max(2, m - 1999), 1, Math.min(2000, m - 1), 11).getValues() : [];
      if (d.user) ar = ar.filter(function (x) { return x[1] === d.user; });
      if (d.kind) ar = ar.filter(function (x) { return x[3] === d.kind; });
      if (d.table) ar = ar.filter(function (x) { return x[4] === d.table; });
      return { ok: true, rows: ar.slice(-400).map(function (x) { return x.slice(0, 10); }), total: m - 1 };

    case 'auditverify':
      if (!can_(u, 'audit.view')) return { ok: false, err: 'forbidden' };
      var av = auditVerify_(); addLog_(u.id, r, 'نظام: تحقق من سلامة سجل التدقيق — ' + (av.ok ? 'سليم' : 'مخالفة عند السطر ' + av.brokenAt));
      return { ok: true, v: av };

    case 'bklist':
      if (!can_(u, 'backup.manage')) return { ok: false, err: 'forbidden' };
      p = sh_('BackupLog'); m = p.getLastRow();
      return { ok: true, rows: m > 1 ? p.getRange(Math.max(2, m - 59), 1, Math.min(60, m - 1), 9).getValues() : [], ext: !!String(setObj_(rd_('Settings')).set.bkExt || '').trim(), trig: ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'dailyBackup'; }) };

    case 'bknow':
      if (!can_(u, 'backup.manage')) return { ok: false, err: 'forbidden' };
      var bn = bkFull_('manual', u.id); addLog_(u.id, r, 'نظام: نسخة احتياطية يدوية — ' + (bn.ok ? 'نجاح' : 'فشل'));
      return bn;

    case 'bkdl': // تنزيل ملف Excel لنسخة مسجّلة في BackupLog فقط (base64)
      if (!can_(u, 'backup.manage')) return { ok: false, err: 'forbidden' };
      var xi = String(d.id || ''), okx = xi && rd_('BackupLog').some(function (z) { return z[8] === xi; });
      if (!okx) throw new Error('ملف Excel غير مسجّل في سجل النسخ');
      var xfile = DriveApp.getFileById(xi), xb = xfile.getBlob();
      addLog_(u.id, r, 'نظام: تنزيل نسخة Excel ' + xfile.getName());
      return { ok: true, name: xfile.getName(), b64: Utilities.base64Encode(xb.getBytes()) };

    case 'bkverify': // استعادة تجريبية غير مُتلِفة لنسخة محدّدة
      if (!can_(u, 'backup.manage')) return { ok: false, err: 'forbidden' };
      var vv = bkVerify_(String(d.id || '')); addLog_(u.id, r, 'نظام: فحص نسخة احتياطية ' + String(d.id).slice(0, 8) + ' — ' + (vv.ok ? 'سليمة' : vv.bad + ' مخالفة'));
      return { ok: true, v: vv };

    case 'bkrestore': // استعادة فعلية: تأكيد نصي + نسخة أمان تلقائية + استبدال الأوراق المختارة فقط
      if (!can_(u, 'backup.manage') || r !== 'admin') return { ok: false, err: 'forbidden' };
      if (String(d.confirm || '') !== 'استعادة') throw new Error('اكتب كلمة «استعادة» للتأكيد');
      var tb2 = (d.tables || []).map(String).filter(function (k) { return HD[k]; }); if (!tb2.length) throw new Error('اختر أوراقًا للاستعادة');
      var dn = bkRestore_(String(d.id || ''), tb2, u.id);
      addLog_(u.id, r, 'نظام: استعادة من نسخة ' + String(d.id).slice(0, 8) + ' للأوراق: ' + dn.join('، '));
      auditAdd_(u, [['استعادة', 'النسخ الاحتياطي', String(d.id).slice(0, 12), '', '', dn.join('، '), 'نسخة أمان سبقت الاستعادة']]);
      return { ok: true, restored: dn };

    case 'since': // جلب تزايدي (الموظف يحصل على بصماته فقط)
      p = sh_('Punches'); m = p.getLastRow();
      var ep = P_.getProperty('EPOCH') || '0', full = String(d.ep) !== ep, from = full ? 0 : (d.from | 0);
      rows = m - 1 > from ? p.getRange(2 + from, 1, m - 1 - from, 5).getValues() : [];
      if (r === 'emp') {
        var own = owner_();
        rows = rows.filter(function (x) { return own(String(x[2] || ''), x[0]) === u.empNo; });
      }
      return { ok: true, rows: rows, total: Math.max(m - 1, 0), ep: ep, full: full };

    case 'delpunch':
      var why = String(d.reason || '').trim().slice(0, 150);
      if (!why) return { ok: false, err: 'سبب الحذف مطلوب' };
      p = sh_('Punches'); m = p.getLastRow();
      if (m < 2) return { ok: true, deleted: 0 };
      var kill = {}, killAny = {}; (d.rows || []).forEach(function (x) { if (x[2] === undefined || x[2] === null) killAny[nid_(x[0]) + '|' + x[1]] = 1; else kill[nid_(x[0]) + '|' + x[1] + '|' + String(x[2])] = 1; });
      var v = p.getRange(2, 1, m - 1, 5).getValues(), gone = [], nw2 = nowS_();
      var keep = v.filter(function (x) { var k = nid_(x[0]) + '|' + x[1]; if (killAny[k] || kill[k + '|' + String(x[2] || '')]) { gone.push([nw2, nid_(x[0]), String(x[1]), String(x[2] || ''), 'حُذفت بواسطة ' + u.id + ': ' + why, String(x[3] || '')]); return false; } return true; });
      if (keep.length !== v.length) {
        p.getRange(2, 1, m - 1, 5).clearContent();
        setTxt_(p, 2, keep);
        var pl2 = sh_('PunchLog'); setTxt_(pl2, pl2.getLastRow() + 1, gone.slice(0, 500)); // الأصل يُؤرشف ولا يُمحى بلا أثر
        P_.setProperty('EPOCH', String(num_('EPOCH') + 1));
        addLog_(u.id, u.role, 'نظام: حذف ' + gone.length + ' حركة بصمة — السبب: ' + why);
      }
      return { ok: true, deleted: v.length - keep.length };

    case 'punchlog': // سجل الحركات المستبعدة (المكررة والمحذوفة) — آخر 300
      if (!staff) return { ok: false, err: 'forbidden' };
      p = sh_('PunchLog'); m = p.getLastRow();
      return { ok: true, rows: m > 1 ? p.getRange(Math.max(2, m - 299), 1, Math.min(300, m - 1), 6).getValues() : [] };

    case 'linkfp': // ربط رقم بصمة غير مطابق بموظف: إجراء موثّق بسبب ومنفّذ، لا تعديلًا صامتًا
      if (!can_(u, 'punches.link')) return { ok: false, err: 'forbidden' };
      var lf = nid_(d.fpId), lr = String(d.reason || '').trim().slice(0, 200), le = String(d.empNo || '').trim();
      if (!lf || !le || !lr) throw new Error('رقم البصمة والموظف وسبب الربط مطلوبة');
      var el = rd_('Employees'), me = el.filter(function (x) { return x[0] === le; })[0];
      if (!me) throw new Error('الموظف غير موجود');
      var ldv = String(d.dev || '').trim().slice(0, 60);
      if (ldv) { // ربط دقيق بجهاز محدد (الأفضل مع تعدد الأجهزة): صف في ورقة FpMap
        var fmr = rd_('FpMap'), cf = fmr.filter(function (x) { return x[2] === ldv && nid_(x[3]) === lf; })[0];
        if (cf) throw new Error('الرقم ' + lf + ' في الجهاز «' + ldv + '» مرتبط أصلًا بالموظف ' + cf[1]);
        var sf = sh_('FpMap'); setTxt_(sf, sf.getLastRow() + 1, [[Date.now(), le, ldv, lf, '', '', '', ('ربط من بصمات غير مطابقة: ' + lr).slice(0, 100)]]);
        P_.deleteProperty('H_FpMap');
        var sm0 = sh_('Movements'); setTxt_(sm0, sm0.getLastRow() + 1, [[Date.now(), le, 'تغيير رقم البصمة', '', lf + ' (' + ldv + ')', nowS_().slice(0, 10), '', 'ربط بصمة غير مطابقة: ' + lr, u.id, nowS_()]]);
        P_.deleteProperty('H_Movements');
        addLog_(u.id, u.role, 'نظام: ربط رقم البصمة ' + lf + ' في الجهاز ' + ldv + ' بالموظف ' + le + ' — السبب: ' + lr);
        return { ok: true, rev: bump_() };
      }
      var other = el.filter(function (x) { return x[0] !== le && (nid_(x[5]) === lf || nid_(x[10]) === lf || nid_(x[13]) === lf); })[0];
      if (other) throw new Error('رقم البصمة ' + lf + ' مرتبط أصلًا بالموظف ' + other[1] + ' (' + other[0] + ')');
      if (nid_(me[5]) === lf || nid_(me[10]) === lf || nid_(me[13]) === lf) throw new Error('الرقم ' + lf + ' مرتبط أصلًا بهذا الموظف');
      var slot = -1; [5, 10, 13].forEach(function (ix) { if (slot < 0 && !nid_(me[ix])) slot = ix; }); // أول خانة فارغة من الثلاث
      if (d.replace) slot = 5; // الاستبدال يطال الرقم الأول
      if (slot < 0) throw new Error('للموظف ثلاثة أرقام بصمة (' + me[5] + ' و ' + me[10] + ' و ' + me[13] + ')؛ أكّد الاستبدال');
      var se = sh_('Employees'), ca = se.getRange(2, 1, Math.max(se.getLastRow() - 1, 1), 1).getValues(), rowE = 0;
      for (var q = 0; q < ca.length; q++) if (String(ca[q][0]) === le) { rowE = q + 2; break; }
      if (!rowE) throw new Error('تعذّر تحديد صف الموظف');
      se.getRange(rowE, slot + 1).setNumberFormat('@').setValue(lf);
      P_.deleteProperty('H_Employees');
      var sm = sh_('Movements'); setTxt_(sm, sm.getLastRow() + 1, [[Date.now(), le, 'تغيير رقم البصمة', String(me[slot] || ''), lf, nowS_().slice(0, 10), '', 'ربط بصمة غير مطابقة: ' + lr, u.id, nowS_()]]);
      P_.deleteProperty('H_Movements');
      addLog_(u.id, u.role, 'نظام: ربط رقم البصمة ' + lf + ' بالموظف ' + le + (slot !== 5 ? ' (كرقم إضافي)' : me[5] ? ' (بدلًا من ' + me[5] + ')' : '') + ' — السبب: ' + lr);
      return { ok: true, rev: bump_() };
  }
  return { ok: false, err: 'unknown action' };
}
