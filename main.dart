import 'dart:convert';
import 'dart:io';

import 'package:excel/excel.dart';
import 'package:file_picker/file_picker.dart';
import 'package:flutter/material.dart';
import 'package:flutter/services.dart';
import 'package:flutter_localizations/flutter_localizations.dart';
import 'package:path/path.dart' as p;
import 'package:path_provider/path_provider.dart';
import 'package:permission_handler/permission_handler.dart';

// ======================= الإعدادات =======================
const String kDir = '/storage/emulated/0/A2R_Sender';
const String kFile = '$kDir/groups.xlsx';
const MethodChannel _ch = MethodChannel('a2r/share');

const Map<String, String> kKinds = {
  'text': 'نص فقط',
  'imageText': 'صورة مع نص',
  'sticker1': 'استيكر واحد',
  'sticker2': 'استيكران',
};
const Map<String, int> kLimit = {
  'text': 0,
  'imageText': 1,
  'sticker1': 1,
  'sticker2': 2,
};

// ======================= الجروبات وملف Excel =======================
// البيانات الأولية: تُكتب في groups.xlsx أول مرة فقط (الاسم|المعلمة). بعدها عدّلي الملف نفسه في الهاتف.
const String kSeed = r'''
سارة122/1914|جهاد السيد ابراهيم محمد
تغريد 112/1914|جهاد السيد ابراهيم محمد
موزة 648/649|مروة السيد
محمد524/554/555/667|هاجر الهواري
زياد 463/554/555/667|محمد إبراهيم أبو الخير
خالد 463/524/ 555/667|محمد إبراهيم أبو الخير
تيم 554/463/524/667|هاجر الهواري
الخليل 632/633/634|معاذ محمد
الخطاب 631/633/634|معاذ محمد
ريان 631/632/634|سهى أحمد
بيان 631/632/633|سهى أحمد
عمران 636|شيماء عبد الفتاح
عثمان 635|شيماء عبد الفتاح
أيوب 61|احمد محمود
أ/ مروة أم لولوة 628|شيماء طارق
خليل أحمد|اسلام صلاح
شموخ 641|ولاء ابو السعود
شهلاء 640|ولاء ابو السعود
وسيم محمد 863/953/1846|فايزة حجاج
أم سارة|فايزة حجاج
عبد العزيز 645|احمد المرشدي
يحيى 644|احمد المرشدي
أصيلة 647|ولاء ابو السعود
رحمة 647|ولاء ابو السعود
ناصر 649/392|مروة السيد
ميرا648/392|مروة السيد
محمد 651/1021/1022/1195|احمد محمود
مهرة 650/1021/1022/1195|سارة علي ربيع
مها عثمان 5774/5896|دعاء يوسف
سعاد|شيماء عبد الفتاح
بدرية|دينا احمد محمد ابراهيم
أم عمر|أسماء رضا
مريم|سارة عصام
يزن 658|نهاد السيد البحيري
سدرة 657|نهاد السيد البحيري
سلطانة|سهام محمد
حلا 661/662|فايزة حجاج
ميرا 660/662|فايزة حجاج
زينة و هيا 660/661|فايزة حجاج
لين أحمد 169/901|منار اشرف
ام رائف|االمعلم معتذر
محمد فهد|شيماء عبد الفتاح
موسى|سارة عصام
أم خالد 554/555/524/463|هناء عبدالله
عبدالرحمن محمد|االمعلم معتذر
مودة مصطفى 444|سارة علي ربيع
نوف راشد|االمعلم معتذر
دانية سلامة 672/673|ايمان عبد الحميد
دعاء سلامة 671/673|ايمان عبد الحميد
قطر الندى 671/672|ايمان عبد الحميد
تميم أحمد|نجلاء سعيد
كريم ربيع 4274|ندي خالد جمال
جنى 685|زينب ابو بكر
عبد الله على|اسلام صلاح
مناجد 679/680|االمعلم معتذر
مناضل 678/680|االمعلم معتذر
وقاد 678/679|االمعلم معتذر
ريتال|اسلام صلاح
سعود|أحمد نجاح
رفيف عمر 688/1420/1429|اسراء زقزوق
فاطمة فهد 847|هند ناصر
هشام 676|أحمد إسماعيل
علي|ايمان عبد الحميد
ديما|جهاد السيد ابراهيم محمد
فاطمة عمر 683/1420/1429|اسراء زقزوق
كنان عمار|االمعلم معتذر
رغد|سارة علي ربيع
علي 692/693/3288|نهاد السيد البحيري
شيخة 691/693/3288|نهاد السيد البحيري
عبد الله 691/692/3288|صلاح سعد
شيخة بنت خميس1829|مني سامي
يوسف سامي|فايزة حجاج
شما 697|االمعلم معتذر
عامر 696|االمعلم معتذر
عليا هشام|وفاء محمد
العز الهشيمي|ندى أحمد
فاطمة بنت سعيد 225/708|ايمان فؤاد مبروك
حميد ياسر 702|االمعلم معتذر
نور ياسر701|سمية فرحات
زهد 200|اسماء ابو مسلم
مريم 705/1507/5183|فايزة حجاج
يامن 704/15075183|فايزة حجاج
سعيد|عائشة طه
المزن 4415|ايمان عبد الحميد
ماريا 225/700|ايمان فؤاد مبروك
مريم 710/711|الزهراء محمد رفاعي
محمد 709/711|الزهراء محمد رفاعي
الغالية 709/710|الزهراء محمد رفاعي
علي|وفاء محمد
ديما 714/865|ريهام ربيع
ريما 713/865|ريهام ربيع
الزين|الزهراء محمد رفاعي
أحمد|محمد عصام المجدي
سلطانة795/796/659|عائشة طه
سارة 420/421|جهاد السيد ابراهيم محمد
علي سليمان|سمية فرحات
أم ألما 578/579|مني سامي
موسى1206/1221|عمرو محمد عبد الصمد
جواهر|دينا احمد محمد ابراهيم
علي 729|هبة أحمد
مصطفى 812|محمد نور
جويل 811|محمد نور
ا. مها 684|سارة عصام
رغد 849|شيماء عبد الفتاح
صالح 848|شيماء عبد الفتاح
وئام|منه الله عماد القلاوي
محمد|االمعلم معتذر
الزهراء 132/856|شيماء عبد الحميد
زكريا 132/853|هبة أحمد
سعيد 858/859|طه عبد العظيم
عبد الله 857/859|طه عبد العظيم
لجين 857/858|احمد المرشدي
محمد سالم|شيماء عبد الفتاح
مازن حمد 3611|نجلاء سعيد
الوئام|اسراء جاد
أم وسيم 642/953|فايزة حجاج
علي713/714|محمد نور
أحمد حسن صبحي 933|االمعلم معتذر
ميرا العيسري|االمعلم معتذر
ريهام|اسماء فوزي
ناصر 3431|شيماء عبد الفتاح
يعقوب 873|شيماء طارق
إبراهيم 872|شيماء طارق
مهدي علي|سماح حسن
محمد 876|داليا أنور
علي 875|داليا أنور
ريان878|هدي احمد حسين
ريناد 877|هدي احمد حسين
منى 882|عائشة طه
الحسن 881|عائشة طه
نجاة|االمعلم معتذر
ميرا 1001/1010|شيماء عبد الفتاح
فاطمة 881/882|ايمان نعمان
عمر|هدي احمد حسين
خليفة|سارة عصام
الجلندي 891/920|رنا محمد سعيد حسين
غادة 890/920|رنا محمد سعيد حسين
محمد|عبدالله أحمد عمارة
فياض897|أحمد جمال
غسان896|أحمد جمال
ثامر 899|وفاء محمد
خالد 898|وفاء محمد
ادم 169/663|منار اشرف
رواء|ايمان عبد الحميد
جواد المقداد|عبدالله أحمد عمارة
هادف|عبدالله أحمد عمارة
حور906/907|هدي احمد حسين
محمد905/907|هدي احمد حسين
علي905/906|هدي احمد حسين
جواهر 913|دينا احمد محمد ابراهيم
أية 908|دينا احمد محمد ابراهيم
دانا|االمعلم معتذر
أية|االمعلم معتذر
إبراهيم 917/918|االمعلم معتذر
علي 916/918|االمعلم معتذر
هدى 916/917|أسماء رضا
غيداء 890/891|ايمان نعمان
شهد|عائشة طه
هالة|ريهام حسين
جهاد 924/925|عبدالله أحمد عمارة
رتيل 923/925|سماح محمود
تميم923/924|سماح محمود
ميرا|االمعلم معتذر
نسرين|االمعلم معتذر
ياسين 936|عاليا محمد
مريم 935|عاليا محمد
ناصر|االمعلم معتذر
أحمد 886/1010|شيماء عبد الفتاح
فيصل 1001/886|شيماء عبد الفتاح
حمدة 1022/650/651/1195|سارة علي ربيع
دانة 1021/650/651/1195|سارة علي ربيع
نبع 65/3387|جهاد السيد ابراهيم محمد
شيماء1199/1200/1201|سارة صلاح
مضاوي1149/1150/1151/1152|الزهراء محمد رفاعي
رفيف1148/1150/1151/1152|الزهراء محمد رفاعي
سما1148/1149/1151/1152|الزهراء محمد رفاعي
جنات 1148/1149/1150/1152|ريهام ربيع
ملك 1148/1149/1150/1151|ريهام ربيع
سلمى|مني سامي
نفيسة|سارة صلاح
نادية 1150|نادية محمد
ياسمين|االمعلم معتذر
محمد|عبدالله أحمد عمارة
يمنى 1161|هدير الشاذلي
تالية 1160|هدير الشاذلي
سارة|رقية مصطفي فتحي
هاشم 1164|عائشة طه
سند 1163|عائشة طه
فاطمة 160|سارة عصام
آدم 1167/1168/1169/1170|ياسين الإمام
عمر 1166/1168/1169/1170|ياسين الإمام
سارة 1166/1167/1169/1170|اسلام صلاح
حمزة 1166/1167/1168/1170|احمد المرشدي
زكريا 1166/1167/1168/1169|احمد المرشدي
سعود|شيماء عبد الحميد
مروان 1184/5951|زينب أحمد
جوري 1183|زينب أحمد
معاش|ميرفت يوسف
عزان|احمد عبد الرحمن
آدم 1317|عوض مصطفى
ريان 1316|عوض مصطفى
فاطمة 1319|ايمان عبد الحميد
عبد العزيز 1318|ايمان عبد الحميد
ريما|اسماء عبد الناصر
محمد 1323|طه صالح
هنا 1322|برديس عزت
رحمة 1679/1680|سارة عصام
شيخة 1326/5564|مروة السيد
رحمة 1325/5564|مروة السيد
سلمان1569|ثروت محمد
جواد 1329|فرحة محمد
كرم1328|فرحة محمد
عمران1331|فاطمة الزهراء محمد حسنين
زين1330|فاطمة الزهراء محمد حسنين
ورد|هاجر مدكور
قيصر|هدي احمد حسين
نصرى|الشيماء بدوي محمد
بشرى|ناهد محمد
إسراء|ايمان نعمان
سلطان سعيد|عبدالرحمن رامون
أنس محمد 1534|االمعلم معتذر
علي محمد 1533|االمعلم معتذر
محمد 1536|خميس عبد الفتاح
عبد الرحمن 1535|خميس عبد الفتاح
هديل 5814/5815|نجاة عبد الستار بدوي
زايد علي 1723|أحمد العربي
آسية|ندي خالد جمال
فاطمة الفارسية|سارة عصام
ضي 2225|فاطمة الزهراء محمد حسنين
عبد الرحمن 1543/1568|عبدالرحمن رامون
رحمة 1542/1568|ندي خالد جمال
الخطاب 1545/1546|االمعلم معتذر
مسك1544/1546|بسملة عماد
هاجر1544/1545|فاطمة ابراهيم
أحمد معتز|معتصم حشاد
عائشة 1439|داليا علي
فاطمة 1327|ثروت محمد
منال|ايه عبد المحسن محمد
عبد الله 694|علي حجازي
وعد الحمدانية|نورة صابر
سدن محمد|ضحى صادق
سجى 1871|وفاء محمد
سلا 1870|وفاء محمد
أميمة 1963|دينا احمد محمد ابراهيم
ريم 1875|ريم محمد
مياسة 1874|ريم محمد
عمر 1877|محمد نور
عزة 1876|مريم الجمال
سما 1879|رقية منصور
عمرو 1878|صلاح سعد
سارة سلطان 1881|سهى أحمد
هاجر سلطان 1880|سهى أحمد
ريم الحارثي|اسلام صلاح
تركية|شيماء طارق
راشد حمد|حسن نجم
ميرا محمد 1886/1887/1894/2944|ريهام ربيع
ماسة محمد 1885/1887/1894/2944|ريهام ربيع
عمر محمد 1885/1886/1894|أحمد جمال
عبد الله 1891|فاطمة محمد عبد الجواد
سعيد 1890|فاطمة محمد عبد الجواد
ولاء شحادة 1885/1886/1887|نورة صابر
جنى خالد 2027|أسماء مصطفى
خليفة خالد 2026|أسماء مصطفى
سيدة|أسماء رضا
خالد بن الوليد 2091|أحمد عادل
راشد بن الوليد 2090|أحمد عادل
محمد وليد 2093/2094/2787/4202|شيماء عبد الفتاح
ريم وليد 2092/20944202|شيماء عبد الفتاح
نور وليد 2092/2093/4202|شيماء عبد الفتاح
هارون قنديل 2096|محمود صلاح
موسى قنديل 2095|محمود صلاح
كريم ياسين|إيمان خالد
رائد 2099/2100|رحمة هريدي
ماجد 2098|رحمة هريدي
أويس 2098|رحمة هريدي
منصور مروان 2102|أحمد جمال
مايد مروان 2101|أحمد جمال
راشد أيمن 2104|أحمد إسماعيل
رفيف أيمن 2103|ندى أحمد
كرم 2106|محمد السيد مرجان
ماسة 2105|إيمان احمد
كريم خليفة 2108|رنا محمد سعيد حسين
كاميليا خليفة 2107|رنا محمد سعيد حسين
أحمد|محمد شعبان
أكرم ثابت|محمود عبده
عبد الله 2117/2118|أحمد إسماعيل
فارس 2116/2118|أحمد إسماعيل
ليلاس|ندى أحمد
راشد|عبدالله أحمد عمارة
عاطف|أحمد عصران
علي حلمي|أحمد إسماعيل
سعيد هشام|اسلام صلاح
ميرة أحمد|أسماء رضا
ريهام مختار|ولاء سمير
يارا مهند 2330|داليا علي
لانا مهند 2331|داليا علي
أم عبدالله العامري|سومة ناجي
عهود صالح|ميرفت نعمان
سارة عبد الرحمن|جهاد رشاد
حسن|عبدالله أحمد عمارة
أمل|فاطمة محمد عبد الجواد
محمد|جهاد السيد ابراهيم محمد
شهم و عبد الله و عبد الملك|محمد عصام المجدي
قيس 2362|عبد الله القاضي
تيمور 2361|عبد الله القاضي
رزان عبد الله|ايمان نعمان
أجناس|ميرفت يوسف
سارة محمد|هند ناصر
طلال الحمصي|مصطفى محمد
مسعود 2549/2550|محمد نور
عز 2548/2550|محمد نور
أنسام 2548/2549|دعاء يوسف
سلمان علي|هند ناصر
يوسف خالد|مصطفى محمد
إيمان|عائشة طه
أمنية|شروق عبد المنعم
سلمى محمد|سارة مصطفى
علي أمجد 2606/2607|مصطفى عبد الحميد
محمد أمجد 2605/2607|مصطفى عبد الحميد
ليان أمجد 2605/2606|نرمين رجب
وداد|ولاء سمير
عبد الله 427|عبد الله الأديب
مجد ماجد|هاجر مدكور
عمر حازم|أية حسين
زينة|نرمين رجب
كيان سعيد 2973/2985|نرمين رجب
ليلى و زين سامر|ساره محمد فرهود
نايا سائد 2617|فاطمة محمد
أحمد سائد 2616|فاطمة محمد
عبد العزيز 2999|أحمد نجاح
يحيى محمود|زينب صبري
فارس صبح|منة الشافعي
أم ميثاء|هبة حسين
حمزة محمد|هدى إبراهيم الخولي
زايد|إيمان الجندي
فرح|هدى إبراهيم الخولي
ألمى علاء|نشوى سعيد
نواف سالم|سهى أحمد
مريم مجيد 2656|حنان صابر
زينه 2630|نرمين رجب
ماسه 2629|نرمين رجب
ريان عبد العزيز|منة الشافعي
اريج 2662|نرمين كارم
نيرفان إبراهيم 2646|عائشة طه
مريم إبراهيم 2645|عائشة طه
مها 2628|حنان صابر
مايد|إيمان الجندي
شيخه|أية حسين
عبد الله 2701/2702|محمد عز الدين
فهد 2700/2702|ولاء سمير
غالية 2700/2701|هاجر فتحي
ورد 2725|ايمان رشدي
جوليا 2724|ايمان رشدي
سند محمد 2727|رحمة زكريا
ماسة محمد 2726|رحمة زكريا
يامن 2730/3882/3883|سمية ياسر
ماريا 2729/3882/3883|سمية ياسر
يوسف|ولاء سمير
الايهم نبهان|فهيمة الدسوقي
شيخة حمد|ايمان عز الدين كامل
يعقوب طه|أسماء رضا
أمل 2092/2093/2094/4202|شيماء عبد الفتاح
شادن سعيد 2613/2985|نرمين رجب
ليان 2618|ندي مجدي
وتين عدي|دعاء علي
خالد عابد|سماح محمود
ديمه 961/962/963|نهاد السيد البحيري
داوود 65/1146|امال عادل
عمر 3604/3605|الشناوي احمد
ميرا 3603/3605|شيماء عبد الفتاح
دان 3603/3604|شيماء عبد الفتاح
فاطمه|مريم طارق
مريم المعمري|ميرفت نعمان
سيلين/2729/2730|مروه احمد
عبد الله 3948|زينب صبري
فهد 3947|زينب صبري
تولين|رحاب صبري
دانا 3951|هاجر أحمد
ديما 3950|هاجر أحمد
رؤى|شيماء طارق
سلمي|لوزة رمضان المصري
ميرا 4309|رنيم احمد
غزل|هاجر عماد
ريان|عبد الله الأديب
رنا|دينا احمد محمد ابراهيم
ريان|هاجر عماد
شريفه|زينب أحمد
الاء 3966/3967|هاجر اشرف
الريم 3965/3967|هاجر اشرف
طارق 3965/3966|محمد سعد
مهره4140|ساره زكي
أحمد 3970|سلمي طارق
اديب 3969|سلمي طارق
خديجه|فاطمة صلاح
حمزه|ايمان رشدي
رضا محمد 3977/4368/4369|فاطمة محمد عبد الجواد
عائشه احمد 3976/4368/4369|فاطمة محمد عبد الجواد
يحي 3981|ندي مجدي
يزن 3980|ندي مجدي
سند|رنيم احمد
محمد|شيماء محمود
فاطمه|شيماء طارق
وسام 3992|عبد الله رضوان
سامر 3991|عبد الله رضوان
الاء|رنا محمد سعيد حسين
محمود ابو كشك|علي حجازي
زينه عريبات|هاجر عماد
تالا 4005|سميه عبد الفتاح
محمد 4004|سميه عبد الفتاح
راشد 3968|إيمان احمد
نور تصحيح تلاوه 2787/2092/2093|سميه عبد الفتاح
عبد الحميد 675|علي حجازي
احمد 4369/3976/3977|فاطمة محمد عبد الجواد
عبيد 4368/3976/3977|فاطمة محمد عبد الجواد
الحارث 707|ايمان عبد الحميد
مهند فيصل|زينب أحمد
مايد|جهاد السيد ابراهيم محمد
لطيفه|جهاد السيد ابراهيم محمد
عبد الله|جهاد السيد ابراهيم محمد
سري 2619|زينب صبري
يوسف 2645/2646|عائشة طه
ناديه 3969/3970|رباب صلاح
سما 705/704/1507|فايزة حجاج
ساره|شيماء عبد الفتاح
ميرا|امل حبيب
نور 5226|سميه عبد الفتاح
عمرو 5225|بسام التابعي
عمر 5228|عبد الله أمر الله
سيلين 5227|امال عادل
محمد|دعاء أحمد
افين|ايه نادي حسن
قمر 5232|هاجر ياسر
طلال 5231|هاجر ياسر
بشري|ميرفت يوسف
صدام|شروق عبد المنعم
ليلي 5241|سهى أحمد
يوسف 5240|سهى أحمد
محمد|حكمة الله مجدي
عمر|اميره سعيد
عصام 5860|احمد مسعد
عز|سميه عبد الفتاح
حسان 5248/5249|دعاء علي
ايه 5247/5249|دعاء علي
مريم 5247/5248|دعاء علي
بدر|حسام احمد
سلمي 5252|هاجر مصطفي
علي 5251|هاجر مصطفي
ردينه5258/5259/5260/6013|هبه محمود
ورده 5257/5259/5260|سمية ياسر
اسيد 5257/5258/5260|سمية ياسر
عدي 5257/5258/5259|سمية ياسر
عبد الله|عبد الرحمن هلال
دلع 5168/5235/5236/5264|نيره عبد العزيز
جني 5168/5235/5236/5263|نيره عبد العزيز
مؤمن|صبحي الجندي
سلامه4413/5267|احمد نوفل
تيم 4413/5266|احمد نوفل
كلثم 5269|سلمي طارق
خلود 5268|سلمي طارق
اكرم|ايه محمود
محمد 5279|دعاء سعيد
اياد 5278|دعاء سعيد
موده 1325/1326|مروة السيد
سميه 5245|إسراء أحمد
مريم 652|مريم عصام ابو شادي
عبير 1183/1184|زينب أحمد
ردينه 5257|نيره عبد العزيز
همس|جهاد السيد ابراهيم محمد
مرزوق|خالد مصطفي التركي
اليمان 635/636|شيماء عبد الفتاح
مريم 648/649|مروة السيد
نواف 6279|احمد المرشدي
خالد 6278|احمد المرشدي
عبد الرحمن|عبدالرحمن الخولي
سلاف 6297|جهاد السيد ابراهيم محمد
راشد 6369/6370|مروه عصام الصبابطي
ادم 6368/6370|مروه عصام الصبابطي
الما 6368/6369|مروه عصام الصبابطي
حمزه|احمد منصور
يمني|مشيره نعمان
حمد 2564|شروق عبد المنعم
دانه 6676|جهاد السيد ابراهيم محمد
جود 6675|جهاد السيد ابراهيم محمد
عبده 4459|زينب أحمد
عبد الله 6740/6741|عادل عبد الرحمن
محمد 6739/6741|عادل عبد الرحمن
مهره 6739/6740|شروق سمير الشاذلي
احمد 6776/6777/6778|محمود نصر الله
مروان 6775/6777/6778|محمود نصر الله
نوف 6775/6776/6778|الاء فتحي
ريم 6775/6776/6777|الاء فتحي
سلمي 5860/5245|احسان عاشور
يارا 112/122|جهاد السيد ابراهيم محمد
شهد 7083/7084|مروه علي
لمي 7082/7084|اسماء عبد العليم
ناصر 7082/7083|محمد سعد
هنا|هاجر فتحي
كوثر|احسان عاشور
عمر|محمد عبد المنعم
تهاني|ايمان نعمان
ساره|مريم عصام ابو شادي
امال|فوزيه فليفل
ميلا|نسرين محمد
اسامه|زينب ابو بكر
مسك|سميه عبد الفتاح
عزان7112/7113|ايمان عز الدين كامل
رغد 7111/7113|ايمان عز الدين كامل
رنا 7111/7112|ايمان عز الدين كامل
فاطمة السعيدي 2507|شيماء عادل
ابراهيم 7292|ريهام ايهاب
شاهه|رشا سمير
الهنوف 2114|سميه محمد
سيف 7102|عيسي محمد
الجورى 2114/7156/7281|سميه محمد
اليازيه2114/7156/7280|سميه محمد
يارا 1153|دعاء هاني
محمد 7116|فردوس عبد العال
مجتبي 4202/2787/2092|سميه عبد الفتاح
شهد|ايه محمود
بيان 7116/7292|ريهام ايهاب
سلطان|حسام حسن
''';

class G {
  final String name, numbers, teacher;
  const G(this.name, this.numbers, this.teacher);
  String get key => '$name|$numbers';
  String get title => numbers.isEmpty ? name : '$name  $numbers';

  /// كلمة البحث في واتساب: أول رقم، وإلا الاسم.
  String get searchKey {
    for (final n in numbers.split('/')) {
      if (n.trim().isNotEmpty) return n.trim();
    }
    return name;
  }

  Map<String, String> toJson() =>
      {'name': name, 'numbers': numbers, 'teacher': teacher};
  factory G.fromJson(Map<String, dynamic> j) => G(
      j['name'] as String, j['numbers'] as String, j['teacher'] as String);
}

String normDigits(String s) {
  const ar = '٠١٢٣٤٥٦٧٨٩', fa = '۰۱۲۳۴۵۶۷۸۹';
  final b = StringBuffer();
  for (final r in s.runes) {
    final ch = String.fromCharCode(r);
    var i = ar.indexOf(ch);
    if (i < 0) i = fa.indexOf(ch);
    b.write(i >= 0 ? '$i' : ch);
  }
  return b.toString();
}

String _clean(String s) => normDigits(s)
    .replaceAll(RegExp(r'[\u200e\u200f\u202a-\u202e]'), '')
    .replaceAll(RegExp(r'\s+'), ' ')
    .trim();

final RegExp _tail = RegExp(r'^(.*?)\s*(\d[\d\s/]*)$');

/// الأعمدة: A = الاسم مع الأرقام (محمد524/554)، B أو C = المعلمة، D = "إيقاف" (أي كلمة تعطّل الصف).
G? parseRow(List<String> c) {
  if (c.isEmpty) return null;
  final raw = _clean(c[0]);
  if (raw.isEmpty) return null;
  if (c.length > 3 && _clean(c[3]).isNotEmpty) return null; // موقوف
  var name = raw, nums = '';
  final m = _tail.firstMatch(raw);
  if (m != null && m.group(1)!.trim().isNotEmpty) {
    name = m.group(1)!.trim();
    nums = m
        .group(2)!
        .split('/')
        .map((e) => e.replaceAll(RegExp(r'[^0-9]'), ''))
        .where((e) => e.isNotEmpty)
        .join('/');
  }
  var teacher = '';
  for (final x in c.skip(1).take(2)) {
    if (_clean(x).isNotEmpty) {
      teacher = _clean(x);
      break;
    }
  }
  return G(name, nums, teacher);
}

({List<G> groups, int dups, int off}) parseSheet(List<List<String>> rows) {
  final seen = <String>{};
  final out = <G>[];
  var dups = 0, off = 0;
  for (var i = 0; i < rows.length; i++) {
    final r = rows[i];
    if (i == 0 && r.isNotEmpty && r[0].contains('العميل')) continue;
    final g = parseRow(r);
    if (g == null) {
      if (r.isNotEmpty && _clean(r[0]).isNotEmpty) off++;
      continue;
    }
    if (!seen.add(g.key)) {
      dups++;
      continue;
    }
    out.add(g);
  }
  return (groups: out, dups: dups, off: off);
}

List<List<String>> readRows(String path) {
  final book = Excel.decodeBytes(File(path).readAsBytesSync());
  if (book.tables.isEmpty) return [];
  final sheet = book.tables[book.tables.keys.first]!;
  return sheet.rows
      .map((r) => r.map((c) => c?.value?.toString() ?? '').toList())
      .toList();
}

Future<bool> _tryPrepare() async {
  try {
    final d = Directory(kDir);
    if (!d.existsSync()) d.createSync(recursive: true);
    final f = File(kFile);
    if (!f.existsSync()) {
      final book = Excel.createExcel();
      final sh = book[book.tables.keys.first];
      sh.appendRow(<CellValue?>[
        TextCellValue('إسم العميل'),
        null,
        TextCellValue('اسم المعلمة/المعلم'),
        TextCellValue('إيقاف (اكتبي أي كلمة لتعطيل الصف)'),
      ]);
      for (final line in kSeed.split('\n')) {
        if (line.trim().isEmpty) continue;
        final i = line.indexOf('|');
        sh.appendRow(<CellValue?>[
          TextCellValue(i < 0 ? line : line.substring(0, i)),
          null,
          TextCellValue(i < 0 ? '' : line.substring(i + 1)),
        ]);
      }
      final bytes = book.encode();
      if (bytes == null) return false;
      f.writeAsBytesSync(bytes);
    }
    return true;
  } catch (_) {
    return false;
  }
}

/// ينشئ المجلد وينسخ ملف Excel أول مرة فقط. يرجع نص الخطأ أو null.
Future<String?> prepareFolder() async {
  if (!await Permission.storage.isGranted &&
      !await Permission.manageExternalStorage.isGranted) {
    await Permission.storage.request();
    if (!await Permission.storage.isGranted &&
        !await Permission.manageExternalStorage.isGranted) {
      await Permission.manageExternalStorage.request();
    }
  }
  if (await _tryPrepare()) return null;
  return 'لا أستطيع الوصول إلى المجلد $kDir\nاسمحي للتطبيق بالوصول إلى الملفات من إعدادات الهاتف ثم اضغطي تحديث.';
}

// ======================= الحفظ المحلي (المسودة + تقدم المهمة) =======================
class Store {
  static Map<String, dynamic> data = {};

  static Future<File> _file() async =>
      File(p.join((await getApplicationSupportDirectory()).path, 'state.json'));

  static Future<void> load() async {
    try {
      final f = await _file();
      if (f.existsSync()) {
        data = jsonDecode(await f.readAsString()) as Map<String, dynamic>;
      }
    } catch (_) {
      data = {};
    }
  }

  static Future<void> save() async {
    final f = await _file();
    await f.writeAsString(jsonEncode(data));
  }

  static Map<String, dynamic>? get job {
    final j = data['job'];
    if (j is Map<String, dynamic> && j['finished'] != true) return j;
    return null;
  }
}

Future<String> persistMedia(String src) async {
  final dir = Directory(
      p.join((await getApplicationSupportDirectory()).path, 'media'));
  await dir.create(recursive: true);
  final dest = p.join(dir.path,
      '${DateTime.now().microsecondsSinceEpoch}${p.extension(src).toLowerCase()}');
  await File(src).copy(dest);
  return dest;
}

void snack(BuildContext c, String m) => ScaffoldMessenger.of(c)
  ..hideCurrentSnackBar()
  ..showSnackBar(SnackBar(content: Text(m)));

Future<bool> ask(BuildContext c, String title, String body,
    {String ok = 'متابعة'}) async {
  final r = await showDialog<bool>(
    context: c,
    builder: (ctx) => AlertDialog(
      title: Text(title),
      content: Text(body),
      actions: [
        TextButton(
            onPressed: () => Navigator.pop(ctx, false),
            child: const Text('إلغاء')),
        FilledButton(onPressed: () => Navigator.pop(ctx, true), child: Text(ok)),
      ],
    ),
  );
  return r ?? false;
}

// ======================= التطبيق =======================
void main() {
  WidgetsFlutterBinding.ensureInitialized();
  runApp(const App());
}

class App extends StatelessWidget {
  const App({super.key});
  @override
  Widget build(BuildContext context) {
    const big = Size.fromHeight(56);
    return MaterialApp(
      title: 'A2R Sender',
      debugShowCheckedModeBanner: false,
      locale: const Locale('ar'),
      supportedLocales: const [Locale('ar')],
      localizationsDelegates: const [
        GlobalMaterialLocalizations.delegate,
        GlobalWidgetsLocalizations.delegate,
        GlobalCupertinoLocalizations.delegate,
      ],
      theme: ThemeData(
        useMaterial3: true,
        colorSchemeSeed: const Color(0xFF1B7F5C),
        filledButtonTheme:
            FilledButtonThemeData(style: FilledButton.styleFrom(minimumSize: big)),
        outlinedButtonTheme: OutlinedButtonThemeData(
            style: OutlinedButton.styleFrom(minimumSize: big)),
      ),
      home: const Home(),
    );
  }
}

// ======================= شاشة الرسالة =======================
class Home extends StatefulWidget {
  const Home({super.key});
  @override
  State<Home> createState() => _HomeState();
}

class _HomeState extends State<Home> with WidgetsBindingObserver {
  final TextEditingController text = TextEditingController();
  String kind = 'text';
  List<String> media = [];
  List<G> groups = [];
  int dups = 0, off = 0;
  String? err;
  bool loading = true;

  @override
  void initState() {
    super.initState();
    WidgetsBinding.instance.addObserver(this);
    _init();
  }

  @override
  void dispose() {
    WidgetsBinding.instance.removeObserver(this);
    text.dispose();
    super.dispose();
  }

  @override
  void didChangeAppLifecycleState(AppLifecycleState s) {
    // بعد تعديل Excel في تطبيق آخر والرجوع هنا، يُعاد قراءة الملف تلقائياً.
    if (s == AppLifecycleState.resumed && !loading) _reload();
  }

  Future<void> _init() async {
    await Store.load();
    final d = Store.data['draft'];
    if (d is Map<String, dynamic>) {
      kind = (d['kind'] as String?) ?? 'text';
      if (!kKinds.containsKey(kind)) kind = 'text';
      text.text = (d['text'] as String?) ?? '';
      media = ((d['media'] as List?) ?? const [])
          .map((e) => e.toString())
          .where((x) => File(x).existsSync())
          .toList();
    }
    await _reload();
  }

  Future<void> _reload() async {
    var e = await prepareFolder();
    var g = <G>[];
    var dp = 0, of = 0;
    if (e == null) {
      try {
        final r = parseSheet(readRows(kFile));
        g = r.groups;
        dp = r.dups;
        of = r.off;
      } catch (x) {
        e = 'تعذر قراءة groups.xlsx: $x';
      }
    }
    if (!mounted) return;
    setState(() {
      groups = g;
      dups = dp;
      off = of;
      err = e;
      loading = false;
    });
  }

  void _saveDraft() {
    Store.data['draft'] = {'kind': kind, 'text': text.text, 'media': media};
    Store.save();
  }

  void _setKind(String k) {
    setState(() {
      kind = k;
      final lim = kLimit[k]!;
      if (media.length > lim) media = media.sublist(0, lim);
    });
    _saveDraft();
  }

  Future<void> _pick() async {
    final room = kLimit[kind]! - media.length;
    if (room <= 0) return;
    final res = await FilePicker.platform
        .pickFiles(type: FileType.image, allowMultiple: room > 1);
    if (res == null) return;
    for (final f in res.files.take(room)) {
      if (f.path != null) media.add(await persistMedia(f.path!));
    }
    if (!mounted) return;
    setState(() {});
    _saveDraft();
  }

  String? _validate() {
    if (kind == 'text' && text.text.trim().isEmpty) return 'اكتبي نص الرسالة';
    final lim = kLimit[kind]!;
    if (lim > 0 && media.length != lim) {
      return lim == 2 ? 'اختاري استيكرين' : 'اختاري ملفاً واحداً';
    }
    return null;
  }

  Future<void> _start() async {
    final v = _validate();
    if (v != null) {
      snack(context, v);
      return;
    }
    if (groups.isEmpty) {
      snack(context, 'لا توجد جروبات في ملف Excel');
      return;
    }
    final old = Store.job != null;
    final ok = await ask(context, 'بدء الإرسال',
        '${kKinds[kind]} إلى ${groups.length} جروب.${old ? '\nسيتم إغلاق المهمة غير المنتهية السابقة.' : ''}',
        ok: 'ابدأ');
    if (!ok || !mounted) return;
    Store.data['job'] = {
      'kind': kind,
      'text': kind == 'sticker1' || kind == 'sticker2' ? '' : text.text.trim(),
      'media': kind == 'text' ? <String>[] : media,
      'groups': groups.map((g) => g.toJson()).toList(),
      'st': List<String>.filled(groups.length, 'pending'),
      'finished': false,
    };
    _saveDraft();
    await Store.save();
    if (!mounted) return;
    await Navigator.push(
        context, MaterialPageRoute(builder: (_) => const SendScreen()));
    if (mounted) setState(() {});
  }

  @override
  Widget build(BuildContext context) {
    final job = Store.job;
    final lim = kLimit[kind]!;
    return Scaffold(
      appBar: AppBar(title: const Text('A2R Sender'), actions: [
        IconButton(
            tooltip: 'تحديث من ملف Excel',
            onPressed: _reload,
            icon: const Icon(Icons.refresh)),
      ]),
      body: ListView(padding: const EdgeInsets.all(16), children: [
        if (loading) const LinearProgressIndicator(),
        if (err != null)
          Card(
              color: Theme.of(context).colorScheme.errorContainer,
              child: Padding(padding: const EdgeInsets.all(12), child: Text(err!)))
        else
          Card(
            child: Padding(
              padding: const EdgeInsets.all(12),
              child: Text('${groups.length} جروب جاهز'
                  '${dups > 0 ? ' • مكرر تم تجاهله: $dups' : ''}'
                  '${off > 0 ? ' • موقوف: $off' : ''}\n'
                  'الملف: $kFile'),
            ),
          ),
        if (job != null) ...[
          const SizedBox(height: 8),
          Card(
            child: Padding(
              padding: const EdgeInsets.all(12),
              child: Column(children: [
                Text(
                    'مهمة غير منتهية: ${(job['st'] as List).where((s) => s != 'pending').length} من ${(job['st'] as List).length}'),
                const SizedBox(height: 8),
                FilledButton(
                    onPressed: () async {
                      await Navigator.push(context,
                          MaterialPageRoute(builder: (_) => const SendScreen()));
                      if (mounted) setState(() {});
                    },
                    child: const Text('استكمال الإرسال')),
              ]),
            ),
          ),
        ],
        const SizedBox(height: 8),
        for (final k in kKinds.keys)
          ListTile(
            selected: kind == k,
            leading: Icon(kind == k
                ? Icons.radio_button_checked
                : Icons.radio_button_off),
            title: Text(kKinds[k]!,
                style: Theme.of(context).textTheme.titleMedium),
            onTap: () => _setKind(k),
          ),
        const Divider(),
        if (kind == 'text' || kind == 'imageText') ...[
          const SizedBox(height: 8),
          TextField(
            controller: text,
            minLines: 4,
            maxLines: 10,
            onChanged: (_) => _saveDraft(),
            decoration: const InputDecoration(
                labelText: 'نص الرسالة', border: OutlineInputBorder()),
          ),
        ],
        if (lim > 0) ...[
          const SizedBox(height: 16),
          Wrap(spacing: 12, runSpacing: 12, children: [
            for (var i = 0; i < media.length; i++)
              Stack(children: [
                ClipRRect(
                  borderRadius: BorderRadius.circular(12),
                  child: Image.file(File(media[i]),
                      width: 110,
                      height: 110,
                      fit: BoxFit.cover,
                      cacheWidth: 250,
                      errorBuilder: (_, __, ___) => const SizedBox(
                          width: 110,
                          height: 110,
                          child: Icon(Icons.broken_image))),
                ),
                PositionedDirectional(
                  top: 0,
                  end: 0,
                  child: IconButton.filled(
                    iconSize: 18,
                    onPressed: () {
                      setState(() => media.removeAt(i));
                      _saveDraft();
                    },
                    icon: const Icon(Icons.close),
                  ),
                ),
              ]),
          ]),
          const SizedBox(height: 12),
          if (media.length < lim)
            OutlinedButton.icon(
              onPressed: _pick,
              icon: const Icon(Icons.add_photo_alternate),
              label: Text('اختيار من الهاتف (${media.length}/$lim)'),
            ),
        ],
        const SizedBox(height: 24),
        FilledButton.icon(
            onPressed: loading ? null : _start,
            icon: const Icon(Icons.play_arrow),
            label: const Text('بدء الإرسال للجروبات')),
      ]),
    );
  }
}

// ======================= شاشة الإرسال =======================
class SendScreen extends StatefulWidget {
  const SendScreen({super.key});
  @override
  State<SendScreen> createState() => _SendScreenState();
}

class _SendScreenState extends State<SendScreen> {
  late Map<String, dynamic> job;
  int? override;
  bool opened = false;

  @override
  void initState() {
    super.initState();
    job = Store.job!;
  }

  List get _st => job['st'] as List;
  List get _groups => job['groups'] as List;
  G _g(int i) => G.fromJson(_groups[i] as Map<String, dynamic>);
  int _count(String s) => _st.where((x) => x == s).length;

  int? get cur {
    if (override != null) return override;
    final i = _st.indexOf('pending');
    return i < 0 ? null : i;
  }

  Future<void> _open(int i) async {
    final media = (job['media'] as List).map((e) => e.toString()).toList();
    for (final m in media) {
      if (!File(m).existsSync()) {
        snack(context, 'ملف الصورة غير موجود، جهزي الرسالة من جديد');
        return;
      }
    }
    await Clipboard.setData(ClipboardData(text: _g(i).searchKey));
    try {
      await _ch.invokeMethod<bool>(
          'share', {'text': job['text'] as String, 'paths': media});
      if (mounted) setState(() => opened = true);
    } on PlatformException catch (e) {
      if (!mounted) return;
      snack(
          context,
          e.code == 'NO_WA_BUSINESS'
              ? 'WhatsApp Business غير مثبت على هذا الهاتف'
              : 'تعذر فتح واتساب: ${e.message}');
    }
  }

  Future<void> _mark(int i, String s) async {
    _st[i] = s;
    await Store.save();
    setState(() {
      override = null;
      opened = false;
    });
  }

  Future<void> _jump(int i) async {
    if (_st[i] == 'done') {
      final ok = await ask(context, 'إعادة الإرسال',
          'تم تأكيد الإرسال لهذا الجروب سابقاً. إرسال مرة أخرى؟',
          ok: 'إعادة');
      if (!ok) return;
      _st[i] = 'pending';
      await Store.save();
    }
    if (mounted) setState(() => override = i);
  }

  Future<void> _finish() async {
    final ok = await ask(context, 'إنهاء المهمة',
        'المتبقي: ${_count('pending')} جروب. سيتم إغلاق المهمة.',
        ok: 'إنهاء');
    if (!ok) return;
    job['finished'] = true;
    await Store.save();
    if (mounted) Navigator.pop(context);
  }

  Icon _icon(String s) => s == 'done'
      ? const Icon(Icons.check_circle, color: Colors.green)
      : s == 'skipped'
          ? const Icon(Icons.skip_next, color: Colors.orange)
          : const Icon(Icons.radio_button_unchecked);

  void _list() {
    showModalBottomSheet<void>(
      context: context,
      isScrollControlled: true,
      builder: (ctx) => DraggableScrollableSheet(
        expand: false,
        initialChildSize: 0.8,
        builder: (_, sc) => ListView.builder(
          controller: sc,
          itemCount: _groups.length,
          itemBuilder: (_, i) {
            final g = _g(i);
            return ListTile(
              leading: _icon(_st[i] as String),
              title: Text(g.title),
              subtitle: g.teacher.isEmpty ? null : Text(g.teacher),
              onTap: () {
                Navigator.pop(ctx);
                _jump(i);
              },
            );
          },
        ),
      ),
    );
  }

  @override
  Widget build(BuildContext context) {
    final total = _groups.length;
    final done = _count('done'), skipped = _count('skipped');
    final i = cur;
    final t = Theme.of(context);
    final media = (job['media'] as List).map((e) => e.toString()).toList();
    final kind = job['kind'] as String;
    final body = (job['text'] as String).trim();

    return Scaffold(
      appBar: AppBar(title: const Text('الإرسال'), actions: [
        IconButton(
            tooltip: 'قائمة الجروبات',
            onPressed: _list,
            icon: const Icon(Icons.list)),
        IconButton(
            tooltip: 'إنهاء المهمة',
            onPressed: _finish,
            icon: const Icon(Icons.stop_circle_outlined)),
      ]),
      body: ListView(padding: const EdgeInsets.all(16), children: [
        LinearProgressIndicator(
            value: total == 0 ? 0 : (done + skipped) / total, minHeight: 8),
        const SizedBox(height: 8),
        Text(
            'تم $done  •  تخطي $skipped  •  متبقي ${_count('pending')}  (من $total)'),
        const SizedBox(height: 12),
        if (i == null) ...[
          const SizedBox(height: 24),
          const Icon(Icons.check_circle, size: 72, color: Colors.green),
          const SizedBox(height: 12),
          Text('اكتمل المرور على كل الجروبات',
              textAlign: TextAlign.center, style: t.textTheme.titleLarge),
          if (skipped > 0)
            const Padding(
              padding: EdgeInsets.only(top: 8),
              child: Text('يمكنك العودة للجروبات المتخطاة من أيقونة القائمة.',
                  textAlign: TextAlign.center),
            ),
          const SizedBox(height: 24),
          FilledButton(onPressed: _finish, child: const Text('إنهاء المهمة')),
        ] else ...[
          Card(
            child: Padding(
              padding: const EdgeInsets.all(16),
              child: Column(
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    Text('الجروب الحالي', style: t.textTheme.labelMedium),
                    const SizedBox(height: 4),
                    Text(_g(i).name, style: t.textTheme.headlineSmall),
                    if (_g(i).numbers.isNotEmpty)
                      Text(_g(i).numbers, style: t.textTheme.titleMedium),
                    if (_g(i).teacher.isNotEmpty)
                      Text('المعلمة/المعلم: ${_g(i).teacher}'),
                    const Divider(height: 24),
                    Row(children: [
                      Expanded(
                          child: Text('ابحثي في واتساب عن:  ${_g(i).searchKey}',
                              style: t.textTheme.titleMedium)),
                      IconButton(
                        tooltip: 'نسخ',
                        icon: const Icon(Icons.copy),
                        onPressed: () async {
                          await Clipboard.setData(
                              ClipboardData(text: _g(i).searchKey));
                          if (context.mounted) snack(context, 'تم النسخ');
                        },
                      ),
                    ]),
                  ]),
            ),
          ),
          const SizedBox(height: 8),
          Card(
            child: Padding(
              padding: const EdgeInsets.all(12),
              child: Row(children: [
                for (final m in media.take(2))
                  Padding(
                    padding: const EdgeInsetsDirectional.only(end: 8),
                    child: Image.file(File(m),
                        width: 48,
                        height: 48,
                        cacheWidth: 100,
                        errorBuilder: (_, __, ___) =>
                            const Icon(Icons.broken_image)),
                  ),
                Expanded(
                    child: Text(
                        body.isEmpty ? kKinds[kind]! : '${kKinds[kind]}: $body',
                        maxLines: 3,
                        overflow: TextOverflow.ellipsis)),
              ]),
            ),
          ),
          const SizedBox(height: 16),
          FilledButton.icon(
              onPressed: () => _open(i),
              icon: const Icon(Icons.chat),
              label: const Text('فتح واتساب بيزنس')),
          if (opened)
            Padding(
              padding: const EdgeInsets.only(top: 12),
              child: Text(
                  'اختاري الجروب في واتساب (تم نسخ كلمة البحث) وأرسلي، ثم ارجعي هنا واضغطي "تم الإرسال".',
                  style: t.textTheme.bodyMedium
                      ?.copyWith(color: t.colorScheme.primary)),
            ),
          const SizedBox(height: 12),
          Row(children: [
            Expanded(
                child: FilledButton.tonalIcon(
                    onPressed: () => _mark(i, 'done'),
                    icon: const Icon(Icons.check),
                    label: const Text('تم الإرسال'))),
            const SizedBox(width: 12),
            Expanded(
                child: OutlinedButton.icon(
                    onPressed: () => _mark(i, 'skipped'),
                    icon: const Icon(Icons.skip_next),
                    label: const Text('تخطي'))),
          ]),
        ],
      ]),
    );
  }
}
