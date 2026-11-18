import 'dart:async';
import 'dart:convert';
import 'dart:io';

import 'package:archive/archive.dart' show ZipDecoder;
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
  'image': 'صورة فقط',
  'imageText': 'صورة مع نص',
  'stickers': 'استيكر أو أكثر',
};
// أقصى عدد ملفات في الرسالة (واتساب يقبل حتى 30). الاستيكرات: واحد أو أكثر بدون عدد محدد.
const Map<String, int> kLimit = {
  'text': 0,
  'image': 1,
  'imageText': 1,
  'stickers': 30,
};

/// اسم نوع الرسالة (يدعم المهام القديمة sticker1/sticker2).
String kindName(String k) =>
    kKinds[k] ?? (k.startsWith('sticker') ? 'استيكر' : k);

// ======================= الجروبات وملف Excel =======================
// لا توجد بيانات مدمجة في التطبيق: المصدر الوحيد هو ملف Excel (يُستورد من التطبيق أو يوضع في المجلد).

class G {
  final String name, numbers, teacher;
  final List<String> cats; // التصنيفات/الأيام (عمود E في Excel)
  const G(this.name, this.numbers, this.teacher, [this.cats = const []]);
  String get key => '$name|$numbers';
  String get title => numbers.isEmpty ? name : '$name  $numbers';

  /// كلمة البحث في واتساب: أول رقم، وإلا الاسم.
  String get searchKey {
    for (final n in numbers.split('/')) {
      if (n.trim().isNotEmpty) return n.trim();
    }
    return name;
  }

  Map<String, dynamic> toJson() => {
        'name': name,
        'numbers': numbers,
        'teacher': teacher,
        'cats': cats,
      };
  factory G.fromJson(Map<String, dynamic> j) => G(
      j['name'] as String,
      j['numbers'] as String,
      j['teacher'] as String,
      ((j['cats'] as List?) ?? const []).map((e) => e.toString()).toList());
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

const List<String> kDays = [
  'السبت',
  'الأحد',
  'الاثنين',
  'الثلاثاء',
  'الأربعاء',
  'الخميس',
  'الجمعة',
];
const String kNoCat = '__none';

/// يقسم خانة التصنيف: السبت/الثلاثاء أو السبت، الثلاثاء أو سطر لكل يوم.
List<String> parseCats(String s) => normDigits(s)
    .split(RegExp(r'[,،;؛/|\n]'))
    .map(_clean)
    .where((e) => e.isNotEmpty)
    .toList();

bool inCat(G g, String? cat) {
  if (cat == null) return true;
  if (cat == kNoCat) return g.cats.isEmpty && !Wa.inAnyWaCat(g);
  final n = _nz(cat);
  if (g.cats.any((c) => _nz(c) == n)) return true;
  return Wa.inWaCat(g, cat);
}

/// الأيام أولاً بترتيب الأسبوع ثم باقي التصنيفات أبجدياً.
List<String> allCats(List<G> gs) {
  final seen = <String, String>{};
  for (final g in gs) {
    for (final c in g.cats) {
      seen.putIfAbsent(_nz(c), () => c);
    }
  }
  for (final k in Wa.cats.keys) {
    seen.putIfAbsent(_nz(k), () => k);
  }
  final out = <String>[];
  for (final d in kDays) {
    final v = seen.remove(_nz(d));
    if (v != null) out.add(v);
  }
  final rest = seen.values.toList()..sort();
  out.addAll(rest);
  return out;
}

/// الأعمدة: A = الاسم مع الأرقام (محمد524/554)، B أو C = المعلمة، D = "إيقاف" (أي كلمة تعطّل الصف)، E = التصنيف/الأيام (السبت/الثلاثاء).
G? parseRow(List<String> c) {
  if (c.isEmpty) return null;
  final raw = _clean(c[0]);
  if (raw.isEmpty) return null;
  if (c.length > 3 && _clean(c[3]).isNotEmpty) return null; // موقوف
  var name = raw, nums = '';
  // كود بصيغة PTN-001325: الرقم للبحث = 1325 (بدون PTN وبدون الأصفار)، والباقي هو الاسم.
  final ptn = RegExp(r'PTN\s*[-_]?\s*0*(\d+)', caseSensitive: false);
  final pm = ptn.allMatches(raw).toList();
  final m = pm.isEmpty ? _tail.firstMatch(raw) : null;
  if (pm.isNotEmpty) {
    nums = pm.map((e) => e.group(1)!).join('/');
    final rest = _clean(raw.replaceAll(ptn, ' '))
        .replaceAll(RegExp(r'^[\s/\-_,،]+|[\s/\-_,،]+$'), '');
    if (rest.isNotEmpty) name = rest;
  } else if (m != null && m.group(1)!.trim().isNotEmpty) {
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
  final cats = c.length > 4 ? parseCats(c[4]) : <String>[];
  return G(name, nums, teacher, cats);
}

({List<G> groups, int dups, int off}) parseSheet(List<List<String>> rows) {
  final seen = <String>{};
  final out = <G>[];
  var dups = 0, off = 0;
  // تصدير المواعيد (صف عناوين فيه "الكود" و"مدرس"): الاسم والكود في الأول، والمعلم من عمود "مدرس".
  var start = 0, tcol = -1;
  for (var i = 0; i < rows.length && i < 5; i++) {
    if (rows[i].isNotEmpty && _clean(rows[i][0]) == 'الكود') {
      start = i + 1;
      tcol = rows[i].indexWhere((e) => _clean(e) == 'مدرس');
      break;
    }
  }
  for (var i = start; i < rows.length; i++) {
    var r = rows[i];
    if (start > 0) {
      r = [r.isEmpty ? '' : r[0], (tcol >= 0 && tcol < r.length) ? r[tcol] : ''];
    }
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

String _xmlUnesc(String s) => s
    .replaceAllMapped(RegExp(r'&#x([0-9a-fA-F]+);'),
        (m) => String.fromCharCode(int.parse(m[1]!, radix: 16)))
    .replaceAllMapped(
        RegExp(r'&#(\d+);'), (m) => String.fromCharCode(int.parse(m[1]!)))
    .replaceAll('&lt;', '<')
    .replaceAll('&gt;', '>')
    .replaceAll('&quot;', '"')
    .replaceAll('&apos;', "'")
    .replaceAll('&amp;', '&');

String _tText(String x) => RegExp(r'<t(?:\s[^>]*)?>([\s\S]*?)</t>')
    .allMatches(x)
    .map((m) => _xmlUnesc(m.group(1)!))
    .join();

int _colIdx(String ref) {
  var n = 0;
  for (final c in ref.toUpperCase().codeUnits) {
    if (c < 65 || c > 90) break;
    n = n * 26 + (c - 64);
  }
  return n - 1;
}

/// قارئ xlsx بسيط وسريع (يدعم inlineStr وsharedStrings) لملفات التصدير التي تفشل مع مكتبة excel.
List<List<String>> readXlsxRaw(List<int> bytes) {
  final arc = ZipDecoder().decodeBytes(bytes);
  String? read(String name) {
    for (final f in arc.files) {
      if (f.isFile && f.name == name) {
        return utf8.decode(f.content as List<int>, allowMalformed: true);
      }
    }
    return null;
  }

  final ss = <String>[];
  final sx = read('xl/sharedStrings.xml');
  if (sx != null) {
    for (final m in RegExp(r'<si\b[^>]*>([\s\S]*?)</si>').allMatches(sx)) {
      ss.add(_tText(m.group(1)!));
    }
  }
  final sheets = <int, String>{};
  for (final f in arc.files) {
    final m = RegExp(r'^xl/worksheets/sheet(\d+)\.xml$').firstMatch(f.name);
    if (f.isFile && m != null) sheets[int.parse(m.group(1)!)] = f.name;
  }
  if (sheets.isEmpty) return [];
  final first = (sheets.keys.toList()..sort()).first;
  final sheet = read(sheets[first]!);
  if (sheet == null) return [];

  final rows = <List<String>>[];
  for (final rm
      in RegExp(r'<row\b[^>]*?(?:/>|>([\s\S]*?)</row>)').allMatches(sheet)) {
    final content = rm.group(1) ?? '';
    final cells = <int, String>{};
    var maxC = -1, auto = 0;
    for (final cm
        in RegExp(r'<c\b([^>]*?)(?:/>|>([\s\S]*?)</c>)').allMatches(content)) {
      final attrs = cm.group(1)!;
      final body = cm.group(2) ?? '';
      final ref = RegExp(r'\br="([A-Za-z]+)\d*"').firstMatch(attrs)?.group(1);
      final idx = ref != null ? _colIdx(ref) : auto;
      auto = idx + 1;
      final t = RegExp(r'\bt="(\w+)"').firstMatch(attrs)?.group(1);
      String v;
      if (t == 'inlineStr') {
        v = _tText(body);
      } else {
        final vm = RegExp(r'<v[^>]*>([\s\S]*?)</v>').firstMatch(body);
        final raw = vm == null ? '' : _xmlUnesc(vm.group(1)!);
        if (t == 's') {
          final k = int.tryParse(raw);
          v = (k != null && k >= 0 && k < ss.length) ? ss[k] : '';
        } else {
          v = raw;
        }
      }
      cells[idx] = v;
      if (idx > maxC) maxC = idx;
    }
    rows.add(List<String>.generate(maxC + 1, (i) => cells[i] ?? ''));
  }
  return rows;
}

List<List<String>> readRows(String path) {
  final bytes = File(path).readAsBytesSync();
  try {
    final r = readXlsxRaw(bytes);
    if (r.isNotEmpty) return r;
  } catch (_) {}
  final book = Excel.decodeBytes(bytes);
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
    return true;
  } catch (_) {
    return false;
  }
}

/// أحدث ملف .xlsx في المجلد (groups.xlsx أو أي ملف يُنسخ يدوياً)، أو null لو لا يوجد.
String? findExcel() {
  try {
    final fs = Directory(kDir)
        .listSync()
        .whereType<File>()
        .where((f) {
          final n = p.basename(f.path).toLowerCase();
          return n.endsWith('.xlsx') && !n.startsWith('~\$') && !n.startsWith('.');
        })
        .toList();
    if (fs.isEmpty) return null;
    fs.sort((x, y) => y.lastModifiedSync().compareTo(x.lastModifiedSync()));
    return fs.first.path;
  } catch (_) {
    return null;
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

// ======================= أسماء الجروبات من واتساب =======================
class Wa {
  static final Map<String, String?> _matchC = {};
  static Map<String, Set<String>>? _setsC;
  static List<String>? _namesC;
  static List<Set<String>>? _tokC;

  /// يُستدعى كلما تغيّرت الأسماء المسحوبة أو التصنيفات.
  static void invalidate() {
    _matchC.clear();
    _setsC = null;
    _namesC = null;
    _tokC = null;
  }

  static List<String> get names =>
      ((Store.data['wa'] as List?) ?? const []).map((e) => e.toString()).toList();

  static Map<String, String> get map {
    final m = Store.data['wamap'];
    if (m is Map) {
      return m.map((k, v) => MapEntry(k.toString(), v.toString()));
    }
    return <String, String>{};
  }

  /// تصنيفات واتساب المسحوبة: اسم التصنيف ← أسماء الجروبات.
  static Map<String, List<String>> get cats {
    final m = Store.data['wacats'];
    if (m is Map) {
      return m.map((k, v) => MapEntry(k.toString(),
          ((v as List?) ?? const []).map((e) => e.toString()).toList()));
    }
    return <String, List<String>>{};
  }

  static Map<String, Set<String>> get _sets {
    var c = _setsC;
    if (c == null) {
      c = <String, Set<String>>{};
      for (final e in cats.entries) {
        c[_nz(e.key)] = e.value.map(_nz).toSet();
      }
      _setsC = c;
    }
    return c;
  }

  /// اسم الجروب في واتساب المقابل لصف الشيت (يدوي ثم تلقائي).
  static String? matchOf(G g) {
    final m = Store.data['wamap'];
    if (m is Map && m[g.key] is String && (m[g.key] as String).isNotEmpty) {
      return m[g.key] as String;
    }
    return _matchC.putIfAbsent(g.key, () {
      final n = _namesC ??= names;
      final t = _tokC ??= n.map((w) => _toks(w).toSet()).toList();
      return autoMatch(g, n, t);
    });
  }

  static bool inWaCat(G g, String cat) {
    final set = _sets[_nz(cat)];
    if (set == null) return false;
    final m = matchOf(g);
    return m != null && set.contains(_nz(m));
  }

  static bool inAnyWaCat(G g) => _sets.keys.any((k) => inWaCat(g, k));
}

String _nz(String s) {
  var t = _clean(s).toLowerCase();
  t = t.replaceAll(RegExp(r'[\u064b-\u065f\u0670\u0640]'), '');
  t = t
      .replaceAll(RegExp(r'[أإآ]'), 'ا')
      .replaceAll('ى', 'ي')
      .replaceAll('ة', 'ه');
  return t;
}

List<String> _toks(String s) => RegExp(r'\p{L}+|\p{N}+', unicode: true)
    .allMatches(_nz(s))
    .map((m) => m.group(0)!)
    .toList();

/// يطابق صف الشيت مع أقرب اسم جروب في واتساب. يرجع null لو لا يوجد تطابق واضح.
String? autoMatch(G g, List<String> wa, [List<Set<String>>? pre]) {
  if (wa.isEmpty) return null;
  final st = <String>[
    ..._toks(g.name),
    ...g.numbers.split('/').where((e) => e.isNotEmpty),
  ];
  if (st.isEmpty) return null;
  String? best;
  var b1 = 0.0, b2 = 0.0;
  for (var wi = 0; wi < wa.length; wi++) {
    final w = wa[wi];
    final wt = pre != null ? pre[wi] : _toks(w).toSet();
    if (wt.isEmpty) continue;
    final hits = st.where((x) => wt.contains(x)).length;
    final extra = wt.length - hits;
    final sc = hits / st.length - 0.02 * (extra < 0 ? 0 : extra);
    if (sc > b1) {
      b2 = b1;
      b1 = sc;
      best = w;
    } else if (sc > b2) {
      b2 = sc;
    }
  }
  if (best == null || b1 < 0.6) return null;
  if (b2 > b1 - 0.15) return null; // أكثر من احتمال: يُترك للاختيار اليدوي
  return best;
}

/// الكلمة التي تُنسخ وتُلصق في بحث واتساب: الاختيار اليدوي ثم المطابقة التلقائية ثم الرقم/الاسم القديم.
String waKey(G g) {
  final m = Store.data['wamap'];
  if (m is Map && m[g.key] is String && (m[g.key] as String).isNotEmpty) {
    return m[g.key] as String;
  }
  return Wa.matchOf(g) ?? g.searchKey;
}

class WaGroupsScreen extends StatefulWidget {
  final List<G> groups;
  const WaGroupsScreen({super.key, required this.groups});
  @override
  State<WaGroupsScreen> createState() => _WaGroupsScreenState();
}

class _WaGroupsScreenState extends State<WaGroupsScreen> {
  /// لصق أسماء الجروبات (سطر لكل اسم) بدل السحب التلقائي.
  Future<void> _addNames() async {
    final namesC = TextEditingController();
    final catC = TextEditingController();
    final ok = await showDialog<bool>(
      context: context,
      builder: (ctx) => AlertDialog(
        title: const Text('إضافة أسماء الجروبات'),
        content: SingleChildScrollView(
          child: Column(mainAxisSize: MainAxisSize.min, children: [
            const Text(
                'انسخي أسماء الجروبات من واتساب والصقيها هنا، كل اسم في سطر.'),
            const SizedBox(height: 8),
            TextField(
              controller: namesC,
              minLines: 6,
              maxLines: 12,
              decoration: const InputDecoration(
                  border: OutlineInputBorder(), hintText: 'اسم الجروب\nاسم الجروب'),
            ),
            const SizedBox(height: 12),
            const Text('اختياري: اسم التصنيف (مثل اليوم) لحفظ الأسماء تحته'),
            const SizedBox(height: 6),
            TextField(
                controller: catC,
                decoration: const InputDecoration(hintText: 'مثال: السبت')),
            const SizedBox(height: 8),
            Wrap(spacing: 6, children: [
              for (final d in kDays)
                ActionChip(label: Text(d), onPressed: () => catC.text = d),
            ]),
          ]),
        ),
        actions: [
          TextButton(
              onPressed: () => Navigator.pop(ctx, false),
              child: const Text('إلغاء')),
          FilledButton(
              onPressed: () => Navigator.pop(ctx, true),
              child: const Text('حفظ')),
        ],
      ),
    );
    if (ok != true) return;
    final got = namesC.text
        .split(RegExp(r'[\r\n]+'))
        .map((e) => e.trim())
        .where((e) => e.isNotEmpty)
        .toSet()
        .toList();
    if (got.isEmpty) return;
    final cat = catC.text.trim();
    if (cat.isNotEmpty) {
      final cm = Wa.cats;
      cm[cat] = got;
      Store.data['wacats'] = cm;
    }
    Store.data['wa'] = <String>{...Wa.names, ...got}.toList();
    Wa.invalidate();
    await Store.save();
    if (mounted) {
      setState(() {});
      snack(context, 'تمت إضافة ${got.length} اسم');
    }
  }

  Future<void> _clearAll() async {
    Store.data['wa'] = <String>[];
    Store.data['wacats'] = <String, dynamic>{};
    Wa.invalidate();
    await Store.save();
    if (mounted) setState(() {});
  }

  Future<void> _pick(G g) async {
    final r = await showDialog<String>(
        context: context,
        builder: (_) => _PickDialog(title: g.title, names: Wa.names));
    if (r == null) return;
    final m = Wa.map;
    if (r.isEmpty) {
      m.remove(g.key);
    } else {
      m[g.key] = r;
    }
    Store.data['wamap'] = m;
    await Store.save();
    if (mounted) setState(() {});
  }

  @override
  Widget build(BuildContext context) {
    final names = Wa.names;
    final map = Wa.map;
    var ok = 0;
    final rows = <Widget>[];
    for (final g in widget.groups) {
      final manual = map[g.key];
      final auto = manual == null ? Wa.matchOf(g) : null;
      final shown = manual ?? auto;
      if (shown != null) ok++;
      rows.add(ListTile(
        title: Text(g.title),
        subtitle: Text(
          shown ??
              (names.isEmpty
                  ? 'لم تتم إضافة أسماء بعد'
                  : 'لا يوجد تطابق واضح — اضغطي للاختيار'),
          style: shown == null ? const TextStyle(color: Colors.red) : null,
        ),
        trailing: Icon(
          manual != null
              ? Icons.edit
              : auto != null
                  ? Icons.check_circle
                  : Icons.help_outline,
          color: shown == null ? Colors.red : Colors.green,
        ),
        onTap: () => _pick(g),
      ));
    }
    return Scaffold(
      appBar: AppBar(title: const Text('أسماء الجروبات من واتساب')),
      body: ListView(padding: const EdgeInsets.all(16), children: [
        Card(
          child: Padding(
            padding: const EdgeInsets.all(12),
            child: Text('الأسماء المحفوظة من واتساب: ${names.length}'),
          ),
        ),
        const SizedBox(height: 8),
        FilledButton.icon(
          onPressed: _addNames,
          icon: const Icon(Icons.paste),
          label: const Text('لصق / إضافة أسماء الجروبات'),
        ),
        if (names.isNotEmpty) ...[
          const SizedBox(height: 8),
          OutlinedButton.icon(
            onPressed: _clearAll,
            icon: const Icon(Icons.delete_sweep),
            label: const Text('مسح كل الأسماء المحفوظة'),
          ),
        ],
        if (Wa.cats.isNotEmpty) ...[
          const SizedBox(height: 8),
          Text('التصنيفات المسحوبة', style: Theme.of(context).textTheme.titleMedium),
          for (final e in Wa.cats.entries)
            ListTile(
              dense: true,
              title: Text('${e.key} (${e.value.length} جروب)'),
              trailing: IconButton(
                icon: const Icon(Icons.delete_outline),
                onPressed: () async {
                  final m = Wa.cats..remove(e.key);
                  Store.data['wacats'] = m;
                  Wa.invalidate();
                  await Store.save();
                  if (mounted) setState(() {});
                },
              ),
            ),
        ],
        const SizedBox(height: 8),
        if (names.isNotEmpty)
          Text('تمت مطابقة $ok من ${widget.groups.length} جروب',
              style: Theme.of(context).textTheme.titleMedium),
        ...rows,
      ]),
    );
  }
}

class _PickDialog extends StatefulWidget {
  final String title;
  final List<String> names;
  const _PickDialog({required this.title, required this.names});
  @override
  State<_PickDialog> createState() => _PickDialogState();
}

class _PickDialogState extends State<_PickDialog> {
  String q = '';
  @override
  Widget build(BuildContext context) {
    final nq = _nz(q);
    final list = widget.names
        .where((n) => nq.isEmpty || _nz(n).contains(nq))
        .toList();
    return AlertDialog(
      title: Text(widget.title),
      content: SizedBox(
        width: double.maxFinite,
        height: 420,
        child: Column(children: [
          TextField(
              decoration: const InputDecoration(
                  hintText: 'بحث', prefixIcon: Icon(Icons.search)),
              onChanged: (v) => setState(() => q = v)),
          Expanded(
              child: ListView.builder(
                  itemCount: list.length,
                  itemBuilder: (_, i) => ListTile(
                      dense: true,
                      title: Text(list[i]),
                      onTap: () => Navigator.pop(context, list[i])))),
        ]),
      ),
      actions: [
        TextButton(
            onPressed: () => Navigator.pop(context, ''),
            child: const Text('إلغاء الربط')),
        TextButton(
            onPressed: () => Navigator.pop(context),
            child: const Text('إغلاق')),
      ],
    );
  }
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
  String? srcPath;
  String? cat; // التصنيف المحدد للإرسال (null = الكل)
  bool loading = true;

  List<String> get _cats => allCats(groups);
  String? get _cat {
    final c = cat;
    if (c == null || c == kNoCat) return c;
    return _cats.any((x) => _nz(x) == _nz(c)) ? c : null;
  }

  List<G> get shown => groups.where((g) => inCat(g, _cat)).toList();

  Future<void> _setCat(String? c) async {
    setState(() => cat = c);
    if (c == null) {
      Store.data.remove('cat');
    } else {
      Store.data['cat'] = c;
    }
    await Store.save();
  }

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
    Wa.invalidate();
    final sc = Store.data['cat'];
    if (sc is String) cat = sc;
    final d = Store.data['draft'];
    if (d is Map<String, dynamic>) {
      kind = (d['kind'] as String?) ?? 'text';
      if (kind == 'sticker1' || kind == 'sticker2') kind = 'stickers';
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
    String? src;
    if (e == null) {
      src = findExcel();
      if (src == null) {
        e = 'لا يوجد ملف جروبات بعد.\nاضغطي "استيراد ملف Excel"، أو ضعي ملف .xlsx داخل مجلد A2R_Sender في الذاكرة الداخلية.';
      } else {
        try {
          final r = parseSheet(readRows(src));
          g = r.groups;
          dp = r.dups;
          of = r.off;
        } catch (x) {
          e = 'تعذر قراءة ${p.basename(src)}: $x';
        }
      }
    }
    if (!mounted) return;
    setState(() {
      srcPath = src;
      groups = g;
      dups = dp;
      off = of;
      err = e;
      loading = false;
    });
  }

  /// استيراد ملف Excel من أي مكان في الهاتف ونسخه مكان groups.xlsx.
  Future<void> _importExcel() async {
    try {
      final res = await FilePicker.platform.pickFiles(
          type: FileType.custom, allowedExtensions: ['xlsx']);
      if (res == null || res.files.isEmpty || res.files.first.path == null) {
        return;
      }
      final src = File(res.files.first.path!);
      final r = parseSheet(readRows(src.path)); // تحقق إن الملف سليم
      if (r.groups.isEmpty) {
        if (mounted) snack(context, 'لا توجد جروبات في هذا الملف');
        return;
      }
      final e = await prepareFolder();
      if (e != null) {
        if (mounted) snack(context, e);
        return;
      }
      final dst = File(kFile);
      if (dst.existsSync()) dst.deleteSync();
      await src.copy(kFile);
      await _reload();
      if (mounted) snack(context, 'تم استيراد ${r.groups.length} جروب');
    } catch (x) {
      if (mounted) snack(context, 'تعذر استيراد الملف: $x');
    }
  }

  void _saveDraft() {
    Store.data['draft'] = {'kind': kind, 'text': text.text, 'media': media};
    Store.save();
  }

  void _setKind(String k) {
    setState(() {
      // الاستيكرات والصور لا تُخلط: عند التبديل بينهما تُفرّغ القائمة.
      if ((k == 'stickers') != (kind == 'stickers')) media = [];
      kind = k;
      final lim = kLimit[k]!;
      if (media.length > lim) media = media.sublist(0, lim);
    });
    _saveDraft();
  }

  Future<void> _pick() async {
    final room = kLimit[kind]! - media.length;
    if (room <= 0) return;
    // الاستيكرات ملفات .webp: تُختار من مدير الملفات وليس من معرض الصور.
    final res = kind == 'stickers'
        ? await FilePicker.platform.pickFiles(
            type: FileType.custom,
            allowedExtensions: ['webp'],
            allowMultiple: true)
        : await FilePicker.platform
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
    if (kind != 'text' && media.isEmpty) {
      return kind == 'stickers' ? 'اختاري استيكراً واحداً على الأقل' : 'اختاري صورة';
    }
    return null;
  }

  Future<void> _start() async {
    final v = _validate();
    if (v != null) {
      snack(context, v);
      return;
    }
    final sel = shown;
    if (sel.isEmpty) {
      snack(context, groups.isEmpty ? 'لا توجد جروبات في ملف Excel' : 'لا توجد جروبات في هذا التصنيف');
      return;
    }
    final catLabel = _cat == null ? '' : ' (${_cat == kNoCat ? 'بدون تصنيف' : _cat})';
    final old = Store.job != null;
    final ok = await ask(context, 'بدء الإرسال',
        '${kindName(kind)}${kind == 'stickers' ? ' (${media.length})' : ''} إلى ${sel.length} جروب$catLabel.${old ? '\nسيتم إغلاق المهمة غير المنتهية السابقة.' : ''}',
        ok: 'ابدأ');
    if (!ok || !mounted) return;
    Store.data['job'] = {
      'kind': kind,
      'text': (kind == 'text' || kind == 'imageText') ? text.text.trim() : '',
      'media': kind == 'text' ? <String>[] : media,
      'groups': sel.map((g) => g.toJson()).toList(),
      'st': List<String>.filled(sel.length, 'pending'),
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
            tooltip: 'أسماء الجروبات من واتساب',
            onPressed: () => Navigator.push(
                context,
                MaterialPageRoute(
                    builder: (_) => WaGroupsScreen(groups: groups))),
            icon: const Icon(Icons.groups)),
        IconButton(
            tooltip: 'استيراد ملف Excel',
            onPressed: _importExcel,
            icon: const Icon(Icons.upload_file)),
        IconButton(
            tooltip: 'تحديث من ملف Excel',
            onPressed: _reload,
            icon: const Icon(Icons.refresh)),
      ]),
      body: ListView(padding: const EdgeInsets.all(16), children: [
        if (loading) const LinearProgressIndicator(),
        OutlinedButton.icon(
          onPressed: _importExcel,
          icon: const Icon(Icons.upload_file),
          label: const Text('استيراد ملف Excel'),
        ),
        const SizedBox(height: 8),
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
                  'الملف: ${srcPath ?? kFile}'
                  '${_cat != null ? '\nالمحدد للإرسال: ${shown.length} جروب' : ''}'),
            ),
          ),
        if (err == null && _cats.isNotEmpty) ...[
          const SizedBox(height: 8),
          Card(
            child: Padding(
              padding: const EdgeInsets.all(12),
              child: Column(
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    Text('الإرسال إلى:',
                        style: Theme.of(context).textTheme.titleSmall),
                    const SizedBox(height: 8),
                    Wrap(spacing: 8, runSpacing: 8, children: [
                      ChoiceChip(
                          label: Text('الكل (${groups.length})'),
                          selected: _cat == null,
                          onSelected: (_) => _setCat(null)),
                      for (final c in _cats)
                        ChoiceChip(
                            label: Text(
                                '$c (${groups.where((g) => inCat(g, c)).length})'),
                            selected: _cat != null &&
                                _cat != kNoCat &&
                                _nz(_cat!) == _nz(c),
                            onSelected: (_) => _setCat(c)),
                      if (groups.any((g) => g.cats.isEmpty && !Wa.inAnyWaCat(g)))
                        ChoiceChip(
                            label: Text(
                                'بدون تصنيف (${groups.where((g) => inCat(g, kNoCat)).length})'),
                            selected: _cat == kNoCat,
                            onSelected: (_) => _setCat(kNoCat)),
                    ]),
                  ]),
            ),
          ),
        ],
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
              icon: Icon(kind == 'stickers'
                  ? Icons.emoji_emotions
                  : Icons.add_photo_alternate),
              label: Text(kind == 'stickers'
                  ? 'إضافة استيكرات .webp (${media.length})'
                  : 'اختيار صورة من الهاتف (${media.length}/$lim)'),
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
  int? forced;
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
    if (forced != null) return forced;
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
    await Clipboard.setData(ClipboardData(text: waKey(_g(i))));
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
      forced = null;
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
    if (mounted) setState(() => forced = i);
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
                          child: Text('ابحثي في واتساب عن:  ${waKey(_g(i))}',
                              style: t.textTheme.titleMedium)),
                      IconButton(
                        tooltip: 'نسخ',
                        icon: const Icon(Icons.copy),
                        onPressed: () async {
                          await Clipboard.setData(
                              ClipboardData(text: waKey(_g(i))));
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
                        body.isEmpty ? kindName(kind) : '${kindName(kind)}: $body',
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
