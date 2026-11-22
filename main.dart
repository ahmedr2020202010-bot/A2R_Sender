import 'dart:async';
import 'dart:convert';
import 'dart:io';

import 'package:archive/archive.dart' show ZipDecoder;
import 'package:excel/excel.dart' show Excel;
import 'package:file_picker/file_picker.dart';
import 'package:flutter/material.dart';
import 'package:flutter/services.dart';
import 'package:flutter_localizations/flutter_localizations.dart';
import 'package:flutter/foundation.dart' show compute;
import 'package:image_picker/image_picker.dart' show ImagePicker, ImageSource;
import 'package:path/path.dart' as p;
import 'package:path_provider/path_provider.dart';

// ======================= الإعدادات =======================
const String kRoot = '/storage/emulated/0';
const String kDir = '$kRoot/A2R_Sender';
const String kFile = '$kDir/groups.xlsx';
const MethodChannel _ch = MethodChannel('a2r/share');

const Map<String, String> kKinds = {
  'text': 'نص فقط',
  'image': 'صورة فقط',
  'imageText': 'صورة مع نص',
  'stickers': 'استيكر أو أكثر',
};

/// أسماء مختصرة لشريط اختيار نوع الرسالة.
const Map<String, String> kKindsShort = {
  'text': 'نص فقط',
  'image': 'صورة',
  'imageText': 'صورة + نص',
  'stickers': 'استيكرات',
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

/// قراءة وتحليل الشيت كاملاً — تُستدعى عبر compute حتى لا يتجمد خيط الواجهة.
({List<G> groups, int dups, int off}) parseSheetFile(String path) =>
    parseSheet(readRows(path));

/// فحص مجلد وسائط (أسماء المجلدات والملفات مرتبة) — يعمل عبر compute خارج خيط الواجهة.
({List<String> dirs, List<String> files, bool ok}) scanMediaDir(
    ({String dir, List<String> exts}) q) {
  try {
    final d = <String>[];
    final f = <String>[];
    for (final e in Directory(q.dir).listSync(followLinks: false)) {
      final n = p.basename(e.path);
      if (n.startsWith('.')) continue;
      if (e is Directory) {
        d.add(e.path);
      } else if (e is File &&
          q.exts.contains(p.extension(n).toLowerCase().replaceFirst('.', ''))) {
        f.add(e.path);
      }
    }
    d.sort((a, b) =>
        p.basename(a).toLowerCase().compareTo(p.basename(b).toLowerCase()));
    f.sort((a, b) =>
        File(b).lastModifiedSync().compareTo(File(a).lastModifiedSync()));
    return (dirs: d, files: f, ok: true);
  } catch (_) {
    return (dirs: const [], files: const [], ok: false);
  }
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

/// يتأكد من صلاحية الملفات وينشئ المجلد. يرجع نص الخطأ أو null.
Future<String?> prepareFolder() async {
  final s = await Perm.status();
  if (s.isNotEmpty && s['storage'] != true) {
    return 'صلاحية الملفات غير ممنوحة.\nاسمحي للتطبيق بالوصول إلى كل الملفات ثم اضغطي تحديث.';
  }
  if (await _tryPrepare()) return null;
  return 'لا أستطيع الوصول إلى المجلد $kDir\nاسمحي للتطبيق بالوصول إلى الملفات من إعدادات الهاتف ثم اضغطي تحديث.';
}

// ======================= الحفظ المحلي (المسودة + تقدم المهمة) =======================
class Store {
  static Map<String, dynamic> data = {};

  static Future<File> _file() async =>
      File(p.join((await getApplicationSupportDirectory()).path, 'state.json'));

  static Future<void>? _once;

  /// تحميل مرة واحدة فقط (التابان يستدعيانه معاً).
  static Future<void> load() => _once ??= _load();

  static Future<void> _load() async {
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

  /// مهمة تاب "رسالة لمصنف".
  static Map<String, dynamic>? get ujob {
    final j = data['ujob'];
    if (j is Map<String, dynamic> && j['finished'] != true) return j;
    return null;
  }
}

int _mediaSeq = 0;

Future<String> persistMedia(String src) async {
  final dir = Directory(
      p.join((await getApplicationSupportDirectory()).path, 'media'));
  await dir.create(recursive: true);
  final ext = p.extension(src).toLowerCase();
  _mediaSeq++;
  final dest = p.join(dir.path,
      '${DateTime.now().microsecondsSinceEpoch}_$_mediaSeq${ext.isEmpty ? '.webp' : ext}');
  final bytes = await File(src).readAsBytes();
  await File(dest).writeAsBytes(bytes, flush: true);
  return dest;
}

/// اختيار نسخة واتساب المستهدفة (بيزنس أو العادي).
class WaApp {
  static const String business = 'com.whatsapp.w4b';
  static const String normal = 'com.whatsapp';

  static String get pkg =>
      Store.data['waPkg'] == normal ? normal : business;

  static bool get isNormal => pkg == normal;

  static String get label => isNormal ? 'واتساب العادي' : 'واتساب بيزنس';

  static Future<void> setPkg(String p) async {
    Store.data['waPkg'] = p == normal ? normal : business;
    await Store.save();
  }

  static List<String> stickerDirs() => isNormal
      ? [
          '$kRoot/Android/media/com.whatsapp/WhatsApp/Media/WhatsApp Stickers',
          '$kRoot/WhatsApp/Media/WhatsApp Stickers',
          '$kRoot/Android/media/com.whatsapp.w4b/WhatsApp Business/Media/WhatsApp Business Stickers',
          '$kRoot/Download',
        ]
      : [
          '$kRoot/Android/media/com.whatsapp.w4b/WhatsApp Business/Media/WhatsApp Business Stickers',
          '$kRoot/WhatsApp Business/Media/WhatsApp Business Stickers',
          '$kRoot/Android/media/com.whatsapp/WhatsApp/Media/WhatsApp Stickers',
          '$kRoot/Download',
        ];
}

/// حذف الصور والاستيكرات المؤقتة غير المستخدمة في أي مسودة أو مهمة نشطة.
Future<({int count, int bytes})> cleanupUnusedMedia() async {
  try {
    final dir = Directory(
        p.join((await getApplicationSupportDirectory()).path, 'media'));
    if (!dir.existsSync()) return (count: 0, bytes: 0);
    final keep = <String>{};
    void collect(dynamic obj) {
      if (obj is Map) {
        final m = obj['media'];
        if (m is List) {
          for (final x in m) {
            keep.add(x.toString());
          }
        }
      }
    }
    collect(Store.data['draft']);
    collect(Store.data['draft2']);
    collect(Store.job);
    collect(Store.ujob);

    var count = 0;
    var bytes = 0;
    for (final e in dir.listSync(followLinks: false)) {
      if (e is File && !keep.contains(e.path)) {
        try {
          bytes += e.lengthSync();
          e.deleteSync();
          count++;
        } catch (_) {}
      }
    }
    return (count: count, bytes: bytes);
  } catch (_) {
    return (count: 0, bytes: 0);
  }
}

/// نافذة الإعدادات السريعة: اختيار نسخة واتساب + تنظيف الملفات المؤقتة.
Future<void> showAppSettings(BuildContext context, VoidCallback onChanged) async {
  await showModalBottomSheet<void>(
    context: context,
    showDragHandle: true,
    builder: (ctx) => StatefulBuilder(
      builder: (ctx, setModal) {
        final t = Theme.of(ctx);
        return SafeArea(
          child: Padding(
            padding: const EdgeInsets.fromLTRB(16, 4, 16, 20),
            child: Column(
              mainAxisSize: MainAxisSize.min,
              crossAxisAlignment: CrossAxisAlignment.stretch,
              children: [
                Text('إعدادات التطبيق', style: t.textTheme.titleLarge),
                const SizedBox(height: 16),
                Text('نسخة واتساب المستخدمة في الإرسال:',
                    style: t.textTheme.titleSmall),
                const SizedBox(height: 8),
                SegmentedButton<String>(
                  segments: const [
                    ButtonSegment(
                      value: WaApp.business,
                      label: Text('واتساب بيزنس'),
                      icon: Icon(Icons.business),
                    ),
                    ButtonSegment(
                      value: WaApp.normal,
                      label: Text('واتساب العادي'),
                      icon: Icon(Icons.chat),
                    ),
                  ],
                  selected: {WaApp.pkg},
                  onSelectionChanged: (v) async {
                    await WaApp.setPkg(v.first);
                    setModal(() {});
                    onChanged();
                  },
                ),
                const SizedBox(height: 16),
                const Divider(),
                const SizedBox(height: 8),
                OutlinedButton.icon(
                  onPressed: () async {
                    final r = await cleanupUnusedMedia();
                    if (!ctx.mounted) return;
                    Navigator.pop(ctx);
                    final mb = (r.bytes / (1024 * 1024)).toStringAsFixed(2);
                    snack(
                        context,
                        r.count == 0
                            ? 'لا توجد ملفات مؤقتة زائدة'
                            : 'تم مسح ${r.count} ملف مؤقت ($mb ميجابايت)');
                  },
                  icon: const Icon(Icons.cleaning_services_outlined),
                  label: const Text('تنظيف الصور والاستيكرات المؤقتة'),
                ),
              ],
            ),
          ),
        );
      },
    ),
  );
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

// ======================= الفقاعة العائمة (خدمة إمكانية الوصول) =======================
/// جسر للجزء الأصلي (Kotlin): الفقاعة تكتب كلمة البحث في واتساب بيزنس وتنقل للجروب التالي.
class Bubble {
  static const List<String> _codes = ['pending', 'done', 'skipped'];
  static int code(String s) => s == 'done' ? 1 : (s == 'skipped' ? 2 : 0);

  static Future<Map<String, dynamic>> status() async {
    try {
      final r = await _ch.invokeMethod<Map>('svcStatus');
      return (r ?? {}).map((k, v) => MapEntry(k.toString(), v));
    } catch (_) {
      return {};
    }
  }

  /// يشغّل الجلسة ويفتح العنصر [at]. يرجع نص الخطأ أو null عند النجاح.
  static Future<String?> start(Map<String, dynamic> session, int at) async {
    try {
      await _ch.invokeMethod<bool>(
          'start', {'session': jsonEncode(session), 'at': at});
      return null;
    } on PlatformException catch (e) {
      return e.code == 'NO_A11Y'
          ? 'فعّلي خدمة A2R Sender من إعدادات إمكانية الوصول أولاً'
          : (e.message ?? 'تعذر تشغيل الفقاعة');
    } catch (e) {
      return '$e';
    }
  }

  static Future<Map<String, dynamic>?> state() async {
    try {
      final s = await _ch.invokeMethod<String>('state');
      if (s == null || s.isEmpty) return null;
      return jsonDecode(s) as Map<String, dynamic>;
    } catch (_) {
      return null;
    }
  }

  static Future<void> stop() async {
    try {
      await _ch.invokeMethod<bool>('stop');
    } catch (_) {}
  }

  static Future<void> clear() async {
    try {
      await _ch.invokeMethod<bool>('clear');
    } catch (_) {}
  }

  static Future<void> setStatus(int i, String st) async {
    try {
      await _ch.invokeMethod<bool>('setStatus', {'i': i, 's': code(st)});
    } catch (_) {}
  }

  static Future<void> open(String what) async {
    try {
      await _ch.invokeMethod<bool>(what);
    } catch (_) {}
  }

  /// يدمج حالات الفقاعة (تم/تخطي) في المهمة المحفوظة [key] = job أو ujob.
  /// في وضع المصنف المفتوح القائمة تكبر مع كل محادثة تفتحها الفقاعة، وتُلتقط أسماؤها.
  static Future<({bool active, bool changed})> sync(String key) async {
    final j = Store.data[key];
    if (j is! Map) return (active: false, changed: false);
    final s = await state();
    if (s == null || s['id'] != key) return (active: false, changed: false);
    var changed = false;
    final st = j['st'];
    final ns = s['st'];
    if (st is List && ns is List && ns.length >= st.length) {
      while (st.length < ns.length) {
        st.add('pending');
        changed = true;
      }
      for (var i = 0; i < ns.length; i++) {
        final n = ns[i];
        final v = _codes[(n is int ? n : 0).clamp(0, 2).toInt()];
        if (st[i] != v) {
          st[i] = v;
          changed = true;
        }
      }
      if (changed) await Store.save();
    }
    final tt = s['tt'];
    final items = j['items'];
    if (tt is List && items is List) {
      var grew = false;
      while (items.length < tt.length) {
        items.add('');
        grew = true;
      }
      for (var i = 0; i < tt.length; i++) {
        final v = tt[i].toString();
        if (v.isNotEmpty && items[i] != v) {
          items[i] = v;
          grew = true;
        }
      }
      if (grew) await Store.save();
    }
    final active = s['active'] == true;
    if (!active) await clear();
    return (active: active, changed: changed);
  }

  static Future<void> syncAll() async {
    await sync('job');
    await sync('ujob');
  }

  /// قبل إنشاء مهمة جديدة: إنهاء أي جلسة قديمة حتى لا تختلط حالاتها.
  static Future<void> reset() async {
    await stop();
    await clear();
  }
}

/// جلسة الإرسال الفردي: وضع المراجعة (البحث في القائمة الرئيسية) أو المشاركة المباشرة.
Map<String, dynamic> jobSession(Map<String, dynamic> job,
    {String mode = 'review'}) {
  final gs = job['groups'] as List;
  final st = job['st'] as List;
  return {
    'id': 'job',
    'mode': mode,
    'pkg': WaApp.pkg,
    'text': job['text'] as String,
    'media': (job['media'] as List).map((e) => e.toString()).toList(),
    'items': List.generate(gs.length, (i) {
      final g = G.fromJson(gs[i] as Map<String, dynamic>);
      return {
        't': g.title,
        'k': waKey(g),
        'l': true,
        's': Bubble.code(st[i] as String),
      };
    }),
  };
}

/// جلسة المصنف المفتوح: المستخدم يفتح المصنف في واتساب بنفسه، والفقاعة تفتح
/// محادثات القائمة الظاهرة واحدة واحدة وتلتقط اسم كل محادثة عند فتحها.
/// الرسائل المصحوبة بوسائط تُرسل من داخل المحادثة عبر المشاركة بالبحث بالاسم الملتقط.
Map<String, dynamic> labelSession(Map<String, dynamic> job) {
  final names = (job['items'] as List).map((e) => e.toString()).toList();
  final st = job['st'] as List;
  return {
    'id': 'ujob',
    'mode': 'listwalk',
    'pkg': WaApp.pkg,
    'text': job['text'] as String,
    'media': (job['media'] as List).map((e) => e.toString()).toList(),
    'seen': names,
    'items': List.generate(
        names.length,
        (i) => {
              't': names[i],
              'k': names[i],
              'l': false,
              's': Bubble.code(i < st.length ? st[i] as String : 'pending'),
            }),
  };
}

/// بطاقة تفعيل الفقاعة وتشغيلها.
class BubbleCard extends StatefulWidget {
  final String startLabel;
  final Future<void> Function() onStart;
  const BubbleCard({super.key, required this.startLabel, required this.onStart});
  @override
  State<BubbleCard> createState() => _BubbleCardState();
}

class _BubbleCardState extends State<BubbleCard> with WidgetsBindingObserver {
  Map<String, dynamic> st = {};

  @override
  void initState() {
    super.initState();
    WidgetsBinding.instance.addObserver(this);
    _load();
  }

  @override
  void dispose() {
    WidgetsBinding.instance.removeObserver(this);
    super.dispose();
  }

  @override
  void didChangeAppLifecycleState(AppLifecycleState s) {
    if (s == AppLifecycleState.resumed) _load();
  }

  Future<void> _load() async {
    final r = await Bubble.status();
    if (mounted) setState(() => st = r);
  }

  @override
  Widget build(BuildContext context) {
    final on = st['a11y'] == true;
    final t = Theme.of(context);
    if (on) {
      // الخدمة مفعّلة: زر التشغيل مباشرة بدون كارت وشرح.
      return Column(
        crossAxisAlignment: CrossAxisAlignment.stretch,
        children: [
          FilledButton.icon(
              onPressed: widget.onStart,
              icon: const Icon(Icons.play_circle_outline),
              label: Text(widget.startLabel)),
          if (st['overlay'] != true)
            TextButton(
                onPressed: () => Bubble.open('openOverlay'),
                child: const Text(
                    'لو الفقاعة لا تفتح واتساب: اسمحي بالعرض فوق التطبيقات')),
        ],
      );
    }
    return Card(
      child: Padding(
        padding: const EdgeInsets.all(12),
        child: Column(crossAxisAlignment: CrossAxisAlignment.start, children: [
          Row(children: [
            Icon(Icons.bubble_chart, color: t.colorScheme.outline),
            const SizedBox(width: 8),
            Expanded(
                child: Text('الفقاعة العائمة (تفعيل لمرة واحدة)',
                    style: t.textTheme.titleSmall)),
          ]),
          const SizedBox(height: 8),
          const Text(
              'تكتب الفقاعة كلمة البحث في واتساب بيزنس بدل اللصق اليدوي، وتنقلك للجروب التالي بدون الرجوع للتطبيق. '
              'فعّلي "A2R Sender" من إعدادات إمكانية الوصول (الخدمات المثبتة). الخدمة مقصورة على واتساب بيزنس '
              'ولا تضغط على الجروب ولا على إرسال.'),
          const SizedBox(height: 8),
          FilledButton.tonalIcon(
              onPressed: () => Bubble.open('openA11y'),
              icon: const Icon(Icons.accessibility_new),
              label: const Text('فتح إعدادات إمكانية الوصول')),
          TextButton(
              onPressed: () => Bubble.open('openAppInfo'),
              child: const Text(
                  'التفعيل مقفول أو ظهرت رسالة "إعداد مقيّد"؟ افتحي معلومات التطبيق')),
        ]),
      ),
    );
  }
}

// ======================= الصلاحيات عند بداية التشغيل =======================
class Perm {
  static Future<Map<String, dynamic>> status() async {
    try {
      final r = await _ch.invokeMethod<Map>('permStatus');
      return (r ?? {}).map((k, v) => MapEntry(k.toString(), v));
    } catch (_) {
      return {};
    }
  }

  static Future<void> reqStorage() async {
    try {
      await _ch.invokeMethod<bool>('reqStorage');
    } catch (_) {}
  }
}

/// شاشة البداية: تطلب الصلاحيات واحدة واحدة (الملفات ← العرض فوق التطبيقات ← خدمة إمكانية الوصول)
/// ثم تفتح التطبيق. تعمل على أندرويد 10 حتى 16.
class Gate extends StatefulWidget {
  const Gate({super.key});
  @override
  State<Gate> createState() => _GateState();
}

class _GateState extends State<Gate> with WidgetsBindingObserver {
  Map<String, dynamic> st = {};
  bool checked = false, skipped = false, entered = false, busy = false;

  @override
  void initState() {
    super.initState();
    WidgetsBinding.instance.addObserver(this);
    _refresh();
  }

  @override
  void dispose() {
    WidgetsBinding.instance.removeObserver(this);
    super.dispose();
  }

  @override
  void didChangeAppLifecycleState(AppLifecycleState s) {
    if (s == AppLifecycleState.resumed) _refresh();
  }

  bool _ok(String k) => st[k] == true;
  bool get _all => _ok('storage') && _ok('overlay') && _ok('a11y');

  Future<void> _refresh() async {
    if (busy) return;
    busy = true;
    try {
      final r = await Perm.status();
      if (!mounted) return;
      setState(() {
        st = r;
        checked = true;
      });
    } finally {
      busy = false;
    }
  }

  Future<void> _request(String k) async {
    if (k == 'storage') {
      await Perm.reqStorage();
    } else if (k == 'overlay') {
      await Bubble.open('openOverlay');
    } else {
      await Bubble.open('openA11y');
    }
  }

  Widget _item(ThemeData t, String k, String title, String why, String btn) {
    final ok = _ok(k);
    return Card(
      child: Padding(
        padding: const EdgeInsets.all(12),
        child: Column(crossAxisAlignment: CrossAxisAlignment.start, children: [
          Row(children: [
            Icon(ok ? Icons.check_circle : Icons.error_outline,
                color: ok ? Colors.green : t.colorScheme.error),
            const SizedBox(width: 8),
            Expanded(child: Text(title, style: t.textTheme.titleMedium)),
          ]),
          const SizedBox(height: 6),
          Text(why),
          if (!ok) ...[
            const SizedBox(height: 8),
            FilledButton.tonal(
                onPressed: () => _request(k), child: Text(btn)),
          ],
        ]),
      ),
    );
  }

  @override
  Widget build(BuildContext context) {
    if (!checked) {
      return const Scaffold(body: Center(child: CircularProgressIndicator()));
    }
    if (_all || st.isEmpty || skipped || entered) {
      entered = true;
      return const Shell();
    }
    final t = Theme.of(context);
    final sdk = st['sdk'];
    return Scaffold(
      appBar: AppBar(title: const Text('تجهيز التطبيق')),
      body: ListView(padding: const EdgeInsets.all(16), children: [
        Text('محتاجين تفعيل 3 صلاحيات عشان التطبيق يشتغل كامل:',
            style: t.textTheme.titleMedium),
        const SizedBox(height: 12),
        _item(
            t,
            'storage',
            '1) الوصول إلى الملفات',
            'لقراءة ملف Excel من مجلد A2R_Sender واختيار الصور والاستيكرات.',
            'السماح بالوصول للملفات'),
        _item(
            t,
            'overlay',
            '2) العرض فوق التطبيقات',
            'عشان الفقاعة تقدر تفتح واتساب بيزنس من الخلفية.',
            'السماح بالعرض فوق التطبيقات'),
        _item(
            t,
            'a11y',
            '3) خدمة إمكانية الوصول',
            'تكتب كلمة البحث في واتساب بيزنس فقط. لا تقرأ محادثاتك ولا تضغط إرسال. '
                'فعّلي "A2R Sender" من الخدمات المثبتة.',
            'فتح إعدادات إمكانية الوصول'),
        if (sdk is int && sdk >= 33 && !_ok('a11y'))
          Padding(
            padding: const EdgeInsets.only(top: 4),
            child: Column(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  const Text(
                      'لو خيار التفعيل رمادي (أندرويد 13+): افتحي معلومات التطبيق ← القائمة ⋮ ← "السماح بالإعدادات المقيّدة"، ثم ارجعي وفعّلي الخدمة.'),
                  TextButton(
                      onPressed: () => Bubble.open('openAppInfo'),
                      child: const Text('فتح معلومات التطبيق')),
                ]),
          ),
        const SizedBox(height: 16),
        OutlinedButton(
            onPressed: () => setState(() => skipped = true),
            child: const Text('متابعة بدون تفعيل (بعض الميزات لن تعمل)')),
      ]),
    );
  }
}

// ======================= تبويب استيكراتي =======================
/// مجلد مكتبة الاستيكرات الخاص بالتطبيق (تخزين مشترك يتحكم فيه المستخدم).
const String kStickerLibDir = '$kDir/stickers';

String _firstExistingDir(List<String> c) {
  for (final d in c) {
    if (Directory(d).existsSync()) return d;
  }
  return c.isNotEmpty ? c.first : kRoot;
}

class StickersTab extends StatefulWidget {
  const StickersTab({super.key});
  @override
  State<StickersTab> createState() => _StickersTabState();
}

class _StickersTabState extends State<StickersTab> {
  List<File> lib = [];
  List<dynamic> packs = [];
  bool busy = false;

  Map get _added =>
      Store.data['stickerAdded'] is Map ? Store.data['stickerAdded'] as Map : {};

  @override
  void initState() {
    super.initState();
    _refresh();
  }

  Future<void> _refresh() async {
    final files = <File>[];
    try {
      final d = Directory(kStickerLibDir);
      if (d.existsSync()) {
        for (final e in d.listSync(followLinks: false)) {
          if (e is File &&
              ['webp', 'png', 'gif', 'jpg', 'jpeg'].contains(
                  p.extension(e.path).toLowerCase().replaceFirst('.', ''))) {
            files.add(e);
          }
        }
        files.sort(
            (a, b) => a.lastModifiedSync().compareTo(b.lastModifiedSync()));
      }
    } catch (_) {}
    var ps = <dynamic>[];
    try {
      final r = await _ch.invokeMethod<String>('packList');
      if (r != null && r.isNotEmpty) ps = jsonDecode(r) as List;
    } catch (_) {}
    if (mounted) setState(() { lib = files; packs = ps; });
  }

  Future<void> _import() async {
    if (await prepareFolder() != null) {
      await Perm.reqStorage();
      return;
    }
    if (!mounted) return;
    final res = await Navigator.push<List<String>>(
      context,
      MaterialPageRoute(
        builder: (_) => MediaBrowser(
          title: 'استيراد استيكرات',
          storeKey: 'dirSticker',
          fallbackDir: _firstExistingDir(WaApp.stickerDirs()),
          exts: const ['webp', 'png', 'gif'],
          max: 200,
        ),
      ),
    );
    if (res == null || res.isEmpty) return;
    var ok = 0, failed = 0;
    final d = Directory(kStickerLibDir);
    await d.create(recursive: true);
    for (final src in res) {
      try {
        final ext = p.extension(src).toLowerCase();
        final dest =
            p.join(d.path, '${DateTime.now().microsecondsSinceEpoch}_$ok$ext');
        await File(src).copy(dest);
        ok++;
      } catch (_) {
        failed++;
      }
    }
    await _refresh();
    if (mounted) {
      snack(context,
          failed == 0 ? 'تم استيراد $ok استيكر للمكتبة' : 'تم استيراد $ok وتعذر $failed');
    }
  }

  Future<void> _buildPacks() async {
    if (lib.isEmpty) {
      snack(context, 'المكتبة فارغة — استوردي استيكرات أولاً');
      return;
    }
    setState(() => busy = true);
    try {
      final r = await _ch.invokeMethod<String>('buildPacks',
          {'paths': lib.map((f) => f.path).toList()});
      final ps = r == null || r.isEmpty ? <dynamic>[] : jsonDecode(r) as List;
      setState(() => packs = ps);
      snack(context, 'تم تجهيز ${ps.length} حزمة — أضيفيها لواتساب من القائمة');
    } on PlatformException catch (e) {
      snack(context, 'تعذر تجهيز الحزم: ${e.message}');
    } finally {
      if (mounted) setState(() => busy = false);
    }
  }

  Future<void> _addPack(Map pk) async {
    try {
      await _ch.invokeMethod<bool>('addPack',
          {'id': pk['id'], 'name': pk['name'], 'pkg': WaApp.pkg});
      final m = Map<String, dynamic>.from(_added);
      m['${pk['id']}'] = pk['version'];
      Store.data['stickerAdded'] = m;
      await Store.save();
      if (mounted) setState(() {});
    } on PlatformException catch (e) {
      if (mounted) {
        snack(
            context,
            e.code == 'NO_WA_BUSINESS'
                ? '${WaApp.label} غير مثبت على هذا الهاتف'
                : 'تعذر فتح واتساب: ${e.message}');
      }
    }
  }

  Future<void> _removeSticker(File f) async {
    final ok = await ask(context, 'حذف من المكتبة',
        'حذف هذا الاستيكر من المكتبة؟ (الحزم المضافة لواتساب لا تتغير حتى تحديثيها)',
        ok: 'حذف');
    if (!ok) return;
    try {
      await f.delete();
    } catch (_) {}
    await _refresh();
  }

  @override
  Widget build(BuildContext context) {
    final t = Theme.of(context);
    return Scaffold(
      appBar: AppBar(title: Row(children: [
        const Text('استيكراتي'),
        const SizedBox(width: 8),
        Container(
          padding: const EdgeInsets.symmetric(horizontal: 10, vertical: 3),
          decoration: BoxDecoration(
            color: t.colorScheme.primary,
            borderRadius: BorderRadius.circular(999),
          ),
          child: Text('${lib.length} استيكر',
              style: TextStyle(color: t.colorScheme.onPrimary, fontSize: 11.5)),
        ),
      ])),
      body: ListView(padding: const EdgeInsets.all(16), children: [
        if (busy) const LinearProgressIndicator(),
        Row(children: [
          Expanded(
            child: OutlinedButton.icon(
                onPressed: busy ? null : _import,
                icon: const Icon(Icons.download),
                label: const Text('استيراد من المجلدات')),
          ),
          const SizedBox(width: 8),
          Expanded(
            child: FilledButton.icon(
                onPressed: busy ? null : _buildPacks,
                icon: const Icon(Icons.add_to_photos),
                label: const Text('إضافة إلى واتساب')),
          ),
        ]),
        const SizedBox(height: 12),
        Card(
          child: Padding(
            padding: const EdgeInsets.all(12),
            child: Column(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  Text('المكتبة', style: t.textTheme.titleMedium),
                  const SizedBox(height: 4),
                  Text('مجلد A2R_Sender/stickers',
                      style: t.textTheme.bodySmall),
                  const SizedBox(height: 10),
                  if (lib.isEmpty)
                    Text('لا توجد استيكرات بعد — استوردي من مجلدات الهاتف.',
                        style: t.textTheme.bodySmall)
                  else
                    GridView.builder(
                      shrinkWrap: true,
                      physics: const NeverScrollableScrollPhysics(),
                      gridDelegate:
                          const SliverGridDelegateWithFixedCrossAxisCount(
                              crossAxisCount: 4,
                              mainAxisSpacing: 8,
                              crossAxisSpacing: 8),
                      itemCount: lib.length,
                      itemBuilder: (_, i) => GestureDetector(
                        onLongPress: () => _removeSticker(lib[i]),
                        child: ClipRRect(
                          borderRadius: BorderRadius.circular(12),
                          child: Container(
                            color: t.colorScheme.surfaceContainerHighest,
                            child: Image.file(lib[i], fit: BoxFit.contain,
                                errorBuilder: (_, __, ___) => const Center(
                                    child: Icon(Icons.broken_image))),
                          ),
                        ),
                      ),
                    ),
                  if (lib.isNotEmpty) ...[
                    const SizedBox(height: 6),
                    Text('ضغطة مطوّلة على أي استيكر لحذفه من المكتبة.',
                        style: t.textTheme.bodySmall),
                  ],
                ]),
          ),
        ),
        const SizedBox(height: 8),
        Container(
          padding: const EdgeInsets.all(10),
          decoration: BoxDecoration(
            color: const Color(0xFFFDF6E4),
            border: Border.all(color: const Color(0xFFF0DFAE)),
            borderRadius: BorderRadius.circular(10),
          ),
          child: const Text(
              'استيكر ناقص من واتساب؟ ابعتيه مرة في أي شات أو انسخيه لمجلد المكتبة — واتساب لا يخزن كل المستلم على التخزين المشترك.',
              style: TextStyle(fontSize: 12, color: Color(0xFF7A5B13))),
        ),
        const SizedBox(height: 12),
        Card(
          child: Padding(
            padding: const EdgeInsets.all(12),
            child: Column(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  Text('الحزم في واتساب', style: t.textTheme.titleMedium),
                  const SizedBox(height: 4),
                  Text('كل حزمة حتى ٣٠ استيكر (حد واتساب الرسمي).',
                      style: t.textTheme.bodySmall),
                  const SizedBox(height: 8),
                  if (packs.isEmpty)
                    Text('جهزي مكتبتك ثم اضغطي «إضافة إلى واتساب».',
                        style: t.textTheme.bodySmall)
                  else
                    for (final pk in packs.cast<Map>()) ...[
                      Builder(builder: (_) {
                        final addedV = _added['${pk['id']}'];
                        final upToDate = addedV != null && addedV == pk['version'];
                        return ListTile(
                          dense: true,
                          contentPadding: EdgeInsets.zero,
                          leading: CircleAvatar(
                              backgroundColor: t.colorScheme.primary,
                              child: const Text('A2R',
                                  style: TextStyle(
                                      color: Colors.white, fontSize: 11))),
                          title: Text('${pk['name']}'),
                          subtitle: Text(
                              '${pk['count']} استيكر${upToDate ? ' — مضافة' : addedV != null ? ' — تغيّرت بعد الإضافة' : ' — لم تُضف'}'),
                          trailing: upToDate
                              ? Text('✓ في واتساب',
                                  style: TextStyle(
                                      color: t.colorScheme.primary,
                                      fontWeight: FontWeight.bold,
                                      fontSize: 12))
                              : TextButton(
                                  onPressed: busy ? null : () => _addPack(pk),
                                  child: Text(addedV != null ? 'تحديث' : 'إضافة')),
                        );
                      }),
                    ],
                ]),
          ),
        ),
        const SizedBox(height: 8),
        Container(
          padding: const EdgeInsets.all(10),
          decoration: BoxDecoration(
            color: const Color(0xFFFDF6E4),
            border: Border.all(color: const Color(0xFFF0DFAE)),
            borderRadius: BorderRadius.circular(10),
          ),
          child: const Text(
              'الإرسال التلقائي للنصوص والصور يعمل كما هو دون تغيير. الاستيكر الحقيقي يُرسل من لوحة واتساب بعد إضافة الحزمة.',
              style: TextStyle(fontSize: 12, color: Color(0xFF7A5B13))),
        ),
      ]),
    );
  }
}

// ======================= الهيكل: تابان =======================
class Shell extends StatefulWidget {
  const Shell({super.key});
  @override
  State<Shell> createState() => _ShellState();
}

class _ShellState extends State<Shell> {
  int tab = 0;
  @override
  Widget build(BuildContext context) {
    final kb = MediaQuery.of(context).viewInsets.bottom > 0;
    return Scaffold(
      body: IndexedStack(
          index: tab, children: const [Home(), LabelTab(), StickersTab()]),
      bottomNavigationBar: kb
          ? null
          : NavigationBar(
              selectedIndex: tab,
              onDestinationSelected: (v) => setState(() => tab = v),
              destinations: const [
                NavigationDestination(
                    icon: Icon(Icons.send), label: 'إرسال لكل جروب'),
                NavigationDestination(
                    icon: Icon(Icons.label_outline), label: 'رسالة لمصنف'),
                NavigationDestination(
                    icon: Icon(Icons.sticky_note_2_outlined),
                    label: 'استيكراتي'),
              ],
            ),
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
      home: const Gate(),
    );
  }
}

// ======================= شاشة الرسالة =======================
class Home extends StatefulWidget {
  const Home({super.key});

  /// للاختبارات فقط: بيانات شيت جاهزة بدل قراءة الملف من الذاكرة.
  static ({List<G> groups, int dups, int off, String src})? debugSheet;
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
    if (s == AppLifecycleState.resumed && !loading) {
      Bubble.syncAll().then((_) {
        if (mounted) setState(() {});
      });
      _reload();
    }
  }

  Future<void> _init() async {
    if (Home.debugSheet == null) {
      await Store.load();
    }
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
    unawaited(cleanupUnusedMedia());
    await _reload();
  }

  Future<void> _reload() async {
    final dbg = Home.debugSheet;
    if (dbg != null) {
      if (!mounted) return;
      setState(() {
        srcPath = dbg.src;
        groups = dbg.groups;
        dups = dbg.dups;
        off = dbg.off;
        err = null;
        loading = false;
      });
      return;
    }
    var e = await prepareFolder();
    var g = <G>[];
    var dp = 0, of = 0;
    String? src;
    if (e == null) {
      src = findExcel();
      if (src == null) {
        e = 'لا يوجد ملف جروبات بعد.\nاضغطي "استيراد ملف Excel" من قائمة ⋮، أو ضعي ملف .xlsx داخل مجلد A2R_Sender في الذاكرة الداخلية.';
      } else {
        try {
          // القراءة والتحليل في الخلفية حتى لا تتجمد الواجهة مع الشيتات الكبيرة.
          final r = await compute(parseSheetFile, src);
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

  String _firstDir(List<String> c) {
    for (final d in c) {
      if (Directory(d).existsSync()) return d;
    }
    return kRoot;
  }

  Future<void> _addMedia(List<String> paths) async {
    var added = 0;
    var failed = 0;
    for (final f in paths) {
      try {
        media.add(await persistMedia(f));
        added++;
      } catch (_) {
        failed++;
      }
    }
    if (!mounted) return;
    setState(() {});
    _saveDraft();
    if (failed > 0) {
      snack(context, 'تمت إضافة $added وتعذر نسخ $failed ملف');
    } else if (added > 0) {
      snack(context, 'تمت إضافة $added ملف بنجاح');
    }
  }

  /// صورة من معرض الصور في الهاتف.
  Future<void> _gallery() async {
    if (kLimit[kind]! - media.length <= 0) return;
    try {
      final x = await ImagePicker().pickImage(source: ImageSource.gallery);
      if (x == null) return;
      await _addMedia([x.path]);
    } catch (e) {
      if (mounted) snack(context, 'تعذر فتح معرض الصور: $e');
    }
  }

  /// تصفح مجلدات الهاتف مع تذكر آخر مجلد (للاستيكرات أو الصور).
  Future<void> _browse() async {
    final room = kLimit[kind]! - media.length;
    if (room <= 0) return;
    if (await prepareFolder() != null) {
      await Perm.reqStorage();
      return;
    }
    if (!mounted) return;
    final st = kind == 'stickers';
    final res = await Navigator.push<List<String>>(
      context,
      MaterialPageRoute(
        builder: (_) => MediaBrowser(
          title: st ? 'اختيار استيكرات' : 'اختيار صورة',
          storeKey: st ? 'dirSticker' : 'dirImage',
          fallbackDir: st
              ? _firstDir(WaApp.stickerDirs())
              : _firstDir(['$kRoot/DCIM/Camera', '$kRoot/Pictures', '$kRoot/Download']),
          exts: st ? ['webp', 'png', 'gif'] : ['jpg', 'jpeg', 'png', 'webp'],
          max: room,
        ),
      ),
    );
    if (res == null || res.isEmpty) return;
    await _addMedia(res);
  }

  /// زر إضافة واحد يفتح اختيار المصدر: معرض الصور أو التصفح من مجلد.
  Future<void> _addChooser() async {
    final src = await showModalBottomSheet<String>(
      context: context,
      showDragHandle: true,
      builder: (ctx) => SafeArea(
        child: Column(mainAxisSize: MainAxisSize.min, children: [
          ListTile(
            leading: const Icon(Icons.photo_library),
            title: const Text('من معرض الصور'),
            onTap: () => Navigator.pop(ctx, 'gallery'),
          ),
          ListTile(
            leading: const Icon(Icons.folder_open),
            title: const Text('من مجلد في الهاتف — آخر مجلد محفوظ'),
            onTap: () => Navigator.pop(ctx, 'browse'),
          ),
          const SizedBox(height: 8),
        ]),
      ),
    );
    if (!mounted) return;
    if (src == 'gallery') {
      await _gallery();
    } else if (src == 'browse') {
      await _browse();
    }
  }

  Future<void> _start() async {
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
    await Bubble.reset();
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
    final t = Theme.of(context);
    return Scaffold(
      appBar: AppBar(
          title: Row(children: [
            const Text('A2R Sender'),
            const SizedBox(width: 8),
            Container(
              padding: const EdgeInsets.symmetric(horizontal: 10, vertical: 3),
              decoration: BoxDecoration(
                color: t.colorScheme.primary,
                borderRadius: BorderRadius.circular(999),
              ),
              child: Text(WaApp.label,
                  style:
                      TextStyle(color: t.colorScheme.onPrimary, fontSize: 11.5)),
            ),
          ]),
          actions: [
            PopupMenuButton<String>(
              tooltip: 'أدوات التطبيق',
              onSelected: (v) {
                if (v == 'import') {
                  _importExcel();
                } else if (v == 'groups') {
                  Navigator.push(
                      context,
                      MaterialPageRoute(
                          builder: (_) => WaGroupsScreen(groups: groups)));
                } else if (v == 'refresh') {
                  _reload();
                } else if (v == 'settings') {
                  showAppSettings(context, () {
                    if (mounted) setState(() {});
                  });
                }
              },
              itemBuilder: (_) => const [
                PopupMenuItem(
                    value: 'import',
                    child: ListTile(
                        dense: true,
                        contentPadding: EdgeInsets.zero,
                        leading: Icon(Icons.upload_file),
                        title: Text('استيراد ملف Excel'))),
                PopupMenuItem(
                    value: 'groups',
                    child: ListTile(
                        dense: true,
                        contentPadding: EdgeInsets.zero,
                        leading: Icon(Icons.groups),
                        title: Text('أسماء الجروبات من واتساب'))),
                PopupMenuItem(
                    value: 'refresh',
                    child: ListTile(
                        dense: true,
                        contentPadding: EdgeInsets.zero,
                        leading: Icon(Icons.refresh),
                        title: Text('تحديث من ملف Excel'))),
                PopupMenuItem(
                    value: 'settings',
                    child: ListTile(
                        dense: true,
                        contentPadding: EdgeInsets.zero,
                        leading: Icon(Icons.settings),
                        title: Text('إعدادات التطبيق'))),
              ],
            ),
          ]),
      body: ListView(padding: const EdgeInsets.all(16), children: [
        if (loading) const LinearProgressIndicator(),
        if (err != null)
          Card(
              color: t.colorScheme.errorContainer,
              child: Padding(
                  padding: const EdgeInsets.all(12),
                  child: Column(
                      crossAxisAlignment: CrossAxisAlignment.start,
                      children: [
                        Text(err!),
                        if (err!.contains('اسمحي')) ...[
                          const SizedBox(height: 8),
                          FilledButton.tonal(
                              onPressed: () => Perm.reqStorage(),
                              child: const Text('منح صلاحية الملفات')),
                        ],
                      ])))
        else
          Card(
            child: Padding(
              padding: const EdgeInsets.all(12),
              child: Column(
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    Row(children: [
                      Expanded(
                        child: Column(
                            crossAxisAlignment: CrossAxisAlignment.start,
                            children: [
                              Text('${groups.length} جروب جاهز',
                                  style: t.textTheme.titleMedium?.copyWith(
                                      fontWeight: FontWeight.bold)),
                              Text(
                                '${p.basename(srcPath ?? kFile)}'
                                '${dups > 0 ? ' • مكرر تم تجاهله: $dups' : ''}'
                                '${off > 0 ? ' • موقوف: $off' : ''}',
                                style: t.textTheme.bodySmall,
                              ),
                            ]),
                      ),
                      IconButton(
                          tooltip: 'تحديث من ملف Excel',
                          onPressed: _reload,
                          icon: const Icon(Icons.refresh)),
                    ]),
                    if (srcPath != null && srcPath != kFile) ...[
                      const SizedBox(height: 6),
                      Container(
                        width: double.infinity,
                        padding: const EdgeInsets.all(8),
                        decoration: BoxDecoration(
                          color: const Color(0xFFFDF3E3),
                          borderRadius: BorderRadius.circular(10),
                          border:
                              Border.all(color: const Color(0xFFF0DCB5)),
                        ),
                        child: Text(
                          'المصدر: ${p.basename(srcPath!)} — التطبيق يستخدم أحدث ملف Excel في المجلد تلقائياً. '
                          'لو هذا ليس الملف الصحيح احذفيه من المجلد أو استورديه من قائمة ⋮.',
                          style: const TextStyle(
                              fontSize: 12, color: Color(0xFF7A5A1E)),
                        ),
                      ),
                    ],
                    if (_cats.isNotEmpty) ...[
                      const SizedBox(height: 10),
                      Text('الإرسال إلى:', style: t.textTheme.titleSmall),
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
                        if (groups
                            .any((g) => g.cats.isEmpty && !Wa.inAnyWaCat(g)))
                          ChoiceChip(
                              label: Text(
                                  'بدون تصنيف (${groups.where((g) => inCat(g, kNoCat)).length})'),
                              selected: _cat == kNoCat,
                              onSelected: (_) => _setCat(kNoCat)),
                      ]),
                    ],
                    if (_cat != null) ...[
                      const SizedBox(height: 6),
                      Text('المحدد للإرسال: ${shown.length} جروب',
                          style: t.textTheme.bodySmall),
                    ],
                  ]),
            ),
          ),
        if (job != null) ...[
          const SizedBox(height: 8),
          FilledButton.tonalIcon(
            onPressed: () async {
              await Navigator.push(context,
                  MaterialPageRoute(builder: (_) => const SendScreen()));
              if (mounted) setState(() {});
            },
            icon: const Icon(Icons.play_arrow),
            label: Text(
                'استكمال المهمة غير المنتهية (${(job['st'] as List).where((s) => s != 'pending').length} من ${(job['st'] as List).length})'),
          ),
        ],
        const SizedBox(height: 12),
        SegmentedButton<String>(
          segments: [
            for (final k in kKinds.keys)
              ButtonSegment(
                  value: k,
                  label:
                      Text(kKindsShort[k]!, overflow: TextOverflow.ellipsis)),
          ],
          selected: {kind},
          showSelectedIcon: false,
          onSelectionChanged: (v) => _setKind(v.first),
        ),
        if (kind == 'stickers')
          Padding(
            padding: const EdgeInsets.only(top: 8),
            child: Text(
                'الاستيكرات هنا تُرسل كصور. لإرسالها كاستيكر حقيقي: تبويب «استيكراتي».',
                style: t.textTheme.bodySmall),
          ),
        if (kind == 'text' || kind == 'imageText') ...[
          const SizedBox(height: 12),
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
          const SizedBox(height: 12),
          Card(
            child: Padding(
              padding: const EdgeInsets.all(12),
              child: Column(
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    if (media.isNotEmpty) ...[
                      Wrap(spacing: 12, runSpacing: 12, children: [
                        for (var i = 0; i < media.length; i++)
                          Stack(children: [
                            ClipRRect(
                              borderRadius: BorderRadius.circular(12),
                              child: Container(
                                width: 120,
                                height: 120,
                                color: Colors.black12,
                                child: Image.file(File(media[i]),
                                    width: 120,
                                    height: 120,
                                    fit: kind == 'stickers'
                                        ? BoxFit.contain
                                        : BoxFit.cover,
                                    cacheWidth: 300,
                                    errorBuilder: (_, __, ___) => const Center(
                                        child: Icon(Icons.broken_image))),
                              ),
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
                      const SizedBox(height: 10),
                    ],
                    if (media.length < lim)
                      OutlinedButton.icon(
                        onPressed: _addChooser,
                        icon: const Icon(Icons.add),
                        label: Text(kind == 'stickers'
                            ? 'إضافة استيكرات (${media.length})'
                            : 'إضافة صورة'),
                      )
                    else
                      Text('وصلتي للحد الأقصى (${media.length})',
                          style: t.textTheme.bodySmall),
                  ]),
            ),
          ),
        ],
        const SizedBox(height: 16),
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

class _SendScreenState extends State<SendScreen> with WidgetsBindingObserver {
  late Map<String, dynamic> job;
  int? forced;
  bool opened = false;
  late String sendMode;

  @override
  void initState() {
    super.initState();
    job = Store.job!;
    final savedMode = Store.data['sendMode'];
    sendMode = (savedMode == 'share' || savedMode == 'review')
        ? savedMode as String
        : 'review';
    WidgetsBinding.instance.addObserver(this);
    _sync();
  }

  void _setSendMode(String m) {
    setState(() => sendMode = m);
    Store.data['sendMode'] = m;
    Store.save();
  }

  @override
  void dispose() {
    WidgetsBinding.instance.removeObserver(this);
    super.dispose();
  }

  @override
  void didChangeAppLifecycleState(AppLifecycleState s) {
    if (s == AppLifecycleState.resumed) _sync();
  }

  /// يأخذ ما أنجزته الفقاعة (تم/تخطي) أثناء وجودك في واتساب.
  Future<void> _sync() async {
    final r = await Bubble.sync('job');
    if (r.changed) forced = null;
    if (mounted) setState(() {});
  }

  /// لا شرط اختيار محتوى لبدء الإرسال: ما لا محتوى فيه يُتخطّى عند التنفيذ.
  bool get _emptyPayload =>
      (job['text'] as String).trim().isEmpty &&
      (job['media'] as List).isEmpty;

  Future<void> _startBubble(int i) async {
    if (_emptyPayload) {
      for (var j = 0; j < _st.length; j++) {
        if (_st[j] == 'pending') _st[j] = 'skipped';
      }
      await Store.save();
      if (!mounted) return;
      snack(context, 'لا محتوى للإرسال — تم تخطي الجروبات');
      setState(() {});
      return;
    }
    final media = (job['media'] as List).map((e) => e.toString()).toList();
    for (final m in media) {
      if (!File(m).existsSync()) {
        snack(context, 'ملف الصورة غير موجود، جهزي الرسالة من جديد');
        return;
      }
    }
    final err = await Bubble.start(jobSession(job, mode: sendMode), i);
    if (!mounted) return;
    if (err != null) {
      snack(context, err);
      return;
    }
    setState(() => opened = true);
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
    if (_emptyPayload) {
      snack(context, 'لا محتوى للإرسال — تم تخطي الجروب');
      await _mark(i, 'skipped');
      return;
    }
    final media = (job['media'] as List).map((e) => e.toString()).toList();
    for (final m in media) {
      if (!File(m).existsSync()) {
        snack(context, 'ملف الصورة غير موجود، جهزي الرسالة من جديد');
        return;
      }
    }
    await Clipboard.setData(ClipboardData(text: waKey(_g(i))));
    try {
      await _ch.invokeMethod<bool>('share',
          {'text': job['text'] as String, 'paths': media, 'pkg': WaApp.pkg});
      if (mounted) setState(() => opened = true);
    } on PlatformException catch (e) {
      if (!mounted) return;
      snack(
          context,
          e.code == 'NO_WA_BUSINESS'
              ? '${WaApp.label} غير مثبت على هذا الهاتف'
              : 'تعذر فتح واتساب: ${e.message}');
    }
  }

  Future<void> _mark(int i, String s) async {
    _st[i] = s;
    await Store.save();
    unawaited(Bubble.setStatus(i, s));
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
      unawaited(Bubble.setStatus(i, 'pending'));
    }
    if (mounted) setState(() => forced = i);
  }

  Future<void> _finish() async {
    final ok = await ask(context, 'إنهاء المهمة',
        'المتبقي: ${_count('pending')} جروب. سيتم إغلاق المهمة.',
        ok: 'إنهاء');
    if (!ok) return;
    await Bubble.reset();
    job['finished'] = true;
    await Store.save();
    unawaited(cleanupUnusedMedia());
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
        Card(
          child: Padding(
            padding: const EdgeInsets.all(12),
            child: Column(children: [
              LinearProgressIndicator(
                  value: total == 0 ? 0 : (done + skipped) / total,
                  minHeight: 8),
              const SizedBox(height: 8),
              Row(
                  mainAxisAlignment: MainAxisAlignment.spaceBetween,
                  children: [
                    Text('تم $done', style: t.textTheme.bodySmall),
                    Text('تخطي $skipped', style: t.textTheme.bodySmall),
                    Text('متبقي ${_count('pending')} من $total',
                        style: t.textTheme.bodySmall),
                  ]),
            ]),
          ),
        ),
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
          const SizedBox(height: 8),
          Card(
            child: Padding(
              padding: const EdgeInsets.all(16),
              child: Column(
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    Text('الجروب الحالي (${i + 1})',
                        style: t.textTheme.labelMedium),
                    const SizedBox(height: 4),
                    Text(_g(i).name, style: t.textTheme.headlineSmall),
                    if (_g(i).numbers.isNotEmpty || _g(i).teacher.isNotEmpty)
                      Text(
                        [
                          _g(i).numbers,
                          if (_g(i).teacher.isNotEmpty)
                            'المعلمة/المعلم: ${_g(i).teacher}',
                        ].where((e) => e.isNotEmpty).join(' • '),
                        style: t.textTheme.bodySmall,
                      ),
                    const SizedBox(height: 10),
                    Container(
                      padding: const EdgeInsets.symmetric(
                          horizontal: 8, vertical: 2),
                      decoration: BoxDecoration(
                        color: t.colorScheme.surfaceContainerHighest
                            .withOpacity(0.4),
                        borderRadius: BorderRadius.circular(10),
                        border:
                            Border.all(color: t.colorScheme.outlineVariant),
                      ),
                      child: Row(children: [
                        Expanded(
                            child: Text('كلمة البحث: ${waKey(_g(i))}',
                                style: t.textTheme.titleSmall)),
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
                    ),
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
                        width: 40,
                        height: 40,
                        cacheWidth: 100,
                        errorBuilder: (_, __, ___) =>
                            const Icon(Icons.broken_image)),
                  ),
                Expanded(
                    child: Text(
                        body.isEmpty
                            ? kindName(kind)
                            : '${kindName(kind)}: $body',
                        maxLines: 1,
                        overflow: TextOverflow.ellipsis,
                        style: t.textTheme.bodySmall)),
              ]),
            ),
          ),
          const SizedBox(height: 8),
          BubbleCard(
              startLabel: 'تشغيل الفقاعة من هذا الجروب',
              onStart: () => _startBubble(i)),
          const SizedBox(height: 8),
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
          const SizedBox(height: 8),
          Card(
            child: ExpansionTile(
              title: const Text('خيارات إضافية'),
              subtitle: const Text('طريقة البحث والفتح اليدوي'),
              childrenPadding: const EdgeInsets.fromLTRB(12, 0, 12, 12),
              children: [
                Align(
                    alignment: AlignmentDirectional.centerStart,
                    child: Text('طريقة البحث في واتساب:',
                        style: t.textTheme.titleSmall)),
                const SizedBox(height: 8),
                SegmentedButton<String>(
                  segments: const [
                    ButtonSegment(
                      value: 'review',
                      label: Text('بحث في الرئيسية (مراجعة)'),
                      icon: Icon(Icons.search),
                    ),
                    ButtonSegment(
                      value: 'share',
                      label: Text('مشاركة مباشرة'),
                      icon: Icon(Icons.share),
                    ),
                  ],
                  selected: {sendMode},
                  onSelectionChanged: (v) => _setSendMode(v.first),
                ),
                const SizedBox(height: 8),
                Text(
                  sendMode == 'review'
                      ? 'يبحث برقم الجروب في القائمة الرئيسية لتتأكد من بيانات الجروب ووصفه، ثم تضغط "تجهيز" من الشريط العائم فتظهر الرسالة جاهزة على الإرسال.'
                      : 'يفتح شاشة المشاركة بالرسالة مباشرة ويبحث برقم الجروب للإرسال السريع.',
                  style: t.textTheme.bodySmall,
                ),
                const Divider(height: 24),
                Align(
                  alignment: AlignmentDirectional.centerStart,
                  child: OutlinedButton.icon(
                      onPressed: () => _open(i),
                      icon: const Icon(Icons.chat),
                      label: Text('فتح ${WaApp.label} (يدوي)')),
                ),
                if (opened)
                  Padding(
                    padding: const EdgeInsets.only(top: 12),
                    child: Text(
                        'اختاري الجروب في واتساب (تم نسخ كلمة البحث) وأرسلي، ثم ارجعي هنا واضغطي "تم الإرسال".',
                        style: t.textTheme.bodyMedium
                            ?.copyWith(color: t.colorScheme.primary)),
                  ),
              ],
            ),
          ),
        ],
      ]),
    );
  }
}


// ======================= متصفح الملفات (استيكرات / صور) مع حفظ آخر مجلد =======================
class MediaBrowser extends StatefulWidget {
  final String title, storeKey, fallbackDir;
  final List<String> exts;
  final int max;
  const MediaBrowser({
    super.key,
    required this.title,
    required this.storeKey,
    required this.fallbackDir,
    required this.exts,
    required this.max,
  });
  @override
  State<MediaBrowser> createState() => _MediaBrowserState();
}

class _MediaBrowserState extends State<MediaBrowser> {
  late String dir;
  List<String> dirs = [];
  List<String> files = [];
  final List<String> sel = [];
  String? error;
  bool scanning = true;

  bool get atRoot => dir == kRoot || dir == '/' || p.dirname(dir) == dir;

  @override
  void initState() {
    super.initState();
    final saved = Store.data[widget.storeKey];
    dir = (saved is String && Directory(saved).existsSync())
        ? saved
        : widget.fallbackDir;
    _scan();
  }

  /// فحص المجلد في الخلفية (compute) حتى لا تتجمد الواجهة في المجلدات الكبيرة.
  void _scan() {
    final target = dir;
    scanning = true;
    error = null;
    Store.data[widget.storeKey] = target; // آخر مسار
    Store.save();
    compute(scanMediaDir, (dir: target, exts: widget.exts)).then((r) {
      if (!mounted || dir != target) return;
      setState(() {
        dirs = r.dirs;
        files = r.files;
        error = r.ok
            ? null
            : 'لا أستطيع فتح هذا المجلد. اسمحي للتطبيق بالوصول إلى الملفات من إعدادات الهاتف.';
        scanning = false;
      });
    });
  }

  void _go(String d) {
    setState(() {
      dir = d;
    });
    _scan();
  }

  void _toggle(String path) {
    setState(() {
      if (sel.contains(path)) {
        sel.remove(path);
      } else if (widget.max == 1) {
        sel
          ..clear()
          ..add(path);
      } else if (sel.length >= widget.max) {
        snack(context, 'الحد الأقصى ${widget.max}');
      } else {
        sel.add(path);
      }
    });
  }

  void _confirm() {
    Navigator.pop(context, List<String>.from(sel));
  }

  Future<void> _onBack() async {
    if (sel.isEmpty) {
      Navigator.pop(context);
      return;
    }
    final r = await showDialog<String>(
      context: context,
      builder: (ctx) => AlertDialog(
        title: const Text('تأكيد الاختيار'),
        content: Text('تم تحديد ${sel.length} ملف. هل تريد اعتمادهم أم إلغاء الاختيار؟'),
        actions: [
          TextButton(
              onPressed: () => Navigator.pop(ctx, 'discard'),
              child: const Text('إلغاء الاختيار')),
          FilledButton(
              onPressed: () => Navigator.pop(ctx, 'keep'),
              child: Text('اعتماد (${sel.length})')),
        ],
      ),
    );
    if (!mounted) return;
    if (r == 'keep') {
      _confirm();
    } else if (r == 'discard') {
      Navigator.pop(context);
    }
  }

  @override
  Widget build(BuildContext context) {
    final cs = Theme.of(context).colorScheme;
    return PopScope(
      canPop: sel.isEmpty,
      onPopInvokedWithResult: (didPop, _) {
        if (!didPop) _onBack();
      },
      child: Scaffold(
        appBar: AppBar(
          leading: IconButton(
            tooltip: 'رجوع',
            icon: const Icon(Icons.arrow_back),
            onPressed: _onBack,
          ),
          title: Text(widget.title),
          actions: [
            IconButton(
              tooltip: 'المجلد الأعلى',
              icon: const Icon(Icons.drive_folder_upload),
              onPressed: atRoot ? null : () => _go(p.dirname(dir)),
            ),
            if (sel.isNotEmpty)
              TextButton(
                onPressed: _confirm,
                child: Text('تم (${sel.length})'),
              ),
          ],
        ),
        body: Column(children: [
          Padding(
            padding: const EdgeInsets.symmetric(horizontal: 10, vertical: 6),
            child: Row(children: [
              IconButton(
                tooltip: 'المجلد الأعلى',
                icon: const Icon(Icons.arrow_upward, size: 20),
                onPressed: atRoot ? null : () => _go(p.dirname(dir)),
              ),
              Expanded(
                child: Text(dir,
                    textDirection: TextDirection.ltr,
                    maxLines: 2,
                    overflow: TextOverflow.ellipsis,
                    style: Theme.of(context).textTheme.bodySmall),
              ),
            ]),
          ),
          const Divider(height: 1),
          Expanded(
            child: scanning
                ? const Center(child: CircularProgressIndicator())
                : error != null
                    ? Center(
                        child: Padding(
                            padding: const EdgeInsets.all(16),
                            child: Text(error!)))
                    : CustomScrollView(slivers: [
                        SliverList(
                          delegate: SliverChildBuilderDelegate(
                            (c, i) => ListTile(
                              dense: true,
                              leading: const Icon(Icons.folder),
                              title: Text(p.basename(dirs[i])),
                              onTap: () => _go(dirs[i]),
                            ),
                            childCount: dirs.length,
                          ),
                        ),
                        if (files.isEmpty)
                          const SliverToBoxAdapter(
                            child: Padding(
                              padding: EdgeInsets.all(16),
                              child: Text('لا توجد ملفات مناسبة في هذا المجلد'),
                            ),
                          ),
                        SliverPadding(
                          padding: const EdgeInsets.all(8),
                          sliver: SliverGrid(
                            gridDelegate:
                                const SliverGridDelegateWithFixedCrossAxisCount(
                                    crossAxisCount: 3,
                                    mainAxisSpacing: 8,
                                    crossAxisSpacing: 8),
                            delegate: SliverChildBuilderDelegate(
                              (c, i) {
                                final path = files[i];
                                final selIdx = sel.indexOf(path);
                                final on = selIdx >= 0;
                                return GestureDetector(
                                  onTap: () => _toggle(path),
                                  child: Container(
                                    decoration: BoxDecoration(
                                      borderRadius: BorderRadius.circular(10),
                                      border: on
                                          ? Border.all(
                                              color: cs.primary, width: 3)
                                          : null,
                                    ),
                                    child: ClipRRect(
                                      borderRadius: BorderRadius.circular(8),
                                      child:
                                          Stack(fit: StackFit.expand, children: [
                                        Container(
                                          color: Colors.black12,
                                          child: Image.file(File(path),
                                              fit: BoxFit.contain,
                                              cacheWidth: 300,
                                              errorBuilder: (_, __, ___) =>
                                                  const Center(
                                                      child: Icon(
                                                          Icons.broken_image))),
                                        ),
                                        if (on)
                                          Container(
                                            color: Colors.black38,
                                            child: Center(
                                              child: CircleAvatar(
                                                radius: 18,
                                                backgroundColor: cs.primary,
                                                child: Text(
                                                  '${selIdx + 1}',
                                                  style: TextStyle(
                                                      color: cs.onPrimary,
                                                      fontWeight:
                                                          FontWeight.bold),
                                                ),
                                              ),
                                            ),
                                          ),
                                      ]),
                                    ),
                                  ),
                                );
                              },
                              childCount: files.length,
                            ),
                          ),
                        ),
                      ]),
          ),
        ]),
        bottomNavigationBar: SafeArea(
          child: Padding(
            padding: const EdgeInsets.fromLTRB(16, 8, 16, 12),
            child: FilledButton.icon(
              onPressed: sel.isEmpty ? null : _confirm,
              icon: const Icon(Icons.check_circle),
              label: Text(sel.isEmpty
                  ? 'اضغط على الاستيكر أو الصورة للتحديد'
                  : 'تأكيد الاختيار (${sel.length})'),
            ),
          ),
        ),
      ),
    );
  }
}

// ======================= تاب: رسالة موحدة لمصنف واتساب =======================
class LabelTab extends StatefulWidget {
  const LabelTab({super.key});

  /// للاختبارات فقط: تخطي انتظار التحميل من الذاكرة.
  static bool debugSkipLoad = false;
  @override
  State<LabelTab> createState() => _LabelTabState();
}

class _LabelTabState extends State<LabelTab> with WidgetsBindingObserver {
  final TextEditingController text = TextEditingController();
  String kind = 'text';
  List<String> media = [];
  bool ready = false;

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
    if (s == AppLifecycleState.resumed && ready) {
      Bubble.syncAll().then((_) {
        if (mounted) setState(() {});
      });
    }
  }

  Future<void> _init() async {
    if (!LabelTab.debugSkipLoad) {
      await Store.load();
    }
    Wa.invalidate();
    final d = Store.data['draft2'];
    if (d is Map<String, dynamic>) {
      kind = (d['kind'] as String?) ?? 'text';
      if (!kKinds.containsKey(kind)) kind = 'text';
      text.text = (d['text'] as String?) ?? '';
      media = ((d['media'] as List?) ?? const [])
          .map((e) => e.toString())
          .where((x) => File(x).existsSync())
          .toList();
    }
    if (!mounted) return;
    setState(() => ready = true);
  }

  void _save() {
    Store.data['draft2'] = {'kind': kind, 'text': text.text, 'media': media};
    Store.save();
  }

  void _setKind(String k) {
    setState(() {
      if ((k == 'stickers') != (kind == 'stickers')) media = [];
      kind = k;
      final lim = kLimit[k]!;
      if (media.length > lim) media = media.sublist(0, lim);
    });
    _save();
  }

  String _firstDir(List<String> c) {
    for (final d in c) {
      if (Directory(d).existsSync()) return d;
    }
    return kRoot;
  }

  Future<void> _addMedia(List<String> paths) async {
    var added = 0;
    var failed = 0;
    for (final f in paths) {
      try {
        media.add(await persistMedia(f));
        added++;
      } catch (_) {
        failed++;
      }
    }
    if (!mounted) return;
    setState(() {});
    _save();
    if (failed > 0) {
      snack(context, 'تمت إضافة $added وتعذر نسخ $failed ملف');
    } else if (added > 0) {
      snack(context, 'تمت إضافة $added ملف بنجاح');
    }
  }

  Future<void> _gallery() async {
    if (kLimit[kind]! - media.length <= 0) return;
    try {
      final x = await ImagePicker().pickImage(source: ImageSource.gallery);
      if (x == null) return;
      await _addMedia([x.path]);
    } catch (e) {
      if (mounted) snack(context, 'تعذر فتح معرض الصور: $e');
    }
  }

  Future<void> _browse() async {
    final room = kLimit[kind]! - media.length;
    if (room <= 0) return;
    if (await prepareFolder() != null) {
      await Perm.reqStorage();
      return;
    }
    if (!mounted) return;
    final st = kind == 'stickers';
    final res = await Navigator.push<List<String>>(
      context,
      MaterialPageRoute(
        builder: (_) => MediaBrowser(
          title: st ? 'اختيار استيكرات' : 'اختيار صورة',
          storeKey: st ? 'dirSticker' : 'dirImage',
          fallbackDir: st
              ? _firstDir(WaApp.stickerDirs())
              : _firstDir(['$kRoot/DCIM/Camera', '$kRoot/Pictures', '$kRoot/Download']),
          exts: st ? ['webp', 'png', 'gif'] : ['jpg', 'jpeg', 'png', 'webp'],
          max: room,
        ),
      ),
    );
    if (res == null || res.isEmpty) return;
    await _addMedia(res);
  }

  /// زر إضافة واحد يفتح اختيار المصدر: معرض الصور أو التصفح من مجلد.
  Future<void> _addChooser() async {
    final src = await showModalBottomSheet<String>(
      context: context,
      showDragHandle: true,
      builder: (ctx) => SafeArea(
        child: Column(mainAxisSize: MainAxisSize.min, children: [
          ListTile(
            leading: const Icon(Icons.photo_library),
            title: const Text('من معرض الصور'),
            onTap: () => Navigator.pop(ctx, 'gallery'),
          ),
          ListTile(
            leading: const Icon(Icons.folder_open),
            title: const Text('من مجلد في الهاتف — آخر مجلد محفوظ'),
            onTap: () => Navigator.pop(ctx, 'browse'),
          ),
          const SizedBox(height: 8),
        ]),
      ),
    );
    if (!mounted) return;
    if (src == 'gallery') {
      await _gallery();
    } else if (src == 'browse') {
      await _browse();
    }
  }

  /// يجهز مهمة "المصنف المفتوح": هو يفتح المصنف في واتساب بنفسه،
  /// والفقاعة تفتح محادثات القائمة الظاهرة واحدة واحدة وتجهز الرسالة.
  Future<void> _start() async {
    final old = Store.ujob != null;
    final ok = await ask(
        context,
        'تجهيز الإرسال',
        '${kindName(kind)}${kind == 'stickers' ? ' (${media.length})' : ''} لكل محادثات المصنف.'
        '\n\nافتحي المصنف بنفسك في واتساب واتركيه مفتوحاً أمامك، ثم شغّلي الفقاعة من الشاشة التالية. '
        'الفقاعة هتفتح كل محادثة في القائمة واحدة واحدة وتجهّز الرسالة.'
        '${old ? '\nسيتم إغلاق المهمة غير المنتهية السابقة.' : ''}',
        ok: 'متابعة');
    if (!ok || !mounted) return;
    await Bubble.reset();
    Store.data['ujob'] = {
      'kind': kind,
      'text': (kind == 'text' || kind == 'imageText') ? text.text.trim() : '',
      'media': kind == 'text' ? <String>[] : media,
      'items': <String>[],
      'st': <String>[],
      'finished': false,
    };
    _save();
    await Store.save();
    if (!mounted) return;
    await Navigator.push(
        context, MaterialPageRoute(builder: (_) => const LabelSendScreen()));
    if (mounted) setState(() {});
  }

  @override
  Widget build(BuildContext context) {
    final t = Theme.of(context);
    final lim = kLimit[kind]!;
    final job = Store.ujob;
    return Scaffold(
      appBar: AppBar(title: const Text('رسالة لمصنف'), actions: [
        IconButton(
            tooltip: 'إعدادات التطبيق',
            onPressed: () => showAppSettings(context, () {
                  if (mounted) setState(() {});
                }),
            icon: const Icon(Icons.settings)),
      ]),
      body: !ready
          ? const Center(child: CircularProgressIndicator())
          : ListView(padding: const EdgeInsets.all(16), children: [
              Card(
                child: Padding(
                  padding: const EdgeInsets.all(12),
                  child: Column(
                      crossAxisAlignment: CrossAxisAlignment.start,
                      children: [
                        Text('إزاي بيشتغل؟', style: t.textTheme.titleSmall),
                        const SizedBox(height: 6),
                        const Text(
                            '١) افتحي المصنف بنفسك في واتساب واتركيه مفتوحاً.\n'
                            '٢) ارجعي هنا واضغطي "تجهيز الإرسال للمصنف".\n'
                            '٣) الفقاعة هتفتح كل محادثة في القائمة اللي قدامك واحدة واحدة وتجهّز الرسالة، وإنتِ بس تراجعي وتبعتي.'),
                      ]),
                ),
              ),
              if (job != null) ...[
                const SizedBox(height: 8),
                FilledButton.tonalIcon(
                  onPressed: () async {
                    await Navigator.push(
                        context,
                        MaterialPageRoute(
                            builder: (_) => const LabelSendScreen()));
                    if (mounted) setState(() {});
                  },
                  icon: const Icon(Icons.play_arrow),
                  label: Text(
                      'استكمال مهمة المصنف غير المنتهية (${(job['st'] as List).where((x) => x != 'pending').length} من ${(job['st'] as List).length})'),
                ),
              ],
              const SizedBox(height: 12),
              SegmentedButton<String>(
                segments: [
                  for (final k in kKinds.keys)
                    ButtonSegment(
                        value: k,
                        label: Text(kKindsShort[k]!,
                            overflow: TextOverflow.ellipsis)),
                ],
                selected: {kind},
                showSelectedIcon: false,
                onSelectionChanged: (v) => _setKind(v.first),
              ),
              if (kind == 'text' || kind == 'imageText') ...[
                const SizedBox(height: 12),
                TextField(
                  controller: text,
                  minLines: 4,
                  maxLines: 10,
                  onChanged: (_) => _save(),
                  decoration: const InputDecoration(
                      labelText: 'نص الرسالة', border: OutlineInputBorder()),
                ),
              ],
              if (lim > 0) ...[
                const SizedBox(height: 12),
                Card(
                  child: Padding(
                    padding: const EdgeInsets.all(12),
                    child: Column(
                        crossAxisAlignment: CrossAxisAlignment.start,
                        children: [
                          if (media.isNotEmpty) ...[
                            Wrap(spacing: 12, runSpacing: 12, children: [
                              for (var i = 0; i < media.length; i++)
                                Stack(children: [
                                  ClipRRect(
                                    borderRadius: BorderRadius.circular(12),
                                    child: Container(
                                      width: 120,
                                      height: 120,
                                      color: Colors.black12,
                                      child: Image.file(File(media[i]),
                                          width: 120,
                                          height: 120,
                                          fit: kind == 'stickers'
                                              ? BoxFit.contain
                                              : BoxFit.cover,
                                          cacheWidth: 300,
                                          errorBuilder: (_, __, ___) =>
                                              const Center(
                                                  child: Icon(
                                                      Icons.broken_image))),
                                    ),
                                  ),
                                  PositionedDirectional(
                                    top: 0,
                                    end: 0,
                                    child: IconButton.filled(
                                      iconSize: 18,
                                      onPressed: () {
                                        setState(() => media.removeAt(i));
                                        _save();
                                      },
                                      icon: const Icon(Icons.close),
                                    ),
                                  ),
                                ]),
                            ]),
                            const SizedBox(height: 10),
                          ],
                          if (media.length < lim)
                            OutlinedButton.icon(
                              onPressed: _addChooser,
                              icon: const Icon(Icons.add),
                              label: Text(kind == 'stickers'
                                  ? 'إضافة استيكرات (${media.length})'
                                  : 'إضافة صورة'),
                            )
                          else
                            Text('وصلتي للحد الأقصى (${media.length})',
                                style: t.textTheme.bodySmall),
                        ]),
                  ),
                ),
              ],
              const SizedBox(height: 16),
              FilledButton.icon(
                  onPressed: _start,
                  icon: const Icon(Icons.play_arrow),
                  label: const Text('تجهيز الإرسال للمصنف')),
              const SizedBox(height: 24),
            ]),
    );
  }
}

class LabelSendScreen extends StatefulWidget {
  const LabelSendScreen({super.key});
  @override
  State<LabelSendScreen> createState() => _LabelSendScreenState();
}

class _LabelSendScreenState extends State<LabelSendScreen>
    with WidgetsBindingObserver {
  late Map<String, dynamic> job;
  bool started = false;

  @override
  void initState() {
    super.initState();
    job = Store.ujob!;
    WidgetsBinding.instance.addObserver(this);
    _sync();
  }

  @override
  void dispose() {
    WidgetsBinding.instance.removeObserver(this);
    super.dispose();
  }

  @override
  void didChangeAppLifecycleState(AppLifecycleState s) {
    if (s == AppLifecycleState.resumed) _sync();
  }

  List get _items => job['items'] as List;
  List get _st => job['st'] as List;
  int _count(String s) => _st.where((x) => x == s).length;
  List<String> get _media =>
      (job['media'] as List).map((e) => e.toString()).toList();

  Future<void> _sync() async {
    await Bubble.sync('ujob');
    if (mounted) setState(() {});
  }

  bool get _emptyPayload =>
      (job['text'] as String).trim().isEmpty && _media.isEmpty;

  bool _mediaOk() {
    for (final m in _media) {
      if (!File(m).existsSync()) {
        snack(context, 'ملف الصورة غير موجود، جهزي الرسالة من جديد');
        return false;
      }
    }
    return true;
  }

  Future<void> _startBubble() async {
    if (_emptyPayload) {
      snack(context, 'لا محتوى للإرسال — أضيفي نصاً أو صورة أو استيكراً من الشاشة السابقة');
      return;
    }
    if (!_mediaOk()) return;
    final err = await Bubble.start(labelSession(job), -1);
    if (!mounted) return;
    if (err != null) {
      snack(context, err);
      return;
    }
    setState(() => started = true);
  }

  Future<void> _manualShare() async {
    if (_emptyPayload) {
      snack(context, 'لا محتوى للإرسال — أضيفي نصاً أو صورة أو استيكراً من الشاشة السابقة');
      return;
    }
    if (!_mediaOk()) return;
    try {
      await _ch.invokeMethod<bool>('share',
          {'text': job['text'] as String, 'paths': _media, 'pkg': WaApp.pkg});
    } on PlatformException catch (e) {
      if (!mounted) return;
      snack(
          context,
          e.code == 'NO_WA_BUSINESS'
              ? '${WaApp.label} غير مثبت على هذا الهاتف'
              : 'تعذر فتح واتساب: ${e.message}');
    }
  }

  Future<void> _setSt(int i, String s) async {
    _st[i] = s;
    await Store.save();
    unawaited(Bubble.setStatus(i, s));
    if (mounted) setState(() {});
  }

  Future<void> _finish() async {
    final ok = await ask(context, 'إنهاء المهمة',
        'المتبقي: ${_count('pending')} محادثة. سيتم إغلاق المهمة والفقاعة.',
        ok: 'إنهاء');
    if (!ok) return;
    await Bubble.reset();
    job['finished'] = true;
    await Store.save();
    unawaited(cleanupUnusedMedia());
    if (mounted) Navigator.pop(context);
  }

  Icon _icon(String s) => s == 'done'
      ? const Icon(Icons.check_circle, color: Colors.green)
      : s == 'skipped'
          ? const Icon(Icons.skip_next, color: Colors.orange)
          : const Icon(Icons.radio_button_unchecked);

  @override
  Widget build(BuildContext context) {
    final t = Theme.of(context);
    final done = _count('done'), skipped = _count('skipped');
    final total = _items.length;
    final media = _media;
    final body = (job['text'] as String).trim();

    return Scaffold(
      appBar: AppBar(title: const Text('إرسال للمصنف المفتوح'), actions: [
        IconButton(
            tooltip: 'إنهاء المهمة',
            onPressed: _finish,
            icon: const Icon(Icons.stop_circle_outlined)),
      ]),
      body: ListView(padding: const EdgeInsets.all(16), children: [
        Card(
          child: Padding(
            padding: const EdgeInsets.all(12),
            child:
                Column(crossAxisAlignment: CrossAxisAlignment.start, children: [
              Row(children: [
                Expanded(
                    child: Text('تم $done  •  تخطي $skipped',
                        style: t.textTheme.titleSmall)),
                if (total > 0)
                  Text('$total محادثة', style: t.textTheme.bodySmall),
              ]),
              const SizedBox(height: 6),
              Text('المحادثات بتظهر هنا أول بأول مع كل واحدة بتفتحها الفقاعة.',
                  style: t.textTheme.bodySmall),
            ]),
          ),
        ),
        const SizedBox(height: 8),
        Card(
          child: Padding(
            padding: const EdgeInsets.all(12),
            child:
                Column(crossAxisAlignment: CrossAxisAlignment.start, children: [
              Text('قبل التشغيل:', style: t.textTheme.titleSmall),
              const SizedBox(height: 6),
              const Text(
                  'افتحي المصنف في واتساب واتركيه مفتوحاً قدامك، ثم اضغطي زر التشغيل بالأسفل. '
                  'الفقاعة هتفتح كل محادثة في القائمة وتجهّز الرسالة. بعد كل إرسال اضغطي "التالي" في الفقاعة.'),
            ]),
          ),
        ),
        const SizedBox(height: 8),
        BubbleCard(
            startLabel: started
                ? 'استكمال بالفقاعة'
                : 'تشغيل الفقاعة على المصنف المفتوح',
            onStart: _startBubble),
        const SizedBox(height: 8),
        OutlinedButton.icon(
            onPressed: _manualShare,
            icon: const Icon(Icons.chat),
            label: const Text('فتح واتساب بالرسالة (يدوي)')),
        const SizedBox(height: 8),
        Card(
          child: Padding(
            padding: const EdgeInsets.all(12),
            child: Row(children: [
              for (final m in media.take(2))
                Padding(
                  padding: const EdgeInsetsDirectional.only(end: 8),
                  child: Image.file(File(m),
                      width: 40,
                      height: 40,
                      cacheWidth: 100,
                      errorBuilder: (_, __, ___) =>
                          const Icon(Icons.broken_image)),
                ),
              Expanded(
                  child: Text(
                      body.isEmpty
                          ? kindName(job['kind'] as String)
                          : '${kindName(job['kind'] as String)}: $body',
                      maxLines: 1,
                      overflow: TextOverflow.ellipsis,
                      style: t.textTheme.bodySmall)),
            ]),
          ),
        ),
        if (total > 0) ...[
          const SizedBox(height: 12),
          Text('المحادثات', style: t.textTheme.titleSmall),
          for (var i = 0; i < total; i++)
            ListTile(
              dense: true,
              leading: _icon(_st[i] as String),
              title: Text(_items[i].toString()),
              trailing: PopupMenuButton<String>(
                onSelected: (v) => _setSt(i, v),
                itemBuilder: (_) => const [
                  PopupMenuItem(value: 'done', child: Text('تعليم: تم')),
                  PopupMenuItem(value: 'skipped', child: Text('تعليم: تخطي')),
                  PopupMenuItem(value: 'pending', child: Text('إعادته للمعلّق')),
                ],
              ),
            ),
        ],
        const SizedBox(height: 24),
      ]),
    );
  }
}
