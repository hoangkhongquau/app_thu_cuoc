import 'dart:convert';
import 'dart:io';
import 'package:csv/csv.dart';
import 'package:excel/excel.dart' hide Border;
import 'package:file_picker/file_picker.dart';
import 'package:flutter/material.dart';
import 'package:flutter/services.dart';
import 'package:geolocator/geolocator.dart';
import 'package:intl/intl.dart';
import 'package:path/path.dart' as p;
import 'package:path_provider/path_provider.dart';
import 'package:share_plus/share_plus.dart';
import 'package:sqflite/sqflite.dart';
import 'package:sqflite_common_ffi/sqflite_ffi.dart';

void main() {
  WidgetsFlutterBinding.ensureInitialized();
  if (Platform.isWindows || Platform.isLinux) {
    sqfliteFfiInit();
    databaseFactory = databaseFactoryFfi;
  }
  runApp(const ThuCuocApp());
}

class ThuCuocApp extends StatelessWidget {
  const ThuCuocApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'Quản Lý Thu Cước',
      debugShowCheckedModeBanner: false,
      theme: ThemeData(
        colorScheme: ColorScheme.fromSeed(seedColor: const Color(0xFFD32F2F)),
        useMaterial3: true,
      ),
      home: const HomeScreen(),
    );
  }
}

// ==================== DATABASE HELPER (SQLITE 2 BẢNG) ====================
class DbHelper {
  static Database? _db;

  static Future<Database> get db async {
    if (_db != null) return _db!;
    final dbPath = await getDatabasesPath();
    _db = await openDatabase(
      p.join(dbPath, 'thu_cuoc_lang_co.db'),
      version: 3,
      onCreate: (db, version) async {
        // Bảng 1: Sổ tay khách hàng (Lưu vĩnh viễn Hình thức & Ghi chú của từng thuê bao)
        await db.execute('''
          CREATE TABLE so_tay_khach_hang (
            ma_tb TEXT PRIMARY KEY,
            ho_ten TEXT,
            dia_chi TEXT,
            sdt TEXT,
            dich_vu TEXT,
            to_thu_goc TEXT,
            khu_vuc TEXT,
            hinh_thuc TEXT,
            ghi_chu TEXT,
            latitude REAL,
            longitude REAL
          )
        ''');

        // Bảng 2: Danh sách cước đang quản lý trong kỳ
        await db.execute('''
          CREATE TABLE khach_hang (
            ma_tb TEXT PRIMARY KEY,
            ho_ten TEXT,
            dia_chi TEXT,
            sdt TEXT,
            dich_vu TEXT,
            to_thu_goc TEXT,
            khu_vuc TEXT,
            hinh_thuc TEXT,
            ghi_chu TEXT,
            dau_ky INTEGER DEFAULT 0,
            ton_cuoc INTEGER DEFAULT 0,
            du_kien_chan TEXT,
            trang_thai INTEGER DEFAULT 0,
            ngay_cap_nhat TEXT,
            gio_cap_nhat TEXT,
            con_ton_cuoc INTEGER DEFAULT 1
          )
        ''');
        // trang_thai: 0=Chưa thu, 1=Đã thu TM, 2=Đã gạch nợ cước, 3=Đã gạch chưa thu, 4=Hẹn lại
      },
      onUpgrade: (db, oldVersion, newVersion) async {
        await db.execute('''
          CREATE TABLE IF NOT EXISTS so_tay_khach_hang (
            ma_tb TEXT PRIMARY KEY,
            ho_ten TEXT,
            dia_chi TEXT,
            sdt TEXT,
            dich_vu TEXT,
            to_thu_goc TEXT,
            khu_vuc TEXT,
            hinh_thuc TEXT,
            ghi_chu TEXT,
            latitude REAL,
            longitude REAL
          )
        ''');
        if (oldVersion < 2) {
          await db.execute(
            'ALTER TABLE khach_hang ADD COLUMN dau_ky INTEGER DEFAULT 0',
          );
          await db.execute(
            'ALTER TABLE khach_hang ADD COLUMN gio_cap_nhat TEXT',
          );
          await db.execute(
            'ALTER TABLE khach_hang ADD COLUMN con_ton_cuoc INTEGER DEFAULT 1',
          );
        }
        final columns = await db.rawQuery(
          'PRAGMA table_info(so_tay_khach_hang)',
        );
        final columnNames = columns.map((column) => column['name']).toSet();
        if (!columnNames.contains('latitude')) {
          await db.execute(
            'ALTER TABLE so_tay_khach_hang ADD COLUMN latitude REAL',
          );
        }
        if (!columnNames.contains('longitude')) {
          await db.execute(
            'ALTER TABLE so_tay_khach_hang ADD COLUMN longitude REAL',
          );
        }
      },
    );
    return _db!;
  }

  static String cleanPhone(String raw) {
    String s = raw.trim();
    if (s.endsWith('.0')) s = s.substring(0, s.length - 2);
    s = s.replaceAll(RegExp(r'[^0-9]'), '');
    if (s.isEmpty || s == '0') return '';
    if (!s.startsWith('0')) s = '0$s';
    return s;
  }

  static int cleanAmount(String raw) {
    String s = raw.trim();
    if (s.endsWith('.0')) s = s.substring(0, s.length - 2);
    s = s.replaceAll(RegExp(r'[^0-9]'), '');
    if (s.isEmpty) return 0;
    int val = int.tryParse(s) ?? 0;
    if (val > 0 && val < 2000) val = val * 1000;
    return val;
  }

  static String normalizeArea(String raw) {
    String lower = raw.toLowerCase();
    if (lower.contains('an cư tân')) return 'An Cư Tân';
    if (lower.contains('lập an')) return 'Lập An';
    if (lower.contains('loan lý')) return 'Loan Lý';
    if (lower.contains('khu vực 1')) return 'Khu vực 1';
    return 'Khác';
  }

  // Đọc danh sách các bảng (sheet) từ file .xlsx hoặc .csv
  static List<List<List<String>>> parseFileToSheets(
    String fileName,
    Uint8List bytes,
  ) {
    final lowerName = fileName.toLowerCase();
    if (lowerName.endsWith('.xlsx')) {
      final excel = Excel.decodeBytes(bytes);
      List<List<List<String>>> allSheets = [];
      for (var tableName in excel.tables.keys) {
        final sheet = excel.tables[tableName];
        if (sheet == null || sheet.rows.length < 2) continue;
        List<List<String>> rows = [];
        for (var row in sheet.rows) {
          rows.add(
            row.map((cell) => cell?.value?.toString().trim() ?? '').toList(),
          );
        }
        allSheets.add(rows);
      }
      return allSheets;
    } else {
      final content = utf8.decode(bytes, allowMalformed: true);
      final rawRows = const CsvToListConverter(
        eol: '\n',
        shouldParseNumbers: false,
      ).convert(content.replaceAll('\r\n', '\n'));
      List<List<String>> rows = rawRows
          .map((r) => r.map((e) => e.toString().trim()).toList())
          .toList();
      return [rows];
    }
  }

  // Nạp dữ liệu vào CSDL với 2 chế độ:
  // isTonCuocFile = false -> Nạp file Đầu kỳ / Học sổ tay từ file cũ
  // isTonCuocFile = true  -> Nạp file Tồn cước hàng ngày (Loại bỏ thuê bao đã tự TT/đã thu trên hệ thống)
  static Future<int> importRowsList(
    List<List<List<String>>> sheets, {
    required bool isTonCuocFile,
  }) async {
    final database = await db;

    // 1. Lấy Sổ tay khách hàng vĩnh viễn
    final soTayRows = await database.query('so_tay_khach_hang');
    final Map<String, Map<String, dynamic>> soTayMap = {
      for (var r in soTayRows)
        r['ma_tb'].toString(): Map<String, dynamic>.from(r),
    };

    // 2. Lấy danh sách cước hiện tại
    final currentRows = await database.query('khach_hang');
    final Map<String, Map<String, dynamic>> currentMap = {
      for (var r in currentRows)
        r['ma_tb'].toString(): Map<String, dynamic>.from(r),
    };

    int totalProcessed = 0;
    Set<String> accountsInLastSheet = {};

    for (int sIdx = 0; sIdx < sheets.length; sIdx++) {
      final rows = sheets[sIdx];
      if (rows.length < 2) continue;

      // Tìm dòng tiêu đề (có thể nằm ở dòng 0 hoặc dòng 1-3 nếu file Excel có tiêu đề phụ)
      int headerRowIdx = -1;
      for (int rIdx = 0; rIdx < rows.length && rIdx < 10; rIdx++) {
        final lineStr = rows[rIdx].join(' ').toLowerCase();
        if (lineStr.contains('số thuê bao') || lineStr.contains('mã account')) {
          headerRowIdx = rIdx;
          break;
        }
      }
      if (headerRowIdx == -1) continue;

      final headers = rows[headerRowIdx].map((e) => e.trim()).toList();
      int idxMa = headers.indexWhere(
        (h) => h.contains('Số thuê bao') || h.contains('Mã Account'),
      );
      int idxTen = headers.indexWhere((h) => h.contains('Họ và tên'));
      int idxDiaChi = headers.indexWhere((h) => h.contains('Địa chỉ'));
      int idxSdt = headers.indexWhere((h) => h.contains('Số điện thoại'));
      int idxDichVu = headers.indexWhere((h) => h == 'Dịch vụ');
      int idxDauKy = headers.indexWhere((h) => h.contains('Đầu kỳ'));
      int idxTonCuoc = headers.indexWhere((h) => h.contains('Tồn cước'));
      int idxToThu = headers.indexWhere((h) => h == 'Tổ thu');
      int idxChan = headers.indexWhere((h) => h.contains('Dự kiến chặn'));
      int idxHinhThuc = headers.indexWhere((h) => h.contains('Hinh Thuc'));
      int idxGhiChu = headers.indexWhere(
        (h) => h.contains('Ghi Chu') || h.contains('Ghi chú'),
      );
      int idxCot1 = headers.indexWhere((h) => h == '1');

      if (idxMa == -1) continue;

      final batch = database.batch();
      final bool isLastSheet = (sIdx == sheets.length - 1);

      for (int i = headerRowIdx + 1; i < rows.length; i++) {
        final r = rows[i];
        if (r.length <= idxMa) continue;
        String maTb = r[idxMa].trim();
        if (maTb.isEmpty || maTb.toLowerCase() == 'số thuê bao') continue;

        if (isLastSheet) {
          accountsInLastSheet.add(maTb);
        }

        final oldSoTay = soTayMap[maTb];
        final oldCurrent = currentMap[maTb];

        String hoTen = idxTen != -1 && r.length > idxTen && r[idxTen].isNotEmpty
            ? r[idxTen].trim()
            : (oldCurrent?['ho_ten'] ?? oldSoTay?['ho_ten'] ?? '');
        String diaChi =
            idxDiaChi != -1 && r.length > idxDiaChi && r[idxDiaChi].isNotEmpty
            ? r[idxDiaChi].trim()
            : (oldCurrent?['dia_chi'] ?? oldSoTay?['dia_chi'] ?? '');
        String sdtRaw = idxSdt != -1 && r.length > idxSdt ? r[idxSdt] : '';
        String sdt = cleanPhone(sdtRaw);
        if (sdt.isEmpty) {
          sdt = oldCurrent?['sdt'] ?? oldSoTay?['sdt'] ?? '';
        }

        String dichVu =
            idxDichVu != -1 && r.length > idxDichVu && r[idxDichVu].isNotEmpty
            ? r[idxDichVu].trim()
            : (oldCurrent?['dich_vu'] ?? oldSoTay?['dich_vu'] ?? '');
        String toThuGoc =
            idxToThu != -1 && r.length > idxToThu && r[idxToThu].isNotEmpty
            ? r[idxToThu].trim()
            : (oldCurrent?['to_thu_goc'] ?? oldSoTay?['to_thu_goc'] ?? '');
        String khuVuc = normalizeArea(toThuGoc);

        int dauKy = idxDauKy != -1 && r.length > idxDauKy
            ? cleanAmount(r[idxDauKy])
            : (oldCurrent?['dau_ky'] ?? 0);
        int tonCuoc = idxTonCuoc != -1 && r.length > idxTonCuoc
            ? cleanAmount(r[idxTonCuoc])
            : 0;

        String duKienChan = '';
        if (idxChan != -1 && r.length > idxChan) {
          duKienChan = r[idxChan].trim().replaceAll('.0', '');
        }

        // Khôi phục Hình thức & Ghi chú từ Sổ tay vĩnh viễn
        String hinhThuc =
            oldCurrent?['hinh_thuc'] ?? oldSoTay?['hinh_thuc'] ?? '';
        String ghiChu = oldCurrent?['ghi_chu'] ?? oldSoTay?['ghi_chu'] ?? '';

        if (idxHinhThuc != -1 && r.length > idxHinhThuc) {
          String htMoi = r[idxHinhThuc].trim();
          if (htMoi.isNotEmpty && !htMoi.contains('ERROR')) {
            hinhThuc = htMoi.toUpperCase();
          }
        }

        if (idxGhiChu != -1 && r.length > idxGhiChu) {
          String gcMoi = r[idxGhiChu].trim();
          if (gcMoi.isNotEmpty && !gcMoi.contains('ERROR')) {
            if (idxHinhThuc == -1) {
              String upper = gcMoi.toUpperCase();
              if ([
                'CK',
                'TT',
                'TTT',
                'HD',
                'IN HÓA ĐƠN',
                'ACCOUNT MỚI',
              ].contains(upper)) {
                hinhThuc = upper;
              } else if (upper.startsWith('TT ')) {
                hinhThuc = 'TT';
                ghiChu = gcMoi.substring(3).trim();
              } else if (upper.startsWith('CK ')) {
                hinhThuc = 'CK';
                ghiChu = gcMoi.substring(3).trim();
              } else {
                ghiChu = gcMoi;
              }
            } else {
              ghiChu = gcMoi;
            }
          }
        }

        // Lưu vào Sổ tay vĩnh viễn nếu thuê bao có Hình thức hoặc Ghi chú
        if (hinhThuc.isNotEmpty || ghiChu.isNotEmpty || oldSoTay != null) {
          final soTayRecord = {
            'ma_tb': maTb,
            'ho_ten': hoTen,
            'dia_chi': diaChi,
            'sdt': sdt,
            'dich_vu': dichVu,
            'to_thu_goc': toThuGoc,
            'khu_vuc': khuVuc,
            'hinh_thuc': hinhThuc,
            'ghi_chu': ghiChu,
            'latitude': oldSoTay?['latitude'],
            'longitude': oldSoTay?['longitude'],
          };
          soTayMap[maTb] = soTayRecord;
          batch.insert(
            'so_tay_khach_hang',
            soTayRecord,
            conflictAlgorithm: ConflictAlgorithm.replace,
          );
        }

        // Xử lý cột '1' theo đúng logic cũ của bạn
        int trangThai = oldCurrent?['trang_thai'] ?? 0;
        String ngayCapNhat = oldCurrent?['ngay_cap_nhat'] ?? '';
        String gioCapNhat = oldCurrent?['gio_cap_nhat'] ?? '';

        if (idxCot1 != -1 && r.length > idxCot1) {
          String c1 = r[idxCot1].trim().toLowerCase();
          if (c1.contains('gạch')) {
            trangThai = 3; // Đã gạch chưa thu
          }
        }

        // Chỉ cập nhật vào bảng khach_hang đang thu ở sheet cuối cùng (hoặc file 1 sheet)
        if (isLastSheet) {
          final khRecord = {
            'ma_tb': maTb,
            'ho_ten': hoTen,
            'dia_chi': diaChi,
            'sdt': sdt,
            'dich_vu': dichVu,
            'to_thu_goc': toThuGoc,
            'khu_vuc': khuVuc,
            'hinh_thuc': hinhThuc,
            'ghi_chu': ghiChu,
            'dau_ky': dauKy,
            'ton_cuoc': tonCuoc,
            'du_kien_chan': duKienChan,
            'trang_thai': trangThai,
            'ngay_cap_nhat': ngayCapNhat,
            'gio_cap_nhat': gioCapNhat,
            'con_ton_cuoc': 1,
          };
          currentMap[maTb] = khRecord;
          batch.insert(
            'khach_hang',
            khRecord,
            conflictAlgorithm: ConflictAlgorithm.replace,
          );
          totalProcessed++;
        }
      }

      await batch.commit(noResult: true);
    }

    // Nếu là nạp file TỒN CƯỚC hàng ngày:
    // Những thuê bao không còn nằm trong file Tồn cước nghĩa là đã thu/tự thanh toán xong trên hệ thống
    // -> Ẩn khỏi danh sách tồn cước (ngoại trừ các thuê bao bạn đang đánh dấu trong ngày hoặc "Gạch chưa thu")
    if (isTonCuocFile && accountsInLastSheet.isNotEmpty) {
      final allCurrent = await database.query('khach_hang');
      final batchClean = database.batch();
      for (var row in allCurrent) {
        String ma = row['ma_tb'].toString();
        int st = (row['trang_thai'] as int?) ?? 0;
        if (!accountsInLastSheet.contains(ma)) {
          // Nếu chưa thu (st == 0) mà biến mất khỏi file tồn cước -> Đánh dấu con_ton_cuoc = 0 để loại ra
          if (st == 0) {
            batchClean.delete(
              'khach_hang',
              where: 'ma_tb = ?',
              whereArgs: [ma],
            );
          } else {
            batchClean.update(
              'khach_hang',
              {'con_ton_cuoc': 0},
              where: 'ma_tb = ?',
              whereArgs: [ma],
            );
          }
        }
      }
      await batchClean.commit(noResult: true);
    }

    return totalProcessed;
  }
}

// ==================== MÀN HÌNH CHÍNH ====================
class HomeScreen extends StatefulWidget {
  const HomeScreen({super.key});

  @override
  State<HomeScreen> createState() => _HomeScreenState();
}

class _NoteEditResult {
  const _NoteEditResult({
    required this.method,
    required this.note,
    required this.clearCoordinates,
  });

  final String method;
  final String note;
  final bool clearCoordinates;
}

class _NoteMethodDialog extends StatefulWidget {
  const _NoteMethodDialog({
    required this.customerName,
    required this.initialMethod,
    required this.initialNote,
    required this.hasCoordinates,
  });

  final String customerName;
  final String initialMethod;
  final String initialNote;
  final bool hasCoordinates;

  @override
  State<_NoteMethodDialog> createState() => _NoteMethodDialogState();
}

class _NoteMethodDialogState extends State<_NoteMethodDialog> {
  late final TextEditingController _noteController;
  late String _selectedMethod;
  bool _clearCoordinates = false;

  @override
  void initState() {
    super.initState();
    _noteController = TextEditingController(text: widget.initialNote);
    _selectedMethod = widget.initialMethod;
  }

  @override
  void dispose() {
    _noteController.dispose();
    super.dispose();
  }

  @override
  Widget build(BuildContext context) {
    return AlertDialog(
      title: Text(
        widget.customerName,
        style: const TextStyle(fontSize: 18, fontWeight: FontWeight.bold),
      ),
      content: Column(
        mainAxisSize: MainAxisSize.min,
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          const Text(
            'Hình thức thu cố định (Lưu vào Sổ tay):',
            style: TextStyle(fontWeight: FontWeight.w600),
          ),
          const SizedBox(height: 8),
          Wrap(
            spacing: 8,
            children: ['TT', 'CK', 'TTT', 'IN HÓA ĐƠN', 'HD'].map((method) {
              return ChoiceChip(
                label: Text(method),
                selected: _selectedMethod == method,
                onSelected: (selected) =>
                    setState(() => _selectedMethod = selected ? method : ''),
              );
            }).toList(),
          ),
          const SizedBox(height: 16),
          TextField(
            controller: _noteController,
            decoration: const InputDecoration(
              labelText: 'Ghi chú đường đi / Nhắc nhở',
              hintText: 'VD: Gần Thu Phượng, Ghé a râu lấy...',
              border: OutlineInputBorder(),
            ),
            maxLines: 2,
          ),
          if (widget.hasCoordinates)
            CheckboxListTile(
              contentPadding: EdgeInsets.zero,
              title: const Text('Xóa tọa độ GPS đã lưu'),
              value: _clearCoordinates,
              onChanged: (value) =>
                  setState(() => _clearCoordinates = value ?? false),
            ),
        ],
      ),
      actions: [
        TextButton(
          onPressed: () {
            FocusScope.of(context).unfocus();
            Navigator.of(context).pop();
          },
          child: const Text('Hủy'),
        ),
        FilledButton(
          onPressed: () {
            FocusScope.of(context).unfocus();
            Navigator.of(context).pop(
              _NoteEditResult(
                method: _selectedMethod,
                note: _noteController.text.trim(),
                clearCoordinates: _clearCoordinates,
              ),
            );
          },
          child: const Text('Lưu lại'),
        ),
      ],
    );
  }
}

class _HomeScreenState extends State<HomeScreen> {
  List<Map<String, dynamic>> _allData = [];
  List<Map<String, dynamic>> _filteredData = [];
  int _soTayCount = 0;
  bool _nearbyMode = false;
  bool _isLocating = false;
  Position? _myPos;
  double? _radiusMeters = 500;
  int _statusBeforeNearby = -1;

  String _searchQuery = '';
  String _selectedArea = 'Tất cả';
  String _selectedMethod =
      'Có đánh dấu HT'; // Mặc định lọc khách có đánh dấu TT/CK/HD...
  int _selectedStatus =
      -1; // -1: Tất cả, 99: Tác động hôm nay, 0: Chưa thu, 1: Đã thu TM, 2: Đã gạch nợ cước, 3: Đã gạch chưa thu, 4: Hẹn lại
  bool _onlyUrgentBlock = false;

  final currencyFmt = NumberFormat.currency(
    locale: 'vi_VN',
    symbol: 'đ',
    decimalDigits: 0,
  );

  String get _todayStr => DateFormat('yyyy-MM-dd').format(DateTime.now());

  @override
  void initState() {
    super.initState();
    _loadData();
  }

  Future<void> _loadData() async {
    final database = await DbHelper.db;
    final rows = await database.query(
      'khach_hang',
      where: 'con_ton_cuoc = ?',
      whereArgs: [1],
      orderBy: 'ho_ten ASC',
    );
    final coordinateRows = await database.query(
      'so_tay_khach_hang',
      columns: ['ma_tb', 'latitude', 'longitude'],
    );
    final coordinatesByAccount = {
      for (final row in coordinateRows) row['ma_tb'].toString(): row,
    };
    final soTayRows = await database.rawQuery(
      'SELECT COUNT(*) as cnt FROM so_tay_khach_hang WHERE hinh_thuc != "" OR ghi_chu != ""',
    );
    final int stCount = Sqflite.firstIntValue(soTayRows) ?? 0;

    final data = rows
        .map((row) {
          final coordinates = coordinatesByAccount[row['ma_tb'].toString()];
          return {
            ...row,
            'latitude': coordinates?['latitude'],
            'longitude': coordinates?['longitude'],
            '_formatted_ton_cuoc': currencyFmt.format(row['ton_cuoc'] ?? 0),
          };
        })
        .toList(growable: false);
    if (!mounted) return;
    setState(() {
      _allData = data;
      _soTayCount = stCount;
      _filteredData = _filterData(data);
    });
  }

  List<Map<String, dynamic>> _filterData(Iterable<Map<String, dynamic>> data) {
    final q = _searchQuery.toLowerCase().trim();
    final today = _todayStr;

    final filtered = data.where((item) {
      if (_nearbyMode) {
        final latitude = item['latitude'];
        final longitude = item['longitude'];
        if (item['trang_thai'] != 0 ||
            _myPos == null ||
            latitude is! num ||
            longitude is! num) {
          return false;
        }
        final distance = Geolocator.distanceBetween(
          _myPos!.latitude,
          _myPos!.longitude,
          latitude.toDouble(),
          longitude.toDouble(),
        );
        if (_radiusMeters != null && distance > _radiusMeters!) return false;
        item['_distanceMeters'] = distance;
      }

      // Nếu chọn tab "Tác động hôm nay" (status == 99)
      if (_nearbyMode) {
        if (item['trang_thai'] != 0) return false;
      } else if (_selectedStatus == 99) {
        final String ngay = (item['ngay_cap_nhat'] ?? '').toString();
        final int st = (item['trang_thai'] as int?) ?? 0;
        if (ngay != today || st == 0) return false;
      } else if (_selectedStatus != -1) {
        if (item['trang_thai'] != _selectedStatus) return false;
      }

      if (!_nearbyMode &&
          _selectedArea != 'Tất cả' &&
          item['khu_vuc'] != _selectedArea) {
        return false;
      }

      if (!_nearbyMode && _selectedMethod != 'Tất cả') {
        String ht = (item['hinh_thuc'] ?? '').toString().toUpperCase().trim();
        String gc = (item['ghi_chu'] ?? '').toString().trim();
        if (_selectedMethod == 'Có đánh dấu HT') {
          if (ht.isEmpty && gc.isEmpty) return false;
        } else if (_selectedMethod == 'Trống') {
          if (ht.isNotEmpty) return false;
        } else if (ht != _selectedMethod) {
          return false;
        }
      }

      if (!_nearbyMode && _onlyUrgentBlock) {
        String chan = (item['du_kien_chan'] ?? '').toString();
        if (chan.isEmpty) return false;
      }

      if (q.isNotEmpty) {
        String ten = (item['ho_ten'] ?? '').toString().toLowerCase();
        String dc = (item['dia_chi'] ?? '').toString().toLowerCase();
        String ma = (item['ma_tb'] ?? '').toString().toLowerCase();
        String sdt = (item['sdt'] ?? '').toString().toLowerCase();
        String gc = (item['ghi_chu'] ?? '').toString().toLowerCase();
        if (!ten.contains(q) &&
            !dc.contains(q) &&
            !ma.contains(q) &&
            !sdt.contains(q) &&
            !gc.contains(q)) {
          return false;
        }
      }
      return true;
    }).toList();
    if (_nearbyMode) {
      filtered.sort(
        (a, b) => (a['_distanceMeters'] as double).compareTo(
          b['_distanceMeters'] as double,
        ),
      );
    }
    return filtered;
  }

  void _showLocationMessage(String message) {
    if (!mounted) return;
    ScaffoldMessenger.of(context)
      ..clearSnackBars()
      ..showSnackBar(SnackBar(content: Text(message)));
  }

  Future<Position?> _getCurrentPosition() async {
    if (!Platform.isAndroid && !Platform.isIOS) {
      _showLocationMessage('Định vị chỉ hỗ trợ trên Android và iPhone.');
      return null;
    }
    try {
      if (!await Geolocator.isLocationServiceEnabled()) {
        _showLocationMessage('Hãy bật dịch vụ định vị trên điện thoại.');
        return null;
      }

      var permission = await Geolocator.checkPermission();
      if (permission == LocationPermission.denied) {
        permission = await Geolocator.requestPermission();
      }
      if (permission == LocationPermission.denied ||
          permission == LocationPermission.deniedForever) {
        _showLocationMessage(
          permission == LocationPermission.deniedForever
              ? 'Quyền vị trí bị từ chối vĩnh viễn. Hãy bật quyền trong Cài đặt.'
              : 'Ứng dụng cần quyền vị trí để định vị.',
        );
        return null;
      }

      return await Geolocator.getCurrentPosition(
        locationSettings: const LocationSettings(
          accuracy: LocationAccuracy.high,
          timeLimit: Duration(seconds: 15),
        ),
      );
    } on Exception catch (error) {
      _showLocationMessage('Không lấy được vị trí GPS: $error');
      return null;
    }
  }

  Future<void> _saveCustomerLocation(Map<String, dynamic> item) async {
    final position = await _getCurrentPosition();
    if (position == null || !mounted) return;

    final database = await DbHelper.db;
    await database.insert('so_tay_khach_hang', {
      'ma_tb': item['ma_tb'],
      'ho_ten': item['ho_ten'] ?? '',
      'dia_chi': item['dia_chi'] ?? '',
      'sdt': item['sdt'] ?? '',
      'dich_vu': item['dich_vu'] ?? '',
      'to_thu_goc': item['to_thu_goc'] ?? '',
      'khu_vuc': item['khu_vuc'] ?? '',
      'hinh_thuc': item['hinh_thuc'] ?? '',
      'ghi_chu': item['ghi_chu'] ?? '',
      'latitude': position.latitude,
      'longitude': position.longitude,
    }, conflictAlgorithm: ConflictAlgorithm.replace);
    await _loadData();
    _showLocationMessage('Đã lưu vị trí GPS cho thuê bao ${item['ma_tb']}.');
  }

  Future<void> _toggleNearbyMode() async {
    if (_nearbyMode) {
      setState(() {
        _nearbyMode = false;
        _myPos = null;
        _selectedStatus = _statusBeforeNearby;
        _filteredData = _filterData(_allData);
      });
      return;
    }
    if (_isLocating) return;
    setState(() => _isLocating = true);
    final position = await _getCurrentPosition();
    if (!mounted) return;
    setState(() => _isLocating = false);
    if (position == null) return;

    setState(() {
      _statusBeforeNearby = _selectedStatus;
      _selectedStatus = 0;
      _nearbyMode = true;
      _myPos = position;
      _filteredData = _filterData(_allData);
    });
  }

  String _formatDistance(double distanceMeters) {
    if (distanceMeters >= 1000) {
      return 'Cách ${(distanceMeters / 1000).toStringAsFixed(1)} km';
    }
    return 'Cách ${distanceMeters.round()} m';
  }

  void _applyFilters([VoidCallback? update]) {
    setState(() {
      update?.call();
      _filteredData = _filterData(_allData);
    });
  }

  Future<void> _updateStatus(String maTb, int newStatus) async {
    final database = await DbHelper.db;
    final now = DateTime.now();
    final today = DateFormat('yyyy-MM-dd').format(now);
    final timeStr = DateFormat('HH:mm').format(now);
    await database.update(
      'khach_hang',
      {
        'trang_thai': newStatus,
        'ngay_cap_nhat': newStatus == 0 ? '' : today,
        'gio_cap_nhat': newStatus == 0 ? '' : timeStr,
      },
      where: 'ma_tb = ?',
      whereArgs: [maTb],
    );
    await _loadData();
  }

  Future<void> _copyAccount(String maTb) async {
    await Clipboard.setData(ClipboardData(text: maTb));
    if (!mounted) return;
    ScaffoldMessenger.of(context).clearSnackBars();
    ScaffoldMessenger.of(context).showSnackBar(
      SnackBar(
        content: Text('Đã sao chép Account: $maTb'),
        duration: const Duration(seconds: 1),
        behavior: SnackBarBehavior.floating,
      ),
    );
  }

  Future<void> _editNoteAndMethod(Map<String, dynamic> item) async {
    final result = await showDialog<_NoteEditResult>(
      context: context,
      builder: (_) => _NoteMethodDialog(
        customerName: (item['ho_ten'] ?? item['ma_tb']).toString(),
        initialMethod: (item['hinh_thuc'] ?? '').toString(),
        initialNote: (item['ghi_chu'] ?? '').toString(),
        hasCoordinates: item['latitude'] != null && item['longitude'] != null,
      ),
    );
    if (result == null || !mounted) return;

    final database = await DbHelper.db;
    final String maTb = item['ma_tb'].toString();

    // Cập nhật đồng thời vào bảng hiện tại và Sổ tay vĩnh viễn
    await database.update(
      'khach_hang',
      {'hinh_thuc': result.method, 'ghi_chu': result.note},
      where: 'ma_tb = ?',
      whereArgs: [maTb],
    );
    await database.insert('so_tay_khach_hang', {
      'ma_tb': maTb,
      'ho_ten': item['ho_ten'] ?? '',
      'dia_chi': item['dia_chi'] ?? '',
      'sdt': item['sdt'] ?? '',
      'dich_vu': item['dich_vu'] ?? '',
      'to_thu_goc': item['to_thu_goc'] ?? '',
      'khu_vuc': item['khu_vuc'] ?? '',
      'hinh_thuc': result.method,
      'ghi_chu': result.note,
      'latitude': result.clearCoordinates ? null : item['latitude'],
      'longitude': result.clearCoordinates ? null : item['longitude'],
    }, conflictAlgorithm: ConflictAlgorithm.replace);
    await _loadData();
  }

  Future<void> _pickAndImportFile() async {
    // Hỏi người dùng loại file muốn nạp: Đầu kỳ hay Tồn cước hàng ngày
    final bool? isTonCuoc = await showDialog<bool>(
      context: context,
      builder: (ctx) => AlertDialog(
        title: const Text(
          'Chọn chế độ nạp file (.xlsx / .csv)',
          style: TextStyle(fontSize: 17, fontWeight: FontWeight.bold),
        ),
        content: Column(
          mainAxisSize: MainAxisSize.min,
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            Text(
              'Sổ tay đang nhớ: $_soTayCount thuê bao có Hình thức / Ghi chú.',
              style: TextStyle(
                fontSize: 13,
                color: Colors.teal.shade800,
                fontWeight: FontWeight.w600,
              ),
            ),
            const SizedBox(height: 12),
            const Text(
              '• File Đầu Kỳ / File Cũ: Nạp toàn bộ danh sách và tự động học các cột Hình thức, Ghi chú lưu vào Sổ tay.\n'
              '• File Tồn Cước Hàng Ngày: Cập nhật tồn cước mới, tự động loại ra các thuê bao đã thu xong trên hệ thống.',
              style: TextStyle(fontSize: 13.5),
            ),
          ],
        ),
        actions: [
          TextButton(
            onPressed: () => Navigator.pop(ctx, null),
            child: const Text('Hủy'),
          ),
          OutlinedButton(
            onPressed: () => Navigator.pop(ctx, false),
            child: const Text('Nạp Đầu Kỳ / Học File Cũ'),
          ),
          FilledButton(
            onPressed: () => Navigator.pop(ctx, true),
            child: const Text('Nạp Tồn Cước Ngày'),
          ),
        ],
      ),
    );

    if (isTonCuoc == null) return;

    final result = await FilePicker.pickFile(
      type: FileType.custom,
      allowedExtensions: ['xlsx', 'csv'],
    );
    if (result != null) {
      final bytes = await result.readAsBytes();
      final sheets = DbHelper.parseFileToSheets(result.name, bytes);
      final count = await DbHelper.importRowsList(
        sheets,
        isTonCuocFile: isTonCuoc,
      );
      await _loadData();
      if (mounted) {
        ScaffoldMessenger.of(context).showSnackBar(
          SnackBar(
            content: Text(
              isTonCuoc
                  ? 'Đã đồng bộ file Tồn cước ($count thuê bao còn tồn)!'
                  : 'Đã nạp Đầu kỳ & lưu Sổ tay ($count thuê bao)!',
            ),
          ),
        );
      }
    }
  }

  Future<void> _exportCsv() async {
    List<List<dynamic>> rows = [
      [
        'Số thuê bao',
        'Trạng thái thu',
        'Hinh Thuc',
        'Ghi Chu',
        'Họ và tên',
        'Địa chỉ TBC',
        'Số điện thoại',
        'Dịch vụ',
        'Tồn cước',
        'Tổ thu',
        'Dự kiến chặn',
        'Ngày tác động',
      ],
    ];

    String statusText(int st) {
      switch (st) {
        case 1:
          return 'Đã thu TM';
        case 2:
          return 'Đã gạch nợ cước';
        case 3:
          return 'đã gạch chưa thu';
        case 4:
          return 'Hẹn lại';
        default:
          return '';
      }
    }

    for (var r in _allData) {
      rows.add([
        r['ma_tb'],
        statusText(r['trang_thai'] ?? 0),
        r['hinh_thuc'],
        r['ghi_chu'],
        r['ho_ten'],
        r['dia_chi'],
        r['sdt'],
        r['dich_vu'],
        r['ton_cuoc'],
        r['to_thu_goc'],
        r['du_kien_chan'],
        r['ngay_cap_nhat'],
      ]);
    }

    String csvData = const ListToCsvConverter().convert(rows);
    final dir = await getTemporaryDirectory();
    final path = p.join(dir.path, 'Ket_Qua_Thu_Cuoc.csv');
    final file = File(path);
    await file.writeAsString('\uFEFF$csvData', encoding: utf8);
    await SharePlus.instance.share(
      ShareParams(files: [XFile(path)], text: 'Báo cáo kết quả thu cước'),
    );
  }

  @override
  Widget build(BuildContext context) {
    int tongTienMat = 0;
    int tongDaGachNoCuoc = 0;
    int tongDaGachChuaThu = 0;
    int soDaThuTienMat = 0;
    int soTacDongHomNay = 0;
    final today = _todayStr;

    for (var r in _allData) {
      int st = r['trang_thai'] ?? 0;
      int tien = r['ton_cuoc'] ?? 0;
      String ngay = (r['ngay_cap_nhat'] ?? '').toString();

      if (st == 1) {
        tongTienMat += tien;
        soDaThuTienMat++;
      } else if (st == 2) {
        tongDaGachNoCuoc += tien;
      } else if (st == 3) {
        tongDaGachChuaThu += tien;
      }
      if (st != 0 && ngay == today) {
        soTacDongHomNay++;
      }
    }

    return Scaffold(
      appBar: AppBar(
        title: Text(
          _selectedStatus == 99
              ? 'Tác động hôm nay ($soTacDongHomNay)'
              : 'Thu Cước (${_filteredData.length}/${_allData.length})',
          style: const TextStyle(fontSize: 18),
        ),
        backgroundColor: const Color(0xFFD32F2F),
        foregroundColor: Colors.white,
        actions: [
          IconButton(
            icon: Icon(
              _isLocating
                  ? Icons.location_searching
                  : _nearbyMode
                  ? Icons.near_me
                  : Icons.near_me_outlined,
            ),
            tooltip: _nearbyMode
                ? 'Tắt tìm nhà quanh đây'
                : 'Định vị và tìm nhà cần thu quanh đây',
            onPressed: _isLocating ? null : _toggleNearbyMode,
          ),
          if (_nearbyMode)
            PopupMenuButton<int>(
              tooltip: 'Chọn bán kính tìm kiếm',
              initialValue: _radiusMeters?.round() ?? 0,
              onSelected: (meters) {
                setState(() {
                  _radiusMeters = meters == 0 ? null : meters.toDouble();
                  _filteredData = _filterData(_allData);
                });
              },
              itemBuilder: (context) => const [
                PopupMenuItem(value: 200, child: Text('Trong 200 m')),
                PopupMenuItem(value: 500, child: Text('Trong 500 m')),
                PopupMenuItem(value: 1000, child: Text('Trong 1 km')),
                PopupMenuItem(value: 0, child: Text('Tất cả nhà có GPS')),
              ],
            ),
          // Nút bật nhanh danh sách các thẻ tác động trong ngày hôm nay
          IconButton(
            icon: Badge(
              isLabelVisible: soTacDongHomNay > 0,
              label: Text('$soTacDongHomNay'),
              child: Icon(
                _selectedStatus == 99 ? Icons.today : Icons.history_edu,
              ),
            ),
            tooltip: 'Xem các thuê bao tác động hôm nay',
            onPressed: () {
              _applyFilters(() {
                if (_selectedStatus == 99) {
                  _nearbyMode = false;
                  _myPos = null;
                  _selectedStatus = -1;
                } else {
                  _nearbyMode = false;
                  _myPos = null;
                  _selectedStatus = 99;
                  _selectedMethod = 'Tất cả';
                  _selectedArea = 'Tất cả';
                }
              });
            },
          ),
          IconButton(
            icon: const Icon(Icons.upload_file),
            tooltip: 'Nạp file Excel (.xlsx) hoặc CSV',
            onPressed: _pickAndImportFile,
          ),
          IconButton(
            icon: const Icon(Icons.download),
            tooltip: 'Xuất file CSV báo cáo',
            onPressed: _exportCsv,
          ),
        ],
      ),
      body: Column(
        children: [
          // 1. BẢNG ĐỐI SOÁT TIỀN QUỸ TRONG NGÀY
          Container(
            padding: const EdgeInsets.symmetric(horizontal: 10, vertical: 8),
            color: Colors.grey.shade100,
            child: Row(
              children: [
                _buildSummaryBox(
                  'Tiền mặt cầm',
                  currencyFmt.format(tongTienMat),
                  Colors.green.shade700,
                ),
                const SizedBox(width: 6),
                _buildSummaryBox(
                  'Đã gạch nợ cước',
                  currencyFmt.format(tongDaGachNoCuoc),
                  Colors.blue.shade700,
                ),
                const SizedBox(width: 6),
                _buildSummaryBox(
                  'Gạch chưa thu',
                  currencyFmt.format(tongDaGachChuaThu),
                  Colors.orange.shade800,
                ),
              ],
            ),
          ),

          // 2. THANH TÌM KIẾM NHANH
          Padding(
            padding: const EdgeInsets.fromLTRB(10, 6, 10, 4),
            child: TextField(
              decoration: InputDecoration(
                hintText:
                    'Tìm tên, mã account, số nhà (389), ghi chú (a râu)...',
                prefixIcon: const Icon(Icons.search),
                suffixIcon: _searchQuery.isNotEmpty
                    ? IconButton(
                        icon: const Icon(Icons.clear),
                        onPressed: () => _applyFilters(() => _searchQuery = ''),
                      )
                    : null,
                isDense: true,
                contentPadding: const EdgeInsets.symmetric(vertical: 10),
                border: OutlineInputBorder(
                  borderRadius: BorderRadius.circular(10),
                ),
              ),
              onChanged: (val) => _applyFilters(() => _searchQuery = val),
            ),
          ),

          // 3. BỘ LỌC KHU VỰC & HÌNH THỨC
          SingleChildScrollView(
            scrollDirection: Axis.horizontal,
            padding: const EdgeInsets.symmetric(horizontal: 10),
            child: Row(
              children: [
                for (final area in [
                  'Tất cả',
                  'An Cư Tân',
                  'Lập An',
                  'Loan Lý',
                  'Khu vực 1',
                ])
                  Padding(
                    padding: const EdgeInsets.only(right: 6),
                    child: ChoiceChip(
                      label: Text(area, style: const TextStyle(fontSize: 12)),
                      selected: _selectedArea == area,
                      onSelected: (_) =>
                          _applyFilters(() => _selectedArea = area),
                    ),
                  ),
              ],
            ),
          ),
          SingleChildScrollView(
            scrollDirection: Axis.horizontal,
            padding: const EdgeInsets.symmetric(horizontal: 10, vertical: 2),
            child: Row(
              children: [
                FilterChip(
                  label: const Text('Sắp chặn', style: TextStyle(fontSize: 12)),
                  selectedColor: Colors.red.shade100,
                  selected: _onlyUrgentBlock,
                  onSelected: (v) => _applyFilters(() => _onlyUrgentBlock = v),
                ),
                const SizedBox(width: 6),
                for (final method in [
                  'Có đánh dấu HT',
                  'Tất cả',
                  'TT',
                  'CK',
                  'TTT',
                  'IN HÓA ĐƠN',
                  'HD',
                  'Trống',
                ])
                  Padding(
                    padding: const EdgeInsets.only(right: 6),
                    child: ChoiceChip(
                      label: Text(method, style: const TextStyle(fontSize: 12)),
                      selected: _selectedMethod == method,
                      onSelected: (_) =>
                          _applyFilters(() => _selectedMethod = method),
                    ),
                  ),
              ],
            ),
          ),

          // 4. BỘ LỌC TRẠNG THÁI THU (CÓ TAB TÁC ĐỘNG HÔM NAY)
          SingleChildScrollView(
            scrollDirection: Axis.horizontal,
            padding: const EdgeInsets.symmetric(horizontal: 10, vertical: 2),
            child: Row(
              children: [
                _buildStatusFilterChip(
                  99,
                  'Hôm nay ($soTacDongHomNay)',
                  highlightColor: Colors.amber.shade200,
                ),
                _buildStatusFilterChip(-1, 'Tất cả (${_allData.length})'),
                _buildStatusFilterChip(0, 'Chưa thu'),
                _buildStatusFilterChip(1, 'Đã thu TM ($soDaThuTienMat)'),
                _buildStatusFilterChip(2, 'Đã gạch nợ cước'),
                _buildStatusFilterChip(3, 'Đã gạch chưa thu'),
                _buildStatusFilterChip(4, 'Hẹn lại'),
              ],
            ),
          ),
          const Divider(height: 4),

          // 5. DANH SÁCH THẺ KHÁCH HÀNG
          Expanded(
            child: _filteredData.isEmpty
                ? Center(
                    child: Padding(
                      padding: const EdgeInsets.all(20),
                      child: Text(
                        _allData.isEmpty
                            ? 'Chưa có dữ liệu.\nBấm biểu tượng tải lên ở góc phải để nạp thẳng file Excel (.xlsx) hoặc CSV.'
                            : 'Không có thuê bao nào khớp bộ lọc hiện tại.\n(Đang lọc: Hình thức = "$_selectedMethod").\nBấm chọn "Tất cả" nếu muốn xem toàn bộ thuê bao.',
                        textAlign: TextAlign.center,
                      ),
                    ),
                  )
                : ListView.builder(
                    itemCount: _filteredData.length,
                    addAutomaticKeepAlives: false,
                    itemBuilder: (ctx, idx) {
                      final item = _filteredData[idx];
                      final String maTb = (item['ma_tb'] ?? '').toString();
                      final int st = item['trang_thai'] ?? 0;
                      final String chan = (item['du_kien_chan'] ?? '')
                          .toString();
                      final String ht = (item['hinh_thuc'] ?? '').toString();
                      final String gc = (item['ghi_chu'] ?? '').toString();
                      final String ngay = (item['ngay_cap_nhat'] ?? '')
                          .toString();
                      final String gio = (item['gio_cap_nhat'] ?? '')
                          .toString();
                      final distanceMeters = item['_distanceMeters'] as double?;

                      return Card(
                        margin: const EdgeInsets.symmetric(
                          horizontal: 10,
                          vertical: 5,
                        ),
                        elevation: 0,
                        child: Padding(
                          padding: const EdgeInsets.all(10),
                          child: Column(
                            crossAxisAlignment: CrossAxisAlignment.start,
                            children: [
                              Row(
                                mainAxisAlignment:
                                    MainAxisAlignment.spaceBetween,
                                children: [
                                  Expanded(
                                    child: Text(
                                      item['ho_ten'] ?? '',
                                      style: const TextStyle(
                                        fontSize: 16,
                                        fontWeight: FontWeight.bold,
                                      ),
                                    ),
                                  ),
                                  Text(
                                    item['_formatted_ton_cuoc'] as String,
                                    style: const TextStyle(
                                      fontSize: 17,
                                      fontWeight: FontWeight.w900,
                                      color: Color(0xFFD32F2F),
                                    ),
                                  ),
                                ],
                              ),
                              const SizedBox(height: 6),

                              // NÚT SAO CHÉP NHANH MÃ ACCOUNT KHÁCH HÀNG
                              Tooltip(
                                message: 'Chạm để sao chép mã Account',
                                child: InkWell(
                                  onTap: () => _copyAccount(maTb),
                                  borderRadius: BorderRadius.circular(6),
                                  child: Container(
                                    padding: const EdgeInsets.symmetric(
                                      horizontal: 8,
                                      vertical: 4,
                                    ),
                                    decoration: BoxDecoration(
                                      color: Colors.red.shade50,
                                      borderRadius: BorderRadius.circular(6),
                                      border: Border.all(
                                        color: Colors.red.shade200,
                                      ),
                                    ),
                                    child: Row(
                                      children: [
                                        const Icon(
                                          Icons.copy,
                                          size: 15,
                                          color: Color(0xFFD32F2F),
                                        ),
                                        const SizedBox(width: 5),
                                        Expanded(
                                          child: Text(
                                            'Account: $maTb',
                                            maxLines: 1,
                                            overflow: TextOverflow.ellipsis,
                                            style: const TextStyle(
                                              fontSize: 13,
                                              fontWeight: FontWeight.bold,
                                              color: Color(0xFFD32F2F),
                                            ),
                                          ),
                                        ),
                                        if (ngay == today &&
                                            st != 0 &&
                                            gio.isNotEmpty)
                                          Text(
                                            'Hôm nay $gio',
                                            style: TextStyle(
                                              fontSize: 11.5,
                                              fontWeight: FontWeight.w600,
                                              color: Colors.green.shade800,
                                            ),
                                          ),
                                      ],
                                    ),
                                  ),
                                ),
                              ),

                              const SizedBox(height: 6),
                              Wrap(
                                spacing: 6,
                                runSpacing: 4,
                                children: [
                                  if (distanceMeters != null)
                                    _buildBadge(
                                      _formatDistance(distanceMeters),
                                      Colors.green.shade700,
                                    ),
                                  _buildBadge(
                                    item['khu_vuc'] ?? '',
                                    Colors.blueGrey,
                                  ),
                                  _buildBadge(
                                    item['dich_vu'] ?? '',
                                    Colors.indigo,
                                  ),
                                  if (ht.isNotEmpty)
                                    _buildBadge(
                                      'HT: $ht',
                                      Colors.teal.shade700,
                                    ),
                                  if (chan.isNotEmpty)
                                    _buildBadge(
                                      'Chặn ngày $chan',
                                      Colors.red.shade700,
                                    ),
                                  if ((item['sdt'] ?? '').toString().isNotEmpty)
                                    _buildBadge(
                                      'SĐT: ${item['sdt']}',
                                      Colors.grey.shade800,
                                    ),
                                ],
                              ),
                              const SizedBox(height: 6),
                              Text(
                                'Đ/c: ${item['dia_chi']}',
                                style: const TextStyle(
                                  fontSize: 13.5,
                                  color: Colors.black87,
                                ),
                              ),
                              if (gc.isNotEmpty)
                                Container(
                                  margin: const EdgeInsets.only(top: 4),
                                  padding: const EdgeInsets.symmetric(
                                    horizontal: 8,
                                    vertical: 4,
                                  ),
                                  decoration: BoxDecoration(
                                    color: Colors.amber.shade100,
                                    borderRadius: BorderRadius.circular(6),
                                  ),
                                  child: Row(
                                    children: [
                                      const Icon(
                                        Icons.sticky_note_2_outlined,
                                        size: 16,
                                        color: Colors.brown,
                                      ),
                                      const SizedBox(width: 4),
                                      Expanded(
                                        child: Text(
                                          'Ghi chú: $gc',
                                          style: const TextStyle(
                                            fontSize: 13,
                                            fontWeight: FontWeight.w600,
                                          ),
                                        ),
                                      ),
                                    ],
                                  ),
                                ),
                              const SizedBox(height: 8),
                              // HÀNG NÚT BẤM TRẠNG THÁI 1 CHẠM
                              Wrap(
                                spacing: 6,
                                runSpacing: 4,
                                children: [
                                  _buildActionBtn(
                                    maTb,
                                    1,
                                    st,
                                    'Tiền mặt',
                                    Colors.green,
                                  ),
                                  _buildActionBtn(
                                    maTb,
                                    2,
                                    st,
                                    'Đã gạch nợ cước',
                                    Colors.blue,
                                  ),
                                  _buildActionBtn(
                                    maTb,
                                    3,
                                    st,
                                    'Gạch chưa thu',
                                    Colors.orange.shade800,
                                  ),
                                  _buildActionBtn(
                                    maTb,
                                    4,
                                    st,
                                    'Hẹn lại',
                                    Colors.purple,
                                  ),
                                  IconButton(
                                    icon: Icon(
                                      item['latitude'] != null &&
                                              item['longitude'] != null
                                          ? Icons.location_on
                                          : Icons.location_on_outlined,
                                      color:
                                          item['latitude'] != null &&
                                              item['longitude'] != null
                                          ? Colors.green.shade700
                                          : Colors.black54,
                                    ),
                                    tooltip: 'Lưu vị trí GPS tại nhà khách',
                                    onPressed: () =>
                                        _saveCustomerLocation(item),
                                  ),
                                  IconButton(
                                    icon: const Icon(
                                      Icons.edit_note,
                                      color: Colors.black54,
                                    ),
                                    tooltip: 'Sửa Hình thức / Ghi chú',
                                    onPressed: () => _editNoteAndMethod(item),
                                  ),
                                ],
                              ),
                            ],
                          ),
                        ),
                      );
                    },
                  ),
          ),
        ],
      ),
    );
  }

  Widget _buildSummaryBox(String title, String value, Color color) {
    return Expanded(
      child: Container(
        padding: const EdgeInsets.all(8),
        decoration: BoxDecoration(
          color: Colors.white,
          borderRadius: BorderRadius.circular(8),
          border: Border.all(color: color.withValues(alpha: 0.4)),
        ),
        child: Column(
          children: [
            Text(
              title,
              style: const TextStyle(fontSize: 11, color: Colors.black54),
            ),
            const SizedBox(height: 2),
            Text(
              value,
              style: TextStyle(
                fontSize: 13.5,
                fontWeight: FontWeight.bold,
                color: color,
              ),
            ),
          ],
        ),
      ),
    );
  }

  Widget _buildBadge(String text, Color color) {
    return Container(
      padding: const EdgeInsets.symmetric(horizontal: 6, vertical: 2),
      decoration: BoxDecoration(
        color: color.withValues(alpha: 0.12),
        borderRadius: BorderRadius.circular(4),
      ),
      child: Text(
        text,
        style: TextStyle(
          fontSize: 11.5,
          fontWeight: FontWeight.w600,
          color: color,
        ),
      ),
    );
  }

  Widget _buildStatusFilterChip(
    int statusVal,
    String label, {
    Color? highlightColor,
  }) {
    return Padding(
      padding: const EdgeInsets.only(right: 6),
      child: ChoiceChip(
        label: Text(label, style: const TextStyle(fontSize: 12)),
        selectedColor: highlightColor,
        selected: _selectedStatus == statusVal,
        onSelected: (_) => _applyFilters(() {
          if (_nearbyMode && statusVal != 0) {
            _nearbyMode = false;
            _myPos = null;
          }
          _selectedStatus = statusVal;
        }),
      ),
    );
  }

  Widget _buildActionBtn(
    String maTb,
    int targetStatus,
    int currentStatus,
    String label,
    Color color,
  ) {
    final bool isSelected = currentStatus == targetStatus;
    return InkWell(
      onTap: () => _updateStatus(maTb, isSelected ? 0 : targetStatus),
      child: Container(
        padding: const EdgeInsets.symmetric(horizontal: 8, vertical: 6),
        decoration: BoxDecoration(
          color: isSelected ? color : Colors.grey.shade200,
          borderRadius: BorderRadius.circular(6),
        ),
        child: Text(
          label,
          style: TextStyle(
            fontSize: 11.5,
            fontWeight: FontWeight.bold,
            color: isSelected ? Colors.white : Colors.black87,
          ),
        ),
      ),
    );
  }
}
