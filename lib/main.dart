import 'dart:convert';
import 'dart:io';
import 'package:csv/csv.dart';
import 'package:file_picker/file_picker.dart';
import 'package:material_ui/material_ui.dart';
import 'package:intl/intl.dart';
import 'package:path/path.dart' as p;
import 'package:path_provider/path_provider.dart';
import 'package:share_plus/share_plus.dart';
import 'package:sqflite/sqflite.dart';

void main() {
  WidgetsFlutterBinding.ensureInitialized();
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

// ==================== DATABASE HELPER (SQLITE OFFLINE) ====================
class DbHelper {
  static Database? _db;

  static Future<Database> get db async {
    if (_db != null) return _db!;
    final dbPath = await getDatabasesPath();
    _db = await openDatabase(
      p.join(dbPath, 'thu_cuoc_lang_co.db'),
      version: 1,
      onCreate: (db, version) async {
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
            ton_cuoc INTEGER,
            du_kien_chan TEXT,
            trang_thai INTEGER DEFAULT 0,
            ngay_cap_nhat TEXT
          )
        ''');
        // trang_thai: 0=Chưa thu, 1=Đã thu TM, 2=Đã thu CK, 3=Đã gạch chưa thu, 4=Hẹn lại
      },
    );
    return _db!;
  }

  // Làm sạch số điện thoại từ file Excel/CSV
  static String cleanPhone(String raw) {
    String s = raw.trim();
    if (s.endsWith('.0')) s = s.substring(0, s.length - 2);
    s = s.replaceAll(RegExp(r'[^0-9]'), '');
    if (s.isEmpty || s == '0') return '';
    if (!s.startsWith('0')) s = '0$s';
    return s;
  }

  // Chuẩn hóa tiền cước (Xử lý cả dạng 190 ở T9 lẫn 190000 ở T10/T11)
  static int cleanAmount(String raw) {
    String s = raw.trim();
    if (s.endsWith('.0')) s = s.substring(0, s.length - 2);
    s = s.replaceAll(RegExp(r'[^0-9]'), '');
    if (s.isEmpty) return 0;
    int val = int.tryParse(s) ?? 0;
    if (val > 0 && val < 2000) val = val * 1000;
    return val;
  }

  // Gom nhóm Tổ thu dài thành 4 khu vực chính tại Lăng Cô / Phú Lộc
  static String normalizeArea(String raw) {
    String lower = raw.toLowerCase();
    if (lower.contains('an cư tân')) return 'An Cư Tân';
    if (lower.contains('lập an')) return 'Lập An';
    if (lower.contains('loan lý')) return 'Loan Lý';
    if (lower.contains('khu vực 1')) return 'Khu vực 1';
    return 'Khác';
  }

  // Đọc và gộp thông minh dữ liệu từ file CSV
  static Future<int> importCsvFile(File file) async {
    final content = await file.readAsString(encoding: utf8);
    final rows = const CsvToListConverter(eol: '\n', shouldParseNumbers: false)
        .convert(content.replaceAll('\r\n', '\n'));
    if (rows.length < 2) return 0;

    final headers = rows.first.map((e) => e.toString().trim()).toList();
    int idxMa = headers.indexWhere((h) => h.contains('Số thuê bao') || h.contains('Mã Account'));
    int idxTen = headers.indexWhere((h) => h.contains('Họ và tên'));
    int idxDiaChi = headers.indexWhere((h) => h.contains('Địa chỉ'));
    int idxSdt = headers.indexWhere((h) => h.contains('Số điện thoại'));
    int idxDichVu = headers.indexWhere((h) => h == 'Dịch vụ');
    int idxTonCuoc = headers.indexWhere((h) => h.contains('Tồn cước'));
    int idxToThu = headers.indexWhere((h) => h == 'Tổ thu');
    int idxChan = headers.indexWhere((h) => h.contains('Dự kiến chặn'));
    int idxHinhThuc = headers.indexWhere((h) => h.contains('Hinh Thuc'));
    int idxGhiChu = headers.indexWhere((h) => h.contains('Ghi Chu') || h.contains('Ghi chú'));
    int idxCot1 = headers.indexWhere((h) => h == '1');

    if (idxMa == -1) return 0;

    final database = await db;
    final existingRows = await database.query('khach_hang');
    final Map<String, Map<String, dynamic>> existingMap = {
      for (var r in existingRows) r['ma_tb'].toString(): r
    };

    int count = 0;
    final batch = database.batch();

    for (int i = 1; i < rows.length; i++) {
      final r = rows[i];
      if (r.length <= idxMa) continue;
      String maTb = r[idxMa].toString().trim();
      if (maTb.isEmpty) continue;

      final old = existingMap[maTb];

      String hoTen = idxTen != -1 && r.length > idxTen ? r[idxTen].toString().trim() : (old?['ho_ten'] ?? '');
      String diaChi = idxDiaChi != -1 && r.length > idxDiaChi ? r[idxDiaChi].toString().trim() : (old?['dia_chi'] ?? '');
      String sdtRaw = idxSdt != -1 && r.length > idxSdt ? r[idxSdt].toString() : '';
      String sdt = cleanPhone(sdtRaw);
      if (sdt.isEmpty && old != null) sdt = old['sdt'] ?? '';

      String dichVu = idxDichVu != -1 && r.length > idxDichVu ? r[idxDichVu].toString().trim() : (old?['dich_vu'] ?? '');
      String toThuGoc = idxToThu != -1 && r.length > idxToThu ? r[idxToThu].toString().trim() : (old?['to_thu_goc'] ?? '');
      String khuVuc = normalizeArea(toThuGoc);
      int tonCuoc = idxTonCuoc != -1 && r.length > idxTonCuoc ? cleanAmount(r[idxTonCuoc].toString()) : 0;

      String duKienChan = '';
      if (idxChan != -1 && r.length > idxChan) {
        duKienChan = r[idxChan].toString().trim().replaceAll('.0', '');
      }

      // Xử lý kế thừa Hình thức & Ghi chú (Bỏ qua #ERROR! của T11)
      String hinhThuc = old?['hinh_thuc'] ?? '';
      String ghiChu = old?['ghi_chu'] ?? '';

      if (idxHinhThuc != -1 && r.length > idxHinhThuc) {
        String htMoi = r[idxHinhThuc].toString().trim();
        if (htMoi.isNotEmpty && !htMoi.contains('ERROR')) {
          hinhThuc = htMoi.toUpperCase();
        }
      }

      if (idxGhiChu != -1 && r.length > idxGhiChu) {
        String gcMoi = r[idxGhiChu].toString().trim();
        if (gcMoi.isNotEmpty && !gcMoi.contains('ERROR')) {
          // Nếu là file T9 (Gộp chung Hình thức và Ghi chú vào 1 cột)
          if (idxHinhThuc == -1) {
            String upper = gcMoi.toUpperCase();
            if (['CK', 'TT', 'TTT', 'HD', 'IN HÓA ĐƠN', 'ACCOUNT MỚI'].contains(upper)) {
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

      // Kiểm tra cột '1' xem có ghi chú "đã gạch chưa thu" không
      int trangThai = 0;
      if (idxCot1 != -1 && r.length > idxCot1) {
        String c1 = r[idxCot1].toString().trim().toLowerCase();
        if (c1.contains('gạch')) trangThai = 3;
      }

      batch.insert(
        'khach_hang',
        {
          'ma_tb': maTb,
          'ho_ten': hoTen,
          'dia_chi': diaChi,
          'sdt': sdt,
          'dich_vu': dichVu,
          'to_thu_goc': toThuGoc,
          'khu_vuc': khuVuc,
          'hinh_thuc': hinhThuc,
          'ghi_chu': ghiChu,
          'ton_cuoc': tonCuoc,
          'du_kien_chan': duKienChan,
          'trang_thai': trangThai,
          'ngay_cap_nhat': '',
        },
        conflictAlgorithm: ConflictAlgorithm.replace,
      );
      count++;
    }

    await batch.commit(noResult: true);
    return count;
  }
}

// ==================== MÀN HÌNH CHÍNH ====================
class HomeScreen extends StatefulWidget {
  const HomeScreen({super.key});

  @override
  State<HomeScreen> createState() => _HomeScreenState();
}

class _HomeScreenState extends State<HomeScreen> {
  List<Map<String, dynamic>> _allData = [];
  List<Map<String, dynamic>> _filteredData = [];

  String _searchQuery = '';
  String _selectedArea = 'Tất cả';
  String _selectedMethod = 'Tất cả';
  int _selectedStatus = -1; // -1: Tất cả, 0: Chưa thu, 1: Đã thu TM, 2: Đã thu CK, 3: Đã gạch chưa thu, 4: Hẹn lại
  bool _onlyUrgentBlock = false;

  final currencyFmt = NumberFormat.currency(locale: 'vi_VN', symbol: 'đ', decimalDigits: 0);

  @override
  void initState() {
    super.initState();
    _loadData();
  }

  Future<void> _loadData() async {
    final database = await DbHelper.db;
    final data = await database.query('khach_hang', orderBy: 'ho_ten ASC');
    setState(() {
      _allData = data;
    });
    _applyFilters();
  }

  void _applyFilters() {
    final q = _searchQuery.toLowerCase().trim();
    setState(() {
      _filteredData = _allData.where((item) {
        if (_selectedArea != 'Tất cả' && item['khu_vuc'] != _selectedArea) return false;

        if (_selectedMethod != 'Tất cả') {
          String ht = (item['hinh_thuc'] ?? '').toString().toUpperCase();
          if (_selectedMethod == 'Trống' && ht.isNotEmpty) return false;
          if (_selectedMethod != 'Trống' && ht != _selectedMethod) return false;
        }

        if (_selectedStatus != -1 && item['trang_thai'] != _selectedStatus) return false;

        if (_onlyUrgentBlock) {
          String chan = (item['du_kien_chan'] ?? '').toString();
          if (chan.isEmpty) return false;
        }

        if (q.isNotEmpty) {
          String ten = (item['ho_ten'] ?? '').toString().toLowerCase();
          String dc = (item['dia_chi'] ?? '').toString().toLowerCase();
          String ma = (item['ma_tb'] ?? '').toString().toLowerCase();
          String sdt = (item['sdt'] ?? '').toString().toLowerCase();
          String gc = (item['ghi_chu'] ?? '').toString().toLowerCase();
          if (!ten.contains(q) && !dc.contains(q) && !ma.contains(q) && !sdt.contains(q) && !gc.contains(q)) {
            return false;
          }
        }
        return true;
      }).toList();
    });
  }

  Future<void> _updateStatus(String maTb, int newStatus) async {
    final database = await DbHelper.db;
    final today = DateFormat('yyyy-MM-dd').format(DateTime.now());
    await database.update(
      'khach_hang',
      {'trang_thai': newStatus, 'ngay_cap_nhat': newStatus == 0 ? '' : today},
      where: 'ma_tb = ?',
      whereArgs: [maTb],
    );
    await _loadData();
  }

  Future<void> _editNoteAndMethod(Map<String, dynamic> item) async {
    String selectedHt = item['hinh_thuc'] ?? '';
    final noteCtrl = TextEditingController(text: item['ghi_chu'] ?? '');

    await showDialog(
      context: context,
      builder: (ctx) => StatefulBuilder(
        builder: (ctx, setDialogState) => AlertDialog(
          title: Text(item['ho_ten'] ?? item['ma_tb'], style: const TextStyle(fontSize: 18, fontWeight: FontWeight.bold)),
          content: Column(
            mainAxisSize: MainAxisSize.min,
            crossAxisAlignment: CrossAxisAlignment.start,
            children: [
              const Text('Hình thức thu cố định:', style: TextStyle(fontWeight: FontWeight.w600)),
              const SizedBox(height: 8),
              Wrap(
                spacing: 8,
                children: ['TT', 'CK', 'TTT', 'IN HÓA ĐƠN', 'HD'].map((ht) {
                  return ChoiceChip(
                    label: Text(ht),
                    selected: selectedHt == ht,
                    onSelected: (val) => setDialogState(() => selectedHt = val ? ht : ''),
                  );
                }).toList(),
              ),
              const SizedBox(height: 16),
              TextField(
                controller: noteCtrl,
                decoration: const InputDecoration(
                  labelText: 'Ghi chú đường đi / Nhắc nhở',
                  hintText: 'VD: Gần Thu Phượng, Ghé a râu lấy...',
                  border: OutlineInputBorder(),
                ),
                maxLines: 2,
              ),
            ],
          ),
          actions: [
            TextButton(onPressed: () => Navigator.pop(ctx), child: const Text('Hủy')),
            FilledButton(
              onPressed: () async {
                final database = await DbHelper.db;
                await database.update(
                  'khach_hang',
                  {'hinh_thuc': selectedHt, 'ghi_chu': noteCtrl.text.trim()},
                  where: 'ma_tb = ?',
                  whereArgs: [item['ma_tb']],
                );
                if (ctx.mounted) Navigator.pop(ctx);
                await _loadData();
              },
              child: const Text('Lưu lại'),
            ),
          ],
        ),
      ),
    );
  }

  Future<void> _pickAndImportCsv() async {
    final result = await FilePicker.platform.pickFiles(
      type: FileType.custom,
      allowedExtensions: ['csv'],
    );
    if (result != null && result.files.single.path != null) {
      final file = File(result.files.single.path!);
      final count = await DbHelper.importCsvFile(file);
      await _loadData();
      if (mounted) {
        ScaffoldMessenger.of(context).showSnackBar(
          SnackBar(content: Text('Đã nạp và đồng bộ thành công $count thuê bao!')),
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
        'Dự kiến chặn'
      ]
    ];

    String statusText(int st) {
      switch (st) {
        case 1:
          return 'Đã thu TM';
        case 2:
          return 'Đã CK';
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
      ]);
    }

    String csvData = const ListToCsvConverter().convert(rows);
    final dir = await getTemporaryDirectory();
    final path = p.join(dir.path, 'Ket_Qua_Thu_Cuoc.csv');
    final file = File(path);
    await file.writeAsString('\uFEFF$csvData', encoding: utf8);
    await Share.shareXFiles([XFile(path)], text: 'Báo cáo kết quả thu cước');
  }

  @override
  Widget build(BuildContext context) {
    // Tính toán tiền quỹ đối soát
    int tongTienMat = 0;
    int tongChuyenKhoan = 0;
    int tongDaGachChuaThu = 0;
    int soDaThu = 0;

    for (var r in _allData) {
      int st = r['trang_thai'] ?? 0;
      int tien = r['ton_cuoc'] ?? 0;
      if (st == 1) {
        tongTienMat += tien;
        soDaThu++;
      } else if (st == 2) {
        tongChuyenKhoan += tien;
        soDaThu++;
      } else if (st == 3) {
        tongDaGachChuaThu += tien;
      }
    }

    return Scaffold(
      appBar: AppBar(
        title: Text('Thu Cước (${_filteredData.length}/${_allData.length})'),
        backgroundColor: const Color(0xFFD32F2F),
        foregroundColor: Colors.white,
        actions: [
          IconButton(
            icon: const Icon(Icons.upload_file),
            tooltip: 'Nạp file CSV',
            onPressed: _pickAndImportCsv,
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
                _buildSummaryBox('Tiền mặt cầm', currencyFmt.format(tongTienMat), Colors.green.shade700),
                const SizedBox(width: 6),
                _buildSummaryBox('Khách đã CK', currencyFmt.format(tongChuyenKhoan), Colors.blue.shade700),
                const SizedBox(width: 6),
                _buildSummaryBox('Gạch chưa thu', currencyFmt.format(tongDaGachChuaThu), Colors.orange.shade800),
              ],
            ),
          ),

          // 2. THANH TÌM KIẾM NHANH
          Padding(
            padding: const EdgeInsets.fromLTRB(10, 6, 10, 4),
            child: TextField(
              decoration: InputDecoration(
                hintText: 'Tìm tên, số nhà (389), kiệt, ghi chú (a râu, B&H)...',
                prefixIcon: const Icon(Icons.search),
                suffixIcon: _searchQuery.isNotEmpty
                    ? IconButton(
                        icon: const Icon(Icons.clear),
                        onPressed: () {
                          setState(() => _searchQuery = '');
                          _applyFilters();
                        },
                      )
                    : null,
                isDense: true,
                contentPadding: const EdgeInsets.symmetric(vertical: 10),
                border: OutlineInputBorder(borderRadius: BorderRadius.circular(10)),
              ),
              onChanged: (val) {
                _searchQuery = val;
                _applyFilters();
              },
            ),
          ),

          // 3. BỘ LỌC KHU VỰC & HÌNH THỨC
          SingleChildScrollView(
            scrollDirection: Axis.horizontal,
            padding: const EdgeInsets.symmetric(horizontal: 10),
            child: Row(
              children: [
                for (final area in ['Tất cả', 'An Cư Tân', 'Lập An', 'Loan Lý', 'Khu vực 1'])
                  Padding(
                    padding: const EdgeInsets.only(right: 6),
                    child: ChoiceChip(
                      label: Text(area, style: const TextStyle(fontSize: 12)),
                      selected: _selectedArea == area,
                      onSelected: (_) {
                        setState(() => _selectedArea = area);
                        _applyFilters();
                      },
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
                  onSelected: (v) {
                    setState(() => _onlyUrgentBlock = v);
                    _applyFilters();
                  },
                ),
                const SizedBox(width: 6),
                for (final method in ['Tất cả', 'TT', 'CK', 'TTT', 'IN HÓA ĐƠN', 'Trống'])
                  Padding(
                    padding: const EdgeInsets.only(right: 6),
                    child: ChoiceChip(
                      label: Text(method, style: const TextStyle(fontSize: 12)),
                      selected: _selectedMethod == method,
                      onSelected: (_) {
                        setState(() => _selectedMethod = method);
                        _applyFilters();
                      },
                    ),
                  ),
              ],
            ),
          ),

          // 4. BỘ LỌC TRẠNG THÁI THU
          SingleChildScrollView(
            scrollDirection: Axis.horizontal,
            padding: const EdgeInsets.symmetric(horizontal: 10, vertical: 2),
            child: Row(
              children: [
                _buildStatusFilterChip(-1, 'Tất cả (${_allData.length})'),
                _buildStatusFilterChip(0, 'Chưa thu'),
                _buildStatusFilterChip(3, 'Đã gạch chưa thu'),
                _buildStatusFilterChip(4, 'Hẹn lại'),
                _buildStatusFilterChip(1, 'Đã thu TM ($soDaThu)'),
                _buildStatusFilterChip(2, 'Đã CK'),
              ],
            ),
          ),
          const Divider(height: 4),

          // 5. DANH SÁCH THẺ KHÁCH HÀNG
          Expanded(
            child: _filteredData.isEmpty
                ? const Center(
                    child: Text(
                      'Chưa có dữ liệu hoặc không tìm thấy.\nBấm biểu tượng tải lên ở góc phải để nạp file CSV.',
                      textAlign: TextAlign.center,
                    ),
                  )
                : ListView.builder(
                    itemCount: _filteredData.length,
                    itemBuilder: (ctx, idx) {
                      final item = _filteredData[idx];
                      final int st = item['trang_thai'] ?? 0;
                      final String chan = (item['du_kien_chan'] ?? '').toString();
                      final String ht = (item['hinh_thuc'] ?? '').toString();
                      final String gc = (item['ghi_chu'] ?? '').toString();

                      return Card(
                        margin: const EdgeInsets.symmetric(horizontal: 10, vertical: 5),
                        elevation: 2,
                        child: Padding(
                          padding: const EdgeInsets.all(10),
                          child: Column(
                            crossAxisAlignment: CrossAxisAlignment.start,
                            children: [
                              Row(
                                mainAxisAlignment: MainAxisAlignment.spaceBetween,
                                children: [
                                  Expanded(
                                    child: Text(
                                      item['ho_ten'] ?? '',
                                      style: const TextStyle(fontSize: 16, fontWeight: FontWeight.bold),
                                    ),
                                  ),
                                  Text(
                                    currencyFmt.format(item['ton_cuoc'] ?? 0),
                                    style: const TextStyle(
                                      fontSize: 17,
                                      fontWeight: FontWeight.w900,
                                      color: Color(0xFFD32F2F),
                                    ),
                                  ),
                                ],
                              ),
                              const SizedBox(height: 4),
                              Wrap(
                                spacing: 6,
                                runSpacing: 4,
                                children: [
                                  _buildBadge(item['khu_vuc'] ?? '', Colors.blueGrey),
                                  _buildBadge(item['dich_vu'] ?? '', Colors.indigo),
                                  if (ht.isNotEmpty) _buildBadge('HT: $ht', Colors.teal.shade700),
                                  if (chan.isNotEmpty) _buildBadge('Chặn ngày $chan', Colors.red.shade700),
                                  if ((item['sdt'] ?? '').toString().isNotEmpty)
                                    _buildBadge('SĐT: ${item['sdt']}', Colors.grey.shade800),
                                ],
                              ),
                              const SizedBox(height: 6),
                              Text(
                                'Đ/c: ${item['dia_chi']}',
                                style: const TextStyle(fontSize: 13.5, color: Colors.black87),
                              ),
                              if (gc.isNotEmpty)
                                Container(
                                  margin: const EdgeInsets.only(top: 4),
                                  padding: const EdgeInsets.symmetric(horizontal: 8, vertical: 4),
                                  decoration: BoxDecoration(
                                    color: Colors.amber.shade100,
                                    borderRadius: BorderRadius.circular(6),
                                  ),
                                  child: Row(
                                    children: [
                                      const Icon(Icons.sticky_note_2_outlined, size: 16, color: Colors.brown),
                                      const SizedBox(width: 4),
                                      Expanded(
                                        child: Text(
                                          'Ghi chú: $gc',
                                          style: const TextStyle(fontSize: 13, fontWeight: FontWeight.w600),
                                        ),
                                      ),
                                    ],
                                  ),
                                ),
                              const SizedBox(height: 8),
                              // HÀNG NÚT BẤM TRẠNG THÁI 1 CHẠM
                              Row(
                                mainAxisAlignment: MainAxisAlignment.spaceBetween,
                                children: [
                                  _buildActionBtn(item['ma_tb'], 1, st, 'Tiền mặt', Colors.green),
                                  _buildActionBtn(item['ma_tb'], 2, st, 'Đã CK', Colors.blue),
                                  _buildActionBtn(item['ma_tb'], 3, st, 'Gạch nợ', Colors.orange.shade800),
                                  _buildActionBtn(item['ma_tb'], 4, st, 'Hẹn lại', Colors.purple),
                                  IconButton(
                                    icon: const Icon(Icons.edit_note, color: Colors.black54),
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
            Text(title, style: const TextStyle(fontSize: 11, color: Colors.black54)),
            const SizedBox(height: 2),
            Text(value, style: TextStyle(fontSize: 13.5, fontWeight: FontWeight.bold, color: color)),
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
      child: Text(text, style: TextStyle(fontSize: 11.5, fontWeight: FontWeight.w600, color: color)),
    );
  }

  Widget _buildStatusFilterChip(int statusVal, String label) {
    return Padding(
      padding: const EdgeInsets.only(right: 6),
      child: ChoiceChip(
        label: Text(label, style: const TextStyle(fontSize: 12)),
        selected: _selectedStatus == statusVal,
        onSelected: (_) {
          setState(() => _selectedStatus = statusVal);
          _applyFilters();
        },
      ),
    );
  }

  Widget _buildActionBtn(String maTb, int targetStatus, int currentStatus, String label, Color color) {
    final bool isSelected = currentStatus == targetStatus;
    return InkWell(
      onTap: () => _updateStatus(maTb, isSelected ? 0 : targetStatus),
      child: Container(
        padding: const EdgeInsets.symmetric(horizontal: 10, vertical: 6),
        decoration: BoxDecoration(
          color: isSelected ? color : Colors.grey.shade200,
          borderRadius: BorderRadius.circular(6),
        ),
        child: Text(
          label,
          style: TextStyle(
            fontSize: 12.5,
            fontWeight: FontWeight.bold,
            color: isSelected ? Colors.white : Colors.black87,
          ),
        ),
      ),
    );
  }
}