import 'dart:async';
import 'dart:convert';
import 'dart:io';

import 'package:excel/excel.dart' hide Border;
import 'package:file_picker/file_picker.dart';
import 'package:flutter/material.dart';
import 'package:flutter/services.dart';
import 'package:path_provider/path_provider.dart';
import 'package:share_plus/share_plus.dart';

void main() {
  runApp(const MyApp());
}

class MyApp extends StatelessWidget {
  const MyApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      debugShowCheckedModeBanner: false,
      title: 'Inventory Audit Pro',
      theme: ThemeData(
        colorScheme: ColorScheme.fromSeed(
          seedColor: Colors.indigo,
          brightness: Brightness.light,
        ),
        useMaterial3: true,
        inputDecorationTheme: InputDecorationTheme(
          border: OutlineInputBorder(borderRadius: BorderRadius.circular(8)),
          contentPadding: const EdgeInsets.symmetric(horizontal: 12, vertical: 8),
          isDense: true,
        ),
      ),
      home: const InventoryAuditApp(),
    );
  }
}

// ---------------------------------------------------------------------------
// OPTIMIZED DATA MODEL WITH REMARKS
// ---------------------------------------------------------------------------

class MaterialItem {
  final int rowIndex;
  final String code;
  final String description;
  final String uom;
  final double systemQuantity;
  final double rateValue;
  
  String physicalQty;
  String remarks;
  
  final FocusNode focusNode = FocusNode();
  final FocusNode remarksFocusNode = FocusNode();

  MaterialItem({
    required this.rowIndex,
    required this.code,
    required this.description,
    required String systemQtyStr,
    required this.uom,
    required String rateStr,
    this.physicalQty = '',
    this.remarks = '',
  }) : 
    systemQuantity = double.tryParse(systemQtyStr.replaceAll(',', '')) ?? 0,
    rateValue = double.tryParse(rateStr.replaceAll(',', '')) ?? 0;

  double get physicalQuantity => double.tryParse(physicalQty) ?? 0;
  double get variance => physicalQuantity - systemQuantity;
  double get varianceValue => variance * rateValue;
  bool get isVerified => physicalQty.isNotEmpty;

  Color get statusColor {
    if (!isVerified) return Colors.transparent;
    if (variance == 0) return Colors.green.withOpacity(0.05); 
    if (variance > 0) return Colors.blue.withOpacity(0.1);    
    return Colors.red.withOpacity(0.1);
  }

  Map<String, dynamic> toJson() => {
    'rowIndex': rowIndex,
    'physicalQty': physicalQty,
    'remarks': remarks,
  };

  void updateFromJson(Map<String, dynamic> json) {
    physicalQty = json['physicalQty'] ?? '';
    remarks = json['remarks'] ?? '';
  }

  @override
  void dispose() {
    focusNode.dispose();
    remarksFocusNode.dispose();
  }
}

// ---------------------------------------------------------------------------
// MAIN SCREEN WITH AUTO-SAVE
// ---------------------------------------------------------------------------

class InventoryAuditApp extends StatefulWidget {
  const InventoryAuditApp({super.key});

  @override
  State<InventoryAuditApp> createState() => _InventoryAuditAppState();
}

class _InventoryAuditAppState extends State<InventoryAuditApp> with WidgetsBindingObserver {
  List<List<TextEditingController>> _data = [];
  String _fileName = '';
  String _fileHash = '';
  
  final Map<int, MaterialItem> _materialMap = {};
  List<MaterialItem> _filteredMaterials = [];

  int? _codeColIndex;
  int? _descColIndex;
  int? _qtyColIndex;
  int? _uomColIndex;
  int? _rateColIndex;
  int? _physicalQtyColIndex;
  int? _remarksColIndex;

  bool _isLoading = false;
  bool _isVerificationMode = false;
  bool _hasUnsavedChanges = false;
  
  final ScrollController _headerScrollController = ScrollController();
  final ScrollController _bodyScrollController = ScrollController();
  final TextEditingController _searchController = TextEditingController();
  Timer? _debounce;
  Timer? _autoSaveTimer;

  final Map<int, TextEditingController> _verificationControllers = {};
  final Map<int, TextEditingController> _remarksControllers = {};

  @override
  void initState() {
    super.initState();
    WidgetsBinding.instance.addObserver(this);
    _syncScrollControllers();
    _loadSavedSession();
  }

  @override
  void dispose() {
    WidgetsBinding.instance.removeObserver(this);
    _searchController.dispose();
    _headerScrollController.dispose();
    _bodyScrollController.dispose();
    _debounce?.cancel();
    _autoSaveTimer?.cancel();
    _disposeDataControllers();
    super.dispose();
  }

  @override
  void didChangeAppLifecycleState(AppLifecycleState state) {
    if (state == AppLifecycleState.paused || state == AppLifecycleState.inactive) {
      _saveSession();
    }
  }

  void _syncScrollControllers() {
    _headerScrollController.addListener(() {
      if (_bodyScrollController.hasClients && 
          _headerScrollController.offset != _bodyScrollController.offset) {
        _bodyScrollController.jumpTo(_headerScrollController.offset);
      }
    });
    _bodyScrollController.addListener(() {
      if (_headerScrollController.hasClients && 
          _bodyScrollController.offset != _headerScrollController.offset) {
        _headerScrollController.jumpTo(_bodyScrollController.offset);
      }
    });
  }

  void _disposeDataControllers() {
    for (var row in _data) {
      for (var c in row) c.dispose();
    }
    for (var c in _verificationControllers.values) c.dispose();
    for (var c in _remarksControllers.values) c.dispose();
    for (var item in _materialMap.values) {
      item.focusNode.dispose();
      item.remarksFocusNode.dispose();
    }
  }

  // -------------------------------------------------------------------------
  // AUTO-SAVE & SESSION MANAGEMENT (FILE-BASED)
  // -------------------------------------------------------------------------

  String _generateFileHash(String fileName) {
    return fileName.hashCode.toString();
  }

  Future<File> _getSessionFile() async {
    final dir = await getApplicationDocumentsDirectory();
    return File('${dir.path}/audit_session.json');
  }

  Future<void> _saveSession() async {
    if (_materialMap.isEmpty || _fileName.isEmpty) return;
    
    try {
      final sessionFile = await _getSessionFile();
      
      final sessionData = {
        'fileName': _fileName,
        'fileHash': _fileHash,
        'mappings': {
          'code': _codeColIndex,
          'desc': _descColIndex,
          'qty': _qtyColIndex,
          'uom': _uomColIndex,
          'rate': _rateColIndex,
          'physical': _physicalQtyColIndex,
          'remarks': _remarksColIndex,
        },
        'materials': _materialMap.map((key, item) => MapEntry(key.toString(), item.toJson())),
        'timestamp': DateTime.now().toIso8601String(),
      };
      
      await sessionFile.writeAsString(jsonEncode(sessionData));
      setState(() => _hasUnsavedChanges = false);
    } catch (e) {
      debugPrint('Save session error: $e');
    }
  }

  Future<void> _loadSavedSession() async {
    try {
      final sessionFile = await _getSessionFile();
      
      if (!await sessionFile.exists()) return;
      
      final sessionStr = await sessionFile.readAsString();
      final session = jsonDecode(sessionStr);
      final savedTime = DateTime.parse(session['timestamp']);
      final hoursSince = DateTime.now().difference(savedTime).inHours;
      
      if (hoursSince > 48) {
        await sessionFile.delete();
        return;
      }
      
      if (!mounted) return;
      
      final shouldRestore = await showDialog<bool>(
        context: context,
        builder: (ctx) => AlertDialog(
          title: const Text('Restore Session?'),
          content: Text(
            'Found unsaved work from ${_formatTimeDiff(savedTime)}\n\n'
            'File: ${session['fileName']}\n\n'
            'Would you like to restore it?'
          ),
          actions: [
            TextButton(
              onPressed: () => Navigator.pop(ctx, false),
              child: const Text('Discard'),
            ),
            ElevatedButton(
              onPressed: () => Navigator.pop(ctx, true),
              child: const Text('Restore'),
            ),
          ],
        ),
      );
      
      if (shouldRestore == true) {
        _showSnackBar('Session restored. Re-import the Excel file to continue.', Colors.green);
        setState(() {
          _fileHash = session['fileHash'];
        });
      } else {
        await sessionFile.delete();
      }
    } catch (e) {
      debugPrint('Load session error: $e');
    }
  }

  String _formatTimeDiff(DateTime time) {
    final diff = DateTime.now().difference(time);
    if (diff.inMinutes < 60) return '${diff.inMinutes} minutes ago';
    if (diff.inHours < 24) return '${diff.inHours} hours ago';
    return '${diff.inDays} days ago';
  }

  void _scheduleAutoSave() {
    _autoSaveTimer?.cancel();
    _autoSaveTimer = Timer(const Duration(seconds: 5), () {
      _saveSession();
    });
  }

  Future<void> _clearSession() async {
    final sessionFile = await _getSessionFile();
    if (await sessionFile.exists()) {
      await sessionFile.delete();
    }
    setState(() => _hasUnsavedChanges = false);
  }

  // -------------------------------------------------------------------------
  // FILE HANDLING - OPTIMIZED
  // -------------------------------------------------------------------------

  Future<void> _importExcel() async {
    try {
      FilePickerResult? result = await FilePicker.platform.pickFiles(
        type: FileType.custom,
        allowedExtensions: ['xlsx', 'xls'],
      );
      
      if (result == null) return;

      setState(() => _isLoading = true);
      await Future.delayed(const Duration(milliseconds: 100));

      final File file = File(result.files.single.path!);
      final Uint8List bytes = await file.readAsBytes();
      final Excel excel = Excel.decodeBytes(bytes);

      _disposeDataControllers();
      
      List<List<TextEditingController>> newData = [];
      String foundSheetName = '';

      for (var table in excel.tables.keys) {
        var sheet = excel.tables[table]!;
        if (sheet.maxRows == 0) continue;
        
        foundSheetName = table;
        
        for (int rowIndex = 0; rowIndex < sheet.maxRows; rowIndex++) {
          List<TextEditingController> rowControllers = [];
          var sheetRow = sheet.row(rowIndex);
          
          for (int colIndex = 0; colIndex < sheet.maxColumns; colIndex++) {
            String cellValue = '';
            if (colIndex < sheetRow.length) {
              final val = sheetRow[colIndex]?.value;
              if (val != null) cellValue = val.toString();
            }
            rowControllers.add(TextEditingController(text: cellValue));
          }
          newData.add(rowControllers);
        }
        break; 
      }

      if (newData.isEmpty) throw Exception("No data found.");

      final newHash = _generateFileHash(result.files.single.name);
      final bool restoreSession = newHash == _fileHash && _fileHash.isNotEmpty;

      setState(() {
        _data = newData;
        _fileName = result.files.single.name;
        _fileHash = newHash;
        _isLoading = false;
        _isVerificationMode = false;
        _searchController.clear();
        if (!restoreSession) {
          _materialMap.clear();
          _filteredMaterials.clear();
        }
        _codeColIndex = null;
        _qtyColIndex = null;
        _physicalQtyColIndex = null;
        _remarksColIndex = null;
      });

      _showSnackBar("Loaded $foundSheetName: ${newData.length} rows", Colors.green);
      
      if (restoreSession) {
        await _restoreSessionData();
      } else {
        Future.delayed(const Duration(milliseconds: 500), _showColumnMappingDialog);
      }

    } catch (e) {
      setState(() => _isLoading = false);
      _showSnackBar("Import Error: ${e.toString()}", Colors.red);
    }
  }

  Future<void> _restoreSessionData() async {
    try {
      final sessionFile = await _getSessionFile();
      if (!await sessionFile.exists()) return;
      
      final sessionStr = await sessionFile.readAsString();
      final session = jsonDecode(sessionStr);
      final mappings = session['mappings'];
      
      setState(() {
        _codeColIndex = mappings['code'];
        _descColIndex = mappings['desc'];
        _qtyColIndex = mappings['qty'];
        _uomColIndex = mappings['uom'];
        _rateColIndex = mappings['rate'];
        _physicalQtyColIndex = mappings['physical'];
        _remarksColIndex = mappings['remarks'];
      });
      
      _buildMaterialIndex();
      
      final materials = session['materials'] as Map<String, dynamic>;
      for (var entry in materials.entries) {
        final rowIndex = int.parse(entry.key);
        if (_materialMap.containsKey(rowIndex)) {
          _materialMap[rowIndex]!.updateFromJson(entry.value);
        }
      }
      
      _applyVerificationDataToSheet();
      
      setState(() => _isVerificationMode = true);
      _showSnackBar('Session restored successfully', Colors.green);
      
    } catch (e) {
      _showSnackBar('Failed to restore session data', Colors.orange);
    }
  }

  void _applyVerificationDataToSheet() {
    for (var item in _materialMap.values) {
      if (_physicalQtyColIndex != null && _physicalQtyColIndex! < _data[item.rowIndex].length) {
        _data[item.rowIndex][_physicalQtyColIndex!].text = item.physicalQty;
      }
      if (_remarksColIndex != null && _remarksColIndex! < _data[item.rowIndex].length) {
        _data[item.rowIndex][_remarksColIndex!].text = item.remarks;
      }
    }
  }

  Future<void> _exportExcel() async {
    if (_data.isEmpty) return;
    setState(() => _isLoading = true);
    
    await Future.delayed(const Duration(milliseconds: 100));

    try {
      var excel = Excel.createExcel();
      Sheet sheet = excel['Audit Report'];
      
      sheet.appendRow([TextCellValue('INVENTORY AUDIT REPORT')]);
      sheet.appendRow([TextCellValue('File: $_fileName')]);
      sheet.appendRow([TextCellValue('Date: ${DateTime.now().toString().split('.')[0]}')]);
      sheet.appendRow([TextCellValue('')]); 

      for (int i = 0; i < _data.length; i++) {
        List<CellValue> rowValues = [];
        for (var controller in _data[i]) {
          String text = controller.text;
          double? num = double.tryParse(text);
          if (num != null) {
            rowValues.add(DoubleCellValue(num));
          } else {
            rowValues.add(TextCellValue(text));
          }
        }
        sheet.appendRow(rowValues);
      }

      if (_materialMap.isNotEmpty) {
        sheet.appendRow([TextCellValue('')]);
        sheet.appendRow([TextCellValue('SUMMARY STATISTICS')]);
        
        int verified = _materialMap.values.where((i) => i.isVerified).length;
        int shortages = _materialMap.values.where((i) => i.variance < 0).length;
        int excesses = _materialMap.values.where((i) => i.variance > 0).length;
        double totalShortageValue = _materialMap.values
          .where((i) => i.variance < 0)
          .fold(0, (sum, i) => sum + i.varianceValue.abs());
        double totalExcessValue = _materialMap.values
          .where((i) => i.variance > 0)
          .fold(0, (sum, i) => sum + i.varianceValue);

        sheet.appendRow([TextCellValue('Total Items'), DoubleCellValue(_materialMap.length.toDouble())]);
        sheet.appendRow([TextCellValue('Verified Items'), DoubleCellValue(verified.toDouble())]);
        sheet.appendRow([TextCellValue('Shortage Count'), DoubleCellValue(shortages.toDouble())]);
        sheet.appendRow([TextCellValue('Excess Count'), DoubleCellValue(excesses.toDouble())]);
        sheet.appendRow([TextCellValue('Total Shortage Value'), DoubleCellValue(totalShortageValue)]);
        sheet.appendRow([TextCellValue('Total Excess Value'), DoubleCellValue(totalExcessValue)]);
      }

      final dir = await getApplicationDocumentsDirectory();
      final String path = '${dir.path}/Audit_${DateTime.now().millisecondsSinceEpoch}.xlsx';
      final File file = File(path);
      await file.writeAsBytes(excel.encode()!);
      
      await Share.shareXFiles([XFile(path)], text: 'Inventory Audit Report');
      setState(() => _isLoading = false);
      
      await _clearSession();
      _showSnackBar('Exported successfully', Colors.green);
      
    } catch (e) {
      setState(() => _isLoading = false);
      _showSnackBar("Export Failed: $e", Colors.red);
    }
  }

  // -------------------------------------------------------------------------
  // COLUMN MAPPING - WITH REMARKS
  // -------------------------------------------------------------------------

  void _showColumnMappingDialog() {
    if (_data.length < 2) return;

    List<String> headers = _data[0].map((c) => c.text.trim()).toList();
    List<String> lowerHeaders = headers.map((h) => h.toLowerCase()).toList();

    int? detect(List<String> keywords) {
      int idx = lowerHeaders.indexWhere((h) => keywords.any((k) => h.contains(k)));
      return idx != -1 ? idx : null;
    }

    int? tempCode = detect(['code', 'sap', 'material', 'item', 'sku']);
    int? tempDesc = detect(['desc', 'name', 'detail']);
    int? tempQty = detect(['sys', 'book', 'current', 'sap qty', 'qty']);
    int? tempUom = detect(['uom', 'unit', 'base']);
    int? tempRate = detect(['rate', 'price', 'cost', 'val']);
    int? tempPhy = detect(['phy', 'act', 'count', 'audit']);
    int? tempRemarks = detect(['remark', 'note', 'comment', 'obs']);

    showDialog(
      context: context,
      barrierDismissible: false,
      builder: (context) => StatefulBuilder(
        builder: (ctx, setDialogState) => AlertDialog(
          title: const Text("Map Columns"),
          content: SingleChildScrollView(
            child: Column(
              mainAxisSize: MainAxisSize.min,
              children: [
                _mappingDropdown("Material Code *", headers, tempCode, (v) => tempCode = v, setDialogState),
                _mappingDropdown("Description", headers, tempDesc, (v) => tempDesc = v, setDialogState),
                _mappingDropdown("System Qty *", headers, tempQty, (v) => tempQty = v, setDialogState),
                _mappingDropdown("UOM", headers, tempUom, (v) => tempUom = v, setDialogState),
                _mappingDropdown("Rate/Price", headers, tempRate, (v) => tempRate = v, setDialogState),
                const Divider(),
                _mappingDropdown("Physical Qty (Output)", headers, tempPhy, (v) => tempPhy = v, setDialogState),
                _mappingDropdown("Remarks (Output)", headers, tempRemarks, (v) => tempRemarks = v, setDialogState),
                const Padding(
                  padding: EdgeInsets.only(top: 8.0),
                  child: Text("* Required fields", style: TextStyle(fontSize: 12, color: Colors.grey)),
                ),
              ],
            ),
          ),
          actions: [
            TextButton(
              onPressed: () => Navigator.pop(context),
              child: const Text("Cancel"),
            ),
            ElevatedButton(
              onPressed: () {
                if (tempCode == null || tempQty == null) {
                  _showSnackBar("Code and System Qty are required.", Colors.orange);
                  return;
                }
                
                setState(() {
                  _codeColIndex = tempCode;
                  _descColIndex = tempDesc;
                  _qtyColIndex = tempQty;
                  _uomColIndex = tempUom;
                  _rateColIndex = tempRate;
                  
                  if (tempPhy == null) {
                    _addColumn("Physical Qty");
                    _physicalQtyColIndex = _data[0].length - 1;
                  } else {
                    _physicalQtyColIndex = tempPhy;
                  }
                  
                  if (tempRemarks == null) {
                    _addColumn("Remarks");
                    _remarksColIndex = _data[0].length - 1;
                  } else {
                    _remarksColIndex = tempRemarks;
                  }
                });

                _buildMaterialIndex();
                Navigator.pop(context);
                setState(() => _isVerificationMode = true);
              },
              child: const Text("Start Verification"),
            ),
          ],
        ),
      ),
    );
  }

  Widget _mappingDropdown(String label, List<String> options, int? value, Function(int?) onChange, StateSetter setStateParams) {
    return Padding(
      padding: const EdgeInsets.symmetric(vertical: 6),
      child: Row(
        children: [
          Expanded(flex: 2, child: Text(label, style: const TextStyle(fontSize: 13))),
          Expanded(
            flex: 3,
            child: DropdownButtonFormField<int>(
              isExpanded: true,
              value: value,
              items: [
                const DropdownMenuItem(value: null, child: Text("- Ignore -", style: TextStyle(color: Colors.grey))),
                ...List.generate(options.length, (index) => DropdownMenuItem(
                  value: index,
                  child: Text(options[index], overflow: TextOverflow.ellipsis, style: const TextStyle(fontSize: 13)),
                )),
              ],
              onChanged: (val) => setStateParams(() => onChange(val)),
              decoration: const InputDecoration(contentPadding: EdgeInsets.symmetric(horizontal: 8, vertical: 0)),
            ),
          ),
        ],
      ),
    );
  }

  void _addColumn(String headerName) {
    setState(() {
      for (int i = 0; i < _data.length; i++) {
        if (i == 0) {
          _data[i].add(TextEditingController(text: headerName));
        } else {
          _data[i].add(TextEditingController(text: ""));
        }
      }
    });
  }

  void _buildMaterialIndex() {
    _materialMap.clear();
    _filteredMaterials.clear();
    _verificationControllers.clear();
    _remarksControllers.clear();

    if (_codeColIndex == null || _qtyColIndex == null) return;

    for (int i = 1; i < _data.length; i++) {
      if (_codeColIndex! >= _data[i].length) continue;

      String code = _data[i][_codeColIndex!].text;
      if (code.trim().isEmpty) continue;

      String desc = (_descColIndex != null && _descColIndex! < _data[i].length) ? _data[i][_descColIndex!].text : "";
      String qty = (_qtyColIndex! < _data[i].length) ? _data[i][_qtyColIndex!].text : "0";
      String uom = (_uomColIndex != null && _uomColIndex! < _data[i].length) ? _data[i][_uomColIndex!].text : "";
      String rate = (_rateColIndex != null && _rateColIndex! < _data[i].length) ? _data[i][_rateColIndex!].text : "0";
      String phy = (_physicalQtyColIndex != null && _physicalQtyColIndex! < _data[i].length) ? _data[i][_physicalQtyColIndex!].text : "";
      String rem = (_remarksColIndex != null && _remarksColIndex! < _data[i].length) ? _data[i][_remarksColIndex!].text : "";

      _materialMap[i] = MaterialItem(
        rowIndex: i,
        code: code,
        description: desc,
        systemQtyStr: qty,
        uom: uom,
        rateStr: rate,
        physicalQty: phy,
        remarks: rem,
      );
    }
    _applySearchFilter();
  }

  void _onSearchChanged(String query) {
    if (_debounce?.isActive ?? false) _debounce!.cancel();
    _debounce = Timer(const Duration(milliseconds: 300), _applySearchFilter);
  }

  void _applySearchFilter() {
    final query = _searchController.text.toLowerCase().trim();
    
    setState(() {
      if (query.isEmpty) {
        _filteredMaterials = _materialMap.values.toList();
      } else {
        _filteredMaterials = _materialMap.values.where((item) {
          return item.code.toLowerCase().contains(query) || 
                 item.description.toLowerCase().contains(query);
        }).toList();
      }
    });
  }

  void _updatePhysicalQty(MaterialItem item, String value) {
    if (_physicalQtyColIndex != null) {
      _data[item.rowIndex][_physicalQtyColIndex!].text = value;
    }
    item.physicalQty = value;
    _hasUnsavedChanges = true;
    _scheduleAutoSave();
  }

  void _updateRemarks(MaterialItem item, String value) {
    if (_remarksColIndex != null) {
      _data[item.rowIndex][_remarksColIndex!].text = value;
    }
    item.remarks = value;
    _hasUnsavedChanges = true;
    _scheduleAutoSave();
  }

  // -------------------------------------------------------------------------
  // UI BUILDERS
  // -------------------------------------------------------------------------

  @override
  Widget build(BuildContext context) {
    return WillPopScope(
      onWillPop: () async {
        if (_hasUnsavedChanges) {
          final shouldExit = await showDialog<bool>(
            context: context,
            builder: (ctx) => AlertDialog(
              title: const Text('Unsaved Changes'),
              content: const Text('You have unsaved changes. Data is auto-saved. Exit anyway?'),
              actions: [
                TextButton(onPressed: () => Navigator.pop(ctx, false), child: const Text('Cancel')),
                TextButton(
                  onPressed: () {
                    _saveSession();
                    Navigator.pop(ctx, true);
                  },
                  child: const Text('Exit', style: TextStyle(color: Colors.red)),
                ),
              ],
            ),
          );
          return shouldExit ?? false;
        }
        return true;
      },
      child: Scaffold(
        appBar: AppBar(
          title: Column(
            crossAxisAlignment: CrossAxisAlignment.start,
            children: [
              Text(_fileName.isEmpty ? 'Inventory Audit' : _fileName, style: const TextStyle(fontSize: 16)),
              if (_hasUnsavedChanges)
                const Text('Auto-saving...', style: TextStyle(fontSize: 11, fontWeight: FontWeight.normal)),
            ],
          ),
          backgroundColor: Colors.indigo,
          foregroundColor: Colors.white,
          elevation: 0,
          actions: [
            if (_data.isNotEmpty)
              IconButton(
                icon: Icon(_isVerificationMode ? Icons.grid_on : Icons.fact_check),
                tooltip: _isVerificationMode ? "Switch to Spreadsheet" : "Switch to Verification",
                onPressed: () {
                  if (!_isVerificationMode && _codeColIndex == null) {
                    _showColumnMappingDialog();
                  } else {
                    if (!_isVerificationMode) _buildMaterialIndex();
                    setState(() => _isVerificationMode = !_isVerificationMode);
                  }
                },
              ),
            PopupMenuButton<String>(
              onSelected: (val) {
                if (val == 'export') _exportExcel();
                if (val == 'map') _showColumnMappingDialog();
                if (val == 'clear') _clearAllCounts();
                if (val == 'save') _saveSession();
              },
              itemBuilder: (ctx) => [
                const PopupMenuItem(value: 'save', child: Row(
                  children: [Icon(Icons.save, size: 18), SizedBox(width: 8), Text('Save Progress')],
                )),
                const PopupMenuItem(value: 'map', child: Text('Remap Columns')),
                const PopupMenuItem(value: 'clear', child: Text('Clear Counts')),
                const PopupMenuDivider(),
                const PopupMenuItem(value: 'export', child: Row(
                  children: [Icon(Icons.download, size: 18), SizedBox(width: 8), Text('Export Excel')],
                )),
              ],
            )
          ],
        ),
        body: Stack(
          children: [
            Column(
              children: [
                if (_data.isNotEmpty) _buildSummaryHeader(),
                if (_data.isNotEmpty) _buildSearchBar(),
                Expanded(
                  child: _data.isEmpty 
                    ? _buildEmptyState() 
                    : _isVerificationMode 
                        ? _buildVerificationView() 
                        : _buildSpreadsheetView(),
                ),
              ],
            ),
            if (_isLoading)
              Container(
                color: Colors.black45,
                child: const Center(child: CircularProgressIndicator(color: Colors.white)),
              ),
          ],
        ),
        floatingActionButton: _data.isEmpty 
          ? FloatingActionButton.extended(
              onPressed: _importExcel,
              icon: const Icon(Icons.upload_file),
              label: const Text("Import Excel"),
              backgroundColor: Colors.indigo,
              foregroundColor: Colors.white,
            )
          : null,
      ),
    );
  }

  Widget _buildEmptyState() {
    return Center(
      child: Column(
        mainAxisAlignment: MainAxisAlignment.center,
        children: [
          Icon(Icons.inventory_2_outlined, size: 80, color: Colors.indigo.shade100),
          const SizedBox(height: 16),
          const Text("No Audit File Loaded", style: TextStyle(fontSize: 18, fontWeight: FontWeight.bold, color: Colors.grey)),
          const SizedBox(height: 8),
          const Text("Import an Excel file to begin stock count", style: TextStyle(color: Colors.grey)),
        ],
      ),
    );
  }

  Widget _buildSummaryHeader() {
    if (!_isVerificationMode || _materialMap.isEmpty) return const SizedBox.shrink();

    int total = _materialMap.length;
    int verified = _materialMap.values.where((i) => i.isVerified).length;
    double progress = total == 0 ? 0 : verified / total;

    return Container(
      padding: const EdgeInsets.symmetric(horizontal: 16, vertical: 12),
      decoration: BoxDecoration(
        color: Colors.white,
        boxShadow: [BoxShadow(color: Colors.black.withOpacity(0.05), blurRadius: 4, offset: const Offset(0, 2))],
      ),
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          Row(
            mainAxisAlignment: MainAxisAlignment.spaceBetween,
            children: [
              Text("Progress: $verified / $total", style: const TextStyle(fontWeight: FontWeight.bold)),
              Text("${(progress * 100).toInt()}%", style: TextStyle(color: Colors.indigo.shade700, fontWeight: FontWeight.bold)),
            ],
          ),
          const SizedBox(height: 6),
          LinearProgressIndicator(value: progress, backgroundColor: Colors.grey.shade200, color: Colors.indigo),
        ],
      ),
    );
  }

  Widget _buildSearchBar() {
    return Container(
      padding: const EdgeInsets.all(8),
      color: Colors.grey.shade50,
      child: TextField(
        controller: _searchController,
        onChanged: _onSearchChanged,
        decoration: InputDecoration(
          hintText: "Search Code, Description...",
          prefixIcon: const Icon(Icons.search, color: Colors.grey),
          filled: true,
          fillColor: Colors.white,
          suffixIcon: _searchController.text.isNotEmpty 
            ? IconButton(
                icon: const Icon(Icons.clear), 
                onPressed: () {
                  _searchController.clear();
                  _onSearchChanged('');
                  FocusScope.of(context).unfocus();
                }) 
            : null,
        ),
      ),
    );
  }

  Widget _buildVerificationView() {
    if (_filteredMaterials.isEmpty) {
      return const Center(child: Text("No items match your search"));
    }

    return ListView.builder(
      padding: const EdgeInsets.only(bottom: 80),
      itemCount: _filteredMaterials.length,
      itemBuilder: (context, index) {
        final item = _filteredMaterials[index];
        return _buildVerificationCard(item, index);
      },
    );
  }

  Widget _buildVerificationCard(MaterialItem item, int listIndex) {
    _verificationControllers[item.rowIndex] ??= TextEditingController(text: item.physicalQty);
    _remarksControllers[item.rowIndex] ??= TextEditingController(text: item.remarks);
    
    final qtyController = _verificationControllers[item.rowIndex]!;
    final remarksController = _remarksControllers[item.rowIndex]!;
    
    if (qtyController.text != item.physicalQty) qtyController.text = item.physicalQty;
    if (remarksController.text != item.remarks) remarksController.text = item.remarks;

    return Card(
      margin: const EdgeInsets.symmetric(horizontal: 8, vertical: 4),
      elevation: 0,
      shape: RoundedRectangleBorder(
        borderRadius: BorderRadius.circular(8),
        side: BorderSide(color: Colors.grey.shade300),
      ),
      color: item.statusColor == Colors.transparent ? Colors.white : item.statusColor,
      child: Padding(
        padding: const EdgeInsets.all(12),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            Row(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Expanded(
                  child: Column(
                    crossAxisAlignment: CrossAxisAlignment.start,
                    children: [
                      Row(
                        children: [
                          Container(
                            padding: const EdgeInsets.symmetric(horizontal: 6, vertical: 2),
                            decoration: BoxDecoration(color: Colors.grey.shade200, borderRadius: BorderRadius.circular(4)),
                            child: Text(item.code, style: const TextStyle(fontWeight: FontWeight.bold, fontSize: 12)),
                          ),
                          const SizedBox(width: 8),
                          Expanded(
                            child: Text("Sys: ${item.systemQuantity} ${item.uom}", 
                              overflow: TextOverflow.ellipsis,
                              style: const TextStyle(fontWeight: FontWeight.w600, color: Colors.black87)
                            ),
                          ),
                        ],
                      ),
                      const SizedBox(height: 4),
                      Text(item.description, maxLines: 2, overflow: TextOverflow.ellipsis, style: const TextStyle(fontSize: 13, color: Colors.black87)),
                      if (item.isVerified && item.variance != 0)
                        Padding(
                          padding: const EdgeInsets.only(top: 4),
                          child: Text(
                            "Variance: ${item.variance > 0 ? '+' : ''}${item.variance}",
                            style: TextStyle(
                              color: item.variance > 0 ? Colors.blue.shade700 : Colors.red.shade700,
                              fontWeight: FontWeight.bold,
                              fontSize: 12,
                            ),
                          ),
                        ),
                    ],
                  ),
                ),
                SizedBox(
                  width: 90,
                  child: TextField(
                    controller: qtyController,
                    focusNode: item.focusNode,
                    keyboardType: const TextInputType.numberWithOptions(decimal: true),
                    textAlign: TextAlign.center,
                    style: const TextStyle(fontSize: 16, fontWeight: FontWeight.bold),
                    textInputAction: TextInputAction.next,
                    decoration: InputDecoration(
                      hintText: "Qty",
                      filled: true,
                      fillColor: Colors.white,
                      border: OutlineInputBorder(borderRadius: BorderRadius.circular(8)),
                      contentPadding: const EdgeInsets.symmetric(horizontal: 8, vertical: 12),
                    ),
                    onChanged: (val) => _updatePhysicalQty(item, val),
                    onSubmitted: (_) => item.remarksFocusNode.requestFocus(),
                  ),
                ),
              ],
            ),
            const SizedBox(height: 8),
            TextField(
              controller: remarksController,
              focusNode: item.remarksFocusNode,
              textInputAction: TextInputAction.next,
              decoration: InputDecoration(
                hintText: "Add remarks (optional)",
                prefixIcon: const Icon(Icons.comment, size: 18),
                filled: true,
                fillColor: Colors.white,
                border: OutlineInputBorder(borderRadius: BorderRadius.circular(8)),
                contentPadding: const EdgeInsets.symmetric(horizontal: 12, vertical: 10),
              ),
              onChanged: (val) => _updateRemarks(item, val),
              onSubmitted: (_) {
                if (listIndex < _filteredMaterials.length - 1) {
                  _filteredMaterials[listIndex + 1].focusNode.requestFocus();
                } else {
                  FocusScope.of(context).unfocus();
                }
              },
            ),
          ],
        ),
      ),
    );
  }

  Widget _buildSpreadsheetView() {
    if (_data.isEmpty) return const SizedBox();

    final headerRow = _data[0];
    final query = _searchController.text.toLowerCase();
    
    final visibleRows = query.isEmpty 
      ? _data.sublist(1) 
      : _data.sublist(1).where((row) => row.any((cell) => cell.text.toLowerCase().contains(query))).toList();

    const double cellWidth = 120.0;
    final double totalWidth = headerRow.length * cellWidth;

    return Column(
      children: [
        SingleChildScrollView(
          controller: _headerScrollController,
          scrollDirection: Axis.horizontal,
          physics: const ClampingScrollPhysics(),
          child: Row(
            children: headerRow.map((c) => Container(
              width: cellWidth,
              height: 50,
              padding: const EdgeInsets.all(4),
              decoration: BoxDecoration(
                color: Colors.grey.shade200,
                border: Border(bottom: BorderSide(color: Colors.grey.shade300), right: BorderSide(color: Colors.grey.shade300)),
              ),
              alignment: Alignment.center,
              child: Text(c.text, style: const TextStyle(fontWeight: FontWeight.bold, fontSize: 12), textAlign: TextAlign.center, maxLines: 2),
            )).toList(),
          ),
        ),
        Expanded(
          child: Scrollbar(
            controller: _bodyScrollController,
            thumbVisibility: true,
            child: SingleChildScrollView(
              controller: _bodyScrollController,
              scrollDirection: Axis.horizontal,
              physics: const ClampingScrollPhysics(),
              child: SizedBox(
                width: totalWidth,
                child: ListView.builder(
                  physics: const AlwaysScrollableScrollPhysics(),
                  itemCount: visibleRows.length,
                  itemBuilder: (context, index) {
                    return _buildRowFast(visibleRows[index], cellWidth);
                  },
                ),
              ),
            ),
          ),
        ),
      ],
    );
  }

  Widget _buildRowFast(List<TextEditingController> row, double width) {
    return Container(
      height: 50,
      decoration: BoxDecoration(
        color: Colors.white,
        border: Border(bottom: BorderSide(color: Colors.grey.shade200)),
      ),
      child: Row(
        children: row.map((controller) {
          return Container(
            width: width,
            padding: const EdgeInsets.symmetric(horizontal: 4),
            decoration: BoxDecoration(
              border: Border(right: BorderSide(color: Colors.grey.shade200)),
            ),
            child: TextField(
              controller: controller,
              style: const TextStyle(fontSize: 13),
              decoration: const InputDecoration(
                border: InputBorder.none,
                isDense: true,
                contentPadding: EdgeInsets.symmetric(horizontal: 4, vertical: 14),
              ),
            ),
          );
        }).toList(),
      ),
    );
  }

  void _clearAllCounts() {
    showDialog(
      context: context,
      builder: (ctx) => AlertDialog(
        title: const Text("Clear All Data?"),
        content: const Text("This will reset all Physical Quantities and Remarks. This cannot be undone."),
        actions: [
          TextButton(onPressed: () => Navigator.pop(ctx), child: const Text("Cancel")),
          TextButton(
            onPressed: () {
              setState(() {
                if (_physicalQtyColIndex != null) {
                  for (int i = 1; i < _data.length; i++) {
                     if (_physicalQtyColIndex! < _data[i].length) {
                       _data[i][_physicalQtyColIndex!].text = "";
                     }
                  }
                }
                if (_remarksColIndex != null) {
                  for (int i = 1; i < _data.length; i++) {
                     if (_remarksColIndex! < _data[i].length) {
                       _data[i][_remarksColIndex!].text = "";
                     }
                  }
                }
                _buildMaterialIndex();
              });
              _clearSession();
              Navigator.pop(ctx);
            }, 
            child: const Text("Clear All", style: TextStyle(color: Colors.red))
          ),
        ],
      )
    );
  }

  void _showSnackBar(String msg, Color color) {
    if (!mounted) return;
    ScaffoldMessenger.of(context).showSnackBar(SnackBar(content: Text(msg), backgroundColor: color, duration: const Duration(seconds: 2)));
  }
}