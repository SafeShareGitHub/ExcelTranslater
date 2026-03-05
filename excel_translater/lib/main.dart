import 'package:flutter/material.dart';
import 'package:file_picker/file_picker.dart';
import 'package:excel/excel.dart' as excel_pkg;
import 'dart:typed_data';
import 'dart:async';
import 'dart:html' as html;

void main() => runApp(const MaterialApp(
      home: ExcelViewer(),
      debugShowCheckedModeBanner: false,
    ));

class ExcelViewer extends StatefulWidget {
  const ExcelViewer({super.key});

  @override
  State<ExcelViewer> createState() => _ExcelViewerState();
}

class _ExcelViewerState extends State<ExcelViewer> {
  List<List<dynamic>> _data = [];
  bool _isLoading = false;
  String _loadingDots = ""; 
  Timer? _timer;

  double _cellWidth = 300.0; 
  double _cellHeight = 45.0;

  void _startLoadingAnimation() {
    _loadingDots = "";
    _timer = Timer.periodic(const Duration(milliseconds: 500), (timer) {
      if (!mounted) return;
      setState(() {
        _loadingDots = (_loadingDots.length < 3) ? "$_loadingDots." : "";
      });
    });
  }

  void _stopLoadingAnimation() {
    _timer?.cancel();
    setState(() => _loadingDots = "");
  }

  Future<void> _pickAndParseExcel() async {
    setState(() => _isLoading = true);
    _startLoadingAnimation();

    try {
      FilePickerResult? result = await FilePicker.platform.pickFiles(
        type: FileType.custom,
        allowedExtensions: ['xlsx'],
        withData: true,
      );

      if (result != null && result.files.first.bytes != null) {
        await Future.delayed(const Duration(milliseconds: 800));
        Uint8List bytes = result.files.first.bytes!;
        var excel = excel_pkg.Excel.decodeBytes(bytes);
        
        List<List<dynamic>> rows = [];
        for (var table in excel.tables.keys) {
          for (var row in excel.tables[table]!.rows) {
            var rowData = row.map((cell) => cell?.value?.toString() ?? "").toList();
            while (rowData.length < 8) rowData.add("");
            rows.add(rowData);
          }
          break; 
        }
        setState(() { _data = rows; });
      }
    } catch (e) {
      debugPrint("Fehler: $e");
    } finally {
      _stopLoadingAnimation();
      setState(() => _isLoading = false);
    }
  }

  // --- DER INTELLIGENTE LOGIK-PARSER ---
  void _processFormulas() {
    if (_data.isEmpty) return;
    setState(() {
      for (var row in _data) {
        if (row.length > 5) {
          String input = row[5].toString().trim();
          if (input.toUpperCase().startsWith("IF(")) {
            row[3] = _parseExcelLogic(input);
          }
        }
      }
    });
  }

  String _parseExcelLogic(String formula) {
    try {
      // 1. Extrahiere den Bedingungsteil aus IF(Bedingung; Dann; Sonst)
      // Wir suchen das erste Klammerpaar nach "IF"
      int firstOpen = formula.indexOf('(');
      int lastSemicolon = formula.lastIndexOf(';');
      // Wir brauchen nur den Teil vor dem vorletzten Semikolon (der die Bedingung ist)
      // Aber Excel-IFs sind tückisch. Wir nehmen den Content und splitten klug.
      String content = formula.substring(firstOpen + 1, formula.length - 1);
      
      // Den "Dann" und "Sonst" Teil abschneiden, um nur die Bedingung zu behalten
      List<String> topLevelParts = _splitByTopLevelSemicolon(content);
      if (topLevelParts.isEmpty) return formula;
      
      String conditionOnly = topLevelParts[0];

      return _recursiveTranslate(conditionOnly);
    } catch (e) {
      return "ERROR: Parsing";
    }
  }

  String _recursiveTranslate(String exp) {
    exp = exp.trim();

    // AND(a;b) -> (translatedA && translatedB)
    if (exp.toUpperCase().startsWith("AND(")) {
      String inner = exp.substring(4, exp.length - 1);
      List<String> parts = _splitByTopLevelSemicolon(inner);
      return "(${parts.map((p) => _recursiveTranslate(p)).join(" && ")})";
    }

    // OR(a;b) -> (translatedA || translatedB)
    if (exp.toUpperCase().startsWith("OR(")) {
      String inner = exp.substring(3, exp.length - 1);
      List<String> parts = _splitByTopLevelSemicolon(inner);
      return "(${parts.map((p) => _recursiveTranslate(p)).join(" || ")})";
    }

    // NOT(a) -> !(translatedA)
    if (exp.toUpperCase().startsWith("NOT(")) {
      String inner = exp.substring(4, exp.length - 1);
      return "!(${_recursiveTranslate(inner)})";
    }

    // Basis-Fall: Variable=Wert -> $Variable$=="Wert"
    if (exp.contains('=')) {
      List<String> sides = exp.split('=');
      String left = sides[0].trim();
      String right = sides[1].trim();
      return "\$$left\$==$right";
    }

    return exp;
  }

  // Hilfsfunktion um Semikolons nur auf der obersten Ebene zu splitten (ignoriert Semikolons in Klammern)
  List<String> _splitByTopLevelSemicolon(String input) {
    List<String> parts = [];
    int bracketLevel = 0;
    int start = 0;
    for (int i = 0; i < input.length; i++) {
      if (input[i] == '(') bracketLevel++;
      if (input[i] == ')') bracketLevel--;
      if (input[i] == ';' && bracketLevel == 0) {
        parts.add(input.substring(start, i));
        start = i + 1;
      }
    }
    parts.add(input.substring(start));
    return parts;
  }

  // --- DOWNLOAD & UI (Unverändert) ---
  void _downloadExcel() {
    if (_data.isEmpty) return;
    var saveExcel = excel_pkg.Excel.createExcel();
    var sheet = saveExcel['Sheet1'];
    for (int r = 0; r < _data.length; r++) {
      for (int c = 0; c < _data[r].length; c++) {
        var cell = sheet.cell(excel_pkg.CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r));
        cell.value = excel_pkg.TextCellValue(_data[r][c].toString());
      }
    }
    final List<int>? bytes = saveExcel.save();
    if (bytes != null) {
      final blob = html.Blob([bytes], 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      final url = html.Url.createObjectUrlFromBlob(blob);
      html.AnchorElement(href: url)..setAttribute("download", "logic_export.xlsx")..click();
      html.Url.revokeObjectUrl(url);
    }
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(title: const Text('Excel Logic Translator'), backgroundColor: Colors.blueGrey[900], foregroundColor: Colors.white),
      body: Stack(children: [
        Column(children: [
          _buildToolbar(),
          const Divider(height: 1),
          Expanded(child: _data.isEmpty ? const Center(child: Text('Datei hochladen...')) : _buildExcelGrid()),
        ]),
        if (_isLoading) _buildLoadingOverlay(),
      ]),
    );
  }

  Widget _buildToolbar() {
    return Container(
      padding: const EdgeInsets.all(10), color: Colors.grey[200],
      child: Wrap(spacing: 10, runSpacing: 10, children: [
        ElevatedButton.icon(onPressed: _isLoading ? null : _pickAndParseExcel, icon: const Icon(Icons.upload), label: const Text("Upload")),
        ElevatedButton.icon(onPressed: (_data.isEmpty || _isLoading) ? null : _processFormulas, icon: const Icon(Icons.transform), label: const Text("Übersetze F -> D"), style: ElevatedButton.styleFrom(backgroundColor: Colors.orange[900], foregroundColor: Colors.white)),
        ElevatedButton.icon(onPressed: (_data.isEmpty || _isLoading) ? null : _downloadExcel, icon: const Icon(Icons.download), label: const Text("Download XLSX"), style: ElevatedButton.styleFrom(backgroundColor: Colors.green[800], foregroundColor: Colors.white)),
        OutlinedButton(onPressed: () => _showSizeDialog(true), child: const Text("Breite")),
        OutlinedButton(onPressed: () => _showSizeDialog(false), child: const Text("Höhe")),
      ]),
    );
  }

  Widget _buildExcelGrid() {
    int colCount = _data.isNotEmpty ? _data.first.length : 8;
    return Scrollbar(thumbVisibility: true, child: SingleChildScrollView(scrollDirection: Axis.vertical, child: SingleChildScrollView(scrollDirection: Axis.horizontal, child: Column(crossAxisAlignment: CrossAxisAlignment.start, children: [
      Row(children: [_buildCell("#", 50, isHeader: true), ...List.generate(colCount, (i) => _buildCell(_getColumnLabel(i), _cellWidth, isHeader: true))]),
      ..._data.asMap().entries.map((entry) => Row(children: [_buildCell((entry.key + 1).toString(), 50, isHeader: true), ...entry.value.map((cell) => _buildCell(cell.toString(), _cellWidth))])).toList(),
    ]))));
  }

  Widget _buildCell(String text, double width, {bool isHeader = false}) {
    return Container(width: width, height: _cellHeight, decoration: BoxDecoration(color: isHeader ? Colors.grey[300] : Colors.white, border: Border.all(color: Colors.grey[400]!, width: 0.5)), alignment: isHeader ? Alignment.center : Alignment.centerLeft, padding: const EdgeInsets.symmetric(horizontal: 5), child: Text(text, overflow: TextOverflow.ellipsis, style: const TextStyle(fontSize: 11)));
  }

  Widget _buildLoadingOverlay() {
    return Container(color: Colors.black54, child: Center(child: Container(width: 300, padding: const EdgeInsets.all(24), decoration: BoxDecoration(color: Colors.white, borderRadius: BorderRadius.circular(12)), child: Column(mainAxisSize: MainAxisSize.min, children: [Text('Übersetze Logik$_loadingDots', style: const TextStyle(fontSize: 18, fontWeight: FontWeight.bold)), const SizedBox(height: 20), const LinearProgressIndicator(backgroundColor: Colors.grey, valueColor: AlwaysStoppedAnimation<Color>(Colors.green), minHeight: 8)]))));
  }

  void _showSizeDialog(bool isWidth) {
    TextEditingController controller = TextEditingController(text: (isWidth ? _cellWidth : _cellHeight).toInt().toString());
    showDialog(context: context, builder: (context) => AlertDialog(title: Text(isWidth ? "Breite" : "Höhe"), content: TextField(controller: controller, keyboardType: TextInputType.number), actions: [TextButton(onPressed: () { setState(() { double? val = double.tryParse(controller.text); if (val != null) { if (isWidth) _cellWidth = val; else _cellHeight = val; } }); Navigator.pop(context); }, child: const Text("OK"))]));
  }

  String _getColumnLabel(int index) { String label = ""; while (index >= 0) { label = String.fromCharCode((index % 26) + 65) + label; index = (index ~/ 26) - 1; } return label; }

  @override
  void dispose() { _timer?.cancel(); super.dispose(); }
}