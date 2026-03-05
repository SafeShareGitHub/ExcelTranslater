import 'package:flutter/material.dart';
import 'package:file_picker/file_picker.dart';
import 'package:excel/excel.dart' as excel_pkg;
import 'dart:typed_data';
import 'dart:async';
import 'dart:html' as html;

/// Hauptanwendung für den Excel Logic Validator.
/// Diese App ermöglicht das Hochladen, Validieren und Exportieren von Excel-Logiken.
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
  // --- STATE VARIABLEN ---
  List<List<dynamic>> _data = [];
  Set<int> _errorRows = {};
  bool _isLoading = false;
  bool _showOnlyErrors = false;
  String _loadingDots = "";
  Timer? _timer;

  // UI Konfiguration
  double _cellWidth = 300.0;
  double _cellHeight = 45.0;

  // --- ANIMATION & HELPER ---
  
  /// Startet die Punkte-Animation im Lade-Overlay.
  void _startLoadingAnimation() {
    _loadingDots = "";
    _timer = Timer.periodic(const Duration(milliseconds: 500), (timer) {
      if (!mounted) return;
      setState(() {
        _loadingDots = (_loadingDots.length < 3) ? "$_loadingDots." : "";
      });
    });
  }

  /// Stoppt die Animation und setzt den Text zurück.
  void _stopLoadingAnimation() {
    _timer?.cancel();
    setState(() => _loadingDots = "");
  }

  // --- REKURSIVE SYNTAX-PRÜFUNG (LOGIK) ---

  /// Prüft eine Zeile auf logische Fehler in Spalte F.
  bool _hasLogicError(String formula) {
    if (formula.isEmpty) return false;
    String trimmed = formula.trim();

    // 1. Grundlegende Syntax: Klammer-Balance
    if (!_isBracketBalanced(trimmed)) return true;

    // 2. Struktur-Check: Muss eine Excel IF-Funktion sein
    if (!trimmed.toUpperCase().startsWith("IF(") || !trimmed.endsWith(")")) return true;

    try {
      // Extraktion des Inhalts innerhalb der äußeren Klammern
      String content = trimmed.substring(3, trimmed.length - 1);
      List<String> parts = _splitByTopLevelSemicolon(content);

      // Eine valide IF-Funktion benötigt exakt 3 Argumente: IF(Bedingung; Dann; Sonst)
      if (parts.length != 3) return true;

      // Rekursive Prüfung des Bedingungsteils (Argument 1)
      return !_isConditionValid(parts[0]);
    } catch (e) {
      debugPrint("Validierungsfehler: $e");
      return true;
    }
  }

  /// Prüft rekursiv, ob die Bedingung (AND, OR, NOT Schachtelungen) syntaktisch korrekt ist.
  bool _isConditionValid(String exp) {
    exp = exp.trim().toUpperCase();
    if (exp.isEmpty) return false;

    // Prüfung von AND- oder OR-Schachtelungen
    if (exp.startsWith("AND(") || exp.startsWith("OR(")) {
      if (!exp.endsWith(")")) return false;
      int startIdx = exp.indexOf("(") + 1;
      String inner = exp.substring(startIdx, exp.length - 1);
      if (inner.trim().isEmpty) return false;

      List<String> subParts = _splitByTopLevelSemicolon(inner);
      for (var p in subParts) {
        if (!_isConditionValid(p)) return false;
      }
      return true;
    }

    // Prüfung von NOT-Schachtelungen
    if (exp.startsWith("NOT(")) {
      if (!exp.endsWith(")")) return false;
      String inner = exp.substring(4, exp.length - 1);
      return _isConditionValid(inner);
    }

    // Basis-Ausdruck: Muss Operatoren oder Inhalt haben
    bool hasOperator = exp.contains("=") ||
        exp.contains(">") ||
        exp.contains("<") ||
        exp.contains("<>");
    return hasOperator || exp.isNotEmpty;
  }

  /// Hilfsfunktion zur Prüfung, ob alle geöffneten Klammern wieder geschlossen wurden.
  bool _isBracketBalanced(String s) {
    int count = 0;
    for (var char in s.runes) {
      if (char == 40) count++; // ASCII for '('
      if (char == 41) count--; // ASCII for ')'
      if (count < 0) return false; // Mehr geschlossene als offene Klammern zu einem Zeitpunkt
    }
    return count == 0;
  }

  /// Validiert den gesamten Datensatz und speichert Fehler-Indizes.
  void _validateAllRows() {
    _errorRows.clear();
    for (int i = 0; i < _data.length; i++) {
      if (_data[i].length > 5) {
        if (_hasLogicError(_data[i][5].toString())) {
          _errorRows.add(i);
        }
      }
    }
    setState(() {});
  }

  // --- CORE FUNKTIONEN (DATEI & VERARBEITUNG) ---

  /// Öffnet den FilePicker und lädt die Excel-Daten in den Speicher.
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
        await Future.delayed(const Duration(milliseconds: 500));
        Uint8List bytes = result.files.first.bytes!;
        var excel = excel_pkg.Excel.decodeBytes(bytes);

        List<List<dynamic>> rows = [];
        for (var table in excel.tables.keys) {
          for (var row in excel.tables[table]!.rows) {
            // Konvertiere Zellen in Strings, fülle leere Zellen auf
            var rowData = row.map((cell) => cell?.value?.toString() ?? "").toList();
            while (rowData.length < 8) rowData.add("");
            rows.add(rowData);
          }
          break; // Nur das erste Tabellenblatt verarbeiten
        }
        setState(() {
          _data = rows;
          _showOnlyErrors = false;
        });
        _validateAllRows();
      }
    } catch (e) {
      debugPrint("Fehler beim Laden: $e");
    } finally {
      _stopLoadingAnimation();
      setState(() => _isLoading = false);
    }
  }

  /// Startet die Übersetzung der Logik von Spalte F nach Spalte D.
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
    _validateAllRows();
  }

  /// Extrahiert den Bedingungsteil einer IF-Formel.
  String _parseExcelLogic(String formula) {
    try {
      int firstOpen = formula.indexOf('(');
      String content = formula.substring(firstOpen + 1, formula.length - 1);
      List<String> topLevelParts = _splitByTopLevelSemicolon(content);
      if (topLevelParts.isEmpty) return formula;
      return _recursiveTranslate(topLevelParts[0]);
    } catch (e) {
      return "ERROR: Parsing";
    }
  }

  /// Übersetzt Excel-Syntax rekursiv in das Zielformat.
  String _recursiveTranslate(String exp) {
    exp = exp.trim();
    if (exp.toUpperCase().startsWith("AND(")) {
      String inner = exp.substring(4, exp.length - 1);
      List<String> parts = _splitByTopLevelSemicolon(inner);
      return "(${parts.map((p) => _recursiveTranslate(p)).join(" && ")})";
    }
    if (exp.toUpperCase().startsWith("OR(")) {
      String inner = exp.substring(3, exp.length - 1);
      List<String> parts = _splitByTopLevelSemicolon(inner);
      return "(${parts.map((p) => _recursiveTranslate(p)).join(" || ")})";
    }
    if (exp.toUpperCase().startsWith("NOT(")) {
      String inner = exp.substring(4, exp.length - 1);
      return "!(${_recursiveTranslate(inner)})";
    }
    if (exp.contains('=')) {
      List<String> sides = exp.split('=');
      String left = sides[0].trim();
      String right = sides.length > 1 ? sides[1].trim() : "";
      return "\$$left\$==$right";
    }
    return exp;
  }

  /// Splittet einen String an Semikolons, ignoriert aber Semikolons innerhalb von Klammern.
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

  /// Generiert die Excel-Datei und triggert den Browser-Download.
  void _downloadExcel() {
    if (_data.isEmpty) return;
    try {
      var saveExcel = excel_pkg.Excel.createExcel();
      var sheet = saveExcel['Sheet1'];

      // Design für Fehlerzeilen (Hintergrundfarbe Rot)
      excel_pkg.CellStyle errorStyle = excel_pkg.CellStyle(
        backgroundColorHex: excel_pkg.ExcelColor.fromHexString("#FFCCCC"),
        fontFamily: excel_pkg.getFontFamily(excel_pkg.FontFamily.Calibri),
      );

      for (int r = 0; r < _data.length; r++) {
        bool isRowError = _errorRows.contains(r);
        for (int c = 0; c < _data[r].length; c++) {
          var cellIndex = excel_pkg.CellIndex.indexByColumnRow(columnIndex: c, rowIndex: r);
          var cell = sheet.cell(cellIndex);
          
          cell.value = excel_pkg.TextCellValue(_data[r][c].toString());
          
          if (isRowError) {
            cell.cellStyle = errorStyle;
          }
        }
      }

      // Web-Download-Fix: Konvertierung in Uint8List und Blob-Trigger
      final List<int>? excelBytes = saveExcel.save();
      if (excelBytes != null) {
        final Uint8List downloadBytes = Uint8List.fromList(excelBytes);
        final blob = html.Blob([downloadBytes], 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        final url = html.Url.createObjectUrlFromBlob(blob);
        
        html.AnchorElement(href: url)
          ..setAttribute("download", "logic_check_results.xlsx")
          ..click();
          
        html.Url.revokeObjectUrl(url);
      }
    } catch (e) {
      debugPrint("Download-Fehler: $e");
    }
  }

  // --- UI WIDGETS ---

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: const Text('Excel Logic Pro Validator'),
        backgroundColor: Colors.blueGrey[900],
        foregroundColor: Colors.white,
      ),
      body: Stack(
        children: [
          Column(
            children: [
              _buildToolbar(),
              const Divider(height: 1),
              Expanded(
                child: _data.isEmpty 
                  ? const Center(child: Text('Bitte XLSX hochladen')) 
                  : _buildExcelGrid(),
              ),
            ],
          ),
          if (_isLoading) _buildLoadingOverlay(),
        ],
      ),
    );
  }

  /// Erstellt die Werkzeugleiste mit Buttons und Status-Switch.
  Widget _buildToolbar() {
    return Container(
      padding: const EdgeInsets.all(10),
      color: Colors.grey[100],
      child: Wrap(
        spacing: 15,
        runSpacing: 10,
        crossAxisAlignment: WrapCrossAlignment.center,
        children: [
          ElevatedButton.icon(
            onPressed: _isLoading ? null : _pickAndParseExcel,
            icon: const Icon(Icons.upload_file),
            label: const Text("Upload"),
          ),
          ElevatedButton.icon(
            onPressed: (_data.isEmpty || _isLoading) ? null : _processFormulas,
            icon: const Icon(Icons.auto_fix_high),
            label: const Text("Übersetze F -> D"),
            style: ElevatedButton.styleFrom(
              backgroundColor: Colors.orange[900],
              foregroundColor: Colors.white,
            ),
          ),
          ElevatedButton.icon(
            onPressed: (_data.isEmpty || _isLoading) ? null : _downloadExcel,
            icon: const Icon(Icons.download),
            label: const Text("Download"),
            style: ElevatedButton.styleFrom(
              backgroundColor: Colors.green[800],
              foregroundColor: Colors.white,
            ),
          ),
          const VerticalDivider(width: 20),
          Row(
            mainAxisSize: MainAxisSize.min,
            children: [
              const Text("Nur Fehler:", style: TextStyle(fontWeight: FontWeight.bold)),
              Switch(
                value: _showOnlyErrors,
                onChanged: (val) => setState(() => _showOnlyErrors = val),
                activeColor: Colors.red,
              ),
              if (_errorRows.isNotEmpty)
                Container(
                  padding: const EdgeInsets.symmetric(horizontal: 8, vertical: 4),
                  decoration: BoxDecoration(
                    color: Colors.red,
                    borderRadius: BorderRadius.circular(12),
                  ),
                  child: Text(
                    "${_errorRows.length} Fehler",
                    style: const TextStyle(
                      color: Colors.white,
                      fontSize: 11,
                      fontWeight: FontWeight.bold,
                    ),
                  ),
                ),
            ],
          ),
          OutlinedButton(
            onPressed: () => _showSizeDialog(true),
            child: const Text("Breite"),
          ),
          OutlinedButton(
            onPressed: () => _showSizeDialog(false),
            child: const Text("Höhe"),
          ),
        ],
      ),
    );
  }

  /// Baut das Daten-Grid basierend auf dem Filterstatus auf.
  Widget _buildExcelGrid() {
    int colCount = _data.isNotEmpty ? _data.first.length : 8;
    List<MapEntry<int, List<dynamic>>> filteredEntries = _data.asMap().entries.where((entry) {
      if (!_showOnlyErrors) return true;
      return _errorRows.contains(entry.key);
    }).toList();

    return Scrollbar(
      thumbVisibility: true,
      child: SingleChildScrollView(
        scrollDirection: Axis.vertical,
        child: SingleChildScrollView(
          scrollDirection: Axis.horizontal,
          child: Column(
            crossAxisAlignment: CrossAxisAlignment.start,
            children: [
              // Header-Zeile
              Row(
                children: [
                  _buildCell("#", 50, isHeader: true),
                  ...List.generate(colCount, (i) => _buildCell(_getColumnLabel(i), _cellWidth, isHeader: true))
                ],
              ),
              // Daten-Zeilen
              ...filteredEntries.map((entry) {
                bool isRowError = _errorRows.contains(entry.key);
                return Row(
                  children: [
                    _buildCell((entry.key + 1).toString(), 50, isHeader: true, isError: isRowError),
                    ...entry.value.asMap().entries.map((cellEntry) {
                      bool isColumnFError = isRowError && cellEntry.key == 5;
                      return _buildCell(cellEntry.value.toString(), _cellWidth, isError: isColumnFError);
                    })
                  ],
                );
              }).toList(),
            ],
          ),
        ),
      ),
    );
  }

  /// Erstellt eine einzelne Zelle mit bedingtem Styling.
  Widget _buildCell(String text, double width, {bool isHeader = false, bool isError = false}) {
    return Container(
      width: width,
      height: _cellHeight,
      decoration: BoxDecoration(
        color: isHeader
            ? (isError ? Colors.red[200] : Colors.grey[300])
            : (isError ? Colors.red[50] : Colors.white),
        border: Border.all(
          color: isError ? Colors.red : Colors.grey[400]!,
          width: isError ? 1.2 : 0.5,
        ),
      ),
      alignment: isHeader ? Alignment.center : Alignment.centerLeft,
      padding: const EdgeInsets.symmetric(horizontal: 5),
      child: Text(
        text,
        overflow: TextOverflow.ellipsis,
        style: TextStyle(
          fontSize: 11,
          fontWeight: isError ? FontWeight.bold : FontWeight.normal,
          color: isError ? Colors.red[900] : Colors.black,
        ),
      ),
    );
  }

  /// Zeigt das Lade-Overlay während asynchronen Operationen.
  Widget _buildLoadingOverlay() {
    return Container(
      color: Colors.black54,
      child: Center(
        child: Container(
          width: 300,
          padding: const EdgeInsets.all(24),
          decoration: BoxDecoration(
            color: Colors.white,
            borderRadius: BorderRadius.circular(12),
          ),
          child: Column(
            mainAxisSize: MainAxisSize.min,
            children: [
              Text(
                'Prüfe Logik$_loadingDots',
                style: const TextStyle(fontSize: 18, fontWeight: FontWeight.bold),
              ),
              const SizedBox(height: 20),
              const LinearProgressIndicator(
                backgroundColor: Colors.grey,
                valueColor: AlwaysStoppedAnimation<Color>(Colors.green),
                minHeight: 8,
              ),
            ],
          ),
        ),
      ),
    );
  }

  /// Dialog zum Ändern der Grid-Dimensionen.
  void _showSizeDialog(bool isWidth) {
    TextEditingController controller = TextEditingController(
      text: (isWidth ? _cellWidth : _cellHeight).toInt().toString()
    );
    showDialog(
      context: context,
      builder: (context) => AlertDialog(
        title: Text(isWidth ? "Zellenbreite" : "Zellenhöhe"),
        content: TextField(
          controller: controller,
          keyboardType: TextInputType.number,
          decoration: const InputDecoration(suffixText: "px"),
        ),
        actions: [
          TextButton(
            onPressed: () {
              setState(() {
                double? val = double.tryParse(controller.text);
                if (val != null) {
                  if (isWidth) _cellWidth = val; else _cellHeight = val;
                }
              });
              Navigator.pop(context);
            },
            child: const Text("Übernehmen"),
          )
        ],
      ),
    );
  }

  /// Konvertiert Index in Excel-Spaltennamen (0 -> A, 26 -> AA).
  String _getColumnLabel(int index) {
    String label = "";
    while (index >= 0) {
      label = String.fromCharCode((index % 26) + 65) + label;
      index = (index ~/ 26) - 1;
    }
    return label;
  }

  @override
  void dispose() {
    _timer?.cancel();
    super.dispose();
  }
}