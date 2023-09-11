import 'package:flutter_excel/excel.dart';

class ExcelHelper {
  late Excel file;
  String sheet;

  ExcelHelper.init({required List<int> dataInBytes, required this.sheet}) {
    file = Excel.decodeBytes(dataInBytes);
  }

  List<List<Data?>> getData() {
    if (file.tables.isNotEmpty) {
      for (var table in file.tables.keys) {
        var rows = file.tables[table];
        if (rows != null) {
          var array = rows.rows;
          array.removeAt(0);
          return array;
        } else {
          return [];
        }
      }
    } else {
      return [];
    }
    return [];
  }

  void insertData(List<dynamic> values) {
    file.insertRowIterables(
        sheet, values, (file.tables[sheet]?.maxRows as int));
  }

  List<int>? getBytes() {
    return file.encode();
  }

  void updateData(int index, List<dynamic> values) {
    for (var c = 0; c < values.length; c++) {
      file.updateCell(
          sheet,
          CellIndex.indexByColumnRow(columnIndex: c, rowIndex: index),
          values[c]);
    }
  }

  void deleteData(index) {
    file.removeRow(sheet, index);
  }
}
