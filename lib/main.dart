import 'dart:io';
import 'dart:typed_data';

import 'package:docx_template/docx_template.dart';
import 'package:file_saver/file_saver.dart';
import 'package:firebase_core/firebase_core.dart';
import 'package:firebase_storage/firebase_storage.dart';
import 'package:firebase_storage_sample_app/firebase_options.dart';
import 'package:flutter/material.dart';
import 'package:flutter_excel/excel.dart';
import 'package:intl/intl.dart';

void main() async {
  WidgetsFlutterBinding.ensureInitialized();
  await Firebase.initializeApp(
    options: DefaultFirebaseOptions.currentPlatform,
  );
  runApp(MyApp());
}

class MyApp extends StatelessWidget {
  // This widget is the root of your application.
  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      debugShowCheckedModeBanner: false,
      title: 'Flutter Demo',
      theme: ThemeData(
        colorScheme: ColorScheme.fromSeed(seedColor: Colors.deepPurple),
        useMaterial3: true,
      ),
      home: MyHomePage(),
    );
  }
}

class MyHomePage extends StatefulWidget {
  @override
  State<MyHomePage> createState() => _MyHomePageState();
}

class _MyHomePageState extends State<MyHomePage> {
  final numberFormat = NumberFormat("#,##0.00", "en_US");
  final storage = FirebaseStorage.instance;
  final storageRef = FirebaseStorage.instance.ref();
  bool _isLoading = false;
  final excelForm = GlobalKey<FormState>();
  final wordForm = GlobalKey<FormState>();
  int indexToEdit = 0;
  List<List<Data?>> dataExcel = [];
  late Excel excelFile;
  TextEditingController _name = TextEditingController();
  TextEditingController _age = TextEditingController();
  TextEditingController _client = TextEditingController();
  TextEditingController _amount = TextEditingController();

  void edit(int index, String name, int age) {
    indexToEdit = index;
    _name.value = TextEditingValue(text: name);
    _age.value = TextEditingValue(text: age.toString());
  }

  Future<void> deleteRow(int index) async {
    excelFile.removeRow("Hoja1", index);
    await uploadExcelFile();
    await downloadFile();
  }

  int findArray(List<int> source, List<int> target) {
    for (int i = 0; i <= source.length - target.length; i++) {
      bool found = true;

      for (int j = 0; j < target.length; j++) {
        if (source[i + j] != target[j]) {
          found = false;
          break;
        }
      }

      if (found) {
        return i;
      }
    }

    return -1; // Si no se encontró la coincidencia
  }

  Future<void> downloadWordFile() async {
    setState(() {
      _isLoading = true;
    });
    try {
      if (wordForm.currentState!.validate()) {
        final islandRef =
            storageRef.child("flutter_invoice_template_sample.docx");
        final Uint8List? data = await islandRef.getData();
        var doc = await DocxTemplate.fromBytes(data as List<int>);
        var monto = [133, 155, 157, 156, 164, 157, 135];

        var content = Content();
        content.add(TextContent("[cliente]", _client.value.text));
        content.add(TextContent(
            "[monto]", numberFormat.format(double.parse(_amount.value.text))));
        var docGenerated = await doc.generate(content);
        print(docGenerated);
        print(monto);
        print("Array!: ${findArray(docGenerated as List<int>, monto)}");
        // final fileGenerated = File('generated.docx');
        // if (docGenerated != null)
        //   await fileGenerated.writeAsBytes(docGenerated,
        //       flush: true, mode: FileMode.append);
        // await FileSaver.instance.saveFile(
        //     name: "Factura",
        //     bytes: await fileGenerated.readAsBytes(),
        //     ext: "docx");
        _amount.value = TextEditingValue.empty;
        _client.value = TextEditingValue.empty;
      }
    } catch (error) {
      print("Error descargando archivo word: $error");
    }
    setState(() {
      _isLoading = false;
    });
  }

  Future<void> downloadExcelFile() async {
    setState(() {
      _isLoading = true;
    });
    final islandRef = storageRef.child("Firebase_storage_sample_archive.xlsx");
    final Uint8List? data = await islandRef.getData();
    Excel.decodeBytes(data?.toList() as List<int>).save(fileName: "Data.xlsx");
    setState(() {
      _isLoading = false;
    });
  }

  Future<void> uploadExcelFile() async {
    final mountainsRef =
        storageRef.child("Firebase_storage_sample_archive.xlsx");
    await mountainsRef.putData(excelFile.encode() as Uint8List);
  }

  Future<void> insert() async {
    setState(() {
      _isLoading = true;
    });
    try {
      if (excelForm.currentState!.validate()) {
        print("Index: $indexToEdit");
        if (indexToEdit == 0) {
          print(excelFile.tables["Hoja1"]?.maxRows);
          excelFile.insertRowIterables(
              "Hoja1",
              [_name.value.text, int.parse(_age.value.text)],
              (excelFile.tables["Hoja1"]?.maxRows as int));
          await uploadExcelFile();
        } else {
          excelFile.updateCell(
              "Hoja1",
              CellIndex.indexByColumnRow(
                  columnIndex: 0, rowIndex: indexToEdit - 1),
              _name.value.text);
          excelFile.updateCell(
              "Hoja1",
              CellIndex.indexByColumnRow(
                  columnIndex: 1, rowIndex: indexToEdit - 1),
              int.parse(_age.value.text));
          await uploadExcelFile();
        }

        await downloadFile();
        // excelFile.save(fileName: "Firebase_storage_sample_archive.xlsx");
      }
    } catch (error) {
      print("Error escribiendo data: $error");
    }

    setState(() {
      _isLoading = false;
    });
  }

  Future<void> downloadFile() async {
    setState(() {
      _isLoading = true;
    });
    try {
      indexToEdit = 0;
      _name.value = TextEditingValue.empty;
      _age.value = TextEditingValue.empty;
      final islandRef =
          storageRef.child("Firebase_storage_sample_archive.xlsx");
      const oneMegabyte = 1024 * 1024;
      final Uint8List? data = await islandRef.getData();
      excelFile = Excel.decodeBytes(data?.toList() as List<int>);
      for (var table in excelFile.tables.keys) {
        print(table); //sheet Name
        print(excelFile.tables[table]?.maxCols);
        print(excelFile.tables[table]?.maxRows);
        var rows = excelFile.tables[table];
        if (rows != null) {
          var array = rows.rows;
          array.removeAt(0);
          setState(() {
            dataExcel = array;
          });
        }
      }
    } catch (e) {
      print("Error downloaded file: $e");
    }

    setState(() {
      _isLoading = false;
    });
  }

  @override
  void initState() {
    super.initState();
    downloadFile();
  }

  @override
  Widget build(BuildContext context) {
    var size = MediaQuery.of(context).size;
    return Scaffold(
      appBar: AppBar(
        backgroundColor: Theme.of(context).colorScheme.inversePrimary,
        title: const Text("Excel and word editing app"),
      ),
      body: _isLoading
          ? const Center(
              child: CircularProgressIndicator(),
            )
          : Container(
              height: size.height,
              width: size.width,
              child: Row(
                mainAxisAlignment: MainAxisAlignment.spaceAround,
                children: [
                  Container(
                    width: size.width * 0.45,
                    child: Column(
                      children: [
                        Container(
                          margin: const EdgeInsets.symmetric(
                            vertical: 15,
                          ),
                          child: Form(
                            key: excelForm,
                            child: Column(
                              children: [
                                Container(
                                  width: size.width * 0.45,
                                  constraints: BoxConstraints(
                                      maxHeight: size.height * 0.35),
                                  child: Row(
                                      mainAxisAlignment:
                                          MainAxisAlignment.spaceAround,
                                      children: [
                                        Container(
                                          width: size.width * 0.15,
                                          child: TextFormField(
                                            controller: _name,
                                            maxLength: 15,
                                            validator: (value) {
                                              if (value == null ||
                                                  value.isEmpty) {
                                                return "Debe de registrar un valor adecuado.";
                                              }
                                              return null;
                                            },
                                            decoration: const InputDecoration(
                                                hintText: "Nombre"),
                                          ),
                                        ),
                                        Container(
                                          width: size.width * 0.15,
                                          child: TextFormField(
                                            keyboardType: TextInputType.number,
                                            maxLength: 3,
                                            controller: _age,
                                            validator: (value) {
                                              if (value == null ||
                                                  value.isEmpty) {
                                                return "Debe de registrar un valor adecuado.";
                                              }
                                              if (int.tryParse(value) == null) {
                                                return "El valor debe de ser un numero";
                                              }
                                              return null;
                                            },
                                            decoration: const InputDecoration(
                                                hintText: "Edad"),
                                          ),
                                        ),
                                      ]),
                                ),
                                Container(
                                  margin: const EdgeInsets.symmetric(
                                    vertical: 10,
                                  ),
                                  child: ElevatedButton.icon(
                                      onPressed: insert,
                                      icon: const Icon(Icons.save),
                                      label: const Text(
                                          "Guardar informacion en excel")),
                                ),
                              ],
                            ),
                          ),
                        ),
                        Container(
                          width: size.width,
                          child: Row(
                            mainAxisAlignment: MainAxisAlignment.spaceBetween,
                            children: [
                              const Text(
                                "Data Excel.",
                                textAlign: TextAlign.left,
                                style: TextStyle(fontWeight: FontWeight.bold),
                              ),
                              TextButton.icon(
                                  onPressed: downloadExcelFile,
                                  icon: const Icon(Icons.download),
                                  label: const Text("Descargar data"))
                            ],
                          ),
                        ),
                        Divider(),
                        Flexible(
                          child: ListView.builder(
                            itemBuilder: (context, index) => ListTile(
                              title:
                                  Text("Nombre: ${dataExcel[index][0]?.value}"),
                              subtitle:
                                  Text("Edad: ${dataExcel[index][1]?.value}"),
                              trailing: Row(
                                mainAxisSize: MainAxisSize.min,
                                mainAxisAlignment:
                                    MainAxisAlignment.spaceAround,
                                children: [
                                  IconButton(
                                    onPressed: () => edit(
                                        index + 2,
                                        dataExcel[index][0]?.value,
                                        dataExcel[index][1]?.value),
                                    icon: Icon(Icons.edit),
                                    color: Colors.orangeAccent,
                                  ),
                                  IconButton(
                                    onPressed: () => deleteRow(index + 1),
                                    icon: Icon(Icons.delete),
                                    color: Colors.redAccent,
                                  ),
                                ],
                              ),
                            ),
                            itemCount: dataExcel.length,
                          ),
                        ),
                      ],
                    ),
                  ),
                  Container(
                    width: size.width * 0.45,
                    child: Column(
                      children: [
                        Container(
                          margin: const EdgeInsets.symmetric(
                            vertical: 15,
                          ),
                          child: Form(
                            key: wordForm,
                            child: Column(
                              children: [
                                Container(
                                  width: size.width * 0.45,
                                  constraints: BoxConstraints(
                                      maxHeight: size.height * 0.35),
                                  child: Column(
                                      mainAxisSize: MainAxisSize.min,
                                      children: [
                                        Container(
                                          width: size.width * 0.45,
                                          child: TextFormField(
                                            controller: _client,
                                            maxLength: 65,
                                            validator: (value) {
                                              if (value == null ||
                                                  value.isEmpty) {
                                                return "Debe de registrar un valor adecuado.";
                                              }
                                              return null;
                                            },
                                            decoration: const InputDecoration(
                                                hintText: "Cliente"),
                                          ),
                                        ),
                                        Container(
                                          width: size.width * 0.45,
                                          child: TextFormField(
                                            keyboardType: TextInputType.number,
                                            maxLength: 9,
                                            controller: _amount,
                                            validator: (value) {
                                              if (value == null ||
                                                  value.isEmpty) {
                                                return "Debe de registrar un valor adecuado.";
                                              }
                                              if (double.tryParse(value) ==
                                                  null) {
                                                return "El valor debe de ser un numero";
                                              }
                                              return null;
                                            },
                                            decoration: const InputDecoration(
                                                hintText: "Monto"),
                                          ),
                                        ),
                                      ]),
                                ),
                                Container(
                                  margin: const EdgeInsets.symmetric(
                                    vertical: 10,
                                  ),
                                  child: ElevatedButton.icon(
                                    onPressed: downloadWordFile,
                                    icon: const Icon(Icons.download),
                                    label: const Text("Descargar factura"),
                                  ),
                                ),
                              ],
                            ),
                          ),
                        ),
                      ],
                    ),
                  ),
                ],
              )),
    );
  }
}
