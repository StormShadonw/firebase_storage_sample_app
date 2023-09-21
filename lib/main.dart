import 'package:dio/dio.dart';
import 'package:file_saver/file_saver.dart';
import 'package:flutter/material.dart';
import 'package:intl/intl.dart';

void main() async {
  WidgetsFlutterBinding.ensureInitialized();
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
  final dio = Dio();
  static const APISERVER = "https://riascoswebapi.azurewebsites.net";
  final _fileName = "Firebase_storage_sample_archive.xlsx";
  final queryParameters = "documentId=Data.xlsx";
  final numberFormat = NumberFormat("#,##0.00", "en_US");
  bool _isLoading = false;
  final excelForm = GlobalKey<FormState>();
  final wordForm = GlobalKey<FormState>();
  int indexToEdit = 0;
  List<dynamic> dataExcel = [];
  final TextEditingController _name = TextEditingController();
  final TextEditingController _age = TextEditingController();
  final TextEditingController _client = TextEditingController();
  final TextEditingController _amount = TextEditingController();

  void edit(int index, String name, int age) {
    indexToEdit = index;
    _name.value = TextEditingValue(text: name);
    _age.value = TextEditingValue(text: age.toString());
  }

  Future<void> deleteRow(int index) async {
    await dio.delete('$APISERVER/Excel/$index?$queryParameters');
    await getExcelData();
  }

  Future<void> downloadWordFile() async {
    setState(() {
      _isLoading = true;
    });
    try {
      if (wordForm.currentState!.validate()) {
        var dio = Dio();
        var response = await dio.post(
          '$APISERVER/Word?documentId=Document.docx',
          options: Options(responseType: ResponseType.bytes),
          data: {
            "customer": _client.value.text,
            "supplier": "Proveedor de prueba",
            "amount": double.parse(_amount.value.text)
          },
        );
        final responseBody = response.data;
        FileSaver.instance.saveFile(name: "Factura.docx", bytes: responseBody);
        _client.value = TextEditingValue.empty;
        _amount.value = TextEditingValue.empty;
      }
    } catch (error) {}

    setState(() {
      _isLoading = false;
    });
  }

  Future<void> downloadExcelFile() async {
    setState(() {
      _isLoading = true;
    });
    var dio = Dio();
    var response = await dio.get(
      '$APISERVER/Excel/download-file?$queryParameters',
      options: Options(responseType: ResponseType.bytes),
    );
    final responseBody = response.data;
    FileSaver.instance.saveFile(name: _fileName, bytes: responseBody);

    setState(() {
      _isLoading = false;
    });
  }

  Future<void> insert() async {
    setState(() {
      _isLoading = true;
    });
    try {
      if (excelForm.currentState!.validate()) {
        if (indexToEdit == 0) {
          var response = await dio.post('$APISERVER/Excel?$queryParameters',
              data: {
                'name': _name.value.text,
                'age': int.parse(_age.value.text)
              });
          if (response.statusCode == 200) {
            await getExcelData();
          }
        } else {
          var response = await dio
              .put('$APISERVER/Excel/$indexToEdit?$queryParameters', data: {
            'name': _name.value.text,
            'age': int.parse(_age.value.text)
          });
          if (response.statusCode == 200) {
            await getExcelData();
          }
        }
      }
    } catch (error) {
      print("Error escribiendo data: $error");
    }

    setState(() {
      _isLoading = false;
    });
  }

  Future<void> getExcelData() async {
    setState(() {
      _isLoading = true;
    });
    try {
      indexToEdit = 0;
      _name.value = TextEditingValue.empty;
      _age.value = TextEditingValue.empty;
      var response = await dio
          .get(
        '$APISERVER/Excel?$queryParameters',
      )
          .onError((error, stackTrace) async {
        return await showDialog(
            context: context,
            builder: (_) => AlertDialog(
                  title: const Text("Error obteniendo data del excel"),
                  content: Text(error.toString()),
                ));
      }).catchError((error) async {
        return await showDialog(
            context: context,
            builder: (_) => AlertDialog(
                  title: const Text("Error obteniendo data del excel"),
                  content: Text(error.toString()),
                ));
      });
      var data = response.data as List<dynamic>;
      data.removeAt(0);
      if (response.statusCode == 200) {
        setState(() {
          dataExcel = data;
        });
      } else {
        print("Error: ${response.statusMessage}");
      }
    } catch (e) {
      print("Error downloaded file: $e");
    }

    setState(() {
      _isLoading = false;
    });
  }

  Future<void> getInitData() async {
    await getExcelData();
  }

  @override
  void initState() {
    super.initState();
    getInitData();
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
                              title: Text("Nombre: ${dataExcel[index][0]}"),
                              subtitle: Text("Edad: ${dataExcel[index][1]}"),
                              trailing: Row(
                                mainAxisSize: MainAxisSize.min,
                                mainAxisAlignment:
                                    MainAxisAlignment.spaceAround,
                                children: [
                                  IconButton(
                                    onPressed: () => edit(
                                        index + 2,
                                        dataExcel[index][0],
                                        dataExcel[index][1]),
                                    icon: Icon(Icons.edit),
                                    color: Colors.orangeAccent,
                                  ),
                                  IconButton(
                                    onPressed: () => deleteRow(index + 2),
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
