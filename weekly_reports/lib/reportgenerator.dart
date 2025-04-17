import 'package:flutter/material.dart';
import 'package:file_picker/file_picker.dart';
import 'dart:io';
import 'package:path/path.dart' as path;
import 'package:process_run/shell.dart';

class Generatorpage extends StatefulWidget {
  const Generatorpage({Key? key}) : super(key: key);

  @override
  State<Generatorpage> createState() => _GeneratorpageState();
}

class _GeneratorpageState extends State<Generatorpage> {
  bool isLoading = false;
  String? selectedFilePath;
  String? outputDirectoryPath;

  // upload button function
  void UploadButton() async {
    FilePickerResult? result = await FilePicker.platform.pickFiles(
      allowMultiple: false,
      type: FileType.custom,
      allowedExtensions: ['xlsx', 'csv'],
    );

    if (result != null) {
      setState(() {
        selectedFilePath = result.files.single.path;
      });
    }
  }

  // output button function
  outputButton() async {
    String? selectedDirectory = await FilePicker.platform.getDirectoryPath();
    if (selectedDirectory != null) {
      setState(() {
        outputDirectoryPath = selectedDirectory;
      });
    }
  }

  // run conversion script
  Future<void> runPythonScripA() async {
    try {
      final shell = Shell();

      // Run the conversion script
      await shell.run('''
      python "${Directory.current.path}/assets/reportgenerator.py" "$selectedFilePath" "$outputDirectoryPath"
    ''');

      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(
          content: Text('Report created successfully'),
          duration: Duration(seconds: 100),
        ),
      );
    } catch (e) {
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(
          content: Text('Error running Python script: ${e.toString()}'),
          duration: Duration(seconds: 100),
        ),
      );
    }
  }

  //run presentation script
  Future<void> runPythonScripB() async {
    try {
      final shell = Shell();

      // Run the conversion script
      await shell.run('''
      python "${Directory.current.path}/assets/graphgenerator.py" "$selectedFilePath" "$outputDirectoryPath"
    ''');

      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(
          content: Text('Report created successfully'),
          duration: Duration(seconds: 100),
        ),
      );
    } catch (e) {
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(
          content: Text('Error running Python script: ${e.toString()}'),
          duration: Duration(seconds: 100),
        ),
      );
    }
  }

  Future<void> generateQRCodePdf() async {
    if (selectedFilePath == null) {
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(
          content: Text('Please select a file first'),
          duration: Duration(seconds: 100),
        ),
      );
      return;
    }

    setState(() {
      isLoading = true;
    });
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: const Text('Weekly report generator'),
        backgroundColor: const Color.fromARGB(255, 232, 216, 252),
      ),
      body: Center(
        child: Column(
          children: [
            Padding(
              padding: const EdgeInsets.all(15.0),
              child: SizedBox(
                width: 1000,
                height: 100,
                child: ElevatedButton(
                  onPressed: () {
                    UploadButton();
                  },
                  child: Text(
                    "Upload Running totals 2025",
                    style: TextStyle(fontSize: 20),
                  ),
                ),
              ),
            ),

            // shows file has been uploaded
            Visibility(
              visible: selectedFilePath != null,
              child: Text(
                'Selected file: ${selectedFilePath != null ? path.basename(selectedFilePath!) : ""}',
                style: TextStyle(fontSize: 15),
              ),
            ),

            Padding(
              padding: const EdgeInsets.all(15.0),
              child: SizedBox(
                width: 1000,
                height: 100,
                child: ElevatedButton(
                  onPressed: () {
                    outputButton();
                  },
                  child: Text(
                    "Select output folder",
                    style: TextStyle(fontSize: 20),
                  ),
                ),
              ),
            ),

            // shows output directory has been selected
            Visibility(
              visible: outputDirectoryPath != null,
              child: Text(
                'Selected directory: ${outputDirectoryPath != null ? path.basename(outputDirectoryPath!) : ""}',
                style: TextStyle(fontSize: 15),
              ),
            ),

            Row(
              mainAxisAlignment: MainAxisAlignment.center,
              children: [
                Padding(
                  padding: const EdgeInsets.all(15.0),
                  child: SizedBox(
                    width: 500,
                    height: 100,
                    child: ElevatedButton(
                      onPressed: () => runPythonScripA(),
                      child: Text(
                        "Create weekly reports",
                        style: TextStyle(fontSize: 20),
                      ),
                    ),
                  ),
                ),
                Padding(
                  padding: const EdgeInsets.all(15.0),
                  child: SizedBox(
                    width: 500,
                    height: 100,
                    child: ElevatedButton(
                      onPressed: () => runPythonScripB(),
                      child: Text(
                        "Create weekly presentation charts",
                        style: TextStyle(fontSize: 20),
                      ),
                    ),
                  ),
                ),
              ],
            ),
          ],
        ),
      ),
    );
  }
}
