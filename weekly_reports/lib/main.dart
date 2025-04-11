import 'package:flutter/material.dart';
import 'package:weekly_reports/landingpage.dart';

class MyApp extends StatelessWidget {
  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      // title: 'QR code app',
      debugShowCheckedModeBanner: false,
      theme: ThemeData(
        scaffoldBackgroundColor: Color.fromARGB(255, 250, 236, 255),
        useMaterial3: true,
        colorScheme: ColorScheme.fromSeed(
          seedColor: const Color.fromARGB(255, 164, 84, 255),
        ),
      ),
      home: Landingpage(),
    );
  }
}

void main() {
  runApp(MyApp());
}
