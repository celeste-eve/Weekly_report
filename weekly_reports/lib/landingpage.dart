import 'package:flutter/material.dart';
import 'package:weekly_reports/reportgenerator.dart';

class Landingpage extends StatefulWidget {
  const Landingpage({super.key});

  @override
  State<Landingpage> createState() => _LandingPageState();
}

class _LandingPageState extends State<Landingpage> {
  // this function is called when a button is pressed
  void pressme(Widget page) {
    Navigator.push(context, MaterialPageRoute(builder: (context) => page));
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: const Text('Weekly reports generator'),
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
                    pressme(Generatorpage());
                  },
                  child: const Text(
                    'Create project weekly reports',
                    style: TextStyle(
                      fontSize: 20,
                      color: Color.fromARGB(255, 53, 14, 59),
                    ),
                  ),
                ),
              ),
            ),
          ],
        ),
      ),
    );
  }
}
