// This is a basic Flutter widget test.
//
// To perform an interaction with a widget in your test, use the WidgetTester
// utility in the flutter_test package. For example, you can send tap and scroll
// gestures. You can also use WidgetTester to find child widgets in the widget
// tree, read text, and verify that the values of widget properties are correct.

import 'package:flutter_test/flutter_test.dart';
import 'package:sqflite_common_ffi/sqflite_ffi.dart';
import 'package:app_thu_cuoc/main.dart';

void main() {
  testWidgets('thu cuoc screen renders', (WidgetTester tester) async {
    sqfliteFfiInit();
    databaseFactory = databaseFactoryFfi;

    await tester.pumpWidget(const ThuCuocApp());
    await tester.pumpAndSettle();

    expect(find.text('Thu Cước (0/0)'), findsOneWidget);
    expect(find.text('Tiền mặt cầm'), findsOneWidget);
  });
}
