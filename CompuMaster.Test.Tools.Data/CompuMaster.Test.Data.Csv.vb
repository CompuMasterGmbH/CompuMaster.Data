Imports NUnit.Framework

Namespace CompuMaster.Test.Data

    <TestFixture(Category:="CSV")> Public Class Csv

        <Test()> Public Sub ReadDataTableFromCsvStringSeparatorSeparated()

            Dim testinputdata As String
            Dim testoutputdata As DataTable

            Dim CurCulture As System.Globalization.CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture

            'German culture - semi-colon as column separator
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("de-DE")

            testinputdata = "ID;""Description"";DateValue" & vbNewLine & _
                "5;""line1 ü content"";2005-08-29" & vbNewLine & _
                ";""line 2 """" content"";""2005-08-27"""
            testoutputdata = CompuMaster.Data.Csv.ReadDataTableFromCsvString(testinputdata, True)
            'Throw New Exception(CompuMaster.Data.DataTables.ConvertToPlainTextTable(testoutputdata))
            NUnit.Framework.Assert.AreEqual(3, testoutputdata.Columns.Count, "JW #100")
            NUnit.Framework.Assert.AreEqual(2, testoutputdata.Rows.Count, "JW #101")
            NUnit.Framework.Assert.AreEqual("ID", testoutputdata.Columns(0).ColumnName, "JW #102")
            NUnit.Framework.Assert.AreEqual("Description", testoutputdata.Columns(1).ColumnName, "JW #103")
            NUnit.Framework.Assert.AreEqual("DateValue", testoutputdata.Columns(2).ColumnName, "JW #104")
            NUnit.Framework.Assert.AreEqual("5", testoutputdata.Rows(0)(0), "JW #105")
            NUnit.Framework.Assert.AreEqual("line1 ü content", testoutputdata.Rows(0)(1), "JW #106")
            NUnit.Framework.Assert.AreEqual("2005-08-29", testoutputdata.Rows(0)(2), "JW #107")
            NUnit.Framework.Assert.AreEqual("", testoutputdata.Rows(1)(0), "JW #108")
            NUnit.Framework.Assert.AreEqual("line 2 "" content", testoutputdata.Rows(1)(1), "JW #109")
            NUnit.Framework.Assert.AreEqual("2005-08-27", testoutputdata.Rows(1)(2), "JW #110")

            testinputdata = "ID;""Description"";DateValue" & vbNewLine & _
                "5;""line1 ü content"";2005-08-29" & ControlChars.Lf & _
                ";""line 2 " & ControlChars.Lf & "newline content"";""2005-08-27"";" & vbNewLine
            testoutputdata = CompuMaster.Data.Csv.ReadDataTableFromCsvString(testinputdata, True)
            NUnit.Framework.Assert.AreEqual(4, testoutputdata.Columns.Count, "JW #200")
            NUnit.Framework.Assert.AreEqual(2, testoutputdata.Rows.Count, "JW #201")
            NUnit.Framework.Assert.AreEqual("ID", testoutputdata.Columns(0).ColumnName, "JW #202")
            NUnit.Framework.Assert.AreEqual("Description", testoutputdata.Columns(1).ColumnName, "JW #203")
            NUnit.Framework.Assert.AreEqual("DateValue", testoutputdata.Columns(2).ColumnName, "JW #204")
            NUnit.Framework.Assert.AreEqual("5", testoutputdata.Rows(0)(0), "JW #205")
            NUnit.Framework.Assert.AreEqual("line1 ü content", testoutputdata.Rows(0)(1), "JW #206")
            NUnit.Framework.Assert.AreEqual("2005-08-29", testoutputdata.Rows(0)(2), "JW #207")
            NUnit.Framework.Assert.AreEqual("", testoutputdata.Rows(0)(3), "JW #211")
            NUnit.Framework.Assert.AreEqual("", testoutputdata.Rows(1)(0), "JW #208")
            NUnit.Framework.Assert.AreEqual("line 2 " & System.Environment.NewLine & "newline content", testoutputdata.Rows(1)(1), "JW #209")
            NUnit.Framework.Assert.AreEqual("2005-08-27", testoutputdata.Rows(1)(2), "JW #210")
            NUnit.Framework.Assert.AreEqual("", testoutputdata.Rows(1)(3), "JW #211")

            'English culture - comma instead of semi-colon
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-GB")

            testinputdata = "ID,""Description"",DateValue" & vbNewLine & _
                "5,""line1 ü content"",2005-08-29" & vbNewLine & _
                ",""line 2 """" content"",""2005-08-27"""
            testoutputdata = CompuMaster.Data.Csv.ReadDataTableFromCsvString(testinputdata, True)
            'Throw New Exception(CompuMaster.Data.DataTables.ConvertToPlainTextTable(testoutputdata))
            NUnit.Framework.Assert.AreEqual(3, testoutputdata.Columns.Count, "JW #700")
            NUnit.Framework.Assert.AreEqual(2, testoutputdata.Rows.Count, "JW #701")
            NUnit.Framework.Assert.AreEqual("ID", testoutputdata.Columns(0).ColumnName, "JW #702")
            NUnit.Framework.Assert.AreEqual("Description", testoutputdata.Columns(1).ColumnName, "JW #703")
            NUnit.Framework.Assert.AreEqual("DateValue", testoutputdata.Columns(2).ColumnName, "JW #704")
            NUnit.Framework.Assert.AreEqual("5", testoutputdata.Rows(0)(0), "JW #705")
            NUnit.Framework.Assert.AreEqual("line1 ü content", testoutputdata.Rows(0)(1), "JW #706")
            NUnit.Framework.Assert.AreEqual("2005-08-29", testoutputdata.Rows(0)(2), "JW #707")
            NUnit.Framework.Assert.AreEqual("", testoutputdata.Rows(1)(0), "JW #708")
            NUnit.Framework.Assert.AreEqual("line 2 "" content", testoutputdata.Rows(1)(1), "JW #709")
            NUnit.Framework.Assert.AreEqual("2005-08-27", testoutputdata.Rows(1)(2), "JW #710")

            testinputdata = "ID,""Description"",DateValue" & vbNewLine & _
                "5,""line1 ü content"",2005-08-29" & ControlChars.Lf & _
                ",""line 2 " & ControlChars.Lf & "newline content"",""2005-08-27""," & vbNewLine
            testoutputdata = CompuMaster.Data.Csv.ReadDataTableFromCsvString(testinputdata, True)
            NUnit.Framework.Assert.AreEqual(4, testoutputdata.Columns.Count, "JW #800")
            NUnit.Framework.Assert.AreEqual(2, testoutputdata.Rows.Count, "JW #801")
            NUnit.Framework.Assert.AreEqual("ID", testoutputdata.Columns(0).ColumnName, "JW #802")
            NUnit.Framework.Assert.AreEqual("Description", testoutputdata.Columns(1).ColumnName, "JW #803")
            NUnit.Framework.Assert.AreEqual("DateValue", testoutputdata.Columns(2).ColumnName, "JW #804")
            NUnit.Framework.Assert.AreEqual("5", testoutputdata.Rows(0)(0), "JW #805")
            NUnit.Framework.Assert.AreEqual("line1 ü content", testoutputdata.Rows(0)(1), "JW #806")
            NUnit.Framework.Assert.AreEqual("2005-08-29", testoutputdata.Rows(0)(2), "JW #807")
            NUnit.Framework.Assert.AreEqual("", testoutputdata.Rows(0)(3), "JW #811")
            NUnit.Framework.Assert.AreEqual("", testoutputdata.Rows(1)(0), "JW #808")
            NUnit.Framework.Assert.AreEqual("line 2 " & System.Environment.NewLine & "newline content", testoutputdata.Rows(1)(1), "JW #809")
            NUnit.Framework.Assert.AreEqual("2005-08-27", testoutputdata.Rows(1)(2), "JW #810")
            NUnit.Framework.Assert.AreEqual("", testoutputdata.Rows(1)(3), "JW #811")

            'Reset to origin culture
            System.Threading.Thread.CurrentThread.CurrentCulture = CurCulture

        End Sub

        <Test()> Public Sub ReadDataTableFromCsvReaderFixedColumnWidths()
            Dim testinputdata As String
            Dim testoutputdata As DataTable
            Dim columnWidths As Integer()

            testinputdata = "ID     Description   DateValue       " & vbNewLine & _
                            "12345678901234567890123456789012" & vbNewLine & _
                            "1234567890123456789012345678901234567890"
            columnWidths = New Integer() {7, 12, 15}
            testoutputdata = CompuMaster.Data.Csv.ReadDataTableFromCsvString(testinputdata, True, columnWidths)
            'Throw New Exception(CompuMaster.Data.DataTables.ConvertToPlainTextTable(testoutputdata))
            NUnit.Framework.Assert.AreEqual(4, testoutputdata.Columns.Count, "JW #300")
            NUnit.Framework.Assert.AreEqual(2, testoutputdata.Rows.Count, "JW #301")
            NUnit.Framework.Assert.AreEqual("ID", testoutputdata.Columns(0).ColumnName, "JW #302")
            NUnit.Framework.Assert.AreEqual("Description", testoutputdata.Columns(1).ColumnName, "JW #303")
            NUnit.Framework.Assert.AreEqual("DateValue", testoutputdata.Columns(2).ColumnName, "JW #304")
            NUnit.Framework.Assert.AreEqual("Column1", testoutputdata.Columns(3).ColumnName, "JW #304a")
            NUnit.Framework.Assert.AreEqual("1234567", testoutputdata.Rows(0)(0), "JW #305")
            NUnit.Framework.Assert.AreEqual("890123456789", testoutputdata.Rows(0)(1), "JW #306")
            NUnit.Framework.Assert.AreEqual("0123456789012", testoutputdata.Rows(0)(2), "JW #307")
            NUnit.Framework.Assert.AreEqual("1234567", testoutputdata.Rows(1)(0), "JW #308")
            NUnit.Framework.Assert.AreEqual("890123456789", testoutputdata.Rows(1)(1), "JW #309")
            NUnit.Framework.Assert.AreEqual("012345678901234", testoutputdata.Rows(1)(2), "JW #310")
            NUnit.Framework.Assert.AreEqual("567890", testoutputdata.Rows(1)(3), "JW #311")

            'testinputdata = "ID;""Description"";DateValue" & vbNewLine & _
            '    "5;""line1 ü content"";2005-08-29" & vbNewLine & _
            '    ";""line 2 " & ControlChars.Lf & "newline content"";""2005-08-27"";" & vbNewLine
            'testoutputdata = CompuMaster.Data.Csv.ReadDataTableFromCsvString(testinputdata, True)
            'NUnit.Framework.Assert.AreEqual(4, testoutputdata.Columns.Count, "JW #400")
            'NUnit.Framework.Assert.AreEqual(2, testoutputdata.Rows.Count, "JW #401")
            'NUnit.Framework.Assert.AreEqual("ID", testoutputdata.Columns(0).ColumnName, "JW #402")
            'NUnit.Framework.Assert.AreEqual("Description", testoutputdata.Columns(1).ColumnName, "JW #403")
            'NUnit.Framework.Assert.AreEqual("DateValue", testoutputdata.Columns(2).ColumnName, "JW #404")
            'NUnit.Framework.Assert.AreEqual("5", testoutputdata.Rows(0)(0), "JW #405")
            'NUnit.Framework.Assert.AreEqual("line1 ü content", testoutputdata.Rows(0)(1), "JW #406")
            'NUnit.Framework.Assert.AreEqual("2005-08-29", testoutputdata.Rows(0)(2), "JW #407")
            'NUnit.Framework.Assert.AreEqual("", testoutputdata.Rows(0)(3), "JW #411")
            'NUnit.Framework.Assert.AreEqual("", testoutputdata.Rows(1)(0), "JW #408")
            'NUnit.Framework.Assert.AreEqual("line 2 " & System.Environment.NewLine & "newline content", testoutputdata.Rows(1)(1), "JW #409")
            'NUnit.Framework.Assert.AreEqual("2005-08-27", testoutputdata.Rows(1)(2), "JW #410")
            'NUnit.Framework.Assert.AreEqual("", testoutputdata.Rows(1)(3), "JW #411")

        End Sub

    End Class

End Namespace