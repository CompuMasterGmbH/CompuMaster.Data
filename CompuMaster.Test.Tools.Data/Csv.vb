Imports NUnit.Framework

Namespace CompuMaster.Test.Data

    <TestFixture(Category:="CSV")> Public Class Csv
        Public Sub New()
        End Sub

        Private _OriginCulture As System.Globalization.CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture
        <TearDown> Public Sub ResetCulture()
            System.Threading.Thread.CurrentThread.CurrentCulture = _OriginCulture
        End Sub

        <Test()> Public Sub ReadDataTableFromCsvStringSeparatorSeparatedMustFailsBecauseOfWrongCulture(<Values("en-US", "en-GB", "ja-JP")> cultureContext As String)
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture(cultureContext)

            Dim testinputdata As String
            Dim testoutputdata As DataTable

            testinputdata = "ID;""Description"";DateValue" & vbNewLine &
                "5;""line1 ü content"";2005-08-29" & vbNewLine &
                ";""line 2 """" content"";""2005-08-27"""
            testoutputdata = CompuMaster.Data.Csv.ReadDataTableFromCsvString(testinputdata, True, System.Threading.Thread.CurrentThread.CurrentCulture)
            'Throw New Exception(CompuMaster.Data.DataTables.ConvertToPlainTextTable(testoutputdata))
            NUnit.Framework.Assert.AreNotEqual(3, testoutputdata.Columns.Count, "JW #100")
            NUnit.Framework.Assert.AreEqual(2, testoutputdata.Rows.Count, "JW #101")
        End Sub

        <Test> Public Sub ReadDataTableFromCsvFileWithColumnSeparatorCharInTextStrings()
            Dim TestFile As String = AssemblyTestEnvironment.TestFileAbsolutePath("testfiles\country-codes.csv")
            'TestFile = "https://raw.githubusercontent.com/datasets/country-codes/master/data/country-codes.csv"

            Dim CountryCodesTable As DataTable = CompuMaster.Data.Csv.ReadDataTableFromCsvFile(TestFile, True, System.Text.Encoding.UTF8, System.Globalization.CultureInfo.InvariantCulture, """"c, False, False)
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToWikiTable(CountryCodesTable))

            NUnit.Framework.Assert.AreEqual(3, CountryCodesTable.Rows.Count)
            Dim ColumnHeaders As String() = New String() {"name", "official_name_en", "official_name_fr", "ISO3166-1-Alpha-2", "ISO3166-1-Alpha-3", "ISO3166-1-numeric", "ITU", "MARC", "WMO", "DS", "Dial", "FIFA", "FIPS", "GAUL", "IOC", "ISO4217-currency_alphabetic_code", "ISO4217-currency_country_name", "ISO4217-currency_minor_unit", "ISO4217-currency_name", "ISO4217-currency_numeric_code", "is_independent", "Capital", "Continent", "TLD", "Languages", "geonameid", "EDGAR"}
            For MyCounter As Integer = 0 To System.Math.Min(CountryCodesTable.Columns.Count, ColumnHeaders.Length) - 1
                NUnit.Framework.Assert.AreEqual(ColumnHeaders(MyCounter), CountryCodesTable.Columns(MyCounter).ColumnName)
            Next
            NUnit.Framework.Assert.AreEqual(ColumnHeaders.Length, CountryCodesTable.Columns.Count)
        End Sub

        <Test()> Public Sub ReadDataTableFromCsvStringSeparatorSeparated(<Values("en-US", "en-GB", "de-DE", "fr-FR", "ja-JP")> cultureContext As String)
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture(cultureContext)

            Dim testinputdata As String
            Dim testoutputdata As DataTable

            testinputdata = "ID;""Description"";DateValue" & vbNewLine &
                "5;""line1 ü content"";2005-08-29" & vbNewLine &
                ";""line 2 """" content"";""2005-08-27"""
            testoutputdata = CompuMaster.Data.Csv.ReadDataTableFromCsvString(testinputdata, True, System.Globalization.CultureInfo.CreateSpecificCulture("de-DE"))
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

            testinputdata = "ID;""Description"";DateValue" & vbNewLine &
                "5;""line1 ü content"";2005-08-29" & ControlChars.Lf &
                ";""line 2 " & ControlChars.Lf & "newline content"";""2005-08-27"";" & vbNewLine
            testoutputdata = CompuMaster.Data.Csv.ReadDataTableFromCsvString(testinputdata, True, System.Globalization.CultureInfo.CreateSpecificCulture("de-DE"))
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

            testinputdata = "ID,""Description"",DateValue" & vbNewLine &
                "5,""line1 ü content"",2005-08-29" & vbNewLine &
                ",""line 2 """" content"",""2005-08-27"""
            testoutputdata = CompuMaster.Data.Csv.ReadDataTableFromCsvString(testinputdata, True, System.Globalization.CultureInfo.CreateSpecificCulture("en-GB"))
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

            testinputdata = "ID,""Description"",DateValue" & vbNewLine &
                "5,""line1 ü content"",2005-08-29" & ControlChars.Lf &
                ",""line 2 " & ControlChars.Lf & "newline content"",""2005-08-27""," & vbNewLine
            testoutputdata = CompuMaster.Data.Csv.ReadDataTableFromCsvString(testinputdata, True, System.Globalization.CultureInfo.CreateSpecificCulture("en-GB"))
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

            'Again the same tests from above - but this time without culture parameter but some manual info on separator etc.

            testinputdata = "ID;""Description"";DateValue" & vbNewLine &
                "5;""line1 ü content"";2005-08-29" & vbNewLine &
                ";""line 2 """" content"";""2005-08-27"""
            testoutputdata = CompuMaster.Data.Csv.ReadDataTableFromCsvString(testinputdata, True, ";"c, """"c, False, False)
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

            testinputdata = "ID;""Description"";DateValue" & vbNewLine &
                "5;""line1 ü content"";2005-08-29" & ControlChars.Lf &
                ";""line 2 " & ControlChars.Lf & "newline content"";""2005-08-27"";" & vbNewLine
            testoutputdata = CompuMaster.Data.Csv.ReadDataTableFromCsvString(testinputdata, True, ";"c, """"c, False, False)
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

            testinputdata = "ID,""Description"",DateValue" & vbNewLine &
                "5,""line1 ü content"",2005-08-29" & vbNewLine &
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

            testinputdata = "ID,""Description"",DateValue" & vbNewLine &
                "5,""line1 ü content"",2005-08-29" & ControlChars.Lf &
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

        End Sub

        <Test()> Public Sub ReadDataTableFromCsvReaderFixedColumnWidths()
            Dim testinputdata As String
            Dim testoutputdata As DataTable
            Dim columnWidths As Integer()

            testinputdata = "ID     Description   DateValue       " & vbNewLine &
                            "12345678901234567890123456789012" & vbNewLine &
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

        <Test> Sub WriteDataTableToCsvFileStringWithTextEncoding()
            Dim t As DataTable = SimpleSampleTable()
            Dim bom As String = System.Text.Encoding.UTF8.GetString(System.Text.Encoding.UTF8.GetPreamble())
            Dim csv As String
            csv = CompuMaster.Data.Csv.WriteDataTableToCsvFileStringWithTextEncoding(t, True)
            Assert.True(csv.Substring(0, bom.Length) = bom, "CSV starts correctly with BOM signature for UTF-8")
            csv = CompuMaster.Data.Csv.WriteDataTableToCsvFileStringWithTextEncoding(t, True, "UTF-8")
            Assert.True(csv.Substring(0, bom.Length) = bom, "CSV starts correctly with BOM signature for UTF-8")
        End Sub

        <Test> Sub WriteDataTableToCsvTextString()
            Dim t As DataTable = SimpleSampleTable()
            Dim bom As String = System.Text.Encoding.UTF8.GetString(System.Text.Encoding.UTF8.GetPreamble())
            Dim csv As String
            csv = CompuMaster.Data.Csv.WriteDataTableToCsvTextString(t, True)
            Assert.False(csv.Substring(0, bom.Length) = bom, "CSV starts invalidly with BOM signature for UTF-8")
        End Sub

        Private Function SimpleSampleTable() As DataTable
            Dim t As New DataTable("root")
            t.Columns.Add("col1")
            t.Columns.Add("col2")
            Dim r As DataRow = t.NewRow
            r(0) = "R1C1"
            r(1) = "R1C2"
            t.Rows.Add(r)
            r = t.NewRow
            r(0) = "R2C1"
            r(1) = "R2C2"
            t.Rows.Add(r)
            Return t
        End Function

        Private Function CellWithLineBreaksAndSpecialCharsSampleTable() As DataTable
            Dim t As New DataTable("root")
            t.Columns.Add("col1")
            t.Columns.Add("col2")
            Dim r As DataRow = t.NewRow
            r(0) = "R1C1""" & vbCrLf
            r(1) = "R1C2""" & vbCr
            t.Rows.Add(r)
            r = t.NewRow
            r(0) = "R2C1""" & vbLf
            r(1) = "R2C2""" & vbTab
            t.Rows.Add(r)
            Return t
        End Function

        Private Function XxlFactorySampleTableMightCausingOutOfMemoryExceptionWhenInStringInComplete(linesInMillion As Double) As DataTable
            Dim LinesInTotal As Integer = linesInMillion * 1000 ^ 2 'x Mio.
            Dim t As New DataTable("root")
            t.Columns.Add("col1")
            t.Columns.Add("col2")
            For MyCounter As Integer = 1 To LinesInTotal
                Dim r As DataRow = t.NewRow
                r(0) = "Line no. " & MyCounter.ToString("000,000,000,000")
                r(1) = "abcdefghijklmnopqrstuvwxyzäöüß|abcdefghijklmnopqrstuvwxyzäöüß|abcdefghijklmnopqrstuvwxyzäöüß|abcdefghijklmnopqrstuvwxyzäöüß|abcdefghijklmnopqrstuvwxyzäöüß|abcdefghijklmnopqrstuvwxyzäöüß|abcdefghijklmnopqrstuvwxyzäöüß|abcdefghijklmnopqrstuvwxyzäöüß"
                t.Rows.Add(r)
                If MyCounter Mod 1000 ^ 2 = 0 Then
                    Console.WriteLine("Created records in test table: " & MyCounter.ToString("#,##0"))
                End If
            Next
            Return t
        End Function

        <Test> Sub WriteXlDataTableToCsvTextStringAndReRead()
            WriteFactoryXxlDataTableToCsvTextStringAndReRead(XxlFactorySampleTableMightCausingOutOfMemoryExceptionWhenInStringInComplete(0.01)) '10,000 lines
        End Sub

        <Test, Ignore("ReadXxl fails and needs additional work")> Sub WriteXxlDataTableToCsvTextStringAndReRead()
            WriteFactoryXxlDataTableToCsvTextStringAndReRead(XxlFactorySampleTableMightCausingOutOfMemoryExceptionWhenInStringInComplete(1)) '1 mio. lines
        End Sub

        <Test, Ignore("WriteXxl already fails")> Sub WriteXxxlDataTableToCsvTextStringAndReRead()
            WriteFactoryXxlDataTableToCsvTextStringAndReRead(XxlFactorySampleTableMightCausingOutOfMemoryExceptionWhenInStringInComplete(500)) '500 mio. lines
        End Sub

        Sub WriteFactoryXxlDataTableToCsvTextStringAndReRead(t As DataTable)
            Dim TempFile As New CompuMaster.Test.Data.TemporaryFile(".csv")
            Console.WriteLine("Using temporary file: " & TempFile.Path)

            'Write to disk
            Dim Start As DateTime = Now
            CompuMaster.Data.Csv.WriteDataTableToCsvFile(TempFile.Path, t)
            Console.WriteLine("CSV written to disk successfully within " & Now.Subtract(Start).TotalSeconds.ToString("#,##0.0") & " sec. with " & t.Rows.Count.ToString("#,##0") & " records and " & (TempFile.FileSize / 1024 / 1024).ToString("#,##0.00") & " MB size")

            'Read from disk
            Start = Now
            Dim t2 As DataTable = CompuMaster.Data.Csv.ReadDataTableFromCsvFile(TempFile.Path, True)
            Console.WriteLine("CSV read from disk successfully within " & Now.Subtract(Start).TotalSeconds.ToString("#,##0.0") & " sec. with " & t2.Rows.Count.ToString("#,##0") & " records")

            'Basic comparisons
            Assert.AreEqual(t.Rows.Count, t2.Rows.Count) 'should be the very same
        End Sub

        <Test> Sub WriteDataTableToCsvTextStringAndReReadAndReWriteWithoutChanges()
            Dim t As DataTable = SimpleSampleTable()
            Dim bom As String = System.Text.Encoding.UTF8.GetString(System.Text.Encoding.UTF8.GetPreamble())
            Dim csv As String
            csv = CompuMaster.Data.Csv.WriteDataTableToCsvTextString(t, True)
            Dim t2 As DataTable = CompuMaster.Data.Csv.ReadDataTableFromCsvString(csv, True)
            Dim csv2 As String
            Assert.AreEqual(t.Columns.Count, t2.Columns.Count) 'should be the very same
            Assert.AreEqual(t.Rows.Count, t2.Rows.Count) 'should be the very same
            csv2 = CompuMaster.Data.Csv.WriteDataTableToCsvTextString(t2, True)
            Assert.AreEqual(csv, csv2) 'should be the very same
        End Sub

        <Test> Sub WriteDataTableToCsvTextStringAndReReadAndReWriteWithoutChangesWithLineBreaksExcelCompatible()
            Dim t As DataTable = CellWithLineBreaksAndSpecialCharsSampleTable()
            Dim bom As String = System.Text.Encoding.UTF8.GetString(System.Text.Encoding.UTF8.GetPreamble())
            Dim csv As String
            csv = CompuMaster.Data.Csv.WriteDataTableToCsvTextString(t, True)
            Dim t2 As DataTable = CompuMaster.Data.Csv.ReadDataTableFromCsvString(csv, True)
            Dim csv2 As String
            Assert.AreEqual(t.Columns.Count, t2.Columns.Count) 'should be the very same
            Assert.AreEqual(t.Rows.Count, t2.Rows.Count) 'should be the very same - except line breaks in cells haven't been converted correctly
            csv2 = CompuMaster.Data.Csv.WriteDataTableToCsvTextString(t2, True)
            Assert.AreEqual(csv, csv2) 'should be the very same
        End Sub

    End Class

End Namespace