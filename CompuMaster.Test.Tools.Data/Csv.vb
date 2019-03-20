Imports NUnit.Framework

Namespace CompuMaster.Test.Data

    <TestFixture(Category:="CSV")> Public Class Csv
        Public Sub New()
        End Sub

        Private _OriginCulture As System.Globalization.CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture
        <TearDown> Public Sub ResetCulture()
            System.Threading.Thread.CurrentThread.CurrentCulture = _OriginCulture
        End Sub

        <Test> Public Sub ReadDataTableFromCsvUrlWithTls12Required()
            Dim Url As String = "https://data.cityofnewyork.us/api/views/kku6-nxdu/rows.csv?accessType=DOWNLOAD"
            Dim CsvCulture As System.Globalization.CultureInfo = System.Globalization.CultureInfo.CreateSpecificCulture("en-US")
            Dim FileEncoding As System.Text.Encoding = Nothing
            Dim dt As DataTable = CompuMaster.Data.Csv.ReadDataTableFromCsvFile(Url, True, FileEncoding, CsvCulture, """"c, False, True)
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
            Assert.Greater(dt.Columns.Count, 0)
            Assert.Greater(dt.Rows.Count, 0)
        End Sub

        ''' <summary>
        ''' Test from a mini-webserver providing a CSV download with missing response header content-type/charset 
        ''' </summary>
        ''' <remarks>
        ''' The CSV file is returned as UTF-8 bytes
        ''' </remarks>
        <Test> Public Sub ReadDataTableFromCsvUrlAtLocalhostWithContentTypeButWithoutCharset(<Values(Nothing, "text/csv", "text/csv; charset=utf-8")> headerContentType As String)
            Dim Url As String = "http://localhost:8035/"
            Dim CsvCulture As System.Globalization.CultureInfo = System.Globalization.CultureInfo.CreateSpecificCulture("en-US")
            Dim FileEncoding As System.Text.Encoding = Nothing
            Dim Headers As New System.Collections.Specialized.NameValueCollection
            If headerContentType <> Nothing Then
                Headers("content-type") = headerContentType
            End If
            Dim ws As New CompuMaster.Test.Tools.TinyWebServerAdvanced.WebServer(AddressOf ReadDataTableLocalhostTestWebserver, Headers, Url)
            Try
                ws.Run()
                Dim dt As DataTable
                'Test 1
                dt = CompuMaster.Data.Csv.ReadDataTableFromCsvFile(Url, True, FileEncoding, CsvCulture, """"c, False, True)
                Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
                Assert.Greater(dt.Columns.Count, 0)
                Assert.Greater(dt.Rows.Count, 0)
                'Test 2
                dt = CompuMaster.Data.Csv.ReadDataTableFromCsvFile(Url, True, "", ","c, """"c, False, True)
                Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
                Assert.Greater(dt.Columns.Count, 0)
                Assert.Greater(dt.Rows.Count, 0)
                'Test 3
                dt = CompuMaster.Data.Csv.ReadDataTableFromCsvFile(Url, True, CType(Nothing, String), ","c, """"c, False, True)
                Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
                Assert.Greater(dt.Columns.Count, 0)
                Assert.Greater(dt.Rows.Count, 0)
                'Test 4
                dt = CompuMaster.Data.Csv.ReadDataTableFromCsvFile(Url, True, CompuMaster.Data.Csv.ReadLineEncodings.RowBreakCrLfOrCr_CellLineBreakLf, CompuMaster.Data.Csv.ReadLineEncodingAutoConversion.NoAutoConversion, "", ","c, """"c, False, False)
                Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
                Assert.Greater(dt.Columns.Count, 0)
                Assert.Greater(dt.Rows.Count, 0)
            Finally
                ws.Stop()
            End Try
        End Sub

        ''' <summary>
        ''' A test CSV file with unicode characters
        ''' </summary>
        ''' <param name="handler"></param>
        ''' <returns></returns>
        Private Shared Function ReadDataTableLocalhostTestWebserver(handler As System.Net.HttpListenerRequest) As String
            Return "Test,Column" & vbNewLine & "1,äöüßÄÖÜ2"
        End Function

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

        <Test> Public Sub ReadDataTableFromCsvFileViaHttpRequestWithCorrectCharsetEncoding(<Values(1, 2, 3)> testType As Byte)
            Const GithubCountryCodesTestUrl As String = "https://raw.githubusercontent.com/datasets/country-codes/master/data/country-codes.csv"
            Dim CheckEntries As String() = New String() {"CHN", "RUS", "FRA", "ZWE"} 'ISO3166-1-Alpha-3

            Dim CsvTableFromUrl As DataTable, CsvStringTableFromUrl As String

            Select Case testType
                Case 1
                    Console.WriteLine("test of column-separator method type with text encoding """" meaning autodetect")
                    CsvTableFromUrl = CompuMaster.Data.Csv.ReadDataTableFromCsvFile(GithubCountryCodesTestUrl, True, "", ","c, """"c, False, False)
                    CompuMaster.Data.DataTables.RemoveRowsWithWithoutRequiredValuesInColumn(CsvTableFromUrl.Columns("ISO3166-1-Alpha-3"), CheckEntries)
                    Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(CsvTableFromUrl))
                    Assert.AreEqual(CheckEntries.Length, CsvTableFromUrl.Rows.Count)
                    CsvStringTableFromUrl = CompuMaster.Data.DataTables.ConvertToPlainTextTable(CsvTableFromUrl)
                    Assert.IsTrue(CsvStringTableFromUrl.Contains("Russian Federation"))
                    Assert.IsTrue(CsvStringTableFromUrl.Contains("俄罗斯联邦"))
                    Assert.IsTrue(CsvStringTableFromUrl.Contains("الاتحاد الروسي"))

                Case 2
                    Console.WriteLine()
                    Console.WriteLine("test of fixed column method type")
                    CsvTableFromUrl = CompuMaster.Data.Csv.ReadDataTableFromCsvFile(GithubCountryCodesTestUrl, True, New Integer() {}, CType(Nothing, System.Text.Encoding), System.Globalization.CultureInfo.GetCultureInfo("en-US"), False)
                    CsvStringTableFromUrl = CompuMaster.Data.DataTables.ConvertToPlainTextTable(CsvTableFromUrl)
                    Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(CsvTableFromUrl))
                    Assert.IsTrue(CsvStringTableFromUrl.Contains("Russian Federation"))
                    Assert.IsTrue(CsvStringTableFromUrl.Contains("俄罗斯联邦"))
                    Assert.IsTrue(CsvStringTableFromUrl.Contains("الاتحاد الروسي"))

                Case 3
                    Console.WriteLine()
                    Console.WriteLine("test of column-separator method type with text encoding Nothing/null meaning autodetect")
                    CsvTableFromUrl = CompuMaster.Data.Csv.ReadDataTableFromCsvFile(GithubCountryCodesTestUrl, True, CType(Nothing, System.Text.Encoding), System.Globalization.CultureInfo.GetCultureInfo("en-US"), """"c, False, False)
                    CompuMaster.Data.DataTables.RemoveRowsWithWithoutRequiredValuesInColumn(CsvTableFromUrl.Columns("ISO3166-1-Alpha-3"), CheckEntries)
                    Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(CsvTableFromUrl))
                    Assert.AreEqual(CheckEntries.Length, CsvTableFromUrl.Rows.Count)
                    CsvStringTableFromUrl = CompuMaster.Data.DataTables.ConvertToPlainTextTable(CsvTableFromUrl)
                    Assert.IsTrue(CsvStringTableFromUrl.Contains("Russian Federation"))
                    Assert.IsTrue(CsvStringTableFromUrl.Contains("俄罗斯联邦"))
                    Assert.IsTrue(CsvStringTableFromUrl.Contains("الاتحاد الروسي"))

                Case Else
                    Throw New NotImplementedException
            End Select

        End Sub

        <Test> Public Sub ReadDataTableFromCsvFileWithColumnSeparatorCharInTextStrings()
            Dim TestFile As String = AssemblyTestEnvironment.TestFileAbsolutePath("testfiles\country-codes.csv")
            System.Console.WriteLine("TestFile=" & TestFile)
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

        <Test> Sub ReadWriteCompareDatableWithStringEncoding()
            Dim Level0CsvData As String = """69100"";"""";"""""""";""Text with quotation mark("""")"";""Space""" & System.Environment.NewLine
            Dim Level1CsvDataTable As DataTable = CompuMaster.Data.Csv.ReadDataTableFromCsvString(Level0CsvData, False, ";"c, """"c, False, False)
            Dim Level2CsvData As String = CompuMaster.Data.Csv.WriteDataTableToCsvTextString(Level1CsvDataTable, False, ";"c, """"c, "."c)
            Assert.AreEqual("Text with quotation mark("")", CType(Level1CsvDataTable.Rows(0)(3), String))
            Assert.AreEqual(Level0CsvData, Level2CsvData)
            Console.WriteLine(Level2CsvData)
            Console.WriteLine("Cell 3: " & CType(Level1CsvDataTable.Rows(0)(2), String))
            Console.WriteLine("Cell 4: " & CType(Level1CsvDataTable.Rows(0)(3), String))
        End Sub
        <Test> Sub WriteDataTableToCsvTextStringRecognizeTextChar()
            Dim t As DataTable = SimpleSampleTable()
            Dim ExpectedValue As String
            Dim csv As String

            csv = CompuMaster.Data.Csv.WriteDataTableToCsvTextString(t, True, CompuMaster.Data.Csv.WriteLineEncodings.RowBreakCrLf_CellLineBreakCr, "||", """"c, ".")
            Console.WriteLine(csv)
            ExpectedValue = """col1""||""col2""" & vbNewLine &
                """R1C1""||""R1C2""" & vbNewLine &
                """R2C1""||""R2C2""" & vbNewLine
            Assert.AreEqual(ExpectedValue, csv, "Not expected: CSV EmptyStringAsRecognizeTextChar")

            'Special: ChrW(0) is handled as NO text recognition character!
            'ExpectedValue = ChrW(0) & "col1" & ChrW(0) & "||" & ChrW(0) & "col2" & ChrW(0) & vbNewLine &
            '    ChrW(0) & "R1C1" & ChrW(0) & "||" & ChrW(0) & "R1C2" & ChrW(0) & vbNewLine &
            '    ChrW(0) & "R2C1" & ChrW(0) & "||" & ChrW(0) & "R2C2" & ChrW(0) & vbNewLine
            ExpectedValue = "col1||col2" & vbNewLine &
                "R1C1||R1C2" & vbNewLine &
                "R2C1||R2C2" & vbNewLine
            csv = CompuMaster.Data.Csv.WriteDataTableToCsvTextString(t, True, CompuMaster.Data.Csv.WriteLineEncodings.RowBreakCrLf_CellLineBreakCr, "||", "", ".")
            Console.WriteLine(csv)
            Assert.AreEqual(ExpectedValue, csv, "Not expected: CSV EmptyStringAsRecognizeTextChar")

            csv = CompuMaster.Data.Csv.WriteDataTableToCsvTextString(t, True, CompuMaster.Data.Csv.WriteLineEncodings.RowBreakCrLf_CellLineBreakCr, "||", ChrW(0), ".")
            Console.WriteLine(csv)
            Assert.AreEqual(ExpectedValue, csv, "Not expected: CSV Chrw(0)AsRecognizeTextChar")

            csv = CompuMaster.Data.Csv.WriteDataTableToCsvTextString(t, True, CompuMaster.Data.Csv.WriteLineEncodings.RowBreakCrLf_CellLineBreakCr, "||", ChrW(1), ".")
            ExpectedValue = ChrW(1) & "col1" & ChrW(1) & "||" & ChrW(1) & "col2" & ChrW(1) & vbNewLine &
                ChrW(1) & "R1C1" & ChrW(1) & "||" & ChrW(1) & "R1C2" & ChrW(1) & vbNewLine &
                ChrW(1) & "R2C1" & ChrW(1) & "||" & ChrW(1) & "R2C2" & ChrW(1) & vbNewLine
            Console.WriteLine(csv)
            Assert.AreEqual(ExpectedValue, csv, "Not expected: CSV Chrw(1)AsRecognizeTextChar")
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
            r(0) = "R1C1""""" & vbCrLf
            r(1) = "R1C2""""" & vbCr
            t.Rows.Add(r)
            r = t.NewRow
            r(0) = "R2C1""""" & vbLf
            r(1) = "R2C2""""" & vbTab
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
            Console.WriteLine(csv)
            Dim t2 As DataTable = CompuMaster.Data.Csv.ReadDataTableFromCsvString(csv, True, CompuMaster.Data.Csv.ReadLineEncodings.RowBreakCrLf_CellLineBreakLf, CompuMaster.Data.Csv.ReadLineEncodingAutoConversion.AutoConvertLineBreakToLf)
            'DataTables.AssertTables(t, t2, "Comparison t vs t2")
            'DataTables.AssertTables(CloneTableWithSearchAndReplaceOnStrings(t, vbNewLine, vbLf), t2, "Comparison t vs t2")
            DataTables.AssertTables(CloneTableWithSearchAndReplaceOnStrings(CloneTableWithSearchAndReplaceOnStrings(t, vbNewLine, vbLf), vbCr, vbLf), t2, "Comparison t vs t2")
            'DataTables.AssertTables(CloneTableWithSearchAndReplaceOnStrings(CloneTableWithSearchAndReplaceOnStrings(t, vbNewLine, vbLf), vbCr, vbLf), CloneTableWithSearchAndReplaceOnStrings(t2, vbNewLine, vbLf), "Comparison t vs t2")
            Dim csv2 As String
            Assert.AreEqual(t.Columns.Count, t2.Columns.Count) 'should be the very same
            Assert.AreEqual(t.Rows.Count, t2.Rows.Count) 'should be the very same - except line breaks in cells haven't been converted correctly
            csv2 = CompuMaster.Data.Csv.WriteDataTableToCsvTextString(t2, True)
            Console.WriteLine(csv2)
            Assert.AreEqual(csv, csv2) 'should be the very same
        End Sub

        <Test> Sub CsvEncode()
            Assert.AreEqual("R1C1""""" & vbLf, CompuMaster.Data.CsvTools.CsvEncode("R1C1""" & vbLf, """"c, CompuMaster.Data.Csv.WriteLineEncodings.RowBreakCr_CellLineBreakLf))
            Assert.AreEqual("R1C1""""" & vbLf & vbLf, CompuMaster.Data.CsvTools.CsvEncode("R1C1""" & vbNewLine & vbCr, """"c, CompuMaster.Data.Csv.WriteLineEncodings.RowBreakCr_CellLineBreakLf))
            Assert.AreEqual("R1C1""""""""" & vbLf & vbLf, CompuMaster.Data.CsvTools.CsvEncode("R1C1""""" & vbNewLine & vbCr, """"c, CompuMaster.Data.Csv.WriteLineEncodings.RowBreakCr_CellLineBreakLf))
            Assert.AreEqual("R1C1""""" & vbLf & vbLf, CompuMaster.Data.CsvTools.CsvEncode("R1C1""" & vbNewLine & vbCr, """"c, CompuMaster.Data.Csv.WriteLineEncodings.RowBreakCrLf_CellLineBreakLf))
            Assert.AreEqual("R1C1""""""""" & vbLf & vbLf, CompuMaster.Data.CsvTools.CsvEncode("R1C1""""" & vbNewLine & vbCr, """"c, CompuMaster.Data.Csv.WriteLineEncodings.RowBreakCrLf_CellLineBreakLf))
        End Sub

        Private Function CloneTableWithSearchAndReplaceOnStrings(table As DataTable, searchValue As String, replaceValue As String) As DataTable
            Dim Result As DataTable = CompuMaster.Data.DataTables.CreateDataTableClone(table)
            For MyCounter As Integer = 0 To Result.Columns.Count - 1
                If Result.Columns(MyCounter).DataType Is GetType(String) Then
                    For MyRowCounter As Integer = 0 To Result.Rows.Count - 1
                        If IsDBNull(Result.Rows(MyRowCounter)(MyCounter)) = False AndAlso Not Result.Rows(MyRowCounter)(MyCounter) Is Nothing Then
                            Result.Rows(MyRowCounter)(MyCounter) = CType(Result.Rows(MyRowCounter)(MyCounter), String).Replace(searchValue, replaceValue)
                        End If
                    Next
                End If
            Next
            Return Result
        End Function

        <Test, Ignore("NotYetImplementedCompletely OR ChallengeNotPossibleToSolve")> Sub SupportReadLineBreaks_SqlServerExportFileWithStandardExportSettings()
            Dim importFile As String = AssemblyTestEnvironment.TestFileAbsolutePath("testfiles\sql-server-export.csv")
            Dim dt As DataTable
            dt = CompuMaster.Data.Csv.ReadDataTableFromCsvFile(importFile, True, CompuMaster.Data.Csv.ReadLineEncodings.RowBreakCrLfOrCrOrLf_CellLineBreakCrLfOrCrOrLf, CompuMaster.Data.Csv.ReadLineEncodingAutoConversion.NoAutoConversion, "UTF-8", ";"c, ControlChars.Quote, False, True)

        End Sub

        <Test, Ignore("NotYetImplementedCompletely")> Sub SupportReadLineBreakCrLfWithCellBreakCrLf()
            Dim Result As DataTable = CompuMaster.Data.Csv.ReadDataTableFromCsvFile(AssemblyTestEnvironment.TestFileAbsolutePath("testfiles\test_linebreak_crlf_cellbreak_crlf.csv"), True, CompuMaster.Data.Csv.ReadLineEncodings.RowBreakCrLfOrCrOrLf_CellLineBreakCrLfOrCrOrLf, CompuMaster.Data.Csv.ReadLineEncodingAutoConversion.NoAutoConversion, "UTF-8", ";"c, """"c, False, True)

            'Test output
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(Result, "|", "|", "+", "=", "-"))

            'Some simple tests with 1-liners
            For MyColCounter As Integer = 0 To 2
                Assert.AreEqual("1linerNoQuotes" & (MyColCounter + 1).ToString, Result.Rows(0)(MyColCounter))
            Next
            For MyColCounter As Integer = 0 To 2
                Assert.AreEqual("1liner" & (MyColCounter + 1).ToString, Result.Rows(1)(MyColCounter))
            Next

            'Some simple tests with multiliners but every cell with quotation marks
            For MyColCounter As Integer = 0 To 2
                Assert.AreEqual("1stLine" & ControlChars.CrLf & "2ndLine", Result.Rows(2)(MyColCounter))
            Next

            'Given are following 2 record rows:
            '---
            'Line 1: 1stLineA1NoQuotes
            'Line 2: 2ndLineA1NoQuotes;1stLineA2NoQuotes
            'Line 3: 2ndLineA2NoQuotes;1stLineA3NoQuotes
            'Line 4: 2ndLineA3NoQuotes
            'Line 5: 1stLineB1NoQuotes
            'Line 6: 2ndLineB1NoQuotes;1stLineB2NoQuotes
            'Line 7: 2ndLineB2NoQuotes;1stLineB3NoQuotes
            'Line 8: 2ndLineB3NoQuotes
            '---
            'Line 4 + 5 values must be considered being part of 2nd record appearing with cell 2ndLineB1NoQuotes in line 6
            For MyColCounter As Integer = 0 To 1
                Assert.AreEqual("1stLineA" & (MyColCounter + 1) & "NoQuotes" & ControlChars.CrLf & "2ndLineA" & (MyColCounter + 1) & "NoQuotes", Result.Rows(3)(MyColCounter))
            Next
            Assert.AreEqual("1stLineA3NoQuotes", Result.Rows(3)(2))
            Assert.AreEqual("2ndLineA3NoQuotes" & ControlChars.CrLf & "1stLineB1NoQuotes" & ControlChars.CrLf & "2ndLineB1NoQuotes", Result.Rows(4)(0))
            For MyColCounter As Integer = 1 To 2
                'Assert.AreEqual("1stLineB" & (MyColCounter + 1) & "NoQuotes" & ControlChars.CrLf & "2ndLineB" & (MyColCounter + 1) & "NoQuotes", Result.Rows(4)(MyColCounter))
            Next

            'Given are following 2 record rows:
            '---
            'Line 1: "1stLineAWith""SemiColon;
            'Line 2: 2ndLineAWith""SemiColon;";"1stLineAWithSemiColon;
            'Line 3: 2ndLineAWith""SemiColon;";"1stLineAWithSemiColon;
            'Line 4: 2ndLineAWith""SemiColon;"
            'Line 5: "1stLineBWith""SemiColon;
            'Line 6: 2ndLineBWith""SemiColon;";"1stLineBWithSemiColon;
            'Line 7: 2ndLineBWith""SemiColon;";"1stLineBWithSemiColon;
            'Line 8: 2ndLineBWith""SemiColon;"
            '---
            'Line 1 and line 5 are correctly recognized as new record rows
            For MyColCounter As Integer = 0 To 2
                Assert.AreEqual("1stLineAWith""SemiColon;" & ControlChars.CrLf & "2ndLineAWith""SemiColon;", Result.Rows(5)(MyColCounter))
            Next
            For MyColCounter As Integer = 0 To 2
                Assert.AreEqual("1stLineBWith""SemiColon;" & ControlChars.CrLf & "2ndLineBWith""SemiColon;", Result.Rows(6)(MyColCounter))
            Next

            'Total table summary
            Assert.AreEqual(3, Result.Columns.Count)
            Assert.AreEqual(7, Result.Rows.Count)
        End Sub

    End Class

End Namespace