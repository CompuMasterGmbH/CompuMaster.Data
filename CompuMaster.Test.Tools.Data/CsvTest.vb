﻿Imports NUnit.Framework
Imports System.Data
Imports CompuMaster.Data.CsvTables

Namespace CompuMaster.Test.Data

#Disable Warning CA1822 ' Member als statisch markieren
    <TestFixture(Category:="CSV")> Public Class CsvTest
        Public Sub New()
        End Sub

        Friend Const CSV_ONLINE_TEST_RESOURCE_EN_US_URL As String = "https://raw.githubusercontent.com/datasets/covid-19/main/data/reference.csv"

        Private ReadOnly _OriginCulture As System.Globalization.CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture
        <TearDown> Public Sub ResetCulture()
            System.Threading.Thread.CurrentThread.CurrentCulture = _OriginCulture
        End Sub

        <Test> Public Sub ReadDataTableFromCsvUrlWithTls12Required()
            Dim Url As String = CSV_ONLINE_TEST_RESOURCE_EN_US_URL
            Dim CsvCulture As System.Globalization.CultureInfo = System.Globalization.CultureInfo.CreateSpecificCulture("en-US")
            Dim FileEncoding As System.Text.Encoding = Nothing
            Dim dt As DataTable = CompuMaster.Data.Csv.ReadDataTableFromCsvFile(Url, True, FileEncoding, CsvCulture, """"c, False, True)
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
            Assert.Greater(dt.Columns.Count, 0)
            Assert.Greater(dt.Rows.Count, 0)
        End Sub

        <Test> Public Sub ReadDataTableFromFixedWidthsCsv()
            Dim TestFile As String = AssemblyTestEnvironment.TestFileAbsolutePath(System.IO.Path.Combine("testfiles", "fixedwidths.csv"))
            Dim StartLine As Integer = 0
            System.Console.WriteLine("TestFile=" & TestFile)
            System.Console.WriteLine("StartLine=" & TestFile)

            Dim CsvCulture As System.Globalization.CultureInfo = System.Globalization.CultureInfo.CreateSpecificCulture("de-DE")
            Dim FileEncoding As System.Text.Encoding = System.Text.Encoding.UTF8
            Dim FileEncodingName As String = "UTF-8"
            Dim FixedWidths = New Integer() {6, 40, 25, 25, 25, 25, 25, 25, 25, 25}
            Dim dt As DataTable

            'CSV-File
#Disable Warning BC40000 ' Typ oder Element ist veraltet
            dt = CompuMaster.Data.Csv.ReadDataTableFromCsvFile(
                TestFile, True, StartLine,
                FixedWidths, FileEncoding, CsvCulture, True)
#Enable Warning BC40000 ' Typ oder Element ist veraltet
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
            Assert.AreEqual(10, dt.Columns.Count)
            Assert.AreEqual(3, dt.Rows.Count)
            Assert.AreEqual("Konto", dt.Columns(0).ColumnName)
            Assert.AreEqual("115", dt.Rows(0)(0))
            Assert.AreEqual("Geschäftsausstattung", dt.Rows(2)(1))
            Assert.AreEqual(DBNull.Value, dt.Rows(2)(9))

#Disable Warning BC40000 ' Typ oder Element ist veraltet
            dt = CompuMaster.Data.Csv.ReadDataTableFromCsvFile(
                TestFile, True, StartLine,
                CompuMaster.Data.Csv.ReadLineEncodings.RowBreakCrLfOrCr_CellLineBreakLf, CompuMaster.Data.Csv.ReadLineEncodingAutoConversion.NoAutoConversion,
                FixedWidths, FileEncoding, CsvCulture, True)
#Enable Warning BC40000 ' Typ oder Element ist veraltet
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
            Assert.AreEqual(10, dt.Columns.Count)
            Assert.AreEqual(3, dt.Rows.Count)
            Assert.AreEqual("Konto", dt.Columns(0).ColumnName)
            Assert.AreEqual("115", dt.Rows(0)(0))
            Assert.AreEqual("Geschäftsausstattung", dt.Rows(2)(1))
            Assert.AreEqual(DBNull.Value, dt.Rows(2)(9))

            dt = CompuMaster.Data.Csv.ReadDataTableFromCsvFile(
                TestFile, True, StartLine,
                FixedWidths, FileEncodingName, True)
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
            Assert.AreEqual(10, dt.Columns.Count)
            Assert.AreEqual(3, dt.Rows.Count)
            Assert.AreEqual("Konto", dt.Columns(0).ColumnName)
            Assert.AreEqual("115", dt.Rows(0)(0))
            Assert.AreEqual("Geschäftsausstattung", dt.Rows(2)(1))
            Assert.AreEqual(DBNull.Value, dt.Rows(2)(9))

            dt = CompuMaster.Data.Csv.ReadDataTableFromCsvFile(
                TestFile, True, StartLine,
                CompuMaster.Data.Csv.ReadLineEncodings.RowBreakCrLfOrCr_CellLineBreakLf, CompuMaster.Data.Csv.ReadLineEncodingAutoConversion.NoAutoConversion,
                FixedWidths, FileEncodingName, True)
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
            Assert.AreEqual(10, dt.Columns.Count)
            Assert.AreEqual(3, dt.Rows.Count)
            Assert.AreEqual("Konto", dt.Columns(0).ColumnName)
            Assert.AreEqual("115", dt.Rows(0)(0))
            Assert.AreEqual("Geschäftsausstattung", dt.Rows(2)(1))
            Assert.AreEqual(DBNull.Value, dt.Rows(2)(9))

            'CSV-String
            Dim CsvData As String = System.IO.File.ReadAllText(TestFile, FileEncoding)
            dt = CompuMaster.Data.Csv.ReadDataTableFromCsvString(
                CsvData, True, StartLine,
                FixedWidths, True)
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
            Assert.AreEqual(10, dt.Columns.Count)
            Assert.AreEqual(3, dt.Rows.Count)
            Assert.AreEqual("Konto", dt.Columns(0).ColumnName)
            Assert.AreEqual("115", dt.Rows(0)(0))
            Assert.AreEqual("Geschäftsausstattung", dt.Rows(2)(1))
            Assert.AreEqual(DBNull.Value, dt.Rows(2)(9))

            dt = CompuMaster.Data.Csv.ReadDataTableFromCsvString(
                CsvData, True, StartLine,
                CompuMaster.Data.Csv.ReadLineEncodings.RowBreakCrLfOrCr_CellLineBreakLf, CompuMaster.Data.Csv.ReadLineEncodingAutoConversion.NoAutoConversion,
                FixedWidths, True)
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
            Assert.AreEqual(10, dt.Columns.Count)
            Assert.AreEqual(3, dt.Rows.Count)
            Assert.AreEqual("Konto", dt.Columns(0).ColumnName)
            Assert.AreEqual("115", dt.Rows(0)(0))
            Assert.AreEqual("Geschäftsausstattung", dt.Rows(2)(1))
            Assert.AreEqual(DBNull.Value, dt.Rows(2)(9))

            dt = CompuMaster.Data.Csv.ReadDataTableFromCsvString(
                CsvData, True, StartLine,
                FixedWidths, True)
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
            Assert.AreEqual(10, dt.Columns.Count)
            Assert.AreEqual(3, dt.Rows.Count)
            Assert.AreEqual("Konto", dt.Columns(0).ColumnName)
            Assert.AreEqual("115", dt.Rows(0)(0))
            Assert.AreEqual("Geschäftsausstattung", dt.Rows(2)(1))
            Assert.AreEqual(DBNull.Value, dt.Rows(2)(9))

            dt = CompuMaster.Data.Csv.ReadDataTableFromCsvString(
                CsvData, True, StartLine,
                CompuMaster.Data.Csv.ReadLineEncodings.RowBreakCrLfOrCr_CellLineBreakLf, CompuMaster.Data.Csv.ReadLineEncodingAutoConversion.NoAutoConversion,
                FixedWidths, True)
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
            Assert.AreEqual(10, dt.Columns.Count)
            Assert.AreEqual(3, dt.Rows.Count)
            Assert.AreEqual("Konto", dt.Columns(0).ColumnName)
            Assert.AreEqual("115", dt.Rows(0)(0))
            Assert.AreEqual("Geschäftsausstattung", dt.Rows(2)(1))
            Assert.AreEqual(DBNull.Value, dt.Rows(2)(9))

        End Sub

        <Test> Public Sub ReadDataTableFromFixedWidthsCsv_WithExtraLinesBefore()
            Dim TestFile As String = AssemblyTestEnvironment.TestFileAbsolutePath(System.IO.Path.Combine("testfiles", "fixedwidths_withExtraLinesBefore.csv"))
            Dim StartLine As Integer = 3
            System.Console.WriteLine("TestFile=" & TestFile)
            System.Console.WriteLine("StartLine=" & TestFile)

            Dim CsvCulture As System.Globalization.CultureInfo = System.Globalization.CultureInfo.CreateSpecificCulture("de-DE")
            Dim FileEncoding As System.Text.Encoding = System.Text.Encoding.UTF8
            Dim FileEncodingName As String = "UTF-8"
            Dim FixedWidths = New Integer() {6, 40, 25, 25, 25, 25, 25, 25, 25, 25}
            Dim dt As DataTable

            'CSV-File
#Disable Warning BC40000 ' Typ oder Element ist veraltet
            dt = CompuMaster.Data.Csv.ReadDataTableFromCsvFile(
                TestFile, True, StartLine,
                FixedWidths, FileEncoding, CsvCulture, True)
#Enable Warning BC40000 ' Typ oder Element ist veraltet
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
            Assert.AreEqual(10, dt.Columns.Count)
            Assert.AreEqual(3, dt.Rows.Count)
            Assert.AreEqual("Konto", dt.Columns(0).ColumnName)
            Assert.AreEqual("115", dt.Rows(0)(0))
            Assert.AreEqual("Geschäftsausstattung", dt.Rows(2)(1))
            Assert.AreEqual(DBNull.Value, dt.Rows(2)(9))

#Disable Warning BC40000 ' Typ oder Element ist veraltet
            dt = CompuMaster.Data.Csv.ReadDataTableFromCsvFile(
                TestFile, True, StartLine,
                CompuMaster.Data.Csv.ReadLineEncodings.RowBreakCrLfOrCr_CellLineBreakLf, CompuMaster.Data.Csv.ReadLineEncodingAutoConversion.NoAutoConversion,
                FixedWidths, FileEncoding, CsvCulture, True)
#Enable Warning BC40000 ' Typ oder Element ist veraltet
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
            Assert.AreEqual(10, dt.Columns.Count)
            Assert.AreEqual(3, dt.Rows.Count)
            Assert.AreEqual("Konto", dt.Columns(0).ColumnName)
            Assert.AreEqual("115", dt.Rows(0)(0))
            Assert.AreEqual("Geschäftsausstattung", dt.Rows(2)(1))
            Assert.AreEqual(DBNull.Value, dt.Rows(2)(9))

            dt = CompuMaster.Data.Csv.ReadDataTableFromCsvFile(
                TestFile, True, StartLine,
                FixedWidths, FileEncodingName, True)
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
            Assert.AreEqual(10, dt.Columns.Count)
            Assert.AreEqual(3, dt.Rows.Count)
            Assert.AreEqual("Konto", dt.Columns(0).ColumnName)
            Assert.AreEqual("115", dt.Rows(0)(0))
            Assert.AreEqual("Geschäftsausstattung", dt.Rows(2)(1))
            Assert.AreEqual(DBNull.Value, dt.Rows(2)(9))

            dt = CompuMaster.Data.Csv.ReadDataTableFromCsvFile(
                TestFile, True, StartLine,
                CompuMaster.Data.Csv.ReadLineEncodings.RowBreakCrLfOrCr_CellLineBreakLf, CompuMaster.Data.Csv.ReadLineEncodingAutoConversion.NoAutoConversion,
                FixedWidths, FileEncodingName, True)
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
            Assert.AreEqual(10, dt.Columns.Count)
            Assert.AreEqual(3, dt.Rows.Count)
            Assert.AreEqual("Konto", dt.Columns(0).ColumnName)
            Assert.AreEqual("115", dt.Rows(0)(0))
            Assert.AreEqual("Geschäftsausstattung", dt.Rows(2)(1))
            Assert.AreEqual(DBNull.Value, dt.Rows(2)(9))

            'CSV-String
            Dim CsvData As String = System.IO.File.ReadAllText(TestFile, FileEncoding)
            dt = CompuMaster.Data.Csv.ReadDataTableFromCsvString(
                CsvData, True, StartLine,
                FixedWidths, True)
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
            Assert.AreEqual(10, dt.Columns.Count)
            Assert.AreEqual(3, dt.Rows.Count)
            Assert.AreEqual("Konto", dt.Columns(0).ColumnName)
            Assert.AreEqual("115", dt.Rows(0)(0))
            Assert.AreEqual("Geschäftsausstattung", dt.Rows(2)(1))
            Assert.AreEqual(DBNull.Value, dt.Rows(2)(9))

            dt = CompuMaster.Data.Csv.ReadDataTableFromCsvString(
                CsvData, True, StartLine,
                CompuMaster.Data.Csv.ReadLineEncodings.RowBreakCrLfOrCr_CellLineBreakLf, CompuMaster.Data.Csv.ReadLineEncodingAutoConversion.NoAutoConversion,
                FixedWidths, True)
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
            Assert.AreEqual(10, dt.Columns.Count)
            Assert.AreEqual(3, dt.Rows.Count)
            Assert.AreEqual("Konto", dt.Columns(0).ColumnName)
            Assert.AreEqual("115", dt.Rows(0)(0))
            Assert.AreEqual("Geschäftsausstattung", dt.Rows(2)(1))
            Assert.AreEqual(DBNull.Value, dt.Rows(2)(9))

            dt = CompuMaster.Data.Csv.ReadDataTableFromCsvString(
                CsvData, True, StartLine,
                FixedWidths, True)
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
            Assert.AreEqual(10, dt.Columns.Count)
            Assert.AreEqual(3, dt.Rows.Count)
            Assert.AreEqual("Konto", dt.Columns(0).ColumnName)
            Assert.AreEqual("115", dt.Rows(0)(0))
            Assert.AreEqual("Geschäftsausstattung", dt.Rows(2)(1))
            Assert.AreEqual(DBNull.Value, dt.Rows(2)(9))

            dt = CompuMaster.Data.Csv.ReadDataTableFromCsvString(
                CsvData, True, StartLine,
                CompuMaster.Data.Csv.ReadLineEncodings.RowBreakCrLfOrCr_CellLineBreakLf, CompuMaster.Data.Csv.ReadLineEncodingAutoConversion.NoAutoConversion,
                FixedWidths, True)
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
            Assert.AreEqual(10, dt.Columns.Count)
            Assert.AreEqual(3, dt.Rows.Count)
            Assert.AreEqual("Konto", dt.Columns(0).ColumnName)
            Assert.AreEqual("115", dt.Rows(0)(0))
            Assert.AreEqual("Geschäftsausstattung", dt.Rows(2)(1))
            Assert.AreEqual(DBNull.Value, dt.Rows(2)(9))

        End Sub

        <Test> Public Sub ReadDataTableFromDatevCsv()
            Dim TestFile As String = AssemblyTestEnvironment.TestFileAbsolutePath(System.IO.Path.Combine("testfiles", "datev.csv"))
            Dim StartLine As Integer = 6
            System.Console.WriteLine("TestFile=" & TestFile)
            System.Console.WriteLine("StartLine=" & TestFile)

            Dim CsvCulture As System.Globalization.CultureInfo = System.Globalization.CultureInfo.CreateSpecificCulture("de-DE")
            Dim FileEncoding As System.Text.Encoding = System.Text.Encoding.UTF8
            Dim FileEncodingName As String = "UTF-8"
            Dim dt As DataTable

            'CSV-File
            dt = CompuMaster.Data.Csv.ReadDataTableFromCsvFile(
                New CsvFileOptions(TestFile, FileEncoding),
                New CsvReadOptionsDynamicColumnSize(True, StartLine, CsvCulture, """"c, True))
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
            Assert.AreEqual(10, dt.Columns.Count)
            Assert.AreEqual(3, dt.Rows.Count)
            Assert.AreEqual("Konto", dt.Columns(0).ColumnName)
            Assert.AreEqual("115", dt.Rows(0)(0))
            Assert.AreEqual("Geschäftsausstattung", dt.Rows(2)(1))
            Assert.AreEqual(DBNull.Value, dt.Rows(2)(9))

            dt = CompuMaster.Data.Csv.ReadDataTableFromCsvFile(
                TestFile, True, StartLine,
                FileEncoding, CsvCulture, """"c, False, True)
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
            Assert.AreEqual(10, dt.Columns.Count)
            Assert.AreEqual(3, dt.Rows.Count)
            Assert.AreEqual("Konto", dt.Columns(0).ColumnName)
            Assert.AreEqual("115", dt.Rows(0)(0))
            Assert.AreEqual("Geschäftsausstattung", dt.Rows(2)(1))
            Assert.AreEqual(DBNull.Value, dt.Rows(2)(9))

            dt = CompuMaster.Data.Csv.ReadDataTableFromCsvFile(
                TestFile, True, StartLine,
                CompuMaster.Data.Csv.ReadLineEncodings.RowBreakCrLfOrCr_CellLineBreakLf, CompuMaster.Data.Csv.ReadLineEncodingAutoConversion.NoAutoConversion,
                FileEncoding, CsvCulture, """"c, False, True)
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
            Assert.AreEqual(10, dt.Columns.Count)
            Assert.AreEqual(3, dt.Rows.Count)
            Assert.AreEqual("Konto", dt.Columns(0).ColumnName)
            Assert.AreEqual("115", dt.Rows(0)(0))
            Assert.AreEqual("Geschäftsausstattung", dt.Rows(2)(1))
            Assert.AreEqual(DBNull.Value, dt.Rows(2)(9))

            dt = CompuMaster.Data.Csv.ReadDataTableFromCsvFile(
                TestFile, True, StartLine,
                FileEncodingName, columnSeparator:=";"c, recognizeTextBy:=""""c, False, True)
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
            Assert.AreEqual(10, dt.Columns.Count)
            Assert.AreEqual(3, dt.Rows.Count)
            Assert.AreEqual("Konto", dt.Columns(0).ColumnName)
            Assert.AreEqual("115", dt.Rows(0)(0))
            Assert.AreEqual("Geschäftsausstattung", dt.Rows(2)(1))
            Assert.AreEqual(DBNull.Value, dt.Rows(2)(9))

            dt = CompuMaster.Data.Csv.ReadDataTableFromCsvFile(
                TestFile, True, StartLine,
                CompuMaster.Data.Csv.ReadLineEncodings.RowBreakCrLfOrCr_CellLineBreakLf, CompuMaster.Data.Csv.ReadLineEncodingAutoConversion.NoAutoConversion,
                FileEncodingName, columnSeparator:=";"c, recognizeTextBy:=""""c, False, True)
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
            Assert.AreEqual(10, dt.Columns.Count)
            Assert.AreEqual(3, dt.Rows.Count)
            Assert.AreEqual("Konto", dt.Columns(0).ColumnName)
            Assert.AreEqual("115", dt.Rows(0)(0))
            Assert.AreEqual("Geschäftsausstattung", dt.Rows(2)(1))
            Assert.AreEqual(DBNull.Value, dt.Rows(2)(9))

            'CSV-String
            Dim CsvData As String = System.IO.File.ReadAllText(TestFile, FileEncoding)
            dt = CompuMaster.Data.Csv.ReadDataTableFromCsvString(
                CsvData, True, StartLine,
                CsvCulture, """"c, False, True)
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
            Assert.AreEqual(10, dt.Columns.Count)
            Assert.AreEqual(3, dt.Rows.Count)
            Assert.AreEqual("Konto", dt.Columns(0).ColumnName)
            Assert.AreEqual("115", dt.Rows(0)(0))
            Assert.AreEqual("Geschäftsausstattung", dt.Rows(2)(1))
            Assert.AreEqual(DBNull.Value, dt.Rows(2)(9))

            dt = CompuMaster.Data.Csv.ReadDataTableFromCsvString(
                CsvData, True, StartLine,
                CompuMaster.Data.Csv.ReadLineEncodings.RowBreakCrLfOrCr_CellLineBreakLf, CompuMaster.Data.Csv.ReadLineEncodingAutoConversion.NoAutoConversion,
                CsvCulture, """"c, False, True)
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
            Assert.AreEqual(10, dt.Columns.Count)
            Assert.AreEqual(3, dt.Rows.Count)
            Assert.AreEqual("Konto", dt.Columns(0).ColumnName)
            Assert.AreEqual("115", dt.Rows(0)(0))
            Assert.AreEqual("Geschäftsausstattung", dt.Rows(2)(1))
            Assert.AreEqual(DBNull.Value, dt.Rows(2)(9))

            dt = CompuMaster.Data.Csv.ReadDataTableFromCsvString(
                CsvData, True, StartLine,
                ";"c, """"c, False, True)
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
            Assert.AreEqual(10, dt.Columns.Count)
            Assert.AreEqual(3, dt.Rows.Count)
            Assert.AreEqual("Konto", dt.Columns(0).ColumnName)
            Assert.AreEqual("115", dt.Rows(0)(0))
            Assert.AreEqual("Geschäftsausstattung", dt.Rows(2)(1))
            Assert.AreEqual(DBNull.Value, dt.Rows(2)(9))

            dt = CompuMaster.Data.Csv.ReadDataTableFromCsvString(
                CsvData, True, StartLine,
                CompuMaster.Data.Csv.ReadLineEncodings.RowBreakCrLfOrCr_CellLineBreakLf, CompuMaster.Data.Csv.ReadLineEncodingAutoConversion.NoAutoConversion,
                ";"c, """"c, False, True)
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
            Assert.AreEqual(10, dt.Columns.Count)
            Assert.AreEqual(3, dt.Rows.Count)
            Assert.AreEqual("Konto", dt.Columns(0).ColumnName)
            Assert.AreEqual("115", dt.Rows(0)(0))
            Assert.AreEqual("Geschäftsausstattung", dt.Rows(2)(1))
            Assert.AreEqual(DBNull.Value, dt.Rows(2)(9))

        End Sub

        ''' <summary>
        ''' Test from a mini-webserver providing a CSV download with missing response header content-type/charset 
        ''' </summary>
        ''' <remarks>
        ''' The CSV file is returned as UTF-8 bytes
        ''' </remarks>
        <Test, NonParallelizable> Public Sub ReadDataTableFromCsvUrlAtLocalhostWithContentTypeButWithoutCharset(<Values(0, 1, 2)> headerContentTypeVariantToExecute As Integer)
            Dim Url As String = "http://localhost:" & 8035 + headerContentTypeVariantToExecute & "/"
            Dim CsvCulture As System.Globalization.CultureInfo = System.Globalization.CultureInfo.CreateSpecificCulture("en-US")
            Dim FileEncoding As System.Text.Encoding = Nothing
            Dim ws As New CompuMaster.Web.TinyWebServerAdvanced.WebServer(AddressOf ReadDataTableLocalhostTestWebserver,
                                                                          Function(handler As System.Net.HttpListenerRequest) As System.Collections.Specialized.NameValueCollection
                                                                              Dim HeaderContentTypeVariants As String() = New String() {Nothing, "text/csv", "text/csv; charset=utf-8"}
                                                                              Dim HeaderContentType As String = HeaderContentTypeVariants(headerContentTypeVariantToExecute)
                                                                              Dim Headers As New System.Collections.Specialized.NameValueCollection
                                                                              If HeaderContentType <> Nothing Then
                                                                                  Headers("content-type") = HeaderContentType
                                                                              End If
                                                                              Return Headers
                                                                          End Function,
                                                                          New String() {Url})
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
        <CodeAnalysis.SuppressMessage("Major Code Smell", "S1172:Unused procedure parameters should be removed", Justification:="<Ausstehend>")>
        Private Shared Function ReadDataTableLocalhostTestWebserver(handler As System.Net.HttpListenerRequest, ParamArray urls As String()) As String
            Return "Test,Column" & ControlChars.CrLf & "1,äöüßÄÖÜ2"
        End Function

        <Test()> Public Sub ReadDataTableFromCsvStringSeparatorSeparatedMustFailsBecauseOfWrongCulture(<Values("en-US", "en-GB", "ja-JP")> cultureContext As String)
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture(cultureContext)

            Dim testinputdata As String
            Dim testoutputdata As DataTable

            testinputdata = "ID;""Description"";DateValue" & ControlChars.CrLf &
                "5;""line1 ü content"";2005-08-29" & ControlChars.CrLf &
                ";""line 2 """" content"";""2005-08-27"""
            testoutputdata = CompuMaster.Data.Csv.ReadDataTableFromCsvString(testinputdata, True, System.Threading.Thread.CurrentThread.CurrentCulture)
            'Throw New Exception(CompuMaster.Data.DataTables.ConvertToPlainTextTable(testoutputdata))
            NUnit.Framework.Assert.AreNotEqual(3, testoutputdata.Columns.Count, "JW #100")
            NUnit.Framework.Assert.AreEqual(2, testoutputdata.Rows.Count, "JW #101")
        End Sub

        <Test> <CodeAnalysis.SuppressMessage("Performance", "CA1825:Avoid zero-length array allocations.", Justification:="<Ausstehend>")>
        Public Sub ReadDataTableFromCsvFileViaHttpRequestWithCorrectCharsetEncoding(<Values(1, 2, 3)> testType As Byte)
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
#Disable Warning BC40000 ' Typ oder Element ist veraltet
                    CsvTableFromUrl = CompuMaster.Data.Csv.ReadDataTableFromCsvFile(GithubCountryCodesTestUrl, True, New Integer() {}, CType(Nothing, System.Text.Encoding), System.Globalization.CultureInfo.GetCultureInfo("en-US"), False)
#Enable Warning BC40000 ' Typ oder Element ist veraltet
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
            Dim TestFile As String = AssemblyTestEnvironment.TestFileAbsolutePath(System.IO.Path.Combine("testfiles", "country-codes.csv"))
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

            testinputdata = "ID;""Description"";DateValue" & ControlChars.CrLf &
                "5;""line1 ü content"";2005-08-29" & ControlChars.CrLf &
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

            testinputdata = "ID;""Description"";DateValue" & ControlChars.CrLf &
                "5;""line1 ü content"";2005-08-29" & ControlChars.Lf &
                ";""line 2 " & ControlChars.Lf & "newline content"";""2005-08-27"";" & ControlChars.CrLf
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

            testinputdata = "ID,""Description"",DateValue" & ControlChars.CrLf &
                "5,""line1 ü content"",2005-08-29" & ControlChars.CrLf &
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

            testinputdata = "ID,""Description"",DateValue" & ControlChars.CrLf &
                "5,""line1 ü content"",2005-08-29" & ControlChars.Lf &
                ",""line 2 " & ControlChars.Lf & "newline content"",""2005-08-27""," & ControlChars.CrLf
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

            testinputdata = "ID;""Description"";DateValue" & ControlChars.CrLf &
                "5;""line1 ü content"";2005-08-29" & ControlChars.CrLf &
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

            testinputdata = "ID;""Description"";DateValue" & ControlChars.CrLf &
                "5;""line1 ü content"";2005-08-29" & ControlChars.Lf &
                ";""line 2 " & ControlChars.Lf & "newline content"";""2005-08-27"";" & ControlChars.CrLf
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

            testinputdata = "ID,""Description"",DateValue" & ControlChars.CrLf &
                "5,""line1 ü content"",2005-08-29" & ControlChars.CrLf &
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

            testinputdata = "ID,""Description"",DateValue" & ControlChars.CrLf &
                "5,""line1 ü content"",2005-08-29" & ControlChars.Lf &
                ",""line 2 " & ControlChars.Lf & "newline content"",""2005-08-27""," & ControlChars.CrLf
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

        <Test()> Public Sub ReadDataTableFromCsvReaderFixedColumnWidths(<NUnit.Framework.Values(ControlChars.CrLf, ControlChars.Cr & "", ControlChars.Lf & "")> newLineStyle As String)
            Dim testinputdata As String
            Dim testoutputdata As DataTable
            Dim columnWidths As Integer()

            testinputdata = "ID     Description   DateValue       " & newLineStyle &
                            "12345678901234567890123456789012" & newLineStyle &
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

            'testinputdata = "ID;""Description"";DateValue" & ControlChars.CrLf & _
            '    "5;""line1 ü content"";2005-08-29" & ControlChars.CrLf & _
            '    ";""line 2 " & ControlChars.Lf & "newline content"";""2005-08-27"";" & ControlChars.CrLf
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
            Assert.AreEqual(New Byte() {239, 187, 191}, System.Text.Encoding.UTF8.GetPreamble())
            Assert.Greater(bom.Length, 0, "Utf8Preamble must contain at least 1 Char")
            Dim csv As String
#Disable Warning BC40000 ' Typ oder Element ist veraltet
            csv = CompuMaster.Data.Csv.WriteDataTableToCsvFileStringWithTextEncoding(t, True)
            Assert.AreEqual(bom, csv.Substring(0, bom.Length), "CSV starts correctly with BOM signature for UTF-8")
            csv = CompuMaster.Data.Csv.WriteDataTableToCsvFileStringWithTextEncoding(t, True, "UTF-8")
            Assert.AreEqual(bom, csv.Substring(0, bom.Length), "CSV starts correctly with BOM signature for UTF-8")
#Enable Warning BC40000 ' Typ oder Element ist veraltet
        End Sub

        <Test> Sub WriteDataTableToCsvTextString()
            Dim t As DataTable = SimpleSampleTable()
            Dim bom As String = System.Text.Encoding.UTF8.GetString(System.Text.Encoding.UTF8.GetPreamble())
            Dim csv As String
            csv = CompuMaster.Data.Csv.WriteDataTableToCsvTextString(t, True)
            Assert.AreNotEqual(bom, csv.Substring(0, bom.Length), "CSV starts invalidly with BOM signature for UTF-8")
        End Sub

        <Test> Sub ReadWriteCompareDatableWithStringEncodingDefault()
            Dim Level0CsvData As String = """69100"";"""";"""""""";""Text with quotation mark("""")"";""Space""" & ControlChars.CrLf
            Dim Level1CsvDataTable As DataTable = CompuMaster.Data.Csv.ReadDataTableFromCsvString(Level0CsvData, False, ";"c, """"c, False, False)
            Dim Level2CsvData As String = CompuMaster.Data.Csv.WriteDataTableToCsvTextString(Level1CsvDataTable, False, ";"c, """"c, "."c)
            Assert.AreEqual("Text with quotation mark("")", CType(Level1CsvDataTable.Rows(0)(3), String))
            Assert.AreEqual(Level0CsvData, Level2CsvData)
            Console.WriteLine(Level2CsvData)
            Console.WriteLine("Cell 3: " & CType(Level1CsvDataTable.Rows(0)(2), String))
            Console.WriteLine("Cell 4: " & CType(Level1CsvDataTable.Rows(0)(3), String))
        End Sub

        <Test> Sub ReadWriteCompareDatableWithStringEncodingOsPlatformDependent()
            Dim Level0CsvData As String = """69100"";"""";"""""""";""Text with quotation mark("""")"";""Space""" & System.Environment.NewLine
            Dim Level1CsvDataTable As DataTable = CompuMaster.Data.Csv.ReadDataTableFromCsvString(Level0CsvData, False, ";"c, """"c, False, False)
            Dim Level2CsvData As String = CompuMaster.Data.Csv.WriteDataTableToCsvTextString(Level1CsvDataTable, False, CompuMaster.Data.Csv.WriteLineEncodings.Auto, ";"c, """"c, "."c)
            Assert.AreEqual("Text with quotation mark("")", CType(Level1CsvDataTable.Rows(0)(3), String))
            Assert.AreEqual(Level0CsvData, Level2CsvData)
            Console.WriteLine(Level2CsvData)
            Console.WriteLine("Cell 3: " & CType(Level1CsvDataTable.Rows(0)(2), String))
            Console.WriteLine("Cell 4: " & CType(Level1CsvDataTable.Rows(0)(3), String))
        End Sub

        <Test> Sub ReadWriteCompareDatableFixedColWidthsWithStringEncodingDefault()
            Dim Level0CsvData As String = "ID   GROSS klein " & ControlChars.CrLf & "12345ABCDEFabcdef" & ControlChars.CrLf & "äöüß AEIOU aeiou " & ControlChars.CrLf
            Dim Level1CsvDataTable As DataTable = CompuMaster.Data.Csv.ReadDataTableFromCsvString(Level0CsvData, False, New Integer() {5, 6, 6}, False)
            Dim Level2CsvData As String = CompuMaster.Data.Csv.ConvertDataTableToTextAsStringBuilder(Level1CsvDataTable, False, CompuMaster.Data.Csv.WriteLineEncodings.RowBreakCrLf_CellLineBreakLf, System.Globalization.CultureInfo.CurrentCulture, New Integer() {5, 6, 6}).ToString
            Assert.AreEqual("ID", CType(Level1CsvDataTable.Rows(0)(0), String))
            Assert.AreEqual("GROSS", CType(Level1CsvDataTable.Rows(0)(1), String))
            Assert.AreEqual("klein", CType(Level1CsvDataTable.Rows(0)(2), String))
            Assert.AreEqual(Level0CsvData, Level2CsvData)
            Console.WriteLine(Level2CsvData)
        End Sub

        <Test> Sub ReadWriteCompareDatableFixedColWidthsWithStringEncodingOsPlatformDependent()
            Dim Level0CsvData As String = "ID   GROSS klein " & System.Environment.NewLine & "12345ABCDEFabcdef" & System.Environment.NewLine & "äöüß AEIOU aeiou " & System.Environment.NewLine
            Dim Level1CsvDataTable As DataTable = CompuMaster.Data.Csv.ReadDataTableFromCsvString(Level0CsvData, False, New Integer() {5, 6, 6}, False)
            Dim Level2CsvData As String = CompuMaster.Data.Csv.ConvertDataTableToTextAsStringBuilder(Level1CsvDataTable, False, CompuMaster.Data.Csv.WriteLineEncodings.Auto, System.Globalization.CultureInfo.CurrentCulture, New Integer() {5, 6, 6}).ToString
            Assert.AreEqual("ID", CType(Level1CsvDataTable.Rows(0)(0), String))
            Assert.AreEqual("GROSS", CType(Level1CsvDataTable.Rows(0)(1), String))
            Assert.AreEqual("klein", CType(Level1CsvDataTable.Rows(0)(2), String))
            Assert.AreEqual(Level0CsvData, Level2CsvData)
            Console.WriteLine(Level2CsvData)
        End Sub

        <Test> Sub ReadWriteCompareDatableFixedColWidthsWithStringEncodingLinux()
            Dim Level0CsvData As String = "ID   GROSS klein " & ControlChars.Lf & "12345ABCDEFabcdef" & ControlChars.Lf & "äöüß AEIOU aeiou " & ControlChars.Lf
            Dim Level1CsvDataTable As DataTable = CompuMaster.Data.Csv.ReadDataTableFromCsvString(Level0CsvData, False, New Integer() {5, 6, 6}, False)
            Dim Level2CsvData As String = CompuMaster.Data.Csv.ConvertDataTableToTextAsStringBuilder(Level1CsvDataTable, False, CompuMaster.Data.Csv.WriteLineEncodings.RowBreakLf_CellLineBreakCr, System.Globalization.CultureInfo.CurrentCulture, New Integer() {5, 6, 6}).ToString
            Assert.AreEqual("ID", CType(Level1CsvDataTable.Rows(0)(0), String))
            Assert.AreEqual("GROSS", CType(Level1CsvDataTable.Rows(0)(1), String))
            Assert.AreEqual("klein", CType(Level1CsvDataTable.Rows(0)(2), String))
            Assert.AreEqual(Level0CsvData, Level2CsvData)
            Console.WriteLine(Level2CsvData)
        End Sub

        <Test> Sub WriteDataTableToCsvTextStringRecognizeTextChar()
            Dim t As DataTable = SimpleSampleTable()
            Dim ExpectedValue As String
            Dim csv As String

            'Line break checks for rows
            ExpectedValue = """col1""||""col2""" & System.Environment.NewLine &
                """R1C1""||""R1C2""" & System.Environment.NewLine &
                """R2C1""||""R2C2""" & System.Environment.NewLine
            csv = CompuMaster.Data.Csv.WriteDataTableToCsvTextString(t, True, CompuMaster.Data.Csv.WriteLineEncodings.Auto, "||", """"c, ".")
            Console.WriteLine(csv)
            Assert.AreEqual(ExpectedValue, csv, "Not expected: CSV EmptyStringAsRecognizeTextChar")

            csv = CompuMaster.Data.Csv.WriteDataTableToCsvTextString(t, True, CompuMaster.Data.Csv.WriteLineEncodings.None, "||", """"c, ".")
            Console.WriteLine(csv)
            Assert.AreEqual(ExpectedValue, csv, "Not expected: CSV EmptyStringAsRecognizeTextChar")

            ExpectedValue = """col1""||""col2""" & ControlChars.CrLf &
                """R1C1""||""R1C2""" & ControlChars.CrLf &
                """R2C1""||""R2C2""" & ControlChars.CrLf
            csv = CompuMaster.Data.Csv.WriteDataTableToCsvTextString(t, True, CompuMaster.Data.Csv.WriteLineEncodings.Default, "||", """"c, ".")
            Console.WriteLine(csv)
            Assert.AreEqual(ExpectedValue, csv, "Not expected: CSV EmptyStringAsRecognizeTextChar")

            csv = CompuMaster.Data.Csv.WriteDataTableToCsvTextString(t, True, CompuMaster.Data.Csv.WriteLineEncodings.RowBreakCrLf_CellLineBreakCr, "||", """"c, ".")
            Console.WriteLine(csv)
            Assert.AreEqual(ExpectedValue, csv, "Not expected: CSV EmptyStringAsRecognizeTextChar")

            ExpectedValue = """col1""||""col2""" & ControlChars.Cr &
                """R1C1""||""R1C2""" & ControlChars.Cr &
                """R2C1""||""R2C2""" & ControlChars.Cr
            csv = CompuMaster.Data.Csv.WriteDataTableToCsvTextString(t, True, CompuMaster.Data.Csv.WriteLineEncodings.RowBreakCr_CellLineBreakLf, "||", """"c, ".")
            Console.WriteLine(csv)
            Assert.AreEqual(ExpectedValue, csv, "Not expected: CSV EmptyStringAsRecognizeTextChar")

            ExpectedValue = """col1""||""col2""" & ControlChars.Lf &
                """R1C1""||""R1C2""" & ControlChars.Lf &
                """R2C1""||""R2C2""" & ControlChars.Lf
            csv = CompuMaster.Data.Csv.WriteDataTableToCsvTextString(t, True, CompuMaster.Data.Csv.WriteLineEncodings.RowBreakLf_CellLineBreakCr, "||", """"c, ".")
            Console.WriteLine(csv)
            Assert.AreEqual(ExpectedValue, csv, "Not expected: CSV EmptyStringAsRecognizeTextChar")

            'Special: ChrW(0) is handled as NO text recognition character!
            'ExpectedValue = ChrW(0) & "col1" & ChrW(0) & "||" & ChrW(0) & "col2" & ChrW(0) & ControlChars.CrLf &
            '    ChrW(0) & "R1C1" & ChrW(0) & "||" & ChrW(0) & "R1C2" & ChrW(0) & ControlChars.CrLf &
            '    ChrW(0) & "R2C1" & ChrW(0) & "||" & ChrW(0) & "R2C2" & ChrW(0) & ControlChars.CrLf
            ExpectedValue = "col1||col2" & ControlChars.CrLf &
                "R1C1||R1C2" & ControlChars.CrLf &
                "R2C1||R2C2" & ControlChars.CrLf
            csv = CompuMaster.Data.Csv.WriteDataTableToCsvTextString(t, True, CompuMaster.Data.Csv.WriteLineEncodings.RowBreakCrLf_CellLineBreakCr, "||", "", ".")
            Console.WriteLine(csv)
            Assert.AreEqual(ExpectedValue, csv, "Not expected: CSV EmptyStringAsRecognizeTextChar")

            csv = CompuMaster.Data.Csv.WriteDataTableToCsvTextString(t, True, CompuMaster.Data.Csv.WriteLineEncodings.RowBreakCrLf_CellLineBreakCr, "||", ChrW(0), ".")
            Console.WriteLine(csv)
            Assert.AreEqual(ExpectedValue, csv, "Not expected: CSV Chrw(0)AsRecognizeTextChar")

            ExpectedValue = ChrW(1) & "col1" & ChrW(1) & "||" & ChrW(1) & "col2" & ChrW(1) & ControlChars.CrLf &
                ChrW(1) & "R1C1" & ChrW(1) & "||" & ChrW(1) & "R1C2" & ChrW(1) & ControlChars.CrLf &
                ChrW(1) & "R2C1" & ChrW(1) & "||" & ChrW(1) & "R2C2" & ChrW(1) & ControlChars.CrLf
            csv = CompuMaster.Data.Csv.WriteDataTableToCsvTextString(t, True, CompuMaster.Data.Csv.WriteLineEncodings.RowBreakCrLf_CellLineBreakCr, "||", ChrW(1), ".")
            Console.WriteLine(csv)
            Assert.AreEqual(ExpectedValue, csv, "Not expected: CSV Chrw(1)AsRecognizeTextChar")

            csv = CompuMaster.Data.Csv.WriteDataTableToCsvTextString(t, True, CompuMaster.Data.Csv.WriteLineEncodings.RowBreakCrLf_CellLineBreakCr, "||", ChrW(1), ".")
            Console.WriteLine(csv)
            Assert.AreEqual(ExpectedValue, csv, "Not expected: CSV Chrw(1)AsRecognizeTextChar")
        End Sub

        <Test> Sub WriteDataTableToCsvTextStringLineEncodings()
            Dim t As DataTable = SimpleSampleTableWithLineBreaksInCells(System.Environment.NewLine)
            Dim ExpectedValue As String, ExpectedRowLineBreak As String, ExpectedCellLineBreak As String
            Dim csv As String

            'Line break checks for multiline cells
            ExpectedRowLineBreak = System.Environment.NewLine
            ExpectedCellLineBreak = System.Environment.NewLine
            ExpectedValue = """col1" & ExpectedCellLineBreak & "Line 2""||""col2" & ExpectedCellLineBreak & "Line 2""" & ExpectedRowLineBreak &
                """R1C1" & ExpectedCellLineBreak & "Line 2""||""R1C2" & ExpectedCellLineBreak & "Line 2""" & ExpectedRowLineBreak &
                """R2C1" & ExpectedCellLineBreak & "Line 2""||""R2C2" & ExpectedCellLineBreak & "Line 2""" & ExpectedRowLineBreak
            csv = CompuMaster.Data.Csv.WriteDataTableToCsvTextString(t, True, CompuMaster.Data.Csv.WriteLineEncodings.None, "||", """"c, ".")
            Console.WriteLine(csv)
            Assert.AreEqual(ExpectedValue, csv, "Linebreaks=PlatformNewLine/UnChanged(PlatformNewLine)")

            ExpectedRowLineBreak = ControlChars.CrLf
            ExpectedCellLineBreak = ControlChars.Lf
            ExpectedValue = """col1" & ExpectedCellLineBreak & "Line 2""||""col2" & ExpectedCellLineBreak & "Line 2""" & ExpectedRowLineBreak &
                """R1C1" & ExpectedCellLineBreak & "Line 2""||""R1C2" & ExpectedCellLineBreak & "Line 2""" & ExpectedRowLineBreak &
                """R2C1" & ExpectedCellLineBreak & "Line 2""||""R2C2" & ExpectedCellLineBreak & "Line 2""" & ExpectedRowLineBreak
            csv = CompuMaster.Data.Csv.WriteDataTableToCsvTextString(t, True, CompuMaster.Data.Csv.WriteLineEncodings.RowBreakCrLf_CellLineBreakLf, "||", """"c, ".")
            Console.WriteLine(csv)
            Assert.AreEqual(ExpectedValue, csv, "Linebreaks=CrLf/Lf")

            csv = CompuMaster.Data.Csv.WriteDataTableToCsvTextString(t, True, CompuMaster.Data.Csv.WriteLineEncodings.Default, "||", """"c, ".")
            Console.WriteLine(csv)
            Assert.AreEqual(ExpectedValue, csv, "Linebreaks=Default")

            ExpectedRowLineBreak = ControlChars.CrLf
            ExpectedCellLineBreak = ControlChars.Cr
            ExpectedValue = """col1" & ExpectedCellLineBreak & "Line 2""||""col2" & ExpectedCellLineBreak & "Line 2""" & ExpectedRowLineBreak &
                """R1C1" & ExpectedCellLineBreak & "Line 2""||""R1C2" & ExpectedCellLineBreak & "Line 2""" & ExpectedRowLineBreak &
                """R2C1" & ExpectedCellLineBreak & "Line 2""||""R2C2" & ExpectedCellLineBreak & "Line 2""" & ExpectedRowLineBreak
            csv = CompuMaster.Data.Csv.WriteDataTableToCsvTextString(t, True, CompuMaster.Data.Csv.WriteLineEncodings.RowBreakCrLf_CellLineBreakCr, "||", """"c, ".")
            Console.WriteLine(csv)
            Assert.AreEqual(ExpectedValue, csv, "Linebreaks=CrLf/Cr")

            ExpectedRowLineBreak = ControlChars.Cr
            ExpectedCellLineBreak = ControlChars.Lf
            ExpectedValue = """col1" & ExpectedCellLineBreak & "Line 2""||""col2" & ExpectedCellLineBreak & "Line 2""" & ExpectedRowLineBreak &
                """R1C1" & ExpectedCellLineBreak & "Line 2""||""R1C2" & ExpectedCellLineBreak & "Line 2""" & ExpectedRowLineBreak &
                """R2C1" & ExpectedCellLineBreak & "Line 2""||""R2C2" & ExpectedCellLineBreak & "Line 2""" & ExpectedRowLineBreak
            csv = CompuMaster.Data.Csv.WriteDataTableToCsvTextString(t, True, CompuMaster.Data.Csv.WriteLineEncodings.RowBreakCr_CellLineBreakLf, "||", """"c, ".")
            Console.WriteLine(csv)
            Assert.AreEqual(ExpectedValue, csv, "Linebreaks=Cr/Lf")

            ExpectedRowLineBreak = ControlChars.Lf
            ExpectedCellLineBreak = ControlChars.Cr
            ExpectedValue = """col1" & ExpectedCellLineBreak & "Line 2""||""col2" & ExpectedCellLineBreak & "Line 2""" & ExpectedRowLineBreak &
                """R1C1" & ExpectedCellLineBreak & "Line 2""||""R1C2" & ExpectedCellLineBreak & "Line 2""" & ExpectedRowLineBreak &
                """R2C1" & ExpectedCellLineBreak & "Line 2""||""R2C2" & ExpectedCellLineBreak & "Line 2""" & ExpectedRowLineBreak
            csv = CompuMaster.Data.Csv.WriteDataTableToCsvTextString(t, True, CompuMaster.Data.Csv.WriteLineEncodings.RowBreakLf_CellLineBreakCr, "||", """"c, ".")
            Console.WriteLine(csv)
            Assert.AreEqual(ExpectedValue, csv, "Linebreaks=Lf/Cr")

            ExpectedRowLineBreak = System.Environment.NewLine
            Select Case ExpectedRowLineBreak
                Case ControlChars.Cr, ControlChars.CrLf
                    ExpectedCellLineBreak = ControlChars.Lf
                Case ControlChars.Lf
                    ExpectedCellLineBreak = ControlChars.Cr
                Case Else
                    Throw New NotImplementedException
            End Select
            ExpectedValue = """col1" & ExpectedCellLineBreak & "Line 2""||""col2" & ExpectedCellLineBreak & "Line 2""" & ExpectedRowLineBreak &
                """R1C1" & ExpectedCellLineBreak & "Line 2""||""R1C2" & ExpectedCellLineBreak & "Line 2""" & ExpectedRowLineBreak &
                """R2C1" & ExpectedCellLineBreak & "Line 2""||""R2C2" & ExpectedCellLineBreak & "Line 2""" & ExpectedRowLineBreak
            csv = CompuMaster.Data.Csv.WriteDataTableToCsvTextString(t, True, CompuMaster.Data.Csv.WriteLineEncodings.Auto, "||", """"c, ".")
            Console.WriteLine(csv)
            Assert.AreEqual(ExpectedValue, csv, "Linebreaks=Auto")
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

        Private Function SimpleSampleTableWithLineBreaksInCells(lineBreak As String) As DataTable
            Dim t As New DataTable("root" & lineBreak & "Line 2")
            t.Columns.Add("col1" & lineBreak & "Line 2")
            t.Columns.Add("col2" & lineBreak & "Line 2")
            Dim r As DataRow = t.NewRow
            r(0) = "R1C1" & lineBreak & "Line 2"
            r(1) = "R1C2" & lineBreak & "Line 2"
            t.Rows.Add(r)
            r = t.NewRow
            r(0) = "R2C1" & lineBreak & "Line 2"
            r(1) = "R2C2" & lineBreak & "Line 2"
            t.Rows.Add(r)
            Return t
        End Function

        Private Function SeveralTypesSampleTable() As DataTable
            Dim t As New DataTable("root")
            t.Columns.Add("col1")
            t.Columns.Add("col2")
            t.Columns.Add("dbl", GetType(Double))
            t.Columns.Add("dec", GetType(Decimal))
            t.Columns.Add("date", GetType(DateTime))
            t.Columns.Add("integer", GetType(Integer))
            't.Columns.Add("bytes", GetType(Byte())) 'not supported: stable/safe byte-array-to-string-conversion and back re-reading -> general CSV design issue?!
            Dim r As DataRow = t.NewRow
            r(0) = "R1C1"
            r(1) = "R1C2"
            r(2) = 1000.0001
            r(3) = 1000.0001D
            r(4) = Now
            r(5) = 1000
            'r(6) = New Byte() {AscW("-"c), AscW(","c), AscW("."c), AscW("-"c)}
            t.Rows.Add(r)
            r = t.NewRow
            r(0) = "R2C1"
            r(1) = "R2C2"
            r(2) = DBNull.Value
            r(3) = DBNull.Value
            r(4) = DBNull.Value
            r(5) = DBNull.Value
            'r(6) = DBNull.Value
            t.Rows.Add(r)
            r = t.NewRow
            r(0) = ""
            r(1) = Nothing
            r(2) = DBNull.Value
            r(3) = DBNull.Value
            r(4) = DBNull.Value
            r(5) = DBNull.Value
            'r(6) = Nothing
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

        Private Shared Sub WriteFactoryXxlDataTableToCsvTextStringAndReRead(t As DataTable)
            Dim TempFile As New CompuMaster.Test.Data.TemporaryTestFile(".csv")
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
            'Dim bom As String = System.Text.Encoding.UTF8.GetString(System.Text.Encoding.UTF8.GetPreamble())
            Dim csv As String
            csv = CompuMaster.Data.Csv.WriteDataTableToCsvTextString(t, True)
            Dim t2 As DataTable = CompuMaster.Data.Csv.ReadDataTableFromCsvString(csv, True)
            Dim csv2 As String
            Assert.AreEqual(t.Columns.Count, t2.Columns.Count) 'should be the very same
            Assert.AreEqual(t.Rows.Count, t2.Rows.Count) 'should be the very same
            csv2 = CompuMaster.Data.Csv.WriteDataTableToCsvTextString(t2, True)
            Assert.AreEqual(csv, csv2) 'should be the very same
        End Sub

        Private Function DataTypesOfColumns(columns As System.Data.DataColumnCollection) As String()
            Dim Result As New System.Collections.Generic.List(Of String)
            For Each column As System.Data.DataColumn In columns
                Result.Add(column.DataType.ToString)
            Next
            Return Result.ToArray
        End Function

        <Test, Ignore("Column data type auto-detection not yet implemented in CSV read")> Sub WriteDataTableToCsvTextStringAndReReadAndReWriteWithoutChanges_DataTypesReDetection()
            Dim t As DataTable = SeveralTypesSampleTable()
            Dim initialColumnDataTypes As String() = DataTypesOfColumns(t.Columns)
            Console.WriteLine("Initial column data types: " & String.Join(", ", initialColumnDataTypes) & ControlChars.CrLf & ControlChars.CrLf)
            'Dim bom As String = System.Text.Encoding.UTF8.GetString(System.Text.Encoding.UTF8.GetPreamble())
            Dim csv As String
            csv = CompuMaster.Data.Csv.WriteDataTableToCsvTextString(t, True)
            Console.WriteLine("Initial table" & ControlChars.CrLf & ControlChars.CrLf)
            Console.WriteLine(csv)
            Dim t2 As DataTable = CompuMaster.Data.Csv.ReadDataTableFromCsvString(csv, True)
            Dim csv2 As String
            Assert.AreEqual(t.Columns.Count, t2.Columns.Count) 'should be the very same
            Assert.AreEqual(t.Rows.Count, t2.Rows.Count) 'should be the very same
            csv2 = CompuMaster.Data.Csv.WriteDataTableToCsvTextString(t2, True)
            Console.WriteLine("Reread column data types: " & String.Join(", ", DataTypesOfColumns(t2.Columns)) & ControlChars.CrLf & ControlChars.CrLf)
            Console.WriteLine("Rewritten table" & ControlChars.CrLf & ControlChars.CrLf)
            Console.WriteLine(csv2)
            Assert.AreEqual(csv, csv2) 'should be the very same
            Assert.AreEqual(initialColumnDataTypes, DataTypesOfColumns(t2.Columns))
        End Sub

        <Test> Sub WriteDataTableToCsvTextStringAndReReadAndReWriteWithoutChanges_ComlumnSeparatorInjectionTrials()
            Dim t As DataTable = SeveralTypesSampleTable()
            Dim initialColumnDataTypes As String() = DataTypesOfColumns(t.Columns)
            Console.WriteLine("Column data types: " & String.Join(", ", initialColumnDataTypes) & ControlChars.CrLf & ControlChars.CrLf)
            'Dim bom As String = System.Text.Encoding.UTF8.GetString(System.Text.Encoding.UTF8.GetPreamble())
            Dim csv As String
            csv = CompuMaster.Data.Csv.WriteDataTableToCsvTextString(t, True)
            Console.WriteLine("Current culture formatted table" & ControlChars.CrLf & ControlChars.CrLf)
            Console.WriteLine(csv)
            Assert.IsFalse(csv.Contains("""1000,0001"""))
            Dim t2 As DataTable = SeveralTypesSampleTable()
            Dim csv2 As String
            Assert.AreEqual(t.Columns.Count, t2.Columns.Count) 'should be the very same
            Assert.AreEqual(t.Rows.Count, t2.Rows.Count) 'should be the very same
            csv2 = CompuMaster.Data.Csv.WriteDataTableToCsvTextString(t2, True, ","c, """"c, ","c)
            Console.WriteLine("Decimal separator equalling to column separator table" & ControlChars.CrLf & ControlChars.CrLf)
            Console.WriteLine(csv2)
            Assert.IsTrue(csv2.Contains("""1000,0001"""))
            Assert.AreNotEqual(csv, csv2) 'should be the very same
        End Sub

        <Test> Sub WriteDataTableToCsvTextStringAndReReadAndReWriteWithoutChangesWithLineBreaksExcelCompatible()
            Dim t As DataTable = CellWithLineBreaksAndSpecialCharsSampleTable()
            'Dim bom As String = System.Text.Encoding.UTF8.GetString(System.Text.Encoding.UTF8.GetPreamble())
            Dim csv As String
            csv = CompuMaster.Data.Csv.WriteDataTableToCsvTextString(t, True)
            Console.WriteLine(csv)
            Dim t2 As DataTable = CompuMaster.Data.Csv.ReadDataTableFromCsvString(csv, True, CompuMaster.Data.Csv.ReadLineEncodings.RowBreakCrLf_CellLineBreakLf, CompuMaster.Data.Csv.ReadLineEncodingAutoConversion.AutoConvertLineBreakToLf)
            'DataTables.AssertTables(t, t2, "Comparison t vs t2")
            'DataTables.AssertTables(CloneTableWithSearchAndReplaceOnStrings(t, ControlChars.CrLf, vbLf), t2, "Comparison t vs t2")
            Assert.Ignore("TODO: Assertion must be re-checked")
            DataTablesTest.AssertTables(CloneTableWithSearchAndReplaceOnStrings(CloneTableWithSearchAndReplaceOnStrings(t, ControlChars.CrLf, vbLf), vbCr, vbLf), t2, "Comparison t vs t2")
            'DataTables.AssertTables(CloneTableWithSearchAndReplaceOnStrings(CloneTableWithSearchAndReplaceOnStrings(t, ControlChars.CrLf, vbLf), vbCr, vbLf), CloneTableWithSearchAndReplaceOnStrings(t2, ControlChars.CrLf, vbLf), "Comparison t vs t2")
            Dim csv2 As String
            Assert.AreEqual(t.Columns.Count, t2.Columns.Count) 'should be the very same
            Assert.AreEqual(t.Rows.Count, t2.Rows.Count) 'should be the very same - except line breaks in cells haven't been converted correctly
            csv2 = CompuMaster.Data.Csv.WriteDataTableToCsvTextString(t2, True)
            Console.WriteLine(csv2)
            Assert.AreEqual(csv, csv2) 'should be the very same
        End Sub

        <Test> Sub CsvEncode()
            Assert.AreEqual("R1C1""""" & vbLf, CompuMaster.Data.CsvTables.CsvTools.CsvEncode("R1C1""" & vbLf, """"c, CompuMaster.Data.Csv.WriteLineEncodings.RowBreakCr_CellLineBreakLf))
            Assert.AreEqual("R1C1""""" & vbLf & vbLf, CompuMaster.Data.CsvTables.CsvTools.CsvEncode("R1C1""" & ControlChars.CrLf & vbCr, """"c, CompuMaster.Data.Csv.WriteLineEncodings.RowBreakCr_CellLineBreakLf))
            Assert.AreEqual("R1C1""""""""" & vbLf & vbLf, CompuMaster.Data.CsvTables.CsvTools.CsvEncode("R1C1""""" & ControlChars.CrLf & vbCr, """"c, CompuMaster.Data.Csv.WriteLineEncodings.RowBreakCr_CellLineBreakLf))
            Assert.AreEqual("R1C1""""" & vbLf & vbLf, CompuMaster.Data.CsvTables.CsvTools.CsvEncode("R1C1""" & ControlChars.CrLf & vbCr, """"c, CompuMaster.Data.Csv.WriteLineEncodings.RowBreakCrLf_CellLineBreakLf))
            Assert.AreEqual("R1C1""""""""" & vbLf & vbLf, CompuMaster.Data.CsvTables.CsvTools.CsvEncode("R1C1""""" & ControlChars.CrLf & vbCr, """"c, CompuMaster.Data.Csv.WriteLineEncodings.RowBreakCrLf_CellLineBreakLf))
        End Sub

        Private Function CloneTableWithSearchAndReplaceOnStrings(table As DataTable, searchValue As String, replaceValue As String) As DataTable
            Dim Result As DataTable = CompuMaster.Data.DataTables.CreateDataTableClone(table)
            For MyCounter As Integer = 0 To Result.Columns.Count - 1
                If Result.Columns(MyCounter).DataType Is GetType(String) Then
                    For MyRowCounter As Integer = 0 To Result.Rows.Count - 1
                        If IsDBNull(Result.Rows(MyRowCounter)(MyCounter)) = False AndAlso Result.Rows(MyRowCounter)(MyCounter) IsNot Nothing Then
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

        <Ignore("NotYetImplementedCompletely")>
        <Test> Sub SupportReadLineBreakCrLfWithCellBreakCrLf()
            Dim Result As DataTable = CompuMaster.Data.Csv.ReadDataTableFromCsvFile(AssemblyTestEnvironment.TestFileAbsolutePath("testfiles\test_linebreak_crlf_cellbreak_crlf.csv"), True, CompuMaster.Data.Csv.ReadLineEncodings.RowBreakCrLfOrCrOrLf_CellLineBreakCrLfOrCrOrLf, CompuMaster.Data.Csv.ReadLineEncodingAutoConversion.NoAutoConversion, "UTF-8", ";"c, """"c, False, True)

            'Test output
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(Result, CompuMaster.Data.ConvertToPlainTextTableOptions.InlineBordersLayoutAnsi))

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

        <Test> Public Sub ReadDataTableFromLexOfficeCsv()
            Dim TestFile As String = AssemblyTestEnvironment.TestFileAbsolutePath(System.IO.Path.Combine("testfiles", "lexoffice.csv"))
            Dim StartLine As Integer = 0
            System.Console.WriteLine("TestFile=" & TestFile)
            System.Console.WriteLine("StartLine=" & TestFile)

            Dim CsvCulture As System.Globalization.CultureInfo = System.Globalization.CultureInfo.CreateSpecificCulture("de-DE")
            Dim FileEncoding As System.Text.Encoding = System.Text.Encoding.UTF8
            Dim FileEncodingName As String = "UTF-8"
            Dim dt As DataTable

            'CSV-File
            dt = CompuMaster.Data.Csv.ReadDataTableFromCsvFile(
                New CsvFileOptions(TestFile, FileEncoding),
                New CsvReadOptionsDynamicColumnSize(True, StartLine, CsvCulture, """"c, True) With {.RecognizeBackslashEscapes = True})
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
            Assert.AreEqual(10, dt.Columns.Count)
            Assert.AreEqual(2218, dt.Rows.Count)
            Assert.AreEqual("Konto", dt.Columns(0).ColumnName)
            Assert.AreEqual("200", dt.Rows(0)(0))
            Assert.AreEqual("Vhddkqrugy Sfjbnrt iyu Tgyropiam", dt.Rows(2)(1))
            Assert.AreEqual(DBNull.Value, dt.Rows(2)(9))

            dt = CompuMaster.Data.Csv.ReadDataTableFromCsvFile(
                New CsvFileOptions(TestFile, FileEncoding),
                New CsvReadOptionsDynamicColumnSize(
                    True, StartLine,
                    CompuMaster.Data.Csv.ReadLineEncodings.RowBreakCrLfOrCr_CellLineBreakLf, CompuMaster.Data.Csv.ReadLineEncodingAutoConversion.NoAutoConversion,
                    CsvCulture, """"c, True) With {.RecognizeBackslashEscapes = True})
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
            Assert.AreEqual(10, dt.Columns.Count)
            Assert.AreEqual(2218, dt.Rows.Count)
            Assert.AreEqual("Konto", dt.Columns(0).ColumnName)
            Assert.AreEqual("200", dt.Rows(0)(0))
            Assert.AreEqual("Vhddkqrugy Sfjbnrt iyu Tgyropiam", dt.Rows(2)(1))
            Assert.AreEqual(DBNull.Value, dt.Rows(2)(9))

            dt = CompuMaster.Data.Csv.ReadDataTableFromCsvFile(
                New CsvFileOptions(TestFile, FileEncoding),
                New CsvReadOptionsDynamicColumnSize(
                    True, StartLine,
                    CsvCulture, recognizeTextBy:=""""c, True) With {.ColumnSeparator = ";"c, .RecognizeBackslashEscapes = True})
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
            Assert.AreEqual(10, dt.Columns.Count)
            Assert.AreEqual(2218, dt.Rows.Count)
            Assert.AreEqual("Konto", dt.Columns(0).ColumnName)
            Assert.AreEqual("200", dt.Rows(0)(0))
            Assert.AreEqual("Vhddkqrugy Sfjbnrt iyu Tgyropiam", dt.Rows(2)(1))
            Assert.AreEqual(DBNull.Value, dt.Rows(2)(9))

            dt = CompuMaster.Data.Csv.ReadDataTableFromCsvFile(
                New CsvFileOptions(TestFile, FileEncoding),
                New CsvReadOptionsDynamicColumnSize(
                    True, StartLine,
                    CompuMaster.Data.Csv.ReadLineEncodings.RowBreakCrLfOrCr_CellLineBreakLf, CompuMaster.Data.Csv.ReadLineEncodingAutoConversion.NoAutoConversion,
                    CsvCulture, recognizeTextBy:=""""c, True) With {.ColumnSeparator = ";"c, .RecognizeBackslashEscapes = True})
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
            Assert.AreEqual(10, dt.Columns.Count)
            Assert.AreEqual(2218, dt.Rows.Count)
            Assert.AreEqual("Konto", dt.Columns(0).ColumnName)
            Assert.AreEqual("200", dt.Rows(0)(0))
            Assert.AreEqual("Vhddkqrugy Sfjbnrt iyu Tgyropiam", dt.Rows(2)(1))
            Assert.AreEqual(DBNull.Value, dt.Rows(2)(9))

            'CSV-String
            Dim CsvData As String = System.IO.File.ReadAllText(TestFile, FileEncoding)
            dt = CompuMaster.Data.Csv.ReadDataTableFromCsvString(
                CsvData,
                New CsvReadOptionsDynamicColumnSize(True, StartLine, CsvCulture, """"c, True) With {.RecognizeBackslashEscapes = True})
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
            Assert.AreEqual(10, dt.Columns.Count)
            Assert.AreEqual(2218, dt.Rows.Count)
            Assert.AreEqual("Konto", dt.Columns(0).ColumnName)
            Assert.AreEqual("200", dt.Rows(0)(0))
            Assert.AreEqual("Vhddkqrugy Sfjbnrt iyu Tgyropiam", dt.Rows(2)(1))
            Assert.AreEqual(DBNull.Value, dt.Rows(2)(9))

            dt = CompuMaster.Data.Csv.ReadDataTableFromCsvString(
                CsvData,
                New CsvReadOptionsDynamicColumnSize(
                    True, StartLine,
                    CsvCulture, """"c, True) With {.RecognizeBackslashEscapes = True})
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
            Assert.AreEqual(10, dt.Columns.Count)
            Assert.AreEqual(2218, dt.Rows.Count)
            Assert.AreEqual("Konto", dt.Columns(0).ColumnName)
            Assert.AreEqual("200", dt.Rows(0)(0))
            Assert.AreEqual("Vhddkqrugy Sfjbnrt iyu Tgyropiam", dt.Rows(2)(1))
            Assert.AreEqual(DBNull.Value, dt.Rows(2)(9))

            dt = CompuMaster.Data.Csv.ReadDataTableFromCsvString(
                CsvData,
                New CsvReadOptionsDynamicColumnSize(
                    True, StartLine,
                    CompuMaster.Data.Csv.ReadLineEncodings.RowBreakCrLfOrCr_CellLineBreakLf, CompuMaster.Data.Csv.ReadLineEncodingAutoConversion.NoAutoConversion,
                    CsvCulture, """"c, True) With {.RecognizeBackslashEscapes = True})
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
            Assert.AreEqual(10, dt.Columns.Count)
            Assert.AreEqual(2218, dt.Rows.Count)
            Assert.AreEqual("Konto", dt.Columns(0).ColumnName)
            Assert.AreEqual("200", dt.Rows(0)(0))
            Assert.AreEqual("Vhddkqrugy Sfjbnrt iyu Tgyropiam", dt.Rows(2)(1))
            Assert.AreEqual(DBNull.Value, dt.Rows(2)(9))

            dt = CompuMaster.Data.Csv.ReadDataTableFromCsvString(
                CsvData,
                New CsvReadOptionsDynamicColumnSize(
                    True, StartLine,
                    CsvCulture, recognizeTextBy:=""""c, True) With {.ColumnSeparator = ";"c, .RecognizeBackslashEscapes = True})
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
            Assert.AreEqual(10, dt.Columns.Count)
            Assert.AreEqual(2218, dt.Rows.Count)
            Assert.AreEqual("Konto", dt.Columns(0).ColumnName)
            Assert.AreEqual("200", dt.Rows(0)(0))
            Assert.AreEqual("Vhddkqrugy Sfjbnrt iyu Tgyropiam", dt.Rows(2)(1))
            Assert.AreEqual(DBNull.Value, dt.Rows(2)(9))

            dt = CompuMaster.Data.Csv.ReadDataTableFromCsvString(
                CsvData,
                New CsvReadOptionsDynamicColumnSize(
                    True, StartLine,
                    CompuMaster.Data.Csv.ReadLineEncodings.RowBreakCrLfOrCr_CellLineBreakLf, CompuMaster.Data.Csv.ReadLineEncodingAutoConversion.NoAutoConversion,
                    CsvCulture, recognizeTextBy:=""""c, True) With {.ColumnSeparator = ";"c, .RecognizeBackslashEscapes = True})
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
            Assert.AreEqual(10, dt.Columns.Count)
            Assert.AreEqual(2218, dt.Rows.Count)
            Assert.AreEqual("Konto", dt.Columns(0).ColumnName)
            Assert.AreEqual("200", dt.Rows(0)(0))
            Assert.AreEqual("Vhddkqrugy Sfjbnrt iyu Tgyropiam", dt.Rows(2)(1))
            Assert.AreEqual(DBNull.Value, dt.Rows(2)(9))

        End Sub

        Private Function TableCopyWithRowNumbering(table As DataTable, startNumber As Int32) As DataTable
            Dim dtWithRowNumbering As DataTable = table.Copy
            CompuMaster.Data.DataTables.AddOrUpdateRowNumbering(dtWithRowNumbering, "Row-No", startNumber)
            Return dtWithRowNumbering
        End Function

    End Class
#Enable Warning CA1822 ' Member als statisch markieren

End Namespace