﻿Option Explicit On
Option Strict On

Imports NUnit.Framework
Imports OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
Imports System.Data

Namespace CompuMaster.Test.Data

#Disable Warning CA1822 ' Member als statisch markieren
    <TestFixture(Category:="DataTables")> Public Class DataTablesTest

#Region "TestComparisons"
        Public Shared Sub AssertTables(table1 As DataTable, table2 As DataTable, assertionTitle As String)
            Assert.AreEqual(table1.Columns.Count, table2.Columns.Count, assertionTitle & ": Column count must be equal")
            Assert.AreEqual(table1.Rows.Count, table2.Rows.Count, assertionTitle & ": Row count must be equal")
            For MyCounter As Integer = 0 To table1.Columns.Count - 1
                Assert.AreEqual(table1.Columns(MyCounter).DataType, table2.Columns(MyCounter).DataType, assertionTitle & ": DataType must be equal for column index " & MyCounter)
            Next
            For MyCounter As Integer = 0 To table1.Columns.Count - 1
                For MyRowCounter As Integer = 0 To table1.Rows.Count - 1
                    Assert.AreEqual(table1.Rows(MyRowCounter)(MyCounter), table2.Rows(MyRowCounter)(MyCounter), assertionTitle & ": Cell value must be equal for row index " & MyRowCounter & ", column index " & MyCounter)
                Next
            Next
        End Sub
#End Region

#Region "Test data"
        Private Function TestTable1() As DataTable
            Dim Result As New DataTable("test1")
            Result.Columns.Add("ID", GetType(Integer))
            Result.Columns.Add("Value", GetType(String))
            Dim newRow As DataRow
            newRow = Result.NewRow
            newRow(0) = 1
            newRow(1) = "Hello world!"
            Result.Rows.Add(newRow)
            newRow = Result.NewRow
            newRow(0) = 2
            newRow(1) = "Gotcha!"
            Result.Rows.Add(newRow)
            newRow = Result.NewRow
            newRow(0) = 3
            newRow(1) = "Hello world!"
            Result.Rows.Add(newRow)
            newRow = Result.NewRow
            newRow(0) = 4
            newRow(1) = "Not a duplicate"
            Result.Rows.Add(newRow)
            newRow = Result.NewRow
            newRow(0) = 5
            newRow(1) = "Hello world!"
            Result.Rows.Add(newRow)
            newRow = Result.NewRow
            newRow(0) = 6
            newRow(1) = "GOTCHA!"
            Result.Rows.Add(newRow)
            newRow = Result.NewRow
            newRow(0) = 7
            newRow(1) = "Gotcha!"
            Result.Rows.Add(newRow)
            Return Result
        End Function

        Private Function TestTable2() As DataTable
            Dim file As String = AssemblyTestEnvironment.TestFileAbsolutePath(System.IO.Path.Combine("testfiles", "Q&A.xlsx"))
            Dim dt As DataTable = CompuMaster.Data.XlsEpplus.ReadDataTableFromXlsFile(file) ', "Rund um das NT")
            Return dt
        End Function

        Private Function TestTable2WithDisabledFirstRowContentAsColumnName() As DataTable
            Dim file As String = AssemblyTestEnvironment.TestFileAbsolutePath(System.IO.Path.Combine("testfiles", "Q&A.xlsx"))
            Dim dt As DataTable = CompuMaster.Data.XlsEpplus.ReadDataTableFromXlsFile(file, False) ', "Rund um das NT")
            Return dt
        End Function

        Private Function TestTable2WithInvariantCultureInColumnNames() As DataTable
            Dim Result = TestTable2()
            For MyCounter As Integer = Result.Columns("Erläuterung").Ordinal + 1 To Result.Columns.Count - 1
                Result.Columns(MyCounter).ColumnName = Result.Columns(MyCounter).ColumnName.Replace(",", "").Replace(".", "")
            Next
            Return Result
        End Function
#End Region

        <Test()> Public Sub AddPrefixesToColumnNames()
            Dim dt As New DataTable
            dt.Columns.Add("hello")
            CompuMaster.Data.DataTables.AddPrefixesToColumnNames(dt, New Integer() {0}, "pref_")
            StringAssert.StartsWith("pref_", dt.Columns.Item(0).ColumnName)
        End Sub

        <Test()> Public Sub AddSufixesToColumnNames()
            Dim dt As New DataTable
            dt.Columns.Add("hello")
            CompuMaster.Data.DataTables.AddSuffixesToColumnNames(dt, New Integer() {0}, "_suf")
            StringAssert.EndsWith("_suf", dt.Columns.Item(0).ColumnName)
        End Sub

        <Test()> Public Sub ColumnIndex()
            Dim dt As New DataTable
            dt.Columns.Add("hey")
            Dim e As DataRow = dt.NewRow
            e.Item(0) = "D"
            dt.Rows.Add(e)
            Assert.AreEqual(0, CompuMaster.Data.DataTables.ColumnIndex(dt.Columns("hey")))
        End Sub

        <Test()> Public Sub CompareValuesOfUnknownType()
            Dim dt As New DataTable
            dt.Columns.Add("DummmyColumn")

            Dim e As Integer = 23
            Dim f As Integer = 24
            Assert.IsTrue(CompuMaster.Data.DataTables.CompareValuesOfUnknownType("D", "D", False))
            Assert.IsTrue(CompuMaster.Data.DataTables.CompareValuesOfUnknownType("D", "D", True))
            Assert.IsFalse(CompuMaster.Data.DataTables.CompareValuesOfUnknownType("D", "E", False))
            Assert.IsFalse(CompuMaster.Data.DataTables.CompareValuesOfUnknownType("D", "E", True))
            Assert.IsFalse(CompuMaster.Data.DataTables.CompareValuesOfUnknownType("D", "d", False))
            Assert.IsFalse(CompuMaster.Data.DataTables.CompareValuesOfUnknownType("É", "é", False))
            Assert.IsTrue(CompuMaster.Data.DataTables.CompareValuesOfUnknownType("É", "é", True))
            Assert.IsFalse(CompuMaster.Data.DataTables.CompareValuesOfUnknownType(e, f, False))
            Assert.IsTrue(CompuMaster.Data.DataTables.CompareValuesOfUnknownType("D", "D", StringComparison.Ordinal))
            Assert.IsTrue(CompuMaster.Data.DataTables.CompareValuesOfUnknownType("D", "D", StringComparison.OrdinalIgnoreCase))
            Assert.IsFalse(CompuMaster.Data.DataTables.CompareValuesOfUnknownType("D", "E", StringComparison.Ordinal))
            Assert.IsFalse(CompuMaster.Data.DataTables.CompareValuesOfUnknownType("D", "E", StringComparison.OrdinalIgnoreCase))
            Assert.IsFalse(CompuMaster.Data.DataTables.CompareValuesOfUnknownType("D", "d", StringComparison.Ordinal))
            Assert.IsTrue(CompuMaster.Data.DataTables.CompareValuesOfUnknownType("D", "d", StringComparison.OrdinalIgnoreCase))
            Assert.IsFalse(CompuMaster.Data.DataTables.CompareValuesOfUnknownType("É", "é", StringComparison.Ordinal))
            Assert.IsTrue(CompuMaster.Data.DataTables.CompareValuesOfUnknownType("É", "é", StringComparison.OrdinalIgnoreCase))
            Assert.IsFalse(CompuMaster.Data.DataTables.CompareValuesOfUnknownType(e, f, False))
        End Sub

        <Test()> Public Sub ConvertColumnValuesIntoArrayList()
            Dim dt As New DataTable
            dt.Columns.Add("DummyColumn")

            Dim row1 As DataRow = dt.NewRow
            row1.Item(0) = 23
            dt.Rows.Add(row1)

            Dim row2 As DataRow = dt.NewRow
            row2.Item(0) = 21
            dt.Rows.Add(row2)

            Dim row3 As DataRow = dt.NewRow
            row3.Item(0) = 23
            dt.Rows.Add(row3)

            Dim row4 As DataRow = dt.NewRow
            row4.Item(0) = 7
            dt.Rows.Add(row4)


            Dim list As ArrayList = CompuMaster.Data.DataTables.ConvertColumnValuesIntoArrayList(dt.Columns.Item(0))
            Assert.AreEqual(4, list.Count)
        End Sub

        <Test(), Ignore("Requires custom connection string to execute")> Public Sub ConvertDataReaderToDataSet()
            Dim MyConn As New System.Data.SqlClient.SqlConnection("SERVER=yoursqlserver;DATABASE=master;PWD=xxxxxxxxxxxxxxxxxxx;UID=sa")
            Dim MyCmd As IDbCommand = MyConn.CreateCommand()
            MyCmd.CommandType = CommandType.Text
            MyCmd.CommandText = "exec sp_databases; Exec sp_tables;"
            Dim Reader As System.Data.IDataReader = CompuMaster.Data.DataQuery.AnyIDataProvider.ExecuteReader(MyCmd, CompuMaster.Data.DataQuery.Automations.AutoOpenAndCloseAndDisposeConnection)
            Dim Data As DataSet = CompuMaster.Data.DataTables.ConvertDataReaderToDataSet(Reader)
            Assert.AreEqual(2, Data.Tables.Count)
        End Sub

#If Not CI_Build Then
        <Test()> Public Sub ConvertDataReaderToDataTable()
            Dim TestFile As String = AssemblyTestEnvironment.TestFileAbsolutePath(System.IO.Path.Combine("testfiles", "test_for_msaccess.mdb"))
            Dim MyConn As IDbConnection = CompuMaster.Data.DataQuery.Connections.MicrosoftAccessConnection(TestFile)
            Dim MyCmd As IDbCommand = MyConn.CreateCommand()
            MyCmd.CommandType = CommandType.Text
            MyCmd.CommandText = "SELECT IntegerLongValue, StringShort, StringMemo FROM [SeveralColumnTypesTest] ORDER BY ID"
            Dim Reader As System.Data.IDataReader = CompuMaster.Data.DataQuery.AnyIDataProvider.ExecuteReader(MyCmd, CompuMaster.Data.DataQuery.Automations.AutoOpenAndCloseAndDisposeConnection)
            Dim Data As DataTable = CompuMaster.Data.DataTables.ConvertDataReaderToDataTable(Reader, "mytablename")
            Assert.AreEqual("mytablename", Data.TableName)
            Assert.AreNotEqual(0, Data.Rows.Count)
        End Sub
#End If

        <Test(), NUnit.Framework.Ignore("NotYetImplemented")> Public Sub ConvertDatasetToXml()
            Throw New NotImplementedException
        End Sub

        <Test()> Public Sub ConvertDataTableToArrayList()
            Dim dt As New DataTable
            dt.Columns.Add("IntegerColumn", GetType(Integer))
            Dim rRow As DataRow = dt.NewRow
            Dim row2 As DataRow = dt.NewRow
            rRow.Item(0) = 23
            row2.Item(0) = 25

            dt.Rows.Add(rRow)
            dt.Rows.Add(row2)

            Dim e As ArrayList = CompuMaster.Data.DataTables.ConvertDataTableToArrayList(dt.Columns.Item(0))
            Assert.AreEqual(2, e.Count)
            Assert.AreEqual(23, e.Item(0))
            Assert.AreEqual(25, e.Item(1))
        End Sub

        <Test()> Public Sub ConvertDataTableToDictionaryEntryArray()
            Dim dt As New DataTable
            dt.Columns.Add("Key", GetType(String))
            dt.Columns.Add("Value", GetType(Integer))
            Dim rRow As DataRow = dt.NewRow
            rRow.Item(0) = "Number"
            rRow.Item(1) = 25
            dt.Rows.Add(rRow)
            Dim e As DictionaryEntry() = CompuMaster.Data.DataTables.ConvertDataTableToDictionaryEntryArray(dt)
            Assert.AreEqual(1, e.GetLength(0))
            Assert.AreEqual("Number", dt.Rows.Item(0).Item(0))
            Assert.AreEqual(25, dt.Rows.Item(0).Item(1))
        End Sub

        <Test()> Public Sub ConvertDataTableToHashtable()
            Dim dt As New DataTable
            dt.Columns.Add("Key")
            dt.Columns.Add("Value")
            dt.Columns(0).Unique = True

            Dim rRow As DataRow = dt.NewRow
            rRow.Item(0) = "a"
            rRow.Item(1) = "z"
            dt.Rows.Add(rRow)
            Dim ht As Hashtable = CompuMaster.Data.DataTables.ConvertDataTableToHashtable(dt)
            Assert.IsTrue(ht.ContainsKey("a"))
            Assert.IsTrue(ht.ContainsValue("z"))
        End Sub

        <Test(), NUnit.Framework.Ignore("NotYetImplemented")> Public Sub ConvertDataTableToWebFormsListItem()
            Dim dt As New DataTable
            dt.Columns.Add("First")
            dt.Columns.Add("second")

            Dim rRow As DataRow = dt.NewRow
            rRow.Item(0) = "FirstEntry"
            rRow.Item(1) = "SecondValue"
            dt.Rows.Add(rRow)
            '        CompuMaster.Data.DataTables.ConvertDataTableToWebFormsListItem(
            Throw New NotImplementedException
        End Sub

        <Test()> Public Sub ConvertDictionaryEntryArrayToDataTable()
            Dim de As DictionaryEntry()
            ReDim de(1)

            de(0).Key = "Hello"
            de(0).Value = "Bye"

            de(1).Key = "Fire"
            de(1).Value = "Water"

            Dim dt As DataTable = CompuMaster.Data.DataTables.ConvertDictionaryEntryArrayToDataTable(de)
            Assert.AreEqual(2, dt.Columns.Count())
            Assert.AreEqual(2, dt.Rows.Count())
            Assert.AreEqual("Hello", dt.Rows.Item(0).Item(0))
            Assert.AreEqual("Fire", dt.Rows.Item(1).Item(0))
            Assert.AreEqual("Bye", dt.Rows.Item(0).Item(1))
            Assert.AreEqual("Water", dt.Rows.Item(1).Item(1))
        End Sub

        <Test(), NUnit.Framework.Ignore("NotYetImplemented")> Public Sub ConvertICollectionToDataTable()
            Throw New NotImplementedException
        End Sub

        <Test()> <CodeAnalysis.SuppressMessage("Style", "IDE0028:Initialisierung der Sammlung vereinfachen", Justification:="<Ausstehend>")>
        Public Sub ConvertIDictionaryToDataTable()
            Dim dict As IDictionary = New System.Collections.Generic.Dictionary(Of String, String)()
            dict.Add("Berlin", "Germany")

            Dim dt As DataTable = CompuMaster.Data.DataTables.ConvertIDictionaryToDataTable(dict)
            Assert.AreEqual(1, dt.Rows.Count())
            Assert.AreEqual(2, dt.Columns.Count())

        End Sub

        <Test()> <CodeAnalysis.SuppressMessage("Style", "IDE0028:Initialisierung der Sammlung vereinfachen", Justification:="<Ausstehend>")>
        Public Sub ConvertNameValueCollectionToDataTable()
            Dim nvc As New System.Collections.Specialized.NameValueCollection
            nvc.Add("Berlin", "Germany")
            nvc.Add("Paris", "France")
            Dim dt As DataTable = CompuMaster.Data.DataTables.ConvertNameValueCollectionToDataTable(nvc)
            Assert.AreEqual(2, dt.Columns.Count())
            Assert.AreEqual(2, dt.Rows.Count())
        End Sub

        <Test()> Public Sub ConvertToHtmlTable()
            Dim NullString As String = Nothing

            Dim dt As New DataTable
            dt.Columns.Add("id", GetType(Integer))
            dt.Columns.Add("Hi", GetType(String))
            dt.Columns.Add("action", GetType(String))
            dt.Columns("action").Caption = "<strong>Action</strong>"
            Dim row As DataRow = dt.NewRow
            row.Item(0) = 23
            row.Item(1) = "Hello <strong>World</strong>"
            row.Item(2) = "<a hef=""/test/"">Test</a>"
            dt.Rows.Add(row)
            Dim RowArray As DataRow() = New DataRow() {row}
            Dim HtmlColumns As String() = New String() {"action"}

            '## Basic HTML output
            Dim ExpectedHtml As String, Html As String
            ExpectedHtml = CompuMaster.Data.DataTables.ConvertToHtmlTable(dt)
            Assert.IsNotNull(ExpectedHtml)
            Assert.IsNotEmpty(ExpectedHtml)
            Html = CompuMaster.Data.DataTables.ConvertToHtmlTable(dt.Rows, dt.TableName)
            Assert.AreEqual(ExpectedHtml, Html)
            Html = CompuMaster.Data.DataTables.ConvertToHtmlTable(RowArray, dt.TableName)
            Assert.AreEqual(ExpectedHtml, Html)

            '## Arguments list for DataTable
            Html = CompuMaster.Data.DataTables.ConvertToHtmlTable(dt, NullString, NullString, NullString)
            Assert.AreEqual(ExpectedHtml, Html)
            Assert.IsTrue(Html.Contains("Hello <strong>World</strong>"))
            Assert.IsTrue(Html.Contains("<strong>Action</strong>"))
            Assert.IsTrue(Html.Contains("<a hef=""/test/"">Test</a>"))

            '- test with HTML encoding
            Html = CompuMaster.Data.DataTables.ConvertToHtmlTable(dt, NullString, NullString, NullString, True)
            Assert.IsTrue(Html.Contains("Hello &lt;strong&gt;World&lt;/strong&gt;"))
            Assert.IsTrue(Html.Contains("&lt;strong&gt;Action&lt;/strong&gt;"))
            Assert.IsTrue(Html.Contains("&lt;a hef=&quot;/test/&quot;&gt;Test&lt;/a&gt;"))

            '- test with HTML encoding except defined columns
            Html = CompuMaster.Data.DataTables.ConvertToHtmlTable(dt, NullString, NullString, NullString, True, HtmlColumns)
            Assert.IsTrue(Html.Contains("Hello &lt;strong&gt;World&lt;/strong&gt;"))
            Assert.IsTrue(Html.Contains("<strong>Action</strong>"))
            Assert.IsTrue(Html.Contains("<a hef=""/test/"">Test</a>"))

            '## Arguments list for DataRowCollection
            Html = CompuMaster.Data.DataTables.ConvertToHtmlTable(dt.Rows, dt.TableName, NullString, NullString, NullString)
            Assert.AreEqual(ExpectedHtml, Html)
            Assert.IsTrue(Html.Contains("Hello <strong>World</strong>"))
            Assert.IsTrue(Html.Contains("<strong>Action</strong>"))
            Assert.IsTrue(Html.Contains("<a hef=""/test/"">Test</a>"))

            '- test with HTML encoding
            Html = CompuMaster.Data.DataTables.ConvertToHtmlTable(dt.Rows, dt.TableName, NullString, NullString, NullString, True)
            Assert.IsTrue(Html.Contains("Hello &lt;strong&gt;World&lt;/strong&gt;"))
            Assert.IsTrue(Html.Contains("&lt;strong&gt;Action&lt;/strong&gt;"))
            Assert.IsTrue(Html.Contains("&lt;a hef=&quot;/test/&quot;&gt;Test&lt;/a&gt;"))

            '- test with HTML encoding except defined columns
            Html = CompuMaster.Data.DataTables.ConvertToHtmlTable(dt, NullString, NullString, NullString, True, HtmlColumns)
            Assert.IsTrue(Html.Contains("Hello &lt;strong&gt;World&lt;/strong&gt;"))
            Assert.IsTrue(Html.Contains("<strong>Action</strong>"))
            Assert.IsTrue(Html.Contains("<a hef=""/test/"">Test</a>"))

            '## Arguments list for DataRow array
            Html = CompuMaster.Data.DataTables.ConvertToHtmlTable(RowArray, dt.TableName, NullString, NullString, NullString)
            Assert.AreEqual(ExpectedHtml, Html)
            Assert.IsTrue(Html.Contains("Hello <strong>World</strong>"))
            Assert.IsTrue(Html.Contains("<strong>Action</strong>"))
            Assert.IsTrue(Html.Contains("<a hef=""/test/"">Test</a>"))

            '- test with HTML encoding
            Html = CompuMaster.Data.DataTables.ConvertToHtmlTable(RowArray, dt.TableName, NullString, NullString, NullString, True)
            Assert.IsTrue(Html.Contains("Hello &lt;strong&gt;World&lt;/strong&gt;"))
            Assert.IsTrue(Html.Contains("&lt;strong&gt;Action&lt;/strong&gt;"))
            Assert.IsTrue(Html.Contains("&lt;a hef=&quot;/test/&quot;&gt;Test&lt;/a&gt;"))

            '- test with HTML encoding except defined columns
            Html = CompuMaster.Data.DataTables.ConvertToHtmlTable(dt, NullString, NullString, NullString, True, HtmlColumns)
            Assert.IsTrue(Html.Contains("Hello &lt;strong&gt;World&lt;/strong&gt;"))
            Assert.IsTrue(Html.Contains("<strong>Action</strong>"))
            Assert.IsTrue(Html.Contains("<a hef=""/test/"">Test</a>"))

        End Sub

        <Test()> Public Sub ConvertToPlainTextTable()
            Dim dt As New DataTable
            dt.Columns.Add("id", GetType(Integer))
            dt.Columns.Add("Hi", GetType(String))

            Dim row As DataRow = dt.NewRow

            row.Item(0) = 23
            row.Item(1) = "Hello World"
            dt.Rows.Add(row)


            Dim html As String = CompuMaster.Data.DataTables.ConvertToPlainTextTable(dt)
            Assert.IsNotEmpty(html)

            Dim dt2 As New DataTable
            dt2.Columns.Add("id", GetType(Integer))
            dt2.Columns.Add("Hi", GetType(String))

            Dim row2 As DataRow = dt2.NewRow
            row.Item(0) = 23
            row.Item(1) = "Hello"

            Dim row3 As DataRow = dt2.NewRow
            row.Item(0) = 21
            row.Item(1) = "hello"

            dt2.Rows.Add(row2)
            dt2.Rows.Add(row3)
            dt2.AcceptChanges()
            dt2.Rows(1).Delete()
            Dim html2 As String = CompuMaster.Data.DataTables.ConvertToPlainTextTable(dt2)
            Assert.IsNotEmpty(html2)

        End Sub

        <Test()> Public Sub ConvertToWikiTable()
            Dim dt As New DataTable
            dt.Columns.Add("id", GetType(Integer))
            dt.Columns.Add("Hi", GetType(String))

            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToWikiTable(dt))

            Dim row As DataRow = dt.NewRow
            row.Item(0) = 23
            row.Item(1) = "Hello World"
            dt.Rows.Add(row)

            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToWikiTable(dt))

            dt = TestTable2()
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToWikiTable(dt))

        End Sub

        <Test()>
        <CodeAnalysis.SuppressMessage("Style", "IDE0028:Initialisierung der Sammlung vereinfachen")>
        Public Sub ConvertToPlainTextTableFixedColumnWidths_DefaultStyles()
            'Prepare test table
            Dim dt As New DataTable
            dt.Columns.Add("id", GetType(Integer))
            dt.Columns.Add("text", GetType(String))
            Dim row As DataRow = dt.NewRow
            row.Item(0) = 23
            row.Item(1) = "Hello World"
            dt.Rows.Add(row)
            'Delegated custom formatting
            dt.Columns.Add("dict", GetType(System.Collections.Generic.Dictionary(Of String, String)))
            dt.Rows(0)("dict") = New System.Collections.Generic.Dictionary(Of String, String)
            Dim dict0 As New System.Collections.Generic.Dictionary(Of String, String)
            dict0.Add("key1", "value1")
            dict0.Add("key2", "value2")
            dt.Rows(0)("dict") = dict0
            'Add empty 2nd row
            dt.Rows.Add(dt.NewRow)

            'Prepare rows array
            Dim RowsArray = New System.Data.DataRow() {dt.Rows(0), dt.Rows(1)}

            'Run tests
            Dim StyleOptions As CompuMaster.Data.ConvertToPlainTextTableOptions
            Console.WriteLine("# Style demos")

            'CheckSuggestedColWidths 
            StyleOptions = New CompuMaster.Data.ConvertToPlainTextTableOptions
            StyleOptions.ColumnFormatting = AddressOf ConvertColumnToString
            Dim ColWidthsSuggested As Integer() = CompuMaster.Data.DataTables.SuggestColumnWidthsForFixedPlainTables(
                dt.Rows, dt,
                StyleOptions)
            Dim SuggestedWidthsPrintTable As New DataTable()
            For MyCounter As Integer = 0 To ColWidthsSuggested.Length - 1
                SuggestedWidthsPrintTable.Columns.Add(dt.Columns(MyCounter).ColumnName)
            Next
            SuggestedWidthsPrintTable.Rows.Add(SuggestedWidthsPrintTable.NewRow)
            For MyCounter As Integer = 0 To SuggestedWidthsPrintTable.Columns.Count - 1
                SuggestedWidthsPrintTable.Rows(0)(MyCounter) = ColWidthsSuggested(MyCounter).ToString
            Next
            Console.WriteLine()
            Console.WriteLine("## SuggestedWidths")
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(SuggestedWidthsPrintTable, StyleOptions))

            StyleOptions = New CompuMaster.Data.ConvertToPlainTextTableOptions
            StyleOptions.ColumnFormatting = AddressOf ConvertColumnToString
            Console.WriteLine()
            Console.WriteLine("## ConvertToPlainTextTableOptions-Default via RowsCollection")
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt, StyleOptions))
            Console.WriteLine()
            Console.WriteLine("## ConvertToPlainTextTableOptions-Default via RowsArray")
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(RowsArray, StyleOptions))

            StyleOptions = CompuMaster.Data.ConvertToPlainTextTableOptions.SimpleLayout
            StyleOptions.ColumnFormatting = AddressOf ConvertColumnToString
            Console.WriteLine()
            Console.WriteLine("## SimpleLayout via RowsCollection")
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt, StyleOptions))
            Console.WriteLine()
            Console.WriteLine("## SimpleLayout via RowsArray")
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(RowsArray, StyleOptions))

            StyleOptions = CompuMaster.Data.ConvertToPlainTextTableOptions.InlineBordersLayoutAnsi
            StyleOptions.ColumnFormatting = AddressOf ConvertColumnToString
            Console.WriteLine()
            Console.WriteLine("## InlineBordersLayout via RowsCollection")
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt, StyleOptions))
            Console.WriteLine()
            Console.WriteLine("## InlineBordersLayout via RowsArray")
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(RowsArray, StyleOptions))

            StyleOptions = CompuMaster.Data.ConvertToPlainTextTableOptions.InlineBordersLayoutNice
            StyleOptions.ColumnFormatting = AddressOf ConvertColumnToString
            Console.WriteLine()
            Console.WriteLine("## InlineBordersLayout via RowsCollection")
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt, StyleOptions))
            Console.WriteLine()
            Console.WriteLine("## InlineBordersLayout via RowsArray")
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(RowsArray, StyleOptions))

        End Sub

        <Test()>
        <CodeAnalysis.SuppressMessage("Style", "IDE0028:Initialisierung der Sammlung vereinfachen")>
        Public Sub ConvertToPlainTextTableFixedColumnWidths()
#Disable Warning BC40000 ' Typ oder Element ist veraltet
            Dim dt As New DataTable
            dt.Columns.Add("id", GetType(Integer))
            dt.Columns.Add("Hi", GetType(String))

            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt, 10))
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt, 5, 20))
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt, " :: ", " :: ", "=##=", "="c, "="c))

            Dim row As DataRow = dt.NewRow
            row.Item(0) = 23
            row.Item(1) = "Hello World"
            dt.Rows.Add(row)

            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt, " :: ", " :: ", "=##=", "="c, "="c))
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt, 10))
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt, 5, 20))

            'Delegated custom formatting
            dt.Columns.Add("dict", GetType(System.Collections.Generic.Dictionary(Of String, String)))
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt, 5, 20, "|", "|", "+", "="c, "-"c, AddressOf ConvertColumnToString))
            dt.Rows(0)("dict") = New System.Collections.Generic.Dictionary(Of String, String)
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt, 5, 20, "|", "|", "+", "="c, "-"c, AddressOf ConvertColumnToString))
            Dim dict0 As New System.Collections.Generic.Dictionary(Of String, String)
            dict0.Add("key1", "value1")
            dict0.Add("key2", "value2")
            dt.Rows(0)("dict") = dict0
            dt.Rows.Add(dt.NewRow)
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt, 5, 20, "|", "|", "+", "="c, "-"c, AddressOf ConvertColumnToString))

            Dim ConvertedPlainTextTable As String
            Dim Expected As String

            Expected =
                "id|Hi         |dict                                                                " & System.Environment.NewLine &
                "--+-----------+--------------------------------------------------------------------" & System.Environment.NewLine &
                "23|Hello World|System.Collections.Generic.Dictionary`2[System.String,System.String]" & System.Environment.NewLine
            ConvertedPlainTextTable = CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt.Rows(0))
            Console.WriteLine(ConvertedPlainTextTable)
            Assert.AreEqual(Expected, ConvertedPlainTextTable)

            Expected =
                "id|Hi|dict" & System.Environment.NewLine &
                "--+--+----" & System.Environment.NewLine &
                "  |  |    " & System.Environment.NewLine
            ConvertedPlainTextTable = CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt.Rows(1))
            Console.WriteLine(ConvertedPlainTextTable)
            Assert.AreEqual(Expected, ConvertedPlainTextTable)

            'Real data table: quiz questions
            dt = TestTable2()
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt, " :: ", " :: ", "=##=", "="c, "="c))
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt, 10))
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt, 5, 20))
#Enable Warning BC40000 ' Typ oder Element ist veraltet
        End Sub


        Private Shared Function OutputOptions(minimumColumnWidth As Integer?, rowNumbering As Boolean) As CompuMaster.Data.ConvertToPlainTextTableOptions
            Dim Result = CompuMaster.Data.ConvertToPlainTextTableOptions.SimpleLayout
            Result.MinimumColumnWidth = minimumColumnWidth
            Result.MaximumColumnWidth = 65535
            Result.RowNumbering = rowNumbering
            Return Result
        End Function

        <Test> Public Sub ConvertToPlainTextTableFixedColumnWidths_RowNumbering(<Values(False, True)> rowNumbering As Boolean)
            Dim dt = TestTable1()

            Dim ConvertedPlainTextTable As String
            Dim Expected As String

            ConvertedPlainTextTable = CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt, OutputOptions(2, rowNumbering))
            Console.WriteLine(ConvertedPlainTextTable)
            If rowNumbering Then
                Expected =
                    "# |ID|Value          " & System.Environment.NewLine &
                    "--+--+---------------" & System.Environment.NewLine &
                    "1 |1 |Hello world!   " & System.Environment.NewLine &
                    "2 |2 |Gotcha!        " & System.Environment.NewLine &
                    "3 |3 |Hello world!   " & System.Environment.NewLine &
                    "4 |4 |Not a duplicate" & System.Environment.NewLine &
                    "5 |5 |Hello world!   " & System.Environment.NewLine &
                    "6 |6 |GOTCHA!        " & System.Environment.NewLine &
                    "7 |7 |Gotcha!        " & System.Environment.NewLine
            Else
                Expected =
                    "ID|Value          " & System.Environment.NewLine &
                    "--+---------------" & System.Environment.NewLine &
                    "1 |Hello world!   " & System.Environment.NewLine &
                    "2 |Gotcha!        " & System.Environment.NewLine &
                    "3 |Hello world!   " & System.Environment.NewLine &
                    "4 |Not a duplicate" & System.Environment.NewLine &
                    "5 |Hello world!   " & System.Environment.NewLine &
                    "6 |GOTCHA!        " & System.Environment.NewLine &
                    "7 |Gotcha!        " & System.Environment.NewLine
            End If
            Assert.AreEqual(Expected, ConvertedPlainTextTable)

            ConvertedPlainTextTable = CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt, OutputOptions(3, rowNumbering))
            Console.WriteLine(ConvertedPlainTextTable)
            If rowNumbering Then
                Expected =
                    "#  |ID |Value          " & System.Environment.NewLine &
                    "---+---+---------------" & System.Environment.NewLine &
                    "1  |1  |Hello world!   " & System.Environment.NewLine &
                    "2  |2  |Gotcha!        " & System.Environment.NewLine &
                    "3  |3  |Hello world!   " & System.Environment.NewLine &
                    "4  |4  |Not a duplicate" & System.Environment.NewLine &
                    "5  |5  |Hello world!   " & System.Environment.NewLine &
                    "6  |6  |GOTCHA!        " & System.Environment.NewLine &
                    "7  |7  |Gotcha!        " & System.Environment.NewLine
            Else
                Expected =
                    "ID |Value          " & System.Environment.NewLine &
                    "---+---------------" & System.Environment.NewLine &
                    "1  |Hello world!   " & System.Environment.NewLine &
                    "2  |Gotcha!        " & System.Environment.NewLine &
                    "3  |Hello world!   " & System.Environment.NewLine &
                    "4  |Not a duplicate" & System.Environment.NewLine &
                    "5  |Hello world!   " & System.Environment.NewLine &
                    "6  |GOTCHA!        " & System.Environment.NewLine &
                    "7  |Gotcha!        " & System.Environment.NewLine
            End If
            Assert.AreEqual(Expected, ConvertedPlainTextTable)

            ConvertedPlainTextTable = CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt, OutputOptions(New Integer?(), rowNumbering)) 'no minimum column width -> should be handled as min. 2 chars (hard-coded)
            Console.WriteLine(ConvertedPlainTextTable)
            If rowNumbering Then
                Expected =
                    "# |ID|Value          " & System.Environment.NewLine &
                    "--+--+---------------" & System.Environment.NewLine &
                    "1 |1 |Hello world!   " & System.Environment.NewLine &
                    "2 |2 |Gotcha!        " & System.Environment.NewLine &
                    "3 |3 |Hello world!   " & System.Environment.NewLine &
                    "4 |4 |Not a duplicate" & System.Environment.NewLine &
                    "5 |5 |Hello world!   " & System.Environment.NewLine &
                    "6 |6 |GOTCHA!        " & System.Environment.NewLine &
                    "7 |7 |Gotcha!        " & System.Environment.NewLine
            Else
                Expected =
                    "ID|Value          " & System.Environment.NewLine &
                    "--+---------------" & System.Environment.NewLine &
                    "1 |Hello world!   " & System.Environment.NewLine &
                    "2 |Gotcha!        " & System.Environment.NewLine &
                    "3 |Hello world!   " & System.Environment.NewLine &
                    "4 |Not a duplicate" & System.Environment.NewLine &
                    "5 |Hello world!   " & System.Environment.NewLine &
                    "6 |GOTCHA!        " & System.Environment.NewLine &
                    "7 |Gotcha!        " & System.Environment.NewLine
            End If
            Assert.AreEqual(Expected, ConvertedPlainTextTable)

        End Sub

        <CodeAnalysis.SuppressMessage("Major Code Smell", "S1172:Unused procedure parameters should be removed", Justification:="Required parameter to fit AddressOf method compatibility")>
        Private Function ConvertColumnToString(column As DataColumn, value As Object) As String
            If IsDBNull(value) Then
                Return Nothing
            ElseIf GetType(System.Collections.Generic.Dictionary(Of String, String)).IsInstanceOfType(value) Then
                Dim dict As System.Collections.Generic.Dictionary(Of String, String) = CType(value, System.Collections.Generic.Dictionary(Of String, String))
                Dim Result As New System.Text.StringBuilder
                For Each keyName As String In dict.Keys
                    If Result.Length <> 0 Then Result.AppendLine()
                    Result.Append(keyName)
                    Result.Append(":"c)
                    Result.Append(dict(keyName))
                Next
                Return Result.ToString
            Else
                Return CType(value, String)
            End If
        End Function

        <Test(), NUnit.Framework.Ignore("NotYetImplemented")> Public Sub ConvertXmlToDataset()
            Throw New NotImplementedException
        End Sub

        <Test> <Obsolete> Public Sub InsertColumnIntoClonedTable()
            Dim c As DataTable
            Dim t As DataTable = TestTable1()
            Assert.AreEqual(New String() {"ID", "Value"}, CompuMaster.Data.DataTables.AllColumnNames(t))

            c = CompuMaster.Data.DataTables.InsertColumnIntoClonedTable(t, 2, New DataColumn("Insert2"))
            Assert.AreEqual(New String() {"ID", "Value"}, CompuMaster.Data.DataTables.AllColumnNames(t), "Origin table must remain untouched")
            Assert.AreEqual(New String() {"ID", "Value", "Insert2"}, CompuMaster.Data.DataTables.AllColumnNames(c))

            c = CompuMaster.Data.DataTables.InsertColumnIntoClonedTable(c, 1, New DataColumn("Insert1"))
            Assert.AreEqual(New String() {"ID", "Insert1", "Value", "Insert2"}, CompuMaster.Data.DataTables.AllColumnNames(c))

            c = CompuMaster.Data.DataTables.InsertColumnIntoClonedTable(c, 0, New DataColumn("Insert0"))
            Assert.AreEqual(New String() {"Insert0", "ID", "Insert1", "Value", "Insert2"}, CompuMaster.Data.DataTables.AllColumnNames(c))
            Assert.AreEqual(New String() {"ID", "Value"}, CompuMaster.Data.DataTables.AllColumnNames(t), "Origin table must remain untouched")
        End Sub

        <Test> Public Sub InsertColumn()
            Dim t As DataTable
            t = TestTable1()
            Assert.AreEqual(New String() {"ID", "Value"}, CompuMaster.Data.DataTables.AllColumnNames(t))

            CompuMaster.Data.DataTables.InsertColumn(t, 2, New DataColumn("Insert2"))
            Assert.AreEqual(New String() {"ID", "Value", "Insert2"}, CompuMaster.Data.DataTables.AllColumnNames(t))

            CompuMaster.Data.DataTables.InsertColumn(t, 1, New DataColumn("Insert1"))
            Assert.AreEqual(New String() {"ID", "Insert1", "Value", "Insert2"}, CompuMaster.Data.DataTables.AllColumnNames(t))

            CompuMaster.Data.DataTables.InsertColumn(t, 0, New DataColumn("Insert0"))
            Assert.AreEqual(New String() {"Insert0", "ID", "Insert1", "Value", "Insert2"}, CompuMaster.Data.DataTables.AllColumnNames(t))
        End Sub

        <Test()> Public Sub CopyDataTableWithSubsetOfRows()
            Dim dt As New DataTable
            dt.Columns.Add("hi")
            dt.Columns.Add("hi2")
            dt.Rows.Add(New String() {"hi", "d"})
            dt.Rows.Add(New String() {"hix", "dix"})
            dt.Rows.Add(New String() {"l", "d"})

            Dim dt2 As DataTable = CompuMaster.Data.DataTables.CopyDataTableWithSubsetOfRows(dt, 2)
            Assert.AreEqual(2, dt2.Rows.Count())
            Assert.AreEqual("hi", dt.Rows.Item(0).Item(0))
            Assert.AreEqual("hix", dt.Rows.Item(1).Item(0))

            dt2 = CompuMaster.Data.DataTables.CopyDataTableWithSubsetOfRows(dt, 1, 2)
            Assert.AreEqual("l", dt2.Rows.Item(1).Item(0))

        End Sub

        <Test()> Public Sub CopyDataTableWithSubsetOfRows_SelectedRows()
            Dim dt As New DataTable
            dt.Columns.Add("hi")
            dt.Columns.Add("hi2")
            dt.Rows.Add(New String() {"hi", "d"})
            dt.Rows.Add(New String() {"hix", "dix"})
            dt.Rows.Add(New String() {"l", "d"})

            Dim dt2 As DataTable = CompuMaster.Data.DataTables.CopyDataTableWithSubsetOfRows(New DataRow() {dt.Rows(0), dt.Rows(1)})
            Assert.AreEqual(2, dt2.Rows.Count())
            Assert.AreEqual("hi", dt.Rows.Item(0).Item(0))
            Assert.AreEqual("hix", dt.Rows.Item(1).Item(0))

            dt2 = CompuMaster.Data.DataTables.CopyDataTableWithSubsetOfRows(New DataRow() {dt.Rows(1), dt.Rows(2)})
            Assert.AreEqual("l", dt2.Rows.Item(1).Item(0))

        End Sub

        <Test()> Public Sub CreateDataRowClone()
            Dim dt As New DataTable
            dt.Columns.Add("id", GetType(Integer))
            dt.Columns.Add("text", GetType(String))

            Dim row As DataRow = dt.NewRow
            row.Item(0) = 9
            row.Item(1) = "test"

            dt.Rows.Add(row)

            Dim row2 As DataRow
            row2 = CompuMaster.Data.DataTables.CreateDataRowClone(dt.Rows(0))

            Assert.AreEqual(9, row2.Item(0))
            Assert.AreEqual("test", row2.Item(1))
        End Sub

        <Test()> Public Sub CreateDataTableClone()
            Dim dt As New DataTable
            Dim dt2 As DataTable


            dt.Columns.Add("id", GetType(Integer))
            dt.Columns.Add("hi", GetType(String))

            Dim row As DataRow = dt.NewRow

            row.Item(0) = 7
            row.Item(1) = "Test"

            dt.Rows.Add(row)

            dt2 = CompuMaster.Data.DataTables.CreateDataTableClone(dt)
            Assert.AreEqual(1, dt2.Rows.Count())
            Assert.AreEqual(2, dt2.Columns.Count())

            Assert.AreEqual(7, dt2.Rows.Item(0).Item(0))
            Assert.AreEqual("Test", dt2.Rows.Item(0).Item(1))

            dt.Rows.Add(New Object() {8, "TestL2"})


            dt2 = CompuMaster.Data.DataTables.CreateDataTableClone(dt, "hi = 'TestL2'")
            Assert.AreEqual(1, dt2.Rows.Count())

            dt2 = CompuMaster.Data.DataTables.CreateDataTableClone(dt, CType(Nothing, String), "id DESC")
            Assert.AreEqual(8, dt2.Rows.Item(0).Item(0))


            dt.Rows.Add(New Object() {9, "L"})
            dt.Rows.Add(New Object() {10, "JJJ"})
            dt.Rows.Add(New Object() {11, "Lcx"})

            dt2 = CompuMaster.Data.DataTables.CreateDataTableClone(dt, Nothing, Nothing, 3)
            Assert.AreEqual(3, dt2.Rows.Count())

            Dim first As New DataTable
            Dim second As New DataTable

            first.Columns.Add("id")
            first.Columns.Add("whatever")
            first.Columns.Add("thirdCol")
            first.PrimaryKey = New DataColumn() {first.Columns(0)}
            first.Rows.Add(New Object() {23, "L", "X"})
            first.Rows.Add(New Object() {21, "p", "a"})
            first.Rows.Add(New Object() {1, "q", "z"})
            CompuMaster.Data.DataTables.CreateDataTableClone(first, second, Nothing, Nothing, Nothing, CompuMaster.Data.DataTables.RequestedRowChanges.DropExistingRowsInDestinationTableAndInsertNewRows, False,
                                                             CompuMaster.Data.DataTables.RequestedSchemaChangesForUnusedColumns.None, CompuMaster.Data.DataTables.RequestedSchemaChangesForExistingColumns.None,
                                                             CompuMaster.Data.DataTables.RequestedSchemaChangesForAdditionalColumns.Add)

            '  Assert.AreEqual(1, second.PrimaryKey.GetLength(0))




            'Testing the merge function
            Dim merge_source As New DataTable
            Dim merge_dest As New DataTable

            Dim oldSource As DataTable
            Dim oldDest As DataTable

            merge_source.Columns.Add("A", GetType(String))
            merge_source.Columns.Add("B", GetType(String))
            merge_source.Columns.Add("C", GetType(String))

            merge_dest.Columns.Add("A", GetType(String))
            merge_dest.Columns.Add("B", GetType(String))




            merge_source.PrimaryKey = New DataColumn() {merge_source.Columns(0)}
            merge_dest.PrimaryKey = New DataColumn() {merge_dest.Columns(0)}


            merge_dest.Rows.Add(New String() {"23", "hello"})
            merge_source.Rows.Add(New String() {"23", "hello!", "missing"})

            oldSource = merge_source
            oldDest = merge_dest

            CompuMaster.Data.DataTables.CreateDataTableClone(merge_source, merge_dest, Nothing, Nothing, Nothing, CompuMaster.Data.DataTables.RequestedRowChanges.KeepExistingRowsInDestinationTableAndAddRemoveUpdateChangedRows,
                                                             True, CompuMaster.Data.DataTables.RequestedSchemaChangesForUnusedColumns.None, CompuMaster.Data.DataTables.RequestedSchemaChangesForExistingColumns.None,
                                                             CompuMaster.Data.DataTables.RequestedSchemaChangesForAdditionalColumns.None)
            Assert.AreEqual(1, merge_dest.Rows.Count())
            Assert.AreEqual(2, merge_dest.Columns.Count())
            Assert.AreEqual("hello!", merge_dest.Rows(0).Item(1))

            merge_source = oldSource
            merge_dest = oldDest

            CompuMaster.Data.DataTables.CreateDataTableClone(merge_source, merge_dest, Nothing, Nothing, Nothing, CompuMaster.Data.DataTables.RequestedRowChanges.KeepExistingRowsInDestinationTableAndAddRemoveUpdateChangedRows,
                                                             True, CompuMaster.Data.DataTables.RequestedSchemaChangesForUnusedColumns.None, CompuMaster.Data.DataTables.RequestedSchemaChangesForExistingColumns.None,
                                                             CompuMaster.Data.DataTables.RequestedSchemaChangesForAdditionalColumns.Add)
            Assert.AreEqual(3, merge_dest.Columns.Count())
            Assert.AreEqual(1, merge_dest.Rows.Count())
            Assert.AreEqual("hello!", merge_dest.Rows(0).Item(1))
            Assert.AreEqual("missing", merge_dest.Rows(0).Item(2))


            Dim merge_source2 As New DataTable
            Dim merge_dest2 As New DataTable


            merge_source2.Columns.Add("A", GetType(String))
            merge_source2.Columns.Add("B", GetType(String))
            merge_source2.Columns.Add("C", GetType(String))

            merge_dest2.Columns.Add("A", GetType(String))
            merge_dest2.Columns.Add("B", GetType(String))




            merge_source2.PrimaryKey = New DataColumn() {merge_source2.Columns(0)}
            merge_dest2.PrimaryKey = New DataColumn() {merge_dest2.Columns(0)}


            merge_dest2.Rows.Add(New String() {"23", "hello"})
            merge_source2.Rows.Add(New String() {"23", "hello!", "missing"})

            CompuMaster.Data.DataTables.CreateDataTableClone(merge_source2, merge_dest2, Nothing, Nothing, Nothing, CompuMaster.Data.DataTables.RequestedRowChanges.KeepExistingRowsInDestinationTableAndAddRemoveUpdateChangedRows,
                                                             True, CompuMaster.Data.DataTables.RequestedSchemaChangesForUnusedColumns.None,
                                                             CompuMaster.Data.DataTables.RequestedSchemaChangesForExistingColumns.Update, CompuMaster.Data.DataTables.RequestedSchemaChangesForAdditionalColumns.None)
            Assert.AreEqual("hello!", merge_dest2.Rows(0).Item(1))
            Assert.AreEqual(2, merge_dest2.Columns.Count)


            Dim merge_source3 As New DataTable
            Dim merge_dest3 As New DataTable

            merge_source3.Columns.Add("A", GetType(Integer))
            merge_source3.Columns.Add("B", GetType(String))
            merge_source3.Columns.Add("C", GetType(String))

            merge_dest3.Columns.Add("A", GetType(Integer))
            merge_dest3.Columns.Add("B", GetType(String))
            merge_dest3.Columns.Add("C", GetType(String))

            merge_source3.PrimaryKey = New DataColumn() {merge_source3.Columns(0)}
            merge_dest3.PrimaryKey = New DataColumn() {merge_dest3.Columns(0)}

            merge_source3.Rows.Add(New Object() {1, "Text1", "KEEP-THIS-RECORD"})
            merge_source3.Rows.Add(New Object() {2, "Text2", "KEEP-THIS-RECORD"})
            merge_source3.Rows.Add(New Object() {3, "Text3", "KEEP-THIS-RECORD"})
            merge_source3.Rows.Add(New Object() {4, "Text4", "DEL-ME"})
            merge_source3.Rows.Add(New Object() {5, "Text5", "DEL-ME"})

            merge_dest3.Rows.Add(New Object() {1, "TextGone!", "K"})
            merge_dest3.Rows.Add(New Object() {9, "Not touched", "B"})
            merge_dest3.Rows.Add(New Object() {10, "...", "B"})
            merge_dest3.Rows.Add(New Object() {2, "TextX", "A"})
            merge_dest3.Rows.Add(New Object() {3, "..ax", "A"})

            merge_source3.AcceptChanges()
            merge_dest3.AcceptChanges()

            Dim Cloned_merge_dest3 As DataTable = CompuMaster.Data.DataTables.CreateDataTableClone(merge_dest3)
            Assert.AreEqual(TableStatistics(merge_dest3), TableStatistics(Cloned_merge_dest3))

            'Show table statistics
            System.Console.WriteLine("TABLE: merge_source3")
            System.Console.WriteLine(TableStatistics(merge_source3))
            System.Console.WriteLine(System.Environment.NewLine)
            System.Console.WriteLine(System.Environment.NewLine)
            System.Console.WriteLine("TABLE: merge_dest3")
            System.Console.WriteLine(TableStatistics(merge_dest3))
            System.Console.WriteLine(System.Environment.NewLine)
            System.Console.WriteLine(System.Environment.NewLine)
            System.Console.WriteLine("TABLE: Cloned_merge_dest3")
            System.Console.WriteLine(TableStatistics(Cloned_merge_dest3))
            System.Console.WriteLine(System.Environment.NewLine)
            System.Console.WriteLine(System.Environment.NewLine)

            'Show table content
            System.Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(merge_source3, "TABLE: merge_source3"))
            System.Console.WriteLine(System.Environment.NewLine)
            System.Console.WriteLine(System.Environment.NewLine)
            System.Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(merge_dest3, "TABLE: merge_dest3"))
            System.Console.WriteLine(System.Environment.NewLine)
            System.Console.WriteLine(System.Environment.NewLine)
            System.Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(Cloned_merge_dest3, "TABLE: Cloned_merge_dest3"))
            System.Console.WriteLine(System.Environment.NewLine)
            System.Console.WriteLine(System.Environment.NewLine)


            CompuMaster.Data.DataTables.CreateDataTableClone(merge_source3, merge_dest3, "C = 'KEEP-THIS-RECORD'", "A ASC", Nothing, CompuMaster.Data.DataTables.RequestedRowChanges.KeepExistingRowsInDestinationTableAndAddRemoveUpdateChangedRows,
                                                             True, CompuMaster.Data.DataTables.RequestedSchemaChangesForUnusedColumns.None,
                                                             CompuMaster.Data.DataTables.RequestedSchemaChangesForExistingColumns.None,
                                                             CompuMaster.Data.DataTables.RequestedSchemaChangesForAdditionalColumns.None)

            System.Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(Cloned_merge_dest3, "TABLE: Cloned_merge_dest3 #20"))
            System.Console.WriteLine(System.Environment.NewLine)
            System.Console.WriteLine(System.Environment.NewLine)


            Assert.AreEqual(5, merge_dest3.Rows.Count)
            Assert.AreEqual(3, merge_dest3.Columns.Count)

            'Checking sort after merge (only source)
            Assert.AreEqual(1, merge_dest3.Rows.Item(0).Item(0))
            Assert.AreEqual(9, merge_dest3.Rows.Item(1).Item(0))
            Assert.AreEqual(10, merge_dest3.Rows.Item(2).Item(0))
            'Assert.AreEqual(4, merge_dest3.Rows.Item(3).Item(0))
            'Assert.AreEqual(5, merge_dest3.Rows.Item(4)l.Item(0))
            Assert.AreEqual(2, merge_dest3.Rows.Item(3).Item(0))
            Assert.AreEqual(3, merge_dest3.Rows.Item(4).Item(0))

            'Ensure in general correct merge
            Assert.AreEqual("Text1", merge_dest3.Rows.Item(0).Item(1))
            Assert.AreEqual("Not touched", merge_dest3.Rows.Item(1).Item(1))
            Assert.AreEqual("...", merge_dest3.Rows.Item(2).Item(1))
            'Assert.AreEqual("Text4", merge_dest3.Rows.Item(3).Item(1))
            'Assert.AreEqual("Text5", merge_dest3.Rows.Item(4).Item(1))
            Assert.AreEqual("Text2", merge_dest3.Rows.Item(3).Item(1))
            Assert.AreEqual("Text3", merge_dest3.Rows.Item(4).Item(1))

            CompuMaster.Data.DataTables.CreateDataTableClone(merge_source3, merge_dest3, "C = 'KEEP-THIS-RECORD'", "A ASC", 3, CompuMaster.Data.DataTables.RequestedRowChanges.KeepExistingRowsInDestinationTableAndAddRemoveUpdateChangedRows,
                                                             True, CompuMaster.Data.DataTables.RequestedSchemaChangesForUnusedColumns.None,
                                                             CompuMaster.Data.DataTables.RequestedSchemaChangesForExistingColumns.None,
                                                             CompuMaster.Data.DataTables.RequestedSchemaChangesForAdditionalColumns.None)

            System.Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(Cloned_merge_dest3, "TABLE: Cloned_merge_dest3 #30"))
            System.Console.WriteLine(System.Environment.NewLine)
            System.Console.WriteLine(System.Environment.NewLine)

            Assert.AreEqual(3, merge_dest3.Rows.Count(), "RowCount check")

            Dim big As New DataTable
            Dim bigCopy As New DataTable
            Dim bigCopy2 As New DataTable

            big.ReadXml(AssemblyTestEnvironment.TestFileAbsolutePath(System.IO.Path.Combine("testfiles", "3000RowsTable.xml")))

            CompuMaster.Data.DataTables.CreateDataTableClone(big, bigCopy, "", "", Nothing, CompuMaster.Data.DataTables.RequestedRowChanges.DropExistingRowsInDestinationTableAndInsertNewRows, False,
                                                             CompuMaster.Data.DataTables.RequestedSchemaChangesForUnusedColumns.Remove, CompuMaster.Data.DataTables.RequestedSchemaChangesForExistingColumns.None,
                                                             CompuMaster.Data.DataTables.RequestedSchemaChangesForAdditionalColumns.Add)
            Assert.AreEqual(big.Columns.Count, bigCopy.Columns.Count)
            Assert.AreEqual(big.Rows.Count, bigCopy.Rows.Count)
            Assert.AreEqual(CompuMaster.Data.DataTables.ConvertToPlainTextTable(big.Rows, "Table"), CompuMaster.Data.DataTables.ConvertToPlainTextTable(bigCopy.Rows, "Table"))

            bigCopy.Rows.Item(0).Item(1) = 29
            bigCopy.Rows.Item(1020).Item(1) = 20
            bigCopy.Rows.Item(2323).Item(1) = 99
            bigCopy.Rows.Item(1000).Item(1) = 22
            bigCopy.Rows.Item(900).Item(1) = 55
            bigCopy.Rows.Item(bigCopy.Rows.Count - 5).Item(1) = 78

            'Show table statistics
            System.Console.WriteLine("TABLE: big")
            System.Console.WriteLine(TableStatistics(big))
            System.Console.WriteLine(System.Environment.NewLine)
            System.Console.WriteLine(System.Environment.NewLine)
            System.Console.WriteLine("TABLE: bigCopy")
            System.Console.WriteLine(TableStatistics(bigCopy))
            System.Console.WriteLine(System.Environment.NewLine)
            System.Console.WriteLine(System.Environment.NewLine)
            System.Console.WriteLine("TABLE: bigCopy2")
            System.Console.WriteLine(TableStatistics(bigCopy2))
            System.Console.WriteLine(System.Environment.NewLine)
            System.Console.WriteLine(System.Environment.NewLine)

            ''Show table content
            'System.Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(big, "TABLE: big"))
            'System.Console.WriteLine(System.Environment.NewLine)
            'System.Console.WriteLine(System.Environment.NewLine)
            'System.Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(big, "TABLE: bigCopy"))
            'System.Console.WriteLine(System.Environment.NewLine)
            'System.Console.WriteLine(System.Environment.NewLine)
            'System.Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(big, "TABLE: bigCopy2"))
            'System.Console.WriteLine(System.Environment.NewLine)
            'System.Console.WriteLine(System.Environment.NewLine)

            CompuMaster.Data.DataTables.CreateDataTableClone(bigCopy, bigCopy2, "", "", Nothing, CompuMaster.Data.DataTables.RequestedRowChanges.DropExistingRowsInDestinationTableAndInsertNewRows,
                                                             False, CompuMaster.Data.DataTables.RequestedSchemaChangesForUnusedColumns.None, CompuMaster.Data.DataTables.RequestedSchemaChangesForExistingColumns.None,
                                                             CompuMaster.Data.DataTables.RequestedSchemaChangesForAdditionalColumns.Add)

            Assert.AreEqual(bigCopy.Rows.Count(), bigCopy2.Rows.Count())
            Assert.AreEqual(bigCopy.Columns.Count, bigCopy2.Columns.Count)
            Assert.AreEqual(20, bigCopy2.Rows.Item(1020).Item(1))
            Assert.AreEqual(55, bigCopy2.Rows.Item(900).Item(1))
            Assert.AreEqual(CompuMaster.Data.DataTables.ConvertToPlainTextTable(bigCopy.Rows, "Table"), CompuMaster.Data.DataTables.ConvertToPlainTextTable(bigCopy2.Rows, "Table"))

            System.Console.WriteLine("TABLE: bigCopy2")
            System.Console.WriteLine(TableStatistics(bigCopy2))
            System.Console.WriteLine(System.Environment.NewLine)
            System.Console.WriteLine(System.Environment.NewLine)

            'TODO: all variations'

        End Sub

        Private Shared Function TableStatistics(table As DataTable) As String
            Dim Result As New System.Text.StringBuilder
            Result.AppendLine("TableName: " & table.TableName)
            Result.AppendLine("Rows")
            Result.AppendLine("* Count: " & table.Rows.Count)
            Result.AppendLine("PrimaryKey Columns")
            Result.AppendLine("* Count: " & table.PrimaryKey.Length)
            For MyCounter As Integer = 0 To table.PrimaryKey.Length - 1
                Dim Col As DataColumn = table.PrimaryKey(MyCounter)
                Result.AppendLine("* PK Column [" & MyCounter + 1 & "]: " & Col.ColumnName)
            Next
            Result.AppendLine("Columns")
            Result.AppendLine("* Count: " & table.Columns.Count)
            For MyCounter As Integer = 0 To table.Columns.Count - 1
                Dim Col As DataColumn = table.Columns(MyCounter)
                Result.AppendLine("* Column [" & MyCounter + 1 & "]: " & Col.DataType.FullName)
                Result.AppendLine("  * ColumnName: " & Col.ColumnName)
                Result.AppendLine("  * AllowDBNull: " & Col.AllowDBNull)
                Result.AppendLine("  * AutoIncreemnt: " & Col.AutoIncrement)
                Result.AppendLine("  * Caption: " & Col.Caption)
                Result.AppendLine("  * DefaultValue: " & CompuMaster.Data.Utils.ObjectNotNothingOrEmptyString(CompuMaster.Data.Utils.NoDBNull(Col.DefaultValue, "<NULL>")).ToString)
                Result.AppendLine("  * Expression: " & Col.Expression)
                Result.AppendLine("  * ReadOnly: " & Col.ReadOnly)
                Result.AppendLine("  * Unique: " & Col.Unique)
            Next
            Return Result.ToString
        End Function

        <Test()> Public Sub FindUniqueValues()
            Dim dt As New DataTable
            Dim row As DataRow
            dt.Columns.Add("Column Name")

            row = dt.NewRow
            row.Item(0) = "Sun"
            dt.Rows.Add(row)

            row = dt.NewRow
            row.Item(0) = "Moon"
            dt.Rows.Add(row)

            row = dt.NewRow
            row.Item(0) = "Sun"
            dt.Rows.Add(row)

            row = dt.NewRow
            row.Item(0) = "Saturn"
            dt.Rows.Add(row)

            row = dt.NewRow
            row.Item(0) = DBNull.Value
            dt.Rows.Add(row)

            row = dt.NewRow
            row.Item(0) = "Saturn"
            dt.Rows.Add(row)

            Dim list As ArrayList
            list = CompuMaster.Data.DataTables.FindUniqueValues(dt.Columns(0), False)
            Assert.AreEqual("Sun", list(0))
            Assert.AreEqual("Moon", list(1))
            Assert.AreEqual("Saturn", list(2))
            Assert.IsTrue(IsDBNull(list(3)), "Expected DbNull value")
            Assert.AreEqual(4, list.Count)
            list = CompuMaster.Data.DataTables.FindUniqueValues(dt.Columns(0), True)
            Assert.AreEqual("Sun", list(0))
            Assert.AreEqual("Moon", list(1))
            Assert.AreEqual("Saturn", list(2))
            Assert.AreEqual(3, list.Count)
            list = CompuMaster.Data.DataTables.FindUniqueValues(dt.Columns(0), False, New Object() {"Saturn", "Moon"})
            Assert.AreEqual("Sun", list(0))
            Assert.IsTrue(IsDBNull(list(1)), "Expected DbNull value")
            Assert.AreEqual(2, list.Count)
            list = CompuMaster.Data.DataTables.FindUniqueValues(dt.Columns(0), True, New Object() {"Saturn", "Moon"})
            Assert.AreEqual("Sun", list(0))
            Assert.AreEqual(1, list.Count)

            Dim stringList As List(Of String)
            stringList = CompuMaster.Data.DataTables.FindUniqueValues(Of String)(dt.Columns(0), False)
            Assert.AreEqual("Sun", stringList(0))
            Assert.AreEqual("Moon", stringList(1))
            Assert.AreEqual("Saturn", stringList(2))
            Assert.IsNull(stringList(3), "Expected DbNull->Null value")
            Assert.AreEqual(4, stringList.Count)
            stringList = CompuMaster.Data.DataTables.FindUniqueValues(Of String)(dt.Columns(0), True)
            Assert.AreEqual("Sun", stringList(0))
            Assert.AreEqual("Moon", stringList(1))
            Assert.AreEqual("Saturn", stringList(2))
            Assert.AreEqual(3, stringList.Count)
            stringList = CompuMaster.Data.DataTables.FindUniqueValues(Of String)(dt.Columns(0), False, New String() {"Saturn", "Moon"})
            Assert.AreEqual("Sun", stringList(0))
            Assert.IsNull(stringList(1), "Expected DbNull->Null value")
            Assert.AreEqual(2, stringList.Count)
            stringList = CompuMaster.Data.DataTables.FindUniqueValues(Of String)(dt.Columns(0), True, New String() {"Saturn", "Moon"})
            Assert.AreEqual("Sun", stringList(0))
            Assert.AreEqual(1, stringList.Count)

            'Add additional duplicates
            row = dt.NewRow
            row.Item(0) = DBNull.Value
            dt.Rows.Add(row)

            row = dt.NewRow
            row.Item(0) = ""
            dt.Rows.Add(row)

            row = dt.NewRow
            row.Item(0) = DBNull.Value
            dt.Rows.Add(row)

            row = dt.NewRow
            row.Item(0) = ""
            dt.Rows.Add(row)

            row = dt.NewRow
            row.Item(0) = "Sun"
            dt.Rows.Add(row)

            row = dt.NewRow
            row.Item(0) = "Moon"
            dt.Rows.Add(row)

            row = dt.NewRow
            row.Item(0) = "Sun"
            dt.Rows.Add(row)

            'Recheck again
            list = CompuMaster.Data.DataTables.FindUniqueValues(dt.Columns(0))
            Assert.AreEqual("Sun", list(0))
            Assert.AreEqual("Moon", list(1))
            Assert.AreEqual("Saturn", list(2))
            Assert.IsTrue(IsDBNull(list(3)), "Expected DbNull value")
            Assert.AreEqual("", list(4))
            Assert.AreEqual(5, list.Count)

            list = CompuMaster.Data.DataTables.FindUniqueValues(dt.Columns(0), True)
            Assert.AreEqual("Sun", list(0))
            Assert.AreEqual("Moon", list(1))
            Assert.AreEqual("Saturn", list(2))
            Assert.AreEqual("", list(3))
            Assert.AreEqual(4, list.Count)

            stringList = CompuMaster.Data.DataTables.FindUniqueValues(Of String)(dt.Columns(0))
            Assert.AreEqual("Sun", stringList(0))
            Assert.AreEqual("Moon", stringList(1))
            Assert.AreEqual("Saturn", stringList(2))
            Assert.IsNull(stringList(3), "Expected DbNull->Null value")
            Assert.AreEqual("", stringList(4))
            Assert.AreEqual(5, stringList.Count)

            stringList = CompuMaster.Data.DataTables.FindUniqueValues(Of String)(dt.Columns(0), True)
            Assert.AreEqual("Sun", stringList(0))
            Assert.AreEqual("Moon", stringList(1))
            Assert.AreEqual("Saturn", stringList(2))
            Assert.AreEqual("", stringList(3))
            Assert.AreEqual(4, stringList.Count)
        End Sub

        <Test()> Public Sub FindUniqueValues_TestStringListContainsNothingOrEmptyString()
            Dim L As List(Of String)

            '1st test suite: start with adding null/Nothing value
            L = New List(Of String)
            Assert.False(L.Contains(""))
            Assert.False(L.Contains(Nothing))
            L.Add(Nothing)
            Assert.False(L.Contains(""))
            Assert.True(L.Contains(Nothing))
            L.Add("")
            Assert.True(L.Contains(""))
            Assert.True(L.Contains(Nothing))

            '2nd test suite: start with adding ""/EmptyString value
            L = New List(Of String)
            Assert.False(L.Contains(""))
            Assert.False(L.Contains(Nothing))
            L.Add("")
            Assert.True(L.Contains(""))
            Assert.False(L.Contains(Nothing))
            L.Add(Nothing)
            Assert.True(L.Contains(""))
            Assert.True(L.Contains(Nothing))

        End Sub

        <Test()> Public Sub LookupUniqueColumnName()
            Dim dt As New DataTable
            dt.Columns.Add("Test")
            dt.Columns.Add("Test2")
            dt.Columns.Add("Test3")
            dt.Columns.Add("Test4")
            dt.Columns.Add("Test14")
            dt.Columns.Add("Test15")
            dt.Columns.Add("1")
            dt.Columns.Add("ID")
            dt.Columns.Add("ClientTable_")
            dt.Columns.Add("ClientTable_ID")
            dt.Columns.Add("ClientTable_ID2")
            dt.Columns.Add("ClientTable_ID23")
            Dim uname As String, lname As String
            lname = "Test"
            uname = CompuMaster.Data.DataTables.LookupUniqueColumnName(dt, lname)
            Console.WriteLine("New free unique column name: " & uname & " (lookup name: " & lname & ")")
            Assert.IsNotEmpty(uname)
            Assert.IsFalse(dt.Columns.Contains(uname))
            lname = "Test1"
            uname = CompuMaster.Data.DataTables.LookupUniqueColumnName(dt, lname)
            Console.WriteLine("New free unique column name: " & uname & " (lookup name: " & lname & ")")
            Assert.IsNotEmpty(uname)
            Assert.IsFalse(dt.Columns.Contains(uname))
            Assert.AreEqual("Test1", uname)
            lname = "Test2"
            uname = CompuMaster.Data.DataTables.LookupUniqueColumnName(dt, lname)
            Console.WriteLine("New free unique column name: " & uname & " (lookup name: " & lname & ")")
            Assert.IsNotEmpty(uname)
            Assert.IsFalse(dt.Columns.Contains(uname))
            lname = "test2"
            uname = CompuMaster.Data.DataTables.LookupUniqueColumnName(dt, lname)
            Console.WriteLine("New free unique column name: " & uname & " (lookup name: " & lname & ")")
            Assert.AreNotEqual("test2", uname)
            Assert.IsNotEmpty(uname)
            Assert.IsFalse(dt.Columns.Contains(uname))
            lname = "Test4"
            uname = CompuMaster.Data.DataTables.LookupUniqueColumnName(dt, lname)
            Console.WriteLine("New free unique column name: " & uname & " (lookup name: " & lname & ")")
            Assert.IsNotEmpty(uname)
            Assert.IsFalse(dt.Columns.Contains(uname))
            lname = "Test14"
            uname = CompuMaster.Data.DataTables.LookupUniqueColumnName(dt, lname)
            Console.WriteLine("New free unique column name: " & uname & " (lookup name: " & lname & ")")
            Assert.IsNotEmpty(uname)
            Assert.IsFalse(dt.Columns.Contains(uname))
            lname = "ClientTable_"
            uname = CompuMaster.Data.DataTables.LookupUniqueColumnName(dt, lname)
            Console.WriteLine("New free unique column name: " & uname & " (lookup name: " & lname & ")")
            Assert.IsNotEmpty(uname)
            Assert.IsFalse(dt.Columns.Contains(uname))
            lname = "1"
            uname = CompuMaster.Data.DataTables.LookupUniqueColumnName(dt, lname)
            Console.WriteLine("New free unique column name: " & uname & " (lookup name: " & lname & ")")
            Assert.IsNotEmpty(uname)
            Assert.IsFalse(dt.Columns.Contains(uname))
            dt.Columns.Add("ClientTable_1")
            lname = "ClientTable_"
            uname = CompuMaster.Data.DataTables.LookupUniqueColumnName(dt, lname)
            Console.WriteLine("New free unique column name: " & uname & " (lookup name: " & lname & ")")
            Assert.IsNotEmpty(uname)
            Assert.IsFalse(dt.Columns.Contains(uname))
            lname = "1"
            uname = CompuMaster.Data.DataTables.LookupUniqueColumnName(dt, lname)
            Console.WriteLine("New free unique column name: " & uname & " (lookup name: " & lname & ")")
            Assert.IsNotEmpty(uname)
            Assert.IsFalse(dt.Columns.Contains(uname))
            lname = "ID"
            uname = CompuMaster.Data.DataTables.LookupUniqueColumnName(dt, lname)
            Console.WriteLine("New free unique column name: " & uname & " (lookup name: " & lname & ")")
            Assert.IsNotEmpty(uname)
            Assert.IsFalse(dt.Columns.Contains(uname))
            lname = "ClientTable_ID"
            uname = CompuMaster.Data.DataTables.LookupUniqueColumnName(dt, lname)
            Console.WriteLine("New free unique column name: " & uname & " (lookup name: " & lname & ")")
            Assert.IsNotEmpty(uname)
            Assert.IsFalse(dt.Columns.Contains(uname))
            lname = "clienttable_id"
            uname = CompuMaster.Data.DataTables.LookupUniqueColumnName(dt, lname)
            Console.WriteLine("New free unique column name: " & uname & " (lookup name: " & lname & ")")
            Assert.AreNotEqual("clienttable_id", uname)
            Assert.IsNotEmpty(uname)
            Assert.IsFalse(dt.Columns.Contains(uname))
            lname = "ClientTable_ID2"
            uname = CompuMaster.Data.DataTables.LookupUniqueColumnName(dt, lname)
            Console.WriteLine("New free unique column name: " & uname & " (lookup name: " & lname & ")")
            Assert.IsNotEmpty(uname)
            Assert.IsFalse(dt.Columns.Contains(uname))
            lname = "ClientTable_ID23"
            uname = CompuMaster.Data.DataTables.LookupUniqueColumnName(dt, lname)
            Console.WriteLine("New free unique column name: " & uname & " (lookup name: " & lname & ")")
            Assert.IsNotEmpty(uname)
            Assert.IsFalse(dt.Columns.Contains(uname))
        End Sub

        <Test> Public Sub LookupUniqueColumnName2()
            Dim DuplicatedColumnNames15 = New String() {"Column", "Column", "Column", "Column", "Column", "Column", "Column", "Column", "Column", "Column", "Column", "Column", "Column", "Column", "Column"}
            Dim ColumnsSimplyNumbered15 = New String() {"Column1", "Column2", "Column3", "Column4", "Column5", "Column6", "Column7", "Column8", "Column9", "Column10", "Column11", "Column12", "Column13", "Column14", "Column15"}
            Assert.AreEqual("Column1", CompuMaster.Data.DataTables.LookupUniqueColumnName(DuplicatedColumnNames15, "Column"))
            Assert.AreEqual("Column1", CompuMaster.Data.DataTables.LookupUniqueColumnName(DuplicatedColumnNames15, "Column1"))
            Assert.AreEqual("Column16", CompuMaster.Data.DataTables.LookupUniqueColumnName(ColumnsSimplyNumbered15, "Column1"))
        End Sub

        <Test()> Public Sub CloneTableAndReArrangeDataColumns()
            Dim dt As New DataTable
            Dim dt2 As DataTable
            dt.Columns.Add("Test1")
            dt.Columns.Add("Test2")

            dt.Rows.Add(New String() {"Test1", "Test3"})
            dt.Rows.Add(New String() {"Test2", "Test4"})
            dt.Rows.Add(New String() {"Test5", "Test6"})
            dt.Rows.Add(New String() {"Test7", "Test8"})
            dt.Rows.Add(New String() {"Test9", "Test10"})

            'TODO: all overloads'

            dt2 = CompuMaster.Data.DataTables.CloneTableAndReArrangeDataColumns(dt, New String() {"Test1"})
            Assert.AreEqual(1, dt2.Columns.Count())
            Assert.AreEqual(5, dt2.Rows.Count())
            StringAssert.IsMatch("Test1", dt2.Columns.Item(0).ColumnName)
            Assert.AreEqual("Test9", dt2.Rows.Item(4).Item(0))

        End Sub

        <Test()> Public Sub AddColumns()
            Dim dt As DataTable
            dt = Me.TestTable1()
            Assert.AreEqual(2, dt.Columns.Count)

            CompuMaster.Data.DataTables.AddColumns(dt, "SomeStringColumn")
            Assert.AreEqual(3, dt.Columns.Count)
            Assert.AreEqual("SomeStringColumn", dt.Columns(2).ColumnName)
            Assert.AreEqual(GetType(String), dt.Columns(2).DataType)

            CompuMaster.Data.DataTables.AddColumns(dt, "SomeIntColumn", GetType(Integer))
            Assert.AreEqual(4, dt.Columns.Count)
            Assert.AreEqual("SomeIntColumn", dt.Columns(3).ColumnName)
            Assert.AreEqual(GetType(Integer), dt.Columns(3).DataType)

            CompuMaster.Data.DataTables.AddColumns(dt, New String() {"SomeStringColumn1", "SomeStringColumn2"})
            Assert.AreEqual(6, dt.Columns.Count)
            Assert.AreEqual("SomeStringColumn1", dt.Columns(4).ColumnName)
            Assert.AreEqual("SomeStringColumn2", dt.Columns(5).ColumnName)
            Assert.AreEqual(GetType(String), dt.Columns(4).DataType)
            Assert.AreEqual(GetType(String), dt.Columns(5).DataType)

            CompuMaster.Data.DataTables.AddColumns(dt, New String() {"ID", "Value"}) 'already existing columns - nothing should be added
            Assert.AreEqual(6, dt.Columns.Count)

            CompuMaster.Data.DataTables.AddColumns(dt, New String() {"SomeIntegerColumn1", "SomeIntegerColumn2"}, GetType(Integer))
            Assert.AreEqual(8, dt.Columns.Count)
            Assert.AreEqual("SomeIntegerColumn1", dt.Columns(6).ColumnName)
            Assert.AreEqual("SomeIntegerColumn2", dt.Columns(7).ColumnName)
            Assert.AreEqual(GetType(Integer), dt.Columns(6).DataType)
            Assert.AreEqual(GetType(Integer), dt.Columns(7).DataType)

        End Sub

        <Test()> Public Sub RemoveColumns()
            Dim dt As New DataTable
            dt.Columns.Add("SomeColumn")

            Dim row1 As DataRow = dt.NewRow
            row1.Item(0) = "D"

            Assert.IsTrue(dt.Columns.Item(0).ColumnName = "SomeColumn")
            CompuMaster.Data.DataTables.RemoveColumns(dt, New String() {"SomeColumn"})
            Assert.AreEqual(0, dt.Columns.Count())
        End Sub

        <Test()> Public Sub RemoveDuplicates()
            Dim dt As New DataTable
            dt.Columns.Add("Something")
            dt.Columns.Add("Something2")

            Dim row1 As DataRow = dt.NewRow
            row1.Item(0) = "Unique"
            row1.Item(0) = "Unique1"

            Dim row2 As DataRow = dt.NewRow
            row2.Item(0) = "NotSoUnique"
            row2.Item(0) = "NotSoUnique"

            Dim row3 As DataRow = dt.NewRow
            row3.Item(0) = "NotSoUnique"
            row3.Item(0) = "NotSoUnique"

            dt.Rows.Add(row1)
            dt.Rows.Add(row2)
            dt.Rows.Add(row3)

            CompuMaster.Data.DataTables.RemoveDuplicates(dt, "Something")

            Assert.AreEqual(2, dt.Rows.Count())
        End Sub

        <Test()> Public Sub RemoveRowsWithColumnValues()
            Dim dt As DataTable

            Dim dtTemplate As New DataTable
            dtTemplate.Columns.Add("Something")
            dtTemplate.Columns.Add("Something2")
            dtTemplate.Rows.Add(New String() {"A", "Z"})
            dtTemplate.Rows.Add(New String() {"B", "Y"})
            dtTemplate.Rows.Add(New String() {"C", "X"})
            dtTemplate.Rows.Add(New Object() {DBNull.Value, "N"})
            Assert.IsNotNull(dtTemplate.Rows(3)(0), "Not Nothing expected")
            Assert.IsTrue(IsDBNull(dtTemplate.Rows(3)(0)), "DBNull expected")
            dtTemplate.Rows.Add(New Object() {"", "N"})
            dtTemplate.Rows.Add(New Object() {Nothing, "N"})
            Assert.IsNotNull(dtTemplate.Rows(5)(0), "Not Nothing expected because of .NET logic to translate into DBNull.value")
            Assert.IsTrue(IsDBNull(dtTemplate.Rows(5)(0)), "DBNull expected because of .NET logic to translate into DBNull.value")

            Console.WriteLine()
            Console.WriteLine("Test 1 with DBNull/Empty/Null")
            dt = CompuMaster.Data.DataTables.CreateDataTableClone(dtTemplate)
            CompuMaster.Data.DataTables.RemoveRowsWithColumnValues(dt.Columns(0), New Object() {DBNull.Value, "", Nothing})
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
            Assert.AreEqual(3, dt.Rows.Count())
            Assert.AreEqual("A", dt.Rows.Item(0).Item(0))
            Assert.AreEqual("B", dt.Rows.Item(1).Item(0))
            Assert.AreEqual("C", dt.Rows.Item(2).Item(0))

            Console.WriteLine()
            Console.WriteLine("Test 2 with DBNull/Empty/Null")
            dt = CompuMaster.Data.DataTables.CreateDataTableClone(dtTemplate)
            CompuMaster.Data.DataTables.RemoveRowsWithColumnValues(dt.Columns(0), New Object() {"A", "B", "C"})
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
            Assert.AreEqual(3, dt.Rows.Count())
            Assert.AreEqual("N", dt.Rows.Item(0).Item(1))
            Assert.AreEqual("N", dt.Rows.Item(1).Item(1))
            Assert.AreEqual("N", dt.Rows.Item(2).Item(1))

            Console.WriteLine()
            Console.WriteLine("Test 3 some simple tests")
            dt = New DataTable
            dt.Columns.Add("Column1")
            dt.Columns.Add("Column2")
            Dim row1 As DataRow = dt.NewRow
            Dim row2 As DataRow = dt.NewRow
            Dim row3 As DataRow = dt.NewRow
            row1.Item(0) = "YouCanDeleteThis"
            row2.Item(0) = "YouCanDeleteThis"
            row3.Item(0) = "ButYouCouldKeepThis"
            dt.Rows.Add(row1)
            dt.Rows.Add(row2)
            dt.Rows.Add(row3)
            CompuMaster.Data.DataTables.RemoveRowsWithColumnValues(dt.Columns.Item(0), New String() {"YouCanDeleteThis"})
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
            Assert.AreEqual(1, dt.Rows.Count())

        End Sub

        <Test> Public Sub RemoveRowsWithWithoutRequiredValuesInColumn()

            Dim dtTemplate As New DataTable
            dtTemplate.Columns.Add("Something")
            dtTemplate.Columns.Add("Something2")
            dtTemplate.Rows.Add(New String() {"A", "Z"})
            dtTemplate.Rows.Add(New String() {"B", "Y"})
            dtTemplate.Rows.Add(New String() {"C", "X"})
            dtTemplate.Rows.Add(New Object() {DBNull.Value, "N"})
            Assert.IsNotNull(dtTemplate.Rows(3)(0), "Not Nothing expected")
            Assert.IsTrue(IsDBNull(dtTemplate.Rows(3)(0)), "DBNull expected")
            dtTemplate.Rows.Add(New Object() {"", "N"})
            dtTemplate.Rows.Add(New Object() {Nothing, "N"})
            Assert.IsNotNull(dtTemplate.Rows(5)(0), "Not Nothing expected because of .NET logic to translate into DBNull.value")
            Assert.IsTrue(IsDBNull(dtTemplate.Rows(5)(0)), "DBNull expected because of .NET logic to translate into DBNull.value")
            Dim dt As DataTable

            Console.WriteLine()
            Console.WriteLine("Test 1 with DBNull/Empty/Null")
            dt = CompuMaster.Data.DataTables.CreateDataTableClone(dtTemplate)
            CompuMaster.Data.DataTables.RemoveRowsWithWithoutRequiredValuesInColumn(dt.Columns(0), New Object() {DBNull.Value, "", Nothing})
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
            Assert.AreEqual(3, dt.Rows.Count())
            Assert.AreEqual("N", dt.Rows.Item(0).Item(1))
            Assert.AreEqual("N", dt.Rows.Item(1).Item(1))
            Assert.AreEqual("N", dt.Rows.Item(2).Item(1))

            Console.WriteLine()
            Console.WriteLine("Test 2 with DBNull/Empty/Null")
            dt = CompuMaster.Data.DataTables.CreateDataTableClone(dtTemplate)
            CompuMaster.Data.DataTables.RemoveRowsWithWithoutRequiredValuesInColumn(dt.Columns(0), New Object() {"A", "B", "C"})
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
            Assert.AreEqual(3, dt.Rows.Count())
            Assert.AreEqual("A", dt.Rows.Item(0).Item(0))
            Assert.AreEqual("B", dt.Rows.Item(1).Item(0))
            Assert.AreEqual("C", dt.Rows.Item(2).Item(0))
        End Sub

        <Test()> Public Sub RemoveRowsWithNoCorrespondingValueInComparisonTable()

            Dim dtTemplate As New DataTable
            dtTemplate.Columns.Add("Something")
            dtTemplate.Columns.Add("Something2")
            dtTemplate.Rows.Add(New String() {"A", "Z"})
            dtTemplate.Rows.Add(New String() {"B", "Y"})
            dtTemplate.Rows.Add(New String() {"C", "X"})
            dtTemplate.Rows.Add(New Object() {DBNull.Value, "N"})

            Dim dt2 As New DataTable
            dt2.Columns.Add("Test")
            dt2.Columns.Add("TestColumn2")
            dt2.Rows.Add(New String() {"A", "Z2"})
            dt2.Rows.Add(New String() {"B", "Y2"})
            dt2.Rows.Add(New String() {"D", "W2"})

            Dim dt As DataTable
            Dim MethodResult As Object

            Console.WriteLine()
            Console.WriteLine("Test 1 with DBNull at source but with removing source rows with DBNull")
            dt = CompuMaster.Data.DataTables.CreateDataTableClone(dtTemplate)
            MethodResult = CompuMaster.Data.DataTables.RemoveRowsWithNoCorrespondingValueInComparisonTable(dt.Columns(0), dt2.Columns(0)).ToArray
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
            Assert.AreEqual(New Object() {"C", DBNull.Value}, MethodResult)
            Assert.AreEqual(2, dt.Rows.Count())
            Assert.AreEqual("A", dt.Rows.Item(0).Item(0))
            Assert.AreEqual("B", dt.Rows.Item(1).Item(0))
            Assert.AreEqual("Z", dt.Rows.Item(0).Item(1))
            Assert.AreEqual("Y", dt.Rows.Item(1).Item(1))

            Console.WriteLine()
            Console.WriteLine("Test 2 with DBNull at source but not at comparison table")
            dt = CompuMaster.Data.DataTables.CreateDataTableClone(dtTemplate)
            MethodResult = CompuMaster.Data.DataTables.RemoveRowsWithNoCorrespondingValueInComparisonTable(dt.Columns(0), dt2.Columns(0), True, False)
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
            Assert.AreEqual(New Object() {"C", DBNull.Value}, MethodResult)
            Assert.AreEqual(2, dt.Rows.Count())
            Assert.AreEqual("A", dt.Rows.Item(0).Item(0))
            Assert.AreEqual("B", dt.Rows.Item(1).Item(0))
            Assert.AreEqual("Z", dt.Rows.Item(0).Item(1))
            Assert.AreEqual("Y", dt.Rows.Item(1).Item(1))

            Console.WriteLine()
            Console.WriteLine("Test 3 with DBNull at both sides")
            dt = CompuMaster.Data.DataTables.CreateDataTableClone(dtTemplate)
            dt2.Rows.Add(New Object() {DBNull.Value, "N2"})
            MethodResult = CompuMaster.Data.DataTables.RemoveRowsWithNoCorrespondingValueInComparisonTable(dt.Columns(0), dt2.Columns(0), True, False)
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
            Assert.AreEqual(New Object() {"C"}, MethodResult)
            Assert.AreEqual(3, dt.Rows.Count())
            Assert.AreEqual("A", dt.Rows.Item(0).Item(0))
            Assert.AreEqual("B", dt.Rows.Item(1).Item(0))
            Assert.AreEqual(DBNull.Value, dt.Rows.Item(2).Item(0))
            Assert.AreEqual("Z", dt.Rows.Item(0).Item(1))
            Assert.AreEqual("Y", dt.Rows.Item(1).Item(1))
            Assert.AreEqual("N", dt.Rows.Item(2).Item(1))

        End Sub

        <Test()> Public Sub RemoveRowsWithCorrespondingValueInComparisonTable()

            Dim dtTemplate As New DataTable
            dtTemplate.Columns.Add("Something")
            dtTemplate.Columns.Add("Something2")
            dtTemplate.Rows.Add(New String() {"A", "Z"})
            dtTemplate.Rows.Add(New String() {"B", "Y"})
            dtTemplate.Rows.Add(New String() {"C", "X"})
            dtTemplate.Rows.Add(New Object() {DBNull.Value, "N"})

            Dim dt2 As New DataTable
            dt2.Columns.Add("Test")
            dt2.Columns.Add("TestColumn2")
            dt2.Rows.Add(New String() {"A", "Z2"})
            dt2.Rows.Add(New String() {"B", "Y2"})
            dt2.Rows.Add(New String() {"D", "W2"})

            Dim dt As DataTable
            Dim MethodResult As ArrayList

            Console.WriteLine()
            Console.WriteLine("Test 1 with DBNull at source but with removing source rows with DBNull")
            dt = CompuMaster.Data.DataTables.CreateDataTableClone(dtTemplate)
            MethodResult = CompuMaster.Data.DataTables.RemoveRowsWithCorrespondingValueInComparisonTable(dt.Columns(0), dt2.Columns(0))
            Assert.AreEqual(3, MethodResult.Count)
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
            Assert.AreEqual(New Object() {"A", "B", DBNull.Value}, MethodResult)
            Assert.AreEqual(1, dt.Rows.Count())
            Assert.AreEqual("C", dt.Rows.Item(0).Item(0))
            Assert.AreEqual("X", dt.Rows.Item(0).Item(1))

            Console.WriteLine()
            Console.WriteLine("Test 2 with DBNull at source but not at comparison table")
            dt = CompuMaster.Data.DataTables.CreateDataTableClone(dtTemplate)
            MethodResult = CompuMaster.Data.DataTables.RemoveRowsWithCorrespondingValueInComparisonTable(dt.Columns(0), dt2.Columns(0), True, False)
            Assert.AreEqual(2, MethodResult.Count)
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
            Assert.AreEqual(2, dt.Rows.Count())
            Assert.AreEqual("C", dt.Rows.Item(0).Item(0))
            Assert.AreEqual("X", dt.Rows.Item(0).Item(1))
            Assert.AreEqual(DBNull.Value, dt.Rows.Item(1).Item(0))
            Assert.AreEqual("N", dt.Rows.Item(1).Item(1))

            Console.WriteLine()
            Console.WriteLine("Test 3 with DBNull at both sides")
            dt = CompuMaster.Data.DataTables.CreateDataTableClone(dtTemplate)
            dt2.Rows.Add(New Object() {DBNull.Value, "N2"})
            MethodResult = CompuMaster.Data.DataTables.RemoveRowsWithCorrespondingValueInComparisonTable(dt.Columns(0), dt2.Columns(0), True, False)
            Assert.AreEqual(3, MethodResult.Count)
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
            Assert.AreEqual(1, dt.Rows.Count())
            Assert.AreEqual("C", dt.Rows.Item(0).Item(0))
            Assert.AreEqual("X", dt.Rows.Item(0).Item(1))

        End Sub

        <Test()> Public Sub RowIndex()
            Dim dt As New DataTable
            dt.Columns.Add("DummyColumn")
            Dim drow As DataRow = dt.NewRow
            dt.Rows.Add(drow)
            Assert.AreEqual(0, CompuMaster.Data.DataTables.RowIndex(drow))

        End Sub


        Private Function CreateCrossJoinTablesTableSet1() As JoinTableSet
            Dim left As New DataTable
            Dim right As New DataTable
            'Cross join test'
            left.Columns.Add("FirstCol")
            left.Columns.Add("SecondCol")
            left.Columns.Add("ThirdCol")
            Dim lDataRow As DataRow = left.NewRow
            lDataRow.Item(0) = "Test"
            lDataRow.Item(1) = "Test2"
            left.Rows.Add(lDataRow)

            right.Columns.Add("FourthCol")
            Dim rDataRow As DataRow = right.NewRow
            rDataRow.Item(0) = "RightTest"
            right.Rows.Add(rDataRow)

            Return New JoinTableSet("CreateCrossJoinTableSet1", left, Nothing, right, Nothing)
        End Function

        Private Function CreateCrossJoinTablesTableSet2() As JoinTableSet
            Dim left As New DataTable
            Dim right As New DataTable
            'Cross join test'
            left.Columns.Add("FirstCol")
            left.Columns.Add("SecondCol")
            left.Columns.Add("ThirdCol")
            Dim lDataRow As DataRow = left.NewRow
            lDataRow.Item(0) = "Test1"
            lDataRow.Item(1) = "Test1Col2"
            left.Rows.Add(lDataRow)
            lDataRow = left.NewRow
            lDataRow.Item(0) = "Test2"
            lDataRow.Item(1) = "Test2Col2"
            left.Rows.Add(lDataRow)

            right.Columns.Add("FourthCol")
            Dim rDataRow As DataRow = right.NewRow
            rDataRow.Item(0) = "RightTest1"
            right.Rows.Add(rDataRow)
            rDataRow = right.NewRow
            rDataRow.Item(0) = "RightTest2"
            right.Rows.Add(rDataRow)
            rDataRow = right.NewRow
            rDataRow.Item(0) = "RightTest3"
            right.Rows.Add(rDataRow)

            Return New JoinTableSet("CreateCrossJoinTableSet2", left, Nothing, right, Nothing)
        End Function

        <Test()> Public Sub CrossJoinTables()

            Dim TestTableSet As JoinTableSet = Me.CreateCrossJoinTablesTableSet1
            TestTableSet.WriteToConsole()

            Dim crossjoined As DataTable = CompuMaster.Data.DataTables.CrossJoinTables(TestTableSet.LeftTable, Nothing, TestTableSet.RightTable, Nothing)
            StringAssert.IsMatch("FourthCol", crossjoined.Columns.Item(3).ColumnName)
            Assert.AreEqual("RightTest", crossjoined.Rows.Item(0).Item(3))
            Assert.IsTrue(IsDBNull(crossjoined.Rows.Item(0).Item(2)))


            Console.WriteLine("FULL-OUTER-JOINED TABLE CONTENTS")
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(crossjoined))
        End Sub

        <Test()> Public Sub SqlJoinTables_CrossJoin()
            SqlJoinTables_CrossJoin_Test1()
            SqlJoinTables_CrossJoin_Test2()
        End Sub

        <CodeAnalysis.SuppressMessage("Performance", "CA1825:Avoid zero-length array allocations.", Justification:="<Ausstehend>")>
        Private Sub SqlJoinTables_CrossJoin_Test1()

            Dim TestTableSet As JoinTableSet = Me.CreateCrossJoinTablesTableSet1
            TestTableSet.WriteToConsole()

            Dim CrossJoined As DataTable = CompuMaster.Data.DataTables.SqlJoinTables(TestTableSet.LeftTable, New String() {}, New String() {}, TestTableSet.RightTable, New String() {}, New String() {}, CompuMaster.Data.DataTables.SqlJoinTypes.Cross)

            Console.WriteLine("CROS-JOINED TABLE CONTENTS")
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(CrossJoined))

            'ResultComparisonValue to be evaluated
            Dim ShallBeResult As String = "JoinedTable" & System.Environment.NewLine &
                    "FirstCol|SecondCol|ThirdCol|FourthCol" & System.Environment.NewLine &
                    "--------+---------+--------+---------" & System.Environment.NewLine &
                    "Test    |Test2    |        |RightTest" & System.Environment.NewLine &
                    ""
            Assert.AreEqual(ShallBeResult, CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(CrossJoined))

        End Sub

        <CodeAnalysis.SuppressMessage("Performance", "CA1825:Avoid zero-length array allocations.", Justification:="<Ausstehend>")>
        Private Sub SqlJoinTables_CrossJoin_Test2()

            Dim TestTableSet As JoinTableSet = Me.CreateCrossJoinTablesTableSet2
            TestTableSet.WriteToConsole()

            Dim CrossJoined As DataTable = CompuMaster.Data.DataTables.SqlJoinTables(TestTableSet.LeftTable, New String() {}, New String() {}, TestTableSet.RightTable, New String() {}, New String() {}, CompuMaster.Data.DataTables.SqlJoinTypes.Cross)

            Console.WriteLine("CROS-JOINED TABLE CONTENTS")
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(CrossJoined))

            'ResultComparisonValue to be evaluated
            Dim ShallBeResult As String = "JoinedTable" & System.Environment.NewLine &
                    "FirstCol|SecondCol|ThirdCol|FourthCol " & System.Environment.NewLine &
                    "--------+---------+--------+----------" & System.Environment.NewLine &
                    "Test1   |Test1Col2|        |RightTest1" & System.Environment.NewLine &
                    "Test1   |Test1Col2|        |RightTest2" & System.Environment.NewLine &
                    "Test1   |Test1Col2|        |RightTest3" & System.Environment.NewLine &
                    "Test2   |Test2Col2|        |RightTest1" & System.Environment.NewLine &
                    "Test2   |Test2Col2|        |RightTest2" & System.Environment.NewLine &
                    "Test2   |Test2Col2|        |RightTest3" & System.Environment.NewLine &
                    ""
            Assert.AreEqual(ShallBeResult, CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(CrossJoined))

        End Sub

        Private Class JoinTableSet
            Public Sub New(tableSetName As String, leftTable As DataTable, leftPrimaryKeyColumns As String(), rightTable As DataTable, rightPrimaryKeyColumns As String())
                Me.TableSetName = tableSetName
                Me.LeftTable = leftTable
                Me.RightTable = rightTable
                Me.LeftPrimaryKeyColumns = leftPrimaryKeyColumns
                Me.RightPrimaryKeyColumns = rightPrimaryKeyColumns
            End Sub
            Public TableSetName As String
            Public LeftTable As DataTable
            Public RightTable As DataTable
            Public LeftPrimaryKeyColumns As String()
            Public RightPrimaryKeyColumns As String()
            Public Sub WriteToConsole()
                'Console.WriteLine("LEFT TABLE CONTENTS OF " & Me.TableSetName)
                'Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(Me.LeftTable))
                'Console.WriteLine()
                'Console.WriteLine("RIGHT TABLE CONTENTS OF " & Me.TableSetName)
                'Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(Me.RightTable))
                'Console.WriteLine()
                Console.WriteLine("LEFT+RIGHT TABLE CONTENTS OF " & Me.TableSetName)
                Console.WriteLine(CompuMaster.Data.Utils.ArrangeTableBlocksBesides(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(Me.LeftTable), CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(Me.RightTable)))
            End Sub
        End Class

#Region "FullOuterJoinTestTables"""
        Private Function CreateFullOuterJoinTableSet1() As JoinTableSet
            'Full outer join test'
            Dim right As New DataTable
            Dim left As New DataTable

            left.Columns.Add("left1")
            left.Columns.Add("left2")
            left.Columns.Add("left3")

            right.Columns.Add("right1")
            right.Columns.Add("right2")

            For j As Integer = 0 To 5
                Dim leftRow As DataRow = left.NewRow
                leftRow.Item(0) = j
                leftRow.Item(1) = j + 1
                leftRow.Item(2) = j + 2
                left.Rows.Add(leftRow)
            Next j

            For i As Integer = 0 To 5
                Dim rightrow As DataRow = right.NewRow
                rightrow.Item(0) = i
                rightrow.Item(1) = i + 10
                right.Rows.Add(rightrow)
            Next i

            Dim addiontalLeft As DataRow = left.NewRow
            addiontalLeft.Item(0) = 567
            addiontalLeft.Item(1) = 65527
            left.Rows.Add(addiontalLeft)
            addiontalLeft = left.NewRow
            addiontalLeft.Item(0) = 678
            addiontalLeft.Item(1) = 65528
            left.Rows.Add(addiontalLeft)
            addiontalLeft = left.NewRow
            left.Rows.Add(addiontalLeft)

            Dim addiontalRight As DataRow = right.NewRow
            addiontalRight.Item(0) = 789
            addiontalRight.Item(1) = 65728
            right.Rows.Add(addiontalRight)
            addiontalRight = right.NewRow
            addiontalRight.Item(0) = 890
            right.Rows.Add(addiontalRight)
            addiontalRight = right.NewRow
            right.Rows.Add(addiontalRight)

            Return New JoinTableSet("CreateFullOuterJoinTableSet1", left, New String() {"left1"}, right, New String() {"right1"})
        End Function

        Private Function CreateFullOuterJoinTableSet2() As JoinTableSet
            'Full outer join test'
            Dim right As New DataTable
            Dim left As New DataTable

            left.Columns.Add("left1")
            left.Columns.Add("left2")
            left.Columns.Add("left3")

            right.Columns.Add("right1")
            right.Columns.Add("right2")

            For j As Integer = 0 To 5
                Dim leftRow As DataRow = left.NewRow
                If j = 1 Then
                    leftRow.Item(0) = DBNull.Value
                Else
                    leftRow.Item(0) = j
                End If
                leftRow.Item(1) = j + 1
                leftRow.Item(2) = j + 2
                left.Rows.Add(leftRow)
            Next j

            For i As Integer = 0 To 5
                Dim rightrow As DataRow = right.NewRow
                If i = 2 Then
                    rightrow.Item(0) = DBNull.Value
                Else
                    rightrow.Item(0) = i
                End If

                If i = 4 Then
                    rightrow.Item(1) = i + 100
                ElseIf i = 5 Then
                    rightrow.Item(1) = DBNull.Value
                Else
                    rightrow.Item(1) = i + 1
                End If
                right.Rows.Add(rightrow)
            Next i

            Dim addiontalLeft As DataRow = left.NewRow
            addiontalLeft.Item(0) = 567
            addiontalLeft.Item(1) = 65527
            left.Rows.Add(addiontalLeft)
            addiontalLeft = left.NewRow
            addiontalLeft.Item(0) = 5
            addiontalLeft.Item(1) = 60
            addiontalLeft.Item(2) = 70
            left.Rows.Add(addiontalLeft)


            Dim addiontalRight As DataRow = right.NewRow
            addiontalRight.Item(0) = 789
            addiontalRight.Item(1) = 65728
            right.Rows.Add(addiontalRight)
            addiontalRight = right.NewRow
            addiontalRight.Item(0) = 3
            addiontalRight.Item(1) = 40
            right.Rows.Add(addiontalRight)
            addiontalRight = right.NewRow
            addiontalRight.Item(0) = 890
            right.Rows.Add(addiontalRight)

            Return New JoinTableSet("CreateFullOuterJoinTableSet2", left, New String() {"left1"}, right, New String() {"right1"})
        End Function
#End Region

        <Test()> Public Sub SqlJoinTables_FullOuter()
            FullOuterJoinTables_TableSet1()
            FullOuterJoinTables_TableSet1WithSameColumnNameInPrimaryKeyAtBothTables()
            FullOuterJoinTables_TableSet1WithSameColumnNameInPrimaryKeyAtBothTablesNamedClientTable_ID()
            FullOuterJoinTables_TableSet2()
        End Sub

        Private Sub FullOuterJoinTables_TableSet1()
            Dim TestTables As JoinTableSet = Me.CreateFullOuterJoinTableSet1
            TestTables.WriteToConsole()

            Dim FullOuterJoined As DataTable
            'FullOuterJoined = CompuMaster.Data.DataTables.FullJoinTables(TestTables.LeftTable, TestTables.LeftPrimaryKeyColumns, TestTables.RightTable, TestTables.RightPrimaryKeyColumns)
            FullOuterJoined = CompuMaster.Data.DataTables.SqlJoinTables(TestTables.LeftTable, TestTables.LeftPrimaryKeyColumns, Nothing, TestTables.RightTable, TestTables.RightPrimaryKeyColumns, Nothing, CompuMaster.Data.DataTables.SqlJoinTypes.FullOuter)

            Console.WriteLine("FULL-OUTER-JOINED TABLE CONTENTS: " & TestTables.TableSetName)
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(FullOuterJoined))

            'Verify column (names)
            Assert.AreEqual(5, FullOuterJoined.Columns.Count()) 'There must be exactly 5 columns '
            For i As Integer = 1 To 3
                Assert.AreEqual("left" & i, FullOuterJoined.Columns(i - 1).ColumnName)
            Next i
            For i As Integer = 1 To 2
                Assert.AreEqual("right" & i, FullOuterJoined.Columns(i + 2).ColumnName)
            Next i

            'Verify row (content)
            Assert.AreEqual(11, FullOuterJoined.Rows.Count())
            For i As Integer = 0 To 5
                Assert.AreEqual(0 + i, CInt(FullOuterJoined.Rows(i)(0)))
                Assert.AreEqual(1 + i, CInt(FullOuterJoined.Rows(i)(1)))
                Assert.AreEqual(2 + i, CInt(FullOuterJoined.Rows(i)(2)))
                Assert.AreEqual(0 + i, CInt(FullOuterJoined.Rows(i)(3)))
                Assert.AreEqual(10 + i, CInt(FullOuterJoined.Rows(i)(4)))
            Next i
            Assert.AreEqual(567, CInt(FullOuterJoined.Rows(6)(0)))
            Assert.AreEqual(65527, CInt(FullOuterJoined.Rows(6)(1)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(6)(2)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(6)(3)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(6)(4)))
            Assert.AreEqual(678, CInt(FullOuterJoined.Rows(7)(0)))
            Assert.AreEqual(65528, CInt(FullOuterJoined.Rows(7)(1)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(7)(2)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(7)(3)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(7)(4)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(8)(0)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(8)(1)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(8)(2)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(8)(3)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(8)(4)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(9)(0)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(9)(1)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(9)(2)))
            Assert.AreEqual(789, CInt(FullOuterJoined.Rows(9)(3)))
            Assert.AreEqual(65728, CInt(FullOuterJoined.Rows(9)(4)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(10)(0)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(10)(1)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(10)(2)))
            Assert.AreEqual(890, CInt(FullOuterJoined.Rows(10)(3)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(10)(4)))

        End Sub

        Private Sub FullOuterJoinTables_TableSet1WithSameColumnNameInPrimaryKeyAtBothTables()
            Dim TestTables As JoinTableSet = Me.CreateFullOuterJoinTableSet1
            TestTables.TableSetName &= "::WithSameColumnNameInPrimaryKeyAtBothTables"
            TestTables.LeftTable.Columns(0).ColumnName = "ID"
            TestTables.RightTable.Columns(0).ColumnName = "ID"
            TestTables.LeftPrimaryKeyColumns = New String() {"ID"}
            TestTables.RightPrimaryKeyColumns = New String() {"ID"}
            TestTables.WriteToConsole()

            Dim FullOuterJoined As DataTable
            'FullOuterJoined = CompuMaster.Data.DataTables.FullJoinTables(TestTables.LeftTable, TestTables.LeftPrimaryKeyColumns, TestTables.RightTable, TestTables.RightPrimaryKeyColumns)
            FullOuterJoined = CompuMaster.Data.DataTables.SqlJoinTables(TestTables.LeftTable, TestTables.LeftPrimaryKeyColumns, Nothing, TestTables.RightTable, TestTables.RightPrimaryKeyColumns, Nothing, CompuMaster.Data.DataTables.SqlJoinTypes.FullOuter)

            Console.WriteLine("FULL-OUTER-JOINED TABLE CONTENTS: " & TestTables.TableSetName)
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(FullOuterJoined))

            'Verify column (names)
            Assert.AreEqual(5, FullOuterJoined.Columns.Count()) 'There must be exactly 5 columns '
            Assert.AreEqual("ID", FullOuterJoined.Columns(0).ColumnName)
            For i As Integer = 2 To 3
                Assert.AreEqual("left" & i, FullOuterJoined.Columns(i - 1).ColumnName)
            Next i
            Assert.AreEqual("ClientTable_ID", FullOuterJoined.Columns(3).ColumnName)
            For i As Integer = 2 To 2
                Assert.AreEqual("right" & i, FullOuterJoined.Columns(i + 2).ColumnName)
            Next i

            'Verify row (content)
            Assert.AreEqual(11, FullOuterJoined.Rows.Count())
            For i As Integer = 0 To 5
                Assert.AreEqual(0 + i, CInt(FullOuterJoined.Rows(i)(0)))
                Assert.AreEqual(1 + i, CInt(FullOuterJoined.Rows(i)(1)))
                Assert.AreEqual(2 + i, CInt(FullOuterJoined.Rows(i)(2)))
                Assert.AreEqual(0 + i, CInt(FullOuterJoined.Rows(i)(3)))
                Assert.AreEqual(10 + i, CInt(FullOuterJoined.Rows(i)(4)))
            Next i
            Assert.AreEqual(567, CInt(FullOuterJoined.Rows(6)(0)))
            Assert.AreEqual(65527, CInt(FullOuterJoined.Rows(6)(1)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(6)(2)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(6)(3)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(6)(4)))
            Assert.AreEqual(678, CInt(FullOuterJoined.Rows(7)(0)))
            Assert.AreEqual(65528, CInt(FullOuterJoined.Rows(7)(1)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(7)(2)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(7)(3)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(7)(4)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(8)(0)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(8)(1)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(8)(2)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(8)(3)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(8)(4)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(9)(0)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(9)(1)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(9)(2)))
            Assert.AreEqual(789, CInt(FullOuterJoined.Rows(9)(3)))
            Assert.AreEqual(65728, CInt(FullOuterJoined.Rows(9)(4)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(10)(0)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(10)(1)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(10)(2)))
            Assert.AreEqual(890, CInt(FullOuterJoined.Rows(10)(3)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(10)(4)))

        End Sub

        Private Sub FullOuterJoinTables_TableSet1WithSameColumnNameInPrimaryKeyAtBothTablesNamedClientTable_ID()
            Dim TestTables As JoinTableSet = Me.CreateFullOuterJoinTableSet1
            TestTables.TableSetName &= "::WithSameColumnNameInPrimaryKeyAtBothTablesNamedClientTable_ID"
            TestTables.LeftTable.Columns(0).ColumnName = "ClientTable_ID"
            TestTables.RightTable.Columns(0).ColumnName = "ClientTable_ID"
            TestTables.LeftPrimaryKeyColumns = New String() {"ClientTable_ID"}
            TestTables.RightPrimaryKeyColumns = New String() {"ClientTable_ID"}
            TestTables.WriteToConsole()

            Dim FullOuterJoined As DataTable
            'FullOuterJoined = CompuMaster.Data.DataTables.FullJoinTables(TestTables.LeftTable, TestTables.LeftPrimaryKeyColumns, TestTables.RightTable, TestTables.RightPrimaryKeyColumns)
            FullOuterJoined = CompuMaster.Data.DataTables.SqlJoinTables(TestTables.LeftTable, TestTables.LeftPrimaryKeyColumns, Nothing, TestTables.RightTable, TestTables.RightPrimaryKeyColumns, Nothing, CompuMaster.Data.DataTables.SqlJoinTypes.FullOuter)

            Console.WriteLine("FULL-OUTER-JOINED TABLE CONTENTS: " & TestTables.TableSetName)
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(FullOuterJoined))

            'Verify column (names)
            Assert.AreEqual(5, FullOuterJoined.Columns.Count()) 'There must be exactly 5 columns '
            Assert.AreEqual("ClientTable_ID", FullOuterJoined.Columns(0).ColumnName)
            For i As Integer = 2 To 3
                Assert.AreEqual("left" & i, FullOuterJoined.Columns(i - 1).ColumnName)
            Next i
            Assert.AreEqual("ClientTable_ID1", FullOuterJoined.Columns(3).ColumnName)
            For i As Integer = 2 To 2
                Assert.AreEqual("right" & i, FullOuterJoined.Columns(i + 2).ColumnName)
            Next i

            'Verify row (content)
            Assert.AreEqual(11, FullOuterJoined.Rows.Count())
            For i As Integer = 0 To 5
                Assert.AreEqual(0 + i, CInt(FullOuterJoined.Rows(i)(0)))
                Assert.AreEqual(1 + i, CInt(FullOuterJoined.Rows(i)(1)))
                Assert.AreEqual(2 + i, CInt(FullOuterJoined.Rows(i)(2)))
                Assert.AreEqual(0 + i, CInt(FullOuterJoined.Rows(i)(3)))
                Assert.AreEqual(10 + i, CInt(FullOuterJoined.Rows(i)(4)))
            Next i
            Assert.AreEqual(567, CInt(FullOuterJoined.Rows(6)(0)))
            Assert.AreEqual(65527, CInt(FullOuterJoined.Rows(6)(1)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(6)(2)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(6)(3)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(6)(4)))
            Assert.AreEqual(678, CInt(FullOuterJoined.Rows(7)(0)))
            Assert.AreEqual(65528, CInt(FullOuterJoined.Rows(7)(1)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(7)(2)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(7)(3)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(7)(4)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(8)(0)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(8)(1)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(8)(2)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(8)(3)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(8)(4)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(9)(0)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(9)(1)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(9)(2)))
            Assert.AreEqual(789, CInt(FullOuterJoined.Rows(9)(3)))
            Assert.AreEqual(65728, CInt(FullOuterJoined.Rows(9)(4)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(10)(0)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(10)(1)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(10)(2)))
            Assert.AreEqual(890, CInt(FullOuterJoined.Rows(10)(3)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(10)(4)))

        End Sub

        Private Sub FullOuterJoinTables_TableSet2()
            Dim TestTables As JoinTableSet = Me.CreateFullOuterJoinTableSet2
            TestTables.WriteToConsole()

            Dim FullOuterJoined As DataTable
            'FullOuterJoined = CompuMaster.Data.DataTables.FullJoinTables(TestTables.LeftTable, TestTables.LeftPrimaryKeyColumns, TestTables.RightTable, TestTables.RightPrimaryKeyColumns)
            FullOuterJoined = CompuMaster.Data.DataTables.SqlJoinTables(TestTables.LeftTable, TestTables.LeftPrimaryKeyColumns, Nothing, TestTables.RightTable, TestTables.RightPrimaryKeyColumns, Nothing, CompuMaster.Data.DataTables.SqlJoinTypes.FullOuter)

            Console.WriteLine("FULL-OUTER-JOINED TABLE CONTENTS: " & TestTables.TableSetName)
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(FullOuterJoined))

            'Verify column (names)
            Assert.AreEqual(5, FullOuterJoined.Columns.Count()) 'There must be exactly 5 columns '
            For i As Integer = 1 To 3
                Assert.AreEqual("left" & i, FullOuterJoined.Columns(i - 1).ColumnName)
            Next i
            For i As Integer = 1 To 2
                Assert.AreEqual("right" & i, FullOuterJoined.Columns(i + 2).ColumnName)
            Next i

            'Verify row (content)
            Assert.AreEqual(12, FullOuterJoined.Rows.Count())
            For i As Integer = 0 To 3
                If i = 1 Then
                    Assert.AreEqual(DBNull.Value, FullOuterJoined.Rows(i)(0))
                Else
                    Assert.AreEqual(0 + i, CInt(FullOuterJoined.Rows(i)(0)))
                End If
                Assert.AreEqual(1 + i, CInt(FullOuterJoined.Rows(i)(1)))
                Assert.AreEqual(2 + i, CInt(FullOuterJoined.Rows(i)(2)))
                If i = 1 Or i = 2 Then
                    Assert.AreEqual(DBNull.Value, FullOuterJoined.Rows(i)(3))
                Else
                    Assert.AreEqual(0 + i, CInt(FullOuterJoined.Rows(i)(3)))
                End If
                If i = 1 Then
                    Assert.AreEqual(3, CInt(FullOuterJoined.Rows(i)(4)))
                ElseIf i = 2 Then
                    Assert.AreEqual(DBNull.Value, FullOuterJoined.Rows(i)(4))
                ElseIf i = 4 Then
                    Assert.AreEqual(1 + i, CInt(FullOuterJoined.Rows(i)(4)))
                End If
            Next i

            Dim ShallBeResult As String = "JoinedTable" & System.Environment.NewLine &
                    "left1|left2|left3|right1|right2" & System.Environment.NewLine &
                    "-----+-----+-----+------+------" & System.Environment.NewLine &
                    "0    |1    |2    |0     |1     " & System.Environment.NewLine &
                    "     |2    |3    |      |3     " & System.Environment.NewLine &
                    "2    |3    |4    |      |      " & System.Environment.NewLine &
                    "3    |4    |5    |3     |4     " & System.Environment.NewLine &
                    "3    |4    |5    |3     |40    " & System.Environment.NewLine &
                    "4    |5    |6    |4     |104   " & System.Environment.NewLine &
                    "5    |6    |7    |5     |      " & System.Environment.NewLine &
                    "567  |65527|     |      |      " & System.Environment.NewLine &
                    "5    |60   |70   |5     |      " & System.Environment.NewLine &
                    "     |     |     |1     |2     " & System.Environment.NewLine &
                    "     |     |     |789   |65728 " & System.Environment.NewLine &
                    "     |     |     |890   |      " & System.Environment.NewLine &
                    ""
            Assert.AreEqual(ShallBeResult, CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(FullOuterJoined))

        End Sub

        Private Function CreateInnerJoinTableSet1(shiftStepFor2ndTable As Integer) As JoinTableSet
            'Inner join test'
            Dim right As New DataTable
            Dim left As New DataTable

            left.Columns.Add("left1")
            left.Columns.Add("left2")
            left.Columns.Add("left3")
            left.Columns.Add("test")

            right.Columns.Add("right1")
            right.Columns.Add("right2")
            right.Columns.Add("test")

            For j As Integer = 0 To 5
                Dim leftRow As DataRow = left.NewRow
                leftRow.Item(0) = j
                leftRow.Item(1) = j + 1
                leftRow.Item(2) = j + 2
                leftRow.Item(3) = j + 3
                left.Rows.Add(leftRow)
            Next j


            For i As Integer = 0 To 5
                Dim rightrow As DataRow = right.NewRow
                rightrow.Item(0) = i
                rightrow.Item(1) = i + 1
                rightrow.Item(2) = i + 2
                If i >= shiftStepFor2ndTable Then right.Rows.Add(rightrow)
            Next i

            Dim ds As New DataSet
            ds.Tables.Add(left)
            ds.Tables.Add(right)
            Dim relation As New DataRelation("InnerJoined", left.Columns(0), right.Columns(0), True)
            'innerJoined = CompuMaster.Data.DataTables.JoinTables(left, leftColumns, right, rightColumns, CompuMaster.Data.DataTables.JoinTypes.Inner)
            ds.Relations.Add(relation)

            Return New JoinTableSet("CreateInnerJoinTableSet1", left, New String() {"left1"}, right, New String() {"right1"})

        End Function

        <Test()> Public Sub InnerJoinTables_UniqueColumnNamesAfterJoin()
            Dim TestTables As JoinTableSet = Me.CreateInnerJoinTableSet1(0)
            TestTables.WriteToConsole()

            Dim InnerJoined As DataTable
            InnerJoined = CompuMaster.Data.DataTables.JoinTables(TestTables.LeftTable, TestTables.RightTable, TestTables.LeftTable.DataSet.Relations(0), CompuMaster.Data.DataTables.JoinTypes.Inner)

            Console.WriteLine("INNER-JOINED TABLE CONTENTS: " & TestTables.TableSetName)
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(InnerJoined))

            'Acceptance Criteria: 
            '- column of 2nd table must be renamed if 1st table already contains the same column name
            '- renamed column (e.g. "column1") must not exist in list of selected columns (incl. PKs) from 1st and 2nd table           
            Assert.AreEqual(0, CInt(InnerJoined.Rows(0)("left1")))
            Assert.AreEqual(3, CInt(InnerJoined.Rows(0)("test")))
            Assert.AreEqual(2, CInt(InnerJoined.Rows(0)("ClientTable_test")))

            Assert.AreEqual("3", InnerJoined.Rows.Item(0).Item("test"))
            Assert.AreEqual("2", InnerJoined.Rows.Item(0).Item("ClientTable_test"))
        End Sub

        <Test()> Public Sub InnerJoinTables_TablesWithEqualPKs()
            Dim TestTables As JoinTableSet = Me.CreateInnerJoinTableSet1(0)
            TestTables.WriteToConsole()

            Dim InnerJoined As DataTable
            InnerJoined = CompuMaster.Data.DataTables.JoinTables(TestTables.LeftTable, TestTables.RightTable, TestTables.LeftTable.DataSet.Relations(0), CompuMaster.Data.DataTables.JoinTypes.Inner)

            Console.WriteLine("INNER-JOINED TABLE CONTENTS: " & TestTables.TableSetName)
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(InnerJoined))

            'Verify column (names)
            For i As Integer = 1 To 3
                Assert.AreEqual("left" & i, InnerJoined.Columns(i - 1).ColumnName)
            Next i
            StringAssert.IsMatch("left1", InnerJoined.Columns(0).ColumnName)
            StringAssert.IsMatch("left2", InnerJoined.Columns(1).ColumnName)
            StringAssert.IsMatch("left3", InnerJoined.Columns(2).ColumnName)
            StringAssert.IsMatch("test", InnerJoined.Columns(3).ColumnName)
            StringAssert.IsMatch("right1", InnerJoined.Columns(4).ColumnName)
            StringAssert.IsMatch("right2", InnerJoined.Columns(5).ColumnName)
            StringAssert.IsMatch("test", InnerJoined.Columns(6).ColumnName)
            Assert.AreEqual(7, InnerJoined.Columns.Count())

            'Verify row count and accessibility
            Assert.AreEqual(6, InnerJoined.Rows.Count())

            'Verify row (content)
            Assert.AreEqual("3", InnerJoined.Rows.Item(0).Item(3))
            Assert.AreEqual("2", InnerJoined.Rows.Item(0).Item(6))

            For RowCounter As Integer = 0 To 5
                Assert.AreEqual(0 + RowCounter, CInt(InnerJoined.Rows(RowCounter)(0)), "RowCounter=" & RowCounter)
                Assert.AreEqual(1 + RowCounter, CInt(InnerJoined.Rows(RowCounter)(1)), "RowCounter=" & RowCounter)
                Assert.AreEqual(2 + RowCounter, CInt(InnerJoined.Rows(RowCounter)(2)), "RowCounter=" & RowCounter)
                Assert.AreEqual(3 + RowCounter, CInt(InnerJoined.Rows(RowCounter)(3)), "RowCounter=" & RowCounter)
                Assert.AreEqual(0 + RowCounter, CInt(InnerJoined.Rows(RowCounter)(4)), "RowCounter=" & RowCounter)
                Assert.AreEqual(1 + RowCounter, CInt(InnerJoined.Rows(RowCounter)(5)), "RowCounter=" & RowCounter)
                Assert.AreEqual(2 + RowCounter, CInt(InnerJoined.Rows(RowCounter)(6)), "RowCounter=" & RowCounter)
            Next RowCounter

        End Sub

        <Test()> Public Sub InnerJoinTables_TablesWithSomeEqualPKs()
            Dim TestTables As JoinTableSet = Me.CreateInnerJoinTableSet1(1)
            TestTables.WriteToConsole()

            Dim InnerJoined As DataTable
            InnerJoined = CompuMaster.Data.DataTables.JoinTables(TestTables.LeftTable, TestTables.RightTable, TestTables.LeftTable.DataSet.Relations(0), CompuMaster.Data.DataTables.JoinTypes.Inner)

            Console.WriteLine("INNER-JOINED TABLE CONTENTS: " & TestTables.TableSetName)
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(InnerJoined))

            'Verify column (names)
            For i As Integer = 1 To 3
                Assert.AreEqual("left" & i, InnerJoined.Columns(i - 1).ColumnName)
            Next i
            StringAssert.IsMatch("left1", InnerJoined.Columns(0).ColumnName)
            StringAssert.IsMatch("left2", InnerJoined.Columns(1).ColumnName)
            StringAssert.IsMatch("left3", InnerJoined.Columns(2).ColumnName)
            StringAssert.IsMatch("test", InnerJoined.Columns(3).ColumnName)
            StringAssert.IsMatch("right1", InnerJoined.Columns(4).ColumnName)
            StringAssert.IsMatch("right2", InnerJoined.Columns(5).ColumnName)
            StringAssert.IsMatch("test", InnerJoined.Columns(6).ColumnName)
            Assert.AreEqual(7, InnerJoined.Columns.Count())

            'Verify row count and accessibility
            Assert.AreEqual(5, InnerJoined.Rows.Count())

            'Verify row (content)
            Assert.AreEqual("4", InnerJoined.Rows.Item(0).Item(3))
            Assert.AreEqual("3", InnerJoined.Rows.Item(0).Item(6))

            For RowCounter As Integer = 0 To 4
                Assert.AreEqual(0 + 1 + RowCounter, CInt(InnerJoined.Rows(RowCounter)(0)), "RowCounter=" & RowCounter)
                Assert.AreEqual(1 + 1 + RowCounter, CInt(InnerJoined.Rows(RowCounter)(1)), "RowCounter=" & RowCounter)
                Assert.AreEqual(2 + 1 + RowCounter, CInt(InnerJoined.Rows(RowCounter)(2)), "RowCounter=" & RowCounter)
                Assert.AreEqual(3 + 1 + RowCounter, CInt(InnerJoined.Rows(RowCounter)(3)), "RowCounter=" & RowCounter)
                Assert.AreEqual(0 + 1 + RowCounter, CInt(InnerJoined.Rows(RowCounter)(4)), "RowCounter=" & RowCounter)
                Assert.AreEqual(1 + 1 + RowCounter, CInt(InnerJoined.Rows(RowCounter)(5)), "RowCounter=" & RowCounter)
                Assert.AreEqual(2 + 1 + RowCounter, CInt(InnerJoined.Rows(RowCounter)(6)), "RowCounter=" & RowCounter)
            Next RowCounter

        End Sub

        Private Function CreateLeftJoinTableSet1() As JoinTableSet
            Dim right As New DataTable
            Dim left As New DataTable

            left.Columns.Add("left1")
            left.Columns.Add("left2")

            right.Columns.Add("right1")
            right.Columns.Add("right2")

            Dim leftRow As DataRow = left.NewRow
            leftRow.Item("left1") = 5
            leftRow.Item("left2") = 10
            left.Rows.Add(leftRow)

            Dim leftRow1 As DataRow = left.NewRow
            leftRow1.Item("left1") = 23
            leftRow1.Item("left2") = 99
            left.Rows.Add(leftRow1)

            Dim rightRow As DataRow = right.NewRow
            rightRow.Item("right1") = 5
            rightRow.Item("right2") = 11
            right.Rows.Add(rightRow)


            Dim rightRow1 As DataRow = right.NewRow
            rightRow1.Item("right1") = 242
            rightRow1.Item("right2") = 2512
            right.Rows.Add(rightRow1)

            Dim rightRow2 As DataRow = right.NewRow
            rightRow2.Item("right1") = 123
            rightRow2.Item("right2") = 1234
            right.Rows.Add(rightRow2)

            left.PrimaryKey = New System.Data.DataColumn() {left.Columns("left1")}
            right.PrimaryKey = New System.Data.DataColumn() {right.Columns("right1")}

            Console.WriteLine("LEFT TABLE CONTENTS")
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(left))
            Console.WriteLine()
            Console.WriteLine("RIGHT TABLE CONTENTS")
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(right))

            Dim ds As New DataSet
            ds.Tables.Add(left)
            ds.Tables.Add(right)

            Dim relation As New DataRelation("LeftJoined", left.Columns(0), right.Columns(0), False)
            ds.Relations.Add(relation)

            Return New JoinTableSet("CreateLeftJoinTableSet1", left, New String() {"left1"}, right, New String() {"right1"})

        End Function

        <Test()> Public Sub LeftJoinTables()
            Dim TestTables As JoinTableSet = Me.CreateLeftJoinTableSet1
            TestTables.WriteToConsole()

            Dim LeftJoined As DataTable
            LeftJoined = CompuMaster.Data.DataTables.JoinTables(TestTables.LeftTable, TestTables.RightTable, TestTables.LeftTable.DataSet.Relations(0), CompuMaster.Data.DataTables.JoinTypes.Left)

            Console.WriteLine("LEFT-JOINED TABLE CONTENTS: " & TestTables.TableSetName)
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(LeftJoined))

            'Verify column (names)
            Assert.AreEqual(4, LeftJoined.Columns.Count())
            StringAssert.IsMatch("left1", LeftJoined.Columns(0).ColumnName)
            StringAssert.IsMatch("left2", LeftJoined.Columns(1).ColumnName)
            StringAssert.IsMatch("right1", LeftJoined.Columns(2).ColumnName)
            StringAssert.IsMatch("right2", LeftJoined.Columns(3).ColumnName)

            'Verify row (content)
            Assert.AreEqual(2, LeftJoined.Rows.Count())
            Assert.IsTrue(IsDBNull(LeftJoined.Rows(1).Item(2)))
            Assert.AreEqual("23", LeftJoined.Rows(1).Item(0))
            Assert.AreEqual("5", LeftJoined.Rows(0).Item(0))
            Assert.AreEqual("11", LeftJoined.Rows(0).Item(3))

            Assert.AreEqual(10, CInt(LeftJoined.Rows(0)(1)))
            Assert.AreEqual(5, CInt(LeftJoined.Rows(0)(2)))
            Assert.AreEqual(99, CInt(LeftJoined.Rows(1)(1)))
        End Sub

        <Test(), Ignore("ToBeImplemented")> Public Sub SqlJoinTables_Inner()
            'result rows always with partner in other table
            Throw New NotImplementedException
        End Sub

        <Test()> Public Sub SqlJoinTables_Left()
            SqlJoinTables_Left_TableSet1()
            SqlJoinTables_Left_TableSet2()
        End Sub

        Private Sub SqlJoinTables_Left_TableSet1()
            Dim TestTables As JoinTableSet = Me.CreateLeftJoinTableSet1
            TestTables.WriteToConsole()

            Dim LeftJoined As DataTable
            LeftJoined = CompuMaster.Data.DataTables.SqlJoinTables(TestTables.LeftTable, TestTables.LeftTable.PrimaryKey, Nothing, TestTables.RightTable, TestTables.RightTable.PrimaryKey, Nothing, CompuMaster.Data.DataTables.SqlJoinTypes.Left)

            Console.WriteLine("LEFT-JOINED TABLE CONTENTS: " & TestTables.TableSetName)
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(LeftJoined))

            'Verify column (names)
            Assert.AreEqual(4, LeftJoined.Columns.Count())
            StringAssert.IsMatch("left1", LeftJoined.Columns(0).ColumnName)
            StringAssert.IsMatch("left2", LeftJoined.Columns(1).ColumnName)
            StringAssert.IsMatch("right1", LeftJoined.Columns(2).ColumnName)
            StringAssert.IsMatch("right2", LeftJoined.Columns(3).ColumnName)

            'Verify row (content)
            Assert.AreEqual(2, LeftJoined.Rows.Count())
            Assert.IsTrue(IsDBNull(LeftJoined.Rows(1).Item(2)))
            Assert.AreEqual("23", LeftJoined.Rows(1).Item(0))
            Assert.AreEqual("5", LeftJoined.Rows(0).Item(0))
            Assert.AreEqual("11", LeftJoined.Rows(0).Item(3))

            Assert.AreEqual(10, CInt(LeftJoined.Rows(0)(1)))
            Assert.AreEqual(5, CInt(LeftJoined.Rows(0)(2)))
            Assert.AreEqual(99, CInt(LeftJoined.Rows(1)(1)))

        End Sub

        <Test> Public Sub SqlJoinTables_ExpectedExceptionNullParameter()
            Assert.Throws(Of ArgumentNullException)(Sub()
                                                        Dim TestTables As JoinTableSet = Me.CreateLeftJoinTableSet1
                                                        CompuMaster.Data.DataTables.SqlJoinTables(TestTables.LeftTable, TestTables.LeftTable.PrimaryKey, Nothing, Nothing, TestTables.RightTable.PrimaryKey, Nothing, CompuMaster.Data.DataTables.SqlJoinTypes.Left)
                                                    End Sub)
        End Sub

        Private Sub SqlJoinTables_Left_TableSet2()
            Dim TestTables As JoinTableSet = Me.CreateFullOuterJoinTableSet2
            TestTables.WriteToConsole()

            Dim LeftJoined As DataTable
            LeftJoined = CompuMaster.Data.DataTables.SqlJoinTables(TestTables.LeftTable, New String() {"left1"}, Nothing, TestTables.RightTable, New String() {"right1"}, Nothing, CompuMaster.Data.DataTables.SqlJoinTypes.Left)

            Console.WriteLine("LEFT-JOINED TABLE CONTENTS: " & TestTables.TableSetName)
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(LeftJoined))

            Dim ShallBeResult As String = "JoinedTable" & System.Environment.NewLine &
                    "left1|left2|left3|right1|right2" & System.Environment.NewLine &
                    "-----+-----+-----+------+------" & System.Environment.NewLine &
                    "0    |1    |2    |0     |1     " & System.Environment.NewLine &
                    "     |2    |3    |      |3     " & System.Environment.NewLine &
                    "2    |3    |4    |      |      " & System.Environment.NewLine &
                    "3    |4    |5    |3     |4     " & System.Environment.NewLine &
                    "3    |4    |5    |3     |40    " & System.Environment.NewLine &
                    "4    |5    |6    |4     |104   " & System.Environment.NewLine &
                    "5    |6    |7    |5     |      " & System.Environment.NewLine &
                    "567  |65527|     |      |      " & System.Environment.NewLine &
                    "5    |60   |70   |5     |      " & System.Environment.NewLine &
                    ""
            Assert.AreEqual(ShallBeResult, CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(LeftJoined))

        End Sub

        <Test()> Public Sub FindDuplicates()
            Dim testTable As DataTable = TestTable1()
            Dim Result As Hashtable
            Result = CompuMaster.Data.DataTables.FindDuplicates(testTable.Columns("value"))

            For Each MyItem As DictionaryEntry In Result
                Console.WriteLine(CType(MyItem.Key, String) & "=" & CType(MyItem.Value, String))
            Next

            Assert.AreEqual(2, Result.Count, "JW #00001")
            For Each MyKey As DictionaryEntry In Result
                If CType(MyKey.Key, String) = "Hello world!" Then
                    Assert.AreEqual(3, MyKey.Value, "JW #00002")
                ElseIf CType(MyKey.Key, String) = "Gotcha!" Then
                    'Gotcha! (not the one in capital letters!)
                    Assert.AreEqual(2, MyKey.Value, "JW #00003")
                Else
                    Assert.Fail("Invalid returned value: " & CType(MyKey.Key, Object).ToString)
                End If
            Next


            Result = CompuMaster.Data.DataTables.FindDuplicates(testTable.Columns("value"), 3)
            Assert.AreEqual(1, Result.Count, "JW #00011")
            For Each MyKey As DictionaryEntry In Result
                Assert.AreEqual("Hello world!", MyKey.Key, "JW #00012")
                Assert.AreEqual(3, MyKey.Value, "JW #00013")
            Next

        End Sub

        <Test()> Public Sub KeepColumnsAndRemoveAllOthers()
            Dim testData As New DataTable
            testData.Columns.Add()
            testData.Columns.Add("test")
            testData.Columns.Add("äöüÄÖÜß")
            testData.Columns.Add("data")
            testData.Columns.Add()

            CompuMaster.Data.DataTables.KeepColumnsAndRemoveAllOthers(testData, New String() {"ÄÖÜäöüß", "data", "istnich", ""})
            Assert.AreEqual(2, testData.Columns.Count)
            Assert.AreEqual("äöüÄÖÜß", testData.Columns(0).ColumnName)
            Assert.AreEqual("data", testData.Columns(1).ColumnName)

        End Sub

        <Test()> Public Sub CopyDataTablesCaseSensitive()
            Dim source As New DataTable
            Dim dest As New DataTable
            source.Columns.Add("TestColumn")
            dest.Columns.Add("Testcolumn")
            Dim r As DataRow = source.NewRow
            r.Item("TestColumn") = "test"
            source.Rows.Add(r)


            Dim r2 As DataRow = dest.NewRow
            r2.Item("Testcolumn") = "test"
            dest.Rows.Add(r2)

            'First, some case insensitive tests'
            CompuMaster.Data.DataTables.CreateDataTableClone(source, dest, Nothing, Nothing, 2, False, False, False, False, True)
            Assert.AreEqual("test", dest.Rows.Item(0).Item("Testcolumn"))
            Assert.Less(dest.Columns.Count(), 2) 'if 2 then the test failed, because we did a CaseInsensitive comparison'
            StringAssert.IsMatch("Testcolumn", dest.Columns.Item(0).ColumnName)

            'Now case sensitive'
            CompuMaster.Data.DataTables.CreateDataTableClone(source, dest, Nothing, Nothing, 2, False, False, False, False, False)
            Assert.AreEqual(2, dest.Columns.Count()) 'Function shouldn't find "TestColumn" in dest (because there we only have Testcolumn (lowercase 'c'), and therefore add it => 2 columns in table'
            StringAssert.IsMatch("Testcolumn", dest.Columns.Item(0).ColumnName)
            StringAssert.IsMatch("TestColumn", dest.Columns.Item(1).ColumnName)

        End Sub

        <Test()> Public Sub ValidateRequiredColumnNames()
            Dim dt As New DataTable
            dt.Columns.Add("hEllo")

            Assert.AreEqual(1, CompuMaster.Data.DataTables.ValidateRequiredColumnNames(dt, New String() {"_suf"}).Length)
            Assert.AreEqual(1, CompuMaster.Data.DataTables.ValidateRequiredColumnNames(dt, New String() {"_suf", "hEllo"}).Length)
            Assert.AreEqual(2, CompuMaster.Data.DataTables.ValidateRequiredColumnNames(dt, New String() {"_suf", "Hello"}).Length)
            Assert.AreEqual(1, CompuMaster.Data.DataTables.ValidateRequiredColumnNames(dt, New String() {"Hello"}).Length)
            Assert.AreEqual(0, CompuMaster.Data.DataTables.ValidateRequiredColumnNames(dt, New String() {"hEllo"}).Length)
            Assert.AreEqual("_suf", CompuMaster.Data.DataTables.ValidateRequiredColumnNames(dt, New String() {"_suf"})(0))

            Assert.AreEqual(1, CompuMaster.Data.DataTables.ValidateRequiredColumnNames(dt, New String() {"_suf"}, False).Length)
            Assert.AreEqual(1, CompuMaster.Data.DataTables.ValidateRequiredColumnNames(dt, New String() {"_suf", "hEllo"}, False).Length)
            Assert.AreEqual(2, CompuMaster.Data.DataTables.ValidateRequiredColumnNames(dt, New String() {"_suf", "Hello"}, False).Length)
            Assert.AreEqual(1, CompuMaster.Data.DataTables.ValidateRequiredColumnNames(dt, New String() {"Hello"}, False).Length)
            Assert.AreEqual(0, CompuMaster.Data.DataTables.ValidateRequiredColumnNames(dt, New String() {"hEllo"}, False).Length)
            Assert.AreEqual("_suf", CompuMaster.Data.DataTables.ValidateRequiredColumnNames(dt, New String() {"_suf"}, False)(0))

            Assert.AreEqual(1, CompuMaster.Data.DataTables.ValidateRequiredColumnNames(dt, New String() {"_suf"}, True).Length)
            Assert.AreEqual(1, CompuMaster.Data.DataTables.ValidateRequiredColumnNames(dt, New String() {"_suf", "hEllo"}, True).Length)
            Assert.AreEqual(1, CompuMaster.Data.DataTables.ValidateRequiredColumnNames(dt, New String() {"_suf", "Hello"}, True).Length)
            Assert.AreEqual(0, CompuMaster.Data.DataTables.ValidateRequiredColumnNames(dt, New String() {"Hello"}, True).Length)
            Assert.AreEqual(0, CompuMaster.Data.DataTables.ValidateRequiredColumnNames(dt, New String() {"hEllo"}, True).Length)
            Assert.AreEqual("_suf", CompuMaster.Data.DataTables.ValidateRequiredColumnNames(dt, New String() {"_suf"}, True)(0))
        End Sub

        <Test()> Public Sub IsEmptyColumn()
            Dim dt As New DataTable
            dt.Columns.Add("col1", GetType(String))
            dt.Columns.Add("col2", GetType(Object))
            dt.Columns.Add("col3", GetType(String()))
            dt.Columns.Add("col4", GetType(DateTime))
            dt.Columns.Add("col5", GetType(Integer))
            dt.Columns.Add("col6", GetType(Boolean))
            dt.Columns.Add("col7", GetType(List(Of String)))
            Dim NewRow As DataRow

            'no rows -> everything must be considered empty
            For Each col As DataColumn In dt.Columns
                Assert.AreEqual(True, CompuMaster.Data.DataTables.IsEmptyColumn(col), col.ColumnName)
            Next

            'DbNull row -> everything must be considered empty
            NewRow = dt.NewRow()
            dt.Rows.Add(NewRow)
            For Each col As DataColumn In dt.Columns
                Assert.AreEqual(True, CompuMaster.Data.DataTables.IsEmptyColumn(col), col.ColumnName)
            Next

            'Null/Nothing row -> everything must be considered empty
            NewRow = dt.NewRow()
            NewRow.ItemArray = New Object() {Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing}
            dt.Rows.Add(NewRow)
            For Each col As DataColumn In dt.Columns
                Assert.AreEqual(True, CompuMaster.Data.DataTables.IsEmptyColumn(col), col.ColumnName)
            Next

            'values row -> everything must be considered NOT empty
            NewRow = dt.NewRow()
            NewRow.ItemArray = New Object() {"test", New Object(), New String() {}, New DateTime(2000, 1, 1, 12, 0, 0), 1, True, New List(Of String)}
            dt.Rows.Add(NewRow)
            For Each col As DataColumn In dt.Columns
                Assert.AreEqual(False, CompuMaster.Data.DataTables.IsEmptyColumn(col), col.ColumnName)
            Next

            'values row with other values -> everything must be considered NOT empty
            NewRow.ItemArray = New Object() {"", New Object(), New String() {""}, DateTime.MinValue, 0, False, New List(Of String)(0)}
            For Each col As DataColumn In dt.Columns
                Assert.AreEqual(False, CompuMaster.Data.DataTables.IsEmptyColumn(col), col.ColumnName)
            Next
        End Sub

        Private Shared Function RemoveEmptyColumns_TestTable(itemArray As Object()) As DataTable
            Dim dt As New DataTable
            dt.Columns.Add("col1", GetType(String))
            dt.Columns.Add("col2", GetType(Object))
            dt.Columns.Add("col3", GetType(String()))
            dt.Columns.Add("col4", GetType(DateTime))
            dt.Columns.Add("col5", GetType(Integer))
            dt.Columns.Add("col6", GetType(Boolean))
            dt.Columns.Add("col7", GetType(List(Of String)))
            If itemArray IsNot Nothing Then
                Dim NewRow As DataRow = dt.NewRow
                NewRow.ItemArray = itemArray
                dt.Rows.Add(NewRow)
            End If
            Return dt
        End Function

        <Test()> Public Sub RemoveEmptyColumns()
            Dim dt As DataTable

            'no rows -> everything must be considered empty
            dt = RemoveEmptyColumns_TestTable(Nothing)
            CompuMaster.Data.DataTables.RemoveEmptyColumns(dt)
            Assert.AreEqual(0, dt.Columns.Count)

            'DbNull row -> everything must be considered empty
            dt = RemoveEmptyColumns_TestTable(New Object() {DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value})
            CompuMaster.Data.DataTables.RemoveEmptyColumns(dt)
            Assert.AreEqual(0, dt.Columns.Count)

            'Null/Nothing row -> everything must be considered empty
            dt = RemoveEmptyColumns_TestTable(New Object() {Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing})
            CompuMaster.Data.DataTables.RemoveEmptyColumns(dt)
            Assert.AreEqual(0, dt.Columns.Count)

            'values row -> everything must be considered NOT empty
            dt = RemoveEmptyColumns_TestTable(New Object() {"test", New Object(), New String() {}, New DateTime(2000, 1, 1, 12, 0, 0), 1, True, New List(Of String)})
            CompuMaster.Data.DataTables.RemoveEmptyColumns(dt)
            Assert.AreEqual(7, dt.Columns.Count)

            'values row with other values -> everything must be considered NOT empty           
            dt = RemoveEmptyColumns_TestTable(New Object() {"", New Object(), New String() {""}, DateTime.MinValue, 0, False, New List(Of String)(0)})
            CompuMaster.Data.DataTables.RemoveEmptyColumns(dt)
            Assert.AreEqual(7, dt.Columns.Count)
        End Sub

        <Test> Public Sub RemoveColumnsExcept()
            Dim Table As DataTable

            Table = Me.TestTable2WithInvariantCultureInColumnNames
            Assert.AreEqual(New String() {"Frage", "Antwort A", "Antwort B", "Antwort C", "Antwort D", "Rubrik", "Richtige Antwort", "Erläuterung", "100 ", "200 ", "500 ", "1000 ", "5000 ", "10000 ", "20000 "}, CompuMaster.Data.DataTables.AllColumnNames(Table))
            CompuMaster.Data.DataTables.RemoveColumnsExcept(Table, Table.Columns(4), Table.Columns(3), Table.Columns(2), Table.Columns(1))
            Assert.AreEqual(New String() {"Antwort A", "Antwort B", "Antwort C", "Antwort D"}, CompuMaster.Data.DataTables.AllColumnNames(Table))

            Table = Me.TestTable2WithInvariantCultureInColumnNames
            Assert.AreEqual(New String() {"Frage", "Antwort A", "Antwort B", "Antwort C", "Antwort D", "Rubrik", "Richtige Antwort", "Erläuterung", "100 ", "200 ", "500 ", "1000 ", "5000 ", "10000 ", "20000 "}, CompuMaster.Data.DataTables.AllColumnNames(Table))
            CompuMaster.Data.DataTables.RemoveColumnsExcept(Table, "Antwort D", "Antwort C", "Antwort B", "Antwort A")
            Assert.AreEqual(New String() {"Antwort A", "Antwort B", "Antwort C", "Antwort D"}, CompuMaster.Data.DataTables.AllColumnNames(Table))
        End Sub

        <Test> Public Sub SortColumns()
            Dim Table As DataTable

            Table = Me.TestTable2WithInvariantCultureInColumnNames
            Assert.AreEqual(New String() {"Frage", "Antwort A", "Antwort B", "Antwort C", "Antwort D", "Rubrik", "Richtige Antwort", "Erläuterung", "100 ", "200 ", "500 ", "1000 ", "5000 ", "10000 ", "20000 "}, CompuMaster.Data.DataTables.AllColumnNames(Table))
#Disable Warning BC40000 ' Typ oder Element ist veraltet
            CompuMaster.Data.DataTables.SortColumns(Table, Table.Columns(4), Table.Columns(3), Table.Columns(2), Table.Columns(1))
#Enable Warning BC40000 ' Typ oder Element ist veraltet
            Assert.AreEqual(New String() {"Antwort D", "Antwort C", "Antwort B", "Antwort A", "Frage", "Rubrik", "Richtige Antwort", "Erläuterung", "100 ", "200 ", "500 ", "1000 ", "5000 ", "10000 ", "20000 "}, CompuMaster.Data.DataTables.AllColumnNames(Table))

            Table = Me.TestTable2WithInvariantCultureInColumnNames
            Assert.AreEqual(New String() {"Frage", "Antwort A", "Antwort B", "Antwort C", "Antwort D", "Rubrik", "Richtige Antwort", "Erläuterung", "100 ", "200 ", "500 ", "1000 ", "5000 ", "10000 ", "20000 "}, CompuMaster.Data.DataTables.AllColumnNames(Table))
            CompuMaster.Data.DataTables.SortColumns(Table, "Antwort D", "Antwort C", "Antwort B", "Antwort A")
            Assert.AreEqual(New String() {"Antwort D", "Antwort C", "Antwort B", "Antwort A", "Frage", "Rubrik", "Richtige Antwort", "Erläuterung", "100 ", "200 ", "500 ", "1000 ", "5000 ", "10000 ", "20000 "}, CompuMaster.Data.DataTables.AllColumnNames(Table))
        End Sub

        <Test> Public Sub ReArrangeColumns()
            Dim Table As DataTable

            Table = Me.TestTable2WithInvariantCultureInColumnNames
            Assert.AreEqual(New String() {"Frage", "Antwort A", "Antwort B", "Antwort C", "Antwort D", "Rubrik", "Richtige Antwort", "Erläuterung", "100 ", "200 ", "500 ", "1000 ", "5000 ", "10000 ", "20000 "}, CompuMaster.Data.DataTables.AllColumnNames(Table))
#Disable Warning BC40000 ' Typ oder Element ist veraltet
            CompuMaster.Data.DataTables.ReArrangeColumns(Table, Table.Columns(4), Table.Columns(3), Table.Columns(2), Table.Columns(1))
#Enable Warning BC40000 ' Typ oder Element ist veraltet
            Assert.AreEqual(New String() {"Antwort D", "Antwort C", "Antwort B", "Antwort A"}, CompuMaster.Data.DataTables.AllColumnNames(Table))

            Table = Me.TestTable2WithInvariantCultureInColumnNames
            Assert.AreEqual(New String() {"Frage", "Antwort A", "Antwort B", "Antwort C", "Antwort D", "Rubrik", "Richtige Antwort", "Erläuterung", "100 ", "200 ", "500 ", "1000 ", "5000 ", "10000 ", "20000 "}, CompuMaster.Data.DataTables.AllColumnNames(Table))
            CompuMaster.Data.DataTables.ReArrangeColumns(Table.Columns(4), Table.Columns(3), Table.Columns(2), Table.Columns(1))
            Assert.AreEqual(New String() {"Antwort D", "Antwort C", "Antwort B", "Antwort A"}, CompuMaster.Data.DataTables.AllColumnNames(Table))

            Table = Me.TestTable2WithInvariantCultureInColumnNames
            Assert.AreEqual(New String() {"Frage", "Antwort A", "Antwort B", "Antwort C", "Antwort D", "Rubrik", "Richtige Antwort", "Erläuterung", "100 ", "200 ", "500 ", "1000 ", "5000 ", "10000 ", "20000 "}, CompuMaster.Data.DataTables.AllColumnNames(Table))
            CompuMaster.Data.DataTables.ReArrangeColumns(Table, "Antwort D", "Antwort C", "Antwort B", "Antwort A")
            Assert.AreEqual(New String() {"Antwort D", "Antwort C", "Antwort B", "Antwort A"}, CompuMaster.Data.DataTables.AllColumnNames(Table))
        End Sub

        <Test()> Public Sub ConvertColumnType(<NUnit.Framework.Values("de-DE", "en-US")> cultureName As String)
            Dim PreviousThreadCulture As System.Globalization.CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture
            Try
                System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.GetCultureInfo(cultureName)

                Dim Table As DataTable = Me.TestTable2WithInvariantCultureInColumnNames
                'Assert status at start
                Assert.AreEqual(New String() {"Frage", "Antwort A", "Antwort B", "Antwort C", "Antwort D", "Rubrik", "Richtige Antwort", "Erläuterung", "100 ", "200 ", "500 ", "1000 ", "5000 ", "10000 ", "20000 "}, CompuMaster.Data.DataTables.AllColumnNames(Table))
                Assert.AreEqual(GetType(Double), Table.Columns("Rubrik").DataType)
                Assert.AreEqual(GetType(Double), Table.Columns("100 ").DataType)

                'Change column type and re-assert
                CompuMaster.Data.DataTables.ConvertColumnType(Table.Columns.Item("Rubrik"), GetType(String), Function(x) If(IsDBNull(x), x, CType(x, Double).ToString))
                Assert.AreEqual(GetType(String), Table.Columns("Rubrik").DataType)
                Assert.AreEqual(New String() {"Frage", "Antwort A", "Antwort B", "Antwort C", "Antwort D", "Rubrik", "Richtige Antwort", "Erläuterung", "100 ", "200 ", "500 ", "1000 ", "5000 ", "10000 ", "20000 "}, CompuMaster.Data.DataTables.AllColumnNames(Table))

                'Change column type and re-assert
                CompuMaster.Data.DataTables.ConvertColumnType(Table.Columns.Item("Rubrik"), GetType(Integer), Function(x) If(IsDBNull(x), x, Integer.Parse(CType(x, String))))
                CompuMaster.Data.DataTables.ConvertColumnType(Table.Columns.Item("100 "), GetType(Boolean), Function(x) If(IsDBNull(x), x, CType(x, String) = "1"))
                Assert.AreEqual(GetType(Integer), Table.Columns("Rubrik").DataType)
                Assert.AreEqual(GetType(Boolean), Table.Columns("100 ").DataType)
                Assert.AreEqual(New String() {"Frage", "Antwort A", "Antwort B", "Antwort C", "Antwort D", "Rubrik", "Richtige Antwort", "Erläuterung", "100 ", "200 ", "500 ", "1000 ", "5000 ", "10000 ", "20000 "}, CompuMaster.Data.DataTables.AllColumnNames(Table))
            Finally
                System.Threading.Thread.CurrentThread.CurrentCulture = PreviousThreadCulture
            End Try
        End Sub

        <Test> Public Sub ConvertToMetaDataTable(<NUnit.Framework.Values("de-DE", "en-US")> cultureName As String)
            Dim PreviousThreadCulture As System.Globalization.CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture
            Try
                System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.GetCultureInfo(cultureName)

                Dim FullDataTable As DataTable = Me.TestTable2WithInvariantCultureInColumnNames
                FullDataTable.Columns("Frage").Caption = "Fragestellung"
                FullDataTable.Columns("Frage").Unique = True
                FullDataTable.Columns("Frage").AllowDBNull = False
                FullDataTable.Columns.Add("FirstLetterOfFrage", GetType(String), "SUBSTRING(ISNULL(Frage, ' '),1,1)")
                Dim MetaTable As DataTable = CompuMaster.Data.DataTables.ConvertToMetaDataTable(FullDataTable, EnumValues(Of CompuMaster.Data.DataTables.MetaDataFields)().ToArray)
                Dim MetaTableStringified = CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(MetaTable, CompuMaster.Data.ConvertToPlainTextTableOptions.SimpleLayout)

                'Assert status at start
                Assert.AreEqual(New String() {"Frage", "Antwort A", "Antwort B", "Antwort C", "Antwort D", "Rubrik", "Richtige Antwort", "Erläuterung", "100 ", "200 ", "500 ", "1000 ", "5000 ", "10000 ", "20000 ", "FirstLetterOfFrage"}, CompuMaster.Data.DataTables.AllColumnNames(MetaTable))
                Assert.AreEqual(GetType(String), MetaTable.Columns("Rubrik").DataType)
                Assert.AreEqual(GetType(String), MetaTable.Columns("100 ").DataType)

                'Check full result of meta information
                Console.WriteLine("## Origin data table (top 5)")
                Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(CompuMaster.Data.DataTables.CopyDataTableWithSubsetOfRows(FullDataTable, 0, 5), CompuMaster.Data.ConvertToPlainTextTableOptions.SimpleLayout))
                Console.WriteLine()
                Console.WriteLine("## Meta data table")
                Console.WriteLine(MetaTableStringified)

                Dim ExpectedMetaTableStringified As String =
                    "Frage        |Antwort A    |Antwort B    |Antwort C    |Antwort D    |Rubrik       |Richtige Antwort|Erläuterung  |100          |200          |500          |1000         |5000         |10000        |20000        |FirstLetterOfFrage               " & System.Environment.NewLine &
                    "-------------+-------------+-------------+-------------+-------------+-------------+----------------+-------------+-------------+-------------+-------------+-------------+-------------+-------------+-------------+---------------------------------" & System.Environment.NewLine &
                    "System.String|System.String|System.String|System.String|System.String|System.Double|System.String   |System.String|System.Double|System.Double|System.Double|System.Double|System.Double|System.Double|System.Double|System.String                    " & System.Environment.NewLine &
                    "Fragestellung|Antwort A    |Antwort B    |Antwort C    |Antwort D    |Rubrik       |Richtige Antwort|Erläuterung  |100          |200          |500          |1000         |5000         |10000        |20000        |FirstLetterOfFrage               " & System.Environment.NewLine &
                    "False        |True         |True         |True         |True         |True         |True            |True         |True         |True         |True         |True         |True         |True         |True         |True                             " & System.Environment.NewLine &
                    "             |             |             |             |             |             |                |             |             |             |             |             |             |             |             |SUBSTRING(ISNULL(Frage, ' '),1,1)" & System.Environment.NewLine &
                    "True         |False        |False        |False        |False        |False        |False           |False        |False        |False        |False        |False        |False        |False        |False        |False                            " & System.Environment.NewLine
                Assert.AreEqual(ExpectedMetaTableStringified, MetaTableStringified)
            Finally
                System.Threading.Thread.CurrentThread.CurrentCulture = PreviousThreadCulture
            End Try
        End Sub

        Private Shared Function EnumValues(Of EnumBaseType As Structure)() As List(Of EnumBaseType)
            Dim Result As New List(Of EnumBaseType)
            For Each Value As EnumBaseType In [Enum].GetValues(GetType(EnumBaseType))
                Result.Add(Value)
            Next
            Return Result
        End Function

        <Test> Public Sub ApplyFirstRowContentToColumnNames()
            Dim FullDataTable As DataTable

            FullDataTable = Me.TestTable2WithDisabledFirstRowContentAsColumnName
            Assert.AreEqual(New String() {"Column1", "Column2", "Column3", "Column4", "Column5", "Column6", "Column7", "Column8", "Column9", "Column10", "Column11", "Column12", "Column13", "Column14", "Column15"}, CompuMaster.Data.DataTables.AllColumnNames(FullDataTable))
            CompuMaster.Data.DataTables.ApplyFirstRowContentToColumnNames(FullDataTable)
            Assert.AreEqual(New String() {"Frage", "Antwort A", "Antwort B", "Antwort C", "Antwort D", "Rubrik", "Richtige Antwort", "Erläuterung", "100", "200", "500", "1000", "5000", "10000", "20000"}, CompuMaster.Data.DataTables.AllColumnNames(FullDataTable))

            FullDataTable = Me.TestTable2WithDisabledFirstRowContentAsColumnName
            FullDataTable.Rows(0)(1) = "Frage"
            CompuMaster.Data.DataTables.ApplyFirstRowContentToColumnNames(FullDataTable)
            Assert.AreEqual(New String() {"Frage", "Frage1", "Antwort B", "Antwort C", "Antwort D", "Rubrik", "Richtige Antwort", "Erläuterung", "100", "200", "500", "1000", "5000", "10000", "20000"}, CompuMaster.Data.DataTables.AllColumnNames(FullDataTable))
        End Sub

    End Class
#Enable Warning CA1822 ' Member als statisch markieren

End Namespace