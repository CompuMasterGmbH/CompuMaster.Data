Option Explicit On
Option Strict On

Imports NUnit.Framework
Imports System.Data

Namespace CompuMaster.Test.Data

#Disable Warning CA1822 ' Member als statisch markieren
    <TestFixture(Category:="TextTables")> Public Class TextTableTest

        <Test> Public Sub CreateFromDataTable()
            Dim TestTable As DataTable
            Dim TextTable As CompuMaster.Data.TextTable

            TestTable = TestTable1()
            TextTable = New CompuMaster.Data.TextTable(TestTable)
            Console.WriteLine(TextTable.ToString())
            Assert.AreEqual(1, TextTable.Headers.Count)
            Assert.AreEqual(TestTable.Columns.Count, TextTable.Headers(0).Count)
            Assert.AreEqual(TestTable.Rows.Count, TextTable.Rows.Count)
            Assert.AreEqual(TestTable.Columns.Count, TextTable.Rows(0).Count)
            Assert.AreEqual(TestTable.Columns.Count, TextTable.Rows(TextTable.Rows.Count - 1).Count)

            TestTable = TestTable1()
            TextTable = New CompuMaster.Data.TextTable(TestTable, Function(column As DataColumn, value As Object) As String
                                                                      If column.ColumnName = "ID" Then
                                                                          Return "ID" & value.ToString
                                                                      Else
                                                                          Return CompuMaster.Data.Utils.NoDBNull(value, CType(Nothing, String))
                                                                      End If
                                                                  End Function)
            Console.WriteLine(TextTable.ToString())
            Assert.AreEqual(1, TextTable.Headers.Count)
            Assert.AreEqual(TestTable.Columns.Count, TextTable.Headers(0).Count)
            Assert.AreEqual(TestTable.Rows.Count, TextTable.Rows.Count)
            Assert.AreEqual(TestTable.Columns.Count, TextTable.Rows(0).Count)
            Assert.AreEqual(TestTable.Columns.Count, TextTable.Rows(TextTable.Rows.Count - 1).Count)

            TestTable = TestTable2()
            TextTable = New CompuMaster.Data.TextTable(TestTable)
            Assert.AreEqual(1, TextTable.Headers.Count)
            Assert.AreEqual(TestTable.Columns.Count, TextTable.Headers(0).Count)
            Assert.AreEqual(TestTable.Rows.Count, TextTable.Rows.Count)
            Assert.AreEqual(TestTable.Columns.Count, TextTable.Rows(0).Count)
            Assert.AreEqual(TestTable.Columns.Count, TextTable.Rows(TextTable.Rows.Count - 1).Count)
        End Sub

        <Test> Public Sub TextTableToString()
            Dim TestTable As DataTable
            Dim TextTable As CompuMaster.Data.TextTable
            Dim Expected, Output As String
            TestTable = TestTable1()
            TextTable = New CompuMaster.Data.TextTable(TestTable)

            Dim ExpectedRowLineBreak As String
            Dim ExpectedCellLineBreak As String

            Output = TextTable.ToString()
            Console.WriteLine(Output)
            ExpectedRowLineBreak = System.Environment.NewLine
            ExpectedCellLineBreak = System.Environment.NewLine
            Expected =
                "ID│Value1         │Val2 " & ExpectedRowLineBreak &
                "══╪═══════════════╪═════" & ExpectedRowLineBreak &
                "1 │Hello world!   │Line1" & ExpectedCellLineBreak &
                "  │               │Line2" & ExpectedRowLineBreak &
                "──┼───────────────┼─────" & ExpectedRowLineBreak &
                "2 │Gotcha!        │     " & ExpectedRowLineBreak &
                "──┼───────────────┼─────" & ExpectedRowLineBreak &
                "3 │Hello world!   │     " & ExpectedRowLineBreak &
                "──┼───────────────┼─────" & ExpectedRowLineBreak &
                "4 │Not a duplicate│     " & ExpectedRowLineBreak &
                "──┼───────────────┼─────" & ExpectedRowLineBreak &
                "5 │Hello world!   │    T" & ExpectedRowLineBreak &
                "──┼───────────────┼─────" & ExpectedRowLineBreak &
                "6 │GOTCHA!        │     " & ExpectedRowLineBreak &
                "──┼───────────────┼─────" & ExpectedRowLineBreak &
                "7 │Gotcha!        │     "
            Assert.AreEqual(AscW("│"), AscW(Output.Substring(2, 1)))
            Assert.AreEqual(AscW(Expected.Substring(2, 1)), AscW(Output.Substring(2, 1)))
            Assert.AreEqual(Expected.Substring(0, 40), Output.Substring(0, 40))
            Assert.AreEqual(Expected, Output)

            ExpectedRowLineBreak = ControlChars.CrLf
            ExpectedCellLineBreak = ControlChars.Lf
            Output = TextTable.ToString(ExpectedRowLineBreak, ExpectedCellLineBreak, "<NULL>", "    ", "═"c, "─"c, "╬", "┼", "║", "│")
            Console.WriteLine(Output)
            Expected =
                "ID║Value1         ║Val2  " & ExpectedRowLineBreak &
                "══╬═══════════════╬══════" & ExpectedRowLineBreak &
                "1 │Hello world!   │Line1 " & ExpectedCellLineBreak &
                "  │               │Line2 " & ExpectedRowLineBreak &
                "──┼───────────────┼──────" & ExpectedRowLineBreak &
                "2 │Gotcha!        │      " & ExpectedRowLineBreak &
                "──┼───────────────┼──────" & ExpectedRowLineBreak &
                "3 │Hello world!   │<NULL>" & ExpectedRowLineBreak &
                "──┼───────────────┼──────" & ExpectedRowLineBreak &
                "4 │Not a duplicate│<NULL>" & ExpectedRowLineBreak &
                "──┼───────────────┼──────" & ExpectedRowLineBreak &
                "5 │Hello world!   │    T " & ExpectedRowLineBreak &
                "──┼───────────────┼──────" & ExpectedRowLineBreak &
                "6 │GOTCHA!        │<NULL>" & ExpectedRowLineBreak &
                "──┼───────────────┼──────" & ExpectedRowLineBreak &
                "7 │Gotcha!        │<NULL>"
            Assert.AreEqual(Expected, Output)

            ExpectedRowLineBreak = ControlChars.Lf
            ExpectedCellLineBreak = ControlChars.Cr
            Output = TextTable.ToString(ExpectedRowLineBreak, ExpectedCellLineBreak, "<NULL>", "    ", "═"c, Nothing, "╬", "┼", "║", "│")
            Console.WriteLine(Output)
            Expected =
                "ID║Value1         ║Val2  " & ExpectedRowLineBreak &
                "══╬═══════════════╬══════" & ExpectedRowLineBreak &
                "1 │Hello world!   │Line1 " & ExpectedCellLineBreak &
                "  │               │Line2 " & ExpectedRowLineBreak &
                "2 │Gotcha!        │      " & ExpectedRowLineBreak &
                "3 │Hello world!   │<NULL>" & ExpectedRowLineBreak &
                "4 │Not a duplicate│<NULL>" & ExpectedRowLineBreak &
                "5 │Hello world!   │    T " & ExpectedRowLineBreak &
                "6 │GOTCHA!        │<NULL>" & ExpectedRowLineBreak &
                "7 │Gotcha!        │<NULL>"
            Assert.AreEqual(Expected, Output)

            ExpectedRowLineBreak = ControlChars.Lf
            ExpectedCellLineBreak = ControlChars.Cr
            Output = TextTable.ToString(ExpectedRowLineBreak, ExpectedCellLineBreak, "<NULL>", "    ", Nothing, "─"c, "╬", "┼", "║", "│")
            Console.WriteLine(Output)
            Expected =
                "ID║Value1         ║Val2  " & ExpectedRowLineBreak &
                "1 │Hello world!   │Line1 " & ExpectedCellLineBreak &
                "  │               │Line2 " & ExpectedRowLineBreak &
                "──┼───────────────┼──────" & ExpectedRowLineBreak &
                "2 │Gotcha!        │      " & ExpectedRowLineBreak &
                "──┼───────────────┼──────" & ExpectedRowLineBreak &
                "3 │Hello world!   │<NULL>" & ExpectedRowLineBreak &
                "──┼───────────────┼──────" & ExpectedRowLineBreak &
                "4 │Not a duplicate│<NULL>" & ExpectedRowLineBreak &
                "──┼───────────────┼──────" & ExpectedRowLineBreak &
                "5 │Hello world!   │    T " & ExpectedRowLineBreak &
                "──┼───────────────┼──────" & ExpectedRowLineBreak &
                "6 │GOTCHA!        │<NULL>" & ExpectedRowLineBreak &
                "──┼───────────────┼──────" & ExpectedRowLineBreak &
                "7 │Gotcha!        │<NULL>"
            Assert.AreEqual(Expected, Output)

            ExpectedRowLineBreak = ControlChars.Lf
            ExpectedCellLineBreak = ControlChars.Cr
            Output = TextTable.ToString(ExpectedRowLineBreak, ExpectedCellLineBreak, "<NULL>", "    ", Nothing, Nothing, "╬", "┼", "║", "│")
            Console.WriteLine(Output)
            Expected =
                "ID║Value1         ║Val2  " & ExpectedRowLineBreak &
                "1 │Hello world!   │Line1 " & ExpectedCellLineBreak &
                "  │               │Line2 " & ExpectedRowLineBreak &
                "2 │Gotcha!        │      " & ExpectedRowLineBreak &
                "3 │Hello world!   │<NULL>" & ExpectedRowLineBreak &
                "4 │Not a duplicate│<NULL>" & ExpectedRowLineBreak &
                "5 │Hello world!   │    T " & ExpectedRowLineBreak &
                "6 │GOTCHA!        │<NULL>" & ExpectedRowLineBreak &
                "7 │Gotcha!        │<NULL>"
            Assert.AreEqual(Expected, Output)

            ExpectedRowLineBreak = ControlChars.Lf
            ExpectedCellLineBreak = ControlChars.Cr
            Output = TextTable.ToString(New Integer() {0, 15, 6}, ExpectedRowLineBreak, ExpectedCellLineBreak, "<NULL>", "    ", Nothing, Nothing, Nothing, "╬", "┼", "║", "│")
            Console.WriteLine(Output)
            Expected =
                "Value1         ║Val2  " & ExpectedRowLineBreak &
                "Hello world!   │Line1 " & ExpectedCellLineBreak &
                "               │Line2 " & ExpectedRowLineBreak &
                "Gotcha!        │      " & ExpectedRowLineBreak &
                "Hello world!   │<NULL>" & ExpectedRowLineBreak &
                "Not a duplicate│<NULL>" & ExpectedRowLineBreak &
                "Hello world!   │    T " & ExpectedRowLineBreak &
                "GOTCHA!        │<NULL>" & ExpectedRowLineBreak &
                "Gotcha!        │<NULL>"
            Assert.AreEqual(Expected, Output)

            ExpectedRowLineBreak = ControlChars.Lf
            ExpectedCellLineBreak = ControlChars.Cr
            Output = TextTable.ToString(New Integer() {0, 4, 4}, ExpectedRowLineBreak, ExpectedCellLineBreak, "<NULL>", "    ", Nothing, Nothing, Nothing, "╬", "┼", "║", "│")
            Console.WriteLine(Output)
            Expected =
                "Valu║Val2" & ExpectedRowLineBreak &
                "Hell│Line" & ExpectedCellLineBreak &
                "    │Line" & ExpectedRowLineBreak &
                "Gotc│    " & ExpectedRowLineBreak &
                "Hell│<NUL" & ExpectedRowLineBreak &
                "Not │<NUL" & ExpectedRowLineBreak &
                "Hell│    " & ExpectedRowLineBreak &
                "GOTC│<NUL" & ExpectedRowLineBreak &
                "Gotc│<NUL"
            Assert.AreEqual(Expected, Output)

            ExpectedRowLineBreak = ControlChars.Lf
            ExpectedCellLineBreak = ControlChars.Cr
            Output = TextTable.ToString(New Integer() {0, 3, 4}, ExpectedRowLineBreak, ExpectedCellLineBreak, "<NULL>", "    ", "...", Nothing, Nothing, "╬", "┼", "║", "│")
            Console.WriteLine(Output)
            Expected =
                "Val║Val2" & ExpectedRowLineBreak &
                "Hel│L..." & ExpectedCellLineBreak &
                "   │L..." & ExpectedRowLineBreak &
                "Got│    " & ExpectedRowLineBreak &
                "Hel│<..." & ExpectedRowLineBreak &
                "Not│<..." & ExpectedRowLineBreak &
                "Hel│ ..." & ExpectedRowLineBreak &
                "GOT│<..." & ExpectedRowLineBreak &
                "Got│<..."
            Assert.AreEqual(Expected, Output)
        End Sub

        <Test> Public Sub SuggestColumnWidths()
            Dim TestTable As DataTable
            Dim TextTable As CompuMaster.Data.TextTable
            TestTable = TestTable1()
            TextTable = New CompuMaster.Data.TextTable(TestTable)

            Assert.AreEqual(New Integer() {2, 15, 6}, TextTable.SuggestedColumnWidths("<NULL>", "    "))
            Assert.AreEqual(New Integer() {2, 15, 9}, TextTable.SuggestedColumnWidths("<NULL>", "        "))
            Assert.AreEqual(New Integer() {2, 15, 5}, TextTable.SuggestedColumnWidths("", ""))
        End Sub

        <Test> Public Sub ToStringWithOptionRowNumberingEnabled()
            Dim TestTable As DataTable
            Dim TextTable As CompuMaster.Data.TextTable
            Dim Expected, Output As String

            TestTable = TestTable1()
            TestTable.Columns(1).ColumnName &= ControlChars.CrLf & "Line2"
            TextTable = New CompuMaster.Data.TextTable(TestTable)
            TextTable.ApplyRowNumbering()

            Dim ExpectedRowLineBreak As String
            Dim ExpectedCellLineBreak As String

            Output = TextTable.ToString()
            Console.WriteLine(Output)

            Assert.AreEqual("#", TextTable.Headers(0).Cells(0).Text)
            For RowCounter As Integer = 0 To TextTable.Rows.Count - 1
                Assert.AreEqual(TextTable.Rows(RowCounter).Cells(1).Text, TextTable.Rows(RowCounter).Cells(0).Text)
                Assert.AreEqual((RowCounter + 1).ToString, TextTable.Rows(RowCounter).Cells(0).Text)
            Next

            ExpectedRowLineBreak = System.Environment.NewLine
            ExpectedCellLineBreak = System.Environment.NewLine
            Expected =
                "#│ID│Value1         │Val2 " & ExpectedCellLineBreak &
                " │  │Line2          │     " & ExpectedRowLineBreak &
                "═╪══╪═══════════════╪═════" & ExpectedRowLineBreak &
                "1│1 │Hello world!   │Line1" & ExpectedCellLineBreak &
                " │  │               │Line2" & ExpectedRowLineBreak &
                "─┼──┼───────────────┼─────" & ExpectedRowLineBreak &
                "2│2 │Gotcha!        │     " & ExpectedRowLineBreak &
                "─┼──┼───────────────┼─────" & ExpectedRowLineBreak &
                "3│3 │Hello world!   │     " & ExpectedRowLineBreak &
                "─┼──┼───────────────┼─────" & ExpectedRowLineBreak &
                "4│4 │Not a duplicate│     " & ExpectedRowLineBreak &
                "─┼──┼───────────────┼─────" & ExpectedRowLineBreak &
                "5│5 │Hello world!   │    T" & ExpectedRowLineBreak &
                "─┼──┼───────────────┼─────" & ExpectedRowLineBreak &
                "6│6 │GOTCHA!        │     " & ExpectedRowLineBreak &
                "─┼──┼───────────────┼─────" & ExpectedRowLineBreak &
                "7│7 │Gotcha!        │     "
            Assert.AreEqual(AscW("│"), AscW(Output.Substring(1, 1)), "Expected """ & "│" & """, but found """ & Output.Substring(1, 1) & """")
            Assert.AreEqual(AscW(Expected.Substring(1, 1)), AscW(Output.Substring(1, 1)))
            Assert.AreEqual(Expected.Substring(0, 40), Output.Substring(0, 40))
            Assert.AreEqual(Expected, Output)
        End Sub

        <Test>
        Public Sub ToExcelStyleTextTable()
            Dim TestTable As DataTable
            Dim TextTable As CompuMaster.Data.TextTable
            Dim Expected, Output As String

            TestTable = TestTable1()
            TestTable.Columns(1).ColumnName &= ControlChars.CrLf & "Line2"
            TextTable = (New CompuMaster.Data.TextTable(TestTable)).ToExcelStyleTextTable

            Dim ExpectedRowLineBreak As String
            Dim ExpectedCellLineBreak As String

            Output = TextTable.ToPlainTextTable()
            Console.WriteLine(Output)

            Assert.AreEqual("#", TextTable.Headers(0).Cells(0).Text)

            ExpectedRowLineBreak = System.Environment.NewLine
            ExpectedCellLineBreak = System.Environment.NewLine
            Expected =
                "# |A |B              |C          " & ExpectedCellLineBreak &
                "--+--+---------------+-----------" & ExpectedRowLineBreak &
                "1 |ID|Value1         |Val2       " & ExpectedRowLineBreak &
                "  |  |Line2          |           " & ExpectedCellLineBreak &
                "2 |1 |Hello world!   |Line1      " & ExpectedRowLineBreak &
                "  |  |               |Line2      " & ExpectedRowLineBreak &
                "3 |2 |Gotcha!        |           " & ExpectedRowLineBreak &
                "4 |3 |Hello world!   |           " & ExpectedRowLineBreak &
                "5 |4 |Not a duplicate|           " & ExpectedRowLineBreak &
                "6 |5 |Hello world!   |T          " & ExpectedRowLineBreak &
                "7 |6 |GOTCHA!        |           " & ExpectedRowLineBreak &
                "8 |7 |Gotcha!        |           " & ExpectedRowLineBreak
            Assert.AreEqual(Expected, Output)
        End Sub

        <Test>
        Public Sub ToPlainTextExcelTable()
            Dim TestTable As DataTable
            Dim TextTable As CompuMaster.Data.TextTable
            Dim Expected, Output As String

            TestTable = TestTable1()
            TestTable.Columns(1).ColumnName &= ControlChars.CrLf & "Line2"
            TextTable = New CompuMaster.Data.TextTable(TestTable)

            Dim ExpectedRowLineBreak As String
            Dim ExpectedCellLineBreak As String

            Output = TextTable.ToPlainTextExcelTable()
            Console.WriteLine(Output)

            ExpectedRowLineBreak = System.Environment.NewLine
            ExpectedCellLineBreak = System.Environment.NewLine
            Expected =
                "# |A |B              |C          " & ExpectedCellLineBreak &
                "--+--+---------------+-----------" & ExpectedRowLineBreak &
                "1 |ID|Value1         |Val2       " & ExpectedRowLineBreak &
                "  |  |Line2          |           " & ExpectedCellLineBreak &
                "2 |1 |Hello world!   |Line1      " & ExpectedRowLineBreak &
                "  |  |               |Line2      " & ExpectedRowLineBreak &
                "3 |2 |Gotcha!        |           " & ExpectedRowLineBreak &
                "4 |3 |Hello world!   |           " & ExpectedRowLineBreak &
                "5 |4 |Not a duplicate|           " & ExpectedRowLineBreak &
                "6 |5 |Hello world!   |T          " & ExpectedRowLineBreak &
                "7 |6 |GOTCHA!        |           " & ExpectedRowLineBreak &
                "8 |7 |Gotcha!        |           " & ExpectedRowLineBreak
            Assert.AreEqual(Expected, Output)
        End Sub

#Region "Test data"
        Private Function TestTable1() As DataTable
            Dim Result As New DataTable("test1")
            Result.Columns.Add("ID", GetType(Integer))
            Result.Columns.Add("Value1", GetType(String))
            Result.Columns.Add("Val2", GetType(String))
            Dim newRow As DataRow
            newRow = Result.NewRow
            newRow(0) = 1
            newRow(1) = "Hello world!"
            newRow(2) = "Line1" & ControlChars.Lf & "Line2"
            Result.Rows.Add(newRow)
            newRow = Result.NewRow
            newRow(0) = 2
            newRow(1) = "Gotcha!"
            newRow(2) = ""
            Result.Rows.Add(newRow)
            newRow = Result.NewRow
            newRow(0) = 3
            newRow(1) = "Hello world!"
            newRow(2) = Nothing
            Result.Rows.Add(newRow)
            newRow = Result.NewRow
            newRow(0) = 4
            newRow(1) = "Not a duplicate"
            Result.Rows.Add(newRow)
            newRow = Result.NewRow
            newRow(0) = 5
            newRow(1) = "Hello world!"
            newRow(2) = ControlChars.Tab & "T"
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
#End Region

    End Class
#Enable Warning CA1822 ' Member als statisch markieren

End Namespace