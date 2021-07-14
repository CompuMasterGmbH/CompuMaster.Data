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