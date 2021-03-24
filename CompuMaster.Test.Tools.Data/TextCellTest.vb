Option Explicit On
Option Strict On

Imports NUnit.Framework

Namespace CompuMaster.Test.Data

    <TestFixture(Category:="TextTables")> Public Class TextCellTest

        <Test> Public Sub LineBreaksCount()
            Dim DbNullText As String = "<NULL>"
            Dim TabText As String = "    "
            Dim Cell As CompuMaster.Data.TextCell

            Cell = New CompuMaster.Data.TextCell("Line 1" & ControlChars.CrLf & ControlChars.Tab & "- Line 2 (with extra chars)" & ControlChars.Cr & "Line 3" & ControlChars.Lf & "Line 4")
            Assert.AreEqual(3, Cell.LineBreaksCount)
            Assert.AreEqual(4, Cell.TextLines(DbNullText, TabText).Length)
            Assert.AreEqual(New String() {
                            "Line 1", "    - Line 2 (with extra chars)", "Line 3", "Line 4"
                            }, Cell.TextLines(DbNullText, TabText))
            Assert.AreEqual(31, Cell.MaxWidth(DbNullText, TabText))
        End Sub

    End Class

End Namespace