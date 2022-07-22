Imports System.Data

Module Program
    Sub Main(args As String())
        'Try to (not) get an OutOfMemoryException
        TestForOutOfMemoryExceptionWithHugeLogData_StringBuilderArguments()
        TestForOutOfMemoryExceptionWithHugeLogData_StringArguments()
    End Sub

    Private Sub TestForOutOfMemoryExceptionWithHugeLogData_StringBuilderArguments()
        Dim PreviewTable As DataTable = VeryLargeTable()

        System.Console.Write("Creating very large plain text table . . . ")
        Dim PreviewTextTable As String = CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(PreviewTable)
        System.Console.WriteLine("DONE")
        System.Console.Write("Creating very large HTML table . . . ")
        Dim PreviewHtmlTable As String = CompuMaster.Data.DataTables.ConvertToHtmlTable(PreviewTable)
        System.Console.WriteLine("DONE")
    End Sub

    Private Sub TestForOutOfMemoryExceptionWithHugeLogData_StringArguments()
        Dim PreviewTable As DataTable = VeryLargeTable()

        System.Console.Write("Creating very large plain text table . . . ")
        Dim PreviewTextTable As String = CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(PreviewTable)
        System.Console.WriteLine("DONE")

        System.Console.Write("Creating very large plain text table . . . ")
        Dim PreviewHtmlTable As String = CompuMaster.Data.DataTables.ConvertToHtmlTable(PreviewTable)
        System.Console.WriteLine("DONE")
    End Sub

    Private Function VeryLargeTable() As DataTable
        Static _Result As DataTable
        If _Result Is Nothing Then
            Const RequiredRowCount As Integer = 1000000
            System.Console.Write("Creating very large table with " & RequiredRowCount.ToString("#,##0") & " rows . . . ")
            Const RawCsvTable As String = "ID;Title;MainKeyword;Schlagwort;PublishOnGis;AlreadyExportedToGis;Ort;Region;Landschaft;Zeitraum;CreationDate;Column1;Column2;Remarks;Info;RemoteID;RemoteDetails;;;;;;;;;;;;;;;;" & ControlChars.CrLf &
            "1;Title data for this item;This item has got a main keyword;And there are additional tags;;;Hawaii, Isle Paradise;Beach;Beautyful ocean;ab 2000;01.01.2004 00:00;Somebody made a photo;Some kind of picture;This will be a long text to get some huge text data for getting OutOfMemoryExceptions;More info to come;ID in 3rd pary system;there was some data;;;;;;;;;;;;;;;;"
            Dim RecordCloningTemplate As DataTable = CompuMaster.Data.Csv.ReadDataTableFromCsvString(RawCsvTable, True, ";"c, """"c, False, False)
            Dim RecordCloningDataArray As Object() = RecordCloningTemplate.Rows(0).ItemArray
            Dim Result As DataTable = RecordCloningTemplate.Clone()
            For MyCounter = 1 To 1000000
                Dim NewRow As DataRow = Result.NewRow
                NewRow.ItemArray = RecordCloningDataArray
                Result.Rows.Add(NewRow)
            Next
            System.Console.WriteLine("DONE")
            _Result = Result
        End If
        Return _Result
    End Function

End Module
