Option Explicit On 
Option Strict On

Namespace CompuMaster.Data

    ''' <summary>
    '''     Provides simplified read access to XLS(X/M/B) files via MS OLE-DB/ODBC Jet Providers
    ''' </summary>
    Public NotInheritable Class XlsReader

        ''' <summary>
        '''     Read from an excel file
        ''' </summary>
        ''' <param name="path">The path of the .XLS file</param>
        ''' <param name="sheetName">A name of a sheet where the read operations shall execute</param>
        ''' <returns>A new and independent datatable with the content of the sheet</returns>
        Public Shared Function ReadDataTableFromXlsFile(ByVal path As String, ByVal sheetName As String) As DataTable
            Return ReadDataTableFromXlsFile(path, sheetName, "SELECT * FROM [" & sheetName & "$]")
        End Function

        ''' <summary>
        '''     Read from an excel file
        ''' </summary>
        ''' <param name="path">The path of the .XLS file</param>
        ''' <param name="resultingDataTableName">A name for the resulting datatable</param>
        ''' <param name="querySql">A query SQL to filter the returned data, e. g. SELECT * FROM [sheetName$], SELECT * FROM [Tabelle1$A1:B10] or SELECT * FROM NamedArea"</param>
        ''' <returns>A new and independent datatable with the content of the sheet</returns>
        ''' <remarks>
        ''' </remarks>
        Public Shared Function ReadDataTableFromXlsFile(ByVal path As String, ByVal resultingDataTableName As String, ByVal querySql As String) As DataTable
            Dim MyConn As IDbConnection = CompuMaster.Data.DataQuery.Connections.MicrosoftExcelConnection(path, True, True)
            Dim MyCmd As IDbCommand = MyConn.CreateCommand
            MyCmd.CommandText = querySql
            Dim Result As DataTable = CompuMaster.Data.DataQuery.AnyIDataProvider.FillDataTable(MyCmd, DataQuery.AnyIDataProvider.Automations.AutoOpenAndCloseAndDisposeConnection, resultingDataTableName)
            For ColCounter As Integer = 0 To Result.Columns.Count - 1
                If Result.Columns(ColCounter).DataType Is GetType(String) Then
                    'Replace \n into system dependent NewLine (e.g. \r\n)
                    For RowCounter As Integer = 0 To Result.Rows.Count - 1
                        If Not IsDBNull(Result.Rows(RowCounter)(ColCounter)) Then
                            Dim value As String
                            value = CType(Result.Rows(RowCounter)(ColCounter), String)
                            If value <> Nothing AndAlso value.IndexOfAny(New Char() {ControlChars.Cr, ControlChars.Lf}) >= 0 Then
                                Result.Rows(RowCounter)(ColCounter) = value.Replace(ControlChars.CrLf, ControlChars.Lf).Replace(ControlChars.Cr, ControlChars.Lf).Replace(ControlChars.Lf, System.Environment.NewLine)
                            End If
                        End If
                    Next
                End If
            Next
            For RowDropCounter As Integer = Result.Rows.Count - 1 To LookupLastContentRowIndex(Result) + 1 Step -1
                Result.Rows.RemoveAt(RowDropCounter)
            Next
            Return Result
        End Function

        ''' <summary>
        ''' Lookup the last content row index (the last content cell might differ from Excel's special cell xlLastCell)
        ''' </summary>
        ''' <param name="table">A table</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Shared Function LookupLastContentRowIndex(ByVal table As DataTable) As Integer
            Dim autoSuggestionLastColumnIndex As Integer = table.Columns.Count - 1
            Dim autoSuggestedResult As Integer = table.Rows.Count - 1
            For rowCounter As Integer = autoSuggestedResult To 0 Step -1
                For colCounter As Integer = 0 To autoSuggestionLastColumnIndex
                    If IsEmptyCell(table, rowCounter, colCounter) = False Then
                        Return rowCounter
                    End If
                Next
            Next
            Return -1
        End Function

        ''' <summary>
        ''' Determine if a cell contains empty content
        ''' </summary>
        ''' <param name="table">A table</param>
        ''' <param name="rowIndex">The row index</param>
        ''' <param name="columnIndex">The column index</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Shared Function IsEmptyCell(ByVal table As DataTable, ByVal rowIndex As Integer, ByVal columnIndex As Integer) As Boolean
            Dim value As Object = table.Rows(rowIndex)(columnIndex)
            If value Is Nothing OrElse IsDBNull(value) Then
                Return True
            ElseIf value.GetType Is GetType(String) AndAlso CType(value, String) = Nothing Then
                Return True
            Else
                Return False
            End If
        End Function

    End Class

End Namespace
