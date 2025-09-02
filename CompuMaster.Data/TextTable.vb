Option Explicit On
Option Strict On

Imports System.Convert
Imports System.Data

Namespace CompuMaster.Data

    'WARNING/TODO: Different logic between DataTables.SuggestColumnWidths and implementation of this class
    '-> TODO: DbNullText implementation oder besser TextTableRenderOptions ?
    '-> TODO: ColumnFormatting nicht in TextTable-Constructor !?
    '-> TODO: above changes require complete redesign of arguments lists for constructor and ToString()

    ''' <summary>
    ''' A table of text cells
    ''' </summary>
    Public Class TextTable

        Public Sub New()
            Me.Headers = New System.Collections.Generic.List(Of TextRow)
            Me.Rows = New System.Collections.Generic.List(Of TextRow)
        End Sub

        Public Sub New(table As DataTable)
            Me.New(table, CType(Nothing, DataTables.DataColumnToString))
        End Sub

        Public Sub New(table As DataTable, columnFormatting As DataTables.DataColumnToString)
            Me.New
            Me.AssignHeadersData(table)
            Me.AssignRowData(table.Rows, columnFormatting)
        End Sub

        Public Sub New(rows As DataRow(), columnFormatting As DataTables.DataColumnToString)
            Me.New
            If rows.Length = 0 Then Return
            Dim Table As DataTable = rows(0).Table
            Me.AssignHeadersData(Table)
            Me.AssignRowData(rows, columnFormatting)
        End Sub

        Public Sub New(headers As System.Collections.Generic.List(Of TextRow), rows As System.Collections.Generic.List(Of TextRow))
            If headers Is Nothing Then Throw New ArgumentNullException(NameOf(headers))
            If rows Is Nothing Then Throw New ArgumentNullException(NameOf(rows))
            Me.Headers = headers
            Me.Rows = rows
        End Sub

        'Public Sub New(table As DataTable, columnFormatting As DataTables.DataColumnToString)
        '    Me.New(table, CType(Nothing, String), columnFormatting)
        'End Sub
        '
        'Public Sub New(rows As DataRow(), columnFormatting As DataTables.DataColumnToString)
        '    Me.New(rows, CType(Nothing, String), columnFormatting)
        'End Sub
        '
        'Public Sub New(table As DataTable, dbNullText As String, columnFormatting As DataTables.DataColumnToString)
        '    Me.New
        '    Me.AssignHeadersData(table)
        '    Me.AssignRowData(table.Rows, dbNullText, columnFormatting)
        'End Sub
        '
        'Public Sub New(rows As DataRow(), dbNullText As String, columnFormatting As DataTables.DataColumnToString)
        '    Me.New
        '    If rows.Length = 0 Then Return
        '    Dim Table As DataTable = rows(0).Table
        '    Me.AssignHeadersData(Table)
        '    Me.AssignRowData(rows, dbNullText, columnFormatting)
        'End Sub

        Private Sub AssignHeadersData(table As DataTable)
            Dim HeaderCells As New System.Collections.Generic.List(Of TextCell)
            For ColCounter As Integer = 0 To table.Columns.Count - 1
                HeaderCells.Add(New TextCell(Utils.StringNotEmptyOrAlternativeValue(table.Columns(ColCounter).Caption, table.Columns(ColCounter).ColumnName)))
            Next
            Me.Headers.Add(New TextRow(HeaderCells))
        End Sub

        Private Sub AssignRowData(tableRows As DataRowCollection, columnFormatting As DataTables.DataColumnToString)
            If tableRows.Count = 0 Then Return
            Dim Table As DataTable = tableRows(0).Table
            For RowCounter As Integer = 0 To tableRows.Count - 1
                Dim Row As DataRow = tableRows(RowCounter)
                AssignRowData(Row, columnFormatting)
            Next
        End Sub

        Private Sub AssignRowData(tableRows As DataRow(), columnFormatting As DataTables.DataColumnToString)
            If tableRows.Length = 0 Then Return
            Dim Table As DataTable = tableRows(0).Table
            For RowCounter As Integer = 0 To tableRows.Length - 1
                Dim Row As DataRow = tableRows(RowCounter)
                AssignRowData(Row, columnFormatting)
            Next
        End Sub

        'Private Sub AssignRowData(tableRows As DataRowCollection, dbNullText As String, columnFormatting As DataTables.DataColumnToString)
        '    If tableRows.Count = 0 Then Return
        '    Dim Table As DataTable = tableRows(0).Table
        '    For RowCounter As Integer = 0 To tableRows.Count - 1
        '        Dim Row As DataRow = tableRows(RowCounter)
        '        AssignRowData(Row, dbNullText, columnFormatting)
        '    Next
        'End Sub
        '
        'Private Sub AssignRowData(tableRows As DataRow(), dbNullText As String, columnFormatting As DataTables.DataColumnToString)
        '    If tableRows.Length = 0 Then Return
        '    Dim Table As DataTable = tableRows(0).Table
        '    For RowCounter As Integer = 0 To tableRows.Length - 1
        '        Dim Row As DataRow = tableRows(RowCounter)
        '        AssignRowData(Row, dbNullText, columnFormatting)
        '    Next
        'End Sub

        Private Sub AssignRowData(row As DataRow, columnFormatting As DataTables.DataColumnToString)
            If columnFormatting Is Nothing Then
                'Fast item copy
                Rows.Add(New TextRow(row.ItemArray))
            Else
                'Formatted item copy
                Dim Cells As New System.Collections.Generic.List(Of TextCell)
                For ColCounter As Integer = 0 To row.Table.Columns.Count - 1
                    Dim column As DataColumn = row.Table.Columns(ColCounter)
                    Dim RenderValue As Object
                    RenderValue = columnFormatting(column, row(column))
                    Cells.Add(New TextCell(String.Format(Threading.Thread.CurrentThread.CurrentCulture, "{0}", RenderValue)))
                Next
                Rows.Add(New TextRow(Cells))
            End If
        End Sub

        'Private Sub AssignRowData(row As DataRow, dbNullText As String, columnFormatting As DataTables.DataColumnToString)
        '    Dim Cells As New System.Collections.Generic.List(Of TextCell)
        '    For ColCounter As Integer = 0 To row.Table.Columns.Count - 1
        '        Dim column As DataColumn = row.Table.Columns(ColCounter)
        '        Dim RawCellValue As Object = row(column)
        '        Dim RenderValue As String
        '        If column.DataType.IsValueType AndAlso Not GetType(String).IsInstanceOfType(column.DataType) Then
        '            'number or date/time
        '            If IsDBNull(RawCellValue) AndAlso dbNullText IsNot Nothing Then
        '                RenderValue = dbNullText
        '            ElseIf columnFormatting IsNot Nothing Then
        '                RenderValue = columnFormatting(column, RawCellValue)
        '            ElseIf IsDBNull(RawCellValue) Then
        '                RenderValue = ""
        '            Else
        '                RenderValue = String.Format(Threading.Thread.CurrentThread.CurrentCulture, "{0}", RawCellValue)
        '            End If
        '        Else
        '            'string or any other object
        '            If IsDBNull(RawCellValue) AndAlso dbNullText IsNot Nothing Then
        '                RenderValue = dbNullText
        '            ElseIf columnFormatting IsNot Nothing Then
        '                RenderValue = columnFormatting(column, RawCellValue)
        '            ElseIf IsDBNull(RawCellValue) Then
        '                RenderValue = ""
        '            Else
        '                RenderValue = String.Format(Threading.Thread.CurrentThread.CurrentCulture, "{0}", RawCellValue)
        '            End If
        '        End If
        '        Cells.Add(New TextCell(String.Format(Threading.Thread.CurrentThread.CurrentCulture, "{0}", RenderValue)))
        '    Next
        '    Rows.Add(New TextRow(Cells))
        'End Sub

        ''' <summary>
        ''' Captions of the table (can be multiple rows)
        ''' </summary>
        ''' <returns></returns>
        Public Property Headers As System.Collections.Generic.List(Of TextRow)

        ''' <summary>
        ''' Rows of the table
        ''' </summary>
        ''' <returns></returns>
        Public Property Rows As System.Collections.Generic.List(Of TextRow)

        ''' <summary>
        ''' Maximum number of columns in the table (including headers and rows)
        ''' </summary>
        ''' <returns></returns>
        Public Function ColumnCount() As Integer
            Dim MaxColumns As Integer = 0
            For MyCounter As Integer = 0 To Me.Headers.Count - 1
                MaxColumns = System.Math.Max(MaxColumns, Me.Headers(MyCounter).Cells.Count)
            Next
            For MyCounter As Integer = 0 To Me.Rows.Count - 1
                MaxColumns = System.Math.Max(MaxColumns, Me.Rows(MyCounter).Cells.Count)
            Next
            Return MaxColumns
        End Function

        ''' <summary>
        ''' Convert TextTable to classic DataTable
        ''' </summary>
        ''' <returns></returns>
        Public Function ToDataTable() As DataTable
            'Collect column names from all header rows (if multiple header rows exist, combine their texts with line breaks)
            Dim ColumnNames As New System.Collections.Generic.List(Of String)
            For RowCounter As Integer = 0 To Me.Headers.Count - 1
                For ColCounter As Integer = 0 To Me.Headers(RowCounter).Cells.Count - 1
                    If ColumnNames.Count <= ColCounter Then
                        'Add new column
                        ColumnNames.Add(Me.Headers(RowCounter).Cells(ColCounter).Text)
                    Else
                        'Column already exists -> extend column name if required
                        Dim ExistingColumnName As String = ColumnNames(ColCounter)
                        If ExistingColumnName <> Nothing Then
                            ColumnNames(ColCounter) = ExistingColumnName & System.Environment.NewLine & Me.Headers(RowCounter).Cells(ColCounter).Text
                        Else
                            ColumnNames(ColCounter) = Me.Headers(RowCounter).Cells(ColCounter).Text
                        End If
                    End If
                Next
            Next

            'Create DataTable with all columns
            Dim Result As New DataTable
            For ColCounter As Integer = 0 To ColumnNames.Count - 1
                Dim UniqueColumnName As String = DataTables.LookupUniqueColumnName(Result, ColumnNames(ColCounter))
                Result.Columns.Add(UniqueColumnName, GetType(String))
            Next

            'Add all rows
            For RowCounter As Integer = 0 To Me.Rows.Count - 1
                Dim NewRow As DataRow = Result.NewRow
                For ColCounter As Integer = 0 To Me.Rows(RowCounter).Count - 1
                    If Me.Rows(RowCounter).Cells(ColCounter).Text <> Nothing Then
                        NewRow(ColCounter) = Me.Rows(RowCounter).Cells(ColCounter).Text
                    End If
                Next
                Result.Rows.Add(NewRow)
            Next
            Return Result
        End Function

        Private Shared Function OutputOptions(rowNumbering As Boolean) As CompuMaster.Data.ConvertToPlainTextTableOptions
            Dim Result = CompuMaster.Data.ConvertToPlainTextTableOptions.SimpleLayout
            Result.MinimumColumnWidth = 2
            Result.MaximumColumnWidth = 65535
            Result.RowNumbering = rowNumbering
            Return Result
        End Function

        Public Function ToPlainTextTable() As String
            Return CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(Me.ToDataTable, OutputOptions(False))
        End Function

        Public Function ToPlainTextTable(rowNumbering As Boolean) As String
            Return CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(Me.ToDataTable, OutputOptions(rowNumbering))
        End Function

        ''' <summary>
        ''' Convert to plain text table with Excel-like column names (A, B, ..., Z, AA, AB, ..., AZ, BA, BB, ...) and row numbers (1-based)
        ''' </summary>
        ''' <returns></returns>
        Public Function ToPlainTextExcelTable() As String
            Return CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(Me.ToExcelStyleTextTable.ToDataTable, OutputOptions(False))
        End Function

        ''' <summary>
        ''' Convert to TextTable with Excel-like column names (A, B, ..., Z, AA, AB, ..., AZ, BA, BB, ...) and row numbers (1-based)
        ''' </summary>
        ''' <returns></returns>
        Public Function ToExcelStyleTextTable() As TextTable
            Return ToExcelStyleTextTable(False)
        End Function

        ''' <summary>
        ''' Convert to TextTable with Excel-like column names (A, B, ..., Z, AA, AB, ..., AZ, BA, BB, ...) and row numbers (1-based)
        ''' </summary>
        ''' <param name="replaceColumnHeaders">True for dropping existing column headers and replace them with Excel style column names, False for moving the existing column headers into regular data rows</param>
        ''' <returns></returns>
        Public Function ToExcelStyleTextTable(replaceColumnHeaders As Boolean) As TextTable
            'Prepare new header row with column letters
            Dim NewHeaderRows As New System.Collections.Generic.List(Of TextRow)
            With Nothing
                'Setup column names in letters
                Dim NewHeaderCells As New TextRow
                Dim MaxColumns As Integer = Me.ColumnCount()
                For MyCounter As Integer = 0 To MaxColumns - 1
                    NewHeaderCells.Cells.Add(New TextCell(ExcelColumnName(MyCounter)))
                Next
                NewHeaderRows.Add(NewHeaderCells)
            End With

            'Prepare new table data
            Dim NewDataRows As New System.Collections.Generic.List(Of TextRow)
            If replaceColumnHeaders = False Then
                'Add all existing header rows as regular data rows
                For RowCounter As Integer = 0 To Me.Headers.Count - 1
                    NewDataRows.Add(Me.Headers(RowCounter).Clone)
                Next
            End If
            With Nothing
                'Add all existing data rows
                For RowCounter As Integer = 0 To Me.Rows.Count - 1
                    NewDataRows.Add(Me.Rows(RowCounter).Clone)
                Next
            End With

            'Create new table
            Dim Result As TextTable = New TextTable()
            Result.Headers = NewHeaderRows
            Result.Rows = NewDataRows

            'Setup row numbers 1-based
            Result.ApplyRowNumbering()

            Return Result
        End Function

        ''' <summary>
        ''' Calculate Excel-like column name (A, B, ..., Z, AA, AB, ..., AZ, BA, BB, ...) for given 0-based column index
        ''' </summary>
        ''' <param name="columnIndex"></param>
        ''' <returns></returns>
        Friend Shared ReadOnly Property ExcelColumnName(columnIndex As Integer) As String
            Get
                If columnIndex < 0 Then Throw New ArgumentOutOfRangeException(NameOf(columnIndex), "Must be a positive value")
                Dim x As Integer = columnIndex + 1
                If x >= 1 AndAlso x <= 26 Then
                    Return Char.ConvertFromUtf32(x + 64)
                Else
                    Return ExcelColumnName(CType(((x - x Mod 26) / 26) - 1, Integer)) & Char.ConvertFromUtf32((x Mod 26) + 64)
                End If
            End Get
        End Property

        ''' <summary>
        ''' Creates a copy of the current table
        ''' </summary>
        ''' <returns></returns>
        Public Function Clone() As TextTable
            Dim NewHeaders As New System.Collections.Generic.List(Of TextRow)
            For MyCounter As Integer = 0 To Me.Headers.Count - 1
                NewHeaders.Add(Me.Headers(MyCounter).Clone)
            Next
            Dim NewRows As New System.Collections.Generic.List(Of TextRow)
            For MyCounter As Integer = 0 To Me.Rows.Count - 1
                NewRows.Add(Me.Rows(MyCounter).Clone)
            Next
            Return New TextTable(NewHeaders, NewRows)
        End Function

        ''' <summary>
        ''' Text representation of table
        ''' </summary>
        ''' <returns></returns>
        Public Overrides Function ToString() As String
            Return Me.ToString(System.Environment.NewLine, System.Environment.NewLine, "", "    ")
        End Function

        ''' <summary>
        ''' Text representation of table
        ''' </summary>
        ''' <param name="rowLineBreak">Line break before next row</param>
        ''' <param name="cellLineBreak">When cells contain line breaks, use this line break at end of line (not at end of row!)</param>
        ''' <param name="dbNullText">Text representation of DbNull.Value, e.g. empty space or a string like "NULL"</param>
        ''' <param name="tabText">Text representation of TAB char, e.g. 4 spaces</param>
        ''' <returns></returns>
        Public Overloads Function ToString(rowLineBreak As String, cellLineBreak As String, dbNullText As String, tabText As String) As String
            Return Me.ToString(rowLineBreak, cellLineBreak, dbNullText, tabText, "═"c, "─"c, "╪", "┼", "│", "│")
        End Function

        ''' <summary>
        ''' Text representation of table
        ''' </summary>
        ''' <param name="rowLineBreak">Line break before next row</param>
        ''' <param name="cellLineBreak">When cells contain line breaks, use this line break at end of line (not at end of row!)</param>
        ''' <param name="dbNullText">Text representation of DbNull.Value, e.g. empty space or a string like "NULL"</param>
        ''' <param name="tabText">Text representation of TAB char, e.g. 4 spaces</param>
        ''' <returns></returns>
        Public Overloads Function ToString(rowLineBreak As String, cellLineBreak As String, dbNullText As String, tabText As String,
                                           verticalSeparatorAfterHeader As Char?, verticalSeparatorCells As Char?,
                                           crossSeparatorAfterHeader As String, crossSeparatorCells As String,
                                           horizontalSeparatorHeadline As String, horizontalSeparatorCells As String
                                          ) As String
            Return Me.ToString(CType(Nothing, Integer()), rowLineBreak, cellLineBreak, dbNullText, tabText, "",
                               verticalSeparatorAfterHeader, verticalSeparatorCells,
                               crossSeparatorAfterHeader, crossSeparatorCells,
                               horizontalSeparatorHeadline, horizontalSeparatorCells
                               )
        End Function

        ''' <summary>
        ''' Text representation of table
        ''' </summary>
        ''' <param name="rowLineBreak">Line break before next row</param>
        ''' <param name="cellLineBreak">When cells contain line breaks, use this line break at end of line (not at end of row!)</param>
        ''' <param name="dbNullText">Text representation of DbNull.Value, e.g. empty space or a string like "NULL"</param>
        ''' <param name="tabText">Text representation of TAB char, e.g. 4 spaces</param>
        ''' <returns></returns>
        Public Overloads Function ToString(columnWidths As Integer(), rowLineBreak As String, cellLineBreak As String, dbNullText As String, tabText As String, suffixIfCellValueIsTooLong As String,
                                           verticalSeparatorAfterHeader As Char?, verticalSeparatorCells As Char?,
                                           crossSeparatorAfterHeader As String, crossSeparatorForCells As String,
                                           horizontalSeparatorAfterHeader As String, horizontalSeparatorForCells As String
                                          ) As String
            Dim Result As New System.Text.StringBuilder
            If columnWidths Is Nothing Then
                columnWidths = SuggestedColumnWidths(dbNullText, tabText)
            End If
            'Add headers
            For HeaderCounter As Integer = 0 To Me.Headers.Count - 1
                Me.Headers(HeaderCounter).AppendEncodedString(Result,
                                                       CellOutputDirection.Standard, CellContentHorizontalAlignment.Left, CellContentVerticalAlignment.Top, " "c,
                                                       columnWidths, horizontalSeparatorAfterHeader, cellLineBreak, dbNullText, tabText, suffixIfCellValueIsTooLong)
                Result.Append(rowLineBreak)
            Next
            'Add header separator line
            If verticalSeparatorAfterHeader.HasValue Then
                'Add header separator
                Dim LineSeparatorHeader As String = ""
                For ColCounter As Integer = 0 To columnWidths.Length - 1
#Disable Warning S1643 ' Strings should not be concatenated using "+" or "&" in a loop
                    If ColCounter <> 0 Then LineSeparatorHeader &= Utils.StringNotEmptyOrAlternativeValue(crossSeparatorAfterHeader, horizontalSeparatorAfterHeader)
                    LineSeparatorHeader &= New String(verticalSeparatorAfterHeader.Value, columnWidths(ColCounter))
#Enable Warning S1643 ' Strings should not be concatenated using "+" or "&" in a loop
                Next
                Result.Append(LineSeparatorHeader)
                Result.Append(rowLineBreak)
            End If
            'Add row data
            For RowCounter As Integer = 0 To Me.Rows.Count - 1
                'Add row lines
                Me.Rows(RowCounter).AppendEncodedString(Result,
                                                       CellOutputDirection.Standard, CellContentHorizontalAlignment.Left, CellContentVerticalAlignment.Top, " "c,
                                                       columnWidths, horizontalSeparatorForCells, cellLineBreak, dbNullText, tabText, suffixIfCellValueIsTooLong)
                If RowCounter <> Me.Rows.Count - 1 Then
                    Result.Append(rowLineBreak)
                End If
                'Add separator line
                If RowCounter <> Me.Rows.Count - 1 AndAlso verticalSeparatorCells.HasValue Then
                    'Add lines in between of the cells area
                    Dim LineSeparatorCells As String = ""
                    For ColCounter As Integer = 0 To columnWidths.Length - 1
#Disable Warning S1643 ' Strings should not be concatenated using "+" or "&" in a loop
                        If ColCounter <> 0 Then LineSeparatorCells &= Utils.StringNotEmptyOrAlternativeValue(crossSeparatorForCells, horizontalSeparatorForCells)
                        LineSeparatorCells &= New String(verticalSeparatorCells.Value, columnWidths(ColCounter))
#Enable Warning S1643 ' Strings should not be concatenated using "+" or "&" in a loop
                    Next
                    Result.Append(LineSeparatorCells)
                    Result.Append(rowLineBreak)
                End If
            Next
            Return Result.ToString
        End Function

        ''' <summary>
        ''' Calculate column widths to provide enough space to contain all content of all cells
        ''' </summary>
        ''' <param name="dbNullText">Text representation of DbNull.Value, e.g. empty space or a string like "NULL"</param>
        ''' <param name="tabText">Text representation of TAB char, e.g. 4 spaces</param>
        ''' <returns></returns>
        Protected Friend Overridable Function SuggestedColumnWidths(dbNullText As String, tabText As String) As Integer()
            Dim Result As New System.Collections.Generic.List(Of Integer)
            For HeaderCounter As Integer = 0 To Me.Headers.Count - 1
                For ColumnCounter As Integer = 0 To Me.Headers(HeaderCounter).Count - 1
                    If ColumnCounter < Result.Count Then
                        Result(ColumnCounter) = System.Math.Max(Result(ColumnCounter), Me.Headers(HeaderCounter).Cells(ColumnCounter).MaxWidth(dbNullText, tabText))
                    Else
                        Result.Add(Me.Headers(HeaderCounter).Cells(ColumnCounter).MaxWidth(dbNullText, tabText))
                    End If
                Next
            Next
            For RowCounter As Integer = 0 To Me.Rows.Count - 1
                For ColumnCounter As Integer = 0 To Me.Rows(RowCounter).Count - 1
                    If ColumnCounter < Result.Count Then
                        Result(ColumnCounter) = System.Math.Max(Result(ColumnCounter), Me.Rows(RowCounter).Cells(ColumnCounter).MaxWidth(dbNullText, tabText))
                    Else
                        Result.Add(Me.Rows(RowCounter).Cells(ColumnCounter).MaxWidth(dbNullText, tabText))
                    End If
                Next
            Next
            Return Result.ToArray
        End Function

        ''' <summary>
        ''' Add a very 1st column named "#" which contains the row number
        ''' </summary>
        Public Sub ApplyRowNumbering()
            For MyRowCounter As Integer = 0 To Me.Headers.Count - 1
                Dim HeaderText As String
                If MyRowCounter = 0 Then
                    HeaderText = "#"
                Else
                    HeaderText = Nothing
                End If
                Me.Headers(MyRowCounter).Cells.Insert(0, New TextCell(HeaderText))
            Next
            For MyRowCounter As Integer = 0 To Me.Rows.Count - 1
                Me.Rows(MyRowCounter).Cells.Insert(0, New TextCell((MyRowCounter + 1).ToString(System.Globalization.CultureInfo.InvariantCulture)))
            Next
        End Sub

        ''' <summary>
        ''' Output direction for all cells
        ''' </summary>
        Public Enum CellOutputDirection As Byte
            Standard = 0
            Reversed = 1
        End Enum

        ''' <summary>
        ''' Horizontal alignment for cell content
        ''' </summary>
        Public Enum CellContentHorizontalAlignment As Byte
            Left = 0
            Right = 1
            Center = 2
        End Enum

        ''' <summary>
        ''' Vertical alignment for cell content
        ''' </summary>
        Public Enum CellContentVerticalAlignment As Byte
            Top = 0
            Bottom = 1
            Middle = 2
        End Enum

    End Class

End Namespace