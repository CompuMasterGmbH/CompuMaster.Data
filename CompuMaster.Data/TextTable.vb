Option Explicit On
Option Strict On

Imports System.Convert
Imports System.Data
Imports System.Web.UI.WebControls

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


        Public Sub New(headers As System.Collections.Generic.List(Of TextRow), rows As System.Collections.Generic.List(Of TextRow))
            If headers Is Nothing Then Throw New ArgumentNullException(NameOf(headers))
            If rows Is Nothing Then Throw New ArgumentNullException(NameOf(rows))
            Me.Headers = headers
            Me.Rows = rows
        End Sub

        Public Property Headers As System.Collections.Generic.List(Of TextRow)
        Public Property Rows As System.Collections.Generic.List(Of TextRow)

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