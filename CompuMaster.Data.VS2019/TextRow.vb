Option Explicit On
Option Strict On

Namespace CompuMaster.Data

    Public Class TextRow

        Public Sub New(row As System.Collections.Generic.List(Of TextCell))
            If row Is Nothing Then Throw New ArgumentNullException(NameOf(row))
            Me.Cells = row
        End Sub

        Public Sub New(itemArray As Object())
            If itemArray Is Nothing Then Throw New ArgumentNullException(NameOf(itemArray))
            Dim ItemCells As New System.Collections.Generic.List(Of TextCell)
            For MyCounter As Integer = 0 To itemArray.Length - 1
                ItemCells.Add(New TextCell(itemArray(MyCounter)))
            Next
            Me.Cells = ItemCells
        End Sub

        Public Sub New()
            Me.Cells = New System.Collections.Generic.List(Of TextCell)
        End Sub

        'Public Property ParentTable As TextTable

        ''' <summary>
        ''' Contents of cells
        ''' </summary>
        Public Property Cells As System.Collections.Generic.List(Of TextCell)

        ''' <summary>
        ''' Count of cells
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property Count As Integer
            Get
                Return Me.Cells.Count
            End Get
        End Property

        ''' <summary>
        ''' Text representation of row
        ''' </summary>
        ''' <param name="cellBreakNewLine"></param>
        ''' <param name="dbNullText"></param>
        ''' <returns></returns>
        Public Function ToEncodedString(columnWidths As Integer(), cellSeparator As Char, cellBreakNewLine As String, dbNullText As String, suffixIfValueMustBeShortened As String) As String
            Dim Result As New System.Text.StringBuilder
            Me.AppendEncodedString(Result,
                                          TextTable.CellOutputDirection.Standard,
                                          TextTable.CellContentHorizontalAlignment.Left, TextTable.CellContentVerticalAlignment.Top, " "c,
                                          columnWidths, cellSeparator, cellBreakNewLine, dbNullText, "    ", suffixIfValueMustBeShortened)
            Return Result.ToString
        End Function


        ''' <summary>
        ''' Text representation of row
        ''' </summary>
        ''' <param name="outputStringBuilder">Write all output to this StringBuilder instance</param>
        ''' <param name="cellDirection">Cell flow left-to-right or right-to-left</param>
        ''' <param name="textHorizontalAlignment">Horizontal alignment for cell text</param>
        ''' <param name="textVerticalAlignment">Vertical alignment for cell text</param>
        ''' <param name="textFillUpChar">When aligning cell text content, use this char (usually a space char) for spacing</param>
        ''' <param name="columnWidths">Widths of all output columns in chars</param>
        ''' <param name="cellSeparator">Separate cells with this char (Char 0 if no separator char shall be used)</param>
        ''' <param name="cellBreakNewLine">When cells contain line breaks, use this line break at end of line (not at end of row!)</param>
        ''' <param name="dbNullText">Text representation of DbNull.Value, e.g. empty space or a string like "NULL"</param>
        ''' <param name="tabText">Text representation of TAB char, e.g. 4 spaces</param>
        Public Sub AppendEncodedString(outputStringBuilder As System.Text.StringBuilder,
                                        cellDirection As TextTable.CellOutputDirection,
                                        textHorizontalAlignment As TextTable.CellContentHorizontalAlignment, textVerticalAlignment As TextTable.CellContentVerticalAlignment, textFillUpChar As Char,
                                        columnWidths As Integer(), cellSeparator As String,
                                        cellBreakNewLine As String, dbNullText As String, tabText As String, suffixIfCellValueIsTooLong As String)

            If columnWidths Is Nothing OrElse columnWidths.Length = 0 Then Throw New ArgumentNullException(NameOf(columnWidths))

            '1st scan: how many lines are required to write this row
            Dim Lines(columnWidths.Length - 1)() As String
            Dim LinesCount As Integer = 1
            For MyCounter As Integer = 0 To columnWidths.Length - 1
                If columnWidths(MyCounter) = 0 Then
                    'Column to be hidden
                    Lines(MyCounter) = New String() {}
                Else
                    'Column to be written
                    Lines(MyCounter) = Me.Cells(MyCounter).TextLines(dbNullText, tabText)
                    LinesCount = System.Math.Max(LinesCount, Lines(MyCounter).Length)
                End If
            Next

            '2nd formatted (spaced+aligned) output 
            Dim CellTextAligned(columnWidths.Length - 1) As System.Collections.Generic.List(Of String)
            For ColumnCounter As Integer = 0 To columnWidths.Length - 1
                If columnWidths(ColumnCounter) = 0 Then
                    'Column to be hidden
                Else
                    'Column to be written
                    CellTextAligned(ColumnCounter) = New System.Collections.Generic.List(Of String)(Lines(ColumnCounter))
                    'Align vertically
                    Select Case textVerticalAlignment
                        Case TextTable.CellContentVerticalAlignment.Top
                            Do While CellTextAligned(ColumnCounter).Count < LinesCount
                                'Insert empty line on bottom
                                CellTextAligned(ColumnCounter).Add("")
                            Loop
                        Case TextTable.CellContentVerticalAlignment.Bottom
                            Do While CellTextAligned(ColumnCounter).Count < LinesCount
                                'Insert empty line on top
                                CellTextAligned(ColumnCounter).Insert(0, "")
                            Loop
                        Case TextTable.CellContentVerticalAlignment.Middle
                            Do While CellTextAligned(ColumnCounter).Count < LinesCount
                                'First, add empty line below
                                CellTextAligned(ColumnCounter).Insert(0, "")
                                'If still required, add empty line above
                                If CellTextAligned(ColumnCounter).Count < LinesCount Then
                                    CellTextAligned(ColumnCounter).Add("")
                                End If
                            Loop
                        Case Else
                            Throw New ArgumentOutOfRangeException(NameOf(textVerticalAlignment))
                    End Select
                    'Align horizontally
                    Dim MaxLength As Integer = columnWidths(ColumnCounter)
                    For LineCounter As Integer = 0 To CellTextAligned(ColumnCounter).Count - 1
                        If CellTextAligned(ColumnCounter).Item(LineCounter).Length > MaxLength Then
                            'Reduce text length
                            CellTextAligned(ColumnCounter).Item(LineCounter) = DataTables.TrimStringToFixedWidth(CellTextAligned(ColumnCounter).Item(LineCounter), columnWidths(ColumnCounter), suffixIfCellValueIsTooLong)
                            CellTextAligned(ColumnCounter).Item(LineCounter) = CellTextAligned(ColumnCounter).Item(LineCounter).Substring(0, MaxLength)
                        End If
                        Select Case textHorizontalAlignment
                            Case TextTable.CellContentHorizontalAlignment.Left
                                If CellTextAligned(ColumnCounter).Item(LineCounter).Length < MaxLength Then
                                    'Add space chars from right
                                    CellTextAligned(ColumnCounter).Item(LineCounter) = CellTextAligned(ColumnCounter).Item(LineCounter) & New String(textFillUpChar, MaxLength - CellTextAligned(ColumnCounter).Item(LineCounter).Length)
                                End If
                            Case TextTable.CellContentHorizontalAlignment.Right
                                If CellTextAligned(ColumnCounter).Item(LineCounter).Length < MaxLength Then
                                    'Add space chars from right
                                    CellTextAligned(ColumnCounter).Item(LineCounter) = New String(textFillUpChar, MaxLength - CellTextAligned(ColumnCounter).Item(LineCounter).Length) & CellTextAligned(ColumnCounter).Item(LineCounter)
                                End If
                            Case TextTable.CellContentHorizontalAlignment.Center
                                If CellTextAligned(ColumnCounter).Item(LineCounter).Length < MaxLength Then
                                    'Add space chars from right
                                    Dim SpacesTotal As Integer = MaxLength - CellTextAligned(ColumnCounter).Item(LineCounter).Length
                                    Dim SpacesRight, SpacesLeft As Integer
                                    If (SpacesTotal And 1) = 0 Then
                                        'even
                                        SpacesRight = SpacesTotal \ 2
                                        SpacesLeft = SpacesRight
                                    Else
                                        'odd
                                        SpacesRight = SpacesTotal \ 2 + 1
                                        SpacesLeft = SpacesTotal \ 2
                                    End If
                                    CellTextAligned(ColumnCounter).Item(LineCounter) = New String(textFillUpChar, SpacesLeft) & CellTextAligned(ColumnCounter).Item(LineCounter) & New String(textFillUpChar, SpacesRight)
                                End If
                            Case Else
                                Throw New ArgumentOutOfRangeException(NameOf(textHorizontalAlignment))
                        End Select
                    Next
                End If
            Next

            '3rd append output
            For LineCounter As Integer = 0 To LinesCount - 1
                Dim LineStarted As Boolean = True
                Dim ColCounterStart, ColCounterEnd, ColCounterStepSize As Integer
                If cellDirection = TextTable.CellOutputDirection.Reversed Then
                    ColCounterStart = columnWidths.Length - 1
                    ColCounterEnd = 0
                    ColCounterStepSize = -1
                Else
                    ColCounterStart = 0
                    ColCounterEnd = columnWidths.Length - 1
                    ColCounterStepSize = 1
                End If
                For ColumnCounter As Integer = ColCounterStart To ColCounterEnd Step ColCounterStepSize
                    If columnWidths(ColumnCounter) = 0 Then
                        'Column to be hidden
                    Else
                        'Column to be written
                        If LineStarted = True Then
                            LineStarted = False
                        Else
                            outputStringBuilder.Append(cellSeparator)
                        End If
                        outputStringBuilder.Append(CellTextAligned(ColumnCounter).Item(LineCounter))
                    End If
                Next
                If LineCounter <> LinesCount - 1 Then
                    outputStringBuilder.Append(cellBreakNewLine)
                End If
            Next
        End Sub

    End Class

End Namespace