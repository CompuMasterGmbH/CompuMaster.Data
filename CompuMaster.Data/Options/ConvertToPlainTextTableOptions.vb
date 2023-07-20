Option Explicit On
Option Strict On

Namespace CompuMaster.Data

    Public Class ConvertToPlainTextTableOptions

        ''' <summary>
        ''' Uses "-", "|" and "+" to separate header and row areas and to separate columns
        ''' </summary>
        ''' <returns></returns>
        Public Shared ReadOnly Property SimpleLayout As ConvertToPlainTextTableOptions
            Get
                Return New ConvertToPlainTextTableOptions() With {
                    .VerticalSeparatorAfterHeader = "-"c,
                    .VerticalSeparatorForCells = New Char?,
                    .CrossSeparatorHeader = "+",
                    .CrossSeparatorCells = Nothing,
                    .HorizontalSeparatorAfterHeader = "|",
                    .HorizontalSeparatorForCells = "|"
                }
            End Get
        End Property

        ''' <summary>
        ''' Uses "=", "|" and "+" to separate header and row areas and to separate columns, uses "-", "|" and "+" to separate rows from other rows
        ''' </summary>
        ''' <returns></returns>
        Public Shared ReadOnly Property InlineBordersLayoutAnsi As ConvertToPlainTextTableOptions
            Get
                Return New ConvertToPlainTextTableOptions() With {
                    .VerticalSeparatorAfterHeader = "="c,
                    .VerticalSeparatorForCells = "-"c,
                    .CrossSeparatorHeader = "+",
                    .CrossSeparatorCells = "+",
                    .HorizontalSeparatorAfterHeader = "|",
                    .HorizontalSeparatorForCells = "|"
                }
            End Get
        End Property

        ''' <summary>
        ''' Uses "=", "|" and "+" to separate header and row areas and to separate columns, uses "-", "|" and "+" to separate rows from other rows
        ''' </summary>
        ''' <returns></returns>
        Public Shared ReadOnly Property InlineBordersLayoutNice As ConvertToPlainTextTableOptions
            Get
                Return New ConvertToPlainTextTableOptions() With {
                    .VerticalSeparatorAfterHeader = "═"c,
                    .VerticalSeparatorForCells = "─"c,
                    .CrossSeparatorHeader = "╪",
                    .CrossSeparatorCells = "┼",
                    .HorizontalSeparatorAfterHeader = "│",
                    .HorizontalSeparatorForCells = "│"
                }
            End Get
        End Property
        ''' <summary>
        ''' An optional label before the table itself
        ''' </summary>
        ''' <returns></returns>
        Public Property TableTitle As String

        ''' <summary>
        ''' Predefined widths for the several columns, where missing values will be suggested based on cell content length
        ''' </summary>
        ''' <returns></returns>
        Public Property FixedColumnWidths As Integer?()

        ''' <summary>
        ''' A horizontal separator string to be used between columns in vertical separator line after the header 
        ''' </summary>
        ''' <returns></returns>
        Public Property HorizontalSeparatorAfterHeader As String = "|"
        ''' <summary>
        ''' A horizontal separator string to be used between columns in row data lines
        ''' </summary>
        ''' <returns></returns>
        Public Property HorizontalSeparatorForCells As String = "|"
        ''' <summary>
        ''' A horizontal separator string to be used between columns in vertical separator lines after the header
        ''' </summary>
        ''' <returns></returns>
        Public Property CrossSeparatorHeader As String = "+"
        ''' <summary>
        ''' A horizontal separator string to be used between columns in vertical separator lines for cells
        ''' </summary>
        ''' <returns></returns>
        Public Property CrossSeparatorCells As String = "+"
        ''' <summary>
        ''' A vertical separator char to be used below of header, e.g. "="
        ''' </summary>
        ''' <returns></returns>
        Public Property VerticalSeparatorAfterHeader As Char? = "-"c
        ''' <summary>
        ''' A vertical separator char to be used between rows, e.g. "-"
        ''' </summary>
        ''' <returns></returns>
        Public Property VerticalSeparatorForCells As Char?

        ''' <summary>
        ''' A suffix which is used when truncating a cell's value if it's too long to fit into a columns width
        ''' </summary>
        ''' <returns></returns>
        Public Property SuffixIfValueMustBeShortened As String = "..."

        Public Property MinimumColumnWidth As Integer?

        Public Property MaximumColumnWidth As Integer?

        Private _OptimalWidthIsFoundWhenPercentageNumberOfRowsFitIntoCell As Double = 100.0
        ''' <summary>
        ''' Suggests column widths for a table using as minimum 2 chars, but minimum header string length, but also either full cell length for number/date/time columns or for all other types this percentage value of all values should be visible completely
        ''' </summary>
        ''' <returns></returns>
        Public Property OptimalWidthIsFoundWhenPercentageNumberOfRowsFitIntoCell As Double
            Get
                Return _OptimalWidthIsFoundWhenPercentageNumberOfRowsFitIntoCell
            End Get
            Set(value As Double)
                If value > 100.0 OrElse value < 0.0 Then Throw New ArgumentOutOfRangeException(NameOf(value), "Must be a value between 0 - 100")
                _OptimalWidthIsFoundWhenPercentageNumberOfRowsFitIntoCell = value
            End Set
        End Property

        ''' <summary>
        ''' Displayed message when no rows are existing (null/Nothing/String.Empty will lead to no message at all)
        ''' </summary>
        ''' <returns></returns>
        Public Property NoRowsFoundMessage As String = "no rows found"

        ''' <summary>
        ''' A text which is used if a cell is DbNull
        ''' </summary>
        ''' <remarks>ColumnFormatting will be called for a DbNull cell only if DbNullText is not assigned; String.Empty prevents calling of ColumnFormatting method</remarks>
        ''' <returns></returns>
        Public Property DbNullText As String

        ''' <summary>
        ''' A callable method to format a value to the desired string representation
        ''' </summary>
        ''' <returns></returns>
        Public Property ColumnFormatting As DataTables.DataColumnToString

        ''' <summary>
        ''' Validate specified options data
        ''' </summary>
        Friend Sub Validate()
            If Len(HorizontalSeparatorForCells) <> Len(HorizontalSeparatorAfterHeader) Then Throw New ArgumentException("Length of verticalSeparatorHeader and verticalSeparatorCells must be equal")
            If (VerticalSeparatorAfterHeader.HasValue OrElse VerticalSeparatorForCells.HasValue) AndAlso Len(CrossSeparatorHeader) <> Len(HorizontalSeparatorAfterHeader) Then Throw New ArgumentException("Length of verticalSeparatorHeader and crossSeparator must be equal since horizontal lines are requested")
        End Sub

        Private Shared Function Len(text As String) As Integer
            If text Is Nothing Then
                Return 0
            Else
                Return text.Length
            End If
        End Function

    End Class

End Namespace