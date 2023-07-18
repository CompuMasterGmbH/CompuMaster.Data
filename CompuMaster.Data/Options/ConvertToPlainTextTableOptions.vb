Option Explicit On
Option Strict On

Namespace CompuMaster.Data

    Public Class ConvertToPlainTextTableOptions

        ''' <summary>
        ''' Uses "-", "|" and "+" to separate header and row areas and to separate columns
        ''' </summary>
        ''' <returns></returns>
        Public Shared ReadOnly Property SimpleLayout As ConvertToPlainTextTableOptions = New ConvertToPlainTextTableOptions() With {
            .VerticalSeparatorAfterHeader = "-"c,
            .VerticalSeparatorForCells = "-"c,
            .CrossSeparatorHeader = "+",
            .CrossSeparatorCells = Nothing,
            .HorizontalSeparatorAfterHeader = "|",
            .HorizontalSeparatorForCells = "|"
        }

        ''' <summary>
        ''' Uses "=", "|" and "+" to separate header and row areas and to separate columns, uses "-", "|" and "+" to separate rows from other rows
        ''' </summary>
        ''' <returns></returns>
        Public Shared ReadOnly Property InlineBordersLayout As ConvertToPlainTextTableOptions = New ConvertToPlainTextTableOptions() With {
            .VerticalSeparatorAfterHeader = "="c,
            .VerticalSeparatorForCells = "-"c,
            .CrossSeparatorHeader = "+",
            .CrossSeparatorCells = "+",
            .HorizontalSeparatorAfterHeader = "|",
            .HorizontalSeparatorForCells = "|"
        }

        ''' <summary>
        ''' An optional label before the table itself
        ''' </summary>
        ''' <returns></returns>
        Public Property TableTitle As String

        ''' <summary>
        ''' Predefined widths for the several columns
        ''' </summary>
        ''' <returns></returns>
        Public Property FixedColumnWidths As Integer()

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

        Public Property OptimalWidthIsFoundWhenPercentageNumberOfRowsFitIntoCell As Decimal = 100D

        ''' <summary>
        ''' Displayed message when no rows are existing (null/Nothing/String.Empty will lead to no message at all)
        ''' </summary>
        ''' <returns></returns>
        Public Property NoRowsFoundMessage As String = "no rows found"

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