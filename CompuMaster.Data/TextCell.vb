Option Explicit On
Option Strict On

Imports CompuMaster.Data.Information

Namespace CompuMaster.Data

    Public Class TextCell

        Public Sub New(value As Object)
            If value Is Nothing OrElse IsDBNull(value) Then
                Me.Text = Nothing
            Else
                Me.Text = value.ToString
            End If
        End Sub

        'Public Property ParentRow As TextRow
        Public Property Text As String

        Friend Function LineBreaksCount() As Integer
            Dim Result As Integer = 0
            For MyCounter As Integer = 0 To Me.Text.Length - 1
                If Me.Text(MyCounter) = ControlChars.Cr Then
                    If Me.Text.Length >= MyCounter + 1 - 1 AndAlso Me.Text(MyCounter + 1) = ControlChars.Lf Then
                        'Count as 1 NewLine and step over to following char position
                        MyCounter += 1
                    End If
                    Result += 1
                ElseIf Me.Text(MyCounter) = ControlChars.Lf Then
                    Result += 1
                End If
            Next
            Return Result
        End Function

        ''' <summary>
        ''' The maximum number of chars in a line
        ''' </summary>
        ''' <param name="dbNullText">Text representation of DbNull.Value, e.g. empty space or a string like "NULL"</param>
        ''' <param name="tabText">Text representation of TAB char, e.g. 4 spaces</param>
        ''' <returns></returns>
        Friend Function MaxWidth(dbNullText As String, tabText As String) As Integer
            Dim TextLines As String() = Me.TextLines(dbNullText, tabText)
            Dim Result As Integer = 0
            For MyCounter As Integer = 0 To TextLines.Length - 1
                Result = System.Math.Max(Result, TextLines(MyCounter).Length)
            Next
            Return Result
        End Function

        ''' <summary>
        ''' Text representation line by line
        ''' </summary>
        ''' <param name="dbNullText">Text representation of DbNull.Value, e.g. empty space or a string like "NULL"</param>
        ''' <param name="tabText">Text representation of TAB char, e.g. 4 spaces</param>
        ''' <returns></returns>
        Friend Function TextLines(dbNullText As String, tabText As String) As String()
            If Me.Text Is Nothing Then
                Return New String() {Utils.StringNotNothingOrEmpty(dbNullText)}
            ElseIf Me.Text = Nothing Then
                Return New String() {""}
            Else
                Dim AllLineBreakStyles As String() = New String() {ControlChars.CrLf, CType(ControlChars.Cr, String), CType(ControlChars.Lf, String)}
                Return Me.Text.Replace(ControlChars.Tab, tabText).Split(AllLineBreakStyles, StringSplitOptions.None)
            End If
        End Function

    End Class

End Namespace