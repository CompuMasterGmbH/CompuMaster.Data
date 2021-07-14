Option Explicit On
Option Strict On

Namespace CompuMaster.Data

    Friend NotInheritable Class ControlChars
        Public Const CrLf As String = Cr & Lf
        Public Const Cr As Char = ChrW(13)
        Public Const Lf As Char = ChrW(10)
        Public Const Tab As Char = ChrW(9)
    End Class

    Friend NotInheritable Class Information

        Public Shared Function IsDate(value As Object) As Boolean
            If value Is Nothing Then
                Return False
            Else
                Return value.GetType Is GetType(DateTime)
            End If
        End Function

        Public Shared Function IsNumeric(value As Object) As Boolean
            If value Is Nothing Then
                Return False
            Else
                Select Case value.GetType
                    Case GetType(Byte), GetType(Single), GetType(Double), GetType(Short), GetType(Integer), GetType(Long), GetType(Decimal), GetType(Boolean)
                        Return True
                    Case Else
                        Return False
                End Select
            End If
        End Function

        Public Shared Function IsNothing(value As Object) As Boolean
            Return value Is Nothing
        End Function

        Public Shared Function IsDBNull(value As Object) As Boolean
            Return System.DBNull.Value.Equals(value)
        End Function

        Public Enum TriState As Integer
            UseDefault = -2
            [True] = -1
            [False] = 0
        End Enum
    End Class

    Friend NotInheritable Class Strings

        Public Shared Function Mid(text As String, startPosition As Integer) As String
            If text Is Nothing Then
                Return Nothing
            ElseIf startPosition > text.Length Then
                Return ""
            Else
                Return text.Substring(startPosition - 1)
            End If
        End Function

        Public Shared Function Mid(text As String, startPosition As Integer, length As Integer) As String
            If text Is Nothing Then
                Return ""
            ElseIf startPosition > text.Length Then
                Return ""
            Else
                Dim MaxReadLength As Integer = System.Math.Min(text.Length - (startPosition - 1), length)
                Return text.Substring(startPosition - 1, MaxReadLength)
            End If
        End Function

        Public Shared Function Replace(value As String, search As String, replacement As String) As String
            If value Is Nothing Then
                Return Nothing
            ElseIf value = Nothing Then
                Return ""
            Else
                Return value.Replace(search, replacement)
            End If
        End Function

        Public Shared Function Space(number As Integer) As String
            Return StrDup(number, " "c)
        End Function

        Public Shared Function Trim(value As String) As String
            If value = Nothing Then
                Return ""
            Else
                Return value.Trim
            End If
        End Function

        Public Shared Function Len(value As Object) As Integer
            If value Is Nothing Then
                Return 0
            ElseIf value.GetType Is GetType(String) Then
                Return CType(value, String).Length
            ElseIf value.GetType Is GetType(Byte) Then
                Return 1
            ElseIf value.GetType Is GetType(Integer) Then
                Return 4
            ElseIf value.GetType Is GetType(Single) Then
                Return 4
            ElseIf value.GetType Is GetType(Long) Then
                Return 8
            ElseIf value.GetType Is GetType(Decimal) Then
                Return 8
            ElseIf value.GetType Is GetType(DateTime) Then
                Return 8
            ElseIf value.GetType Is GetType(TimeSpan) Then
                Return 0
            Else
                Throw New NotSupportedException
            End If
        End Function

        Public Shared Function StrDup(number As Integer, character As Char) As String
            Dim Result As New System.Text.StringBuilder
            For MyCounter As Integer = 0 To number - 1
                Result.Append(character)
            Next
            Return Result.ToString
        End Function

        Public Shared Function InStr(value As String, search As String) As Integer
            If value = Nothing Then
                Return 0
            Else
                Return value.IndexOf(search) + 1
            End If
        End Function

        Public Shared Function LSet(value As String, length As Integer) As String
            If value = Nothing AndAlso length > 0 Then
                Return Space(length)
            ElseIf value = Nothing Then
                Return ""
            ElseIf length = 0 Then
                Return String.Empty
            ElseIf length < value.Length Then
                Return value.Substring(0, length)
            Else
                Return value & Space(length - value.Length)
            End If
        End Function

        Public Shared Function RSet(value As String, length As Integer) As String
            If value = Nothing AndAlso length > 0 Then
                Return Space(length)
            ElseIf value = Nothing Then
                Return ""
            ElseIf length = 0 Then
                Return String.Empty
            ElseIf length < value.Length Then
                Return value.Substring(0, length)
            Else
                Return Space(length - value.Length) & value
            End If
        End Function

        Public Shared Function LCase(value As String) As String
            If value = Nothing Then
                Return String.Empty
            Else
                Return value.ToLowerInvariant()
            End If
        End Function

        Public Shared Function UCase(value As String) As String
            If value = Nothing Then
                Return String.Empty
            Else
                Return value.ToUpperInvariant()
            End If
        End Function

    End Class

End Namespace