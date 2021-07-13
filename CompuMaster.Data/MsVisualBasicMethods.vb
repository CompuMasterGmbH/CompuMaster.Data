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
            Throw New NotImplementedException
        End Function

        Public Shared Function IsNumeric(value As Object) As Boolean
            Throw New NotImplementedException
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
            If text = Nothing Then
                Return String.Empty
            Else
                Return text.Substring(startPosition - 1)
            End If
        End Function

        Public Shared Function Mid(text As String, startPosition As Integer, length As Integer) As String
            If text = Nothing Then
                Return String.Empty
            Else
                Return text.Substring(startPosition - 1, length)
            End If
        End Function

        Public Shared Function Replace(value As String, search As String, replacement As String) As String
            If value = Nothing Then
                Return value
            Else
                Return value.Replace(search, replacement)
            End If
        End Function

        Public Shared Function Space(number As Integer) As String
            Throw New NotImplementedException
        End Function

        Public Shared Function Trim(value As String) As String
            Throw New NotImplementedException
        End Function

        Public Shared Function Len(value As String) As Integer
            If value = Nothing Then
                Return 0
            Else
                Return value.Length
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
            Return value.IndexOf(search) + 1
        End Function

        Public Shared Function LSet(value As String, length As Integer) As String
            Throw New NotImplementedException
        End Function

        Public Shared Function RSet(value As String, length As Integer) As String
            Throw New NotImplementedException
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