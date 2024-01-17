Option Explicit On
Option Strict On

Imports System.IO
Imports System.Data
Imports CompuMaster.Data.Strings
Imports System.Text

Namespace CompuMaster.Data.Model.Csv

    Public Class CsvOptions

        Public Class CsvOptions

            Protected Sub New(filePath As String, fileEncodingName As String)
                Me.FilePath = filePath
                Me.FileEncodingName = fileEncodingName
            End Sub

            Protected Sub New(filePath As String, fileEncoding As System.Text.Encoding)
                Me.FilePath = filePath
                Me.FileEncoding = fileEncoding
            End Sub

            Private _FilePath As String
            Public Property FilePath As String
                Get
                    Return _FilePath
                End Get
                Private Set(value As String)
                    _FilePath = value
                End Set
            End Property

            Private _FileEncoding As System.Text.Encoding
            Public Property FileEncoding As System.Text.Encoding
                Get
                    Return _FileEncoding
                End Get
                Private Set(value As System.Text.Encoding)
                    _FileEncoding = value
                End Set
            End Property

            Private _FileEncodingName As String
            Public Property FileEncodingName As String
                Get
                    Return _FileEncodingName
                End Get
                Private Set(value As String)
                    _FileEncodingName = value
                End Set
            End Property

            Protected Shared Sub CheckFileExists(path As String)
                If File.Exists(path) Then
                    'do nothing for now
                ElseIf path.ToLowerInvariant.StartsWith("http://", StringComparison.Ordinal) OrElse path.ToLowerInvariant.StartsWith("https://", StringComparison.Ordinal) Then
                    'do nothing for now
                Else
                    Throw New System.IO.FileNotFoundException("File not found", path)
                End If
            End Sub

            Public Overridable Sub Validate()
                CheckFileExists(Me.FilePath)
            End Sub

            'Friend Function GetStreamReader() As StreamReader
            '    If File.Exists(Me.FilePath) Then
            '        'do nothing for now
            '    ElseIf Me.filePath.ToLowerInvariant.StartsWith("http://", StringComparison.Ordinal) OrElse Me.filePath.ToLowerInvariant.StartsWith("https://", StringComparison.Ordinal) Then
            '        Return string2Reader(Utils.ReadStringDataFromUri(Me.FilePath, Me.FileEncodingName))
            '    Else
            '        Throw New System.IO.FileNotFoundException("File not found", Me.FilePath)
            '    End If
            '
            '    If Me.FileEncoding IsNot Nothing Then
            '        Return New StreamReader(Me.FilePath, Me.FileEncoding)
            '    ElseIf Me.FileEncodingName = "" Then
            '        Return New StreamReader(Me.FilePath, System.Text.Encoding.Default)
            '    Else
            '        Return New StreamReader(Me.FilePath, System.Text.Encoding.GetEncoding(Me.FileEncodingName))
            '    End If
            'End Function
        End Class

        Public Class CsvFixedColumnSizeOptions
            Inherits CsvOptions

            Protected Sub New(filePath As String, fileEncodingName As String)
                MyBase.New(filePath, fileEncodingName)
            End Sub

            Protected Sub New(filePath As String, fileEncoding As System.Text.Encoding)
                MyBase.New(filePath, fileEncoding)
            End Sub

            Private _FixedColumnLengths As Integer
            Public Property FixedColumnLengths As Integer
                Get
                    Return _FixedColumnLengths
                End Get
                Private Set(value As Integer)
                    _FixedColumnLengths = value
                End Set
            End Property

        End Class

        Public Class CsvDynamicColumnSizeOptions
            Inherits CsvOptions

            Protected Sub New(filePath As String, fileEncodingName As String)
                MyBase.New(filePath, fileEncodingName)
            End Sub

            Protected Sub New(filePath As String, fileEncoding As System.Text.Encoding)
                MyBase.New(filePath, fileEncoding)
            End Sub

            Private _ColumnSeparator As Char
            Public Property ColumnSeparator As Char
                Get
                    Return _ColumnSeparator
                End Get
                Private Set(value As Char)
                    _ColumnSeparator = value
                End Set
            End Property

        End Class

        'Public Class CsvFixedColumnSizeReadOptions
        '    Inherits CsvFixedColumnSizeOptions
        '
        'End Class
        '
        'Public Class CsvDynamicColumnSizeReadOptions
        '    Inherits CsvDynamicColumnSizeOptions
        '
        'End Class

        'Public Class CsvReadOptions
        '    Inherits CsvOptions
        '
        'End Class
        '
        'Public Class CsvWriteOptions
        '    Inherits CsvOptions
        '
        'End Class

    End Class

End Namespace