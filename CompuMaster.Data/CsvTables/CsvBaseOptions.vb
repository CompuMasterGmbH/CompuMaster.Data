Option Explicit On
Option Strict On

Imports System.IO
Imports System.Data
Imports CompuMaster.Data.Strings
Imports System.Text

Namespace CompuMaster.Data.CsvTables

    Public Class CsvBaseOptions

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

        ''' <summary>
        ''' Read the CSV table from an URI (HTTP/HTTPS) and return the content as string
        ''' </summary>
        ''' <remarks>Due to memory consumption for buffering the full CSV table string, this method doesn't allow super-large tables</remarks>
        ''' <returns></returns>
        Friend Function GetCsvTableStringFromUri() As String
            If Me.FilePath.ToLowerInvariant.StartsWith("http://", StringComparison.Ordinal) OrElse Me.FilePath.ToLowerInvariant.StartsWith("https://", StringComparison.Ordinal) Then
                Return Utils.ReadStringDataFromUri(Me.FilePath, Me.FileEncodingName)
            Else
                Throw New System.NotSupportedException("This method is only supported for URI paths, please use " & NameOf(GetCsvTableStreamReader) & " instead")
            End If
        End Function

        ''' <summary>
        ''' Open a stream reader for the CSV table file (allows to read the CSV without loading the full content into memory, therefore it (better) supports very large tables)
        ''' </summary>
        ''' <returns></returns>
        Friend Function GetCsvTableStreamReader() As StreamReader
            If File.Exists(Me.FilePath) Then
                'do nothing for now
            ElseIf Me.FilePath.ToLowerInvariant.StartsWith("http://", StringComparison.Ordinal) OrElse Me.FilePath.ToLowerInvariant.StartsWith("https://", StringComparison.Ordinal) Then
                Throw New System.NotSupportedException("For Uri resources, please use " & NameOf(GetCsvTableStringFromUri) & " instead")
                'Return string2Reader(Utils.ReadStringDataFromUri(Me.FilePath, Me.FileEncodingName))
            Else
                Throw New System.IO.FileNotFoundException("File not found", Me.FilePath)
            End If

            If Me.FileEncoding IsNot Nothing Then
                Return New StreamReader(Me.FilePath, Me.FileEncoding)
            ElseIf Me.FileEncodingName = "" Then
                Return New StreamReader(Me.FilePath, System.Text.Encoding.Default)
            Else
                Return New StreamReader(Me.FilePath, System.Text.Encoding.GetEncoding(Me.FileEncodingName))
            End If
        End Function
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

End Namespace