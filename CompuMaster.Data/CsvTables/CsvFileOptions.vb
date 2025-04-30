Option Explicit On
Option Strict On

Imports System.IO
Imports System.Data
Imports CompuMaster.Data.Strings
Imports System.Text

Namespace CompuMaster.Data.CsvTables

    Public Class CsvFileOptions

        ''' <summary>
        ''' Create a new instance of CsvFileOptions with default encoding (System.Text.Encoding.Default)        
        ''' </summary>
        ''' <param name="filePath"></param>
        Public Sub New(filePath As String)
            Me.FilePath = filePath
            Me.FileEncoding = System.Text.Encoding.Default
        End Sub

        ''' <summary>
        ''' Create a new instance of CsvFileOptions
        ''' </summary>
        ''' <param name="filePath"></param>
        ''' <param name="fileEncodingName"></param>
        Public Sub New(filePath As String, fileEncodingName As String)
            Me.FilePath = filePath
            Me.FileEncodingName = fileEncodingName
        End Sub

        ''' <summary>
        ''' Create a new instance of CsvFileOptions
        ''' </summary>
        ''' <param name="filePath"></param>
        ''' <param name="fileEncoding"></param>
        Public Sub New(filePath As String, fileEncoding As System.Text.Encoding)
            Me.FilePath = filePath
            Me.FileEncoding = fileEncoding
        End Sub

        Private _FilePath As String
        ''' <summary>
        ''' The file path or Uri (http:// or https://) of the CSV table
        ''' </summary>
        ''' <returns></returns>
        Public Property FilePath As String
            Get
                Return _FilePath
            End Get
            Private Set(value As String)
                _FilePath = value
            End Set
        End Property

        ''' <summary>
        ''' Local file exists (remote resources are not checked)
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property FileExistsLocally As Boolean?
            Get
                If Me.IsRemoteUriResource Then
                    Return Nothing
                Else
                    Return File.Exists(Me.FilePath)
                End If
            End Get
        End Property

        Private _FileEncoding As System.Text.Encoding
        ''' <summary>
        ''' The file encoding (e.g. System.Text.Encoding.UTF8, System.Text.Encoding.Default, System.Text.Encoding.GetEncoding("iso-8859-1"), System.Text.Encoding.GetEncoding("windows-1252"), System.Text.Encoding.GetEncoding("utf-16"), System.Text.Encoding.GetEncoding("utf-32"), System.Text.Encoding.GetEncoding("ascii"))
        ''' </summary>
        ''' <returns></returns>
        Public Property FileEncoding As System.Text.Encoding
            Get
                Return _FileEncoding
            End Get
            Private Set(value As System.Text.Encoding)
                _FileEncoding = value
            End Set
        End Property

        Private _FileEncodingName As String
        ''' <summary>
        ''' The file encoding name (e.g. "utf-8", "iso-8859-1", "windows-1252", "utf-16", "utf-32", "ascii", "default")
        ''' </summary>
        ''' <returns></returns>
        Public Property FileEncodingName As String
            Get
                Return _FileEncodingName
            End Get
            Private Set(value As String)
                _FileEncodingName = value
            End Set
        End Property

        ''' <summary>
        ''' Remote resources from http:// or https:// URIs
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property IsRemoteUriResource() As Boolean
            Get
                Return PathIsRemoteUriResource(Me.FilePath)
            End Get
        End Property

        ''' <summary>
        ''' Remote resources from http:// or https:// URIs
        ''' </summary>
        ''' <returns></returns>
        Private Shared Function PathIsRemoteUriResource(path As String) As Boolean
            Return path.ToLowerInvariant.StartsWith("http://", StringComparison.Ordinal) OrElse path.ToLowerInvariant.StartsWith("https://", StringComparison.Ordinal)
        End Function

        Protected Shared Sub CheckFileExists(path As String)
            If File.Exists(path) Then
                'do nothing for now
            ElseIf PathIsRemoteUriResource(path) Then
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
            If PathIsRemoteUriResource(Me.FilePath) Then
                Dim EncodingWebName As String
                If FileEncoding Is Nothing Then
                    EncodingWebName = FileEncodingName
                Else
                    EncodingWebName = FileEncoding.WebName
                End If
                Return Utils.ReadStringDataFromUri(Me.FilePath, EncodingWebName)
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
            ElseIf PathIsRemoteUriResource(Me.FilePath) Then
                'Throw New System.NotSupportedException("For Uri resources, please use " & NameOf(GetCsvTableStringFromUri) & " instead")
                Return ConvertStringToStreamReader(Utils.ReadStringDataFromUri(Me.FilePath, Me.FileEncodingName))
            Else
                Throw New System.IO.FileNotFoundException("File not found", Me.FilePath)
            End If

            ' Datei im Read-Only-Modus öffnen
            Dim MyFileStream As New FileStream(Me.FilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
            If Me.FileEncoding IsNot Nothing Then
                Return New StreamReader(MyFileStream, Me.FileEncoding)
            ElseIf Me.FileEncodingName = "" Then
                Return New StreamReader(MyFileStream, System.Text.Encoding.Default)
            Else
                Return New StreamReader(MyFileStream, System.Text.Encoding.GetEncoding(Me.FileEncodingName))
            End If
        End Function

        Friend Shared Function ConvertStringToStreamReader(csvContent As String) As StreamReader
            Return New StreamReader(New MemoryStream(System.Text.Encoding.Unicode.GetBytes(csvContent)), System.Text.Encoding.Unicode, False)
        End Function

        ''' <summary>
        ''' The recommended DataTable's name from the file path (without file extension)
        ''' </summary>
        ''' <returns></returns>
        Friend ReadOnly Property RecommendedTableNameFromFilePath As String
            Get
                Return Path.GetFileNameWithoutExtension(Me.FilePath)
            End Get
        End Property

    End Class

End Namespace