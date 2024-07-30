Option Explicit On
Option Strict On

Imports System.IO
Imports System.Data
Imports CompuMaster.Data.Strings
Imports System.Text

Namespace CompuMaster.Data.CsvTables

    Public Class CsvDynamicColumnSizeBaseOptions
        Inherits CsvBaseOptions

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

End Namespace