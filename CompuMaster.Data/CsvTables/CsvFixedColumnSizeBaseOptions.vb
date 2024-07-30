Option Explicit On
Option Strict On

Imports System.IO
Imports System.Data
Imports CompuMaster.Data.Strings
Imports System.Text

Namespace CompuMaster.Data.CsvTables

    Public Class CsvFixedColumnSizeBaseOptions
        Inherits CsvBaseOptions

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

End Namespace