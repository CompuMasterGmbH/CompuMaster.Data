Option Explicit On 
Option Strict On

Namespace CompuMaster.Data.DataQuery

    ''' <summary>
    ''' Provides access to a temporary file path plus an automatic cleanup of the test file after this class instance when disposing
    ''' </summary>
    ''' <remarks></remarks>
    Friend Class TestFile
        Implements IDisposable

        Private path As String

        Public ReadOnly Property FilePath() As String
            Get
                Return path
            End Get
        End Property

        Public Enum TestFileType As Byte
            MsExcel95Xls
            MsExcel2007Xlsx
            MsAccess
        End Enum

        Public Sub New(ByVal fileType As TestFileType)
            Dim TempFile As String
            TempFile = System.IO.Path.GetTempFileName
            If fileType = TestFileType.MsExcel95Xls Then
                CompuMaster.Data.DatabaseManagement.CreateMsExcelFile(TempFile, DatabaseManagement.MsExcelFileType.MsExcel95Xls)
            ElseIf fileType = TestFileType.MsExcel2007Xlsx Then
                CompuMaster.Data.DatabaseManagement.CreateMsExcelFile(TempFile, DatabaseManagement.MsExcelFileType.MsExcel2007Xlsx)
            ElseIf fileType = TestFileType.MsAccess Then
                CompuMaster.Data.DatabaseManagement.CreateDatabaseFile(TempFile, DatabaseManagement.DatabaseFileType.MsAccess2002Mdb)
            Else
                Throw New ArgumentException("Invalid value for parameter fileType", "fileType")
            End If
            path = TempFile
        End Sub

        Public Sub Dispose() Implements System.IDisposable.Dispose
            Try
                If System.IO.File.Exists(path) Then System.IO.File.Delete(path)
            Catch
            End Try
        End Sub

    End Class

End Namespace