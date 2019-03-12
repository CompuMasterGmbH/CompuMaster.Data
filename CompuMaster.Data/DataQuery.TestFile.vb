Option Explicit On 
Option Strict On

Namespace CompuMaster.Data.DataQuery

    ''' <summary>
    ''' Provides access to a temporary file path plus an automatic cleanup of the test file after this class instance when disposing
    ''' </summary>
    ''' <remarks></remarks>
    Friend Class TestFile
        Implements IDisposable

        Private disposed As Boolean
        Private path As String

        Public ReadOnly Property FilePath() As String
            Get
                Return path
            End Get
        End Property

        Public Enum TestFileType As Byte
            MsExcel95Xls = 0
            MsExcel2007Xlsx = 1
            <Obsolete("Use MsAccessMdb instead"), System.ComponentModel.EditorBrowsable(ComponentModel.EditorBrowsableState.Never)> MsAccess = 2
            MsAccessMdb = 3
            MsAccessAccdb = 4
        End Enum

        Public Sub New(ByVal fileType As TestFileType)
            Dim TempFile As String
            TempFile = System.IO.Path.GetTempFileName()
            If fileType = TestFileType.MsExcel95Xls Then
                TempFile = TempFile & ".xls"
                CompuMaster.Data.DatabaseManagement.CreateMsExcelFile(TempFile, DatabaseManagement.MsExcelFileType.MsExcel95Xls)
            ElseIf fileType = TestFileType.MsExcel2007Xlsx Then
                TempFile = TempFile & ".xlsx"
                CompuMaster.Data.DatabaseManagement.CreateMsExcelFile(TempFile, DatabaseManagement.MsExcelFileType.MsExcel2007Xlsx)
            ElseIf fileType = TestFileType.MsAccessmdb Then
                TempFile = TempFile & ".mdb"
                CompuMaster.Data.DatabaseManagement.CreateDatabaseFile(TempFile, DatabaseManagement.DatabaseFileType.MsAccess2002Mdb)
            ElseIf fileType = TestFileType.MsAccessaccdb Then
                TempFile = TempFile & ".accdb"
                CompuMaster.Data.DatabaseManagement.CreateDatabaseFile(TempFile, DatabaseManagement.DatabaseFileType.MsAccess2007Accdb)
            Else
                Throw New ArgumentException("Invalid value for parameter fileType", "fileType")
            End If
            path = TempFile
        End Sub

        Protected Overridable Sub Dispose(disposing As Boolean)
            If Me.disposed Then
                Return
            End If
            Try
                If System.IO.File.Exists(path) Then
                    System.IO.File.Delete(path)
                End If
            Catch
            End Try
            Me.disposed = True
            Me.path = Nothing
        End Sub
        Public Sub Dispose() Implements System.IDisposable.Dispose
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub

        Protected Overrides Sub Finalize()
            Dispose(False)
        End Sub



    End Class

End Namespace