Public Class AssemblyTestEnvironment

    Public Shared Function TestDirectory() As String
        Return System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly.Location)
    End Function

    Public Shared Function TestFileAbsolutePath(relativePath As String) As String
        'Return System.IO.Path.Combine(System.Environment.CurrentDirectory, "testfiles\test_for_msaccess.mdb")
        Return System.IO.Path.Combine(TestDirectory(), relativePath)
    End Function

    Public Shared Function TestFileAbsolutePath(relativePath As String, ensureDirectoryExists As Boolean, deleteExistingTargetFileIfExists As Boolean) As String
        Dim Result As String = TestFileAbsolutePath(relativePath)
        If ensureDirectoryExists Then
            Dim DirName As String = System.IO.Path.GetDirectoryName(Result)
            If System.IO.Directory.Exists(DirName) = False Then
                System.IO.Directory.CreateDirectory(DirName)
            End If
        End If
        If deleteExistingTargetFileIfExists Then
            If System.IO.File.Exists(Result) = True Then
                System.IO.File.Delete(Result)
            End If
        End If
        Return Result
    End Function

End Class
