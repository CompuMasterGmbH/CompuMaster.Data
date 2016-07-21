Public Class AssemblyTestEnvironment

    Public Shared Function TestDirectory() As String
        Return System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly.Location)
    End Function

    Public Shared Function TestFileAbsolutePath(relativePath As String) As String
        'Return System.IO.Path.Combine(System.Environment.CurrentDirectory, "testfiles\test_for_msaccess.mdb")
        Return System.IO.Path.Combine(TestDirectory(), relativePath)
    End Function

End Class
