Imports NUnit.Framework

Namespace CompuMaster.Test.Data

    <TestFixture(Category:="PlatformTools")> Public Class PlatformToolsTest

        <Test> Public Sub InstalledOleDbProviders()
            Dim IsMonoRuntime As Boolean = Type.GetType("Mono.Runtime") IsNot Nothing
            If Not IsMonoRuntime AndAlso System.Environment.OSVersion.Platform <> PlatformID.Win32NT Then
                Assert.Throws(Of System.PlatformNotSupportedException)(
                    Sub()
                        CompuMaster.Data.DataQuery.PlatformTools.InstalledOleDbProviders()
                    End Sub)
            Else
                Dim Result As DictionaryEntry() = CompuMaster.Data.DataQuery.PlatformTools.InstalledOleDbProviders
                If Result Is Nothing OrElse Result.Length = 0 Then
                    'Mono .NET Framework and/or Non-Windows platforms (e.g. Linux) don't support this feature currently
                    Assert.Ignore("Platform " & System.Environment.OSVersion.Platform.ToString & " doesn't provide OleDbProvider list")
                Else
                    'There should be at least 1 entry being found
                    For Each item As DictionaryEntry In Result
                        Console.WriteLine(item.Key & "=" & item.Value)
                    Next
                    Assert.NotZero(Result.Length)
                End If
            End If
        End Sub

        <Test> Public Sub InstalledOdbcDrivers()
            Dim Result As String() = Nothing
            Try
                Result = CompuMaster.Data.DataQuery.PlatformTools.InstalledOdbcDrivers(CompuMaster.Data.DataQuery.PlatformTools.TargetPlatform.Current)
            Catch ex As PlatformNotSupportedException
                Assert.Ignore("Platform doesn't provide OdbcDriver list")
            Catch ex As NotImplementedException
                'Mono .NET Framework and/or Non-Windows platforms (e.g. Linux) don't support this feature currently
                Assert.Ignore("Platform doesn't provide OdbcDriver list")
            End Try
            'There should be at least 1 entry being found
            For Each item As String In Result
                Console.WriteLine(item)
            Next
            Assert.NotZero(Result.Length)
        End Sub

    End Class

End Namespace