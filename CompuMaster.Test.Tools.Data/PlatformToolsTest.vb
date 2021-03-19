Imports NUnit.Framework

Namespace CompuMaster.Test.Data

    <TestFixture(Category:="PlatformTools")> Public Class PlatformToolsTest

        <Test> Public Sub InstalledOleDbProviders()
            Dim Result As DictionaryEntry() = CompuMaster.Data.DataQuery.PlatformTools.InstalledOleDbProviders
            If Result Is Nothing Then
                'Mono .NET Framework and/or Non-Windows platforms (e.g. Linux) don't support this feature currently
                Assert.Ignore("Platform doesn't provide OleDbProvider list")
            Else
                'There should be at least 1 entry being found
                For Each item As DictionaryEntry In Result
                    Console.WriteLine(item.Key & "=" & item.Value)
                Next
                Assert.NotZero(Result.Length)
            End If
        End Sub

        <Test> Public Sub InstalledOdbcDrivers()
            Dim Result As String() = Nothing
            Try
                Result = CompuMaster.Data.DataQuery.PlatformTools.InstalledOdbcDrivers(CompuMaster.Data.DataQuery.PlatformTools.TargetPlatform.Current)
            Catch ex As NotImplementedException
                'Mono .NET Framework and/or Non-Windows platforms (e.g. Linux) don't support this feature currently
                Assert.Ignore("Platform doesn't provide OleDbProvider list")
            End Try
            'There should be at least 1 entry being found
            For Each item As String In Result
                Console.WriteLine(item)
            Next
            Assert.NotZero(Result.Length)
        End Sub

    End Class

End Namespace