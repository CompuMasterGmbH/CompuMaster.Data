Imports NUnit.Framework

Namespace CompuMaster.Test.Data

    <TestFixture(Category:="PlatformTools")> Public Class PlatformToolsTest

        <Test> Public Sub InstalledOleDbProviders()
            Dim Result As DictionaryEntry() = CompuMaster.Data.DataQuery.PlatformTools.InstalledOleDbProviders
            If Result Is Nothing Then
                Assert.Ignore("Platform doesn't provide OleDbProvider list")
            Else
                For Each item As DictionaryEntry In Result
                    Console.WriteLine(item.Key & "=" & item.Value)
                Next
                Assert.NotZero(Result.Length)
            End If
        End Sub

        <Test> Public Sub InstalledOdbcDrivers()
            Dim Result As String() = CompuMaster.Data.DataQuery.PlatformTools.InstalledOdbcDrivers(CompuMaster.Data.DataQuery.PlatformTools.TargetPlatform.Current)
            If Result Is Nothing Then
                Assert.Ignore("Platform doesn't provide OleDbProvider list")
            Else
                For Each item As String In Result
                    Console.WriteLine(item)
                Next
                Assert.NotZero(Result.Length)
            End If
        End Sub

    End Class

End Namespace