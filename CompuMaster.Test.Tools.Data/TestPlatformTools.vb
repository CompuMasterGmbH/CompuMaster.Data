Imports NUnit.Framework

<TestFixture()> Public Class TestPlatformTools

    <Test> Public Sub InstalledOleDbProviders()
        Dim result As DictionaryEntry() = CompuMaster.Data.DataQuery.PlatformTools.InstalledOleDbProviders
        For Each item As DictionaryEntry In result
            Console.WriteLine(item.Key & "=" & item.Value)
        Next
    End Sub

    <Test> Public Sub InstalledOdbcDrivers()
        Dim result As String() = CompuMaster.Data.DataQuery.PlatformTools.InstalledOdbcDrivers(CompuMaster.Data.DataQuery.PlatformTools.TargetPlatform.Current)
        For Each item As String In result
            Console.WriteLine(item)
        Next
    End Sub

End Class
