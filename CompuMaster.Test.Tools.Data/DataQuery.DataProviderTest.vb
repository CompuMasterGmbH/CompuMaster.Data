Imports NUnit.Framework
Imports System.Collections.Generic

Namespace CompuMaster.Test.Data.DataQuery

#Disable Warning CA1822 ' Member als statisch markieren
    <TestFixture(Category:="DataQueryDataProvider")> Public Class DataQueryDataProviderTest

        <OneTimeSetUp> Public Sub LoadSystemDataAssembly()
            Dim SqlConnType As Type = GetType(System.Data.SqlClient.SqlConnection)
            CompuMaster.Data.DataQuery.AnyIDataProvider.CreateConnection(SqlConnType.Assembly.GetName.Name, SqlConnType.FullName)
            'CompuMaster.Data.DataQuery.AnyIDataProvider.CreateConnection("System.Data", "System.Data.SqlClient.SqlConnection")
        End Sub

        Private Sub LoadedAssembliesInCurrentAppDomain()
            Console.WriteLine("Assemblies:")
            Console.WriteLine(ListOfAssemblies(AppDomain.CurrentDomain.GetAssemblies))
            Console.WriteLine()
            Console.WriteLine("ReflectionOnly-Assemblies:")
            Console.WriteLine(ListOfAssemblies(AppDomain.CurrentDomain.ReflectionOnlyGetAssemblies))
        End Sub

        Private Function ListOfAssemblies(asms As Reflection.Assembly()) As String
            Dim Result As String = ""
            For MyCounter As Integer = 0 To asms.Length - 1
                Result &= asms(MyCounter).FullName & ControlChars.CrLf
            Next
            Return Result
        End Function

        <Test> Public Sub AvailableDataProvidersTest()
#Disable Warning BC42024 ' Nicht verwendete lokale Variable
            'Force assemblies being loaded into app domain
            Dim odbc As System.Data.Odbc.OdbcConnection
            Dim oledb As System.Data.OleDb.OleDbConnection
            Dim sqlclient As System.Data.SqlClient.SqlConnection
#Enable Warning BC42024 ' Nicht verwendete lokale Variable
            LoadedAssembliesInCurrentAppDomain()
            Dim providers As List(Of CompuMaster.Data.DataQuery.DataProvider) = CompuMaster.Data.DataQuery.DataProvider.AvailableDataProviders
            Console.WriteLine()
            Console.WriteLine("Found AvailableDataProviders:")
            For Each MyProvider As CompuMaster.Data.DataQuery.DataProvider In providers
                Console.WriteLine(MyProvider.Title & " - " & MyProvider.ConnectionType.FullName & " - " & MyProvider.Assembly.FullName)
            Next
            Assert.GreaterOrEqual(providers.Count, 3)
        End Sub

        <Test> Public Sub LookupDataProviderTest()
            Dim provider As CompuMaster.Data.DataQuery.DataProvider

            provider = CompuMaster.Data.DataQuery.DataProvider.LookupDataProvider("gibt es nicht")
            Assert.IsNull(provider)

            provider = CompuMaster.Data.DataQuery.DataProvider.LookupDataProvider("ODBC")
            Assert.IsNotNull(provider)
            Assert.AreEqual(GetType(System.Data.Odbc.OdbcConnection), provider.CreateConnection.GetType)
            Assert.AreEqual(GetType(System.Data.Odbc.OdbcCommand), provider.CreateCommand.GetType)
            Assert.IsNotNull(provider.CreateCommandBuilder)
            Assert.AreEqual(GetType(System.Data.Odbc.OdbcCommandBuilder), provider.CreateCommandBuilder.GetType)
            Assert.IsNotNull(provider.CreateDataAdapter)
            Assert.AreEqual(GetType(System.Data.Odbc.OdbcDataAdapter), provider.CreateDataAdapter.GetType)
            Console.WriteLine("ODBC WORKING")

            provider = CompuMaster.Data.DataQuery.DataProvider.LookupDataProvider("OleDb")
            Dim IsMonoRuntime As Boolean = Type.GetType("Mono.Runtime") IsNot Nothing
            If IsMonoRuntime AndAlso System.Environment.OSVersion.Platform <> PlatformID.Win32NT Then
                Assert.IsNull(provider)
                Console.WriteLine("OleDb NOT SUPPORTED at Mono/" & System.Environment.OSVersion.Platform.ToString)
            ElseIf Not IsMonoRuntime AndAlso System.Environment.OSVersion.Platform <> PlatformID.Win32NT Then
                Assert.IsTrue(provider.Title = "OleDb")
                Assert.Throws(Of System.PlatformNotSupportedException)(
                    Sub()
                        provider.CreateConnection()
                    End Sub)
                Console.WriteLine("OleDb NOT SUPPORTED at Non-Windows-Platform " & System.Environment.OSVersion.Platform.ToString)
            Else
                Assert.IsNotNull(provider)
                Assert.AreEqual(GetType(System.Data.OleDb.OleDbConnection), provider.CreateConnection.GetType)
                Assert.AreEqual(GetType(System.Data.OleDb.OleDbCommand), provider.CreateCommand.GetType)
                Assert.IsNotNull(provider.CreateCommandBuilder)
                Assert.AreEqual(GetType(System.Data.OleDb.OleDbCommandBuilder), provider.CreateCommandBuilder.GetType)
                Assert.IsNotNull(provider.CreateDataAdapter)
                Assert.AreEqual(GetType(System.Data.OleDb.OleDbDataAdapter), provider.CreateDataAdapter.GetType)
                Console.WriteLine("OleDb WORKING")
            End If

            provider = CompuMaster.Data.DataQuery.DataProvider.LookupDataProvider("SqlClient")
            Assert.IsNotNull(provider)
            Assert.AreEqual(GetType(System.Data.SqlClient.SqlConnection), provider.CreateConnection.GetType)
            Assert.AreEqual(GetType(System.Data.SqlClient.SqlCommand), provider.CreateCommand.GetType)
            Assert.IsNotNull(provider.CreateCommandBuilder)
            Assert.AreEqual(GetType(System.Data.SqlClient.SqlCommandBuilder), provider.CreateCommandBuilder.GetType)
            Assert.IsNotNull(provider.CreateDataAdapter)
            Assert.AreEqual(GetType(System.Data.SqlClient.SqlDataAdapter), provider.CreateDataAdapter.GetType)
            Console.WriteLine("SqlClient WORKING")
        End Sub

    End Class
#Enable Warning CA1822 ' Member als statisch markieren

End Namespace
