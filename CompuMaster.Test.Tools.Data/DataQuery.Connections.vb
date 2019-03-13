Imports NUnit.Framework

Namespace CompuMaster.Test.Data.DataQuery

    <TestFixture(Category:="DB Connections")> Public Class Connections

#If Not NET_1_1 Then
        <Test(), Obsolete> Public Sub DataException()
            Dim ex As New CompuMaster.Data.DataQuery.AnyIDataProvider.DataException(Nothing, Nothing)
            Assert.Pass("No exception thrown - all is perfect :-)")
        End Sub

        <Test()> Public Sub CloseAndDisposeConnectionNpgSql()
            Dim conn As System.Data.IDbConnection
            conn = New Npgsql.NpgsqlConnection
            conn.Dispose()
            CompuMaster.Data.DataQuery.AnyIDataProvider.CloseConnection(conn) 'should not throw an exception
            CompuMaster.Data.DataQuery.AnyIDataProvider.CloseAndDisposeConnection(conn) 'should not throw an exception
            Assert.AreEqual(System.Data.ConnectionState.Closed, conn.State)
            Assert.Pass("No exception thrown - all is perfect :-)")
        End Sub
        <Test()> Public Sub CloseAndDisposeConnectionMsSql()
            Dim conn As System.Data.IDbConnection
            conn = New System.Data.SqlClient.SqlConnection
            conn.Dispose()
            CompuMaster.Data.DataQuery.AnyIDataProvider.CloseConnection(conn) 'should not throw an exception
            CompuMaster.Data.DataQuery.AnyIDataProvider.CloseAndDisposeConnection(conn) 'should not throw an exception
            Assert.AreEqual(System.Data.ConnectionState.Closed, conn.State)
            Assert.Pass("No exception thrown - all is perfect :-)")
        End Sub
#End If

        <Test()> Public Sub LoadAndUseConnectionFromExternalAssembly()
            'TODO: Unabhängigkeit von spezifischer Workstation mit Lw. G:
            'TODO: Sinnvolles Testing
        End Sub

#If Not CI_Build Then
        Private Sub ReadMsAccessDatabaseMdb_Execute(path As String)
            CompuMaster.Data.DataQuery.Connections.ProbeOleDbOrOdbcProviderVerboseMode = True 'add some additional output to console
            Console.WriteLine("Trying to find appropriate data provider for platform " & PlatformDependentProcessBitNumber())
            Dim TestFile As String = AssemblyTestEnvironment.TestFileAbsolutePath(path)
            Console.WriteLine("Trying to open database: " & TestFile)
            Assert.True(System.IO.File.Exists(TestFile), "ERROR IN TEST: File not found: " & TestFile)
            Dim MyConn As IDbConnection = CompuMaster.Data.DataQuery.Connections.MicrosoftAccessConnection(TestFile)
            Console.WriteLine("Evaluated data provider connection string for current platform: " & MyConn.ConnectionString)
            Dim MyCmd As IDbCommand = MyConn.CreateCommand()
            MyCmd.CommandType = CommandType.Text
            MyCmd.CommandText = "SELECT * FROM TestData"
            Dim table As DataTable = CompuMaster.Data.DataQuery.FillDataTable(MyCmd, CompuMaster.Data.DataQuery.Automations.AutoOpenAndCloseAndDisposeConnection, "testdata")
            Assert.AreEqual(3, table.Rows.Count, "Row count for table TestData")
        End Sub

        <Test()> Public Sub ReadMsAccessDatabaseMdb()
            ReadMsAccessDatabaseMdb_Execute("testfiles\test_for_msaccess.mdb")
        End Sub

        <Test()> Public Sub ReadMsAccessDatabaseMdb2000()
            ReadMsAccessDatabaseMdb_Execute("testfiles\test_for_msaccess_2000.mdb")
        End Sub

        <Test()> Public Sub ReadMsAccessDatabaseMdb2002UpTo2003()
            ReadMsAccessDatabaseMdb_Execute("testfiles\test_for_msaccess_2002-2003.mdb")
        End Sub

        <Test()> Public Sub ReadMsAccessDatabaseAccdb()
            ReadMsAccessDatabaseMdb_Execute("testfiles\test_for_msaccess.accdb")
        End Sub

        <Test> Public Sub TextCsvConnection()
            CompuMaster.Data.DataQuery.Connections.ProbeOleDbOrOdbcProviderVerboseMode = True 'add some additional output to console
            Dim conn As IDbConnection = CompuMaster.Data.DataQuery.Connections.TextCsvConnection(AssemblyTestEnvironment.TestFileAbsolutePath("testfiles"))
            Assert.NotNull(conn, "CSV provider not found")
            Console.WriteLine("Evaluated data provider connection string for current platform: " & conn.ConnectionString)
            Dim Cmd As IDbCommand = conn.CreateCommand()
            Cmd.CommandType = CommandType.Text
            Cmd.CommandText = "SELECT * FROM [country-codes.csv]"
            Dim table As DataTable = CompuMaster.Data.DataQuery.FillDataTable(Cmd, CompuMaster.Data.DataQuery.Automations.AutoOpenAndCloseAndDisposeConnection, "testdata")
            Assert.AreEqual(3, table.Rows.Count, "Row count for table TestData")
        End Sub

        <Test()> Public Sub EnumerateTablesAndViewsInOdbcDbDataSource()
            Dim TestDir As String = AssemblyTestEnvironment.TestFileAbsolutePath("testfiles")
            Dim conn As IDbConnection = CompuMaster.Data.DataQuery.Connections.TextCsvConnection(TestDir)
            Try
                conn.Open()
                Dim tables As CompuMaster.Data.DataQuery.Connections.OdbcTableDescriptor() = CompuMaster.Data.DataQuery.Connections.EnumerateTablesAndViewsInOdbcDataSource(CType(conn, System.Data.Odbc.OdbcConnection))
                Dim tableNames As New System.Collections.Generic.List(Of String)
                Dim TestDataTable As CompuMaster.Data.DataQuery.Connections.OdbcTableDescriptor = Nothing
                For Each table As CompuMaster.Data.DataQuery.Connections.OdbcTableDescriptor In tables
                    Console.WriteLine(table.ToString)
                    tableNames.Add(table.ToString)
                    If table.ToString = "[country-codes.csv]" Then
                        TestDataTable = table
                    End If
                Next
                Assert.AreNotEqual(0, tables.Length)
                Assert.IsNotNull(TestDataTable, "Table TestData not found")
                Assert.AreEqual("country-codes.csv", TestDataTable.TableName)
                Assert.AreEqual(Nothing, TestDataTable.SchemaName)
                Assert.AreEqual("[country-codes.csv]", TestDataTable.ToString)
            Finally
                CompuMaster.Data.DataQuery.CloseAndDisposeConnection(conn)
            End Try

            Dim TestFile As String = AssemblyTestEnvironment.TestFileAbsolutePath("testfiles\test_for_msaccess.mdb")
            conn = CompuMaster.Data.DataQuery.Connections.MicrosoftAccessOdbcConnection(TestFile)
            Try
                conn.Open()
                Dim tables As CompuMaster.Data.DataQuery.Connections.OdbcTableDescriptor() = CompuMaster.Data.DataQuery.Connections.EnumerateTablesAndViewsInOdbcDataSource(CType(conn, System.Data.Odbc.OdbcConnection))
                Dim tableNames As New System.Collections.Generic.List(Of String)
                Dim TestDataTable As CompuMaster.Data.DataQuery.Connections.OdbcTableDescriptor = Nothing
                For Each table As CompuMaster.Data.DataQuery.Connections.OdbcTableDescriptor In tables
                    Console.WriteLine(table.ToString)
                    tableNames.Add(table.ToString)
                    If table.ToString = "[TestData]" Then
                        TestDataTable = table
                    End If
                Next
                Assert.AreNotEqual(0, tables.Length)
                Assert.IsNotNull(TestDataTable, "Table TestData not found")
                Assert.AreEqual("TestData", TestDataTable.TableName)
                Assert.AreEqual(Nothing, TestDataTable.SchemaName)
                Assert.AreEqual("[TestData]", TestDataTable.ToString)
            Finally
                CompuMaster.Data.DataQuery.CloseAndDisposeConnection(conn)
            End Try
        End Sub

        <Test()> Public Sub EnumerateTablesAndViewsInOleDbDataSource()
            Dim TestFile As String = AssemblyTestEnvironment.TestFileAbsolutePath("testfiles\test_for_msaccess.mdb")
            Dim conn As IDbConnection = CompuMaster.Data.DataQuery.Connections.MicrosoftAccessConnection(TestFile)
            If CType(conn, Object).GetType Is GetType(System.Data.OleDb.OleDbConnection) Then
                Try
                    conn.Open()
                    Dim tables As CompuMaster.Data.DataQuery.Connections.OleDbTableDescriptor() = CompuMaster.Data.DataQuery.Connections.EnumerateTablesAndViewsInOleDbDataSource(CType(conn, System.Data.OleDb.OleDbConnection))
                    Dim tableNames As New System.Collections.Generic.List(Of String)
                    Dim TestDataTable As CompuMaster.Data.DataQuery.Connections.OleDbTableDescriptor = Nothing
                    For Each table As CompuMaster.Data.DataQuery.Connections.OleDbTableDescriptor In tables
                        Console.WriteLine(table.ToString)
                        tableNames.Add(table.ToString)
                        If table.ToString = "[TestData]" Then
                            TestDataTable = table
                        End If
                    Next
                    Assert.AreNotEqual(0, tables.Length)
                    Assert.IsNotNull(TestDataTable, "Table TestData not found")
                    Assert.AreEqual("TestData", TestDataTable.TableName)
                    Assert.AreEqual(Nothing, TestDataTable.SchemaName)
                    Assert.AreEqual("[TestData]", TestDataTable.ToString)
                Finally
                    CompuMaster.Data.DataQuery.CloseAndDisposeConnection(conn)
                End Try
            Else
                Assert.Fail("Test environment doesn't contain OleDb provider for current platform x64/x32 - reconfigure test server!")
            End If
        End Sub
#End If

#If Not CI_Build Then
        <Test()> Public Sub MicrosoftExcelOdbcConnection()
            Console.WriteLine("Trying to find appropriate data provider for platform " & PlatformDependentProcessBitNumber())
            Dim TestFile As String = AssemblyTestEnvironment.TestFileAbsolutePath("testfiles\test_for_lastcell_e50aka95.xls")
            Dim conn As IDbConnection = CompuMaster.Data.DataQuery.Connections.MicrosoftExcelOdbcConnection(TestFile, False, True)
            Console.WriteLine("Evaluated data provider connection string for current platform: " & conn.ConnectionString)
            If CType(conn, Object).GetType Is GetType(System.Data.Odbc.OdbcConnection) Then
                Try
                    CompuMaster.Data.DataQuery.OpenConnection(conn)
                    Assert.AreNotEqual(0, CompuMaster.Data.DataQuery.Connections.EnumerateTablesAndViewsInOdbcDataSource(conn).Length)
                Finally
                    CompuMaster.Data.DataQuery.CloseAndDisposeConnection(conn)
                End Try
                Assert.Pass("Excel XLS opened at " & PlatformDependentProcessBitNumber() & " platform")
            Else
                Assert.Fail("Failed to open Excel XLS at " & PlatformDependentProcessBitNumber() & " platform")
            End If
        End Sub

        <Test()> Public Sub MicrosoftExcelConnectionMatrixByProviderAndExcelFileFormatVersion()
            Dim TestFails As Boolean = False
            Console.WriteLine("Trying to find appropriate data provider for platform " & PlatformDependentProcessBitNumber())
            Console.WriteLine()
            Dim TestFiles As New Generic.Dictionary(Of String, String)
            TestFiles.Add("XLS95", "testfiles\test_for_lastcell_e50aka95.xls")
            TestFiles.Add("XLS97", "testfiles\test_for_lastcell_e70aka97-2003.xls")
            TestFiles.Add("XLSX2007", "testfiles\test_for_lastcell_e12aka2007.xlsx")
            TestFiles.Add("XLSB2007", "testfiles\test_for_lastcell_e12aka2007.xlsb")
            TestFiles.Add("XLSM2007", "testfiles\test_for_lastcell_e12aka2007.xlsm")
            'OLE DB checks
            Console.WriteLine("Executing OLEDB checks")
            For Each TestFile As Generic.KeyValuePair(Of String, String) In TestFiles
                Dim CurrentTestFile As String = AssemblyTestEnvironment.TestFileAbsolutePath(TestFile.Value)
                Console.Write("Checking " & TestFile.Key & ": ")
                Dim FoundProviderLookupException As Exception = Nothing
                Dim conn As IDbConnection = Nothing
                Try
                    conn = CompuMaster.Data.DataQuery.Connections.MicrosoftExcelOleDbConnection(CurrentTestFile, False, True)
                Catch ex As Exception
                    FoundProviderLookupException = ex
                    TestFails = True
                End Try
                If FoundProviderLookupException IsNot Nothing Then
                    Console.WriteLine("FAILED ON PROVIDER LOOKUP: " & FoundProviderLookupException.Message)
                Else
                    Console.WriteLine(MicrosoftAccessOrExcelConnectionMatrixByProviderAndAccessOrExcelFileFormatVersion_TryOpenConnectionTest(conn, TestFails))
                End If
            Next
            'ODBC checks
            Console.WriteLine()
            Console.WriteLine("Executing ODBC checks")
            For Each TestFile As Generic.KeyValuePair(Of String, String) In TestFiles
                Dim CurrentTestFile As String = AssemblyTestEnvironment.TestFileAbsolutePath(TestFile.Value)
                Console.Write("Checking " & TestFile.Key & ": ")
                Dim FoundProviderLookupException As Exception = Nothing
                Dim conn As IDbConnection = Nothing
                Try
                    conn = CompuMaster.Data.DataQuery.Connections.MicrosoftExcelOdbcConnection(CurrentTestFile, False, True)
                Catch ex As Exception
                    FoundProviderLookupException = ex
                    TestFails = True
                End Try
                If FoundProviderLookupException IsNot Nothing Then
                    Console.WriteLine("FAILED ON PROVIDER LOOKUP: " & FoundProviderLookupException.Message)
                Else
                    Console.WriteLine(MicrosoftAccessOrExcelConnectionMatrixByProviderAndAccessOrExcelFileFormatVersion_TryOpenConnectionTest(conn, TestFails))
                End If
            Next
            'Auto-Lookup checks
            Console.WriteLine()
            Console.WriteLine("Executing Auto-Lookup checks")
            For Each TestFile As Generic.KeyValuePair(Of String, String) In TestFiles
                Dim CurrentTestFile As String = AssemblyTestEnvironment.TestFileAbsolutePath(TestFile.Value)
                Console.Write("Checking " & TestFile.Key & ": ")
                Dim FoundProviderLookupException As Exception = Nothing
                Dim conn As IDbConnection = Nothing
                Try
                    conn = CompuMaster.Data.DataQuery.Connections.MicrosoftExcelConnection(CurrentTestFile, False, True)
                Catch ex As Exception
                    FoundProviderLookupException = ex
                    TestFails = True
                End Try
                If FoundProviderLookupException IsNot Nothing Then
                    Console.WriteLine("FAILED ON PROVIDER LOOKUP: " & FoundProviderLookupException.Message)
                Else
                    Console.WriteLine(MicrosoftAccessOrExcelConnectionMatrixByProviderAndAccessOrExcelFileFormatVersion_TryOpenConnectionTest(conn, TestFails))
                End If
            Next
            If TestFails = True Then
                Assert.Fail("Some errors occured")
            Else
                Assert.Pass("All files tested successfully with OLEDB + ODBC")
            End If

        End Sub

        <Test()> Public Sub MicrosoftAccessConnectionMatrixByProviderAndAccessFileFormatVersion()
            Dim TestFails As Boolean = False
            Console.WriteLine("Trying to find appropriate data provider for platform " & PlatformDependentProcessBitNumber())
            Console.WriteLine()
            Dim TestFiles As New Generic.Dictionary(Of String, String)
            TestFiles.Add("MDB", "testfiles\test_for_msaccess.mdb")
            TestFiles.Add("MDB2000", "testfiles\test_for_msaccess_2000.mdb")
            TestFiles.Add("MDB2002-2003", "testfiles\test_for_msaccess_2002-2003.mdb")
            TestFiles.Add("ACCDB", "testfiles\test_for_msaccess.accdb")
            'OLE DB checks
            Console.WriteLine("Executing OLEDB checks")
            For Each TestFile As Generic.KeyValuePair(Of String, String) In TestFiles
                Dim CurrentTestFile As String = AssemblyTestEnvironment.TestFileAbsolutePath(TestFile.Value)
                Console.Write("Checking " & TestFile.Key & ": ")
                Dim FoundProviderLookupException As Exception = Nothing
                Dim conn As IDbConnection = Nothing
                Try
                    conn = CompuMaster.Data.DataQuery.Connections.MicrosoftAccessOleDbConnection(CurrentTestFile)
                Catch ex As Exception
                    FoundProviderLookupException = ex
                    TestFails = True
                End Try
                If FoundProviderLookupException IsNot Nothing Then
                    Console.WriteLine("FAILED ON PROVIDER LOOKUP: " & FoundProviderLookupException.Message)
                Else
                    Console.WriteLine(MicrosoftAccessOrExcelConnectionMatrixByProviderAndAccessOrExcelFileFormatVersion_TryOpenConnectionTest(conn, TestFails))
                End If
            Next
            'ODBC checks
            Console.WriteLine()
            Console.WriteLine("Executing ODBC checks")
            For Each TestFile As Generic.KeyValuePair(Of String, String) In TestFiles
                Dim CurrentTestFile As String = AssemblyTestEnvironment.TestFileAbsolutePath(TestFile.Value)
                Console.Write("Checking " & TestFile.Key & ": ")
                Dim FoundProviderLookupException As Exception = Nothing
                Dim conn As IDbConnection = Nothing
                Try
                    conn = CompuMaster.Data.DataQuery.Connections.MicrosoftAccessOdbcConnection(CurrentTestFile)
                Catch ex As Exception
                    FoundProviderLookupException = ex
                    TestFails = True
                End Try
                If FoundProviderLookupException IsNot Nothing Then
                    Console.WriteLine("FAILED ON PROVIDER LOOKUP: " & FoundProviderLookupException.Message)
                Else
                    Console.WriteLine(MicrosoftAccessOrExcelConnectionMatrixByProviderAndAccessOrExcelFileFormatVersion_TryOpenConnectionTest(conn, TestFails))
                End If
            Next
            'Auto-Lookup checks
            Console.WriteLine()
            Console.WriteLine("Executing Auto-Lookup checks")
            For Each TestFile As Generic.KeyValuePair(Of String, String) In TestFiles
                Dim CurrentTestFile As String = AssemblyTestEnvironment.TestFileAbsolutePath(TestFile.Value)
                Console.Write("Checking " & TestFile.Key & ": ")
                Dim FoundProviderLookupException As Exception = Nothing
                Dim conn As IDbConnection = Nothing
                Try
                    conn = CompuMaster.Data.DataQuery.Connections.MicrosoftAccessConnection(CurrentTestFile)
                Catch ex As Exception
                    FoundProviderLookupException = ex
                    TestFails = True
                End Try
                If FoundProviderLookupException IsNot Nothing Then
                    Console.WriteLine("FAILED ON PROVIDER LOOKUP: " & FoundProviderLookupException.Message)
                Else
                    Console.WriteLine(MicrosoftAccessOrExcelConnectionMatrixByProviderAndAccessOrExcelFileFormatVersion_TryOpenConnectionTest(conn, TestFails))
                End If
            Next
            If TestFails = True Then
                Assert.Fail("Some errors occured")
            Else
                Assert.Pass("All files tested successfully with OLEDB + ODBC")
            End If
        End Sub

        Private Function MicrosoftAccessOrExcelConnectionMatrixByProviderAndAccessOrExcelFileFormatVersion_TryOpenConnectionTest(conn As IDbConnection, ByRef TestFails As Boolean) As String
            Dim Result As String = Nothing
            If CType(conn, Object).GetType Is GetType(System.Data.OleDb.OleDbConnection) Then
                Result = "OLEDB"
            ElseIf CType(conn, Object).GetType Is GetType(System.Data.Odbc.OdbcConnection) Then
                Result = "ODBC"
            Else
                Result = "UNKNOWN DATA PROVIDER"
            End If

            Try
                CompuMaster.Data.DataQuery.OpenConnection(conn)
                Result &= " WORKING"
            Catch ex As Exception
                Result &= " FAILED ON OPENING: " & ex.Message
                TestFails = True
            Finally
                CompuMaster.Data.DataQuery.CloseAndDisposeConnection(conn)
            End Try
            Return Result
        End Function

        <Test()> Public Sub MicrosoftExcelOledbConnection()
            Console.WriteLine("Trying to find appropriate data provider for platform " & PlatformDependentProcessBitNumber())
            Dim TestFile As String = AssemblyTestEnvironment.TestFileAbsolutePath("testfiles\test_for_lastcell_e50aka95.xls")
            Dim conn As IDbConnection = CompuMaster.Data.DataQuery.Connections.MicrosoftExcelOleDbConnection(TestFile, False, True)
            Console.WriteLine("Evaluated data provider connection string for current platform: " & conn.ConnectionString)
            If CType(conn, Object).GetType Is GetType(System.Data.OleDb.OleDbConnection) Then
                Try
                    CompuMaster.Data.DataQuery.OpenConnection(conn)
                    Assert.AreNotEqual(0, CompuMaster.Data.DataQuery.Connections.EnumerateTablesAndViewsInOleDbDataSource(conn).Length)
                Finally
                    CompuMaster.Data.DataQuery.CloseAndDisposeConnection(conn)
                End Try
                Assert.Pass("Excel XLS opened at " & PlatformDependentProcessBitNumber() & " platform")
            Else
                Assert.Fail("Failed to open Excel XLS at " & PlatformDependentProcessBitNumber() & " platform")
            End If
        End Sub

        <Test()> Public Sub MicrosoftExcelConnection()
            Console.WriteLine("Trying to find appropriate data provider for platform " & PlatformDependentProcessBitNumber())
            Dim TestFile As String = AssemblyTestEnvironment.TestFileAbsolutePath("testfiles\test_for_lastcell_e50aka95.xls")
            Dim conn As IDbConnection = CompuMaster.Data.DataQuery.Connections.MicrosoftExcelConnection(TestFile, False, True)
            Console.WriteLine("Evaluated data provider connection string for current platform: " & conn.ConnectionString)
            If CType(conn, Object).GetType Is GetType(System.Data.OleDb.OleDbConnection) Then
                Try
                    CompuMaster.Data.DataQuery.OpenConnection(conn)
                    Assert.AreNotEqual(0, CompuMaster.Data.DataQuery.Connections.EnumerateTablesAndViewsInOleDbDataSource(conn).Length)
                Finally
                    CompuMaster.Data.DataQuery.CloseAndDisposeConnection(conn)
                End Try
                Assert.Pass("Excel XLS opened at " & PlatformDependentProcessBitNumber() & " platform")
                Console.WriteLine("Excel XLS opened at " & PlatformDependentProcessBitNumber() & " platform")
            Else
                Assert.Fail("Failed to open Excel XLS at " & PlatformDependentProcessBitNumber() & " platform")
            End If
        End Sub

        <Test()> Public Sub MicrosoftAccessOdbcConnection()
            Console.WriteLine("Trying to find appropriate data provider for platform " & PlatformDependentProcessBitNumber())
            Dim TestFile As String = AssemblyTestEnvironment.TestFileAbsolutePath("testfiles\test_for_msaccess.mdb")
            Dim conn As IDbConnection = CompuMaster.Data.DataQuery.Connections.MicrosoftAccessOdbcConnection(TestFile)
            Console.WriteLine("Evaluated data provider connection string for current platform: " & conn.ConnectionString)
            If CType(conn, Object).GetType Is GetType(System.Data.Odbc.OdbcConnection) Then
                Try
                    CompuMaster.Data.DataQuery.OpenConnection(conn)
                    Assert.AreNotEqual(0, CompuMaster.Data.DataQuery.Connections.EnumerateTablesAndViewsInOdbcDataSource(conn).Length)
                Finally
                    CompuMaster.Data.DataQuery.CloseAndDisposeConnection(conn)
                End Try
                Assert.Pass("Access MDB opened at " & PlatformDependentProcessBitNumber() & " platform")
            Else
                Assert.Fail("Failed to open Access MDB at " & PlatformDependentProcessBitNumber() & " platform")
            End If
        End Sub

        <Test()> Public Sub MicrosoftAccessOledbConnection()
            Console.WriteLine("Trying to find appropriate data provider for platform " & PlatformDependentProcessBitNumber())
            Dim TestFile As String = AssemblyTestEnvironment.TestFileAbsolutePath("testfiles\test_for_msaccess.mdb")
            Dim conn As IDbConnection = CompuMaster.Data.DataQuery.Connections.MicrosoftAccessOleDbConnection(TestFile)
            Console.WriteLine("Evaluated data provider connection string for current platform: " & conn.ConnectionString)
            If CType(conn, Object).GetType Is GetType(System.Data.OleDb.OleDbConnection) Then
                Try
                    CompuMaster.Data.DataQuery.OpenConnection(conn)
                    Assert.AreNotEqual(0, CompuMaster.Data.DataQuery.Connections.EnumerateTablesAndViewsInOleDbDataSource(conn).Length)
                Finally
                    CompuMaster.Data.DataQuery.CloseAndDisposeConnection(conn)
                End Try
                Assert.Pass("Access MDB opened at " & PlatformDependentProcessBitNumber() & " platform")
            Else
                Assert.Fail("Failed to open Access MDB at " & PlatformDependentProcessBitNumber() & " platform")
            End If
        End Sub

        <Test()> Public Sub MicrosoftAccessConnection_MediumTrust()
            'Permission required to read the providers application name And access config
            Dim permissions As New System.Security.PermissionSet(System.Security.Permissions.PermissionState.None)
            permissions.AddPermission(New System.Web.AspNetHostingPermission(System.Web.AspNetHostingPermissionLevel.Minimal))
            permissions.AddPermission(New System.Security.Permissions.FileIOPermission(System.Security.Permissions.PermissionState.Unrestricted))
            permissions.Assert()

            Console.WriteLine("Current trust level for code security: " & GetCurrentTrustLevel.ToString)
            Console.WriteLine("Current trust level for app domain security: IsUnrestricted=" & AppDomain.CurrentDomain.ApplicationTrust.DefaultGrantSet.PermissionSet.IsUnrestricted())
            If (System.Web.AspNetHostingPermissionLevel.Medium <> GetCurrentTrustLevel()) Then
                Assert.Ignore("Code access security trust level must be set to medium trust for this test")
            End If
            MicrosoftAccessConnection()

            System.Security.PermissionSet.RevertAssert()

        End Sub

        Private Function GetCurrentTrustLevel() As System.Web.AspNetHostingPermissionLevel
            Dim CheckTrustLevels As System.Web.AspNetHostingPermissionLevel()
            CheckTrustLevels = New System.Web.AspNetHostingPermissionLevel() {
                    System.Web.AspNetHostingPermissionLevel.Unrestricted,
                    System.Web.AspNetHostingPermissionLevel.High,
                    System.Web.AspNetHostingPermissionLevel.Medium,
                    System.Web.AspNetHostingPermissionLevel.Low,
                    System.Web.AspNetHostingPermissionLevel.Minimal
                }
            For Each trustLevel As System.Web.AspNetHostingPermissionLevel In CheckTrustLevels
                Try
                    Dim TestPermissionLevel As System.Web.AspNetHostingPermission = New System.Web.AspNetHostingPermission(trustLevel)
                    TestPermissionLevel.Demand()
                Catch ex As System.Security.SecurityException
                    Continue For
                End Try
                Return trustLevel
            Next

            Return System.Web.AspNetHostingPermissionLevel.None
        End Function

        <Test()> Public Sub MicrosoftAccessConnection()
            Console.WriteLine("Trying to find appropriate data provider for platform " & PlatformDependentProcessBitNumber())
            Dim TestFile As String = AssemblyTestEnvironment.TestFileAbsolutePath("testfiles\test_for_msaccess.mdb")
            Dim conn As IDbConnection = CompuMaster.Data.DataQuery.Connections.MicrosoftAccessConnection(TestFile)
            Console.WriteLine("Evaluated data provider connection string for current platform: " & conn.ConnectionString)
            If CType(conn, Object).GetType Is GetType(System.Data.OleDb.OleDbConnection) Then
                Try
                    CompuMaster.Data.DataQuery.OpenConnection(conn)
                    Assert.AreNotEqual(0, CompuMaster.Data.DataQuery.Connections.EnumerateTablesAndViewsInOleDbDataSource(conn).Length)
                Finally
                    CompuMaster.Data.DataQuery.CloseAndDisposeConnection(conn)
                End Try
                Assert.Pass("Access MDB opened at " & PlatformDependentProcessBitNumber() & " platform")
            Else
                Assert.Fail("Failed to open Access MDB at " & PlatformDependentProcessBitNumber() & " platform")
            End If
        End Sub
#End If

        Private Function PlatformDependentProcessBitNumber() As String
            If Environment.Is64BitProcess Then
                Return "x64"
            Else
                Return "x32"
            End If
        End Function

#If Not CI_Build Then
        <Test()> Public Sub ReadMsAccessDatabaseEnumeratedTable()
            Dim TestFile As String = AssemblyTestEnvironment.TestFileAbsolutePath("testfiles\test_for_msaccess.mdb")
            Dim MyConn As IDbConnection = CompuMaster.Data.DataQuery.Connections.MicrosoftAccessConnection(TestFile)
            Dim table As DataTable
            Try
                MyConn.Open()
                Dim MyCmd As IDbCommand = MyConn.CreateCommand()
                Dim tableIdentifier As String = CompuMaster.Data.DataQuery.Connections.EnumerateTablesAndViewsInOleDbDataSource(MyConn)(0).ToString
                MyCmd.CommandType = CommandType.Text
                MyCmd.CommandText = "SELECT * FROM " & tableIdentifier
                table = CompuMaster.Data.DataQuery.FillDataTable(MyCmd, CompuMaster.Data.DataQuery.Automations.AutoOpenAndCloseAndDisposeConnection, tableIdentifier)
                Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTable(table))
            Finally
                CompuMaster.Data.DataQuery.CloseAndDisposeConnection(MyConn)
            End Try
            Assert.AreNotEqual(0, table.Columns.Count, "Column count for random, enumerated table")
        End Sub
#End If

    End Class

End Namespace