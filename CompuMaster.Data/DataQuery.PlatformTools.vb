Option Explicit On 
Option Strict On

Namespace CompuMaster.Data.DataQuery

    ''' <summary>
    ''' Identify the runtime platform
    ''' </summary>
    ''' <remarks></remarks>
    Public NotInheritable Class PlatformTools

#Disable Warning CA1027 ' Mark enums with FlagsAttribute
        Public Enum ClrRuntimePlatform As Short
            x32 = 32
            x64 = 64
        End Enum

        Public Enum TargetPlatform As Short
            Current = 0
            x64 = 64
            x32 = 32
        End Enum
#Enable Warning CA1027 ' Mark enums with FlagsAttribute

        ''' <summary>
        ''' Indicates wether the current application runs in 32bit mode or in 64bit mode (relevant e.g. for the ODBC drivers to load)
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function CurrentClrRuntime() As ClrRuntimePlatform
            If IntPtr.Size = 4 Then
                '32bit CLR
                Return ClrRuntimePlatform.x32
            ElseIf IntPtr.Size = 8 Then
                '64bit CLR
                Return ClrRuntimePlatform.x64
            End If
        End Function

        ''' <summary>
        ''' Create a provider-independent IDbConnection based on the given textual information
        ''' </summary>
        ''' <param name="provider">SqlClient, ODBC or OleDB</param>
        ''' <param name="connectionString"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function CreateDataConnection(ByVal provider As String, ByVal connectionString As String) As IDbConnection
            Dim MyConn As System.Data.IDbConnection
            If provider = "SqlClient" Then
                MyConn = New System.Data.SqlClient.SqlConnection(connectionString)
            ElseIf provider = "ODBC" Then
                MyConn = New System.Data.Odbc.OdbcConnection(connectionString)
            ElseIf provider = "OleDB" Then
                MyConn = New System.Data.OleDb.OleDbConnection(connectionString)
            Else
                Throw New Exception("Invalid data provider")
            End If
            Return MyConn
        End Function

        ''' <summary>
        ''' Enumerates the ODBC drivers currently installed on the running machine
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function InstalledOdbcDrivers(platform As TargetPlatform) As String()
            If System.Environment.OSVersion.Platform = PlatformID.Win32NT Then
                If platform = TargetPlatform.x32 AndAlso CurrentClrRuntime() = ClrRuntimePlatform.x64 Then
                    Return Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SOFTWARE\\Wow6432Node\\ODBC\\ODBCINST.INI\ODBC Drivers").GetValueNames()
                ElseIf platform = TargetPlatform.x64 AndAlso CurrentClrRuntime() = ClrRuntimePlatform.x32 Then
                    Throw New NotSupportedException("64 bit data not available on 32 bit platform")
                Else
                    Return Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SOFTWARE\\ODBC\\ODBCINST.INI\ODBC Drivers").GetValueNames()
                End If
            Else
                Throw New NotImplementedException("Support for this method at the current platform not yet implemented")
            End If
        End Function

        ''' <summary>
        ''' Write some additional details on installed OleDb providers to console output
        ''' </summary>
        Friend Shared Sub ConsolOutputListOfInstalledOleDbProviders()
            Dim InstalledProviders As Generic.List(Of String) = InstalledOleDbProvidersList()
            Console.WriteLine("Installed OleDB providers: " & InstalledProviders.Count)
            If InstalledProviders.Count > 0 Then
                Console.WriteLine("- " & String.Join(System.Environment.NewLine & "- ", InstalledProviders.ToArray))
            End If
        End Sub

        ''' <summary>
        ''' List of names of all installed OleDbProviders
        ''' </summary>
        ''' <returns></returns>
        Public Shared Function InstalledOleDbProvidersList() As Generic.List(Of String)
            Try
                Dim ProviderReader As System.Data.OleDb.OleDbDataReader = System.Data.OleDb.OleDbEnumerator.GetRootEnumerator
                Dim ProviderTable As DataTable = CompuMaster.Data.DataTables.ConvertDataReaderToDataTable(ProviderReader)
                Return CompuMaster.Data.DataTables.ConvertColumnValuesIntoList(Of String)(ProviderTable.Columns("SOURCES_NAME"))
            Catch ex As NotImplementedException
                Return New Generic.List(Of String)
            End Try
        End Function

        ''' <summary>
        ''' List of names and descriptions of all installed OleDbProviders
        ''' </summary>
        ''' <returns></returns>
        Public Shared Function InstalledOleDbProvidersWithDescription() As Generic.List(Of Generic.KeyValuePair(Of String, String))
            Try
                Dim ProviderReader As System.Data.OleDb.OleDbDataReader = System.Data.OleDb.OleDbEnumerator.GetRootEnumerator
                Dim ProviderTable As DataTable = CompuMaster.Data.DataTables.ConvertDataReaderToDataTable(ProviderReader)
                Return CompuMaster.Data.DataTables.ConvertColumnValuesIntoList(Of String, String)(ProviderTable.Columns("SOURCES_NAME"), ProviderTable.Columns("SOURCES_DESCRIPTION"))
            Catch ex As NotImplementedException
                Return New Generic.List(Of Generic.KeyValuePair(Of String, String))
            End Try
        End Function

        ''' <summary>
        ''' Test for the existance of a provider with the specified beginning of name
        ''' </summary>
        ''' <param name="nameMustStartWith"></param>
        ''' <returns></returns>
        Public Shared Function ProbeOleDbProvider(nameMustStartWith As String) As Boolean
            Dim Providers As Generic.List(Of String) = InstalledOleDbProvidersList()
            For MyCounter As Integer = 0 To Providers.Count - 1
                If Providers(MyCounter).StartsWith(nameMustStartWith) Then Return True
            Next
            Return False
        End Function

        ''' <summary>
        ''' Enumerates the OLE DB providers currently installed on the running machine for the current process (64bit vs. 32bit)
        ''' </summary>
        ''' <returns>The DictionaryEntry contains the name of the OLE DB provider in the key field, the value field contains the provider description</returns>
        Public Shared Function InstalledOleDbProviders() As DictionaryEntry()
            Try
                Dim ProviderReader As System.Data.OleDb.OleDbDataReader = System.Data.OleDb.OleDbEnumerator.GetRootEnumerator
                Dim ProviderTable As DataTable = CompuMaster.Data.DataTables.ConvertDataReaderToDataTable(ProviderReader)
                Return CompuMaster.Data.DataTables.ConvertDataTableToDictionaryEntryArray(ProviderTable.Columns("SOURCES_NAME"), ProviderTable.Columns("SOURCES_DESCRIPTION"))
            Catch ex As NotImplementedException
                Return New DictionaryEntry() {}
            End Try
        End Function

        ''' <summary>
        ''' Find the newest available provider name starting with "Microsoft.ACE.OLEDB."
        ''' </summary>
        ''' <returns></returns>
        Public Shared Function FindLatestMsOfficeAceOleDbProviderName() As String
            Dim MatchingProviders As Generic.List(Of String)
            MatchingProviders = DataQuery.PlatformTools.InstalledOleDbProvidersList.FindAll(
                Function(value As String) As Boolean
                    Return value.StartsWith("Microsoft.ACE.OLEDB.")
                End Function)
            If MatchingProviders.Count = 0 Then
                Return Nothing
            Else
                MatchingProviders.Sort()
                Return MatchingProviders(MatchingProviders.Count - 1)
            End If
        End Function

        ''' <summary>
        ''' Find the newest available provider name starting with "Microsoft.Jet.OLEDB."
        ''' </summary>
        ''' <returns></returns>
        Public Shared Function FindLatestMsOfficeJetOleDbProviderName() As String
            Dim MatchingProviders As Generic.List(Of String)
            MatchingProviders = DataQuery.PlatformTools.InstalledOleDbProvidersList.FindAll(
                Function(value As String) As Boolean
                    Return value.StartsWith("Microsoft.Jet.OLEDB.")
                End Function)
            If MatchingProviders.Count = 0 Then
                Return Nothing
            Else
                MatchingProviders.Sort()
                Return MatchingProviders(MatchingProviders.Count - 1)
            End If
        End Function

    End Class

End Namespace
