Option Explicit On 
Option Strict On

Namespace CompuMaster.Data.DataQuery

    ''' <summary>
    ''' Identify the runtime platform
    ''' </summary>
    ''' <remarks></remarks>
    Public Class PlatformTools

        Public Enum ClrRuntimePlatform As Short
            x32 = 32
            x64 = 64
        End Enum

        Public Enum TargetPlatform As Short
            Current = 0
            x64 = 64
            x32 = 32
        End Enum

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
            Dim MyConn As System.Data.IDbConnection = Nothing
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
            If  platform = TargetPlatform.x32 AndAlso CurrentClrRuntime() = ClrRuntimePlatform.x64 Then
                Return Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SOFTWARE\\Wow6432Node\\ODBC\\ODBCINST.INI\ODBC Drivers").GetValueNames()
            ElseIf platform = TargetPlatform.x64 AndAlso CurrentClrRuntime() = ClrRuntimePlatform.x32 Then
                Throw New NotSupportedException("64 bit data not available on 32 bit platform")
            Else
                Return Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SOFTWARE\\ODBC\\ODBCINST.INI\ODBC Drivers").GetValueNames()
            End If
        End Function

        ''' <summary>
        ''' Enumerates the OLE DB providers currently installed on the running machine
        ''' </summary>
        ''' <returns>The DictionaryEntry contains the name of the OLE DB provider in the key field, the value field contains the provider description</returns>
        ''' <remarks>
        ''' <para>DELAY WARNING: the enumeration by registry keys will take approx. 1,000 ms (!)</para>
        ''' <para>CONTENT WARNING: the enumeration will return ALL registered providers, but you may not be able to use them because of 32bit vs. 64bit loading problems</para>
        ''' </remarks>
        Public Shared Function InstalledOleDbProviders() As DictionaryEntry()
            Dim Providers As New ArrayList

            ' I am only interested in the CLSID subtree
            'Dim keyCLSID As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey("CLSID", False)
            Dim keyCLSID As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey("CLSID", Microsoft.Win32.RegistryKeyPermissionCheck.ReadSubTree, Security.AccessControl.RegistryRights.ReadKey)
            Dim keys() As String = keyCLSID.GetSubKeyNames()
            Dim de As DictionaryEntry
            Dim i As Int32

            Dim AccessErrors As New Generic.List(Of Exception)
            ' Search through the tree just one level
            For i = 0 To keys.Length - 1
                Dim key As Microsoft.Win32.RegistryKey = Nothing
                Try
                    'Dim key As Microsoft.Win32.RegistryKey = keyCLSID.OpenSubKey(keys(i), False)
                    key = keyCLSID.OpenSubKey(keys(i), Microsoft.Win32.RegistryKeyPermissionCheck.ReadSubTree, Security.AccessControl.RegistryRights.ReadKey)

                    ' Search for OLE DB Providers
                    de = SearchKeys(key)
                    If Not (de.Key Is Nothing) Then
                        ' Found one, add it to the Dictionary
                        Providers.Add(de)
                    End If
                Catch ex As Exception
                    AccessErrors.Add(New Exception("ERROR at " & keyCLSID.ToString & "\" & keys(i)))
                Finally
                    If Not key Is Nothing Then key.Close()
                End Try
            Next
            If AccessErrors.Count > keys.Length / 40 Then '1 access error is usual at Win10 - error situation is with more than 40% errors on all existing sub-keys
                Throw New Exception("AccessErrors=" & AccessErrors.Count)
            End If
            Return CType(Providers.ToArray(GetType(DictionaryEntry)), DictionaryEntry())
        End Function

        ''' <summary>
        ''' Search OLE DB provider registry keys
        ''' </summary>
        ''' <param name="key"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Shared Function SearchKeys(ByVal key As Microsoft.Win32.RegistryKey) As DictionaryEntry
            Dim de As DictionaryEntry

            Try
                'Tries to find the "OLE DB Provider" key
                'Dim key2 As Microsoft.Win32.RegistryKey = key.OpenSubKey("OLE DB Provider", False)
                Dim key2 As Microsoft.Win32.RegistryKey = key.OpenSubKey("OLE DB Provider", Microsoft.Win32.RegistryKeyPermissionCheck.ReadSubTree, Security.AccessControl.RegistryRights.ReadKey)

                If Not (key2 Is Nothing) Then
                    ' Found it, fills the DictionaryEntry
                    de = New DictionaryEntry()
                    Dim sValues() As String = key2.GetValueNames()
                    de.Key = key.OpenSubKey("ProgID", False).GetValue(sValues(0))
                    Dim sValues2() As String = key2.GetValueNames()
                    de.Value = key2.GetValue(sValues2(0))
                    key2.Close()
                End If
            Catch ex As Exception
                Throw New Exception("ERROR at " & key.ToString & "\OLE DB Provider")
            End Try
            Return de
        End Function

    End Class

End Namespace
