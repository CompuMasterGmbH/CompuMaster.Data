Option Explicit On
Option Strict On

Imports System.Data
Imports CompuMaster.Data.Information
'Imports Microsoft.VisualBasic

Namespace CompuMaster.Data.DataQuery

    ''' <summary>
    ''' A factory for common data connection types, usable on most platforms
    ''' </summary>
    ''' <remarks></remarks>
    Public NotInheritable Class Connections

        ''' <summary>
        ''' A most probably working Microsoft Access connection which uses the most appropriate, installed OleDB provider of the current machine
        ''' </summary>
        ''' <param name="path">The path of the file</param>
        ''' <returns>An OleDB data connection to the requested file</returns>
        ''' <remarks></remarks>
        ''' <exception cref="Office2010x64OleDbOdbcEngineRequiredException" />
        Public Shared Function MicrosoftAccessOleDbConnection(ByVal path As String) As IDbConnection
            If path = Nothing Then Throw New ArgumentNullException(NameOf(path))
            Return MicrosoftAccessOleDbConnection(path, "")
        End Function

        ''' <summary>
        ''' A most probably working Microsoft Access connection which uses the most appropriate, installed OleDB provider of the current machine
        ''' </summary>
        ''' <param name="path">The path of the file</param>
        ''' <param name="databasePassword"></param>
        ''' <returns>An OleDB data connection to the requested file</returns>
        ''' <remarks></remarks>
        ''' <exception cref="Office2010x64OleDbOdbcEngineRequiredException" />
        Public Shared Function MicrosoftAccessOleDbConnection(ByVal path As String, ByVal databasePassword As String) As IDbConnection
            If path = Nothing Then Throw New ArgumentNullException(NameOf(path))
            'Lookup OleDb provider (fast!)
            Dim FoundProvider As String
            FoundProvider = PlatformTools.FindLatestMsOfficeAceOleDbProviderName
            If FoundProvider <> Nothing Then
                Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("OleDB", "Provider=" & FoundProvider & ";Data Source=" & path & ";User Id=admin;Password=" & databasePassword & ";")
            End If
            If path.ToLowerInvariant.EndsWith(".mdb") OrElse path.ToLowerInvariant.EndsWith(".mde") Then
                'Try to lookup MS Jet provider which is still fine for this file type
                FoundProvider = PlatformTools.FindLatestMsOfficeJetOleDbProviderName
                If FoundProvider <> Nothing Then
                    Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("OleDB", "Provider=" & FoundProvider & ";Data Source=" & path & ";Jet OLEDB:Database Password=" & databasePassword & ";")
                End If
            End If
            'Probe OleDb provider
            Dim TestFile As TestFile
            Dim MsJetProviderIsSufficient As Boolean
            Try
                'Try to create a temporary file - might fail in environments which are not fully trusted
                If path.ToLowerInvariant.EndsWith(".mdb") OrElse path.ToLowerInvariant.EndsWith(".mde") Then
                    TestFile = New TestFile(DataQuery.TestFile.TestFileType.MsAccessMdb)
                    MsJetProviderIsSufficient = True
                Else
                    TestFile = New TestFile(DataQuery.TestFile.TestFileType.MsAccessAccdb)
                    MsJetProviderIsSufficient = False
                End If
            Catch ex As Exception
                'Non-full-trust or not enough priviledges/rights to create the temporary file
                'Use a safe way to get a working data connection to the MS Access database
                If CompuMaster.Data.DataQuery.PlatformTools.CurrentClrRuntime = CompuMaster.Data.DataQuery.PlatformTools.ClrRuntimePlatform.x64 Then
                    '64bit - Requires Office 2010 JET drivers, but are called "Microsoft.ACE.OLEDB.12.0", too
                    Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("OleDB", "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & path & ";User Id=admin;Password=" & databasePassword & ";")
                ElseIf path.ToLowerInvariant.EndsWith(".accdb") OrElse path.ToLowerInvariant.EndsWith(".accde") Then
                    '32bit - Requires Office 2007 JET drivers
                    Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("OleDB", "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & path & ";User Id=admin;Password=" & databasePassword & ";")
                Else
                    '32bit - Requires basic JET drivers
                    Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("OleDB", "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & path & ";Jet OLEDB:Database Password=" & databasePassword & ";")
                End If
            End Try
            Try
                For MyCounter As Integer = (System.DateTime.Now.Year + 1 - 2000) To 15 Step -1 'try all MS Office releases since 2015 up to current year + 1
                    If ProbeOleDBProvider(MicrosoftAccessConnectionProviderWorkingStatusForACEDynList(MyCounter), "Provider=Microsoft.ACE.OLEDB." & MyCounter & ".0;Data Source=" & TestFile.FilePath & ";User Id=admin;Password=" & databasePassword & ";") Then
                        Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("OleDB", "Provider=Microsoft.ACE.OLEDB." & MyCounter & ".0;Data Source=" & path & ";User Id=admin;Password=" & databasePassword & ";")
                    End If
                Next
                If ProbeOleDBProvider(MicrosoftAccessConnectionProviderWorkingStatusForACE14, "Provider=Microsoft.ACE.OLEDB.14.0;Data Source=" & TestFile.FilePath & ";User Id=admin;Password=" & databasePassword & ";") Then
                    Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("OleDB", "Provider=Microsoft.ACE.OLEDB.14.0;Data Source=" & path & ";Jet OLEDB:Database Password=" & databasePassword & ";")
                ElseIf ProbeOleDBProvider(MicrosoftAccessConnectionProviderWorkingStatusForACE12, "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & TestFile.FilePath & ";User Id=admin;Password=" & databasePassword & ";") Then
                    Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("OleDB", "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & path & ";Jet OLEDB:Database Password=" & databasePassword & ";")
                ElseIf MsJetProviderIsSufficient AndAlso ProbeOleDBProvider(MicrosoftAccessConnectionProviderWorkingStatusForJet4, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & TestFile.FilePath & ";User Id=admin;Password=" & databasePassword & ";") Then
                    Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("OleDB", "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & path & ";Jet OLEDB:Database Password=" & databasePassword & ";")
                Else
                    'Let the application find the exception with the most modern provider
                    If CompuMaster.Data.DataQuery.PlatformTools.CurrentClrRuntime = CompuMaster.Data.DataQuery.PlatformTools.ClrRuntimePlatform.x64 Then
                        '64bit - Requires Office 2010 JET drivers
                        Throw New Office2010x64OleDbOdbcEngineRequiredException()
                    ElseIf path.ToLower.EndsWith(".accdb") Then
                        '32bit - Requires Office 2007 JET drivers
                        Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("OleDB", "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & path & ";Jet OLEDB:Database Password=" & databasePassword & ";")
                    Else
                        '32bit - Requires basic JET drivers
                        Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("OleDB", "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & path & ";Jet OLEDB:Database Password=" & databasePassword & ";")
                    End If
                End If
            Finally
                If TestFile IsNot Nothing Then
                    TestFile.Dispose()
                End If
            End Try
        End Function

        ''' <summary>
        ''' A most probably working Microsoft Access connection which uses the most appropriate, installed ODBC driver of the current machine
        ''' </summary>
        ''' <param name="path">The path of the file</param>
        ''' <returns>An ODBC data connection to the requested file</returns>
        ''' <remarks></remarks>
        ''' <exception cref="Office2010x64OleDbOdbcEngineRequiredException" />
        Public Shared Function MicrosoftAccessOdbcConnection(ByVal path As String) As IDbConnection
            If path = Nothing Then Throw New ArgumentNullException(NameOf(path))
            Return MicrosoftAccessOdbcConnection(path, "")
        End Function

        ''' <summary>
        ''' A most probably working Microsoft Access connection which uses the most appropriate, installed ODBC driver of the current machine
        ''' </summary>
        ''' <param name="path">The path of the file</param>
        ''' <param name="databasePassword"></param>
        ''' <returns>An ODBC data connection to the requested file</returns>
        ''' <remarks></remarks>
        ''' <exception cref="Office2010x64OleDbOdbcEngineRequiredException" />
        Public Shared Function MicrosoftAccessOdbcConnection(ByVal path As String, ByVal databasePassword As String) As IDbConnection
            If path = Nothing Then Throw New ArgumentNullException(NameOf(path))
            Dim TestFile As TestFile
            Dim MsJetProviderIsSufficient As Boolean
            Try
                'Try to create a temporary file - might fail in environments which are not fully trusted
                If path.ToLowerInvariant.EndsWith(".mdb") OrElse path.ToLowerInvariant.EndsWith(".mde") Then
                    TestFile = New TestFile(DataQuery.TestFile.TestFileType.MsAccessMdb)
                    MsJetProviderIsSufficient = True
                Else
                    TestFile = New TestFile(DataQuery.TestFile.TestFileType.MsAccessAccdb)
                    MsJetProviderIsSufficient = False
                End If
            Catch ex As Exception
                'Non-full-trust or not enough priviledges/rights to create the temporary file
                'Use a safe way to get a working data connection to the MS Access database
                If CompuMaster.Data.DataQuery.PlatformTools.CurrentClrRuntime = CompuMaster.Data.DataQuery.PlatformTools.ClrRuntimePlatform.x64 Then
                    '64bit 
                    Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("ODBC", "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" & path & ";Uid=Admin;Pwd=;")
                ElseIf path.ToLowerInvariant.EndsWith(".accdb") OrElse path.ToLowerInvariant.EndsWith(".accde") Then
                    '32bit but .accdb
                    Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("ODBC", "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" & path & ";Uid=Admin;Pwd=;")
                Else
                    '32bit .mdb
                    Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("ODBC", "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & path & ";Uid=Admin;Pwd=;")
                End If
            End Try
            Try
                If CompuMaster.Data.DataQuery.PlatformTools.CurrentClrRuntime = CompuMaster.Data.DataQuery.PlatformTools.ClrRuntimePlatform.x64 AndAlso ProbeOdbcDBProvider(MicrosoftAccessConnectionProviderWorkingStatusForOdbcDriver, "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" & TestFile.FilePath & ";Uid=Admin;Pwd=" & databasePassword & ";") Then
                    Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("ODBC", "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" & path & ";Uid=Admin;Pwd=" & databasePassword & ";")
                ElseIf MsJetProviderIsSufficient AndAlso CompuMaster.Data.DataQuery.PlatformTools.CurrentClrRuntime = CompuMaster.Data.DataQuery.PlatformTools.ClrRuntimePlatform.x32 AndAlso ProbeOdbcDBProvider(MicrosoftAccessConnectionProviderWorkingStatusForOdbcDriver, "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & TestFile.FilePath & ";Uid=Admin;Pwd=" & databasePassword & ";") Then
                    Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("ODBC", "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & path & ";Uid=Admin;Pwd=" & databasePassword & ";")
                Else
                    'Let the application find the exception with the most modern provider
                    If CompuMaster.Data.DataQuery.PlatformTools.CurrentClrRuntime = CompuMaster.Data.DataQuery.PlatformTools.ClrRuntimePlatform.x64 Then
                        '64bit 
                        Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("ODBC", "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" & path & ";Uid=Admin;Pwd=" & databasePassword & ";")
                    ElseIf path.ToLowerInvariant.EndsWith(".accdb") OrElse path.ToLowerInvariant.EndsWith(".accde") Then
                        '32bit but .accdb
                        Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("ODBC", "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" & path & ";Uid=Admin;Pwd=" & databasePassword & ";")
                    Else
                        '32bit .mdb
                        Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("ODBC", "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & path & ";Uid=Admin;Pwd=" & databasePassword & ";")
                    End If
                End If
            Finally
                If TestFile IsNot Nothing Then
                    TestFile.Dispose()
                End If
            End Try
        End Function

        ''' <summary>
        ''' A most probably working Microsoft Access connection which uses the most appropriate, installed ODBC driver of the current machine
        ''' </summary>
        ''' <param name="path">The path of the directory with multiple text files</param>
        ''' <returns>An ODBC data connection to the requested file</returns>
        ''' <remarks></remarks>
        ''' <exception cref="Office2010x64OleDbOdbcEngineRequiredException" />
        Public Shared Function TextCsvConnection(ByVal path As String) As IDbConnection
            If path = Nothing Then Throw New ArgumentNullException(NameOf(path))
            If System.IO.Directory.Exists(path) = False Then Throw New ArgumentException("Path must be a an existing directory", NameOf(path))
            Dim TestFile As TestFile
            Try
                'Try to create a temporary file - might fail in environments which are not fully trusted
                TestFile = New TestFile(DataQuery.TestFile.TestFileType.TextCsv)
            Catch ex As Exception
                'Non-full-trust or not enough priviledges/rights to create the temporary file
                'Use a safe way to get a working data connection to the MS Access database
                If CompuMaster.Data.DataQuery.PlatformTools.CurrentClrRuntime = CompuMaster.Data.DataQuery.PlatformTools.ClrRuntimePlatform.x64 Then
                    '64bit 
                    Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("ODBC", "Driver={Microsoft Access Text Driver (*.txt, *.csv)};Dbq=" & path & ";Extensions=asc,csv,tab,txt;")
                Else
                    '32bit
                    Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("ODBC", "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" & path & ";Extensions=asc,csv,tab,txt;")
                End If
            End Try
            Try
                Dim CsvDataDir As String = System.IO.Path.GetDirectoryName(TestFile.FilePath)
                If ProbeOdbcDBProvider(MicrosoftAccessTextCsvConnectionProviderWorkingStatusForOdbcDriver, "Driver={Microsoft Access Text Driver (*.txt, *.csv)};Dbq=" & CsvDataDir & ";Extensions=asc,csv,tab,txt;") Then
                    Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("ODBC", "Driver={Microsoft Access Text Driver (*.txt, *.csv)};Dbq=" & path & ";Extensions=asc,csv,tab,txt;")
                ElseIf ProbeOdbcDBProvider(MicrosoftTextCsvConnectionProviderWorkingStatusForOdbcDriver, "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" & CsvDataDir & ";Extensions=asc,csv,tab,txt;") Then
                    Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("ODBC", "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" & path & ";Extensions=asc,csv,tab,txt;")
                Else
                    Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("ODBC", "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" & path & ";Extensions=asc,csv,tab,txt;")
                End If
            Finally
                If TestFile IsNot Nothing Then
                    TestFile.Dispose()
                End If
            End Try
        End Function

        ''' <summary>
        ''' A most probably working Microsoft Access connection which uses the most appropriate, installed OleDB provider or ODBC driver of the current machine
        ''' </summary>
        ''' <param name="path">The path of the file</param>
        ''' <returns>An OleDB or ODBC data connection to the requested file</returns>
        ''' <remarks></remarks>
        ''' <exception cref="Office2010x64OleDbOdbcEngineRequiredException" />
        Public Shared Function MicrosoftAccessConnection(ByVal path As String) As IDbConnection
            If path = Nothing Then Throw New ArgumentNullException(NameOf(path))
            Return MicrosoftAccessConnection(path, "")
        End Function

        ''' <summary>
        ''' A most probably working Microsoft Access connection which uses the most appropriate, installed OleDB provider or ODBC driver of the current machine
        ''' </summary>
        ''' <param name="path">The path of the file</param>
        ''' <param name="databasePassword"></param>
        ''' <returns>An OleDB or ODBC data connection to the requested file</returns>
        ''' <remarks></remarks>
        ''' <exception cref="Office2010x64OleDbOdbcEngineRequiredException" />
        Public Shared Function MicrosoftAccessConnection(ByVal path As String, ByVal databasePassword As String) As IDbConnection
            If path = Nothing Then Throw New ArgumentNullException(NameOf(path))
            'Lookup OleDb provider (fast!)
            Dim FoundProvider As String
            FoundProvider = PlatformTools.FindLatestMsOfficeAceOleDbProviderName
            If FoundProvider <> Nothing Then
                Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("OleDB", "Provider=" & FoundProvider & ";Data Source=" & path & ";User Id=admin;Password=" & databasePassword & ";")
            End If
            If path.ToLowerInvariant.EndsWith(".mdb") OrElse path.ToLowerInvariant.EndsWith(".mde") Then
                'Try to lookup MS Jet provider which is still fine for this file type
                FoundProvider = PlatformTools.FindLatestMsOfficeJetOleDbProviderName
                If FoundProvider <> Nothing Then
                    Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("OleDB", "Provider=" & FoundProvider & ";Data Source=" & path & ";Jet OLEDB:Database Password=" & databasePassword & ";")
                End If
            End If
            'Probe OleDb provider
            Dim TestFile As TestFile
            Dim MsJetProviderIsSufficient As Boolean
            Try
                'Try to create a temporary file - might fail in environments which are not fully trusted
                If path.ToLowerInvariant.EndsWith(".mdb") OrElse path.ToLowerInvariant.EndsWith(".mde") Then
                    TestFile = New TestFile(DataQuery.TestFile.TestFileType.MsAccessMdb)
                    MsJetProviderIsSufficient = True
                Else
                    TestFile = New TestFile(DataQuery.TestFile.TestFileType.MsAccessAccdb)
                    MsJetProviderIsSufficient = False
                End If
            Catch ex As Exception
                'Non-full-trust or not enough priviledges/rights to create the temporary file
                'Use a safe way to get a working data connection to the MS Access database
                If CompuMaster.Data.DataQuery.PlatformTools.CurrentClrRuntime = CompuMaster.Data.DataQuery.PlatformTools.ClrRuntimePlatform.x64 Then
                    '64bit - Requires Office 2010 JET drivers, but are called "Microsoft.ACE.OLEDB.12.0", too
                    Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("OleDB", "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & path & ";User Id=admin;Password=;")
                ElseIf path.ToLowerInvariant.EndsWith(".accdb") OrElse path.ToLowerInvariant.EndsWith(".accde") Then
                    '32bit - Requires Office 2007 JET drivers
                    Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("OleDB", "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & path & ";User Id=admin;Password=;")
                Else
                    '32bit - Requires basic JET drivers
                    Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("ODBC", "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & path & ";Uid=Admin;Pwd=;")
                End If
            End Try
            Try
                For MyCounter As Integer = (System.DateTime.Now.Year + 1 - 2000) To 15 Step -1 'try all MS Office releases since 2015 up to current year + 1
                    If ProbeOleDBProvider(MicrosoftAccessConnectionProviderWorkingStatusForACEDynList(MyCounter), "Provider=Microsoft.ACE.OLEDB." & MyCounter & ".0;Data Source=" & TestFile.FilePath & ";User Id=admin;Password=" & databasePassword & ";") Then
                        Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("OleDB", "Provider=Microsoft.ACE.OLEDB." & MyCounter & ".0;Data Source=" & path & ";User Id=admin;Password=" & databasePassword & ";")
                    End If
                Next
                If ProbeOleDBProvider(MicrosoftAccessConnectionProviderWorkingStatusForACE14, "Provider=Microsoft.ACE.OLEDB.14.0;Data Source=" & TestFile.FilePath & ";User Id=admin;Password=" & databasePassword & ";") Then
                    Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("OleDB", "Provider=Microsoft.ACE.OLEDB.14.0;Data Source=" & path & ";Jet OLEDB:Database Password=" & databasePassword & ";")
                ElseIf ProbeOleDBProvider(MicrosoftAccessConnectionProviderWorkingStatusForACE12, "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & TestFile.FilePath & ";User Id=admin;Password=" & databasePassword & ";") Then
                    Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("OleDB", "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & path & ";Jet OLEDB:Database Password=" & databasePassword & ";")
                ElseIf MsJetProviderIsSufficient AndAlso ProbeOleDBProvider(MicrosoftAccessConnectionProviderWorkingStatusForJet4, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & TestFile.FilePath & ";User Id=admin;Password=" & databasePassword & ";") Then
                    Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("OleDB", "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & path & ";Jet OLEDB:Database Password=" & databasePassword & ";")
                ElseIf CompuMaster.Data.DataQuery.PlatformTools.CurrentClrRuntime = CompuMaster.Data.DataQuery.PlatformTools.ClrRuntimePlatform.x64 AndAlso ProbeOdbcDBProvider(MicrosoftAccessConnectionProviderWorkingStatusForOdbcDriver, "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" & TestFile.FilePath & ";Uid=Admin;Pwd=" & databasePassword & ";") Then
                    Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("ODBC", "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" & path & ";Uid=Admin;Pwd=" & databasePassword & ";")
                ElseIf MsJetProviderIsSufficient AndAlso CompuMaster.Data.DataQuery.PlatformTools.CurrentClrRuntime = CompuMaster.Data.DataQuery.PlatformTools.ClrRuntimePlatform.x32 AndAlso ProbeOdbcDBProvider(MicrosoftAccessConnectionProviderWorkingStatusForOdbcDriver, "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & TestFile.FilePath & ";Uid=Admin;Pwd=" & databasePassword & ";") Then
                    Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("ODBC", "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & path & ";Uid=Admin;Pwd=" & databasePassword & ";")
                Else
                    'Let the application find the exception with the most modern provider
                    If CompuMaster.Data.DataQuery.PlatformTools.CurrentClrRuntime = CompuMaster.Data.DataQuery.PlatformTools.ClrRuntimePlatform.x64 Then
                        '64bit - Requires Office 2010 JET drivers
                        Throw New Office2010x64OleDbOdbcEngineRequiredException()
                    ElseIf path.ToLowerInvariant.EndsWith(".accdb") OrElse path.ToLowerInvariant.EndsWith(".accde") Then
                        '32bit - Requires Office 2007 JET drivers
                        Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("OleDB", "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & path & ";Jet OLEDB:Database Password=" & databasePassword & ";")
                    Else
                        '32bit - Requires basic JET drivers
                        Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("OleDB", "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & path & ";Jet OLEDB:Database Password=" & databasePassword & ";")
                    End If
                End If
            Finally
                If TestFile IsNot Nothing Then
                    TestFile.Dispose()
                End If
            End Try
        End Function

        ''' <summary>
        ''' A most probably working Microsoft Excel connection which uses the most appropriate, installed OleDB provider or ODBC driver of the current machine
        ''' </summary>
        ''' <param name="path">The path of the file</param>
        ''' <param name="firstRowContainsHeaders"></param>
        ''' <param name="readAllColumnsAsTextOnly"></param>
        ''' <returns>An OleDB or ODBC data connection to the requested file</returns>
        ''' <remarks></remarks>
        ''' <exception cref="Office2010x64OleDbOdbcEngineRequiredException" />
        Public Shared Function MicrosoftExcelConnection(ByVal path As String, ByVal firstRowContainsHeaders As Boolean, ByVal readAllColumnsAsTextOnly As Boolean) As IDbConnection
            If path = Nothing Then Throw New ArgumentNullException(NameOf(path))
            'Lookup OleDb provider (fast!)
            Dim FoundProvider As String
            FoundProvider = PlatformTools.FindLatestMsOfficeAceOleDbProviderName
            If FoundProvider <> Nothing Then
                Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("OleDB", "Provider=" & FoundProvider & ";Data Source=" & path & ";Extended Properties=""Excel 12.0 Xml;HDR=" & BoolIf(firstRowContainsHeaders, "YES", "NO") & ";" & BoolIf(readAllColumnsAsTextOnly, "IMEX=1", "IMEX=0") & """;")
            End If
            If path.ToLowerInvariant.EndsWith(".mdb") OrElse path.ToLowerInvariant.EndsWith(".mde") Then
                'Try to lookup MS Jet provider which is still fine for this file type
                FoundProvider = PlatformTools.FindLatestMsOfficeJetOleDbProviderName
                If FoundProvider <> Nothing Then
                    Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("OleDB", "Provider=" & FoundProvider & ";Data Source=" & path & ";Extended Properties=""Excel 8.0;HDR=" & BoolIf(firstRowContainsHeaders, "YES", "NO") & ";" & BoolIf(readAllColumnsAsTextOnly, "IMEX=1", "") & """;")
                End If
            End If
            'Probe OleDb provider
            Dim TestFile As TestFile = Nothing
            Try
                Dim MsJetProviderIsSufficient As Boolean
                If path.ToLowerInvariant.EndsWith(".xls") Then
                    TestFile = New TestFile(DataQuery.TestFile.TestFileType.MsExcel95Xls)
                    MsJetProviderIsSufficient = True
                Else
                    TestFile = New TestFile(DataQuery.TestFile.TestFileType.MsExcel2007Xlsx)
                    MsJetProviderIsSufficient = False
                End If
                For MyCounter As Integer = (System.DateTime.Now.Year + 1 - 2000) To 15 Step -1 'try all MS Office releases since 2015 up to current year + 1
                    If ProbeOleDBProvider(MicrosoftExcelConnectionProviderWorkingStatusForACEDynList(MyCounter), "Provider=Microsoft.ACE.OLEDB." & MyCounter & ".0;Data Source=" & TestFile.FilePath & ";Extended Properties=""Excel " & MyCounter & ".0 Xml;HDR=YES;IMEX=0"";") Then
                        'If ProbeOleDBProvider(MicrosoftExcelConnectionProviderWorkingStatusForACEDynList(MyCounter), "Provider=Microsoft.ACE.OLEDB." & MyCounter & ".0;Data Source=" & TestFile.FilePath & ";HDR=YES;IMEX=0"";") Then
                        Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("OleDB", "Provider=Microsoft.ACE.OLEDB." & MyCounter & ".0;Data Source=" & path & ";Extended Properties=""Excel " & MyCounter & ".0 Xml;HDR=" & BoolIf(firstRowContainsHeaders, "YES", "NO") & ";" & BoolIf(readAllColumnsAsTextOnly, "IMEX=1", "IMEX=0") & """;")
                    End If
                Next
                If ProbeOleDBProvider(MicrosoftExcelConnectionProviderWorkingStatusForACE14, "Provider=Microsoft.ACE.OLEDB.14.0;Data Source=" & TestFile.FilePath & ";Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=0"";") Then
                    Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("OleDB", "Provider=Microsoft.ACE.OLEDB.14.0;Data Source=" & path & ";Extended Properties=""Excel 12.0 Xml;HDR=" & BoolIf(firstRowContainsHeaders, "YES", "NO") & ";" & BoolIf(readAllColumnsAsTextOnly, "IMEX=1", "IMEX=0") & """;")
                ElseIf ProbeOleDBProvider(MicrosoftExcelConnectionProviderWorkingStatusForACE12, "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & TestFile.FilePath & ";Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=0"";") Then
                    Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("OleDB", "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & path & ";Extended Properties=""Excel 12.0 Xml;HDR=" & BoolIf(firstRowContainsHeaders, "YES", "NO") & ";" & BoolIf(readAllColumnsAsTextOnly, "IMEX=1", "IMEX=0") & """;")
                ElseIf MsJetProviderIsSufficient AndAlso ProbeOleDBProvider(MicrosoftExcelConnectionProviderWorkingStatusForJet4, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & TestFile.FilePath & ";Extended Properties=""Excel 8.0;HDR=YES;IMEX=0"";") Then
                    Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("OleDB", "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & path & ";Extended Properties=""Excel 8.0;HDR=" & BoolIf(firstRowContainsHeaders, "YES", "NO") & ";" & BoolIf(readAllColumnsAsTextOnly, "IMEX=1", "IMEX=0") & """;")
                ElseIf ProbeOdbcDBProvider(MicrosoftExcelConnectionProviderWorkingStatusForOdbcDriver, "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};Dbq=" & TestFile.FilePath & ";") Then
                    Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("ODBC", "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};Dbq=" & path & ";" & BoolIf(firstRowContainsHeaders, "FirstRowHasNames=1", "FirstRowHasNames=0") & ";" & BoolIf(readAllColumnsAsTextOnly, "ReadOnly=1", "ReadOnly=0") & """;")
                ElseIf MsJetProviderIsSufficient AndAlso ProbeOdbcDBProvider(MicrosoftExcelConnectionProviderWorkingStatusForOdbcDriver, "Driver={Microsoft Excel Driver (*.xls)};Dbq=" & TestFile.FilePath & ";") Then
                    Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("ODBC", "Driver={Microsoft Excel Driver (*.xls)};Dbq=" & path & ";" & BoolIf(firstRowContainsHeaders, "FirstRowHasNames=1", "FirstRowHasNames=0") & ";" & BoolIf(readAllColumnsAsTextOnly, "ReadOnly=1", "ReadOnly=0") & """;")
                Else
                    'Let the application find the exception with the most modern provider
                    If CompuMaster.Data.DataQuery.PlatformTools.CurrentClrRuntime = CompuMaster.Data.DataQuery.PlatformTools.ClrRuntimePlatform.x64 Then
                        '64bit - Requires Office 2010 JET drivers
                        Throw New Office2010x64OleDbOdbcEngineRequiredException()
                    ElseIf path.ToLower.EndsWith(".xlsx") OrElse path.ToLower.EndsWith(".xlsb") OrElse path.ToLower.EndsWith(".xlsm") Then
                        '32bit - Requires Office 2007 JET drivers
                        Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("OleDB", "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & path & ";Extended Properties=""Excel 12.0 Xml;HDR=" & BoolIf(firstRowContainsHeaders, "YES", "NO") & ";" & BoolIf(readAllColumnsAsTextOnly, "IMEX=1", "") & """;")
                    Else
                        '32bit - Requires basic JET drivers
                        Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("OleDB", "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & path & ";Extended Properties=""Excel 8.0;HDR=" & BoolIf(firstRowContainsHeaders, "YES", "NO") & ";" & BoolIf(readAllColumnsAsTextOnly, "IMEX=1", "") & """;")
                    End If
                End If
            Finally
                If TestFile IsNot Nothing Then
                    TestFile.Dispose()
                End If
            End Try

        End Function

        ''' <summary>
        ''' A most probably working Microsoft Excel connection which uses the most appropriate, installed ODBC driver of the current machine
        ''' </summary>
        ''' <param name="path">The path of the file</param>
        ''' <param name="firstRowContainsHeaders"></param>
        ''' <param name="readAllColumnsAsTextOnly"></param>
        ''' <returns>An ODBC data connection to the requested file</returns>
        ''' <remarks></remarks>
        ''' <exception cref="Office2010x64OleDbOdbcEngineRequiredException" />
        Public Shared Function MicrosoftExcelOdbcConnection(ByVal path As String, ByVal firstRowContainsHeaders As Boolean, ByVal readAllColumnsAsTextOnly As Boolean) As IDbConnection
            If path = Nothing Then Throw New ArgumentNullException(NameOf(path))
            Dim TestFile As TestFile = Nothing
            Try
                Dim MsJetProviderIsSufficient As Boolean
                If path.ToLowerInvariant.EndsWith(".xls") Then
                    TestFile = New TestFile(DataQuery.TestFile.TestFileType.MsExcel95Xls)
                    MsJetProviderIsSufficient = True
                Else
                    TestFile = New TestFile(DataQuery.TestFile.TestFileType.MsExcel2007Xlsx)
                    MsJetProviderIsSufficient = False
                End If
                If ProbeOdbcDBProvider(MicrosoftExcelConnectionProviderWorkingStatusForOdbcDriver, "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};Dbq=" & TestFile.FilePath & ";") Then
                    Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("ODBC", "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};Dbq=" & path & ";" & BoolIf(firstRowContainsHeaders, "FirstRowHasNames=1", "FirstRowHasNames=0") & ";" & BoolIf(readAllColumnsAsTextOnly, "ReadOnly=1", "ReadOnly=0") & """;")
                ElseIf MsJetProviderIsSufficient AndAlso ProbeOdbcDBProvider(MicrosoftExcelConnectionProviderWorkingStatusForOdbcDriver, "Driver={Microsoft Excel Driver (*.xls)};Dbq=" & TestFile.FilePath & ";") Then
                    Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("ODBC", "Driver={Microsoft Excel Driver (*.xls)};Dbq=" & path & ";" & BoolIf(firstRowContainsHeaders, "FirstRowHasNames=1", "FirstRowHasNames=0") & ";" & BoolIf(readAllColumnsAsTextOnly, "ReadOnly=1", "ReadOnly=0") & """;")
                Else
                    'Let the application find the exception with the most modern provider
                    If CompuMaster.Data.DataQuery.PlatformTools.CurrentClrRuntime = CompuMaster.Data.DataQuery.PlatformTools.ClrRuntimePlatform.x64 Then
                        '64bit - Requires Office 2010 JET drivers
                        Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("ODBC", "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};Dbq=" & path & ";" & BoolIf(firstRowContainsHeaders, "FirstRowHasNames=1", "FirstRowHasNames=0") & ";" & BoolIf(readAllColumnsAsTextOnly, "ReadOnly=1", "ReadOnly=0") & """;")
                    ElseIf path.ToLower.EndsWith(".xlsx") OrElse path.ToLower.EndsWith(".xlsb") OrElse path.ToLower.EndsWith(".xlsm") Then
                        '32bit - Requires Office 2007 JET drivers
                        Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("ODBC", "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};Dbq=" & path & ";" & BoolIf(firstRowContainsHeaders, "FirstRowHasNames=1", "FirstRowHasNames=0") & ";" & BoolIf(readAllColumnsAsTextOnly, "ReadOnly=1", "ReadOnly=0") & """;")
                    Else
                        '32bit - Requires basic JET drivers
                        Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("ODBC", "Driver={Microsoft Excel Driver (*.xls)};Dbq=" & path & ";" & BoolIf(firstRowContainsHeaders, "FirstRowHasNames=1", "FirstRowHasNames=0") & ";" & BoolIf(readAllColumnsAsTextOnly, "ReadOnly=1", "ReadOnly=0") & """;")
                    End If
                End If
            Finally
                If TestFile IsNot Nothing Then
                    TestFile.Dispose()
                End If
            End Try
        End Function

        ''' <summary>
        ''' A most probably working Microsoft Excel connection which uses the most appropriate, installed OleDB provider of the current machine
        ''' </summary>
        ''' <param name="path">The path of the file</param>
        ''' <param name="firstRowContainsHeaders"></param>
        ''' <param name="readAllColumnsAsTextOnly"></param>
        ''' <returns>An OleDB data connection to the requested file</returns>
        ''' <remarks></remarks>
        ''' <exception cref="Office2010x64OleDbOdbcEngineRequiredException" />
        Public Shared Function MicrosoftExcelOleDbConnection(ByVal path As String, ByVal firstRowContainsHeaders As Boolean, ByVal readAllColumnsAsTextOnly As Boolean) As IDbConnection
            If path = Nothing Then Throw New ArgumentNullException(NameOf(path))
            'Lookup OleDb provider (fast!)
            Dim FoundProvider As String
            FoundProvider = PlatformTools.FindLatestMsOfficeAceOleDbProviderName
            If FoundProvider <> Nothing Then
                Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("OleDB", "Provider=" & FoundProvider & ";Data Source=" & path & ";Extended Properties=""Excel 12.0 Xml;HDR=" & BoolIf(firstRowContainsHeaders, "YES", "NO") & ";" & BoolIf(readAllColumnsAsTextOnly, "IMEX=1", "IMEX=0") & """;")
            End If
            If path.ToLowerInvariant.EndsWith(".mdb") OrElse path.ToLowerInvariant.EndsWith(".mde") Then
                'Try to lookup MS Jet provider which is still fine for this file type
                FoundProvider = PlatformTools.FindLatestMsOfficeJetOleDbProviderName
                If FoundProvider <> Nothing Then
                    Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("OleDB", "Provider=" & FoundProvider & ";Data Source=" & path & ";Extended Properties=""Excel 8.0;HDR=" & BoolIf(firstRowContainsHeaders, "YES", "NO") & ";" & BoolIf(readAllColumnsAsTextOnly, "IMEX=1", "") & """;")
                End If
            End If
            'Probe OleDb provider
            Dim TestFile As TestFile = Nothing
            Try
                Dim MsJetProviderIsSufficient As Boolean
                If path.ToLowerInvariant.EndsWith(".xls") Then
                    TestFile = New TestFile(DataQuery.TestFile.TestFileType.MsExcel95Xls)
                    MsJetProviderIsSufficient = True
                Else
                    TestFile = New TestFile(DataQuery.TestFile.TestFileType.MsExcel2007Xlsx)
                    MsJetProviderIsSufficient = False
                End If
                For MyCounter As Integer = 20 To 15 Step -1
                    If ProbeOleDbProvider(MicrosoftExcelConnectionProviderWorkingStatusForACEDynList(MyCounter), "Provider=Microsoft.ACE.OLEDB." & MyCounter & ".0;Data Source=" & TestFile.FilePath & ";Extended Properties=""Excel " & MyCounter & ".0 Xml;HDR=YES;IMEX=0"";") Then
                        Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("OleDB", "Provider=Microsoft.ACE.OLEDB." & MyCounter & ".0;Data Source=" & path & ";Extended Properties=""Excel " & MyCounter & ".0 Xml;HDR=" & BoolIf(firstRowContainsHeaders, "YES", "NO") & ";" & BoolIf(readAllColumnsAsTextOnly, "IMEX=1", "IMEX=0") & """;")
                    End If
                Next
                If ProbeOleDbProvider(MicrosoftExcelConnectionProviderWorkingStatusForACE14, "Provider=Microsoft.ACE.OLEDB.14.0;Data Source=" & TestFile.FilePath & ";Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=0"";") Then
                    Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("OleDB", "Provider=Microsoft.ACE.OLEDB.14.0;Data Source=" & path & ";Extended Properties=""Excel 12.0 Xml;HDR=" & BoolIf(firstRowContainsHeaders, "YES", "NO") & ";" & BoolIf(readAllColumnsAsTextOnly, "IMEX=1", "IMEX=0") & """;")
                ElseIf ProbeOleDbProvider(MicrosoftExcelConnectionProviderWorkingStatusForACE12, "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & TestFile.FilePath & ";Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=0"";") Then
                    Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("OleDB", "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & path & ";Extended Properties=""Excel 12.0 Xml;HDR=" & BoolIf(firstRowContainsHeaders, "YES", "NO") & ";" & BoolIf(readAllColumnsAsTextOnly, "IMEX=1", "IMEX=0") & """;")
                ElseIf MsJetProviderIsSufficient AndAlso ProbeOleDbProvider(MicrosoftExcelConnectionProviderWorkingStatusForJet4, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & TestFile.FilePath & ";Extended Properties=""Excel 8.0;HDR=YES;IMEX=0"";") Then
                    Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("OleDB", "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & path & ";Extended Properties=""Excel 8.0;HDR=" & BoolIf(firstRowContainsHeaders, "YES", "NO") & ";" & BoolIf(readAllColumnsAsTextOnly, "IMEX=1", "IMEX=0") & """;")
                Else
                    'Let the application find the exception with the most modern provider
                    If CompuMaster.Data.DataQuery.PlatformTools.CurrentClrRuntime = CompuMaster.Data.DataQuery.PlatformTools.ClrRuntimePlatform.x64 Then
                        '64bit - Requires Office 2010 JET drivers
                        Throw New Office2010x64OleDbOdbcEngineRequiredException()
                    ElseIf path.ToLower.EndsWith(".xlsx") OrElse path.ToLower.EndsWith(".xlsb") OrElse path.ToLower.EndsWith(".xlsm") Then
                        '32bit - Requires Office 2007 JET drivers
                        Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("OleDB", "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & path & ";Extended Properties=""Excel 12.0 Xml;HDR=" & BoolIf(firstRowContainsHeaders, "YES", "NO") & ";" & BoolIf(readAllColumnsAsTextOnly, "IMEX=1", "") & """;")
                    Else
                        '32bit - Requires basic JET drivers
                        Return CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("OleDB", "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & path & ";Extended Properties=""Excel 8.0;HDR=" & BoolIf(firstRowContainsHeaders, "YES", "NO") & ";" & BoolIf(readAllColumnsAsTextOnly, "IMEX=1", "") & """;")
                    End If
                End If
            Finally
                If TestFile IsNot Nothing Then
                    TestFile.Dispose()
                End If
            End Try
        End Function

        Private Shared ReadOnly _MicrosoftAccessConnectionProviderWorkingStatusForACEDynList As New Generic.Dictionary(Of Integer, TriState)
        Private Shared ReadOnly Property MicrosoftAccessConnectionProviderWorkingStatusForACEDynList(officeMainVersion As Integer) As TriState
            Get
                If _MicrosoftAccessConnectionProviderWorkingStatusForACEDynList.ContainsKey(officeMainVersion) = False Then
                    _MicrosoftAccessConnectionProviderWorkingStatusForACEDynList(officeMainVersion) = TriState.UseDefault
                End If
                Return _MicrosoftAccessConnectionProviderWorkingStatusForACEDynList(officeMainVersion)
            End Get
        End Property

        Private Shared ReadOnly _MicrosoftExcelConnectionProviderWorkingStatusForACEDynList As New Generic.Dictionary(Of Integer, TriState)
        Private Shared ReadOnly Property MicrosoftExcelConnectionProviderWorkingStatusForACEDynList(officeMainVersion As Integer) As TriState
            Get
                If _MicrosoftExcelConnectionProviderWorkingStatusForACEDynList.ContainsKey(officeMainVersion) = False Then
                    _MicrosoftExcelConnectionProviderWorkingStatusForACEDynList(officeMainVersion) = TriState.UseDefault
                End If
                Return _MicrosoftExcelConnectionProviderWorkingStatusForACEDynList(officeMainVersion)
            End Get
        End Property
        Private Shared MicrosoftExcelConnectionProviderWorkingStatusForACE14 As TriState = TriState.UseDefault
        Private Shared MicrosoftExcelConnectionProviderWorkingStatusForACE12 As TriState = TriState.UseDefault
        Private Shared MicrosoftExcelConnectionProviderWorkingStatusForJet4 As TriState = TriState.UseDefault
        Private Shared MicrosoftExcelConnectionProviderWorkingStatusForOdbcDriver As TriState = TriState.UseDefault
        Private Shared MicrosoftAccessConnectionProviderWorkingStatusForACE14 As TriState = TriState.UseDefault
        Private Shared MicrosoftAccessConnectionProviderWorkingStatusForACE12 As TriState = TriState.UseDefault
        Private Shared MicrosoftAccessConnectionProviderWorkingStatusForJet4 As TriState = TriState.UseDefault
        Private Shared MicrosoftAccessConnectionProviderWorkingStatusForOdbcDriver As TriState = TriState.UseDefault
        Private Shared MicrosoftAccessTextCsvConnectionProviderWorkingStatusForOdbcDriver As TriState = TriState.UseDefault
        Private Shared MicrosoftTextCsvConnectionProviderWorkingStatusForOdbcDriver As TriState = TriState.UseDefault

        ''' <summary>
        ''' Verbose mode creates some additional console output on probe exceptions
        ''' </summary>
        Friend Shared ProbeOleDbOrOdbcProviderVerboseMode As Boolean = False

        ''' <summary>
        ''' Test an OLE DB connection if result hasn't been cached, yet
        ''' </summary>
        ''' <param name="resultCache">A reference to a cache variable</param>
        ''' <param name="connectionString">A working connection string using the provider which shall be tested</param>
        ''' <returns>True if the connectionstring works, False if not</returns>
        ''' <remarks></remarks>
        Private Shared Function ProbeOleDbProvider(ByRef resultCache As TriState, ByVal connectionString As String) As Boolean
            If resultCache = TriState.UseDefault Then
                Dim TestConnection As System.Data.IDbConnection = Nothing
                Try
                    TestConnection = CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("OleDB", connectionString)
                    TestConnection.Open()
                    resultCache = TriState.True
                Catch ex As Exception
                    If ProbeOleDbOrOdbcProviderVerboseMode = True Then
                        Console.WriteLine("Exception at ProbeOleDBProvider: " & ex.Message)
                        PlatformTools.ConsolOutputListOfInstalledOleDbProviders()
                    End If
                    resultCache = TriState.False
                Finally
                    CompuMaster.Data.DataQuery.CloseAndDisposeConnection(TestConnection)
                End Try
            End If
            If resultCache = TriState.True Then
                If ProbeOleDbOrOdbcProviderVerboseMode = True Then Console.WriteLine("Working ProbeOleDBProvider (cached result): " & connectionString)
                Return True
            Else
                If ProbeOleDbOrOdbcProviderVerboseMode = True Then Console.WriteLine("Failing ProbeOleDBProvider (cached result): " & connectionString)
                Return False
            End If
        End Function

        ''' <summary>
        ''' Test an ODBC connection if result hasn't been cached, yet
        ''' </summary>
        ''' <param name="resultCache">A reference to a cache variable</param>
        ''' <param name="connectionString">A working connection string using the provider which shall be tested</param>
        ''' <returns>True if the connectionstring works, False if not</returns>
        ''' <remarks></remarks>
        Private Shared Function ProbeOdbcDBProvider(ByRef resultCache As TriState, ByVal connectionString As String) As Boolean
            If resultCache = TriState.UseDefault Then
                Dim TestConnection As System.Data.IDbConnection = Nothing
                Try
                    TestConnection = CompuMaster.Data.DataQuery.PlatformTools.CreateDataConnection("ODBC", connectionString)
                    TestConnection.Open()
                    resultCache = TriState.True
                Catch ex As Exception
                    If ProbeOleDbOrOdbcProviderVerboseMode = True Then Console.WriteLine("Exception at ProbeOdbcProvider: " & ex.Message)
                    resultCache = TriState.False
                Finally
                    CompuMaster.Data.DataQuery.CloseAndDisposeConnection(TestConnection)
                End Try
            End If
            If resultCache = TriState.True Then
                If ProbeOleDbOrOdbcProviderVerboseMode = True Then Console.WriteLine("Working ProbeOdbcProvider (cached result): " & connectionString)
                Return True
            Else
                If ProbeOleDbOrOdbcProviderVerboseMode = True Then Console.WriteLine("Failing ProbeOdbcProvider (cached result): " & connectionString)
                Return False
            End If
        End Function

        Private Shared Function BoolIf(ByVal expression As Boolean, ByVal trueValue As String, ByVal falseValue As String) As String
            If expression Then Return trueValue Else Return falseValue
        End Function

        ''' <summary>
        ''' Represents a table identifier in an OleDB data source
        ''' </summary>
        ''' <remarks></remarks>
        Public Class OleDbTableDescriptor

            Friend Sub New(ByVal schemaName As String, ByVal tableName As String)
                If tableName = Nothing Then Throw New ArgumentNullException(NameOf(tableName))
                _SchemaName = schemaName
                _TableName = tableName
            End Sub

            Private _SchemaName As String
            ''' <summary>
            ''' The schema name (if supported by the data source)
            ''' </summary>
            ''' <value></value>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Property SchemaName() As String
                Get
                    Return _SchemaName
                End Get
                Set(ByVal value As String)
                    _SchemaName = value
                End Set
            End Property

            Private _TableName As String
            ''' <summary>
            ''' The table name
            ''' </summary>
            ''' <value></value>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Property TableName() As String
                Get
                    Return _TableName
                End Get
                Set(ByVal value As String)
                    _TableName = value
                End Set
            End Property

            ''' <summary>
            ''' The full table identifier as it can be used in select statements
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Overrides Function ToString() As String
                If SchemaName = Nothing Then
                    Return "[" & TableName & "]"
                Else
                    Return "[" & SchemaName & "].[" & TableName & "]"
                End If
            End Function

        End Class

        ''' <summary>
        ''' Represents a table identifier in an ODBC data source
        ''' </summary>
        ''' <remarks></remarks>
        Public Class OdbcTableDescriptor

            Friend Sub New(ByVal schemaName As String, ByVal tableName As String)
                If tableName = Nothing Then Throw New ArgumentNullException(NameOf(tableName))
                _SchemaName = schemaName
                _TableName = tableName
            End Sub

            Private _SchemaName As String
            ''' <summary>
            ''' The schema name (if supported by the data source)
            ''' </summary>
            ''' <value></value>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Property SchemaName() As String
                Get
                    Return _SchemaName
                End Get
                Set(ByVal value As String)
                    _SchemaName = value
                End Set
            End Property

            Private _TableName As String
            ''' <summary>
            ''' The table name
            ''' </summary>
            ''' <value></value>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Property TableName() As String
                Get
                    Return _TableName
                End Get
                Set(ByVal value As String)
                    _TableName = value
                End Set
            End Property

            ''' <summary>
            ''' The full table identifier as it can be used in select statements
            ''' </summary>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Overrides Function ToString() As String
                If SchemaName = Nothing Then
                    Return "[" & TableName & "]"
                Else
                    Return "[" & SchemaName & "].[" & TableName & "]"
                End If
            End Function

        End Class

        ''' <summary>
        ''' Enumerate all tables/views which can be used for SQL SELECT statements
        ''' </summary>
        ''' <param name="openedConnection"></param>
        ''' <returns>The DictionaryEntry contains the table/view name in the key field, the schema name in the value field</returns>
        ''' <remarks></remarks>
        Public Shared Function EnumerateTablesAndViewsInOleDbDataSource(ByVal openedConnection As System.Data.OleDb.OleDbConnection) As OleDbTableDescriptor()
            Dim DbSchema As DataTable = openedConnection.GetSchema()
            Dim DbSchemaCollections As String() = CType(CompuMaster.Data.DataTables.ConvertDataTableToArrayList(DbSchema).ToArray(GetType(String)), String())
            Dim Result As New ArrayList
            If Array.IndexOf(DbSchemaCollections, "Tables") >= 0 Then
                Dim tables As DataTable = openedConnection.GetSchema("Tables")
                For MyCounter As Integer = 0 To tables.Rows.Count - 1
                    Result.Add(New OleDbTableDescriptor(Utils.NoDBNull(tables.Rows(MyCounter)("TABLE_SCHEMA"), CType(Nothing, String)), Utils.NoDBNull(tables.Rows(MyCounter)("TABLE_NAME"), CType(Nothing, String))))
                Next
            End If
            Return CType(Result.ToArray(GetType(OleDbTableDescriptor)), OleDbTableDescriptor())
        End Function

        ''' <summary>
        ''' Enumerate all tables/views which can be used for SQL SELECT statements
        ''' </summary>
        ''' <param name="openedConnection"></param>
        ''' <returns>The DictionaryEntry contains the table/view name in the key field, the schema name in the value field</returns>
        ''' <remarks></remarks>
        Public Shared Function EnumerateTablesAndViewsInOdbcDataSource(ByVal openedConnection As System.Data.Odbc.OdbcConnection) As OdbcTableDescriptor()
            Dim DbSchema As DataTable = openedConnection.GetSchema()
            Dim DbSchemaCollections As String() = CType(CompuMaster.Data.DataTables.ConvertDataTableToArrayList(DbSchema).ToArray(GetType(String)), String())
            Dim Result As New ArrayList
            If Array.IndexOf(DbSchemaCollections, "Tables") >= 0 Then
                Dim tables As DataTable = openedConnection.GetSchema("Tables")
                For MyCounter As Integer = 0 To tables.Rows.Count - 1
                    Result.Add(New OdbcTableDescriptor(Utils.NoDBNull(tables.Rows(MyCounter)("TABLE_SCHEM"), CType(Nothing, String)), Utils.NoDBNull(tables.Rows(MyCounter)("TABLE_NAME"), CType(Nothing, String))))
                Next
            End If
            Return CType(Result.ToArray(GetType(OdbcTableDescriptor)), OdbcTableDescriptor())
        End Function

        ''' <summary>
        ''' In case that no usable data drivers are available on a x64 platform, this exception will be fired
        ''' </summary>
        ''' <remarks></remarks>
        <CodeAnalysis.SuppressMessage("Design", "CA1032:Implement standard exception constructors", Justification:="<Ausstehend>")>
        <CodeAnalysis.SuppressMessage("Usage", "CA2237:Mark ISerializable types with serializable", Justification:="<Ausstehend>")>
        Public Class Office2010x64OleDbOdbcEngineRequiredException
            Inherits System.Exception

            Friend Sub New()
            End Sub

            ''' <summary>
            ''' A collection of probe results for alternative providers at the running machine
            ''' </summary>
            ''' <returns></returns>
            <Obsolete("Never filled any more")> Public Property AlternativeProvidersProbeResults As System.Collections.Specialized.NameValueCollection

            ''' <summary>
            ''' Details on the exception
            ''' </summary>
            ''' <returns></returns>
            Public Overrides ReadOnly Property Message() As String
                Get
                    Return "Microsoft Access Database Engine 2010 x64 Redistributable or newer required to use OleDB/ODBC x64 drivers, please follow " & RecommendedDownloadLink
                End Get
            End Property

            ''' <summary>
            ''' Recommended link for the user to download and install missing components on x64 systems
            ''' </summary>
            ''' <value></value>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public ReadOnly Property RecommendedDownloadLink() As String
                Get
                    Return "http://www.microsoft.com/downloads/en/details.aspx?familyid=C06B8369-60DD-4B64-A44B-84B371EDE16D&displaylang=en"
                End Get
            End Property

        End Class

    End Class

End Namespace
