Option Explicit On 
Option Strict On

Namespace CompuMaster.Data

    Public NotInheritable Class DatabaseManagement

        ''' <summary>
        ''' Create a database file on the specified location, supported file types are .mdb, .accdb, .xls, .xlsx, .xlsm, .xlsb
        ''' </summary>
        ''' <param name="path">The path of the new database file</param>
        ''' <remarks>The folder for the file should already exist and be writable. The file format of the database will be the latest known file type version which is recognized with the file extension.</remarks>
        Public Shared Sub CreateDatabaseFile(ByVal path As String)
            Select Case System.IO.Path.GetExtension(path).ToLower
                Case ".accdb"
                    CreateDatabaseFile(path, DatabaseFileType.MsAccess2007Accdb)
                Case ".mdb"
                    CreateDatabaseFile(path, DatabaseFileType.MsAccess2002Mdb)
                Case ".xls"
                    CreateMsExcelFile(path, MsExcelFileType.MsExcel97Xls)
                Case ".xlsx"
                    CreateMsExcelFile(path, MsExcelFileType.MsExcel2007Xlsx)
                Case ".xlsb"
                    CreateMsExcelFile(path, MsExcelFileType.MsExcel2007Xlsb)
                Case ".xlsm"
                    CreateMsExcelFile(path, MsExcelFileType.MsExcel2007Xlsm)
                Case Else
                    Throw New NotSupportedException("Database file format unknown for selected filename extension """ & System.IO.Path.GetExtension(path) & """")
            End Select
        End Sub

        ''' <summary>
        ''' Create a database file on the specified location, supported file types are .mdb, .accdb
        ''' </summary>
        ''' <param name="path">The path of the new database file</param>
        ''' <remarks>The folder for the file should already exist and be writable. </remarks>
        Public Shared Sub CreateDatabaseFile(ByVal path As String, ByVal databaseFormat As DatabaseFileType)
            Select Case databaseFormat
                Case DatabaseFileType.MsAccess2007Accdb
                    WriteAllBytes(path, LoadBinaryResource("template.accdb"))
                Case DatabaseFileType.MsAccess2002Mdb
                    WriteAllBytes(path, LoadBinaryResource("template.mdb"))
                Case Else
                    Throw New NotSupportedException("Database file format unknown or not supported yet")
            End Select
        End Sub

        Public Shared Sub CreateTextCsvDatabaseFile(ByVal path As String)
            Dim ParentPath As String = System.IO.Path.GetDirectoryName(path)
            If System.IO.Directory.Exists(ParentPath) = False Then System.IO.Directory.CreateDirectory(ParentPath)
            System.IO.File.WriteAllText(path, "Column1" & vbNewLine & "Value1" & vbNewLine & "Value2")
        End Sub

        ''' <summary>
        ''' Create a database file on the specified location, supported file types are .mdb, .accdb
        ''' </summary>
        ''' <param name="path">The path of the new database file</param>
        ''' <remarks>The folder for the file should already exist and be writable. </remarks>
        Public Shared Sub CreateMsExcelFile(ByVal path As String, ByVal excelVersion As MsExcelFileType)
            Select Case excelVersion
                Case MsExcelFileType.MsExcel95Xls
                    WriteAllBytes(path, LoadBinaryResource("template_e95.xls"))
                Case MsExcelFileType.MsExcel97Xls
                    WriteAllBytes(path, LoadBinaryResource("template_e97.xls"))
                Case MsExcelFileType.MsExcel2007Xlsx
                    WriteAllBytes(path, LoadBinaryResource("template_e2007.xlsx"))
                Case MsExcelFileType.MsExcel2007Xlsb
                    WriteAllBytes(path, LoadBinaryResource("template_e2007.xlsb"))
                Case MsExcelFileType.MsExcel2007Xlsm
                    WriteAllBytes(path, LoadBinaryResource("template_e2007.xlsm"))
                Case Else
                    Throw New NotSupportedException("Excel file format unknown")
            End Select
        End Sub

        Public Enum DatabaseFileType As Byte
            MsAccess95Mdb = 0
            MsAccess97Mdb = 1
            MsAccess2000Mdb = 2
            MsAccess2002Mdb = 3
            MsAccess2007Accdb = 4
        End Enum

        Public Enum MsExcelFileType As Byte
            MsExcel95Xls = 0
            MsExcel97Xls = 1
            MsExcel2007Xlsx = 2
            MsExcel2007Xlsb = 3
            MsExcel2007Xlsm = 4
        End Enum

        ''' <summary>
        ''' Write all bytes to a binary file
        ''' </summary>
        ''' <param name="path">The file path for the output</param>
        ''' <param name="bytes">File output data</param>
        ''' <remarks>An existing file will be overwritten</remarks>
        Private Shared Sub WriteAllBytes(ByVal path As String, ByVal bytes As Byte())
            If (bytes Is Nothing) Then
                Throw New ArgumentNullException(NameOf(bytes))
            End If
            Dim stream As System.IO.FileStream = New System.IO.FileStream(path, System.IO.FileMode.Create, System.IO.FileAccess.Write, System.IO.FileShare.Read)
            stream.Write(bytes, 0, bytes.Length)
            stream.Close()
        End Sub

        ''' <summary>
        ''' Read an embedded, binary resource file
        ''' </summary>
        ''' <param name="embeddedFileName">The name of the resouces</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Shared Function LoadBinaryResource(ByVal embeddedFileName As String) As Byte()
            Dim stream As System.IO.Stream = Nothing
            Dim buffer As Byte()
            Try
                stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream(embeddedFileName)
                If stream Is Nothing Then Throw New Exception("Embedded resource not found: " & embeddedFileName)
                ReDim buffer(CInt(stream.Length) - 1)
                stream.Read(buffer, 0, CInt(stream.Length))
            Catch ex As Exception
                Throw New Exception("Failure while loading resource name """ & embeddedFileName & """" & vbNewLine & "Available resource names are: " & String.Join(",", System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceNames), ex)
            Finally
                If stream IsNot Nothing Then stream.Close()
            End Try
            Return buffer
        End Function

    End Class

End Namespace