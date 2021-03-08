Option Explicit On 
Option Strict On

Namespace CompuMaster.Data

    ''' <summary>
    ''' Provide methods for transferring data from and back to a remote database on a data connection
    ''' </summary>
    ''' <remarks></remarks>
    Public NotInheritable Class Manipulation

        ''' <summary>
        ''' DDL languages
        ''' </summary>
        ''' <remarks></remarks>
        Public Enum DdlLanguage
            ''' <summary>
            ''' Do not use automations for creating of tables or columns
            ''' </summary>
            ''' <remarks></remarks>
            NoDDL = 0
            ''' <summary>
            ''' Use the DDL syntax for maintenance of MS Jet Engines like MS Access files
            ''' </summary>
            ''' <remarks></remarks>
            MSJetEngine = 1
            ''' <summary>
            ''' Use the DDL syntax for maintenance of MS SQL Server databases
            ''' </summary>
            ''' <remarks></remarks>
            MSSqlServer = 2

            ''' <summary>
            ''' Use the DDL syntax for maintenance of PostgreSQL databases
            ''' </summary>
            ''' <remarks></remarks>
            PostgreSQL = 3
        End Enum

        ''' <summary>
        ''' A container for a DataTable with its IDataAdapter and IDbCommand
        ''' </summary>
        ''' <remarks></remarks>
        <Obsolete("Use CompuMaster.Data.Manipulation instead", False), ComponentModel.EditorBrowsable(ComponentModel.EditorBrowsableState.Never)> Public Class DataManipulationResults
            Inherits CompuMaster.Data.DataManipulationResult
            <Obsolete("Use CompuMaster.Data.Manipulation instead", True)> Public Sub New()
                MyBase.New(Nothing, Nothing, Nothing, Nothing)
            End Sub
            Friend Sub New(ByVal command As System.Data.IDbCommand, ByVal dataAdapter As System.Data.IDbDataAdapter, commandBuilder As System.Data.Common.DbCommandBuilder)
                MyBase.New(Nothing, command, dataAdapter, commandBuilder)
            End Sub
            Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
                'Do not dispose connections and commands due to internal behaviour as well as compatibility
            End Sub
        End Class

        ''' <summary>
        ''' Write tables of a dataset with their rows into tables on a data connection
        ''' </summary>
        ''' <param name="dataSet">A dataset whose tables shall be transferred to the data connection</param>
        ''' <param name="dataConnection">An opened connection to the data source</param>
        ''' <param name="ddlLanguage">A DDL language which shall be used for creating/extending a table on the data connection</param>
        ''' <param name="dropExistingRowsInDestinationTable"></param>
        ''' <remarks>Missing columns will be added automatically. In case that a column already exist on the remote database and its datatype doesn't match the datatype in the source table, there might be thrown an exception while data transfer.</remarks>
        Public Shared Sub WriteDataSetToDataConnection(ByVal dataSet As DataSet, ByVal dataConnection As IDbConnection, ByVal ddlLanguage As DdlLanguage, ByVal dropExistingRowsInDestinationTable As Boolean)
            For MyCounter As Integer = 0 To dataSet.Tables.Count - 1
                Dim MyTable As DataTable = dataSet.Tables(MyCounter)
                WriteDataTableToDataConnection(MyTable, dataConnection, ddlLanguage, dropExistingRowsInDestinationTable)
            Next
        End Sub

        ''' <summary>
        ''' Write tables of a dataset with their rows into tables on a data connection
        ''' </summary>
        ''' <param name="dataSet">A dataset whose tables shall be transferred to the data connection</param>
        ''' <param name="dataConnection">An opened connection to the data source</param>
        ''' <param name="ddlLanguage">A DDL language which shall be used for creating/extending a table on the data connection</param>
        ''' <param name="dropExistingRowsInDestinationTable"></param>
        ''' <param name="connectionBehaviour">Automations regarding the connection state</param>
        ''' <remarks>Missing columns will be added automatically. In case that a column already exist on the remote database and its datatype doesn't match the datatype in the source table, there might be thrown an exception while data transfer.</remarks>
        Public Shared Sub WriteDataSetToDataConnection(ByVal dataSet As DataSet, ByVal dataConnection As IDbConnection, ByVal ddlLanguage As DdlLanguage, ByVal dropExistingRowsInDestinationTable As Boolean, ByVal connectionBehaviour As CompuMaster.Data.DataQuery.Automations)
            If dataConnection Is Nothing Then Throw New ArgumentNullException(NameOf(dataConnection))
            Try
                'Auto-Open
                Select Case connectionBehaviour
                    Case DataQuery.AnyIDataProvider.Automations.AutoOpenAndCloseAndDisposeConnection, DataQuery.AnyIDataProvider.Automations.AutoOpenConnection
                        CompuMaster.Data.DataQuery.OpenConnection(dataConnection)
                    Case Else
                        'Do Nothing
                End Select
                'Write to database
                WriteDataSetToDataConnection(dataSet, dataConnection, ddlLanguage, dropExistingRowsInDestinationTable)
            Finally
                'Auto-Open
                Select Case connectionBehaviour
                    Case DataQuery.AnyIDataProvider.Automations.AutoCloseAndDisposeConnection
                        CompuMaster.Data.DataQuery.CloseConnection(dataConnection)
                    Case DataQuery.AnyIDataProvider.Automations.AutoOpenAndCloseAndDisposeConnection
                        CompuMaster.Data.DataQuery.CloseAndDisposeConnection(dataConnection)
                    Case Else
                        'Do Nothing
                End Select
            End Try
        End Sub


        ''' <summary>
        ''' Write a datatable with its rows into a table on a data connection
        ''' </summary>
        ''' <param name="table">The table which shall be transferred to the data connection</param>
        ''' <param name="dataConnection">An opened connection to the data source</param>
        ''' <param name="ddlLanguage">A DDL language which shall be used for creating/extending a table on the data connection</param>
        ''' <param name="dropExistingRowsInDestinationTable">If True, all existing rows will be removed first before new rows from the source table will be imported</param>
        ''' <remarks>If the table doesn't exist on the data connection, it will be created automatically if supported by the DDL language. Missing columns will be added automatically. In case that a column already exist on the remote database and its datatype doesn't match the datatype in the source table, there might be thrown an exception while data transfer.</remarks>
        Public Shared Sub WriteDataTableToDataConnection(ByVal table As DataTable, ByVal dataConnection As IDbConnection, ByVal ddlLanguage As DdlLanguage, ByVal dropExistingRowsInDestinationTable As Boolean)
            If table.TableName = Nothing Then
                Throw New ArgumentNullException("table.TableName", "A table name is required")
            End If
            WriteDataTableToDataConnection(table, table.TableName, dataConnection, ddlLanguage, dropExistingRowsInDestinationTable)
        End Sub

        ''' <summary>
        ''' Write a datatable with its rows into a table on a data connection
        ''' </summary>
        ''' <param name="sourceTable">The table which shall be transferred to the data connection</param>
        ''' <param name="remoteTableName"></param>
        ''' <param name="dataConnection">An opened connection to the data source</param>
        ''' <param name="ddlLanguage">A DDL language which shall be used for creating/extending a table on the data connection</param>
        ''' <param name="dropExistingRowsInDestinationTable">If True, all existing rows will be removed first before new rows from the source table will be imported</param>
        ''' <remarks>If the table doesn't exist on the data connection, it will be created automatically if supported by the DDL language. Missing columns will be added automatically. In case that a column already exist on the remote database and its datatype doesn't match the datatype in the source table, there might be thrown an exception while data transfer.</remarks>
        Public Shared Sub WriteDataTableToDataConnection(ByVal sourceTable As DataTable, ByVal remoteTableName As String, ByVal dataConnection As IDbConnection, ByVal ddlLanguage As DdlLanguage, ByVal dropExistingRowsInDestinationTable As Boolean)
            If dataConnection Is Nothing Then Throw New ArgumentNullException(NameOf(dataConnection))
            If dataConnection.State <> ConnectionState.Open Then Throw New ArgumentException("dataConnection.ConnectionState is not open")
            If remoteTableName = Nothing Then
                remoteTableName = sourceTable.TableName
            End If
            If remoteTableName = Nothing Then
                Throw New ArgumentNullException(NameOf(remoteTableName))
            End If
            Dim RemoteTable As DataTable = LoadTableStructureWith1RowFromConnection(remoteTableName, dataConnection, True)

            'Create remote table if required
            If RemoteTable Is Nothing Then
                Dim PrimaryKeyColumn As String
                If sourceTable.PrimaryKey IsNot Nothing AndAlso sourceTable.PrimaryKey.Length = 1 AndAlso sourceTable.PrimaryKey(0).AutoIncrement = True AndAlso sourceTable.PrimaryKey(0).Unique = True AndAlso sourceTable.PrimaryKey(0).DataType Is GetType(Integer) Then
                    PrimaryKeyColumn = sourceTable.PrimaryKey(0).ColumnName
                Else
                    PrimaryKeyColumn = CompuMaster.Data.DataTablesTools.LookupUniqueColumnName(sourceTable, "PrimaryKeyID")
                End If
                Dim ColumnCreationCommandText As String = CreateTableCommandText(remoteTableName, PrimaryKeyColumn, ddlLanguage)
                CompuMaster.Data.DataQuery.ExecuteNonQuery(dataConnection, ColumnCreationCommandText, CommandType.Text, Nothing, CompuMaster.Data.DataQuery.Automations.None, 0)
                RemoteTable = LoadTableStructureWith1RowFromConnection(remoteTableName, dataConnection, False)
            End If

            'Extend schema if required
            Dim extendSchemaCommandText As String = AddMissingColumnsCommandText(sourceTable, RemoteTable, ddlLanguage)
            If extendSchemaCommandText <> Nothing Then
                CompuMaster.Data.DataQuery.ExecuteNonQuery(dataConnection, extendSchemaCommandText, CommandType.Text, Nothing, CompuMaster.Data.DataQuery.Automations.None, 0)
            End If
            RemoteTable = LoadTableStructureWith1RowFromConnection(remoteTableName, dataConnection, False)

            'Load data for manipulation process
            Dim datacontainer As CompuMaster.Data.DataManipulationResult = Nothing
            datacontainer = LoadTableDataForManipulationViaCode(dataConnection, remoteTableName)
            'Dim recordCountBefore As Integer = datacontainer.Table.Rows.Count
            'Manipulate table
            CompuMaster.Data.DataTables.CreateDataTableClone(sourceTable, datacontainer.Table, "", "", 0, False, dropExistingRowsInDestinationTable, False, False, True)
            'Now let's write the changes back to the database
            UpdateCodeManipulatedData(datacontainer)

        End Sub

        ''' <summary>
        ''' Write a datatable with its rows into a table on a data connection
        ''' </summary>
        ''' <param name="sourceTable">The table which shall be transferred to the data connection</param>
        ''' <param name="dataConnection">An opened connection to the data source</param>
        ''' <param name="ddlLanguage">A DDL language which shall be used for creating/extending a table on the data connection</param>
        ''' <param name="dropExistingRowsInDestinationTable">If True, all existing rows will be removed first before new rows from the source table will be imported</param>
        ''' <param name="connectionBehaviour">Automations regarding the connection state</param>
        ''' <remarks>If the table doesn't exist on the data connection, it will be created automatically if supported by the DDL language. Missing columns will be added automatically. In case that a column already exist on the remote database and its datatype doesn't match the datatype in the source table, there might be thrown an exception while data transfer.</remarks>
        Public Shared Sub WriteDataTableToDataConnection(ByVal sourceTable As DataTable, ByVal dataConnection As IDbConnection, ByVal ddlLanguage As DdlLanguage, ByVal dropExistingRowsInDestinationTable As Boolean, ByVal connectionBehaviour As CompuMaster.Data.DataQuery.Automations)
            If dataConnection Is Nothing Then Throw New ArgumentNullException(NameOf(dataConnection))
            Try
                'Auto-Open
                Select Case connectionBehaviour
                    Case DataQuery.AnyIDataProvider.Automations.AutoOpenAndCloseAndDisposeConnection, DataQuery.AnyIDataProvider.Automations.AutoOpenConnection
                        CompuMaster.Data.DataQuery.OpenConnection(dataConnection)
                    Case Else
                        'Do Nothing
                End Select
                'Write to database
                WriteDataTableToDataConnection(sourceTable, dataConnection, ddlLanguage, dropExistingRowsInDestinationTable)
            Finally
                'Auto-Close
                Select Case connectionBehaviour
                    Case DataQuery.AnyIDataProvider.Automations.AutoCloseAndDisposeConnection
                        CompuMaster.Data.DataQuery.CloseConnection(dataConnection)
                    Case DataQuery.AnyIDataProvider.Automations.AutoOpenAndCloseAndDisposeConnection
                        CompuMaster.Data.DataQuery.CloseAndDisposeConnection(dataConnection)
                    Case Else
                        'Do Nothing
                End Select
            End Try
        End Sub

        ''' <summary>
        ''' Write a datatable with its rows into a table on a data connection
        ''' </summary>
        ''' <param name="sourceTable">The table which shall be transferred to the data connection</param>
        ''' <param name="remoteTableName"></param>
        ''' <param name="dataConnection">An opened connection to the data source</param>
        ''' <param name="ddlLanguage">A DDL language which shall be used for creating/extending a table on the data connection</param>
        ''' <param name="dropExistingRowsInDestinationTable">If True, all existing rows will be removed first before new rows from the source table will be imported</param>
        ''' <param name="connectionBehaviour">Automations regarding the connection state</param>
        ''' <remarks>If the table doesn't exist on the data connection, it will be created automatically if supported by the DDL language. Missing columns will be added automatically. In case that a column already exist on the remote database and its datatype doesn't match the datatype in the source table, there might be thrown an exception while data transfer.</remarks>
        Public Shared Sub WriteDataTableToDataConnection(ByVal sourceTable As DataTable, ByVal remoteTableName As String, ByVal dataConnection As IDbConnection, ByVal ddlLanguage As DdlLanguage, ByVal dropExistingRowsInDestinationTable As Boolean, ByVal connectionBehaviour As CompuMaster.Data.DataQuery.Automations)
            If dataConnection Is Nothing Then Throw New ArgumentNullException(NameOf(dataConnection))
            Try
                'Auto-Open
                Select Case connectionBehaviour
                    Case DataQuery.AnyIDataProvider.Automations.AutoOpenAndCloseAndDisposeConnection, DataQuery.AnyIDataProvider.Automations.AutoOpenConnection
                        CompuMaster.Data.DataQuery.OpenConnection(dataConnection)
                    Case Else
                        'Do Nothing
                End Select
                'Write to database
                WriteDataTableToDataConnection(sourceTable, remoteTableName, dataConnection, ddlLanguage, dropExistingRowsInDestinationTable)
            Finally
                'Auto-Open
                Select Case connectionBehaviour
                    Case DataQuery.AnyIDataProvider.Automations.AutoCloseAndDisposeConnection
                        CompuMaster.Data.DataQuery.CloseConnection(dataConnection)
                    Case DataQuery.AnyIDataProvider.Automations.AutoOpenAndCloseAndDisposeConnection
                        CompuMaster.Data.DataQuery.CloseAndDisposeConnection(dataConnection)
                    Case Else
                        'Do Nothing
                End Select
            End Try
        End Sub

        ''' <summary>
        ''' Create a script for creating an empty table with just a single primary ID key field
        ''' </summary>
        ''' <param name="tableName">The table name which shall be created</param>
        ''' <param name="primaryColumnName">The name for the primary, auto-increment ID field</param>
        ''' <param name="ddlLanguage">The DDL language which shall be used</param>
        ''' <returns>A string containing a command text which can be executed against a data connection</returns>
        ''' <remarks></remarks>
        Private Shared Function CreateTableCommandText(ByVal tableName As String, ByVal primaryColumnName As String, ByVal ddlLanguage As DdlLanguage) As String
            Dim OpenBrackets, CloseBrackets As String
            If tableName.IndexOf("[") >= 0 AndAlso tableName.IndexOf("]") >= 0 Then
                'table name already in a well-formed syntax, e.g. "dbo.[Test - 123]"
                OpenBrackets = Nothing
                CloseBrackets = Nothing
            Else
                'table name (e.g. "Test - 123") requires a well-formed syntax (e.g. [Test - 123])
                OpenBrackets = "["
                CloseBrackets = "]"
            End If

            If ddlLanguage = DdlLanguage.PostgreSQL Then
                OpenBrackets = """"
                CloseBrackets = """"
            End If

            Select Case ddlLanguage
                Case DdlLanguage.MSJetEngine
                    Return "CREATE TABLE " & OpenBrackets & tableName & CloseBrackets & " ([" & primaryColumnName & "] AUTOINCREMENT, Primary Key ([" & primaryColumnName & "]))"
                Case DdlLanguage.MSSqlServer
                    Return "CREATE TABLE " & OpenBrackets & tableName & CloseBrackets & " ([" & primaryColumnName & "] int NOT NULL IDENTITY (1, 1) PRIMARY KEY)"
                Case DdlLanguage.PostgreSQL
                    Return "CREATE TABLE " & OpenBrackets & tableName & CloseBrackets & " ( " & primaryColumnName & " SERIAL NOT NULL PRIMARY KEY)"
                Case Else
                    Throw New NotSupportedException("CreateTableCommandText not supported for " & ddlLanguage.ToString)
            End Select
        End Function

        Private Shared Function IsStringWithA2ZOnly(value As String) As Boolean
            Dim pattern As String = "^[a-zA-Z]+$"
            Dim reg As New System.Text.RegularExpressions.Regex(pattern)
            Return reg.IsMatch(value)
        End Function

        ''' <summary>
        ''' Create an SQL command text to create missing columns on the remote database
        ''' </summary>
        ''' <param name="sourceTable">The table which shall be written into the remote database</param>
        ''' <param name="destinationTable">The table as it is currently on the remote database</param>
        ''' <param name="ddlLanguage">The SQL language which shall be used</param>
        ''' <returns>A valid command text to create missing columns on the remote database</returns>
        ''' <remarks>
        ''' This function doesn't create any column update commands to change existing columns; it just creates commands for adding additional columns.
        ''' </remarks>
        Public Shared Function AddMissingColumnsCommandText(ByVal sourceTable As DataTable, ByVal destinationTable As DataTable, ByVal ddlLanguage As DdlLanguage) As String
            Return AddMissingColumnsCommandText(sourceTable, destinationTable, destinationTable.TableName, ddlLanguage)
        End Function

        ''' <summary>
        ''' Create an SQL command text to create missing columns on the remote database
        ''' </summary>
        ''' <param name="sourceTable">The table which shall be written into the remote database</param>
        ''' <param name="destinationTable">The table as it is currently on the remote database</param>
        ''' <param name="remoteTableName">The table name at the remote server</param>
        ''' <param name="ddlLanguage">The SQL language which shall be used</param>
        ''' <returns>A valid command text to create missing columns on the remote database</returns>
        ''' <remarks>
        ''' This function doesn't create any column update commands to change existing columns; it just creates commands for adding additional columns.
        ''' </remarks>
        <CodeAnalysis.SuppressMessage("Minor Code Smell", "S1643:Strings should not be concatenated using ""+"" or ""&"" in a loop", Justification:="<Ausstehend>")>
        Public Shared Function AddMissingColumnsCommandText(ByVal sourceTable As DataTable, ByVal destinationTable As DataTable, remoteTableName As String, ByVal ddlLanguage As DdlLanguage) As String
            Dim OpenBrackets, CloseBrackets As String
            If remoteTableName.IndexOf("[") >= 0 AndAlso remoteTableName.IndexOf("]") >= 0 Then
                'table name already in a well-formed syntax, e.g. "dbo.[Test - 123]"
                OpenBrackets = Nothing
                CloseBrackets = Nothing
            Else
                'table name (e.g. "Test - 123") requires a well-formed syntax (e.g. [Test - 123])
                OpenBrackets = "["
                CloseBrackets = "]"
            End If
            Select Case ddlLanguage

                Case DdlLanguage.PostgreSQL
                    Dim ColumnCreationArguments As String = Nothing
                    OpenBrackets = """"
                    CloseBrackets = """"
                    Dim TableOpenBrackets As String
                    Dim TableCloseBrackets As String
                    If remoteTableName.Contains(".""") = False Then
                        TableOpenBrackets = """"
                        TableCloseBrackets = """"
                    Else
                        TableOpenBrackets = Nothing
                        TableCloseBrackets = Nothing
                    End If
                    For Each MyColumn As DataColumn In sourceTable.Columns
                        '  Dim myColumnName As String = IIf(MyColumn.ColumnName.Contains(" "), MyColumn
                        If destinationTable.Columns.Contains(MyColumn.ColumnName) = False Then
                            If ColumnCreationArguments <> Nothing Then ColumnCreationArguments &= ", ADD COLUMN "
                            If IsStringWithA2ZOnly(MyColumn.ColumnName) = False Then
                                ColumnCreationArguments &= OpenBrackets & MyColumn.ColumnName & CloseBrackets & " "
                            Else
                                ColumnCreationArguments &= MyColumn.ColumnName & " "
                            End If

                            Select Case MyColumn.DataType.Name
                                Case "String"
                                    ColumnCreationArguments &= "text"
                                Case "DateTime"
                                    ColumnCreationArguments &= "timestamp"
                                Case "Boolean"
                                    ColumnCreationArguments &= "boolean"
                                Case "Byte"
                                    ColumnCreationArguments &= "smallint"
                                Case "Int16"
                                    ColumnCreationArguments &= "smallint"
                                Case "Int32"
                                    ColumnCreationArguments &= "int"
                                Case "Int64"
                                    ColumnCreationArguments &= "bigint"
                                Case "Double"
                                    ColumnCreationArguments &= "numeric (16,4)"
                                Case "Decimal"
                                    ColumnCreationArguments &= "numeric (16,4)"
                                Case Else
                                    Throw New NotSupportedException("Data type """ & MyColumn.DataType.Name & """ for column """ & MyColumn.ColumnName & """ not supported for auto-adding in database")
                            End Select
                            ColumnCreationArguments &= " NULL" & vbNewLine
                        End If
                    Next
                    If ColumnCreationArguments <> Nothing Then ColumnCreationArguments = "ALTER TABLE " & TableOpenBrackets & remoteTableName & TableCloseBrackets & " ADD COLUMN " & ColumnCreationArguments
                    Return ColumnCreationArguments
                    'Return Nothing
                Case DdlLanguage.MSJetEngine
                    Dim ColumnCreationArguments As String = Nothing
                    For Each MyColumn As DataColumn In sourceTable.Columns
                        If destinationTable.Columns.Contains(MyColumn.ColumnName) = False Then
                            If ColumnCreationArguments <> Nothing Then ColumnCreationArguments &= ","
                            ColumnCreationArguments &= "[" & MyColumn.ColumnName & "] "
                            Select Case MyColumn.DataType.Name
                                Case "String"
                                    ColumnCreationArguments &= "memo"
                                Case "DateTime"
                                    ColumnCreationArguments &= "date"
                                Case "Boolean"
                                    ColumnCreationArguments &= "bit"
                                Case "Byte"
                                    ColumnCreationArguments &= "integer"
                                Case "Int16"
                                    ColumnCreationArguments &= "integer"
                                Case "Int32"
                                    ColumnCreationArguments &= "long"
                                Case "Int64"
                                    ColumnCreationArguments &= "long"
                                Case "Single"
                                    ColumnCreationArguments &= "single"
                                Case "Double"
                                    ColumnCreationArguments &= "double"
                                Case "Decimal"
                                    ColumnCreationArguments &= "double"
                                Case Else
                                    Throw New NotSupportedException("Data type """ & MyColumn.DataType.Name & """ for column """ & MyColumn.ColumnName & """ not supported for auto-adding in database")
                            End Select
                            ColumnCreationArguments &= " NULL" & vbNewLine
                        End If
                    Next
                    If ColumnCreationArguments <> Nothing Then ColumnCreationArguments = "ALTER TABLE " & OpenBrackets & remoteTableName & CloseBrackets & " ADD " & ColumnCreationArguments
                    Return ColumnCreationArguments
                Case DdlLanguage.MSSqlServer
                    Dim ColumnCreationArguments As String = Nothing
                    For Each MyColumn As DataColumn In sourceTable.Columns
                        If destinationTable.Columns.Contains(MyColumn.ColumnName) = False Then
                            If ColumnCreationArguments <> Nothing Then ColumnCreationArguments &= ","
                            ColumnCreationArguments &= "[" & MyColumn.ColumnName & "] "
                            Select Case MyColumn.DataType.Name
                                Case "String"
                                    ColumnCreationArguments &= "ntext"
                                Case "DateTime"
                                    ColumnCreationArguments &= "datetime"
                                Case "Boolean"
                                    ColumnCreationArguments &= "bit"
                                Case "Byte"
                                    ColumnCreationArguments &= "tinyint"
                                Case "Int16"
                                    ColumnCreationArguments &= "smallint"
                                Case "Int32"
                                    ColumnCreationArguments &= "int"
                                Case "Int64"
                                    ColumnCreationArguments &= "bigint"
                                Case "Double"
                                    ColumnCreationArguments &= "decimal (16,4)"
                                Case "Decimal"
                                    ColumnCreationArguments &= "decimal (16,4)"
                                Case Else
                                    Throw New NotSupportedException("Data type """ & MyColumn.DataType.Name & """ for column """ & MyColumn.ColumnName & """ not supported for auto-adding in database")
                            End Select
                            ColumnCreationArguments &= " NULL" & vbNewLine
                        End If
                    Next
                    If ColumnCreationArguments <> Nothing Then ColumnCreationArguments = "ALTER TABLE " & OpenBrackets & remoteTableName & CloseBrackets & " ADD " & ColumnCreationArguments
                    Return ColumnCreationArguments
                Case Else
                    Throw New NotSupportedException("AddMissingColumnsCommandText not supported for " & ddlLanguage.ToString)
            End Select
        End Function

        ''' <summary>
        ''' Load table data from the data connection in a mode for submitting changes in a later step
        ''' </summary>
        ''' <param name="dataConnection">An opened connection to the data source</param>
        ''' <param name="tableName">The name of a table on the database</param>
        ''' <returns>An DataManipulationResults object with the returned data</returns>
        ''' <remarks></remarks>
        Public Shared Function LoadTableDataForManipulationViaCode(ByVal dataConnection As IDbConnection, ByVal tableName As String) As CompuMaster.Data.DataManipulationResult
            Return LoadTableDataForManipulationViaCode(dataConnection, tableName, 0, False)
        End Function

        ''' <summary>
        ''' Load table data from the data connection in a mode for submitting changes in a later step
        ''' </summary>
        ''' <param name="dataConnection">An opened connection to the data source</param>
        ''' <param name="tableName">The name of a table on the database</param>
        ''' <param name="commandTimeout">A timeout for the command in seconds</param>
        ''' <returns>An DataManipulationResults object with the returned data</returns>
        ''' <remarks></remarks>
        Public Shared Function LoadTableDataForManipulationViaCode(ByVal dataConnection As IDbConnection, ByVal tableName As String, ByVal commandTimeout As Integer) As CompuMaster.Data.DataManipulationResult
            Return LoadTableDataForManipulationViaCode(dataConnection, tableName, commandTimeout, False)
        End Function

        ''' <summary>
        ''' Load table data from the data connection in a mode for submitting changes in a later step
        ''' </summary>
        ''' <param name="dataConnection">An opened connection to the data source</param>
        ''' <param name="tableName">The name of a table on the database</param>
        ''' <param name="commandTimeout">A timeout for the command in seconds</param>
        ''' <param name="isSafeTableName">Table name is already in a well-formed syntax for the underlying provider, e.g. dbo.[Test - 123] or public."Test - 123"</param>
        ''' <returns>An DataManipulationResults object with the returned data</returns>
        ''' <remarks></remarks>
        Public Shared Function LoadTableDataForManipulationViaCode(ByVal dataConnection As IDbConnection, ByVal tableName As String, ByVal commandTimeout As Integer, isSafeTableName As Boolean) As CompuMaster.Data.DataManipulationResult
            Dim OpenBrackets, CloseBrackets As String
            If tableName.IndexOf("[") >= 0 AndAlso tableName.IndexOf("]") >= 0 Then
                'table name already in a well-formed syntax, e.g. "dbo.[Test - 123]"
                OpenBrackets = Nothing
                CloseBrackets = Nothing
            Else
                'table name (e.g. "Test - 123") requires a well-formed syntax (e.g. [Test - 123])
                OpenBrackets = "["
                CloseBrackets = "]"
            End If
            If CType(dataConnection, Object).GetType.ToString = "Npgsql.NpgsqlConnection" Then
                If tableName.Contains(".""") = False Then
                    OpenBrackets = """"
                    CloseBrackets = """"
                Else
                    OpenBrackets = Nothing
                    CloseBrackets = Nothing
                End If
            End If
            If isSafeTableName Then
                'table name already in a well-formed syntax, e.g. "dbo.[Test - 123]"
                OpenBrackets = Nothing
                CloseBrackets = Nothing
            End If

            'Prepare the command 
            Dim MyCmd As System.Data.IDbCommand = dataConnection.CreateCommand
            MyCmd.CommandText = "SELECT * FROM " & OpenBrackets & tableName & CloseBrackets
            MyCmd.CommandType = CommandType.Text
            MyCmd.CommandTimeout = commandTimeout

            Return LoadDataForManipulationViaCode(dataConnection, MyCmd)

        End Function

        ''' <summary>
        ''' Load data from the data connection in a mode for submitting changes in a later step
        ''' </summary>
        ''' <param name="dataConnection">An opened connection to the data source</param>
        ''' <param name="command">A prepared command object</param>
        ''' <returns>An DataManipulationResults object with the returned data</returns>
        ''' <remarks></remarks>
        Private Shared Function LoadDataForManipulationViaCode(ByVal dataConnection As IDbConnection, ByVal command As IDbCommand) As CompuMaster.Data.DataManipulationResult

            Dim Result As CompuMaster.Data.DataManipulationResult
            Dim CurrentProvider As Data.DataQuery.DataProvider = Data.DataQuery.DataProvider.LookupDataProvider(dataConnection)
            If CurrentProvider IsNot Nothing Then
                Dim MyCmdsPrepareDA As System.Data.IDbDataAdapter = CurrentProvider.CreateDataAdapter(command)
                Dim MyCmdsPrepareCmdBuilder As System.Data.Common.DbCommandBuilder = CurrentProvider.CreateCommandBuilder(MyCmdsPrepareDA)
                Dim MyDA As System.Data.IDbDataAdapter = CurrentProvider.CreateDataAdapter(command)
                'MyDA.MissingSchemaAction = MissingSchemaAction.Add
                'MyDA.MissingMappingAction = MissingMappingAction.Passthrough

                'Load the data
                Dim MyCmdBuilder As System.Data.Common.DbCommandBuilder = CurrentProvider.CreateCommandBuilder(MyDA)
                Result = New CompuMaster.Data.DataManipulationResult(command, MyDA, MyCmdBuilder)
                CType(MyDA, System.Data.Common.DbDataAdapter).Fill(Result.Table)

                'Create required data manipulation commands
                MyDA.DeleteCommand = MyCmdsPrepareCmdBuilder.GetDeleteCommand()
                MyDA.InsertCommand = MyCmdsPrepareCmdBuilder.GetInsertCommand()
                MyDA.UpdateCommand = MyCmdsPrepareCmdBuilder.GetUpdateCommand()

                'Do some provider-specific stuff for better compatibility/stability
                Select Case CurrentProvider.Title
                    Case "SqlClient", "OleDb", "ODBC"
                        'Auto-Fix delete/insert/update commands to support field names with reserved names by adding brackets [ ] around the field names
                        'Dim remoteColumnNames As String() = LookupColumnNamesOnRemoteTable(MyDA.InsertCommand.CommandText, MyDA.DeleteCommand.CommandText)
                        Dim remoteColumnNames As String() = LookupColumnNamesOnRemoteTable(Result.Table)
                        For MyCounter As Integer = 0 To remoteColumnNames.Length - 1
                            Dim remoteTableColumnName As String = remoteColumnNames(MyCounter)
                            AutoFixCommandColumnNames(MyDA.DeleteCommand, MyDA.InsertCommand, MyDA.UpdateCommand, remoteTableColumnName)
                        Next
                    Case "Npgsql"
                        MyDA.DeleteCommand.UpdatedRowSource = UpdateRowSource.None
                        MyDA.InsertCommand.UpdatedRowSource = UpdateRowSource.None
                        MyDA.UpdateCommand.UpdatedRowSource = UpdateRowSource.None
                End Select
            Else
                Throw New NotSupportedException("Data provider not supported yet")
            End If

            Return Result

        End Function

        ''' <summary>
        ''' Auto-Fix delete/insert/update commands to support field names with reserved names by adding brackets [ ] around the field names
        ''' </summary>
        ''' <param name="DeleteCommand"></param>
        ''' <param name="InsertCommand"></param>
        ''' <param name="UpdateCommand"></param>
        ''' <param name="remoteTableColumnName"></param>
        Private Shared Sub AutoFixCommandColumnNames(DeleteCommand As IDbCommand, InsertCommand As IDbCommand, UpdateCommand As IDbCommand, remoteTableColumnName As String)
            DeleteCommand.CommandText = Replace(DeleteCommand.CommandText, "(" & remoteTableColumnName & " = ", "([" & remoteTableColumnName & "] = ")
            DeleteCommand.CommandText = Replace(DeleteCommand.CommandText, ", " & remoteTableColumnName & " = ", ", [" & remoteTableColumnName & "] = ")
            DeleteCommand.CommandText = Replace(DeleteCommand.CommandText, "[[" & remoteTableColumnName & "]] = ", "[" & remoteTableColumnName & "] = ")
            InsertCommand.CommandText = Replace(InsertCommand.CommandText, "(" & remoteTableColumnName & ",", "([" & remoteTableColumnName & "],")
            InsertCommand.CommandText = Replace(InsertCommand.CommandText, "(" & remoteTableColumnName & ")", "([" & remoteTableColumnName & "])")
            InsertCommand.CommandText = Replace(InsertCommand.CommandText, ", " & remoteTableColumnName & ")", ", [" & remoteTableColumnName & "])")
            InsertCommand.CommandText = Replace(InsertCommand.CommandText, ", " & remoteTableColumnName & ", ", ", [" & remoteTableColumnName & "], ")
            InsertCommand.CommandText = Replace(InsertCommand.CommandText, "[[" & remoteTableColumnName & "]]", "[" & remoteTableColumnName & "]")
            UpdateCommand.CommandText = Replace(UpdateCommand.CommandText, " " & remoteTableColumnName & " = ", " [" & remoteTableColumnName & "] = ")
        End Sub

        ''' <summary>
        ''' Lookup a full set of column names used in a table
        ''' </summary>
        ''' <param name="table">A data table</param>
        ''' <returns>An array of strings with the column names of the data table</returns>
        Private Shared Function LookupColumnNamesOnRemoteTable(ByVal table As DataTable) As String()
            Dim Result As New ArrayList
            For MyCounter As Integer = 0 To table.Columns.Count - 1
                Result.Add(table.Columns(MyCounter).ColumnName)
            Next
            Return CType(Result.ToArray(GetType(String)), String())
        End Function

        '''' <summary>
        '''' Lookup a full set of column names used in INSERT+UPDATE+DELETE statements
        '''' </summary>
        '''' <param name="insertCommandCreatedByCommandBuilder">An insert command as it has been created by the CommandBuilder</param>
        '''' <param name="deleteCommandCreatedByCommandBuilder">A delete command as it has been created by the CommandBuilder</param>
        '''' <returns>A list of all columns in a table</returns>
        '''' <remarks>The INSERT statement doesn't contain the auto increment keys, typically the primary ID key. The DELETE statement only contains the primary ID key(s) in its WHERE clause. Summarized, both statements contain the full list of column names used in all command texts.</remarks>
        'Private Shared Function LookupColumnNamesOnRemoteTable(ByVal insertCommandCreatedByCommandBuilder As String, ByVal deleteCommandCreatedByCommandBuilder As String) As String()
        '    'Lookup data fields from INSERT command
        '    If insertCommandCreatedByCommandBuilder = Nothing OrElse _
        '            Not insertCommandCreatedByCommandBuilder.StartsWith("INSERT INTO ") OrElse _
        '            insertCommandCreatedByCommandBuilder.IndexOf("("c) <= 0 Then
        '        Throw New NotSupportedException("INSERT statement not supported for lookup of column names")
        '    End If
        '    Dim ColumnNames As String = insertCommandCreatedByCommandBuilder.Substring(insertCommandCreatedByCommandBuilder.IndexOf("("c) + 1, insertCommandCreatedByCommandBuilder.IndexOf(")"c) - insertCommandCreatedByCommandBuilder.IndexOf("("c) - 1)
        '    Dim Result As New ArrayList(ColumnNames.Split(","c))
        '    For MyCounter As Integer = 0 To Result.Count - 1
        '        Result.Item(MyCounter) = CType(Result.Item(MyCounter), String).Trim
        '        If CType(Result.Item(MyCounter), String).StartsWith("[") AndAlso CType(Result.Item(MyCounter), String).EndsWith("]") Then
        '            Result.Item(MyCounter) = CType(Result.Item(MyCounter), String).Substring(1, CType(Result.Item(MyCounter), String).Length - 2)
        '        End If
        '    Next

        '    'Lookup primary ID fields from DELETE command
        '    'TODO: identify column names from WHERE clause in deleteCommandCreatedByCommandBuilder (for single-field-PKs as well as for multiple-field-PKs)
        '    'TODO: add to result array but without duplicates (what is a duplicate - case sensitive or case insensitive? Is it the same with all database-DDLs?)

        '    'Return results
        '    Return CType(Result.ToArray(GetType(String)), String())
        'End Function

        'Private Shared Function FindTableMappingForDatasetTableName(ByVal tableMappingCollection As System.Data.Common.DataTableMappingCollection, ByVal tableName As String) As System.Data.Common.DataTableMapping
        '    For MyCounter As Integer = 0 To tableMappingCollection.Count - 1
        '        If tableMappingCollection(MyCounter).DataSetTable = tableName Then
        '            Return tableMappingCollection(MyCounter)
        '        End If
        '    Next
        '    Return Nothing
        'End Function

        ''' <summary>
        ''' Query the data from the data connection in a mode for submitting changes in a later step
        ''' </summary>
        ''' <param name="dataConnection">An opened connection to the data source</param>
        ''' <param name="selectStatement">The name of a table on the database</param>
        ''' <returns>An DataManipulationResults object with the returned data</returns>
        ''' <remarks></remarks>
        Public Shared Function LoadQueryDataForManipulationViaCode(ByVal dataConnection As IDbConnection, ByVal selectStatement As String) As CompuMaster.Data.DataManipulationResult
            Return LoadQueryDataForManipulationViaCode(dataConnection, selectStatement, 0)
        End Function

        ''' <summary>
        ''' Query the data from the data connection in a mode for submitting changes in a later step
        ''' </summary>
        ''' <param name="dataConnection">An opened connection to the data source</param>
        ''' <param name="selectStatement">The name of a table on the database</param>
        ''' <param name="commandTimeout">A timeout for the command in seconds</param>
        ''' <returns>An DataManipulationResults object with the returned data</returns>
        ''' <remarks></remarks>
        Public Shared Function LoadQueryDataForManipulationViaCode(ByVal dataConnection As IDbConnection, ByVal selectStatement As String, ByVal commandTimeout As Integer) As CompuMaster.Data.DataManipulationResult
            'Prepare the command 
            Dim MyCmd As System.Data.IDbCommand = dataConnection.CreateCommand
            MyCmd.CommandText = selectStatement
            MyCmd.CommandType = CommandType.Text
            MyCmd.CommandTimeout = commandTimeout
            Return LoadDataForManipulationViaCode(dataConnection, MyCmd)
        End Function

        ''' <summary>
        ''' Query the data from the data connection in a mode for submitting changes in a later step
        ''' </summary>
        ''' <param name="command">A command with an opened connection to the data source</param>
        ''' <returns>An DataManipulationResults object with the returned data</returns>
        ''' <remarks></remarks>
        Public Shared Function LoadQueryDataForManipulationViaCode(ByVal command As IDbCommand) As CompuMaster.Data.DataManipulationResult
            Return LoadDataForManipulationViaCode(command.Connection, command)
        End Function


        ''' <summary>
        '''     Write back changes to the data connection
        ''' </summary>
        ''' <param name="container"></param>
        Public Shared Sub UpdateCodeManipulatedData(ByVal container As CompuMaster.Data.DataManipulationResult)
            UpdateCodeManipulatedData(container, True)
        End Sub

        ''' <summary>
        ''' An exception on data update
        ''' </summary>
        <CodeAnalysis.SuppressMessage("Design", "CA1032:Implement standard exception constructors", Justification:="<Ausstehend>")>
        <CodeAnalysis.SuppressMessage("Usage", "CA2237:Mark ISerializable types with serializable", Justification:="SqlDataAdapter possibly not Seriaizable (as is)")>
        Public Class SqlDataAdapterException
            Inherits Exception

            ''' <summary>
            ''' Re-encapsulate exception
            ''' </summary>
            ''' <param name="message"></param>
            ''' <param name="innerException"></param>
            Protected Friend Sub New(message As String, innerException As SqlDataAdapterException)
                Me.New(message, innerException.SqlDataAdapter, innerException.InnerException)
            End Sub

            Public Sub New(sqlDataAdapter As System.Data.IDbDataAdapter, innerException As System.Exception)
                Me.New("Data update failed", sqlDataAdapter, innerException)
            End Sub

            Public Sub New(message As String, sqlDataAdapter As System.Data.IDbDataAdapter, innerException As System.Exception)
                MyBase.New(message, innerException)
                Me.SqlDataAdapter = sqlDataAdapter
            End Sub

            ''' <summary>
            ''' The instance of SqlDataAdapter that should execute the update actions
            ''' </summary>
            ''' <returns></returns>
            Public Property SqlDataAdapter As System.Data.IDbDataAdapter

            ''' <summary>
            ''' Summary of command texts used by SqlDataAdapter
            ''' </summary>
            ''' <returns></returns>
            Public Function SqlDataAdapterCommands() As String
                Dim Result As String = ""
                Result &= "Command for SELECT: " & Me.SqlDataAdapter.SelectCommand.CommandText & System.Environment.NewLine
                Result &= "Command for INSERT: " & Me.SqlDataAdapter.InsertCommand.CommandText & System.Environment.NewLine
                Result &= "Command for UPDATE: " & Me.SqlDataAdapter.UpdateCommand.CommandText & System.Environment.NewLine
                Result &= "Command for DELETE: " & Me.SqlDataAdapter.DeleteCommand.CommandText
                Return Result
            End Function

        End Class

        ''' <summary>
        '''     Write back changes to the data connection
        ''' </summary>
        ''' <param name="container"></param>
        ''' <param name="useTransactionsIfAvailable"></param>
        ''' <remarks></remarks>
        Public Shared Sub UpdateCodeManipulatedData(ByVal container As CompuMaster.Data.DataManipulationResult, ByVal useTransactionsIfAvailable As Boolean)
            If CType(container.Command.Connection, Object).GetType.ToString = "System.Data.SqlClient.SqlConnection" Then
                Dim trans As System.Data.SqlClient.SqlTransaction = Nothing
                If useTransactionsIfAvailable = True Then trans = CType(container.Command.Connection.BeginTransaction(), System.Data.SqlClient.SqlTransaction)
                Try
                    'Assign current transaction to SELECT statement
                    CType(container.DataAdapter, SqlClient.SqlDataAdapter).SelectCommand.Transaction = trans
                    'Create missing update command statements
                    Dim sqlBuilder As New SqlClient.SqlCommandBuilder(CType(container.DataAdapter, SqlClient.SqlDataAdapter))
                    If trans IsNot Nothing Then
                        'ATTENTION: using manually created commands leads to not supported situation of columns with NOT NULL but with DEFAULT values
                        'result will be to trials of insertions of NULLs when there would be a default value 
                        'and which will lead to an exception when inserting
                        'BY: JW / 2010-12-23
                        If CType(container.DataAdapter, SqlClient.SqlDataAdapter).UpdateCommand Is Nothing Then
                            CType(container.DataAdapter, SqlClient.SqlDataAdapter).UpdateCommand = sqlBuilder.GetUpdateCommand
                        End If
                        If CType(container.DataAdapter, SqlClient.SqlDataAdapter).InsertCommand Is Nothing Then
                            CType(container.DataAdapter, SqlClient.SqlDataAdapter).InsertCommand = sqlBuilder.GetInsertCommand
                        End If
                        If CType(container.DataAdapter, SqlClient.SqlDataAdapter).DeleteCommand Is Nothing Then
                            CType(container.DataAdapter, SqlClient.SqlDataAdapter).DeleteCommand = sqlBuilder.GetDeleteCommand
                        End If
                        'Assign current transaction
                        CType(container.DataAdapter, SqlClient.SqlDataAdapter).UpdateCommand.Transaction = trans
                        CType(container.DataAdapter, SqlClient.SqlDataAdapter).DeleteCommand.Transaction = trans
                        CType(container.DataAdapter, SqlClient.SqlDataAdapter).InsertCommand.Transaction = trans
                    Else
                        'ATTENTION: provided container.DataAdapter.InsertCommand will be dropped, here!
                        'CAUSE: Do NOT provide the insert command because if the InsertCommand hasn't been provided,
                        'the dataAdapter.Update method will use customized, internal InsertCommands per each row 
                        'so that it supports no-NULLs-columns with default values 
                        '(otherwise if created manually, it always tries to insert a NULL value instead of just using the table column's default)
                        'ALSO SEE: 2 bottom posts on http://www.dotnetmonster.com/Uwe/Forum.aspx/dotnet-ado-net/4884/Inside-SqlCommandBuilder
                        'BY: JW / 2010-12-23
                        CType(container.DataAdapter, SqlClient.SqlDataAdapter).InsertCommand = Nothing
                    End If
                    'Update data
                    Try
                        CType(container.DataAdapter, SqlClient.SqlDataAdapter).Update(container.Table)
                    Catch ex As Exception
                        Throw New SqlDataAdapterException(container.DataAdapter, ex)
                    End Try
                    If trans IsNot Nothing Then
                        trans.Commit()
                        trans.Dispose()
                    End If
                Catch ex As Exception
                    If trans IsNot Nothing Then
                        trans.Rollback()
                        trans.Dispose()
                    End If
                    If ex.GetType Is GetType(SqlDataAdapterException) Then
                        Dim InnerEx As SqlDataAdapterException = CType(ex, SqlDataAdapterException)
                        Throw New SqlDataAdapterException("Error found - transaction has been rolled back", InnerEx)
                    Else
                        Throw New Exception("Error found - transaction has been rolled back", ex)
                    End If
                End Try
            Else
                Dim providers As System.Collections.Generic.List(Of Data.DataQuery.DataProvider)
                Dim CurrentProvider As Data.DataQuery.DataProvider = Nothing
                providers = Data.DataQuery.DataProvider.AvailableDataProviders()
                For MyCounter As Integer = 0 To providers.Count - 1
                    If providers(MyCounter).ConnectionType Is CType(container.Command.Connection, Object).GetType Then
                        CurrentProvider = providers(MyCounter)
                    End If
                Next
                If CurrentProvider IsNot Nothing Then
                    Try
                        CType(container.DataAdapter, System.Data.Common.DbDataAdapter).Update(container.Table)
                    Catch ex As Exception
                        Throw New SqlDataAdapterException(container.DataAdapter, ex)
                    End Try
                Else
                    Throw New NotSupportedException("Data provider not supported yet")
                End If
            End If
        End Sub

        ''' <summary>
        ''' Load a first row from the remote connection to receive list of columns
        ''' </summary>
        ''' <param name="tableName"></param>
        ''' <param name="dataConnection"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Shared Function LoadTableStructureWith1RowFromConnection(ByVal tableName As String, ByVal dataConnection As IDbConnection, ByVal ignoreExceptions As Boolean) As DataTable
            Dim OpenBrackets, CloseBrackets As String
            If tableName.IndexOf("[") >= 0 AndAlso tableName.IndexOf("]") >= 0 Then
                'table name already in a well-formed syntax, e.g. "dbo.[Test - 123]"
                OpenBrackets = Nothing
                CloseBrackets = Nothing
            Else
                'table name (e.g. "Test - 123") requires a well-formed syntax (e.g. [Test - 123])
                OpenBrackets = "["
                CloseBrackets = "]"
            End If
            If CType(dataConnection, Object).GetType.ToString = "Npgsql.NpgsqlConnection" Then
                If tableName.Contains(".""") = False Then
                    OpenBrackets = """"
                    CloseBrackets = """"
                Else
                    OpenBrackets = Nothing
                    CloseBrackets = Nothing
                End If
            End If

            Try
                Dim MyTable As DataTable
                Dim MyCmd As IDbCommand = dataConnection.CreateCommand()
                MyCmd.CommandText = "SELECT * FROM " & OpenBrackets & tableName & CloseBrackets
                MyCmd.CommandType = CommandType.Text
                MyCmd.Connection = dataConnection
                MyTable = CompuMaster.Data.DataQuery.FillDataTable(MyCmd, CompuMaster.Data.DataQuery.Automations.None, tableName)
                Return MyTable
            Catch ex As Exception
                If ignoreExceptions Then
                    Return Nothing
                Else
                    Throw New Exception("Error reading from table """ & tableName & """", ex)
                End If
            End Try
        End Function

    End Class

End Namespace
