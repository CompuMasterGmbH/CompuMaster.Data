Option Explicit On 
Option Strict On

Imports System.Data
Imports CompuMaster.Data.Information

Namespace CompuMaster.Data.DataQuery

    ''' <summary>
    '''     Querying data from all available types of data sources
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    Friend Class NamespaceDoc
        'UPDATE FOLLOWING LINE FOR EVERY CHANGE TO TRACK THE VERSION NUMBER INSIDE THIS DOCUMENT
        'Last change on V3.50 - 2009-06-25 JW
    End Class

    ''' <summary>
    '''     Common routines to query data from any data provider
    ''' </summary>
    <CodeAnalysis.SuppressMessage("Major Code Smell", "S1066:Collapsible ""if"" statements should be merged", Justification:="<Ausstehend>")>
    Public Module AnyIDataProvider

        ''' <summary>
        '''     Create a new database connection by reflection of a type name
        ''' </summary>
        ''' <param name="provider">A data provider</param>
        ''' <returns>The created connection object as an IDbConnection</returns>
        Public Function CreateConnection(provider As DataProvider) As IDbConnection
            Return provider.CreateConnection()
        End Function

        ''' <summary>
        '''     Create a new database connection by reflection of a type name
        ''' </summary>
        ''' <param name="provider">A data provider</param>
        ''' <param name="connectionString">A connection string to be used for this connection</param>
        ''' <returns>The created connection object as an IDbConnection</returns>
        Public Function CreateConnection(provider As DataProvider, ByVal connectionString As String) As IDbConnection
            Dim Result As IDbConnection = CreateConnection(provider)
            Result.ConnectionString = connectionString
            Return Result
        End Function

        ''' <summary>
        '''     Create a new database connection by reflection of a type name
        ''' </summary>
        ''' <param name="assemblyName">The assembly which implements the desired connection type</param>
        ''' <param name="connectionTypeName">The case-insensitive type name of the connection class, e. g. System.Data.SqlClient.SqlConnection</param>
        ''' <returns>The created connection object as an IDbConnection</returns>
        ''' <remarks>
        '''     Errors will be thrown in case of unresolvable parameter values or if the created type can't be casted into an IDbConnection.
        ''' </remarks>
        <CodeAnalysis.SuppressMessage("Major Code Smell", "S3385:""Exit"" statements should not be used", Justification:="<Ausstehend>")>
        Public Function CreateConnection(ByVal assemblyName As String, ByVal connectionTypeName As String) As IDbConnection
            Dim connectionType As Type = Nothing
            Dim runningAssembly As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly
            Dim referencedAssemblies As System.Reflection.AssemblyName() = runningAssembly.GetReferencedAssemblies

            For Each currentAssemblyName As System.Reflection.AssemblyName In referencedAssemblies
                If currentAssemblyName.Name = assemblyName Then
                    Dim referencedAssembly As System.Reflection.Assembly
                    referencedAssembly = System.Reflection.Assembly.Load(currentAssemblyName.FullName)
                    connectionType = referencedAssembly.GetType(connectionTypeName, True, True)
                    Exit For
                End If
            Next
            If connectionType IsNot Nothing Then
                Return CType(Activator.CreateInstance(connectionType), IDbConnection)
            Else
                Dim referencedAssembly As System.Reflection.Assembly
                referencedAssembly = System.Reflection.Assembly.LoadFile(assemblyName)
                Dim t As Type = referencedAssembly.GetType(connectionTypeName)
                If (t IsNot Nothing) Then
                    Return CType(Activator.CreateInstance(t), IDbConnection)
                    'Dim m As System.Runtime.Remoting.ObjectHandle = Activator.CreateInstanceFrom(assemblyName, connectionTypeName)
                    'If ((m) IsNot Nothing) Then
                    '    Return CType(m.Unwrap, IDbConnection)
                    'End If
                End If
            End If
            Throw New Exception("Class not found: " & assemblyName & "::" & connectionTypeName)
        End Function

        ''' <summary>
        '''     Create a new database connection by reflection of a type name
        ''' </summary>
        ''' <param name="assemblyName">The assembly which implements the desired connection type</param>
        ''' <param name="connectionTypeName">The case-insensitive type name of the connection class, e. g. System.Data.SqlClient.SqlConnection</param>
        ''' <param name="connectionString">A connection string to be used for this connection</param>
        ''' <returns>The created connection object as an IDbConnection</returns>
        ''' <remarks>
        '''     Errors will be thrown in case of unresolvable parameter values or if the created type can't be casted into an IDbConnection.
        ''' </remarks>
        Public Function CreateConnection(ByVal assemblyName As String, ByVal connectionTypeName As String, ByVal connectionString As String) As IDbConnection
            Dim Result As IDbConnection = CreateConnection(assemblyName, connectionTypeName)
            Result.ConnectionString = connectionString
            Return Result
        End Function

        ''' <summary>
        '''     Automations for the connection in charge
        ''' </summary>
        Public Enum Automations
            None = 0
            AutoOpenConnection = 1
            AutoCloseAndDisposeConnection = 2
            AutoOpenAndCloseAndDisposeConnection = 3
        End Enum

        ''' <summary>
        '''     Executes a command without returning any data
        ''' </summary>
        ''' <param name="dbConnection">The connection to the database</param>
        ''' <param name="commandText">The command text</param>
        ''' <param name="commandType">The command type</param>
        ''' <param name="sqlParameters">An optional list of SqlParameters</param>
        ''' <param name="automations">Automation options for the connection</param>
        ''' <param name="commandTimeout">A timeout value in seconds for the command object (negative values will be ignored and leave the timeout value on default)</param>
        Public Sub ExecuteNonQuery(ByVal dbConnection As IDbConnection, ByVal commandText As String, ByVal commandType As System.Data.CommandType, ByVal sqlParameters As IDataParameter(), ByVal automations As Automations, ByVal commandTimeout As Integer)
            Dim MyConn As IDbConnection = dbConnection
            Dim MyCmd As IDbCommand = MyConn.CreateCommand
            MyCmd.CommandText = commandText
            MyCmd.CommandType = commandType
            If commandTimeout >= 0 Then
                MyCmd.CommandTimeout = commandTimeout
            End If
            If sqlParameters IsNot Nothing Then
                For Each MySqlParam As IDataParameter In sqlParameters
                    MyCmd.Parameters.Add(MySqlParam)
                Next
            End If
            Dim Result As Object
            Try
                If automations = Automations.AutoOpenAndCloseAndDisposeConnection OrElse automations = Automations.AutoOpenConnection Then
                    If MyConn IsNot Nothing AndAlso MyConn.State <> ConnectionState.Open Then
                        MyConn.Open()
                    End If
                End If
                Result = MyCmd.ExecuteNonQuery
                If MyCmd IsNot Nothing Then
                    MyCmd.Dispose()
                End If
            Catch ex As Exception
                Throw New CompuMaster.Data.DataQuery.DataException(MyCmd, ex)
            Finally
                If automations = Automations.AutoOpenAndCloseAndDisposeConnection OrElse automations = Automations.AutoCloseAndDisposeConnection Then
                    If MyConn IsNot Nothing Then
                        If MyConn.State <> ConnectionState.Closed Then
                            MyConn.Close()
                        End If
                        MyConn.Dispose()
                    End If
                End If
            End Try
        End Sub

        ''' <summary>
        '''     Executes a command without returning any data
        ''' </summary>
        ''' <param name="dbCommand">The command with an assigned connection property value</param>
        ''' <param name="automations">Automation options for the connection</param>
        Public Sub ExecuteNonQuery(ByVal dbCommand As IDbCommand, ByVal automations As Automations)
            Dim MyCmd As IDbCommand = dbCommand
            Dim MyConn As IDbConnection = MyCmd.Connection
            Dim Result As Integer
            Try
                If automations = Automations.AutoOpenAndCloseAndDisposeConnection OrElse automations = Automations.AutoOpenConnection Then
                    If MyConn IsNot Nothing AndAlso MyConn.State <> ConnectionState.Open Then
                        MyConn.Open()
                    End If
                End If
                Result = MyCmd.ExecuteNonQuery
            Catch ex As Exception
                Throw New CompuMaster.Data.DataQuery.DataException(MyCmd, ex)
            Finally
                If automations = Automations.AutoOpenAndCloseAndDisposeConnection OrElse automations = Automations.AutoCloseAndDisposeConnection Then
                    If MyConn IsNot Nothing Then
                        If MyConn.State <> ConnectionState.Closed Then
                            MyConn.Close()
                        End If
                        MyConn.Dispose()
                    End If
                End If
            End Try
        End Sub

        ''' <summary>
        '''     Executes a command without returning any data
        ''' </summary>
        ''' <param name="dbConnection">The connection to the database</param>
        ''' <param name="commandText">The command text</param>
        ''' <param name="commandType">The command type</param>
        ''' <param name="sqlParameters">An optional list of SqlParameters</param>
        ''' <param name="automations">Automation options for the connection</param>
        Public Sub ExecuteNonQuery(ByVal dbConnection As IDbConnection, ByVal commandText As String, ByVal commandType As System.Data.CommandType, ByVal sqlParameters As IDataParameter(), ByVal automations As Automations)
            ExecuteNonQuery(dbConnection, commandText, commandType, sqlParameters, automations, -1)
        End Sub

        ''' <summary>
        '''     Executes a command without returning any data
        ''' </summary>
        ''' <param name="dbConnection">The connection to the database</param>
        ''' <param name="commandText">The command text</param>
        ''' <param name="commandType">The command type</param>
        ''' <param name="sqlParameters">An optional list of SqlParameters</param>
        Public Sub ExecuteNonQuery(ByVal dbConnection As IDbConnection, ByVal commandText As String, ByVal commandType As System.Data.CommandType, ByVal sqlParameters As IDataParameter())
            ExecuteNonQuery(dbConnection, commandText, commandType, sqlParameters, Automations.AutoOpenAndCloseAndDisposeConnection)
        End Sub

        ''' <summary>
        '''     Executes a command scalar and returns the value
        ''' </summary>
        ''' <param name="dbConnection">The connection to the database</param>
        ''' <param name="commandText">The command text</param>
        ''' <param name="commandType">The command type</param>
        ''' <param name="sqlParameters">An optional list of SqlParameters</param>
        ''' <param name="automations">Automation options for the connection</param>
        ''' <returns></returns>
        Public Function ExecuteScalar(ByVal dbConnection As IDbConnection, ByVal commandText As String, ByVal commandType As System.Data.CommandType, ByVal sqlParameters As IDataParameter(), ByVal automations As Automations) As Object
            Dim MyConn As IDbConnection = dbConnection
            Dim MyCmd As IDbCommand = MyConn.CreateCommand
            MyCmd.CommandText = commandText
            MyCmd.CommandType = commandType
            If sqlParameters IsNot Nothing Then
                For Each MySqlParam As IDataParameter In sqlParameters
                    MyCmd.Parameters.Add(MySqlParam)
                Next
            End If
            Dim Result As Object
            Try
                If automations = Automations.AutoOpenAndCloseAndDisposeConnection OrElse automations = Automations.AutoOpenConnection Then
                    If MyConn IsNot Nothing AndAlso MyConn.State <> ConnectionState.Open Then
                        MyConn.Open()
                    End If
                End If
                Result = MyCmd.ExecuteScalar
                If MyCmd IsNot Nothing Then
                    MyCmd.Dispose()
                End If
            Catch ex As Exception
                Throw New CompuMaster.Data.DataQuery.DataException(MyCmd, ex)
            Finally
                If automations = Automations.AutoOpenAndCloseAndDisposeConnection OrElse automations = Automations.AutoCloseAndDisposeConnection Then
                    If MyConn IsNot Nothing Then
                        If MyConn.State <> ConnectionState.Closed Then
                            MyConn.Close()
                        End If
                        MyConn.Dispose()
                    End If
                End If
            End Try
            Return Result
        End Function

        ''' <summary>
        '''     Executes a command scalar and returns the value
        ''' </summary>
        ''' <param name="dbCommand">The command with an assigned connection property value</param>
        ''' <param name="automations">Automation options for the connection</param>
        Public Function ExecuteScalar(ByVal dbCommand As IDbCommand, ByVal automations As Automations) As Object
            Dim MyCmd As IDbCommand = dbCommand
            Dim MyConn As IDbConnection = MyCmd.Connection
            Dim Result As Object
            Try
                If automations = Automations.AutoOpenAndCloseAndDisposeConnection OrElse automations = Automations.AutoOpenConnection Then
                    If MyCmd.Connection.State <> ConnectionState.Open Then
                        MyConn.Open()
                    End If
                End If
                Result = MyCmd.ExecuteScalar
            Catch ex As Exception
                Throw New CompuMaster.Data.DataQuery.DataException(MyCmd, ex)
            Finally
                If automations = Automations.AutoOpenAndCloseAndDisposeConnection OrElse automations = Automations.AutoCloseAndDisposeConnection Then
                    If MyConn IsNot Nothing Then
                        If MyConn.State <> ConnectionState.Closed Then
                            MyConn.Close()
                        End If
                        MyConn.Dispose()
                    End If
                End If
            End Try
            Return Result
        End Function

        ''' <summary>
        '''     Executes a command scalar and returns the value
        ''' </summary>
        ''' <param name="dbConnection">The connection to the database</param>
        ''' <param name="commandText">The command text</param>
        ''' <param name="commandType">The command type</param>
        ''' <param name="sqlParameters">An optional list of SqlParameters</param>
        ''' <returns></returns>
        Public Function ExecuteScalar(ByVal dbConnection As IDbConnection, ByVal commandText As String, ByVal commandType As System.Data.CommandType, ByVal sqlParameters As IDataParameter()) As Object
            Return ExecuteScalar(dbConnection, commandText, commandType, sqlParameters, Automations.AutoOpenAndCloseAndDisposeConnection)
        End Function

        ''' <summary>
        '''     Executes a command with a data reader and returns the values of the first column
        ''' </summary>
        ''' <param name="dbConnection">The connection to the database</param>
        ''' <param name="commandText">The command text</param>
        ''' <param name="commandType">The command type</param>
        ''' <param name="sqlParameters">An optional list of SqlParameters</param>
        ''' <param name="automations">Automation options for the connection</param>
        ''' <returns></returns>
        Public Function ExecuteReaderAndPutFirstColumnIntoArrayList(ByVal dbConnection As IDbConnection, ByVal commandText As String, ByVal commandType As System.Data.CommandType, ByVal sqlParameters As IDataParameter(), ByVal automations As Automations) As ArrayList
            Dim MyConn As IDbConnection = dbConnection
            Dim MyCmd As IDbCommand = MyConn.CreateCommand
            MyCmd.CommandText = commandText
            MyCmd.CommandType = commandType
            If sqlParameters IsNot Nothing Then
                For Each MySqlParam As IDataParameter In sqlParameters
                    MyCmd.Parameters.Add(MySqlParam)
                Next
            End If
            Return ExecuteReaderAndPutFirstColumnIntoArrayList(MyCmd, automations)
        End Function

        ''' <summary>
        '''     Executes a command with a data reader and returns the values of the first column
        ''' </summary>
        ''' <param name="dbCommand">The command object which shall be executed</param>
        ''' <param name="automations">Automation options for the connection</param>
        ''' <returns></returns>
        Public Function ExecuteReaderAndPutFirstColumnIntoArrayList(ByVal dbCommand As IDbCommand, ByVal automations As Automations) As ArrayList
            Dim MyConn As IDbConnection = dbCommand.Connection
            Dim MyCmd As IDbCommand = dbCommand
            Dim MyReader As IDataReader = Nothing
            Dim Result As New ArrayList
            Try
                If automations = Automations.AutoOpenAndCloseAndDisposeConnection OrElse automations = Automations.AutoOpenConnection Then
                    If MyConn IsNot Nothing AndAlso MyConn.State <> ConnectionState.Open Then
                        MyConn.Open()
                    End If
                End If
                MyReader = MyCmd.ExecuteReader
                While MyReader.Read
                    Result.Add(MyReader(0))
                End While
            Catch ex As Exception
                Throw New CompuMaster.Data.DataQuery.DataException(MyCmd, ex)
            Finally
                If MyReader IsNot Nothing AndAlso Not MyReader.IsClosed Then
                    MyReader.Close()
                End If
                If automations = Automations.AutoOpenAndCloseAndDisposeConnection OrElse automations = Automations.AutoCloseAndDisposeConnection Then
                    If MyConn IsNot Nothing Then
                        If MyConn.State <> ConnectionState.Closed Then
                            MyConn.Close()
                        End If
                        MyConn.Dispose()
                    End If
                End If
            End Try
            Return Result
        End Function

        Public Function ExecuteReaderAndPutFirstColumnIntoGenericList(Of TValue)(ByVal dbCommand As IDbCommand, ByVal automations As Automations) As System.Collections.Generic.List(Of TValue)
            Dim MyConn As IDbConnection = dbCommand.Connection
            Dim MyCmd As IDbCommand = dbCommand
            Dim MyReader As IDataReader = Nothing
            Dim Result As New System.Collections.Generic.List(Of TValue)
            Try
                If automations = Automations.AutoOpenAndCloseAndDisposeConnection OrElse automations = Automations.AutoOpenConnection Then
                    If MyConn IsNot Nothing AndAlso MyConn.State <> ConnectionState.Open Then
                        MyConn.Open()
                    End If
                End If
                MyReader = MyCmd.ExecuteReader
                While MyReader.Read
                    If IsDBNull(MyReader(0)) Then
                        Result.Add(Nothing)
                    Else
                        Result.Add(CType(MyReader(0), TValue))
                    End If
                End While
            Catch ex As Exception
                Throw New CompuMaster.Data.DataQuery.DataException(MyCmd, ex)
            Finally
                If MyReader IsNot Nothing AndAlso Not MyReader.IsClosed Then
                    MyReader.Close()
                End If
                If automations = Automations.AutoOpenAndCloseAndDisposeConnection OrElse automations = Automations.AutoCloseAndDisposeConnection Then
                    If MyConn IsNot Nothing Then
                        If MyConn.State <> ConnectionState.Closed Then
                            MyConn.Close()
                        End If
                        MyConn.Dispose()
                    End If
                End If
            End Try
            Return Result
        End Function

        Public Function ExecuteReaderAndPutFirstColumnIntoGenericNullableList(Of TValue As Structure)(ByVal dbCommand As IDbCommand, ByVal automations As Automations) As System.Collections.Generic.List(Of TValue?)
            Dim MyConn As IDbConnection = dbCommand.Connection
            Dim MyCmd As IDbCommand = dbCommand
            Dim MyReader As IDataReader = Nothing
            Dim Result As New System.Collections.Generic.List(Of TValue?)
            Try
                If automations = Automations.AutoOpenAndCloseAndDisposeConnection OrElse automations = Automations.AutoOpenConnection Then
                    If MyConn IsNot Nothing AndAlso MyConn.State <> ConnectionState.Open Then
                        MyConn.Open()
                    End If
                End If
                MyReader = MyCmd.ExecuteReader
                While MyReader.Read
                    If IsDBNull(MyReader(0)) Then
                        Result.Add(New TValue?()) 'Empty T --> .HasValue = False
                    Else
                        Result.Add(New TValue?(CType(MyReader(0), TValue)))
                    End If
                End While
            Catch ex As Exception
                Throw New CompuMaster.Data.DataQuery.DataException(MyCmd, ex)
            Finally
                If MyReader IsNot Nothing AndAlso Not MyReader.IsClosed Then
                    MyReader.Close()
                End If
                If automations = Automations.AutoOpenAndCloseAndDisposeConnection OrElse automations = Automations.AutoCloseAndDisposeConnection Then
                    If MyConn IsNot Nothing Then
                        If MyConn.State <> ConnectionState.Closed Then
                            MyConn.Close()
                        End If
                        MyConn.Dispose()
                    End If
                End If
            End Try
            Return Result
        End Function

        ''' <summary>
        '''     Executes a command with a data reader and returns the values of the first column
        ''' </summary>
        ''' <param name="dbConnection">The connection to the database</param>
        ''' <param name="commandText">The command text</param>
        ''' <param name="commandType">The command type</param>
        ''' <param name="sqlParameters">An optional list of SqlParameters</param>
        ''' <returns></returns>
        Public Function ExecuteReaderAndPutFirstColumnIntoArrayList(ByVal dbConnection As IDbConnection, ByVal commandText As String, ByVal commandType As System.Data.CommandType, ByVal sqlParameters As IDataParameter()) As ArrayList
            Return ExecuteReaderAndPutFirstColumnIntoArrayList(dbConnection, commandText, commandType, sqlParameters, Automations.AutoOpenAndCloseAndDisposeConnection)
        End Function

        ''' <summary>
        '''     Executes a command with a data reader and returns the values of the first two columns
        ''' </summary>
        ''' <param name="dbCommand">The prepared command to the database</param>
        ''' <param name="automations">Automation options for the connection</param>
        ''' <returns>An array of DictionaryEntry with the values of the first column as the key element and the second column values in the value element of the DictionaryEntry</returns>
        Public Function ExecuteReaderAndPutFirstTwoColumnsIntoDictionaryEntryArray(ByVal dbCommand As IDbCommand, ByVal automations As Automations) As DictionaryEntry()
            Dim Result As New ArrayList
            Dim MyCmd As IDbCommand = dbCommand
            Dim MyReader As IDataReader = Nothing
            Try
                If automations = Automations.AutoOpenAndCloseAndDisposeConnection OrElse automations = Automations.AutoOpenConnection Then
                    If MyCmd.Connection.State <> ConnectionState.Open Then
                        MyCmd.Connection.Open()
                    End If
                End If
                MyReader = MyCmd.ExecuteReader
                While MyReader.Read
                    Result.Add(New DictionaryEntry(MyReader(0), MyReader(1)))
                End While
            Catch ex As Exception
                Throw New CompuMaster.Data.DataQuery.DataException(MyCmd, ex)
            Finally
                If MyReader IsNot Nothing AndAlso Not MyReader.IsClosed Then
                    MyReader.Close()
                End If
                If automations = Automations.AutoOpenAndCloseAndDisposeConnection OrElse automations = Automations.AutoCloseAndDisposeConnection Then
                    If MyCmd.Connection IsNot Nothing Then
                        If MyCmd.Connection.State <> ConnectionState.Closed Then
                            MyCmd.Connection.Close()
                        End If
                        MyCmd.Connection.Dispose()
                    End If
                End If
            End Try
            Return CType(Result.ToArray(GetType(DictionaryEntry)), DictionaryEntry())
        End Function

        ''' <summary>
        '''     Executes a command with a data reader and returns the values of the first two columns
        ''' </summary>
        ''' <param name="dbConnection">The connection to the database</param>
        ''' <param name="commandText">The command text</param>
        ''' <param name="commandType">The command type</param>
        ''' <param name="sqlParameters">An optional list of SqlParameters</param>
        ''' <param name="automations">Automation options for the connection</param>
        ''' <returns>An array of DictionaryEntry with the values of the first column as the key element and the second column values in the value element of the DictionaryEntry</returns>
        Public Function ExecuteReaderAndPutFirstTwoColumnsIntoDictionaryEntryArray(ByVal dbConnection As IDbConnection, ByVal commandText As String, ByVal commandType As System.Data.CommandType, ByVal sqlParameters As IDataParameter(), ByVal automations As Automations) As DictionaryEntry()
            Dim Result As New ArrayList
            Dim MyConn As IDbConnection = dbConnection
            Dim MyCmd As IDbCommand = MyConn.CreateCommand
            MyCmd.CommandText = commandText
            MyCmd.CommandType = commandType
            If sqlParameters IsNot Nothing Then
                For Each MySqlParam As IDataParameter In sqlParameters
                    MyCmd.Parameters.Add(MySqlParam)
                Next
            End If
            Dim MyReader As IDataReader = Nothing
            Try
                If automations = Automations.AutoOpenAndCloseAndDisposeConnection OrElse automations = Automations.AutoOpenConnection Then
                    If MyConn IsNot Nothing AndAlso MyConn.State <> ConnectionState.Open Then
                        MyConn.Open()
                    End If
                End If
                MyReader = MyCmd.ExecuteReader
                While MyReader.Read
                    Result.Add(New DictionaryEntry(MyReader(0), MyReader(1)))
                End While
            Catch ex As Exception
                Throw New CompuMaster.Data.DataQuery.DataException(MyCmd, ex)
            Finally
                If MyReader IsNot Nothing AndAlso Not MyReader.IsClosed Then
                    MyReader.Close()
                End If
                If automations = Automations.AutoOpenAndCloseAndDisposeConnection OrElse automations = Automations.AutoCloseAndDisposeConnection Then
                    If MyConn IsNot Nothing Then
                        If MyConn.State <> ConnectionState.Closed Then
                            MyConn.Close()
                        End If
                        MyConn.Dispose()
                    End If
                End If
            End Try
            Return CType(Result.ToArray(GetType(DictionaryEntry)), DictionaryEntry())
        End Function

        ''' <summary>
        '''     Executes a command with a data reader and returns the values of the first column
        ''' </summary>
        ''' <param name="dbCommand">The command object which shall be executed</param>
        ''' <param name="automations">Automation options for the connection</param>
        ''' <returns>A hashtable with the values of the first column in the hashtable's key field and the second column values in the hashtable's value field</returns>
        ''' <remarks>
        ''' ATTENTION: Please note that multiple but equal values from the first column will result in 1 key/value pair since hashtables use a unique key and override the value with the last assignment. Alternatively you may want to receive an array of DictionaryEntry.
        ''' </remarks>
        Public Function ExecuteReaderAndPutFirstTwoColumnsIntoHashtable(ByVal dbCommand As IDbCommand, ByVal automations As Automations) As Hashtable
            Dim Result As New Hashtable
            Dim MyCmd As IDbCommand = dbCommand
            Dim MyReader As IDataReader = Nothing
            Try
                If automations = Automations.AutoOpenAndCloseAndDisposeConnection OrElse automations = Automations.AutoOpenConnection Then
                    If MyCmd.Connection.State <> ConnectionState.Open Then
                        MyCmd.Connection.Open()
                    End If
                End If
                MyReader = MyCmd.ExecuteReader
                While MyReader.Read
                    Result.Add(MyReader(0), MyReader(1))
                End While
            Catch ex As Exception
                Throw New CompuMaster.Data.DataQuery.DataException(MyCmd, ex)
            Finally
                If MyReader IsNot Nothing AndAlso Not MyReader.IsClosed Then
                    MyReader.Close()
                End If
                If automations = Automations.AutoOpenAndCloseAndDisposeConnection OrElse automations = Automations.AutoCloseAndDisposeConnection Then
                    If MyCmd.Connection IsNot Nothing Then
                        If MyCmd.Connection.State <> ConnectionState.Closed Then
                            MyCmd.Connection.Close()
                        End If
                        MyCmd.Connection.Dispose()
                    End If
                End If
            End Try
            Return Result
        End Function

        ''' <summary>
        '''     Executes a command with a data reader and returns the values of the first two columns
        ''' </summary>
        ''' <param name="dbConnection">The connection to the database</param>
        ''' <param name="commandText">The command text</param>
        ''' <param name="commandType">The command type</param>
        ''' <param name="sqlParameters">An optional list of SqlParameters</param>
        ''' <param name="automations">Automation options for the connection</param>
        ''' <returns>A hashtable with the values of the first column in the hashtable's key field and the second column values in the hashtable's value field</returns>
        ''' <remarks>
        ''' ATTENTION: Please note that multiple but equal values from the first column will result in 1 key/value pair since hashtables use a unique key and override the value with the last assignment. Alternatively you may want to receive an array of DictionaryEntry.
        ''' </remarks>
        Public Function ExecuteReaderAndPutFirstTwoColumnsIntoHashtable(ByVal dbConnection As IDbConnection, ByVal commandText As String, ByVal commandType As System.Data.CommandType, ByVal sqlParameters As IDataParameter(), ByVal automations As Automations) As Hashtable
            Dim MyConn As IDbConnection = dbConnection
            Dim MyCmd As IDbCommand = MyConn.CreateCommand
            MyCmd.CommandText = commandText
            MyCmd.CommandType = commandType
            If sqlParameters IsNot Nothing Then
                For Each MySqlParam As IDataParameter In sqlParameters
                    MyCmd.Parameters.Add(MySqlParam)
                Next
            End If
            Dim MyReader As IDataReader = Nothing
            Dim Result As New Hashtable
            Try
                If automations = Automations.AutoOpenAndCloseAndDisposeConnection OrElse automations = Automations.AutoOpenConnection Then
                    If MyConn IsNot Nothing AndAlso MyConn.State <> ConnectionState.Open Then
                        MyConn.Open()
                    End If
                End If
                MyReader = MyCmd.ExecuteReader
                While MyReader.Read
                    Result.Add(MyReader(0), MyReader(1))
                End While
            Catch ex As Exception
                Throw New CompuMaster.Data.DataQuery.DataException(MyCmd, ex)
            Finally
                If MyReader IsNot Nothing AndAlso Not MyReader.IsClosed Then
                    MyReader.Close()
                End If
                If automations = Automations.AutoOpenAndCloseAndDisposeConnection OrElse automations = Automations.AutoCloseAndDisposeConnection Then
                    If MyConn IsNot Nothing Then
                        If MyConn.State <> ConnectionState.Closed Then
                            MyConn.Close()
                        End If
                        MyConn.Dispose()
                    End If
                End If
            End Try
            Return Result
        End Function

        ''' <summary>
        '''     Executes a command with a data reader and returns the values of the first two columns
        ''' </summary>
        ''' <param name="dbCommand">The command object which shall be executed</param>
        ''' <param name="automations">Automation options for the connection</param>
        ''' <returns>A list of KeyValuePairs with the values of the first column in the key field and the second column values in the value field, NULL values are initialized with null (Nothing in VisualBasic)</returns>
        Public Function ExecuteReaderAndPutFirstTwoColumnsIntoGenericKeyValuePairs(Of TKey, TValue)(ByVal dbCommand As IDbCommand, ByVal automations As Automations) As System.Collections.Generic.List(Of System.Collections.Generic.KeyValuePair(Of TKey, TValue))
            Dim Result As New System.Collections.Generic.List(Of System.Collections.Generic.KeyValuePair(Of TKey, TValue))
            Dim MyCmd As IDbCommand = dbCommand
            Dim MyReader As IDataReader = Nothing
            Try
                If automations = Automations.AutoOpenAndCloseAndDisposeConnection OrElse automations = Automations.AutoOpenConnection Then
                    If MyCmd.Connection.State <> ConnectionState.Open Then
                        MyCmd.Connection.Open()
                    End If
                End If
                MyReader = MyCmd.ExecuteReader
                While MyReader.Read
                    Dim key As TKey, value As TValue
                    If IsDBNull(MyReader(0)) Then
                        key = Nothing
                    Else
                        key = CType(MyReader(0), TKey)
                    End If
                    If IsDBNull(MyReader(1)) Then
                        value = Nothing
                    Else
                        value = CType(MyReader(1), TValue)
                    End If
                    Result.Add(New System.Collections.Generic.KeyValuePair(Of TKey, TValue)(key, value))
                End While
            Catch ex As Exception
                Throw New CompuMaster.Data.DataQuery.DataException(MyCmd, ex)
            Finally
                If MyReader IsNot Nothing AndAlso Not MyReader.IsClosed Then
                    MyReader.Close()
                End If
                If automations = Automations.AutoOpenAndCloseAndDisposeConnection OrElse automations = Automations.AutoCloseAndDisposeConnection Then
                    If MyCmd.Connection IsNot Nothing Then
                        If MyCmd.Connection.State <> ConnectionState.Closed Then
                            MyCmd.Connection.Close()
                        End If
                        MyCmd.Connection.Dispose()
                    End If
                End If
            End Try
            Return Result
        End Function

        ''' <summary>
        '''     Executes a command with a data reader and returns the values of the first two columns
        ''' </summary>
        ''' <param name="dbCommand">The command object which shall be executed</param>
        ''' <param name="automations">Automation options for the connection</param>
        ''' <returns>A list of KeyValuePairs with the values of the first column in the key field and the second column values in the value field</returns>
        Public Function ExecuteReaderAndPutFirstTwoColumnsIntoGenericNullableKeyValuePairs(Of TKey As Structure, TValue As Structure)(ByVal dbCommand As IDbCommand, ByVal automations As Automations) As System.Collections.Generic.List(Of System.Collections.Generic.KeyValuePair(Of TKey?, TValue?))
            Dim Result As New System.Collections.Generic.List(Of System.Collections.Generic.KeyValuePair(Of TKey?, TValue?))
            Dim MyCmd As IDbCommand = dbCommand
            Dim MyReader As IDataReader = Nothing
            Try
                If automations = Automations.AutoOpenAndCloseAndDisposeConnection OrElse automations = Automations.AutoOpenConnection Then
                    If MyCmd.Connection.State <> ConnectionState.Open Then
                        MyCmd.Connection.Open()
                    End If
                End If
                MyReader = MyCmd.ExecuteReader
                While MyReader.Read
                    Dim key As TKey?, value As TValue?
                    If IsDBNull(MyReader(0)) Then
                        key = New TKey?
                    Else
                        key = New TKey?(CType(MyReader(0), TKey))
                    End If
                    If IsDBNull(MyReader(1)) Then
                        value = New TValue?
                    Else
                        value = New TValue?(CType(MyReader(1), TValue))
                    End If
                    Result.Add(New System.Collections.Generic.KeyValuePair(Of TKey?, TValue?)(key, value))
                End While
            Catch ex As Exception
                Throw New CompuMaster.Data.DataQuery.DataException(MyCmd, ex)
            Finally
                If MyReader IsNot Nothing AndAlso Not MyReader.IsClosed Then
                    MyReader.Close()
                End If
                If automations = Automations.AutoOpenAndCloseAndDisposeConnection OrElse automations = Automations.AutoCloseAndDisposeConnection Then
                    If MyCmd.Connection IsNot Nothing Then
                        If MyCmd.Connection.State <> ConnectionState.Closed Then
                            MyCmd.Connection.Close()
                        End If
                        MyCmd.Connection.Dispose()
                    End If
                End If
            End Try
            Return Result
        End Function

        ''' <summary>
        '''     Executes a command with a data reader and returns the values of the first two columns
        ''' </summary>
        ''' <param name="dbCommand">The command object which shall be executed</param>
        ''' <param name="automations">Automation options for the connection</param>
        ''' <returns>A dictionary of KeyValuePairs with the values of the first column in the key field and the second column values in the value field, NULL values are initialized with null (Nothing in VisualBasic)</returns>
        ''' <remarks>
        ''' ATTENTION: Please note that multiple but equal values from the first column will result in 1 key/value pair since hashtables use a unique key and override the value with the last assignment. Alternatively you may want to receive a List of KeyValuePairs.
        ''' </remarks>
        Public Function ExecuteReaderAndPutFirstTwoColumnsIntoGenericDictionary(Of TKey, TValue)(ByVal dbCommand As IDbCommand, ByVal automations As Automations) As System.Collections.Generic.Dictionary(Of TKey, TValue)
            Dim Result As New System.Collections.Generic.Dictionary(Of TKey, TValue)
            Dim MyCmd As IDbCommand = dbCommand
            Dim MyReader As IDataReader = Nothing
            Try
                If automations = Automations.AutoOpenAndCloseAndDisposeConnection OrElse automations = Automations.AutoOpenConnection Then
                    If MyCmd.Connection.State <> ConnectionState.Open Then
                        MyCmd.Connection.Open()
                    End If
                End If
                MyReader = MyCmd.ExecuteReader
                While MyReader.Read
                    Dim key As TKey, value As TValue
                    If IsDBNull(MyReader(0)) Then
                        key = Nothing
                    Else
                        key = CType(MyReader(0), TKey)
                    End If
                    If IsDBNull(MyReader(1)) Then
                        value = Nothing
                    Else
                        value = CType(MyReader(1), TValue)
                    End If
                    If Result.ContainsKey(key) Then
                        Result(key) = value
                    Else
                        Result.Add(key, value)
                    End If
                End While
            Catch ex As Exception
                Throw New CompuMaster.Data.DataQuery.DataException(MyCmd, ex)
            Finally
                If MyReader IsNot Nothing AndAlso Not MyReader.IsClosed Then
                    MyReader.Close()
                End If
                If automations = Automations.AutoOpenAndCloseAndDisposeConnection OrElse automations = Automations.AutoCloseAndDisposeConnection Then
                    If MyCmd.Connection IsNot Nothing Then
                        If MyCmd.Connection.State <> ConnectionState.Closed Then
                            MyCmd.Connection.Close()
                        End If
                        MyCmd.Connection.Dispose()
                    End If
                End If
            End Try
            Return Result
        End Function

        ''' <summary>
        '''     Executes a command with a data reader and returns the values of the first two columns
        ''' </summary>
        ''' <param name="dbCommand">The command object which shall be executed</param>
        ''' <param name="automations">Automation options for the connection</param>
        ''' <returns>A dictionary of KeyValuePairs with the values of the first column in the key field and the second column values in the value field</returns>
        ''' <remarks>
        ''' ATTENTION: Please note that multiple but equal values from the first column will result in 1 key/value pair since hashtables use a unique key and override the value with the last assignment. Alternatively you may want to receive a List of KeyValuePairs.
        ''' </remarks>
        Public Function ExecuteReaderAndPutFirstTwoColumnsIntoGenericNullableDictionary(Of TKey, TValue As Structure)(ByVal dbCommand As IDbCommand, ByVal automations As Automations) As System.Collections.Generic.Dictionary(Of TKey, TValue?)
            Dim Result As New System.Collections.Generic.Dictionary(Of TKey, TValue?)
            Dim MyCmd As IDbCommand = dbCommand
            Dim MyReader As IDataReader = Nothing
            Try
                If automations = Automations.AutoOpenAndCloseAndDisposeConnection OrElse automations = Automations.AutoOpenConnection Then
                    If MyCmd.Connection.State <> ConnectionState.Open Then
                        MyCmd.Connection.Open()
                    End If
                End If
                MyReader = MyCmd.ExecuteReader
                While MyReader.Read
                    Dim key As TKey, value As TValue?
                    If IsDBNull(MyReader(0)) Then
                        key = Nothing
                    Else
                        key = CType(MyReader(0), TKey)
                    End If
                    If IsDBNull(MyReader(1)) Then
                        value = New TValue?
                    Else
                        value = New TValue?(CType(MyReader(1), TValue))
                    End If
                    If Result.ContainsKey(key) Then
                        Result(key) = value
                    Else
                        Result.Add(key, value)
                    End If
                End While
            Catch ex As Exception
                Throw New CompuMaster.Data.DataQuery.DataException(MyCmd, ex)
            Finally
                If MyReader IsNot Nothing AndAlso Not MyReader.IsClosed Then
                    MyReader.Close()
                End If
                If automations = Automations.AutoOpenAndCloseAndDisposeConnection OrElse automations = Automations.AutoCloseAndDisposeConnection Then
                    If MyCmd.Connection IsNot Nothing Then
                        If MyCmd.Connection.State <> ConnectionState.Closed Then
                            MyCmd.Connection.Close()
                        End If
                        MyCmd.Connection.Dispose()
                    End If
                End If
            End Try
            Return Result
        End Function

        ''' <summary>
        '''     Executes a command with a data reader and returns the values of the first two columns
        ''' </summary>
        ''' <param name="dbConnection">The connection to the database</param>
        ''' <param name="commandText">The command text</param>
        ''' <param name="commandType">The command type</param>
        ''' <param name="sqlParameters">An optional list of SqlParameters</param>
        ''' <returns>A hashtable with the values of the first column in the hashtable's key field and the second column values in the hashtable's value field</returns>
        ''' <remarks>
        ''' ATTENTION: Please note that multiple but equal values from the first column will result in 1 key/value pair since hashtables use a unique key and override the value with the last assignment. Alternatively you may want to receive an array of DictionaryEntry.
        ''' </remarks>
        Public Function ExecuteReaderAndPutFirstTwoColumnsIntoHashtable(ByVal dbConnection As IDbConnection, ByVal commandText As String, ByVal commandType As System.Data.CommandType, ByVal sqlParameters As IDataParameter()) As Hashtable
            Return ExecuteReaderAndPutFirstTwoColumnsIntoHashtable(dbConnection, commandText, commandType, sqlParameters, Automations.AutoOpenAndCloseAndDisposeConnection)
        End Function

        ''' <summary>
        '''     Executes a command and return the data reader object for it
        ''' </summary>
        ''' <param name="dbConnection">The connection to the database</param>
        ''' <param name="commandText">The command text</param>
        ''' <param name="commandType">The command type</param>
        ''' <param name="sqlParameters">An optional list of SqlParameters</param>
        ''' <param name="automations">Automation options for the connection</param>
        ''' <param name="commandTimeout">A timeout value in seconds for the command object (negative values will be ignored and leave the timeout value on default)</param>
        ''' <returns></returns>
        ''' <remarks>
        '''     Automations can only open a connection, but never close. This is because you have to close the connection by yourself AFTER you walked through the data reader.
        ''' </remarks>
        Public Function ExecuteReader(ByVal dbConnection As IDbConnection, ByVal commandText As String, ByVal commandType As System.Data.CommandType, ByVal sqlParameters As IDataParameter(), ByVal automations As Automations, ByVal commandTimeout As Integer) As IDataReader
            If automations = Automations.AutoCloseAndDisposeConnection OrElse automations = Automations.AutoOpenAndCloseAndDisposeConnection Then
                Throw New Exception("Can't close a data reader automatically since data has to be read first")
            End If

            Dim MyConn As IDbConnection = dbConnection
            Dim MyCmd As IDbCommand = MyConn.CreateCommand
            MyCmd.CommandText = commandText
            MyCmd.CommandType = commandType
            If commandTimeout >= 0 Then
                MyCmd.CommandTimeout = commandTimeout
            End If
            If sqlParameters IsNot Nothing Then
                For Each MySqlParam As IDataParameter In sqlParameters
                    MyCmd.Parameters.Add(MySqlParam)
                Next
            End If
            Dim Result As IDataReader
            Try
                If automations = Automations.AutoOpenAndCloseAndDisposeConnection OrElse automations = Automations.AutoOpenConnection Then
                    If MyConn IsNot Nothing AndAlso MyConn.State <> ConnectionState.Open Then
                        MyConn.Open()
                    End If
                End If
                If automations = Automations.AutoCloseAndDisposeConnection OrElse automations = Automations.AutoOpenAndCloseAndDisposeConnection Then
                    Result = MyCmd.ExecuteReader(CommandBehavior.CloseConnection)
                Else
                    Result = MyCmd.ExecuteReader()
                End If
            Catch ex As Exception
                Throw New CompuMaster.Data.DataQuery.DataException(MyCmd, ex)
            Finally
                'Keep the connection opened since the reader still requires processing
            End Try
            Return Result
        End Function

        ''' <summary>
        '''     Executes a command and return the data reader object for it
        ''' </summary>
        ''' <param name="dbConnection">The connection to the database</param>
        ''' <param name="commandText">The command text</param>
        ''' <param name="commandType">The command type</param>
        ''' <param name="sqlParameters">An optional list of SqlParameters</param>
        ''' <returns></returns>
        Public Function ExecuteReader(ByVal dbConnection As IDbConnection, ByVal commandText As String, ByVal commandType As System.Data.CommandType, ByVal sqlParameters As IDataParameter(), ByVal automations As Automations) As IDataReader
            Dim MyConn As IDbConnection = dbConnection
            Dim MyCmd As IDbCommand = MyConn.CreateCommand
            MyCmd.CommandText = commandText
            MyCmd.CommandType = commandType
            If sqlParameters IsNot Nothing Then
                For Each MySqlParam As IDataParameter In sqlParameters
                    MyCmd.Parameters.Add(MySqlParam)
                Next
            End If
            Dim Result As IDataReader
            Try
                If automations = Automations.AutoOpenAndCloseAndDisposeConnection OrElse automations = Automations.AutoOpenConnection Then
                    If MyConn IsNot Nothing AndAlso MyConn.State <> ConnectionState.Open Then
                        MyConn.Open()
                    End If
                End If
                If automations = Automations.AutoCloseAndDisposeConnection OrElse automations = Automations.AutoOpenAndCloseAndDisposeConnection Then
                    Result = MyCmd.ExecuteReader(CommandBehavior.CloseConnection)
                Else
                    Result = MyCmd.ExecuteReader()
                End If
            Catch ex As Exception
                Throw New CompuMaster.Data.DataQuery.DataException(MyCmd, ex)
            Finally
                'Keep the connection opened since the reader still requires processing
            End Try
            Return Result
        End Function

        ''' <summary>
        '''     Executes a command and return the data reader object for it
        ''' </summary>
        ''' <param name="dbCommand">The command with an assigned connection property value</param>
        ''' <param name="automations">Automation options for the connection</param>
        Public Function ExecuteReader(ByVal dbCommand As IDbCommand, ByVal automations As Automations) As IDataReader
            Dim MyCmd As IDbCommand = dbCommand
            Dim MyConn As IDbConnection = MyCmd.Connection
            Dim Result As IDataReader
            Try
                If automations = Automations.AutoOpenAndCloseAndDisposeConnection OrElse automations = Automations.AutoOpenConnection Then
                    If MyCmd.Connection.State <> ConnectionState.Open Then
                        MyConn.Open()
                    End If
                End If
                If automations = Automations.AutoCloseAndDisposeConnection OrElse automations = Automations.AutoOpenAndCloseAndDisposeConnection Then
                    Result = MyCmd.ExecuteReader(CommandBehavior.CloseConnection)
                Else
                    Result = MyCmd.ExecuteReader()
                End If
            Catch ex As Exception
                Throw New CompuMaster.Data.DataQuery.DataException(MyCmd, ex)
            Finally
                'Keep the connection opened since the reader still requires processing
            End Try
            Return Result
        End Function

        ''' <summary>
        '''     Fill a new data table with the result of a command
        ''' </summary>
        ''' <param name="dbCommand">The command object</param>
        ''' <param name="automations">Automation options for the connection</param>
        ''' <param name="tableName">The name for the new table</param>
        ''' <returns></returns>
        Public Function FillDataTable(ByVal dbCommand As IDbCommand, ByVal automations As Automations, ByVal tableName As String) As System.Data.DataTable
            Dim MyReader As IDataReader = Nothing
            Dim Result As New System.Data.DataTable
            Dim dbConnection As IDbConnection = dbCommand.Connection
            Try
                If automations = Automations.AutoOpenAndCloseAndDisposeConnection OrElse automations = Automations.AutoOpenConnection Then
                    If dbConnection.State <> ConnectionState.Open Then
                        dbConnection.Open()
                    End If
                End If
                'Attention: ExecuteReader doesn't allow auto-close of the connection
                Dim Automation As Automations
                If automations = Automations.AutoCloseAndDisposeConnection Then
                    Automation = Automations.None
                ElseIf automations = Automations.AutoOpenAndCloseAndDisposeConnection Then
                    Automation = Automations.AutoOpenConnection
                End If
                'Execute the reader
                MyReader = CompuMaster.Data.DataQuery.AnyIDataProvider.ExecuteReader(dbCommand, Automation)
                'Convert the reader to a data table
                Result = CompuMaster.Data.DataTablesTools.ConvertDataReaderToDataTable(MyReader, tableName)
            Catch ex As Exception
                Throw New CompuMaster.Data.DataQuery.DataException(dbCommand, ex)
            Finally
                If MyReader IsNot Nothing AndAlso MyReader.IsClosed = False Then
                    MyReader.Close()
                End If
                If automations = Automations.AutoOpenAndCloseAndDisposeConnection OrElse automations = Automations.AutoCloseAndDisposeConnection Then
                    CloseAndDisposeConnection(dbConnection)
                End If
            End Try
            Return Result
        End Function

        ''' <summary>
        '''     Fill a new data table with the result of a command
        ''' </summary>
        ''' <param name="dbConnection">The connection to the database</param>
        ''' <param name="commandText">The command text</param>
        ''' <param name="commandType">The command type</param>
        ''' <param name="sqlParameters">An optional list of SqlParameters</param>
        ''' <param name="automations">Automation options for the connection</param>
        ''' <param name="tableName">The name for the new table</param>
        ''' <param name="commandTimeout">A timeout value in seconds for the command object (negative values will be ignored and leave the timeout value on default)</param>
        ''' <returns></returns>
        Public Function FillDataTable(ByVal dbConnection As IDbConnection, ByVal commandText As String, ByVal commandType As System.Data.CommandType, ByVal sqlParameters As IDataParameter(), ByVal automations As Automations, ByVal tableName As String, ByVal commandTimeout As Integer) As System.Data.DataTable
            Dim MyCmd As IDbCommand = dbConnection.CreateCommand
            MyCmd.CommandType = commandType
            If commandTimeout >= 0 Then 'never assign a -1 value
                MyCmd.CommandTimeout = commandTimeout
            End If
            MyCmd.CommandText = commandText
            If sqlParameters IsNot Nothing Then
                For Each MySqlParam As IDataParameter In sqlParameters
                    MyCmd.Parameters.Add(MySqlParam)
                Next
            End If
            Return FillDataTable(MyCmd, automations, tableName)
        End Function

        ''' <summary>
        '''     Fill a new data table with the result of a command
        ''' </summary>
        ''' <param name="dbConnection">The connection to the database</param>
        ''' <param name="commandText">The command text</param>
        ''' <param name="commandType">The command type</param>
        ''' <param name="sqlParameters">An optional list of SqlParameters</param>
        ''' <param name="automations">Automation options for the connection</param>
        ''' <param name="tableName">The name for the new table</param>
        ''' <returns></returns>
        Public Function FillDataTable(ByVal dbConnection As IDbConnection, ByVal commandText As String, ByVal commandType As System.Data.CommandType, ByVal sqlParameters As IDataParameter(), ByVal automations As Automations, ByVal tableName As String) As System.Data.DataTable
            Return FillDataTable(dbConnection, commandText, commandType, sqlParameters, automations, tableName, -1)
        End Function

        ''' <summary>
        '''     Fill a new data table with the result of a command
        ''' </summary>
        ''' <param name="dbConnection">The connection to the database</param>
        ''' <param name="commandText">The command text</param>
        ''' <param name="commandType">The command type</param>
        ''' <param name="sqlParameters">An optional list of SqlParameters</param>
        ''' <param name="automations">Automation options for the connection</param>
        ''' <returns></returns>
        Public Function FillDataTable(ByVal dbConnection As IDbConnection, ByVal commandText As String, ByVal commandType As System.Data.CommandType, ByVal sqlParameters As IDataParameter(), ByVal automations As Automations) As System.Data.DataTable
            Return FillDataTable(dbConnection, commandText, commandType, sqlParameters, automations, Nothing)
        End Function

        ''' <summary>
        '''     Fill a new data table with the result of a command
        ''' </summary>
        ''' <param name="dbCommand">The command object</param>
        ''' <param name="automations">Automation options for the connection</param>
        ''' <returns></returns>
        Public Function FillDataTables(ByVal dbCommand As IDbCommand, ByVal automations As Automations) As System.Data.DataTable()
            Dim MyReader As IDataReader = Nothing
            Dim Results As New ArrayList
            Dim dbConnection As IDbConnection = dbCommand.Connection
            Try
                If automations = Automations.AutoOpenAndCloseAndDisposeConnection OrElse automations = Automations.AutoOpenConnection Then
                    If dbConnection.State <> ConnectionState.Open Then
                        dbConnection.Open()
                    End If
                End If
                'Attention: ExecuteReader doesn't allow auto-close of the connection
                Dim Automation As Automations
                If automations = Automations.AutoCloseAndDisposeConnection Then
                    Automation = Automations.None
                ElseIf automations = Automations.AutoOpenAndCloseAndDisposeConnection Then
                    Automation = Automations.AutoOpenConnection
                End If
                'Execute the reader
                MyReader = CompuMaster.Data.DataQuery.AnyIDataProvider.ExecuteReader(dbCommand, Automation)
                'Convert the reader to data tables
                Dim Result As System.Data.DataSet = CompuMaster.Data.DataTablesTools.ConvertDataReaderToDataSet(MyReader)
                For MyCounter As Integer = 0 To Result.Tables.Count - 1
                    Results.Add(Result.Tables(MyCounter))
                Next
            Catch ex As Exception
                Throw New CompuMaster.Data.DataQuery.DataException(dbCommand, ex)
            Finally
                If MyReader IsNot Nothing AndAlso MyReader.IsClosed = False Then
                    MyReader.Close()
                End If
                If automations = Automations.AutoOpenAndCloseAndDisposeConnection OrElse automations = Automations.AutoCloseAndDisposeConnection Then
                    CloseAndDisposeConnection(dbConnection)
                End If
            End Try
            Return CType(Results.ToArray(GetType(DataTable)), DataTable())
        End Function

        ''' <summary>
        '''     Securely close and dispose a database connection
        ''' </summary>
        ''' <param name="connection">The connection to close and dispose</param>
        Public Sub CloseAndDisposeConnection(ByVal connection As IDbConnection)
            If connection IsNot Nothing Then
                If CType(connection, Object).GetType Is GetType(SqlClient.SqlConnection) Then
                    If connection.State <> ConnectionState.Closed Then
                        connection.Close()
                    End If
                    connection.Dispose()
                Else
                    Try
                        If connection.State <> ConnectionState.Closed Then
                            connection.Close()
                        End If
                        connection.Dispose()
                    Catch ex As System.ObjectDisposedException
                        'Ignore - happens e.g. by NpgSql provider till 2.0.11.94
                    End Try
                End If
            End If
        End Sub

        ''' <summary>
        '''     Securely close a database connection
        ''' </summary>
        ''' <param name="connection">The connection to close</param>
        Public Sub CloseConnection(ByVal connection As IDbConnection)
            If connection IsNot Nothing Then
                If CType(connection, Object) Is GetType(SqlClient.SqlConnection) Then
                    If connection.State <> ConnectionState.Closed Then
                        connection.Close()
                    End If
                Else
                    Try
                        If connection.State <> ConnectionState.Closed Then
                            connection.Close()
                        End If
                    Catch ex As System.ObjectDisposedException
                        'Ignore - happens e.g. by NpgSql provider till 2.0.11.94
                    End Try
                End If
            End If
        End Sub

        ''' <summary>
        '''     Open a database connection if it is not already opened
        ''' </summary>
        ''' <param name="connection">The connection to open</param>
        Public Sub OpenConnection(ByVal connection As IDbConnection)
            If connection Is Nothing Then
                Throw New ArgumentNullException(NameOf(connection))
            End If
            If connection.State <> ConnectionState.Open Then
                connection.Open()
            End If
        End Sub

    End Module

End Namespace
