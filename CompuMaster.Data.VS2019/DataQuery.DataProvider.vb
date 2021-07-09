Option Explicit On
Option Strict On

Imports System.Collections.Generic
Imports System.Data

Namespace CompuMaster.Data.DataQuery

    ''' <summary>
    ''' A data provider which implements the System.Data.IDbConnection/IDbCommand interface
    ''' </summary>
    Public Class DataProvider

        Public Sub New(assembly As System.Reflection.Assembly, connectionType As System.Type, commandType As System.Type, commandBuilderType As System.Type, dataAdapterType As System.Type)
            Me.Assembly = assembly
            Me.ConnectionType = connectionType
            Me.CommandType = commandType
            Me.CommandBuilderType = commandBuilderType
            Me.DataAdapterType = dataAdapterType
        End Sub

        Public ReadOnly Property Assembly As System.Reflection.Assembly

        Private ReadOnly Property AssemblyName As String
            Get
                Return Me.Assembly.FullName.Substring(0, Me.Assembly.FullName.IndexOf(","c))
            End Get
        End Property
        Public ReadOnly Property ConnectionType As System.Type
        Public ReadOnly Property CommandBuilderType As System.Type
        Public ReadOnly Property DataAdapterType As System.Type

        Public Function CreateConnection() As IDbConnection
            Return CType(Activator.CreateInstance(Me.ConnectionType), IDbConnection)
        End Function
        Public Function CreateConnection(connectionString As String) As IDbConnection
            Dim Result As IDbConnection = Me.CreateConnection
            Result.ConnectionString = connectionString
            Return Result
        End Function

        Public ReadOnly Property CommandType As System.Type

        Public Function CreateCommand() As IDbCommand
            Return CType(Activator.CreateInstance(Me.CommandType), IDbCommand)
        End Function
        Public Function CreateCommand(sql As String) As IDbCommand
            Dim Result As IDbCommand = Me.CreateCommand
            Result.CommandText = sql
            Return Result
        End Function
        Public Function CreateCommand(sql As String, connectionString As String) As IDbCommand
            Dim Result As IDbCommand = Me.CreateCommand
            Result.CommandText = sql
            Result.Connection = Me.CreateConnection(connectionString)
            Return Result
        End Function
        Public Function CreateCommandBuilder() As System.Data.Common.DbCommandBuilder
            If Me.CommandBuilderType Is Nothing Then
                Return Nothing
            Else
                Return CType(Activator.CreateInstance(Me.CommandBuilderType), System.Data.Common.DbCommandBuilder)
            End If
        End Function
        Public Function CreateCommandBuilder(dataAdapter As System.Data.IDbDataAdapter) As System.Data.Common.DbCommandBuilder
            Dim Result As System.Data.Common.DbCommandBuilder = Me.CreateCommandBuilder
            Result.DataAdapter = CType(dataAdapter, System.Data.Common.DbDataAdapter)
            Return Result
        End Function
        Public Function CreateDataAdapter() As System.Data.IDbDataAdapter
            If Me.DataAdapterType Is Nothing Then
                Return Nothing
            Else
                Return CType(Activator.CreateInstance(Me.DataAdapterType), System.Data.IDbDataAdapter)
            End If
        End Function
        Public Function CreateDataAdapter(selectCommand As System.Data.IDbCommand) As System.Data.IDbDataAdapter
            Dim Result As System.Data.IDbDataAdapter = Me.CreateDataAdapter
            Result.SelectCommand = selectCommand
            Return Result
        End Function


        ''' <summary>
        ''' A common title for the data connector
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property Title As String
            Get
                Dim Result As String
                Result = Strings.Replace(Me.ConnectionType.Name, "Connection", "",,, CompareMethod.Text)
                If Result = "Sql" AndAlso Me.AssemblyName.ToLowerInvariant = "system.data" Then
                    Result = "SqlClient"
                ElseIf Result = "Odbc" AndAlso Me.AssemblyName.ToLowerInvariant = "system.data" Then
                    Result = "ODBC"
                End If
                Return Result
            End Get
        End Property

        Public Overrides Function ToString() As String
            Return Me.Title
        End Function

        Public Shared Function LookupDataProvider(connection As System.Data.IDbConnection) As DataProvider
            Dim availableProviders As List(Of DataProvider) = AvailableDataProviders()
            For MyCounter As Integer = 0 To availableProviders.Count - 1
                If CType(connection, Object).GetType Is availableProviders(MyCounter).ConnectionType Then Return availableProviders(MyCounter)
            Next
            Return Nothing
        End Function
        Public Shared Function LookupDataProvider(connection As System.Data.IDbConnection, appDomain As AppDomain) As DataProvider
            Dim availableProviders As List(Of DataProvider) = AvailableDataProviders(appDomain)
            For MyCounter As Integer = 0 To availableProviders.Count - 1
                If CType(connection, Object).GetType Is availableProviders(MyCounter).ConnectionType Then Return availableProviders(MyCounter)
            Next
            Return Nothing
        End Function

        Public Shared Function LookupDataProvider(title As String) As DataProvider
            Dim availableProviders As List(Of DataProvider) = AvailableDataProviders()
            For MyCounter As Integer = 0 To availableProviders.Count - 1
                If title = availableProviders(MyCounter).Title Then Return availableProviders(MyCounter)
            Next
            Return Nothing
        End Function
        Public Shared Function LookupDataProvider(title As String, appDomain As AppDomain) As DataProvider
            Dim availableProviders As List(Of DataProvider) = AvailableDataProviders(appDomain)
            For MyCounter As Integer = 0 To availableProviders.Count - 1
                If title = availableProviders(MyCounter).Title Then Return availableProviders(MyCounter)
            Next
            Return Nothing
        End Function
        Public Shared Function LookupDataProvider(title As String, assembly As System.Reflection.Assembly) As DataProvider
            Dim availableProviders As List(Of DataProvider) = AvailableDataProviders(assembly)
            For MyCounter As Integer = 0 To availableProviders.Count - 1
                If title = availableProviders(MyCounter).Title Then Return availableProviders(MyCounter)
            Next
            Return Nothing
        End Function

        Public Shared Function AvailableDataProviders() As List(Of DataProvider)
            Return AvailableDataProviders(AppDomain.CurrentDomain)
        End Function

        Public Shared Function AvailableDataProviders(appDomain As AppDomain) As List(Of DataProvider)
            Dim AlreadyLoadedAssemblies As System.Reflection.Assembly()
            If appDomain IsNot Nothing Then
                AlreadyLoadedAssemblies = appDomain.GetAssemblies
            Else
                AlreadyLoadedAssemblies = System.AppDomain.CurrentDomain.GetAssemblies
            End If
            Dim Result As New List(Of DataProvider)
            For Each asm As System.Reflection.Assembly In AlreadyLoadedAssemblies
                Dim asmName As String = asm.FullName.Substring(0, asm.FullName.IndexOf(","c)) 'asm.GetName.Name.ToLowerInvariant
                Dim TryToFindDataConnectorsInAssembly As Boolean
                If asmName = "system.data" OrElse asmName = "system.data.oracleclient" Then
                    TryToFindDataConnectorsInAssembly = True
                ElseIf asmName = "system" OrElse asmName.StartsWith("system.") OrElse asmName.StartsWith("compumaster.data") OrElse asmName.StartsWith("mscorlib") OrElse asmName.StartsWith("mono.security") OrElse asmName.StartsWith("vshost") Then
                    TryToFindDataConnectorsInAssembly = False
                Else
                    TryToFindDataConnectorsInAssembly = True
                End If
                If TryToFindDataConnectorsInAssembly Then
                    'Analyse all loaded assemblies for available data connectors
                    Result.AddRange(AvailableDataProviders(asm))
                End If
            Next
            Return Result
        End Function

        ''' <summary>
        ''' GetAssemblyTypes might throw System.Reflection.ReflectionTypeLoadException - by .NET exception handling it's required to put this into a separate method to be able to catch this exception
        ''' </summary>
        ''' <param name="assembly"></param>
        ''' <returns></returns>
        Private Shared Function GetAssemblyTypes(assembly As System.Reflection.Assembly) As System.Type()
            Return assembly.GetTypes()
        End Function
        Private Shared Function GetAssemblyTypesSafely(assembly As System.Reflection.Assembly) As System.Type()
            Dim AssemblyTypes As System.Type()
            Try
                AssemblyTypes = GetAssemblyTypes(assembly)
            Catch ex As System.Reflection.ReflectionTypeLoadException
                AssemblyTypes = ex.Types
            End Try
            Return AssemblyTypes
        End Function
        ''' <summary>
        ''' GetTypeInterfaces might throw System.Reflection.ReflectionTypeLoadException - by .NET exception handling it's required to put this into a separate method to be able to catch this exception
        ''' </summary>
        ''' <param name="type"></param>
        ''' <returns></returns>
        Private Shared Function GetTypeInterfaces(type As System.Type) As System.Type()
            Return type.GetInterfaces()
        End Function
        Private Shared Function GetTypeInterfacesSafely(type As System.Type) As System.Type()
            Dim TypeInterfaces As System.Type()
            Try
                TypeInterfaces = GetTypeInterfaces(type)
            Catch ex As System.Reflection.ReflectionTypeLoadException
                TypeInterfaces = ex.Types
            End Try
            Return TypeInterfaces
        End Function
        Private Shared Function GetTypePublicProperties(type As System.Type) As System.Reflection.PropertyInfo()
            Return type.GetProperties(Reflection.BindingFlags.Public)
        End Function
        Private Shared Function GetTypePublicPropertiesSafely(type As System.Type) As System.Reflection.PropertyInfo()
            Dim TypeProperties As System.Reflection.PropertyInfo()
            Try
                TypeProperties = GetTypePublicProperties(type)
            Catch ex As System.Reflection.ReflectionTypeLoadException
                'just ignore this type
                Return New System.Reflection.PropertyInfo() {}
            End Try
            Return TypeProperties
        End Function


        Public Shared Function AvailableDataProviders(assembly As System.Reflection.Assembly) As List(Of DataProvider)
            Dim IsMonoRuntime As Boolean = Type.GetType("Mono.Runtime") IsNot Nothing
            Dim Result As New List(Of DataProvider)
            For Each t As Type In GetAssemblyTypesSafely(assembly)
                If IsMonoRuntime AndAlso t IsNot Nothing AndAlso t Is GetType(System.Data.OleDb.OleDbConnection) Then
                    'Mono runtime throws NotImplementedExceptions for OleDb stubs, but more important their Dispose method also throw NotImplementedExceptions causing garbage collector to crash causing AppDomain to crash
                    'Workaround for now because of https://github.com/mono/mono/issues/20975: don't use OleDbConnection at Mono at all
                ElseIf t IsNot Nothing Then
                    For Each iface As Type In GetTypeInterfacesSafely(t)
                        If iface Is GetType(System.Data.IDbConnection) AndAlso t IsNot GetType(System.Data.Common.DbConnection) Then
                            Try
                                Dim IDbCommandType As System.Type
                                Try
                                    IDbCommandType = CType(Activator.CreateInstance(t), IDbConnection).CreateCommand.GetType
                                Catch ex As System.Reflection.TargetInvocationException
                                    IDbCommandType = FindIDbCommandType(assembly, t)
                                End Try
                                Dim IDbDataAdapterType As System.Type = FindIDbDataApaterType(assembly, IDbCommandType)
                                Dim DbCommandBuilderType As System.Type = FindDbCommandBuilder(assembly, IDbDataAdapterType)
                                Dim Provider As New DataProvider(assembly, t, IDbCommandType, DbCommandBuilderType, IDbDataAdapterType)
                                Result.Add(Provider)
                            Catch ex As NotImplementedException
                                'Ignore OleDbProviders on Mono .NET throwing NotImplementedExceptions
                            End Try
                        End If
                    Next
                End If
            Next
            Return Result
        End Function

        Public Shared ReadOnly Property AvailableDataProvidersFoundExceptions As New List(Of DataProviderDetectionException)
        Public Shared ReadOnly Property AvailableDataProvidersFoundExceptions(assembly As System.Reflection.Assembly) As DataProviderDetectionException
            Get
                For MyCounter As Integer = 0 To DataProvider.AvailableDataProvidersFoundExceptions.Count - 1
                    If DataProvider.AvailableDataProvidersFoundExceptions()(MyCounter).Assembly Is assembly Then
                        Return DataProvider.AvailableDataProvidersFoundExceptions()(MyCounter)
                    End If
                Next
                Return Nothing
            End Get
        End Property

#Disable Warning CA1032 ' Implement standard exception constructors
#Disable Warning CA2237 ' Mark ISerializable types with serializable
        Public Class DataProviderDetectionException
#Enable Warning CA2237 ' Mark ISerializable types with serializable
#Enable Warning CA1032 ' Implement standard exception constructors
            Inherits Exception

            Friend Sub New(assembly As System.Reflection.Assembly, innerException As Exception)
                MyBase.New("Reflection of assembly " & assembly.FullName & " failed", innerException)
                Me.Assembly = assembly
            End Sub

            Public Property Assembly As System.Reflection.Assembly
        End Class

        Private Shared Function FindIDbCommandType(assembly As System.Reflection.Assembly, targetForConnection As System.Type) As System.Type
            For Each t As Type In GetAssemblyTypesSafely(assembly)
                If t IsNot Nothing Then
                    For Each iface As Type In GetTypeInterfaces(t)
                        If iface Is GetType(System.Data.IDbCommand) AndAlso t IsNot GetType(System.Data.Common.DbCommand) Then
                            For Each tProperty As Reflection.PropertyInfo In GetTypePublicPropertiesSafely(t)
                                If tProperty.PropertyType Is targetForConnection Then
                                    Return t
                                End If
                            Next
                        End If
                    Next
                End If
            Next
            Return Nothing
        End Function

        Private Shared Function FindIDbDataApaterType(assembly As System.Reflection.Assembly, targetForCommand As System.Type) As System.Type
            For Each t As Type In GetAssemblyTypesSafely(assembly)
                If t IsNot Nothing Then
                    For Each iface As Type In GetTypeInterfaces(t)
                        If iface Is GetType(System.Data.IDbDataAdapter) AndAlso t IsNot GetType(System.Data.Common.DbDataAdapter) Then
                            For Each tContructor As Reflection.ConstructorInfo In t.GetConstructors
                                For Each tConstructorParameter As Reflection.ParameterInfo In tContructor.GetParameters
                                    If tConstructorParameter.ParameterType Is targetForCommand Then
                                        Return t
                                    End If
                                Next
                            Next
                        End If
                    Next
                End If
            Next
            Return Nothing
        End Function

        Private Shared Function FindDbCommandBuilder(assembly As System.Reflection.Assembly, targetForDataAdapter As System.Type) As System.Type
            For Each t As Type In GetAssemblyTypesSafely(assembly)
                If t IsNot Nothing Then
                    For Each tParent As Type In AllBaseTypes(t)
                        If tParent Is GetType(System.Data.Common.DbCommandBuilder) AndAlso t IsNot GetType(System.Data.Common.DbCommandBuilder) Then
                            For Each tContructor As Reflection.ConstructorInfo In t.GetConstructors
                                For Each tConstructorParameter As Reflection.ParameterInfo In tContructor.GetParameters
                                    If tConstructorParameter.ParameterType Is targetForDataAdapter Then
                                        Return t
                                    End If
                                Next
                            Next
                        End If
                    Next
                End If
            Next
            Return Nothing
        End Function

        Private Shared Function AllBaseTypes(item As Type) As List(Of Type)
            Dim Result As New List(Of Type)
            AllBaseTypes(Result, item)
            Return Result
        End Function
        Private Shared Sub AllBaseTypes(resultList As List(Of Type), item As Type)
            If item.BaseType IsNot Nothing Then
                resultList.Add(item.BaseType)
                AllBaseTypes(resultList, item.BaseType)
            End If
        End Sub

    End Class
End Namespace