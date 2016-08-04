Option Explicit On
Option Strict On

Imports System.Collections.Generic

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
        'Private ReadOnly Property AssemblyPath As String
        '    Get
        '        Return Assembly.Location
        '    End Get
        'End Property
        Private ReadOnly Property AssemblyName As String
            Get
                Return Me.Assembly.FullName.Substring(0, Me.Assembly.FullName.IndexOf(","c))
            End Get
        End Property
        Public ReadOnly Property ConnectionType As System.Type
        Public ReadOnly Property CommandBuilderType As System.Type
        Public ReadOnly Property DataAdapterType As System.Type
        'Private ReadOnly Property ConnectionTypeName As String
        '    Get
        '        Return Me.ConnectionType.Name
        '    End Get
        'End Property
        Public Function CreateConnection() As IDbConnection
            Return CType(Activator.CreateInstance(Me.ConnectionType), IDbConnection)
        End Function

        Public ReadOnly Property CommandType As System.Type
        'Private ReadOnly Property CommandTypeName As String
        '    Get
        '        Return Me.CommandType.Name
        '    End Get
        'End Property
        Public Function CreateCommand() As IDbCommand
            Return CType(Activator.CreateInstance(Me.CommandType), IDbCommand)
        End Function
        Public Function CreateCommandBuilder() As System.Data.Common.DbCommandBuilder
            If Me.CommandBuilderType Is Nothing Then
                Return Nothing
            Else
                Return CType(Activator.CreateInstance(Me.CommandBuilderType), System.Data.Common.DbCommandBuilder)
            End If
        End Function
        Public Function CreateDataAdapter() As System.Data.IDbDataAdapter
            If Me.DataAdapterType Is Nothing Then
                Return Nothing
            Else
                Return CType(Activator.CreateInstance(Me.DataAdapterType), System.Data.IDbDataAdapter)
            End If
        End Function

        'Public ReadOnly Property CommandBuilderType As System.Type
        '    Get
        '        Static BufferedResult As System.Type
        '        If BufferedResult Is Nothing Then
        '            BufferedResult = Me.CreateConnection.CreateCommand.GetType
        '        End If
        '        Return BufferedResult
        '    End Get
        'End Property
        'Public ReadOnly Property DataAdapterType As System.Type
        '    Get
        '        Static BufferedResult As System.Type
        '        If BufferedResult Is Nothing Then
        '            BufferedResult = System.Data.DbProviderFactories.GetFactory(Me.CreateConnection.CreateCommand
        '        End If
        '        Return BufferedResult
        '    End Get
        'End Property

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
            Dim AlreadyLoadedAssemblies As System.Reflection.Assembly() = AppDomain.CurrentDomain.GetAssemblies
            Dim Result As New List(Of DataProvider)
            For Each asm As System.Reflection.Assembly In AlreadyLoadedAssemblies
                Dim asmName As String = asm.FullName.Substring(0,asm.fullname.IndexOf(","c)) 'asm.GetName.Name.ToLowerInvariant
                Dim TryToFindDataConnectorsInAssembly As Boolean
                If asmName = "system.data" OrElse asmName = "system.data.oracleclient" Then
                    TryToFindDataConnectorsInAssembly = True
                ElseIf asmName = "system" OrElse asmName.StartsWith("system.") OrElse asmName.StartsWith("compumaster.data") OrElse asmName.StartsWith("digitalrune.windows") OrElse asmName.StartsWith("mscorlib") OrElse asmName.StartsWith("mono.security") OrElse asmName.StartsWith("microsoft.") OrElse asmName.StartsWith("vshost") Then
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

        Public Shared Function AvailableDataProviders(assembly As System.Reflection.Assembly) As List(Of DataProvider)
            Dim Result As New List(Of DataProvider)
            For Each t As Type In assembly.GetTypes()
                For Each iface As Type In t.GetInterfaces()
                    If iface Is GetType(System.Data.IDbConnection) AndAlso t IsNot GetType(System.Data.Common.DbConnection) Then
                        Dim IDbCommandType As System.Type = CType(Activator.CreateInstance(t), IDbConnection).CreateCommand.GetType
                        Dim IDbDataAdapterType As System.Type = FindIDbDataApaterType(assembly, IDbCommandType)
                        Dim DbCommandBuilderType As System.Type = FindDbCommandBuilder(assembly, IDbDataAdapterType)
                        Dim Provider As New DataProvider(assembly, t, IDbCommandType, DbCommandBuilderType, IDbDataAdapterType)
                        Result.Add(Provider)
                    End If
                Next
            Next
            Return Result
        End Function

        Private Shared Function FindIDbDataApaterType(assembly As System.Reflection.Assembly, targetForCommand As System.Type) As System.Type
            For Each t As Type In assembly.GetTypes()
                For Each iface As Type In t.GetInterfaces()
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
            Next
            Return Nothing
        End Function

        Private Shared Function FindDbCommandBuilder(assembly As System.Reflection.Assembly, targetForDataAdapter As System.Type) As System.Type
            For Each t As Type In assembly.GetTypes()
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