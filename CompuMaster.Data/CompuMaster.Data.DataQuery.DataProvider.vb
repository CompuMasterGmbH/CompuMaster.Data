Option Explicit On
Option Strict On

Imports System.Collections.Generic

Namespace CompuMaster.Data.DataQuery

    ''' <summary>
    ''' A data provider which implements the System.Data.IDbConnection/IDbCommand interface
    ''' </summary>
    Public Class DataProvider

        Public Sub New(assembly As System.Reflection.Assembly, connectionType As System.Type)
            Me.Assembly = assembly
            Me.ConnectionType = connectionType
        End Sub

        Public ReadOnly Property Assembly As System.Reflection.Assembly
        Public ReadOnly Property AssemblyPath As String
            Get
                Return Assembly.Location
            End Get
        End Property
        Public ReadOnly Property AssemblyName As String
            Get
                Return Assembly.GetName.Name
            End Get
        End Property
        Public ReadOnly Property ConnectionType As System.Type
        Public ReadOnly Property ConnectionTypeName As String
            Get
                Return Me.ConnectionType.Name
            End Get
        End Property
        Public Function CreateConnection() As IDbConnection
            Return CType(Activator.CreateInstance(Me.ConnectionType), IDbConnection)
        End Function

        Public ReadOnly Property CommandType As System.Type
            Get
                Static BufferedResult As System.Type
                If BufferedResult Is Nothing Then
                    BufferedResult = Me.CreateConnection.CreateCommand.GetType
                End If
                Return BufferedResult
            End Get
        End Property
        Public Function CreateCommand() As IDbCommand
            Return CType(Activator.CreateInstance(Me.CommandType), IDbCommand)
        End Function

        ''' <summary>
        ''' A common title for the data connector
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property Title As String
            Get
                Dim Result As String
                Result = Strings.Replace(Me.ConnectionTypeName, "Connection", "",,, CompareMethod.Text)
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

        Public Shared Function AvailableDataProviders() As List(Of DataProvider)
            Return AvailableDataProviders(AppDomain.CurrentDomain)
        End Function

        Public Shared Function AvailableDataProviders(appDomain As AppDomain) As List(Of DataProvider)
            Dim AlreadyLoadedAssemblies As System.Reflection.Assembly() = AppDomain.CurrentDomain.GetAssemblies
            Dim Result As New List(Of DataProvider)
            For Each asm As System.Reflection.Assembly In AlreadyLoadedAssemblies
                Dim asmName As String = asm.GetName.Name.ToLowerInvariant
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
                    If iface Is GetType(System.Data.IDbConnection) And t IsNot GetType(System.Data.Common.DbConnection) Then
                        Dim Provider As New DataProvider(assembly, t)
                        Result.Add(Provider)
                    End If
                Next
            Next
            Return Result
        End Function

    End Class
End Namespace