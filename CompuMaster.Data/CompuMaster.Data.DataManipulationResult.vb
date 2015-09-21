Option Explicit On
Option Strict On

Namespace CompuMaster.Data

    ''' <summary>
    ''' A container for a DataTable with its IDataAdapter and IDbCommand
    ''' </summary>
    ''' <remarks></remarks>
    Public Class DataManipulationResult
        Implements IDisposable

        Protected WithEvents DataTable As System.Data.DataTable
        Protected UpdateDataAdapter As System.Data.IDataAdapter
        Protected SelectCommand As System.Data.IDbCommand

        ''' <summary>
        ''' Create an empty instance 
        ''' </summary>
        ''' <remarks></remarks>
        Protected Sub New()
        End Sub

        ''' <summary>
        ''' Fill a DataManipulationResult using the given SELECT command for updating a modifed version on a later step
        ''' </summary>
        ''' <param name="command">The SELECT command for retrieving the data</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal command As System.Data.IDbCommand)
            Dim Result As CompuMaster.Data.DataManipulationResult = CompuMaster.Data.Manipulation.LoadQueryDataForManipulationViaCode(command)
            Me.UpdateDataAdapter = Result.DataAdapter
            Me.SelectCommand = Result.Command
            Me.DataTable = Result.Table
        End Sub

        ''' <summary>
        ''' Create a new instance of DataManipulationResults for updating queried data on a later step
        ''' </summary>
        ''' <param name="command">The SELECT command for retrieving the data</param>
        ''' <param name="dataAdapter">An instance of data adapter using the SELECT command</param>
        ''' <remarks></remarks>
        Friend Sub New(ByVal command As System.Data.IDbCommand, ByVal dataAdapter As System.Data.IDataAdapter)
            Me.New(Nothing, command, dataAdapter)
        End Sub

        ''' <summary>
        ''' Create a new instance of DataManipulationResults for updating queried data on a later step
        ''' </summary>
        ''' <param name="table">A new table which shall contain the queried data</param>
        ''' <param name="command">The SELECT command for retrieving the data</param>
        ''' <param name="dataAdapter">An instance of data adapter using the SELECT command</param>
        ''' <remarks></remarks>
        Friend Sub New(ByVal table As System.Data.DataTable, ByVal command As System.Data.IDbCommand, ByVal dataAdapter As System.Data.IDataAdapter)
            If table Is Nothing Then table = New System.Data.DataTable("livedataclone")
            If command Is Nothing Then Throw New ArgumentNullException("command")
            If dataAdapter Is Nothing Then Throw New ArgumentNullException("dataAdapter")
            Me.DataTable = table
            Me.SelectCommand = command
            Me.UpdateDataAdapter = dataAdapter
        End Sub

        ''' <summary>
        ''' Save all changes to the data source (requires an opened connection)
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub UpdateChanges()
            CompuMaster.Data.Manipulation.UpdateCodeManipulatedData(Me, False)
        End Sub

        ''' <summary>
        ''' The table which holds the queried data
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Table() As System.Data.DataTable
            Get
                Return DataTable
            End Get
        End Property

        ''' <summary>
        ''' The data adapter which is responsible to upload the changed data
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property DataAdapter() As System.Data.IDataAdapter
            Get
                Return UpdateDataAdapter
            End Get
        End Property

        ''' <summary>
        ''' The select command
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Command() As System.Data.IDbCommand
            Get
                Return SelectCommand
            End Get
        End Property

        ''' <summary>
        ''' Clean up
        ''' </summary>
        ''' <param name="disposing"></param>
        ''' <remarks></remarks>
        Protected Overridable Sub Dispose(ByVal disposing As Boolean)
            If disposing Then
                If Not SelectCommand Is Nothing Then
                    CompuMaster.Data.DataQuery.CloseAndDisposeConnection(SelectCommand.Connection)
                    SelectCommand.Dispose()
                    DataTable = Nothing
                End If
            End If
        End Sub

#Region " IDisposable Support "
        Public Sub Dispose() Implements IDisposable.Dispose
            ' Ändern Sie diesen Code nicht. Fügen Sie oben in Dispose(ByVal disposing As Boolean) Bereinigungscode ein.
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
#End Region

        Private Sub DataTable_RowChanged(ByVal sender As Object, ByVal e As System.Data.DataRowChangeEventArgs) Handles DataTable.RowChanged
            If Not Me.Table.GetChanges() Is Nothing Then
                RaiseEvent DataChanged()
            End If
        End Sub

        Private Sub DataTable_RowDeleted(ByVal sender As Object, ByVal e As System.Data.DataRowChangeEventArgs) Handles DataTable.RowDeleted
            If Not Me.Table.GetChanges() Is Nothing Then
                RaiseEvent DataChanged()
            End If
        End Sub

        ''' <summary>
        ''' The data in the table has been changed and is available for saving/uploading back
        ''' </summary>
        ''' <remarks></remarks>
        Public Event DataChanged()

    End Class

End Namespace
