Option Strict On
Option Explicit On

Imports CompuMaster.Data


Namespace CompuMaster.Data.Windows

    Friend Class Utils

        ''' <summary>
        ''' Execute query and provide results
        ''' </summary>
        Public Shared Function LoadDataForManipulationViaQuickEdit(ByVal selectCommand As System.Data.IDbCommand) As DataManipulationResult
            Dim Result As DataManipulationResult
            Try
                DataQuery.OpenConnection(selectCommand.Connection)
                Result = New DataManipulationResult(selectCommand, Nothing)
            Finally
                DataQuery.CloseConnection(selectCommand.Connection)
            End Try
            Return Result
        End Function

        ''' <summary>
        ''' Save changed data to the data source
        ''' </summary>
        Public Shared Sub SaveData(ByVal data As DataManipulationResult)
            Try
                DataQuery.OpenConnection(data.Command.Connection)
                data.UpdateChanges()
            Finally
                DataQuery.CloseConnection(data.Command.Connection)
            End Try
        End Sub

        ''' <summary>
        ''' Close the currently used connection
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Sub CloseAndDisposeQuickEditDataContainer(ByVal _DataContainer As DataManipulationResult)
            If _DataContainer IsNot Nothing AndAlso _DataContainer.Command IsNot Nothing Then
                DataQuery.CloseAndDisposeConnection(_DataContainer.Command.Connection)
                If _DataContainer.Command IsNot Nothing Then
                    _DataContainer.Command.Dispose()
                End If
            End If
        End Sub

    End Class

End Namespace