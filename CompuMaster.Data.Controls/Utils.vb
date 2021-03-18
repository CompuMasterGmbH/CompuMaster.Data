Option Strict On
Option Explicit On

Namespace CompuMaster.Data.Windows

    Friend Class Utils

        ''' <summary>
        ''' Execute query and provide results
        ''' </summary>
        Public Shared Function LoadDataForManipulationViaQuickEdit(ByVal selectCommand As System.Data.IDbCommand) As CompuMaster.Data.DataManipulationResult
            Dim Result As CompuMaster.Data.DataManipulationResult
            Try
                CompuMaster.Data.DataQuery.OpenConnection(selectCommand.Connection)
                Result = New CompuMaster.Data.DataManipulationResult(selectCommand, Nothing)
            Finally
                CompuMaster.Data.DataQuery.CloseConnection(selectCommand.Connection)
            End Try
            Return Result
        End Function

        ''' <summary>
        ''' Save changed data to the data source
        ''' </summary>
        Public Shared Sub SaveData(ByVal data As CompuMaster.Data.DataManipulationResult)
            Try
                CompuMaster.Data.DataQuery.OpenConnection(data.Command.Connection)
                data.UpdateChanges()
            Finally
                CompuMaster.Data.DataQuery.CloseConnection(data.Command.Connection)
            End Try
        End Sub

        ''' <summary>
        ''' Close the currently used connection
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Sub CloseAndDisposeQuickEditDataContainer(ByVal _DataContainer As CompuMaster.Data.DataManipulationResult)
            If _DataContainer IsNot Nothing AndAlso _DataContainer.Command IsNot Nothing Then
                CompuMaster.Data.DataQuery.CloseAndDisposeConnection(_DataContainer.Command.Connection)
                If _DataContainer.Command IsNot Nothing Then
                    _DataContainer.Command.Dispose()
                End If
            End If
        End Sub

    End Class

End Namespace