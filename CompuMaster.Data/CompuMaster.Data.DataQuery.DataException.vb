Option Explicit On
Option Strict On

Namespace CompuMaster.Data.DataQuery

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Data execution exceptions with details on the executed IDbCommand
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[adminwezel]	23.06.2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Class DataException
        Inherits System.Exception

        Private _commandText As String
        Private _command As IDbCommand

        Friend Sub New(ByVal command As IDbCommand, ByVal innerException As Exception)
            MyBase.New("Data layer exception", innerException)
            _command = command
            If Not _command Is Nothing Then
                _commandText = "ConnectionString (without sensitive data): " & Utils.ConnectionStringWithoutPasswords(command.Connection.ConnectionString) & vbNewLine
                _commandText = "CommandType: " & command.CommandType.ToString & vbNewLine
                _commandText &= "CommandText:" & vbNewLine & command.CommandText
                _commandText &= vbNewLine & vbNewLine
                If command.Parameters.Count > 0 Then
                    _commandText &= "Parameters:" & vbNewLine & ConvertParameterCollectionToString(command.Parameters) & vbNewLine
                Else
                    _commandText &= "Parameters:" & vbNewLine & "The parameters collection is empty" & vbNewLine
                End If
            End If
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''     Convert the collection with all the parameters to a plain text string
        ''' </summary>
        ''' <param name="parameters">An IDataParameterCollection of a IDbCommand</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[adminwezel]	23.06.2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Private Function ConvertParameterCollectionToString(ByVal parameters As System.Data.IDataParameterCollection) As String
            Dim Result As String = Nothing
            For MyCounter As Integer = 0 To parameters.Count - 1
                Result &= "Parameter " & MyCounter & ": "
                Try
                    Result &= CType(parameters(MyCounter), IDataParameter).ParameterName & ": "
                    Try
                        If CType(parameters(MyCounter), IDataParameter).Value Is Nothing Then
                            Result &= "{null}"
                        ElseIf IsDBNull(CType(parameters(MyCounter), IDataParameter).Value) Then
                            Result &= "{DBNull.Value}"
                        Else
                            Result &= CType(parameters(MyCounter), IDataParameter).Value.ToString
                        End If
                    Catch
                        Result &= "{" & CType(parameters(MyCounter), IDataParameter).Value.GetType.ToString & "}"
                    End Try
                Catch
                    Result &= "{" & parameters(MyCounter).GetType.ToString & "}"
                End Try
                Result &= vbNewLine
            Next
            Return Result
        End Function

        Public ReadOnly Property Command As IDbCommand
            Get
                Return _command
            End Get
        End Property

        ''' <summary>
        ''' The complete and detailed exception information inclusive the command text
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        Public Overrides Function ToString() As String
            Return MyBase.ToString & vbNewLine & vbNewLine & _commandText
        End Function

        ''' <summary>
        ''' Provides simplified overview on most important details on command text and command environment
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property DetailsOnCommandEnvironment As String
            Get
                Return _commandText
            End Get
        End Property

        ''' <summary>
        ''' Detailed exception details without details on command environment
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ToStringWithoutExceptionDetails() As String
            Return MyBase.ToString
        End Function

    End Class

End Namespace
