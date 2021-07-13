Option Explicit On
Option Strict On

Imports System.Data
Imports CompuMaster.Data.Information

Namespace CompuMaster.Data.DataQuery

#Disable Warning CA2237 ' Mark ISerializable types with serializable
#Disable Warning CA1032 ' Implement standard exception constructors
    ''' <summary>
    '''     Data execution exceptions with details on the executed IDbCommand
    ''' </summary>
    Public Class DataException
#Enable Warning CA1032 ' Implement standard exception constructors
#Enable Warning CA2237 ' Mark ISerializable types with serializable
        Inherits System.Exception

        Private ReadOnly _commandText As String
        Private ReadOnly _command As IDbCommand

        Friend Sub New(ByVal command As IDbCommand, ByVal innerException As Exception)
            MyBase.New("Data layer exception", innerException)
            _command = command
            If _command IsNot Nothing Then
                _commandText = "ConnectionString (without sensitive data): " & Utils.ConnectionStringWithoutPasswords(command.Connection.ConnectionString) & ControlChars.CrLf
                _commandText = "CommandType: " & command.CommandType.ToString & ControlChars.CrLf
                _commandText &= "CommandText:" & ControlChars.CrLf & command.CommandText
                _commandText &= ControlChars.CrLf & ControlChars.CrLf
                If command.Parameters.Count > 0 Then
                    _commandText &= "Parameters:" & ControlChars.CrLf & ConvertParameterCollectionToString(command.Parameters) & ControlChars.CrLf
                Else
                    _commandText &= "Parameters:" & ControlChars.CrLf & "The parameters collection is empty" & ControlChars.CrLf
                End If
            End If
        End Sub

        ''' <summary>
        '''     Convert the collection with all the parameters to a plain text string
        ''' </summary>
        ''' <param name="parameters">An IDataParameterCollection of a IDbCommand</param>
        ''' <returns></returns>
        Private Shared Function ConvertParameterCollectionToString(ByVal parameters As System.Data.IDataParameterCollection) As String
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
                Result &= ControlChars.CrLf
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
            Return MyBase.ToString & ControlChars.CrLf & ControlChars.CrLf & _commandText
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
