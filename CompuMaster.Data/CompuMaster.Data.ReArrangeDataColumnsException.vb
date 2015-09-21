Option Explicit On
Option Strict On

Namespace CompuMaster.Data

    ''' <summary>
    '''     An exception which gets thrown when converting data in the ReArrangeDataColumns methods
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[wezel]	14.04.2005	Created
    ''' </history>
    Friend Class ReArrangeDataColumnsException
        Inherits Exception

        Public Sub New(ByVal rowIndex As Integer, ByVal columnIndex As Integer, ByVal sourceColumnType As Type, ByVal targetColumnType As Type, ByVal problematicValue As Object, ByVal innerException As Exception)
            MyBase.New("Data conversion exception", innerException)
            _RowIndex = rowIndex
            _ColumnIndex = columnIndex
            _sourceColumnType = sourceColumnType
            _targetColumnType = targetColumnType
            _problematicValue = problematicValue
        End Sub

        Private _sourceColumnType As Type
        Public ReadOnly Property SourceColumnType() As Type
            Get
                Return _sourceColumnType
            End Get
        End Property

        Private _targetColumnType As Type
        Public ReadOnly Property TargetColumnType() As Type
            Get
                Return _targetColumnType
            End Get
        End Property

        Private _problematicValue As Object
        Public ReadOnly Property ProblematicValue() As Object
            Get
                Return _problematicValue
            End Get
        End Property

        Private _RowIndex As Integer
        Public ReadOnly Property RowIndex() As Integer
            Get
                Return _RowIndex
            End Get
        End Property

        Private _ColumnIndex As Integer
        Public ReadOnly Property ColumnIndex() As Integer
            Get
                Return _ColumnIndex
            End Get
        End Property

        Public Overrides ReadOnly Property Message() As String
            Get
                Dim Result As String
                Result = "Conversion exception in row index " & RowIndex & " and column index " & ColumnIndex & "." & System.Environment.NewLine
                Result &= "Data type in source table: " & SourceColumnType.ToString & System.Environment.NewLine
                Result &= "Data type in destination table: " & TargetColumnType.ToString & System.Environment.NewLine
                If ProblematicValue.GetType Is GetType(String) Then
                    Result &= "The problematic value is: " & CType(ProblematicValue, String) & System.Environment.NewLine
                End If
                If Not Me.InnerException Is Nothing Then
                    Result &= System.Environment.NewLine & "This is the inner exception message: " & System.Environment.NewLine & Me.InnerException.Message
                End If
                Return Result
            End Get
        End Property

    End Class

End Namespace