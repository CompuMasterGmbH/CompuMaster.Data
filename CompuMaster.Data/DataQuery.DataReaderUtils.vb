Option Explicit On 
Option Strict On

Namespace CompuMaster.Data.DataQuery

    ''' <summary>
    ''' Methods for simplifying the handling with data readers
    ''' </summary>
    Public Class DataReaderUtils

        ''' <summary>
        ''' Lookup if the reader contains a result column with the requested name
        ''' </summary>
        ''' <param name="reader">A data reader object</param>
        ''' <param name="columnName">The name of the column which shall be identified</param>
        ''' <returns>True if the column exist else False</returns>
        Public Shared Function ContainsColumn(ByVal reader As IDataReader, ByVal columnName As String) As Boolean
            If reader Is Nothing Then Throw New ArgumentNullException("reader", "Parameter reader is required")
            If columnName = Nothing Then Throw New ArgumentNullException("columnName", "Parameter columnName can't be an empty value")
            For MyCounter As Integer = 0 To reader.FieldCount - 1
                If LCase(reader.GetName(MyCounter)) = LCase(columnName) Then Return True
            Next
            Return False
        End Function

        ''' <summary>
        ''' Return the column names of a data reader as a String array
        ''' </summary>
        ''' <param name="reader">A data reader object</param>
        ''' <returns></returns>
        Public Shared Function ColumnNames(ByVal reader As IDataReader) As String()
            If reader Is Nothing Then Return Nothing
            Dim Result As New ArrayList
            For MyCounter As Integer = 0 To reader.FieldCount - 1
                Result.Add(reader.GetName(MyCounter))
            Next
            Return CType(Result.ToArray(GetType(String)), String())
        End Function

        ''' <summary>
        ''' Return the column data types of a data reader as an array
        ''' </summary>
        ''' <param name="reader">A data reader object</param>
        ''' <returns></returns>
        Public Shared Function DataTypes(ByVal reader As IDataReader) As Type()
            If reader Is Nothing Then Return Nothing
            Dim Result As New ArrayList
            For MyCounter As Integer = 0 To reader.FieldCount - 1
                Result.Add(reader.GetFieldType(MyCounter))
            Next
            Return CType(Result.ToArray(GetType(Type)), Type())
        End Function

    End Class

End Namespace
