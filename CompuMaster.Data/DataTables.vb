Option Explicit On 
Option Strict On

Namespace CompuMaster.Data

    ''' <summary>
    '''     Common DataTable operations
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    <CodeAnalysis.SuppressMessage("Major Code Smell", "S3385:""Exit"" statements should not be used", Justification:="<Ausstehend>")>
    Public NotInheritable Class DataTables

        ''' <summary>
        ''' Remove rows from a table which don't match with a given range of values in a defined column
        ''' </summary>
        ''' <param name="column">The column whose values shall be verified</param>
        ''' <param name="values">The values which are required to keep a row; all rows without a matching value will be removed</param>
        ''' <remarks>Please note: String comparison is case-sensitive</remarks>
        Public Shared Sub RemoveRowsWithWithoutRequiredValuesInColumn(ByVal column As System.Data.DataColumn, ByVal values As Object())
            Dim RowsOkay As New ArrayList
            For RowCounter As Integer = column.Table.Rows.Count - 1 To 0 Step -1
                Dim rowValue As Object = column.Table.Rows(RowCounter)(column)
                For ValueCounter As Integer = 0 To values.Length - 1
                    Dim IgnoreValueChecks As Boolean = False
                    If IsDBNull(values(ValueCounter)) AndAlso IsDBNull(rowValue) Then
                        RowsOkay.Add(RowCounter)
                        Exit For
                    ElseIf Not IsDBNull(values(ValueCounter)) AndAlso IsDBNull(rowValue) Then
                        IgnoreValueChecks = True
                    ElseIf IsDBNull(values(ValueCounter)) AndAlso Not IsDBNull(rowValue) Then
                        IgnoreValueChecks = True
                    End If
                    If IgnoreValueChecks = False Then
                        If column.DataType Is GetType(String) Then
                            If CType(values(ValueCounter), String) = CType(rowValue, String) Then
                                RowsOkay.Add(RowCounter)
                                Exit For
                            End If
                        ElseIf column.DataType Is GetType(Int16) Then
                            If CType(values(ValueCounter), Int16) = CType(rowValue, Int16) Then
                                RowsOkay.Add(RowCounter)
                                Exit For
                            End If
                        ElseIf column.DataType Is GetType(Int32) Then
                            If CType(values(ValueCounter), Int32) = CType(rowValue, Int32) Then
                                RowsOkay.Add(RowCounter)
                                Exit For
                            End If
                        ElseIf column.DataType Is GetType(Int64) Then
                            If CType(values(ValueCounter), Int64) = CType(rowValue, Int64) Then
                                RowsOkay.Add(RowCounter)
                                Exit For
                            End If
                        ElseIf column.DataType Is GetType(Boolean) Then
                            If CType(values(ValueCounter), Boolean) = CType(rowValue, Boolean) Then
                                RowsOkay.Add(RowCounter)
                                Exit For
                            End If
                        ElseIf column.DataType Is GetType(UInt16) Then
                            If CType(values(ValueCounter), System.UInt16) = CType(rowValue, System.UInt16) Then
                                RowsOkay.Add(RowCounter)
                                Exit For
                            End If
                        ElseIf column.DataType Is GetType(UInt32) Then
                            If CType(values(ValueCounter), UInt32) = CType(rowValue, UInt32) Then
                                RowsOkay.Add(RowCounter)
                                Exit For
                            End If
                        ElseIf column.DataType Is GetType(UInt64) Then
                            If CType(values(ValueCounter), UInt64) = CType(rowValue, UInt64) Then
                                RowsOkay.Add(RowCounter)
                                Exit For
                            End If
                        ElseIf column.DataType Is GetType(TimeSpan) Then
                            If CType(values(ValueCounter), TimeSpan) = CType(rowValue, TimeSpan) Then
                                RowsOkay.Add(RowCounter)
                                Exit For
                            End If
                        ElseIf column.DataType Is GetType(Date) Then
                            If CType(values(ValueCounter), Date) = CType(rowValue, Date) Then
                                RowsOkay.Add(RowCounter)
                                Exit For
                            End If
                        ElseIf column.DataType Is GetType(Decimal) Then
                            If CType(values(ValueCounter), Decimal) = CType(rowValue, Decimal) Then
                                RowsOkay.Add(RowCounter)
                                Exit For
                            End If
                        ElseIf column.DataType Is GetType(Single) Then
                            If CType(values(ValueCounter), Single) = CType(rowValue, Single) Then
                                RowsOkay.Add(RowCounter)
                                Exit For
                            End If
                        ElseIf column.DataType Is GetType(Double) Then
                            If CType(values(ValueCounter), Double) = CType(rowValue, Double) Then
                                RowsOkay.Add(RowCounter)
                                Exit For
                            End If
                        Else
                            'object type
                            If values(ValueCounter) Is rowValue Then
                                RowsOkay.Add(RowCounter)
                                Exit For
                            End If
                        End If
                    End If
                Next
            Next
            'Now delete rows without mark
            For RowCounter As Integer = column.Table.Rows.Count - 1 To 0 Step -1
                If RowsOkay.Contains(RowCounter) = False Then
                    column.Table.Rows.RemoveAt(RowCounter)
                End If
            Next
        End Sub

        ''' <summary>
        ''' Remove rows from a table without any value in specified column
        ''' </summary>
        ''' <param name="column">The column whose values shall be verified</param>
        ''' <remarks></remarks>
        Public Shared Sub RemoveRowsWithDbNullValues(ByVal column As System.Data.DataColumn)
            RemoveRowsWithColumnValues(column, New Object() {DBNull.Value})
        End Sub

        ''' <summary>
        ''' Remove rows from a table with a given range of values in a defined column
        ''' </summary>
        ''' <param name="column">The column whose values shall be verified</param>
        ''' <param name="values">The values which lead to a removal of a row</param>
        ''' <remarks></remarks>
        Public Shared Sub RemoveRowsWithColumnValues(ByVal column As System.Data.DataColumn, ByVal values As Object())
            For RowCounter As Integer = column.Table.Rows.Count - 1 To 0 Step -1
                Dim rowValue As Object = column.Table.Rows(RowCounter)(column)
                For ValueCounter As Integer = 0 To values.Length - 1
                    Dim IgnoreValueChecks As Boolean = False
                    If IsDBNull(values(ValueCounter)) AndAlso IsDBNull(rowValue) Then
                        column.Table.Rows.RemoveAt(RowCounter)
                        Exit For
                    ElseIf Not IsDBNull(values(ValueCounter)) AndAlso IsDBNull(rowValue) Then
                        IgnoreValueChecks = True
                    ElseIf IsDBNull(values(ValueCounter)) AndAlso Not IsDBNull(rowValue) Then
                        IgnoreValueChecks = True
                    End If
                    If IgnoreValueChecks = False Then
                        If column.DataType Is GetType(String) Then
                            If CType(values(ValueCounter), String) = CType(rowValue, String) Then
                                column.Table.Rows.RemoveAt(RowCounter)
                                Exit For
                            End If
                        ElseIf column.DataType Is GetType(Int16) Then
                            If CType(values(ValueCounter), Int16) = CType(rowValue, Int16) Then
                                column.Table.Rows.RemoveAt(RowCounter)
                                Exit For
                            End If
                        ElseIf column.DataType Is GetType(Int32) Then
                            If CType(values(ValueCounter), Int32) = CType(rowValue, Int32) Then
                                column.Table.Rows.RemoveAt(RowCounter)
                                Exit For
                            End If
                        ElseIf column.DataType Is GetType(Int64) Then
                            If CType(values(ValueCounter), Int64) = CType(rowValue, Int64) Then
                                column.Table.Rows.RemoveAt(RowCounter)
                                Exit For
                            End If
                        ElseIf column.DataType Is GetType(Boolean) Then
                            If CType(values(ValueCounter), Boolean) = CType(rowValue, Boolean) Then
                                column.Table.Rows.RemoveAt(RowCounter)
                                Exit For
                            End If
                        ElseIf column.DataType Is GetType(UInt16) Then
                            If CType(values(ValueCounter), System.UInt16) = CType(rowValue, System.UInt16) Then
                                column.Table.Rows.RemoveAt(RowCounter)
                                Exit For
                            End If
                        ElseIf column.DataType Is GetType(UInt32) Then
                            If CType(values(ValueCounter), UInt32) = CType(rowValue, UInt32) Then
                                column.Table.Rows.RemoveAt(RowCounter)
                                Exit For
                            End If
                        ElseIf column.DataType Is GetType(UInt64) Then
                            If CType(values(ValueCounter), UInt64) = CType(rowValue, UInt64) Then
                                column.Table.Rows.RemoveAt(RowCounter)
                                Exit For
                            End If
                        ElseIf column.DataType Is GetType(TimeSpan) Then
                            If CType(values(ValueCounter), TimeSpan) = CType(rowValue, TimeSpan) Then
                                column.Table.Rows.RemoveAt(RowCounter)
                                Exit For
                            End If
                        ElseIf column.DataType Is GetType(Date) Then
                            If CType(values(ValueCounter), Date) = CType(rowValue, Date) Then
                                column.Table.Rows.RemoveAt(RowCounter)
                                Exit For
                            End If
                        ElseIf column.DataType Is GetType(Decimal) Then
                            If CType(values(ValueCounter), Decimal) = CType(rowValue, Decimal) Then
                                column.Table.Rows.RemoveAt(RowCounter)
                                Exit For
                            End If
                        ElseIf column.DataType Is GetType(Single) Then
                            If CType(values(ValueCounter), Single) = CType(rowValue, Single) Then
                                column.Table.Rows.RemoveAt(RowCounter)
                                Exit For
                            End If
                        ElseIf column.DataType Is GetType(Double) Then
                            If CType(values(ValueCounter), Double) = CType(rowValue, Double) Then
                                column.Table.Rows.RemoveAt(RowCounter)
                                Exit For
                            End If
                        Else
                            'object type
                            If values(ValueCounter) Is rowValue Then
                                column.Table.Rows.RemoveAt(RowCounter)
                                Exit For
                            End If
                        End If
                    End If
                Next
            Next
        End Sub

        ''' <summary>
        ''' Convert a column into another data type by using an own function (a delegate function) for converting the values
        ''' </summary>
        ''' <param name="column"></param>
        ''' <param name="newDataType"></param>
        ''' <param name="delegateForConversion"></param>
        ''' <remarks></remarks>
        Public Shared Sub ConvertColumnType(ByVal column As DataColumn, ByVal newDataType As Type, ByVal delegateForConversion As TypeConverter)
            If column.Table.Columns.CanRemove(column) = False Then
                Throw New Exception("A column shall be converted which can't be removed; replacement failed")
            End If
            Dim newCol As DataColumn = column.Table.Columns.Add(CompuMaster.Data.DataTables.LookupUniqueColumnName(column.Table, column.ColumnName), newDataType)
            'Copy column settings as far as possible
            newCol.ReadOnly = column.ReadOnly
            newCol.Unique = column.Unique
            newCol.MaxLength = column.MaxLength
            newCol.DateTimeMode = column.DateTimeMode
            newCol.ColumnMapping = column.ColumnMapping
            newCol.Prefix = column.Prefix
            newCol.Caption = column.Caption
            newCol.AllowDBNull = column.AllowDBNull
            'Convert the column content
            For MyCounter As Integer = 0 To column.Table.Rows.Count - 1
                If Not IsDBNull(column.Table.Rows(MyCounter)(column)) Then
                    column.Table.Rows(MyCounter)(newCol) = delegateForConversion.Invoke(column.Table.Rows(MyCounter)(column))
                End If
            Next
            'Remove the old column
            Dim oldColName As String = column.ColumnName
            column.Table.Columns.Remove(column)
            'Rename the new column to the old name
            newCol.ColumnName = oldColName
        End Sub

        ''' <summary>
        ''' A delegate function for converting values from one type into another type
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Delegate Function TypeConverter(ByVal value As Object) As Object

        ''' <summary>
        '''     Drop all columns except the required ones
        ''' </summary>
        ''' <param name="table">A data table containing some columns</param>
        ''' <param name="remainingColumns">A list of column names which shall not be removed</param>
        ''' <remarks>
        '''     If the list of the remaining columns contains some column names which are not existing, then those column names will be ignored. There will be no exception in this case.
        '''     The names of the columns are handled case-insensitive.
        ''' </remarks>
        Public Shared Sub KeepColumnsAndRemoveAllOthers(ByVal table As DataTable, ByVal remainingColumns As String())
            CompuMaster.Data.DataTablesTools.KeepColumnsAndRemoveAllOthers(table, remainingColumns)
        End Sub

        ''' <summary>
        ''' A list item which can be consumed by list controls in namespaces System.Windows
        ''' </summary>
        Public Class WinFormsListControlItem

            Private _Key As Object
            Public Property Key() As Object
                Get
                    Return _Key
                End Get
                Set(ByVal Value As Object)
                    _Key = Value
                End Set
            End Property

            Public Overrides Function ToString() As String
                If Value Is Nothing Then
                    Return String.Empty
                Else
                    Return Value.ToString
                End If
            End Function

            Private _Value As Object
            Public Property Value() As Object
                Get
                    Return _Value
                End Get
                Set(ByVal Value As Object)
                    _Value = Value
                End Set
            End Property

            Public Sub New()
            End Sub

            Public Sub New(ByVal key As Object, ByVal value As Object)
                Me.Key = key
                Me.Value = value
            End Sub

        End Class

        ''' <summary>
        '''     Lookup the row index for a data row in a data table
        ''' </summary>
        ''' <param name="dataRow">The data row whose index number is required</param>
        ''' <returns>An index number for the given data row</returns>
        Public Shared Function RowIndex(ByVal dataRow As DataRow) As Integer
            Return CompuMaster.Data.DataTablesTools.RowIndex(dataRow)
        End Function

        ''' <summary>
        '''     Lookup the column index for a data column in a data table
        ''' </summary>
        ''' <param name="column">The data column whose index number is required</param>
        ''' <returns>An index number for the given column</returns>
        Public Shared Function ColumnIndex(ByVal column As DataColumn) As Integer
            Return CompuMaster.Data.DataTablesTools.ColumnIndex(column)
        End Function

        ''' <summary>
        '''     Find duplicate values in a given row and calculate the number of occurances of each value in the table
        ''' </summary>
        ''' <param name="column">A column of a datatable</param>
        ''' <returns>A hashtable containing the origin column value as key and the number of occurances as value</returns>
        Public Shared Function FindDuplicates(ByVal column As DataColumn) As Hashtable
            Return CompuMaster.Data.DataTablesTools.FindDuplicates(column)
        End Function

        ''' <summary>
        '''     Find duplicate values in a given row and calculate the number of occurances of each value in the table
        ''' </summary>
        ''' <param name="column">A column of a datatable</param>
        ''' <param name="minOccurances">Only values with occurances equal or more than this number will be returned</param>
        ''' <returns>A hashtable containing the origin column value as key and the number of occurances as value</returns>
        Public Shared Function FindDuplicates(ByVal column As DataColumn, ByVal minOccurances As Integer) As Hashtable
            Return CompuMaster.Data.DataTablesTools.FindDuplicates(column, minOccurances)
        End Function

        ''' <summary>
        '''     Find duplicate values in a given row and calculate the number of occurances of each value in the table
        ''' </summary>
        ''' <param name="column">A column of a datatable</param>
        ''' <returns>A hashtable containing the origin column value as key and the number of occurances as value</returns>
        Public Shared Function FindDuplicates(Of t)(ByVal column As DataColumn) As System.Collections.Generic.Dictionary(Of t, Integer)
            Return CompuMaster.Data.DataTablesTools.FindDuplicates(Of t)(column)
        End Function

        ''' <summary>
        '''     Find duplicate values in a given row and calculate the number of occurances of each value in the table
        ''' </summary>
        ''' <param name="column">A column of a datatable</param>
        ''' <param name="minOccurances">Only values with occurances equal or more than this number will be returned</param>
        ''' <returns>A hashtable containing the origin column value as key and the number of occurances as value</returns>
        Public Shared Function FindDuplicates(Of t)(ByVal column As DataColumn, ByVal minOccurances As Integer) As System.Collections.Generic.Dictionary(Of t, Integer)
            Return CompuMaster.Data.DataTablesTools.FindDuplicates(Of t)(column, minOccurances)
        End Function

        ''' <summary>
        ''' Remove rows with duplicate values in a given column
        ''' </summary>
        ''' <param name="dataTable">A datatable with duplicate values</param>
        ''' <param name="columnName">column name of the datatable which contains the duplicate values</param>
        ''' <returns>A datatable with unique records in the specified column</returns>
        Public Shared Function RemoveDuplicates(ByVal dataTable As DataTable, ByVal columnName As String) As DataTable
            Return CompuMaster.Data.DataTablesTools.RemoveDuplicates(dataTable, columnName)
        End Function 'RemoveDuplicateRows

        ''' <summary>
        '''     Convert the first two columns into objects which can be consumed by the ListControl objects in the System.Windows.Forms namespaces
        ''' </summary>
        ''' <param name="datatable">The datatable which contains a key column and a value column for the list control</param>
        ''' <returns>An array of WinFormsListControlItem</returns>
        Public Shared Function ConvertDataTableToWinFormsListControlItem(ByVal datatable As DataTable) As WinFormsListControlItem()
            Dim Result As WinFormsListControlItem() = Nothing
            Dim Source As CompuMaster.Data.DataTablesTools.ListControlItem()
            Source = CompuMaster.Data.DataTablesTools.ConvertDataTableToListControlItem(datatable)
            If Source.Length > 0 Then
                ReDim Result(Source.Length - 1)
                For MyCounter As Integer = 0 To Source.Length - 1
                    Dim NewValue As New WinFormsListControlItem With {
                        .Key = Source(MyCounter).Key,
                        .Value = Source(MyCounter).Value
                    }
                    Result(MyCounter) = NewValue
                Next
            End If
            Return Result
        End Function

        ''' <summary>
        '''     Convert the first two columns into objects which can be consumed by the ListControl objects in the System.Web.WebControls namespaces
        ''' </summary>
        ''' <param name="datatable">The datatable which contains a key column and a value column for the list control</param>
        ''' <returns>An array of System.Web.UI.WebControls.ListItem for consumption in many list controls of the System.Web namespace</returns>
        Public Shared Function ConvertDataTableToWebFormsListItem(ByVal datatable As DataTable) As System.Web.UI.WebControls.ListItem()
            Dim Result As System.Web.UI.WebControls.ListItem() = Nothing
            Dim Source As CompuMaster.Data.DataTablesTools.ListControlItem()
            Source = CompuMaster.Data.DataTablesTools.ConvertDataTableToListControlItem(datatable)
            If Source.Length > 0 Then
                ReDim Result(Source.Length - 1)
                For MyCounter As Integer = 0 To Source.Length - 1
                    Dim NewValue As New System.Web.UI.WebControls.ListItem With {
                        .Value = CType(Source(MyCounter).Key, String),
                        .Text = CType(Source(MyCounter).Value, String)
                    }
                    Result(MyCounter) = NewValue
                Next
            End If
            Return Result
        End Function

        ''' <summary>
        '''     Convert a dataset to an xml string with data and schema information
        ''' </summary>
        ''' <param name="dataset"></param>
        ''' <returns></returns>
        Public Shared Function ConvertDatasetToXml(ByVal dataset As DataSet) As String
            Return CompuMaster.Data.DataTablesTools.ConvertDatasetToXml(dataset)
        End Function

        ''' <summary>
        '''     Convert an xml string to a dataset
        ''' </summary>
        ''' <param name="xml"></param>
        ''' <returns></returns>
        Public Shared Function ConvertXmlToDataset(ByVal xml As String) As DataSet
            Return CompuMaster.Data.DataTablesTools.ConvertXmlToDataset(xml)
        End Function

        ''' <summary>
        '''     Create a new data table clone with only some first rows
        ''' </summary>
        ''' <param name="SourceTable">The source table to be copied</param>
        ''' <param name="NumberOfRows">The number of rows to be copied</param>
        ''' <returns>The new clone of the datatable</returns>
        Public Shared Function CopyDataTableWithSubsetOfRows(ByVal SourceTable As DataTable, ByVal NumberOfRows As Integer) As DataTable
            Return CompuMaster.Data.DataTablesTools.GetDataTableWithSubsetOfRows(SourceTable, NumberOfRows)
        End Function

        ''' <summary>
        '''     Create a new data table clone with only some first rows
        ''' </summary>
        ''' <param name="SourceTable">The source table to be copied</param>
        ''' <param name="StartAtRow">The position where to start the copy process, the first row is at 0</param>
        ''' <param name="NumberOfRows">The number of rows to be copied</param>
        ''' <returns>The new clone of the datatable</returns>
        Public Shared Function CopyDataTableWithSubsetOfRows(ByVal SourceTable As DataTable, ByVal StartAtRow As Integer, ByVal NumberOfRows As Integer) As DataTable
            Return CompuMaster.Data.DataTablesTools.GetDataTableWithSubsetOfRows(SourceTable, StartAtRow, NumberOfRows)
        End Function

        ''' <summary>
        '''     Remove those rows in the source column which haven't got the same value in the compare table
        ''' </summary>
        ''' <param name="sourceColumn">This is the column of the master table where all operations shall be executed</param>
        ''' <param name="valuesMustExistInThisColumnToKeepTheSourceRow">This is the comparison value against the source table's column</param>
        ''' <returns>An arraylist of removed keys</returns>
        ''' <remarks>
        '''     Strings will be compared case-insensitive, DBNull values in the source table will always be removed
        '''     Attention: result of this function is not an arraylist containing keys!
        '''                result of this funciton is an arraylist containing object arrays of keys of removed rows!
        ''' </remarks>
        Public Shared Function RemoveRowsWithNoCorrespondingValueInComparisonTable(ByVal sourceColumn As DataColumn,
                                                                                   ByVal valuesMustExistInThisColumnToKeepTheSourceRow As DataColumn) As ArrayList
            Return RemoveRowsWithNoCorrespondingValueInComparisonTable(sourceColumn, valuesMustExistInThisColumnToKeepTheSourceRow, True, True)
        End Function

        'TO BE IMPLEMENTED: Version with multiple comparison columns
        '''' <summary>
        ''''     Remove those rows in the source column which haven't got the same value in the compare table
        '''' </summary>
        '''' <param name="sourceColumns">These are the columns of the master table where all operations shall be executed</param>
        '''' <param name="valuesMustExistInTheseColumnsToKeepTheSourceRow">These are the comparison values against the source table's columns</param>
        '''' <param name="ignoreCaseInStrings">Strings will be compared case-insensitive</param>
        '''' <param name="alwaysRemoveDBNullValues">Always remove the source row when it contains a DBNull value</param>
        '''' <returns>An arraylist with object arrays containing all key values of a row in the order of the source columns</returns>
        '''' <remarks>
        ''''     Attention: result of this function is not an arraylist containing keys!
        ''''                result of this funciton is an arraylist containing object arrays of keys of removed rows!
        '''' </remarks>
        'Friend Shared Function RemoveRowsWithNoCorrespondingValueInComparisonTable(ByVal sourceColumns As DataColumn(), ByVal valuesMustExistInTheseColumnsToKeepTheSourceRow As DataColumn(), ByVal ignoreCaseInStrings As Boolean, ByVal alwaysRemoveDBNullValues As Boolean) As System.Collections.Generic.List(Of Object())

        '    'parameters validation
        '    If sourceColumns Is Nothing Then
        '        Throw New ArgumentNullException("sourceColumns", "Required column: sourceColumns")
        '    ElseIf valuesMustExistInTheseColumnsToKeepTheSourceRow Is Nothing Then
        '        Throw New ArgumentNullException("valuesMustExistInTheseColumnsToKeepTheSourceRow", "Required column: valuesMustExistInTheseColumnsToKeepTheSourceRow")
        '    ElseIf sourceColumns.Length <> valuesMustExistInTheseColumnsToKeepTheSourceRow.Length Then
        '        Throw New ArgumentOutOfRangeException("Key definition of both tables must contain the same number of keys")
        '    Else
        '        'ToDo: additional testings
        '        '- Are table references of all source columns the same?
        '        If sourceColumns.Length > 1 Then
        '            For MyCounter As Integer = 1 To sourceColumns.Length - 1
        '                If sourceColumns(MyCounter).Table IsNot sourceColumns(0).Table Then
        '                    Throw New ArgumentException("sourceColumn", "All source columns must be related to the same table")
        '                End If
        '            Next
        '        End If
        '        '- Are table references of all comparison columns the same?
        '        If valuesMustExistInTheseColumnsToKeepTheSourceRow.Length > 1 Then
        '            For MyCounter As Integer = 1 To valuesMustExistInTheseColumnsToKeepTheSourceRow.Length - 1
        '                If valuesMustExistInTheseColumnsToKeepTheSourceRow(MyCounter).Table IsNot valuesMustExistInTheseColumnsToKeepTheSourceRow(0).Table Then
        '                    Throw New ArgumentException("valuesMustExistInTheseColumnsToKeepTheSourceRow", "All comparison columns must be related to the same table")
        '                End If
        '            Next
        '        End If
        '        '- Are all keys in the source table matching the same datatype as in the comparison table?
        '        '- Additional checks see already implemented functions
        '        For MyCounter As Integer = 0 To sourceColumns.Length - 1
        '            If Not sourceColumns(MyCounter).DataType Is valuesMustExistInTheseColumnsToKeepTheSourceRow(MyCounter).DataType Then
        '                Throw New InvalidCastException("Data type mismatch: both tables must use the same data types for the comparison columns: """ & sourceColumns(MyCounter).ColumnName & """ vs. """ & valuesMustExistInTheseColumnsToKeepTheSourceRow(MyCounter).ColumnName & """")
        '            End If
        '        Next
        '    End If

        '    'TODO: implementation

        'End Function

        ''' <summary>
        '''     Remove those rows in the source column which haven't got the same value in the compare table
        ''' </summary>
        ''' <param name="sourceColumn">This is the column of the master table where all operations shall be executed</param>
        ''' <param name="valuesMustExistInThisColumnToKeepTheSourceRow">This is the comparison value against the source table's column</param>
        ''' <param name="ignoreCaseInStrings">Strings will be compared case-insensitive</param>
        ''' <param name="alwaysRemoveDBNullValues">Always remove the source row when it contains a DBNull value</param>
        ''' <returns>An arraylist of removed keys</returns>
        ''' <remarks>
        '''     Attention: result of this function is not an arraylist containing keys!
        '''                result of this funciton is an arraylist containing object arrays of keys of removed rows!
        ''' </remarks>
        Public Shared Function RemoveRowsWithNoCorrespondingValueInComparisonTable(ByVal sourceColumn As DataColumn,
                                                                                   ByVal valuesMustExistInThisColumnToKeepTheSourceRow As DataColumn,
                                                                                   ByVal ignoreCaseInStrings As Boolean,
                                                                                   ByVal alwaysRemoveDBNullValues As Boolean) As ArrayList

            'parameters validation
            If sourceColumn Is Nothing Then
                Throw New ArgumentNullException(NameOf(sourceColumn), "Required column: sourceColumn")
            ElseIf valuesMustExistInThisColumnToKeepTheSourceRow Is Nothing Then
                Throw New ArgumentNullException(NameOf(valuesMustExistInThisColumnToKeepTheSourceRow), "Required column: valuesMustExistInThisColumnToKeepTheSourceRow")
            ElseIf sourceColumn.DataType IsNot valuesMustExistInThisColumnToKeepTheSourceRow.DataType Then
                Throw New InvalidCastException("Data type mismatch: both tables must use the same data types for the comparison columns")
            End If

            'Prepare local variables
            Dim Result As New ArrayList 'Contains all keys which have been removed
            Dim sourceTable As DataTable = sourceColumn.Table
            Dim comparisonTable As DataTable = valuesMustExistInThisColumnToKeepTheSourceRow.Table

            'Loop through the source table and try to find matches in the comparison table
            For MyCounter As Integer = sourceTable.Rows.Count - 1 To 0 Step -1
                Dim MatchFound As Boolean = False
                If sourceColumn.DataType Is GetType(String) Then
                    'Compare strings
                    For MyCompCounter As Integer = 0 To comparisonTable.Rows.Count - 1
                        If IsDBNull(sourceTable.Rows(MyCounter)(sourceColumn)) Then
                            If alwaysRemoveDBNullValues Then
                                'Remove this line from source table because it contains a DBNull and those rows shall be removed
                                MatchFound = False
                                Exit For
                            Else
                                If IsDBNull(comparisonTable.Rows(MyCompCounter)(valuesMustExistInThisColumnToKeepTheSourceRow)) Then
                                    'This is a match, keep that row!
                                    MatchFound = True
                                    Exit For
                                Else
                                    'Not identical, continue search
                                End If
                            End If
                        ElseIf IsDBNull(comparisonTable.Rows(MyCompCounter)(valuesMustExistInThisColumnToKeepTheSourceRow)) Then
                            'Not identical, continue search
                        ElseIf String.Compare(CType(sourceTable.Rows(MyCounter)(sourceColumn), String), CType(comparisonTable.Rows(MyCompCounter)(valuesMustExistInThisColumnToKeepTheSourceRow), String), ignoreCaseInStrings, System.Globalization.CultureInfo.InvariantCulture) = 0 Then
                            'Case insensitive comparison resulted to successful match
                            MatchFound = True
                            Exit For
                        Else
                            'Not identical, continue search
                        End If
                    Next
                ElseIf sourceColumn.DataType.IsValueType Then
                    'Compare value types
                    For MyCompCounter As Integer = 0 To comparisonTable.Rows.Count - 1
                        If IsDBNull(sourceTable.Rows(MyCounter)(sourceColumn)) Then
                            If alwaysRemoveDBNullValues Then
                                'Remove this line from source table because it contains a DBNull and those rows shall be removed
                                MatchFound = False
                                Exit For
                            Else
                                If IsDBNull(comparisonTable.Rows(MyCompCounter)(valuesMustExistInThisColumnToKeepTheSourceRow)) Then
                                    'This is a match, keep that row!
                                    MatchFound = True
                                    Exit For
                                Else
                                    'Not identical, continue search
                                End If
                            End If
                        ElseIf IsDBNull(comparisonTable.Rows(MyCompCounter)(valuesMustExistInThisColumnToKeepTheSourceRow)) Then
                            'Not identical, continue search
                        ElseIf CType(sourceTable.Rows(MyCounter)(sourceColumn), ValueType).Equals(comparisonTable.Rows(MyCompCounter)(valuesMustExistInThisColumnToKeepTheSourceRow)) Then
                            'Values are equal
                            MatchFound = True
                            Exit For
                        Else
                            'Not identical, continue search
                        End If
                    Next
                ElseIf sourceColumn.DataType.IsValueType = False Then
                    'Compare objects
                    For MyCompCounter As Integer = 0 To comparisonTable.Rows.Count - 1
                        If IsDBNull(sourceTable.Rows(MyCounter)(sourceColumn)) Then
                            If alwaysRemoveDBNullValues Then
                                'Remove this line from source table because it contains a DBNull and those rows shall be removed
                                MatchFound = False
                                Exit For
                            Else
                                If IsDBNull(comparisonTable.Rows(MyCompCounter)(valuesMustExistInThisColumnToKeepTheSourceRow)) Then
                                    'This is a match, keep that row!
                                    MatchFound = True
                                    Exit For
                                Else
                                    'Not identical, continue search
                                End If
                            End If
                        ElseIf IsDBNull(comparisonTable.Rows(MyCompCounter)(valuesMustExistInThisColumnToKeepTheSourceRow)) Then
                            'Not identical, continue search
                        ElseIf sourceTable.Rows(MyCounter)(sourceColumn) Is comparisonTable.Rows(MyCompCounter)(valuesMustExistInThisColumnToKeepTheSourceRow) Then
                            'Objects are the same
                            MatchFound = True
                            Exit For
                        Else
                            'Not identical, continue search
                        End If
                    Next
                End If
                If MatchFound = False Then
                    'Add the key of the row to the result list
                    Result.Insert(0, sourceTable.Rows(MyCounter)(sourceColumn))
                    'No match found leads to removal of the row in the source table
                    sourceTable.Rows.RemoveAt(MyCounter)
                End If
            Next
            Return Result
        End Function

        ''' <summary>
        '''     Remove those rows in the source column which haven't got the same value in the compare table
        ''' </summary>
        ''' <param name="sourceColumn">This is the column of the master table where all operations shall be executed</param>
        ''' <param name="valuesMustExistInThisColumnToKeepTheSourceRow">This is the comparison value against the source table's column</param>
        ''' <returns>An arraylist of removed keys</returns>
        ''' <remarks>
        '''     Strings will be compared case-insensitive, DBNull values in the source table will always be removed
        '''     Attention: result of this function is not an arraylist containing keys!
        '''                result of this funciton is an arraylist containing object arrays of keys of removed rows!
        ''' </remarks>
        Public Shared Function RemoveRowsWithCorrespondingValueInComparisonTable(ByVal sourceColumn As DataColumn,
                                                                                   ByVal valuesMustExistInThisColumnToKeepTheSourceRow As DataColumn) As ArrayList
            Return RemoveRowsWithCorrespondingValueInComparisonTable(sourceColumn, valuesMustExistInThisColumnToKeepTheSourceRow, True, True)
        End Function

        ''' <summary>
        '''     Remove those rows in the source column which have got the same value in the compare table
        ''' </summary>
        ''' <param name="sourceColumn">This is the column of the master table where all operations shall be executed</param>
        ''' <param name="valuesMustExistInThisColumnToKeepTheSourceRow">This is the comparison value against the source table's column</param>
        ''' <param name="ignoreCaseInStrings">Strings will be compared case-insensitive</param>
        ''' <param name="alwaysRemoveDBNullValues">Always remove the source row when it contains a DBNull value</param>
        ''' <returns>An arraylist of removed keys</returns>
        ''' <remarks>
        '''     Attention: result of this function is not an arraylist containing keys!
        '''                result of this funciton is an arraylist containing object arrays of keys of removed rows!
        ''' </remarks>
        Public Shared Function RemoveRowsWithCorrespondingValueInComparisonTable(ByVal sourceColumn As DataColumn,
                                                                                   ByVal valuesMustExistInThisColumnToKeepTheSourceRow As DataColumn,
                                                                                   ByVal ignoreCaseInStrings As Boolean,
                                                                                   ByVal alwaysRemoveDBNullValues As Boolean) As ArrayList

            'parameters validation
            If sourceColumn Is Nothing Then
                Throw New ArgumentNullException(NameOf(sourceColumn), "Required column: sourceColumn")
            ElseIf valuesMustExistInThisColumnToKeepTheSourceRow Is Nothing Then
                Throw New ArgumentNullException(NameOf(valuesMustExistInThisColumnToKeepTheSourceRow), "Required column: valuesMustExistInThisColumnToKeepTheSourceRow")
            ElseIf sourceColumn.DataType IsNot valuesMustExistInThisColumnToKeepTheSourceRow.DataType Then
                Throw New InvalidCastException("Data type mismatch: both tables must use the same data types for the comparison columns")
            End If

            'Prepare local variables
            Dim Result As New ArrayList 'Contains all keys which have been removed
            Dim sourceTable As DataTable = sourceColumn.Table
            Dim comparisonTable As DataTable = valuesMustExistInThisColumnToKeepTheSourceRow.Table

            'Loop through the source table and try to find matches in the comparison table
            For MyCounter As Integer = sourceTable.Rows.Count - 1 To 0 Step -1
                Dim MatchFound As Boolean = False
                If sourceColumn.DataType Is GetType(String) Then
                    'Compare strings
                    For MyCompCounter As Integer = 0 To comparisonTable.Rows.Count - 1
                        If IsDBNull(sourceTable.Rows(MyCounter)(sourceColumn)) Then
                            If alwaysRemoveDBNullValues Then
                                'Remove this line from source table because it contains a DBNull and those rows shall be removed
                                MatchFound = True
                                Exit For
                            Else
                                If IsDBNull(comparisonTable.Rows(MyCompCounter)(valuesMustExistInThisColumnToKeepTheSourceRow)) Then
                                    'This is a match, drop that row!
                                    MatchFound = True
                                    Exit For
                                Else
                                    'Not identical, continue search
                                End If
                            End If
                        ElseIf IsDBNull(comparisonTable.Rows(MyCompCounter)(valuesMustExistInThisColumnToKeepTheSourceRow)) Then
                            'Not identical, continue search
                        ElseIf String.Compare(CType(sourceTable.Rows(MyCounter)(sourceColumn), String), CType(comparisonTable.Rows(MyCompCounter)(valuesMustExistInThisColumnToKeepTheSourceRow), String), ignoreCaseInStrings, System.Globalization.CultureInfo.InvariantCulture) = 0 Then
                            'Case insensitive comparison resulted to successful match
                            MatchFound = True
                            Exit For
                        Else
                            'Not identical, continue search
                        End If
                    Next
                ElseIf sourceColumn.DataType.IsValueType Then
                    'Compare value types
                    For MyCompCounter As Integer = 0 To comparisonTable.Rows.Count - 1
                        If IsDBNull(sourceTable.Rows(MyCounter)(sourceColumn)) Then
                            If alwaysRemoveDBNullValues Then
                                'Remove this line from source table because it contains a DBNull and those rows shall be removed
                                MatchFound = True
                                Exit For
                            Else
                                If IsDBNull(comparisonTable.Rows(MyCompCounter)(valuesMustExistInThisColumnToKeepTheSourceRow)) Then
                                    'This is a match, drop that row!
                                    MatchFound = True
                                    Exit For
                                Else
                                    'Not identical, continue search
                                End If
                            End If
                        ElseIf IsDBNull(comparisonTable.Rows(MyCompCounter)(valuesMustExistInThisColumnToKeepTheSourceRow)) Then
                            'Not identical, continue search
                        ElseIf CType(sourceTable.Rows(MyCounter)(sourceColumn), ValueType).Equals(comparisonTable.Rows(MyCompCounter)(valuesMustExistInThisColumnToKeepTheSourceRow)) Then
                            'Values are equal
                            MatchFound = True
                            Exit For
                        Else
                            'Not identical, continue search
                        End If
                    Next
                ElseIf sourceColumn.DataType.IsValueType = False Then
                    'Compare objects
                    For MyCompCounter As Integer = 0 To comparisonTable.Rows.Count - 1
                        If IsDBNull(sourceTable.Rows(MyCounter)(sourceColumn)) Then
                            If alwaysRemoveDBNullValues Then
                                'Remove this line from source table because it contains a DBNull and those rows shall be removed
                                MatchFound = True
                                Exit For
                            Else
                                If IsDBNull(comparisonTable.Rows(MyCompCounter)(valuesMustExistInThisColumnToKeepTheSourceRow)) Then
                                    'This is a match, drop that row!
                                    MatchFound = True
                                    Exit For
                                Else
                                    'Not identical, continue search
                                End If
                            End If
                        ElseIf IsDBNull(comparisonTable.Rows(MyCompCounter)(valuesMustExistInThisColumnToKeepTheSourceRow)) Then
                            'Not identical, continue search
                        ElseIf sourceTable.Rows(MyCounter)(sourceColumn) Is comparisonTable.Rows(MyCompCounter)(valuesMustExistInThisColumnToKeepTheSourceRow) Then
                            'Objects are the same
                            MatchFound = True
                            Exit For
                        Else
                            'Not identical, continue search
                        End If
                    Next
                End If
                If MatchFound = True Then
                    'Add the key of the row to the result list
                    Result.Insert(0, sourceTable.Rows(MyCounter)(sourceColumn))
                    'No match found leads to removal of the row in the source table
                    sourceTable.Rows.RemoveAt(MyCounter)
                End If
            Next
            Return Result
        End Function

        ''' <summary>
        '''     Creates a complete clone of a DataRow with structure as well as data
        ''' </summary>
        ''' <param name="sourceRow">The source row to be copied</param>
        ''' <returns>The new clone of the DataRow</returns>
        ''' <remarks>
        '''     The resulting DataRow has got the schema from the sourceRow's DataTable.
        ''' </remarks>
        Public Shared Function CreateDataRowClone(ByVal sourceRow As DataRow) As DataRow
            Return CompuMaster.Data.DataTablesTools.CreateDataRowClone(sourceRow)
        End Function

        ''' <summary>
        ''' Create a table union from 2 tables
        ''' </summary>
        ''' <param name="firstTable"></param>
        ''' <param name="secondTable"></param>
        ''' <returns></returns>
        Public Shared Function UnionDataTables(ByVal firstTable As DataTable, secondTable As DataTable) As DataTable
            Dim Result As DataTable = CompuMaster.Data.DataTablesTools.GetDataTableClone(firstTable)
            CreateDataTableClone(secondTable, Result, String.Empty)
            Return Result
        End Function

        ''' <summary>
        '''     Creates a complete clone of a DataTable with structure as well as data
        ''' </summary>
        ''' <param name="SourceTable">The source table to be copied</param>
        ''' <returns>The new clone of the datatable</returns>
        Public Shared Function CreateDataTableClone(ByVal SourceTable As DataTable) As DataTable
            Return CompuMaster.Data.DataTablesTools.GetDataTableClone(SourceTable)
        End Function

        ''' <summary>
        '''     Creates a complete clone of a DataTable with structure as well as data
        ''' </summary>
        ''' <param name="SourceTable">The source table to be copied</param>
        ''' <param name="sourceRowFilter">An additional row filter for the source table, for all rows set it to null (Nothing in VisualBasic)</param>
        ''' <returns>The new clone of the datatable</returns>
        Public Shared Function CreateDataTableClone(ByVal SourceTable As DataTable, ByVal sourceRowFilter As String) As DataTable
            Return CompuMaster.Data.DataTablesTools.GetDataTableClone(SourceTable, sourceRowFilter)
        End Function

        ''' <summary>
        '''     Creates a complete clone of a DataTable with structure as well as data
        ''' </summary>
        ''' <param name="SourceTable">The source table to be copied</param>
        ''' <param name="sourceRowFilter">An additional row filter for the source table, for all rows set it to null (Nothing in VisualBasic)</param>
        ''' <param name="sourceSortExpression">An additional sort command for the source table</param>
        ''' <returns>The new clone of the datatable</returns>
        Public Shared Function CreateDataTableClone(ByVal SourceTable As DataTable, ByVal sourceRowFilter As String, ByVal sourceSortExpression As String) As DataTable
            Return CompuMaster.Data.DataTablesTools.GetDataTableClone(SourceTable, sourceRowFilter, sourceSortExpression)
        End Function

        ''' <summary>
        '''     Creates a complete clone of a DataTable with structure as well as data
        ''' </summary>
        ''' <param name="SourceTable">The source table to be copied</param>
        ''' <param name="sourceRowFilter">An additional row filter for the source table, for all rows set it to null (Nothing in VisualBasic)</param>
        ''' <param name="sourceSortExpression">An additional sort command for the source table</param>
        ''' <param name="topRows">After row filtering, how many rows from top shall be returned as maximum? (0 = all rows)</param>
        ''' <returns>The new clone of the datatable</returns>
        Public Shared Function CreateDataTableClone(ByVal SourceTable As DataTable, ByVal sourceRowFilter As String, ByVal sourceSortExpression As String,
                                                    ByVal topRows As Integer) As DataTable
            Return CompuMaster.Data.DataTablesTools.GetDataTableClone(SourceTable, sourceRowFilter, sourceSortExpression, topRows)
        End Function

        ''' <summary>
        ''' Clear destination table rows, remove columns not existing in source table and copy rows/columns from source table into destination table
        ''' </summary>
        ''' <param name="sourceTable">The source table to be copied</param>
        ''' <param name="destinationTable">The destination of all operations; the destination table will be a clone of the source table at the end</param>
        ''' <param name="sourceRowFilter">An additional row filter for the source table. For all rows set it to null (Nothing in VisualBasic)</param>
        ''' <param name="sourceSortExpression">An additional sort command for the source table</param>
        ''' <param name="topRows">After row filtering, how many rows from top shall be returned as maximum? (0 = all rows)</param>
        ''' <param name="overwritePropertiesOfExistingColumns">Shall the data type or any other settings of an existing table be modified to match the source column's definition?</param>
        ''' <remarks>
        '''     All rows of the destination table will be removed, first.
        ''' </remarks>
        Public Shared Sub CreateDataTableClone(ByVal sourceTable As DataTable, ByVal destinationTable As DataTable, ByVal sourceRowFilter As String,
                                               ByVal sourceSortExpression As String, ByVal topRows As Integer, ByVal overwritePropertiesOfExistingColumns As Boolean)
            CreateDataTableClone(sourceTable, destinationTable, sourceRowFilter, sourceSortExpression, topRows, overwritePropertiesOfExistingColumns, True, True)
        End Sub

        ''' <summary>
        ''' Copy a source table into destination table while preserving existing column data types, preserving content and adding rows/columns as required
        ''' </summary>
        ''' <param name="sourceTable">The source table to be copied</param>
        ''' <param name="destinationTable">The destination of all operations; the destination table will be a clone of the source table at the end</param>
        ''' <param name="sourceRowFilter">An additional row filter for the source table. For all rows set it to null (Nothing in VisualBasic)</param>
        ''' <remarks>
        ''' </remarks>
        Public Shared Sub CreateDataTableClone(ByVal sourceTable As DataTable, ByVal destinationTable As DataTable, ByVal sourceRowFilter As String)
            CreateDataTableClone(sourceTable, destinationTable, sourceRowFilter, Nothing, 0, False, False, False)
        End Sub

        ''' <summary>
        '''     Creates a complete clone of a DataTable with structure as well as data
        ''' </summary>
        ''' <param name="sourceTable">The source table to be copied</param>
        ''' <param name="destinationTable">The destination of all operations; the destination table will be a clone of the source table at the end</param>
        ''' <param name="sourceRowFilter">An additional row filter for the source table. For all rows set it to null (Nothing in VisualBasic)</param>
        ''' <param name="sourceSortExpression">An additional sort command for the source table</param>
        ''' <param name="topRows">After row filtering, how many rows from top shall be returned as maximum? (0 = all rows)</param>
        ''' <param name="overwritePropertiesOfExistingColumns">Shall the data type or any other settings of an existing table be modified to match the source column's definition?</param>
        ''' <param name="dropExistingRowsInDestinationTable">Remove the existing rows of the destination table, first</param>
        ''' <param name="removeUnusedColumnsFromDestinationTable">Remove the existing columns of the destination table which are not present in the source table</param>
        Public Shared Sub CreateDataTableClone(ByVal sourceTable As DataTable, ByVal destinationTable As DataTable, ByVal sourceRowFilter As String,
                                               ByVal sourceSortExpression As String, ByVal topRows As Integer, ByVal overwritePropertiesOfExistingColumns As Boolean,
                                               ByVal dropExistingRowsInDestinationTable As Boolean, ByVal removeUnusedColumnsFromDestinationTable As Boolean)
            CreateDataTableClone(sourceTable, destinationTable, sourceRowFilter, sourceSortExpression, topRows, overwritePropertiesOfExistingColumns, dropExistingRowsInDestinationTable,
                                 removeUnusedColumnsFromDestinationTable, False, False)
        End Sub

        ''' <summary>
        '''     Creates a complete clone of a DataTable with structure as well as data
        ''' </summary>
        ''' <param name="sourceTable">The source table to be copied</param>
        ''' <param name="destinationTable">The destination of all operations; the destination table will be a clone of the source table at the end</param>
        ''' <param name="sourceRowFilter">An additional row filter for the source table. For all rows set it to null (Nothing in VisualBasic)</param>
        ''' <param name="sourceSortExpression">An additional sort command for the source table</param>
        ''' <param name="topRows">After row filtering, how many rows from top shall be returned as maximum? (0 = all rows)</param>
        ''' <param name="overwritePropertiesOfExistingColumns">Shall the data type or any other settings of an existing table be modified to match the source column's definition?</param>
        ''' <param name="dropExistingRowsInDestinationTable">Remove the existing rows of the destination table, first</param>
        ''' <param name="removeUnusedColumnsFromDestinationTable">Remove the existing columns of the destination table which are not present in the source table</param>
        ''' <param name="dontExtendSchemaOfDestinatonTable">If true: do not add columns from the source table not existing in the destination table.</param>
        Public Shared Sub CreateDataTableClone(ByVal sourceTable As DataTable, ByVal destinationTable As DataTable, ByVal sourceRowFilter As String,
                                               ByVal sourceSortExpression As String, ByVal topRows As Integer, ByVal overwritePropertiesOfExistingColumns As Boolean,
                                               ByVal dropExistingRowsInDestinationTable As Boolean, ByVal removeUnusedColumnsFromDestinationTable As Boolean,
                                               ByVal dontExtendSchemaOfDestinatonTable As Boolean)
            CreateDataTableClone(sourceTable, destinationTable, sourceRowFilter, sourceSortExpression, topRows, overwritePropertiesOfExistingColumns, dropExistingRowsInDestinationTable,
                                 removeUnusedColumnsFromDestinationTable, dontExtendSchemaOfDestinatonTable, False)
        End Sub

        ''' <summary>
        '''     Creates a complete clone of a DataTable with structure as well as data
        ''' </summary>
        ''' <param name="sourceTable">The source table to be copied</param>
        ''' <param name="destinationTable">The destination of all operations; the destination table will be a clone of the source table at the end</param>
        ''' <param name="sourceRowFilter">An additional row filter for the source table. For all rows set it to null (Nothing in VisualBasic)</param>
        ''' <param name="sourceSortExpression">An additional sort command for the source table</param>
        ''' <param name="topRows">After row filtering, how many rows from top shall be returned as maximum? (0 = all rows)</param>
        ''' <param name="overwritePropertiesOfExistingColumns">Shall the data type or any other settings of an existing table be modified to match the source column's definition?</param>
        ''' <param name="dropExistingRowsInDestinationTable">Remove the existing rows of the destination table, first</param>
        ''' <param name="removeUnusedColumnsFromDestinationTable">Remove the existing columns of the destination table which are not present in the source table</param>
        ''' <param name="dontExtendSchemaOfDestinatonTable">If true: don't add columns from the source table not existing in the destination table.</param>
        ''' <param name="caseInsensitiveColumnNames">Specifies whether case insensitivity should matter for column names</param>
        Public Shared Sub CreateDataTableClone(ByVal sourceTable As DataTable, ByVal destinationTable As DataTable, ByVal sourceRowFilter As String,
                                               ByVal sourceSortExpression As String, ByVal topRows As Integer, ByVal overwritePropertiesOfExistingColumns As Boolean,
                                               ByVal dropExistingRowsInDestinationTable As Boolean,
                                               ByVal removeUnusedColumnsFromDestinationTable As Boolean,
                                               ByVal dontExtendSchemaOfDestinatonTable As Boolean, ByVal caseInsensitiveColumnNames As Boolean)
            Dim destinationSchemaChangesForUnusedColumns As RequestedSchemaChangesForUnusedColumns
            Dim destinationSchemaChangesForExistingColumns As RequestedSchemaChangesForExistingColumns
            Dim destinationSchemaChangesForAdditionalColumns As RequestedSchemaChangesForAdditionalColumns
            Dim rowChanges As RequestedRowChanges

            If overwritePropertiesOfExistingColumns Then
                destinationSchemaChangesForExistingColumns = RequestedSchemaChangesForExistingColumns.Update
            End If

            If dropExistingRowsInDestinationTable Then
                rowChanges = RequestedRowChanges.DropExistingRowsInDestinationTableAndInsertNewRows
            Else
                rowChanges = RequestedRowChanges.KeepExistingRowsInDestinationTableAndInsertNewRows
            End If

            If dontExtendSchemaOfDestinatonTable Then
                destinationSchemaChangesForAdditionalColumns = RequestedSchemaChangesForAdditionalColumns.None
            Else
                destinationSchemaChangesForAdditionalColumns = RequestedSchemaChangesForAdditionalColumns.Add
            End If

            If removeUnusedColumnsFromDestinationTable Then
                destinationSchemaChangesForUnusedColumns = RequestedSchemaChangesForUnusedColumns.Remove
            End If

            CreateDataTableClone(sourceTable, destinationTable, sourceRowFilter, sourceSortExpression, topRows, rowChanges, caseInsensitiveColumnNames,
                                 destinationSchemaChangesForUnusedColumns, destinationSchemaChangesForExistingColumns, destinationSchemaChangesForAdditionalColumns)

        End Sub

        Public Enum RequestedSchemaChangesForUnusedColumns As Byte
            None = 0
            ''' <summary>
            ''' Remove columns from the destination table not existing in the source table
            ''' </summary>
            ''' <remarks></remarks>
            Remove = 1
        End Enum

        Public Enum RequestedSchemaChangesForExistingColumns As Byte
            None = 0
            ''' <summary>
            ''' Column properties like datatype will be changed to match with the source column properties (Attention: data conversion might throw conversion exceptions!)
            ''' </summary>
            ''' <remarks></remarks>
            Update = 1
        End Enum

        Public Enum RequestedSchemaChangesForAdditionalColumns As Byte
            None = 0
            ''' <summary>
            ''' Add missing columns in the destination table which exist in the source table
            ''' </summary>
            ''' <remarks></remarks>
            Add = 1
        End Enum

        Public Enum RequestedRowChanges As Byte
            ''' <summary>
            ''' All rows of the destination table will be kept, rows from the source table will be added
            ''' </summary>
            ''' <remarks></remarks>
            KeepExistingRowsInDestinationTableAndInsertNewRows = 0
            ''' <summary>
            ''' All rows of the destination table will be removed, rows from the source table will be added
            ''' </summary>
            ''' <remarks>This behaviour can be considered as a &quot;replacing&quot; method.</remarks>
            DropExistingRowsInDestinationTableAndInsertNewRows = 1
            ''' <summary>
            ''' Update, delete and insert rows to match the source table's row collection. In other words: perform a merge.
            ''' </summary>
            ''' <remarks>After merging, the destination table will have an exact copy of the row collection of the source table. Please note: this value doesn't affect the column collection. Changes in the destination table done before merging won't be preserved</remarks>
            KeepExistingRowsInDestinationTableAndAddRemoveUpdateChangedRows = 2
        End Enum

        ''' <summary>
        '''     Creates a complete clone of a DataTable with structure as well as data
        ''' </summary>
        ''' <param name="sourceTable">The source table to be copied</param>
        ''' <param name="destinationTable">The destination of all operations; the destination table will be a clone of the source table at the end</param>
        ''' <param name="sourceRowFilter">An additional row filter for the source table. For all rows set it to null (Nothing in VisualBasic)</param>
        ''' <param name="sourceSortExpression">An additional sort command for the source table</param>
        ''' <param name="topRows">After row filtering/merging, how many rows from top shall be returned as maximum? (0 = all rows)</param>
        ''' <param name="rowChanges">Enum specifing the changes to be performed on the destination row </param>
        ''' <param name="caseInsensitiveColumnNames">Specifies whether case insensitivity should matter for column names</param>
        ''' <param name="destinationSchemaChangesForUnusedColumns">Remove the existing columns of the destination table which are not present in the source table</param>
        ''' <param name="destinationSchemaChangesForExistingColumns">If true: do not add columns from the source table not existing in the destination table.</param>
        ''' <param name="destinationSchemaChangesForAdditionalColumns">Specifies if we should compare columns case insensitive when we check whether all columns exist in the destination table. This parameter has no effect if the previous is true.</param>
        Public Shared Sub CreateDataTableClone(ByVal sourceTable As DataTable, ByVal destinationTable As DataTable, ByVal sourceRowFilter As String,
                                               ByVal sourceSortExpression As String, ByVal topRows As Integer, rowChanges As RequestedRowChanges,
                                               caseInsensitiveColumnNames As Boolean,
                                               destinationSchemaChangesForUnusedColumns As RequestedSchemaChangesForUnusedColumns,
                                               destinationSchemaChangesForExistingColumns As RequestedSchemaChangesForExistingColumns,
                                               destinationSchemaChangesForAdditionalColumns As RequestedSchemaChangesForAdditionalColumns)

            If RequestedRowChanges.DropExistingRowsInDestinationTableAndInsertNewRows = rowChanges Then
                'Drop existing rows
                For MyRowCounter As Integer = destinationTable.Rows.Count - 1 To 0 Step -1
                    destinationTable.Rows(MyRowCounter).Delete()
                Next
            End If

            'Define column set of destination table to column set of source table
            If RequestedSchemaChangesForUnusedColumns.Remove = destinationSchemaChangesForUnusedColumns Then
                '1. Remove columns not required anymore
                For MyDestTableCounter As Integer = destinationTable.Columns.Count - 1 To 0 Step -1
                    Dim columnExistsInSource As Boolean = False
                    For MySourceTableCounter As Integer = 0 To sourceTable.Columns.Count - 1
                        If destinationTable.Columns(MyDestTableCounter).ColumnName = sourceTable.Columns(MySourceTableCounter).ColumnName Then
                            columnExistsInSource = True
                        End If
                    Next
                    If columnExistsInSource = False Then
                        destinationTable.Columns.RemoveAt(MyDestTableCounter)
                    End If
                Next
            End If

            '2. Update existing, matching columns to be of the same data type
            If RequestedSchemaChangesForExistingColumns.Update = destinationSchemaChangesForExistingColumns Then
                For MyDestTableCounter As Integer = 0 To destinationTable.Columns.Count - 1
                    For MySourceTableCounter As Integer = 0 To sourceTable.Columns.Count - 1

                        Dim sourceTableColumn As DataColumn = sourceTable.Columns(MySourceTableCounter)
                        Dim destTableColumn As DataColumn = destinationTable.Columns(MyDestTableCounter)

                        If destTableColumn.ColumnName = sourceTableColumn.ColumnName Then
                            destTableColumn.AllowDBNull = sourceTableColumn.AllowDBNull
                            destTableColumn.AutoIncrement = sourceTableColumn.AutoIncrement
                            destTableColumn.AutoIncrementSeed = sourceTableColumn.AutoIncrementSeed
                            destTableColumn.AutoIncrementStep = sourceTableColumn.AutoIncrementStep
                            destTableColumn.Caption = sourceTableColumn.Caption
                            destTableColumn.ColumnMapping = sourceTableColumn.ColumnMapping
                            destTableColumn.DataType = sourceTableColumn.DataType
                            destTableColumn.DefaultValue = sourceTableColumn.DefaultValue
                            destTableColumn.ExtendedProperties.Clear()
                            For Each key As Object In sourceTableColumn.ExtendedProperties
                                destTableColumn.ExtendedProperties.Add(key, sourceTable.Columns(MySourceTableCounter).ExtendedProperties(key))
                            Next
                            destTableColumn.MaxLength = sourceTableColumn.MaxLength
                            destTableColumn.Namespace = sourceTableColumn.Namespace
                            destTableColumn.Prefix = sourceTableColumn.Prefix
                            destTableColumn.ReadOnly = sourceTableColumn.ReadOnly
                            destTableColumn.Unique = sourceTableColumn.Unique
                            If RequestedRowChanges.KeepExistingRowsInDestinationTableAndAddRemoveUpdateChangedRows <> rowChanges AndAlso Array.IndexOf(destinationTable.PrimaryKey, destinationTable.Columns) <> -1 Then
                                destTableColumn.Expression = sourceTableColumn.Expression
                            End If
                        End If
                    Next
                Next
            End If
            '3. Add missing columns
            If RequestedSchemaChangesForAdditionalColumns.Add = destinationSchemaChangesForAdditionalColumns Then
                For MySourceTableCounter As Integer = 0 To sourceTable.Columns.Count - 1
                    Dim columnExistsInDestination As Boolean = False
                    For MyDestTableCounter As Integer = 0 To destinationTable.Columns.Count - 1
                        If caseInsensitiveColumnNames Then
                            If destinationTable.Columns(MyDestTableCounter).ColumnName.ToUpper = sourceTable.Columns(MySourceTableCounter).ColumnName.ToUpper Then
                                columnExistsInDestination = True
                                Exit For
                            End If
                        Else
                            If destinationTable.Columns(MyDestTableCounter).ColumnName = sourceTable.Columns(MySourceTableCounter).ColumnName Then
                                columnExistsInDestination = True
                                Exit For
                            End If
                        End If
                    Next
                    If columnExistsInDestination = False Then
                        'Add missing column
                        Dim MyDestTableCounter As Integer 'for the new column index
                        MyDestTableCounter = destinationTable.Columns.Add(sourceTable.Columns(MySourceTableCounter).ColumnName, sourceTable.Columns(MySourceTableCounter).DataType).Ordinal

                        Dim sourceTableColumn As DataColumn = sourceTable.Columns(MySourceTableCounter)
                        Dim destTableColumn As DataColumn = destinationTable.Columns(MyDestTableCounter)

                        If destTableColumn.AllowDBNull = True AndAlso sourceTableColumn.AllowDBNull = False AndAlso destinationTable.Rows.Count > 0 Then
                            Try
                                destTableColumn.AllowDBNull = sourceTableColumn.AllowDBNull
                            Catch dataEx As Exception
                                Throw New Exception("Can't convert added column in destination table to NOT NULLABLE (rows already exist with assigned empty default values)", dataEx)
                            End Try
                        ElseIf destTableColumn.AllowDBNull <> sourceTableColumn.AllowDBNull Then
                            destTableColumn.AllowDBNull = sourceTableColumn.AllowDBNull
                        End If
                        destTableColumn.AutoIncrement = sourceTableColumn.AutoIncrement
                        destTableColumn.AutoIncrementSeed = sourceTableColumn.AutoIncrementSeed
                        destTableColumn.AutoIncrementStep = sourceTableColumn.AutoIncrementStep
                        destTableColumn.Caption = sourceTableColumn.Caption
                        destTableColumn.ColumnMapping = sourceTableColumn.ColumnMapping
                        destTableColumn.DefaultValue = sourceTableColumn.DefaultValue


                        destTableColumn.Expression = sourceTableColumn.Expression
                        destTableColumn.ExtendedProperties.Clear()
                        For Each key As Object In sourceTableColumn.ExtendedProperties
                            destTableColumn.ExtendedProperties.Add(key, sourceTableColumn.ExtendedProperties(key))
                        Next
                        destTableColumn.MaxLength = sourceTableColumn.MaxLength
                        destTableColumn.Namespace = sourceTableColumn.Namespace
                        destTableColumn.Prefix = sourceTableColumn.Prefix
                        destTableColumn.ReadOnly = sourceTableColumn.ReadOnly
                        destTableColumn.Unique = sourceTableColumn.Unique
                    End If
                Next
            End If
            'Copy related rows from source table to destination table row by row and column by column
            Dim MySrcTableRows As DataRow() = sourceTable.Select(sourceRowFilter, sourceSortExpression)


            If topRows = Nothing Then
                'All rows
                topRows = Integer.MaxValue
            End If

            'Copy rows
            If rowChanges <> RequestedRowChanges.KeepExistingRowsInDestinationTableAndAddRemoveUpdateChangedRows Then
                If MySrcTableRows IsNot Nothing Then

                    Dim srcTableColumnsList(sourceTable.Columns.Count - 1) As String
                    For i As Integer = 0 To sourceTable.Columns.Count - 1
                        srcTableColumnsList(i) = sourceTable.Columns(i).ColumnName
                    Next

                    'Copy row by row
                    For MySrcRowCounter As Integer = 1 To MySrcTableRows.Length
                        If MySrcRowCounter > topRows Then
                            Exit For
                        Else
                            'TODO: consider ignoreColumnNameCasing
                            'TODO: performance improvement proofed with 3000lines-table
                            Dim MyNewDestTableRow As DataRow = destinationTable.NewRow

                            'Copy column by column
                            For MyColCounter As Integer = 0 To sourceTable.Columns.Count - 1
                                Dim colName As String = srcTableColumnsList(MyColCounter)
                                'if we didn't extend the schema in the destination table we need to check here whether the Column actually exists.
                                If destinationTable.Columns.Contains(colName) Then
                                    MyNewDestTableRow(colName) = MySrcTableRows(MySrcRowCounter - 1)(MyColCounter)
                                End If
                            Next
                            destinationTable.Rows.Add(MyNewDestTableRow)
                        End If
                    Next
                End If
            Else 'Merging'
                Dim sourceView As DataView = sourceTable.DefaultView
                sourceView.RowFilter = sourceRowFilter
                sourceView.Sort = sourceSortExpression

                destinationTable.Merge(sourceView.ToTable(), False, MissingSchemaAction.Ignore)

                If topRows < destinationTable.Rows.Count Then
                    For MyCounter As Integer = destinationTable.Rows.Count - 1 To topRows Step -1
                        destinationTable.Rows.RemoveAt(MyCounter)
                    Next
                End If
            End If
        End Sub

        ''' <summary>
        '''     Remove the specified columns if they exist
        ''' </summary>
        ''' <param name="datatable">A datatable where the operations shall be made</param>
        ''' <param name="columnNames">The names of the columns which shall be removed</param>
        ''' <remarks>
        '''     The columns will only be removed if they exist. If a column name doesn't exist, it will be ignored.
        ''' </remarks>
        Public Shared Sub RemoveColumns(ByVal datatable As System.Data.DataTable, ByVal columnNames As String())
            CompuMaster.Data.DataTablesTools.RemoveColumns(datatable, columnNames)
        End Sub

        ''' <summary>
        '''     Creates a clone of a dataview but as a new data table
        ''' </summary>
        ''' <param name="data">The data view to create the data table from</param>
        ''' <returns></returns>
        Public Shared Function ConvertDataViewToDataTable(ByVal data As DataView) As System.Data.DataTable
            Return CompuMaster.Data.DataTablesTools.ConvertDataViewToDataTable(data)
        End Function

        ''' <summary>
        ''' Copy the values of a data column into an arraylist
        ''' </summary>
        ''' <param name="column">The column which contains the data</param>
        ''' <returns>An array containing data with type of the column's datatype OR with type of DBNull</returns>
        ''' <remarks></remarks>
        Public Shared Function ConvertColumnValuesIntoArrayList(ByVal column As DataColumn) As ArrayList
            Return ConvertDataTableToArrayList(column.Table, column.Ordinal)
        End Function

        ''' <summary>
        '''     Convert a data table to an arraylist
        ''' </summary>
        ''' <param name="column">The column which shall be used to fill the arraylist</param>
        ''' <returns>An array containing data with type of the column's datatype OR with type of DBNull</returns>
        Public Shared Function ConvertDataTableToArrayList(ByVal column As DataColumn) As ArrayList
            Return ConvertDataTableToArrayList(column.Table, column.Ordinal)
        End Function

        ''' <summary>
        '''     Convert a data table to an arraylist
        ''' </summary>
        ''' <param name="data">The first column of this data table will be used</param>
        ''' <returns>An array containing data with type of the column's datatype OR with type of DBNull</returns>
        Public Shared Function ConvertDataTableToArrayList(ByVal data As DataTable) As ArrayList
            Return ConvertDataTableToArrayList(data, 0)
        End Function

        ''' <summary>
        '''     Convert a data table to an arraylist
        ''' </summary>
        ''' <param name="data">The data table with the content</param>
        ''' <param name="selectedColumnIndex">The column which shall be used to fill the arraylist</param>
        ''' <returns>An array containing data with type of the column's datatype OR with type of DBNull</returns>
        Public Shared Function ConvertDataTableToArrayList(ByVal data As DataTable, ByVal selectedColumnIndex As Integer) As ArrayList
            Dim Result As New ArrayList
            For MyCounter As Integer = 0 To data.Rows.Count - 1
                Result.Add(data.Rows(MyCounter)(selectedColumnIndex))
            Next
            Return Result
        End Function

        ''' <summary>
        ''' Copy the values of a data column into an arraylist (except DBNull values)
        ''' </summary>
        ''' <param name="column">The column which contains the data</param>
        ''' <returns>An array containing data with type of the column's datatype OR with type of DBNull</returns>
        ''' <remarks></remarks>
        Public Shared Function ConvertColumnValuesIntoList(Of T)(ByVal column As DataColumn) As Generic.List(Of T)
            Return ConvertDataTableToList(Of T)(column.Table, column.Ordinal)
        End Function

        ''' <summary>
        '''     Convert a data table column to a generic list (except DBNull values)
        ''' </summary>
        ''' <param name="column">The column which shall be used</param>
        ''' <returns>An array containing data with type of the column's datatype OR with type of DBNull</returns>
        ''' <remarks>
        ''' </remarks>
        Public Shared Function ConvertDataTableToList(Of T)(ByVal column As DataColumn) As Generic.List(Of T)
            Return ConvertDataTableToList(Of T)(column.Table, column.Ordinal)
        End Function

        ''' <summary>
        '''     Convert a data table column to a generic list (except DBNull values)
        ''' </summary>
        ''' <param name="data">The first column of this data table will be used</param>
        ''' <returns>An array containing data with type of the column's datatype OR with type of DBNull</returns>
        ''' <remarks>
        ''' </remarks>
        Public Shared Function ConvertDataTableToList(Of T)(ByVal data As DataTable) As Generic.List(Of T)
            Return ConvertDataTableToList(Of T)(data, 0)
        End Function

        ''' <summary>
        '''     Convert a data table column to a generic list (except DBNull values)
        ''' </summary>
        ''' <param name="data">The data table with the content</param>
        ''' <param name="selectedColumnIndex">The column which shall be usedt</param>
        ''' <returns>An array containing data with type of the column's datatype OR with type of DBNull</returns>
        ''' <remarks>
        ''' </remarks>
        Public Shared Function ConvertDataTableToList(Of T)(ByVal data As DataTable, ByVal selectedColumnIndex As Integer) As Generic.List(Of T)
            Dim Result As New System.Collections.Generic.List(Of T)
            For MyCounter As Integer = 0 To data.Rows.Count - 1
                If Not IsDBNull(data.Rows(MyCounter)(selectedColumnIndex)) Then
                    Result.Add(CType(data.Rows(MyCounter)(selectedColumnIndex), T))
                End If
            Next
            Return Result
        End Function

        ''' <summary>
        ''' Copy the values of a data column into an arraylist (except rows with DBNull values in both columns)
        ''' </summary>
        ''' <param name="column1">The column which contains the data</param>
        ''' <param name="column2">The column which contains the data</param>
        ''' <returns>An array containing data with type of the column's datatype OR with type of DBNull</returns>
        ''' <remarks></remarks>
        Public Shared Function ConvertColumnValuesIntoList(Of T1, T2)(ByVal column1 As DataColumn, ByVal column2 As DataColumn) As Generic.List(Of Generic.KeyValuePair(Of T1, T2))
            If column1.Table IsNot column2.Table Then Throw New ArgumentException("Tables of both columns must be the same")
            Return ConvertDataTableToList(Of T1, T2)(column1.Table, column1.Ordinal, column2.Ordinal)
        End Function

        ''' <summary>
        '''     Convert a data table column to a generic list (except rows with DBNull values in both columns)
        ''' </summary>
        ''' <param name="column1">The column which contains the data</param>
        ''' <param name="column2">The column which contains the data</param>
        ''' <returns>An array containing data with type of the column's datatype OR with type of DBNull</returns>
        ''' <remarks>
        ''' </remarks>
        Public Shared Function ConvertDataTableToList(Of T1, T2)(ByVal column1 As DataColumn, ByVal column2 As DataColumn) As Generic.List(Of Generic.KeyValuePair(Of T1, T2))
            Return ConvertColumnValuesIntoList(Of T1, T2)(column1, column2)
        End Function

        ''' <summary>
        '''     Convert a data table column to a generic list (except rows with DBNull values in both columns)
        ''' </summary>
        ''' <param name="data">The first column of this data table will be used</param>
        ''' <returns>An array containing data with type of the column's datatype OR with type of DBNull</returns>
        ''' <remarks>
        ''' </remarks>
        Public Shared Function ConvertDataTableToList(Of T1, T2)(ByVal data As DataTable) As Generic.List(Of Generic.KeyValuePair(Of T1, T2))
            Return ConvertDataTableToList(Of T1, T2)(data, 0, 1)
        End Function

        ''' <summary>
        '''     Convert a data table column to a generic list (except rows with DBNull values in both columns)
        ''' </summary>
        ''' <param name="data">The data table with the content</param>
        ''' <param name="column1Index">The column which shall be used</param>
        ''' <param name="column2Index">The column which shall be used</param>
        ''' <returns>An array containing data with type of the column's datatype OR with type of DBNull</returns>
        ''' <remarks>
        ''' </remarks>
        Public Shared Function ConvertDataTableToList(Of T1, T2)(ByVal data As DataTable, ByVal column1Index As Integer, ByVal column2Index As Integer) As Generic.List(Of Generic.KeyValuePair(Of T1, T2))
            Dim Result As New System.Collections.Generic.List(Of Generic.KeyValuePair(Of T1, T2))
            For MyCounter As Integer = 0 To data.Rows.Count - 1
                If Not IsDBNull(data.Rows(MyCounter)(column1Index)) OrElse Not IsDBNull(data.Rows(MyCounter)(column2Index)) Then
                    Result.Add(New Generic.KeyValuePair(Of T1, T2)(Utils.NoDBNull(Of T1)(data.Rows(MyCounter)(column1Index)), Utils.NoDBNull(Of T2)(CType(data.Rows(MyCounter)(column2Index), T2))))
                End If
            Next
            Return Result
        End Function

        ''' <summary>
        '''     Convert a data table to a hash table
        ''' </summary>
        ''' <param name="keyColumn">This is the key column from the data table and MUST BE UNIQUE</param>
        ''' <param name="valueColumn">A column which contains the values</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' ATTENTION: the very first column is used as key column and must be unique therefore
        ''' </remarks>
        Public Shared Function ConvertDataTableToHashtable(ByVal keyColumn As DataColumn, ByVal valueColumn As DataColumn) As Hashtable
            Return CompuMaster.Data.DataTablesTools.ConvertDataTableToHashtable(keyColumn, valueColumn)
        End Function

        ''' <summary>
        '''     Convert a data table to a hash table
        ''' </summary>
        ''' <param name="data">The first two columns of this data table will be used</param>
        ''' <returns></returns>
        ''' <remarks>
        '''     ATTENTION: the very first column is used as key column and must be unique therefore
        ''' </remarks>
        Public Shared Function ConvertDataTableToHashtable(ByVal data As DataTable) As Hashtable
            Return CompuMaster.Data.DataTablesTools.ConvertDataTableToHashtable(data)
        End Function

        ''' <summary>
        '''     Convert a data table to a hash table
        ''' </summary>
        ''' <param name="data">The data table with the content</param>
        ''' <param name="keyColumnIndex">This is the key column from the data table and MUST BE UNIQUE</param>
        ''' <param name="valueColumnIndex">A column which contains the values</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' ATTENTION: the very first column is used as key column and must be unique therefore
        ''' </remarks>
        Public Shared Function ConvertDataTableToHashtable(ByVal data As DataTable, ByVal keyColumnIndex As Integer, ByVal valueColumnIndex As Integer) As Hashtable
            Return CompuMaster.Data.DataTablesTools.ConvertDataTableToHashtable(data, keyColumnIndex, valueColumnIndex)
        End Function

        ''' <summary>
        '''     Convert a data table to an array of dictionary entries
        ''' </summary>
        ''' <param name="data">The first two columns of this data table will be used</param>
        ''' <returns></returns>
        ''' <remarks>
        '''     The very first column is used as key column, the second one as the value column
        ''' </remarks>
        Public Shared Function ConvertDataTableToDictionaryEntryArray(ByVal data As DataTable) As DictionaryEntry()
            Return CompuMaster.Data.DataTablesTools.ConvertDataTableToDictionaryEntryArray(data)
        End Function

        ''' <summary>
        '''     Convert a data table to an array of dictionary entries
        ''' </summary>
        ''' <param name="keyColumn">This is the key column from the data table</param>
        ''' <param name="valueColumn">A column which contains the values</param>
        ''' <returns></returns>
        Public Shared Function ConvertDataTableToDictionaryEntryArray(ByVal keyColumn As DataColumn, ByVal valueColumn As DataColumn) As DictionaryEntry()
            Return CompuMaster.Data.DataTablesTools.ConvertDataTableToDictionaryEntryArray(keyColumn, valueColumn)
        End Function

        ''' <summary>
        '''     Convert a data table to an array of dictionary entries
        ''' </summary>
        ''' <param name="data">The data table with the content</param>
        ''' <param name="keyColumnIndex">This is the key column from the data table</param>
        ''' <param name="valueColumnIndex">A column which contains the values</param>
        ''' <returns></returns>
        Public Shared Function ConvertDataTableToDictionaryEntryArray(ByVal data As DataTable, ByVal keyColumnIndex As Integer, ByVal valueColumnIndex As Integer) As DictionaryEntry()
            Return CompuMaster.Data.DataTablesTools.ConvertDataTableToDictionaryEntryArray(data, keyColumnIndex, valueColumnIndex)
        End Function

        ''' <summary>
        '''     Convert an ICollection to a datatable
        ''' </summary>
        ''' <param name="collection">An ICollection with some content</param>
        ''' <returns>Datatable with column &quot;key&quot; and &quot;value&quot;</returns>
        Public Shared Function ConvertICollectionToDataTable(ByVal collection As ICollection) As DataTable
            Return CompuMaster.Data.DataTablesTools.ConvertICollectionToDataTable(collection)
        End Function

        ''' <summary>
        '''     Convert an IDictionary to a datatable
        ''' </summary>
        ''' <param name="dictionary">An IDictionary with some content</param>
        ''' <returns>Datatable with column &quot;key&quot; and &quot;value&quot;</returns>
        Public Shared Function ConvertIDictionaryToDataTable(ByVal dictionary As IDictionary) As DataTable
            Return CompuMaster.Data.DataTablesTools.ConvertIDictionaryToDataTable(dictionary)
        End Function

        ''' <summary>
        '''     Convert an IDictionary to a datatable
        ''' </summary>
        ''' <param name="dictionary">An IDictionary with some content</param>
        ''' <param name="keyIsUnique">If true, the key column in the data table will be marked as unique</param>
        ''' <returns>Datatable with column &quot;key&quot; and &quot;value&quot;</returns>
        Public Shared Function ConvertIDictionaryToDataTable(ByVal dictionary As IDictionary, ByVal keyIsUnique As Boolean) As DataTable
            Return CompuMaster.Data.DataTablesTools.ConvertIDictionaryToDataTable(dictionary, keyIsUnique)
        End Function

        ''' <summary>
        '''     Convert an array of DictionaryEntry to a datatable
        ''' </summary>
        ''' <param name="dictionaryEntries">An array of DictionaryEntry with some content</param>
        ''' <returns>Datatable with column &quot;key&quot; and &quot;value&quot;</returns>
        ''' <remarks>
        ''' </remarks>
        Public Shared Function ConvertDictionaryEntryArrayToDataTable(ByVal dictionaryEntries As DictionaryEntry()) As DataTable
            Return CompuMaster.Data.DataTablesTools.ConvertDictionaryEntryArrayToDataTable(dictionaryEntries, False)
        End Function

        ''' <summary>
        '''     Convert an array of DictionaryEntry to a datatable
        ''' </summary>
        ''' <param name="dictionaryEntries">An array of DictionaryEntry with some content</param>
        ''' <param name="keyIsUnique">If true, the key column in the data table will be marked as unique</param>
        ''' <returns>Datatable with column &quot;key&quot; and &quot;value&quot;</returns>
        ''' <remarks>
        ''' </remarks>
        Public Shared Function ConvertDictionaryEntryArrayToDataTable(ByVal dictionaryEntries As DictionaryEntry(), ByVal keyIsUnique As Boolean) As DataTable
            Return CompuMaster.Data.DataTablesTools.ConvertDictionaryEntryArrayToDataTable(dictionaryEntries, keyIsUnique)
        End Function

        ''' <summary>
        '''     Convert a NameValueCollection to a datatable
        ''' </summary>
        ''' <param name="nameValueCollection">An name-value-collection with some content</param>
        ''' <returns>Datatable with column &quot;key&quot; and &quot;value&quot;</returns>
        Public Shared Function ConvertNameValueCollectionToDataTable(ByVal nameValueCollection As Specialized.NameValueCollection) As DataTable
            Return CompuMaster.Data.DataTablesTools.ConvertNameValueCollectionToDataTable(nameValueCollection)
        End Function

        ''' <summary>
        '''     Convert a NameValueCollection to a datatable
        ''' </summary>
        ''' <param name="nameValueCollection">An name-value-collection with some content</param>
        ''' <param name="keyIsUnique">If true, the key column in the data table will be marked as unique</param>
        ''' <returns>Datatable with column &quot;key&quot; and &quot;value&quot;</returns>
        Public Shared Function ConvertNameValueCollectionToDataTable(ByVal nameValueCollection As Specialized.NameValueCollection, ByVal keyIsUnique As Boolean) As DataTable
            Return CompuMaster.Data.DataTablesTools.ConvertNameValueCollectionToDataTable(nameValueCollection, keyIsUnique)
        End Function

        ''' <summary>
        '''     Simplified creation of a DataTable by definition of a SQL statement and a connection string
        ''' </summary>
        ''' <param name="strSQL">The SQL statement to retrieve the data</param>
        ''' <param name="ConnectionString">The connection string to the data source</param>
        ''' <param name="NameOfNewDataTable">The name of the new DataTable</param>
        ''' <returns>A filled DataTable</returns>
        <Obsolete("Use DataQuery namespace instead"), ComponentModel.EditorBrowsable(ComponentModel.EditorBrowsableState.Never)>
        Public Shared Function GetDataTableViaODBC(ByVal strSQL As String, ByVal ConnectionString As String, ByVal NameOfNewDataTable As String) As DataTable
            Return CompuMaster.Data.DataTablesTools.GetDataTableViaODBC(strSQL, ConnectionString, NameOfNewDataTable)
        End Function

        ''' <summary>
        '''     Simplified creation of a DataTable by definition of a SQL statement and a connection string
        ''' </summary>
        ''' <param name="strSQL">The SQL statement to retrieve the data</param>
        ''' <param name="ConnectionString">The connection string to the data source</param>
        ''' <param name="NameOfNewDataTable">The name of the new DataTable</param>
        ''' <returns>A filled DataTable</returns>
        <Obsolete("Use DataQuery namespace instead"), ComponentModel.EditorBrowsable(ComponentModel.EditorBrowsableState.Never)>
        Public Shared Function GetDataTableViaSqlClient(ByVal strSQL As String, ByVal ConnectionString As String, ByVal NameOfNewDataTable As String) As DataTable
            Return CompuMaster.Data.DataTablesTools.GetDataTableViaSqlClient(strSQL, ConnectionString, NameOfNewDataTable)
        End Function

        ''' <summary>
        '''     Lookup a new unique column name for a data table
        ''' </summary>
        ''' <param name="dataTable">The data table which shall get a new data column</param>
        ''' <param name="suggestedColumnName">A column name suggestion</param>
        ''' <returns>The suggested column name as it is or modified column name to be unique</returns>
        ''' <remarks>
        ''' </remarks>
        Public Shared Function LookupUniqueColumnName(ByVal dataTable As DataTable, ByVal suggestedColumnName As String) As String
            Return CompuMaster.Data.DataTablesTools.LookupUniqueColumnName(dataTable, suggestedColumnName)
        End Function

        ''' <summary>
        '''     Lookup a new unique column name for a data table
        ''' </summary>
        ''' <param name="dataTable">The data table which shall get a new data column</param>
        ''' <param name="suggestedColumnName">A column name suggestion</param>
        ''' <returns>The suggested column name as it is or modified column name to be unique</returns>
        ''' <remarks>
        ''' </remarks>
        <Obsolete("Use the correct method name without typing error"), ComponentModel.EditorBrowsable(ComponentModel.EditorBrowsableState.Never)>
        Public Shared Function LookupUnqiueColumnName(ByVal dataTable As DataTable, ByVal suggestedColumnName As String) As String
            Return CompuMaster.Data.DataTablesTools.LookupUniqueColumnName(dataTable, suggestedColumnName)
        End Function

        ''' <summary>
        '''     Return a string with all columns and rows as an html table
        ''' </summary>
        ''' <param name="dataTable">The datatable to retrieve the content from</param>
        ''' <returns>If no rows have been processed, the return string is nothing</returns>
        Public Shared Function ConvertToHtmlTable(ByVal dataTable As DataTable) As String
            Return CompuMaster.Data.DataTablesTools.ConvertToHtmlTable(dataTable)
        End Function

        ''' <summary>
        '''     Return a string with all columns and rows as an html table
        ''' </summary>
        ''' <param name="rows">The rows to be processed</param>
        ''' <param name="label">An optional title of the rows</param>
        ''' <returns>If no rows have been processed, the return string is nothing</returns>
        Public Shared Function ConvertToHtmlTable(ByVal rows As DataRowCollection, ByVal label As String) As String
            Return CompuMaster.Data.DataTablesTools.ConvertToHtmlTable(rows, label)
        End Function

        ''' <summary>
        '''     Return a string with all columns and rows as an html table
        ''' </summary>
        ''' <param name="rows">The rows to be processed</param>
        ''' <param name="label">An optional title of the rows</param>
        ''' <returns>If no rows have been processed, the return string is nothing</returns>
        Public Shared Function ConvertToHtmlTable(ByVal rows() As DataRow, ByVal label As String) As String
            Return CompuMaster.Data.DataTablesTools.ConvertToHtmlTable(rows, label)
        End Function

        ''' <summary>
        '''     Return a string with all columns and rows as an html table
        ''' </summary>
        ''' <param name="dataTable">The datatable to retrieve the content from</param>
        ''' <param name="titleTagOpener">The opening tag in front of the table's title</param>
        ''' <param name="titleTagEnd">The closing tag after the table title</param>
        ''' <param name="additionalTableAttributes">Additional attributes for the rendered table</param>
        ''' <returns>If no rows have been processed, the return string is nothing</returns>
        Public Shared Function ConvertToHtmlTable(ByVal dataTable As DataTable, ByVal titleTagOpener As String, ByVal titleTagEnd As String,
                                                  ByVal additionalTableAttributes As String) As String
            Return CompuMaster.Data.DataTablesTools.ConvertToHtmlTable(dataTable, titleTagOpener, titleTagEnd, additionalTableAttributes)
        End Function

        <Obsolete("Subject of removal in a future version", True), System.ComponentModel.EditorBrowsable(ComponentModel.EditorBrowsableState.Never)>
        Public Shared Function ConvertToHtmlTable(ByVal dataTable As DataTable, ByVal titleTag As String, ByVal additionalTableAttributes As String) As String
            Return CompuMaster.Data.DataTablesTools.ConvertToHtmlTable(dataTable, "<" & titleTag & ">", "</" & titleTag & ">", additionalTableAttributes)
        End Function

        ''' <summary>
        '''     Return a string with all columns and rows as an html table
        ''' </summary>
        ''' <param name="dataTable">The datatable to retrieve the content from</param>
        ''' <param name="titleTagOpener">The opening tag in front of the table's title</param>
        ''' <param name="titleTagEnd">The closing tag after the table title</param>
        ''' <param name="additionalTableAttributes">Additional attributes for the rendered table</param>
        ''' <param name="htmlEncodeCellContentAndLineBreaks">Encode all output to valid HTML</param>
        ''' <returns>If no rows have been processed, the return string is nothing</returns>
        Public Shared Function ConvertToHtmlTable(ByVal dataTable As DataTable, ByVal titleTagOpener As String, ByVal titleTagEnd As String,
                                                  ByVal additionalTableAttributes As String, ByVal htmlEncodeCellContentAndLineBreaks As Boolean) As String
            Return CompuMaster.Data.DataTablesTools.ConvertToHtmlTable(dataTable.Rows, dataTable.TableName, titleTagOpener, titleTagEnd, additionalTableAttributes, htmlEncodeCellContentAndLineBreaks)
        End Function

        ''' <summary>
        '''     Return a string with all columns and rows as an html table
        ''' </summary>
        ''' <param name="rows">The rows to be processed</param>
        ''' <param name="label">An optional title of the rows</param>
        ''' <param name="titleTagOpener">The opening tag in front of the table's title</param>
        ''' <param name="titleTagEnd">The closing tag after the table title</param>
        ''' <param name="additionalTableAttributes">Additional attributes for the rendered table</param>
        ''' <returns>If no rows have been processed, the return string is nothing</returns>
        Public Shared Function ConvertToHtmlTable(ByVal rows As DataRowCollection, ByVal label As String, ByVal titleTagOpener As String, ByVal titleTagEnd As String,
                                                  ByVal additionalTableAttributes As String) As String
            Return CompuMaster.Data.DataTablesTools.ConvertToHtmlTable(rows, label, titleTagOpener, titleTagEnd, additionalTableAttributes)
        End Function

        <Obsolete("Subject of removal in a future version", True), System.ComponentModel.EditorBrowsable(ComponentModel.EditorBrowsableState.Never)>
        Public Shared Function ConvertToHtmlTable(ByVal rows As DataRowCollection, ByVal label As String, ByVal titleTag As String, ByVal additionalTableAttributes As String) As String
            Return CompuMaster.Data.DataTablesTools.ConvertToHtmlTable(rows, label, "<" & titleTag & ">", "</" & titleTag & ">", additionalTableAttributes)
        End Function

        ''' <summary>
        '''     Return a string with all columns and rows as an html table
        ''' </summary>
        ''' <param name="rows">The rows to be processed</param>
        ''' <param name="label">An optional title of the rows</param>
        ''' <param name="titleTagOpener">The opening tag in front of the table's title</param>
        ''' <param name="titleTagEnd">The closing tag after the table title</param>
        ''' <param name="additionalTableAttributes">Additional attributes for the rendered table</param>
        ''' <returns>If no rows have been processed, the return string is nothing</returns>
        Public Shared Function ConvertToHtmlTable(ByVal rows() As DataRow, ByVal label As String, ByVal titleTagOpener As String, ByVal titleTagEnd As String,
                                                  ByVal additionalTableAttributes As String) As String
            Return CompuMaster.Data.DataTablesTools.ConvertToHtmlTable(rows, label, titleTagOpener, titleTagEnd, additionalTableAttributes)
        End Function

        <Obsolete("Subject of removal in a future version", True), System.ComponentModel.EditorBrowsable(ComponentModel.EditorBrowsableState.Never)>
        Public Shared Function ConvertToHtmlTable(ByVal rows() As DataRow, ByVal label As String, ByVal titleTag As String, ByVal additionalTableAttributes As String) As String
            Return CompuMaster.Data.DataTablesTools.ConvertToHtmlTable(rows, label, "<" & titleTag & ">", "</" & titleTag & ">", additionalTableAttributes)
        End Function

        ''' <summary>
        '''     Return a string with all columns and rows, helpfull for debugging purposes
        ''' </summary>
        ''' <param name="dataTable">The datatable to retrieve the content from</param>
        ''' <returns>All rows are tab separated. If no rows have been processed, the user will get notified about this fact</returns>
        Public Shared Function ConvertToPlainTextTable(ByVal dataTable As DataTable) As String
            Return CompuMaster.Data.DataTablesTools.ConvertToPlainTextTable(dataTable)
        End Function

        ''' <summary>
        '''     Return a string with all columns and rows, helpfull for debugging purposes
        ''' </summary>
        ''' <param name="dataTable">The datatable to retrieve the content from</param>
        ''' <param name="fixedColumnWidths">The column sizes in chars</param>
        ''' <returns>All rows are tab separated. If no rows have been processed, the user will get notified about this fact</returns>
        <Obsolete("Use ConvertToPlainTextTableFixedColumnWidths instead", False), ComponentModel.EditorBrowsable(ComponentModel.EditorBrowsableState.Never)>
        Public Shared Function ConvertToPlainTextTable(ByVal dataTable As DataTable, ByVal fixedColumnWidths As Integer()) As String
            Return ConvertToPlainTextTable(dataTable.Rows, dataTable.TableName, fixedColumnWidths)
        End Function

        ''' <summary>
        '''     Return a string with all columns and rows, helpfull for debugging purposes
        ''' </summary>
        ''' <param name="dataRows">The datatable to retrieve the content from</param>
        ''' <returns>All rows are separated by fixed width. If no rows have been processed, the user will get notified about this fact</returns>
        ''' <remarks></remarks>
        Public Shared Function ConvertToPlainTextTableFixedColumnWidths(ByVal dataRows As DataRow()) As String
            Dim TableName As String = ""
            If dataRows.Length > 0 Then
                TableName = dataRows(0).Table.TableName
            End If
            Return ConvertToPlainTextTableWithFixedColumnWidthsInternal(dataRows, TableName, SuggestColumnWidthsForFixedPlainTables(dataRows, 100.0), Nothing)
        End Function

        ''' <summary>
        '''     Return a string with all columns and rows, helpfull for debugging purposes
        ''' </summary>
        ''' <param name="dataRows">The datatable to retrieve the content from</param>
        ''' <param name="tableTitle">The headline for the table</param>
        ''' <returns>All rows are separated by fixed width. If no rows have been processed, the user will get notified about this fact</returns>
        ''' <remarks></remarks>
        Public Shared Function ConvertToPlainTextTableFixedColumnWidths(ByVal dataRows As DataRow(), tableTitle As String) As String
            Return ConvertToPlainTextTableWithFixedColumnWidthsInternal(dataRows, tableTitle, SuggestColumnWidthsForFixedPlainTables(dataRows, 100.0), Nothing)
        End Function

        ''' <summary>
        '''     Return a string with all columns and rows, helpfull for debugging purposes
        ''' </summary>
        ''' <param name="dataRows">The datatable to retrieve the content from</param>
        ''' <returns>All rows are separated by fixed width. If no rows have been processed, the user will get notified about this fact</returns>
        ''' <remarks></remarks>
        Public Shared Function ConvertToPlainTextTableFixedColumnWidths(ByVal dataRows As DataRow(), columnFormatting As DataColumnToString) As String
            Dim TableName As String = ""
            If dataRows.Length > 0 Then
                TableName = dataRows(0).Table.TableName
            End If
            Return ConvertToPlainTextTableWithFixedColumnWidthsInternal(dataRows, TableName, SuggestColumnWidthsForFixedPlainTables(dataRows, Nothing, 100.0, columnFormatting), columnFormatting)
        End Function

        ''' <summary>
        '''     Return a string with all columns and rows, helpfull for debugging purposes
        ''' </summary>
        ''' <param name="dataRow">The data row to retrieve the content from</param>
        ''' <returns>All rows are separated by fixed width. If no rows have been processed, the user will get notified about this fact</returns>
        ''' <remarks></remarks>
        Public Shared Function ConvertToPlainTextTableFixedColumnWidths(ByVal dataRow As DataRow) As String
            Return ConvertToPlainTextTableWithFixedColumnWidthsInternal(New System.Data.DataRow() {dataRow}, dataRow.Table.TableName, SuggestColumnWidthsForFixedPlainTables(New System.Data.DataRow() {dataRow}, 100.0), Nothing)
        End Function

        ''' <summary>
        '''     Return a string with all columns and rows, helpfull for debugging purposes
        ''' </summary>
        ''' <param name="dataRow">The data row to retrieve the content from</param>
        ''' <param name="tableTitle">The headline for the table</param>
        ''' <returns>All rows are separated by fixed width. If no rows have been processed, the user will get notified about this fact</returns>
        ''' <remarks></remarks>
        Public Shared Function ConvertToPlainTextTableFixedColumnWidths(ByVal dataRow As DataRow, tableTitle As String) As String
            Return ConvertToPlainTextTableWithFixedColumnWidthsInternal(New System.Data.DataRow() {dataRow}, tableTitle, SuggestColumnWidthsForFixedPlainTables(New System.Data.DataRow() {dataRow}, 100.0), Nothing)
        End Function

        ''' <summary>
        '''     Return a string with all columns and rows, helpfull for debugging purposes
        ''' </summary>
        ''' <param name="dataRow">The data row to retrieve the content from</param>
        ''' <returns>All rows are separated by fixed width. If no rows have been processed, the user will get notified about this fact</returns>
        ''' <remarks></remarks>
        Public Shared Function ConvertToPlainTextTableFixedColumnWidths(ByVal dataRow As DataRow, columnFormatting As DataColumnToString) As String
            Return ConvertToPlainTextTableWithFixedColumnWidthsInternal(New System.Data.DataRow() {dataRow}, dataRow.Table.TableName, SuggestColumnWidthsForFixedPlainTables(New System.Data.DataRow() {dataRow}, Nothing, 100.0, columnFormatting), columnFormatting)
        End Function

        ''' <summary>
        '''     Return a string with all columns and rows, helpfull for debugging purposes
        ''' </summary>
        ''' <param name="dataTable">The datatable to retrieve the content from</param>
        ''' <returns>All rows are separated by fixed width. If no rows have been processed, the user will get notified about this fact</returns>
        ''' <remarks></remarks>
        Public Shared Function ConvertToPlainTextTableFixedColumnWidths(ByVal dataTable As DataTable) As String
            Return ConvertToPlainTextTableWithFixedColumnWidthsInternal(dataTable.Rows, dataTable.TableName, SuggestColumnWidthsForFixedPlainTables(dataTable.Rows, 100.0), Nothing)
        End Function

        ''' <summary>
        '''     Return a string with all columns and rows, helpfull for debugging purposes
        ''' </summary>
        ''' <param name="dataTable">The datatable to retrieve the content from</param>
        ''' <param name="tableTitle">The headline for the table</param>
        ''' <returns>All rows are separated by fixed width. If no rows have been processed, the user will get notified about this fact</returns>
        ''' <remarks></remarks>
        Public Shared Function ConvertToPlainTextTableFixedColumnWidths(ByVal dataTable As DataTable, ByVal tableTitle As String) As String
            Return ConvertToPlainTextTableWithFixedColumnWidthsInternal(dataTable.Rows, tableTitle, SuggestColumnWidthsForFixedPlainTables(dataTable.Rows, 100.0), Nothing)
        End Function

        ''' <summary>
        '''     Return a string with all columns and rows, helpfull for debugging purposes
        ''' </summary>
        ''' <param name="dataTable">The datatable to retrieve the content from</param>
        ''' <returns>All rows are separated by fixed width. If no rows have been processed, the user will get notified about this fact</returns>
        ''' <remarks></remarks>
        Public Shared Function ConvertToPlainTextTableFixedColumnWidths(ByVal dataTable As DataTable, columnFormatting As DataColumnToString) As String
            Return ConvertToPlainTextTableWithFixedColumnWidthsInternal(dataTable.Rows, dataTable.TableName, SuggestColumnWidthsForFixedPlainTables(dataTable.Rows, Nothing, 100.0, columnFormatting), columnFormatting)
        End Function

        ''' <summary>
        '''     Return a string with all columns and rows, helpfull for debugging purposes
        ''' </summary>
        ''' <param name="dataTable">The datatable to retrieve the content from</param>
        ''' <returns>All rows are separated by fixed width. If no rows have been processed, the user will get notified about this fact</returns>
        ''' <remarks></remarks>
        Public Shared Function ConvertToPlainTextTableFixedColumnWidths(ByVal dataTable As DataTable, ByVal fixedColumnWidths As Integer()) As String
            Return ConvertToPlainTextTableWithFixedColumnWidthsInternal(dataTable.Rows, dataTable.TableName, fixedColumnWidths, Nothing)
        End Function

        ''' <summary>
        '''     Return a string with all columns and rows, helpfull for debugging purposes
        ''' </summary>
        ''' <param name="dataTable">The datatable to retrieve the content from</param>
        ''' <returns>All rows are separated by fixed width. If no rows have been processed, the user will get notified about this fact</returns>
        ''' <remarks></remarks>
        Public Shared Function ConvertToPlainTextTableFixedColumnWidths(ByVal dataTable As DataTable, ByVal standardColumnWidth As Integer) As String
            Dim columnWidths As New System.Collections.Generic.List(Of Integer)
            For MyCounter As Integer = 0 To dataTable.Columns.Count - 1
                columnWidths.Add(standardColumnWidth)
            Next
            Return ConvertToPlainTextTableWithFixedColumnWidthsInternal(dataTable.Rows, dataTable.TableName, columnWidths.ToArray, Nothing)
        End Function

        ''' <summary>
        '''     Return a string with all columns and rows, helpfull for debugging purposes
        ''' </summary>
        ''' <param name="dataTable">The datatable to retrieve the content from</param>
        ''' <returns>All rows are separated by fixed width. If no rows have been processed, the user will get notified about this fact</returns>
        ''' <remarks></remarks>
        Public Shared Function ConvertToPlainTextTableFixedColumnWidths(ByVal dataTable As DataTable, ByVal minimumColumnWidth As Integer,
                                                                        maximumColumnWidth As Integer) As String
            Return ConvertToPlainTextTableFixedColumnWidths(dataTable, minimumColumnWidth, maximumColumnWidth, Nothing)
        End Function

        ''' <summary>
        '''     Return a string with all columns and rows, helpfull for debugging purposes
        ''' </summary>
        ''' <param name="dataTable">The datatable to retrieve the content from</param>
        ''' <returns>All rows are separated by fixed width. If no rows have been processed, the user will get notified about this fact</returns>
        ''' <remarks></remarks>
        Public Shared Function ConvertToPlainTextTableFixedColumnWidths(ByVal dataTable As DataTable, ByVal minimumColumnWidth As Integer,
                                                                        maximumColumnWidth As Integer,
                                                                        columnFormatting As DataColumnToString) As String
            Dim columnWidths As Integer() = SuggestColumnWidthsForFixedPlainTables(dataTable.Rows, dataTable, 100.0, columnFormatting)
            If columnWidths Is Nothing Then
                Dim newWidths(dataTable.Columns.Count - 1) As Integer
                For MyCounter As Integer = 0 To dataTable.Columns.Count - 1
                    newWidths(MyCounter) = minimumColumnWidth
                Next
                columnWidths = newWidths
            Else
                For MyCounter As Integer = 0 To columnWidths.Length - 1
                    If columnWidths(MyCounter) < minimumColumnWidth Then columnWidths(MyCounter) = minimumColumnWidth
                    If columnWidths(MyCounter) > maximumColumnWidth Then columnWidths(MyCounter) = maximumColumnWidth
                Next
            End If
            Return ConvertToPlainTextTableWithFixedColumnWidthsInternal(dataTable.Rows, dataTable.TableName, columnWidths, columnFormatting)
        End Function

        ''' <summary>
        '''     Return a string with all columns and rows, helpfull for debugging purposes
        ''' </summary>
        ''' <param name="dataTable">The datatable to retrieve the content from</param>
        ''' <returns>All rows are separated by fixed width. If no rows have been processed, the user will get notified about this fact</returns>
        ''' <remarks></remarks>
        Public Shared Function ConvertToPlainTextTableFixedColumnWidths(ByVal dataTable As DataTable, verticalSeparatorHeader As String,
                                                                        verticalSeparatorCells As String, crossSeparator As String,
                                                                        horizontalSeparatorHeadline As Char, horizontalSeparatorCells As Char) As String
            Return ConvertToPlainTextTableWithFixedColumnWidthsInternal(dataTable.Rows, dataTable.TableName, SuggestColumnWidthsForFixedPlainTables(dataTable.Rows), verticalSeparatorHeader, verticalSeparatorCells,
                                                                 crossSeparator, horizontalSeparatorHeadline, horizontalSeparatorCells, Nothing)
        End Function

        ''' <summary>
        '''     Return a string with all columns and rows, helpfull for debugging purposes
        ''' </summary>
        ''' <param name="dataTable">The datatable to retrieve the content from</param>
        ''' <returns>All rows are separated by fixed width. If no rows have been processed, the user will get notified about this fact</returns>
        ''' <remarks></remarks>
        Public Shared Function ConvertToPlainTextTableFixedColumnWidths(ByVal dataTable As DataTable, ByVal fixedColumnWidths As Integer(),
                                                                        verticalSeparatorHeader As String, verticalSeparatorCells As String,
                                                                        crossSeparator As String, horizontalSeparatorHeadline As Char,
                                                                        horizontalSeparatorCells As Char) As String
            Return ConvertToPlainTextTableWithFixedColumnWidthsInternal(dataTable.Rows, dataTable.TableName, fixedColumnWidths, verticalSeparatorHeader, verticalSeparatorCells, crossSeparator, horizontalSeparatorHeadline,
                                                                 horizontalSeparatorCells, Nothing)
        End Function

        ''' <summary>
        '''     Return a string with all columns and rows, helpfull for debugging purposes
        ''' </summary>
        ''' <param name="dataTable">The datatable to retrieve the content from</param>
        ''' <returns>All rows are separated by fixed width. If no rows have been processed, the user will get notified about this fact</returns>
        ''' <remarks></remarks>
        Public Shared Function ConvertToPlainTextTableFixedColumnWidths(ByVal dataTable As DataTable, ByVal standardColumnWidth As Integer,
                                                                        verticalSeparatorHeader As String, verticalSeparatorCells As String,
                                                                        crossSeparator As String, horizontalSeparatorHeadline As Char,
                                                                        horizontalSeparatorCells As Char) As String
            Dim columnWidths(dataTable.Columns.Count - 1) As Integer
            For MyCounter As Integer = 0 To dataTable.Columns.Count - 1
                columnWidths(MyCounter) = standardColumnWidth
            Next
            Return ConvertToPlainTextTableWithFixedColumnWidthsInternal(dataTable.Rows, dataTable.TableName, columnWidths, verticalSeparatorHeader, verticalSeparatorCells,
                                                                 crossSeparator, horizontalSeparatorHeadline, horizontalSeparatorCells, Nothing)
        End Function

        ''' <summary>
        '''     Return a string with all columns and rows, helpfull for debugging purposes
        ''' </summary>
        ''' <param name="dataTable">The datatable to retrieve the content from</param>
        ''' <returns>All rows are separated by fixed width. If no rows have been processed, the user will get notified about this fact</returns>
        ''' <remarks></remarks>
        Public Shared Function ConvertToPlainTextTableFixedColumnWidths(ByVal dataTable As DataTable, ByVal minimumColumnWidth As Integer,
                                                                        maximumColumnWidth As Integer, verticalSeparatorHeader As String,
                                                                        verticalSeparatorCells As String, crossSeparator As String,
                                                                        horizontalSeparatorHeadline As Char, horizontalSeparatorCells As Char) As String
            Dim columnWidths As Integer() = SuggestColumnWidthsForFixedPlainTables(dataTable.Rows, dataTable, 100.0, Nothing)
            If columnWidths Is Nothing Then
                Dim newWidths(dataTable.Columns.Count - 1) As Integer
                For MyCounter As Integer = 0 To dataTable.Columns.Count - 1
                    newWidths(MyCounter) = minimumColumnWidth
                Next
                columnWidths = newWidths
            Else
                For MyCounter As Integer = 0 To columnWidths.Length - 1
                    If columnWidths(MyCounter) < minimumColumnWidth Then columnWidths(MyCounter) = minimumColumnWidth
                    If columnWidths(MyCounter) > maximumColumnWidth Then columnWidths(MyCounter) = maximumColumnWidth
                Next
            End If
            Return ConvertToPlainTextTableWithFixedColumnWidthsInternal(dataTable.Rows, dataTable.TableName, columnWidths, verticalSeparatorHeader,
                                                                 verticalSeparatorCells, crossSeparator, horizontalSeparatorHeadline,
                                                                 horizontalSeparatorCells, Nothing)
        End Function

        ''' <summary>
        '''     Return a string with all columns and rows, helpfull for debugging purposes
        ''' </summary>
        ''' <param name="dataTable">The datatable to retrieve the content from</param>
        ''' <returns>All rows are separated by fixed width. If no rows have been processed, the user will get notified about this fact</returns>
        ''' <remarks></remarks>
        Public Shared Function ConvertToPlainTextTableFixedColumnWidths(ByVal dataTable As DataTable, ByVal minimumColumnWidth As Integer,
                                                                        maximumColumnWidth As Integer, verticalSeparatorHeader As String,
                                                                        verticalSeparatorCells As String, crossSeparator As String,
                                                                        horizontalSeparatorHeadline As Char, horizontalSeparatorCells As Char,
                                                                        columnFormatting As DataColumnToString) As String
            Dim columnWidths As Integer() = SuggestColumnWidthsForFixedPlainTables(dataTable.Rows, dataTable, 100.0, columnFormatting)
            If columnWidths Is Nothing Then
                Dim newWidths(dataTable.Columns.Count - 1) As Integer
                For MyCounter As Integer = 0 To dataTable.Columns.Count - 1
                    newWidths(MyCounter) = minimumColumnWidth
                Next
                columnWidths = newWidths
            Else
                For MyCounter As Integer = 0 To columnWidths.Length - 1
                    If columnWidths(MyCounter) < minimumColumnWidth Then columnWidths(MyCounter) = minimumColumnWidth
                    If columnWidths(MyCounter) > maximumColumnWidth Then columnWidths(MyCounter) = maximumColumnWidth
                Next
            End If
            Return ConvertToPlainTextTableWithFixedColumnWidthsInternal(dataTable.Rows, dataTable.TableName, columnWidths, verticalSeparatorHeader,
                                                                 verticalSeparatorCells, crossSeparator, horizontalSeparatorHeadline,
                                                                 horizontalSeparatorCells, columnFormatting)
        End Function

        ''' <summary>
        ''' Create a well-formed table for Wiki
        ''' </summary>
        ''' <param name="rows">A set of DataRows</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ConvertToWikiTable(ByVal rows As DataRowCollection) As String
            If rows.Count = 0 Then
                Return Nothing
            Else
                Return ConvertToWikiTable(rows(0).Table)
            End If
        End Function

        ''' <summary>
        ''' Create a well-formed table for Wiki
        ''' </summary>
        ''' <param name="rows">A set of DataRows</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ConvertToWikiTable(ByVal rows As DataRowCollection, columnFormatting As DataColumnToString) As String
            If rows.Count = 0 Then
                Return Nothing
            Else
                Return ConvertToWikiTable(rows(0).Table, columnFormatting)
            End If
        End Function

        ''' <summary>
        ''' Create a well-formed table for Wiki
        ''' </summary>
        ''' <param name="table">A table</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ConvertToWikiTable(ByVal table As DataTable) As String
            Return ConvertToWikiTable(table, Nothing)
        End Function

        ''' <summary>
        ''' Create a well-formed table for Wiki
        ''' </summary>
        ''' <param name="table">A table</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ConvertToWikiTable(ByVal table As DataTable, columnFormatting As DataColumnToString) As String
            'For DokuWiki, use
            Const verticalSeparatorHeader As String = " ^ "
            Const verticalSeparatorCells As String = " | "
            Dim fixedColumnWidths As Integer() = SuggestColumnWidthsForFixedPlainTables(table.Rows, table, 100.0, columnFormatting)
            Dim Result As New System.Text.StringBuilder
            Dim rows As DataRowCollection = table.Rows
            'Add table name
            If rows.Count <= 0 Then
                Result.Append("no rows found" & System.Environment.NewLine)
                Return Result.ToString
            End If
            'Add column headers
            For ColCounter As Integer = 0 To System.Math.Min(rows(0).Table.Columns.Count, fixedColumnWidths.Length) - 1
                Dim column As DataColumn = rows(0).Table.Columns(ColCounter)
                Dim textAlignmentRight As Boolean
                Select Case column.DataType.Name
                    Case GetType(Int16).Name, GetType(Int32).Name, GetType(Int64).Name, GetType(Single).Name, GetType(Decimal).Name, GetType(Double).Name, GetType(UInt16).Name, GetType(UInt32).Name, GetType(UInt64).Name
                        textAlignmentRight = True
                    Case Else
                        textAlignmentRight = False
                End Select
                If ColCounter = 0 Then
                    Result.Append(verticalSeparatorHeader.TrimStart)
                Else
                    Result.Append(verticalSeparatorHeader)
                End If
                If textAlignmentRight = True Then Result.Append(" ")
                If column.Caption <> Nothing Then
                    Result.Append(TrimStringToFixedWidth(String.Format("{0}", column.Caption), fixedColumnWidths(ColCounter)))
                Else
                    Result.Append(TrimStringToFixedWidth(String.Format("{0}", column.ColumnName), fixedColumnWidths(ColCounter)))
                End If
                If ColCounter = table.Columns.Count - 1 Then Result.Append(verticalSeparatorCells.TrimEnd)
            Next
            Result.Append(System.Environment.NewLine)
            'Add table rows
            For Each row As DataRow In rows
                For ColCounter As Integer = 0 To System.Math.Min(row.Table.Columns.Count, fixedColumnWidths.Length) - 1
                    Dim column As DataColumn = row.Table.Columns(ColCounter)
                    Dim textAlignmentRight As Boolean
                    Select Case column.DataType.Name
                        Case GetType(Int16).Name, GetType(Int32).Name, GetType(Int64).Name, GetType(Single).Name, GetType(Decimal).Name, GetType(Double).Name, GetType(UInt16).Name, GetType(UInt32).Name, GetType(UInt64).Name
                            textAlignmentRight = True
                        Case Else
                            textAlignmentRight = False
                    End Select
                    If ColCounter = 0 Then
                        Result.Append(verticalSeparatorCells.TrimStart)
                    Else
                        Result.Append(verticalSeparatorCells)
                    End If
                    If textAlignmentRight = True Then Result.Append(" ")
                    Dim RenderValue As Object
                    If columnFormatting Is Nothing Then
                        RenderValue = row(column)
                    Else
                        RenderValue = columnFormatting(column, row(column))
                    End If
                    Result.Append(TrimStringToFixedWidth(String.Format("{0}", RenderValue), fixedColumnWidths(ColCounter)))
                    If ColCounter = table.Columns.Count - 1 Then Result.Append(verticalSeparatorCells.TrimEnd)
                Next
                Result.Append(System.Environment.NewLine)
            Next
            Return Result.ToString
        End Function

        ''' <summary>
        ''' Suggests column widths for a table using as minimum 2 chars, but minimum header string length, but also either full cell length for number/date/time columns or for all other types 80 % of all values should be visible completely
        ''' </summary>
        ''' <param name="table">A table</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function SuggestColumnWidthsForFixedPlainTables(table As System.Data.DataTable) As Integer()
            Return SuggestColumnWidthsForFixedPlainTables(table.Rows, table, 80, Nothing)
        End Function

        ''' <summary>
        ''' Suggests column widths for a table using as minimum 2 chars, but minimum header string length, but also either full cell length for number/date/time columns or for all other types a given % value of all values should be visible completely
        ''' </summary>
        ''' <param name="table">A table</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function SuggestColumnWidthsForFixedPlainTables(table As System.Data.DataTable, optimalWidthWhenPercentageNumberOfRowsFitIntoCell As Double) As Integer()
            Return SuggestColumnWidthsForFixedPlainTables(table.Rows, table, optimalWidthWhenPercentageNumberOfRowsFitIntoCell, Nothing)
        End Function

        ''' <summary>
        ''' Suggests column widths for a table using as minimum 2 chars, but minimum header string length, but also either full cell length for number/date/time columns or for all other types 80 % of all values should be visible completely
        ''' </summary>
        ''' <param name="rows">A set of DataRows</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function SuggestColumnWidthsForFixedPlainTables(rows As System.Data.DataRowCollection) As Integer()
            If rows.Count = 0 Then
                Return Nothing
            Else
                Return SuggestColumnWidthsForFixedPlainTables(rows, rows(0).Table, 80, Nothing)
            End If
        End Function

        ''' <summary>
        ''' Suggests column widths for a table using as minimum 2 chars, but minimum header string length, but also either full cell length for number/date/time columns or for all other types a given % value of all values should be visible completely
        ''' </summary>
        ''' <param name="rows">A set of DataRows</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function SuggestColumnWidthsForFixedPlainTables(rows As System.Data.DataRowCollection, optimalWidthWhenPercentageNumberOfRowsFitIntoCell As Double) As Integer()
            If rows.Count = 0 Then
                Return Nothing
            Else
                Return SuggestColumnWidthsForFixedPlainTables(rows, rows(0).Table, optimalWidthWhenPercentageNumberOfRowsFitIntoCell, Nothing)
            End If
        End Function

        ''' <summary>
        ''' Suggests column widths for a table using as minimum 2 chars, but minimum header string length, but also either full cell length for number/date/time columns or for all other types 80 % of all values should be visible completely
        ''' </summary>
        ''' <param name="rows">A set of DataRows</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function SuggestColumnWidthsForFixedPlainTables(rows As System.Data.DataRow()) As Integer()
            If rows.Length = 0 Then
                Return Nothing
            Else
                Return SuggestColumnWidthsForFixedPlainTables(rows, rows(0).Table, 80, Nothing)
            End If
        End Function

        ''' <summary>
        ''' Suggests column widths for a table using as minimum 2 chars, but minimum header string length, but also either full cell length for number/date/time columns or for all other types a given % value of all values should be visible completely
        ''' </summary>
        ''' <param name="rows">A set of DataRows</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function SuggestColumnWidthsForFixedPlainTables(rows As System.Data.DataRow(), optimalWidthWhenPercentageNumberOfRowsFitIntoCell As Double) As Integer()
            If rows.Length = 0 Then
                Return Nothing
            Else
                Return SuggestColumnWidthsForFixedPlainTables(rows, rows(0).Table, optimalWidthWhenPercentageNumberOfRowsFitIntoCell, Nothing)
            End If
        End Function

        ''' <summary>
        ''' Suggests column widths for a table using as minimum 2 chars, but minimum header string length, but also either full cell length for number/date/time columns or for all other types 80 % of all values should be visible completely
        ''' </summary>
        ''' <param name="table">A table</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function SuggestColumnWidthsForFixedPlainTables(table As System.Data.DataTable, columnFormatting As DataColumnToString) As Integer()
            Return SuggestColumnWidthsForFixedPlainTables(table.Rows, table, 80, columnFormatting)
        End Function

        ''' <summary>
        ''' Suggests column widths for a table using as minimum 2 chars, but minimum header string length, but also either full cell length for number/date/time columns or for all other types 80 % of all values should be visible completely
        ''' </summary>
        ''' <param name="rows">A set of DataRows</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function SuggestColumnWidthsForFixedPlainTables(rows As System.Data.DataRowCollection, columnFormatting As DataColumnToString) As Integer()
            If rows.Count = 0 Then
                Return Nothing
            Else
                Return SuggestColumnWidthsForFixedPlainTables(rows, rows(0).Table, 80, columnFormatting)
            End If
        End Function

        ''' <summary>
        ''' Suggests column widths for a table using as minimum 2 chars, but minimum header string length, but also either full cell length for number/date/time columns or for all other types 80 % of all values should be visible completely
        ''' </summary>
        ''' <param name="rows">A set of DataRows</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function SuggestColumnWidthsForFixedPlainTables(rows As System.Data.DataRow(), columnFormatting As DataColumnToString) As Integer()
            If rows.Length = 0 Then
                Return Nothing
            Else
                Return SuggestColumnWidthsForFixedPlainTables(rows, rows(0).Table, 80, columnFormatting)
            End If
        End Function
        ''' <summary>
        ''' Suggests column widths for a table using as minimum 2 chars, but minimum header string length, but also either full cell length for number/date/time columns or for all other types a given % value of all values should be visible completely
        ''' </summary>
        ''' <param name="rows">A set of DataRows</param>
        ''' <param name="table">A table</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Shared Function SuggestColumnWidthsForFixedPlainTables(rows As System.Data.DataRow(), ByVal table As DataTable,
                                                                       optimalWidthWhenPercentageNumberOfRowsFitIntoCell As Double,
                                                                       columnFormatting As DataColumnToString) As Integer()
            If table Is Nothing AndAlso rows IsNot Nothing AndAlso rows.Length > 0 Then
                table = rows(0).Table
            End If
            Dim colWidths As New ArrayList
            For ColCounter As Integer = 0 To table.Columns.Count - 1
                Dim MinWidthForHeader As Integer
                If table.Columns(ColCounter).Caption <> Nothing Then
                    MinWidthForHeader = (String.Format("{0}", table.Columns(ColCounter).Caption)).Length
                Else
                    MinWidthForHeader = (String.Format("{0}", table.Columns(ColCounter).ColumnName)).Length
                End If
                Dim MinWidthForCells As Integer
                If rows.Length > 0 Then
                    If table.Columns(ColCounter).DataType.IsValueType AndAlso Not GetType(String).IsInstanceOfType(table.Columns(ColCounter).DataType) Then
                        'number or date/time
                        For RowCounter As Integer = 0 To table.Rows.Count - 1
                            MinWidthForCells = System.Math.Max(MinWidthForCells, String.Format("{0}", table.Rows(RowCounter)(ColCounter)).Length)
                        Next
                    Else
                        'string or any other object
                        Dim cellWidths(table.Rows.Count - 1) As Integer
                        For RowCounter As Integer = 0 To table.Rows.Count - 1
                            Dim RenderValue As Object
                            If columnFormatting Is Nothing Then
                                RenderValue = table.Rows(RowCounter)(ColCounter)
                            Else
                                RenderValue = columnFormatting(table.Columns(ColCounter), table.Rows(RowCounter)(ColCounter))
                            End If
                            cellWidths(RowCounter) = String.Format("{0}", RenderValue).Length
                        Next
                        MinWidthForCells = MaxValueOfFirstXPercent(cellWidths, optimalWidthWhenPercentageNumberOfRowsFitIntoCell)
                    End If
                End If
                colWidths.Add(System.Math.Max(2, System.Math.Max(MinWidthForHeader, MinWidthForCells)))
            Next
            Return CType(colWidths.ToArray(GetType(Integer)), Integer())
        End Function

        ''' <summary>
        ''' Suggests column widths for a table using as minimum 2 chars, but minimum header string length, but also either full cell length for number/date/time columns or for all other types a given % value of all values should be visible completely
        ''' </summary>
        ''' <param name="rows">A set of DataRows</param>
        ''' <param name="table">A table</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Shared Function SuggestColumnWidthsForFixedPlainTables(rows As System.Data.DataRowCollection, table As DataTable,
                                                                       optimalWidthWhenPercentageNumberOfRowsFitIntoCell As Double,
                                                                       columnFormatting As DataColumnToString) As Integer()
            Dim colWidths As New ArrayList
            For ColCounter As Integer = 0 To table.Columns.Count - 1
                Dim MinWidthForHeader As Integer
                If table.Columns(ColCounter).Caption <> Nothing Then
                    MinWidthForHeader = (String.Format("{0}", table.Columns(ColCounter).Caption)).Length
                Else
                    MinWidthForHeader = (String.Format("{0}", table.Columns(ColCounter).ColumnName)).Length
                End If
                Dim MinWidthForCells As Integer
                If rows.Count > 0 Then
                    If table.Columns(ColCounter).DataType.IsValueType AndAlso Not GetType(String).IsInstanceOfType(table.Columns(ColCounter).DataType) Then
                        'number or date/time
                        MinWidthForCells = 1
                        For RowCounter As Integer = 0 To table.Rows.Count - 1
                            MinWidthForCells = System.Math.Max(MinWidthForCells, String.Format("{0}", table.Rows(RowCounter)(ColCounter)).Length)
                        Next
                    Else
                        'string or any other object
                        Dim cellWidths(table.Rows.Count - 1) As Integer
                        For RowCounter As Integer = 0 To table.Rows.Count - 1
                            Dim RenderValue As Object
                            If columnFormatting Is Nothing Then
                                RenderValue = table.Rows(RowCounter)(ColCounter)
                            Else
                                RenderValue = columnFormatting(table.Columns(ColCounter), table.Rows(RowCounter)(ColCounter))
                            End If
                            cellWidths(RowCounter) = String.Format("{0}", RenderValue).Length
                        Next
                        MinWidthForCells = MaxValueOfFirstXPercent(cellWidths, optimalWidthWhenPercentageNumberOfRowsFitIntoCell)
                    End If
                End If
                colWidths.Add(System.Math.Max(2, System.Math.Max(MinWidthForHeader, MinWidthForCells)))
            Next
            Return CType(colWidths.ToArray(GetType(Integer)), Integer())
        End Function

        ''' <summary>
        '''     Return a string with all columns and rows, helpfull for debugging purposes
        ''' </summary>
        ''' <param name="rows">The rows to be processed</param>
        ''' <param name="label">An optional title of the rows</param>
        ''' <param name="fixedColumnWidths">The column sizes in chars</param>
        ''' <returns>All rows are with fixed column withs. If no rows have been processed, the user will get notified about this fact</returns>
        ''' <remarks></remarks>
        Private Shared Function ConvertToPlainTextTableWithFixedColumnWidthsInternal(ByVal rows As DataRow(), ByVal label As String,
                                                                              ByVal fixedColumnWidths As Integer(),
                                                                              columnFormatting As DataColumnToString) As String
            Const vSeparatorHeader As String = "|"
            Const hSeparatorHeader As Char = "-"c
            Const hSeparatorCells As Char = Nothing
            Const vSeparatorCells As String = "|"
            Const crossSeparator As String = "+"
            Return ConvertToPlainTextTableWithFixedColumnWidthsInternal(rows, label, fixedColumnWidths, vSeparatorHeader, vSeparatorCells, crossSeparator, hSeparatorHeader, hSeparatorCells, columnFormatting)
        End Function

        ''' <summary>
        '''     Return a string with all columns and rows, helpfull for debugging purposes
        ''' </summary>
        ''' <param name="rows">The rows to be processed</param>
        ''' <param name="label">An optional title of the rows</param>
        ''' <param name="fixedColumnWidths">The column sizes in chars</param>
        ''' <returns>All rows are with fixed column withs. If no rows have been processed, the user will get notified about this fact</returns>
        ''' <remarks></remarks>
        Private Shared Function ConvertToPlainTextTableWithFixedColumnWidthsInternal(ByVal rows As DataRowCollection, ByVal label As String,
                                                                              ByVal fixedColumnWidths As Integer(),
                                                                              columnFormatting As DataColumnToString) As String
            Const vSeparatorHeader As String = "|"
            Const hSeparatorHeader As Char = "-"c
            Const hSeparatorCells As Char = Nothing
            Const vSeparatorCells As String = "|"
            Const crossSeparator As String = "+"
            Return ConvertToPlainTextTableWithFixedColumnWidthsInternal(rows, label, fixedColumnWidths, vSeparatorHeader, vSeparatorCells, crossSeparator, hSeparatorHeader, hSeparatorCells, columnFormatting)
        End Function

        ''' <summary>
        '''     Return a string with all columns and rows, helpfull for debugging purposes
        ''' </summary>
        ''' <param name="rows">The rows to be processed</param>
        ''' <param name="label">An optional title of the rows</param>
        ''' <param name="fixedColumnWidths">The column sizes in chars</param>
        ''' <param name="verticalSeparatorHeader"></param>
        ''' <param name="verticalSeparatorCells"></param>
        ''' <param name="crossSeparator"></param>
        ''' <param name="horizontalSeparatorHeadline"></param>
        ''' <param name="horizontalSeparatorCells"></param>
        ''' <returns>All rows are with fixed column withs. If no rows have been processed, the user will get notified about this fact</returns>
        ''' <remarks></remarks>
        Private Shared Function ConvertToPlainTextTableWithFixedColumnWidthsInternal(ByVal rows As DataRow(), ByVal label As String,
                                                                              ByVal fixedColumnWidths As Integer(), verticalSeparatorHeader As String,
                                                                              verticalSeparatorCells As String, crossSeparator As String,
                                                                              horizontalSeparatorHeadline As Char, horizontalSeparatorCells As Char,
                                                                              columnFormatting As DataColumnToString) As String
            If Len(verticalSeparatorCells) <> Len(verticalSeparatorHeader) Then Throw New ArgumentException("Length of verticalSeparatorHeader and verticalSeparatorCells must be equal")
            If (Char.GetNumericValue(horizontalSeparatorHeadline) > 0 OrElse Char.GetNumericValue(horizontalSeparatorCells) > 0) AndAlso Len(crossSeparator) <> Len(verticalSeparatorHeader) Then Throw New ArgumentException("Length of verticalSeparatorHeader and crossSeparator must be equal since horizontal lines are requested")
            Dim Result As New System.Text.StringBuilder
            'Add table name
            If label <> "" Then
                Result.Append(String.Format("{0}", label) & System.Environment.NewLine)
            End If
            If rows.Length <= 0 Then
                Result.Append("no rows found" & System.Environment.NewLine)
                Return Result.ToString
            End If
            'Add column headers
            For ColCounter As Integer = 0 To System.Math.Min(rows(0).Table.Columns.Count, fixedColumnWidths.Length) - 1
                Dim column As DataColumn = rows(0).Table.Columns(ColCounter)
                If ColCounter <> 0 Then Result.Append(verticalSeparatorHeader)
                If column.Caption <> Nothing Then
                    Result.Append(TrimStringToFixedWidth(String.Format("{0}", column.Caption), fixedColumnWidths(ColCounter)))
                Else
                    Result.Append(TrimStringToFixedWidth(String.Format("{0}", column.ColumnName), fixedColumnWidths(ColCounter)))
                End If
            Next
            Result.Append(System.Environment.NewLine)
            If horizontalSeparatorHeadline <> Nothing Then
                'Add header separator
                Dim LineSeparatorHeader As String = ""
                For ColCounter As Integer = 0 To System.Math.Min(rows(0).Table.Columns.Count, fixedColumnWidths.Length) - 1
#Disable Warning S1643 ' Strings should not be concatenated using "+" or "&" in a loop
                    If ColCounter <> 0 Then LineSeparatorHeader &= crossSeparator
                    LineSeparatorHeader &= Strings.StrDup(fixedColumnWidths(ColCounter), horizontalSeparatorHeadline)
#Enable Warning S1643 ' Strings should not be concatenated using "+" or "&" in a loop
                Next
                Result.Append(LineSeparatorHeader)
                Result.Append(System.Environment.NewLine)
            End If
            'Add table rows
            For Each row As DataRow In rows
                For ColCounter As Integer = 0 To System.Math.Min(row.Table.Columns.Count, fixedColumnWidths.Length) - 1
                    Dim column As DataColumn = row.Table.Columns(ColCounter)
                    If ColCounter <> 0 Then Result.Append(verticalSeparatorCells)
                    Dim RenderValue As Object
                    If columnFormatting Is Nothing Then
                        RenderValue = row(column)
                    Else
                        RenderValue = columnFormatting(column, row(column))
                    End If
                    Result.Append(TrimStringToFixedWidth(String.Format("{0}", RenderValue), fixedColumnWidths(ColCounter)))
                Next
                Result.Append(System.Environment.NewLine)
                If horizontalSeparatorCells <> Nothing Then
                    'Add lines in between of the cells area
                    Dim LineSeparatorCells As String = ""
                    For ColCounter As Integer = 0 To System.Math.Min(rows(0).Table.Columns.Count, fixedColumnWidths.Length) - 1
#Disable Warning S1643 ' Strings should not be concatenated using "+" or "&" in a loop
                        If ColCounter <> 0 Then LineSeparatorCells &= crossSeparator
                        LineSeparatorCells &= Strings.StrDup(fixedColumnWidths(ColCounter), horizontalSeparatorCells)
#Enable Warning S1643 ' Strings should not be concatenated using "+" or "&" in a loop
                    Next
                    Result.Append(LineSeparatorCells)
                    Result.Append(System.Environment.NewLine)
                End If
            Next
            Return Result.ToString
        End Function

        Public Delegate Function DataColumnToString(column As System.Data.DataColumn, value As Object) As String

        ''' <summary>
        '''     Return a string with all columns and rows, helpfull for debugging purposes
        ''' </summary>
        ''' <param name="rows">The rows to be processed</param>
        ''' <param name="label">An optional title of the rows</param>
        ''' <param name="fixedColumnWidths">The column sizes in chars</param>
        ''' <param name="verticalSeparatorHeader"></param>
        ''' <param name="verticalSeparatorCells"></param>
        ''' <param name="crossSeparator"></param>
        ''' <param name="horizontalSeparatorHeadline"></param>
        ''' <param name="horizontalSeparatorCells"></param>
        ''' <returns>All rows are with fixed column withs. If no rows have been processed, the user will get notified about this fact</returns>
        ''' <remarks></remarks>
        Private Shared Function ConvertToPlainTextTableWithFixedColumnWidthsInternal(ByVal rows As DataRowCollection, ByVal label As String,
                                                                              ByVal fixedColumnWidths As Integer(), verticalSeparatorHeader As String,
                                                                              verticalSeparatorCells As String, crossSeparator As String,
                                                                              horizontalSeparatorHeadline As Char, horizontalSeparatorCells As Char,
                                                                              columnFormatting As DataColumnToString) As String
            If Len(verticalSeparatorCells) <> Len(verticalSeparatorHeader) Then Throw New ArgumentException("Length of verticalSeparatorHeader and verticalSeparatorCells must be equal")
            If (Char.GetNumericValue(horizontalSeparatorHeadline) > 0 OrElse Char.GetNumericValue(horizontalSeparatorCells) > 0) AndAlso Len(crossSeparator) <> Len(verticalSeparatorHeader) Then Throw New ArgumentException("Length of verticalSeparatorHeader and crossSeparator must be equal since horizontal lines are requested")
            Dim Result As New System.Text.StringBuilder
            'Add table name
            If label <> "" Then
                Result.Append(String.Format("{0}", label) & System.Environment.NewLine)
            End If
            If rows.Count <= 0 Then
                Result.Append("no rows found" & System.Environment.NewLine)
                Return Result.ToString
            End If
            'Add column headers
            For ColCounter As Integer = 0 To System.Math.Min(rows(0).Table.Columns.Count, fixedColumnWidths.Length) - 1
                Dim column As DataColumn = rows(0).Table.Columns(ColCounter)
                If ColCounter <> 0 Then Result.Append(verticalSeparatorHeader)
                If column.Caption <> Nothing Then
                    Result.Append(TrimStringToFixedWidth(String.Format("{0}", column.Caption), fixedColumnWidths(ColCounter)))
                Else
                    Result.Append(TrimStringToFixedWidth(String.Format("{0}", column.ColumnName), fixedColumnWidths(ColCounter)))
                End If
            Next
            Result.Append(System.Environment.NewLine)
            If horizontalSeparatorHeadline <> Nothing Then
                'Add header separator
                Dim LineSeparatorHeader As String = ""
                For ColCounter As Integer = 0 To System.Math.Min(rows(0).Table.Columns.Count, fixedColumnWidths.Length) - 1
#Disable Warning S1643 ' Strings should not be concatenated using "+" or "&" in a loop
                    If ColCounter <> 0 Then LineSeparatorHeader &= crossSeparator
                    LineSeparatorHeader &= Strings.StrDup(fixedColumnWidths(ColCounter), horizontalSeparatorHeadline)
#Enable Warning S1643 ' Strings should not be concatenated using "+" or "&" in a loop
                Next
                Result.Append(LineSeparatorHeader)
                Result.Append(System.Environment.NewLine)
            End If
            'Add table rows
            For Each row As DataRow In rows
                For ColCounter As Integer = 0 To System.Math.Min(row.Table.Columns.Count, fixedColumnWidths.Length) - 1
                    Dim column As DataColumn = row.Table.Columns(ColCounter)
                    If ColCounter <> 0 Then Result.Append(verticalSeparatorCells)
                    Dim RenderValue As Object
                    If columnFormatting Is Nothing Then
                        RenderValue = row(column)
                    Else
                        RenderValue = columnFormatting(column, row(column))
                    End If
                    Result.Append(TrimStringToFixedWidth(String.Format("{0}", RenderValue), fixedColumnWidths(ColCounter)))
                Next
                Result.Append(System.Environment.NewLine)
                If horizontalSeparatorCells <> Nothing Then
                    'Add lines in between of the cells area
                    Dim LineSeparatorCells As String = ""
                    For ColCounter As Integer = 0 To System.Math.Min(rows(0).Table.Columns.Count, fixedColumnWidths.Length) - 1
#Disable Warning S1643 ' Strings should not be concatenated using "+" or "&" in a loop
                        If ColCounter <> 0 Then LineSeparatorCells &= crossSeparator
                        LineSeparatorCells &= Strings.StrDup(fixedColumnWidths(ColCounter), horizontalSeparatorCells)
#Enable Warning S1643 ' Strings should not be concatenated using "+" or "&" in a loop
                    Next
                    Result.Append(LineSeparatorCells)
                    Result.Append(System.Environment.NewLine)
                End If
            Next
            Return Result.ToString
        End Function

        ''' <summary>
        ''' Trim the string to a fixed width and concat a string which is too long with triple-dot at the end
        ''' </summary>
        ''' <param name="value"></param>
        ''' <param name="width"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Shared Function TrimStringToFixedWidth(ByVal value As String, ByVal width As Integer) As String
            If value Is Nothing Then value = String.Empty
            If value.Length > width AndAlso width > 3 Then
                Return value.Substring(0, width - 3) & "..."
            Else
                Return Strings.LSet(value, width)
            End If
        End Function

        ''' <summary>
        ''' Lookup a value which is at a given % value position
        ''' </summary>
        ''' <param name="values"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Shared Function MaxValueOfFirstXPercent(values As Integer(), optimalWidthWhenPercentageNumberOfRowsFitIntoCell As Double) As Integer
            If optimalWidthWhenPercentageNumberOfRowsFitIntoCell < 0 OrElse optimalWidthWhenPercentageNumberOfRowsFitIntoCell > 100 Then
                Throw New ArgumentOutOfRangeException(NameOf(optimalWidthWhenPercentageNumberOfRowsFitIntoCell))
            End If
            'Dim sl As New System.Collections.Generic.SortedList(Of Integer, Integer)
            Dim sl As New System.Collections.SortedList
            For MyCounter As Integer = 0 To values.Length - 1
                If sl.ContainsKey(values(MyCounter)) = False Then
                    sl.Add(values(MyCounter), 1)
                End If
            Next
            Dim IndexAtXPercent As Integer
            If optimalWidthWhenPercentageNumberOfRowsFitIntoCell = 100 Then
                IndexAtXPercent = sl.Count - 1
            Else
                IndexAtXPercent = CType((sl.Count - 1) * optimalWidthWhenPercentageNumberOfRowsFitIntoCell / 100, Integer)
            End If
            'Return sl.Keys(IndexAtXPercent)
            Return CInt(sl.GetKey(IndexAtXPercent))
        End Function


        ''' <summary>
        '''     Return a string with all columns and rows, helpfull for debugging purposes
        ''' </summary>
        ''' <param name="rows">The rows to be processed</param>
        ''' <param name="label">An optional title of the rows</param>
        ''' <returns>All rows are tab separated. If no rows have been processed, the user will get notified about this fact</returns>
        Public Shared Function ConvertToPlainTextTable(ByVal rows() As DataRow, ByVal label As String) As String
            Return CompuMaster.Data.DataTablesTools.ConvertToPlainTextTable(rows, label)
        End Function

        ''' <summary>
        '''     Return a string with all columns and rows, helpfull for debugging purposes
        ''' </summary>
        ''' <param name="rows">The rows to be processed</param>
        ''' <param name="label">An optional title of the rows</param>
        ''' <returns>All rows are tab separated. If no rows have been processed, the user will get notified about this fact</returns>
        Public Shared Function ConvertToPlainTextTable(ByVal rows As DataRowCollection, ByVal label As String) As String
            Return CompuMaster.Data.DataTablesTools.ConvertToPlainTextTable(rows, label)
        End Function

        ''' <summary>
        '''     Return a string with all columns and rows, helpfull for debugging purposes
        ''' </summary>
        ''' <param name="rows">The rows to be processed</param>
        ''' <param name="label">An optional title of the rows</param>
        ''' <param name="fixedColumnWidths">The column sizes in chars</param>
        ''' <returns>All rows are tab separated. If no rows have been processed, the user will get notified about this fact</returns>
        <Obsolete("Use ConvertToPlainTextTableFixedColumnWidths instead", False), ComponentModel.EditorBrowsable(ComponentModel.EditorBrowsableState.Never)>
        Public Shared Function ConvertToPlainTextTable(ByVal rows As DataRowCollection, ByVal label As String, ByVal fixedColumnWidths As Integer()) As String
            Return ConvertToPlainTextTableWithFixedColumnWidthsInternal(rows, label, fixedColumnWidths, Nothing)
        End Function

        ''' <summary>
        '''     Return a string with all columns for the specified row in vertical arrangement, helpfull for debugging purposes
        ''' </summary>
        ''' <param name="row">The row to be processed</param>
        ''' <param name="label">An optional title of the rows</param>
        ''' <returns>All columns captions/names are separated from their values by a &quot;: &quot;. All columns are arranged vertically.</returns>
        Public Shared Function ConvertToPlainTextTable(ByVal row As DataRow, ByVal label As String) As String
            Const ColumnSeparator As String = ": "
            Dim MaxLengthOfColumnTitle As Integer = 0
            For Each column As DataColumn In row.Table.Columns
                If column.Caption <> Nothing Then
                    MaxLengthOfColumnTitle = System.Math.Max(MaxLengthOfColumnTitle, column.Caption.Length)
                Else
                    MaxLengthOfColumnTitle = System.Math.Max(MaxLengthOfColumnTitle, column.ColumnName.Length)
                End If
            Next
            Dim Result As New System.Text.StringBuilder
            If label <> "" Then
                Result.Append(String.Format("{0}", label) & System.Environment.NewLine)
            End If
            For Each column As DataColumn In row.Table.Columns
                If column.Caption <> Nothing Then
                    Result.Append(Strings.RSet(String.Format("{0}", column.Caption), MaxLengthOfColumnTitle))
                Else
                    Result.Append(Strings.RSet(String.Format("{0}", column.ColumnName), MaxLengthOfColumnTitle))
                End If
                Result.Append(ColumnSeparator)
                Result.Append(String.Format("{0}", row(column)))
                Result.Append(System.Environment.NewLine)
            Next
            Result.Append(System.Environment.NewLine)
            Return Result.ToString
        End Function

        ''' <summary>
        '''     Convert any opened datareader into a dataset
        ''' </summary>
        ''' <param name="dataReader">An already opened dataReader</param>
        ''' <returns>A dataset containing all datatables the dataReader was able to read</returns>
        Public Shared Function ConvertDataReaderToDataSet(ByVal datareader As IDataReader) As DataSet
            Return CompuMaster.Data.DataTablesTools.ConvertDataReaderToDataSet(datareader)
        End Function

        ''' <summary>
        '''     Convert any opened datareader into a data table
        ''' </summary>
        ''' <param name="dataReader">An already opened dataReader</param>
        ''' <returns>A data table containing all data the dataReader was able to read</returns>
        Public Shared Function ConvertDataReaderToDataTable(ByVal dataReader As IDataReader) As DataTable
            Return CompuMaster.Data.DataTablesTools.ConvertDataReaderToDataTable(dataReader)
        End Function

        ''' <summary>
        '''     Convert any opened datareader into a data table
        ''' </summary>
        ''' <param name="dataReader">An already opened dataReader</param>
        ''' <param name="tableName">The name for the new table</param>
        ''' <returns>A data table containing all data the dataReader was able to read</returns>
        Public Shared Function ConvertDataReaderToDataTable(ByVal dataReader As IDataReader, ByVal tableName As String) As DataTable
            Return CompuMaster.Data.DataTablesTools.ConvertDataReaderToDataTable(dataReader, tableName)
        End Function

        ''' <summary>
        '''     Table join types
        ''' </summary>
        Public Enum JoinTypes As Integer
            Inner = 0
            Left = 1
        End Enum

        Public Enum SqlJoinTypes As Byte
            Inner = 0
            Left = 1
            Right = 2
            FullOuter = 3
            Cross = 4
        End Enum

        ''' <summary>
        '''     Execute a table join on two tables of the same dataset based on the first relation found
        ''' </summary>
        ''' <param name="leftParentTable"></param>
        ''' <param name="rightChildTable"></param>
        ''' <param name="joinType">Inner or left join</param>
        ''' <returns></returns>
        Public Shared Function JoinTables(ByVal leftParentTable As DataTable, ByVal rightChildTable As DataTable, ByVal joinType As JoinTypes) As DataTable
            Return CompuMaster.Data.DataTablesTools.JoinTables(leftParentTable, rightChildTable, CType(joinType, CompuMaster.Data.DataTablesTools.JoinTypes))
        End Function

        ''' <summary>
        '''     Execute a table join on two tables of the same dataset which have got a defined relation
        ''' </summary>
        ''' <param name="leftParentTable">The left or parent table</param>
        ''' <param name="rightChildTable">The right or child table</param>
        ''' <param name="relation">A data table relation which shall be used for the joining</param>
        ''' <param name="joinType">Inner or left join</param>
        ''' <returns></returns>
        ''' <remarks>
        '''     The selected columns are: 
        '''     <ul>
        '''         <li>all columns from the left parent table</li>
        '''         <li>INNER JOIN: those columns from the right child table which are not member of the relation in charge</li>
        '''         <li>LEFT JOIN: all columns from the right child table</li>
        '''     </ul>
        ''' </remarks>
        Public Shared Function JoinTables(ByVal leftParentTable As DataTable, ByVal rightChildTable As DataTable,
                                          ByVal relation As DataRelation, ByVal joinType As JoinTypes) As DataTable
            Return CompuMaster.Data.DataTablesTools.JoinTables(leftParentTable, rightChildTable, relation, CType(joinType, CompuMaster.Data.DataTablesTools.JoinTypes))
        End Function

        ''' <summary>
        '''     Execute a table join on two tables of the same dataset which have got a defined relation
        ''' </summary>
        ''' <param name="leftParentTable">The left or parent table</param>
        ''' <param name="leftTableColumnsToCopy">An array of columns to copy from the left table</param>
        ''' <param name="rightChildTable">The right or child table</param>
        ''' <param name="rightTableColumnsToCopy">An array of columns to copy from the right table</param>
        ''' <param name="joinType">Inner or left join</param>
        ''' <returns></returns>
        Public Shared Function JoinTables(ByVal leftParentTable As DataTable, ByVal leftTableColumnsToCopy As DataColumn(), ByVal rightChildTable As DataTable,
                                          ByVal rightTableColumnsToCopy As DataColumn(), ByVal joinType As JoinTypes) As DataTable
            Return CompuMaster.Data.DataTablesTools.JoinTables(leftParentTable, leftTableColumnsToCopy, rightChildTable, rightTableColumnsToCopy, CType(joinType, CompuMaster.Data.DataTablesTools.JoinTypes))
        End Function

        ''' <summary>
        '''     Execute a table join on two tables of the same dataset which have got a defined relation
        ''' </summary>
        ''' <param name="leftParentTable">The left or parent table</param>
        ''' <param name="indexesOfLeftTableColumnsToCopy">An array of column indexes to copy from the left table</param>
        ''' <param name="rightChildTable">The right or child table</param>
        ''' <param name="indexesOfRightTableColumnsToCopy">An array of column indexes to copy from the right table</param>
        ''' <param name="joinType">Inner or left join</param>
        ''' <returns></returns>
        Public Shared Function JoinTables(ByVal leftParentTable As DataTable, ByVal indexesOfLeftTableColumnsToCopy As Integer(),
                                          ByVal rightChildTable As DataTable, ByVal indexesOfRightTableColumnsToCopy As Integer(),
                                          ByVal joinType As JoinTypes) As DataTable
            Return CompuMaster.Data.DataTablesTools.JoinTables(leftParentTable, indexesOfLeftTableColumnsToCopy, rightChildTable, indexesOfRightTableColumnsToCopy, CType(joinType, CompuMaster.Data.DataTablesTools.JoinTypes))
        End Function

        ''' <summary>
        '''     Execute a table join on two tables of the same dataset which have got a defined relation
        ''' </summary>
        ''' <param name="leftParentTable">The left or parent table</param>
        ''' <param name="indexesOfLeftTableColumnsToCopy">An array of column indexes to copy from the left table</param>
        ''' <param name="rightChildTable">The right or child table</param>
        ''' <param name="indexesOfRightTableColumnsToCopy">An array of column indexes to copy from the right table</param>
        ''' <param name="relation">A data table relation which shall be used for the joining</param>
        ''' <param name="joinType">Inner or left join</param>
        ''' <returns></returns>
        Public Shared Function JoinTables(ByVal leftParentTable As DataTable, ByVal indexesOfLeftTableColumnsToCopy As Integer(),
                                          ByVal rightChildTable As DataTable, ByVal indexesOfRightTableColumnsToCopy As Integer(),
                                          ByVal relation As DataRelation, ByVal joinType As JoinTypes) As DataTable
            Return CompuMaster.Data.DataTablesTools.JoinTables(leftParentTable, indexesOfLeftTableColumnsToCopy, rightChildTable, indexesOfRightTableColumnsToCopy, relation, CType(joinType,
                                                               CompuMaster.Data.DataTablesTools.JoinTypes))
        End Function

        ''' <summary>
        '''     Cross join of two tables
        ''' </summary>
        ''' <param name="leftTable">A first datatable</param>
        ''' <param name="indexesOfLeftTableColumnsToCopy">An array of column indexes to copy from the left table</param>
        ''' <param name="rightTable">A second datatable</param>
        ''' <param name="indexesOfRightTableColumnsToCopy">An array of column indexes to copy from the right table</param>
        ''' <returns></returns>
        Public Shared Function CrossJoinTables(ByVal leftTable As DataTable, ByVal indexesOfLeftTableColumnsToCopy As Integer(),
                                               ByVal rightTable As DataTable, ByVal indexesOfRightTableColumnsToCopy As Integer()) As DataTable
            Return CompuMaster.Data.DataTablesTools.CrossJoinTables(leftTable, indexesOfLeftTableColumnsToCopy, rightTable, indexesOfRightTableColumnsToCopy)
        End Function

        ''' <summary>
        ''' Create a new table using a full outer join
        ''' </summary>
        ''' <param name="leftTable">1st table</param>
        ''' <param name="rightTable">2nd table</param>
        ''' <returns></returns>
        ''' <remarks>The primary key columns of both tables are used to find the corrorresponding matches</remarks>
        <Obsolete("Use CompuMaster.Data.DataTables.SqlJoinTables instead", False), System.ComponentModel.EditorBrowsable(ComponentModel.EditorBrowsableState.Never)>
        Public Shared Function FullJoinTables(ByVal leftTable As DataTable, ByVal rightTable As DataTable) As DataTable
            Return FullJoinTables(leftTable, leftTable.PrimaryKey, rightTable, rightTable.PrimaryKey)
        End Function

        ''' <summary>
        ''' Create a new table using a full outer join
        ''' </summary>
        ''' <param name="leftTable">1st table</param>
        ''' <param name="leftKeyColumns">The key columns which shall be used for finding matches in the 2nd table</param>
        ''' <param name="rightTable">2nd table</param>
        ''' <param name="rightKeyColumns">The key columns which shall be used for finding matches in the 1st table</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <Obsolete("Use CompuMaster.Data.DataTables.SqlJoinTables instead", False), System.ComponentModel.EditorBrowsable(ComponentModel.EditorBrowsableState.Never)>
        Public Shared Function FullJoinTables(ByVal leftTable As DataTable, ByVal leftKeyColumns As String(), ByVal rightTable As DataTable,
                                              ByVal rightKeyColumns As String()) As DataTable
            Dim leftIndexes As New ArrayList, rightIndexes As New ArrayList
            For MyCounter As Integer = 0 To leftKeyColumns.Length - 1
                leftIndexes.Add(leftTable.Columns(leftKeyColumns(MyCounter)).Ordinal)
            Next
            For MyCounter As Integer = 0 To rightKeyColumns.Length - 1
                rightIndexes.Add(rightTable.Columns(rightKeyColumns(MyCounter)).Ordinal)
            Next
            Return FullJoinTables(leftTable, CType(leftIndexes.ToArray(GetType(Integer)), Integer()), rightTable, CType(rightIndexes.ToArray(GetType(Integer)), Integer()))
        End Function

        ''' <summary>
        ''' Create a new table using a full outer join
        ''' </summary>
        ''' <param name="leftTable">1st table</param>
        ''' <param name="leftKeyColumns">The key columns which shall be used for finding matches in the 2nd table</param>
        ''' <param name="rightTable">2nd table</param>
        ''' <param name="rightKeyColumns">The key columns which shall be used for finding matches in the 1st table</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <Obsolete("Use CompuMaster.Data.DataTables.SqlJoinTables instead", False), System.ComponentModel.EditorBrowsable(ComponentModel.EditorBrowsableState.Never)>
        Public Shared Function FullJoinTables(ByVal leftTable As DataTable, ByVal leftKeyColumns As DataColumn(), ByVal rightTable As DataTable,
                                              ByVal rightKeyColumns As DataColumn()) As DataTable
            Dim leftIndexes As New ArrayList, rightIndexes As New ArrayList
            For MyCounter As Integer = 0 To leftKeyColumns.Length - 1
                If leftTable Is leftKeyColumns(MyCounter).Table Then
                    Throw New Exception("Table mismatch: data column must be referencing to the same data table")
                End If
                leftIndexes.Add(leftKeyColumns(MyCounter).Ordinal)
            Next
            For MyCounter As Integer = 0 To rightKeyColumns.Length - 1
                If rightTable Is rightKeyColumns(MyCounter).Table Then
                    Throw New Exception("Table mismatch: data column must be referencing to the same data table")
                End If
                rightIndexes.Add(rightKeyColumns(MyCounter).Ordinal)
            Next
            Return FullJoinTables(leftTable, CType(leftIndexes.ToArray(GetType(Integer)), Integer()), rightTable, CType(rightIndexes.ToArray(GetType(Integer)), Integer()))
        End Function

        ''' <summary>
        ''' Create a new table using a full outer join and case-insensitive string-comparison mode
        ''' </summary>
        ''' <param name="leftTable">1st table</param>
        ''' <param name="leftKeyColumnIndexes">The key columns which shall be used for finding matches in the 2nd table</param>
        ''' <param name="rightTable">2nd table</param>
        ''' <param name="rightKeyColumnIndexes">The key columns which shall be used for finding matches in the 1st table</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <Obsolete("Use CompuMaster.Data.DataTables.SqlJoinTables instead", False), System.ComponentModel.EditorBrowsable(ComponentModel.EditorBrowsableState.Never)>
        Public Shared Function FullJoinTables(ByVal leftTable As DataTable, ByVal leftKeyColumnIndexes As Integer(), ByVal rightTable As DataTable,
                                              ByVal rightKeyColumnIndexes As Integer()) As DataTable
            Return FullJoinTables(leftTable, leftKeyColumnIndexes, rightTable, rightKeyColumnIndexes, True)
        End Function

        ''' <summary>
        ''' Create a new table using a full outer join
        ''' </summary>
        ''' <param name="leftTable">1st table</param>
        ''' <param name="leftKeyColumnIndexes">The key columns which shall be used for finding matches in the 2nd table</param>
        ''' <param name="rightTable">2nd table</param>
        ''' <param name="rightKeyColumnIndexes">The key columns which shall be used for finding matches in the 1st table</param>
        ''' <param name="compareStringsCaseInsensitive">True to compare strings case insensitive, False for case sensitive</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <Obsolete("Use CompuMaster.Data.DataTables.SqlJoinTables instead", False), System.ComponentModel.EditorBrowsable(ComponentModel.EditorBrowsableState.Never)>
        Public Shared Function FullJoinTables(ByVal leftTable As DataTable, ByVal leftKeyColumnIndexes As Integer(), ByVal rightTable As DataTable,
                                              ByVal rightKeyColumnIndexes As Integer(), ByVal compareStringsCaseInsensitive As Boolean) As DataTable
            'Parameter validation
            If leftTable Is Nothing Then Throw New ArgumentException("Missing argument", NameOf(leftTable))
            If leftKeyColumnIndexes Is Nothing OrElse leftKeyColumnIndexes.Length = 0 Then Throw New ArgumentException("Missing argument", NameOf(leftKeyColumnIndexes))
            If rightTable Is Nothing Then Throw New ArgumentException("Missing argument", NameOf(rightTable))
            If rightKeyColumnIndexes Is Nothing OrElse rightKeyColumnIndexes.Length = 0 Then Throw New ArgumentException("Missing argument", NameOf(rightKeyColumnIndexes))
            If leftKeyColumnIndexes.Length <> rightKeyColumnIndexes.Length Then Throw New Exception("Count of leftKeyColumnIndexes must be equal to count of rightKeyColumnIndexes")
            For MyCounter As Integer = 0 To leftKeyColumnIndexes.Length - 1
                If leftTable.Columns(leftKeyColumnIndexes(MyCounter)).DataType IsNot rightTable.Columns(rightKeyColumnIndexes(MyCounter)).DataType Then
                    Throw New Exception("Data types of key columns must be equal")
                End If
            Next
            With Nothing 'Ensure unique index numbers
                Dim leftIndexes As New ArrayList, rightIndexes As New ArrayList
                For MyCounter As Integer = 0 To leftKeyColumnIndexes.Length - 1
                    If leftIndexes.Contains(leftKeyColumnIndexes(MyCounter)) = True Then
                        Throw New Exception("Duplicate data column with index " & leftKeyColumnIndexes(MyCounter))
                    Else
                        leftIndexes.Add(leftKeyColumnIndexes(MyCounter))
                    End If
                Next
                For MyCounter As Integer = 0 To rightKeyColumnIndexes.Length - 1
                    If rightIndexes.Contains(rightKeyColumnIndexes(MyCounter)) = True Then
                        Throw New Exception("Duplicate data column with index " & rightKeyColumnIndexes(MyCounter))
                    Else
                        rightIndexes.Add(rightKeyColumnIndexes(MyCounter))
                    End If
                Next
            End With

            'Prepare Result table scheme
            Dim Result As New DataTable("JoinedTable")
            CopyColumnScheme(leftTable, Result, True)
            CopyColumnScheme(rightTable, Result, False)

            'Hint for way of implementation: FULL OUTER JOIN = LEFT JOIN (where RIGHT IS NULL) + RIGHT JOIN (where LEFT IS NULL)

            Dim AssignedRowIndexesOfRightTable As New ArrayList
            For LeftTableCounter As Integer = 0 To leftTable.Rows.Count - 1
                Dim RightTableRowFoundWithRowIndex As Integer = -1
                'Compare to find row matches
                For RightTableCounter As Integer = 0 To rightTable.Rows.Count - 1
                    Dim leftRow As DataRow = leftTable.Rows(LeftTableCounter)
                    Dim rightRow As DataRow = rightTable.Rows(RightTableCounter)
                    Dim ComparisonResult As Boolean = True
                    For KeyColCounter As Integer = 0 To leftKeyColumnIndexes.Length - 1
                        If CompareValuesOfUnknownType(leftRow(leftKeyColumnIndexes(KeyColCounter)), rightRow(rightKeyColumnIndexes(KeyColCounter)), True) = False Then
                            ComparisonResult = False
                            Exit For
                        End If
                    Next
                    If ComparisonResult = True Then
                        RightTableRowFoundWithRowIndex = RightTableCounter
                        Exit For
                    End If
                Next
                'Add the row as a match has been found
                If RightTableRowFoundWithRowIndex = -1 Then
                    'No right row has been found -> left table row only
                    Dim NewRow As DataRow = Result.NewRow
                    For MyCounter As Integer = 0 To leftTable.Columns.Count - 1
                        NewRow(MyCounter) = leftTable.Rows(LeftTableCounter)(MyCounter)
                    Next
                    For RightColCopyCounter As Integer = 0 To rightTable.Columns.Count - 1
                        NewRow(leftTable.Columns.Count + RightColCopyCounter) = DBNull.Value
                    Next
                    Result.Rows.Add(NewRow)
                Else
                    'Right row has been found -> combine left and right table rows
                    Dim NewRow As DataRow = Result.NewRow
                    For LeftColCopyCounter As Integer = 0 To leftTable.Columns.Count - 1
                        NewRow(LeftColCopyCounter) = leftTable.Rows(LeftTableCounter)(LeftColCopyCounter)
                    Next
                    For RightColCopyCounter As Integer = 0 To rightTable.Columns.Count - 1
                        NewRow(leftTable.Columns.Count + RightColCopyCounter) = rightTable.Rows(RightTableRowFoundWithRowIndex)(RightColCopyCounter)
                    Next
                    Result.Rows.Add(NewRow)
                    'Mark the right row as being assigned
                    AssignedRowIndexesOfRightTable.Add(RightTableRowFoundWithRowIndex)
                End If
            Next
            For RightTableCounter As Integer = 0 To rightTable.Rows.Count - 1
                If AssignedRowIndexesOfRightTable.Contains(RightTableCounter) Then
                    'Already assigned - we don't need that row here
                Else
                    'Row must be appended to the result table
                    Dim NewRow As DataRow = Result.NewRow
                    For LeftColCopyCounter As Integer = 0 To leftTable.Columns.Count - 1
                        NewRow(LeftColCopyCounter) = DBNull.Value
                    Next
                    For RightColCopyCounter As Integer = 0 To rightTable.Columns.Count - 1
                        NewRow(leftTable.Columns.Count + RightColCopyCounter) = rightTable.Rows(RightTableCounter)(RightColCopyCounter)
                    Next
                    Result.Rows.Add(NewRow)
                End If
            Next
            Return Result
        End Function

        ''' <summary>
        ''' Copy the column collection from a template table to a destination table
        ''' </summary>
        ''' <param name="templateTable"></param>
        ''' <param name="destinationTable"></param>
        ''' <remarks>The data scheme is copied, but contraints are removed</remarks>
        Private Shared Sub CopyColumnScheme(ByVal templateTable As DataTable, ByVal destinationTable As DataTable, ByVal initialSchemaFill As Boolean)
            For MyColCounter As Integer = 0 To templateTable.Columns.Count - 1
                Dim col As DataColumn = CloneDataColumn(templateTable.Columns(MyColCounter))
                If initialSchemaFill Then
                    'Never change the column names (even if the template table already contains 2 or more columns with the same name)
                Else
                    'Change the column names to provide unique column names
                    Dim newColName As String = LookupUniqueColumnName(destinationTable, col.ColumnName)
                    If col.ColumnName <> newColName Then
                        'Also change the caption of the column (because it's equal to its name)
                        If col.Caption = col.ColumnName Then
                            col.Caption = newColName
                        End If
                        col.ColumnName = newColName
                    End If
                End If
                destinationTable.Columns.Add(col)
            Next
        End Sub

        ''' <summary>
        ''' Create a clone of a DataColumn except identities, mappings and constraints
        ''' </summary>
        ''' <param name="templateColumn"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Shared Function CloneDataColumn(ByVal templateColumn As DataColumn) As DataColumn
            Dim Result As New DataColumn With {
                .AllowDBNull = templateColumn.AllowDBNull,
                .AutoIncrement = False,
                .Caption = templateColumn.Caption,
                .ColumnName = templateColumn.ColumnName,
                .DataType = templateColumn.DataType,
                .DefaultValue = templateColumn.DefaultValue,
                .MaxLength = templateColumn.MaxLength,
                .ReadOnly = templateColumn.ReadOnly,
                .Unique = False,
                .DateTimeMode = templateColumn.DateTimeMode
            }
            Return Result
        End Function

        ''' <summary>
        ''' Ensure the string is a valid value (never a null (Nothing in VisualBasic))
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns>String.Empty for values which are null (Nothing in VisualBasic) or otherwise the value as it is</returns>
        ''' <remarks></remarks>
        Private Shared Function StringNotNothingOrEmpty(ByVal value As String) As String
            If value Is Nothing Then
                Return String.Empty
            Else
                Return value
            End If
        End Function

        ''' <summary>
        ''' Compare 2 values of unknown but same type
        ''' </summary>
        ''' <param name="value1">1st value</param>
        ''' <param name="value2">2nd value</param>
        ''' <param name="compareStringsCaseInsensitive">True to compare strings case insensitive, False for case sensitive</param>
        ''' <returns></returns>
        ''' <remarks>Comparisons with DBNull.Value will return False or True, never DBNull.Value</remarks>
        Public Shared Function CompareValuesOfUnknownType(ByVal value1 As Object, ByVal value2 As Object, ByVal compareStringsCaseInsensitive As Boolean) As Boolean
            Dim TypeCheckValue As Object
            If value1 Is Nothing Then
                TypeCheckValue = value2
            Else
                TypeCheckValue = value1
            End If
            If value1 Is DBNull.Value OrElse value2 Is DBNull.Value Then
                'At least 1 DBNull is present
                If value1 Is DBNull.Value AndAlso value2 Is DBNull.Value Then
                    'DBNull at both sides lead to True result
                    Return True
                Else
                    'DBNull only at one 1 side leads to False result
                    Return False
                End If
            ElseIf value1 Is Nothing AndAlso value2 Is Nothing Then
                Return True
            ElseIf TypeCheckValue.GetType Is GetType(String) Then
                'Strings
                If compareStringsCaseInsensitive = False Then
                    If CType(value1, String) <> CType(value2, String) Then
                        Return False
                    End If
                Else
                    If StringNotNothingOrEmpty(CType(value1, String)).ToLower(Globalization.CultureInfo.InvariantCulture) <> StringNotNothingOrEmpty(CType(value2, String)).ToLower(Globalization.CultureInfo.InvariantCulture) Then
                        Return False
                    End If
                End If
            ElseIf TypeCheckValue.GetType Is GetType(System.Double) Then
                'Doubles
                If CType(value1, System.Double) <> CType(value2, System.Double) Then
                    Return False
                End If
            ElseIf TypeCheckValue.GetType Is GetType(System.Decimal) Then
                'Decimals
                If CType(value1, System.Decimal) <> CType(value2, System.Decimal) Then
                    Return False
                End If
            ElseIf TypeCheckValue.GetType Is GetType(System.DateTime) Then
                'Datetime
                If CType(value1, System.DateTime) <> CType(value2, System.DateTime) Then
                    Return False
                End If
            ElseIf TypeCheckValue.GetType Is GetType(System.Int16) OrElse value1 Is GetType(System.Int32) OrElse value1 Is GetType(System.Int64) Then
                'Intxx
                If CType(value1, System.Int64) <> CType(value2, System.Int64) Then
                    Return False
                End If
            ElseIf TypeCheckValue.GetType Is GetType(System.UInt16) OrElse value1 Is GetType(System.UInt32) OrElse value1 Is GetType(System.UInt64) Then
                'UIntxx
                If CType(value1, System.UInt64).CompareTo(CType(value2, System.UInt64)) <> 0 Then
                    Return False
                End If
            Else
                'Other data types
                If value1 IsNot value2 Then
                    'Other data types which do not require textual handling
                    Return False
                End If
            End If
            Return True
        End Function

        ''' <summary>
        ''' Find unique values in a column
        ''' </summary>
        ''' <param name="column">The DataColumn which holds the data</param>
        ''' <returns></returns>
        Public Shared Function FindUniqueValues(ByVal column As DataColumn) As ArrayList
            Return FindUniqueValues(column, False)
        End Function

        ''' <summary>
        ''' Returns unique values in a column
        ''' </summary>
        ''' <param name="column">The DataColumn which holds the data</param>
        ''' <param name="ignoreDBNull">True never results a DBNull value</param>
        ''' <returns></returns>
        Public Shared Function FindUniqueValues(ByVal column As DataColumn, ByVal ignoreDBNull As Boolean) As ArrayList
            Dim table As DataTable = column.Table
            Dim Result As New ArrayList
            For MyCounter As Integer = 0 To table.Rows.Count - 1
                If ignoreDBNull = True AndAlso IsDBNull(table.Rows(MyCounter)(column)) Then
                    'do not add DbNulls to result
                ElseIf Not Result.Contains(table.Rows(MyCounter)(column)) Then
                    Result.Add(table.Rows(MyCounter)(column))
                End If
            Next
            Return Result
        End Function

        ''' <summary>
        ''' Find unique values in a column
        ''' </summary>
        ''' <param name="column">The DataColumn which holds the data</param>
        ''' <returns></returns>
        Public Shared Function FindUniqueValues(Of T)(ByVal column As DataColumn) As System.Collections.Generic.List(Of T)
            Return FindUniqueValues(Of T)(column, False)
        End Function

        ''' <summary>
        ''' Returns unique values in a column
        ''' </summary>
        ''' <param name="column">The DataColumn which holds the data</param>
        ''' <param name="ignoreDBNull">True never results a DBNull value</param>
        ''' <returns></returns>
        Public Shared Function FindUniqueValues(Of T)(ByVal column As DataColumn, ByVal ignoreDBNull As Boolean) As System.Collections.Generic.List(Of T)
            Dim table As DataTable = column.Table
            Dim Result As New System.Collections.Generic.List(Of T)
            For MyCounter As Integer = 0 To table.Rows.Count - 1
                If ignoreDBNull = True AndAlso IsDBNull(table.Rows(MyCounter)(column)) Then
                    'do not add DbNulls to result
                ElseIf ignoreDBNull = False AndAlso IsDBNull(table.Rows(MyCounter)(column)) Then
                    Result.Add(Nothing)
                Else
                    Dim Value As T = CType(table.Rows(MyCounter)(column), T)
                    If Not Result.Contains(Value) Then
                        Result.Add(Value)
                    End If
                End If
            Next
            Return Result
        End Function

        ''' <summary>
        '''     Add the specified columns if they don't exist
        ''' </summary>
        ''' <param name="datatable">A datatable where the operations shall be made</param>
        ''' <param name="columnName">The name of the String column which shall be added</param>
        ''' <remarks>
        '''     The columns will only be added if they don't exist. If a column name exists, it will be ignored.
        ''' </remarks>
        Public Shared Sub AddColumns(ByVal datatable As System.Data.DataTable, ByVal columnName As String)
            Dim NewColumn As New DataColumn(columnName, GetType(String))
            AddColumns(datatable, New DataColumn() {NewColumn})
        End Sub

        ''' <summary>
        '''     Add the specified columns if they don't exist
        ''' </summary>
        ''' <param name="datatable">A datatable where the operations shall be made</param>
        ''' <param name="columnName">The name of the column which shall be added</param>
        ''' <param name="dataType">The type of the column which shall be added</param>
        ''' <remarks>
        '''     The columns will only be added if they don't exist. If a column name exists, it will be ignored.
        ''' </remarks>
        Public Shared Sub AddColumns(ByVal datatable As System.Data.DataTable, ByVal columnName As String, dataType As Type)
            Dim NewColumn As New DataColumn(columnName, dataType)
            AddColumns(datatable, New DataColumn() {NewColumn})
        End Sub

        ''' <summary>
        '''     Add the specified columns if they don't exist
        ''' </summary>
        ''' <param name="datatable">A datatable where the operations shall be made</param>
        ''' <param name="columnNames">The names of the columns which shall be added</param>
        ''' <remarks>
        '''     The columns will only be added if they don't exist. If a column name exists, it will be ignored.
        ''' </remarks>
        Public Shared Sub AddColumns(ByVal datatable As System.Data.DataTable, ByVal columnNames As String())
            Dim NewColumns As New System.Collections.Generic.List(Of DataColumn)
            For Each ColumnName As String In columnNames
                NewColumns.Add(New DataColumn(ColumnName, GetType(String)))
            Next
            AddColumns(datatable, NewColumns.ToArray)
        End Sub

        ''' <summary>
        '''     Add the specified columns if they don't exist
        ''' </summary>
        ''' <param name="datatable">A datatable where the operations shall be made</param>
        ''' <param name="columns">The columns which shall be added</param>
        ''' <remarks>
        '''     The columns will only be added if they don't exist. If a column name exists, it will be ignored.
        ''' </remarks>
        Public Shared Sub AddColumns(ByVal datatable As System.Data.DataTable, ByVal columns As DataColumn())
            For Each Column As DataColumn In columns
                If datatable.Columns.Contains(Column.ColumnName) = False Then
                    datatable.Columns.Add(Column)
                End If
            Next
        End Sub

        ''' <summary>
        '''     Add the specified columns if they don't exist
        ''' </summary>
        ''' <param name="datatable">A datatable where the operations shall be made</param>
        ''' <param name="columnNames">The columns which shall be added</param>
        ''' <remarks>
        '''     The columns will only be added if they don't exist. If a column name exists, it will be ignored.
        ''' </remarks>
        Public Shared Sub AddColumns(ByVal datatable As System.Data.DataTable, ByVal columnNames As String(), dataType As Type)
            Dim NewColumns As New System.Collections.Generic.List(Of DataColumn)
            For Each ColumnName As String In columnNames
                NewColumns.Add(New DataColumn(ColumnName, dataType))
            Next
            AddColumns(datatable, NewColumns.ToArray)
        End Sub

        ''' <summary>
        '''     Add a prefix to the names of the columns
        ''' </summary>
        ''' <param name="dataTable">A datatable which shall be exported</param>
        ''' <param name="columnIndexes">An array of column indexes</param>
        ''' <param name="prefix">e. g. "orders."</param>
        Public Shared Sub AddPrefixesToColumnNames(ByVal dataTable As DataTable, ByVal columnIndexes As Integer(), ByVal prefix As String)
            CompuMaster.Data.DataTablesTools.AddPrefixesToColumnNames(dataTable, columnIndexes, prefix)
        End Sub

        ''' <summary>
        '''     Add a suffix to the names of the columns
        ''' </summary>
        ''' <param name="dataTable">A datatable which shall be exported</param>
        ''' <param name="columnIndexes">An array of column indexes</param>
        ''' <param name="suffix">e. g. "-orders"</param>
        Public Shared Sub AddSuffixesToColumnNames(ByVal dataTable As DataTable, ByVal columnIndexes As Integer(), ByVal suffix As String)
            CompuMaster.Data.DataTablesTools.AddSuffixesToColumnNames(dataTable, columnIndexes, suffix)
        End Sub

        ''' <summary>
        '''     An exception which gets thrown when converting data in the ReArrangeDataColumns methods
        ''' </summary>
        <CodeAnalysis.SuppressMessage("Design", "CA1032:Implement standard exception constructors", Justification:="<Ausstehend>")>
        <CodeAnalysis.SuppressMessage("Usage", "CA2237:Mark ISerializable types with serializable", Justification:="<Ausstehend>")>
        Public Class ReArrangeDataColumnsException
            Inherits Exception

            Private ReadOnly MyCMToolsReArrangeDataColumnsException As CompuMaster.Data.ReArrangeDataColumnsException

            Public Sub New(ByVal rowIndex As Integer, ByVal columnIndex As Integer, ByVal sourceColumnType As Type, ByVal targetColumnType As Type,
                           ByVal problematicValue As Object, ByVal innerException As Exception)
                MyCMToolsReArrangeDataColumnsException = New CompuMaster.Data.ReArrangeDataColumnsException(rowIndex, columnIndex, sourceColumnType, targetColumnType,
                                                                                                            problematicValue, innerException)
            End Sub

            Public ReadOnly Property TargetColumnType() As Type
                Get
                    Return MyCMToolsReArrangeDataColumnsException.TargetColumnType
                End Get
            End Property

            Public ReadOnly Property ProblematicValue() As Object
                Get
                    Return MyCMToolsReArrangeDataColumnsException.ProblematicValue
                End Get
            End Property

            Public ReadOnly Property RowIndex() As Integer
                Get
                    Return MyCMToolsReArrangeDataColumnsException.RowIndex
                End Get
            End Property

            Public ReadOnly Property ColumnIndex() As Integer
                Get
                    Return MyCMToolsReArrangeDataColumnsException.ColumnIndex
                End Get
            End Property

            Public Overrides ReadOnly Property Message() As String
                Get
                    Return MyCMToolsReArrangeDataColumnsException.Message
                End Get
            End Property

        End Class

        ''' <summary>
        '''     Rearrange columns
        ''' </summary>
        ''' <param name="source">The source table with data</param>
        ''' <param name="columnsToCopy">An array of column names which shall be copied in the specified order from the source table</param>
        ''' <returns>A new and independent data table with copied data</returns>
        Public Shared Function ReArrangeDataColumns(ByVal source As DataTable, ByVal columnsToCopy As String()) As DataTable
            Return CompuMaster.Data.DataTablesTools.ReArrangeDataColumns(source, columnsToCopy)
        End Function

        ''' <summary>
        '''     Rearrange columns and also change their data types
        ''' </summary>
        ''' <param name="source">The source table with data</param>
        ''' <param name="destinationColumnSet">An array of columns as they shall be inserted into the result</param>
        ''' <returns>A new and independent data table with copied data</returns>
        ''' <remarks>
        '''     The copy process requires that the names of the destination columns can be found in the columns collection of the source table. 
        ''' </remarks>
        ''' <example>
        '''     <code language="vb">
        '''         ReArrangeDataColumns(source, New System.Data.DataColumn() {New DataColumn("column1Name", GetType(String)), New DataColumn("column2Name", GetType(Integer))})
        '''     </code>
        ''' </example>
        Public Shared Function ReArrangeDataColumns(ByVal source As DataTable, ByVal destinationColumnSet As DataColumn()) As DataTable
            Return CompuMaster.Data.DataTablesTools.ReArrangeDataColumns(source, destinationColumnSet)
        End Function

        ''' <summary>
        '''     Rearrange columns and also change their data types
        ''' </summary>
        ''' <param name="source">The source table with data</param>
        ''' <param name="destinationColumnSet">An array of columns as they shall be inserted into the result</param>
        ''' <param name="ignoreConversionExceptionAndLogThemHere">In case of data conversion exceptions, log them here instead of throwing them immediately</param>
        ''' <returns>A new and independent data table with copied data</returns>
        ''' <remarks>
        '''     The copy process requires that the names of the destination columns can be found in the columns collection of the source table. 
        ''' </remarks>
        ''' <example>
        '''     <code language="vb">
        '''         ReArrangeDataColumns(source, New System.Data.DataColumn() {New DataColumn("column1Name", GetType(String)), New DataColumn("column2Name", GetType(Integer))})
        '''     </code>
        ''' </example>
        Public Shared Function ReArrangeDataColumns(ByVal source As DataTable, ByVal destinationColumnSet As DataColumn(),
                                                    ByVal ignoreConversionExceptionAndLogThemHere As ArrayList) As DataTable
            Return CompuMaster.Data.DataTablesTools.ReArrangeDataColumns(source, destinationColumnSet, ignoreConversionExceptionAndLogThemHere)
        End Function

        ''' <summary>
        ''' All columns of the table
        ''' </summary>
        ''' <param name="table">A table</param>
        ''' <returns></returns>
        Public Shared Function AllColumns(table As System.Data.DataTable) As System.Data.DataColumn()
            Dim Result As New System.Collections.Generic.List(Of System.Data.DataColumn)
            For MyCounter As Integer = 0 To table.Columns.Count - 1
                Result.Add(table.Columns(MyCounter))
            Next
            Return Result.ToArray
        End Function

        ''' <summary>
        ''' All column names of the table
        ''' </summary>
        ''' <param name="table">A table</param>
        ''' <returns></returns>
        Public Shared Function AllColumnNames(table As System.Data.DataTable) As String()
            Dim Result As New System.Collections.Generic.List(Of String)
            For MyCounter As Integer = 0 To table.Columns.Count - 1
                Result.Add(table.Columns(MyCounter).ColumnName)
            Next
            Return Result.ToArray
        End Function

        ''' <summary>
        '''     Execute a table join on two tables using their primary key columns (independent from their dataset, independent from their registered relations, without requirement for existing parent items (unlike to .NET standard behaviour) more like SQL behaviour)
        ''' </summary>
        ''' <param name="leftTable">The left table</param>
        ''' <param name="rightTable">The right table</param>
        ''' <param name="joinType">Inner, left, right or full join</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        Public Shared Function SqlJoinTables(ByVal leftTable As DataTable,
                                          ByVal rightTable As DataTable,
                                          ByVal joinType As SqlJoinTypes) As DataTable
            Return SqlJoinTables(leftTable, CType(Nothing, DataColumn()), CType(Nothing, DataColumn()), rightTable, CType(Nothing, DataColumn()), CType(Nothing, DataColumn()), joinType)
        End Function

        ''' <summary>
        '''     Execute a table join on two tables (independent from their dataset, independent from their registered relations, without requirement for existing parent items (unlike to .NET standard behaviour) more like SQL behaviour)
        ''' </summary>
        ''' <param name="leftTable">The left table</param>
        ''' <param name="leftTableKeys">An array of columns to be used as key columns for join (null/Nothing/empty array uses PrimaryKeys)</param>
        ''' <param name="rightTable">The right table</param>
        ''' <param name="rightTableKeys">An array of columns to be used as key columns for join (null/Nothing/empty array uses PrimaryKeys)</param>
        ''' <param name="joinType">Inner, left, right or full join</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        Public Shared Function SqlJoinTables(ByVal leftTable As DataTable, leftTableKeys As System.Data.DataColumn(),
                                          ByVal rightTable As DataTable, rightTableKeys As System.Data.DataColumn(),
                                          ByVal joinType As SqlJoinTypes) As DataTable
            Return SqlJoinTables(leftTable, leftTableKeys, Nothing, rightTable, rightTableKeys, Nothing, joinType)
        End Function

        ''' <summary>
        '''     Execute a table join on two tables (independent from their dataset, independent from their registered relations, without requirement for existing parent items (unlike to .NET standard behaviour) more like SQL behaviour)
        ''' </summary>
        ''' <param name="leftTable">The left table</param>
        ''' <param name="leftTableKey">A column to be used as key columns for join (null/Nothing uses PrimaryKeys)</param>
        ''' <param name="rightTable">The right table</param>
        ''' <param name="rightTableKey">A column to be used as key columns for join (null/Nothing uses PrimaryKeys)</param>
        ''' <param name="joinType">Inner, left, right or full join</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        Public Shared Function SqlJoinTables(ByVal leftTable As DataTable, leftTableKey As System.Data.DataColumn,
                                          ByVal rightTable As DataTable, rightTableKey As System.Data.DataColumn,
                                          ByVal joinType As SqlJoinTypes) As DataTable
            Dim LeftTableKeys As System.Data.DataColumn()
            If leftTableKey Is Nothing Then
                LeftTableKeys = Nothing
            Else
                LeftTableKeys = New System.Data.DataColumn() {leftTableKey}
            End If
            Dim RightTableKeys As System.Data.DataColumn()
            If rightTableKey Is Nothing Then
                RightTableKeys = Nothing
            Else
                RightTableKeys = New System.Data.DataColumn() {rightTableKey}
            End If
            Return SqlJoinTables(leftTable, LeftTableKeys, Nothing, rightTable, RightTableKeys, Nothing, joinType)
        End Function

        ''' <summary>
        '''     Execute a table join on two tables (independent from their dataset, independent from their registered relations, without requirement for existing parent items (unlike to .NET standard behaviour) more like SQL behaviour)
        ''' </summary>
        ''' <param name="leftTable">The left table</param>
        ''' <param name="leftTableKey">A column to be used as key columns for join (null/Nothing uses PrimaryKeys)</param>
        ''' <param name="leftTableColumnsToCopy">An array of columns to copy from the left table (null/Nothing uses all columns, empty array uses no columns)</param>
        ''' <param name="rightTable">The right table</param>
        ''' <param name="rightTableKey">A column to be used as key columns for join (null/Nothing uses PrimaryKeys)</param>
        ''' <param name="rightTableColumnsToCopy">An array of columns to copy from the right table (null/Nothing uses all columns, empty array uses no columns)</param>
        ''' <param name="joinType">Inner, left, right or full join</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        Public Shared Function SqlJoinTables(ByVal leftTable As DataTable, leftTableKey As System.Data.DataColumn, ByVal leftTableColumnsToCopy As System.Data.DataColumn(),
                                          ByVal rightTable As DataTable, rightTableKey As System.Data.DataColumn, ByVal rightTableColumnsToCopy As System.Data.DataColumn(),
                                          ByVal joinType As SqlJoinTypes) As DataTable
            Dim LeftTableKeys As System.Data.DataColumn()
            If leftTableKey Is Nothing Then
                LeftTableKeys = Nothing
            Else
                LeftTableKeys = New System.Data.DataColumn() {leftTableKey}
            End If
            Dim RightTableKeys As System.Data.DataColumn()
            If rightTableKey Is Nothing Then
                RightTableKeys = Nothing
            Else
                RightTableKeys = New System.Data.DataColumn() {rightTableKey}
            End If
            Return SqlJoinTables(leftTable, LeftTableKeys, leftTableColumnsToCopy, rightTable, RightTableKeys, rightTableColumnsToCopy, joinType)
        End Function

        ''' <summary>
        '''     Execute a table join on two tables (independent from their dataset, independent from their registered relations, without requirement for existing parent items (unlike to .NET standard behaviour) more like SQL behaviour)
        ''' </summary>
        ''' <param name="leftTable">The left table</param>
        ''' <param name="leftTableKeys">An array of columns to be used as key columns for join (null/Nothing/empty array uses PrimaryKeys)</param>
        ''' <param name="leftTableColumnsToCopy">An array of columns to copy from the left table (null/Nothing uses all columns, empty array uses no columns)</param>
        ''' <param name="rightTable">The right table</param>
        ''' <param name="rightTableKeys">An array of columns to be used as key columns for join (null/Nothing/empty array uses PrimaryKeys)</param>
        ''' <param name="rightTableColumnsToCopy">An array of columns to copy from the right table (null/Nothing uses all columns, empty array uses no columns)</param>
        ''' <param name="joinType">Inner, left, right or full join</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        Public Shared Function SqlJoinTables(ByVal leftTable As DataTable, leftTableKeys As System.Data.DataColumn(), ByVal leftTableColumnsToCopy As System.Data.DataColumn(),
                                          ByVal rightTable As DataTable, rightTableKeys As System.Data.DataColumn(), ByVal rightTableColumnsToCopy As System.Data.DataColumn(),
                                          ByVal joinType As SqlJoinTypes) As DataTable
            Return SqlJoinTables(leftTable, leftTableKeys, leftTableColumnsToCopy, rightTable, rightTableKeys, rightTableColumnsToCopy, joinType, False)
        End Function

        ''' <summary>
        '''     Execute a table join on two tables (independent from their dataset, independent from their registered relations, without requirement for existing parent items (unlike to .NET standard behaviour) more like SQL behaviour)
        ''' </summary>
        ''' <param name="leftTable">The left table</param>
        ''' <param name="leftTableKeys">An array of columns to be used as key columns for join (null/Nothing/empty array uses PrimaryKeys)</param>
        ''' <param name="leftTableColumnsToCopy">An array of columns to copy from the left table (null/Nothing uses all columns, empty array uses no columns)</param>
        ''' <param name="rightTable">The right table</param>
        ''' <param name="rightTableKeys">An array of columns to be used as key columns for join (null/Nothing/empty array uses PrimaryKeys)</param>
        ''' <param name="rightTableColumnsToCopy">An array of columns to copy from the right table (null/Nothing uses all columns, empty array uses no columns)</param>
        ''' <param name="joinType">Inner, left, right or full join</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        Public Shared Function SqlJoinTables(ByVal leftTable As DataTable, leftTableKeys As System.Data.DataColumn(), ByVal leftTableColumnsToCopy As System.Data.DataColumn(),
                                          ByVal rightTable As DataTable, rightTableKeys As System.Data.DataColumn(), ByVal rightTableColumnsToCopy As System.Data.DataColumn(),
                                          ByVal joinType As SqlJoinTypes, compareStringsCaseInsensitive As Boolean) As DataTable
            'Check required arguments
            If leftTable Is Nothing Then
                Throw New ArgumentNullException(NameOf(leftTable), "Left table is a required parameter")
            End If
            If rightTable Is Nothing Then
                Throw New ArgumentNullException(NameOf(rightTable), "Right table is a required parameter")
            End If

            'Auto-complete required arguments
            If leftTableColumnsToCopy Is Nothing Then
                leftTableColumnsToCopy = AllColumns(leftTable)
            End If
            If rightTableColumnsToCopy Is Nothing Then
                rightTableColumnsToCopy = AllColumns(rightTable)
            End If
            If (leftTableColumnsToCopy Is Nothing OrElse leftTableColumnsToCopy.Length = 0) AndAlso (rightTableColumnsToCopy Is Nothing OrElse rightTableColumnsToCopy.Length = 0) Then
                'Show all columns in case left and right side are without explicit definition
                leftTableColumnsToCopy = AllColumns(leftTable)
                rightTableColumnsToCopy = AllColumns(rightTable)
            End If
            If leftTableKeys Is Nothing OrElse leftTableKeys.Length = 0 Then
                leftTableKeys = leftTable.PrimaryKey
            End If
            If rightTableKeys Is Nothing OrElse rightTableKeys.Length = 0 Then
                rightTableKeys = rightTable.PrimaryKey
            End If

            'Execute the SQL-JOIN
            If joinType = SqlJoinTypes.Cross Then
                'Execute CrossJoin
                Dim indexesOfLeftTableColumnsToCopy As New ArrayList(), indexesOfRightTableColumnsToCopy As New ArrayList()
                For MyCounter As Integer = 0 To leftTableColumnsToCopy.Length - 1
                    indexesOfLeftTableColumnsToCopy.Add(leftTableColumnsToCopy(MyCounter).Ordinal)
                Next
                For MyCounter As Integer = 0 To rightTableColumnsToCopy.Length - 1
                    indexesOfRightTableColumnsToCopy.Add(rightTableColumnsToCopy(MyCounter).Ordinal)
                Next
                Return Data.DataTablesTools.CrossJoinTables(leftTable, CType(indexesOfLeftTableColumnsToCopy.ToArray(GetType(Integer)), Integer()), rightTable, CType(indexesOfRightTableColumnsToCopy.ToArray(GetType(Integer)), Integer()))
            ElseIf joinType = SqlJoinTypes.Right Then
                'Execute RightJoin

                'Pre-Check required arguments
                If leftTableKeys.Length <> rightTableKeys.Length Then Throw New ArgumentException("leftTableKeys and rightTableKeys must have got the same amount of keys")
                If leftTableKeys.Length = 0 Then Throw New ArgumentNullException("leftTableKeys must have got at least 1 key")
                For MyCounter As Integer = 0 To leftTableKeys.Length - 1
                    If leftTableKeys(MyCounter).Table IsNot leftTable Then Throw New ArgumentException("All leftTableKeys must be columns of leftTable")
                Next
                For MyCounter As Integer = 0 To rightTableKeys.Length - 1
                    If rightTableKeys(MyCounter).Table IsNot rightTable Then Throw New ArgumentException("All rightTableKeys must be columns of rightTable")
                Next
                For MyCounter As Integer = 0 To leftTableKeys.Length - 1
                    If Not SqlJoin_AreCompatibleComparisonColumns(leftTableKeys(MyCounter).DataType, rightTableKeys(MyCounter).DataType) Then Throw New ArgumentException("Columns [" & leftTableKeys(MyCounter).ColumnName & "] And [" & rightTableKeys(MyCounter).ColumnName & "] can't be compared: datatype mismatch")
                Next

                'Inverse RightJoin to LeftJoin
#Disable Warning S2234 ' Parameters should be passed in the correct order
                Return SqlJoinTables(rightTable, rightTableKeys, rightTableColumnsToCopy, leftTable, leftTableKeys, leftTableColumnsToCopy, SqlJoinTypes.Left, compareStringsCaseInsensitive)
#Enable Warning S2234 ' Parameters should be passed in the correct order
            Else
                'Execute Inner, Left or FullOuter Join

                'Pre-Check required arguments
                If leftTableKeys.Length <> rightTableKeys.Length Then Throw New ArgumentException("leftTableKeys and rightTableKeys must have got the same amount of keys")
                If leftTableKeys.Length = 0 Then Throw New ArgumentNullException("leftTableKeys must have got at least 1 key")
                For MyCounter As Integer = 0 To leftTableKeys.Length - 1
                    If leftTableKeys(MyCounter).Table IsNot leftTable Then Throw New ArgumentException("All leftTableKeys must be columns of leftTable")
                Next
                For MyCounter As Integer = 0 To rightTableKeys.Length - 1
                    If rightTableKeys(MyCounter).Table IsNot rightTable Then Throw New ArgumentException("All rightTableKeys must be columns of rightTable")
                Next
                For MyCounter As Integer = 0 To leftTableKeys.Length - 1
                    If Not SqlJoin_AreCompatibleComparisonColumns(leftTableKeys(MyCounter).DataType, rightTableKeys(MyCounter).DataType) Then Throw New ArgumentException("Columns [" & leftTableKeys(MyCounter).ColumnName & "] And [" & rightTableKeys(MyCounter).ColumnName & "] can't be compared: datatype mismatch")
                Next

                'Prepare column wrap table
                Dim LeftTableColumnWraps As Integer()
                Dim colWraps As New ArrayList
                For ColCounter As Integer = 0 To leftTableColumnsToCopy.Length - 1
                    colWraps.Add(leftTableColumnsToCopy(ColCounter).Ordinal)
                Next
                LeftTableColumnWraps = CType(colWraps.ToArray(GetType(Integer)), Integer())


                'Prepare the result table by copying the parent table
                Dim Result As DataTable = leftTable.Clone
                Result.TableName = "JoinedTable"
                Result.PrimaryKey = Nothing

                'Remove left table columns which are not required any more
                For MyCounter As Integer = Result.Columns.Count - 1 To 0 Step -1
                    Dim KeepThisColumn As Boolean = False
                    For MyColCounter As Integer = 0 To LeftTableColumnWraps.Length - 1
                        If LeftTableColumnWraps(MyColCounter) = MyCounter Then
                            KeepThisColumn = True
                            Exit For
                        End If
                    Next
                    'Remove unique constraints to allow duplicate values now in the joined result table
                    If Result.Columns(MyCounter).Unique = True Then
                        Result.Columns(MyCounter).Unique = False
                    End If
                    'Remove unnecessary columns
                    If KeepThisColumn = False Then
                        Result.Columns.Remove(Result.Columns(MyCounter))
                    End If
                Next

                'Add the right columns
                Dim RightTableColumnWraps As Integer()
                Dim colWrapsR As New ArrayList
                For ColCounter As Integer = 0 To rightTableColumnsToCopy.Length - 1
                    colWrapsR.Add(rightTableColumnsToCopy(ColCounter).Ordinal)
                Next
                RightTableColumnWraps = CType(colWrapsR.ToArray(GetType(Integer)), Integer())
                For MyCounter As Integer = 0 To RightTableColumnWraps.Length - 1
                    Dim MyColumn As DataColumn = rightTable.Columns(RightTableColumnWraps(MyCounter))
                    Dim UniqueColumnName As String = LookupUniqueColumnName(Result, MyColumn.ColumnName)
                    Dim ColumnCaption As String = MyColumn.Caption
                    Dim ColumnType As System.Type = MyColumn.DataType
                    Result.Columns.Add(UniqueColumnName, ColumnType).Caption = ColumnCaption
                Next

                Dim FoundRelatedRightRows As New ArrayList
                'Fill the rows now with the missing data
                For MyLeftTableRowCounter As Integer = 0 To leftTable.Rows.Count - 1
                    Dim MyLeftRow As DataRow = leftTable.Rows(MyLeftTableRowCounter)
                    Dim MyRightRows As DataRow() = SqlJoin_GetRightTableRows(MyLeftRow, rightTable, leftTableKeys, rightTableKeys, compareStringsCaseInsensitive)
                    If joinType = SqlJoinTypes.FullOuter Then
                        'only required for FullOuterJoin
                        For MyRightRowCounter As Integer = 0 To MyRightRows.Length - 1
                            Dim RowIndex As Integer = rightTable.Rows.IndexOf(MyRightRows(MyRightRowCounter))
                            If FoundRelatedRightRows.Contains(RowIndex) = False Then
                                FoundRelatedRightRows.Add(RowIndex)
                            End If
                        Next
                    End If

                    If MyRightRows.Length = 0 Then
                        'Data only on left side
                        Select Case joinType
                            Case SqlJoinTypes.Left, SqlJoinTypes.FullOuter
                                Dim NewRow As DataRow = Result.NewRow
                                'Copy only data from parent table
                                For MyColCounter As Integer = 0 To LeftTableColumnWraps.Length - 1
                                    NewRow(MyColCounter) = MyLeftRow(LeftTableColumnWraps(MyColCounter))
                                Next
                                'Add the new row, now
                                Result.Rows.Add(NewRow)
                            Case SqlJoinTypes.Inner
                                'don't add this row
                        End Select
                    Else
                        'Data found on both sides
                        For RowInserts As Integer = 0 To MyRightRows.Length - 1
                            Dim NewRow As DataRow = Result.NewRow
                            'Copy data from left table row
                            For MyColCounter As Integer = 0 To LeftTableColumnWraps.Length - 1
                                NewRow(MyColCounter) = MyLeftRow(LeftTableColumnWraps(MyColCounter))
                            Next
                            'Copy data from this right row
                            Dim MyAdditionalRightRow As DataRow = MyRightRows(RowInserts)
                            For MyColCounter As Integer = 0 To RightTableColumnWraps.Length - 1
                                NewRow(LeftTableColumnWraps.Length + MyColCounter) = MyAdditionalRightRow(RightTableColumnWraps(MyColCounter))
                            Next
                            'Add the new row, now
                            Result.Rows.Add(NewRow)
                        Next
                    End If

                Next

                'FullOuterJoin: Add rows from right table which haven't got a reference in left table
                If joinType = SqlJoinTypes.FullOuter Then
                    'only required for FullOuterJoin
                    For MyRightRowCounter As Integer = 0 To rightTable.Rows.Count - 1
                        If FoundRelatedRightRows.Contains(MyRightRowCounter) = False Then
                            Dim NewRow As DataRow = Result.NewRow
                            'Copy data from left table row
                            For MyColCounter As Integer = 0 To LeftTableColumnWraps.Length - 1
                                NewRow(MyColCounter) = DBNull.Value
                            Next
                            'Copy data from this right row
                            Dim MyAdditionalRightRow As DataRow = rightTable.Rows(MyRightRowCounter)
                            For MyColCounter As Integer = 0 To RightTableColumnWraps.Length - 1
                                NewRow(LeftTableColumnWraps.Length + MyColCounter) = MyAdditionalRightRow(RightTableColumnWraps(MyColCounter))
                            Next
                            'Add the new row, now
                            Result.Rows.Add(NewRow)
                        End If
                    Next
                End If

                Return Result
            End If
        End Function

        ''' <summary>
        ''' Find rows in a table with the specified values in its key columns
        ''' </summary>
        ''' <param name="searchedValue">A value which must be present in the key column of the table</param>
        ''' <param name="table">The table which is to be filtered</param>
        ''' <returns>All rows which match with the searched values</returns>
        Public Shared Function FindRowsInTable(searchedValue As Object, table As DataTable) As DataRow()
            If table.PrimaryKey.Length = 0 Then
                Throw New ArgumentException("The table doesn't contain a primary key definition")
            ElseIf table.PrimaryKey.Length <> 1 Then
                Throw New ArgumentException("A single searched value is specified, but the table contains a primary key collection with more than 1 key column")
            End If
            Return FindRowsInTable(New Object() {searchedValue}, table, table.PrimaryKey)
        End Function

        ''' <summary>
        ''' Find rows in a table with the specified values in its key columns
        ''' </summary>
        ''' <param name="searchedValueSet">A set of values which must be present in the key columns of the table</param>
        ''' <param name="table">The table which is to be filtered</param>
        ''' <returns>All rows which match with the searched values</returns>
        Public Shared Function FindRowsInTable(searchedValueSet As Object(), table As DataTable) As DataRow()
            Return FindRowsInTable(searchedValueSet, table, table.PrimaryKey)
        End Function

        ''' <summary>
        ''' Find rows in a table with the specified values in its key columns
        ''' </summary>
        ''' <param name="searchedValue">A value which must be present in the key column of the table</param>
        ''' <param name="table">The table which is to be filtered</param>
        ''' <param name="keyColumn">The key column of the table</param>
        ''' <returns>All rows which match with the searched values</returns>
        Public Shared Function FindRowsInTable(searchedValue As Object, table As DataTable, keyColumn As String) As DataRow()
            If keyColumn = Nothing Then
                Return FindRowsInTable(New Object() {searchedValue}, table, table.PrimaryKey)
            Else
                Return FindRowsInTable(New Object() {searchedValue}, table, New String() {keyColumn})
            End If
        End Function

        ''' <summary>
        ''' Find rows in a table with the specified values in its key columns
        ''' </summary>
        ''' <param name="searchedValueSet">A set of values which must be present in the key columns of the table</param>
        ''' <param name="table">The table which is to be filtered</param>
        ''' <param name="keyColumns">The key columns of the table</param>
        ''' <returns>All rows which match with the searched values</returns>
        Public Shared Function FindRowsInTable(searchedValueSet As Object(), table As DataTable, keyColumns As String()) As DataRow()
            If keyColumns Is Nothing OrElse keyColumns.Length = 0 Then
                Return FindRowsInTable(searchedValueSet, table, table.PrimaryKey)
            Else
                Dim MyKeyColumns As New System.Collections.Generic.List(Of DataColumn)
                For MyCounter As Integer = 0 To keyColumns.Length - 1
                    MyKeyColumns.Add(table.Columns(keyColumns(MyCounter)))
                Next
                Return FindRowsInTable(searchedValueSet, table, MyKeyColumns.ToArray)
            End If
        End Function

        ''' <summary>
        ''' Find rows in a table with the specified values in its key columns
        ''' </summary>
        ''' <param name="searchedValue">A value which must be present in the key column of the table</param>
        ''' <param name="table">The table which is to be filtered</param>
        ''' <param name="keyColumnIndex">The key column of the table</param>
        ''' <returns>All rows which match with the searched values</returns>
        Public Shared Function FindRowsInTable(searchedValue As Object, table As DataTable, keyColumnIndex As Integer) As DataRow()
            Return FindRowsInTable(New Object() {searchedValue}, table, table.Columns(keyColumnIndex))
        End Function

        ''' <summary>
        ''' Find rows in a table with the specified values in its key columns
        ''' </summary>
        ''' <param name="searchedValueSet">A set of values which must be present in the key columns of the table</param>
        ''' <param name="table">The table which is to be filtered</param>
        ''' <param name="keyColumnIndexes">The key columns of the table</param>
        ''' <returns>All rows which match with the searched values</returns>
        Public Shared Function FindRowsInTable(searchedValueSet As Object(), table As DataTable, keyColumnIndexes As Integer()) As DataRow()
            If keyColumnIndexes Is Nothing OrElse keyColumnIndexes.Length = 0 Then
                Return FindRowsInTable(searchedValueSet, table, table.PrimaryKey)
            Else
                Dim MyKeyColumns As New System.Collections.Generic.List(Of DataColumn)
                For MyCounter As Integer = 0 To keyColumnIndexes.Length - 1
                    MyKeyColumns.Add(table.Columns(keyColumnIndexes(MyCounter)))
                Next
                Return FindRowsInTable(searchedValueSet, table, MyKeyColumns.ToArray)
            End If
        End Function

        ''' <summary>
        ''' Find rows in a table with the specified values in its key columns
        ''' </summary>
        ''' <param name="searchedValue">A value which must be present in the key column of the table</param>
        ''' <param name="table">The table which is to be filtered</param>
        ''' <param name="keyColumn">The key column of the table</param>
        ''' <returns>All rows which match with the searched values</returns>
        Public Shared Function FindRowsInTable(searchedValue As Object, table As DataTable, keyColumn As DataColumn) As DataRow()
            Return FindRowsInTable(New Object() {searchedValue}, table, New DataColumn() {keyColumn})
        End Function

        ''' <summary>
        ''' Find rows in a table with the specified values in its key columns
        ''' </summary>
        ''' <param name="searchedValueSet">A set of values which must be present in the key columns of the table</param>
        ''' <param name="table">The table which is to be filtered</param>
        ''' <param name="keyColumns">The key columns of the table</param>
        ''' <returns>All rows which match with the searched values</returns>
        Public Shared Function FindRowsInTable(searchedValueSet As Object(), table As DataTable, keyColumns As DataColumn()) As DataRow()
            Return FindRowsInTable(searchedValueSet, table, keyColumns, False)
        End Function

        ''' <summary>
        ''' Find rows in a table with the specified values in its key columns
        ''' </summary>
        ''' <param name="searchedValueSet">A set of values which must be present in the key columns of the table</param>
        ''' <param name="table">The table which is to be filtered</param>
        ''' <param name="keyColumns">The key columns of the table</param>
        ''' <returns>All rows which match with the searched values</returns>
        Public Shared Function FindRowsInTable(searchedValueSet As Object(), table As DataTable, keyColumns As DataColumn(), compareStringsCaseInsensitive As Boolean) As DataRow()
            Dim MyKeyColumns As DataColumn() = keyColumns
            If MyKeyColumns Is Nothing OrElse MyKeyColumns.Length = 0 Then
                MyKeyColumns = table.PrimaryKey
            End If
            If MyKeyColumns Is Nothing OrElse MyKeyColumns.Length = 0 Then
                Throw New ArgumentException("Key columns haven't been specified and table doesn't contain a primary key defintion")
            ElseIf searchedValueSet Is Nothing Then
                Throw New ArgumentNullException(NameOf(searchedValueSet), "Required argument: searchedValueSet")
            ElseIf searchedValueSet.Length <> MyKeyColumns.Length Then
                Throw New ArgumentException("Array lengths must be equal: searchedValueSet and keyColumns")
            End If
            Dim Result As New System.Collections.Generic.List(Of DataRow)
            For MyRowCounter As Integer = 0 To table.Rows.Count - 1
                Dim IsMatch As Boolean = True
                For MyKeyCounter As Integer = 0 To MyKeyColumns.Length - 1
                    If SqlJoin_IsEqual(searchedValueSet(MyKeyCounter), table.Rows(MyRowCounter)(MyKeyColumns(MyKeyCounter)), compareStringsCaseInsensitive) = False Then
                        IsMatch = False
                        Exit For
                    End If
                Next
                If IsMatch Then
                    Result.Add(table.Rows(MyRowCounter))
                End If
            Next
            Return Result.ToArray
        End Function

        ''' <summary>
        ''' Find matching rows in a foreign table with the values in specified columns of a source table row
        ''' </summary>
        ''' <param name="sourceRow"></param>
        ''' <param name="foreignTable"></param>
        ''' <returns></returns>
        Public Shared Function FindMatchingRowsInForeignTable(sourceRow As DataRow, foreignTable As DataTable) As DataRow()
            Return FindMatchingRowsInForeignTable(sourceRow, foreignTable, sourceRow.Table.PrimaryKey, foreignTable.PrimaryKey)
        End Function

        ''' <summary>
        ''' Find matching rows in a foreign table with the values in specified columns of a source table row
        ''' </summary>
        ''' <param name="sourceRow"></param>
        ''' <param name="foreignTable"></param>
        ''' <param name="sourceRowKeyColumn"></param>
        ''' <param name="foreignTableKeyColumn"></param>
        ''' <returns></returns>
        Public Shared Function FindMatchingRowsInForeignTable(sourceRow As DataRow, foreignTable As DataTable, sourceRowKeyColumn As String, foreignTableKeyColumn As String) As DataRow()
            If sourceRowKeyColumn = Nothing Then Throw New ArgumentNullException("leftKeyColumn")
            If foreignTableKeyColumn = Nothing Then Throw New ArgumentNullException("rightKeyColumn")
            Return FindMatchingRowsInForeignTable(sourceRow, foreignTable, New String() {sourceRowKeyColumn}, New String() {foreignTableKeyColumn}, False)
        End Function

        ''' <summary>
        ''' Find matching rows in a foreign table with the values in specified columns of a source table row
        ''' </summary>
        ''' <param name="sourceRow"></param>
        ''' <param name="foreignTable"></param>
        ''' <param name="sourceRowKeyColumn"></param>
        ''' <param name="foreignTableKeyColumn"></param>
        ''' <returns></returns>
        Public Shared Function FindMatchingRowsInForeignTable(sourceRow As DataRow, foreignTable As DataTable, sourceRowKeyColumn As String, foreignTableKeyColumn As String, compareStringsCaseInsensitive As Boolean) As DataRow()
            If sourceRowKeyColumn = Nothing Then Throw New ArgumentNullException("leftKeyColumn")
            If foreignTableKeyColumn = Nothing Then Throw New ArgumentNullException("rightKeyColumn")
            Return FindMatchingRowsInForeignTable(sourceRow, foreignTable, New String() {sourceRowKeyColumn}, New String() {foreignTableKeyColumn}, compareStringsCaseInsensitive)
        End Function

        ''' <summary>
        ''' Find matching rows in a foreign table with the values in specified columns of a source table row
        ''' </summary>
        ''' <param name="sourceRow"></param>
        ''' <param name="foreignTable"></param>
        ''' <param name="sourceRowKeyColumns"></param>
        ''' <param name="foreignTableKeyColumns"></param>
        ''' <returns></returns>
        Public Shared Function FindMatchingRowsInForeignTable(sourceRow As DataRow, foreignTable As DataTable, sourceRowKeyColumns As String(), foreignTableKeyColumns As String()) As DataRow()
            Return FindMatchingRowsInForeignTable(sourceRow, foreignTable, sourceRowKeyColumns, foreignTableKeyColumns, False)
        End Function

        ''' <summary>
        ''' Find matching rows in a foreign table with the values in specified columns of a source table row
        ''' </summary>
        ''' <param name="sourceRow"></param>
        ''' <param name="foreignTable"></param>
        ''' <param name="sourceRowKeyColumns"></param>
        ''' <param name="foreignTableKeyColumns"></param>
        ''' <returns></returns>
        Public Shared Function FindMatchingRowsInForeignTable(sourceRow As DataRow, foreignTable As DataTable, sourceRowKeyColumns As String(), foreignTableKeyColumns As String(), compareStringsCaseInsensitive As Boolean) As DataRow()
            Dim MyLeftKeys As DataColumn()
            If sourceRowKeyColumns Is Nothing OrElse sourceRowKeyColumns.Length = 0 Then
                MyLeftKeys = sourceRow.Table.PrimaryKey
            Else
                Dim MyLeftKeyColumns As New System.Collections.Generic.List(Of DataColumn)
                For MyCounter As Integer = 0 To sourceRowKeyColumns.Length - 1
                    MyLeftKeyColumns.Add(sourceRow.Table.Columns(sourceRowKeyColumns(MyCounter)))
                Next
                MyLeftKeys = MyLeftKeyColumns.ToArray
            End If
            Dim MyRightKeys As DataColumn()
            If foreignTableKeyColumns Is Nothing OrElse foreignTableKeyColumns.Length = 0 Then
                MyRightKeys = foreignTable.PrimaryKey
            Else
                Dim MyrightKeyColumns As New System.Collections.Generic.List(Of DataColumn)
                For MyCounter As Integer = 0 To foreignTableKeyColumns.Length - 1
                    MyrightKeyColumns.Add(foreignTable.Columns(foreignTableKeyColumns(MyCounter)))
                Next
                MyRightKeys = MyrightKeyColumns.ToArray
            End If
            Return FindMatchingRowsInForeignTable(sourceRow, foreignTable, MyLeftKeys, MyRightKeys, compareStringsCaseInsensitive)
        End Function

        ''' <summary>
        ''' Find matching rows in a foreign table with the values in specified columns of a source table row
        ''' </summary>
        ''' <param name="sourceRow"></param>
        ''' <param name="foreignTable"></param>
        ''' <param name="sourceRowKeyColumnIndex"></param>
        ''' <param name="foreignTableKeyColumnIndex"></param>
        ''' <returns></returns>
        Public Shared Function FindMatchingRowsInForeignTable(sourceRow As DataRow, foreignTable As DataTable, sourceRowKeyColumnIndex As Integer, foreignTableKeyColumnIndex As Integer) As DataRow()
            Return FindMatchingRowsInForeignTable(sourceRow, foreignTable, New Integer() {sourceRowKeyColumnIndex}, New Integer() {foreignTableKeyColumnIndex}, False)
        End Function

        ''' <summary>
        ''' Find matching rows in a foreign table with the values in specified columns of a source table row
        ''' </summary>
        ''' <param name="sourceRow"></param>
        ''' <param name="foreignTable"></param>
        ''' <param name="sourceRowKeyColumnIndex"></param>
        ''' <param name="foreignTableKeyColumnIndex"></param>
        ''' <returns></returns>
        Public Shared Function FindMatchingRowsInForeignTable(sourceRow As DataRow, foreignTable As DataTable, sourceRowKeyColumnIndex As Integer, foreignTableKeyColumnIndex As Integer, compareStringsCaseInsensitive As Boolean) As DataRow()
            Return FindMatchingRowsInForeignTable(sourceRow, foreignTable, New Integer() {sourceRowKeyColumnIndex}, New Integer() {foreignTableKeyColumnIndex}, compareStringsCaseInsensitive)
        End Function

        ''' <summary>
        ''' Find matching rows in a foreign table with the values in specified columns of a source table row
        ''' </summary>
        ''' <param name="sourceRow"></param>
        ''' <param name="foreignTable"></param>
        ''' <param name="sourceRowKeyColumnIndexes"></param>
        ''' <param name="foreignTableKeyColumnIndexes"></param>
        ''' <returns></returns>
        Public Shared Function FindMatchingRowsInForeignTable(sourceRow As DataRow, foreignTable As DataTable, sourceRowKeyColumnIndexes As Integer(), foreignTableKeyColumnIndexes As Integer()) As DataRow()
            Return FindMatchingRowsInForeignTable(sourceRow, foreignTable, sourceRowKeyColumnIndexes, foreignTableKeyColumnIndexes, False)
        End Function
        ''' <summary>
        ''' Find matching rows in a foreign table with the values in specified columns of a source table row
        ''' </summary>
        ''' <param name="sourceRow"></param>
        ''' <param name="foreignTable"></param>
        ''' <param name="sourceRowKeyColumnIndexes"></param>
        ''' <param name="foreignTableKeyColumnIndexes"></param>
        ''' <returns></returns>
        Public Shared Function FindMatchingRowsInForeignTable(sourceRow As DataRow, foreignTable As DataTable, sourceRowKeyColumnIndexes As Integer(), foreignTableKeyColumnIndexes As Integer(), compareStringsCaseInsensitive As Boolean) As DataRow()
            Dim MyLeftKeys As DataColumn()
            If sourceRowKeyColumnIndexes Is Nothing OrElse sourceRowKeyColumnIndexes.Length = 0 Then
                MyLeftKeys = sourceRow.Table.PrimaryKey
            Else
                Dim MyLeftKeyColumns As New System.Collections.Generic.List(Of DataColumn)
                For MyCounter As Integer = 0 To sourceRowKeyColumnIndexes.Length - 1
                    MyLeftKeyColumns.Add(sourceRow.Table.Columns(sourceRowKeyColumnIndexes(MyCounter)))
                Next
                MyLeftKeys = MyLeftKeyColumns.ToArray
            End If
            Dim MyRightKeys As DataColumn()
            If foreignTableKeyColumnIndexes Is Nothing OrElse foreignTableKeyColumnIndexes.Length = 0 Then
                MyRightKeys = foreignTable.PrimaryKey
            Else
                Dim MyrightKeyColumns As New System.Collections.Generic.List(Of DataColumn)
                For MyCounter As Integer = 0 To foreignTableKeyColumnIndexes.Length - 1
                    MyrightKeyColumns.Add(foreignTable.Columns(foreignTableKeyColumnIndexes(MyCounter)))
                Next
                MyRightKeys = MyrightKeyColumns.ToArray
            End If
            Return FindMatchingRowsInForeignTable(sourceRow, foreignTable, MyLeftKeys, MyRightKeys, compareStringsCaseInsensitive)
        End Function

        ''' <summary>
        ''' Find matching rows in a foreign table with the values in specified columns of a source table row
        ''' </summary>
        ''' <param name="sourceRow"></param>
        ''' <param name="foreignTable"></param>
        ''' <param name="sourceRowKeyColumn"></param>
        ''' <param name="foreignTableKeyColumn"></param>
        ''' <returns></returns>
        Public Shared Function FindMatchingRowsInForeignTable(sourceRow As DataRow, foreignTable As DataTable, sourceRowKeyColumn As DataColumn, foreignTableKeyColumn As DataColumn) As DataRow()
            If sourceRowKeyColumn Is Nothing Then Throw New ArgumentNullException("leftKeyColumn")
            If foreignTableKeyColumn Is Nothing Then Throw New ArgumentNullException("rightKeyColumn")
            Return FindMatchingRowsInForeignTable(sourceRow, foreignTable, New DataColumn() {sourceRowKeyColumn}, New DataColumn() {foreignTableKeyColumn}, False)
        End Function

        ''' <summary>
        ''' Find matching rows in a foreign table with the values in specified columns of a source table row
        ''' </summary>
        ''' <param name="sourceRow"></param>
        ''' <param name="foreignTable"></param>
        ''' <param name="sourceRowKeyColumn"></param>
        ''' <param name="foreignTableKeyColumn"></param>
        ''' <returns></returns>
        Public Shared Function FindMatchingRowsInForeignTable(sourceRow As DataRow, foreignTable As DataTable, sourceRowKeyColumn As DataColumn, foreignTableKeyColumn As DataColumn, compareStringsCaseInsensitive As Boolean) As DataRow()
            If sourceRowKeyColumn Is Nothing Then Throw New ArgumentNullException("leftKeyColumn")
            If foreignTableKeyColumn Is Nothing Then Throw New ArgumentNullException("rightKeyColumn")
            Return FindMatchingRowsInForeignTable(sourceRow, foreignTable, New DataColumn() {sourceRowKeyColumn}, New DataColumn() {foreignTableKeyColumn}, compareStringsCaseInsensitive)
        End Function

        ''' <summary>
        ''' Find matching rows in a foreign table with the values in specified columns of a source table row
        ''' </summary>
        ''' <param name="sourceRow"></param>
        ''' <param name="foreignTable"></param>
        ''' <param name="sourceRowKeyColumns"></param>
        ''' <param name="foreignTableKeyColumns"></param>
        ''' <returns></returns>
        Public Shared Function FindMatchingRowsInForeignTable(sourceRow As DataRow, foreignTable As DataTable, sourceRowKeyColumns As DataColumn(), foreignTableKeyColumns As DataColumn()) As DataRow()
            Return SqlJoin_GetRightTableRows(sourceRow, foreignTable, sourceRowKeyColumns, foreignTableKeyColumns, False)
        End Function

        ''' <summary>
        ''' Find matching rows in a foreign table with the values in specified columns of a source table row
        ''' </summary>
        ''' <param name="sourceRow"></param>
        ''' <param name="foreignTable"></param>
        ''' <param name="sourceRowKeyColumns"></param>
        ''' <param name="foreignTableKeyColumns"></param>
        ''' <param name="compareStringsCaseInsensitive"></param>
        ''' <returns></returns>
        Public Shared Function FindMatchingRowsInForeignTable(sourceRow As DataRow, foreignTable As DataTable, sourceRowKeyColumns As DataColumn(), foreignTableKeyColumns As DataColumn(), compareStringsCaseInsensitive As Boolean) As DataRow()
            Return SqlJoin_GetRightTableRows(sourceRow, foreignTable, sourceRowKeyColumns, foreignTableKeyColumns, compareStringsCaseInsensitive)
        End Function

        Private Shared Function SqlJoin_GetRightTableRows(leftRow As DataRow, rightTable As DataTable, leftKeys As DataColumn(), rightKeys As DataColumn(), compareStringsCaseInsensitive As Boolean) As DataRow()
            Dim Result As New System.Collections.Generic.List(Of DataRow)
            For MyRowCounter As Integer = 0 To rightTable.Rows.Count - 1
                Dim IsMatch As Boolean = True
                For MyKeyCounter As Integer = 0 To leftKeys.Length - 1
                    If SqlJoin_IsEqual(leftRow(leftKeys(MyKeyCounter)), rightTable.Rows(MyRowCounter)(rightKeys(MyKeyCounter)), compareStringsCaseInsensitive) = False Then
                        IsMatch = False
                        Exit For
                    End If
                Next
                If IsMatch Then
                    Result.Add(rightTable.Rows(MyRowCounter))
                End If
            Next
            Return Result.ToArray
        End Function

        Private Shared Function SqlJoin_IsEqual(value1 As Object, value2 As Object, compareStringsCaseInsensitive As Boolean) As Boolean
            If IsDBNull(value1) Xor IsDBNull(value2) Then
                Return False
            ElseIf IsDBNull(value1) AndAlso IsDBNull(value2) Then
                Return True
            ElseIf value1.GetType Is GetType(String) AndAlso value2.GetType Is GetType(String) AndAlso compareStringsCaseInsensitive Then
                Return LCase(CType(value1, String)) = LCase(CType(value2, String))
            ElseIf value1.GetType Is GetType(String) AndAlso value2.GetType Is GetType(String) Then
                Return CType(value1, String) = CType(value2, String)
            ElseIf value1.GetType Is GetType(System.Decimal) OrElse value2.GetType Is GetType(System.Decimal) Then
                Return CType(value1, System.Decimal) = CType(value2, System.Decimal)
            ElseIf value1.GetType Is GetType(System.Double) OrElse value2.GetType Is GetType(System.Double) OrElse value1.GetType Is GetType(System.Single) OrElse value2.GetType Is GetType(System.Single) Then
                Return CType(value1, System.Double) = CType(value2, System.Double)
            ElseIf IsNumeric(value1) AndAlso IsNumeric(value2) Then
                Return CType(value1, System.Int64) = CType(value2, System.Int64)
            ElseIf IsDate(value1) AndAlso IsDate(value2) Then
                Return CType(value1, DateTime) = CType(value2, DateTime)
            Else
                Return Object.Equals(value1, value2)
            End If
        End Function

        Private Shared Function SqlJoin_AreCompatibleComparisonColumns(leftColumnDataType As System.Type, rightColumnDataType As System.Type) As Boolean
            If leftColumnDataType Is GetType(String) AndAlso rightColumnDataType Is GetType(String) Then
                Return True
            ElseIf leftColumnDataType Is GetType(String) Xor rightColumnDataType Is GetType(String) Then
                Return False
            ElseIf leftColumnDataType.IsValueType AndAlso rightColumnDataType.IsValueType Then
                Return True
            ElseIf leftColumnDataType.IsValueType Xor rightColumnDataType.IsValueType Then
                Return False
            Else
                Return True
            End If
        End Function

        ''' <summary>
        '''     Execute a table join on two tables (independent from their dataset, independent from their registered relations, without requirement for existing parent items (unlike to .NET standard behaviour) more like SQL behaviour)
        ''' </summary>
        ''' <param name="leftTable">The left table</param>
        ''' <param name="indexesOfLeftTableKeys">An array of column indexes to be used as key columns for join (null/Nothing/empty array uses PrimaryKeys)</param>
        ''' <param name="rightTable">The right table</param>
        ''' <param name="indexesOfRightTableKeys">An array of column indexes to be used as key columns for join (null/Nothing/empty array uses PrimaryKeys)</param>
        ''' <param name="joinType">Inner, left, right or full join</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        Public Shared Function SqlJoinTables(ByVal leftTable As DataTable, indexesOfLeftTableKeys As Integer(),
                                          ByVal rightTable As DataTable, indexesOfRightTableKeys As Integer(),
                                          ByVal joinType As SqlJoinTypes) As DataTable
            Return SqlJoinTables(leftTable, indexesOfLeftTableKeys, Nothing, rightTable, indexesOfRightTableKeys, Nothing, joinType)
        End Function

        ''' <summary>
        '''     Execute a table join on two tables (independent from their dataset, independent from their registered relations, without requirement for existing parent items (unlike to .NET standard behaviour) more like SQL behaviour)
        ''' </summary>
        ''' <param name="leftTable">The left table</param>
        ''' <param name="indexesOfLeftTableKey">A column index to be used as key columns for join</param>
        ''' <param name="rightTable">The right table</param>
        ''' <param name="indexesOfRightTableKey">A column index to be used as key columns for join</param>
        ''' <param name="joinType">Inner, left, right or full join</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        Public Shared Function SqlJoinTables(ByVal leftTable As DataTable, indexesOfLeftTableKey As Integer,
                                          ByVal rightTable As DataTable, indexesOfRightTableKey As Integer,
                                          ByVal joinType As SqlJoinTypes) As DataTable
            Return SqlJoinTables(leftTable, New Integer() {indexesOfLeftTableKey}, Nothing, rightTable, New Integer() {indexesOfRightTableKey}, Nothing, joinType)
        End Function

        ''' <summary>
        '''     Execute a table join on two tables (independent from their dataset, independent from their registered relations, without requirement for existing parent items (unlike to .NET standard behaviour) more like SQL behaviour)
        ''' </summary>
        ''' <param name="leftTable">The left table</param>
        ''' <param name="indexesOfLeftTableKey">A column index to be used as key columns for join</param>
        ''' <param name="indexesOfLeftTableColumnsToCopy">An array of column indexes to copy from the left table (null/Nothing uses all columns, empty array uses no columns)</param>
        ''' <param name="rightTable">The right table</param>
        ''' <param name="indexesOfRightTableKey">A column index to be used as key columns for join</param>
        ''' <param name="indexesOfRightTableColumnsToCopy">An array of column indexes to copy from the right table (null/Nothing uses all columns, empty array uses no columns)</param>
        ''' <param name="joinType">Inner, left, right or full join</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        Public Shared Function SqlJoinTables(ByVal leftTable As DataTable, indexesOfLeftTableKey As Integer, ByVal indexesOfLeftTableColumnsToCopy As Integer(),
                                          ByVal rightTable As DataTable, indexesOfRightTableKey As Integer, ByVal indexesOfRightTableColumnsToCopy As Integer(),
                                          ByVal joinType As SqlJoinTypes) As DataTable
            Return SqlJoinTables(leftTable, New Integer() {indexesOfLeftTableKey}, indexesOfLeftTableColumnsToCopy, rightTable, New Integer() {indexesOfRightTableKey}, indexesOfRightTableColumnsToCopy, joinType)
        End Function

        ''' <summary>
        '''     Execute a table join on two tables (independent from their dataset, independent from their registered relations, without requirement for existing parent items (unlike to .NET standard behaviour) more like SQL behaviour)
        ''' </summary>
        ''' <param name="leftTable">The left table</param>
        ''' <param name="indexesOfLeftTableKeys">An array of column indexes to be used as key columns for join (null/Nothing/empty array uses PrimaryKeys)</param>
        ''' <param name="indexesOfLeftTableColumnsToCopy">An array of column indexes to copy from the left table (null/Nothing uses all columns, empty array uses no columns)</param>
        ''' <param name="rightTable">The right table</param>
        ''' <param name="indexesOfRightTableKeys">An array of column indexes to be used as key columns for join (null/Nothing/empty array uses PrimaryKeys)</param>
        ''' <param name="indexesOfRightTableColumnsToCopy">An array of column indexes to copy from the right table (null/Nothing uses all columns, empty array uses no columns)</param>
        ''' <param name="joinType">Inner, left, right or full join</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        Public Shared Function SqlJoinTables(ByVal leftTable As DataTable, indexesOfLeftTableKeys As Integer(), ByVal indexesOfLeftTableColumnsToCopy As Integer(),
                                          ByVal rightTable As DataTable, indexesOfRightTableKeys As Integer(), ByVal indexesOfRightTableColumnsToCopy As Integer(),
                                          ByVal joinType As SqlJoinTypes) As DataTable

            If leftTable Is Nothing Then
                Throw New ArgumentNullException(NameOf(leftTable), "Left table is a required parameter")
            End If
            If rightTable Is Nothing Then
                Throw New ArgumentNullException(NameOf(rightTable), "Right table is a required parameter")
            End If

            Dim leftKeys As New ArrayList, rightKeys As New ArrayList, leftColumns As New ArrayList, rightColumns As New ArrayList
            Dim newLeftKeys As DataColumn() = Nothing, newRightKeys As DataColumn() = Nothing, newLeftColumns As DataColumn() = Nothing, newRightColumns As DataColumn() = Nothing
            If indexesOfLeftTableKeys IsNot Nothing Then
                For MyCounter As Integer = 0 To indexesOfLeftTableKeys.Length - 1
                    leftKeys.Add(leftTable.Columns(indexesOfLeftTableKeys(MyCounter)))
                Next
                newLeftKeys = CType(leftKeys.ToArray(GetType(System.Data.DataColumn)), System.Data.DataColumn())
            End If
            If indexesOfLeftTableColumnsToCopy IsNot Nothing Then
                For MyCounter As Integer = 0 To indexesOfLeftTableColumnsToCopy.Length - 1
                    leftColumns.Add(leftTable.Columns(indexesOfLeftTableColumnsToCopy(MyCounter)))
                Next
                newLeftColumns = CType(leftColumns.ToArray(GetType(System.Data.DataColumn)), System.Data.DataColumn())
            End If
            If indexesOfRightTableKeys IsNot Nothing Then
                For MyCounter As Integer = 0 To indexesOfRightTableKeys.Length - 1
                    rightKeys.Add(rightTable.Columns(indexesOfRightTableKeys(MyCounter)))
                Next
                newRightKeys = CType(rightKeys.ToArray(GetType(System.Data.DataColumn)), System.Data.DataColumn())
            End If
            If indexesOfRightTableColumnsToCopy IsNot Nothing Then
                For MyCounter As Integer = 0 To indexesOfRightTableColumnsToCopy.Length - 1
                    rightColumns.Add(rightTable.Columns(indexesOfRightTableColumnsToCopy(MyCounter)))
                Next
                newRightColumns = CType(rightColumns.ToArray(GetType(System.Data.DataColumn)), System.Data.DataColumn())
            End If

            Return SqlJoinTables(leftTable, newLeftKeys, newLeftColumns,
                              rightTable, newRightKeys, newRightColumns,
                              joinType)

        End Function

        ''' <summary>
        '''     Execute a table join on two tables (independent from their dataset, independent from their registered relations, without requirement for existing parent items (unlike to .NET standard behaviour) more like SQL behaviour)
        ''' </summary>
        ''' <param name="leftTable">The left table</param>
        ''' <param name="leftTableKeys">An array of columns to be used as key columns for join (null/Nothing/empty array uses PrimaryKeys)</param>
        ''' <param name="rightTable">The right table</param>
        ''' <param name="rightTableKeys">An array of columns to be used as key columns for join (null/Nothing/empty array uses PrimaryKeys)</param>
        ''' <param name="joinType">Inner, left, right or full join</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        Public Shared Function SqlJoinTables(ByVal leftTable As DataTable, leftTableKeys As String(),
                                          ByVal rightTable As DataTable, rightTableKeys As String(),
                                          ByVal joinType As SqlJoinTypes) As DataTable
            Return SqlJoinTables(leftTable, leftTableKeys, Nothing, rightTable, rightTableKeys, Nothing, joinType)
        End Function

        ''' <summary>
        '''     Execute a table join on two tables (independent from their dataset, independent from their registered relations, without requirement for existing parent items (unlike to .NET standard behaviour) more like SQL behaviour)
        ''' </summary>
        ''' <param name="leftTable">The left table</param>
        ''' <param name="leftTableKey">A column to be used as key column for join</param>
        ''' <param name="rightTable">The right table</param>
        ''' <param name="rightTableKey">A column to be used as key column for join</param>
        ''' <param name="joinType">Inner, left, right or full join</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        Public Shared Function SqlJoinTables(ByVal leftTable As DataTable, leftTableKey As String,
                                          ByVal rightTable As DataTable, rightTableKey As String,
                                          ByVal joinType As SqlJoinTypes) As DataTable
            If leftTableKey = Nothing Then Throw New ArgumentNullException(NameOf(leftTableKey))
            If rightTableKey = Nothing Then Throw New ArgumentNullException(NameOf(rightTableKey))
            Return SqlJoinTables(leftTable, New String() {leftTableKey}, Nothing, rightTable, New String() {rightTableKey}, Nothing, joinType)
        End Function

        ''' <summary>
        '''     Execute a table join on two tables (independent from their dataset, independent from their registered relations, without requirement for existing parent items (unlike to .NET standard behaviour) more like SQL behaviour)
        ''' </summary>
        ''' <param name="leftTable">The left table</param>
        ''' <param name="leftTableKey">A column to be used as key column for join</param>
        ''' <param name="leftTableColumnsToCopy">An array of columns to copy from the left table (null/Nothing uses all columns, empty array uses no columns)</param>
        ''' <param name="rightTable">The right table</param>
        ''' <param name="rightTableKey">A column to be used as key column for join</param>
        ''' <param name="rightTableColumnsToCopy">An array of columns to copy from the right table (null/Nothing uses all columns, empty array uses no columns)</param>
        ''' <param name="joinType">Inner, left, right or full join</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        Public Shared Function SqlJoinTables(ByVal leftTable As DataTable, leftTableKey As String, ByVal leftTableColumnsToCopy As String(),
                                          ByVal rightTable As DataTable, rightTableKey As String, ByVal rightTableColumnsToCopy As String(),
                                          ByVal joinType As SqlJoinTypes) As DataTable
            If leftTableKey = Nothing Then Throw New ArgumentNullException(NameOf(leftTableKey))
            If rightTableKey = Nothing Then Throw New ArgumentNullException(NameOf(rightTableKey))
            Return SqlJoinTables(leftTable, New String() {leftTableKey}, leftTableColumnsToCopy, rightTable, New String() {rightTableKey}, rightTableColumnsToCopy, joinType)
        End Function

        ''' <summary>
        '''     Execute a table join on two tables (independent from their dataset, independent from their registered relations, without requirement for existing parent items (unlike to .NET standard behaviour) more like SQL behaviour)
        ''' </summary>
        ''' <param name="leftTable">The left table</param>
        ''' <param name="leftTableKeys">An array of columns to be used as key columns for join (null/Nothing/empty array uses PrimaryKeys)</param>
        ''' <param name="leftTableColumnsToCopy">An array of columns to copy from the left table (null/Nothing uses all columns, empty array uses no columns)</param>
        ''' <param name="rightTable">The right table</param>
        ''' <param name="rightTableKeys">An array of columns to be used as key columns for join (null/Nothing/empty array uses PrimaryKeys)</param>
        ''' <param name="rightTableColumnsToCopy">An array of columns to copy from the right table (null/Nothing uses all columns, empty array uses no columns)</param>
        ''' <param name="joinType">Inner, left, right or full join</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        Public Shared Function SqlJoinTables(ByVal leftTable As DataTable, leftTableKeys As String(), ByVal leftTableColumnsToCopy As String(),
                                          ByVal rightTable As DataTable, rightTableKeys As String(), ByVal rightTableColumnsToCopy As String(),
                                          ByVal joinType As SqlJoinTypes) As DataTable

            If leftTable Is Nothing Then
                Throw New ArgumentNullException(NameOf(leftTable), "Left table is a required parameter")
            End If
            If rightTable Is Nothing Then
                Throw New ArgumentNullException(NameOf(rightTable), "Right table is a required parameter")
            End If

            Dim leftKeys As New ArrayList, rightKeys As New ArrayList, leftColumns As New ArrayList, rightColumns As New ArrayList
            Dim newLeftKeys As DataColumn() = Nothing, newRightKeys As DataColumn() = Nothing, newLeftColumns As DataColumn() = Nothing, newRightColumns As DataColumn() = Nothing
            If leftTableKeys IsNot Nothing Then
                For MyCounter As Integer = 0 To leftTableKeys.Length - 1
                    leftKeys.Add(leftTable.Columns(leftTableKeys(MyCounter)))
                Next
                newLeftKeys = CType(leftKeys.ToArray(GetType(System.Data.DataColumn)), System.Data.DataColumn())
            End If
            If leftTableColumnsToCopy IsNot Nothing Then
                For MyCounter As Integer = 0 To leftTableColumnsToCopy.Length - 1
                    leftColumns.Add(leftTable.Columns(leftTableColumnsToCopy(MyCounter)))
                Next
                newLeftColumns = CType(leftColumns.ToArray(GetType(System.Data.DataColumn)), System.Data.DataColumn())
            End If
            If rightTableKeys IsNot Nothing Then
                For MyCounter As Integer = 0 To rightTableKeys.Length - 1
                    rightKeys.Add(rightTable.Columns(rightTableKeys(MyCounter)))
                Next
                newRightKeys = CType(rightKeys.ToArray(GetType(System.Data.DataColumn)), System.Data.DataColumn())
            End If
            If rightTableColumnsToCopy IsNot Nothing Then
                For MyCounter As Integer = 0 To rightTableColumnsToCopy.Length - 1
                    rightColumns.Add(rightTable.Columns(rightTableColumnsToCopy(MyCounter)))
                Next
                newRightColumns = CType(rightColumns.ToArray(GetType(System.Data.DataColumn)), System.Data.DataColumn())
            End If

            Return SqlJoinTables(leftTable, newLeftKeys, newLeftColumns,
                              rightTable, newRightKeys, newRightColumns,
                              joinType)

        End Function

        ''' <summary>
        ''' Check that all required columns are available in specified table
        ''' </summary>
        ''' <param name="table">A table which must contain the columns</param>
        ''' <param name="requiredColumnNames">Column names that must exist in table</param>
        ''' <returns></returns>
        Public Shared Function ValidateRequiredColumnNames(table As DataTable, requiredColumnNames As String()) As String()
            Return ValidateRequiredColumnNames(table, requiredColumnNames, False)
        End Function

        ''' <summary>
        ''' Check that all required columns are available in specified table
        ''' </summary>
        ''' <param name="table">A table which must contain the columns</param>
        ''' <param name="requiredColumnNames">Column names that must exist in table</param>
        ''' <param name="ignoreCase">Ignore upper/lower case (invariant) of column names</param>
        ''' <returns></returns>
        Public Shared Function ValidateRequiredColumnNames(table As DataTable, requiredColumnNames As String(), ignoreCase As Boolean) As String()
            If requiredColumnNames Is Nothing OrElse requiredColumnNames.Length = 0 Then Return New String() {} 'Shortcut result

            Dim AvailableColumns As New System.Collections.Generic.List(Of String)
            For MyCounter As Integer = 0 To table.Columns.Count - 1
                If ignoreCase Then
                    AvailableColumns.Add(table.Columns(MyCounter).ColumnName.ToLowerInvariant)
                Else
                    AvailableColumns.Add(table.Columns(MyCounter).ColumnName)
                End If
            Next

            Dim MissingColumns As New System.Collections.Generic.List(Of String)
            For MyCounter As Integer = 0 To requiredColumnNames.Length - 1
                If ignoreCase AndAlso AvailableColumns.Contains(requiredColumnNames(MyCounter).ToLowerInvariant) = False Then
                    MissingColumns.Add(requiredColumnNames(MyCounter))
                ElseIf Not ignoreCase AndAlso AvailableColumns.Contains(requiredColumnNames(MyCounter)) = False Then
                    MissingColumns.Add(requiredColumnNames(MyCounter))
                End If
            Next

            Return MissingColumns.ToArray
        End Function

        ''' <summary>
        ''' Reset all cells of a column to DbNull.Value
        ''' </summary>
        ''' <param name="column"></param>
        Public Shared Sub ClearColumnValues(column As DataColumn)
            FillColumnWithStaticValue(column, DBNull.Value)
        End Sub

        ''' <summary>
        ''' Fill all cells of a column with a static value
        ''' </summary>
        ''' <param name="column"></param>
        ''' <param name="value"></param>
        Public Shared Sub FillColumnWithStaticValue(column As DataColumn, value As Object)
            FillColumnWithStaticValue(column, value, False)
        End Sub

        ''' <summary>
        ''' Fill all cells of a column with a static value
        ''' </summary>
        ''' <param name="column"></param>
        ''' <param name="value"></param>
        Public Shared Sub FillColumnWithStaticValue(column As DataColumn, value As Object, onlyIfValueIsDbNull As Boolean)
            Dim Table As DataTable = column.Table
            Dim ColOrdinal As Integer = column.Ordinal
            For MyCounter As Integer = 0 To Table.Rows.Count - 1
                If onlyIfValueIsDbNull Then
                    If IsDBNull(Table.Rows(MyCounter)(ColOrdinal)) Then
                        Table.Rows(MyCounter)(ColOrdinal) = value
                    End If
                Else
                    Table.Rows(MyCounter)(ColOrdinal) = value
                End If
            Next
        End Sub

        ''' <summary>
        ''' Calculate a new value based on a row's content
        ''' </summary>
        ''' <param name="row"></param>
        ''' <returns></returns>
        Public Delegate Function CalculateColumnValue(row As DataRow) As Object

        ''' <summary>
        ''' Fill all cells of a column with a calculated value based on current row data
        ''' </summary>
        ''' <param name="column"></param>
        ''' <param name="valueSetter"></param>
        Public Shared Sub FillColumnWithCalculatedValue(column As DataColumn, valueSetter As CalculateColumnValue)
            Dim Table As DataTable = column.Table
            Dim ColOrdinal As Integer = column.Ordinal
            For MyCounter As Integer = 0 To Table.Rows.Count - 1
                Table.Rows(MyCounter)(ColOrdinal) = valueSetter(Table.Rows(MyCounter))
            Next
        End Sub

        ''' <summary>
        ''' Remove all rows matching the filter expression
        ''' </summary>
        ''' <param name="table">A table</param>
        ''' <param name="filterExpression"></param>
        Public Shared Sub RemoveRowsByFilter(table As DataTable, filterExpression As String)
            Dim FoundRows As DataRow() = table.Select(filterExpression)
            For MyCounter As Integer = 0 To FoundRows.Length - 1
                table.Rows.Remove(FoundRows(MyCounter))
            Next
        End Sub

    End Class

End Namespace