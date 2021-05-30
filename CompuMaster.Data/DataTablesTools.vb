Option Explicit On 
Option Strict On

Imports System.Collections.Generic

Namespace CompuMaster.Data

    ''' <summary>
    '''     Common DataTable operations
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    <CodeAnalysis.SuppressMessage("Major Code Smell", "S3385:""Exit"" statements should not be used", Justification:="<Ausstehend>")>
    Friend Class DataTablesTools

        ''' <summary>
        ''' Remove rows with duplicate values in a given column
        ''' </summary>
        ''' <param name="dataTable">A datatable with duplicate values</param>
        ''' <param name="columnName">Column name of the datatable which contains the duplicate values</param>
        ''' <returns>A datatable with unique records in the specified column</returns>
        Friend Shared Function RemoveDuplicates(ByVal dataTable As DataTable, ByVal columnName As String) As DataTable
            Dim hTable As New Hashtable
            Dim duplicateList As New ArrayList

            'Add list of all the unique item value to hashtable, which stores combination of key, value pair.
            'And add duplicate item value in arraylist.
            Dim drow As DataRow
            For Each drow In dataTable.Rows
                If hTable.Contains(drow(columnName)) Then
                    duplicateList.Add(drow)
                Else
                    hTable.Add(drow(columnName), String.Empty)
                End If
            Next drow
            'Removing a list of duplicate items from datatable.
            Dim daRow As DataRow
            For Each daRow In duplicateList
                dataTable.Rows.Remove(daRow)
            Next daRow
            'Datatable which contains unique records will be return as output.
            Return dataTable
        End Function 'RemoveDuplicateRows

        ''' <summary>
        '''     Drop all columns except the required ones
        ''' </summary>
        ''' <param name="table">A data table containing some columns</param>
        ''' <param name="remainingColumns">A list of column names which shall not be removed</param>
        ''' <remarks>
        '''     If the list of the remaining columns contains some column names which are not existing, then those column names will be ignored. There will be no exception in this case.
        '''     The names of the columns are handled case-insensitive.
        ''' </remarks>
        Friend Shared Sub KeepColumnsAndRemoveAllOthers(ByVal table As DataTable, ByVal remainingColumns As String())
            Dim KeepColFlags(table.Columns.Count - 1) As Boolean
            'Identify unwanted columns
            For MyKeepColCounter As Integer = 0 To remainingColumns.Length - 1
                If remainingColumns(MyKeepColCounter) <> Nothing Then
                    For MyColCounter As Integer = 0 To table.Columns.Count - 1
                        If table.Columns(remainingColumns(MyKeepColCounter)) Is table.Columns(MyColCounter) Then
                            KeepColFlags(MyColCounter) = True
                        End If
                    Next
                End If
            Next
            'Remove unwanted columns
            For MyCounter As Integer = KeepColFlags.Length - 1 To 0 Step -1
                If KeepColFlags(MyCounter) = False Then
                    table.Columns.RemoveAt(MyCounter)
                End If
            Next
        End Sub

        ''' <summary>
        '''     Lookup the row index for a data row in a data table
        ''' </summary>
        ''' <param name="dataRow">The data row whose index number is required</param>
        ''' <returns>An index number for the given data row</returns>
        Friend Shared Function RowIndex(ByVal dataRow As DataRow) As Integer
            If dataRow.Table Is Nothing Then
                Throw New Exception("DataRow must be part of a table to retrieve its row index")
            End If
            For MyCounter As Integer = 0 To dataRow.Table.Rows.Count - 1
                If dataRow.Table.Rows(MyCounter) Is dataRow Then
                    Return MyCounter
                End If
            Next
            Throw New Exception("Unexpected error: provided data row can't be identified in its data table. Please contact your software vendor.")
        End Function

        ''' <summary>
        '''     Lookup the column index for a data column in a data table
        ''' </summary>
        ''' <param name="column">The data column whose index number is required</param>
        ''' <returns>An index number for the given column</returns>
        Friend Shared Function ColumnIndex(ByVal column As DataColumn) As Integer
            If column.Table Is Nothing Then
                Throw New Exception("DataColumn must be part of a table to retrieve its column index")
            End If
            For MyCounter As Integer = 0 To column.Table.Columns.Count - 1
                If column.Table.Columns(MyCounter) Is column Then
                    Return MyCounter
                End If
            Next
            Throw New Exception("Unexpected error: provided data column can't be identified in its data table. Please contact your software vendor.")
        End Function

        ''' <summary>
        '''     Find duplicate values in a given row and calculate the number of occurances of each value in the table
        ''' </summary>
        ''' <param name="column">A column of a datatable</param>
        ''' <returns>A hashtable containing the origin column value as key and the number of occurances as value</returns>
        Friend Shared Function FindDuplicates(ByVal column As DataColumn) As Hashtable
            Return FindDuplicates(column, 2)
        End Function

        ''' <summary>
        '''     Find duplicate values in a given row and calculate the number of occurances of each value in the table
        ''' </summary>
        ''' <param name="column">A column of a datatable</param>
        ''' <param name="minOccurances">Only values with occurances equal or more than this number will be returned</param>
        ''' <returns>A hashtable containing the origin column value as key and the number of occurances as value</returns>
        Friend Shared Function FindDuplicates(ByVal column As DataColumn, ByVal minOccurances As Integer) As Hashtable

            Dim Table As DataTable = column.Table
            Dim Result As New Hashtable

            'Find all elements and count their duplicates number
            For MyCounter As Integer = 0 To Table.Rows.Count - 1
                Dim key As Object = Table.Rows(MyCounter)(column)
                If Result.ContainsKey(key) Then
                    'Increase counter for this existing value by 1
                    Result.Item(key) = CType(Result.Item(key), Integer) + 1
                Else
                    'Add new element
                    Result.Add(key, 1)
                End If
            Next

            'Remove all elements with occurances lesser than the required number
            Dim removeTheseKeys As New ArrayList
            For Each MyKey As DictionaryEntry In Result
                If CType(MyKey.Value, Integer) < minOccurances Then
                    removeTheseKeys.Add(MyKey.Key)
                End If
            Next
            For MyCounter As Integer = 0 To removeTheseKeys.Count - 1
                Result.Remove(removeTheseKeys(MyCounter))
            Next

            Return Result

        End Function

        ''' <summary>
        '''     Find duplicate values in a given row and calculate the number of occurances of each value in the table
        ''' </summary>
        ''' <param name="column">A column of a datatable</param>
        ''' <returns>A hashtable containing the origin column value as key and the number of occurances as value</returns>
        Friend Shared Function FindDuplicates(Of T)(ByVal column As DataColumn) As System.Collections.Generic.Dictionary(Of T, Integer)
            Return FindDuplicates(Of T)(column, 2)
        End Function

        ''' <summary>
        '''     Find duplicate values in a given row and calculate the number of occurances of each value in the table
        ''' </summary>
        ''' <param name="column">A column of a datatable</param>
        ''' <param name="minOccurances">Only values with occurances equal or more than this number will be returned</param>
        ''' <returns>A hashtable containing the origin column value as key and the number of occurances as value</returns>
        Friend Shared Function FindDuplicates(Of T)(ByVal column As DataColumn, ByVal minOccurances As Integer) As System.Collections.Generic.Dictionary(Of T, Integer)

            Dim Table As DataTable = column.Table
            Dim Result As New System.Collections.Generic.Dictionary(Of T, Integer)

            'Find all elements and count their duplicates number
            For MyCounter As Integer = 0 To Table.Rows.Count - 1
                Dim key As T = CType(Table.Rows(MyCounter)(column), T)
                If Result.ContainsKey(key) Then
                    'Increase counter for this existing value by 1
                    Result.Item(key) = CType(Result.Item(key), Integer) + 1
                Else
                    'Add new element
                    Result.Add(key, 1)
                End If
            Next

            'Remove all elements with occurances lesser than the required number
            Dim removeTheseKeys As New System.Collections.Generic.List(Of T)
            For Each MyKey As System.Collections.Generic.KeyValuePair(Of T, Integer) In Result
                If CType(MyKey.Value, Integer) < minOccurances Then
                    removeTheseKeys.Add(MyKey.Key)
                End If
            Next
            For MyCounter As Integer = 0 To removeTheseKeys.Count - 1
                Result.Remove(removeTheseKeys(MyCounter))
            Next

            Return Result

        End Function

        ''' <summary>
        '''     Convert the first two columns into objects which can be consumed by the ListControl objects in the System.Windows.Forms or System.Web.WebControl namespaces
        ''' </summary>
        ''' <param name="datatable">The datatable which contains a key column and a value column for the list control</param>
        ''' <returns>An array of ListControlItem</returns>
        Friend Shared Function ConvertDataTableToListControlItem(ByVal datatable As DataTable) As ListControlItem()
            If datatable Is Nothing Then
                Return Nothing
            ElseIf datatable.Rows.Count = 0 Then
                Return New ListControlItem() {}
            Else
                Dim Result As ListControlItem()
                ReDim Result(datatable.Rows.Count - 1)
                For MyCounter As Integer = 0 To datatable.Rows.Count - 1
                    Result(MyCounter) = New ListControlItem With {
                        .Key = datatable.Rows(MyCounter)(0),
                        .Value = datatable.Rows(MyCounter)(1)
                    }
                Next
                Return Result
            End If
        End Function

        ''' <summary>
        ''' A list item which can be consumed by list controls in namespaces System.Windows as well as in System.Web
        ''' </summary>
        Friend Class ListControlItem

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
        End Class

        ''' <summary>
        '''     Convert a dataset to an xml string with data and schema information
        ''' </summary>
        ''' <param name="dataset"></param>
        ''' <returns></returns>
        Friend Shared Function ConvertDatasetToXml(ByVal dataset As DataSet) As String
            Dim sbuilder As New System.Text.StringBuilder
            Dim xmlSW As New System.IO.StringWriter(sbuilder)
            dataset.WriteXml(xmlSW, XmlWriteMode.WriteSchema)
            xmlSW.Close()
            Return sbuilder.ToString
        End Function

        ''' <summary>
        '''     Convert an xml string to a dataset
        ''' </summary>
        ''' <param name="xml"></param>
        ''' <returns></returns>
        Friend Shared Function ConvertXmlToDataset(ByVal xml As String) As DataSet
            Dim reader As New System.IO.StringReader(xml)
            Dim DataSet As New DataSet
            DataSet.ReadXml(reader, XmlReadMode.Auto)
            reader.Close()
            Return DataSet
        End Function

        ''' <summary>
        '''     Create a new data table clone with only some first rows
        ''' </summary>
        ''' <param name="SourceTable">The source table to be copied</param>
        ''' <param name="NumberOfRows">The number of rows to be copied</param>
        ''' <returns>The new clone of the datatable</returns>
        Friend Shared Function GetDataTableWithSubsetOfRows(ByVal SourceTable As DataTable, ByVal NumberOfRows As Integer) As DataTable
            Return GetDataTableWithSubsetOfRows(SourceTable, 0, NumberOfRows)
        End Function

        ''' <summary>
        '''     Create a new data table clone with only some first rows
        ''' </summary>
        ''' <param name="SourceTable">The source table to be copied</param>
        ''' <param name="StartAtRow">The position where to start the copy process, the first row is at 0</param>
        ''' <param name="NumberOfRows">The number of rows to be copied</param>
        ''' <returns>The new clone of the datatable</returns>
        Friend Shared Function GetDataTableWithSubsetOfRows(ByVal SourceTable As DataTable, ByVal StartAtRow As Integer, ByVal NumberOfRows As Integer) As DataTable
            Dim Result As DataTable = SourceTable.Clone
            Dim MyRows As DataRowCollection = SourceTable.Rows
            Dim LastRowIndex As Integer

            If NumberOfRows = Integer.MaxValue Then
                'Read to end
                LastRowIndex = MyRows.Count - 1
            Else
                'Read only the given number of rows
                LastRowIndex = StartAtRow + NumberOfRows - 1
                'Verify that we're not going to read more rows than existant
                If LastRowIndex >= MyRows.Count Then
                    LastRowIndex = MyRows.Count - 1
                End If
            End If

            If MyRows IsNot Nothing AndAlso MyRows.Count > 0 Then
                For MyRowCounter As Integer = StartAtRow To LastRowIndex
                    Dim MyNewRow As DataRow = Result.NewRow
                    MyNewRow.ItemArray = MyRows(MyRowCounter).ItemArray
                    Result.Rows.Add(MyNewRow)
                Next
            End If

            Return Result

        End Function

        ''' <summary>
        '''     Creates a complete clone of a DataRow with structure as well as data
        ''' </summary>
        ''' <param name="sourceRow">The source row to be copied</param>
        ''' <returns>The new clone of the DataRow</returns>
        ''' <remarks>
        '''     The resulting DataRow has got the schema from the sourceRow's DataTable, but it hasn't been added to the table yet.
        ''' </remarks>
        Public Shared Function CreateDataRowClone(ByVal sourceRow As DataRow) As DataRow
            If sourceRow Is Nothing Then Throw New ArgumentNullException(NameOf(sourceRow))
            Dim Result As DataRow = sourceRow.Table.NewRow
            Result.ItemArray = sourceRow.ItemArray
            Return Result
        End Function

        ''' <summary>
        '''     Creates a complete clone of a DataTable with structure as well as data
        ''' </summary>
        ''' <param name="SourceTable">The source table to be copied</param>
        ''' <returns>The new clone of the datatable</returns>
        Friend Shared Function GetDataTableClone(ByVal SourceTable As DataTable) As DataTable
            Return GetDataTableClone(SourceTable, Nothing, Nothing)
        End Function

        ''' <summary>
        '''     Creates a complete clone of a DataTable with structure as well as data
        ''' </summary>
        ''' <param name="SourceTable">The source table to be copied</param>
        ''' <param name="RowFilter">An additional row filter, for all rows set it to null (Nothing in VisualBasic)</param>
        ''' <returns>The new clone of the datatable</returns>
        Friend Shared Function GetDataTableClone(ByVal SourceTable As DataTable, ByVal RowFilter As String) As DataTable
            Return GetDataTableClone(SourceTable, RowFilter, Nothing)
        End Function

        ''' <summary>
        '''     Creates a complete clone of a DataTable with structure as well as data
        ''' </summary>
        ''' <param name="SourceTable">The source table to be copied</param>
        ''' <param name="RowFilter">An additional row filter, for all rows set it to null (Nothing in VisualBasic)</param>
        ''' <param name="Sort">An additional sort command</param>
        ''' <returns>The new clone of the datatable</returns>
        Friend Shared Function GetDataTableClone(ByVal SourceTable As DataTable, ByVal RowFilter As String, ByVal Sort As String) As DataTable
            Return GetDataTableClone(SourceTable, RowFilter, Sort, Nothing)
        End Function

        ''' <summary>
        '''     Creates a complete clone of a DataTable with structure as well as data
        ''' </summary>
        ''' <param name="SourceTable">The source table to be copied</param>
        ''' <param name="RowFilter">An additional row filter, for all rows set it to null (Nothing in VisualBasic)</param>
        ''' <param name="Sort">An additional sort command</param>
        ''' <param name="topRows">How many rows from top shall be returned as maximum?</param>
        ''' <returns>The new clone of the datatable</returns>
        Friend Shared Function GetDataTableClone(ByVal SourceTable As DataTable, ByVal RowFilter As String, ByVal Sort As String, ByVal topRows As Integer) As DataTable
            Dim Result As DataTable = SourceTable.Clone
            Dim MyRows As DataRow() = SourceTable.Select(RowFilter, Sort)

            If topRows = Nothing Then
                'All rows
                topRows = Integer.MaxValue
            End If

            If MyRows IsNot Nothing Then
                For MyCounter As Integer = 1 To MyRows.Length
                    If MyCounter > topRows Then
                        Exit For
                    Else
                        Dim MyNewRow As DataRow = Result.NewRow
                        MyNewRow.ItemArray = MyRows(MyCounter - 1).ItemArray
                        Result.Rows.Add(MyNewRow)
                    End If
                Next
            End If

            Return Result

        End Function

        ''' <summary>
        '''     Creates a clone of a dataview but as a new data table
        ''' </summary>
        ''' <param name="data">The data view to create the data table from</param>
        ''' <returns></returns>
        Friend Shared Function ConvertDataViewToDataTable(ByVal data As DataView) As System.Data.DataTable
            Dim Result As DataTable = data.Table.Clone
            'Dim MyRows As DataRowView() = data.Item

            If data.Count > 0 Then
                For MyCounter As Integer = 1 To data.Count
                    Dim MyNewRow As DataRow = Result.NewRow
                    MyNewRow.ItemArray = data.Item(MyCounter - 1).Row.ItemArray
                    Result.Rows.Add(MyNewRow)
                Next
            End If

            Return Result

        End Function

        ''' <summary>
        '''     Convert an ArrayList to a datatable
        ''' </summary>
        ''' <param name="arrayList">An ArrayList with some content</param>
        ''' <returns>Datatable with column &quot;key&quot; and &quot;value&quot;</returns>
        <Obsolete("use ConvertICollectionToDataTable instead", False)> Friend Shared Function ConvertArrayListToDataTable(ByVal arrayList As ArrayList) As DataTable
            Return ConvertICollectionToDataTable(arrayList)
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
        Friend Shared Function ConvertDataTableToHashtable(ByVal keyColumn As DataColumn, ByVal valueColumn As DataColumn) As Hashtable
            If keyColumn.Table IsNot valueColumn.Table Then
                Throw New Exception("Key column and value column must be from the same table")
            End If
            Return ConvertDataTableToHashtable(keyColumn.Table, keyColumn.Ordinal, valueColumn.Ordinal)
        End Function

        ''' <summary>
        '''     Convert a data table to a hash table
        ''' </summary>
        ''' <param name="data">The first two columns of this data table will be used</param>
        ''' <returns></returns>
        ''' <remarks>
        '''     ATTENTION: the very first column is used as key column and must be unique therefore
        ''' </remarks>
        Friend Shared Function ConvertDataTableToHashtable(ByVal data As DataTable) As Hashtable
            Return ConvertDataTableToHashtable(data, 0, 1)
        End Function

        ''' <summary>
        '''     Convert a data table to a hash table
        ''' </summary>
        ''' <param name="data">The data table with the content</param>
        ''' <param name="keyColumnIndex">This is the key column from the data table and MUST BE UNIQUE (make it unique, first!)</param>
        ''' <param name="valueColumnIndex">A column which contains the values</param>
        ''' <returns></returns>
        ''' <remarks>
        '''     ATTENTION: the very first column is used as key column and must be unique therefore
        ''' </remarks>
        Friend Shared Function ConvertDataTableToHashtable(ByVal data As DataTable, ByVal keyColumnIndex As Integer, ByVal valueColumnIndex As Integer) As Hashtable
            If data.Columns(keyColumnIndex).Unique = False Then
                Throw New Exception("The hashtable requires your key column to be a unique column - make it a unique column, first!")
            End If
            Dim Result As New Hashtable
            For MyCounter As Integer = 0 To data.Rows.Count - 1
                Result.Add(data.Rows(MyCounter)(keyColumnIndex), data.Rows(MyCounter)(valueColumnIndex))
            Next
            Return Result
        End Function

        ''' <summary>
        '''     Convert a data table to an array of dictionary entries
        ''' </summary>
        ''' <param name="data">The first two columns of this data table will be used</param>
        ''' <returns></returns>
        ''' <remarks>
        '''     The very first column is used as key column, the second one as the value column
        ''' </remarks>
        Friend Shared Function ConvertDataTableToDictionaryEntryArray(ByVal data As DataTable) As DictionaryEntry()
            Return ConvertDataTableToDictionaryEntryArray(data, 0, 1)
        End Function

        ''' <summary>
        '''     Convert a data table to an array of dictionary entries
        ''' </summary>
        ''' <param name="keyColumn">This is the key column from the data table</param>
        ''' <param name="valueColumn">A column which contains the values</param>
        ''' <returns></returns>
        Friend Shared Function ConvertDataTableToDictionaryEntryArray(ByVal keyColumn As DataColumn, ByVal valueColumn As DataColumn) As DictionaryEntry()
            If keyColumn.Table IsNot valueColumn.Table Then
                Throw New Exception("Key column and value column must be from the same table")
            End If
            Return ConvertDataTableToDictionaryEntryArray(keyColumn.Table, keyColumn.Ordinal, valueColumn.Ordinal)
        End Function

        ''' <summary>
        '''     Convert a data table to an array of dictionary entries
        ''' </summary>
        ''' <param name="data">The data table with the content</param>
        ''' <param name="keyColumnIndex">This is the key column from the data table</param>
        ''' <param name="valueColumnIndex">A column which contains the values</param>
        ''' <returns></returns>
        Friend Shared Function ConvertDataTableToDictionaryEntryArray(ByVal data As DataTable, ByVal keyColumnIndex As Integer, ByVal valueColumnIndex As Integer) As DictionaryEntry()
            Dim Result As DictionaryEntry()
            ReDim Result(data.Rows.Count - 1)
            For MyCounter As Integer = 0 To data.Rows.Count - 1
                Result(MyCounter) = New DictionaryEntry(data.Rows(MyCounter)(keyColumnIndex), data.Rows(MyCounter)(valueColumnIndex))
            Next
            Return Result
        End Function

        ''' <summary>
        '''     Convert a hashtable to a datatable
        ''' </summary>
        ''' <param name="hashtable">A hashtable with some content</param>
        ''' <returns>Datatable with column &quot;key&quot; and &quot;value&quot;</returns>
        <Obsolete("use ConvertIDictionaryToDataTable instead and pay attention to parameter keyIsUnique", False)>
        Friend Shared Function ConvertHashtableToDataTable(ByVal hashtable As Hashtable) As DataTable
            Return ConvertIDictionaryToDataTable(hashtable, True)
        End Function

        ''' <summary>
        '''     Convert an ICollection to a datatable
        ''' </summary>
        ''' <param name="collection">An ICollection with some content</param>
        ''' <returns>Datatable with column &quot;key&quot; and &quot;value&quot;</returns>
        Friend Shared Function ConvertICollectionToDataTable(ByVal collection As ICollection) As DataTable
            Dim Result As New DataTable
            Result.Columns.Add(New DataColumn("value"))

            For Each MyKey As Object In collection
                Dim MyRow As DataRow = Result.NewRow
                MyRow(0) = MyKey
                Result.Rows.Add(MyRow)
            Next

            Return Result
        End Function

        ''' <summary>
        '''     Convert an IDictionary to a datatable
        ''' </summary>
        ''' <param name="dictionary">An IDictionary with some content</param>
        ''' <returns>Datatable with column &quot;key&quot; and &quot;value&quot;</returns>
        Friend Shared Function ConvertIDictionaryToDataTable(ByVal dictionary As IDictionary) As DataTable
            Return ConvertIDictionaryToDataTable(dictionary, False)
        End Function

        ''' <summary>
        '''     Convert an IDictionary to a datatable
        ''' </summary>
        ''' <param name="dictionary">An IDictionary with some content</param>
        ''' <param name="keyIsUnique">If true, the key column in the data table will be marked as unique</param>
        ''' <returns>Datatable with column &quot;key&quot; and &quot;value&quot;</returns>
        Friend Shared Function ConvertIDictionaryToDataTable(ByVal dictionary As IDictionary, ByVal keyIsUnique As Boolean) As DataTable
            Dim Result As New DataTable
            Result.Columns.Add(New DataColumn("key"))
            Result.Columns("key").Unique = keyIsUnique
            Result.Columns.Add(New DataColumn("value"))

            For Each MyKey As Object In dictionary.Keys
                Dim MyRow As DataRow = Result.NewRow
                MyRow(0) = MyKey
                MyRow(1) = dictionary(MyKey)
                Result.Rows.Add(MyRow)
            Next

            Return Result
        End Function

        ''' <summary>
        '''     Convert an array of DictionaryEntry to a datatable
        ''' </summary>
        ''' <param name="dictionaryEntries">An array of DictionaryEntry with some content</param>
        ''' <param name="keyIsUnique">If true, the key column in the data table will be marked as unique</param>
        ''' <returns>Datatable with column &quot;key&quot; and &quot;value&quot;</returns>
        ''' <remarks>
        ''' </remarks>
        Friend Shared Function ConvertDictionaryEntryArrayToDataTable(ByVal dictionaryEntries As DictionaryEntry(), ByVal keyIsUnique As Boolean) As DataTable
            Dim Result As New DataTable
            Result.Columns.Add(New DataColumn("key"))
            Result.Columns("key").Unique = keyIsUnique
            Result.Columns.Add(New DataColumn("value"))

            For Each entry As DictionaryEntry In dictionaryEntries
                Dim MyRow As DataRow = Result.NewRow
                MyRow(0) = entry.Key
                MyRow(1) = entry.Value
                Result.Rows.Add(MyRow)
            Next

            Return Result
        End Function

        ''' <summary>
        '''     Convert a NameValueCollection to a datatable
        ''' </summary>
        ''' <param name="nameValueCollection">An IDictionary with some content</param>
        ''' <returns>Datatable with column &quot;key&quot; and &quot;value&quot;</returns>
        Friend Shared Function ConvertNameValueCollectionToDataTable(ByVal nameValueCollection As Specialized.NameValueCollection) As DataTable
            Return ConvertNameValueCollectionToDataTable(nameValueCollection, False)
        End Function

        ''' <summary>
        '''     Convert a NameValueCollection to a datatable
        ''' </summary>
        ''' <param name="nameValueCollection">An IDictionary with some content</param>
        ''' <param name="keyIsUnique">If true, the key column in the data table will be marked as unique</param>
        ''' <returns>Datatable with column &quot;key&quot; and &quot;value&quot;</returns>
        Friend Shared Function ConvertNameValueCollectionToDataTable(ByVal nameValueCollection As Specialized.NameValueCollection, ByVal keyIsUnique As Boolean) As DataTable
            Dim Result As New DataTable
            Result.Columns.Add(New DataColumn("key"))
            Result.Columns("key").Unique = keyIsUnique
            Result.Columns.Add(New DataColumn("value"))

            For Each MyKey As String In nameValueCollection.Keys
                Dim MyRow As DataRow = Result.NewRow
                MyRow(0) = MyKey
                MyRow(1) = nameValueCollection(MyKey)
                Result.Rows.Add(MyRow)
            Next

            Return Result
        End Function

        ''' <summary>
        '''     Simplified creation of a DataTable by definition of a SQL statement and a connection string
        ''' </summary>
        ''' <param name="strSQL">The SQL statement to retrieve the data</param>
        ''' <param name="ConnectionString">The connection string to the data source</param>
        ''' <param name="NameOfNewDataTable">The name of the new DataTable</param>
        ''' <returns>A filled DataTable</returns>
        Friend Shared Function GetDataTableViaODBC(ByVal strSQL As String, ByVal ConnectionString As String, ByVal NameOfNewDataTable As String) As DataTable

            Dim MyConn As New Odbc.OdbcConnection(ConnectionString)
            Dim MyDataTable As New DataTable(NameOfNewDataTable)
            Dim MyCmd As New Odbc.OdbcCommand(strSQL, MyConn)
            Dim MyDA As New Odbc.OdbcDataAdapter(MyCmd)
            Try
                MyConn.Open()
                MyDA.Fill(MyDataTable)
            Finally
                If MyDA IsNot Nothing Then
                    MyDA.Dispose()
                End If
                If MyCmd IsNot Nothing Then
                    MyCmd.Dispose()
                End If
                If MyConn IsNot Nothing Then
                    If MyConn.State <> ConnectionState.Closed Then
                        MyConn.Close()
                    End If
                    MyConn.Dispose()
                End If
            End Try

            Return MyDataTable

        End Function

        ''' <summary>
        '''     Simplified creation of a DataTable by definition of a SQL statement and a connection string
        ''' </summary>
        ''' <param name="strSQL">The SQL statement to retrieve the data</param>
        ''' <param name="ConnectionString">The connection string to the data source</param>
        ''' <param name="NameOfNewDataTable">The name of the new DataTable</param>
        ''' <returns>A filled DataTable</returns>
        Friend Shared Function GetDataTableViaSqlClient(ByVal strSQL As String, ByVal ConnectionString As String, ByVal NameOfNewDataTable As String) As DataTable

            Dim MyConn As New System.Data.SqlClient.SqlConnection(ConnectionString)
            Dim MyDataTable As New DataTable(NameOfNewDataTable)
            Dim MyCmd As New System.Data.SqlClient.SqlCommand(strSQL, MyConn)
            Dim MyDA As New System.Data.SqlClient.SqlDataAdapter(MyCmd)

            Try
                MyConn.Open()
                MyDA.Fill(MyDataTable)
            Finally
                If MyDA IsNot Nothing Then
                    MyDA.Dispose()
                End If
                If MyCmd IsNot Nothing Then
                    MyCmd.Dispose()
                End If
                If MyConn IsNot Nothing Then
                    If MyConn.State <> ConnectionState.Closed Then
                        MyConn.Close()
                    End If
                    MyConn.Dispose()
                End If
            End Try

            Return MyDataTable

        End Function

        ''' <summary>
        '''     Return a string with all columns and rows as an html table
        ''' </summary>
        ''' <param name="rows">The rows to be processed</param>
        ''' <param name="label">An optional title of the rows</param>
        ''' <param name="titleTagOpener">The opening tag in front of the table's title</param>
        ''' <param name="titleTagEnd">The closing tag after the table title</param>
        ''' <param name="additionalTableAttributes">Additional attributes for the rendered table</param>
        ''' <param name="htmlEncodeCellContentAndLineBreaks">Encode all output to valid HTML</param>
        ''' <returns>If no rows have been processed, the return string is nothing</returns>
        Friend Shared Function ConvertToHtmlTable(ByVal rows As DataRowCollection, ByVal label As String, ByVal titleTagOpener As String,
                                                  ByVal titleTagEnd As String, ByVal additionalTableAttributes As String,
                                                  ByVal htmlEncodeCellContentAndLineBreaks As Boolean,
                                                  disableHtmlEncodingForColumns As String()) As String
            Dim RowList As New List(Of DataRow)
            For MyCounter As Integer = 0 To rows.Count - 1
                RowList.Add(rows.Item(MyCounter))
            Next
            Return ConvertToHtmlTable(RowList, label, titleTagOpener, titleTagEnd, additionalTableAttributes, htmlEncodeCellContentAndLineBreaks, disableHtmlEncodingForColumns)
        End Function

        ''' <summary>
        '''     Return a string with all columns and rows as an html table
        ''' </summary>
        ''' <param name="rows">The rows to be processed</param>
        ''' <param name="label">An optional title of the rows</param>
        ''' <param name="titleTagOpener">The opening tag in front of the table's title</param>
        ''' <param name="titleTagEnd">The closing tag after the table title</param>
        ''' <param name="additionalTableAttributes">Additional attributes for the rendered table</param>
        ''' <param name="htmlEncodeCellContentAndLineBreaks">Encode all output to valid HTML</param>
        ''' <returns>If no rows have been processed, the return string is nothing</returns>
        Friend Shared Function ConvertToHtmlTable(ByVal rows As DataRow(), ByVal label As String, ByVal titleTagOpener As String,
                                                  ByVal titleTagEnd As String, ByVal additionalTableAttributes As String,
                                                  ByVal htmlEncodeCellContentAndLineBreaks As Boolean,
                                                  disableHtmlEncodingForColumns As String()) As String
            Dim RowList As New List(Of DataRow)(rows)
            Return ConvertToHtmlTable(RowList, label, titleTagOpener, titleTagEnd, additionalTableAttributes, htmlEncodeCellContentAndLineBreaks, disableHtmlEncodingForColumns)
        End Function

        ''' <summary>
        '''     Return a string with all columns and rows as an html table
        ''' </summary>
        ''' <param name="rows">The rows to be processed</param>
        ''' <param name="label">An optional title of the rows</param>
        ''' <param name="titleTagOpener">The opening tag in front of the table's title</param>
        ''' <param name="titleTagEnd">The closing tag after the table title</param>
        ''' <param name="additionalTableAttributes">Additional attributes for the rendered table</param>
        ''' <param name="htmlEncodeCellContentAndLineBreaks">Encode all output to valid HTML</param>
        ''' <returns>If no rows have been processed, the return string is nothing</returns>
        Friend Shared Function ConvertToHtmlTable(ByVal rows As List(Of DataRow), ByVal label As String, ByVal titleTagOpener As String,
                                                  ByVal titleTagEnd As String, ByVal additionalTableAttributes As String,
                                                  ByVal htmlEncodeCellContentAndLineBreaks As Boolean,
                                                  disableHtmlEncodingForColumns As String()) As String
            If disableHtmlEncodingForColumns Is Nothing Then
                disableHtmlEncodingForColumns = New String() {}
            End If
            If titleTagOpener Is Nothing AndAlso titleTagEnd Is Nothing Then
                titleTagOpener = "<H1>"
                titleTagEnd = "</H1>"
            End If

            Dim Result As New System.Text.StringBuilder
            If label IsNot Nothing Then
                Result.Append(titleTagOpener)
                If htmlEncodeCellContentAndLineBreaks Then
                    Result.Append(HtmlEncodeLineBreaks(System.Web.HttpUtility.HtmlEncode(String.Format("{0}", label))))
                Else
                    Result.Append(String.Format("{0}", label))
                End If
                Result.Append(titleTagEnd & System.Environment.NewLine)
            End If
            If rows.Count <= 0 Then
                Return Nothing
            End If
            Result.Append("<TABLE ")
            Result.Append(additionalTableAttributes)
            Result.Append("><TR>")
            For Each column As DataColumn In rows(0).Table.Columns
                Dim HtmlEncodeCellContentDisabledForCurrentColumn As Boolean = (
                    htmlEncodeCellContentAndLineBreaks = False OrElse
                    Not (Array.IndexOf(Of String)(disableHtmlEncodingForColumns, column.ColumnName) = -1)
                    )
                Result.Append("<TH>")
                If column.Caption <> Nothing Then
                    If htmlEncodeCellContentAndLineBreaks AndAlso HtmlEncodeCellContentDisabledForCurrentColumn = False Then
                        Result.Append(HtmlEncodeLineBreaks(System.Web.HttpUtility.HtmlEncode(String.Format("{0}", column.Caption))))
                    Else
                        Result.Append(String.Format("{0}", column.Caption))
                    End If
                Else
                    If htmlEncodeCellContentAndLineBreaks AndAlso HtmlEncodeCellContentDisabledForCurrentColumn = False Then
                        Result.Append(HtmlEncodeLineBreaks(System.Web.HttpUtility.HtmlEncode(String.Format("{0}", column.ColumnName))))
                    Else
                        Result.Append(String.Format("{0}", column.ColumnName))
                    End If
                End If
                Result.Append("</TH>")
            Next
            Result.Append("</TR>")
            Result.Append(System.Environment.NewLine)
            For Each row As DataRow In rows
                Result.Append("<TR>")
                For Each column As DataColumn In row.Table.Columns
                    Dim HtmlEncodeCellContentDisabledForCurrentColumn As Boolean = (
                        htmlEncodeCellContentAndLineBreaks = False OrElse
                        Not (Array.IndexOf(Of String)(disableHtmlEncodingForColumns, column.ColumnName) = -1)
                        )
                    Result.Append("<TD>")
                    If htmlEncodeCellContentAndLineBreaks AndAlso HtmlEncodeCellContentDisabledForCurrentColumn = False Then
                        Result.Append(HtmlEncodeLineBreaks(System.Web.HttpUtility.HtmlEncode(String.Format("{0}", row(column)))))
                    Else
                        Result.Append(String.Format("{0}", row(column)))
                    End If
                    Result.Append("</TD>")
                Next
                Result.Append("</TR>")
                Result.Append(System.Environment.NewLine)
            Next
            Result.Append("</TABLE>")
            Return Result.ToString
        End Function

        ''' <summary>
        '''     Converts all line breaks into HTML line breaks (&quot;&lt;br&gt;&quot;)
        ''' </summary>
        ''' <param name="Text"></param>
        ''' <returns></returns>
        ''' <remarks>
        '''     Supported line breaks are linebreaks of Windows, MacOS as well as Linux/Unix.
        ''' </remarks>
        Private Shared Function HtmlEncodeLineBreaks(ByVal Text As String) As String
            Return Text.Replace(ControlChars.CrLf, "<br>").Replace(ControlChars.Cr, "<br>").Replace(ControlChars.Lf, "<br>")
        End Function

        ''' <summary>
        '''     Return a string with all columns and rows, helpfull for debugging purposes
        ''' </summary>
        ''' <param name="dataTable">The datatable to retrieve the content from</param>
        ''' <returns>All rows are tab separated. If no rows have been processed, the user will get notified about this fact</returns>
        Friend Shared Function ConvertToPlainTextTable(ByVal dataTable As DataTable) As String
            Return ConvertToPlainTextTableInternal(dataTable.Rows, dataTable.TableName)
        End Function

        ''' <summary>
        '''     Return a string with all columns and rows, helpfull for debugging purposes
        ''' </summary>
        ''' <param name="rows">The rows to be processed</param>
        ''' <param name="label">An optional title of the rows</param>
        ''' <returns>All rows are tab separated. If no rows have been processed, the user will get notified about this fact</returns>
        Friend Shared Function ConvertToPlainTextTable(ByVal rows() As DataRow, ByVal label As String) As String
            Const separator As Char = ControlChars.Tab
            Dim Result As New System.Text.StringBuilder
            If label <> "" Then
                Result.Append(String.Format("{0}", label) & System.Environment.NewLine)
            End If
            If rows.Length <= 0 Then
                Result.Append("no rows found" & System.Environment.NewLine)
                Return Result.ToString
            End If
            For Each column As DataColumn In rows(0).Table.Columns
                If column.Ordinal <> 0 Then Result.Append(separator)
                If column.Caption <> Nothing Then
                    Result.Append(String.Format("{0}", column.Caption))
                Else
                    Result.Append(String.Format("{0}", column.ColumnName))
                End If
            Next
            Result.Append(System.Environment.NewLine)
            For Each row As DataRow In rows
                For Each column As DataColumn In row.Table.Columns
                    If column.Ordinal <> 0 Then Result.Append(separator)
                    Result.Append(String.Format("{0}", row(column)))
                Next
                Result.Append(System.Environment.NewLine)
            Next
            Return Result.ToString
        End Function

        ''' <summary>
        '''     Return a string with all columns and rows, helpfull for debugging purposes
        ''' </summary>
        ''' <param name="rows">The rows to be processed</param>
        ''' <param name="label">An optional title of the rows</param>
        ''' <returns>All rows are tab separated. If no rows have been processed, the user will get notified about this fact</returns>
        Private Shared Function ConvertToPlainTextTableInternal(ByVal rows As DataRowCollection, ByVal label As String) As String
            Const separator As Char = ControlChars.Tab
            Dim Result As New System.Text.StringBuilder
            If label <> "" Then
                Result.Append(String.Format("{0}", label) & System.Environment.NewLine)
            End If
            If rows.Count <= 0 Then
                Result.Append("no rows found" & System.Environment.NewLine)
                Return Result.ToString
            End If
            For Each column As DataColumn In rows(0).Table.Columns
                If column.Ordinal <> 0 Then Result.Append(separator)
                If column.Caption <> Nothing Then
                    Result.Append(String.Format("{0}", column.Caption))
                Else
                    Result.Append(String.Format("{0}", column.ColumnName))
                End If
            Next
            Result.Append(System.Environment.NewLine)
            For Each row As DataRow In rows
                If row.RowState <> DataRowState.Deleted Then
                    For Each column As DataColumn In row.Table.Columns
                        If column.Ordinal <> 0 Then Result.Append(separator)
                        Result.Append(String.Format("{0}", row(column)))
                    Next
                    Result.Append(System.Environment.NewLine)
                End If
            Next
            Return Result.ToString
        End Function

        ''' <summary>
        '''     Remove the specified columns if they exist
        ''' </summary>
        ''' <param name="datatable">A datatable where the operations shall be made</param>
        ''' <param name="columnNames">The names of the columns which shall be removed</param>
        ''' <remarks>
        '''     The columns will only be removed if they exist. If a column name doesn't exist, it will be ignored.
        ''' </remarks>
        Public Shared Sub RemoveColumns(ByVal datatable As System.Data.DataTable, ByVal columnNames As String())
            If columnNames IsNot Nothing Then
                For MyRemoveCounter As Integer = 0 To columnNames.Length - 1
                    For MyColumnsCounter As Integer = datatable.Columns.Count - 1 To 0 Step -1
                        If datatable.Columns(MyColumnsCounter).ColumnName = columnNames(MyRemoveCounter) Then
                            datatable.Columns.RemoveAt(MyColumnsCounter)
                        End If
                    Next
                Next
            End If
        End Sub

        ''' <summary>
        '''     Return a string with all columns and rows, helpfull for debugging purposes
        ''' </summary>
        ''' <param name="rows">The rows to be processed</param>
        ''' <param name="label">An optional title of the rows</param>
        ''' <returns>All rows are tab separated. If no rows have been processed, the user will get notified about this fact</returns>
        Friend Shared Function ConvertToPlainTextTable(ByVal rows As DataRowCollection, ByVal label As String) As String
            Return ConvertToPlainTextTableInternal(rows, label)
        End Function

        ''' <summary>
        '''     Convert any opened datareader into a dataset
        ''' </summary>
        ''' <param name="dataReader">An already opened dataReader</param>
        ''' <returns>A dataset containing all datatables the dataReader was able to read</returns>
        Friend Shared Function ConvertDataReaderToDataSet(ByVal datareader As IDataReader) As DataSet
            Dim Result As New DataSet
            Dim DRA As New DataReaderAdapter
            DRA.FillFromReader(Result, datareader)
            Return Result
        End Function

        ''' <summary>
        '''     Convert any opened datareader into a data table
        ''' </summary>
        ''' <param name="dataReader">An already opened dataReader</param>
        ''' <returns>A data table containing all data the dataReader was able to read</returns>
        Friend Shared Function ConvertDataReaderToDataTable(ByVal dataReader As IDataReader) As DataTable
            Return ConvertDataReaderToDataTable(dataReader, Nothing)
        End Function

        ''' <summary>
        '''     Convert any opened datareader into a data table
        ''' </summary>
        ''' <param name="dataReader">An already opened dataReader</param>
        ''' <param name="tableName">The name for the new table</param>
        ''' <returns>A data table containing all data the dataReader was able to read</returns>
        Friend Shared Function ConvertDataReaderToDataTable(ByVal dataReader As IDataReader, ByVal tableName As String) As DataTable

            Dim Result As DataTable
            Dim DRA As New DataReaderAdapter
            If tableName Is Nothing Then
                Result = New DataTable
            Else
                Result = New DataTable(tableName)
            End If
            DRA.FillFromReader(Result, dataReader)
            Return Result

        End Function

        ''' <summary>
        '''     A data adapter for data readers making the real conversion
        ''' </summary>
        Private Class DataReaderAdapter
            Inherits System.Data.Common.DbDataAdapter

            Friend Function FillFromReader(ByVal dataTable As DataTable, ByVal dataReader As IDataReader) As Integer
                Return Me.Fill(dataTable, dataReader)
            End Function

            Friend Function FillFromReader(ByVal dataSet As DataSet, ByVal dataReader As IDataReader) As Integer
                Return Me.Fill(dataSet, "Table", dataReader, 0, 0)
            End Function

            Protected Overrides Function CreateRowUpdatedEvent(ByVal dataRow As System.Data.DataRow, ByVal command As System.Data.IDbCommand,
                                                               ByVal statementType As System.Data.StatementType,
                                                               ByVal tableMapping As System.Data.Common.DataTableMapping) As System.Data.Common.RowUpdatedEventArgs
                Return Nothing
            End Function

            Protected Overrides Function CreateRowUpdatingEvent(ByVal dataRow As System.Data.DataRow, ByVal command As System.Data.IDbCommand,
                                                                ByVal statementType As System.Data.StatementType,
                                                                ByVal tableMapping As System.Data.Common.DataTableMapping) As System.Data.Common.RowUpdatingEventArgs
                Return Nothing
            End Function

            Protected Overrides Sub OnRowUpdated(ByVal value As System.Data.Common.RowUpdatedEventArgs)
                'don't execute base code
            End Sub

            Protected Overrides Sub OnRowUpdating(ByVal value As System.Data.Common.RowUpdatingEventArgs)
                'don't execute base code
            End Sub

        End Class

        ''' <summary>
        '''     Table join types
        ''' </summary>
        Friend Enum JoinTypes As Integer
            ''' <summary>
            '''     The result contains only those rows which exist in both tables
            ''' </summary>
            Inner = 0
            ''' <summary>
            '''     The result contains all rows of the left, parent table and only those rows of the other table which are related to the rows of the left table
            ''' </summary>
            Left = 1
            '''' <summary>
            ''''     The result contains all rows of the left, parent table and all rows of the right, child table. Missing values on the other side will be of value DBNull 
            '''' </summary>
            'Full = 2
            '''' <summary>
            ''''     The result contains all rows of the right, parent table and only those rows of the other table which are related to the rows of the right table
            '''' </summary>
            'Right=3
        End Enum

        ''' <summary>
        '''     Execute a table join on two tables of the same dataset based on the first relation found
        ''' </summary>
        ''' <param name="leftParentTable"></param>
        ''' <param name="rightChildTable"></param>
        ''' <param name="joinType">Inner or left join</param>
        ''' <returns></returns>
        Friend Shared Function JoinTables(ByVal leftParentTable As DataTable, ByVal rightChildTable As DataTable, ByVal joinType As JoinTypes) As DataTable

            'Find the appropriate relation information
            Dim ActiveRelation As DataRelation = Nothing
            For MyRelCounter As Integer = 0 To leftParentTable.ChildRelations.Count - 1
                If leftParentTable.ChildRelations(MyRelCounter).ChildTable Is rightChildTable Then
                    ActiveRelation = leftParentTable.ChildRelations(MyRelCounter)
                    Exit For
                End If
            Next

            Return JoinTables(leftParentTable, rightChildTable, ActiveRelation, joinType)

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
        '''     <list>
        '''         <item>all columns from the left parent table</item>
        '''         <item>INNER JOIN: those columns from the right child table which are not member of the relation in charge</item>
        '''         <item>LEFT JOIN: all columns from the right child table</item>
        '''     </list>
        ''' </remarks>
        Friend Shared Function JoinTables(ByVal leftParentTable As DataTable, ByVal rightChildTable As DataTable, ByVal relation As DataRelation,
                                          ByVal joinType As JoinTypes) As DataTable

            'Verify parameters
            If leftParentTable Is Nothing OrElse rightChildTable Is Nothing Then
                Throw New Exception("One or both table references are null")
            End If
            If leftParentTable.DataSet Is Nothing OrElse leftParentTable.DataSet IsNot rightChildTable.DataSet Then
                Throw New Exception("Both tables must be member of the same dataset")
            End If
            If relation Is Nothing Then
                Throw New Exception("No relation defined between the two tables")
            End If

            Dim rightColumns As Integer() = Nothing
            If joinType = JoinTypes.Inner Then
                Dim ChildColumnsToAdd As New ArrayList 'Contains the column index values of the child table which shall be copied. These columns are regulary all columns (LEFT JOIN) except those required for the relation (for INNER JOIN; they would lead to duplicate columns)
                'Collect the data columns which shall be added from the client table to the parent table
                Dim LeftRelationColumns As DataColumn() = relation.ParentColumns
                Dim RelationColumns As DataColumn() = relation.ChildColumns
                For MyColCounter As Integer = 0 To rightChildTable.Columns.Count - 1
                    Dim IsRelationColumn As Boolean = False
                    'Verify that we don't add relation columns which would lead to duplicate data since those columns must also be in the parent table
                    For MyRelColCounter As Integer = 0 To RelationColumns.Length - 1
                        Dim rightColumn As DataColumn = rightChildTable.Columns(MyColCounter)
                        If rightColumn Is RelationColumns(MyRelColCounter) AndAlso LeftRelationColumns(MyRelColCounter).ColumnName = RelationColumns(MyRelColCounter).ColumnName Then 'if the name is equal then we don't need to add this column again, otherwise we do
                            IsRelationColumn = True
                        End If
                    Next
                    If IsRelationColumn = False Then
                        'Add the column index to the list of ToAdd-columns
                        ChildColumnsToAdd.Add(MyColCounter)
                    End If
                Next
                rightColumns = CType(ChildColumnsToAdd.ToArray(GetType(Integer)), Integer())
            End If

            'Return the results
            Return JoinTables(leftParentTable, Nothing, rightChildTable, rightColumns, relation, joinType)

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
        Friend Shared Function JoinTables(ByVal leftParentTable As DataTable, ByVal leftTableColumnsToCopy As DataColumn(),
                                          ByVal rightChildTable As DataTable, ByVal rightTableColumnsToCopy As DataColumn(), ByVal joinType As JoinTypes) As DataTable

            'Find the appropriate relation information
            Dim ActiveRelation As DataRelation = Nothing
            For MyRelCounter As Integer = 0 To leftParentTable.ChildRelations.Count - 1
                If leftParentTable.ChildRelations(MyRelCounter).ChildTable Is rightChildTable Then
                    ActiveRelation = leftParentTable.ChildRelations(MyRelCounter)
                    Exit For
                End If
            Next

            'Find required column indexes
            Dim LeftColumns As Integer() = Nothing
            Dim RightColumns As Integer() = Nothing
            If leftTableColumnsToCopy IsNot Nothing Then
                Dim indexesOfLeftTableColumnsToCopy As New ArrayList
                For MyCounter As Integer = 0 To leftTableColumnsToCopy.Length - 1
                    indexesOfLeftTableColumnsToCopy.Add(leftTableColumnsToCopy(MyCounter).Ordinal)
                Next
                LeftColumns = CType(indexesOfLeftTableColumnsToCopy.ToArray(GetType(Integer)), Integer())
            End If
            If rightTableColumnsToCopy IsNot Nothing Then
                Dim indexesOfRightTableColumnsToCopy As New ArrayList
                For MyCounter As Integer = 0 To rightTableColumnsToCopy.Length - 1
                    indexesOfRightTableColumnsToCopy.Add(rightTableColumnsToCopy(MyCounter).Ordinal)
                Next
                RightColumns = CType(indexesOfRightTableColumnsToCopy.ToArray(GetType(Integer)), Integer())
            End If

            'Return the results
            Return JoinTables(leftParentTable, LeftColumns, rightChildTable, RightColumns, ActiveRelation, joinType)

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
        Friend Shared Function JoinTables(ByVal leftParentTable As DataTable, ByVal indexesOfLeftTableColumnsToCopy As Integer(),
                                          ByVal rightChildTable As DataTable, ByVal indexesOfRightTableColumnsToCopy As Integer(),
                                          ByVal joinType As JoinTypes) As DataTable

            'Find the appropriate relation information
            Dim ActiveRelation As DataRelation = Nothing
            For MyRelCounter As Integer = 0 To leftParentTable.ChildRelations.Count - 1
                If leftParentTable.ChildRelations(MyRelCounter).ChildTable Is rightChildTable Then
                    ActiveRelation = leftParentTable.ChildRelations(MyRelCounter)
                    Exit For
                End If
            Next

            Return JoinTables(leftParentTable, indexesOfLeftTableColumnsToCopy, rightChildTable, indexesOfRightTableColumnsToCopy, ActiveRelation, joinType)

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
        Friend Shared Function JoinTables(ByVal leftParentTable As DataTable, ByVal indexesOfLeftTableColumnsToCopy As Integer(),
                                          ByVal rightChildTable As DataTable, ByVal indexesOfRightTableColumnsToCopy As Integer(),
                                          ByVal relation As DataRelation, ByVal joinType As JoinTypes) As DataTable

            'Verify parameters
            If leftParentTable Is Nothing OrElse rightChildTable Is Nothing Then
                Throw New Exception("One or both table references are null")
            End If
            If leftParentTable.DataSet Is Nothing OrElse leftParentTable.DataSet IsNot rightChildTable.DataSet Then
                Throw New Exception("Both tables must be member of the same dataset")
            End If
            If relation Is Nothing Then
                Throw New Exception("No relation defined between the two tables")
            End If

            'Prepare column wrap table
            Dim LeftTableColumnWraps As Integer()
            If indexesOfLeftTableColumnsToCopy Is Nothing Then
                Dim LeftColumnsToCopy As New ArrayList
                'Add all columns from left table
                For ColCounter As Integer = 0 To leftParentTable.Columns.Count - 1
                    LeftColumnsToCopy.Add(ColCounter)
                Next
                LeftTableColumnWraps = CType(LeftColumnsToCopy.ToArray(GetType(Integer)), Integer())
            Else
                'Add all columns as defined by indexesOfLeftTableColumnsToCopy
                Dim colWraps As New ArrayList
                For ColCounter As Integer = 0 To indexesOfLeftTableColumnsToCopy.Length - 1
                    Try
                        colWraps.Add(leftParentTable.Columns.Item(indexesOfLeftTableColumnsToCopy(ColCounter)).Ordinal)
                    Catch
                        Throw New Exception("Column index can't be found in source table's column collection: " & indexesOfLeftTableColumnsToCopy(ColCounter))
                    End Try
                Next
                LeftTableColumnWraps = CType(colWraps.ToArray(GetType(Integer)), Integer())
            End If

            'Prepare the result table by copying the parent table
            Dim Result As DataTable = leftParentTable.Clone
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
            If indexesOfRightTableColumnsToCopy Is Nothing Then
                Dim RightColumnsToCopy As New ArrayList
                For MyCounter As Integer = 0 To rightChildTable.Columns.Count - 1
                    RightColumnsToCopy.Add(MyCounter)
                Next
                RightTableColumnWraps = CType(RightColumnsToCopy.ToArray(GetType(Integer)), Integer())
            Else
                RightTableColumnWraps = CType(indexesOfRightTableColumnsToCopy.Clone, Integer())
            End If
            For MyCounter As Integer = 0 To RightTableColumnWraps.Length - 1
                Dim MyColumn As DataColumn = rightChildTable.Columns(RightTableColumnWraps(MyCounter))
                Dim UniqueColumnName As String = LookupUniqueColumnName(Result, MyColumn.ColumnName)
                Dim ColumnCaption As String = MyColumn.Caption
                Dim ColumnType As System.Type = MyColumn.DataType
                Result.Columns.Add(UniqueColumnName, ColumnType).Caption = ColumnCaption
            Next

            'Fill the rows now with the missing data
            For MyLeftTableRowCounter As Integer = 0 To leftParentTable.Rows.Count - 1
                Dim MyLeftRow As DataRow = leftParentTable.Rows(MyLeftTableRowCounter)
                Dim MyRightRows As DataRow() = MyLeftRow.GetChildRows(relation)

                If MyRightRows.Length = 0 Then
                    'Data only on left hand (parent) side
                    If joinType = JoinTypes.Left Then
                        Dim NewRow As DataRow = Result.NewRow
                        'Copy only data from parent table
                        For MyColCounter As Integer = 0 To LeftTableColumnWraps.Length - 1
                            NewRow(MyColCounter) = MyLeftRow(LeftTableColumnWraps(MyColCounter))
                        Next
                        'Add the new row, now
                        Result.Rows.Add(NewRow)
                    End If
                Else
                    'Data found on both sides
                    For RowInserts As Integer = 0 To MyRightRows.Length - 1
                        Dim NewRow As DataRow = Result.NewRow
                        'Copy data from parent table row
                        For MyColCounter As Integer = 0 To LeftTableColumnWraps.Length - 1
                            NewRow(MyColCounter) = MyLeftRow(LeftTableColumnWraps(MyColCounter))
                        Next
                        'Copy data from this child row
                        Dim MyRightRowsChild As DataRow = MyRightRows(RowInserts)
                        For MyColCounter As Integer = 0 To RightTableColumnWraps.Length - 1
                            NewRow(LeftTableColumnWraps.Length + MyColCounter) = MyRightRowsChild(RightTableColumnWraps(MyColCounter))
                        Next
                        'Add the new row, now
                        Result.Rows.Add(NewRow)
                    Next
                End If

            Next

            Return Result

        End Function

        ''' <summary>
        '''     Cross join of two tables
        ''' </summary>
        ''' <param name="leftTable">A first datatable</param>
        ''' <param name="indexesOfLeftTableColumnsToCopy">An array of column indexes to copy from the left table</param>
        ''' <param name="rightTable">A second datatable</param>
        ''' <param name="indexesOfRightTableColumnsToCopy">An array of column indexes to copy from the right table</param>
        ''' <returns></returns>
        Friend Shared Function CrossJoinTables(ByVal leftTable As DataTable, ByVal indexesOfLeftTableColumnsToCopy As Integer(),
                                               ByVal rightTable As DataTable, ByVal indexesOfRightTableColumnsToCopy As Integer()) As DataTable
            'TODO: verify/fix exceptions when left AND right table contain rows not matching to the other side (FULL OUTER JOIN situations)
            'TODO: above ToDo "verify/fix exceptions" might not be applicable here?!? --> to remove ?!?

            'Verify parameters
            If leftTable Is Nothing OrElse rightTable Is Nothing Then
                Throw New Exception("One or both table references are null")
            End If

            'Prepare column wrap table
            Dim LeftTableColumnWraps As Integer()
            If indexesOfLeftTableColumnsToCopy Is Nothing Then
                Dim LeftColumnsToCopy As New ArrayList
                'Add all columns from left table
                For ColCounter As Integer = 0 To leftTable.Columns.Count - 1
                    LeftColumnsToCopy.Add(ColCounter)
                Next
                LeftTableColumnWraps = CType(LeftColumnsToCopy.ToArray(GetType(Integer)), Integer())
            Else
                'Add all columns as defined by indexesOfLeftTableColumnsToCopy
                Dim colWraps As New ArrayList
                For ColCounter As Integer = 0 To indexesOfLeftTableColumnsToCopy.Length - 1
                    Try
                        colWraps.Add(leftTable.Columns.Item(indexesOfLeftTableColumnsToCopy(ColCounter)).Ordinal)
                    Catch
                        Throw New Exception("Column index can't be found in source table's column collection: " & indexesOfLeftTableColumnsToCopy(ColCounter))
                    End Try
                Next
                LeftTableColumnWraps = CType(colWraps.ToArray(GetType(Integer)), Integer())
            End If

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
                If KeepThisColumn = False Then
                    If Result.Columns(MyCounter).Unique = True Then
                        Result.Columns(MyCounter).Unique = False
                    End If
                    Result.Columns.Remove(Result.Columns(MyCounter))
                End If
            Next

            'Add the right columns
            Dim RightTableColumnWraps As Integer()
            If indexesOfRightTableColumnsToCopy Is Nothing Then
                Dim RightColumnsToCopy As New ArrayList
                For MyCounter As Integer = 0 To rightTable.Columns.Count - 1
                    RightColumnsToCopy.Add(MyCounter)
                Next
                RightTableColumnWraps = CType(RightColumnsToCopy.ToArray(GetType(Integer)), Integer())
            Else
                RightTableColumnWraps = CType(indexesOfRightTableColumnsToCopy.Clone, Integer())
            End If
            For MyCounter As Integer = 0 To RightTableColumnWraps.Length - 1
                Dim MyColumn As DataColumn = rightTable.Columns(RightTableColumnWraps(MyCounter))
                Dim UniqueColumnName As String = LookupUniqueColumnName(Result, MyColumn.ColumnName)
                Dim ColumnCaption As String = MyColumn.Caption
                Dim ColumnType As System.Type = MyColumn.DataType
                Result.Columns.Add(UniqueColumnName, ColumnType).Caption = ColumnCaption
            Next

            'Fill the rows now with the missing data
            For MyLeftTableRowCounter As Integer = 0 To leftTable.Rows.Count - 1
                For MyRightTableRowCounter As Integer = 0 To rightTable.Rows.Count - 1
                    Dim MyLeftRow As DataRow = leftTable.Rows(MyLeftTableRowCounter)
                    Dim MyRightRow As DataRow = rightTable.Rows(MyRightTableRowCounter)

                    'Data found on both sides
                    Dim NewRow As DataRow = Result.NewRow

                    'Copy data from parent table row
                    For MyColCounter As Integer = 0 To LeftTableColumnWraps.Length - 1
                        NewRow(MyColCounter) = MyLeftRow(LeftTableColumnWraps(MyColCounter))
                    Next

                    'Copy data from this child row
                    For MyColCounter As Integer = 0 To RightTableColumnWraps.Length - 1
                        NewRow(LeftTableColumnWraps.Length + MyColCounter) = MyRightRow(RightTableColumnWraps(MyColCounter))
                    Next

                    'Add the new row, now
                    Result.Rows.Add(NewRow)

                Next
            Next

            Return Result

        End Function

        ''' <summary>
        '''     Add a prefix to the names of the columns
        ''' </summary>
        ''' <param name="dataTable">A datatable which shall be exported</param>
        ''' <param name="columnIndexes">An array of column indexes</param>
        ''' <param name="prefix">e. g. "orders."</param>
        Friend Shared Sub AddPrefixesToColumnNames(ByVal dataTable As DataTable, ByVal columnIndexes As Integer(), ByVal prefix As String)

            'all columns if nothing is given
            If columnIndexes Is Nothing Then
                ReDim columnIndexes(dataTable.Columns.Count - 1)
                For MyCounter As Integer = 0 To dataTable.Columns.Count - 1
                    columnIndexes(MyCounter) = MyCounter
                Next
            End If

            For MyCounter As Integer = 0 To columnIndexes.Length - 1
                dataTable.Columns(columnIndexes(MyCounter)).ColumnName = prefix & dataTable.Columns(columnIndexes(MyCounter)).ColumnName
            Next

        End Sub

        ''' <summary>
        '''     Add a suffix to the names of the columns
        ''' </summary>
        ''' <param name="dataTable">A datatable which shall be exported</param>
        ''' <param name="columnIndexes">An array of column indexes</param>
        ''' <param name="suffix">e. g. "-orders"</param>
        Friend Shared Sub AddSuffixesToColumnNames(ByVal dataTable As DataTable, ByVal columnIndexes As Integer(), ByVal suffix As String)

            'all columns if nothing is given
            If columnIndexes Is Nothing Then
                ReDim columnIndexes(dataTable.Columns.Count - 1)
                For MyCounter As Integer = 0 To dataTable.Columns.Count - 1
                    columnIndexes(MyCounter) = MyCounter
                Next
            End If

            For MyCounter As Integer = 0 To columnIndexes.Length - 1
#Disable Warning S1643 ' Strings should not be concatenated using "+" or "&" in a loop
                dataTable.Columns(columnIndexes(MyCounter)).ColumnName = dataTable.Columns(columnIndexes(MyCounter)).ColumnName & suffix
#Enable Warning S1643 ' Strings should not be concatenated using "+" or "&" in a loop
            Next

        End Sub

        ''' <summary>
        '''     Lookup a new unique column name for a data table
        ''' </summary>
        ''' <param name="dataTable">The data table which shall get a new data column</param>
        ''' <param name="suggestedColumnName">A column name suggestion</param>
        ''' <returns>The suggested column name as it is or modified column name to be unique</returns>
        Friend Shared Function LookupUniqueColumnName(ByVal dataTable As DataTable, ByVal suggestedColumnName As String) As String

            Dim ColumnNameAlreadyExistant As Boolean = False
            For MyCounter As Integer = 0 To dataTable.Columns.Count - 1
                If String.Compare(suggestedColumnName, dataTable.Columns(MyCounter).ColumnName, True) = 0 Then
                    ColumnNameAlreadyExistant = True
                End If
            Next

            If ColumnNameAlreadyExistant = False Then
                'Exit function
                Return suggestedColumnName
            Else
                'Add prefix "ClientTable_" or add/increase a counter at the end
                If suggestedColumnName.StartsWith("ClientTable_") Then
                    'Find the position range of an already existing counter at the end of the string - if there is a number
                    Dim NumberPositionIndex As Integer = -1
                    For NumberPartCounter As Integer = suggestedColumnName.Length - 1 To 0 Step -1
                        If Char.IsNumber(suggestedColumnName.Chars(NumberPartCounter)) = False Then
                            NumberPositionIndex = NumberPartCounter + 1 'Next char behind the current char
                            Exit For
                        End If
                    Next
                    'Read out the value of the counter
                    Dim NumberCounterValue As Integer
                    If NumberPositionIndex = -1 OrElse NumberPositionIndex + 1 > suggestedColumnName.Length Then
                        'Attach a new counter value
                        NumberCounterValue = 1
                        suggestedColumnName &= NumberCounterValue.ToString
                    Else
                        'Update the counter value
                        NumberCounterValue = CType(suggestedColumnName.Substring(NumberPositionIndex), Integer) + 1
                        suggestedColumnName = suggestedColumnName.Substring(0, NumberPositionIndex) & NumberCounterValue.ToString
                    End If
                Else
                    'Add new prefix
                    suggestedColumnName = "ClientTable_" & suggestedColumnName
                End If
                'Revalidate uniqueness by running recursively
                suggestedColumnName = LookupUniqueColumnName(dataTable, suggestedColumnName)
            End If

            Return suggestedColumnName

        End Function

        ''' <summary>
        '''     Rearrange columns
        ''' </summary>
        ''' <param name="source">The source table with data</param>
        ''' <param name="columnsToCopy">An array of column names which shall be copied in the specified order from the source table</param>
        ''' <returns>A new and independent data table with copied data</returns>
        Friend Shared Function ReArrangeDataColumns(ByVal source As DataTable, ByVal columnsToCopy As String()) As DataTable
            Dim columns As New ArrayList
            For MyCounter As Integer = 0 To columnsToCopy.Length - 1
                columns.Add(New DataColumn(columnsToCopy(MyCounter), source.Columns(columnsToCopy(MyCounter)).DataType))
            Next
            Return ReArrangeDataColumns(source, CType(columns.ToArray(GetType(System.Data.DataColumn)), System.Data.DataColumn()))
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
            Return ReArrangeDataColumns(source, destinationColumnSet, Nothing)
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

            'Parameter validation
            If source Is Nothing Then
                Throw New ArgumentNullException(NameOf(source))
            End If
            If destinationColumnSet Is Nothing Then
                Throw New ArgumentNullException(NameOf(destinationColumnSet))
            ElseIf destinationColumnSet.Length = 0 Then
                Throw New ArgumentException("empty array not allowed", NameOf(destinationColumnSet))
            End If

            'Prepare new datatable
            Dim Result As New DataTable(source.TableName)
            For MyCounter As Integer = 0 To destinationColumnSet.Length - 1
                Result.Columns.Add(destinationColumnSet(MyCounter))
            Next

            'Prepare column wrap table
            Dim colWraps As New ArrayList
            For ColCounter As Integer = 0 To destinationColumnSet.Length - 1
                Try
                    colWraps.Add(source.Columns.Item(destinationColumnSet(ColCounter).ColumnName).Ordinal)
                Catch
                    Throw New Exception("Column name can't be found in source table's column collection: " & destinationColumnSet(ColCounter).ColumnName)
                End Try
            Next
            Dim ColumnWraps As Integer() = CType(colWraps.ToArray(GetType(Integer)), Integer())

            'Copy content
            For RowCounter As Integer = 0 To source.Rows.Count - 1

                Dim MyNewRow As DataRow = Result.NewRow

                For ColCounter As Integer = 0 To destinationColumnSet.Length - 1
                    Dim sourceData As Object = source.Rows(RowCounter)(ColumnWraps(ColCounter))
                    If sourceData Is Nothing OrElse IsDBNull(sourceData) = True Then
                        MyNewRow(ColCounter) = sourceData
                    Else
                        Try
                            MyNewRow(ColCounter) = sourceData
                        Catch ex As Exception
                            Dim conversionException As ReArrangeDataColumnsException
                            conversionException = New ReArrangeDataColumnsException(RowCounter, ColCounter, source.Columns(ColumnWraps(ColCounter)).DataType, Result.Columns(ColCounter).DataType, source.Rows(RowCounter)(ColumnWraps(ColCounter)), ex)
                            If ignoreConversionExceptionAndLogThemHere Is Nothing Then
                                Throw conversionException
                            Else
                                ignoreConversionExceptionAndLogThemHere.Add(conversionException)
                            End If
                        End Try
                    End If
                Next

                Result.Rows.Add(MyNewRow)

            Next

            'Return new datatable
            Return Result

        End Function

    End Class

End Namespace