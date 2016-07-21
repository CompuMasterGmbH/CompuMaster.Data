Imports NUnit.Framework

Namespace CompuMaster.Test.Data

    <TestFixture(Category:="DataTables")> Public Class DataTables

#Region "Test data"
        Private Function _TestTable1() As DataTable
            Dim Result As New DataTable("test1")
            Result.Columns.Add("ID", GetType(Integer))
            Result.Columns.Add("Value", GetType(String))
            Dim newRow As DataRow
            newRow = Result.NewRow
            newRow(0) = 1
            newRow(1) = "Hello world!"
            Result.Rows.Add(newRow)
            newRow = Result.NewRow
            newRow(0) = 2
            newRow(1) = "Gotcha!"
            Result.Rows.Add(newRow)
            newRow = Result.NewRow
            newRow(0) = 3
            newRow(1) = "Hello world!"
            Result.Rows.Add(newRow)
            newRow = Result.NewRow
            newRow(0) = 4
            newRow(1) = "Not a duplicate"
            Result.Rows.Add(newRow)
            newRow = Result.NewRow
            newRow(0) = 5
            newRow(1) = "Hello world!"
            Result.Rows.Add(newRow)
            newRow = Result.NewRow
            newRow(0) = 6
            newRow(1) = "GOTCHA!"
            Result.Rows.Add(newRow)
            newRow = Result.NewRow
            newRow(0) = 7
            newRow(1) = "Gotcha!"
            Result.Rows.Add(newRow)
            Return Result
        End Function

        Private Function _TestTable2() As DataTable
            Dim file As String = AssemblyTestEnvironment.TestFileAbsolutePath("testfiles\Q&A.xls")
            Dim dt As DataTable = CompuMaster.Data.XlsReader.ReadDataTableFromXlsFile(file, "Rund um das NT")
            Return dt
        End Function
#End Region

        <Test()> Public Sub AddPrefixesToColumnNames()
            Dim dt As New DataTable
            dt.Columns.Add("hello")
            CompuMaster.Data.DataTables.AddPrefixesToColumnNames(dt, New Integer() {0}, "pref_")
            StringAssert.StartsWith("pref_", dt.Columns.Item(0).ColumnName)
        End Sub

        <Test()> Public Sub AddSufixesToColumnNames()
            Dim dt As New DataTable
            dt.Columns.Add("hello")
            CompuMaster.Data.DataTables.AddSuffixesToColumnNames(dt, New Integer() {0}, "_suf")
            StringAssert.EndsWith("_suf", dt.Columns.Item(0).ColumnName)
        End Sub

        <Test()> Public Sub ColumnIndex()
            Dim dt As New DataTable
            dt.Columns.Add("hey")
            Dim e As DataRow = dt.NewRow
            e.Item(0) = "D"
            dt.Rows.Add(e)
            Assert.AreEqual(0, CompuMaster.Data.DataTables.ColumnIndex(dt.Columns("hey")))
        End Sub
        <Test()> Public Sub CompareValuesOfUnknownType()
            Dim dt As New DataTable
            dt.Columns.Add("DummmyColumn")

            Dim e As Integer = 23
            Dim f As Integer = 24
            Assert.IsTrue(CompuMaster.Data.DataTables.CompareValuesOfUnknownType("D", "D", False))
            Assert.IsFalse(CompuMaster.Data.DataTables.CompareValuesOfUnknownType("D", "d", False))
            Assert.IsFalse(CompuMaster.Data.DataTables.CompareValuesOfUnknownType(e, f, False))
        End Sub

        <Test()> Public Sub ConvertColumnType()
            Dim dt As New DataTable
            dt.Columns.Add("Amount", GetType(Integer))
            ' CompuMaster.Data.DataTables.ConvertColumnType(dt.Columns.Item(0), GetType(Boolean),
        End Sub

        <Test()> Public Sub ConvertColumnValuesIntoArrayList()
            Dim dt As New DataTable
            dt.Columns.Add("DummyColumn")

            Dim row1 As DataRow = dt.NewRow
            row1.Item(0) = 23
            dt.Rows.Add(row1)

            Dim row2 As DataRow = dt.NewRow
            row2.Item(0) = 21
            dt.Rows.Add(row2)

            Dim row3 As DataRow = dt.NewRow
            row3.Item(0) = 23
            dt.Rows.Add(row3)

            Dim row4 As DataRow = dt.NewRow
            row4.Item(0) = 7
            dt.Rows.Add(row4)


            Dim list As ArrayList = CompuMaster.Data.DataTables.ConvertColumnValuesIntoArrayList(dt.Columns.Item(0))
            Assert.AreEqual(4, list.Count)
        End Sub

        <Test()> Public Sub ConvertDataReaderToDataSet()
            Dim MyConn As New System.Data.SqlClient.SqlConnection("SERVER=sql2012;DATABASE=master;PWD=xxxxxxxxxxxxxxxxxxx;UID=sa")
            Dim MyCmd As IDbCommand = MyConn.CreateCommand()
            MyCmd.CommandType = CommandType.Text
            MyCmd.CommandText = "exec sp_databases; Exec sp_tables;"
            Dim Reader As System.Data.IDataReader = CompuMaster.Data.DataQuery.AnyIDataProvider.ExecuteReader(MyCmd, CompuMaster.Data.DataQuery.Automations.AutoOpenAndCloseAndDisposeConnection)
            Dim Data As DataSet = CompuMaster.Data.DataTables.ConvertDataReaderToDataSet(Reader)
            Assert.AreEqual(2, Data.Tables.Count)
        End Sub

        <Test()> Public Sub ConvertDataReaderToDataTable()
            Dim TestFile As String = AssemblyTestEnvironment.TestFileAbsolutePath("testfiles\test_for_msaccess.mdb")
            Dim MyConn As IDbConnection = CompuMaster.Data.DataQuery.Connections.MicrosoftAccessConnection(TestFile)
            Dim MyCmd As IDbCommand = MyConn.CreateCommand()
            MyCmd.CommandType = CommandType.Text
            MyCmd.CommandText = "SELECT IntegerLongValue, StringShort, StringMemo FROM [SeveralColumnTypesTest] ORDER BY ID"
            Dim Reader As System.Data.IDataReader = CompuMaster.Data.DataQuery.AnyIDataProvider.ExecuteReader(MyCmd, CompuMaster.Data.DataQuery.Automations.AutoOpenAndCloseAndDisposeConnection)
            Dim Data As DataTable = CompuMaster.Data.DataTables.ConvertDataReaderToDataTable(Reader, "mytablename")
            Assert.AreEqual("mytablename", Data.TableName)
            Assert.AreNotEqual(0, Data.Rows.Count)
        End Sub

        <Test(), NUnit.Framework.Ignore("NotYetImplemented")> Public Sub ConvertDatasetToXml()
            Throw New NotImplementedException
        End Sub

        <Test()> Public Sub ConvertDataTableToArrayList()
            Dim dt As New DataTable
            dt.Columns.Add("IntegerColumn", GetType(Integer))
            Dim rRow As DataRow = dt.NewRow
            Dim row2 As DataRow = dt.NewRow
            rRow.Item(0) = 23
            row2.Item(0) = 25

            dt.Rows.Add(rRow)
            dt.Rows.Add(row2)

            Dim e As ArrayList = CompuMaster.Data.DataTables.ConvertDataTableToArrayList(dt.Columns.Item(0))
            Assert.AreEqual(2, e.Count)
            Assert.AreEqual(23, e.Item(0))
            Assert.AreEqual(25, e.Item(1))
        End Sub

        <Test()> Public Sub ConvertDataTableToDictionaryEntryArray()
            Dim dt As New DataTable
            dt.Columns.Add("Key", GetType(String))
            dt.Columns.Add("Value", GetType(Integer))
            Dim rRow As DataRow = dt.NewRow
            rRow.Item(0) = "Number"
            rRow.Item(1) = 25
            dt.Rows.Add(rRow)
            Dim e As DictionaryEntry() = CompuMaster.Data.DataTables.ConvertDataTableToDictionaryEntryArray(dt)
            Assert.AreEqual(1, e.GetLength(0))
            StringAssert.IsMatch("Number", dt.Rows.Item(0).Item(0))
            Assert.AreEqual(25, dt.Rows.Item(0).Item(1))
        End Sub

        <Test()> Public Sub ConvertDataTableToHashtable()
            Dim dt As New DataTable
            dt.Columns.Add("Key")
            dt.Columns.Add("Value")
            dt.Columns(0).Unique = True

            Dim rRow As DataRow = dt.NewRow
            rRow.Item(0) = "a"
            rRow.Item(1) = "z"
            dt.Rows.Add(rRow)
            Dim ht As Hashtable = CompuMaster.Data.DataTables.ConvertDataTableToHashtable(dt)
            Assert.IsTrue(ht.ContainsKey("a"))
            Assert.IsTrue(ht.ContainsValue("z"))
        End Sub

        <Test(), NUnit.Framework.Ignore("NotYetImplemented")> Public Sub ConvertDataTableToWebFormsListItem()
            Dim dt As New DataTable
            dt.Columns.Add("First")
            dt.Columns.Add("second")

            Dim rRow As DataRow = dt.NewRow
            rRow.Item(0) = "FirstEntry"
            rRow.Item(1) = "SecondValue"
            dt.Rows.Add(rRow)
            '        CompuMaster.Data.DataTables.ConvertDataTableToWebFormsListItem(
            Throw New NotImplementedException
        End Sub

        <Test()> Public Sub ConvertDictionaryEntryArrayToDataTable()
            Dim de As DictionaryEntry()
            ReDim de(1)

            de(0).Key = "Hello"
            de(0).Value = "Bye"

            de(1).Key = "Fire"
            de(1).Value = "Water"

            Dim dt As DataTable = CompuMaster.Data.DataTables.ConvertDictionaryEntryArrayToDataTable(de)
            Assert.AreEqual(2, dt.Columns.Count())
            Assert.AreEqual(2, dt.Rows.Count())
            StringAssert.IsMatch("Hello", dt.Rows.Item(0).Item(0))
            StringAssert.IsMatch("Fire", dt.Rows.Item(1).Item(0))
            StringAssert.IsMatch("Bye", dt.Rows.Item(0).Item(1))
            StringAssert.IsMatch("Water", dt.Rows.Item(1).Item(1))
        End Sub

        <Test(), NUnit.Framework.Ignore("NotYetImplemented")> Public Sub ConvertICollectionToDataTable()
            Throw New NotImplementedException
        End Sub

#If NET_1_1 = False Then
        <Test()> Public Sub ConvertIDictionaryToDataTable()
            Dim dict As IDictionary = New System.Collections.Generic.Dictionary(Of String, String)()
            dict.Add("Berlin", "Germany")

            Dim dt As DataTable = CompuMaster.Data.DataTables.ConvertIDictionaryToDataTable(dict)
            Assert.AreEqual(1, dt.Rows.Count())
            Assert.AreEqual(2, dt.Columns.Count())

        End Sub
#End If

        <Test()> Public Sub ConvertNameValueCollectionToDataTable()
            Dim nvc As New System.Collections.Specialized.NameValueCollection
            nvc.Add("Berlin", "Germany")
            nvc.Add("Paris", "France")
            Dim dt As DataTable = CompuMaster.Data.DataTables.ConvertNameValueCollectionToDataTable(nvc)
            Assert.AreEqual(2, dt.Columns.Count())
            Assert.AreEqual(2, dt.Rows.Count())
        End Sub

        <Test()> Public Sub ConvertToHtmlTable()
            Dim dt As New DataTable
            dt.Columns.Add("id", GetType(Integer))
            dt.Columns.Add("Hi", GetType(String))
            Dim row As DataRow = dt.NewRow
            row.Item(0) = 23
            row.Item(1) = "Hello World"

            Dim html As String = CompuMaster.Data.DataTables.ConvertToHtmlTable(dt)

            'TODO: ...'

        End Sub

        <Test()> Public Sub ConvertToPlainTextTable()
            Dim dt As New DataTable
            dt.Columns.Add("id", GetType(Integer))
            dt.Columns.Add("Hi", GetType(String))

            Dim row As DataRow = dt.NewRow

            row.Item(0) = 23
            row.Item(1) = "Hello World"
            dt.Rows.Add(row)


            Dim html As String = CompuMaster.Data.DataTables.ConvertToPlainTextTable(dt)

            Dim dt2 As New DataTable
            dt2.Columns.Add("id", GetType(Integer))
            dt2.Columns.Add("Hi", GetType(String))

            Dim row2 As DataRow = dt2.NewRow
            row.item(0) = 23
            row.Item(1) = "Hello"

            Dim row3 As DataRow = dt2.NewRow
            row.Item(0) = 21
            row.Item(1) = "hello"

            dt2.Rows.Add(row2)
            dt2.Rows.Add(row3)
            dt2.AcceptChanges()
            dt2.Rows(1).Delete()
            Dim html2 As String = CompuMaster.Data.DataTables.ConvertToPlainTextTable(dt2)

        End Sub

        <Test()> Public Sub ConvertToWikiTable()
            Dim dt As New DataTable
            dt.Columns.Add("id", GetType(Integer))
            dt.Columns.Add("Hi", GetType(String))

            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToWikiTable(dt))

            Dim row As DataRow = dt.NewRow
            row.Item(0) = 23
            row.Item(1) = "Hello World"
            dt.Rows.Add(row)

            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToWikiTable(dt))

            dt = _TestTable2()
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToWikiTable(dt))

        End Sub

        <Test()> Public Sub ConvertToPlainTextTableFixedColumnWidths()
            Dim dt As New DataTable
            dt.Columns.Add("id", GetType(Integer))
            dt.Columns.Add("Hi", GetType(String))

            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt, 10))
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt, 5, 20))
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt, " :: ", " :: ", "=##=", "=", "="))

            Dim row As DataRow = dt.NewRow
            row.Item(0) = 23
            row.Item(1) = "Hello World"
            dt.Rows.Add(row)

            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt, " :: ", " :: ", "=##=", "=", "="))
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt, 10))
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt, 5, 20))

            dt = _TestTable2()
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt))
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt, " :: ", " :: ", "=##=", "=", "="))
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt, 10))
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(dt, 5, 20))

        End Sub


        <Test(), NUnit.Framework.Ignore("NotYetImplemented")> Public Sub ConvertXmlToDataset()
            Throw New NotImplementedException
        End Sub


        <Test()> Public Sub CopyDataTableWithSubsetOfRows()
            Dim dt As New DataTable
            dt.Columns.Add("hi")
            dt.Columns.Add("hi2")
            dt.Rows.Add(New String() {"hi", "d"})
            dt.Rows.Add(New String() {"hix", "dix"})
            dt.Rows.Add(New String() {"l", "d"})

            Dim dt2 As DataTable = CompuMaster.Data.DataTables.CopyDataTableWithSubsetOfRows(dt, 2)
            Assert.AreEqual(2, dt2.Rows.Count())
            StringAssert.IsMatch("hi", dt.Rows.Item(0).Item(0))
            StringAssert.IsMatch("hix", dt.Rows.Item(1).Item(0))

            dt2 = CompuMaster.Data.DataTables.CopyDataTableWithSubsetOfRows(dt, 1, 2)
            StringAssert.IsMatch("l", dt2.Rows.Item(1).Item(0))

        End Sub

        <Test()> Public Sub CreateDataRowClone()
            Dim dt As New DataTable
            dt.Columns.Add("id", GetType(Integer))
            dt.Columns.Add("text", GetType(String))

            Dim row As DataRow = dt.NewRow
            row.Item(0) = 9
            row.Item(1) = "test"

            dt.Rows.Add(row)

            Dim row2 As DataRow
            row2 = CompuMaster.Data.DataTables.CreateDataRowClone(dt.Rows(0))

            Assert.AreEqual(9, row2.Item(0))
            StringAssert.IsMatch("test", row2.Item(1))
        End Sub

        <Test()> Public Sub CreateDataTableClone()
            Dim dt As New DataTable
            Dim dt2 As DataTable


            dt.Columns.Add("id", GetType(Integer))
            dt.Columns.Add("hi", GetType(String))

            Dim row As DataRow = dt.NewRow

            row.Item(0) = 7
            row.Item(1) = "Test"

            dt.Rows.Add(row)

            dt2 = CompuMaster.Data.DataTables.CreateDataTableClone(dt)
            Assert.AreEqual(1, dt2.Rows.Count())
            Assert.AreEqual(2, dt2.Columns.Count())

            Assert.AreEqual(7, dt2.Rows.Item(0).Item(0))
            StringAssert.IsMatch("Test", dt2.Rows.Item(0).Item(1))

            dt.Rows.Add(New Object() {8, "TestL2"})


            dt2 = CompuMaster.Data.DataTables.CreateDataTableClone(dt, "hi = 'TestL2'")
            Assert.AreEqual(1, dt2.Rows.Count())

            dt2 = CompuMaster.Data.DataTables.CreateDataTableClone(dt, Nothing, "id DESC")
            Assert.AreEqual(8, dt2.Rows.Item(0).Item(0))


            dt.Rows.Add(New Object() {9, "L"})
            dt.Rows.Add(New Object() {10, "JJJ"})
            dt.Rows.Add(New Object() {11, "Lcx"})

            dt2 = CompuMaster.Data.DataTables.CreateDataTableClone(dt, Nothing, Nothing, 3)
            Assert.AreEqual(3, dt2.Rows.Count())

            Dim first As New DataTable
            Dim second As New DataTable

            first.Columns.Add("id")
            first.Columns.Add("whatever")
            first.Columns.Add("thirdCol")
            first.PrimaryKey = New DataColumn() {first.Columns(0)}
            first.Rows.Add(New Object() {23, "L", "X"})
            first.Rows.Add(New Object() {21, "p", "a"})
            first.Rows.Add(New Object() {1, "q", "z"})
            CompuMaster.Data.DataTables.CreateDataTableClone(first, second, Nothing, Nothing, Nothing, CompuMaster.Data.DataTables.RequestedRowChanges.DropExistingRowsInDestinationTableAndInsertNewRows, False,
                                                             CompuMaster.Data.DataTables.RequestedSchemaChangesForUnusedColumns.None, CompuMaster.Data.DataTables.RequestedSchemaChangesForExistingColumns.None,
                                                             CompuMaster.Data.DataTables.RequestedSchemaChangesForAdditionalColumns.Add)

            '  Assert.AreEqual(1, second.PrimaryKey.GetLength(0))




            'Testing the merge function
            Dim merge_source As New DataTable
            Dim merge_dest As New DataTable

            Dim oldSource As DataTable
            Dim oldDest As DataTable

            merge_source.Columns.Add("A", GetType(String))
            merge_source.Columns.Add("B", GetType(String))
            merge_source.Columns.Add("C", GetType(String))

            merge_dest.Columns.Add("A", GetType(String))
            merge_dest.Columns.Add("B", GetType(String))




            merge_source.PrimaryKey = New DataColumn() {merge_source.Columns(0)}
            merge_dest.PrimaryKey = New DataColumn() {merge_dest.Columns(0)}


            merge_dest.Rows.Add(New String() {"23", "hello"})
            merge_source.Rows.Add(New String() {"23", "hello!", "missing"})

            oldSource = merge_source
            oldDest = merge_dest

            CompuMaster.Data.DataTables.CreateDataTableClone(merge_source, merge_dest, Nothing, Nothing, Nothing, CompuMaster.Data.DataTables.RequestedRowChanges.KeepExistingRowsInDestinationTableAndAddRemoveUpdateChangedRows,
                                                             True, CompuMaster.Data.DataTables.RequestedSchemaChangesForUnusedColumns.None, CompuMaster.Data.DataTables.RequestedSchemaChangesForExistingColumns.None,
                                                             CompuMaster.Data.DataTables.RequestedSchemaChangesForAdditionalColumns.None)
            Assert.AreEqual(1, merge_dest.Rows.Count())
            Assert.AreEqual(2, merge_dest.Columns.Count())
            StringAssert.IsMatch("hello!", merge_dest.Rows(0).Item(1))

            merge_source = oldSource
            merge_dest = oldDest

            CompuMaster.Data.DataTables.CreateDataTableClone(merge_source, merge_dest, Nothing, Nothing, Nothing, CompuMaster.Data.DataTables.RequestedRowChanges.KeepExistingRowsInDestinationTableAndAddRemoveUpdateChangedRows,
                                                             True, CompuMaster.Data.DataTables.RequestedSchemaChangesForUnusedColumns.None, CompuMaster.Data.DataTables.RequestedSchemaChangesForExistingColumns.None,
                                                             CompuMaster.Data.DataTables.RequestedSchemaChangesForAdditionalColumns.Add)
            Assert.AreEqual(3, merge_dest.Columns.Count())
            Assert.AreEqual(1, merge_dest.Rows.Count())
            StringAssert.IsMatch("hello!", merge_dest.Rows(0).Item(1))
            StringAssert.IsMatch("missing", merge_dest.Rows(0).Item(2))


            Dim merge_source2 As New DataTable
            Dim merge_dest2 As New DataTable


            merge_source2.Columns.Add("A", GetType(String))
            merge_source2.Columns.Add("B", GetType(String))
            merge_source2.Columns.Add("C", GetType(String))

            merge_dest2.Columns.Add("A", GetType(String))
            merge_dest2.Columns.Add("B", GetType(String))




            merge_source2.PrimaryKey = New DataColumn() {merge_source2.Columns(0)}
            merge_dest2.PrimaryKey = New DataColumn() {merge_dest2.Columns(0)}


            merge_dest2.Rows.Add(New String() {"23", "hello"})
            merge_source2.Rows.Add(New String() {"23", "hello!", "missing"})

            CompuMaster.Data.DataTables.CreateDataTableClone(merge_source2, merge_dest2, Nothing, Nothing, Nothing, CompuMaster.Data.DataTables.RequestedRowChanges.KeepExistingRowsInDestinationTableAndAddRemoveUpdateChangedRows,
                                                             True, CompuMaster.Data.DataTables.RequestedSchemaChangesForUnusedColumns.None,
                                                             CompuMaster.Data.DataTables.RequestedSchemaChangesForExistingColumns.Update, CompuMaster.Data.DataTables.RequestedSchemaChangesForAdditionalColumns.None)
            StringAssert.IsMatch("hello!", merge_dest2.Rows(0).Item(1))
            Assert.AreEqual(2, merge_dest2.Columns.Count)


            Dim merge_source3 As New DataTable
            Dim merge_dest3 As New DataTable

            merge_source3.Columns.Add("A", GetType(Integer))
            merge_source3.Columns.Add("B", GetType(String))
            merge_source3.Columns.Add("C", GetType(String))

            merge_dest3.Columns.Add("A", GetType(Integer))
            merge_dest3.Columns.Add("B", GetType(String))
            merge_dest3.Columns.Add("C", GetType(String))

            merge_source3.PrimaryKey = New DataColumn() {merge_source3.Columns(0)}
            merge_dest3.PrimaryKey = New DataColumn() {merge_dest3.Columns(0)}

            merge_source3.Rows.Add(New Object() {1, "Text1", "A"})
            merge_source3.Rows.Add(New Object() {2, "Text2", "A"})
            merge_source3.Rows.Add(New Object() {3, "Text3", "A"})
            merge_source3.Rows.Add(New Object() {4, "Text4", "B"})
            merge_source3.Rows.Add(New Object() {5, "Text5", "B"})

            merge_dest3.Rows.Add(New Object() {1, "TextGone!", "A"})
            merge_dest3.Rows.Add(New Object() {9, "Not touched", "B"})
            merge_dest3.Rows.Add(New Object() {10, "...", "B"})
            merge_dest3.Rows.Add(New Object() {2, "TextX", "A"})
            merge_dest3.Rows.Add(New Object() {3, "..ax", "A"})

            merge_source3.AcceptChanges()
            merge_dest3.AcceptChanges()

            CompuMaster.Data.DataTables.CreateDataTableClone(merge_source3, merge_dest3, "C = 'A'", "A ASC", Nothing, CompuMaster.Data.DataTables.RequestedRowChanges.KeepExistingRowsInDestinationTableAndAddRemoveUpdateChangedRows,
                                                             True, CompuMaster.Data.DataTables.RequestedSchemaChangesForUnusedColumns.None,
                                                             CompuMaster.Data.DataTables.RequestedSchemaChangesForExistingColumns.None,
                                                             CompuMaster.Data.DataTables.RequestedSchemaChangesForAdditionalColumns.None)


            Assert.AreEqual(5, merge_dest3.Rows.Count)
            Assert.AreEqual(3, merge_dest3.Columns.Count)


            'Checking sort after merge (only source 
            Assert.AreEqual(1, merge_dest3.Rows.Item(0).Item(0))
            Assert.AreEqual(9, merge_dest3.Rows.Item(1).Item(0))
            Assert.AreEqual(10, merge_dest3.Rows.Item(2).Item(0))
            'Assert.AreEqual(4, merge_dest3.Rows.Item(3).Item(0))
            'Assert.AreEqual(5, merge_dest3.Rows.Item(4)l.Item(0))
            Assert.AreEqual(2, merge_dest3.Rows.Item(3).Item(0))
            Assert.AreEqual(3, merge_dest3.Rows.Item(4).Item(0))


            'Ensure in general correct merge'
            StringAssert.IsMatch("Text1", merge_dest3.Rows.Item(0).Item(1))
            StringAssert.IsMatch("Not touched", merge_dest3.Rows.Item(1).Item(1))
            StringAssert.IsMatch("...", merge_dest3.Rows.Item(2).Item(1))
            'StringAssert.IsMatch("Text4", merge_dest3.Rows.Item(3).Item(1))
            'StringAssert.IsMatch("Text5", merge_dest3.Rows.Item(4).Item(1))
            StringAssert.IsMatch("Text2", merge_dest3.Rows.Item(3).Item(1))
            StringAssert.IsMatch("Text3", merge_dest3.Rows.Item(4).Item(1))

#If Not NET_1_1 Then
            Dim big As New DataTable
            Dim bigCopy As New DataTable
            Dim bigCopy2 As New DataTable

            big.ReadXml(System.Environment.CurrentDirectory & "\testfiles/3000RowsTable.xml")

            CompuMaster.Data.DataTables.CreateDataTableClone(big, bigCopy, "", "", Nothing, CompuMaster.Data.DataTables.RequestedRowChanges.DropExistingRowsInDestinationTableAndInsertNewRows, False,
                                                             CompuMaster.Data.DataTables.RequestedSchemaChangesForUnusedColumns.Remove, CompuMaster.Data.DataTables.RequestedSchemaChangesForExistingColumns.None,
                                                             CompuMaster.Data.DataTables.RequestedSchemaChangesForAdditionalColumns.Add)
            Assert.AreEqual(big.Columns.Count, bigCopy.Columns.Count)
            Assert.AreEqual(big.Rows.Count, bigCopy.Rows.Count)

            bigCopy.Rows.Item(0).Item(1) = 29
            bigCopy.Rows.Item(1020).Item(1) = 20
            bigCopy.Rows.Item(2323).Item(1) = 99
            bigCopy.Rows.Item(1000).Item(1) = 22
            bigCopy.Rows.Item(900).Item(1) = 55
            bigCopy.Rows.Item(bigCopy.Rows.Count - 5).Item(1) = 78

            CompuMaster.Data.DataTables.CreateDataTableClone(bigCopy, bigCopy2, "", "", Nothing, CompuMaster.Data.DataTables.RequestedRowChanges.DropExistingRowsInDestinationTableAndInsertNewRows,
                                                             False, CompuMaster.Data.DataTables.RequestedSchemaChangesForUnusedColumns.None, CompuMaster.Data.DataTables.RequestedSchemaChangesForExistingColumns.None,
                                                             CompuMaster.Data.DataTables.RequestedSchemaChangesForAdditionalColumns.Add)

            Assert.AreEqual(bigCopy.Rows.Count(), bigCopy2.Rows.Count())
            Assert.AreEqual(bigCopy.Columns.Count, bigCopy2.Columns.Count)
            Assert.AreEqual(20, bigCopy2.Rows.Item(1020).Item(1))
            Assert.AreEqual(55, bigCopy2.Rows.Item(900).Item(1))
#End If

            'TODO: all variations'
        End Sub

        <Test()> Public Sub FindUniqueValues()
            Dim dt As New DataTable
            Dim row As DataRow
            dt.Columns.Add("Column Name")

            row = dt.NewRow
            row.Item(0) = "Sun"
            dt.Rows.Add(row)

            row = dt.NewRow
            row.Item(0) = "Moon"
            dt.Rows.Add(row)

            row = dt.NewRow
            row.Item(0) = "Sun"
            dt.Rows.Add(row)

            row = dt.NewRow
            row.Item(0) = "Saturn"
            dt.Rows.Add(row)

            row = dt.NewRow
            row.Item(0) = DBNull.Value
            dt.Rows.Add(row)

            row = dt.NewRow
            row.Item(0) = "Saturn"
            dt.Rows.Add(row)

            Dim list As ArrayList
            list = CompuMaster.Data.DataTables.FindUniqueValues(dt.Columns(0), False)
            StringAssert.IsMatch("Sun", list(0))
            StringAssert.IsMatch("Moon", list(1))
            StringAssert.IsMatch("Saturn", list(2))
            Assert.IsTrue(IsDBNull(list(3)), "Expected DbNull value")
            Assert.AreEqual(4, list.Count)
            list = CompuMaster.Data.DataTables.FindUniqueValues(dt.Columns(0), True)
            StringAssert.IsMatch("Sun", list(0))
            StringAssert.IsMatch("Moon", list(1))
            StringAssert.IsMatch("Saturn", list(2))
            Assert.AreEqual(3, list.Count)
        End Sub

        <Test()> Public Sub LookupUniqueColumnName()
            Dim dt As New DataTable
            dt.Columns.Add("Test")
            dt.Columns.Add("Test2")
            dt.Columns.Add("Test3")
            dt.Columns.Add("Test4")
            dt.Columns.Add("Test14")
            dt.Columns.Add("Test15")
            dt.Columns.Add("1")
            dt.Columns.Add("ID")
            dt.Columns.Add("ClientTable_")
            dt.Columns.Add("ClientTable_ID")
            dt.Columns.Add("ClientTable_ID2")
            dt.Columns.Add("ClientTable_ID23")
            Dim uname As String, lname As String
            lname = "Test"
            uname = CompuMaster.Data.DataTables.LookupUniqueColumnName(dt, lname)
            Console.WriteLine("New free unique column name: " & uname & " (lookup name: " & lname & ")")
            Assert.IsNotEmpty(uname)
            Assert.IsFalse(dt.Columns.Contains(uname))
            lname = "Test1"
            uname = CompuMaster.Data.DataTables.LookupUniqueColumnName(dt, lname)
            Console.WriteLine("New free unique column name: " & uname & " (lookup name: " & lname & ")")
            Assert.IsNotEmpty(uname)
            Assert.IsFalse(dt.Columns.Contains(uname))
            Assert.AreEqual("Test1", uname)
            lname = "Test2"
            uname = CompuMaster.Data.DataTables.LookupUniqueColumnName(dt, lname)
            Console.WriteLine("New free unique column name: " & uname & " (lookup name: " & lname & ")")
            Assert.IsNotEmpty(uname)
            Assert.IsFalse(dt.Columns.Contains(uname))
            lname = "test2"
            uname = CompuMaster.Data.DataTables.LookupUniqueColumnName(dt, lname)
            Console.WriteLine("New free unique column name: " & uname & " (lookup name: " & lname & ")")
            Assert.AreNotEqual("test2", uname)
            Assert.IsNotEmpty(uname)
            Assert.IsFalse(dt.Columns.Contains(uname))
            lname = "Test4"
            uname = CompuMaster.Data.DataTables.LookupUniqueColumnName(dt, lname)
            Console.WriteLine("New free unique column name: " & uname & " (lookup name: " & lname & ")")
            Assert.IsNotEmpty(uname)
            Assert.IsFalse(dt.Columns.Contains(uname))
            lname = "Test14"
            uname = CompuMaster.Data.DataTables.LookupUniqueColumnName(dt, lname)
            Console.WriteLine("New free unique column name: " & uname & " (lookup name: " & lname & ")")
            Assert.IsNotEmpty(uname)
            Assert.IsFalse(dt.Columns.Contains(uname))
            lname = "ClientTable_"
            uname = CompuMaster.Data.DataTables.LookupUniqueColumnName(dt, lname)
            Console.WriteLine("New free unique column name: " & uname & " (lookup name: " & lname & ")")
            Assert.IsNotEmpty(uname)
            Assert.IsFalse(dt.Columns.Contains(uname))
            lname = "1"
            uname = CompuMaster.Data.DataTables.LookupUniqueColumnName(dt, lname)
            Console.WriteLine("New free unique column name: " & uname & " (lookup name: " & lname & ")")
            Assert.IsNotEmpty(uname)
            Assert.IsFalse(dt.Columns.Contains(uname))
            dt.Columns.Add("ClientTable_1")
            lname = "ClientTable_"
            uname = CompuMaster.Data.DataTables.LookupUniqueColumnName(dt, lname)
            Console.WriteLine("New free unique column name: " & uname & " (lookup name: " & lname & ")")
            Assert.IsNotEmpty(uname)
            Assert.IsFalse(dt.Columns.Contains(uname))
            lname = "1"
            uname = CompuMaster.Data.DataTables.LookupUniqueColumnName(dt, lname)
            Console.WriteLine("New free unique column name: " & uname & " (lookup name: " & lname & ")")
            Assert.IsNotEmpty(uname)
            Assert.IsFalse(dt.Columns.Contains(uname))
            lname = "ID"
            uname = CompuMaster.Data.DataTables.LookupUniqueColumnName(dt, lname)
            Console.WriteLine("New free unique column name: " & uname & " (lookup name: " & lname & ")")
            Assert.IsNotEmpty(uname)
            Assert.IsFalse(dt.Columns.Contains(uname))
            lname = "ClientTable_ID"
            uname = CompuMaster.Data.DataTables.LookupUniqueColumnName(dt, lname)
            Console.WriteLine("New free unique column name: " & uname & " (lookup name: " & lname & ")")
            Assert.IsNotEmpty(uname)
            Assert.IsFalse(dt.Columns.Contains(uname))
            lname = "clienttable_id"
            uname = CompuMaster.Data.DataTables.LookupUniqueColumnName(dt, lname)
            Console.WriteLine("New free unique column name: " & uname & " (lookup name: " & lname & ")")
            Assert.AreNotEqual("clienttable_id", uname)
            Assert.IsNotEmpty(uname)
            Assert.IsFalse(dt.Columns.Contains(uname))
            lname = "ClientTable_ID2"
            uname = CompuMaster.Data.DataTables.LookupUniqueColumnName(dt, lname)
            Console.WriteLine("New free unique column name: " & uname & " (lookup name: " & lname & ")")
            Assert.IsNotEmpty(uname)
            Assert.IsFalse(dt.Columns.Contains(uname))
            lname = "ClientTable_ID23"
            uname = CompuMaster.Data.DataTables.LookupUniqueColumnName(dt, lname)
            Console.WriteLine("New free unique column name: " & uname & " (lookup name: " & lname & ")")
            Assert.IsNotEmpty(uname)
            Assert.IsFalse(dt.Columns.Contains(uname))
        End Sub

        <Test()> Public Sub ReArrangeDataColumns()
            Dim dt As New DataTable
            Dim dt2 As DataTable
            dt.Columns.Add("Test1")
            dt.Columns.Add("Test2")

            dt.Rows.Add(New String() {"Test1", "Test3"})
            dt.Rows.Add(New String() {"Test2", "Test4"})
            dt.Rows.Add(New String() {"Test5", "Test6"})
            dt.Rows.Add(New String() {"Test7", "Test8"})
            dt.Rows.Add(New String() {"Test9", "Test10"})


            'TODO: all overloads'

            dt2 = CompuMaster.Data.DataTables.ReArrangeDataColumns(dt, New String() {"Test1"})
            Assert.AreEqual(1, dt2.Columns.Count())
            Assert.AreEqual(5, dt2.Rows.Count())
            StringAssert.IsMatch("Test1", dt2.Columns.Item(0).ColumnName)
            StringAssert.IsMatch("Test9", dt2.Rows.Item(4).Item(0))

        End Sub

        <Test()> Public Sub RemoveColumns()
            Dim dt As New DataTable
            dt.Columns.Add("SomeColumn")

            Dim row1 As DataRow = dt.NewRow
            row1.Item(0) = "D"

            Assert.IsTrue(dt.Columns.Item(0).ColumnName = "SomeColumn")
            CompuMaster.Data.DataTables.RemoveColumns(dt, New String() {"SomeColumn"})
            Assert.AreEqual(0, dt.Columns.Count())
        End Sub

        <Test()> Public Sub RemoveDuplicates()
            Dim dt As New DataTable
            dt.Columns.Add("Something")
            dt.Columns.Add("Something2")

            Dim row1 As DataRow = dt.NewRow
            row1.Item(0) = "Unique"
            row1.Item(0) = "Unique1"

            Dim row2 As DataRow = dt.NewRow
            row2.Item(0) = "NotSoUnique"
            row2.Item(0) = "NotSoUnique"

            Dim row3 As DataRow = dt.NewRow
            row3.Item(0) = "NotSoUnique"
            row3.Item(0) = "NotSoUnique"

            dt.Rows.Add(row1)
            dt.Rows.Add(row2)
            dt.Rows.Add(row3)

            CompuMaster.Data.DataTables.RemoveDuplicates(dt, "Something")

            Assert.AreEqual(2, dt.Rows.Count())
        End Sub

        <Test()> Public Sub RemoveRowsWithColumnValues()
            Dim dt As New DataTable
            dt.Columns.Add("Column1")
            dt.Columns.Add("Column2")

            Dim row1 As DataRow = dt.NewRow
            Dim row2 As DataRow = dt.NewRow
            Dim row3 As DataRow = dt.NewRow


            row1.Item(0) = "YouCanDeleteThis"
            row2.Item(0) = "YouCanDeleteThis"
            row3.Item(0) = "ButYouCouldKeepThis"

            dt.Rows.Add(row1)
            dt.Rows.Add(row2)
            dt.Rows.Add(row3)

            CompuMaster.Data.DataTables.RemoveRowsWithColumnValues(dt.Columns.Item(0), New String() {"YouCanDeleteThis"})
            Assert.AreEqual(1, dt.Rows.Count())
        End Sub

        <Test()> Public Sub RemoveRowsWithNoCorrespondingValueInComparisonTable()
            Dim dt As New DataTable
            dt.Columns.Add("Something")
            dt.Columns.Add("Something2")


            dt.Rows.Add(New String() {"A", "Z"})
            dt.Rows.Add(New String() {"B", "Y"})
            dt.Rows.Add(New String() {"C", "X"})

            Dim dt2 As New DataTable

            dt2.Columns.Add("Test")
            dt2.Columns.Add("TestColumn2")

            dt2.Rows.Add(New String() {"A", "Z"})
            dt2.Rows.Add(New String() {"B", "Y"})
            dt2.Rows.Add(New String() {"D", "W"})


            CompuMaster.Data.DataTables.RemoveRowsWithNoCorrespondingValueInComparisonTable(dt.Columns(0), dt2.Columns(0))
            Assert.AreEqual(2, dt.Rows.Count())
            StringAssert.IsMatch("A", dt.Rows.Item(0).Item(0))
            StringAssert.IsMatch("B", dt.Rows.Item(1).Item(0))
            StringAssert.IsMatch("Z", dt.Rows.Item(0).Item(1))
            StringAssert.IsMatch("Y", dt.Rows.Item(1).Item(1))

        End Sub

        <Test()> Public Sub RowIndex()
            Dim dt As New DataTable
            dt.Columns.Add("DummyColumn")
            Dim drow As DataRow = dt.NewRow
            dt.Rows.Add(drow)
            Assert.AreEqual(0, CompuMaster.Data.DataTables.RowIndex(drow))

        End Sub


        Private Function CreateCrossJoinTablesTableSet1() As JoinTableSet
            Dim left As New DataTable
            Dim right As New DataTable
            'Cross join test'
            left.Columns.Add("FirstCol")
            left.Columns.Add("SecondCol")
            left.Columns.Add("ThirdCol")
            Dim lDataRow As DataRow = left.NewRow
            lDataRow.Item(0) = "Test"
            lDataRow.Item(1) = "Test2"
            left.Rows.Add(lDataRow)

            right.Columns.Add("FourthCol")
            Dim rDataRow As DataRow = right.NewRow
            rDataRow.Item(0) = "RightTest"
            right.Rows.Add(rDataRow)

            Return New JoinTableSet("CreateCrossJoinTableSet1", left, Nothing, right, Nothing)
        End Function

        Private Function CreateCrossJoinTablesTableSet2() As JoinTableSet
            Dim left As New DataTable
            Dim right As New DataTable
            'Cross join test'
            left.Columns.Add("FirstCol")
            left.Columns.Add("SecondCol")
            left.Columns.Add("ThirdCol")
            Dim lDataRow As DataRow = left.NewRow
            lDataRow.Item(0) = "Test1"
            lDataRow.Item(1) = "Test1Col2"
            left.Rows.Add(lDataRow)
            lDataRow = left.NewRow
            lDataRow.Item(0) = "Test2"
            lDataRow.Item(1) = "Test2Col2"
            left.Rows.Add(lDataRow)

            right.Columns.Add("FourthCol")
            Dim rDataRow As DataRow = right.NewRow
            rDataRow.Item(0) = "RightTest1"
            right.Rows.Add(rDataRow)
            rDataRow = right.NewRow
            rDataRow.Item(0) = "RightTest2"
            right.Rows.Add(rDataRow)
            rDataRow = right.NewRow
            rDataRow.Item(0) = "RightTest3"
            right.Rows.Add(rDataRow)

            Return New JoinTableSet("CreateCrossJoinTableSet2", left, Nothing, right, Nothing)
        End Function

        <Test()> Public Sub CrossJoinTables()

            Dim TestTableSet As JoinTableSet = Me.CreateCrossJoinTablesTableSet1
            TestTableSet.WriteToConsole()

            Dim crossjoined As DataTable = CompuMaster.Data.DataTables.CrossJoinTables(TestTableSet.LeftTable, Nothing, TestTableSet.RightTable, Nothing)
            StringAssert.IsMatch("FourthCol", crossjoined.Columns.Item(3).ColumnName)
            StringAssert.IsMatch("RightTest", crossjoined.Rows.Item(0).Item(3))
            Assert.IsTrue(IsDBNull(crossjoined.Rows.Item(0).Item(2)))


            Console.WriteLine("FULL-OUTER-JOINED TABLE CONTENTS")
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(crossjoined))
        End Sub

        <Test()> Public Sub SqlJoinTables_CrossJoin()
            SqlJoinTables_CrossJoin_Test1()
            SqlJoinTables_CrossJoin_Test2()
        End Sub
        Private Sub SqlJoinTables_CrossJoin_Test1()

            Dim TestTableSet As JoinTableSet = Me.CreateCrossJoinTablesTableSet1
            TestTableSet.WriteToConsole()

            Dim CrossJoined As DataTable = CompuMaster.Data.DataTables.SqlJoinTables(TestTableSet.LeftTable, New String() {}, New String() {}, TestTableSet.RightTable, New String() {}, New String() {}, CompuMaster.Data.DataTables.SqlJoinTypes.Cross)

            Console.WriteLine("CROS-JOINED TABLE CONTENTS")
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(CrossJoined))

            'TODO: ResultComparisonValue to be evaluated
            Dim ShallBeResult As String = "JoinedTable" & vbNewLine &
                    "FirstCol|SecondCol|ThirdCol|FourthCol" & vbNewLine &
                    "--------+---------+--------+---------" & vbNewLine &
                    "Test    |Test2    |        |RightTest" & vbNewLine &
                    ""
            Assert.AreEqual(ShallBeResult, CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(CrossJoined))

        End Sub

        Private Sub SqlJoinTables_CrossJoin_Test2()

            Dim TestTableSet As JoinTableSet = Me.CreateCrossJoinTablesTableSet2
            TestTableSet.WriteToConsole()

            Dim CrossJoined As DataTable = CompuMaster.Data.DataTables.SqlJoinTables(TestTableSet.LeftTable, New String() {}, New String() {}, TestTableSet.RightTable, New String() {}, New String() {}, CompuMaster.Data.DataTables.SqlJoinTypes.Cross)

            Console.WriteLine("CROS-JOINED TABLE CONTENTS")
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(CrossJoined))

            'TODO: ResultComparisonValue to be evaluated
            Dim ShallBeResult As String = "JoinedTable" & vbNewLine &
                    "FirstCol|SecondCol|ThirdCol|FourthCol " & vbNewLine &
                    "--------+---------+--------+----------" & vbNewLine &
                    "Test1   |Test1Col2|        |RightTest1" & vbNewLine &
                    "Test1   |Test1Col2|        |RightTest2" & vbNewLine &
                    "Test1   |Test1Col2|        |RightTest3" & vbNewLine &
                    "Test2   |Test2Col2|        |RightTest1" & vbNewLine &
                    "Test2   |Test2Col2|        |RightTest2" & vbNewLine &
                    "Test2   |Test2Col2|        |RightTest3" & vbNewLine &
                    ""
            Assert.AreEqual(ShallBeResult, CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(CrossJoined))

        End Sub

        Private Class JoinTableSet
            Public Sub New(tableSetName As String, leftTable As DataTable, leftPrimaryKeyColumns As String(), rightTable As DataTable, rightPrimaryKeyColumns As String())
                Me.TableSetName = tableSetName
                Me.LeftTable = leftTable
                Me.RightTable = rightTable
                Me.LeftPrimaryKeyColumns = leftPrimaryKeyColumns
                Me.RightPrimaryKeyColumns = rightPrimaryKeyColumns
            End Sub
            Public TableSetName As String
            Public LeftTable As DataTable
            Public RightTable As DataTable
            Public LeftPrimaryKeyColumns As String()
            Public RightPrimaryKeyColumns As String()
            Public Sub WriteToConsole()
                'Console.WriteLine("LEFT TABLE CONTENTS OF " & Me.TableSetName)
                'Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(Me.LeftTable))
                'Console.WriteLine()
                'Console.WriteLine("RIGHT TABLE CONTENTS OF " & Me.TableSetName)
                'Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(Me.RightTable))
                'Console.WriteLine()
                Console.WriteLine("LEFT+RIGHT TABLE CONTENTS OF " & Me.TableSetName)
                Console.WriteLine(CompuMaster.Data.Utils.ArrangeTableBlocksBesides(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(Me.LeftTable), CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(Me.RightTable)))
            End Sub
        End Class

#Region "FullOuterJoinTestTables"""
        Private Function CreateFullOuterJoinTableSet1() As JoinTableSet
            'Full outer join test'
            Dim fullOuterJoined As New DataTable
            Dim right As New DataTable
            Dim left As New DataTable

            left.Columns.Add("left1")
            left.Columns.Add("left2")
            left.Columns.Add("left3")

            right.Columns.Add("right1")
            right.Columns.Add("right2")

            For j As Integer = 0 To 5
                Dim leftRow As DataRow = left.NewRow
                leftRow.Item(0) = j
                leftRow.Item(1) = j + 1
                leftRow.Item(2) = j + 2
                left.Rows.Add(leftRow)
            Next j

            For i As Integer = 0 To 5
                Dim rightrow As DataRow = right.NewRow
                rightrow.Item(0) = i
                rightrow.Item(1) = i + 10
                right.Rows.Add(rightrow)
            Next i

            Dim addiontalLeft As DataRow = left.NewRow
            addiontalLeft.Item(0) = 567
            addiontalLeft.Item(1) = 65527
            left.Rows.Add(addiontalLeft)
            addiontalLeft = left.NewRow
            addiontalLeft.Item(0) = 678
            addiontalLeft.Item(1) = 65528
            left.Rows.Add(addiontalLeft)
            addiontalLeft = left.NewRow
            left.Rows.Add(addiontalLeft)

            Dim addiontalRight As DataRow = right.NewRow
            addiontalRight.Item(0) = 789
            addiontalRight.Item(1) = 65728
            right.Rows.Add(addiontalRight)
            addiontalRight = right.NewRow
            addiontalRight.Item(0) = 890
            right.Rows.Add(addiontalRight)
            addiontalRight = right.NewRow
            right.Rows.Add(addiontalRight)

            Return New JoinTableSet("CreateFullOuterJoinTableSet1", left, New String() {"left1"}, right, New String() {"right1"})
        End Function

        Private Function CreateFullOuterJoinTableSet2() As JoinTableSet
            'Full outer join test'
            Dim fullOuterJoined As New DataTable
            Dim right As New DataTable
            Dim left As New DataTable

            left.Columns.Add("left1")
            left.Columns.Add("left2")
            left.Columns.Add("left3")

            right.Columns.Add("right1")
            right.Columns.Add("right2")

            For j As Integer = 0 To 5
                Dim leftRow As DataRow = left.NewRow
                If j = 1 Then
                    leftRow.Item(0) = DBNull.Value
                Else
                    leftRow.Item(0) = j
                End If
                leftRow.Item(1) = j + 1
                leftRow.Item(2) = j + 2
                left.Rows.Add(leftRow)
            Next j

            For i As Integer = 0 To 5
                Dim rightrow As DataRow = right.NewRow
                If i = 2 Then
                    rightrow.Item(0) = DBNull.Value
                Else
                    rightrow.Item(0) = i
                End If

                If i = 4 Then
                    rightrow.Item(1) = i + 100
                ElseIf i = 5 Then
                    rightrow.Item(1) = DBNull.Value
                Else
                    rightrow.Item(1) = i + 1
                End If
                right.Rows.Add(rightrow)
            Next i

            Dim addiontalLeft As DataRow = left.NewRow
            addiontalLeft.Item(0) = 567
            addiontalLeft.Item(1) = 65527
            left.Rows.Add(addiontalLeft)
            addiontalLeft = left.NewRow
            addiontalLeft.Item(0) = 5
            addiontalLeft.Item(1) = 60
            addiontalLeft.Item(2) = 70
            left.Rows.Add(addiontalLeft)


            Dim addiontalRight As DataRow = right.NewRow
            addiontalRight.Item(0) = 789
            addiontalRight.Item(1) = 65728
            right.Rows.Add(addiontalRight)
            addiontalRight = right.NewRow
            addiontalRight.Item(0) = 3
            addiontalRight.Item(1) = 40
            right.Rows.Add(addiontalRight)
            addiontalRight = right.NewRow
            addiontalRight.Item(0) = 890
            right.Rows.Add(addiontalRight)

            Return New JoinTableSet("CreateFullOuterJoinTableSet2", left, New String() {"left1"}, right, New String() {"right1"})
        End Function
#End Region

        <Test()> Public Sub SqlJoinTables_FullOuter()
            FullOuterJoinTables_TableSet1()
            FullOuterJoinTables_TableSet1WithSameColumnNameInPrimaryKeyAtBothTables()
            FullOuterJoinTables_TableSet1WithSameColumnNameInPrimaryKeyAtBothTablesNamedClientTable_ID()
            FullOuterJoinTables_TableSet2()
        End Sub

        Private Sub FullOuterJoinTables_TableSet1()
            Dim TestTables As JoinTableSet = Me.CreateFullOuterJoinTableSet1
            TestTables.WriteToConsole()

            Dim FullOuterJoined As DataTable
            'FullOuterJoined = CompuMaster.Data.DataTables.FullJoinTables(TestTables.LeftTable, TestTables.LeftPrimaryKeyColumns, TestTables.RightTable, TestTables.RightPrimaryKeyColumns)
            FullOuterJoined = CompuMaster.Data.DataTables.SqlJoinTables(TestTables.LeftTable, TestTables.LeftPrimaryKeyColumns, Nothing, TestTables.RightTable, TestTables.RightPrimaryKeyColumns, Nothing, CompuMaster.Data.DataTables.SqlJoinTypes.FullOuter)

            Console.WriteLine("FULL-OUTER-JOINED TABLE CONTENTS: " & TestTables.TableSetName)
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(FullOuterJoined))

            'Verify column (names)
            Assert.AreEqual(5, FullOuterJoined.Columns.Count()) 'There must be exactly 5 columns '
            For i As Integer = 1 To 3
                Assert.AreEqual("left" & i, FullOuterJoined.Columns(i - 1).ColumnName)
            Next i
            For i As Integer = 1 To 2
                Assert.AreEqual("right" & i, FullOuterJoined.Columns(i + 2).ColumnName)
            Next i

            'Verify row (content)
            Assert.AreEqual(11, FullOuterJoined.Rows.Count())
            For i As Integer = 0 To 5
                Assert.AreEqual(0 + i, CInt(FullOuterJoined.Rows(i)(0)))
                Assert.AreEqual(1 + i, CInt(FullOuterJoined.Rows(i)(1)))
                Assert.AreEqual(2 + i, CInt(FullOuterJoined.Rows(i)(2)))
                Assert.AreEqual(0 + i, CInt(FullOuterJoined.Rows(i)(3)))
                Assert.AreEqual(10 + i, CInt(FullOuterJoined.Rows(i)(4)))
            Next i
            Assert.AreEqual(567, CInt(FullOuterJoined.Rows(6)(0)))
            Assert.AreEqual(65527, CInt(FullOuterJoined.Rows(6)(1)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(6)(2)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(6)(3)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(6)(4)))
            Assert.AreEqual(678, CInt(FullOuterJoined.Rows(7)(0)))
            Assert.AreEqual(65528, CInt(FullOuterJoined.Rows(7)(1)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(7)(2)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(7)(3)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(7)(4)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(8)(0)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(8)(1)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(8)(2)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(8)(3)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(8)(4)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(9)(0)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(9)(1)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(9)(2)))
            Assert.AreEqual(789, CInt(FullOuterJoined.Rows(9)(3)))
            Assert.AreEqual(65728, CInt(FullOuterJoined.Rows(9)(4)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(10)(0)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(10)(1)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(10)(2)))
            Assert.AreEqual(890, CInt(FullOuterJoined.Rows(10)(3)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(10)(4)))

        End Sub

        Private Sub FullOuterJoinTables_TableSet1WithSameColumnNameInPrimaryKeyAtBothTables()
            Dim TestTables As JoinTableSet = Me.CreateFullOuterJoinTableSet1
            TestTables.TableSetName &= "::WithSameColumnNameInPrimaryKeyAtBothTables"
            TestTables.LeftTable.Columns(0).ColumnName = "ID"
            TestTables.RightTable.Columns(0).ColumnName = "ID"
            TestTables.LeftPrimaryKeyColumns = New String() {"ID"}
            TestTables.RightPrimaryKeyColumns = New String() {"ID"}
            TestTables.WriteToConsole()

            Dim FullOuterJoined As DataTable
            'FullOuterJoined = CompuMaster.Data.DataTables.FullJoinTables(TestTables.LeftTable, TestTables.LeftPrimaryKeyColumns, TestTables.RightTable, TestTables.RightPrimaryKeyColumns)
            FullOuterJoined = CompuMaster.Data.DataTables.SqlJoinTables(TestTables.LeftTable, TestTables.LeftPrimaryKeyColumns, Nothing, TestTables.RightTable, TestTables.RightPrimaryKeyColumns, Nothing, CompuMaster.Data.DataTables.SqlJoinTypes.FullOuter)

            Console.WriteLine("FULL-OUTER-JOINED TABLE CONTENTS: " & TestTables.TableSetName)
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(FullOuterJoined))

            'Verify column (names)
            Assert.AreEqual(5, FullOuterJoined.Columns.Count()) 'There must be exactly 5 columns '
            Assert.AreEqual("ID", FullOuterJoined.Columns(0).ColumnName)
            For i As Integer = 2 To 3
                Assert.AreEqual("left" & i, FullOuterJoined.Columns(i - 1).ColumnName)
            Next i
            Assert.AreEqual("ClientTable_ID", FullOuterJoined.Columns(3).ColumnName)
            For i As Integer = 2 To 2
                Assert.AreEqual("right" & i, FullOuterJoined.Columns(i + 2).ColumnName)
            Next i

            'Verify row (content)
            Assert.AreEqual(11, FullOuterJoined.Rows.Count())
            For i As Integer = 0 To 5
                Assert.AreEqual(0 + i, CInt(FullOuterJoined.Rows(i)(0)))
                Assert.AreEqual(1 + i, CInt(FullOuterJoined.Rows(i)(1)))
                Assert.AreEqual(2 + i, CInt(FullOuterJoined.Rows(i)(2)))
                Assert.AreEqual(0 + i, CInt(FullOuterJoined.Rows(i)(3)))
                Assert.AreEqual(10 + i, CInt(FullOuterJoined.Rows(i)(4)))
            Next i
            Assert.AreEqual(567, CInt(FullOuterJoined.Rows(6)(0)))
            Assert.AreEqual(65527, CInt(FullOuterJoined.Rows(6)(1)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(6)(2)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(6)(3)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(6)(4)))
            Assert.AreEqual(678, CInt(FullOuterJoined.Rows(7)(0)))
            Assert.AreEqual(65528, CInt(FullOuterJoined.Rows(7)(1)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(7)(2)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(7)(3)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(7)(4)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(8)(0)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(8)(1)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(8)(2)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(8)(3)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(8)(4)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(9)(0)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(9)(1)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(9)(2)))
            Assert.AreEqual(789, CInt(FullOuterJoined.Rows(9)(3)))
            Assert.AreEqual(65728, CInt(FullOuterJoined.Rows(9)(4)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(10)(0)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(10)(1)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(10)(2)))
            Assert.AreEqual(890, CInt(FullOuterJoined.Rows(10)(3)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(10)(4)))

        End Sub

        Private Sub FullOuterJoinTables_TableSet1WithSameColumnNameInPrimaryKeyAtBothTablesNamedClientTable_ID()
            Dim TestTables As JoinTableSet = Me.CreateFullOuterJoinTableSet1
            TestTables.TableSetName &= "::WithSameColumnNameInPrimaryKeyAtBothTablesNamedClientTable_ID"
            TestTables.LeftTable.Columns(0).ColumnName = "ClientTable_ID"
            TestTables.RightTable.Columns(0).ColumnName = "ClientTable_ID"
            TestTables.LeftPrimaryKeyColumns = New String() {"ClientTable_ID"}
            TestTables.RightPrimaryKeyColumns = New String() {"ClientTable_ID"}
            TestTables.WriteToConsole()

            Dim FullOuterJoined As DataTable
            'FullOuterJoined = CompuMaster.Data.DataTables.FullJoinTables(TestTables.LeftTable, TestTables.LeftPrimaryKeyColumns, TestTables.RightTable, TestTables.RightPrimaryKeyColumns)
            FullOuterJoined = CompuMaster.Data.DataTables.SqlJoinTables(TestTables.LeftTable, TestTables.LeftPrimaryKeyColumns, Nothing, TestTables.RightTable, TestTables.RightPrimaryKeyColumns, Nothing, CompuMaster.Data.DataTables.SqlJoinTypes.FullOuter)

            Console.WriteLine("FULL-OUTER-JOINED TABLE CONTENTS: " & TestTables.TableSetName)
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(FullOuterJoined))

            'Verify column (names)
            Assert.AreEqual(5, FullOuterJoined.Columns.Count()) 'There must be exactly 5 columns '
            Assert.AreEqual("ClientTable_ID", FullOuterJoined.Columns(0).ColumnName)
            For i As Integer = 2 To 3
                Assert.AreEqual("left" & i, FullOuterJoined.Columns(i - 1).ColumnName)
            Next i
            Assert.AreEqual("ClientTable_ID1", FullOuterJoined.Columns(3).ColumnName)
            For i As Integer = 2 To 2
                Assert.AreEqual("right" & i, FullOuterJoined.Columns(i + 2).ColumnName)
            Next i

            'Verify row (content)
            Assert.AreEqual(11, FullOuterJoined.Rows.Count())
            For i As Integer = 0 To 5
                Assert.AreEqual(0 + i, CInt(FullOuterJoined.Rows(i)(0)))
                Assert.AreEqual(1 + i, CInt(FullOuterJoined.Rows(i)(1)))
                Assert.AreEqual(2 + i, CInt(FullOuterJoined.Rows(i)(2)))
                Assert.AreEqual(0 + i, CInt(FullOuterJoined.Rows(i)(3)))
                Assert.AreEqual(10 + i, CInt(FullOuterJoined.Rows(i)(4)))
            Next i
            Assert.AreEqual(567, CInt(FullOuterJoined.Rows(6)(0)))
            Assert.AreEqual(65527, CInt(FullOuterJoined.Rows(6)(1)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(6)(2)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(6)(3)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(6)(4)))
            Assert.AreEqual(678, CInt(FullOuterJoined.Rows(7)(0)))
            Assert.AreEqual(65528, CInt(FullOuterJoined.Rows(7)(1)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(7)(2)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(7)(3)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(7)(4)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(8)(0)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(8)(1)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(8)(2)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(8)(3)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(8)(4)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(9)(0)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(9)(1)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(9)(2)))
            Assert.AreEqual(789, CInt(FullOuterJoined.Rows(9)(3)))
            Assert.AreEqual(65728, CInt(FullOuterJoined.Rows(9)(4)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(10)(0)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(10)(1)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(10)(2)))
            Assert.AreEqual(890, CInt(FullOuterJoined.Rows(10)(3)))
            Assert.AreEqual(True, IsDBNull(FullOuterJoined.Rows(10)(4)))

        End Sub

        Private Sub FullOuterJoinTables_TableSet2()
            Dim TestTables As JoinTableSet = Me.CreateFullOuterJoinTableSet2
            TestTables.WriteToConsole()

            Dim FullOuterJoined As DataTable
            'FullOuterJoined = CompuMaster.Data.DataTables.FullJoinTables(TestTables.LeftTable, TestTables.LeftPrimaryKeyColumns, TestTables.RightTable, TestTables.RightPrimaryKeyColumns)
            FullOuterJoined = CompuMaster.Data.DataTables.SqlJoinTables(TestTables.LeftTable, TestTables.LeftPrimaryKeyColumns, Nothing, TestTables.RightTable, TestTables.RightPrimaryKeyColumns, Nothing, CompuMaster.Data.DataTables.SqlJoinTypes.FullOuter)

            Console.WriteLine("FULL-OUTER-JOINED TABLE CONTENTS: " & TestTables.TableSetName)
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(FullOuterJoined))

            'Verify column (names)
            Assert.AreEqual(5, FullOuterJoined.Columns.Count()) 'There must be exactly 5 columns '
            For i As Integer = 1 To 3
                Assert.AreEqual("left" & i, FullOuterJoined.Columns(i - 1).ColumnName)
            Next i
            For i As Integer = 1 To 2
                Assert.AreEqual("right" & i, FullOuterJoined.Columns(i + 2).ColumnName)
            Next i

            'Verify row (content)
            Assert.AreEqual(12, FullOuterJoined.Rows.Count())
            For i As Integer = 0 To 3
                If i = 1 Then
                    Assert.AreEqual(DBNull.Value, FullOuterJoined.Rows(i)(0))
                Else
                    Assert.AreEqual(0 + i, CInt(FullOuterJoined.Rows(i)(0)))
                End If
                Assert.AreEqual(1 + i, CInt(FullOuterJoined.Rows(i)(1)))
                Assert.AreEqual(2 + i, CInt(FullOuterJoined.Rows(i)(2)))
                If i = 1 Or i = 2 Then
                    Assert.AreEqual(DBNull.Value, FullOuterJoined.Rows(i)(3))
                Else
                    Assert.AreEqual(0 + i, CInt(FullOuterJoined.Rows(i)(3)))
                End If
                If i = 1 Then
                    Assert.AreEqual(3, CInt(FullOuterJoined.Rows(i)(4)))
                ElseIf i = 2 Then
                    Assert.AreEqual(DBNull.Value, FullOuterJoined.Rows(i)(4))
                ElseIf i = 4
                    Assert.AreEqual(1 + i, CInt(FullOuterJoined.Rows(i)(4)))
                End If
            Next i

            Dim ShallBeResult As String = "JoinedTable" & vbNewLine &
                    "left1|left2|left3|right1|right2" & vbNewLine &
                    "-----+-----+-----+------+------" & vbNewLine &
                    "0    |1    |2    |0     |1     " & vbNewLine &
                    "     |2    |3    |      |3     " & vbNewLine &
                    "2    |3    |4    |      |      " & vbNewLine &
                    "3    |4    |5    |3     |4     " & vbNewLine &
                    "3    |4    |5    |3     |40    " & vbNewLine &
                    "4    |5    |6    |4     |104   " & vbNewLine &
                    "5    |6    |7    |5     |      " & vbNewLine &
                    "567  |65527|     |      |      " & vbNewLine &
                    "5    |60   |70   |5     |      " & vbNewLine &
                    "     |     |     |1     |2     " & vbNewLine &
                    "     |     |     |789   |65728 " & vbNewLine &
                    "     |     |     |890   |      " & vbNewLine &
                    ""
            Assert.AreEqual(ShallBeResult, CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(FullOuterJoined))

        End Sub

        Private Function CreateInnerJoinTableSet1() As JoinTableSet
            'Inner join test'
            Dim right As New DataTable
            Dim left As New DataTable

            left.Columns.Add("left1")
            left.Columns.Add("left2")
            left.Columns.Add("left3")
            left.Columns.Add("test")

            right.Columns.Add("right1")
            right.Columns.Add("right2")
            right.Columns.Add("test")

            For j As Integer = 0 To 5
                Dim leftRow As DataRow = left.NewRow
                leftRow.Item(0) = j
                leftRow.Item(1) = j + 1
                leftRow.Item(2) = j + 2
                leftRow.Item(3) = j + 3
                left.Rows.Add(leftRow)
            Next j


            For i As Integer = 0 To 5
                Dim rightrow As DataRow = right.NewRow
                rightrow.Item(0) = i
                rightrow.Item(1) = i + 1
                rightrow.Item(2) = i + 2
                right.Rows.Add(rightrow)
            Next i

            Dim ds As New DataSet
            ds.Tables.Add(left)
            ds.Tables.Add(right)
            Dim relation As New DataRelation("InnerJoined", left.Columns(0), right.Columns(0), True)
            'innerJoined = CompuMaster.Data.DataTables.JoinTables(left, leftColumns, right, rightColumns, CompuMaster.Data.DataTables.JoinTypes.Inner)
            ds.Relations.Add(relation)

            Return New JoinTableSet("CreateInnerJoinTableSet1", left, New String() {"left1"}, right, New String() {"right1"})

        End Function

        <Test()> Public Sub InnerJoinTables()
            Dim TestTables As JoinTableSet = Me.CreateInnerJoinTableSet1
            TestTables.WriteToConsole()

            Dim InnerJoined As DataTable
            InnerJoined = CompuMaster.Data.DataTables.JoinTables(TestTables.LeftTable, TestTables.RightTable, TestTables.LeftTable.DataSet.Relations(0), CompuMaster.Data.DataTables.JoinTypes.Inner)

            Console.WriteLine("INNER-JOINED TABLE CONTENTS: " & TestTables.TableSetName)
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(InnerJoined))

            'Verify column (names)
            For i As Integer = 1 To 3
                Assert.AreEqual("left" & i, InnerJoined.Columns(i - 1).ColumnName)
            Next i
            StringAssert.IsMatch("left1", InnerJoined.Columns(0).ColumnName)
            StringAssert.IsMatch("left2", innerJoined.Columns(1).ColumnName)
            StringAssert.IsMatch("left3", innerJoined.Columns(2).ColumnName)
            StringAssert.IsMatch("test", innerJoined.Columns(3).ColumnName)
            StringAssert.IsMatch("right1", innerJoined.Columns(4).ColumnName)
            StringAssert.IsMatch("right2", innerJoined.Columns(5).ColumnName)
            StringAssert.IsMatch("test", innerJoined.Columns(6).ColumnName)
            Assert.AreEqual(7, InnerJoined.Columns.Count())

            'Verify row (content)
            Assert.AreEqual(6, InnerJoined.Rows.Count())

            StringAssert.IsMatch("3", innerJoined.Rows.Item(0).Item(3))
            StringAssert.IsMatch("2", innerJoined.Rows.Item(0).Item(6))

            For i As Integer = 0 To 5
                Assert.AreEqual(0 + i, CInt(innerJoined.Rows(i)(0)))
                Assert.AreEqual(1 + i, CInt(innerJoined.Rows(i)(1)))
                Assert.AreEqual(2 + i, CInt(innerJoined.Rows(i)(2)))
                Assert.AreEqual(0 + i, CInt(innerJoined.Rows(i)(3)))
                Assert.AreEqual(1 + i, CInt(innerJoined.Rows(i)(4)))
                Assert.AreEqual(2 + i, CInt(innerJoined.Rows(i)(5)))
            Next i

            StringAssert.IsMatch("3", innerJoined.Rows.Item(0).Item("test"))
            StringAssert.IsMatch("2", innerJoined.Rows.Item(0).Item(6))

            Assert.AreEqual("", innerJoined.Rows(1)(2))
            Assert.AreEqual("", innerJoined.Rows(1)(3))



            'Throw New NotImplementedException
            'TODO: implementation of test code for InnerJoinTables -done
            'TODO: implementation of test code for RightJoinTables
            'TODO: implementation of test code for FullJoinTables -done
            'TODO: implementation of test code for CrossJoinTables -done

        End Sub

        Private Function CreateLeftJoinTableSet1() As JoinTableSet
            Dim right As New DataTable
            Dim left As New DataTable

            left.Columns.Add("left1")
            left.Columns.Add("left2")

            right.Columns.Add("right1")
            right.Columns.Add("right2")

            Dim leftRow As DataRow = left.NewRow
            leftRow.Item("left1") = 5
            leftRow.Item("left2") = 10
            left.Rows.Add(leftRow)

            Dim leftRow1 As DataRow = left.NewRow
            leftRow1.Item("left1") = 23
            leftRow1.Item("left2") = 99
            left.Rows.Add(leftRow1)

            Dim rightRow As DataRow = right.NewRow
            rightRow.Item("right1") = 5
            rightRow.Item("right2") = 11
            right.Rows.Add(rightRow)


            Dim rightRow1 As DataRow = right.NewRow
            rightRow1.Item("right1") = 242
            rightRow1.Item("right2") = 2512
            right.Rows.Add(rightRow1)

            Dim rightRow2 As DataRow = right.NewRow
            rightRow2.Item("right1") = 123
            rightRow2.Item("right2") = 1234
            right.Rows.Add(rightRow2)

            left.PrimaryKey = New System.Data.DataColumn() {left.Columns("left1")}
            right.PrimaryKey = New System.Data.DataColumn() {right.Columns("right1")}

            Console.WriteLine("LEFT TABLE CONTENTS")
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(left))
            Console.WriteLine()
            Console.WriteLine("RIGHT TABLE CONTENTS")
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(right))

            Dim ds As New DataSet
            ds.Tables.Add(left)
            ds.Tables.Add(right)

            Dim relation As New DataRelation("LeftJoined", left.Columns(0), right.Columns(0), False)
            ds.Relations.Add(relation)

            Return New JoinTableSet("CreateLeftJoinTableSet1", left, New String() {"left1"}, right, New String() {"right1"})

        End Function

        <Test()> Public Sub LeftJoinTables()
            Dim TestTables As JoinTableSet = Me.CreateLeftJoinTableSet1
            TestTables.WriteToConsole()

            Dim LeftJoined As DataTable
            LeftJoined = CompuMaster.Data.DataTables.JoinTables(TestTables.LeftTable, TestTables.RightTable, TestTables.LeftTable.DataSet.Relations(0), CompuMaster.Data.DataTables.JoinTypes.Left)

            Console.WriteLine("LEFT-JOINED TABLE CONTENTS: " & TestTables.TableSetName)
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(LeftJoined))

            'Verify column (names)
            Assert.AreEqual(4, LeftJoined.Columns.Count())
            StringAssert.IsMatch("left1", LeftJoined.Columns(0).ColumnName)
            StringAssert.IsMatch("left2", LeftJoined.Columns(1).ColumnName)
            StringAssert.IsMatch("right1", LeftJoined.Columns(2).ColumnName)
            StringAssert.IsMatch("right2", LeftJoined.Columns(3).ColumnName)

            'Verify row (content)
            Assert.AreEqual(2, LeftJoined.Rows.Count())
            Assert.IsTrue(IsDBNull(LeftJoined.Rows(1).Item(2)))
            StringAssert.IsMatch("23", LeftJoined.Rows(1).Item(0))
            StringAssert.IsMatch("5", LeftJoined.Rows(0).Item(0))
            StringAssert.IsMatch("11", LeftJoined.Rows(0).Item(3))

            Assert.AreEqual(10, CInt(LeftJoined.Rows(0)(1)))
            Assert.AreEqual(5, CInt(LeftJoined.Rows(0)(2)))
            Assert.AreEqual(99, CInt(LeftJoined.Rows(1)(1)))
        End Sub

        <Test()> Public Sub SqlJoinTables_Left()
            SqlJoinTables_Left_TableSet1()
            SqlJoinTables_Left_TableSet2()
        End Sub

        Private Sub SqlJoinTables_Left_TableSet1()
            Dim TestTables As JoinTableSet = Me.CreateLeftJoinTableSet1
            TestTables.WriteToConsole()

            Dim LeftJoined As DataTable
            LeftJoined = CompuMaster.Data.DataTables.SqlJoinTables(TestTables.LeftTable, TestTables.LeftTable.PrimaryKey, Nothing, TestTables.RightTable, TestTables.RightTable.PrimaryKey, Nothing, CompuMaster.Data.DataTables.SqlJoinTypes.Left)

            Console.WriteLine("LEFT-JOINED TABLE CONTENTS: " & TestTables.TableSetName)
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(LeftJoined))

            'Verify column (names)
            Assert.AreEqual(4, LeftJoined.Columns.Count())
            StringAssert.IsMatch("left1", LeftJoined.Columns(0).ColumnName)
            StringAssert.IsMatch("left2", LeftJoined.Columns(1).ColumnName)
            StringAssert.IsMatch("right1", LeftJoined.Columns(2).ColumnName)
            StringAssert.IsMatch("right2", LeftJoined.Columns(3).ColumnName)

            'Verify row (content)
            Assert.AreEqual(2, LeftJoined.Rows.Count())
            Assert.IsTrue(IsDBNull(LeftJoined.Rows(1).Item(2)))
            StringAssert.IsMatch("23", LeftJoined.Rows(1).Item(0))
            StringAssert.IsMatch("5", LeftJoined.Rows(0).Item(0))
            StringAssert.IsMatch("11", LeftJoined.Rows(0).Item(3))

            Assert.AreEqual(10, CInt(LeftJoined.Rows(0)(1)))
            Assert.AreEqual(5, CInt(LeftJoined.Rows(0)(2)))
            Assert.AreEqual(99, CInt(LeftJoined.Rows(1)(1)))

        End Sub

        <Test> Public Sub SqlJoinTables_ExpectedExceptionNullParameter()
            Assert.Throws(Of ArgumentNullException)(Sub()
                                                        Dim TestTables As JoinTableSet = Me.CreateLeftJoinTableSet1
                                                        CompuMaster.Data.DataTables.SqlJoinTables(TestTables.LeftTable, TestTables.LeftTable.PrimaryKey, Nothing, Nothing, TestTables.RightTable.PrimaryKey, Nothing, CompuMaster.Data.DataTables.SqlJoinTypes.Left)
                                                    End Sub)
        End Sub

        Private Sub SqlJoinTables_Left_TableSet2()
            Dim TestTables As JoinTableSet = Me.CreateFullOuterJoinTableSet2
            TestTables.WriteToConsole()

            Dim LeftJoined As DataTable
            LeftJoined = CompuMaster.Data.DataTables.SqlJoinTables(TestTables.LeftTable, New String() {"left1"}, Nothing, TestTables.RightTable, New String() {"right1"}, Nothing, CompuMaster.Data.DataTables.SqlJoinTypes.Left)

            Console.WriteLine("LEFT-JOINED TABLE CONTENTS: " & TestTables.TableSetName)
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(LeftJoined))

            Dim ShallBeResult As String = "JoinedTable" & vbNewLine &
                    "left1|left2|left3|right1|right2" & vbNewLine &
                    "-----+-----+-----+------+------" & vbNewLine &
                    "0    |1    |2    |0     |1     " & vbNewLine &
                    "     |2    |3    |      |3     " & vbNewLine &
                    "2    |3    |4    |      |      " & vbNewLine &
                    "3    |4    |5    |3     |4     " & vbNewLine &
                    "3    |4    |5    |3     |40    " & vbNewLine &
                    "4    |5    |6    |4     |104   " & vbNewLine &
                    "5    |6    |7    |5     |      " & vbNewLine &
                    "567  |65527|     |      |      " & vbNewLine &
                    "5    |60   |70   |5     |      " & vbNewLine &
                    ""
            Assert.AreEqual(ShallBeResult, CompuMaster.Data.DataTables.ConvertToPlainTextTableFixedColumnWidths(LeftJoined))

        End Sub

        <Test()> Public Sub FindDuplicates()
            Dim testTable As DataTable = _TestTable1()
            Dim Result As Hashtable
            Result = CompuMaster.Data.DataTables.FindDuplicates(testTable.Columns("value"))

            For Each MyItem As DictionaryEntry In Result
                Console.WriteLine(MyItem.Key & "=" & MyItem.Value)
            Next

            Assert.AreEqual(2, Result.Count, "JW #00001")
            For Each MyKey As DictionaryEntry In Result
                If MyKey.Key = "Hello world!" Then
                    Assert.AreEqual(3, MyKey.Value, "JW #00002")
                ElseIf MyKey.Key = "Gotcha!" Then
                    'Gotcha! (not the one in capital letters!)
                    Assert.AreEqual(2, MyKey.Value, "JW #00003")
                Else
                    Assert.Fail("Invalid returned value: " & CType(MyKey.Key, Object).ToString)
                End If
            Next


            Result = CompuMaster.Data.DataTables.FindDuplicates(testTable.Columns("value"), 3)
            Assert.AreEqual(1, Result.Count, "JW #00011")
            For Each MyKey As DictionaryEntry In Result
                Assert.AreEqual("Hello world!", MyKey.Key, "JW #00012")
                Assert.AreEqual(3, MyKey.Value, "JW #00013")
            Next

        End Sub

        <Test()> Public Sub KeepColumnsAndRemoveAllOthers()
            Dim testData As New DataTable
            testData.Columns.Add()
            testData.Columns.Add("test")
            testData.Columns.Add("‰ˆ¸ƒ÷‹ﬂ")
            testData.Columns.Add("data")
            testData.Columns.Add()

            CompuMaster.Data.DataTables.KeepColumnsAndRemoveAllOthers(testData, New String() {"ƒ÷‹‰ˆ¸ﬂ", "data", "istnich", ""})
            Assert.AreEqual(2, testData.Columns.Count)
            Assert.AreEqual("‰ˆ¸ƒ÷‹ﬂ", testData.Columns(0).ColumnName)
            Assert.AreEqual("data", testData.Columns(1).ColumnName)

        End Sub

        <Test()> Public Sub CopyDataTablesCaseSensitive()
            Dim source As New DataTable
            Dim dest As New DataTable
            source.Columns.Add("TestColumn")
            dest.Columns.Add("Testcolumn")
            Dim r As DataRow = source.NewRow
            r.Item("TestColumn") = "test"
            source.Rows.Add(r)


            Dim r2 As DataRow = dest.NewRow
            r2.Item("Testcolumn") = "test"
            dest.Rows.Add(r2)

            'First, some case insensitive tests'
            CompuMaster.Data.DataTables.CreateDataTableClone(source, dest, Nothing, Nothing, 2, False, False, False, False, True)
            StringAssert.IsMatch("test", dest.Rows.Item(0).Item("Testcolumn"))
            Assert.Less(dest.Columns.Count(), 2) 'if 2 then the test failed, because we did a CaseInsensitive comparison'
            StringAssert.IsMatch("Testcolumn", dest.Columns.Item(0).ColumnName)

            'Now case sensitive'
            CompuMaster.Data.DataTables.CreateDataTableClone(source, dest, Nothing, Nothing, 2, False, False, False, False, False)
            Assert.AreEqual(2, dest.Columns.Count()) 'Function shouldn't find "TestColumn" in dest (because there we only have Testcolumn (lowercase 'c'), and therefore add it => 2 columns in table'
            StringAssert.IsMatch("Testcolumn", dest.Columns.Item(0).ColumnName)
            StringAssert.IsMatch("TestColumn", dest.Columns.Item(1).ColumnName)

        End Sub

    End Class

End Namespace