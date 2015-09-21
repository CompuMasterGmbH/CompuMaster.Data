Imports NUnit.Framework

Namespace CompuMaster.Test.Data.DataQuery

    <TestFixture()> Public Class DataQueryAnyIDataProvider

        <Test()> Public Sub ExecuteReaderAndPutFirstColumnIntoGenericList()
            Dim TestFile As String = System.IO.Path.Combine(System.Environment.CurrentDirectory, "testfiles\test_for_msaccess.mdb")
            Dim MyConn As IDbConnection = CompuMaster.Data.DataQuery.Connections.MicrosoftAccessConnection(TestFile)
            Dim MyCmd As IDbCommand = MyConn.CreateCommand()
            MyCmd.CommandType = CommandType.Text
            MyCmd.CommandText = "SELECT IntegerLongValue FROM [SeveralColumnTypesTest] ORDER BY ID"
            Dim IntList As System.Collections.Generic.List(Of Integer) = CompuMaster.Data.DataQuery.AnyIDataProvider.ExecuteReaderAndPutFirstColumnIntoGenericList(Of Integer)(MyCmd, CompuMaster.Data.DataQuery.Automations.AutoOpenAndCloseAndDisposeConnection)
            Assert.AreEqual(123456789, IntList(0), "Row content check IntegerLongValue column")
            Assert.AreEqual(987654321, IntList(1), "Row content check IntegerLongValue column")
            Assert.AreEqual(0, IntList(2), "Row content check IntegerLongValue column")
            Assert.AreEqual(0, IntList(3), "Row content check IntegerLongValue column")
            Assert.AreEqual(Nothing, IntList(4), "Row content check IntegerLongValue column")
        End Sub

        <Test()> Public Sub ExecuteReaderAndPutFirstColumnIntoGenericNullableList()
            Dim TestFile As String = System.IO.Path.Combine(System.Environment.CurrentDirectory, "testfiles\test_for_msaccess.mdb")
            Dim MyConn As IDbConnection = CompuMaster.Data.DataQuery.Connections.MicrosoftAccessConnection(TestFile)
            Dim MyCmd As IDbCommand = MyConn.CreateCommand()
            MyCmd.CommandType = CommandType.Text
            MyCmd.CommandText = "SELECT IntegerLongValue FROM [SeveralColumnTypesTest] ORDER BY ID"
            Dim IntList As System.Collections.Generic.List(Of Integer?) = CompuMaster.Data.DataQuery.AnyIDataProvider.ExecuteReaderAndPutFirstColumnIntoGenericNullableList(Of Integer)(MyCmd, CompuMaster.Data.DataQuery.Automations.AutoOpenAndCloseAndDisposeConnection)
            Assert.AreEqual(123456789, IntList(0), "Row content check IntegerLongValue column")
            Assert.AreEqual(987654321, IntList(1), "Row content check IntegerLongValue column")
            Assert.AreEqual(0, IntList(2), "Row content check IntegerLongValue column")
            Assert.AreEqual(True, IntList(2).HasValue, "Row content check IntegerLongValue column")
            Assert.AreEqual(0, IntList(3), "Row content check IntegerLongValue column")
            Assert.AreEqual(True, IntList(3).HasValue, "Row content check IntegerLongValue column")
            Assert.AreEqual(Nothing, IntList(4), "Row content check IntegerLongValue column")
            Assert.AreEqual(False, IntList(4).HasValue, "Row content check IntegerLongValue column") 'Access DBs don't know a NULLABLE BOOLEAN --> it's a FALSE value instead of a DbNull value
        End Sub

        <Test()> Public Sub ExecuteReaderAndPutFirstTwoColumnsIntoGenericNullableKeyValuePairs()
            Dim TestFile As String = System.IO.Path.Combine(System.Environment.CurrentDirectory, "testfiles\test_for_msaccess.mdb")
            Dim MyConn As IDbConnection = CompuMaster.Data.DataQuery.Connections.MicrosoftAccessConnection(TestFile)
            Dim MyCmd As IDbCommand = MyConn.CreateCommand()
            MyCmd.CommandType = CommandType.Text
            MyCmd.CommandText = "SELECT IntegerLongValue, BooleanValue FROM [SeveralColumnTypesTest] ORDER BY ID"
            Dim IntegerStringDictionary As System.Collections.Generic.List(Of System.Collections.Generic.KeyValuePair(Of Integer?, Boolean?)) = CompuMaster.Data.DataQuery.AnyIDataProvider.ExecuteReaderAndPutFirstTwoColumnsIntoGenericNullableKeyValuePairs(Of Integer, Boolean)(MyCmd, CompuMaster.Data.DataQuery.Automations.AutoOpenAndCloseAndDisposeConnection)
            Assert.AreEqual(123456789, IntegerStringDictionary(0).Key, "Row content check IntegerLongValue column")
            Assert.AreEqual(987654321, IntegerStringDictionary(1).Key, "Row content check IntegerLongValue column")
            Assert.AreEqual(0, IntegerStringDictionary(2).Key, "Row content check IntegerLongValue column")
            Assert.AreEqual(True, IntegerStringDictionary(2).Key.HasValue, "Row content check IntegerLongValue column")
            Assert.AreEqual(0, IntegerStringDictionary(3).Key, "Row content check IntegerLongValue column")
            Assert.AreEqual(True, IntegerStringDictionary(3).Key.HasValue, "Row content check IntegerLongValue column")
            Assert.AreEqual(Nothing, IntegerStringDictionary(4).Key, "Row content check IntegerLongValue column")
            Assert.AreEqual(False, IntegerStringDictionary(4).Key.HasValue, "Row content check IntegerLongValue column")
            Assert.AreEqual(False, IntegerStringDictionary(0).Value, "Row content check BooleanValue column")
            Assert.AreEqual(True, IntegerStringDictionary(1).Value, "Row content check BooleanValue column")
            Assert.AreEqual(False, IntegerStringDictionary(2).Value, "Row content check BooleanValue column")
            Assert.AreEqual(True, IntegerStringDictionary(3).Value, "Row content check BooleanValue column")
            Assert.AreEqual(True, IntegerStringDictionary(3).Value.HasValue, "Row content check BooleanValue column")
            Assert.AreEqual(False, IntegerStringDictionary(4).Value, "Row content check BooleanValue column") 'Access DBs don't know a NULLABLE BOOLEAN --> it's a FALSE value instead of a DbNull value
            Assert.AreEqual(True, IntegerStringDictionary(4).Value.HasValue, "Row content check BooleanValue column") 'Access DBs don't know a NULLABLE BOOLEAN --> it's a FALSE value PRESENT
        End Sub

        <Test()> Public Sub ExecuteReaderAndPutFirstTwoColumnsIntoGenericKeyValuePairs()
            Dim TestFile As String = System.IO.Path.Combine(System.Environment.CurrentDirectory, "testfiles\test_for_msaccess.mdb")
            Dim MyConn As IDbConnection = CompuMaster.Data.DataQuery.Connections.MicrosoftAccessConnection(TestFile)
            Dim MyCmd As IDbCommand = MyConn.CreateCommand()
            MyCmd.CommandType = CommandType.Text
            MyCmd.CommandText = "SELECT IntegerLongValue, StringShort, StringMemo FROM [SeveralColumnTypesTest] ORDER BY ID"
            Dim IntegerStringDictionary As System.Collections.Generic.List(Of System.Collections.Generic.KeyValuePair(Of Integer, String)) = CompuMaster.Data.DataQuery.AnyIDataProvider.ExecuteReaderAndPutFirstTwoColumnsIntoGenericKeyValuePairs(Of Integer, String)(MyCmd, CompuMaster.Data.DataQuery.Automations.AutoOpenAndCloseAndDisposeConnection)
            Dim key As Integer = IntegerStringDictionary(0).Key
            Assert.AreEqual(123456789, IntegerStringDictionary(0).Key, "Row content check IntegerLongValue column")
            Assert.AreEqual(987654321, IntegerStringDictionary(1).Key, "Row content check IntegerLongValue column")
            Assert.AreEqual(0, IntegerStringDictionary(2).Key, "Row content check IntegerLongValue column")
            Assert.AreEqual(0, IntegerStringDictionary(3).Key, "Row content check IntegerLongValue column")
            Assert.AreEqual(Nothing, IntegerStringDictionary(4).Key, "Row content check IntegerLongValue column")
            Assert.AreEqual("text short 1", IntegerStringDictionary(0).Value, "Row content check StringShortValue column")
            Assert.AreEqual("text short 2", IntegerStringDictionary(1).Value, "Row content check StringShortValue column")
            Assert.AreEqual("text short 3", IntegerStringDictionary(2).Value, "Row content check StringShortValue column")
            Assert.AreEqual("text short 4", IntegerStringDictionary(3).Value, "Row content check StringShortValue column")
            Assert.AreEqual(Nothing, IntegerStringDictionary(4).Value, "Row content check StringShortValue column")
        End Sub

        <Test()> Public Sub ExecuteReaderAndPutFirstTwoColumnsIntoGenericNullableDictionary()
            Dim TestFile As String = System.IO.Path.Combine(System.Environment.CurrentDirectory, "testfiles\test_for_msaccess.mdb")
            Dim MyConn As IDbConnection = CompuMaster.Data.DataQuery.Connections.MicrosoftAccessConnection(TestFile)
            Dim MyCmd As IDbCommand = MyConn.CreateCommand()
            MyCmd.CommandType = CommandType.Text
            MyCmd.CommandText = "SELECT IntegerLongValue, BooleanValue FROM [SeveralColumnTypesTest] ORDER BY ID"
            Dim IntegerStringDictionary As System.Collections.Generic.Dictionary(Of Integer, Boolean?) = CompuMaster.Data.DataQuery.AnyIDataProvider.ExecuteReaderAndPutFirstTwoColumnsIntoGenericNullableDictionary(Of Integer, Boolean)(MyCmd, CompuMaster.Data.DataQuery.Automations.AutoOpenAndCloseAndDisposeConnection)
            Dim Keys As New System.Collections.Generic.List(Of Integer)
            For Each value As Integer In IntegerStringDictionary.Keys
                Keys.Add(value)
            Next
            Assert.AreEqual(123456789, Keys(0), "Row content check IntegerLongValue column")
            Assert.AreEqual(987654321, Keys(1), "Row content check IntegerLongValue column")
            Assert.AreEqual(0, Keys(2), "Row content check IntegerLongValue column")
            Assert.AreEqual(0, Keys(2), "Row content check IntegerLongValue column")
            Assert.AreEqual(False, IntegerStringDictionary(Keys(0)), "Row content check BooleanValue column")
            Assert.AreEqual(True, IntegerStringDictionary(Keys(1)), "Row content check BooleanValue column")
            Assert.AreEqual(False, IntegerStringDictionary(Keys(2)), "Row content check BooleanValue column")
            Assert.AreEqual(3, Keys.Count, "Keys count")
            Assert.AreEqual(3, IntegerStringDictionary.Count, "Records count")
        End Sub

        <Test()> Public Sub ExecuteReaderAndPutFirstTwoColumnsIntoGenericDictionary()
            Dim TestFile As String = System.IO.Path.Combine(System.Environment.CurrentDirectory, "testfiles\test_for_msaccess.mdb")
            Dim MyConn As IDbConnection = CompuMaster.Data.DataQuery.Connections.MicrosoftAccessConnection(TestFile)
            Dim MyCmd As IDbCommand = MyConn.CreateCommand()
            MyCmd.CommandType = CommandType.Text
            MyCmd.CommandText = "SELECT IntegerLongValue, StringShort, StringMemo FROM [SeveralColumnTypesTest] ORDER BY ID"
            Dim IntegerStringDictionary As System.Collections.Generic.Dictionary(Of Integer, String) = CompuMaster.Data.DataQuery.AnyIDataProvider.ExecuteReaderAndPutFirstTwoColumnsIntoGenericDictionary(Of Integer, String)(MyCmd, CompuMaster.Data.DataQuery.Automations.AutoOpenAndCloseAndDisposeConnection)
            Dim Keys As New System.Collections.Generic.List(Of Integer)
            For Each value As Integer In IntegerStringDictionary.Keys
                Keys.Add(value)
            Next
            Assert.AreEqual(123456789, Keys(0), "Row content check IntegerLongValue column")
            Assert.AreEqual("text short 1", IntegerStringDictionary(Keys(0)), "Row content check StringShortValue column")
            Assert.AreEqual(987654321, Keys(1), "Row content check IntegerLongValue column")
            Assert.AreEqual("text short 2", IntegerStringDictionary(Keys(1)), "Row content check StringShortValue column")
            Assert.AreEqual(0, Keys(2), "Row content check IntegerLongValue column")
            Assert.AreEqual(Nothing, IntegerStringDictionary(Keys(2)), "Row content check StringShortValue column")
            Assert.AreEqual(3, Keys.Count, "Keys count")
            Assert.AreEqual(3, IntegerStringDictionary.Count, "Records count")
        End Sub

    End Class

End Namespace