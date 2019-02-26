Imports NUnit.Framework

Namespace CompuMaster.Test.Data

    <TestFixture(Category:="XLS Reader")> Public Class XlsReader

#If Not CI_Build Then
        <Test()> Public Sub ReadLastCell()
            Dim TestFile As String = AssemblyTestEnvironment.TestFileAbsolutePath("testfiles\test_for_lastcell_e50aka95.xls")

            'Read and compare written test data
            '==================================

            'read the existing file, auto-detect column-types, take datatable and compare it with the written data: it should be always the same (or must be argumented and discussed with Jochen why it isn't)
            'the number of columns and rows should be always 2
            Dim ReReadData As DataTable
            ReReadData = CompuMaster.Data.XlsReader.ReadDataTableFromXlsFile(TestFile, "test")
            Assert.AreEqual(0, ReReadData.Rows.Count, "SaveAndReadEmptyStates #10") 'because last 4 lines only contains DBNull/nothing/empty string values
            Assert.AreEqual(1, ReReadData.Columns.Count, "SaveAndReadEmptyStates #11") 'but the column "string" has been defined by the column header
        End Sub

        <Test()> Public Sub ReadTestFileQnA()
            Dim file As String = AssemblyTestEnvironment.TestFileAbsolutePath("testfiles\Q&A.xls")
            Dim dt As DataTable = CompuMaster.Data.XlsReader.ReadDataTableFromXlsFile(file, "Rund um das NT")
            Assert.AreEqual(35, dt.Rows.Count, "Row-Length")
        End Sub

        <Test()> Public Sub ReadFormatExcel95()
            Dim TestFile As String = AssemblyTestEnvironment.TestFileAbsolutePath("testfiles\test_for_lastcell_e50aka95.xls")
            CompuMaster.Data.XlsReader.ReadDataTableFromXlsFile(TestFile, "test")
        End Sub

        <Test()> Public Sub ReadFormatExcel97()
            Dim TestFile As String = AssemblyTestEnvironment.TestFileAbsolutePath("testfiles\test_for_lastcell_e70aka97-2003.xls")
            CompuMaster.Data.XlsReader.ReadDataTableFromXlsFile(TestFile, "test")
        End Sub

        <Test()> Public Sub ReadFormatExcel2007xlsx()
            Dim TestFile As String = AssemblyTestEnvironment.TestFileAbsolutePath("testfiles\test_for_lastcell_e12aka2007.xlsx")
            CompuMaster.Data.XlsReader.ReadDataTableFromXlsFile(TestFile, "test")
        End Sub

        <Test()> Public Sub ReadFormatExcel2007xlsb()
            Dim TestFile As String = AssemblyTestEnvironment.TestFileAbsolutePath("testfiles\test_for_lastcell_e12aka2007.xlsb")
            CompuMaster.Data.XlsReader.ReadDataTableFromXlsFile(TestFile, "test")
        End Sub

        <Test()> Public Sub ReadFormatExcel2007xlsm()
            Dim TestFile As String = AssemblyTestEnvironment.TestFileAbsolutePath("testfiles\test_for_lastcell_e12aka2007.xlsm")
            CompuMaster.Data.XlsReader.ReadDataTableFromXlsFile(TestFile, "test")
        End Sub
#End If

    End Class

End Namespace