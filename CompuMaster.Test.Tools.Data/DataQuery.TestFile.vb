Imports NUnit.Framework
Imports System.Collections.Generic

Namespace CompuMaster.Test.Data.DataQuery

    <TestFixture(Category:="DataQueryTestFile")> Public Class DataQueryTestFile

        Private Shared Sub CreateAndDispose(ByVal testFileType As CompuMaster.Data.DataQuery.TestFile.TestFileType)
            Dim testFile As New CompuMaster.Data.DataQuery.TestFile(testFileType)
            Assert.IsTrue(System.IO.File.Exists(testFile.FilePath), "File exists")
            Dim filepath As String = testFile.FilePath
            testFile.Dispose()
            Assert.IsFalse(System.IO.File.Exists(filepath), "File was deleted")

            Dim cachedPath As String = Nothing
            Using tsFile As New CompuMaster.Data.DataQuery.TestFile(testFileType)
                Assert.IsTrue(System.IO.File.Exists(tsFile.FilePath), "File exists")
                cachedPath = tsFile.FilePath
            End Using
            Assert.IsFalse(System.IO.File.Exists(cachedPath), "File was deleted")
        End Sub

        '        <Test> Public Sub CreateAndDisposeMsAccess()
        '#Disable Warning BC40000 ' Typ oder Element ist veraltet
        '            CreateAndDispose(CompuMaster.Data.DataQuery.TestFile.TestFileType.MsAccess)
        '#Enable Warning BC40000 ' Typ oder Element ist veraltet
        '        End Sub

        <Test> Public Sub CreateAndDisposeMsAccessAccdb()
            CreateAndDispose(CompuMaster.Data.DataQuery.TestFile.TestFileType.MsAccessAccdb)
        End Sub

        <Test> Public Sub CreateAndDisposeMsAccessMdb()
            CreateAndDispose(CompuMaster.Data.DataQuery.TestFile.TestFileType.MsAccessMdb)
        End Sub

        <Test> Public Sub CreateAndDisposeMsExcel95Xls()
            CreateAndDispose(CompuMaster.Data.DataQuery.TestFile.TestFileType.MsExcel95Xls)
        End Sub

        <Test> Public Sub CreateAndDisposeMsExcel2007Xlsx()
            CreateAndDispose(CompuMaster.Data.DataQuery.TestFile.TestFileType.MsExcel2007Xlsx)
        End Sub


    End Class

End Namespace
