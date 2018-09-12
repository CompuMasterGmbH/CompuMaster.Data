Imports NUnit.Framework
Imports System.Collections.Generic

Namespace CompuMaster.Test.Data.DataQuery

    <TestFixture(Category:="DataQueryTestFile")> Public Class DataQueryTestFile

        Private Sub CreateAndDispose(ByVal testFileType As CompuMaster.Data.DataQuery.TestFile.TestFileType)
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
        <Test> Public Sub CreateAndDisposeMsAccess()
            CreateAndDispose(CompuMaster.Data.DataQuery.TestFile.TestFileType.MsAccess)
        End Sub

        <Test> Public Sub CreateAndDisposeMsExcel95Xls()
            CreateAndDispose(CompuMaster.Data.DataQuery.TestFile.TestFileType.MsExcel95Xls)
        End Sub

        <Test> Public Sub CreateAndDisposeMsExcel2007Xlsx()
            CreateAndDispose(CompuMaster.Data.DataQuery.TestFile.TestFileType.MsExcel2007Xlsx)
        End Sub


    End Class

End Namespace
