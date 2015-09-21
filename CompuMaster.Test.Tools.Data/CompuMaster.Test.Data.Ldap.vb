Imports NUnit.Framework

Namespace CompuMaster.Test.Data

    <TestFixture()> Public Class Ldap

        <Test()> Public Sub Query()
            Dim testTable As DataTable = CompuMaster.Data.Ldap.Query("compumaster", "(objectCategory=user)")
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTable(testTable))
            testTable = CompuMaster.Data.Ldap.Query("CN=Jochen Wezel,CN=Users,DC=lan,DC=compumaster,DC=de", "(objectCategory=user)")
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTable(testTable))
        End Sub

    End Class


    <NUnit.Framework.TestFixture> Public Class MiniTests

        <NUnit.Framework.Test> Public Sub TestIsStringWithA2ZOnly()
            Assert.AreEqual(True, IsStringWithA2ZOnly("akbkDED"))
            Assert.AreEqual(False, IsStringWithA2ZOnly("akbkDED "))
            Assert.AreEqual(False, IsStringWithA2ZOnly("akbküDED"))
        End Sub

        Private Function IsStringWithA2ZOnly(value As String) As Boolean
            Dim pattern As String = "^[a-zA-Z]+$"
            Dim reg As New System.Text.RegularExpressions.Regex(pattern)
            Return reg.IsMatch(value)
        End Function

    End Class

End Namespace