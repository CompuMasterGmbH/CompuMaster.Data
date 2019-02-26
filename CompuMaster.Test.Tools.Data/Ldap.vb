Imports NUnit.Framework

Namespace CompuMaster.Test.Data

    <TestFixture(Category:="LDAP with security")> Public Class Ldap

        <Test()> Public Sub CurrentRootDomain()
            Console.WriteLine("First domain in forest=" & CompuMaster.Data.Ldap.GetRootDomain)
        End Sub

        <Test()> Public Sub CurrentDomains()
            Console.WriteLine("Domains in current forest:" & vbNewLine & Strings.Join(CompuMaster.Data.Ldap.GetDomains, vbNewLine))
        End Sub

#If Not CI_Build Then
        <Test()> Public Sub Query()
            Dim testTable As DataTable = CompuMaster.Data.Ldap.Query("compumaster", "(objectCategory=user)")
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTable(testTable))
            Assert.Greater(testTable.Rows.Count, 1)
            testTable = CompuMaster.Data.Ldap.Query("CN=Jochen Wezel,OU=Emmelshausen,OU=Users - CompuMaster,DC=lan,DC=compumaster,DC=de", "(objectCategory=user)")
            Console.WriteLine(CompuMaster.Data.DataTables.ConvertToPlainTextTable(testTable))
            Assert.AreEqual(testTable.Rows.Count, 1)
        End Sub
#End If

    End Class

    <TestFixture(Category:="LDAP with security", Ignore:="Required custom user credentials")> Class LdapWithSecurity

        <Test, Category("LDAP")> Public Sub QueryMoreThan1000Entries()
            Dim RecordCount As Integer = CompuMaster.Data.Ldap.QueryRecordCount("yourdomain.com", "(objectCategory=user)", "yourdomain\user", "yourpassword")
            Console.WriteLine(RecordCount)
            Assert.Greater(RecordCount, 1000)
            Dim testTable As DataTable = CompuMaster.Data.Ldap.Query("yourdomain.com", "(objectCategory=user)", "yourdomain\user", "yourpassword")
            Assert.Greater(testTable.Rows.Count, 1000)
            testTable = CompuMaster.Data.Ldap.Query("CN=Users,DC=yourdomain,DC=com", "(objectCategory=user)")
            Assert.Greater(testTable.Rows.Count, 1000)
        End Sub

    End Class


    <NUnit.Framework.TestFixture(Category:="LDAP with security")> Public Class MiniTests

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