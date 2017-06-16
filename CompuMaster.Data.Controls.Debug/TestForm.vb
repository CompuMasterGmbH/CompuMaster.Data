Public Class TestForm

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.Text &= " (" & CompuMaster.Data.DataQuery.PlatformTools.CurrentClrRuntime.ToString & ")"
        Me.TextBoxConnectionString.Text = CreateNewDbConnection.ConnectionString
        LoadQuery("")
    End Sub

    Private Function CreateNewDbConnection() As IDbConnection
        Return CompuMaster.Data.DataQuery.Connections.MicrosoftAccessOleDbConnection(System.IO.Path.Combine(My.Application.Info.DirectoryPath, "app_data\country.mdb"))
    End Function

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        LoadQuery("indepyear < 1900")
    End Sub

    Sub LoadQuery(ByVal filter As String)
        Dim sql As String = "SELECT * FROM [country]"
        If filter <> Nothing Then sql &= " WHERE " & filter
        sql &= ";"
        Dim MyConn As IDbConnection = Me.CreateNewDbConnection
        Dim MyCmd As IDbCommand = MyConn.CreateCommand
        MyCmd.CommandText = sql
        Me.DataGridView1.SelectCommand = MyCmd
        Me.DataGridView1.LoadData()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        LoadQuery("indepyear >= 1900")
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        LoadQuery("")
    End Sub

    Private Sub AdditionalTestEntryToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AdditionalTestEntryToolStripMenuItem.Click
        MsgBox("Context menu item AdditionalTestEntryToolStripMenuItem clicked." & vbNewLine & "Test successful if content menu item itself should be visible as well as its text property", MsgBoxStyle.Information)
    End Sub

End Class
