Public Class TestForm

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        LoadQuery("")
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        LoadQuery("indepyear < 1900")
    End Sub

    Sub LoadQuery(ByVal filter As String)
        Dim sql As String = "SELECT * FROM country"
        If filter <> Nothing Then sql &= " WHERE " & filter
        sql &= ";"
        'Dim MyCmd As New SqlClient.SqlCommand(sql, New SqlClient.SqlConnection(Me.TextBoxConnectionString.Text))
        Dim MyCmd As New Npgsql.NpgsqlCommand(sql, New Npgsql.NpgsqlConnection(Me.TextBoxConnectionString.Text))
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
