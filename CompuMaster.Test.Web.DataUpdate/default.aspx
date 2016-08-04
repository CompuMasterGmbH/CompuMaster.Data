<%@ Page Language="VB" %>

<%@ Register TagPrefix="CMDataWeb" Namespace="CompuMaster.Data.Web" Assembly="CompuMaster.Data.Web" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<script runat="server">

    Function JitCreateAccessDbOrCreateDbConnectionString() As String
        'e.g. "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=J:\temp\test.mdb;User Id=admin;Password=;"
        Dim TestDbPath As String = Server.MapPath("app_data/test.mdb")
        If System.IO.File.Exists(TestDbPath) = False Then
            CompuMaster.Data.DatabaseManagement.CreateDatabaseFile(TestDbPath)
            Dim InitialContentDbConn As System.Data.IDbConnection = CompuMaster.Data.DataQuery.Connections.MicrosoftAccessConnection(TestDbPath)
            Dim Result As String
            Result = InitialContentDbConn.ConnectionString
            CompuMaster.Data.Manipulation.WriteDataTableToDataConnection(Me.InitialTableDataInDb, InitialContentDbConn, CompuMaster.Data.Manipulation.DdlLanguage.MSJetEngine, True, CompuMaster.Data.DataQuery.AnyIDataProvider.Automations.AutoOpenAndCloseAndDisposeConnection)
            'Dim MyCmd As System.Data.IDbCommand = InitialContentDbConn.CreateCommand()
            'MyCmd.CommandText = "SELECT 'text' AS What2Setup, 'demo' AS Value2Setup, True AS TranslationRequired, 1 AS ID, 0 AS LangID, 1000 AS SortID"
            'CompuMaster.Data.DataQuery.ExecuteNonQuery(MyCmd, CompuMaster.Data.DataQuery.AnyIDataProvider.Automations.AutoOpenAndCloseAndDisposeConnection)
            'Return Result
        End If
        Return CompuMaster.Data.DataQuery.Connections.MicrosoftAccessConnection(TestDbPath).ConnectionString
    End Function

    Private Function InitialTableDataInDb() As System.Data.DataTable
        Dim Result As New System.Data.DataTable("Internationalizations")
        Result.Columns.Add("ID", GetType(Integer))
        Result.Columns("ID").Unique = True
        Result.Columns("ID").AutoIncrement = True
        Result.PrimaryKey = New System.Data.DataColumn() {Result.Columns("ID")}
        Result.Columns.Add("What2Setup", GetType(String))
        Result.Columns.Add("Value2Setup", GetType(String))
        Result.Columns.Add("TranslationRequired", GetType(Boolean))
        Result.Columns.Add("LangID", GetType(Integer))
        Result.Columns.Add("SortID", GetType(Integer))
        Dim Row As System.Data.DataRow
        Row = Result.NewRow
        Row("ID") = 1
        Row("What2Setup") = "text"
        Row("Value2Setup") = "DEMO"
        Row("TranslationRequired") = True
        Row("LangID") = 0
        Row("SortID") = 1000
        Result.Rows.Add(Row)
        Row = Result.NewRow
        Row("ID") = 2
        Row("What2Setup") = "text"
        Row("Value2Setup") = "demo"
        Row("TranslationRequired") = True
        Row("LangID") = 3
        Row("SortID") = 1000
        Result.Rows.Add(Row)
        Return Result
    End Function

    Sub ValueInits(sender As Object, e As System.EventArgs) Handles MyBase.Init
        InitErrorInfo.Text = ""
        Try
            Me.TextBoxQueryForManipulationViaQuickEdit.Text = "SELECT What2Setup, [Value2Setup], [TranslationRequired], [ID] FROM [Internationalizations] WHERE LangID = 3 AND TranslationRequired <> False ORDER BY [SortID], [What2Setup]"
            Me.DropDownListDataProvider.SelectedValue = "OleDb"
            Me.TextBoxConnectionString.Text = JitCreateAccessDbOrCreateDbConnectionString()
        Catch ex As Exception
            InitErrorInfo.Text = ex.ToString.Replace(vbNewLine, "<br />")
        End Try
    End Sub

    Sub AssignValues(sender As Object, e As System.EventArgs) Handles MyBase.InitComplete
        GridEditView.DataConnectionString = Me.TextBoxConnectionString.Text
        GridEditView.DataProviderName = Me.DropDownListDataProvider.SelectedValue
        GridEditView.DataSelectCommand = Me.TextBoxQueryForManipulationViaQuickEdit.Text
    End Sub

    Sub StatusInfo(sender As Object, e As EventArgs) Handles MyBase.PreRender
        Me.ExecutionReport.Text = GridEditView.UpdateStatus
    End Sub

    Protected Overrides Sub OnLoad(e As EventArgs)
        Try
            MyBase.OnLoad(e)
        Catch ex As Exception
            If InitErrorInfo.Text <> "" Then InitErrorInfo.Text &= "<br /><br />"
            InitErrorInfo.Text = ex.ToString.Replace(vbNewLine, "<br />")
        End Try
    End Sub

    Private Sub GridEditViewDataLoadExceptionCheck(sender As Object, e As System.EventArgs) Handles MyBase.LoadComplete
        Try
            GridEditView.DataLoadExceptionCheck()
        Catch ex As Exception
            If InitErrorInfo.Text <> "" Then InitErrorInfo.Text &= "<br /><br />"
            InitErrorInfo.Text = ex.ToString.Replace(vbNewLine, "<br />")
        End Try
    End Sub

    Sub BindData(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            GridEditView.DataLoad()
            If GridEditView.DataLoadException IsNot Nothing Then
                AccessDataSource1.SelectCommand = Nothing
                GridView1.Enabled = False
            End If
            GridEditView.DataLoadExceptionCheck()
            Dim sqlCmdBuilder As New System.Data.OleDb.OleDbCommandBuilder(GridEditView.QuickEditDataContainer.DataAdapter)
            AccessDataSource1.UpdateCommand = sqlCmdBuilder.GetUpdateCommand().CommandText
        Catch ex As Exception
            If InitErrorInfo.Text <> "" Then InitErrorInfo.Text &= "<br /><br />"
            InitErrorInfo.Text = ex.ToString.Replace(vbNewLine, "<br />")
        End Try
    End Sub

</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>CompuMaster.Data.Manipulation.ByCode - WebEdition</title>
    <style type="text/css">
        BODY {
            font-family: Arial;
        }
    </style>
</head>
<body>
    <form id="FormDataManipulationByQuickEdit" runat="server">
        <h1>Live Data Manipulation - Demo</h1>
        <asp:Label runat="server" ID="InitErrorInfo" ForeColor="Red" />
        <h2>Datasource provider</h2>
        <asp:DropDownList ID="DropDownListDataProvider" runat="server">
            <asp:ListItem Selected="True">SqlClient</asp:ListItem>
            <asp:ListItem>OleDb</asp:ListItem>
            <asp:ListItem>ODBC</asp:ListItem>
        </asp:DropDownList>
        <br />
        <h2>ConnectionString</h2>
        <asp:TextBox ID="TextBoxConnectionString" runat="server" Height="50px" TextMode="SingleLine" Text="SERVER=localhost;DATABASE=master;PWD=xxxxxxxxx;UID=sa" Width="546px"></asp:TextBox>
        <br />
        <h2>SQL Select statement</h2>
        <asp:TextBox ID="TextBoxQueryForManipulationViaQuickEdit" runat="server" Height="50px" TextMode="MultiLine"
            Width="546px" Text="SELECT * FROM [dbo].[myTable]"></asp:TextBox>
        <br />
        <h2>Edit grid</h2>
        <CMDataWeb:GridEditQuery ID="GridEditView" runat="server" AllowPaging="True"
            AllowSorting="True" AutoGenerateColumns="False" DataKeyNames="ID">
            <Columns>
                <asp:BoundField DataField="ID" HeaderText="ID" InsertVisible="False"
                    ReadOnly="True" SortExpression="ID" />
                <asp:BoundField DataField="What2Setup" HeaderText="What2Setup"
                    ReadOnly="True" SortExpression="What2Setup" />
                <asp:BoundField DataField="Value2Setup" HeaderText="Value2Setup"
                    SortExpression="Value2Setup" />
                <asp:CheckBoxField DataField="TranslationRequired"
                    HeaderText="TranslationRequired" SortExpression="TranslationRequired" />
                <asp:CommandField ShowEditButton="True" />
            </Columns>
        </CMDataWeb:GridEditQuery>
        <br />
        <br />
        <h2>Preview data</h2>
        <asp:GridView ID="GridViewPreviewData" runat="server" EnableViewState="False">
        </asp:GridView>
        <br />
        <asp:RadioButtonList ID="RadioButtonListPreviewMode" runat="server">
            <asp:ListItem Selected="True" Value="True">Preview data!</asp:ListItem>
            <asp:ListItem Value="False">Execute now!</asp:ListItem>
        </asp:RadioButtonList>
        <br />
        <br />
        <h2>Execution Report</h2>
        <asp:TextBox ID="ExecutionReport" runat="server" Height="327px" Width="546px" TextMode="MultiLine" />
        <br />
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False"
            DataKeyNames="ID" DataSourceID="AccessDataSource1">
            <Columns>
                <asp:BoundField DataField="ID" HeaderText="ID" InsertVisible="False"
                    ReadOnly="True" SortExpression="ID" />
                <asp:BoundField DataField="What2Setup" HeaderText="What2Setup"
                    SortExpression="What2Setup" />
                <asp:CheckBoxField DataField="TranslationRequired"
                    HeaderText="TranslationRequired" SortExpression="TranslationRequired" />
                <asp:BoundField DataField="Value2Setup" HeaderText="Value2Setup"
                    SortExpression="Value2Setup" />
                <asp:CommandField ShowEditButton="True" />
            </Columns>
        </asp:GridView>
        <asp:AccessDataSource ID="AccessDataSource1" runat="server"
            DataFile="~/app_data/test.mdb"
            SelectCommand="SELECT What2Setup, [Value2Setup], [TranslationRequired], [ID] FROM [Internationalizations] WHERE LangID = 3 AND TranslationRequired <> False ORDER BY [SortID], [What2Setup]"></asp:AccessDataSource>
        <asp:ObjectDataSource ID="ObjectDataSource1" runat="server"></asp:ObjectDataSource>

    </form>
</body>
</html>
