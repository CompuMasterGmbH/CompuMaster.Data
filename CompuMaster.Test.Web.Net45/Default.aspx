<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Default.aspx.vb" Inherits="_Default" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <h1>CompuMaster.Data Test .Net 4.5 with medium trust level (=typical ASP.NET permission set)</h1>
            <h2>Available data providers (<%= Me.AvailableDataProvidersListbox.Items.Count %> data providers)</h2>
            <asp:RadioButtonList runat="server" ID="AvailableDataProvidersListbox" AutoPostBack="true" />
            <p>
                WARNING: Oracle library might not load in medium trust levels (requires full trust or customized trust levels as per 2016-08-04). 
                <ul>
                    <li>If it fails, it might cause the assembly load mechanism to also fail. As a result, all local binaries aren't loaded, but the assemblies from Global Assembly Cache.</li>
                    <li>You might want to play around with the Oracle library to remove it from the bin folder temporary. As soon as the Oracle dll doesn't interrupt the assembly loader, the MySql and Npgsql libraries will be loaded successfully and can be used as expected.</li>
                </ul>
            </p>
            <asp:Label runat="server" ID="AvailableDataProvidersError" ForeColor="Red" />
            <asp:Panel runat="server" ID="FeatureShow" Visible="false">
                <h2>Available features</h2>
                <ul>
                    <li>IDbConnection:
                        <asp:CheckBox runat="server" ID="FeatureInfoConnection" Enabled="false" />
                        <asp:Label runat="server" ID="FeatureInfoConnectionError" ForeColor="Red" /></li>
                    <li>IDbCommand:
                        <asp:CheckBox runat="server" ID="FeatureInfoCommand" Enabled="false" />
                        <asp:Label runat="server" ID="FeatureInfoCommandError" ForeColor="Red" /></li>
                    <li>DbCommandBuilder:
                        <asp:CheckBox runat="server" ID="FeatureInfoCommandBuilder" Enabled="false" />
                        <asp:Label runat="server" ID="FeatureInfoCommandBuilderError" ForeColor="Red" /></li>
                    <li>IDbDataAdapter:
                        <asp:CheckBox runat="server" ID="FeatureInfoDbDataAdapter" Enabled="false" />
                        <asp:Label runat="server" ID="FeatureInfoDbDataAdapterError" ForeColor="Red" /></li>
                </ul>
            </asp:Panel>
            <h2>Loaded assemblies in current AppDomain (<%= Me.LoadedAssembliesList.Items.Count %> assemblies)</h2>
            <ul>
                <asp:BulletedList runat="server" ID="LoadedAssembliesList" />
            </ul>
            <h2>Really good guides to trust levels, etc.</h2>
            <ul>
                <li><a href="http://www.codemag.com/article/0801031" target="_blank">http://www.codemag.com/article/0801031</a></li>
            </ul>
        </div>
    </form>
</body>
</html>
