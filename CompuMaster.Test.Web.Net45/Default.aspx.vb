
Partial Class _Default
    Inherits System.Web.UI.Page

    Private Sub _Default_Init(sender As Object, e As EventArgs) Handles Me.Init
        Try
            Me.AvailableDataProvidersListbox.Items.Clear()
            For Each dataProvider As CompuMaster.Data.DataQuery.DataProvider In AvailableDataProvidersInCurrentAppDomain
                Me.AvailableDataProvidersListbox.Items.Add(New ListItem(dataProvider.Title))
            Next
            Me.AvailableDataProvidersError.Text = ""
        Catch ex As Exception
            Me.AvailableDataProvidersError.Text = "ERROR: " & ex.ToString.Replace(vbNewLine, "<br />")
        End Try
    End Sub

    Private Sub _Default_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.LoadedAssembliesList.Items.Clear()
        Dim AlreadyLoadedAssemblies As System.Reflection.Assembly() = AppDomain.CurrentDomain.GetAssemblies
        For Each asm As System.Reflection.Assembly In AlreadyLoadedAssemblies
            Me.LoadedAssembliesList.Items.Add(New ListItem(asm.FullName))
        Next
    End Sub

    Private ReadOnly Property AvailableDataProvidersInCurrentAppDomain As System.Collections.Generic.List(Of CompuMaster.Data.DataQuery.DataProvider)
        Get
            Static Result As System.Collections.Generic.List(Of CompuMaster.Data.DataQuery.DataProvider)
            If Result Is Nothing Then
                Result = CompuMaster.Data.DataQuery.DataProvider.AvailableDataProviders()
            End If
            Return Result
        End Get
    End Property

    Private Sub AvailableDataProvidersListbox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles AvailableDataProvidersListbox.SelectedIndexChanged
        Dim SelectedDataProvider As CompuMaster.Data.DataQuery.DataProvider
        SelectedDataProvider = CompuMaster.Data.DataQuery.DataProvider.LookupDataProvider(AvailableDataProvidersListbox.SelectedValue)
        If SelectedDataProvider.ConnectionType Is Nothing Then
            Me.FeatureInfoConnection.Checked = False
            Me.FeatureInfoConnection.Text = "N/A"
        Else
            Me.FeatureInfoConnection.Text = SelectedDataProvider.ConnectionType.ToString
            Dim Conn As System.Data.IDbConnection
            Try
                Conn = SelectedDataProvider.CreateConnection
                Me.FeatureInfoConnection.Checked = True
                Me.FeatureInfoConnectionError.Text = ""
            Catch ex As Exception
                Me.FeatureInfoConnection.Checked = False
                Me.FeatureInfoConnectionError.Text = ex.ToString.Replace(vbNewLine, "<br />")
            End Try
        End If
        If SelectedDataProvider.CommandType Is Nothing Then
            Me.FeatureInfoCommand.Checked = False
            Me.FeatureInfoCommand.Text = "N/A"
        Else
            Me.FeatureInfoCommand.Text = SelectedDataProvider.CommandType.ToString
            Dim Cmd As System.Data.IDbCommand
            Try
                Cmd = SelectedDataProvider.CreateCommand
                Me.FeatureInfoCommand.Checked = True
                Me.FeatureInfoCommandError.Text = ""
            Catch ex As Exception
                Me.FeatureInfoCommand.Checked = False
                Me.FeatureInfoCommandError.Text = ex.ToString.Replace(vbNewLine, "<br />")
            End Try
        End If
        If SelectedDataProvider.CommandBuilderType Is Nothing Then
            Me.FeatureInfoCommandBuilder.Checked = False
            Me.FeatureInfoCommandBuilder.Text = "N/A"
        Else
            Me.FeatureInfoCommandBuilder.Text = SelectedDataProvider.CommandBuilderType.ToString
            Dim CB As System.Data.Common.DbCommandBuilder
            Try
                CB = SelectedDataProvider.CreateCommandBuilder
                Me.FeatureInfoCommandBuilder.Checked = True
                Me.FeatureInfoCommandBuilderError.Text = ""
            Catch ex As Exception
                Me.FeatureInfoCommandBuilder.Checked = False
                Me.FeatureInfoCommandBuilderError.Text = ex.ToString.Replace(vbNewLine, "<br />")
            End Try
        End If
        If SelectedDataProvider.DataAdapterType Is Nothing Then
            Me.FeatureInfoDbDataAdapter.Checked = False
            Me.FeatureInfoDbDataAdapter.Text = "N/A"
        Else
            Me.FeatureInfoDbDataAdapter.Text = SelectedDataProvider.DataAdapterType.ToString
            Dim DA As System.Data.IDbDataAdapter
            Try
                DA = SelectedDataProvider.CreateDataAdapter
                Me.FeatureInfoDbDataAdapter.Checked = True
                Me.FeatureInfoDbDataAdapterError.Text = ""
            Catch ex As Exception
                Me.FeatureInfoDbDataAdapter.Checked = False
                Me.FeatureInfoDbDataAdapterError.Text = ex.ToString.Replace(vbNewLine, "<br />")
            End Try
        End If

        Me.FeatureShow.Visible = True
    End Sub
End Class
