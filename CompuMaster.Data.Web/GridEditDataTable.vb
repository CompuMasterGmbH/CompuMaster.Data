Imports System
Imports System.ComponentModel
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls

Namespace CompuMaster.Data.Web

    'Maybe the better way of implementation???
    'Public Class DataTableSource
    '    Inherits System.Web.UI.WebControls.ObjectDataSource
    'End Class

    <DefaultProperty("Text"), ToolboxData("<{0}:ServerControl1 runat=server></{0}:ServerControl1>")> _
    Public Class GridEditDataTable
        Inherits System.Web.UI.WebControls.GridView

        Protected Overrides Sub OnLoad(ByVal e As System.EventArgs)
            MyBase.OnLoad(e)
            If _IsDataBound = False Then Me.DataBind()
        End Sub

        Public Class DataManipulationResults
            Public Table As System.Data.DataTable
            Public DataAdapter As System.Data.IDataAdapter
            Public Command As System.Data.IDbCommand
        End Class

        Private _SuppressEventRowCreated As Boolean = False

        Protected Overrides Sub DataBindChildren()
            _SuppressEventRowCreated = True
            MyBase.DataBindChildren()
            _SuppressEventRowCreated = False
        End Sub

        Protected Overrides Sub PerformSelect()
            _SuppressEventRowCreated = True
            MyBase.PerformSelect()
            _SuppressEventRowCreated = False
        End Sub

        Sub DataAssign()

            Dim data As System.Data.DataTable
            data = CompuMaster.Data.DataTables.CreateDataTableClone(Me.DataTable, "", SortArgument)
            'data.rows(0)("Value2Setup") = "line1" & vbnewline & "<b>line2</b>"
            For MyColCounter As Integer = 0 To data.Columns.Count - 1
                If data.Columns(MyColCounter).DataType Is GetType(String) Then
                    For MyRowCounter As Integer = 0 To data.Rows.Count - 1
                        data.Rows(MyRowCounter)(MyColCounter) = "$data:text/plain$" & Utils.Nz(data.Rows(MyRowCounter)(MyColCounter), "")
                    Next
                End If
            Next
            Me.DataSource = data

            For MyCounter As Integer = 0 To Me.Rows.Count - 1
                For MyColCounter As Integer = 0 To Me.Rows(MyCounter).Cells.Count - 1
                    If Me.Rows(MyCounter).Cells(MyColCounter).Text.StartsWith("$data:text/plain$") Then
                        Me.Rows(MyCounter).Cells(MyColCounter).Text = "<nobr>" & Utils.HTMLEncodeLineBreaks((Me.Rows(MyCounter).Cells(MyColCounter).Text.Substring("$data:text/plain$".Length()))) & "</nobr>"
                    End If
                Next
            Next

            _IsDataBound = False

        End Sub

        Protected Overrides Sub DataBind(ByVal raiseOnDataBinding As Boolean)
            _SuppressEventRowCreated = True
            MyBase.DataBind(raiseOnDataBinding)
            _SuppressEventRowCreated = False
        End Sub

        Public Property DataTable() As DataTable
            Get
                Return ViewState("DataTable")
            End Get
            Set(ByVal value As DataTable)
                ViewState("DataTable") = value
                DataAssign()
            End Set
        End Property

        Private _IsDataBound As Boolean = False

        Public QuickEditDataContainer As DataManipulationResults = Nothing
        Dim QuickEditRecordCountBefore As Integer = 0

        Public Event DataTableUpdated()

        Private Sub DataGridTableEditor_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles Me.RowCreated
            If _SuppressEventRowCreated = True Then Return
            If QuickEditDataContainer Is Nothing Then
                Me.DataAssign()
                Me.DataBind()
            End If
            'e.Row
            RaiseEvent DataTableUpdated()
        End Sub

        Private Sub DataGridTableEditor_RowDeleted(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeletedEventArgs) Handles Me.RowDeleted
            If QuickEditDataContainer Is Nothing Then
                Me.DataAssign()
                Me.DataBind()
            End If
            'e.Keys
            RaiseEvent DataTableUpdated()
        End Sub

        Private Sub GridEditView_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles Me.RowDeleting

        End Sub

        'Private Sub GridEditView_RowEditing(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewEditEventArgs) Handles Me.RowEditing
        '    MyBase.h()
        'End Sub

        Private Sub DataGridTableEditor_RowUpdated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewUpdatedEventArgs) Handles Me.RowUpdated
            If QuickEditDataContainer Is Nothing Then
                Me.DataAssign()
                Me.DataBind()
            End If
            'e.Keys
            'e.NewValues
            RaiseEvent DataTableUpdated()
        End Sub

        Private Sub GridEditView_RowUpdating(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewUpdateEventArgs) Handles Me.RowUpdating

        End Sub

        Private Property SortArgument() As String
            Get
                Return ViewState("SortArgument")
            End Get
            Set(ByVal value As String)
                ViewState("SortArgument") = value
            End Set
        End Property

        Private Sub GridSorting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewSortEventArgs) Handles Me.Sorting
            Dim NewSorting As String = e.SortExpression & " " & IIf(e.SortDirection = SortDirection.Descending, "DESC", "ASC")
            HttpContext.Current.Response.Write(NewSorting & "<br>")
            If SortArgument = NewSorting Then
                'Reverse direction
                If e.SortDirection = SortDirection.Ascending Then
                    e.SortDirection = SortDirection.Descending
                Else
                    e.SortDirection = SortDirection.Ascending
                End If
                NewSorting = e.SortExpression & " " & IIf(e.SortDirection = SortDirection.Descending, "DESC", "ASC")
            End If
            HttpContext.Current.Response.Write(NewSorting & "<br>")
            SortArgument = NewSorting
            Me.DataAssign()
            Me.DataBind()
        End Sub

        Private Sub GridRowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles Me.RowCommand
            Select Case e.CommandName
                Case "Accept"
                    Dim DataRowID As Integer
                    DataRowID = Integer.Parse(CType(e.CommandSource, GridView).Rows(CType(e.CommandArgument, Integer)).Cells(3).Text)
                    HttpContext.Current.Response.Write("Accepted ID: " & DataRowID)
                Case "Reject"
                    Dim DataRowID As Integer
                    DataRowID = Integer.Parse(CType(e.CommandSource, GridView).Rows(CType(e.CommandArgument, Integer)).Cells(3).Text)
                    HttpContext.Current.Response.Write("Rejected ID: " & DataRowID)
                    HttpContext.Current.Response.Write("<<")
                    HttpContext.Current.Response.Write(CType(e.CommandSource, GridView).SelectedPersistedDataKey.Value.ToString & "<<")
                    HttpContext.Current.Response.Write(e.CommandArgument.ToString())
                    HttpContext.Current.Response.Write(">>")
                    HttpContext.Current.Response.Write(e.CommandSource.ToString())
                Case Else
                    'Throw New Exception("Invalid grid command " & e.CommandName)
            End Select
        End Sub

        Protected Sub GridRowEditing(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewEditEventArgs) Handles Me.RowEditing
            Me.EditIndex = e.NewEditIndex
        End Sub

        Protected Sub GridRowCancelingEdit(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCancelEditEventArgs) Handles Me.RowCancelingEdit
            Me.EditIndex = -1
        End Sub

    End Class

End Namespace