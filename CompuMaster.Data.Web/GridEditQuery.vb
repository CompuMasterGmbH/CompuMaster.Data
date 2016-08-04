Imports System.ComponentModel
Imports System.Web.UI
Imports System.Web.UI.WebControls

Namespace CompuMaster.Data.Web

	<DefaultProperty("Text"), ToolboxData("<{0}:ServerControl1 runat=server></{0}:ServerControl1>")> _
	Public Class GridEditQuery
		Inherits System.Web.UI.WebControls.GridView

		<Bindable(True), Category("Data"), DefaultValue(""), Localizable(False)> Public Property DataProviderName() As String
			Get
				Return CType(ViewState("ProviderName"), String)
			End Get
			Set(ByVal value As String)
				Select Case LCase(value)
					Case "sqlclient", "odbc", "oledb"
						ViewState("ProviderName") = value
					Case Else
						Throw New ArgumentOutOfRangeException("value", value, "ProviderName must be one of these values: SqlClient, Odbc, OleDb")
				End Select
				_IsDataLoaded = False
			End Set
		End Property

		<Bindable(True), Category("Data"), DefaultValue(""), Localizable(False)> Public Property DataConnectionString() As String
			Get
				Return CType(ViewState("ConnectionString"), String)
			End Get
			Set(ByVal value As String)
				ViewState("ConnectionString") = value
				_IsDataLoaded = False
			End Set
		End Property

		<Bindable(True), Category("Data"), DefaultValue(""), Localizable(False)> Public Property DataSelectCommand() As String
			Get
				Return CType(ViewState("SelectCommand"), String)
			End Get
			Set(ByVal value As String)
				ViewState("SelectCommand") = value
				_IsDataLoaded = False
			End Set
		End Property

        Protected Overrides Sub OnLoad(ByVal e As System.EventArgs)
            Me.DataLoadException = Nothing
            Try
                Me.DataLoad()
                MyBase.OnLoad(e)
                Me.DataBind()
            Catch ex As Exception
                Me.DataLoadException = ex
            End Try
        End Sub

        Private Sub DataLoadExceptionCheckAfterPageLoad(sender As Object, ByVal e As System.EventArgs) Handles Me.PreRender
            If Me.DataLoadException IsNot Nothing Then Throw New Exception("DataLoadException occured on loading of control, but not handled by page on time", Me.DataLoadException)
        End Sub

        Public Sub DataLoadExceptionCheck()
            If Me.DataLoadException IsNot Nothing Then
                Dim ThrowEx As Exception = Me.DataLoadException
                Me.DataLoadException = Nothing
                Throw New Exception("Exception loading control data", ThrowEx)
            End If
        End Sub

        Public Property DataLoadException As Exception

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
		Protected Overrides Sub DataBind(ByVal raiseOnDataBinding As Boolean)
			If _IsDataLoaded = False Then DataLoad()
			_SuppressEventRowCreated = True
			MyBase.DataBind(raiseOnDataBinding)
			_SuppressEventRowCreated = False
		End Sub

        Private Function LoadDataForManipulationViaQuickEdit() As CompuMaster.Data.DataManipulationResult
            Dim Provider As CompuMaster.Data.DataQuery.DataProvider = CompuMaster.Data.DataQuery.DataProvider.LookupDataProvider(Me.DataProviderName)
            Dim MyCmd As System.Data.IDbCommand = Provider.CreateCommand(Me.DataSelectCommand, Me.DataConnectionString)
            MyCmd.CommandType = CommandType.Text
            MyCmd.CommandTimeout = 300 '5 minutes
            Return CompuMaster.Data.Manipulation.LoadQueryDataForManipulationViaCode(MyCmd)
        End Function

        Private _IsDataLoaded As Boolean = False
		Public Sub DataLoad()
			'Release any data loaded before
			CloseAndDisposeQuickEditDataContainer()
            'Load data for manipulation process
            QuickEditDataContainer = LoadDataForManipulationViaQuickEdit()
            QuickEditRecordCountBefore = QuickEditDataContainer.Table.Rows.Count
			Me.DataSource = QuickEditDataContainer.Table
			'Me.DataSource = New ObjectDataSource()
			_IsDataLoaded = True
			_UpdateStatus = "Table loaded: " & QuickEditDataContainer.Table.TableName
		End Sub

		Private Class os
			Inherits ObjectDataSource

			Private Sub os_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

			End Sub

			Public Sub New()
				'Me.SetDesignModeState()
			End Sub
		End Class

        Public QuickEditDataContainer As CompuMaster.Data.DataManipulationResult = Nothing
        Dim QuickEditRecordCountBefore As Integer = 0

		Private QuickUploadCounter As Integer = 0
		Public Sub QuickEditUploadChanges()

			Try
				Dim recordsAffected As DataTable = QuickEditDataContainer.Table.GetChanges
				Dim recordCountAffected As Integer
				If Not recordsAffected Is Nothing Then
					recordCountAffected = QuickEditDataContainer.Table.GetChanges.Rows.Count()
				End If

                'Now let's write the changes back to the database
                CompuMaster.Data.Manipulation.UpdateCodeManipulatedData(QuickEditDataContainer)
                Dim recordCountAfter As Integer = QuickEditDataContainer.Table.Rows.Count

				'Show successful execution state
				'_UpdateStatus = "Data has been successfully uploaded" & vbNewLine & "Records loaded: " & QuickEditRecordCountBefore & vbNewLine & "Records affected: " & recordCountAffected & vbNewLine & "Records after changes: " & recordCountAfter
				_UpdateStatus &= "Data has been successfully uploaded" & vbNewLine & System.Environment.StackTrace & vbNewLine & "Records loaded: " & QuickEditRecordCountBefore & vbNewLine & "Records affected: " & recordCountAffected & vbNewLine & "Records after changes: " & recordCountAfter
				QuickUploadCounter += 1
				_UpdateStatus &= vbNewLine & "Uploads: " & QuickUploadCounter

			Catch ex As Exception
				If ex.GetType Is GetType(System.Data.SqlClient.SqlException) Then
					Dim sqlEx As System.Data.SqlClient.SqlException = CType(ex, System.Data.SqlClient.SqlException)
					Throw New Exception("Data update error", New Exception(sqlEx.Message & " -- Line " & sqlEx.LineNumber, sqlEx))
				Else
					Throw New Exception("Data update error", ex)
				End If
			End Try

			'Reload changed data and bind
			DataLoad()
			DataBind()

		End Sub

		Private _UpdateStatus As String
		Public ReadOnly Property UpdateStatus() As String
			Get
				Return _UpdateStatus
			End Get
		End Property

		Private Sub CloseAndDisposeQuickEditDataContainer()

			If Not QuickEditDataContainer Is Nothing AndAlso Not QuickEditDataContainer.Command Is Nothing Then
				If Not QuickEditDataContainer.Command.Connection Is Nothing Then
					If QuickEditDataContainer.Command.Connection.State <> ConnectionState.Closed Then
						QuickEditDataContainer.Command.Connection.Close()
					End If
					QuickEditDataContainer.Command.Connection.Dispose()
				End If
				If Not QuickEditDataContainer.Command Is Nothing Then
					QuickEditDataContainer.Command.Dispose()
				End If
			End If

		End Sub

		Private Sub GridEditQuery_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles Me.PageIndexChanging
			Me.PageIndex = e.NewPageIndex
			'Bind data to the GridView control.
			Me.DataBind()
		End Sub

		Private Sub GridEditQuery_RowCancelingEdit(sender As Object, e As GridViewCancelEditEventArgs) Handles Me.RowCancelingEdit
			'Reset the edit index.
			Me.EditIndex = -1
			'Bind data to the GridView control.
			Me.DataBind()
		End Sub

		Private Sub DataGridTableEditor_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles Me.RowCreated
			If _SuppressEventRowCreated = True Then Return
			If QuickEditDataContainer Is Nothing Then Me.DataLoad()
			'e.Row
			QuickEditUploadChanges()
		End Sub

		Private Sub DataGridTableEditor_RowDeleted(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeletedEventArgs) Handles Me.RowDeleted
			If QuickEditDataContainer Is Nothing Then Me.DataLoad()
			'e.Keys
			QuickEditUploadChanges()
		End Sub

		Private Sub GridEditView_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles Me.RowDeleting

		End Sub

		Private Sub GridEditQuery_RowEditing(sender As Object, e As GridViewEditEventArgs) Handles Me.RowEditing
			'Set the edit index.
			Me.EditIndex = e.NewEditIndex
			'Bind data to the GridView control.
			Me.DataBind()
		End Sub

		Private Function GetCombinationNotSupportedException(ControlType As Type, FieldType As Type) As NotSupportedException
			Return New NotSupportedException(Utils.BuildString("Control type", ControlType, """ doesn't fit with column data type """, FieldType, """ or comibnation not implemented yet"))
		End Function

		Private Sub RowUpdating_HandleBoundFields(Column As BoundField, Row As GridViewRow, InnerControl As WebControl)
			Dim FieldCol As System.Web.UI.WebControls.BoundField = CType(Column, System.Web.UI.WebControls.BoundField)

			If FieldCol.ReadOnly = False Then
				'Update editable column values
				Dim drow As System.Data.DataRow = Me.QuickEditDataContainer.Table.Rows(Row.DataItemIndex)
				Dim ControlType As Type = InnerControl.GetType()
				Dim FieldType As Type = Me.QuickEditDataContainer.Table.Columns(FieldCol.DataField).DataType

				If ControlType Is GetType(TextBox) Then
					If Me.Page.Request(InnerControl.UniqueID).Length = 0 Then
						If Me.QuickEditDataContainer.Table.Columns(FieldCol.DataField).AllowDBNull Then
							drow(FieldCol.DataField) = DBNull.Value
						Else
							drow(FieldCol.DataField) = ""
						End If
					ElseIf FieldType Is GetType(String) Then
						Dim newValue As String = Me.Page.Request(InnerControl.UniqueID)
						drow(FieldCol.DataField) = newValue
					ElseIf FieldType Is GetType(Int32) Then
						Dim newValue As Int32
						Dim tBox As TextBox = CType(InnerControl, TextBox)
						If Int32.TryParse(tBox.Text, newValue) Then
							drow(FieldCol.DataField) = newValue
						Else
							Throw New NotSupportedException(Utils.BuildString("Value of ", tBox.UniqueID, " is not Int32 or equal."))
						End If
					Else
						Throw GetCombinationNotSupportedException(ControlType, FieldType)
					End If
				ElseIf ControlType Is GetType(DropDownList) Then
					If FieldType Is GetType(String) Then
						Dim newValue As String = Me.Page.Request(InnerControl.UniqueID)
						drow(FieldCol.DataField) = newValue
					Else
						Throw GetCombinationNotSupportedException(ControlType, FieldType)
					End If
				ElseIf ControlType Is GetType(CheckBox) Then
					If FieldType Is GetType(Boolean) Then
						Throw GetCombinationNotSupportedException(ControlType, FieldType)
						'TODO: Retrieve Checkbox Value from RequestForm data correctly (because CheckBox.Checked contains the old data before change)
						'Dim newValue As Boolean = CType(row.Cells(MyCounter).Controls(0), CheckBox).Checked
						'drow(FieldCol.DataField) = newValue
					Else
						Throw GetCombinationNotSupportedException(ControlType, FieldType)
					End If
				Else
					Throw GetCombinationNotSupportedException(ControlType, FieldType)
				End If
			End If
		End Sub

		Private Sub RowUpdating_HandleTemplateFields(Column As TemplateField, Row As GridViewRow, InnerControl As Control)
			Dim FieldCol As System.Web.UI.WebControls.TemplateField = CType(Column, System.Web.UI.WebControls.TemplateField)

			'Update editable column values
			Dim drow As System.Data.DataRow = Me.QuickEditDataContainer.Table.Rows(Row.DataItemIndex)
			Dim ControlType As Type = InnerControl.GetType()
			Dim FieldType As Type

			If ControlType Is GetType(BoundDropDown) Then
				Dim cControl As BoundDropDown = CType(InnerControl, BoundDropDown)

				FieldType = Me.QuickEditDataContainer.Table.Columns(cControl.BoundField).DataType

				If FieldType Is GetType(String) Then
					Dim newValue As String = Me.Page.Request(InnerControl.UniqueID)
					drow(cControl.BoundField) = newValue
				Else
					GetCombinationNotSupportedException(ControlType, FieldType)
				End If
			Else
				Throw New NotSupportedException(Utils.BuildString("Support for control type """, ControlType, """ not implemented yet"))
			End If
		End Sub

		Private Sub DataGridTableEditor_RowUpdating(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewUpdateEventArgs) Handles Me.RowUpdating
			If QuickEditDataContainer Is Nothing Then Me.DataLoad()

			'Update the values.
			Dim row As GridViewRow = Me.Rows(e.RowIndex)
			'QuickEditDataContainer.Table.Rows(row.DataItemIndex)("")
			For MyCounter As Integer = 0 To Me.Columns.Count - 1
				Dim col As Object = Me.Columns(MyCounter)

				If row.Cells(MyCounter).Controls.Count > 0 Then
					If col.GetType() Is GetType(System.Web.UI.WebControls.BoundField) Then
						RowUpdating_HandleBoundFields(col, row, row.Cells(MyCounter).Controls(0))
					ElseIf col.GetType() Is GetType(TemplateField) Then
						If row.Cells(MyCounter).Controls(1).GetType() Is GetType(LiteralControl) Then
							RowUpdating_HandleTemplateFields(col, row, row.Cells(MyCounter).Controls(0))
						Else
							RowUpdating_HandleTemplateFields(col, row, row.Cells(MyCounter).Controls(1))
						End If
					End If
				End If
			Next

			'Reset the edit index.
			Me.EditIndex = -1

			Me.QuickEditUploadChanges()

		End Sub



		Private Property SortArgument() As String
			Get
				Return ViewState("SortArgument")
			End Get
			Set(ByVal value As String)
				ViewState("SortArgument") = value
			End Set
		End Property

		Private Sub GridEditQuery_Sorting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewSortEventArgs) Handles Me.Sorting
			Dim NewSorting As String = e.SortExpression & " " & IIf(e.SortDirection = SortDirection.Descending, "DESC", "ASC")
			If SortArgument = NewSorting Then
				'Reverse direction
				If e.SortDirection = SortDirection.Ascending Then
					e.SortDirection = SortDirection.Descending
				Else
					e.SortDirection = SortDirection.Ascending
				End If
				Me.Sort(e.SortExpression, e.SortDirection)
				NewSorting = e.SortExpression & " " & IIf(e.SortDirection = SortDirection.Descending, "DESC", "ASC")
			End If
			SortArgument = NewSorting
			Me.DataBind()
		End Sub

		Private Sub DataGridTableEditor_Unload(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Unload
			'Release any data loaded before
			CloseAndDisposeQuickEditDataContainer()
		End Sub

	End Class

End Namespace