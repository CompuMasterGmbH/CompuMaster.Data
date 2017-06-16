Option Strict On
Option Explicit On

Imports System.Runtime.InteropServices

Namespace CompuMaster.Data.Windows

    'TODO: implement usual clipboard commands: Ctrl+C,V,X, Ctrl+Ins/Shift+Ins/Shift-Del
    'TODO: implement Ctrl+Space, Shift+Space, Ctrl+Shift+Space
    'TODO: fix bug that VS designer often moves Friend WithEvents DataGridView1 As CompuMaster.Data.Windows.DataGridView into a Dim DataGridView1 As CompuMaster.Data.Windows.DataGridView in InitializeComponent method (making it private and inaccessible, leading to compiler errors); e.g. just add an additional button into the form with the DataGridView and double-click the button to add event handler code: at this time you should already find compiler errors in your existing code which refers to the DataGridView
    'TODO: fix known bugs (ask JW)
    ''' <summary>
    ''' A data grid view class implementing easy access for quick edit data manipulation
    ''' </summary>
    ''' <remarks></remarks>
    <System.ComponentModel.DesignTimeVisible(True)> _
    <System.Windows.Forms.DockingAttribute(System.Windows.Forms.DockingBehavior.Ask)> _
    <System.ComponentModel.Designer("System.Windows.Forms.Design.DataGridViewDesigner, System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a"), _
    ComVisible(True), _
    ClassInterface(ClassInterfaceType.AutoDispatch), _
    System.ComponentModel.DefaultEvent("CellContentClick"), _
    System.Reflection.DefaultMember("Item"), _
    System.ComponentModel.ComplexBindingProperties("DataSource", "DataMember"), _
    System.ComponentModel.Editor("System.Windows.Forms.Design.DataGridViewComponentEditor, System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", GetType(System.ComponentModel.ComponentEditor)),
    System.ComponentModel.Description("CompuMaster.Data.DataGridView")
    > _
    Public Class DataGridView
        Inherits System.Windows.Forms.DataGridView
        Implements System.ComponentModel.ISupportInitialize

        Public Sub New()
            ' Dieser Aufruf ist für den Designer erforderlich.
            InitializeComponent()

            ' Fügen Sie Initialisierungen nach dem InitializeComponent()-Aufruf hinzu.
            Me.MultiSelect = True
            Me.DefaultCellStyle.WrapMode = System.Windows.Forms.DataGridViewTriState.True
            Me.AlternatingRowsDefaultCellStyle.WrapMode = System.Windows.Forms.DataGridViewTriState.True

        End Sub

        ''' <summary>
        ''' The SELECT command for querying the data
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property SelectCommand As IDbCommand

        ''' <summary>
        ''' Immediately save/upload changes to the data source after a row change occured
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property SaveDataChangesAfterEveryRowChange As Boolean = True

        Private WithEvents _DataContainer As CompuMaster.Data.DataManipulationResult = Nothing
        ''' <summary>
        ''' The data container holds all necessary data for edit and upload
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property DataContainer As CompuMaster.Data.DataManipulationResult
            Get
                Return _DataContainer
            End Get
        End Property

        ''' <summary>
        ''' DataSource of base datagrid should not be used any more, use DataContainer or DataSourceObject property instead
        ''' </summary>
        ''' <returns></returns>
        <Obsolete("Use DataContainer or DataSourceObject instead", True)>
        Public Shadows Property DataSource As Object
            Get
                Return MyBase.DataSource
            End Get
            Set(value As Object)
                MyBase.DataSource = value
            End Set
        End Property

        ''' <summary>
        ''' The DataSource object as used by the underlying DataGridView
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property DataSourceObject As Object
            Get
                Return MyBase.DataSource
            End Get
        End Property

        ''' <summary>
        ''' Fires when the DataGridView changed data to the underlying data table
        ''' </summary>
        ''' <remarks></remarks>
        Public Event DataChanged()

        ''' <summary>
        ''' Auto-Enter into cells if there is no active multi-cell-selection
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub DataGridView_CellEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles Me.CellEnter
            If IsLoading = False AndAlso Not Me.DataContainer Is Nothing AndAlso Me._DataContainer.Table.Rows.Count > 0 AndAlso Not Me.SelectedCells Is Nothing AndAlso Me.SelectedCells.Count <= 1 Then
                Me.BeginEdit(True)
            End If
        End Sub

        ''' <summary>
        ''' Write back all changed data based on the DataChanged event
        ''' </summary>
        ''' <remarks></remarks>
        Protected Overridable Sub OnDataChanged() Handles Me.DataChanged
            If Me.SaveDataChangesAfterEveryRowChange Then
                Try
                    Me.SaveDataChanges()
                Catch ex As Exception
                    If Not ex.InnerException Is Nothing Then
                        MsgBox(ex.InnerException.Message, MsgBoxStyle.Critical)
                    Else
                        MsgBox(ex.Message, MsgBoxStyle.Critical)
                    End If
                End Try
            End If
        End Sub

        ''' <summary>
        ''' Load the data from the data source using the given SelectCommand and its connection
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub LoadData()
            Try
                Me.LoadData(True)
            Catch ex As Exception
                System.Windows.Forms.MessageBox.Show(ex.Message, "Error loading data", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
            End Try
        End Sub

        ''' <summary>
        ''' Indicates that data (re-)loading is in progress
        ''' </summary>
        ''' <remarks></remarks>
        Private IsLoading As Boolean = False

        ''' <summary>
        ''' Load the data from the data source using the given SelectCommand and its connection
        ''' </summary>
        ''' <param name="autoResizeColumnsAndRows"></param>
        ''' <remarks></remarks>
        Private Sub LoadData(ByVal autoResizeColumnsAndRows As Boolean)
            IsLoading = True
            PreserveColumnSettings()
            PreserveCurrentCell()
            PreserveScrollPosition()
            PreserveSortOrder()
            MyBase.DataSource = Nothing
            If SelectCommand Is Nothing Then Throw New InvalidOperationException("SelectCommand is a required property")
            If SelectCommand.Connection Is Nothing Then Throw New InvalidOperationException("SelectCommand requires a valid connection")
            _DataContainer = Utils.LoadDataForManipulationViaQuickEdit(Me.SelectCommand)
            RestoreColumnSettings()
            MyBase.DataSource = _DataContainer.Table
            RestoreSortOrder()
            RestoreRowSettings()
            RestoreCurrentCell()
            If autoResizeColumnsAndRows Then
                Me.AutoResizeColumns(System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells)
                Me.AutoResizeRows(System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells)
            End If
            Me.PerformLayout()
            RestoreScrollPostion()
            IsLoading = False
            Me.DataGridView_CellEnter(Nothing, Nothing) 'Enter the current cell
        End Sub

        ''' <summary>
        ''' The preserved sorting column index
        ''' </summary>
        ''' <remarks></remarks>
        Private PreservedSortedColumnIndex As Integer
        ''' <summary>
        ''' The preserved sort order direction
        ''' </summary>
        ''' <remarks></remarks>
        Private PreservedSortOrder As System.Windows.Forms.SortOrder

        ''' <summary>
        ''' Preserve the current sorting
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub PreserveSortOrder()
            If SortedColumn Is Nothing Then
                PreservedSortedColumnIndex = -1
            Else
                PreservedSortedColumnIndex = Me.SortedColumn.Index
            End If
            PreservedSortOrder = Me.SortOrder
        End Sub
        ''' <summary>
        ''' Restore the preserved sorting
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub RestoreSortOrder()
            If PreservedSortOrder = System.Windows.Forms.SortOrder.Ascending Then
                Me.Sort(Me.Columns(PreservedSortedColumnIndex), System.ComponentModel.ListSortDirection.Ascending)
            ElseIf PreservedSortOrder = System.Windows.Forms.SortOrder.Descending Then
                Me.Sort(Me.Columns(PreservedSortedColumnIndex), System.ComponentModel.ListSortDirection.Descending)
            End If
        End Sub

        ''' <summary>
        ''' A collection of cloned columns for later re-establishing after a reload of data
        ''' </summary>
        ''' <remarks></remarks>
        Private PreservedColumns As System.Windows.Forms.DataGridViewColumn()
        ''' <summary>
        ''' A collection of cloned rows for later re-establishing after a reload of data
        ''' </summary>
        ''' <remarks></remarks>
        Private PreservedRows As System.Windows.Forms.DataGridViewRow()

        ''' <summary>
        ''' Save a cloned collection of the columns and rows for later restoring of collection and row heights
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub PreserveColumnSettings()
            ReDim PreservedColumns(Me.Columns.Count - 1)
            For MyCounter As Integer = 0 To Me.Columns.Count - 1
                PreservedColumns(MyCounter) = CType(Me.Columns(MyCounter).Clone, System.Windows.Forms.DataGridViewColumn)
                If Me.Columns(MyCounter).DisplayIndex <> -1 Then PreservedColumns(MyCounter).DisplayIndex = Me.Columns(MyCounter).DisplayIndex
            Next
            ReDim PreservedRows(Me.Rows.Count - 1)
            For MyCounter As Integer = 0 To Me.Rows.Count - 1
                PreservedRows(MyCounter) = CType(Me.Rows(MyCounter).Clone, System.Windows.Forms.DataGridViewRow)
            Next
        End Sub

        ''' <summary>
        ''' Restore the preserved columns collection by cloning each pre-defined column from the origin collection
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub RestoreColumnSettings()
            If Me.Columns.Count = 0 Then 'Should always be the case since Me.DataSource has been set to Nothing and caused by that, the columns collection has been wiped
                For MyCounter As Integer = 0 To Me.PreservedColumns.Length - 1
                    Me.Columns.AddRange(CType(PreservedColumns(MyCounter).Clone, System.Windows.Forms.DataGridViewColumn))
                Next
                For MyCounter As Integer = 0 To Me.PreservedColumns.Length - 1
                    If Me.Columns(MyCounter).DisplayIndex <> -1 Then Me.Columns(MyCounter).DisplayIndex = PreservedColumns(MyCounter).DisplayIndex
                Next
            End If
        End Sub

        ''' <summary>
        ''' Restore the preserved row heights
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub RestoreRowSettings()
            For MyCounter As Integer = 0 To System.Math.Min(Me.PreservedRows.Length - 1, Me.Rows.Count - 1)
                Me.Rows(MyCounter).Height = Me.PreservedRows(MyCounter).Height
            Next
        End Sub

        ''' <summary>
        ''' The preserved row index of the current position in the gridview
        ''' </summary>
        ''' <remarks></remarks>
        Private PreservedCurrentCellRowIndex As Integer
        ''' <summary>
        ''' The preserved column index of the current position in the gridview
        ''' </summary>
        ''' <remarks></remarks>
        Private PreservedCurrentCellColumnIndex As Integer

        ''' <summary>
        ''' Preserve the current position in the gridview
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub PreserveCurrentCell()
            If Me.CurrentCell Is Nothing Then
                PreservedCurrentCellRowIndex = -1
                PreservedCurrentCellColumnIndex = -1
            Else
                PreservedCurrentCellRowIndex = Me.CurrentCell.RowIndex
                PreservedCurrentCellColumnIndex = Me.CurrentCell.ColumnIndex
            End If
        End Sub

        ''' <summary>
        ''' Restore the preserved position in the gridview so that current view and position doesn't jump unexpectedly
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub RestoreCurrentCell()
            If PreservedCurrentCellRowIndex <> -1 AndAlso PreservedCurrentCellColumnIndex <> -1 AndAlso PreservedCurrentCellRowIndex < Me.Rows.Count AndAlso PreservedCurrentCellColumnIndex < Me.Columns.Count Then
                Me.CurrentCell = Me.Rows(PreservedCurrentCellRowIndex).Cells(PreservedCurrentCellColumnIndex)
            ElseIf PreservedCurrentCellRowIndex <> -1 AndAlso PreservedCurrentCellColumnIndex <> -1 Then
                'Reduce row/cell position to max row/count amount
                Me.CurrentCell = Me.Rows(System.Math.Min(System.Math.Max(Me.Rows.Count - 1, 0), PreservedCurrentCellRowIndex)).Cells(System.Math.Min(System.Math.Max(Me.Columns.Count - 1, 0), PreservedCurrentCellColumnIndex))
            End If
        End Sub

        ''' <summary>
        ''' The preserved vertical scrollbar value of the current position in the gridview
        ''' </summary>
        ''' <remarks></remarks>
        Private PreservedCurrentVerticalScrollBarValue As Integer
        ''' <summary>
        ''' The preserved horizontal scrollbar value of the current position in the gridview
        ''' </summary>
        ''' <remarks></remarks>
        Private PreservedCurrentHorizontalScrollBarValue As Integer
        ''' <summary>
        ''' The preserved row index of the first displayed cell in the gridview
        ''' </summary>
        ''' <remarks></remarks>
        Private PreservedFirstDisplayedScrollingRowIndex As Integer
        ''' <summary>
        ''' The preserved column index of the first displayed cell in the gridview
        ''' </summary>
        ''' <remarks></remarks>
        Private PreservedFirstDisplayedScrollingColumnIndex As Integer

        ''' <summary>
        ''' Preserve the scroll position of the gridview
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub PreserveScrollPosition()
            PreservedCurrentVerticalScrollBarValue = Me.VerticalScrollBar.Value
            PreservedCurrentHorizontalScrollBarValue = Me.HorizontalScrollBar.Value
            PreservedFirstDisplayedScrollingRowIndex = Me.FirstDisplayedScrollingRowIndex
            PreservedFirstDisplayedScrollingColumnIndex = Me.FirstDisplayedScrollingColumnIndex
        End Sub

        ''' <summary>
        ''' Restore the scroll position of the gridview
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub RestoreScrollPostion()
            Me.VerticalScrollBar.Value = System.Math.Min(PreservedCurrentVerticalScrollBarValue, Me.VerticalScrollBar.Maximum)
            Me.HorizontalScrollBar.Value = System.Math.Min(PreservedCurrentHorizontalScrollBarValue, Me.HorizontalScrollBar.Maximum)
            Me.FirstDisplayedScrollingRowIndex = System.Math.Max(System.Math.Min(PreservedFirstDisplayedScrollingRowIndex, Me.Rows.Count - 1), 0)
            Me.FirstDisplayedScrollingColumnIndex = System.Math.Max(PreservedFirstDisplayedScrollingColumnIndex, 0)
        End Sub

        ''' <summary>
        ''' Save/upload the data changes
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub SaveDataChanges()
            Dim CurrentCursor As System.Windows.Forms.Cursor
            CurrentCursor = Me.Cursor
            Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
            Try
                'Now let's write the changes back to the database
                Utils.SaveData(_DataContainer)
                'Reload data from data source
                Me.LoadData(False)
            Catch ex As DBConcurrencyException
                'Reload data from data source
                Me.LoadData(False)
            Finally
                Me.Cursor = CurrentCursor
            End Try
        End Sub

        ''' <summary>
        ''' Ignore errors when binary arrays such as image, timestamp or binary fields will be tryed to be presented as a picture
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks>The default behaviour of .NET will be changed here to ignore such errors per default for increased stability of this control. Otherwise, the default behaviour of .NET would try to draw those fields and run into exceptions sometimes.</remarks>
        Public Property IgnoreErrorsOnLayoutingBinaryValues As Boolean = True

        ''' <summary>
        ''' Ignore byte-array/image errors
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub DataGridViewQuickEdit_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles Me.DataError
            If CType(MyBase.DataSource, DataTable).Columns(e.ColumnIndex).DataType Is GetType(Byte()) Then
                'Ignore this error
                'typically, the DataGrid trys to display Byte-arrays as image which may fail depending on provided data.
                'e.g. upsize_ts columns will always fail here
                e.ThrowException = False
            End If
        End Sub

        ''' <summary>
        ''' Forward DataChanged event
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub _DataContainer_DataChanged() Handles _DataContainer.DataChanged
            RaiseEvent DataChanged()
        End Sub

        Private Sub DataGridView_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
            If e.Control = True And e.KeyCode = System.Windows.Forms.Keys.A Then
                For MyCounter As Integer = 0 To Me.Rows.Count - 1 - CType(IIf(Me.NewRowIndex >= 0, 1, 0), Integer)
                    Me.SetSelectedRowCore(MyCounter, True)
                Next
            End If
        End Sub

        ''' <summary>
        ''' Prevent paint exceptions due to scrolling position issues (bug in MS gridview)
        ''' </summary>
        ''' <param name="e"></param>
        ''' <remarks>Workaround recommendation source: https://connect.microsoft.com/VisualStudio/feedback/details/673075/datagridview-control-crashed-gui-shows-white-area-with-red-cross-app-shows-error-below </remarks>
        Protected Overrides Sub OnPaint(ByVal e As System.Windows.Forms.PaintEventArgs)
            Try
                MyBase.OnPaint(e)
            Catch ex As Exception
                Me.Invalidate()
            End Try
        End Sub

    End Class

End Namespace