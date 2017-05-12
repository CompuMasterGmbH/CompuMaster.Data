Option Explicit On
Option Strict On

Imports System.Windows.Forms
Imports System.Drawing

Namespace CompuMaster.Data.Windows

    ''' <summary>
    ''' Context menu default entries
    ''' </summary>
    ''' <remarks>
    ''' This Code comes mostly from CompuMaster.Data.Manipulation
    ''' </remarks>
    <System.ComponentModel.DesignTimeVisible(True)> _
    Public Class CellCopyingContextMenu
        Inherits ContextMenuStrip

        Private Grid As System.Windows.Forms.DataGridView
        Private WithEvents CopyFullTableToClipboardToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
        Private WithEvents CopySelectedCellsToClipboardWithHeadersToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
        Private WithEvents CopySelectedCellsToClipboardWithoutHeadersToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
        Private WithEvents ExportCultureOptionsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
        Private WithEvents urrentCultureToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
        Private WithEvents InvariantCultureToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
        Private WithEvents CultureOfOperatingSystemToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
        Private WithEvents EnglishCultureenUSToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
        Private WithEvents CurrentCultureToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem

        Public Sub New()
            MyBase.New()
            Me.Initialize()
        End Sub

        Public Sub New(grid As System.Windows.Forms.DataGridView)
            MyBase.New()
            Me.Grid = grid
            Me.Initialize()
        End Sub

        Public Sub New(container As System.ComponentModel.IContainer)
            MyBase.New(container)
            Me.Initialize()
        End Sub

        Public Sub New(container As System.ComponentModel.IContainer, grid As System.Windows.Forms.DataGridView)
            MyBase.New(container)
            Me.Grid = grid
            Me.Initialize()
        End Sub

        Private Sub CellCopyingContextMenu_ItemAdded(sender As Object, e As ToolStripItemEventArgs) Handles Me.ItemAdded
            Me.ResumeLayout()
        End Sub

        <System.ComponentModel.Category("Data")> Public Property DataGridView As System.Windows.Forms.DataGridView
            Get
                Return Me.Grid
            End Get
            Set(value As System.Windows.Forms.DataGridView)
                Me.Grid = value
                If Me.Grid IsNot Nothing AndAlso Me.Grid.ContextMenuStrip Is Nothing Then
                    Me.Grid.ContextMenuStrip = Me
                End If
            End Set
        End Property

        Public Property CopyFullTableText As String = "Copy full table to clipboard"
        Public Property CopySelectedCellsWithHeadersText As String = "Copy selected cells to clipboard (with headers)"
        Public Property CopySelectedCellsWithoutHeadersText As String = "Copy selected cells to clipboard (without headers)"
        Public Property ExportCultureOptionsText As String = "Export culture options"
        Public Property CurrentCultureText As String = "Current culture"
        Public Property EnglishCultureText As String = "English culture (en-US)"
        Public Property InternationCultureText As String = "International culture"
        Public Property OsCultureText As String = "Culture of operating system"

        Private Sub Initialize()
            Me.CopySelectedCellsToClipboardWithHeadersToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
            Me.CopySelectedCellsToClipboardWithoutHeadersToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
            Me.CopyFullTableToClipboardToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
            Me.ExportCultureOptionsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
            Me.CurrentCultureToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
            Me.EnglishCultureenUSToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
            Me.InvariantCultureToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
            Me.CultureOfOperatingSystemToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()

            'Me.Items.Add(Me.CopySelectedCellsToClipboardWithHeadersToolStripMenuItem)
            Me.Name = "ContextMenuStripDataGrid"
            Me.Size = New System.Drawing.Size(338, 92)

            Me.CopyFullTableToClipboardToolStripMenuItem.Name = "CopyFullTableToClipboardToolStripMenuItem"
            Me.CopyFullTableToClipboardToolStripMenuItem.Size = New System.Drawing.Size(337, 22)
            Me.CopyFullTableToClipboardToolStripMenuItem.Text = CopyFullTableText

            Me.CopySelectedCellsToClipboardWithHeadersToolStripMenuItem.Name = "CopySelectedCellsToClipboardWithHeadersToolStripMenuItem"
            Me.CopySelectedCellsToClipboardWithHeadersToolStripMenuItem.Size = New System.Drawing.Size(337, 22)
            Me.CopySelectedCellsToClipboardWithHeadersToolStripMenuItem.Text = CopySelectedCellsWithHeadersText


            Me.CopySelectedCellsToClipboardWithoutHeadersToolStripMenuItem.Name = "CopySelectedCellsToClipboardWithoutHeadersToolStripMenuItem"
            Me.CopySelectedCellsToClipboardWithoutHeadersToolStripMenuItem.Size = New System.Drawing.Size(337, 22)
            Me.CopySelectedCellsToClipboardWithoutHeadersToolStripMenuItem.Text = CopySelectedCellsWithoutHeadersText

            Me.ExportCultureOptionsToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CurrentCultureToolStripMenuItem, Me.EnglishCultureenUSToolStripMenuItem, Me.InvariantCultureToolStripMenuItem, Me.CultureOfOperatingSystemToolStripMenuItem})
            Me.ExportCultureOptionsToolStripMenuItem.Name = "ExportCultureOptionsToolStripMenuItem"
            Me.ExportCultureOptionsToolStripMenuItem.Size = New System.Drawing.Size(337, 22)
            Me.ExportCultureOptionsToolStripMenuItem.Text = ExportCultureOptionsText

            Me.CurrentCultureToolStripMenuItem.Checked = True
            Me.CurrentCultureToolStripMenuItem.CheckState = System.Windows.Forms.CheckState.Checked
            Me.CurrentCultureToolStripMenuItem.Name = "CurrentCultureToolStripMenuItem"
            Me.CurrentCultureToolStripMenuItem.Size = New System.Drawing.Size(221, 22)
            Me.CurrentCultureToolStripMenuItem.Text = CurrentCultureText

            Me.EnglishCultureenUSToolStripMenuItem.Name = "EnglishCultureenUSToolStripMenuItem"
            Me.EnglishCultureenUSToolStripMenuItem.Size = New System.Drawing.Size(221, 22)
            Me.EnglishCultureenUSToolStripMenuItem.Text = EnglishCultureText

            Me.InvariantCultureToolStripMenuItem.Name = "InvariantCultureToolStripMenuItem"
            Me.InvariantCultureToolStripMenuItem.Size = New System.Drawing.Size(221, 22)
            Me.InvariantCultureToolStripMenuItem.Text = InternationCultureText

            Me.CultureOfOperatingSystemToolStripMenuItem.Name = "CultureOfOperatingSystemToolStripMenuItem"
            Me.CultureOfOperatingSystemToolStripMenuItem.Size = New System.Drawing.Size(221, 22)
            Me.CultureOfOperatingSystemToolStripMenuItem.Text = OsCultureText

            Me.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CopyFullTableToClipboardToolStripMenuItem, Me.CopySelectedCellsToClipboardWithHeadersToolStripMenuItem, Me.CopySelectedCellsToClipboardWithoutHeadersToolStripMenuItem, Me.ExportCultureOptionsToolStripMenuItem})

            InitCultureContextMenu()

        End Sub

        'TODO: implement in CM.Data?
        ''' <summary>
        ''' Creates a DataTabe only with the cells the user has selected.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function CreateDataTableFromMarkedCells() As DataTable
            If Me.Grid Is Nothing Then Throw New Exception("DataGridView hasn't been assigned, yet")
            Dim dataSourceTable As DataTable = CType(Me.Grid.DataSource, DataTable)
            Dim newTable As New DataTable
            Dim addedRows As New System.Collections.Generic.List(Of Integer)
            Dim neededCols As New System.Collections.Generic.List(Of Integer) 'actually needed columns'

            For i As Integer = 0 To dataSourceTable.Columns.Count() - 1
                Dim col As DataColumn = dataSourceTable.Columns(i)
                newTable.Columns.Add(col.ColumnName, col.DataType)
            Next

            Dim selectedCells As DataGridViewSelectedCellCollection = Grid.SelectedCells 'unsorted :/'
            Dim gridviewCells(selectedCells.Count - 1) As GridViewCellMetaInfo

            Dim counter As Integer = 0
            For Each cell As DataGridViewCell In selectedCells
                gridviewCells(counter) = New GridViewCellMetaInfo(cell.RowIndex, cell.ColumnIndex, cell.OwningRow)
                counter += 1
            Next
            Array.Sort(gridviewCells)

            For Each cell As GridViewCellMetaInfo In gridviewCells
                Dim rowIndex As Integer = cell.rowIndex
                Dim colIndex As Integer = cell.columnIndex
                If Not neededCols.Contains(colIndex) Then neededCols.Add(colIndex)
                If Not addedRows.Contains(rowIndex) Then
                    Dim newRow As DataRow = newTable.NewRow
                    Dim rowCells As DataGridViewCellCollection = cell.owningRow.Cells
                    For Each rowCell As DataGridViewCell In rowCells
                        If rowCell.Selected Then
                            newRow(rowCell.ColumnIndex) = rowCell.Value
                        End If
                    Next
                    newTable.Rows.Add(newRow)
                    addedRows.Add(rowIndex)
                End If
            Next

            For i As Integer = newTable.Columns.Count() - 1 To 0 Step -1
                If Not neededCols.Contains(i) Then newTable.Columns.RemoveAt(i)
            Next

            newTable.AcceptChanges()
            Return newTable
        End Function

        Private Sub ContextMenuStripDataGrid_Opening(sender As System.Object, e As System.ComponentModel.CancelEventArgs) Handles Me.Opening
            If Me.DesignMode Then
                CopySelectedCellsToClipboardWithHeadersToolStripMenuItem.Visible = True
                CopySelectedCellsToClipboardWithoutHeadersToolStripMenuItem.Visible = True
            ElseIf Not Grid Is Nothing Then
                If Grid.AreAllCellsSelected(False) Then
                    CopySelectedCellsToClipboardWithHeadersToolStripMenuItem.Visible = False
                    CopySelectedCellsToClipboardWithoutHeadersToolStripMenuItem.Visible = False
                Else
                    CopySelectedCellsToClipboardWithHeadersToolStripMenuItem.Visible = True
                    CopySelectedCellsToClipboardWithoutHeadersToolStripMenuItem.Visible = True
                End If
            End If
        End Sub

        ''' <summary>
        ''' Copies the full table into the clipboard
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub CopyFullTableToClipboardToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles CopyFullTableToClipboardToolStripMenuItem.Click
            If Grid.DataSource IsNot Nothing Then
                Clipboard.SetDataObject(CompuMaster.Data.Csv.ConvertDataTableToTextAsStringBuilder(CType(Grid.DataSource, DataTable), True, PreferredCulture(), ControlChars.Tab).ToString)
            End If
        End Sub

        ''' <summary>
        ''' Copies the selected cells into the clipboard including headers
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub CopySelectedCellsToClipboardwithHeadersToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles CopySelectedCellsToClipboardWithHeadersToolStripMenuItem.Click
            Dim newTable As DataTable = CreateDataTableFromMarkedCells()
            Clipboard.SetDataObject(CompuMaster.Data.Csv.ConvertDataTableToTextAsStringBuilder(newTable, True, PreferredCulture(), ControlChars.Tab).ToString)
        End Sub

        ''' <summary>
        ''' Copies the selected cells into the clipboard without headers
        ''' </summary>e
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub CopySelectedCellsToClipboardwithoutHeadersToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles CopySelectedCellsToClipboardWithoutHeadersToolStripMenuItem.Click
            Dim newTable As DataTable = CreateDataTableFromMarkedCells()
            Clipboard.SetDataObject(CompuMaster.Data.Csv.ConvertDataTableToTextAsStringBuilder(newTable, False, PreferredCulture(), ControlChars.Tab).ToString)
        End Sub

        ''' <summary>
        '''  Returns the selected culture
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function PreferredCulture() As Globalization.CultureInfo
            If InvariantCultureToolStripMenuItem.Checked = True Then
                Return Globalization.CultureInfo.InvariantCulture
            ElseIf CultureOfOperatingSystemToolStripMenuItem.Checked = True Then
                Return Globalization.CultureInfo.InstalledUICulture
            ElseIf EnglishCultureenUSToolStripMenuItem.Checked Then
                Return Globalization.CultureInfo.GetCultureInfo("en-US")
            ElseIf CurrentCultureToolStripMenuItem.Checked Then
                Return Globalization.CultureInfo.CurrentCulture
            Else
                Throw New NotImplementedException("Culture not defined")
            End If
        End Function

        Private Sub CurrentCultureToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles CurrentCultureToolStripMenuItem.Click
            ResetCultureSelection()
            CurrentCultureToolStripMenuItem.Checked = True
        End Sub

        Private Sub InvariantCultureToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles InvariantCultureToolStripMenuItem.Click
            ResetCultureSelection()
            InvariantCultureToolStripMenuItem.Checked = True
        End Sub

        Private Sub ResetCultureSelection()
            CurrentCultureToolStripMenuItem.Checked = False
            InvariantCultureToolStripMenuItem.Checked = False
            EnglishCultureenUSToolStripMenuItem.Checked = False
            CultureOfOperatingSystemToolStripMenuItem.Checked = False
        End Sub

        Private Sub InitCultureContextMenu()
            CurrentCultureToolStripMenuItem.Text &= " (" & Globalization.CultureInfo.CurrentCulture.DisplayName & ")"
            CultureOfOperatingSystemToolStripMenuItem.Text &= " (" & Globalization.CultureInfo.InstalledUICulture.DisplayName & ")"
        End Sub

        Private Sub CultureOfOperatingSystemToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles CultureOfOperatingSystemToolStripMenuItem.Click
            ResetCultureSelection()
            CultureOfOperatingSystemToolStripMenuItem.Checked = True
        End Sub

        Private Sub EnglishCultureenUSToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles EnglishCultureenUSToolStripMenuItem.Click
            ResetCultureSelection()
            EnglishCultureenUSToolStripMenuItem.Checked = True
        End Sub

        Protected Overrides Sub Dispose(disposing As Boolean)
            If disposing Then
                Me.CopySelectedCellsToClipboardWithHeadersToolStripMenuItem.Dispose()
                Me.CopySelectedCellsToClipboardWithoutHeadersToolStripMenuItem.Dispose()
                Me.ExportCultureOptionsToolStripMenuItem.Dispose()
                Me.CurrentCultureToolStripMenuItem.Dispose()
                Me.EnglishCultureenUSToolStripMenuItem.Dispose()
                Me.InvariantCultureToolStripMenuItem.Dispose()
                Me.CultureOfOperatingSystemToolStripMenuItem.Dispose()
            End If
            MyBase.Dispose(disposing)
        End Sub

        ''' <summary>
        ''' Required class to sort the selected cells collection 
        ''' </summary>
        ''' <remarks></remarks>
        Private Class GridViewCellMetaInfo
            Implements IComparable
            Public rowIndex As Integer
            Public columnIndex As Integer
            Public owningRow As DataGridViewRow

            Public Sub New(ByVal rowIndex As Integer, ByVal colIndex As Integer, ByVal owningRow As DataGridViewRow)
                Me.rowIndex = rowIndex
                Me.columnIndex = colIndex
                Me.owningRow = owningRow
            End Sub

            Public Function CompareTo(obj As Object) As Integer Implements IComparable.CompareTo
                If TypeOf obj Is GridViewCellMetaInfo Then
                    Dim otherStruct As GridViewCellMetaInfo = CType(obj, GridViewCellMetaInfo)
                    If rowIndex < otherStruct.rowIndex Then
                        Return -1
                    ElseIf rowIndex > otherStruct.rowIndex Then
                        Return 1
                    Else
                        Return 0
                    End If
                End If
                Throw New ArgumentException("obj is not a GridViewCell")
            End Function
        End Class

    End Class

End Namespace

