<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TestForm
    Inherits System.Windows.Forms.Form

    'Das Formular überschreibt den Löschvorgang, um die Komponentenliste zu bereinigen.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Wird vom Windows Form-Designer benötigt.
    Private components As System.ComponentModel.IContainer

    'Hinweis: Die folgende Prozedur ist für den Windows Form-Designer erforderlich.
    'Das Bearbeiten ist mit dem Windows Form-Designer möglich.  
    'Das Bearbeiten mit dem Code-Editor ist nicht möglich.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.DataGridView1 = New CompuMaster.Data.Windows.DataGridView()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBoxConnectionString = New System.Windows.Forms.TextBox()
        Me.Button6 = New System.Windows.Forms.Button()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.CellCopyingContextMenu1 = New CompuMaster.Data.Windows.CellCopyingContextMenu(Me.components)
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.AdditionalTestEntryToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.CellCopyingContextMenu1.SuspendLayout()
        Me.SuspendLayout()
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToOrderColumns = True
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridView1.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.DataGridView1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridView1.DefaultCellStyle = DataGridViewCellStyle2
        Me.DataGridView1.IgnoreErrorsOnLayoutingBinaryValues = True
        Me.DataGridView1.Location = New System.Drawing.Point(3, 3)
        Me.DataGridView1.MultiSelect = False
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.SaveDataChangesAfterEveryRowChange = True
        Me.DataGridView1.SelectCommand = Nothing
        Me.DataGridView1.Size = New System.Drawing.Size(696, 499)
        Me.DataGridView1.TabIndex = 0
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(-3, -1)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.DataGridView1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TextBoxConnectionString)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Button6)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Button5)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Button4)
        Me.SplitContainer1.Size = New System.Drawing.Size(1039, 505)
        Me.SplitContainer1.SplitterDistance = 702
        Me.SplitContainer1.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(3, 6)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(94, 13)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Connection String:"
        '
        'TextBoxConnectionString
        '
        Me.TextBoxConnectionString.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TextBoxConnectionString.Location = New System.Drawing.Point(3, 22)
        Me.TextBoxConnectionString.Name = "TextBoxConnectionString"
        Me.TextBoxConnectionString.Size = New System.Drawing.Size(314, 20)
        Me.TextBoxConnectionString.TabIndex = 4
        '
        'Button6
        '
        Me.Button6.Location = New System.Drawing.Point(-1, 166)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(178, 23)
        Me.Button6.TabIndex = 3
        Me.Button6.Text = "No Filter"
        Me.Button6.UseVisualStyleBackColor = True
        '
        'Button5
        '
        Me.Button5.Location = New System.Drawing.Point(-1, 125)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(178, 23)
        Me.Button5.TabIndex = 3
        Me.Button5.Text = "Filter: Independence < 1900"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(-1, 87)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(178, 23)
        Me.Button4.TabIndex = 2
        Me.Button4.Text = "Filter: Independence >= 1900"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'CellCopyingContextMenu1
        '
        Me.CellCopyingContextMenu1.CopyFullTableText = "Copy full table to clipboard"
        Me.CellCopyingContextMenu1.CopySelectedCellsWithHeadersText = "Copy selected cells to clipboard (with headers)"
        Me.CellCopyingContextMenu1.CopySelectedCellsWithoutHeadersText = "Copy selected cells to clipboard (without headers)"
        Me.CellCopyingContextMenu1.CurrentCultureText = "Current culture"
        Me.CellCopyingContextMenu1.DataGridView = Nothing
        Me.CellCopyingContextMenu1.EnglishCultureText = "English culture (en-US)"
        Me.CellCopyingContextMenu1.ExportCultureOptionsText = "Export culture options"
        Me.CellCopyingContextMenu1.InternationCultureText = "International culture"
        Me.CellCopyingContextMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripSeparator1, Me.AdditionalTestEntryToolStripMenuItem})
        Me.CellCopyingContextMenu1.Name = "ContextMenuStripDataGrid"
        Me.CellCopyingContextMenu1.OsCultureText = "Culture of operating system"
        Me.CellCopyingContextMenu1.Size = New System.Drawing.Size(338, 120)
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(334, 6)
        '
        'AdditionalTestEntryToolStripMenuItem
        '
        Me.AdditionalTestEntryToolStripMenuItem.Name = "AdditionalTestEntryToolStripMenuItem"
        Me.AdditionalTestEntryToolStripMenuItem.Size = New System.Drawing.Size(337, 22)
        Me.AdditionalTestEntryToolStripMenuItem.Text = "Additional test entry"
        '
        'TestForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1032, 503)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Name = "TestForm"
        Me.Text = "Form1"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.CellCopyingContextMenu1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents Button6 As System.Windows.Forms.Button
    Friend WithEvents CellCopyingContextMenu1 As CompuMaster.Data.Windows.CellCopyingContextMenu
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents AdditionalTestEntryToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Protected WithEvents DataGridView1 As CompuMaster.Data.Windows.DataGridView
    Friend WithEvents Label1 As Label
    Friend WithEvents TextBoxConnectionString As TextBox
End Class
