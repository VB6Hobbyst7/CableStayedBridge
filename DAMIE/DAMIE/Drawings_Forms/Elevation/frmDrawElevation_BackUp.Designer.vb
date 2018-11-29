<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDrawElevation_BackUp
    Inherits System.Windows.Forms.Form

    'Form 重写 Dispose，以清理组件列表。
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

    'Windows 窗体设计器所必需的
    Private components As System.ComponentModel.IContainer

    '注意: 以下过程是 Windows 窗体设计器所必需的
    '可以使用 Windows 窗体设计器修改它。
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.lbChooseExcav = New System.Windows.Forms.Label()
        Me.lstbxChooseRegion = New System.Windows.Forms.ListBox()
        Me.btnGenerate = New System.Windows.Forms.Button()
        Me.OpenFileDlg_DataExcel = New System.Windows.Forms.OpenFileDialog()
        Me.btnChooseAll = New System.Windows.Forms.Button()
        Me.btnChooseNone = New System.Windows.Forms.Button()
        Me.BGW_Generate = New System.ComponentModel.BackgroundWorker()
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ToolStripMenuItemRefresh = New System.Windows.Forms.ToolStripMenuItem()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'lbChooseExcav
        '
        Me.lbChooseExcav.AutoSize = True
        Me.lbChooseExcav.Location = New System.Drawing.Point(11, 12)
        Me.lbChooseExcav.Name = "lbChooseExcav"
        Me.lbChooseExcav.Size = New System.Drawing.Size(89, 12)
        Me.lbChooseExcav.TabIndex = 2
        Me.lbChooseExcav.Text = "进行对比的区域"
        '
        'lstbxChooseRegion
        '
        Me.lstbxChooseRegion.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lstbxChooseRegion.FormattingEnabled = True
        Me.lstbxChooseRegion.HorizontalScrollbar = True
        Me.lstbxChooseRegion.ItemHeight = 12
        Me.lstbxChooseRegion.Location = New System.Drawing.Point(13, 32)
        Me.lstbxChooseRegion.Name = "lstbxChooseRegion"
        Me.lstbxChooseRegion.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple
        Me.lstbxChooseRegion.Size = New System.Drawing.Size(237, 184)
        Me.lstbxChooseRegion.TabIndex = 3
        '
        'btnGenerate
        '
        Me.btnGenerate.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnGenerate.Location = New System.Drawing.Point(165, 264)
        Me.btnGenerate.Name = "btnGenerate"
        Me.btnGenerate.Size = New System.Drawing.Size(85, 25)
        Me.btnGenerate.TabIndex = 4
        Me.btnGenerate.Text = "Generate"
        Me.btnGenerate.UseVisualStyleBackColor = True
        '
        'OpenFileDlg_DataExcel
        '
        Me.OpenFileDlg_DataExcel.FileName = "OpenFileDialog1"
        '
        'btnChooseAll
        '
        Me.btnChooseAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnChooseAll.Location = New System.Drawing.Point(12, 229)
        Me.btnChooseAll.Name = "btnChooseAll"
        Me.btnChooseAll.Size = New System.Drawing.Size(75, 25)
        Me.btnChooseAll.TabIndex = 5
        Me.btnChooseAll.Text = "全选"
        Me.btnChooseAll.UseVisualStyleBackColor = True
        '
        'btnChooseNone
        '
        Me.btnChooseNone.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnChooseNone.Location = New System.Drawing.Point(13, 264)
        Me.btnChooseNone.Name = "btnChooseNone"
        Me.btnChooseNone.Size = New System.Drawing.Size(75, 25)
        Me.btnChooseNone.TabIndex = 5
        Me.btnChooseNone.Text = "全不选"
        Me.btnChooseNone.UseVisualStyleBackColor = True
        '
        'BGW_Generate
        '
        Me.BGW_Generate.WorkerReportsProgress = True
        Me.BGW_Generate.WorkerSupportsCancellation = True
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripMenuItemRefresh})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(117, 26)
        '
        'ToolStripMenuItemRefresh
        '
        Me.ToolStripMenuItemRefresh.Name = "ToolStripMenuItemRefresh"
        Me.ToolStripMenuItemRefresh.Size = New System.Drawing.Size(116, 22)
        Me.ToolStripMenuItemRefresh.Text = "刷新(&R)"
        '
        'frmDrawElevation
        '
        Me.AcceptButton = Me.btnGenerate
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(262, 301)
        Me.Controls.Add(Me.btnChooseNone)
        Me.Controls.Add(Me.btnChooseAll)
        Me.Controls.Add(Me.btnGenerate)
        Me.Controls.Add(Me.lstbxChooseRegion)
        Me.Controls.Add(Me.lbChooseExcav)
        Me.MinimumSize = New System.Drawing.Size(278, 339)
        Me.Name = "frmDrawElevation"
        Me.Text = "生成剖面图"
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lbChooseExcav As System.Windows.Forms.Label
    Friend WithEvents lstbxChooseRegion As System.Windows.Forms.ListBox
    Friend WithEvents btnGenerate As System.Windows.Forms.Button
    Friend WithEvents OpenFileDlg_DataExcel As System.Windows.Forms.OpenFileDialog
    Friend WithEvents btnChooseAll As System.Windows.Forms.Button
    Friend WithEvents btnChooseNone As System.Windows.Forms.Button
    Friend WithEvents BGW_Generate As System.ComponentModel.BackgroundWorker
    Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents ToolStripMenuItemRefresh As System.Windows.Forms.ToolStripMenuItem
End Class
