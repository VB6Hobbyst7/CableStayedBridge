
Namespace GlobalApp_Form
    <Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
    Partial Class APPLICATION_MAINFORM
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
            Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(APPLICATION_MAINFORM))
            Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
            Me.BackgroundWorker = New System.ComponentModel.BackgroundWorker()
            Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
            Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
            Me.ProgressBar1 = New System.Windows.Forms.ToolStripProgressBar()
            Me.StatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
            Me.NotifyIcon1 = New System.Windows.Forms.NotifyIcon(Me.components)
            Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
            Me.TlStrpBtn_Roll = New System.Windows.Forms.ToolStripButton()
            Me.ToolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator()
            Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
            Me.MenuItemFile = New System.Windows.Forms.ToolStripMenuItem()
            Me.MenuItem_NewProject = New System.Windows.Forms.ToolStripMenuItem()
            Me.MenuItem_OpenProject = New System.Windows.Forms.ToolStripMenuItem()
            Me.MenuItem_EditProject = New System.Windows.Forms.ToolStripMenuItem()
            Me.MenuItem_CloseProject = New System.Windows.Forms.ToolStripMenuItem()
            Me.ToolStripSeparator4 = New System.Windows.Forms.ToolStripSeparator()
            Me.MenuItem_SaveProject = New System.Windows.Forms.ToolStripMenuItem()
            Me.MenuItem_SaveAsProject = New System.Windows.Forms.ToolStripMenuItem()
            Me.ToolStripSeparator5 = New System.Windows.Forms.ToolStripSeparator()
            Me.MenuItemExport = New System.Windows.Forms.ToolStripMenuItem()
            Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
            Me.MenuItemPreference = New System.Windows.Forms.ToolStripMenuItem()
            Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
            Me.MenuItemExit = New System.Windows.Forms.ToolStripMenuItem()
            Me.MenuItemEdit = New System.Windows.Forms.ToolStripMenuItem()
            Me.MenuItemDrawingPoints = New System.Windows.Forms.ToolStripMenuItem()
            Me.MenuItemExtractData = New System.Windows.Forms.ToolStripMenuItem()
            Me.ToolStripMenuItemExtractDataFromExcel = New System.Windows.Forms.ToolStripMenuItem()
            Me.ToolStripMenuItemExtractDataFromWord = New System.Windows.Forms.ToolStripMenuItem()
            Me.VisioToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
            Me.MenuItemNew = New System.Windows.Forms.ToolStripMenuItem()
            Me.MenuItemSectionalView = New System.Windows.Forms.ToolStripMenuItem()
            Me.MenuItemPlanView = New System.Windows.Forms.ToolStripMenuItem()
            Me.MenuItemMntData_Incline = New System.Windows.Forms.ToolStripMenuItem()
            Me.MenuItemMntData_Others = New System.Windows.Forms.ToolStripMenuItem()
            Me.MenuItem_Window = New System.Windows.Forms.ToolStripMenuItem()
            Me.MenuItem_Arrange_Vertical = New System.Windows.Forms.ToolStripMenuItem()
            Me.MenuItem_Arrange_Horizontal = New System.Windows.Forms.ToolStripMenuItem()
            Me.MenuItem_Arrange_Cascade = New System.Windows.Forms.ToolStripMenuItem()
            Me.TSMenuItem_Union = New System.Windows.Forms.ToolStripMenuItem()
            Me.MenuItemHelp = New System.Windows.Forms.ToolStripMenuItem()
            Me.AboutAMEToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
            Me.StatusStrip1.SuspendLayout()
            Me.ToolStrip1.SuspendLayout()
            Me.MenuStrip1.SuspendLayout()
            Me.SuspendLayout()
            '
            'OpenFileDialog1
            '
            Me.OpenFileDialog1.FileName = "OpenFileDialog1"
            '
            'BackgroundWorker
            '
            Me.BackgroundWorker.WorkerReportsProgress = True
            Me.BackgroundWorker.WorkerSupportsCancellation = True
            '
            'StatusStrip1
            '
            Me.StatusStrip1.BackColor = System.Drawing.Color.FromArgb(CType(CType(88, Byte), Integer), CType(CType(53, Byte), Integer), CType(CType(98, Byte), Integer))
            Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ProgressBar1, Me.StatusLabel1})
            Me.StatusStrip1.Location = New System.Drawing.Point(0, 540)
            Me.StatusStrip1.Name = "StatusStrip1"
            Me.StatusStrip1.Size = New System.Drawing.Size(784, 22)
            Me.StatusStrip1.TabIndex = 3
            Me.StatusStrip1.Text = "StatusStrip1"
            '
            'ProgressBar1
            '
            Me.ProgressBar1.Name = "ProgressBar1"
            Me.ProgressBar1.Size = New System.Drawing.Size(250, 16)
            Me.ProgressBar1.Visible = False
            '
            'StatusLabel1
            '
            Me.StatusLabel1.ForeColor = System.Drawing.SystemColors.ButtonFace
            Me.StatusLabel1.Name = "StatusLabel1"
            Me.StatusLabel1.Size = New System.Drawing.Size(17, 17)
            Me.StatusLabel1.Text = "..."
            Me.StatusLabel1.Visible = False
            '
            'NotifyIcon1
            '
            Me.NotifyIcon1.Icon = CType(resources.GetObject("NotifyIcon1.Icon"), System.Drawing.Icon)
            Me.NotifyIcon1.Text = "基坑群实测数据动态分析"
            '
            'ToolStrip1
            '
            Me.ToolStrip1.BackColor = System.Drawing.Color.White
            Me.ToolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden
            Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.TlStrpBtn_Roll, Me.ToolStripSeparator3})
            Me.ToolStrip1.Location = New System.Drawing.Point(0, 25)
            Me.ToolStrip1.Name = "ToolStrip1"
            Me.ToolStrip1.RenderMode = System.Windows.Forms.ToolStripRenderMode.Professional
            Me.ToolStrip1.Size = New System.Drawing.Size(784, 25)
            Me.ToolStrip1.TabIndex = 5
            Me.ToolStrip1.Text = "ToolStrip1"
            '
            'TlStrpBtn_Roll
            '
            Me.TlStrpBtn_Roll.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
            Me.TlStrpBtn_Roll.Image = Global.DAMIE.My.Resources.Resources.btn_Roll
            Me.TlStrpBtn_Roll.ImageTransparentColor = System.Drawing.Color.Magenta
            Me.TlStrpBtn_Roll.Name = "TlStrpBtn_Roll"
            Me.TlStrpBtn_Roll.Size = New System.Drawing.Size(23, 22)
            Me.TlStrpBtn_Roll.Text = "ToolStripButton1"
            Me.TlStrpBtn_Roll.ToolTipText = "同步滚动"
            '
            'ToolStripSeparator3
            '
            Me.ToolStripSeparator3.Name = "ToolStripSeparator3"
            Me.ToolStripSeparator3.Size = New System.Drawing.Size(6, 25)
            '
            'MenuStrip1
            '
            Me.MenuStrip1.BackColor = System.Drawing.SystemColors.Control
            Me.MenuStrip1.BackgroundImage = Global.DAMIE.My.Resources.Resources.菜单栏
            Me.MenuStrip1.GripMargin = New System.Windows.Forms.Padding(0)
            Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MenuItemFile, Me.MenuItemEdit, Me.MenuItemNew, Me.MenuItem_Window, Me.TSMenuItem_Union})
            Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
            Me.MenuStrip1.Name = "MenuStrip1"
            Me.MenuStrip1.RenderMode = System.Windows.Forms.ToolStripRenderMode.System
            Me.MenuStrip1.Size = New System.Drawing.Size(784, 25)
            Me.MenuStrip1.TabIndex = 1
            Me.MenuStrip1.Text = "主菜单栏"
            '
            'MenuItemFile
            '
            Me.MenuItemFile.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MenuItem_NewProject, Me.MenuItem_OpenProject, Me.MenuItem_EditProject, Me.MenuItem_CloseProject, Me.ToolStripSeparator4, Me.MenuItem_SaveProject, Me.MenuItem_SaveAsProject, Me.ToolStripSeparator5, Me.MenuItemExport, Me.ToolStripSeparator1, Me.MenuItemPreference, Me.ToolStripSeparator2, Me.MenuItemExit})
            Me.MenuItemFile.Name = "MenuItemFile"
            Me.MenuItemFile.Size = New System.Drawing.Size(58, 21)
            Me.MenuItemFile.Text = "文件(&F)"
            '
            'MenuItem_NewProject
            '
            Me.MenuItem_NewProject.Name = "MenuItem_NewProject"
            Me.MenuItem_NewProject.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.N), System.Windows.Forms.Keys)
            Me.MenuItem_NewProject.Size = New System.Drawing.Size(199, 22)
            Me.MenuItem_NewProject.Text = "新建项目(&N)"
            '
            'MenuItem_OpenProject
            '
            Me.MenuItem_OpenProject.Name = "MenuItem_OpenProject"
            Me.MenuItem_OpenProject.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.O), System.Windows.Forms.Keys)
            Me.MenuItem_OpenProject.Size = New System.Drawing.Size(199, 22)
            Me.MenuItem_OpenProject.Text = "打开项目(&O)"
            '
            'MenuItem_EditProject
            '
            Me.MenuItem_EditProject.Name = "MenuItem_EditProject"
            Me.MenuItem_EditProject.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.E), System.Windows.Forms.Keys)
            Me.MenuItem_EditProject.Size = New System.Drawing.Size(199, 22)
            Me.MenuItem_EditProject.Text = "编辑项目"
            '
            'MenuItem_CloseProject
            '
            Me.MenuItem_CloseProject.Name = "MenuItem_CloseProject"
            Me.MenuItem_CloseProject.Size = New System.Drawing.Size(199, 22)
            Me.MenuItem_CloseProject.Text = "关闭项目"
            '
            'ToolStripSeparator4
            '
            Me.ToolStripSeparator4.Name = "ToolStripSeparator4"
            Me.ToolStripSeparator4.Size = New System.Drawing.Size(196, 6)
            '
            'MenuItem_SaveProject
            '
            Me.MenuItem_SaveProject.Name = "MenuItem_SaveProject"
            Me.MenuItem_SaveProject.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.S), System.Windows.Forms.Keys)
            Me.MenuItem_SaveProject.Size = New System.Drawing.Size(199, 22)
            Me.MenuItem_SaveProject.Text = "保存(&S)"
            '
            'MenuItem_SaveAsProject
            '
            Me.MenuItem_SaveAsProject.Name = "MenuItem_SaveAsProject"
            Me.MenuItem_SaveAsProject.ShortcutKeys = CType(((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.Shift) _
                Or System.Windows.Forms.Keys.S), System.Windows.Forms.Keys)
            Me.MenuItem_SaveAsProject.Size = New System.Drawing.Size(199, 22)
            Me.MenuItem_SaveAsProject.Text = "另存为..."
            '
            'ToolStripSeparator5
            '
            Me.ToolStripSeparator5.Name = "ToolStripSeparator5"
            Me.ToolStripSeparator5.Size = New System.Drawing.Size(196, 6)
            '
            'MenuItemExport
            '
            Me.MenuItemExport.Name = "MenuItemExport"
            Me.MenuItemExport.Size = New System.Drawing.Size(199, 22)
            Me.MenuItemExport.Text = "结果输出到Word(&E)..."
            '
            'ToolStripSeparator1
            '
            Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
            Me.ToolStripSeparator1.Size = New System.Drawing.Size(196, 6)
            '
            'MenuItemPreference
            '
            Me.MenuItemPreference.Name = "MenuItemPreference"
            Me.MenuItemPreference.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.P), System.Windows.Forms.Keys)
            Me.MenuItemPreference.Size = New System.Drawing.Size(199, 22)
            Me.MenuItemPreference.Text = "选项(&P)"
            '
            'ToolStripSeparator2
            '
            Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
            Me.ToolStripSeparator2.Size = New System.Drawing.Size(196, 6)
            '
            'MenuItemExit
            '
            Me.MenuItemExit.Name = "MenuItemExit"
            Me.MenuItemExit.ShortcutKeys = CType((System.Windows.Forms.Keys.Alt Or System.Windows.Forms.Keys.F4), System.Windows.Forms.Keys)
            Me.MenuItemExit.Size = New System.Drawing.Size(199, 22)
            Me.MenuItemExit.Text = "退出(&Q)"
            '
            'MenuItemEdit
            '
            Me.MenuItemEdit.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MenuItemDrawingPoints, Me.MenuItemExtractData})
            Me.MenuItemEdit.Name = "MenuItemEdit"
            Me.MenuItemEdit.Size = New System.Drawing.Size(59, 21)
            Me.MenuItemEdit.Text = "编辑(&E)"
            '
            'MenuItemDrawingPoints
            '
            Me.MenuItemDrawingPoints.Name = "MenuItemDrawingPoints"
            Me.MenuItemDrawingPoints.Size = New System.Drawing.Size(139, 22)
            Me.MenuItemDrawingPoints.Text = "绘制测点(&P)"
            '
            'MenuItemExtractData
            '
            Me.MenuItemExtractData.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripMenuItemExtractDataFromExcel, Me.ToolStripMenuItemExtractDataFromWord, Me.VisioToolStripMenuItem})
            Me.MenuItemExtractData.Name = "MenuItemExtractData"
            Me.MenuItemExtractData.Size = New System.Drawing.Size(139, 22)
            Me.MenuItemExtractData.Text = "数据提取..."
            '
            'ToolStripMenuItemExtractDataFromExcel
            '
            Me.ToolStripMenuItemExtractDataFromExcel.Image = Global.DAMIE.My.Resources.Resources.DatafromExcel
            Me.ToolStripMenuItemExtractDataFromExcel.Name = "ToolStripMenuItemExtractDataFromExcel"
            Me.ToolStripMenuItemExtractDataFromExcel.Size = New System.Drawing.Size(142, 22)
            Me.ToolStripMenuItemExtractDataFromExcel.Text = "Excel (&E) ..."
            '
            'ToolStripMenuItemExtractDataFromWord
            '
            Me.ToolStripMenuItemExtractDataFromWord.Image = Global.DAMIE.My.Resources.Resources.DataFromWord
            Me.ToolStripMenuItemExtractDataFromWord.Name = "ToolStripMenuItemExtractDataFromWord"
            Me.ToolStripMenuItemExtractDataFromWord.Size = New System.Drawing.Size(142, 22)
            Me.ToolStripMenuItemExtractDataFromWord.Text = "Word(&W) ..."
            '
            'VisioToolStripMenuItem
            '
            Me.VisioToolStripMenuItem.Image = Global.DAMIE.My.Resources.Resources.IdToShape
            Me.VisioToolStripMenuItem.Name = "VisioToolStripMenuItem"
            Me.VisioToolStripMenuItem.Size = New System.Drawing.Size(142, 22)
            Me.VisioToolStripMenuItem.Text = "Visio (&V) ..."
            '
            'MenuItemNew
            '
            Me.MenuItemNew.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MenuItemSectionalView, Me.MenuItemPlanView, Me.MenuItemMntData_Incline, Me.MenuItemMntData_Others})
            Me.MenuItemNew.Name = "MenuItemNew"
            Me.MenuItemNew.Size = New System.Drawing.Size(61, 21)
            Me.MenuItemNew.Text = "绘图(&D)"
            '
            'MenuItemSectionalView
            '
            Me.MenuItemSectionalView.Name = "MenuItemSectionalView"
            Me.MenuItemSectionalView.Size = New System.Drawing.Size(178, 22)
            Me.MenuItemSectionalView.Text = "开挖剖面图(&S)"
            '
            'MenuItemPlanView
            '
            Me.MenuItemPlanView.Name = "MenuItemPlanView"
            Me.MenuItemPlanView.Size = New System.Drawing.Size(178, 22)
            Me.MenuItemPlanView.Text = "开挖平面图(&P)"
            '
            'MenuItemMntData_Incline
            '
            Me.MenuItemMntData_Incline.Name = "MenuItemMntData_Incline"
            Me.MenuItemMntData_Incline.Size = New System.Drawing.Size(178, 22)
            Me.MenuItemMntData_Incline.Text = "测斜曲线图(&M)"
            '
            'MenuItemMntData_Others
            '
            Me.MenuItemMntData_Others.Name = "MenuItemMntData_Others"
            Me.MenuItemMntData_Others.Size = New System.Drawing.Size(178, 22)
            Me.MenuItemMntData_Others.Text = "其他监测曲线图(&O)"
            '
            'MenuItem_Window
            '
            Me.MenuItem_Window.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MenuItem_Arrange_Vertical, Me.MenuItem_Arrange_Horizontal, Me.MenuItem_Arrange_Cascade})
            Me.MenuItem_Window.Name = "MenuItem_Window"
            Me.MenuItem_Window.Size = New System.Drawing.Size(64, 21)
            Me.MenuItem_Window.Text = "窗口(&W)"
            '
            'MenuItem_Arrange_Vertical
            '
            Me.MenuItem_Arrange_Vertical.Name = "MenuItem_Arrange_Vertical"
            Me.MenuItem_Arrange_Vertical.Size = New System.Drawing.Size(140, 22)
            Me.MenuItem_Arrange_Vertical.Text = "垂直并排(&V)"
            '
            'MenuItem_Arrange_Horizontal
            '
            Me.MenuItem_Arrange_Horizontal.Name = "MenuItem_Arrange_Horizontal"
            Me.MenuItem_Arrange_Horizontal.Size = New System.Drawing.Size(140, 22)
            Me.MenuItem_Arrange_Horizontal.Text = "水平并排(&V)"
            '
            'MenuItem_Arrange_Cascade
            '
            Me.MenuItem_Arrange_Cascade.Name = "MenuItem_Arrange_Cascade"
            Me.MenuItem_Arrange_Cascade.Size = New System.Drawing.Size(140, 22)
            Me.MenuItem_Arrange_Cascade.Text = "层叠(&C)"
            '
            'TSMenuItem_Union
            '
            Me.TSMenuItem_Union.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MenuItemHelp, Me.AboutAMEToolStripMenuItem})
            Me.TSMenuItem_Union.Name = "TSMenuItem_Union"
            Me.TSMenuItem_Union.Size = New System.Drawing.Size(61, 21)
            Me.TSMenuItem_Union.Text = "帮助(&H)"
            '
            'MenuItemHelp
            '
            Me.MenuItemHelp.Name = "MenuItemHelp"
            Me.MenuItemHelp.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.F1), System.Windows.Forms.Keys)
            Me.MenuItemHelp.Size = New System.Drawing.Size(184, 22)
            Me.MenuItemHelp.Text = "View Help"
            '
            'AboutAMEToolStripMenuItem
            '
            Me.AboutAMEToolStripMenuItem.Name = "AboutAMEToolStripMenuItem"
            Me.AboutAMEToolStripMenuItem.Size = New System.Drawing.Size(184, 22)
            Me.AboutAMEToolStripMenuItem.Text = "About DAMIE"
            '
            'APPLICATION_MAINFORM
            '
            Me.AllowDrop = True
            Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.BackColor = System.Drawing.SystemColors.Control
            Me.ClientSize = New System.Drawing.Size(784, 562)
            Me.Controls.Add(Me.ToolStrip1)
            Me.Controls.Add(Me.StatusStrip1)
            Me.Controls.Add(Me.MenuStrip1)
            Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
            Me.IsMdiContainer = True
            Me.MainMenuStrip = Me.MenuStrip1
            Me.Name = "APPLICATION_MAINFORM"
            Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
            Me.Text = "DAMIE"
            Me.StatusStrip1.ResumeLayout(False)
            Me.StatusStrip1.PerformLayout()
            Me.ToolStrip1.ResumeLayout(False)
            Me.ToolStrip1.PerformLayout()
            Me.MenuStrip1.ResumeLayout(False)
            Me.MenuStrip1.PerformLayout()
            Me.ResumeLayout(False)
            Me.PerformLayout()

        End Sub
        Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
        Friend WithEvents MenuItemFile As System.Windows.Forms.ToolStripMenuItem
        Friend WithEvents MenuItemExit As System.Windows.Forms.ToolStripMenuItem
        Friend WithEvents MenuItemEdit As System.Windows.Forms.ToolStripMenuItem
        Friend WithEvents MenuItem_Window As System.Windows.Forms.ToolStripMenuItem
        Friend WithEvents MenuItem_Arrange_Horizontal As System.Windows.Forms.ToolStripMenuItem
        Friend WithEvents MenuItem_Arrange_Cascade As System.Windows.Forms.ToolStripMenuItem
        Friend WithEvents TSMenuItem_Union As System.Windows.Forms.ToolStripMenuItem
        Friend WithEvents MenuItemExport As System.Windows.Forms.ToolStripMenuItem
        Friend WithEvents ToolStripSeparator2 As System.Windows.Forms.ToolStripSeparator
        Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
        Friend WithEvents MenuItemNew As System.Windows.Forms.ToolStripMenuItem
        Friend WithEvents MenuItemSectionalView As System.Windows.Forms.ToolStripMenuItem
        Friend WithEvents MenuItemPlanView As System.Windows.Forms.ToolStripMenuItem
        Friend WithEvents MenuItemMntData_Incline As System.Windows.Forms.ToolStripMenuItem
        Friend WithEvents MenuItemMntData_Others As System.Windows.Forms.ToolStripMenuItem
        Friend WithEvents MenuItem_Arrange_Vertical As System.Windows.Forms.ToolStripMenuItem
        Friend WithEvents MenuItemHelp As System.Windows.Forms.ToolStripMenuItem
        Friend WithEvents AboutAMEToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
        Friend WithEvents BackgroundWorker As System.ComponentModel.BackgroundWorker
        Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
        Friend WithEvents MenuItemPreference As System.Windows.Forms.ToolStripMenuItem
        Friend WithEvents MenuItemDrawingPoints As System.Windows.Forms.ToolStripMenuItem
        Friend WithEvents MenuItem_NewProject As System.Windows.Forms.ToolStripMenuItem
        Friend WithEvents MenuItem_OpenProject As System.Windows.Forms.ToolStripMenuItem
        Friend WithEvents ToolStripSeparator4 As System.Windows.Forms.ToolStripSeparator
        Friend WithEvents MenuItem_EditProject As System.Windows.Forms.ToolStripMenuItem
        Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
        Friend WithEvents MenuItem_SaveProject As System.Windows.Forms.ToolStripMenuItem
        Friend WithEvents MenuItem_SaveAsProject As System.Windows.Forms.ToolStripMenuItem
        Friend WithEvents ToolStripSeparator5 As System.Windows.Forms.ToolStripSeparator
        Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
        Friend WithEvents ProgressBar1 As System.Windows.Forms.ToolStripProgressBar
        Friend WithEvents StatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
        Friend WithEvents NotifyIcon1 As System.Windows.Forms.NotifyIcon
        Friend WithEvents MenuItemExtractData As System.Windows.Forms.ToolStripMenuItem
        Friend WithEvents ToolStripMenuItemExtractDataFromExcel As System.Windows.Forms.ToolStripMenuItem
        Friend WithEvents ToolStripMenuItemExtractDataFromWord As System.Windows.Forms.ToolStripMenuItem
        Friend WithEvents ToolStrip1 As System.Windows.Forms.ToolStrip
        Friend WithEvents TlStrpBtn_Roll As System.Windows.Forms.ToolStripButton
        Friend WithEvents ToolStripSeparator3 As System.Windows.Forms.ToolStripSeparator
        Friend WithEvents VisioToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
        Friend WithEvents MenuItem_CloseProject As System.Windows.Forms.ToolStripMenuItem
    End Class
End Namespace