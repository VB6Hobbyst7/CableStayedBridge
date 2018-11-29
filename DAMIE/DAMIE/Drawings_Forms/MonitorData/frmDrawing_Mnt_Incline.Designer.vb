<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDrawing_Mnt_Incline
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.btnChooseMonitorData = New System.Windows.Forms.Button()
        Me.ComboBox_ExcavID = New System.Windows.Forms.ComboBox()
        Me.Label_Component_Elevation = New System.Windows.Forms.Label()
        Me.btnGenerate = New System.Windows.Forms.Button()
        Me.chkBoxOpenNewExcel = New System.Windows.Forms.CheckBox()
        Me.ListBoxWorksheetsName = New System.Windows.Forms.ListBox()
        Me.ComboBox_ExcavRegion = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.btnDrawMonitorPoints = New System.Windows.Forms.Button()
        Me.BGWK_NewDrawing = New System.ComponentModel.BackgroundWorker()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.RadioButton_Dynamic = New System.Windows.Forms.RadioButton()
        Me.RadioButton_Max_Depth = New System.Windows.Forms.RadioButton()
        Me.Panel_Dynamic = New System.Windows.Forms.Panel()
        Me.ComboBox_MntType = New System.Windows.Forms.ComboBox()
        Me.Label_MntType = New System.Windows.Forms.Label()
        Me.ComboBox_WorkingStage = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Panel_Static = New System.Windows.Forms.Panel()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.ComboBoxOpenedWorkbook = New System.Windows.Forms.ComboBox()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.GroupBox1.SuspendLayout()
        Me.Panel_Dynamic.SuspendLayout()
        Me.Panel_Static.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnChooseMonitorData
        '
        Me.btnChooseMonitorData.Location = New System.Drawing.Point(13, 13)
        Me.btnChooseMonitorData.Name = "btnChooseMonitorData"
        Me.btnChooseMonitorData.Size = New System.Drawing.Size(75, 23)
        Me.btnChooseMonitorData.TabIndex = 0
        Me.btnChooseMonitorData.Text = "监测数据"
        Me.btnChooseMonitorData.UseVisualStyleBackColor = True
        '
        'ComboBox_ExcavID
        '
        Me.ComboBox_ExcavID.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_ExcavID.FormattingEnabled = True
        Me.ComboBox_ExcavID.Location = New System.Drawing.Point(5, 38)
        Me.ComboBox_ExcavID.Name = "ComboBox_ExcavID"
        Me.ComboBox_ExcavID.Size = New System.Drawing.Size(90, 20)
        Me.ComboBox_ExcavID.TabIndex = 1
        '
        'Label_Component_Elevation
        '
        Me.Label_Component_Elevation.AutoSize = True
        Me.Label_Component_Elevation.Location = New System.Drawing.Point(3, 18)
        Me.Label_Component_Elevation.Name = "Label_Component_Elevation"
        Me.Label_Component_Elevation.Size = New System.Drawing.Size(59, 12)
        Me.Label_Component_Elevation.TabIndex = 3
        Me.Label_Component_Elevation.Text = "构件-标高"
        '
        'btnGenerate
        '
        Me.btnGenerate.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnGenerate.Location = New System.Drawing.Point(354, 290)
        Me.btnGenerate.Name = "btnGenerate"
        Me.btnGenerate.Size = New System.Drawing.Size(75, 23)
        Me.btnGenerate.TabIndex = 4
        Me.btnGenerate.Text = "生成"
        Me.btnGenerate.UseVisualStyleBackColor = True
        '
        'chkBoxOpenNewExcel
        '
        Me.chkBoxOpenNewExcel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.chkBoxOpenNewExcel.AutoSize = True
        Me.chkBoxOpenNewExcel.Location = New System.Drawing.Point(319, 69)
        Me.chkBoxOpenNewExcel.Name = "chkBoxOpenNewExcel"
        Me.chkBoxOpenNewExcel.Size = New System.Drawing.Size(90, 16)
        Me.chkBoxOpenNewExcel.TabIndex = 6
        Me.chkBoxOpenNewExcel.Text = "打开新Excel"
        Me.chkBoxOpenNewExcel.UseVisualStyleBackColor = True
        '
        'ListBoxWorksheetsName
        '
        Me.ListBoxWorksheetsName.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ListBoxWorksheetsName.FormattingEnabled = True
        Me.ListBoxWorksheetsName.ItemHeight = 12
        Me.ListBoxWorksheetsName.Location = New System.Drawing.Point(12, 51)
        Me.ListBoxWorksheetsName.Name = "ListBoxWorksheetsName"
        Me.ListBoxWorksheetsName.Size = New System.Drawing.Size(161, 256)
        Me.ListBoxWorksheetsName.TabIndex = 7
        '
        'ComboBox_ExcavRegion
        '
        Me.ComboBox_ExcavRegion.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_ExcavRegion.FormattingEnabled = True
        Me.ComboBox_ExcavRegion.Location = New System.Drawing.Point(5, 141)
        Me.ComboBox_ExcavRegion.Name = "ComboBox_ExcavRegion"
        Me.ComboBox_ExcavRegion.Size = New System.Drawing.Size(90, 20)
        Me.ComboBox_ExcavRegion.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(3, 119)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(53, 12)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "施工进度"
        '
        'btnDrawMonitorPoints
        '
        Me.btnDrawMonitorPoints.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnDrawMonitorPoints.Location = New System.Drawing.Point(354, 261)
        Me.btnDrawMonitorPoints.Name = "btnDrawMonitorPoints"
        Me.btnDrawMonitorPoints.Size = New System.Drawing.Size(75, 23)
        Me.btnDrawMonitorPoints.TabIndex = 8
        Me.btnDrawMonitorPoints.Text = "绘制测点"
        Me.btnDrawMonitorPoints.UseVisualStyleBackColor = True
        '
        'BGWK_NewDrawing
        '
        Me.BGWK_NewDrawing.WorkerReportsProgress = True
        Me.BGWK_NewDrawing.WorkerSupportsCancellation = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.ComboBox1)
        Me.GroupBox1.Controls.Add(Me.CheckBox1)
        Me.GroupBox1.Controls.Add(Me.RadioButton_Dynamic)
        Me.GroupBox1.Controls.Add(Me.RadioButton_Max_Depth)
        Me.GroupBox1.Location = New System.Drawing.Point(313, 119)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(117, 131)
        Me.GroupBox1.TabIndex = 9
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "绘图类型"
        '
        'RadioButton_Dynamic
        '
        Me.RadioButton_Dynamic.AutoSize = True
        Me.RadioButton_Dynamic.Checked = True
        Me.RadioButton_Dynamic.Location = New System.Drawing.Point(6, 46)
        Me.RadioButton_Dynamic.Name = "RadioButton_Dynamic"
        Me.RadioButton_Dynamic.Size = New System.Drawing.Size(83, 16)
        Me.RadioButton_Dynamic.TabIndex = 0
        Me.RadioButton_Dynamic.TabStop = True
        Me.RadioButton_Dynamic.Text = "形状动态图"
        Me.RadioButton_Dynamic.UseVisualStyleBackColor = True
        '
        'RadioButton_Max_Depth
        '
        Me.RadioButton_Max_Depth.AutoSize = True
        Me.RadioButton_Max_Depth.Location = New System.Drawing.Point(7, 21)
        Me.RadioButton_Max_Depth.Name = "RadioButton_Max_Depth"
        Me.RadioButton_Max_Depth.Size = New System.Drawing.Size(83, 16)
        Me.RadioButton_Max_Depth.TabIndex = 0
        Me.RadioButton_Max_Depth.Text = "最值走势图"
        Me.RadioButton_Max_Depth.UseVisualStyleBackColor = True
        '
        'Panel_Dynamic
        '
        Me.Panel_Dynamic.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel_Dynamic.Controls.Add(Me.Label_Component_Elevation)
        Me.Panel_Dynamic.Controls.Add(Me.ComboBox_ExcavID)
        Me.Panel_Dynamic.Controls.Add(Me.ComboBox_ExcavRegion)
        Me.Panel_Dynamic.Controls.Add(Me.Label2)
        Me.Panel_Dynamic.Location = New System.Drawing.Point(186, 101)
        Me.Panel_Dynamic.Name = "Panel_Dynamic"
        Me.Panel_Dynamic.Size = New System.Drawing.Size(106, 194)
        Me.Panel_Dynamic.TabIndex = 10
        '
        'ComboBox_MntType
        '
        Me.ComboBox_MntType.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ComboBox_MntType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_MntType.DropDownWidth = 90
        Me.ComboBox_MntType.FormattingEnabled = True
        Me.ComboBox_MntType.Location = New System.Drawing.Point(191, 71)
        Me.ComboBox_MntType.Name = "ComboBox_MntType"
        Me.ComboBox_MntType.Size = New System.Drawing.Size(90, 20)
        Me.ComboBox_MntType.TabIndex = 1
        '
        'Label_MntType
        '
        Me.Label_MntType.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label_MntType.AutoSize = True
        Me.Label_MntType.Location = New System.Drawing.Point(189, 51)
        Me.Label_MntType.Name = "Label_MntType"
        Me.Label_MntType.Size = New System.Drawing.Size(77, 12)
        Me.Label_MntType.TabIndex = 3
        Me.Label_MntType.Text = "监测数据类型"
        '
        'ComboBox_WorkingStage
        '
        Me.ComboBox_WorkingStage.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_WorkingStage.DropDownWidth = 90
        Me.ComboBox_WorkingStage.FormattingEnabled = True
        Me.ComboBox_WorkingStage.Location = New System.Drawing.Point(5, 41)
        Me.ComboBox_WorkingStage.Name = "ComboBox_WorkingStage"
        Me.ComboBox_WorkingStage.Size = New System.Drawing.Size(90, 20)
        Me.ComboBox_WorkingStage.TabIndex = 1
        Me.ToolTip1.SetToolTip(Me.ComboBox_WorkingStage, "用来在绘制测斜位移的最值走势图时，在图表中标开挖工况信息")
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(3, 21)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(53, 12)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "开挖工况"
        '
        'Panel_Static
        '
        Me.Panel_Static.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel_Static.Controls.Add(Me.Label3)
        Me.Panel_Static.Controls.Add(Me.ComboBox_WorkingStage)
        Me.Panel_Static.Location = New System.Drawing.Point(186, 98)
        Me.Panel_Static.Name = "Panel_Static"
        Me.Panel_Static.Size = New System.Drawing.Size(106, 73)
        Me.Panel_Static.TabIndex = 11
        '
        'ComboBoxOpenedWorkbook
        '
        Me.ComboBoxOpenedWorkbook.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ComboBoxOpenedWorkbook.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxOpenedWorkbook.FormattingEnabled = True
        Me.ComboBoxOpenedWorkbook.Location = New System.Drawing.Point(94, 15)
        Me.ComboBoxOpenedWorkbook.Name = "ComboBoxOpenedWorkbook"
        Me.ComboBoxOpenedWorkbook.Size = New System.Drawing.Size(335, 20)
        Me.ComboBoxOpenedWorkbook.TabIndex = 12
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(6, 77)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(60, 16)
        Me.CheckBox1.TabIndex = 1
        Me.CheckBox1.Text = "自定义"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'ComboBox1
        '
        Me.ComboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Items.AddRange(New Object() {"空间分布1"})
        Me.ComboBox1.Location = New System.Drawing.Point(6, 100)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(90, 20)
        Me.ComboBox1.TabIndex = 2
        '
        'frmDrawing_Mnt_Incline
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(442, 318)
        Me.Controls.Add(Me.ComboBoxOpenedWorkbook)
        Me.Controls.Add(Me.Panel_Static)
        Me.Controls.Add(Me.Label_MntType)
        Me.Controls.Add(Me.ComboBox_MntType)
        Me.Controls.Add(Me.Panel_Dynamic)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.btnDrawMonitorPoints)
        Me.Controls.Add(Me.ListBoxWorksheetsName)
        Me.Controls.Add(Me.chkBoxOpenNewExcel)
        Me.Controls.Add(Me.btnGenerate)
        Me.Controls.Add(Me.btnChooseMonitorData)
        Me.MinimumSize = New System.Drawing.Size(333, 303)
        Me.Name = "frmDrawing_Mnt_Incline"
        Me.Text = "测斜曲线绘制"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.Panel_Dynamic.ResumeLayout(False)
        Me.Panel_Dynamic.PerformLayout()
        Me.Panel_Static.ResumeLayout(False)
        Me.Panel_Static.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnChooseMonitorData As System.Windows.Forms.Button
    Friend WithEvents ComboBox_ExcavID As System.Windows.Forms.ComboBox
    Friend WithEvents Label_Component_Elevation As System.Windows.Forms.Label
    Friend WithEvents btnGenerate As System.Windows.Forms.Button
    Friend WithEvents chkBoxOpenNewExcel As System.Windows.Forms.CheckBox
    Friend WithEvents ListBoxWorksheetsName As System.Windows.Forms.ListBox
    Friend WithEvents ComboBox_ExcavRegion As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnDrawMonitorPoints As System.Windows.Forms.Button
    Friend WithEvents BGWK_NewDrawing As System.ComponentModel.BackgroundWorker
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents RadioButton_Dynamic As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton_Max_Depth As System.Windows.Forms.RadioButton
    Friend WithEvents Panel_Dynamic As System.Windows.Forms.Panel
    Friend WithEvents ComboBox_MntType As System.Windows.Forms.ComboBox
    Friend WithEvents Label_MntType As System.Windows.Forms.Label
    Friend WithEvents ComboBox_WorkingStage As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Panel_Static As System.Windows.Forms.Panel
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents ComboBoxOpenedWorkbook As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
End Class
