<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDrawing_Mnt_Others
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
        Me.btnGenerate = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.RbtnStaticWithTime = New System.Windows.Forms.RadioButton()
        Me.RbtnDynamic = New System.Windows.Forms.RadioButton()
        Me.chkBoxOpenNewExcel = New System.Windows.Forms.CheckBox()
        Me.ListBoxPointsName = New System.Windows.Forms.ListBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.listSheetsName = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.ComboBox_WorkingStage = New System.Windows.Forms.ComboBox()
        Me.btnDrawMonitorPoints = New System.Windows.Forms.Button()
        Me.BGWK_NewDrawing = New System.ComponentModel.BackgroundWorker()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ComboBox_MntType = New System.Windows.Forms.ComboBox()
        Me.Panel_Static = New System.Windows.Forms.Panel()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.ComboBoxOpenedWorkbook = New System.Windows.Forms.ComboBox()
        Me.GroupBox1.SuspendLayout()
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
        'btnGenerate
        '
        Me.btnGenerate.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnGenerate.Location = New System.Drawing.Point(336, 355)
        Me.btnGenerate.Name = "btnGenerate"
        Me.btnGenerate.Size = New System.Drawing.Size(75, 23)
        Me.btnGenerate.TabIndex = 4
        Me.btnGenerate.Text = "Generate"
        Me.btnGenerate.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.RbtnStaticWithTime)
        Me.GroupBox1.Controls.Add(Me.RbtnDynamic)
        Me.GroupBox1.Location = New System.Drawing.Point(312, 114)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(99, 73)
        Me.GroupBox1.TabIndex = 5
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "曲线图的类型"
        '
        'RbtnStaticWithTime
        '
        Me.RbtnStaticWithTime.AutoSize = True
        Me.RbtnStaticWithTime.Location = New System.Drawing.Point(7, 22)
        Me.RbtnStaticWithTime.Name = "RbtnStaticWithTime"
        Me.RbtnStaticWithTime.Size = New System.Drawing.Size(47, 16)
        Me.RbtnStaticWithTime.TabIndex = 1
        Me.RbtnStaticWithTime.TabStop = True
        Me.RbtnStaticWithTime.Text = "静态"
        Me.ToolTip1.SetToolTip(Me.RbtnStaticWithTime, "以时间为X轴，查看每一个测点在整个施工过程中的变化")
        Me.RbtnStaticWithTime.UseVisualStyleBackColor = True
        '
        'RbtnDynamic
        '
        Me.RbtnDynamic.AutoSize = True
        Me.RbtnDynamic.Location = New System.Drawing.Point(7, 45)
        Me.RbtnDynamic.Name = "RbtnDynamic"
        Me.RbtnDynamic.Size = New System.Drawing.Size(47, 16)
        Me.RbtnDynamic.TabIndex = 0
        Me.RbtnDynamic.TabStop = True
        Me.RbtnDynamic.Text = "动态"
        Me.ToolTip1.SetToolTip(Me.RbtnDynamic, "以测点为X轴，动态查看每一天的变化情况")
        Me.RbtnDynamic.UseVisualStyleBackColor = True
        '
        'chkBoxOpenNewExcel
        '
        Me.chkBoxOpenNewExcel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.chkBoxOpenNewExcel.AutoSize = True
        Me.chkBoxOpenNewExcel.Location = New System.Drawing.Point(321, 328)
        Me.chkBoxOpenNewExcel.Name = "chkBoxOpenNewExcel"
        Me.chkBoxOpenNewExcel.Size = New System.Drawing.Size(90, 16)
        Me.chkBoxOpenNewExcel.TabIndex = 6
        Me.chkBoxOpenNewExcel.Text = "打开新Excel"
        Me.chkBoxOpenNewExcel.UseVisualStyleBackColor = True
        '
        'ListBoxPointsName
        '
        Me.ListBoxPointsName.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ListBoxPointsName.FormattingEnabled = True
        Me.ListBoxPointsName.HorizontalScrollbar = True
        Me.ListBoxPointsName.ItemHeight = 12
        Me.ListBoxPointsName.Location = New System.Drawing.Point(106, 94)
        Me.ListBoxPointsName.Name = "ListBoxPointsName"
        Me.ListBoxPointsName.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.ListBoxPointsName.Size = New System.Drawing.Size(185, 280)
        Me.ListBoxPointsName.TabIndex = 7
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(11, 59)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(89, 12)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "选择数据工作表"
        '
        'listSheetsName
        '
        Me.listSheetsName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.listSheetsName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.listSheetsName.FormattingEnabled = True
        Me.listSheetsName.Location = New System.Drawing.Point(106, 56)
        Me.listSheetsName.Name = "listSheetsName"
        Me.listSheetsName.Size = New System.Drawing.Size(188, 20)
        Me.listSheetsName.TabIndex = 9
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(11, 94)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(77, 12)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "选择相应测点"
        '
        'ComboBox_WorkingStage
        '
        Me.ComboBox_WorkingStage.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_WorkingStage.FormattingEnabled = True
        Me.ComboBox_WorkingStage.Location = New System.Drawing.Point(3, 20)
        Me.ComboBox_WorkingStage.Name = "ComboBox_WorkingStage"
        Me.ComboBox_WorkingStage.Size = New System.Drawing.Size(90, 20)
        Me.ComboBox_WorkingStage.TabIndex = 1
        Me.ToolTip1.SetToolTip(Me.ComboBox_WorkingStage, "用来在绘制测斜位移的最值走势图时，在图表中标开挖工况信息")
        '
        'btnDrawMonitorPoints
        '
        Me.btnDrawMonitorPoints.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnDrawMonitorPoints.Location = New System.Drawing.Point(12, 355)
        Me.btnDrawMonitorPoints.Name = "btnDrawMonitorPoints"
        Me.btnDrawMonitorPoints.Size = New System.Drawing.Size(75, 23)
        Me.btnDrawMonitorPoints.TabIndex = 10
        Me.btnDrawMonitorPoints.Text = "绘制测点"
        Me.btnDrawMonitorPoints.UseVisualStyleBackColor = True
        '
        'BGWK_NewDrawing
        '
        Me.BGWK_NewDrawing.WorkerReportsProgress = True
        Me.BGWK_NewDrawing.WorkerSupportsCancellation = True
        '
        'Label1
        '
        Me.Label1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(311, 56)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(77, 12)
        Me.Label1.TabIndex = 12
        Me.Label1.Text = "监测数据类型"
        '
        'ComboBox_MntType
        '
        Me.ComboBox_MntType.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ComboBox_MntType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_MntType.FormattingEnabled = True
        Me.ComboBox_MntType.IntegralHeight = False
        Me.ComboBox_MntType.Location = New System.Drawing.Point(313, 76)
        Me.ComboBox_MntType.Name = "ComboBox_MntType"
        Me.ComboBox_MntType.Size = New System.Drawing.Size(90, 20)
        Me.ComboBox_MntType.TabIndex = 11
        '
        'Panel_Static
        '
        Me.Panel_Static.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel_Static.Controls.Add(Me.Label4)
        Me.Panel_Static.Controls.Add(Me.ComboBox_WorkingStage)
        Me.Panel_Static.Location = New System.Drawing.Point(312, 206)
        Me.Panel_Static.Name = "Panel_Static"
        Me.Panel_Static.Size = New System.Drawing.Size(99, 46)
        Me.Panel_Static.TabIndex = 13
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(1, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(53, 12)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "开挖工况"
        '
        'ComboBoxOpenedWorkbook
        '
        Me.ComboBoxOpenedWorkbook.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ComboBoxOpenedWorkbook.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxOpenedWorkbook.FormattingEnabled = True
        Me.ComboBoxOpenedWorkbook.Location = New System.Drawing.Point(95, 15)
        Me.ComboBoxOpenedWorkbook.Name = "ComboBoxOpenedWorkbook"
        Me.ComboBoxOpenedWorkbook.Size = New System.Drawing.Size(324, 20)
        Me.ComboBoxOpenedWorkbook.TabIndex = 14
        '
        'frmDrawing_Mnt_Others
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(431, 395)
        Me.Controls.Add(Me.ComboBoxOpenedWorkbook)
        Me.Controls.Add(Me.Panel_Static)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ComboBox_MntType)
        Me.Controls.Add(Me.chkBoxOpenNewExcel)
        Me.Controls.Add(Me.btnGenerate)
        Me.Controls.Add(Me.btnDrawMonitorPoints)
        Me.Controls.Add(Me.listSheetsName)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.ListBoxPointsName)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.btnChooseMonitorData)
        Me.MinimumSize = New System.Drawing.Size(340, 300)
        Me.Name = "frmDrawing_Mnt_Others"
        Me.Text = "其他监测曲线"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.Panel_Static.ResumeLayout(False)
        Me.Panel_Static.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnChooseMonitorData As System.Windows.Forms.Button
    Friend WithEvents btnGenerate As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents RbtnStaticWithTime As System.Windows.Forms.RadioButton
    Friend WithEvents RbtnDynamic As System.Windows.Forms.RadioButton
    Friend WithEvents chkBoxOpenNewExcel As System.Windows.Forms.CheckBox
    Friend WithEvents ListBoxPointsName As System.Windows.Forms.ListBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents listSheetsName As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents btnDrawMonitorPoints As System.Windows.Forms.Button
    Friend WithEvents BGWK_NewDrawing As System.ComponentModel.BackgroundWorker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ComboBox_MntType As System.Windows.Forms.ComboBox
    Friend WithEvents Panel_Static As System.Windows.Forms.Panel
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents ComboBox_WorkingStage As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBoxOpenedWorkbook As System.Windows.Forms.ComboBox
End Class
