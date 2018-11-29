Imports DAMIE.AME_UserControl

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmRolling
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
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.LabelDate = New System.Windows.Forms.Label()
        Me.btnRefresh = New System.Windows.Forms.Button()
        Me.btnOutPut = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.btn_GroupHandle = New System.Windows.Forms.Button()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.CheckBox_PlanView = New System.Windows.Forms.CheckBox()
        Me.ProgressBar_PlanView = New System.Windows.Forms.ProgressBar()
        Me.CheckBox_SectionalView = New System.Windows.Forms.CheckBox()
        Me.ProgressBar_SectionalView = New System.Windows.Forms.ProgressBar()
        Me.Panel_Roll = New System.Windows.Forms.Panel()
        Me.btnRoll = New System.Windows.Forms.Button()
        Me.NumChanging = New DAMIE.AME_UserControl.UsrCtrl_NumberChanging()
        Me.Calendar_Construction = New System.Windows.Forms.MonthCalendar()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.ListBoxMonitorData = New DAMIE.AME_UserControl.ListBox_NoReplyForKeyDown()
        Me.GroupBox1.SuspendLayout
        Me.Panel2.SuspendLayout
        Me.Panel_Roll.SuspendLayout
        Me.SuspendLayout
        '
        'Label4
        '
        Me.Label4.AutoSize = true
        Me.Label4.Location = New System.Drawing.Point(3, 44)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(53, 12)
        Me.Label4.TabIndex = 18
        Me.Label4.Text = "施工日期"
        '
        'Label3
        '
        Me.Label3.AutoSize = true
        Me.Label3.Location = New System.Drawing.Point(276, 21)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(77, 12)
        Me.Label3.TabIndex = 18
        Me.Label3.Text = "监测数据曲线"
        '
        'LabelDate
        '
        Me.LabelDate.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.LabelDate.Font = New System.Drawing.Font("Times New Roman", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0,Byte))
        Me.LabelDate.Location = New System.Drawing.Point(156, 40)
        Me.LabelDate.Name = "LabelDate"
        Me.LabelDate.Size = New System.Drawing.Size(93, 21)
        Me.LabelDate.TabIndex = 14
        Me.LabelDate.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btnRefresh
        '
        Me.btnRefresh.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right),System.Windows.Forms.AnchorStyles)
        Me.btnRefresh.Location = New System.Drawing.Point(371, 316)
        Me.btnRefresh.Name = "btnRefresh"
        Me.btnRefresh.Size = New System.Drawing.Size(75, 23)
        Me.btnRefresh.TabIndex = 19
        Me.btnRefresh.Text = "刷新(&R)"
        Me.btnRefresh.UseVisualStyleBackColor = true
        '
        'btnOutPut
        '
        Me.btnOutPut.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right),System.Windows.Forms.AnchorStyles)
        Me.btnOutPut.Location = New System.Drawing.Point(371, 345)
        Me.btnOutPut.Name = "btnOutPut"
        Me.btnOutPut.Size = New System.Drawing.Size(75, 23)
        Me.btnOutPut.TabIndex = 20
        Me.btnOutPut.Text = "输出..."
        Me.btnOutPut.UseVisualStyleBackColor = true
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom)  _
            Or System.Windows.Forms.AnchorStyles.Left)  _
            Or System.Windows.Forms.AnchorStyles.Right),System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.btn_GroupHandle)
        Me.GroupBox1.Controls.Add(Me.Panel2)
        Me.GroupBox1.Controls.Add(Me.Panel_Roll)
        Me.GroupBox1.Controls.Add(Me.btnOutPut)
        Me.GroupBox1.Controls.Add(Me.btnRefresh)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.ListBoxMonitorData)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.MinimumSize = New System.Drawing.Size(459, 380)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(459, 380)
        Me.GroupBox1.TabIndex = 19
        Me.GroupBox1.TabStop = false
        Me.GroupBox1.Text = "选择要进行同步滚动和结果输出的对象"
        '
        'btn_GroupHandle
        '
        Me.btn_GroupHandle.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right),System.Windows.Forms.AnchorStyles)
        Me.btn_GroupHandle.Location = New System.Drawing.Point(290, 345)
        Me.btn_GroupHandle.Name = "btn_GroupHandle"
        Me.btn_GroupHandle.Size = New System.Drawing.Size(75, 23)
        Me.btn_GroupHandle.TabIndex = 26
        Me.btn_GroupHandle.Text = "批量操作"
        Me.btn_GroupHandle.UseVisualStyleBackColor = true
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.CheckBox_PlanView)
        Me.Panel2.Controls.Add(Me.ProgressBar_PlanView)
        Me.Panel2.Controls.Add(Me.CheckBox_SectionalView)
        Me.Panel2.Controls.Add(Me.ProgressBar_SectionalView)
        Me.Panel2.Location = New System.Drawing.Point(12, 36)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(258, 69)
        Me.Panel2.TabIndex = 25
        '
        'CheckBox_PlanView
        '
        Me.CheckBox_PlanView.AutoSize = true
        Me.CheckBox_PlanView.Location = New System.Drawing.Point(6, 11)
        Me.CheckBox_PlanView.Name = "CheckBox_PlanView"
        Me.CheckBox_PlanView.Size = New System.Drawing.Size(84, 16)
        Me.CheckBox_PlanView.TabIndex = 23
        Me.CheckBox_PlanView.Text = "开挖平面图"
        Me.CheckBox_PlanView.UseVisualStyleBackColor = true
        '
        'ProgressBar_PlanView
        '
        Me.ProgressBar_PlanView.Location = New System.Drawing.Point(97, 10)
        Me.ProgressBar_PlanView.Name = "ProgressBar_PlanView"
        Me.ProgressBar_PlanView.Size = New System.Drawing.Size(143, 16)
        Me.ProgressBar_PlanView.TabIndex = 24
        '
        'CheckBox_SectionalView
        '
        Me.CheckBox_SectionalView.AutoSize = true
        Me.CheckBox_SectionalView.Location = New System.Drawing.Point(6, 44)
        Me.CheckBox_SectionalView.Name = "CheckBox_SectionalView"
        Me.CheckBox_SectionalView.Size = New System.Drawing.Size(84, 16)
        Me.CheckBox_SectionalView.TabIndex = 23
        Me.CheckBox_SectionalView.Text = "开挖剖面图"
        Me.CheckBox_SectionalView.UseVisualStyleBackColor = true
        '
        'ProgressBar_SectionalView
        '
        Me.ProgressBar_SectionalView.Location = New System.Drawing.Point(97, 43)
        Me.ProgressBar_SectionalView.Name = "ProgressBar_SectionalView"
        Me.ProgressBar_SectionalView.Size = New System.Drawing.Size(143, 16)
        Me.ProgressBar_SectionalView.TabIndex = 24
        '
        'Panel_Roll
        '
        Me.Panel_Roll.Controls.Add(Me.btnRoll)
        Me.Panel_Roll.Controls.Add(Me.LabelDate)
        Me.Panel_Roll.Controls.Add(Me.NumChanging)
        Me.Panel_Roll.Controls.Add(Me.Calendar_Construction)
        Me.Panel_Roll.Controls.Add(Me.Label4)
        Me.Panel_Roll.Controls.Add(Me.Label5)
        Me.Panel_Roll.Location = New System.Drawing.Point(12, 117)
        Me.Panel_Roll.Name = "Panel_Roll"
        Me.Panel_Roll.Size = New System.Drawing.Size(258, 251)
        Me.Panel_Roll.TabIndex = 22
        '
        'btnRoll
        '
        Me.btnRoll.Location = New System.Drawing.Point(63, 36)
        Me.btnRoll.Name = "btnRoll"
        Me.btnRoll.Size = New System.Drawing.Size(75, 23)
        Me.btnRoll.TabIndex = 22
        Me.btnRoll.Text = "滚动"
        Me.btnRoll.UseVisualStyleBackColor = true
        '
        'NumChanging
        '
        Me.NumChanging.BackColor = System.Drawing.Color.Transparent
        Me.NumChanging.Location = New System.Drawing.Point(61, 7)
        Me.NumChanging.Name = "NumChanging"
        Me.NumChanging.Size = New System.Drawing.Size(190, 21)
        Me.NumChanging.TabIndex = 21
        Me.NumChanging.unit = DAMIE.AME_UserControl.UsrCtrl_NumberChanging.YearMonthDay.Days
        '
        'Calendar_Construction
        '
        Me.Calendar_Construction.Location = New System.Drawing.Point(5, 65)
        Me.Calendar_Construction.MaxDate = New Date(2014, 10, 5, 0, 0, 0, 0)
        Me.Calendar_Construction.MinDate = New Date(2013, 1, 1, 0, 0, 0, 0)
        Me.Calendar_Construction.Name = "Calendar_Construction"
        Me.Calendar_Construction.ShowTodayCircle = false
        Me.Calendar_Construction.ShowWeekNumbers = true
        Me.Calendar_Construction.TabIndex = 13
        '
        'Label5
        '
        Me.Label5.AutoSize = true
        Me.Label5.Location = New System.Drawing.Point(3, 10)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(53, 12)
        Me.Label5.TabIndex = 18
        Me.Label5.Text = "增减日期"
        '
        'ListBoxMonitorData
        '
        Me.ListBoxMonitorData.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom)  _
            Or System.Windows.Forms.AnchorStyles.Left)  _
            Or System.Windows.Forms.AnchorStyles.Right),System.Windows.Forms.AnchorStyles)
        Me.ListBoxMonitorData.FormattingEnabled = true
        Me.ListBoxMonitorData.HorizontalScrollbar = true
        Me.ListBoxMonitorData.ItemHeight = 12
        Me.ListBoxMonitorData.Location = New System.Drawing.Point(276, 39)
        Me.ListBoxMonitorData.MinimumSize = New System.Drawing.Size(4, 184)
        Me.ListBoxMonitorData.Name = "ListBoxMonitorData"
        Me.ListBoxMonitorData.ParentControl = Me.NumChanging
        Me.ListBoxMonitorData.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.ListBoxMonitorData.Size = New System.Drawing.Size(170, 256)
        Me.ListBoxMonitorData.TabIndex = 17
        '
        'frmRolling
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6!, 12!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(483, 404)
        Me.Controls.Add(Me.GroupBox1)
        Me.MinimumSize = New System.Drawing.Size(457, 432)
        Me.Name = "frmRolling"
        Me.Text = "动态同步控制"
        Me.GroupBox1.ResumeLayout(false)
        Me.GroupBox1.PerformLayout
        Me.Panel2.ResumeLayout(false)
        Me.Panel2.PerformLayout
        Me.Panel_Roll.ResumeLayout(false)
        Me.Panel_Roll.PerformLayout
        Me.ResumeLayout(false)

End Sub


    Friend WithEvents ListBoxMonitorData As ListBox_NoReplyForKeyDown

    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents LabelDate As System.Windows.Forms.Label
    Friend WithEvents btnRefresh As System.Windows.Forms.Button
    Friend WithEvents btnOutPut As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Calendar_Construction As System.Windows.Forms.MonthCalendar
    Friend WithEvents NumChanging As DAMIE.AME_UserControl.UsrCtrl_NumberChanging
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents CheckBox_PlanView As System.Windows.Forms.CheckBox
    Friend WithEvents ProgressBar_PlanView As System.Windows.Forms.ProgressBar
    Friend WithEvents CheckBox_SectionalView As System.Windows.Forms.CheckBox
    Friend WithEvents ProgressBar_SectionalView As System.Windows.Forms.ProgressBar
    Friend WithEvents Panel_Roll As System.Windows.Forms.Panel
    Friend WithEvents btn_GroupHandle As System.Windows.Forms.Button
    Friend WithEvents btnRoll As System.Windows.Forms.Button
End Class
