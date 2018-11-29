<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Diafrm_Output_Word
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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.CheckBox_PlanView = New System.Windows.Forms.CheckBox()
        Me.ProgressBar_PlanView = New System.Windows.Forms.ProgressBar()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.CheckBox_SectionalView = New System.Windows.Forms.CheckBox()
        Me.ProgressBar_SectionalView = New System.Windows.Forms.ProgressBar()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.ListBoxMonitor_Static = New System.Windows.Forms.ListBox()
        Me.ListBoxMonitor_Dynamic = New System.Windows.Forms.ListBox()
        Me.LabelDate = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.btnExport = New System.Windows.Forms.Button()
        Me.ChkBxSelect = New System.Windows.Forms.CheckBox()
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.CheckBox_PlanView)
        Me.GroupBox1.Controls.Add(Me.ProgressBar_PlanView)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.CheckBox_SectionalView)
        Me.GroupBox1.Controls.Add(Me.ProgressBar_SectionalView)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.ListBoxMonitor_Static)
        Me.GroupBox1.Controls.Add(Me.ListBoxMonitor_Dynamic)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(454, 245)
        Me.GroupBox1.TabIndex = 20
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "选择要进行同步滚动和结果输出的对象"
        '
        'CheckBox_PlanView
        '
        Me.CheckBox_PlanView.AutoSize = True
        Me.CheckBox_PlanView.Location = New System.Drawing.Point(9, 30)
        Me.CheckBox_PlanView.Name = "CheckBox_PlanView"
        Me.CheckBox_PlanView.Size = New System.Drawing.Size(84, 16)
        Me.CheckBox_PlanView.TabIndex = 23
        Me.CheckBox_PlanView.Text = "开挖平面图"
        Me.CheckBox_PlanView.UseVisualStyleBackColor = True
        '
        'ProgressBar_PlanView
        '
        Me.ProgressBar_PlanView.Location = New System.Drawing.Point(100, 29)
        Me.ProgressBar_PlanView.Name = "ProgressBar_PlanView"
        Me.ProgressBar_PlanView.Size = New System.Drawing.Size(100, 16)
        Me.ProgressBar_PlanView.TabIndex = 24
        '
        'Label5
        '
        Me.Label5.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(234, 59)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(95, 12)
        Me.Label5.TabIndex = 18
        Me.Label5.Text = "监测曲线 - 静态"
        '
        'CheckBox_SectionalView
        '
        Me.CheckBox_SectionalView.AutoSize = True
        Me.CheckBox_SectionalView.Location = New System.Drawing.Point(236, 30)
        Me.CheckBox_SectionalView.Name = "CheckBox_SectionalView"
        Me.CheckBox_SectionalView.Size = New System.Drawing.Size(84, 16)
        Me.CheckBox_SectionalView.TabIndex = 23
        Me.CheckBox_SectionalView.Text = "开挖剖面图"
        Me.CheckBox_SectionalView.UseVisualStyleBackColor = True
        '
        'ProgressBar_SectionalView
        '
        Me.ProgressBar_SectionalView.Location = New System.Drawing.Point(327, 29)
        Me.ProgressBar_SectionalView.Name = "ProgressBar_SectionalView"
        Me.ProgressBar_SectionalView.Size = New System.Drawing.Size(100, 16)
        Me.ProgressBar_SectionalView.TabIndex = 24
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(7, 59)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(95, 12)
        Me.Label3.TabIndex = 18
        Me.Label3.Text = "监测曲线 - 动态"
        '
        'ListBoxMonitor_Static
        '
        Me.ListBoxMonitor_Static.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ListBoxMonitor_Static.FormattingEnabled = True
        Me.ListBoxMonitor_Static.HorizontalScrollbar = True
        Me.ListBoxMonitor_Static.ItemHeight = 12
        Me.ListBoxMonitor_Static.Location = New System.Drawing.Point(236, 74)
        Me.ListBoxMonitor_Static.Name = "ListBoxMonitor_Static"
        Me.ListBoxMonitor_Static.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.ListBoxMonitor_Static.Size = New System.Drawing.Size(191, 160)
        Me.ListBoxMonitor_Static.TabIndex = 17
        '
        'ListBoxMonitor_Dynamic
        '
        Me.ListBoxMonitor_Dynamic.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.ListBoxMonitor_Dynamic.FormattingEnabled = True
        Me.ListBoxMonitor_Dynamic.HorizontalScrollbar = True
        Me.ListBoxMonitor_Dynamic.ItemHeight = 12
        Me.ListBoxMonitor_Dynamic.Location = New System.Drawing.Point(9, 74)
        Me.ListBoxMonitor_Dynamic.Name = "ListBoxMonitor_Dynamic"
        Me.ListBoxMonitor_Dynamic.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.ListBoxMonitor_Dynamic.Size = New System.Drawing.Size(191, 160)
        Me.ListBoxMonitor_Dynamic.TabIndex = 17
        '
        'LabelDate
        '
        Me.LabelDate.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.LabelDate.Font = New System.Drawing.Font("Times New Roman", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelDate.Location = New System.Drawing.Point(69, 269)
        Me.LabelDate.Name = "LabelDate"
        Me.LabelDate.Size = New System.Drawing.Size(93, 21)
        Me.LabelDate.TabIndex = 14
        Me.LabelDate.Text = "2014/09/28"
        Me.LabelDate.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(10, 274)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(53, 12)
        Me.Label4.TabIndex = 18
        Me.Label4.Text = "施工日期"
        '
        'btnExport
        '
        Me.btnExport.Location = New System.Drawing.Point(391, 269)
        Me.btnExport.Name = "btnExport"
        Me.btnExport.Size = New System.Drawing.Size(75, 23)
        Me.btnExport.TabIndex = 19
        Me.btnExport.Text = "结果输出"
        Me.btnExport.UseVisualStyleBackColor = True
        '
        'ChkBxSelect
        '
        Me.ChkBxSelect.AutoSize = True
        Me.ChkBxSelect.Location = New System.Drawing.Point(183, 273)
        Me.ChkBxSelect.Name = "ChkBxSelect"
        Me.ChkBxSelect.Size = New System.Drawing.Size(138, 16)
        Me.ChkBxSelect.TabIndex = 21
        Me.ChkBxSelect.Text = "Select/DeSelect All"
        Me.ChkBxSelect.ThreeState = True
        Me.ChkBxSelect.UseVisualStyleBackColor = True
        '
        'BackgroundWorker1
        '
        Me.BackgroundWorker1.WorkerReportsProgress = True
        Me.BackgroundWorker1.WorkerSupportsCancellation = True
        '
        'Diafrm_Output_Word
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(478, 301)
        Me.Controls.Add(Me.ChkBxSelect)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.btnExport)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.LabelDate)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Diafrm_Output_Word"
        Me.Text = "输出"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents ListBoxMonitor_Static As System.Windows.Forms.ListBox
    Friend WithEvents ListBoxMonitor_Dynamic As System.Windows.Forms.ListBox
    Friend WithEvents LabelDate As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents btnExport As System.Windows.Forms.Button
    Friend WithEvents ChkBxSelect As System.Windows.Forms.CheckBox
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents CheckBox_PlanView As System.Windows.Forms.CheckBox
    Friend WithEvents ProgressBar_PlanView As System.Windows.Forms.ProgressBar
    Friend WithEvents CheckBox_SectionalView As System.Windows.Forms.CheckBox
    Friend WithEvents ProgressBar_SectionalView As System.Windows.Forms.ProgressBar
End Class
