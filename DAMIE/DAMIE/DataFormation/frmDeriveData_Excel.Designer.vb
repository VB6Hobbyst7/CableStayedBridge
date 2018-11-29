<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDeriveData_Excel
    Inherits frmDeriveData
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmDeriveData_Excel))
        Me.SuspendLayout()
        '
        'btnExport
        '
        '
        'frmDeriveData_Excel
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.ClientSize = New System.Drawing.Size(539, 396)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmDeriveData_Excel"
        Me.Text = "从Excel中提取数据"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    ''' <summary>
    ''' 只有在程序运行时才能显示出来的界面更新效果
    ''' </summary>
    Private Sub InitializeComponent_ActivateAtRuntime()
        Dim SheetName = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Dim RangeName = New System.Windows.Forms.DataGridViewTextBoxColumn()
        'SheetName
        SheetName.HeaderText = "工作表名称"
        SheetName.Name = "SheetName"
        SheetName.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        SheetName.ToolTipText = "提取的工作表名称包含于要进行检索的工作表名称，比如输入""CX""，则会提取工作簿中第一个名称中含有""CX""的工作表。" & vbCrLf &
            "每一个工作表名称都会在用来保存数据的工作簿中创建一个对应的工作表。"
        SheetName.Width = 183
        '
        'RangeName
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        DataGridViewCellStyle1.ForeColor = System.Drawing.Color.Blue
        RangeName.DefaultCellStyle = DataGridViewCellStyle1
        RangeName.HeaderText = "区域范围"
        RangeName.Name = "RangeName"
        RangeName.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        RangeName.ToolTipText = "示例： A1:C3,如果要引用一张表中不连续的两个区域，可以使用""A1:A3,C1:C3"""
        '
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle2.Font = New System.Drawing.Font("SimSun", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.MyDataGridView1.RowTemplate.Height = 23
        Me.MyDataGridView1.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle2
        Me.MyDataGridView1.ColumnHeadersHeight = 25
        Me.MyDataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.MyDataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {SheetName, RangeName})

    End Sub

End Class
