<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDeriveData_Word
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmDeriveData_Word))
        Me.SuspendLayout()
        '
        'btnExport
        '
        '
        'frmDeriveData_Word
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.ClientSize = New System.Drawing.Size(539, 396)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmDeriveData_Word"
        Me.Text = "从Word中提取数据"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    ''' <summary>
    ''' 只有在程序运行时才能显示出来的界面更新效果
    ''' </summary>
    Private Sub InitializeComponent_ActivateAtRuntime()

        Dim PointName As System.Windows.Forms.DataGridViewTextBoxColumn
        Dim DataOffset As System.Windows.Forms.DataGridViewTextBoxColumn
        Dim SearchDirection As System.Windows.Forms.DataGridViewComboBoxColumn


        PointName = New System.Windows.Forms.DataGridViewTextBoxColumn()
        DataOffset = New System.Windows.Forms.DataGridViewTextBoxColumn()
        SearchDirection = New System.Windows.Forms.DataGridViewComboBoxColumn()

        'PointName
        '
        PointName.HeaderText = "点位特征名"
        PointName.Name = "PointName"
        PointName.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        PointName.ToolTipText = "特征名是包含于实际的监测点位的，比如：特征名CX会在Word文档中搜索包含有CX的所有测点，如TCX01。"
        '


        'SearchDirection
        '
        SearchDirection.HeaderText = "搜索"
        SearchDirection.Items.AddRange(New Object() {"按行", "按列"})
        SearchDirection.Name = "SearchDirection"
        '

        'DataOffset
        '
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        DataGridViewCellStyle2.ForeColor = System.Drawing.Color.Blue
        DataOffset.DefaultCellStyle = DataGridViewCellStyle2
        DataOffset.HeaderText = "数据偏移"
        DataOffset.Name = "DataOffset"
        DataOffset.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        DataOffset.ToolTipText = "如果点位的数据在点位单元格的右侧且与之相邻，则为1"
        DataOffset.Width = 80


        'MyDataGridView1
        '
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("SimSun", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.MyDataGridView1.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.MyDataGridView1.ColumnHeadersHeight = 25
        Me.MyDataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.MyDataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {PointName, DataOffset, SearchDirection})
        Me.MyDataGridView1.Location = New System.Drawing.Point(11, 181)
        Me.MyDataGridView1.RowTemplate.Height = 23
        Me.MyDataGridView1.RowTemplate.Resizable = System.Windows.Forms.DataGridViewTriState.[False]
        Me.MyDataGridView1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.MyDataGridView1.Size = New System.Drawing.Size(346, 110)
        Me.MyDataGridView1.TabIndex = 14
    End Sub



End Class
