<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DiaFrm_PointsTreeView
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
        Me.TreeViewPoints = New System.Windows.Forms.TreeView()
        Me.ListBoxChosenItems = New System.Windows.Forms.ListBox()
        Me.BtnOk = New System.Windows.Forms.Button()
        Me.BtnRemove = New System.Windows.Forms.Button()
        Me.BtnClear = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.btnPreview = New System.Windows.Forms.Button()
        Me.BtnAdd = New System.Windows.Forms.Button()
        Me.BtnColpsAll = New System.Windows.Forms.Button()
        Me.btnClearCheckedNode = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'TreeViewPoints
        '
        Me.TreeViewPoints.CheckBoxes = True
        Me.TreeViewPoints.FullRowSelect = True
        Me.TreeViewPoints.Indent = 20
        Me.TreeViewPoints.ItemHeight = 14
        Me.TreeViewPoints.Location = New System.Drawing.Point(12, 37)
        Me.TreeViewPoints.Name = "TreeViewPoints"
        Me.TreeViewPoints.ShowLines = False
        Me.TreeViewPoints.Size = New System.Drawing.Size(210, 340)
        Me.TreeViewPoints.TabIndex = 0
        '
        'ListBoxChosenItems
        '
        Me.ListBoxChosenItems.FormattingEnabled = True
        Me.ListBoxChosenItems.HorizontalScrollbar = True
        Me.ListBoxChosenItems.ItemHeight = 12
        Me.ListBoxChosenItems.Location = New System.Drawing.Point(309, 37)
        Me.ListBoxChosenItems.Name = "ListBoxChosenItems"
        Me.ListBoxChosenItems.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.ListBoxChosenItems.Size = New System.Drawing.Size(214, 340)
        Me.ListBoxChosenItems.TabIndex = 1
        '
        'BtnOk
        '
        Me.BtnOk.Location = New System.Drawing.Point(448, 392)
        Me.BtnOk.Name = "BtnOk"
        Me.BtnOk.Size = New System.Drawing.Size(75, 23)
        Me.BtnOk.TabIndex = 0
        Me.BtnOk.Text = "确定"
        Me.BtnOk.UseVisualStyleBackColor = True
        '
        'BtnRemove
        '
        Me.BtnRemove.Location = New System.Drawing.Point(228, 207)
        Me.BtnRemove.Name = "BtnRemove"
        Me.BtnRemove.Size = New System.Drawing.Size(75, 23)
        Me.BtnRemove.TabIndex = 3
        Me.BtnRemove.Text = "<== 移除"
        Me.BtnRemove.UseVisualStyleBackColor = True
        '
        'BtnClear
        '
        Me.BtnClear.Location = New System.Drawing.Point(228, 273)
        Me.BtnClear.Name = "BtnClear"
        Me.BtnClear.Size = New System.Drawing.Size(75, 23)
        Me.BtnClear.TabIndex = 3
        Me.BtnClear.Text = "清空(&C)"
        Me.BtnClear.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 13)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(113, 12)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "选择单个或多个测点"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(307, 13)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(53, 12)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "选择结果"
        '
        'btnPreview
        '
        Me.btnPreview.Location = New System.Drawing.Point(313, 392)
        Me.btnPreview.Name = "btnPreview"
        Me.btnPreview.Size = New System.Drawing.Size(75, 23)
        Me.btnPreview.TabIndex = 3
        Me.btnPreview.Text = "应用(&P)"
        Me.btnPreview.UseVisualStyleBackColor = True
        '
        'BtnAdd
        '
        Me.BtnAdd.Location = New System.Drawing.Point(228, 159)
        Me.BtnAdd.Name = "BtnAdd"
        Me.BtnAdd.Size = New System.Drawing.Size(75, 23)
        Me.BtnAdd.TabIndex = 3
        Me.BtnAdd.Text = "==> 更新"
        Me.BtnAdd.UseVisualStyleBackColor = True
        '
        'BtnColpsAll
        '
        Me.BtnColpsAll.Location = New System.Drawing.Point(14, 392)
        Me.BtnColpsAll.Name = "BtnColpsAll"
        Me.BtnColpsAll.Size = New System.Drawing.Size(117, 23)
        Me.BtnColpsAll.TabIndex = 3
        Me.BtnColpsAll.Text = "Collapse All(&C)"
        Me.BtnColpsAll.UseVisualStyleBackColor = True
        '
        'btnClearCheckedNode
        '
        Me.btnClearCheckedNode.Location = New System.Drawing.Point(147, 392)
        Me.btnClearCheckedNode.Name = "btnClearCheckedNode"
        Me.btnClearCheckedNode.Size = New System.Drawing.Size(75, 23)
        Me.btnClearCheckedNode.TabIndex = 5
        Me.btnClearCheckedNode.Text = "清空选择"
        Me.btnClearCheckedNode.UseVisualStyleBackColor = True
        '
        'DiaFrm_PointsTreeView
        '
        Me.AcceptButton = Me.BtnOk
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(539, 427)
        Me.Controls.Add(Me.btnClearCheckedNode)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.BtnColpsAll)
        Me.Controls.Add(Me.btnPreview)
        Me.Controls.Add(Me.BtnClear)
        Me.Controls.Add(Me.BtnAdd)
        Me.Controls.Add(Me.BtnRemove)
        Me.Controls.Add(Me.BtnOk)
        Me.Controls.Add(Me.ListBoxChosenItems)
        Me.Controls.Add(Me.TreeViewPoints)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Name = "DiaFrm_PointsTreeView"
        Me.Text = "选择测点标志"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TreeViewPoints As System.Windows.Forms.TreeView
    Friend WithEvents ListBoxChosenItems As System.Windows.Forms.ListBox
    Friend WithEvents BtnOk As System.Windows.Forms.Button
    Friend WithEvents BtnRemove As System.Windows.Forms.Button
    Friend WithEvents BtnClear As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnPreview As System.Windows.Forms.Button
    Friend WithEvents BtnAdd As System.Windows.Forms.Button
    Friend WithEvents BtnColpsAll As System.Windows.Forms.Button
    Friend WithEvents btnClearCheckedNode As System.Windows.Forms.Button
End Class
