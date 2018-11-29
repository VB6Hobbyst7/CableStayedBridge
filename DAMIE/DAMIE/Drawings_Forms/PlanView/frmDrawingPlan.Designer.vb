<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDrawingPlan
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBoxInfoBoxID = New System.Windows.Forms.TextBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.TextBoxAllRegions = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.TextBoxPageName = New System.Windows.Forms.TextBox()
        Me.BtnGenerate = New System.Windows.Forms.Button()
        Me.TextBoxFilePath = New System.Windows.Forms.TextBox()
        Me.btnChooseVisioPlanView = New System.Windows.Forms.Button()
        Me.ChkBx_PointInfo = New System.Windows.Forms.CheckBox()
        Me.txtbx_ShapeName_MonitorPointTag = New System.Windows.Forms.TextBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.txtbx_Pt_UR_CAD_Y = New System.Windows.Forms.TextBox()
        Me.txtbx_Pt_BL_CAD_Y = New System.Windows.Forms.TextBox()
        Me.txtbx_Pt_BL_CAD_X = New System.Windows.Forms.TextBox()
        Me.txtbx_Pt_UR_ShapeID = New System.Windows.Forms.TextBox()
        Me.txtbx_Pt_UR_CAD_X = New System.Windows.Forms.TextBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.txtbx_Pt_BL_ShapeID = New System.Windows.Forms.TextBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Btn_Import = New System.Windows.Forms.Button()
        Me.Btn_Export = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout
        Me.Panel1.SuspendLayout
        Me.TableLayoutPanel1.SuspendLayout
        Me.SuspendLayout
        '
        'Label1
        '
        Me.Label1.AutoSize = true
        Me.Label1.Location = New System.Drawing.Point(6, 23)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(65, 12)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "信息文本框"
        Me.ToolTip1.SetToolTip(Me.Label1, "记录开挖信息的文本框的ID值")
        '
        'TextBoxInfoBoxID
        '
        Me.TextBoxInfoBoxID.Location = New System.Drawing.Point(100, 20)
        Me.TextBoxInfoBoxID.Name = "TextBoxInfoBoxID"
        Me.TextBoxInfoBoxID.Size = New System.Drawing.Size(72, 21)
        Me.TextBoxInfoBoxID.TabIndex = 1
        Me.TextBoxInfoBoxID.Text = "5079"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.TextBoxAllRegions)
        Me.GroupBox1.Controls.Add(Me.TextBoxInfoBoxID)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Location = New System.Drawing.Point(6, 84)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(356, 56)
        Me.GroupBox1.TabIndex = 2
        Me.GroupBox1.TabStop = false
        Me.GroupBox1.Text = "特征形状ID值"
        '
        'TextBoxAllRegions
        '
        Me.TextBoxAllRegions.Location = New System.Drawing.Point(272, 20)
        Me.TextBoxAllRegions.Name = "TextBoxAllRegions"
        Me.TextBoxAllRegions.Size = New System.Drawing.Size(72, 21)
        Me.TextBoxAllRegions.TabIndex = 1
        Me.TextBoxAllRegions.Text = "5078"
        '
        'Label2
        '
        Me.Label2.AutoSize = true
        Me.Label2.Location = New System.Drawing.Point(178, 23)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(53, 12)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "所有分区"
        Me.ToolTip1.SetToolTip(Me.Label2, "所有分区的组合形状的ID值")
        '
        'Label4
        '
        Me.Label4.AutoSize = true
        Me.Label4.Location = New System.Drawing.Point(12, 48)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(53, 12)
        Me.Label4.TabIndex = 0
        Me.Label4.Text = "页面名称"
        Me.ToolTip1.SetToolTip(Me.Label4, "开挖平面图在Visio中的页面名称")
        '
        'Label7
        '
        Me.Label7.AutoSize = true
        Me.Label7.Location = New System.Drawing.Point(5, 9)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(209, 12)
        Me.Label7.TabIndex = 0
        Me.Label7.Text = "测点主控形状中表示编号的形状的Name"
        Me.ToolTip1.SetToolTip(Me.Label7, "visio中在监测点的主控形状中，用来显示测点编号的形状的Name属性")
        '
        'Label8
        '
        Me.Label8.AutoSize = true
        Me.Label8.Location = New System.Drawing.Point(3, 20)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(95, 12)
        Me.Label8.TabIndex = 0
        Me.Label8.Text = "Visio中的形状ID"
        Me.ToolTip1.SetToolTip(Me.Label8, "Visio平面图中用于坐标变换的两个定位点的形状ID，"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"这两个点分别代表ABCD基坑群的左下角与右上角。")
        '
        'Label9
        '
        Me.Label9.AutoSize = true
        Me.Label9.Location = New System.Drawing.Point(3, 40)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(101, 12)
        Me.Label9.TabIndex = 1
        Me.Label9.Text = "CAD中点的坐标(X)"
        Me.ToolTip1.SetToolTip(Me.Label9, "CAD平面图中用于坐标变换的两个定位点的坐标，这两个点分别代表ABCD基坑群的左下角与右上角。")
        '
        'Label10
        '
        Me.Label10.AutoSize = true
        Me.Label10.Location = New System.Drawing.Point(3, 60)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(101, 12)
        Me.Label10.TabIndex = 2
        Me.Label10.Text = "CAD中点的坐标(Y)"
        Me.ToolTip1.SetToolTip(Me.Label10, "CAD平面图中用于坐标变换的两个定位点的坐标，这两个点分别代表ABCD基坑群的左下角与右上角。")
        '
        'TextBoxPageName
        '
        Me.TextBoxPageName.Location = New System.Drawing.Point(106, 45)
        Me.TextBoxPageName.Name = "TextBoxPageName"
        Me.TextBoxPageName.Size = New System.Drawing.Size(72, 21)
        Me.TextBoxPageName.TabIndex = 1
        Me.TextBoxPageName.Text = "开挖平面"
        '
        'BtnGenerate
        '
        Me.BtnGenerate.Location = New System.Drawing.Point(293, 320)
        Me.BtnGenerate.Name = "BtnGenerate"
        Me.BtnGenerate.Size = New System.Drawing.Size(75, 23)
        Me.BtnGenerate.TabIndex = 3
        Me.BtnGenerate.Text = "确定"
        Me.BtnGenerate.UseVisualStyleBackColor = true
        '
        'TextBoxFilePath
        '
        Me.TextBoxFilePath.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left)  _
            Or System.Windows.Forms.AnchorStyles.Right),System.Windows.Forms.AnchorStyles)
        Me.TextBoxFilePath.Location = New System.Drawing.Point(88, 13)
        Me.TextBoxFilePath.MinimumSize = New System.Drawing.Size(210, 21)
        Me.TextBoxFilePath.Name = "TextBoxFilePath"
        Me.TextBoxFilePath.Size = New System.Drawing.Size(274, 21)
        Me.TextBoxFilePath.TabIndex = 5
        Me.TextBoxFilePath.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'btnChooseVisioPlanView
        '
        Me.btnChooseVisioPlanView.Location = New System.Drawing.Point(6, 12)
        Me.btnChooseVisioPlanView.Name = "btnChooseVisioPlanView"
        Me.btnChooseVisioPlanView.Size = New System.Drawing.Size(75, 23)
        Me.btnChooseVisioPlanView.TabIndex = 4
        Me.btnChooseVisioPlanView.Text = "Visio文档"
        Me.btnChooseVisioPlanView.UseVisualStyleBackColor = true
        '
        'ChkBx_PointInfo
        '
        Me.ChkBx_PointInfo.AutoSize = true
        Me.ChkBx_PointInfo.Location = New System.Drawing.Point(6, 158)
        Me.ChkBx_PointInfo.Name = "ChkBx_PointInfo"
        Me.ChkBx_PointInfo.Size = New System.Drawing.Size(72, 16)
        Me.ChkBx_PointInfo.TabIndex = 7
        Me.ChkBx_PointInfo.Text = "测点信息"
        Me.ChkBx_PointInfo.UseVisualStyleBackColor = true
        '
        'txtbx_ShapeName_MonitorPointTag
        '
        Me.txtbx_ShapeName_MonitorPointTag.Location = New System.Drawing.Point(220, 6)
        Me.txtbx_ShapeName_MonitorPointTag.Name = "txtbx_ShapeName_MonitorPointTag"
        Me.txtbx_ShapeName_MonitorPointTag.Size = New System.Drawing.Size(75, 21)
        Me.txtbx_ShapeName_MonitorPointTag.TabIndex = 1
        Me.txtbx_ShapeName_MonitorPointTag.Text = "Tag"
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.TableLayoutPanel1)
        Me.Panel1.Controls.Add(Me.Label7)
        Me.Panel1.Controls.Add(Me.txtbx_ShapeName_MonitorPointTag)
        Me.Panel1.Location = New System.Drawing.Point(6, 180)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(362, 120)
        Me.Panel1.TabIndex = 8
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 3
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 33.33333!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 33.33334!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 33.33334!))
        Me.TableLayoutPanel1.Controls.Add(Me.txtbx_Pt_UR_CAD_Y, 2, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.txtbx_Pt_BL_CAD_Y, 1, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.txtbx_Pt_BL_CAD_X, 1, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.txtbx_Pt_UR_ShapeID, 2, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.txtbx_Pt_UR_CAD_X, 2, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.Label9, 0, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.Label10, 0, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.Label12, 2, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.txtbx_Pt_BL_ShapeID, 1, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.Label8, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.Label11, 1, 0)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 35)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 4
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 25!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 25!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 25!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 25!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(360, 83)
        Me.TableLayoutPanel1.TabIndex = 3
        '
        'txtbx_Pt_UR_CAD_Y
        '
        Me.txtbx_Pt_UR_CAD_Y.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtbx_Pt_UR_CAD_Y.Location = New System.Drawing.Point(240, 61)
        Me.txtbx_Pt_UR_CAD_Y.Margin = New System.Windows.Forms.Padding(1)
        Me.txtbx_Pt_UR_CAD_Y.Name = "txtbx_Pt_UR_CAD_Y"
        Me.txtbx_Pt_UR_CAD_Y.Size = New System.Drawing.Size(119, 21)
        Me.txtbx_Pt_UR_CAD_Y.TabIndex = 9
        Me.txtbx_Pt_UR_CAD_Y.Text = "201852.14"
        '
        'txtbx_Pt_BL_CAD_Y
        '
        Me.txtbx_Pt_BL_CAD_Y.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtbx_Pt_BL_CAD_Y.Location = New System.Drawing.Point(120, 61)
        Me.txtbx_Pt_BL_CAD_Y.Margin = New System.Windows.Forms.Padding(1)
        Me.txtbx_Pt_BL_CAD_Y.Name = "txtbx_Pt_BL_CAD_Y"
        Me.txtbx_Pt_BL_CAD_Y.Size = New System.Drawing.Size(118, 21)
        Me.txtbx_Pt_BL_CAD_Y.TabIndex = 9
        Me.txtbx_Pt_BL_CAD_Y.Text = "-119668.436"
        '
        'txtbx_Pt_BL_CAD_X
        '
        Me.txtbx_Pt_BL_CAD_X.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtbx_Pt_BL_CAD_X.Location = New System.Drawing.Point(120, 41)
        Me.txtbx_Pt_BL_CAD_X.Margin = New System.Windows.Forms.Padding(1)
        Me.txtbx_Pt_BL_CAD_X.Name = "txtbx_Pt_BL_CAD_X"
        Me.txtbx_Pt_BL_CAD_X.Size = New System.Drawing.Size(118, 21)
        Me.txtbx_Pt_BL_CAD_X.TabIndex = 9
        Me.txtbx_Pt_BL_CAD_X.Text = "309598.527"
        '
        'txtbx_Pt_UR_ShapeID
        '
        Me.txtbx_Pt_UR_ShapeID.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtbx_Pt_UR_ShapeID.Location = New System.Drawing.Point(240, 21)
        Me.txtbx_Pt_UR_ShapeID.Margin = New System.Windows.Forms.Padding(1)
        Me.txtbx_Pt_UR_ShapeID.Name = "txtbx_Pt_UR_ShapeID"
        Me.txtbx_Pt_UR_ShapeID.Size = New System.Drawing.Size(119, 21)
        Me.txtbx_Pt_UR_ShapeID.TabIndex = 9
        Me.txtbx_Pt_UR_ShapeID.Text = "217"
        '
        'txtbx_Pt_UR_CAD_X
        '
        Me.txtbx_Pt_UR_CAD_X.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtbx_Pt_UR_CAD_X.Location = New System.Drawing.Point(240, 41)
        Me.txtbx_Pt_UR_CAD_X.Margin = New System.Windows.Forms.Padding(1)
        Me.txtbx_Pt_UR_CAD_X.Name = "txtbx_Pt_UR_CAD_X"
        Me.txtbx_Pt_UR_CAD_X.Size = New System.Drawing.Size(119, 21)
        Me.txtbx_Pt_UR_CAD_X.TabIndex = 9
        Me.txtbx_Pt_UR_CAD_X.Text = "536642.644"
        '
        'Label12
        '
        Me.Label12.AutoSize = true
        Me.Label12.Location = New System.Drawing.Point(242, 0)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(53, 12)
        Me.Label12.TabIndex = 4
        Me.Label12.Text = "右上角点"
        '
        'txtbx_Pt_BL_ShapeID
        '
        Me.txtbx_Pt_BL_ShapeID.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtbx_Pt_BL_ShapeID.Location = New System.Drawing.Point(120, 21)
        Me.txtbx_Pt_BL_ShapeID.Margin = New System.Windows.Forms.Padding(1)
        Me.txtbx_Pt_BL_ShapeID.Name = "txtbx_Pt_BL_ShapeID"
        Me.txtbx_Pt_BL_ShapeID.Size = New System.Drawing.Size(118, 21)
        Me.txtbx_Pt_BL_ShapeID.TabIndex = 5
        Me.txtbx_Pt_BL_ShapeID.Text = "197"
        '
        'Label11
        '
        Me.Label11.AutoSize = true
        Me.Label11.Location = New System.Drawing.Point(122, 0)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(53, 12)
        Me.Label11.TabIndex = 3
        Me.Label11.Text = "左下角点"
        '
        'Btn_Import
        '
        Me.Btn_Import.Location = New System.Drawing.Point(6, 320)
        Me.Btn_Import.Name = "Btn_Import"
        Me.Btn_Import.Size = New System.Drawing.Size(63, 23)
        Me.Btn_Import.TabIndex = 9
        Me.Btn_Import.Text = "导入"
        Me.Btn_Import.UseVisualStyleBackColor = true
        '
        'Btn_Export
        '
        Me.Btn_Export.Location = New System.Drawing.Point(75, 320)
        Me.Btn_Export.Name = "Btn_Export"
        Me.Btn_Export.Size = New System.Drawing.Size(63, 23)
        Me.Btn_Export.TabIndex = 9
        Me.Btn_Export.Text = "导出"
        Me.Btn_Export.UseVisualStyleBackColor = true
        '
        'btnCancel
        '
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(212, 320)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 10
        Me.btnCancel.Text = "取消"
        Me.btnCancel.UseVisualStyleBackColor = true
        '
        'frmDrawingPlan
        '
        Me.AcceptButton = Me.BtnGenerate
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6!, 12!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(377, 355)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.Btn_Export)
        Me.Controls.Add(Me.Btn_Import)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.ChkBx_PointInfo)
        Me.Controls.Add(Me.TextBoxFilePath)
        Me.Controls.Add(Me.btnChooseVisioPlanView)
        Me.Controls.Add(Me.BtnGenerate)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.TextBoxPageName)
        Me.Controls.Add(Me.Label4)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = false
        Me.Name = "frmDrawingPlan"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "开挖平面图"
        Me.GroupBox1.ResumeLayout(false)
        Me.GroupBox1.PerformLayout
        Me.Panel1.ResumeLayout(false)
        Me.Panel1.PerformLayout
        Me.TableLayoutPanel1.ResumeLayout(false)
        Me.TableLayoutPanel1.PerformLayout
        Me.ResumeLayout(false)
        Me.PerformLayout

End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBoxInfoBoxID As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents TextBoxAllRegions As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TextBoxPageName As System.Windows.Forms.TextBox
    Friend WithEvents BtnGenerate As System.Windows.Forms.Button
    Friend WithEvents TextBoxFilePath As System.Windows.Forms.TextBox
    Friend WithEvents btnChooseVisioPlanView As System.Windows.Forms.Button
    Friend WithEvents ChkBx_PointInfo As System.Windows.Forms.CheckBox
    Friend WithEvents txtbx_ShapeName_MonitorPointTag As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents txtbx_Pt_BL_ShapeID As System.Windows.Forms.TextBox
    Friend WithEvents txtbx_Pt_UR_CAD_Y As System.Windows.Forms.TextBox
    Friend WithEvents txtbx_Pt_BL_CAD_Y As System.Windows.Forms.TextBox
    Friend WithEvents txtbx_Pt_BL_CAD_X As System.Windows.Forms.TextBox
    Friend WithEvents txtbx_Pt_UR_ShapeID As System.Windows.Forms.TextBox
    Friend WithEvents txtbx_Pt_UR_CAD_X As System.Windows.Forms.TextBox
    Friend WithEvents Btn_Import As System.Windows.Forms.Button
    Friend WithEvents Btn_Export As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
End Class
