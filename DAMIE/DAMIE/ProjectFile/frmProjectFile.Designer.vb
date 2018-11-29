Namespace DataBase

    <Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
    Partial Class frmProjectFile
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
            Me.LstbxSheetsProgressInWkbk = New System.Windows.Forms.ListBox()
            Me.LstbxSheetsProgressInProject = New System.Windows.Forms.ListBox()
            Me.Label1 = New System.Windows.Forms.Label()
            Me.Label2 = New System.Windows.Forms.Label()
            Me.BtnAddSheet = New System.Windows.Forms.Button()
            Me.BtnRemoveSheet = New System.Windows.Forms.Button()
            Me.Label3 = New System.Windows.Forms.Label()
            Me.Label5 = New System.Windows.Forms.Label()
            Me.CmbbxSectional = New System.Windows.Forms.ComboBox()
            Me.Label4 = New System.Windows.Forms.Label()
            Me.LineShape2 = New Microsoft.VisualBasic.PowerPacks.LineShape()
            Me.LineShape1 = New Microsoft.VisualBasic.PowerPacks.LineShape()
            Me.PanelGeneral = New System.Windows.Forms.Panel()
            Me.CmbbxWorkingStage = New System.Windows.Forms.ComboBox()
            Me.Label9 = New System.Windows.Forms.Label()
            Me.CmbbxPointCoordinates = New System.Windows.Forms.ComboBox()
            Me.CmbbxPlan = New System.Windows.Forms.ComboBox()
            Me.ShapeContainer2 = New Microsoft.VisualBasic.PowerPacks.ShapeContainer()
            Me.LineShape3 = New Microsoft.VisualBasic.PowerPacks.LineShape()
            Me.PanelFather = New System.Windows.Forms.Panel()
            Me.GroupBox1 = New System.Windows.Forms.GroupBox()
            Me.CmbbxProgressWkbk = New System.Windows.Forms.ComboBox()
            Me.Label6 = New System.Windows.Forms.Label()
            Me.btnAddWorkbook = New System.Windows.Forms.Button()
            Me.LstBxWorkbooks = New System.Windows.Forms.ListBox()
            Me.Label7 = New System.Windows.Forms.Label()
            Me.BtnRemoveWorkbook = New System.Windows.Forms.Button()
            Me.btnOk = New System.Windows.Forms.Button()
            Me.btnCancel = New System.Windows.Forms.Button()
            Me.Label8 = New System.Windows.Forms.Label()
            Me.Panel1 = New System.Windows.Forms.Panel()
            Me.LabelProjectFilePath = New System.Windows.Forms.Label()
            Me.PanelGeneral.SuspendLayout()
            Me.PanelFather.SuspendLayout()
            Me.GroupBox1.SuspendLayout()
            Me.Panel1.SuspendLayout()
            Me.SuspendLayout()
            '
            'LstbxSheetsProgressInWkbk
            '
            Me.LstbxSheetsProgressInWkbk.FormattingEnabled = True
            Me.LstbxSheetsProgressInWkbk.HorizontalScrollbar = True
            Me.LstbxSheetsProgressInWkbk.ItemHeight = 12
            Me.LstbxSheetsProgressInWkbk.Location = New System.Drawing.Point(12, 80)
            Me.LstbxSheetsProgressInWkbk.Name = "LstbxSheetsProgressInWkbk"
            Me.LstbxSheetsProgressInWkbk.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
            Me.LstbxSheetsProgressInWkbk.Size = New System.Drawing.Size(129, 172)
            Me.LstbxSheetsProgressInWkbk.TabIndex = 0
            '
            'LstbxSheetsProgressInProject
            '
            Me.LstbxSheetsProgressInProject.FormattingEnabled = True
            Me.LstbxSheetsProgressInProject.HorizontalScrollbar = True
            Me.LstbxSheetsProgressInProject.ItemHeight = 12
            Me.LstbxSheetsProgressInProject.Location = New System.Drawing.Point(149, 80)
            Me.LstbxSheetsProgressInProject.Name = "LstbxSheetsProgressInProject"
            Me.LstbxSheetsProgressInProject.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
            Me.LstbxSheetsProgressInProject.Size = New System.Drawing.Size(129, 172)
            Me.LstbxSheetsProgressInProject.TabIndex = 0
            '
            'Label1
            '
            Me.Label1.AutoSize = True
            Me.Label1.Location = New System.Drawing.Point(10, 58)
            Me.Label1.Name = "Label1"
            Me.Label1.Size = New System.Drawing.Size(89, 12)
            Me.Label1.TabIndex = 3
            Me.Label1.Text = "文件中的工作表"
            '
            'Label2
            '
            Me.Label2.AutoSize = True
            Me.Label2.Location = New System.Drawing.Point(147, 58)
            Me.Label2.Name = "Label2"
            Me.Label2.Size = New System.Drawing.Size(89, 12)
            Me.Label2.TabIndex = 3
            Me.Label2.Text = "项目中的工作表"
            '
            'BtnAddSheet
            '
            Me.BtnAddSheet.Location = New System.Drawing.Point(284, 132)
            Me.BtnAddSheet.Name = "BtnAddSheet"
            Me.BtnAddSheet.Size = New System.Drawing.Size(61, 23)
            Me.BtnAddSheet.TabIndex = 4
            Me.BtnAddSheet.Text = "Add"
            Me.BtnAddSheet.UseVisualStyleBackColor = True
            '
            'BtnRemoveSheet
            '
            Me.BtnRemoveSheet.Location = New System.Drawing.Point(284, 176)
            Me.BtnRemoveSheet.Name = "BtnRemoveSheet"
            Me.BtnRemoveSheet.Size = New System.Drawing.Size(61, 23)
            Me.BtnRemoveSheet.TabIndex = 4
            Me.BtnRemoveSheet.Text = "Remove"
            Me.BtnRemoveSheet.UseVisualStyleBackColor = True
            '
            'Label3
            '
            Me.Label3.AutoSize = True
            Me.Label3.Location = New System.Drawing.Point(3, 11)
            Me.Label3.Name = "Label3"
            Me.Label3.Size = New System.Drawing.Size(53, 12)
            Me.Label3.TabIndex = 1
            Me.Label3.Text = "剖面标高"
            '
            'Label5
            '
            Me.Label5.AutoSize = True
            Me.Label5.Location = New System.Drawing.Point(3, 122)
            Me.Label5.Name = "Label5"
            Me.Label5.Size = New System.Drawing.Size(53, 12)
            Me.Label5.TabIndex = 1
            Me.Label5.Text = "测点坐标"
            '
            'CmbbxSectional
            '
            Me.CmbbxSectional.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
            Me.CmbbxSectional.FormattingEnabled = True
            Me.CmbbxSectional.Location = New System.Drawing.Point(62, 8)
            Me.CmbbxSectional.Name = "CmbbxSectional"
            Me.CmbbxSectional.Size = New System.Drawing.Size(271, 20)
            Me.CmbbxSectional.TabIndex = 2
            '
            'Label4
            '
            Me.Label4.AutoSize = True
            Me.Label4.Location = New System.Drawing.Point(3, 68)
            Me.Label4.Name = "Label4"
            Me.Label4.Size = New System.Drawing.Size(53, 12)
            Me.Label4.TabIndex = 1
            Me.Label4.Text = "开挖分块"
            '
            'LineShape2
            '
            Me.LineShape2.BorderColor = System.Drawing.SystemColors.ControlDark
            Me.LineShape2.Cursor = System.Windows.Forms.Cursors.Default
            Me.LineShape2.Enabled = False
            Me.LineShape2.Name = "LineShape1"
            Me.LineShape2.X1 = 12
            Me.LineShape2.X2 = 344
            Me.LineShape2.Y1 = 48
            Me.LineShape2.Y2 = 48
            '
            'LineShape1
            '
            Me.LineShape1.BorderColor = System.Drawing.SystemColors.ControlDark
            Me.LineShape1.Enabled = False
            Me.LineShape1.Name = "LineShape1"
            Me.LineShape1.X1 = 12
            Me.LineShape1.X2 = 344
            Me.LineShape1.Y1 = 104
            Me.LineShape1.Y2 = 104
            '
            'PanelGeneral
            '
            Me.PanelGeneral.Controls.Add(Me.CmbbxWorkingStage)
            Me.PanelGeneral.Controls.Add(Me.Label9)
            Me.PanelGeneral.Controls.Add(Me.CmbbxPointCoordinates)
            Me.PanelGeneral.Controls.Add(Me.CmbbxPlan)
            Me.PanelGeneral.Controls.Add(Me.Label3)
            Me.PanelGeneral.Controls.Add(Me.Label5)
            Me.PanelGeneral.Controls.Add(Me.Label4)
            Me.PanelGeneral.Controls.Add(Me.CmbbxSectional)
            Me.PanelGeneral.Controls.Add(Me.ShapeContainer2)
            Me.PanelGeneral.Location = New System.Drawing.Point(3, 3)
            Me.PanelGeneral.Name = "PanelGeneral"
            Me.PanelGeneral.Size = New System.Drawing.Size(358, 215)
            Me.PanelGeneral.TabIndex = 8
            '
            'CmbbxWorkingStage
            '
            Me.CmbbxWorkingStage.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
            Me.CmbbxWorkingStage.FormattingEnabled = True
            Me.CmbbxWorkingStage.Location = New System.Drawing.Point(65, 175)
            Me.CmbbxWorkingStage.Name = "CmbbxWorkingStage"
            Me.CmbbxWorkingStage.Size = New System.Drawing.Size(271, 20)
            Me.CmbbxWorkingStage.TabIndex = 2
            '
            'Label9
            '
            Me.Label9.AutoSize = True
            Me.Label9.Location = New System.Drawing.Point(3, 178)
            Me.Label9.Name = "Label9"
            Me.Label9.Size = New System.Drawing.Size(53, 12)
            Me.Label9.TabIndex = 1
            Me.Label9.Text = "开挖工况"
            '
            'CmbbxPointCoordinates
            '
            Me.CmbbxPointCoordinates.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
            Me.CmbbxPointCoordinates.FormattingEnabled = True
            Me.CmbbxPointCoordinates.Location = New System.Drawing.Point(65, 119)
            Me.CmbbxPointCoordinates.Name = "CmbbxPointCoordinates"
            Me.CmbbxPointCoordinates.Size = New System.Drawing.Size(271, 20)
            Me.CmbbxPointCoordinates.TabIndex = 2
            '
            'CmbbxPlan
            '
            Me.CmbbxPlan.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
            Me.CmbbxPlan.FormattingEnabled = True
            Me.CmbbxPlan.Location = New System.Drawing.Point(65, 65)
            Me.CmbbxPlan.Name = "CmbbxPlan"
            Me.CmbbxPlan.Size = New System.Drawing.Size(271, 20)
            Me.CmbbxPlan.TabIndex = 2
            '
            'ShapeContainer2
            '
            Me.ShapeContainer2.Location = New System.Drawing.Point(0, 0)
            Me.ShapeContainer2.Margin = New System.Windows.Forms.Padding(0)
            Me.ShapeContainer2.Name = "ShapeContainer2"
            Me.ShapeContainer2.Shapes.AddRange(New Microsoft.VisualBasic.PowerPacks.Shape() {Me.LineShape3, Me.LineShape2, Me.LineShape1})
            Me.ShapeContainer2.Size = New System.Drawing.Size(358, 215)
            Me.ShapeContainer2.TabIndex = 4
            Me.ShapeContainer2.TabStop = False
            '
            'LineShape3
            '
            Me.LineShape3.BorderColor = System.Drawing.SystemColors.ControlDark
            Me.LineShape3.Cursor = System.Windows.Forms.Cursors.Default
            Me.LineShape3.Enabled = False
            Me.LineShape3.Name = "LineShape1"
            Me.LineShape3.X1 = 12
            Me.LineShape3.X2 = 344
            Me.LineShape3.Y1 = 160
            Me.LineShape3.Y2 = 160
            '
            'PanelFather
            '
            Me.PanelFather.AutoScroll = True
            Me.PanelFather.AutoScrollMargin = New System.Drawing.Size(0, 10)
            Me.PanelFather.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
            Me.PanelFather.Controls.Add(Me.GroupBox1)
            Me.PanelFather.Controls.Add(Me.PanelGeneral)
            Me.PanelFather.Location = New System.Drawing.Point(12, 146)
            Me.PanelFather.Name = "PanelFather"
            Me.PanelFather.Size = New System.Drawing.Size(389, 260)
            Me.PanelFather.TabIndex = 10
            '
            'GroupBox1
            '
            Me.GroupBox1.Controls.Add(Me.CmbbxProgressWkbk)
            Me.GroupBox1.Controls.Add(Me.Label6)
            Me.GroupBox1.Controls.Add(Me.Label1)
            Me.GroupBox1.Controls.Add(Me.Label2)
            Me.GroupBox1.Controls.Add(Me.LstbxSheetsProgressInWkbk)
            Me.GroupBox1.Controls.Add(Me.BtnRemoveSheet)
            Me.GroupBox1.Controls.Add(Me.BtnAddSheet)
            Me.GroupBox1.Controls.Add(Me.LstbxSheetsProgressInProject)
            Me.GroupBox1.Location = New System.Drawing.Point(3, 224)
            Me.GroupBox1.Name = "GroupBox1"
            Me.GroupBox1.Size = New System.Drawing.Size(358, 260)
            Me.GroupBox1.TabIndex = 9
            Me.GroupBox1.TabStop = False
            Me.GroupBox1.Text = "施工进度"
            '
            'CmbbxProgressWkbk
            '
            Me.CmbbxProgressWkbk.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
            Me.CmbbxProgressWkbk.FormattingEnabled = True
            Me.CmbbxProgressWkbk.Location = New System.Drawing.Point(107, 26)
            Me.CmbbxProgressWkbk.Name = "CmbbxProgressWkbk"
            Me.CmbbxProgressWkbk.Size = New System.Drawing.Size(229, 20)
            Me.CmbbxProgressWkbk.TabIndex = 6
            '
            'Label6
            '
            Me.Label6.AutoSize = True
            Me.Label6.Location = New System.Drawing.Point(12, 30)
            Me.Label6.Name = "Label6"
            Me.Label6.Size = New System.Drawing.Size(89, 12)
            Me.Label6.TabIndex = 5
            Me.Label6.Text = "选择工作簿文件"
            '
            'btnAddWorkbook
            '
            Me.btnAddWorkbook.Location = New System.Drawing.Point(305, 27)
            Me.btnAddWorkbook.Name = "btnAddWorkbook"
            Me.btnAddWorkbook.Size = New System.Drawing.Size(75, 23)
            Me.btnAddWorkbook.TabIndex = 11
            Me.btnAddWorkbook.Text = "Add"
            Me.btnAddWorkbook.UseVisualStyleBackColor = True
            '
            'LstBxWorkbooks
            '
            Me.LstBxWorkbooks.FormattingEnabled = True
            Me.LstBxWorkbooks.HorizontalScrollbar = True
            Me.LstBxWorkbooks.ItemHeight = 12
            Me.LstBxWorkbooks.Location = New System.Drawing.Point(6, 27)
            Me.LstBxWorkbooks.Name = "LstBxWorkbooks"
            Me.LstBxWorkbooks.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple
            Me.LstBxWorkbooks.Size = New System.Drawing.Size(293, 64)
            Me.LstBxWorkbooks.TabIndex = 12
            '
            'Label7
            '
            Me.Label7.AutoSize = True
            Me.Label7.Location = New System.Drawing.Point(6, 6)
            Me.Label7.Name = "Label7"
            Me.Label7.Size = New System.Drawing.Size(89, 12)
            Me.Label7.TabIndex = 13
            Me.Label7.Text = "数据工作簿列表"
            '
            'BtnRemoveWorkbook
            '
            Me.BtnRemoveWorkbook.Location = New System.Drawing.Point(305, 68)
            Me.BtnRemoveWorkbook.Name = "BtnRemoveWorkbook"
            Me.BtnRemoveWorkbook.Size = New System.Drawing.Size(75, 23)
            Me.BtnRemoveWorkbook.TabIndex = 11
            Me.BtnRemoveWorkbook.Text = "Remove"
            Me.BtnRemoveWorkbook.UseVisualStyleBackColor = True
            '
            'btnOk
            '
            Me.btnOk.Location = New System.Drawing.Point(326, 421)
            Me.btnOk.Name = "btnOk"
            Me.btnOk.Size = New System.Drawing.Size(75, 23)
            Me.btnOk.TabIndex = 14
            Me.btnOk.Text = "确定"
            Me.btnOk.UseVisualStyleBackColor = True
            '
            'btnCancel
            '
            Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
            Me.btnCancel.Location = New System.Drawing.Point(243, 421)
            Me.btnCancel.Name = "btnCancel"
            Me.btnCancel.Size = New System.Drawing.Size(75, 23)
            Me.btnCancel.TabIndex = 14
            Me.btnCancel.Text = "取消(&C)"
            Me.btnCancel.UseVisualStyleBackColor = True
            '
            'Label8
            '
            Me.Label8.AutoSize = True
            Me.Label8.Location = New System.Drawing.Point(14, 15)
            Me.Label8.Name = "Label8"
            Me.Label8.Size = New System.Drawing.Size(71, 12)
            Me.Label8.TabIndex = 16
            Me.Label8.Text = "项目文件 : "
            '
            'Panel1
            '
            Me.Panel1.Controls.Add(Me.Label7)
            Me.Panel1.Controls.Add(Me.btnAddWorkbook)
            Me.Panel1.Controls.Add(Me.BtnRemoveWorkbook)
            Me.Panel1.Controls.Add(Me.LstBxWorkbooks)
            Me.Panel1.Location = New System.Drawing.Point(12, 43)
            Me.Panel1.Name = "Panel1"
            Me.Panel1.Size = New System.Drawing.Size(389, 94)
            Me.Panel1.TabIndex = 17
            '
            'LabelProjectFilePath
            '
            Me.LabelProjectFilePath.AutoSize = True
            Me.LabelProjectFilePath.Location = New System.Drawing.Point(79, 15)
            Me.LabelProjectFilePath.Name = "LabelProjectFilePath"
            Me.LabelProjectFilePath.Size = New System.Drawing.Size(53, 12)
            Me.LabelProjectFilePath.TabIndex = 18
            Me.LabelProjectFilePath.Text = "FilePath"
            '
            'frmProjectFile
            '
            Me.AcceptButton = Me.btnOk
            Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.CancelButton = Me.btnCancel
            Me.ClientSize = New System.Drawing.Size(419, 456)
            Me.Controls.Add(Me.LabelProjectFilePath)
            Me.Controls.Add(Me.Label8)
            Me.Controls.Add(Me.Panel1)
            Me.Controls.Add(Me.btnCancel)
            Me.Controls.Add(Me.btnOk)
            Me.Controls.Add(Me.PanelFather)
            Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
            Me.HelpButton = True
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.Name = "frmProjectFile"
            Me.Text = "New Project"
            Me.PanelGeneral.ResumeLayout(False)
            Me.PanelGeneral.PerformLayout()
            Me.PanelFather.ResumeLayout(False)
            Me.GroupBox1.ResumeLayout(False)
            Me.GroupBox1.PerformLayout()
            Me.Panel1.ResumeLayout(False)
            Me.Panel1.PerformLayout()
            Me.ResumeLayout(False)
            Me.PerformLayout()

        End Sub
        Friend WithEvents LstbxSheetsProgressInWkbk As System.Windows.Forms.ListBox
        Friend WithEvents LstbxSheetsProgressInProject As System.Windows.Forms.ListBox
        Friend WithEvents Label1 As System.Windows.Forms.Label
        Friend WithEvents Label2 As System.Windows.Forms.Label
        Friend WithEvents BtnAddSheet As System.Windows.Forms.Button
        Friend WithEvents BtnRemoveSheet As System.Windows.Forms.Button
        Friend WithEvents CmbbxSectional As System.Windows.Forms.ComboBox
        Friend WithEvents Label5 As System.Windows.Forms.Label
        Friend WithEvents Label4 As System.Windows.Forms.Label
        Friend WithEvents Label3 As System.Windows.Forms.Label
        Friend WithEvents LineShape2 As Microsoft.VisualBasic.PowerPacks.LineShape
        Friend WithEvents LineShape1 As Microsoft.VisualBasic.PowerPacks.LineShape
        Friend WithEvents PanelGeneral As System.Windows.Forms.Panel
        Friend WithEvents CmbbxPointCoordinates As System.Windows.Forms.ComboBox
        Friend WithEvents CmbbxPlan As System.Windows.Forms.ComboBox
        Friend WithEvents ShapeContainer2 As Microsoft.VisualBasic.PowerPacks.ShapeContainer
        Friend WithEvents PanelFather As System.Windows.Forms.Panel
        Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
        Friend WithEvents CmbbxProgressWkbk As System.Windows.Forms.ComboBox
        Friend WithEvents Label6 As System.Windows.Forms.Label
        Friend WithEvents btnAddWorkbook As System.Windows.Forms.Button
        Friend WithEvents LstBxWorkbooks As System.Windows.Forms.ListBox
        Friend WithEvents Label7 As System.Windows.Forms.Label
        Friend WithEvents BtnRemoveWorkbook As System.Windows.Forms.Button
        Friend WithEvents btnOk As System.Windows.Forms.Button
        Friend WithEvents btnCancel As System.Windows.Forms.Button
        Friend WithEvents Label8 As System.Windows.Forms.Label
        Friend WithEvents Panel1 As System.Windows.Forms.Panel
        Friend WithEvents LabelProjectFilePath As System.Windows.Forms.Label
        Friend WithEvents CmbbxWorkingStage As System.Windows.Forms.ComboBox
        Friend WithEvents Label9 As System.Windows.Forms.Label
        Friend WithEvents LineShape3 As Microsoft.VisualBasic.PowerPacks.LineShape
    End Class
End Namespace