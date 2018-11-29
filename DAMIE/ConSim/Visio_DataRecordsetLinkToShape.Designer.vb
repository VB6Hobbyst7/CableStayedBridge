<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class DataRecordsetLinkToShape
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DataRecordsetLinkToShape))
        Me.BtnChooseVsoDoc = New System.Windows.Forms.Button()
        Me.txtbxVsoDoc = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.btnLink = New System.Windows.Forms.Button()
        Me.ComboBox_Page = New System.Windows.Forms.ComboBox()
        Me.ComboBox_DataRs = New System.Windows.Forms.ComboBox()
        Me.btnValidate = New System.Windows.Forms.Button()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ComboBox_Column_ShapeID = New System.Windows.Forms.ComboBox()
        Me.SuspendLayout()
        '
        'BtnChooseVsoDoc
        '
        Me.BtnChooseVsoDoc.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnChooseVsoDoc.BackColor = System.Drawing.SystemColors.Control
        Me.BtnChooseVsoDoc.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnChooseVsoDoc.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BtnChooseVsoDoc.ForeColor = System.Drawing.SystemColors.InfoText
        Me.BtnChooseVsoDoc.Location = New System.Drawing.Point(319, 9)
        Me.BtnChooseVsoDoc.Name = "BtnChooseVsoDoc"
        Me.BtnChooseVsoDoc.Size = New System.Drawing.Size(74, 23)
        Me.BtnChooseVsoDoc.TabIndex = 5
        Me.BtnChooseVsoDoc.Text = "选择..."
        Me.BtnChooseVsoDoc.UseVisualStyleBackColor = False
        '
        'txtbxVsoDoc
        '
        Me.txtbxVsoDoc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtbxVsoDoc.BackColor = System.Drawing.Color.White
        Me.txtbxVsoDoc.Location = New System.Drawing.Point(86, 9)
        Me.txtbxVsoDoc.Margin = New System.Windows.Forms.Padding(0)
        Me.txtbxVsoDoc.Name = "txtbxVsoDoc"
        Me.txtbxVsoDoc.Size = New System.Drawing.Size(222, 21)
        Me.txtbxVsoDoc.TabIndex = 4
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(11, 12)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(59, 12)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Visio绘图"
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(5, 51)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(65, 12)
        Me.Label5.TabIndex = 13
        Me.Label5.Text = "数据记录集"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(213, 51)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(53, 12)
        Me.Label6.TabIndex = 13
        Me.Label6.Text = "绘图页面"
        '
        'btnLink
        '
        Me.btnLink.Location = New System.Drawing.Point(318, 91)
        Me.btnLink.Name = "btnLink"
        Me.btnLink.Size = New System.Drawing.Size(75, 23)
        Me.btnLink.TabIndex = 0
        Me.btnLink.Text = "链接"
        Me.ToolTip1.SetToolTip(Me.btnLink, "以数据记录集中的形状ID为指标，将其链接到页面的对应ID的形状上。")
        Me.btnLink.UseVisualStyleBackColor = True
        '
        'ComboBox_Page
        '
        Me.ComboBox_Page.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_Page.FormattingEnabled = True
        Me.ComboBox_Page.Location = New System.Drawing.Point(272, 48)
        Me.ComboBox_Page.Name = "ComboBox_Page"
        Me.ComboBox_Page.Size = New System.Drawing.Size(121, 20)
        Me.ComboBox_Page.TabIndex = 12
        '
        'ComboBox_DataRs
        '
        Me.ComboBox_DataRs.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_DataRs.FormattingEnabled = True
        Me.ComboBox_DataRs.Location = New System.Drawing.Point(75, 48)
        Me.ComboBox_DataRs.Name = "ComboBox_DataRs"
        Me.ComboBox_DataRs.Size = New System.Drawing.Size(121, 20)
        Me.ComboBox_DataRs.TabIndex = 12
        '
        'btnValidate
        '
        Me.btnValidate.Location = New System.Drawing.Point(232, 91)
        Me.btnValidate.Name = "btnValidate"
        Me.btnValidate.Size = New System.Drawing.Size(75, 23)
        Me.btnValidate.TabIndex = 14
        Me.btnValidate.Text = "验证"
        Me.ToolTip1.SetToolTip(Me.btnValidate, "验证页面的形状中是否有对应的形状ID。")
        Me.btnValidate.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(5, 91)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(65, 29)
        Me.Label1.TabIndex = 15
        Me.Label1.Text = "形状ID" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "所在的字段"
        '
        'ComboBox_Column_ShapeID
        '
        Me.ComboBox_Column_ShapeID.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_Column_ShapeID.FormattingEnabled = True
        Me.ComboBox_Column_ShapeID.Location = New System.Drawing.Point(75, 91)
        Me.ComboBox_Column_ShapeID.Name = "ComboBox_Column_ShapeID"
        Me.ComboBox_Column_ShapeID.Size = New System.Drawing.Size(121, 20)
        Me.ComboBox_Column_ShapeID.TabIndex = 12
        '
        'DataRecordsetLinkToShape
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(406, 125)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnValidate)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.btnLink)
        Me.Controls.Add(Me.BtnChooseVsoDoc)
        Me.Controls.Add(Me.ComboBox_Column_ShapeID)
        Me.Controls.Add(Me.ComboBox_Page)
        Me.Controls.Add(Me.txtbxVsoDoc)
        Me.Controls.Add(Me.ComboBox_DataRs)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "DataRecordsetLinkToShape"
        Me.Text = "数据记录集链接到形状"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents BtnChooseVsoDoc As System.Windows.Forms.Button
    Friend WithEvents txtbxVsoDoc As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents btnLink As System.Windows.Forms.Button
    Friend WithEvents ComboBox_Page As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox_DataRs As System.Windows.Forms.ComboBox
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents btnValidate As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ComboBox_Column_ShapeID As System.Windows.Forms.ComboBox
End Class
