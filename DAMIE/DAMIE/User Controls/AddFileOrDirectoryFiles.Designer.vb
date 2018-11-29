Namespace AME_UserControl


    <Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
    Partial Class AddFileOrDirectoryFiles
        Inherits System.Windows.Forms.UserControl

        'UserControl overrides dispose to clean up the component list.
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
            Me.PanelAddFileOrDir = New System.Windows.Forms.Panel()
            Me.lbAddFile = New System.Windows.Forms.Label()
            Me.lbAddDir = New System.Windows.Forms.Label()
            Me.btnAdd = New System.Windows.Forms.Button()
            Me.PanelAddFileOrDir.SuspendLayout()
            Me.SuspendLayout()
            '
            'PanelAddFileOrDir
            '
            Me.PanelAddFileOrDir.Controls.Add(Me.lbAddFile)
            Me.PanelAddFileOrDir.Controls.Add(Me.lbAddDir)
            Me.PanelAddFileOrDir.Location = New System.Drawing.Point(2, 24)
            Me.PanelAddFileOrDir.Name = "PanelAddFileOrDir"
            Me.PanelAddFileOrDir.Size = New System.Drawing.Size(96, 41)
            Me.PanelAddFileOrDir.TabIndex = 7
            Me.PanelAddFileOrDir.Visible = False
            '
            'lbAddFile
            '
            Me.lbAddFile.BackColor = System.Drawing.Color.White
            Me.lbAddFile.Font = New System.Drawing.Font("SimSun", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
            Me.lbAddFile.ForeColor = System.Drawing.SystemColors.InfoText
            Me.lbAddFile.Location = New System.Drawing.Point(0, 0)
            Me.lbAddFile.Name = "lbAddFile"
            Me.lbAddFile.Size = New System.Drawing.Size(98, 21)
            Me.lbAddFile.TabIndex = 0
            Me.lbAddFile.Text = "添加文件"
            Me.lbAddFile.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            '
            'lbAddDir
            '
            Me.lbAddDir.BackColor = System.Drawing.Color.White
            Me.lbAddDir.Font = New System.Drawing.Font("SimSun", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
            Me.lbAddDir.ForeColor = System.Drawing.SystemColors.InfoText
            Me.lbAddDir.Location = New System.Drawing.Point(0, 20)
            Me.lbAddDir.Name = "lbAddDir"
            Me.lbAddDir.Size = New System.Drawing.Size(98, 21)
            Me.lbAddDir.TabIndex = 0
            Me.lbAddDir.Text = "添加文件夹"
            Me.lbAddDir.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            '
            'btnAdd
            '
            Me.btnAdd.Location = New System.Drawing.Point(1, 1)
            Me.btnAdd.Name = "btnAdd"
            Me.btnAdd.Size = New System.Drawing.Size(98, 24)
            Me.btnAdd.TabIndex = 6
            Me.btnAdd.Text = "添加"
            Me.btnAdd.UseVisualStyleBackColor = True
            '
            'AddFileOrDirectoryFiles
            '
            Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.BackColor = System.Drawing.Color.Transparent
            Me.Controls.Add(Me.PanelAddFileOrDir)
            Me.Controls.Add(Me.btnAdd)
            Me.Margin = New System.Windows.Forms.Padding(0)
            Me.Name = "AddFileOrDirectoryFiles"
            Me.Size = New System.Drawing.Size(100, 68)
            Me.PanelAddFileOrDir.ResumeLayout(False)
            Me.ResumeLayout(False)

        End Sub
        Friend WithEvents PanelAddFileOrDir As System.Windows.Forms.Panel
        Friend WithEvents lbAddFile As System.Windows.Forms.Label
        Friend WithEvents lbAddDir As System.Windows.Forms.Label
        Friend WithEvents btnAdd As System.Windows.Forms.Button

    End Class
End Namespace