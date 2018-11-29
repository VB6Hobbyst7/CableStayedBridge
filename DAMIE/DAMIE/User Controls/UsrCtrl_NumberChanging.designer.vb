Namespace AME_UserControl
    <Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
    Partial Class UsrCtrl_NumberChanging
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
            Me.btnNext = New System.Windows.Forms.Button()
            Me.TextBoxNumber = New System.Windows.Forms.TextBox()
            Me.btnPrevious = New System.Windows.Forms.Button()
            Me.cbUnit = New System.Windows.Forms.ComboBox()
            Me.SuspendLayout()
            '
            'btnNext
            '
            Me.btnNext.BackColor = System.Drawing.SystemColors.ButtonFace
            Me.btnNext.Location = New System.Drawing.Point(148, -1)
            Me.btnNext.Name = "btnNext"
            Me.btnNext.Size = New System.Drawing.Size(40, 21)
            Me.btnNext.TabIndex = 2
            Me.btnNext.Text = "->"
            Me.btnNext.UseVisualStyleBackColor = False
            '
            'TextBoxNumber
            '
            Me.TextBoxNumber.BackColor = System.Drawing.Color.White
            Me.TextBoxNumber.BorderStyle = System.Windows.Forms.BorderStyle.None
            Me.TextBoxNumber.Font = New System.Drawing.Font("SimSun", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
            Me.TextBoxNumber.Location = New System.Drawing.Point(44, 2)
            Me.TextBoxNumber.Name = "TextBoxNumber"
            Me.TextBoxNumber.Size = New System.Drawing.Size(35, 16)
            Me.TextBoxNumber.TabIndex = 0
            '
            'btnPrevious
            '
            Me.btnPrevious.BackColor = System.Drawing.SystemColors.ButtonFace
            Me.btnPrevious.Location = New System.Drawing.Point(-1, -1)
            Me.btnPrevious.Name = "btnPrevious"
            Me.btnPrevious.Size = New System.Drawing.Size(40, 21)
            Me.btnPrevious.TabIndex = 1
            Me.btnPrevious.Text = "<-"
            Me.btnPrevious.UseVisualStyleBackColor = False
            '
            'cbUnit
            '
            Me.cbUnit.BackColor = System.Drawing.SystemColors.Control
            Me.cbUnit.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
            Me.cbUnit.FormattingEnabled = True
            Me.cbUnit.Location = New System.Drawing.Point(84, 0)
            Me.cbUnit.Name = "cbUnit"
            Me.cbUnit.Size = New System.Drawing.Size(60, 20)
            Me.cbUnit.TabIndex = 3
            '
            'UsrCtrl_NumberChanging
            '
            Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.BackColor = System.Drawing.Color.Transparent
            Me.Controls.Add(Me.cbUnit)
            Me.Controls.Add(Me.btnPrevious)
            Me.Controls.Add(Me.btnNext)
            Me.Controls.Add(Me.TextBoxNumber)
            Me.Name = "UsrCtrl_NumberChanging"
            Me.Size = New System.Drawing.Size(190, 21)
            Me.ResumeLayout(False)
            Me.PerformLayout()

        End Sub
        Private WithEvents btnNext As System.Windows.Forms.Button
        Private WithEvents btnPrevious As System.Windows.Forms.Button
        Private WithEvents TextBoxNumber As System.Windows.Forms.TextBox
        Private WithEvents cbUnit As System.Windows.Forms.ComboBox

    End Class
End Namespace