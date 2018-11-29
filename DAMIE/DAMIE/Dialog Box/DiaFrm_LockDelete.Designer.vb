Namespace Miscellaneous

    <Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
    Partial Class DiaFrm_LockDelete
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
            Me.btn1 = New System.Windows.Forms.Button()
            Me.btn2 = New System.Windows.Forms.Button()
            Me.LabelPrompt = New System.Windows.Forms.Label()
            Me.btn3 = New System.Windows.Forms.Button()
            Me.SuspendLayout()
            '
            'btn1
            '
            Me.btn1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
            Me.btn1.Location = New System.Drawing.Point(12, 38)
            Me.btn1.Name = "btn1"
            Me.btn1.Size = New System.Drawing.Size(75, 23)
            Me.btn1.TabIndex = 0
            Me.btn1.Text = "Lock"
            Me.btn1.UseVisualStyleBackColor = True
            '
            'btn2
            '
            Me.btn2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
            Me.btn2.Location = New System.Drawing.Point(116, 38)
            Me.btn2.Name = "btn2"
            Me.btn2.Size = New System.Drawing.Size(75, 23)
            Me.btn2.TabIndex = 1
            Me.btn2.Text = "Delete"
            Me.btn2.UseVisualStyleBackColor = True
            '
            'LabelPrompt
            '
            Me.LabelPrompt.AutoEllipsis = True
            Me.LabelPrompt.AutoSize = True
            Me.LabelPrompt.Location = New System.Drawing.Point(12, 9)
            Me.LabelPrompt.Name = "LabelPrompt"
            Me.LabelPrompt.Size = New System.Drawing.Size(41, 12)
            Me.LabelPrompt.TabIndex = 3
            Me.LabelPrompt.Text = "PROMPT"
            Me.LabelPrompt.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            '
            'btn3
            '
            Me.btn3.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
            Me.btn3.DialogResult = System.Windows.Forms.DialogResult.Cancel
            Me.btn3.Location = New System.Drawing.Point(220, 38)
            Me.btn3.Name = "btn3"
            Me.btn3.Size = New System.Drawing.Size(75, 23)
            Me.btn3.TabIndex = 2
            Me.btn3.Text = "Ignore"
            Me.btn3.UseVisualStyleBackColor = True
            '
            'DiaFrm_LockDelete
            '
            Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.CancelButton = Me.btn3
            Me.ClientSize = New System.Drawing.Size(307, 70)
            Me.Controls.Add(Me.LabelPrompt)
            Me.Controls.Add(Me.btn3)
            Me.Controls.Add(Me.btn2)
            Me.Controls.Add(Me.btn1)
            Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.Name = "DiaFrm_LockDelete"
            Me.Text = "Operate on series"
            Me.TopMost = True
            Me.ResumeLayout(False)
            Me.PerformLayout()

        End Sub
        Friend WithEvents btn1 As System.Windows.Forms.Button
        Friend WithEvents btn2 As System.Windows.Forms.Button
        Friend WithEvents LabelPrompt As System.Windows.Forms.Label
        Friend WithEvents btn3 As System.Windows.Forms.Button
    End Class


End Namespace
