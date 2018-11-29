Namespace UI
    <Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
    Partial Class UI_BackGround
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
            Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(UI_BackGround))
            Me.PictureBoxAME = New System.Windows.Forms.PictureBox()
            Me.PictureBoxBackGround = New System.Windows.Forms.PictureBox()
            CType(Me.PictureBoxAME, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.PictureBoxBackGround, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.SuspendLayout()
            '
            'PictureBoxAME
            '
            Me.PictureBoxAME.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            Me.PictureBoxAME.BackColor = System.Drawing.Color.Transparent
            Me.PictureBoxAME.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
            Me.PictureBoxAME.Image = CType(resources.GetObject("PictureBoxAME.Image"), System.Drawing.Image)
            Me.PictureBoxAME.Location = New System.Drawing.Point(376, 246)
            Me.PictureBoxAME.Name = "PictureBoxAME"
            Me.PictureBoxAME.Size = New System.Drawing.Size(108, 45)
            Me.PictureBoxAME.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
            Me.PictureBoxAME.TabIndex = 9
            Me.PictureBoxAME.TabStop = False
            '
            'PictureBoxBackGround
            '
            Me.PictureBoxBackGround.BackColor = System.Drawing.Color.Transparent
            Me.PictureBoxBackGround.BackgroundImage = Global.DAMIE.My.Resources.Resources.线条背景
            Me.PictureBoxBackGround.Dock = System.Windows.Forms.DockStyle.Fill
            Me.PictureBoxBackGround.Location = New System.Drawing.Point(0, 0)
            Me.PictureBoxBackGround.Name = "PictureBoxBackGround"
            Me.PictureBoxBackGround.Size = New System.Drawing.Size(496, 303)
            Me.PictureBoxBackGround.TabIndex = 8
            Me.PictureBoxBackGround.TabStop = False
            '
            'UI_BackGround
            '
            Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.ClientSize = New System.Drawing.Size(496, 303)
            Me.ControlBox = False
            Me.Controls.Add(Me.PictureBoxAME)
            Me.Controls.Add(Me.PictureBoxBackGround)
            Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.Name = "UI_BackGround"
            Me.ShowInTaskbar = False
            Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
            CType(Me.PictureBoxAME, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.PictureBoxBackGround, System.ComponentModel.ISupportInitialize).EndInit()
            Me.ResumeLayout(False)
            Me.PerformLayout()

        End Sub
        Friend WithEvents PictureBoxAME As System.Windows.Forms.PictureBox
        Friend WithEvents PictureBoxBackGround As System.Windows.Forms.PictureBox

    End Class
End Namespace