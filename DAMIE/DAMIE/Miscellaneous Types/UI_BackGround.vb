Namespace UI
    Public NotInheritable Class UI_BackGround
        Public Sub New()

            ' This call is required by the designer.
            InitializeComponent()

            ' Add any initialization after the InitializeComponent() call.

            '设置图片的相对透明的父控件与相对位置
            With Me.PictureBoxAME
                .Parent = Me.PictureBoxBackGround
                .Top = .Top - .Parent.Top
                .Left = .Left - .Parent.Left
            End With

            '设置背景图片的背景色
            Me.BackColor = System.Drawing.Color.FromArgb(195, 195, 195)
        End Sub

    End Class

End Namespace