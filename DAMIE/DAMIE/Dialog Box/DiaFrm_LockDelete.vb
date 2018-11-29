Namespace Miscellaneous

    Public Enum M_DialogResult
        Lock
        Delete
        Cancel
    End Enum


    Public Class DiaFrm_LockDelete

#Region "  ---  Declarations & Definitions"

#Region "  ---  Types"

#End Region

#Region "  ---  Events"

#End Region

#Region "  ---  Constants"

#End Region

#Region "  ---  Properties"

        Public WriteOnly Property Prompt As String
            Set(value As String)
                Me.LabelPrompt.Text = value
            End Set
        End Property
        Public WriteOnly Property Title As String
            Set(value As String)
                Me.Text = value
            End Set
        End Property
        Public WriteOnly Property Text_Button1 As String
            Set(value As String)
                Me.btn1.Text = value
            End Set
        End Property
        Public WriteOnly Property Text_Button2 As String
            Set(value As String)
                Me.btn2.Text = value
            End Set
        End Property
        Public WriteOnly Property Text_Button3 As String
            Set(value As String)
                Me.btn3.Text = value
            End Set
        End Property

#End Region

#Region "  ---  Fields"

        ''' <summary>
        ''' 最后的返回值
        ''' </summary>
        ''' <remarks></remarks>
        Private result As M_DialogResult '= M_DialogResult.ignore

#End Region

#End Region

#Region "  ---  构造函数与窗体的加载、打开与关闭"
        Public Sub New()

            ' This call is required by the designer.
            InitializeComponent()

            ' Add any initialization after the InitializeComponent() call.
            Me.TopMost = True
            Me.StartPosition = FormStartPosition.CenterParent
        End Sub
#End Region


        Private Delegate Sub ShowDialogHandler(ByVal Prompt As String, ByVal Title As String)
        ''' <summary>
        '''  打开对话框
        ''' </summary>
        ''' <param name="Prompt">提示</param>
        ''' <param name="Title">窗口标题</param>
        ''' <param name="Text_Button1">第一个按钮的文字</param>
        ''' <param name="Text_Button2">第二个按钮的文字</param>
        ''' <param name="Text_Button3">第三个按钮的文字</param>
        ''' <returns>选择的操作方式：lock、delete、Cancel</returns>
        ''' <remarks></remarks>
        Public Shadows Function ShowDialog(ByVal Prompt As String, Optional ByVal Title As String = "TIP", _
                                           Optional Text_Button1 As String = "Lock", _
                                           Optional Text_Button2 As String = "Delete", _
                                           Optional Text_Button3 As String = "Cancel") As M_DialogResult
            If Me.InvokeRequired Then
                '非UI线程，再次封送该方法到UI线程
                Debug.Print("非UI线程，再次封送该方法到UI线程")
                Me.BeginInvoke(New ShowDialogHandler(AddressOf Me.ShowDialog), {Prompt, Title})
            Else
                With Me
                    .Title = Title
                    .Prompt = Prompt
                    .Text_Button1 = Text_Button1
                    .Text_Button2 = Text_Button2
                    .Text_Button3 = Text_Button3
                    '
                    .btn1.Focus()
                End With
                '
                MyBase.ShowDialog()
                Me.Dispose()
            End If
            Return result
        End Function

#Region "  ---  界面操作"
        '点击按钮
        Private Sub btnLock_Click(sender As Object, e As EventArgs) Handles btn1.Click
            result = Miscellaneous.M_DialogResult.Lock
            Me.Close()
        End Sub
        Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btn2.Click
            result = Miscellaneous.M_DialogResult.Delete
            Me.Close()
        End Sub
        Private Sub btnIgnore_Click(sender As Object, e As EventArgs) Handles btn3.Click
            result = Miscellaneous.M_DialogResult.Cancel
            Me.Close()
        End Sub

        '控制窗口的尺寸
        Private Sub LabelPrompt_SizeChanged(sender As Object, e As EventArgs) Handles LabelPrompt.SizeChanged
            Me.Height = Me.LabelPrompt.Height + 96
        End Sub
#End Region
    End Class


End Namespace