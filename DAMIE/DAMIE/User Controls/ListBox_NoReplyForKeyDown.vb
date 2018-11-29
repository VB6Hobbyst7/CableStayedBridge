Imports System.ComponentModel
Imports eZstd.eZAPI

Namespace AME_UserControl

    ''' <summary>
    ''' 列表框控件，但是屏蔽了对于列表框的键盘事件。
    ''' 除非按下的键与列表框中的元素的第一个字符相同，否则此键按下不生效。
    ''' </summary>
    ''' <remarks></remarks>
    Friend Class ListBox_NoReplyForKeyDown
        Inherits System.Windows.Forms.ListBox

#Region "  ---  Properties"

        ''' <summary>
        ''' 当此列表框拥有焦点，并响应键盘按下的事件时，自动将焦点转移到此属性所指定的控件上，并向其发送此键盘消息。
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <Browsable(True), Description("当此列表框拥有焦点，并响应键盘按下的事件时，自动将焦点转移到此属性所指定的控件上，并向其发送此键盘消息。")>
        Public Property ParentControl As Control

#End Region

        Protected Overrides Sub WndProc(ByRef m As Message)
            If m.Msg = WindowsMessages.WM_KEYDOWN Then
                If ParentControl IsNot Nothing Then
                    ParentControl.Focus()
                    APIMessage.SendMessage(ParentControl.FindForm.Handle, _
                                           m.Msg, m.WParam, m.LParam)
                End If
            Else
                MyBase.WndProc(m)
            End If

        End Sub

    End Class
End Namespace