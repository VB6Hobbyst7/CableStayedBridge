Imports eZstd.eZAPI.APIWindows
Imports eZstd.eZAPI

Public Class Test_WindowProc
    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
        Static i As Decimal = 0 : i += 0.01 : Debug.Print(i)
        ' 当点击窗口的“关闭”按钮时，将窗口最小化
        If m.Msg = WindowsMessages.WM_SYSCOMMAND And m.WParam = SysCommands.SC_CLOSE Then
            Static j As Decimal = 0 : j += 0.01 : Debug.Print("进入" & j)
            ' 屏蔽传入的消息事件()
            Me.DefWndProc(m)
            Me.WindowState = FormWindowState.Minimized
            Return
        End If
        MyBase.WndProc(m)
    End Sub
End Class