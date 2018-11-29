Namespace AME_UserControl
    ''' <summary>
    ''' 自定义控件：用来添加文件或者批量添加文件夹中的指定文件
    ''' </summary>
    ''' <remarks></remarks>
    Public Class AddFileOrDirectoryFiles

        ''' <summary>
        ''' 添加文件
        ''' </summary>
        ''' <remarks></remarks>
        Public Event AddFile As EventHandler
        ''' <summary>
        ''' 批量添加文件夹中的指定文件
        ''' </summary>
        ''' <remarks></remarks>
        Public Event AddFilesFromDirectory As EventHandler

        ''' <summary>
        ''' 提供给外部调用，用来从外部隐藏“添加文件”与“添加文件夹”两个标签
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub HideLabel()
            Me.PanelAddFileOrDir.Visible = False
        End Sub

        ''' <summary>
        ''' 添加文件
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub _AddFile(sender As Object, e As EventArgs) Handles lbAddFile.Click, btnAdd.Click
            RaiseEvent AddFile(sender, e)
        End Sub
        ''' <summary>
        ''' 添加文件夹中的文件
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub _AddDire(sender As Object, e As EventArgs) Handles lbAddDir.Click
            RaiseEvent AddFilesFromDirectory(sender, e)
        End Sub

#Region "  ---  UI界面显示"

        Private Sub PanelAdd_MouseLeave(sender As Object, e As EventArgs) _
        Handles Me.MouseLeave, btnAdd.Click, lbAddDir.Click, lbAddFile.Click, Me.LostFocus, Me.MouseLeave
            Me.HideLabel()
        End Sub

        Private Sub btnAdd_MouseEnter(sender As Object, e As EventArgs) Handles btnAdd.MouseEnter
            Me.PanelAddFileOrDir.Visible = True
        End Sub
        Private Sub colorFocused(sender As Object, e As EventArgs) Handles lbAddFile.MouseEnter, lbAddDir.MouseEnter
            Dim lb As System.Windows.Forms.Label = DirectCast(sender, System.Windows.Forms.Label)
            lb.BackColor = Color.LightBlue
        End Sub
        Private Sub colorLostFocus(sender As Object, e As EventArgs) Handles lbAddFile.MouseLeave, lbAddDir.MouseLeave
            Dim lb As System.Windows.Forms.Label = DirectCast(sender, System.Windows.Forms.Label)
            lb.BackColor = Color.White
        End Sub

#End Region
    End Class
End Namespace