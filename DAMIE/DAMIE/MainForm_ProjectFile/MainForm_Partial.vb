Imports System.Threading
Imports System.ComponentModel
Imports DAMIE.Miscellaneous
Namespace GlobalApp_Form

    Partial Class APPLICATION_MAINFORM

#Region "  --- 进度条的显示与隐藏 "

        ''' <summary>
        ''' 委托：在主程序界面上显示进度条与进度信息
        ''' </summary>
        ''' <remarks></remarks>
        Private Delegate Sub ShowProgressBar_MarqueeHandler()
        ''' <summary>
        ''' 在主程序界面上显示进度条与进度信息
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub ShowProgressBar_Marquee()
            With Me
                If .InvokeRequired Then
                    '非UI线程，再次封送该方法到UI线程
                    Me.BeginInvoke(New ShowProgressBar_MarqueeHandler(AddressOf Me.ShowProgressBar_Marquee))

                Else
                    'UI线程，进度更新
                    '主程序界面的进度条的UI显示

                    With .StatusLabel1
                        .Text = "Please Wait..."
                        .Visible = True
                    End With

                    With .ProgressBar1
                        .Visible = True
                        .Style = ProgressBarStyle.Marquee
                        '获取或设置进度块在进度栏内滚动所用的时间段，以毫秒为单位，其值越小，图形滚动得越快。
                        '只对于Style属性为Marquee的ProgressBar控件有效。
                        .MarqueeAnimationSpeed = 30
                    End With
                End If
            End With
        End Sub

        Private Delegate Sub ShowProgressBar_ContinueHander(ByVal Value As Integer)
        Public Sub ShowProgressBar_Continue(ByVal Value As Integer)
            With Me
                If .InvokeRequired Then
                    '非UI线程，再次封送该方法到UI线程
                    Me.BeginInvoke(New ShowProgressBar_ContinueHander(AddressOf Me.ShowProgressBar_Continue), {Value})
                Else
                    'UI线程，进度更新
                    '主程序界面的进度条的UI显示

                    With .StatusLabel1
                        .Text = "Please Wait..."
                        .Visible = True
                    End With

                    With .ProgressBar1
                        .Visible = True
                        .Style = ProgressBarStyle.Continuous
                        '获取或设置进度块在进度栏内滚动所用的时间段，以毫秒为单位，其值越小，图形滚动得越快。
                        '只对于Style属性为Marquee的ProgressBar控件有效。
                        .Value = Value
                    End With
                End If
            End With
        End Sub

        ''' <summary>
        ''' 委托：隐藏主程序界面中的进度条与进度信息
        ''' </summary>
        ''' <remarks></remarks>
        Private Delegate Sub HideProgressHandler(ByVal state As String)
        ''' <summary>
        ''' 隐藏主程序界面中的进度条与进度信息
        ''' </summary>
        ''' <param name="state">要显示在进度信息标签中的文本</param>
        ''' <remarks></remarks>
        Public Sub HideProgress(ByVal state As String)
            With Me
                If .InvokeRequired Then
                    '非UI线程，再次封送该方法到UI线程
                    Me.BeginInvoke(New HideProgressHandler(AddressOf Me.HideProgress), state)
                Else
                    'UI线程，进度更新
                    '主程序界面的进度条的UI显示
                    With Me.ProgressBar1
                        .Value = 100
                        .Visible = False
                    End With
                    Me.StatusLabel1.Text = state
                    With myTimer
                        .Interval = 500
                        .Start()
                    End With
                End If
            End With
        End Sub

        ' This is the method to run when the timer is raised.
        ''' <summary>
        ''' 闹钟触发的次数
        ''' </summary>
        ''' <remarks></remarks>
        Private alarmCounter As Integer = 1
        Private WithEvents myTimer As New System.Windows.Forms.Timer()
        ''' <summary>
        ''' 在指定的时间间隔内触发闹钟事件
        ''' </summary>
        ''' <param name="myObject"></param>
        ''' <param name="myEventArgs"></param>
        ''' <remarks></remarks>
        Private Sub TimerEventProcessor(myObject As Object, ByVal myEventArgs As EventArgs) Handles myTimer.Tick
            If alarmCounter = 2 Then
                With Me
                    .ProgressBar1.Visible = False
                    .StatusLabel1.Visible = False
                End With
                myTimer.Stop()
                alarmCounter = 1
            Else
                alarmCounter += 1
            End If
        End Sub

#End Region

#Region "  --- 控制主程序界面的显示 "

        Private Delegate Sub MainUI_ProjectNotOpenedHandler()
        ''' <summary>
        ''' 设置主程序中还没有任何项目文件打开时的UI界面
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub MainUI_ProjectNotOpened()
            If Me.InvokeRequired Then
                Me.BeginInvoke(New MainUI_ProjectNotOpenedHandler(AddressOf Me.MainUI_ProjectNotOpened))
            Else
                With Me
                    'Add What you want to do here.
                    '禁用菜单项 ------ 文件
                    .MenuItem_EditProject.Enabled = False
                    .MenuItem_SaveProject.Enabled = False
                    .MenuItem_SaveAsProject.Enabled = False
                    .MenuItemExport.Enabled = False
                    '禁用菜单项 ------ 编辑
                    .MenuItemDrawingPoints.Enabled = False
                    '禁用菜单项 ------ 绘图
                    .MenuItemSectionalView.Enabled = False
                    .MenuItemPlanView.Enabled = False
                    '禁用工具栏 ------ 同步滚动
                    .TlStrpBtn_Roll.Enabled = False
                End With
            End If
        End Sub

        Private Delegate Sub MainUI_ProjectOpenedHandler()
        ''' <summary>
        ''' 设置项目文档打开后的主程序的UI界面
        ''' </summary>
        ''' <remarks></remarks>
        Friend Sub MainUI_ProjectOpened()
            If Me.InvokeRequired Then
                Me.BeginInvoke(New MainUI_ProjectOpenedHandler(AddressOf Me.MainUI_ProjectOpened))
            Else
                With Me
                    'Add What you want to do here.
                    '启用菜单项 ------ 文件
                    .MenuItem_EditProject.Enabled = True
                    .MenuItem_SaveProject.Enabled = True
                    .MenuItem_SaveAsProject.Enabled = True
                    .MenuItemExport.Enabled = True
                    '启用菜单项 ------ 编辑
                    .MenuItemDrawingPoints.Enabled = True
                    '启用菜单项 ------ 绘图
                    .MenuItemSectionalView.Enabled = True
                    .MenuItemPlanView.Enabled = True
                    '启用工具栏 ------ 同步滚动
                    .TlStrpBtn_Roll.Enabled = True
                End With
            End If
        End Sub

        Private Delegate Sub MainUI_RollingObjectCreatedHandler()
        Friend Sub MainUI_RollingObjectCreated()
            If Me.InvokeRequired Then
                Me.BeginInvoke(New MainUI_RollingObjectCreatedHandler(AddressOf Me.MainUI_RollingObjectCreated))
            Else
                With Me
                    'Add What you want to do here.
                    '启用工具栏 ------ 同步滚动
                    .TlStrpBtn_Roll.Enabled = True
                End With
            End If
        End Sub
#End Region

    End Class
End Namespace