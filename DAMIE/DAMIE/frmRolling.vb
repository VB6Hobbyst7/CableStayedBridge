Imports System.Drawing
Imports DAMIE.AME_UserControl
Imports DAMIE.Miscellaneous
Imports DAMIE.All_Drawings_In_Application
Imports DAMIE.GlobalApp_Form

''' <summary>
''' 图形滚动窗口
''' </summary>
''' <remarks>最关键的触发图形滚动的事件是</remarks>
Public Class frmRolling

#Region "  ---  Declarations & Definitions"

#Region "  ---  Types"
    ''' <summary>
    ''' 用来进行滚动显示的绘图对象
    ''' </summary>
    ''' <remarks></remarks>
    Public Structure Drawings_For_Rolling
        Public PlanView As ClsDrawing_PlanView
        Public SectionalView As ClsDrawing_ExcavationElevation
        Public RollingMnt As List(Of clsDrawing_Mnt_RollingBase)

        Public Function Count() As UShort
            Dim SumUp As Short = RollingMnt.Count
            If PlanView IsNot Nothing Then
                SumUp += 1
            End If
            If SectionalView IsNot Nothing Then
                SumUp += 1
            End If
            Return SumUp
        End Function

        Public Sub New(Sender As frmRolling)
            RollingMnt = New List(Of clsDrawing_Mnt_RollingBase)
        End Sub
    End Structure

#End Region

    ''' <summary>
    ''' 日期的显示格式：2013/11/2
    ''' </summary>
    ''' <remarks></remarks>
    Const DateFormat_yyyyMd As String = "yyyy/M/d"

#Region "  ---  Properties"
    Private F_FocusOn_DateSpan As DateSpan
    ''' <summary>
    ''' 同步设置滚动条与日历的时间跨度值
    ''' </summary>
    ''' <value></value>
    ''' <remarks>只写属性,代表滚动界面中选择的图表对象的时间跨度的最大并集</remarks>
    Public Property FocusOn_DateSpan As DateSpan
        Get
            Return Me.F_FocusOn_DateSpan
        End Get
        Private Set(value As DateSpan)
            '日历的最小日期不能早于1753年1月1日。
            If Date.Compare(value.FinishedDate, New Date(1753, 1, 1)) > 0 Then
                Try
                    If Date.Compare(value.StartedDate, Calendar_Construction.MaxDate) < 0 Then
                        Calendar_Construction.MinDate = value.StartedDate
                        Calendar_Construction.MaxDate = value.FinishedDate
                    Else  '如果新的最小值比旧的最大值还要大，则先设置最大值，以避免出现先设置最小值时，它比最大值还要大的异常
                        Calendar_Construction.MaxDate = value.FinishedDate
                        Calendar_Construction.MinDate = value.StartedDate
                    End If
                Catch ex As ArgumentOutOfRangeException
                    Debug.Print(vbCrLf & ex.Message &
                                vbCrLf & "报错位置：" & ex.TargetSite.Name &
                                "日历范围设置出错，但不用进行修正处理。")
                Finally
                    Me.F_FocusOn_DateSpan = value
                End Try
            End If
        End Set
    End Property

    ''' <summary>
    ''' 在滚动时记录的“当天”的日期值。
    ''' </summary>
    ''' <remarks></remarks>
    Dim F_Rollingday As Date
    ''' <summary>
    ''' !!!触发图形滚动事件。
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Rollingday As Date
        Get
            Return F_Rollingday
        End Get
        Set(value As Date) ' 改变RollingDay属性，以触发图形滚动事件。
            If Date.Compare(value, Me.F_Rollingday) <> 0 Then
                '触发图形滚动事件。
                Call Rolling(value, Me.F_arrThreadDelegete)
                '触发DateChanged事件，在事件中会更新日期label的值。
                Me.LabelDate.Text = value.ToString(DateFormat_yyyyMd)
            End If
            Me.F_Rollingday = value
        End Set
    End Property

#End Region

#Region "  ---  Fields"

    ''' <summary>
    ''' 窗口中是否选择了用来进行滚动的图形
    ''' </summary>
    ''' <remarks></remarks>
    Private F_blnHasDrawingToRoll As Boolean

    ''' <summary>
    ''' 列表中选择的用来进行滚动显示的绘图对象
    ''' </summary>
    ''' <remarks></remarks>
    Public F_SelectedDrawings As Drawings_For_Rolling

    ''' <summary>
    ''' 批量操作绘图曲线的对话框。用来对多条曲线进行锁定或者删除
    ''' </summary>
    ''' <remarks></remarks>
    Private F_frmHandleMultipleCurve As New frmLockCurves

#End Region

#End Region

#Region "  ---  构造函数与窗体的加载、打开与关闭"

    ''' <summary>
    ''' 构造函数
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        '
        Me.FocusOn_DateSpan = Nothing
        Me.F_Rollingday = Date.Today
        CheckBox_PlanView.Enabled = False
        CheckBox_SectionalView.Enabled = False
        '
        Me.F_SelectedDrawings = New Drawings_For_Rolling(Me)
    End Sub

    ''' <summary>
    ''' 窗体加载
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub frmRolling_Load(sender As Object, e As EventArgs) Handles Me.Load
        '窗口于所有控件之前先接收键盘事件
        '如果此属性设置为True，当键盘按下时，此窗口会接收 KeyPress, KeyDown和KeyUp事件，
        '当窗口完全执行完这些键盘事件的方法后，键盘事件才会接着传递给拥有焦点的那个控件。
        Me.KeyPreview = True
        '在窗口第一次出现时，根据是否有可滚动的图形设置初始界面。
        Call Me.OnRollingDrawingsRefreshed
    End Sub

    ''' <summary>
    ''' 窗体关闭之前的事件
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>在关闭窗口时将其隐藏</remarks>
    Private Sub frmRolling_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        '如果是子窗口自己要关闭，则将其隐藏
        '如果是mdi父窗口要关闭，则不隐藏，而由父窗口去结束整个进程
        If Not e.CloseReason = CloseReason.MdiFormClosing Then
            Me.Hide()
            e.Cancel = True
        End If
    End Sub

#End Region

#Region "  ---  刷新UI界面——所有滚动窗口"

    ''' <summary>
    ''' 委托：在非UI线程中刷新窗口控件时，用Me.BeginInvoke来转到创建此窗口的线程中来执行。
    ''' </summary>
    ''' <remarks></remarks>
    Private Delegate Sub RollingDrawingsRefreshedHandler()
    ''' <summary>
    ''' 在此方法中，触发了滚动窗口的RollingDrawingsRefreshed事件：
    ''' 刷新滚动窗口中的列表框的数据与其UI显示
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub OnRollingDrawingsRefreshed() Handles btnRefresh.Click
        If Me.InvokeRequired Then
            Me.BeginInvoke(New RollingDrawingsRefreshedHandler(AddressOf Me.OnRollingDrawingsRefreshed))
        Else
            With Me
                Try
                    Me.F_SelectedDrawings = RefreshUI_RollingDrawings()
                    Call SelectedRollingDrawingsChanged(Me.F_SelectedDrawings)
                Catch ex As Exception
                    MessageBox.Show("在添加或者删除程序中的图表时，滚动窗口的界面刷新出错。" & vbCrLf & ex.Message & vbCrLf & "报错位置：" & ex.TargetSite.Name, _
      "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End With
        End If
    End Sub

    ''' <summary>
    ''' 在整个程序中的可以滚动的图表发生增加或者减少时触发的事件：
    ''' 刷新窗口中的列表框的数据与其UI显示
    ''' </summary>
    ''' <remarks></remarks>
    Private Function RefreshUI_RollingDrawings() As Drawings_For_Rolling
        Dim SelectedDrawings As New Drawings_For_Rolling(Me)
        ' 主程序中所有可以滚动的图形的汇总()
        Dim RollingMethods As RollingEnabledDrawings = GlobalApplication.Application.ExposeRollingDrawings()

        '
        Dim plan As ClsDrawing_PlanView = RollingMethods.PlanView
        SelectedDrawings.PlanView = plan
        With CheckBox_PlanView
            .Tag = plan
            If plan IsNot Nothing Then
                .Checked = True
                .Enabled = True
            Else
                .Checked = False
                .Enabled = False
            End If
        End With
        '
        Dim Sectional As ClsDrawing_ExcavationElevation = RollingMethods.SectionalView
        SelectedDrawings.SectionalView = Sectional
        With CheckBox_SectionalView
            .Tag = Sectional
            If Sectional IsNot Nothing Then
                .Checked = True
                .Enabled = True
            Else
                .Checked = False
                .Enabled = False
            End If
        End With

        ' --------------  为窗口中的控件赋值  ------------------------
        With RollingMethods
            '
            With Me.ListBoxMonitorData
                Dim listMnt As New List(Of LstbxDisplayAndItem)

                For Each M As clsDrawing_Mnt_RollingBase In RollingMethods.MonitorData
                    listMnt.Add(New LstbxDisplayAndItem(DisplayedText:=M.Chart_App_Title, _
                                                                          Value:=M))
                Next
                .DisplayMember = LstbxDisplayAndItem.DisplayMember
                .DataSource = listMnt
                SelectedDrawings.RollingMnt.Clear()
                For Each item As LstbxDisplayAndItem In .SelectedItems
                    SelectedDrawings.RollingMnt.Add(DirectCast(item.Value, clsDrawing_Mnt_RollingBase))
                Next
            End With
        End With

        Return SelectedDrawings
    End Function

#End Region

#Region "  ---  刷新进行滚动的线程"

    '选择复选框或者列表项——更新选择的图表对象
    Private Sub CheckBox_PlanView_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox_PlanView.CheckedChanged
        If CheckBox_PlanView.Checked Then
            Me.F_SelectedDrawings.PlanView = DirectCast(CheckBox_PlanView.Tag, ClsDrawing_PlanView)
        Else
            Me.F_SelectedDrawings.PlanView = Nothing
        End If
        '
        Call SelectedRollingDrawingsChanged(Me.F_SelectedDrawings)
    End Sub
    Private Sub CheckBox_SectionalView_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox_SectionalView.CheckedChanged
        If CheckBox_SectionalView.Checked Then
            Me.F_SelectedDrawings.SectionalView = DirectCast(CheckBox_SectionalView.Tag, ClsDrawing_ExcavationElevation)
        Else
            Me.F_SelectedDrawings.SectionalView = Nothing
        End If
        '
        Call SelectedRollingDrawingsChanged(Me.F_SelectedDrawings)
    End Sub
    Private Sub ListBoxMonitorData_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBoxMonitorData.SelectedIndexChanged
        Me.F_SelectedDrawings.RollingMnt.Clear()
        '
        Dim Items = ListBoxMonitorData.SelectedItems
        Dim Drawing As clsDrawing_Mnt_RollingBase
        For Each lstboxItem As LstbxDisplayAndItem In Items
            Drawing = DirectCast(lstboxItem.Value, clsDrawing_Mnt_RollingBase)
            Me.F_SelectedDrawings.RollingMnt.Add(Drawing)
        Next
        '
        Call SelectedRollingDrawingsChanged(Me.F_SelectedDrawings)
    End Sub

    '选择复选框或者列表项——更新滚动线程与窗口界面
    ''' <summary>
    ''' ！选择的图形发生改变时，更新滚动线程与窗口界面。
    ''' </summary>
    ''' <param name="Selected_Drawings">更新后的要进行滚动的图形</param>
    ''' <remarks>此方法不能直接Handle复选框的CheckedChanged或者列表框的SelectedIndexChanged事件，
    ''' 因为此方法必须是在更新了Me.F_SelectedDrawings属性之后，才能去更新窗口界面。</remarks>
    Private Sub SelectedRollingDrawingsChanged(ByVal Selected_Drawings As Drawings_For_Rolling)
        '
        Call ConstructRollingThreads(Selected_Drawings)
        Me.FocusOn_DateSpan = Refesh_FocusOn_DateSpan(Selected_Drawings)
        '
        If Selected_Drawings.Count > 0 Then
            F_blnHasDrawingToRoll = True
        Else
            F_blnHasDrawingToRoll = False
        End If
        Call RefreshUI_SelectedDrawings(F_blnHasDrawingToRoll, Me.F_FocusOn_DateSpan)
        '
    End Sub

#Region "  ---  子方法"

    ''' <summary>
    ''' 为每一个图形滚动设置一个线程
    ''' </summary>
    ''' <remarks></remarks>
    Private F_arrThread() As Thread
    ''' <summary>
    ''' 每一个图形滚动的线程所指向的方法
    ''' </summary>
    ''' <remarks></remarks>
    Private F_arrThreadDelegete() As ParameterizedThreadStart
    ''' <summary>
    ''' 为每一个选择的要进行滚动的图形构造一个线程
    ''' </summary>
    ''' <param name="Selected_Drawings"></param>
    ''' <remarks></remarks>
    Private Sub ConstructRollingThreads(ByVal Selected_Drawings As Drawings_For_Rolling)
        Call AbortAllThread(F_arrThread)

        ' -------------------------------------  包含所有要进行滚动的线程的数组 
        Dim SelectedDrawingsCount As Short = Selected_Drawings.Count
        '
        ReDim F_arrThread(0 To SelectedDrawingsCount - 1)
        ReDim F_arrThreadDelegete(0 To SelectedDrawingsCount - 1)
        ''
        Dim btThread As Short = 0
        '
        '
        ' ------------------------------------------ 标高剖面图
        Dim ele As ClsDrawing_ExcavationElevation = Selected_Drawings.SectionalView
        If ele IsNot Nothing Then
            '为线程数组中的元素线程赋值
            F_arrThreadDelegete(btThread) = New ParameterizedThreadStart(AddressOf ele.Rolling)
            Dim thd As New Thread(F_arrThreadDelegete(btThread))
            thd.Name = "滚动标高剖面图"
            F_arrThread(btThread) = thd
            F_arrThread(btThread).IsBackground = True
            btThread += 1
        End If

        ' ------------------------------------------ 开挖平面图
        Dim Plan As ClsDrawing_PlanView = Selected_Drawings.PlanView
        If Plan IsNot Nothing Then
            '为线程数组中的元素线程赋值
            F_arrThreadDelegete(btThread) = New ParameterizedThreadStart(AddressOf Plan.Rolling)
            Dim thd As New Thread(F_arrThreadDelegete(btThread))
            thd.Name = "滚动开挖平面图"
            F_arrThread(btThread) = thd
            F_arrThread(btThread).IsBackground = True
            btThread += 1
        End If

        ' ------------------------------------------ 监测曲线图
        For Each Moni As clsDrawing_Mnt_RollingBase In Selected_Drawings.RollingMnt
            '为线程数组中的元素线程赋值
            F_arrThreadDelegete(btThread) = New ParameterizedThreadStart(AddressOf Moni.Rolling)
            Dim thd As New Thread(F_arrThreadDelegete(btThread))
            thd.Name = "滚动监测曲线图"
            F_arrThread(btThread) = thd
            F_arrThread(btThread).IsBackground = True
            btThread += 1
        Next
    End Sub

    ''' <summary>
    ''' !!! 在每一次更改选择项时刷新滚动条和日历的时间跨度，以及终结旧线程，创建新的线程
    ''' </summary>
    ''' <remarks></remarks>
    Private Function Refesh_FocusOn_DateSpan(ByVal Selected_Drawings As Drawings_For_Rolling) As DateSpan
        ' -------------------------------------  包含所有要进行滚动的线程的数组 
        '
        Dim blnDateSpanInitialized As Boolean = False
        Dim DateSpan As DateSpan
        '
        ' ------------------------------------------ 标高剖面图
        Dim ele As ClsDrawing_ExcavationElevation = Selected_Drawings.SectionalView
        If ele IsNot Nothing Then
            '更新滚动界面的时间跨度
            DateSpan = RenewTimeSpan(blnDateSpanInitialized, DateSpan, ele.DateSpan)
        End If

        ' ------------------------------------------ 开挖平面图
        Dim Plan As ClsDrawing_PlanView = Selected_Drawings.PlanView
        If Plan IsNot Nothing Then
            '更新滚动界面的时间跨度
            DateSpan = RenewTimeSpan(blnDateSpanInitialized, DateSpan, Plan.DateSpan)
        End If

        ' ------------------------------------------ 监测曲线图
        For Each Moni As clsDrawing_Mnt_RollingBase In Selected_Drawings.RollingMnt
            '更新滚动界面的时间跨度
            DateSpan = RenewTimeSpan(blnDateSpanInitialized, DateSpan, Moni.DateSpan)
        Next
        Return DateSpan
    End Function
    ''' <summary>
    ''' 更新滚动窗口中的日期跨度
    ''' </summary>
    ''' <param name="blnDateSpanInitialized">引用类型，指示滚动窗口中是否已经有了最初的作为基准日期跨度的值</param>
    ''' <param name="OldDateSpan">已经有的日期跨度的值</param>
    ''' <param name="ObjectDateSpan">进行扩充的日期跨度</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function RenewTimeSpan(ByRef blnDateSpanInitialized As Boolean, _
                                 ByVal OldDateSpan As DateSpan, ByVal ObjectDateSpan As DateSpan) As DateSpan
        Dim NewDateSpan As DateSpan
        If Not blnDateSpanInitialized Then
            NewDateSpan = ObjectDateSpan
            blnDateSpanInitialized = True
        Else
            NewDateSpan = ExpandDateSpan(OldDateSpan, ObjectDateSpan)
        End If
        Return NewDateSpan
    End Function

#End Region

#End Region

#Region "  ---  触发图形滚动的操作"

    ''' <summary>
    ''' 1.通过键盘的方向键来进行图形的滚动。按下键盘时列表框不要处于激活状态。
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>要注意的是，在按下键盘时，如果窗口中的列表框是处于激活状态，
    ''' 则程序会同时执行ListBox.SelectedIndexChanged与此处的KeyDown，那么，
    ''' 其结果就是先启动了一个图形的滚动线程，然后第二个事件又马上将此线程关闭了。</remarks>
    Private Sub frmRolling_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If F_blnHasDrawingToRoll Then
            If e.KeyCode = Keys.Left Then
                Call Me.NumChanging_ValueMinused()
            ElseIf e.KeyCode = Keys.Right Then
                Call Me.NumChanging_ValueAdded()
            End If
        End If
    End Sub

    ''' <summary>
    ''' 2.日期的增加
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub NumChanging_ValueAdded() Handles NumChanging.ValueAdded
        Dim NewDay As Date
        With NumChanging
            Select Case .unit
                Case UsrCtrl_NumberChanging.YearMonthDay.Days
                    NewDay = Me.F_Rollingday.AddDays(.Value_TimeSpan)
                Case UsrCtrl_NumberChanging.YearMonthDay.Months
                    NewDay = Me.F_Rollingday.AddMonths(.Value_TimeSpan)
                Case UsrCtrl_NumberChanging.YearMonthDay.Years
                    NewDay = Me.F_Rollingday.AddYears(.Value_TimeSpan)
            End Select
        End With

        '限制日期的跨度在最大跨度范围之内
        If Date.Compare(NewDay, Me.FocusOn_DateSpan.FinishedDate) > 0 Then
            '这里没有判断与最早日期的比较，是因为这里的操作是将日期增加
            NewDay = Me.FocusOn_DateSpan.FinishedDate
        End If
        Me.Rollingday = NewDay

    End Sub
    ''' <summary>
    ''' 2.日期的后退
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub NumChanging_ValueMinused() Handles NumChanging.ValueMinused
        Dim Newday As Date
        With NumChanging
            Select Case .unit
                Case UsrCtrl_NumberChanging.YearMonthDay.Days
                    Newday = Me.F_Rollingday.AddDays(-.Value_TimeSpan)
                Case UsrCtrl_NumberChanging.YearMonthDay.Months
                    Newday = Me.F_Rollingday.AddMonths(-.Value_TimeSpan)
                Case UsrCtrl_NumberChanging.YearMonthDay.Years
                    Newday = Me.F_Rollingday.AddYears(-.Value_TimeSpan)
            End Select
        End With

        '限制日期的跨度在最大跨度范围之内
        If Date.Compare(Newday, Me.FocusOn_DateSpan.StartedDate) < 0 Then
            '这里没有判断与最晚日期的比较，是因为这里的操作是将日期减小
            Newday = Me.FocusOn_DateSpan.StartedDate
        End If
        '改变RollingDay属性，以触发图形滚动事件。
        Me.Rollingday = Newday
    End Sub

    ''' <summary>
    ''' 3.从日历中选择指定日期
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Calendar_Construction_DateSelected(sender As Object, e As DateRangeEventArgs) Handles Calendar_Construction.DateSelected
        With Me.Calendar_Construction
            Dim NewDate As Date = .SelectionEnd.Date
            '改变RollingDay属性，以触发图形滚动事件。
            Me.Rollingday = NewDate
        End With
    End Sub

    ''' <summary>
    ''' 4.在窗口中选择的进行滚动的窗口发生改变时刷新UI界面：日历、日期前进或后退
    ''' </summary>
    ''' <param name="blnHasDrawingToRoll">指示是否有要进行滚动的图形</param>
    ''' <param name="DateSpan">当前选择的滚动图形的日期跨度的并集。
    ''' 如果窗口中没有要进行滚动的图形，则此参数值无效</param>
    ''' <remarks></remarks>
    Private Sub RefreshUI_SelectedDrawings(ByVal blnHasDrawingToRoll As Boolean, ByVal DateSpan As DateSpan)
        If Not blnHasDrawingToRoll Then      '如果窗口中没有选择任何一项
            '窗口的界面刷新
            Me.Panel_Roll.Enabled = False
            Me.Calendar_Construction.Visible = False
            '设置当天的日期
            Me.LabelDate.Text = Me.F_Rollingday.ToString(DateFormat_yyyyMd)
        Else
            '窗口的界面刷新
            Me.Panel_Roll.Enabled = True
            Me.Calendar_Construction.Visible = True
            '设置当天的日期
            Dim NewDay As Date = Me.F_Rollingday
            '如果超出日期跨度的范围，则设置其为跨度的边界
            If Date.Compare(Me.F_Rollingday, DateSpan.StartedDate) < 0 Then
                NewDay = DateSpan.StartedDate
            ElseIf Date.Compare(Me.F_Rollingday, DateSpan.FinishedDate) > 0 Then
                NewDay = DateSpan.FinishedDate
            End If
            '如果文本框前后的字符不变，那么此时是否会触发其TextChanged事件呢——
            '会，因为程序只会只会判断是否进行了重新赋值，而不会判断前后的值是否相同。
            Me.LabelDate.Text = NewDay.ToString(DateFormat_yyyyMd)
        End If
    End Sub

#End Region

#Region "  ---  图形滚动"

    ''' <summary>
    ''' 【关键】滚动时的执行方法
    ''' </summary>
    ''' <param name="ThisDay">进行滚动的当天的日期</param>
    ''' <param name="arrThreadDelegete">每一个图形滚动的线程所指向的方法</param>
    ''' <remarks>在进行滚动时，先判断是否还有线程没有执行完，如果有，则要先将未结束的线程取消掉，然后再重新启动线程。
    ''' 而线程的数量与调用的方法是不变的。</remarks>
    Private Sub Rolling(ByVal ThisDay As Date, ByVal arrThreadDelegete As Threading.ParameterizedThreadStart())

        '取消未结束的线程
        Call AbortAllThread(Me.F_arrThread)
        Try
            '创建新线程并重新启动
            If arrThreadDelegete IsNot Nothing Then
                For i As Integer = 0 To arrThreadDelegete.Count - 1
                    Me.F_arrThread(i) = New Thread(arrThreadDelegete(i))
                    Me.F_arrThread(i).Start(ThisDay)
                Next
            End If
        Catch ex As Exception
            Debug.Print("图形滚动出错！——StartRolling")
        End Try
    End Sub

    ''' <summary>
    ''' 取消所有当前正在运行的线程
    ''' </summary>
    ''' <param name="Threads"></param>
    ''' <remarks></remarks>
    Private Sub AbortAllThread(ByVal Threads() As Thread)
        If Threads IsNot Nothing Then
            For Each thd As Thread In Threads
                If thd.IsAlive Then
                    '在Abort方法中会自动抛出异常：ThreadAbortException
                    thd.Abort()
                    Debug.Print("线程成功终止。")
                End If
            Next
        End If
    End Sub

#End Region

#Region "  --- 一般界面操作"

    Private Sub ProgressBar_PlanView_Click(sender As Object, e As EventArgs) Handles ProgressBar_PlanView.Click
        With CheckBox_PlanView
            If .Enabled Then
                If .Checked Then
                    .Checked = False
                Else
                    .Checked = True
                End If
            End If
        End With
    End Sub
    Private Sub ProgressBar_SectionalView_Click(sender As Object, e As EventArgs) Handles ProgressBar_SectionalView.Click
        With CheckBox_SectionalView
            If .Enabled Then
                If .Checked Then
                    .Checked = False
                Else
                    .Checked = True
                End If
            End If
        End With
    End Sub

    ''' <summary>
    ''' 输出到Word
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnOutPut_Click(sender As Object, e As EventArgs) Handles btnOutPut.Click
        Call APPLICATION_MAINFORM.MainForm.ExportToWord(Nothing, Nothing)
    End Sub

    ''' <summary>
    ''' 进行批量操作
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btn_GroupHandle_Click(sender As Object, e As EventArgs) Handles btn_GroupHandle.Click
        Me.F_frmHandleMultipleCurve.ShowDialog()
    End Sub
#End Region

End Class
