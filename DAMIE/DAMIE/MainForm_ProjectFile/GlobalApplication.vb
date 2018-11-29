Imports Microsoft.Office.Interop
Imports DAMIE.Miscellaneous
Imports DAMIE.DataBase
Imports DAMIE.All_Drawings_In_Application

Namespace GlobalApp_Form

    Public Class GlobalApplication ' GlobalApplication

#Region "  ---  属性值的定义"

        ''' <summary>
        ''' 用来索引用来保存全局数据的类实例
        ''' </summary>
        Private Shared F_Application As GlobalApplication 'GlobalApplication.Application
        ''' <summary>
        ''' 共享属性，用来索引用来保存全局数据的类实例
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared ReadOnly Property Application As GlobalApplication
            Get
                Return F_Application
            End Get
        End Property


        ''' <summary>
        ''' 整个程序中用来放置各种隐藏的Excel数据文档的Application对象
        ''' </summary>
        ''' <remarks></remarks>F:\基坑数据\程序编写\群坑分析\AME\MainForm_ProjectFile\GlobalApplication.vb
        Private F_ExcelApplication_DB As Excel.Application
        ''' <summary>
        ''' 读取或设置整个程序中用来放置各种隐藏的Excel数据文档的Application对象
        ''' </summary>
        ''' <value></value>
        ''' <returns>一个Excel.Application对象，用来装载整个程序中的所有隐藏的后台数据的Excel文档</returns>
        ''' <remarks></remarks>
        Public Property ExcelApplication_DB As Excel.Application
            Get
                '如果此时还没有打开装载Excel数据库工作簿的Excel程序，则先创建一个Excel程序
                If Me.F_ExcelApplication_DB Is Nothing Then
                    Me.F_ExcelApplication_DB = New Excel.Application
                    Me.F_ExcelApplication_DB.Visible = False
                End If
                '
                Return Me.F_ExcelApplication_DB
            End Get
            Set(value As Excel.Application)
                Me.F_ExcelApplication_DB = value
            End Set
        End Property

        ''' <summary>
        ''' 进行同步滚动的时间跨度，用来给出时间滚动条与日历的范围。
        ''' </summary>
        ''' <remarks></remarks>
        Private F_DateSpan As DateSpan
        ''' <summary>
        ''' 进行同步滚动的时间跨度，用来给出时间滚动条与日历的范围。
        ''' 这个全局的DateSpan值是只有扩充，不会缩减的。即在程序中每新添加一个图表，
        ''' 就会对其进行扩充，而将图表删除时，这个日期跨度值并不会缩减。
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property DateSpan As DateSpan
            Get
                Return F_DateSpan
            End Get
        End Property


#Region "  -！-！-  绘图程序界面"

        ''' <summary>
        ''' 用来绘制剖面标高图的那一个Excel程序
        ''' </summary>
        ''' Excel程序界面  ---  标高剖面图
        Private F_ElevationDrawing As ClsDrawing_ExcavationElevation
        ''' <summary>
        ''' 用来绘制剖面标高图的那一个Excel程序
        ''' </summary>
        Public Property ElevationDrawing As ClsDrawing_ExcavationElevation
            Get
                Return F_ElevationDrawing
            End Get
            Set(value As ClsDrawing_ExcavationElevation)
                F_ElevationDrawing = value
                '刷新滚动窗口的列表框的界面显示
                APPLICATION_MAINFORM.MainForm.Form_Rolling.OnRollingDrawingsRefreshed()
            End Set
        End Property

        'Excel程序界面  ---  集合  ---  监测曲线绘图
        Private F_MntDrawing_ExcelApps As New Dictionary_AutoKey(Of Cls_ExcelForMonitorDrawing)
        ''' <summary>
        ''' Excel程序界面  ---  集合  ---  监测曲线绘图
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property MntDrawing_ExcelApps As Dictionary_AutoKey(Of Cls_ExcelForMonitorDrawing)
            Get
                Return F_MntDrawing_ExcelApps
            End Get
            Set(value As Dictionary_AutoKey(Of Cls_ExcelForMonitorDrawing))
                Me.F_MntDrawing_ExcelApps = value
            End Set
        End Property

        'Visio程序界面  ---  开挖平面图
        Private F_PlanView_VisioWindow As ClsDrawing_PlanView
        Public Property PlanView_VisioWindow As ClsDrawing_PlanView
            Get
                Return F_PlanView_VisioWindow
            End Get
            Set(value As ClsDrawing_PlanView)
                F_PlanView_VisioWindow = value
                '刷新滚动窗口的列表框的界面显示
                APPLICATION_MAINFORM.MainForm.Form_Rolling.OnRollingDrawingsRefreshed()
            End Set
        End Property

#End Region

#Region "  ---  项目文件与数据库"

        ''' <summary>
        ''' 程序中正在运行的那个项目文件
        ''' </summary>
        ''' <remarks></remarks>
        Private F_ProjectFileContents As clsProjectFile
        ''' <summary>
        ''' 程序中正在运行的那个项目文件
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property ProjectFile As clsProjectFile
            Get
                If Me.F_ProjectFileContents Is Nothing Then
                    Me.F_ProjectFileContents = New clsProjectFile
                End If
                Return Me.F_ProjectFileContents
            End Get
            Set(value As clsProjectFile)
                Me.F_ProjectFileContents = value
            End Set
        End Property

        ''' <summary>
        ''' 主程序的数据库文件
        ''' </summary>
        Private F_DataBase As ClsData_DataBase
        ''' <summary>
        ''' 主程序的数据库文件
        ''' </summary>
        ''' <remarks>用来加载所有Excel的数据文档，在整个过程中不可见，即在隐藏的状态下进行数据的交换与存储。</remarks>
        Public Property DataBase As ClsData_DataBase
            Get
                Return F_DataBase
            End Get
            Set(value As ClsData_DataBase)
                F_DataBase = value
            End Set
        End Property

#End Region


#End Region

#Region "  ---  主程序的加载与关闭"

        Public Sub New()
            '设置一个初始值blnTimeSpanInitialized，说明MainForm.TimeSpan还没有被初始化过，后期应该先对其赋值以进行初始化，然后再进行比较
            Me.blnTimeSpanInitialized = False
            GlobalApplication.F_Application = Me
        End Sub

#End Region

#Region "  ---  遍历程序中的绘图"

        ''' <summary>
        ''' 主程序中所有的绘图，包括不能进行滚动的图形。比如开挖平面图，监测曲线图，开挖剖面图。
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ExposeAllDrawings() As AmeDrawings
            Dim Excel_Monitor As Dictionary_AutoKey(Of Cls_ExcelForMonitorDrawing)
            '
            Dim Visio_PlanView As ClsDrawing_PlanView = Nothing
            Dim Elevation As ClsDrawing_ExcavationElevation
            Dim MonitorData As New List(Of ClsDrawing_Mnt_Base)
            '
            With GlobalApplication.Application
                Elevation = .ElevationDrawing
                Excel_Monitor = .MntDrawing_ExcelApps
            End With
            '

            'Me.F_AllDrawingsCount = 0
            '
            '对象名称与对象
            Try
                '剖面图
                Elevation = GlobalApplication.Application.ElevationDrawing
                If Elevation IsNot Nothing Then
                    'Me.F_AllDrawingsCount += 1
                End If
            Catch err As Exception

            End Try

            Try
                '平面图
                Visio_PlanView = GlobalApplication.Application.PlanView_VisioWindow
                If Visio_PlanView IsNot Nothing Then
                    'Me.F_AllDrawingsCount += 1
                End If
            Catch err As Exception

            End Try

            Try
                '监测曲线图
                For Each MonitorSheets As Cls_ExcelForMonitorDrawing In Excel_Monitor.Values
                    For Each sht As ClsDrawing_Mnt_Base In MonitorSheets.Mnt_Drawings.Values

                        MonitorData.Add(sht)
                        'Me.F_AllDrawingsCount += 1
                    Next
                Next
            Catch err As Exception
            End Try
            Dim AllDrawings As New AmeDrawings(SectionalView:=Elevation, _
                                               PlanView:=Visio_PlanView, _
                                               MonitorData:=MonitorData)
            Return AllDrawings
        End Function

        ''' <summary>
        ''' 提取出主程序中含有Rolling方法的对象
        ''' </summary>
        ''' <returns>其中的每一个元素都是一个字典，以对象名称来索引对象
        ''' 其中有三个元素，依次代表：剖面图、平面图和监测曲线图，
        ''' 它不是指从mainform中提取出来的三个App的属性值，而是从这三个App属性中挑选出来的，正确地带有Rolling方法的相关类的实例对象。
        ''' </returns>
        ''' <remarks></remarks>
        Public Function ExposeRollingDrawings() As RollingEnabledDrawings
            '
            Dim Elevation As ClsDrawing_ExcavationElevation
            Dim Visio_PlanView As ClsDrawing_PlanView
            Dim Excel_Monitor As Dictionary_AutoKey(Of Cls_ExcelForMonitorDrawing)
            With GlobalApplication.Application
                Elevation = .ElevationDrawing
                Visio_PlanView = .PlanView_VisioWindow
                Excel_Monitor = .MntDrawing_ExcelApps
            End With
            '
            '对象名称与对象
            Try
                '剖面图
                If Elevation IsNot Nothing Then
                    'RollingDrawings.SectionalView.Add(Elevation)
                End If
            Catch err As Exception

            End Try

            Try
                '平面图
                If Visio_PlanView IsNot Nothing Then
                    'RollingDrawings.PlanView.Add(Visio_PlanView)
                End If
            Catch err As Exception

            End Try
            Dim RollingMntDrawings As New List(Of clsDrawing_Mnt_RollingBase)
            Try
                '监测曲线图
                For Each MonitorSheets As Cls_ExcelForMonitorDrawing In Excel_Monitor.Values
                    For Each sht As ClsDrawing_Mnt_Base In MonitorSheets.Mnt_Drawings.Values
                        If sht.CanRoll Then
                            RollingMntDrawings.Add(sht)
                        End If
                    Next
                Next
                ''数组中的每一个元素都是一个字典，以对象名称来索引对象
            Catch err As Exception
            End Try
            '
            Dim RollingDrawings As New RollingEnabledDrawings(Elevation, Visio_PlanView, RollingMntDrawings)
            Return RollingDrawings
        End Function

#End Region

        ''' <summary>
        ''' 布尔值，用来指示主程序的时间跨度是否已被初始化。
        ''' </summary>
        ''' <remarks>在操作上，如果还没有被初始化，则以添加的形式进行初始化；
        ''' 如果已初始化，则是比较并扩展的形式进行初始化</remarks>
        Private blnTimeSpanInitialized As Boolean = False
        ''' <summary>
        ''' 将主程序的TimeSpan属性与指定的TimeSpan值进行比较，并将主程序的TimeSpan扩展为二者的并集的大区间
        ''' </summary>
        ''' <param name="ComparedDateSpan">要与主程序的TimeSpan进行比较的TimeSpan值</param>
        ''' <remarks></remarks>
        Public Sub refreshGlobalDateSpan(ByVal ComparedDateSpan As DateSpan)
            With Me
                If Not blnTimeSpanInitialized Then
                    .F_DateSpan = ComparedDateSpan
                    blnTimeSpanInitialized = True
                Else
                    .F_DateSpan = ExpandDateSpan(.DateSpan, ComparedDateSpan)
                End If
            End With
        End Sub

        ''' <summary>
        ''' 绘制或删除测点在Visio中的对应形状。
        ''' </summary>
        ''' <remarks>一个Form对象，用来绘制或删除测点在Visio中的对应形状。</remarks>
        Private Form_TreeView_MonitorPoints As DiaFrm_PointsTreeView
        ''' <summary>
        ''' 打开用于在Visio中绘制监测点位形状的对话框
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub DrawingPointsInVisio()
            '是否有可以正常打开的，用来绘制监测点位图的窗口对象。
            If Me.F_PlanView_VisioWindow IsNot Nothing Then     '如果Visio绘图已经打开
                If Me.F_PlanView_VisioWindow.HasMonitorPointsInfo Then
                    ' --------- 判断绘制测点图的窗口是否存在并有效
                    Dim blnPointFormValidated As Boolean = True
                    If Me.Form_TreeView_MonitorPoints Is Nothing Then
                        blnPointFormValidated = False
                    Else
                        If Me.Form_TreeView_MonitorPoints.IsDisposed Then
                            blnPointFormValidated = False
                        End If
                    End If
                    ' --------- 判断绘制测点图的窗口是否存在并有效

                    If blnPointFormValidated Then
                        Me.Form_TreeView_MonitorPoints.ShowDialog()
                    Else
                        If Me.DataBase IsNot Nothing Then
                            If Me.DataBase.sheet_Points_Coordinates IsNot Nothing Then
                                Me.Form_TreeView_MonitorPoints = New DiaFrm_PointsTreeView( _
                                                                                         Me.DataBase.sheet_Points_Coordinates, _
                                                                                         Me.F_PlanView_VisioWindow)
                                Me.Form_TreeView_MonitorPoints.ShowDialog()
                            End If
                        Else
                            MessageBox.Show("没有找到监测点位的数据信息所在的工作表！", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        End If
                    End If
                Else
                    MessageBox.Show("在Visio绘图中没有指定测点信息！", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Exit Sub
                End If
            Else        '说明Visio平面图还没有打开
                If MessageBox.Show("在程序中未检测到Visio平面图，" & vbCrLf & "是否要打开平面图", "Warning", _
                                   MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) _
                               = System.Windows.Forms.DialogResult.OK Then
                    Call Me.DrawVisioPlanView()
                End If
            End If

        End Sub

        ''' <summary>
        ''' 生成Visio的平面开挖图
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub DrawVisioPlanView()
            '判断程序中是否有Visio绘图中必须的与每一个分场形状相对应的“完成日期”的数据。
            Dim blnHasFinishedDate As Boolean = False
            If Me.DataBase IsNot Nothing Then
                If Me.DataBase.ShapeIDAndFinishedDate IsNot Nothing Then
                    APPLICATION_MAINFORM.MainForm.Form_VisioPlanView.ShowDialog()
                    blnHasFinishedDate = True
                End If
            End If
            If Not blnHasFinishedDate Then
                MessageBox.Show("程序中没有与Visio中的形状相对应的""完成日期""的数据！", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        End Sub

        ''' <summary>
        ''' 程序中所有图表的窗口的句柄值，用来对窗口进行禁用或者启用窗口
        ''' </summary>
        ''' <param name="AllDrawings"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetWindwosHandles(ByVal AllDrawings As AmeDrawings) As IntPtr()
            Dim arrHandles(0 To AllDrawings.Count - 1) As IntPtr
            Try
                Dim index As Integer = 0
                With AllDrawings
                    If .PlanView IsNot Nothing Then
                        arrHandles(index) = .PlanView.Application.Window.WindowHandle32
                        index += 1
                    End If
                    If .SectionalView IsNot Nothing Then
                        arrHandles(index) = .SectionalView.Application.Hwnd
                        index += 1
                    End If
                    For Each mnt As ClsDrawing_Mnt_Base In .MonitorData
                        arrHandles(index) = mnt.Application.Hwnd
                        index += 1
                    Next
                End With
            Catch ex As Exception
                MessageBox.Show("获取窗口句柄出错" & vbCrLf & ex.Message & vbCrLf & "报错位置：" & ex.TargetSite.Name, _
                "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
            Return arrHandles
        End Function

    End Class

    ''' <summary>
    ''' 主程序中含有Rolling方法,而且可以正确地执行滚动方法的对象
    ''' </summary>
    ''' <remarks>其中的每一个元素都是一个字典，以对象名称来索引对象
    ''' 其中有三个元素，依次代表：剖面图、平面图和监测曲线图，
    ''' 它不是指从mainform中提取出来的三个App的属性值，而是从这三个App属性中挑选出来的，正确地带有Rolling方法的相关类的实例对象。</remarks>
    Public Structure RollingEnabledDrawings
        ''' <summary>
        ''' 剖面图
        ''' </summary>
        ''' <remarks></remarks>
        Public SectionalView As ClsDrawing_ExcavationElevation
        ''' <summary>
        ''' 平面图
        ''' </summary>
        ''' <remarks></remarks>
        Public PlanView As ClsDrawing_PlanView
        ''' <summary>
        ''' 监测曲线图
        ''' </summary>
        ''' <remarks></remarks>
        Public MonitorData As List(Of clsDrawing_Mnt_RollingBase)

        ''' <summary>
        ''' 主程序中所有可以滚动的图表的数量
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Count() As UShort
            Dim btCount As Byte = 0
            '
            If Me.SectionalView IsNot Nothing Then
                btCount += 1
            End If
            '
            If Me.PlanView IsNot Nothing Then
                btCount += 1
            End If
            '
            If Me.MonitorData IsNot Nothing Then
                btCount += Me.MonitorData.Count
            End If
            '
            Return btCount
        End Function

        ''' <summary>
        ''' 构造函数
        ''' </summary>
        ''' <param name="SectionalView">剖面图</param>
        ''' <param name="PlanView">平面图</param>
        ''' <param name="MonitorData">监测曲线图</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal SectionalView As ClsDrawing_ExcavationElevation, _
                        ByVal PlanView As ClsDrawing_PlanView, _
                        ByVal MonitorData As List(Of clsDrawing_Mnt_RollingBase))
            Me.SectionalView = SectionalView
            Me.PlanView = PlanView
            Me.MonitorData = MonitorData
        End Sub

    End Structure

    ''' <summary>
    ''' 主程序中所有的绘图，包括不能进行滚动的图形。比如开挖平面图，监测曲线图，开挖剖面图。
    ''' </summary>
    ''' <remarks></remarks>
    Public Structure AmeDrawings
        ''' <summary>
        ''' 剖面图
        ''' </summary>
        ''' <remarks></remarks>
        Public SectionalView As ClsDrawing_ExcavationElevation
        ''' <summary>
        ''' 平面图
        ''' </summary>
        ''' <remarks></remarks>
        Public PlanView As ClsDrawing_PlanView
        ''' <summary>
        ''' 监测曲线图，其中的每一种监测曲线图的具体类型可以通过其Type属性来进行判断。
        ''' </summary>
        ''' <remarks></remarks>
        Public MonitorData As List(Of ClsDrawing_Mnt_Base)

        ''' <summary>
        ''' 主程序中所有图表的数量
        ''' </summary>
        ''' <returns>主程序中所有图表的数量</returns>
        ''' <remarks></remarks>
        Public Function Count() As UShort

            Dim btCount As Byte = 0
            '
            If Me.SectionalView IsNot Nothing Then
                btCount += 1
            End If
            '
            If Me.PlanView IsNot Nothing Then
                btCount += 1
            End If
            '
            If Me.MonitorData IsNot Nothing Then
                btCount += Me.MonitorData.Count
            End If
            '
            Return btCount
        End Function

        ''' <summary>
        ''' 构造函数
        ''' </summary>
        ''' <param name="SectionalView">剖面图</param>
        ''' <param name="PlanView">平面图</param>
        ''' <param name="MonitorData">监测曲线图</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal SectionalView As ClsDrawing_ExcavationElevation, _
                        ByVal PlanView As ClsDrawing_PlanView, _
                        ByVal MonitorData As List(Of ClsDrawing_Mnt_Base))
            Me.SectionalView = SectionalView
            Me.PlanView = PlanView
            Me.MonitorData = MonitorData
        End Sub

    End Structure

End Namespace
