Imports System.ComponentModel
Imports DAMIE.DataBase
Imports DAMIE.GlobalApp_Form
Imports DAMIE.Miscellaneous
Imports Microsoft.Office.Interop.Visio

Namespace All_Drawings_In_Application

    ''' <summary>
    ''' Visio的开挖平面图
    ''' </summary>
    ''' <remarks></remarks>
    Public Class ClsDrawing_PlanView
        Implements IAllDrawings, IRolling

#Region "  ---  Declarations & Definitions"

#Region "  ---  Types"

        ''' <summary>
        ''' Visio平面图中，与监测点位相关的信息（不是开挖平面），用来在Visio平面图中绘制测点。
        ''' </summary>
        ''' <remarks></remarks>
        Public Structure MonitorPointsInformation

            ' ------ 不删！！！ ------ 不删！！！ ------ 不删！！！ ------
            ' ''' <summary>
            ' ''' visio中包含两个定位点的组合形状的形状名
            ' ''' </summary>
            ' ''' <remarks></remarks>
            'Const cstShapeName_OverView As String = "OVERVIEW"
            ' ''' <summary>
            ' ''' visio中在监测点的主控形状中，用来显示测点编号的形状的Name属性。
            ' ''' </summary>
            ' ''' <remarks></remarks>
            'Const cstShapeName_MonitorPointTag As String = "Tag"
            ' ''' <summary>
            ' ''' Visio平面图中用于坐标变换的两个定位点的形状ID，
            ' ''' 这两个点分别代表ABCD基坑群的左下角与右上角。
            ' ''' </summary>
            ' ''' <remarks></remarks>
            'Const cstShapeName_ConversionPoint1 As String = "Location Reference 1"
            'Const cstShapeName_ConversionPoint2 As String = "Location Reference 2"
            ''CAD平面图中用于坐标变换的两个定位点的坐标，这两个点分别代表ABCD基坑群的左下角与右上角。
            ''Const cstCADLocation_Point1 As New PointF(309598.527, -119668.436)
            ''Const cstCADLocation_Point2 As New PointF(536642.644, 201852.14)
            'Const cstCADLocation_x_Point1 As Single = 309598.527
            'Const cstCADLocation_y_Point1 As Single = -119668.436
            'Const cstCADLocation_x_Point2 As Single = 536642.644
            'Const cstCADLocation_y_Point2 As Single = 201852.14
            '

            ''' <summary>
            ''' visio中在监测点的主控形状中，用来显示测点编号的形状的Name属性。
            ''' </summary>
            ''' <remarks>在表示监测点位的主控形状的模板中，每一个监测点位的主控形状中，
            ''' 都有一个子形状，其Name属性为Tag，以此来索引此文本形状。</remarks>
            Public ShapeName_MonitorPointTag As String
            ''' <summary>
            ''' Visio平面图中用于坐标变换的两个定位点的形状ID，
            ''' 这两个点分别代表ABCD基坑群的左下角与右上角。
            ''' </summary>
            ''' <remarks></remarks>
            Public pt_Visio_BottomLeft_ShapeID As Integer
            ''' <summary>
            ''' Visio平面图中用于坐标变换的两个定位点的形状ID，
            ''' 这两个点分别代表ABCD基坑群的左下角与右上角。
            ''' </summary>
            ''' <remarks></remarks>
            Public pt_Visio_UpRight_ShapeID As Integer
            ''' <summary>
            ''' CAD平面图中用于坐标变换的两个定位点的坐标，这两个点分别代表ABCD基坑群的左下角与右上角。
            ''' </summary>
            ''' <remarks></remarks>
            Public pt_CAD_BottomLeft As PointF
            ''' <summary>
            ''' CAD平面图中用于坐标变换的两个定位点的坐标，这两个点分别代表ABCD基坑群的左下角与右上角。
            ''' </summary>
            ''' <remarks></remarks>
            Public pt_CAD_UpRight As PointF
            '"OVERVIEW","Tag",197,217,New PointF(309598.527, -119668.436),New PointF(536642.644, 201852.14)"

        End Structure

#End Region

#Region "  ---  Constants"

        Const cstDrawingTag As String = "平面开挖分区"
#End Region

#Region "  ---  Events"

        ''' <summary>
        ''' 鼠标在Visio绘图的窗口中双击
        ''' </summary>
        ''' <remarks></remarks>
        Private Event WindowDoubleClick As Action

#End Region

#Region "  --- Properties"

        ''' <summary>
        ''' Visio程序界面
        ''' </summary>
        ''' <remarks></remarks>
        Private WithEvents P_Application As Application
        ''' <summary>
        ''' Visio程序界面
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Application As Application
            Get
                Return P_Application
            End Get
            Private Set(value As Application)
                '在关闭Visio程序时不弹出保存文档的对话框
                value.AlertResponse = 7     '0表示弹出任何警告或模型对话框，7表示不弹出对话框而默认选择IDNo。
                P_Application = value
            End Set
        End Property

        ''' <summary>
        ''' 进行绘画的那一个页面
        ''' </summary>
        ''' <remarks></remarks>
        Private P_Page As Page
        ''' <summary>
        ''' 进行绘画的那一个页面
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Page As Page
            Get
                Return P_Page
            End Get
            Set(value As Page)
                P_Page = value
            End Set
        End Property

        ''' <summary>
        ''' Visio绘图中的窗口，此窗口并不限制于页面所在的范围，而是可以扩展到页面范围以外。
        ''' </summary>
        ''' <remarks></remarks>
        Private WithEvents P_Window As Window
        ''' <summary>
        ''' Visio绘图中的窗口，此窗口并不限制于页面所在的范围，而是可以扩展到页面范围以外。
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Window As Window
            Get
                Return Me.P_Window
            End Get
            Set(value As Window)
                Me.P_Window = value
            End Set
        End Property

        ''' <summary>
        ''' 工作表“开挖分块”中的“形状名”与“完成日期”两列的数据，
        ''' 其中分别以向量的形式记录了这两列数据。向量的第一个元素的下标值为0
        ''' </summary>
        ''' <remarks></remarks>
        Private P_ShapeID_FinishedDate As clsData_ShapeID_FinishedDate
        ''' <summary>
        ''' 工作表“开挖分块”中的“形状名”与“完成日期”两列的数据，
        ''' 其中分别以向量的形式记录了这两列数据。向量的第一个元素的下标值为0
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property ShapeID_FinishedDate As clsData_ShapeID_FinishedDate
            Get
                Return P_ShapeID_FinishedDate
            End Get
            Set(value As clsData_ShapeID_FinishedDate)
                P_ShapeID_FinishedDate = value
            End Set
        End Property

        Private P_DateSpan As DateSpan
        ''' <summary>
        ''' 进行同步滚动的时间跨度，用来给出时间滚动条与日历的范围。
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property DateSpan As DateSpan Implements IRolling.DateSpan
            Get
                Return Me.P_DateSpan
            End Get
            Private Set(value As DateSpan)
                ' 扩展MainForm.TimeSpan的区间
                Call GlobalApplication.Application.refreshGlobalDateSpan(value)
                Me.P_DateSpan = value
            End Set
        End Property

#Region "  ---  开挖平面图的标签信息"

        Private P_DrawingType As DrawingType
        ''' <summary>
        ''' 此图表所属的类型，由枚举DrawingType提供
        ''' </summary>
        ''' <value></value>
        ''' <returns>DrawingType枚举类型，指示此图形所属的类型</returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Type As DrawingType Implements IAllDrawings.Type
            Get
                Return P_DrawingType
            End Get
        End Property

        ''' <summary>
        ''' 此绘图画面的标签，用来进行索引
        ''' </summary>
        ''' <remarks></remarks>
        Private P_Name As String
        ''' <summary>
        ''' 此绘图画面的标签，用来进行索引
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Name As String
            Get
                Return Me.P_Name
            End Get
        End Property

        Private P_UniqueID As Long
        ''' <summary>
        ''' 图表的全局独立ID
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks>以当时时间的Tick属性来定义</remarks>
        Public ReadOnly Property UniqueID As Long Implements IAllDrawings.UniqueID
            Get
                Return P_UniqueID
            End Get
        End Property

#End Region

#Region "  ---  在开挖平面图中绘制测点"
        ''' <summary>
        ''' 此Visio平面图中是否可以用来绘制监测点位的信息
        ''' </summary>
        ''' <remarks></remarks>
        Private P_HasMonitorPointsInfo As Boolean = False
        ''' <summary>
        ''' 此Visio平面图中是否可以用来绘制监测点位的信息
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property HasMonitorPointsInfo As Boolean
            Get
                Return Me.P_HasMonitorPointsInfo
            End Get
            Set(value As Boolean)
                Me.P_HasMonitorPointsInfo = value
            End Set
        End Property

        ''' <summary>
        ''' Visio平面图中，与监测点位相关的信息（不是开挖平面），用来在Visio平面图中绘制测点。
        ''' </summary>
        ''' <remarks></remarks>
        Private P_MonitorPointsInfo As MonitorPointsInformation
        ''' <summary>
        ''' Visio平面图中，与监测点位相关的信息（不是开挖平面），用来在Visio平面图中绘制测点。
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property MonitorPointsInfo As MonitorPointsInformation
            Get
                Return Me.P_MonitorPointsInfo
            End Get
        End Property

#End Region

#End Region

#Region "  ---  Fields"

        ''' <summary>
        ''' 后台工作线程
        ''' </summary>
        ''' <remarks></remarks>
        Private WithEvents F_BackgroundWorker As New BackgroundWorker

        ''' <summary>
        ''' 开挖平面图在Visio中的页面名称
        ''' </summary>
        ''' <remarks></remarks>
        Private F_PageName_PlanView As String '= "开挖平面"
        ''' <summary>
        ''' 所有分区的组合形状的ID值
        ''' </summary>
        ''' <remarks></remarks>
        Private F_ShapeID_AllRegions As Integer '= 5078
        ''' <summary>
        ''' 记录开挖信息的文本框
        ''' </summary>
        ''' <remarks></remarks>
        Private F_InfoBoxID As Integer '= 5079
        '
        ''' <summary>
        ''' 绘图页面中的“所有形状”的集合
        ''' </summary>
        ''' <remarks></remarks>
        Private F_shps As Shapes
        ''' <summary>
        ''' 图层：“显示”
        ''' </summary>
        ''' <remarks></remarks>
        Dim F_layer_Show As Layer
        ''' <summary>
        ''' 图层：“显示”
        ''' </summary>
        ''' <remarks></remarks>
        Dim F_layer_Hide As Layer
#End Region

#End Region

        ''' <summary>
        ''' 构造函数
        ''' </summary>
        ''' <param name="strFilePath">要打开的Visio文档的绝对路径</param>
        ''' <param name="type">Visio平面图所属的绘图类型</param>
        ''' <param name="PageName_PlanView">开挖平面图在Visio中的页面名称</param>
        ''' <param name="ShapeID_AllRegions">所有分区的组合形状的ID值</param>
        ''' <param name="InfoBoxID">记录开挖信息的文本框</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal strFilePath As String, ByVal type As DrawingType,
                        ByVal PageName_PlanView As String,
                        ByVal ShapeID_AllRegions As Integer, ByVal InfoBoxID As Integer,
                        ByVal HasMonitorPointsInfo As Boolean,
                        ByVal MonitorPointsInfo As MonitorPointsInformation)
            '开挖平面图信息
            Me.F_PageName_PlanView = PageName_PlanView
            Me.F_ShapeID_AllRegions = ShapeID_AllRegions
            Me.F_InfoBoxID = InfoBoxID
            '监测点位信息
            Me.P_HasMonitorPointsInfo = HasMonitorPointsInfo
            Me.P_MonitorPointsInfo = MonitorPointsInfo
            '
            Me.P_ShapeID_FinishedDate = GlobalApplication.Application.DataBase.ShapeIDAndFinishedDate

            With Me.F_BackgroundWorker
                .WorkerReportsProgress = True
                .WorkerSupportsCancellation = True
                '开始绘图
                If Not .IsBusy Then
                    '在工作线程中执行绘图操作
                    Call .RunWorkerAsync({strFilePath, type})
                End If
            End With

        End Sub

#Region "  ---  打开Visio平面图"

        Private Sub F_BackgroundWorker_DoWork(sender As Object, e As DoWorkEventArgs) Handles F_BackgroundWorker.DoWork
            '主程序界面的进度条的UI显示     
            APPLICATION_MAINFORM.MainForm.ShowProgressBar_Marquee()
            '
            Dim strFilePath As String = e.Argument(0)
            Dim type As DrawingType = e.Argument(1)
            '执行具体的绘图操作
            Try
                Me.Application = NewVsoApp()
            Catch ex As Exception
                Debug.Print(ex.Message)
                e.Cancel = True
                Exit Sub
            End Try


            '
            Try
                Dim vsoDoc As Document = OpenDocument(Me.P_Application, strFilePath)
                If vsoDoc Is Nothing Then
                    e.Cancel = True
                    Exit Sub
                End If
                '将指定page（开挖平面）的窗口指定给对应变量WndDrawing
                '在NewApplication方法中，已经将指定page对象的窗口进行了激活。
                Me.P_Page = vsoDoc.Pages.Item(F_PageName_PlanView)
                '------------设置visioWindow的TimeSpan区间
                '！这里必须将数组进行Clone！以复制一个新副本，不然后面的Sort方法就会改变MainForm.DataBase中的数组的排序。
                Dim TimeRange() As Date = P_ShapeID_FinishedDate.FinishedDate.Clone
                Me.DateSpan = Me.GetAndRefreshDateSpan(TimeRange)
                '设置此图形所属的类型
                Me.P_DrawingType = type
                '此图形的独立ID
                Me.P_UniqueID = GetUniqueID()
                Me.P_Name = cstDrawingTag
                '
                ' ----------------- 创建索引
                Call GenerateReference(Me.P_Page, F_layer_Show, F_layer_Hide)
                '---------------- 设置形状的显示与隐藏
                F_shps = Me.P_Page.Shapes
                Dim shpAllRegion As Shape
                shpAllRegion = F_shps.ItemFromID(F_ShapeID_AllRegions)
                Call SetVisible(Me.P_Page, shpAllRegion, F_layer_Show, F_layer_Hide)
            Catch ex As Exception
                Debug.Print("创建新的Visio对象出错！")
            Finally
                '显示窗口与刷新屏幕
                If Me.P_Application IsNot Nothing Then
                    With Me.P_Application
                        .Visible = True
                        .ShowChanges = True
                    End With
                End If
            End Try
        End Sub

        ''' <summary>
        ''' 创建一个新的Visio.Application对象
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function NewVsoApp() As Application
            If Me.Application IsNot Nothing Then
                Return Me.P_Application
            Else

                '创建新Visio窗口并为其命名
                Dim vsoApp = New Application()

                ' Dim vsoApp  = CreateObject("visio.application")
                '但是，这两种方法都有一个问题，就是在刚执行完这一句后，Visio程序的界面就出现了，
                '即使后面紧跟着vsoApp.Visible = False，在UI显示上还是会出现一个界面由出现到隐藏的闪动。
                vsoApp.Visible = False
                '
                Return vsoApp
            End If
        End Function

        ''' <summary>
        ''' 打开一个Visio文档
        ''' </summary>
        ''' <param name="vsoApp"></param>
        ''' <param name="FilePath">文档的绝对路径</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function OpenDocument(ByVal vsoApp As Application, ByVal FilePath As String) As Document
            Dim vsoDoc As Document
            Try
                '有可能出现此visio文档已经在系统中打开的情况
                vsoDoc = vsoApp.Documents.Open(FilePath)
                'Visio界面美化
                Call VsoApplicationBeauty(vsoApp, vsoDoc)
            Catch ex As Exception
                vsoApp.Quit()
                Me.P_Application = Nothing
                MessageBox.Show("选择的文件已经打开，请将其手动关闭并重新打开。", "Warning", MessageBoxButtons.OK,
                                MessageBoxIcon.Warning)
                Return Nothing
            End Try
            Return vsoDoc
        End Function

        ''' <summary>
        ''' 获取此文档所对应的施工日期跨度
        ''' </summary>
        ''' <remarks></remarks>
        Private Function GetAndRefreshDateSpan(ByVal TimeRange() As Date) As DateSpan
            Array.Sort(TimeRange)
            Dim DS As DateSpan
            DS.StartedDate = TimeRange(0)
            DS.FinishedDate = TimeRange(UBound(TimeRange))
            Return DS
        End Function

        ''' <summary>
        ''' ！创建索引，将变量索引到Visio中的形状集合与相应图层。
        ''' </summary>
        ''' <param name="DrawingPage"></param>
        ''' <param name="layer_Show">页面中用来放置“显示”的对象的图层</param>
        ''' <param name="Layer_Hide">页面中用来放置“隐藏”的对象的图层</param>
        ''' <remarks></remarks>
        Private Sub GenerateReference(ByVal DrawingPage As Page, ByRef layer_Show As Layer, ByRef Layer_Hide As Layer)
            '创建“显示”与“隐藏”图层的索引
            '如果文档中已经有相应的图层，则直接索引，如果文档中还没有这两个图层，则添加这两个图层
            Const cstLayerName_Visible As String = "Show"
            Const cstLayerName_InVisible As String = "Hide"
            With DrawingPage.Layers
                '当要添加的图层已经存在时，会直接返回那个图层，并且保留那个图层所有的设置信息。
                layer_Show = .Add(cstLayerName_Visible)
                Layer_Hide = .Add(cstLayerName_InVisible)
            End With

            '设置“隐藏”与“显示”图层的可见性
            layer_Show.CellsC(VisCellIndices.visLayerVisible).ResultIU = 1
            Layer_Hide.CellsC(VisCellIndices.visLayerVisible).ResultIU = 0
        End Sub

        ''' <summary>
        ''' 设置文档中开挖分块形状的显示与隐藏
        ''' </summary>
        ''' <param name="DrawingPage"></param>
        ''' <param name="shpAllRegion">所有开挖分块形状所在的组合形状</param>
        ''' <param name="layer_Show"></param>
        ''' <param name="Layer_Hide"></param>
        ''' <remarks>
        ''' 对于文档中的形状的显示与隐藏的处理方法:
        ''' 基本原理：如果一个形状同时位于多个图层中，只要这些图层中有一个图层可见，则此形状可见。
        ''' 1、先将所有的分块形状都从所有的图层中删除，然后将它们添加到“Show”与“Hide”这两个图层中。
        ''' 2、在后续的操作中，如果将形状从“Show”中移除，则此形状会被隐藏，如果将形状添加进“Show”中，则此形状会显示出来；</remarks>
        Private Sub SetVisible(ByVal DrawingPage As Page, ByVal shpAllRegion As Shape, ByVal layer_Show As Layer, ByVal Layer_Hide As Layer)

            With DrawingPage
                For Each l As Layer In .Layers
                    '如果要移除的形状对象本来就不在该图层中，则此代码不生效，也不会报错。
                    l.Remove(shpAllRegion, 0)
                Next
            End With
            Layer_Hide.Add(shpAllRegion, 0)
            'Layer.Add说明：如果形状为组并且 fPresMems 不为零，则该组的组件形状保留它们当前的图层分配，并且将组件形状添加到该图层。
            '如果 fPresMems 为零，则组件形状将重新分配到该图层，并且丢失它们当前的图层分配。
            layer_Show.Add(shpAllRegion, 1)
            ' --------------------------------------------------------
        End Sub

        ''' <summary>
        ''' 进行Visio界面的美化
        ''' </summary>
        ''' <param name="VsoApp"></param>
        ''' <param name="vsoDoc">进行美化时,要先打开Visio文档,如果没有打开任何文档,
        ''' 则Application的ActiveDocument与ActiveWindow属性会返回Nothing.</param>
        ''' <remarks></remarks>
        Private Sub VsoApplicationBeauty(ByVal VsoApp As Application, ByVal vsoDoc As Document)

            With VsoApp
                .Visible = False
                .ShowChanges = False

                '激活绘图页面
                Dim ActivePage As Page
                If Me.F_PageName_PlanView IsNot Nothing Then
                    '如果visio中还没有打开任何文档，则ActiveDocument返回nothing
                    ActivePage = vsoDoc.Pages.Item(F_PageName_PlanView)
                Else
                    ActivePage = vsoDoc.Pages.Item(1)
                End If

                '如果visio中还没有打开任何文档，则ActiveWindow返回nothing
                .ActiveWindow.Page = ActivePage
                Me.P_Window = .ActiveWindow
                With .ActiveWindow
                    '窗口最大化
                    .WindowState = VisWindowStates.visWSMaximized

                    .ShowScrollBars = False
                    .ShowRulers = False
                    .ShowPageTabs = False
                    '禁止鼠标滚动，但是可以用鼠标滚动+Ctrl进行缩放。
                    '.ScrollLock = True
                    '让窗口的显示适应页面
                    .ViewFit = VisWindowFit.visFitPage
                    '打开“形状”窗格，以供后面用DoCmd进行切换时将其隐藏。
                    .Windows.ItemFromID(VisWinTypes.visWinIDShapeSearch).Visible = True
                End With

                '隐藏状态栏与工具栏
                .ShowStatusBar = False
                .ShowToolbar = False

                '对于“形状”窗格进行特殊处理——切换显示
                .DoCmd(VisUICmds.visCmdShapesWindow)

                '关闭所有的子窗口
                For Each wnd As Window In P_Application.ActiveWindow.Windows
                    wnd.Visible = False
                Next
            End With
        End Sub

        Private Sub F_BackgroundWorker_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles F_BackgroundWorker.RunWorkerCompleted

            If e.Cancelled OrElse Me.Application Is Nothing Then   '说明Visio文档打开异常
                '在绘图完成后，隐藏进度
                With APPLICATION_MAINFORM.MainForm
                    .HideProgress("Visio文档打开异常。")
                    '启动同步滚动按钮
                    .MainUI_RollingObjectCreated()
                End With

                '---将对象传递给mainform的属性中
                With GlobalApplication.Application
                    .PlanView_VisioWindow = Nothing
                    '标记：程序中已经打开的Visio平面图
                    ' .HasVisioPlanView = False
                End With
            Else                  '正常地打开了Visio平面图        
                '在绘图完成后，隐藏进度条
                APPLICATION_MAINFORM.MainForm.HideProgress("Done")

                '---将对象传递给mainform的属性中
                GlobalApplication.Application.PlanView_VisioWindow = Me

            End If
        End Sub
#End Region

#Region "  ---  图形滚动"

        ''' <summary>
        ''' 图形滚动
        ''' </summary>
        ''' <param name="dateThisDay"></param>
        ''' <remarks>在Visio中的数据记录集上，用来控制形状的显示与否的，就只有“形状ID”与“完成日期”这两列的数据。
        ''' Layer.Add说明：如果形状为组并且 fPresMems 不为零，则该组合形状的子形状保留它们当前的图层分配，并且将子形状也添加到该图层。
        '''               如果 fPresMems 为零，则子形状将重新分配到该图层，并且丢失它们状原来的图层分配信息，即此时子形状与组合形状的在的图层完全相同。
        ''' Layer.Remove说明：如果形状为一个组合，而 fPresMems 为非零值，该组合的成员形状将不会受到影响。
        '''                  如果 fPresMems 为零 (0)，该组合的成员形状也将从图层中移除。</remarks>
        Public Sub Rolling(ByVal dateThisDay As Date) Implements IRolling.Rolling

            '工作表“开挖分块”中的“形状ID”与“完成日期”两列的数据，其中分别以向量的形式记录了这两列数据。
            '向量的第一个元素的下标值为0()
            Dim arrShapeID() As Integer = P_ShapeID_FinishedDate.ShapeID
            Dim arrFinishedDate() As Date = P_ShapeID_FinishedDate.FinishedDate
            '
            Dim shp As Shape
            Dim FinishedDate As Date
            Dim CompareDate As Integer
            SyncLock Me
                P_Application.ShowChanges = False
                P_Application.ScreenUpdating = False
                For IRow As Integer = LBound(arrShapeID) To UBound(arrShapeID)
                    '形状的索引是从
                    shp = F_shps.ItemFromID(arrShapeID(IRow))
                    FinishedDate = arrFinishedDate(IRow)
                    CompareDate = Date.Compare(dateThisDay, FinishedDate)
                    If CompareDate >= 0 Then          '说明今天已经过了这个基坑区域完成的日期，所以应该将这个区域显示出来
                        Me.F_layer_Show.Add(shp, 0)
                    Else                              '说明今天还没挖到相应的区域，所以应该将这个区域隐藏起来
                        'F_layer_Show.Add(shp, 1)   '将组合形状本身添加进图层，而不添加其子形状。
                        F_layer_Show.Remove(shp, 0)
                    End If
                Next

                '在记录信息的文本框中显示相关信息
                F_shps.ItemFromID(F_InfoBoxID).Characters.Text = "施工日期：" & dateThisDay.ToString("yyyy/MM/dd")

                '刷新屏幕
                P_Application.ShowChanges = True
            End SyncLock
        End Sub

#End Region

        ''' <summary>
        ''' 绘图界面被删除时引发的事件
        ''' </summary>
        ''' <param name="app"></param>
        ''' <remarks>此事件会在Visio弹出是否要保存文档的对话框之后才被触发，所以不能在此事件中去设置app.AlertResponse。</remarks>
        Private Sub Application_Quit(ByVal app As Application) Handles P_Application.BeforeQuit
            For Each doc As Document In app.Documents
                doc.Close()
            Next
            '
            Me.P_Application = Nothing
            '
            With GlobalApplication.Application
                .PlanView_VisioWindow = Nothing
                '.HasVisioPlanView = False
            End With
        End Sub

#Region "  ---  窗口双击"

        Private DoubleClick_StartTime As DateTime
        Private DoubleClick_FirstClickInitialized As Boolean

        Private Sub DoubleClick_Up(Button As Integer, KeyButtonState As Integer,
                                   x As Double, y As Double, ByRef CancelDefault As Boolean) Handles P_Window.MouseUp
            If Button = 1 Then      'Left mouse button released
                If DoubleClick_FirstClickInitialized Then
                    Dim DoubleClickInterval As Integer = 300    '定义双击的时间间隔
                    Dim Interval As Integer = Now.Subtract(DoubleClick_StartTime).Milliseconds
                    If Interval < DoubleClickInterval Then
                        RaiseEvent WindowDoubleClick()
                        Debug.Print("双击成功！")
                        DoubleClick_FirstClickInitialized = False
                    Else
                        DoubleClick_StartTime = Now
                        DoubleClick_FirstClickInitialized = True
                    End If
                Else
                    DoubleClick_StartTime = Now
                    DoubleClick_FirstClickInitialized = True
                End If
            End If
        End Sub

        ''' <summary>
        ''' 在窗口中双击时让窗口适应页面，即缩放页面使之在窗口中完全显示。
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub ViewFit() Handles Me.WindowDoubleClick
            Me.P_Window.ViewFit = VisWindowFit.visFitPage
        End Sub

#End Region

        ''' <summary>
        ''' 关闭绘图的Visio文档以及其所在的Application程序
        ''' </summary>
        ''' <param name="SaveChanges">在关闭文档时是否保存修改的内容</param>
        ''' <remarks></remarks>
        Public Sub Close(Optional ByVal SaveChanges As Boolean = False) Implements IAllDrawings.Close
            Try
                With Me
                    Dim vsoDoc As Document = .Page.Document
                    vsoDoc.Close()
                    '此时开挖平面图已经关闭
                    .Application.Quit()
                End With
            Catch ex As Exception
                MessageBox.Show("关闭Visio开挖平面图出错！" & vbCrLf & ex.Message,
                          "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End Try
        End Sub

    End Class
End Namespace