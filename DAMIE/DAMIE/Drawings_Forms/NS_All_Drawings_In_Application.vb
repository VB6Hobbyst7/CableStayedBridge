'程序中的所有与绘图图表相关的信息
Namespace All_Drawings_In_Application

#Region "  ---  Interfaces"

    ''' <summary>
    ''' 程序中的所有类型的图表，包括剖面图、平面图以及各种监测曲线图
    ''' </summary>
    ''' <remarks></remarks>
    Public Interface IAllDrawings

        ''' <summary>
        ''' 图表的类型，根据图表类型、监测数据类型的不同，每一种不同的图表都有不同的类型
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        ReadOnly Property Type As DrawingType

        ''' <summary>
        ''' 图表的全局独立ID
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks>以当时时间的Tick属性来定义</remarks>
        ReadOnly Property UniqueID As Long

        ''' <summary>
        ''' 关闭绘图的文档以及其所在的Application程序
        ''' </summary>
        ''' <param name="SaveChanges">在关闭文档时是否保存修改的内容</param>
        ''' <remarks></remarks>
        Sub Close(Optional ByVal SaveChanges As Boolean = False)

    End Interface


    ''' <summary>
    ''' 接口，用来进行图形随着施工进度的变化而发生的相应的滚动
    ''' </summary>
    ''' <remarks></remarks>
    Public Interface IRolling
        ''' <summary>
        ''' 进行同步滚动的日期跨度，用来给出时间滚动条与日历的范围。
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Property DateSpan As DAMIE.Miscellaneous.DateSpan

        ''' <summary>
        ''' 图形随着施工进度的变化而发生的相应的滚动
        ''' </summary>
        ''' <param name="dateThisDay">进行滚动的日期</param>
        ''' <remarks></remarks>
        Sub Rolling(ByVal dateThisDay As Date)

    End Interface

#End Region

#Region "  ---  Enumerators"

    ''' <summary>
    ''' 枚举程序中的图表的类型
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum DrawingType As Byte
        ''' <summary>
        ''' 动态监测曲线图
        ''' </summary>
        ''' <remarks></remarks>
        Monitor_Dynamic
        ''' <summary>
        ''' 静态监测曲线图
        ''' </summary>
        ''' <remarks></remarks>
        Monitor_Static
        ''' <summary>
        ''' 测斜曲线图，进行动态滚动
        ''' </summary>
        ''' <remarks></remarks>
        Monitor_Incline_Dynamic
        ''' <summary>
        ''' 测斜曲线图中，用来提取其某一测点的位移最值的变化及对应的深度
        ''' </summary>
        ''' <remarks></remarks>
        Monitor_Incline_MaxMinDepth
        ''' <summary>
        ''' 开挖平面图
        ''' </summary>
        ''' <remarks></remarks>
        Vso_PlanView
        ''' <summary>
        ''' 开挖剖面图
        ''' </summary>
        ''' <remarks></remarks>
        Xls_SectionalView
    End Enum

    ''' <summary>
    ''' 监测数据的类型
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum MntType As Byte

        ''' <summary>
        ''' 一般的任意监测项目
        ''' </summary>
        ''' <remarks></remarks>
        General

        ''' <summary>
        ''' 测斜
        ''' </summary>
        ''' <remarks></remarks>
        Incline


        ''' <summary>
        ''' 立柱垂直位移
        ''' </summary>
        ''' <remarks></remarks>
        Column

        ''' <summary>
        ''' 支撑轴力
        ''' </summary>
        ''' <remarks></remarks>
        Struts

        ''' <summary>
        ''' 水位
        ''' </summary>
        ''' <remarks></remarks>
        WaterLevel

        ''' <summary>
        ''' 墙顶水平位移
        ''' </summary>
        ''' <remarks></remarks>
        WallTop_Horizontal

        ''' <summary>
        ''' 墙顶竖直位移
        ''' </summary>
        ''' <remarks></remarks>
        WallTop_Vertical

        ''' <summary>
        ''' 地表垂直位移
        ''' </summary>
        ''' <remarks></remarks>
        EarthSurface
    End Enum


    ''' <summary>
    ''' 指定的施工日期在某监测项目的日期集合中所处的状态
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum TodayState

        ''' <summary>
        ''' 指定的这一天比此项监测数据所记录的最早日期还要早
        ''' </summary>
        BeforeStartDay
        ''' <summary>
        ''' 指定的这一天比此项监测数据所记录的最晚的日期还要晚
        ''' </summary>
        AfterFinishedDay
        ''' <summary>
        ''' 指定的这一天在此项监测数据所记录的时间跨度之内，而且这一天有监测数据
        ''' </summary>
        DateMatched
        ''' <summary>
        ''' 指定的这一天在此项监测数据所记录的时间跨度之内，但是这一天没有对应的监测数据
        ''' </summary>
        DateNotFound
    End Enum

#End Region

    Public Structure ChartSize
        ''' <summary>
        ''' 画布的高度
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property ChartHeight As Integer
        ''' <summary>
        ''' 画布的宽度
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property ChartWidth As Integer
        ''' <summary>
        ''' 由画布边界扩展到Excel界面的尺寸的高度的增量
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property MarginOut_Height As Integer
        ''' <summary>
        ''' 由画布边界扩展到Excel界面的尺寸的宽度的增量
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property MarginOut_Width As Integer

        ''' <summary>
        ''' 构造函数
        ''' </summary>
        ''' <param name="ChartHeight">画布的高度</param>
        ''' <param name="ChartWidth">画布的宽度</param>
        ''' <param name="MarginOut_Height">由画布边界扩展到Excel界面的尺寸的高度的增量</param>
        ''' <param name="MarginOut_Width">由画布边界扩展到Excel界面的尺寸的宽度的增量</param>
        ''' <remarks></remarks>
        Public Sub New(ChartHeight As Integer, ChartWidth As Integer, _
                       MarginOut_Height As Integer, _
                       MarginOut_Width As Integer)
            Me.ChartHeight = ChartHeight
            Me.ChartWidth = ChartWidth
            Me.MarginOut_Height = MarginOut_Height
            Me.MarginOut_Width = MarginOut_Width
        End Sub
    End Structure

    ''' <summary>
    ''' 每一个监测曲线图的相关信息，包括测斜曲线图与其他监测曲线图（如水位，立柱，圈梁等）
    ''' </summary>
    ''' <remarks></remarks>d
    Public Structure MonitorInfo

        ''' <summary>
        ''' 监测项目名称，如“测斜”、“水位”等
        ''' </summary>
        ''' <remarks></remarks>
        Public MonitorItem As String

        ''' <summary>
        ''' 基坑编号，如“A1、B、C1、C2、D”
        ''' </summary>
        ''' <remarks></remarks>
        Public ExcavationRegion As String

        ''' <summary>
        ''' 监测点位名称，如“CX5”，这一字段只有在测斜曲线图中才有定义
        ''' </summary>
        ''' <remarks></remarks>
        Public PointName As String

        ''' <summary>
        ''' 构造函数
        ''' </summary>
        ''' <param name="MonitorItem">监测项目名称，如“测斜”、“水位”等</param>
        ''' <param name="ExcavationRegion">基坑编号，如“A1、B、C1、C2、D”</param>
        ''' <param name="PointName">监测点位名称，如“CX5”，这一字段只有在测斜曲线图中才有定义</param>
        ''' <remarks></remarks>
        Public Sub New(MonitorItem As String, ExcavationRegion As String, Optional PointName As String = Nothing)
            Me.MonitorItem = MonitorItem
            Me.ExcavationRegion = ExcavationRegion
            Me.PointName = PointName
        End Sub

    End Structure

End Namespace