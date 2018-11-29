Imports Microsoft.Office.Interop.Excel
Imports DAMIE.Miscellaneous
Imports DAMIE.Constants
Imports DAMIE.All_Drawings_In_Application
Imports DAMIE.GlobalApp_Form

Namespace All_Drawings_In_Application
    ''' <summary>
    ''' 在测斜数据中，某一测点在整个施工跨度内，每一天的位移最值，以及对应的深度
    ''' </summary>
    ''' <remarks></remarks>
    Public Class ClsDrawing_Mnt_MaxMinDepth
        Inherits clsDrawing_Mnt_StaticBase

#Region "  ---  Types"


        ''' <summary>
        ''' 在位移极值走势图中，所需要的所有数据
        ''' </summary>
        ''' <remarks></remarks>
        Public Class DateMaxMinDepth
            ''' <summary>
            ''' 这里的Date数据不能用Date来保存，而应该用函数ToOADate将Date转换为等效的Double类型。
            ''' 这是因为：在Excel的Chart中，如果要设置一个坐标轴标签以日期型显示，
            ''' 则不能将其以Date型数组进行赋值，而应该以对应的Double类型的数组进行赋值。
            ''' </summary>
            ''' <remarks></remarks>
            Public ConstructionDate As Double()
            Public Max As Object()
            Public Min As Object()
            Public Depth_Max As Object()
            Public Depth_Min As Object()

            ''' <summary>
            ''' 输入的数组的元素个数必须相同。元素类型必须为Object，以避免出现因为没有数据而被误写为0.0的错误。
            ''' </summary>
            ''' <param name="ConstructionDate">
            ''' 这里的Date数据不能用Date来保存，而应该用函数ToOADate将Date转换为等效的Double类型。
            ''' 这是因为：在Excel的Chart中，如果要设置一个坐标轴标签以日期型显示，
            ''' 则不能将其以Date型数组进行赋值，而应该以对应的Double类型的数组进行赋值。</param>
            ''' <param name="max"></param>
            ''' <param name="min"></param>
            ''' <param name="depth_max"></param>
            ''' <param name="depth_min"></param>
            ''' <remarks></remarks>
            Public Sub New(ConstructionDate As Double(), max As Object(), min As Object(), _
                           depth_max As Object(), depth_min As Object())
                Me.ConstructionDate = ConstructionDate
                Me.Max = max
                Me.Min = min
                Me.Depth_Max = depth_max
                Me.Depth_Min = depth_min

            End Sub
        End Class


#End Region

#Region "  ---  Properties"

        ''' <summary>
        ''' 绘图界面与画布的尺寸
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Protected Overrides Property ChartSize_sugested As ChartSize
            Get
                Return New ChartSize(Data_Drawing_Format.Drawing_Incline_DMMD.ChartHeight, _
                            Data_Drawing_Format.Drawing_Incline_DMMD.ChartWidth, _
                            Data_Drawing_Format.Drawing_Incline_DMMD.MarginOut_Height, _
                            Data_Drawing_Format.Drawing_Incline_DMMD.MarginOut_Width)
            End Get
            Set(value As ChartSize)
                Call ExcelFunction.SetLocation_Size(Me.ChartSize_sugested, Me.Chart, Me.Application)
            End Set
        End Property

#End Region

        ''' <summary>
        ''' 构造函数
        ''' </summary>
        ''' <param name="DataSheet">图表对应的数据工作表</param>
        ''' <param name="DrawingChart">Excel图形所在的Chart对象</param>
        ''' <param name="ParentApp">此图表所在的Excel类的实例对象</param>
        ''' <param name="type">此图表所属的类型，由枚举drawingtype提供</param>
        ''' <param name="CanRoll">是图表是否可以滚动，即是动态图还是静态图</param>
        ''' <param name="Info">图表中用来显示相关信息的那个文本框对象</param>
        ''' <param name="DrawingTag">每一个监测曲线图的相关信息</param>
        ''' <param name="MonitorType">监测数据的类型，比如测斜数据、立柱垂直位移数据、支撑轴力数据等</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal DataSheet As Worksheet, _
                       ByVal DrawingChart As Chart, ByVal ParentApp As Cls_ExcelForMonitorDrawing, _
                       ByVal type As DrawingType, ByVal CanRoll As Boolean, ByVal Info As TextFrame2, _
                       ByVal DrawingTag As MonitorInfo, ByVal MonitorType As MntType, ByVal AllDate As Double(), ByVal Data As DateMaxMinDepth)
            Call MyBase.New(DataSheet, DrawingChart, ParentApp, type, CanRoll, Info, DrawingTag, MonitorType, AllDate)
            '  -----------------------------------
            '
            Me.F_dicSeries = New Dictionary(Of Series, Object())
            Dim Sc As SeriesCollection = DrawingChart.SeriesCollection
            With Me.F_dicSeries
                .Add(Sc.Item(1), Data.Max)
                .Add(Sc.Item(2), Data.Min)
                .Add(Sc.Item(3), Data.Depth_Max)
                .Add(Sc.Item(4), Data.Depth_Min)
            End With
        End Sub

    End Class
End Namespace