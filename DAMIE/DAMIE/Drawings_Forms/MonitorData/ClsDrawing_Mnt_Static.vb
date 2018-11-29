Imports Microsoft.Office.Interop.Excel
Imports DAMIE.Miscellaneous
Imports DAMIE.Miscellaneous.GeneralMethods
Imports DAMIE.Constants
Imports DAMIE.All_Drawings_In_Application
Imports System.Threading

Namespace All_Drawings_In_Application
    Public Class ClsDrawing_Mnt_Static
        Inherits clsDrawing_Mnt_StaticBase

#Region "  ---  Constants"
        '图表网格与坐标值划分
        Public Const cstChartParts_Y As Byte = 10       '图表Y轴（位移）划分的区段数
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
                Return New ChartSize(Data_Drawing_Format.Drawing_Mnt_Others.ChartHeight, _
                            Data_Drawing_Format.Drawing_Mnt_Others.ChartWidth, _
                            Data_Drawing_Format.Drawing_Mnt_Others.MarginOut_Height, _
                            Data_Drawing_Format.Drawing_Mnt_Others.MarginOut_Width)
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
        Sub New(ByVal DataSheet As Worksheet, ByVal DrawingChart As Chart, _
                ByVal ParentApp As Cls_ExcelForMonitorDrawing, _
                ByVal type As DrawingType, ByVal CanRoll As Boolean, _
                ByVal Info As TextFrame2, ByVal DrawingTag As MonitorInfo, ByVal MonitorType As MntType, _
                ByVal AllselectedData As Dictionary(Of Series, Object()), ByVal arrAllDate() As Double)
            Call MyBase.New(DataSheet, DrawingChart, ParentApp, type, CanRoll, Info, DrawingTag, MonitorType, arrAllDate)
            '  -----------------------------------
            '
            Me.F_dicSeries = AllselectedData
        End Sub

    End Class
End Namespace