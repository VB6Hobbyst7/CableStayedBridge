Imports System.Configuration
Imports System.Windows.Forms
Imports DAMIE.All_Drawings_In_Application
Namespace Miscellaneous
    ''' <summary>
    ''' 与程序的UI显示相关的属性
    ''' ApplicationSettingsBase类中可以保存的数据类型：
    ''' 1、基本数据类型，如integer、string、single等；
    ''' 2、 基本数据类型组成的一维数组（不能是二维或多维数组，但是可以在一维数组中嵌套一维数组，比如以行向量作为列向量的元素来构造二维数组。）；
    ''' 3、
    ''' ApplicationSettingsBase类中不能保存的数据类型：
    ''' 1、泛型Dictionary
    ''' 2、</summary>
    ''' <remarks></remarks>
    Public NotInheritable Class mySettings_UI
        Inherits System.Configuration.ApplicationSettingsBase

        ''' <summary>
        ''' 主程序界面的窗口状态
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <UserScopedSettingAttribute()> _
        Public Property WindowState As FormWindowState
            Get
                Return Me("WindowState")
            End Get
            Set(value As FormWindowState)
                Me("WindowState") = value
            End Set
        End Property

        ''' <summary>
        ''' 主程序界面的窗口位置
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <UserScopedSettingAttribute(), DefaultSettingValueAttribute("0,0")> _
        Public Property WindowLocation As Point
            Get
                Return Me("WindowLocation")
            End Get
            Set(value As Point)
                Me("WindowLocation") = value
            End Set
        End Property

        ''' <summary>
        ''' 主程序界面的窗口尺寸
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <UserScopedSettingAttribute(), DefaultSettingValueAttribute("1000, 650")> _
        Public Property WindowSize As Size
            Get
                Return Me("WindowSize")
            End Get
            Set(value As Size)
                Me("WindowSize") = value
            End Set
        End Property

    End Class

    ''' <summary>
    ''' 与程序的数据相关的属性
    ''' ApplicationSettingsBase类中可以保存的数据类型：
    ''' 1、基本数据类型，如integer、string、single等；
    ''' 2、 基本数据类型组成的一维数组（不能是二维或多维数组，但是可以在一维数组中嵌套一维数组，比如以行向量作为列向量的元素来构造二维数组。）；
    ''' 3、
    ''' ApplicationSettingsBase类中不能保存的数据类型：
    ''' 1、泛型Dictionary
    ''' 2、    ''' </summary>
    ''' <remarks></remarks>
    Public Class mySettings_Application
        Inherits System.Configuration.ApplicationSettingsBase

        ''' <summary>
        ''' 在Visio中的绘制监测点位信息时所需要的参数值
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <UserScopedSettingAttribute()> _
        Public Property MonitorPointsInfo As ClsDrawing_PlanView.MonitorPointsInformation
            Get
                Return Me("MonitorPointsInfo")
            End Get
            Set(value As ClsDrawing_PlanView.MonitorPointsInformation)
                Me("MonitorPointsInfo") = value
            End Set
        End Property

        ''' <summary>
        ''' 是否要强制将Excel单元格中的空字符串转换为Nothing。因为对于一个单元格而言，如果其中只有几个空字符，
        ''' 那么将它转换为Object时，它可以会是String类型，而不是Nothing。这样在画图时，它会以数据0.0显示。
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <UserScopedSettingAttribute(), DefaultSettingValueAttribute("True")> _
        Public Property CheckForEmpty As Boolean
            Get
                Return Me("CheckForEmpty")
            End Get
            Set(value As Boolean)
                Me("CheckForEmpty") = value
            End Set
        End Property

        ''' <summary>
        ''' 对于进行滚动的曲线进行批量处理
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <UserScopedSettingAttribute()> _
        Public Property Curve_BatchProcessing As Object()
            Get
                Return Me("Curve_BatchProcessing")
            End Get
            Set(value As Object())
                Me("Curve_BatchProcessing") = value
            End Set
        End Property

    End Class

End Namespace

