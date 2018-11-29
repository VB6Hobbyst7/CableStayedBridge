Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports DAMIE.Miscellaneous
Imports DAMIE.Constants
Imports DAMIE.DataBase
Imports DAMIE.All_Drawings_In_Application
Imports DAMIE.GlobalApp_Form
Imports DAMIE.Constants.Data_Drawing_Format

Public Class ClsDrawing_ExcavationElevation
    Implements IAllDrawings, IRolling

#Region "  ---  Declarations and Definitions"

#Region "  ---  Properties"

#Region "  ---  Excel中的对象"

    Private WithEvents P_ExcelApp As Application
    ''' <summary>
    ''' 进行绘图的Excel程序
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Application As Excel.Application
        Get
            If Me.P_ExcelApp Is Nothing Then
                Me.P_ExcelApp = New Application
            End If
            Return Me.P_ExcelApp
        End Get
        Set(value As Excel.Application)
            '不弹出警告对话框
            value.DisplayAlerts = False
            '获取Excel的进程
            Me.P_ExcelApp = value
        End Set
    End Property

    ''' <summary>
    ''' 进行绘图的工作表
    ''' </summary>
    ''' <remarks></remarks>
    Private P_Sheet_Drawing As Excel.Worksheet
    ''' <summary>
    ''' 进行绘图的工作表
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Sheet_Drawing As Excel.Worksheet
        Get
            Return Me.P_Sheet_Drawing
        End Get
    End Property

    ''' <summary>
    ''' 进行绘图的Chart对象
    ''' </summary>
    ''' <remarks></remarks>
    Private P_Chart As Excel.Chart
    ''' <summary>
    ''' 进行绘图的Chart对象
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Chart As Excel.Chart
        Get
            Return Me.P_Chart
        End Get
    End Property

    ''' <summary>
    ''' 记录数据信息的文本框
    ''' </summary>
    ''' <remarks></remarks>
    Private F_textbox_Info As TextFrame2
    ''' <summary>
    ''' 记录数据信息的文本框
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Information As TextFrame2
        Get
            Return F_textbox_Info
        End Get
    End Property

#End Region

    Private P_Type As DrawingType
    Public ReadOnly Property Type As DrawingType Implements IAllDrawings.Type
        Get
            Return Me.P_Type
        End Get
    End Property

    Private P_UniqueID As Long
    Public ReadOnly Property UniqueID As Long Implements IAllDrawings.UniqueID
        Get
            Return P_UniqueID
        End Get
    End Property

    Private P_DateSpan As DateSpan
    Public Property DateSpan As DateSpan Implements IRolling.DateSpan
        Get
            Return Me.P_DateSpan
        End Get
        Set(value As DateSpan)
            Me.P_DateSpan = value
            ' 扩展MainForm.TimeSpan的区间
            Call GlobalApplication.Application.refreshGlobalDateSpan(value)
        End Set
    End Property

#End Region

#Region "  ---  Fields"

    ''' <summary>
    ''' 此图表中选择的所有的基坑区域对象
    ''' </summary>
    ''' <remarks></remarks>
    Private F_Regions As List(Of clsData_ProcessRegionData)

    ''' <summary>
    ''' 记录每一个基坑区域的坑底标高的数据系列
    ''' </summary>
    ''' <remarks></remarks>
    Private F_Series_Static As Series

    ''' <summary>
    ''' 记录每一个基坑区域的即时开挖标高的数据系列
    ''' </summary>
    ''' <remarks></remarks>
    Private F_Series_Depth As Series

#End Region

#End Region

    ''' <summary>
    ''' 构造函数
    ''' </summary>
    ''' <param name="Series_Static">表示基坑区域的坑底深度的数据系列</param>
    ''' <param name="Series_Depth">表示基坑区域的即时开挖标高的数据系列</param>
    ''' <param name="ChosenRegion">此绘图中所包含的矩形方块与对应的数据范围</param>
    ''' <param name="textbox">记录信息的文本框</param>
    ''' <param name="type">此图表所属的类型，由枚举DrawingType提供</param>
    Public Sub New(ByVal Series_Static As Series, ByVal Series_Depth As Series, _
                ByVal ChosenRegion As List(Of clsData_ProcessRegionData), ByVal DateSpan As DateSpan, _
                ByVal textbox As Excel.TextFrame2, ByVal type As DrawingType)
        With Me
            .F_Series_Static = Series_Static
            .F_Series_Depth = Series_Depth
            .F_textbox_Info = textbox
            .P_Chart = Series_Static.Parent.Parent
            .P_Sheet_Drawing = Me.P_Chart.Parent.parent
            .Application = Me.P_Sheet_Drawing.Application
            .F_Regions = ChosenRegion
            .DateSpan = DateSpan
            .P_Type = type
            .Application.Caption = ""
            .P_UniqueID = GetUniqueID()
        End With

        ' -------------------------------------------------------------
        GlobalApplication.Application.ElevationDrawing = Me
    End Sub

    Public Sub Rolling(dateThisDay As Date) Implements IRolling.Rolling
        Dim lockobject As New Object
        SyncLock lockobject
            Dim RegionsCount As UShort = Me.F_Regions.Count
            If RegionsCount > 0 Then
                Me.Application.ScreenUpdating = False       '禁用excel界面
                Dim series_Depth As Series = Me.F_Series_Depth
                Dim Region As clsData_ProcessRegionData
                Dim Pt As Excel.Point
                Dim Depths(0 To RegionsCount - 1) As Single
                For i As UShort = 0 To RegionsCount - 1
                    Region = Me.F_Regions.Item(i)
                    Pt = series_Depth.Points.item(i + 1)
                    With Region
                        If .HasBottomDate Then
                            If dateThisDay.CompareTo(.BottomDate) > 0 Then    '说明已经开挖到基坑底，并在向上进行结构物的施工
                                With Pt

                                    .Format.Fill.ForeColor.RGB = RGB(255, 0, 0)
                                End With
                            Else                                '说明还未开挖到基坑底，并在向下开挖
                                With Pt

                                    .Format.Fill.ForeColor.RGB = RGB(0, 0, 255)
                                End With
                            End If
                        End If
                        With .Date_Elevation
                            Try
                                Depths(i) = Region.Date_Elevation.Item(dateThisDay)
                            Catch ex As KeyNotFoundException
                                Dim ClosestDate As Date = ClsData_DataBase.FindTheClosestDateInSortedList(Region.Date_Elevation.Keys, dateThisDay)
                                Depths(i) = Region.Date_Elevation.Item(ClosestDate)
                            End Try
                        End With
                    End With
                Next
                series_Depth.Values = Depths
                '刷新日期放置在最后，以免由于耗时过长而出现误判
                Me.F_textbox_Info.TextRange.Text = dateThisDay.ToString(AMEApplication.DateFormat)
            End If
            Me.Application.ScreenUpdating = True        '刷新excel界面
        End SyncLock
    End Sub

#Region "  ---  通用子方法"

    Public Sub Close(Optional SaveChanges As Boolean = False) Implements IAllDrawings.Close
        Try
            With Me
                Dim wkbk As Excel.Workbook = .P_Sheet_Drawing.Parent
                wkbk.Close(False)
                .Application.Quit()
            End With
        Catch ex As Exception
            MessageBox.Show("关闭开挖剖面图出错！" & vbCrLf & ex.Message, _
                            "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub

    ''' <summary>
    ''' 绘图界面被删除时引发的事件
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Application_Quit(ByVal Wb As Workbook, ByRef Cancel As Boolean) Handles P_ExcelApp.WorkbookBeforeClose
        GlobalApplication.Application.ElevationDrawing = Nothing
    End Sub

#End Region

End Class
