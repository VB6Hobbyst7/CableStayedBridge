Imports Microsoft.Office.Interop.Excel
Imports DAMIE.Miscellaneous
Imports DAMIE.Constants
Imports DAMIE.GlobalApp_Form

Namespace All_Drawings_In_Application
    ''' <summary>
    ''' 监测曲线图中，除了测斜曲线图外，其他的可以动态滚动的曲线图。
    ''' </summary>
    ''' <remarks></remarks>
    Public Class ClsDrawing_Mnt_OtherDynamics
        Inherits clsDrawing_Mnt_RollingBase

#Region "  ---  常数值定义"

        '图表网格与坐标值划分
        Public Const cstChartParts_Y As Byte = 10       '图表Y轴（位移）划分的区段数
        Public Const cstAxisTitle_X_Dynamic As String = Constants.AxisLabels.Points
        Public Const cstAxisTitle_Y As String = Constants.AxisLabels.Displacement_mm

#End Region

#Region "  ---  属性值的定义"

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

        Private P_dicDate_ChosenDatum As New Dictionary(Of Date, Object())
        ''' <summary>
        ''' 以每一天的日期来索引这一天的监测数据
        ''' </summary>
        ''' <value></value>
        ''' <returns>返回一个字典，其关键字为监测数据表中有数据的每一天的日期，
        ''' 对应的值为当天每一个被选择的监测点的监测数据</returns>
        ''' <remarks>监测数据只包含列表中选择了的监测点</remarks>
        Public Property DateAndDatum As Dictionary(Of Date, Object())
            Get
                Return P_dicDate_ChosenDatum
            End Get
            Set(value As Dictionary(Of Date, Object()))
                P_dicDate_ChosenDatum = value
            End Set
        End Property



#End Region

        ''' <summary>
        ''' 构造函数
        ''' </summary>
        ''' <param name="DataSheet">图表对应的数据工作表</param>
        ''' <param name="DrawingChart">Excel图形所在的Chart对象</param>
        ''' <param name="ParentApp">此图表所在的Excel类的实例对象</param>
        ''' <param name="DateSpan">此图表的TimeSpan跨度</param>
        ''' <param name="CanRoll">是图表是否可以滚动，即是动态图还是静态图</param>
        ''' <param name="Date_ChosenDatum">一个字典，其关键字为监测数据表中有数据的每一天的日期，
        ''' 对应的值为当天每一个被选择的监测点的监测数据，监测数据只包含列表中选择了的监测点</param>
        ''' <param name="Info">记录数据信息的文本框</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal DataSheet As Worksheet, ByVal DrawingChart As Chart, _
                       ByVal ParentApp As Cls_ExcelForMonitorDrawing, ByVal DateSpan As DateSpan, _
                       ByVal type As DrawingType, ByVal CanRoll As Boolean, ByVal Info As TextFrame2, _
                       ByVal DrawingTag As MonitorInfo, ByVal MonitorType As MntType, _
                       ByVal Date_ChosenDatum As Dictionary(Of Date, Object()), _
                       ByVal SeriesTag As SeriesTag)
            MyBase.New(DataSheet, DrawingChart, ParentApp, type, CanRoll, Info, DrawingTag, MonitorType, DateSpan, SeriesTag)

            '  ------------------------------------
            '为进行滚动的那条数据曲线添加数据标签
            With Me.MovingSeries
                '在数据点旁边显示数据值
                .ApplyDataLabels()
                '设置数据标签的格式
                Dim dataLBs As DataLabels = .DataLabels
                With dataLBs
                    .NumberFormat = "0.00"
                    With .Format.TextFrame2.TextRange.Font
                        .Size = 8
                        .Fill.ForeColor.RGB = RGB(0, 0, 0)
                        .Name = AMEApplication.FontName_TNR
                    End With
                End With

            End With
            P_dicDate_ChosenDatum = Date_ChosenDatum  '包括第一列，但是不包括第一行的日期。
            ' '' -----对图例进行更新---------
            'Call LegendRefresh(Me.List_HasCurve)

        End Sub

        ''' <summary>
        ''' 图形滚动
        ''' </summary>
        ''' <param name="dateThisday"></param>
        ''' <remarks></remarks>
        Public Overrides Sub Rolling(ByVal dateThisday As Date)
            MyBase.F_RollingDate = dateThisday
            Dim lockobject As New Object
            SyncLock lockobject
                Dim app = Me.Chart.Application
                app.ScreenUpdating = False

                ' ------------------- 绘制监测曲线图
                Dim Allday = Me.P_dicDate_ChosenDatum.Keys
                '考察选定的日期是否有数据
                Dim State As TodayState
                Dim closedDay As Date
                '
                If Date.Compare(dateThisday, Me.DateSpan.StartedDate) < 0 Then
                    State = TodayState.BeforeStartDay
                ElseIf Date.Compare(dateThisday, Me.DateSpan.FinishedDate) > 0 Then
                    State = TodayState.AfterFinishedDay
                ElseIf Allday.Contains(dateThisday) Then       '如果搜索的那一天有数据
                    State = TodayState.DateMatched
                    closedDay = dateThisday
                Else   '搜索的那一天没有数据，则查找选定的日期附近最近的一天，并绘制其监测曲线
                    State = TodayState.DateNotFound
                    Dim sortedlist_AllDays As New SortedSet(Of Date)(Allday)
                    closedDay = MyBase.GetClosestDay(sortedlist_AllDays, dateThisday)
                End If
                '
                Call Me.CurveRolling(dateThisday, State, closedDay)
                app.ScreenUpdating = True
            End SyncLock
        End Sub

        ''' <summary>
        ''' 绘制距选定的日期（包括选择的日期）最近的日期的监测数据
        ''' </summary>
        ''' <param name="dateThisDay">选定的施工日期</param>
        ''' <param name="State">指示选定的日期所处的状态</param>
        ''' <param name="Closestday">距离选定日期最近的日期，如果选定的日期在工作表中有数据，则等效于dateThisDay</param>
        ''' <remarks></remarks>
        Private Sub CurveRolling(ByVal dateThisDay As Date, ByVal State As TodayState, ByVal Closestday As Date)
            Dim series As Series = Me.MovingSeries ' Me.Chart.SeriesCollection(1)
            Const strDateFormat As String = AMEApplication.DateFormat
            F_DicSeries_Tag.Item(1).ConstructionDate = dateThisDay  '刷新滚动的曲线所代表的日期
            '绘制曲线图
            With series
                '.XValues = dicDate_ChosenDatum.Item(dateThisday)           'X轴的数据
                '
                Select Case State

                    Case TodayState.BeforeStartDay
                        .Values = {Nothing}     '不能设置为Series.Value=vbEmpty，因为这会将x轴标签中的某一个值设置为0.0。
                        .Name = dateThisDay.ToString(strDateFormat) & " :早于" &
                            Me.DateSpan.StartedDate.ToString(strDateFormat)

                    Case TodayState.DateNotFound
                        .Values = P_dicDate_ChosenDatum.Item(Closestday)
                        .Name = Closestday.ToString(strDateFormat) & "(" & dateThisDay.ToString(strDateFormat) & ":No Data" & ")"

                    Case TodayState.AfterFinishedDay
                        .Values = {Nothing}
                        .Name = dateThisDay.ToString(strDateFormat) & " :晚于" &
                            Me.DateSpan.FinishedDate.ToString(strDateFormat)
                    Case TodayState.DateMatched
                        .Values = P_dicDate_ChosenDatum.Item(dateThisDay)
                        .Name = dateThisDay.ToString(strDateFormat)

                End Select
            End With
        End Sub
    End Class

End Namespace