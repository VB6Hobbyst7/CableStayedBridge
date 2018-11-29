Imports DAMIE.Constants
Imports DAMIE.DataBase
Imports DAMIE.Miscellaneous
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Excel

Namespace All_Drawings_In_Application
    ''' <summary>
    ''' 测斜数据的曲线图，此
    ''' </summary>
    ''' <remarks></remarks>
    Public Class ClsDrawing_Mnt_Incline
        Inherits clsDrawing_Mnt_RollingBase

#Region "  ---  Types"

        ''' <summary>
        ''' 数据列series对象所对应的一些信息对象
        ''' </summary>
        ''' <remarks>它包括每一条数据系列的对象，及其对应的日期、开挖深度的标志线、记录开挖深度值的文本框</remarks>
        Public Class SeriesTag_Incline
            Inherits SeriesTag

            ''' <summary>
            ''' 与数据系列相关联的挖深直线
            ''' </summary>
            Private P_DepthLine As Microsoft.Office.Interop.Excel.Shape
            ''' <summary>
            ''' 与数据系列相关联的挖深直线
            ''' </summary>
            ''' <value></value>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Property DepthLine As Microsoft.Office.Interop.Excel.Shape
                Get
                    Return Me.P_DepthLine
                End Get
                Set(value As Microsoft.Office.Interop.Excel.Shape)
                    Me.P_DepthLine = value
                End Set
            End Property

            ''' <summary>
            ''' 与数据系列相关联的文本框
            ''' </summary>
            ''' 开挖深度信息的文本框
            Private P_DepthText As Microsoft.Office.Interop.Excel.Shape
            ''' <summary>
            ''' 与数据系列相关联的文本框。设置文本框中的文字： DepthTextbox.TextFrame2.TextRange.Text
            ''' </summary>
            ''' <value></value>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Property DepthTextbox As Microsoft.Office.Interop.Excel.Shape
                Get
                    Return Me.P_DepthText
                End Get
                Set(value As Microsoft.Office.Interop.Excel.Shape)
                    Me.P_DepthText = value
                End Set
            End Property

            ''' <summary>
            ''' 构造函数
            ''' </summary>
            ''' <param name="Series"></param>
            ''' <param name="ConstructionDate"></param>
            ''' <param name="DepthLine">与数据系列相关联的挖深直线</param>
            ''' <param name="DepthText">与数据系列相关联的文本框</param>
            ''' <remarks></remarks>
            Public Sub New(ByVal Series As Series, ByVal ConstructionDate As Date,
                         Optional ByVal DepthLine As Microsoft.Office.Interop.Excel.Shape = Nothing,
                         Optional ByVal DepthText As Microsoft.Office.Interop.Excel.TextFrame2 = Nothing)
                Call MyBase.New(Series, ConstructionDate)
                Me.P_DepthLine = DepthLine
                Me.P_DepthText = DepthText
            End Sub

        End Class

#End Region

#Region " --- Properties"

        ''' <summary>
        ''' 绘图界面与画布的尺寸
        ''' </summary>
        ''' <value></value>
        ''' <returns>此数组有四个元素，分别代表：画布的高度、宽度；由画布边界扩展到Excel界面的尺寸的高度和宽度的增量</returns>
        ''' <remarks></remarks>
        Protected Overrides Property ChartSize_sugested As ChartSize
            Get
                Return New ChartSize(Data_Drawing_Format.Drawing_Incline.ChartHeight,
                                Data_Drawing_Format.Drawing_Incline.ChartWidth,
                                Data_Drawing_Format.Drawing_Incline.MarginOut_Height,
                                Data_Drawing_Format.Drawing_Incline.MarginOut_Width)
            End Get
            Set(value As ChartSize)
                Call ExcelFunction.SetLocation_Size(Me.ChartSize_sugested, Me.Chart, Me.Application)
            End Set
        End Property

        ''' <summary>
        ''' 重写基类的属性：图例框的尺寸与位置
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Protected Overrides ReadOnly Property Legend_Location As LegendLocation
            Get
                Return New LegendLocation(LegendHeight:=Data_Drawing_Format.Drawing_Incline.Legend_Height,
                                          LegendWidth:=Data_Drawing_Format.Drawing_Incline.Legend_Width)
            End Get
        End Property

        ''' <summary>
        ''' 监测曲线的数据范围返回的Range对象中，包括此工作表的UsedRange中的第一列，
        ''' 即深度的数据；但是不包括第一行的日期数据
        ''' </summary>
        ''' <remarks></remarks>
        Private P_rgMntData As Range

        ''' <summary>
        ''' 在施工进度工作表中，每一个基坑区域相关的各种信息，比如区域名称，区域的描述，区域数据的Range对象，区域所属的基坑ID及其ID的数据等
        ''' </summary>
        ''' <remarks></remarks>
        Private P_ProcessRegionData As clsData_ProcessRegionData
        ''' <summary>
        ''' 在施工进度工作表中，每一个基坑区域相关的各种信息，比如区域名称，区域的描述，区域数据的Range对象，区域所属的基坑ID及其ID的数据等
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property ProcessRegionData As clsData_ProcessRegionData
            Get
                Return Me.P_ProcessRegionData
            End Get
        End Property

        ''' <summary>
        ''' 指示是否要在进行滚动时指示开挖标高的标识线旁给出文字说明，比如“开挖标高”等。
        ''' </summary>
        Public ReadOnly Property ShowLabelsWhileRolling As Boolean
            Get
                Return False
            End Get
        End Property

        ''' <summary> 测斜管的顶部的绝对标高。在测斜管的监测数据中，深度值是相对于测斜管顶部的深度，
        ''' 而在监测曲线图中绘制开挖深度或者构件深度时，其中的深度是按绝对标高给出来的，所以需要此属性的值来进行二者之间的转换。
        ''' </summary>
        Private Property _inclineTopElevaion As Single
        ''' <summary> 测斜管的顶部的绝对标高。在测斜管的监测数据中，深度值是相对于测斜管顶部的深度，
        ''' 而在监测曲线图中绘制开挖深度或者构件深度时，其中的深度是按绝对标高给出来的，所以需要此属性的值来进行二者之间的转换。
        ''' </summary>
        Public ReadOnly Property InclineTopElevaion As Single
            Get
                Return _inclineTopElevaion
            End Get
        End Property

#End Region

#Region "  ---  Fields"

        ''' <summary>
        ''' 此测斜点的监测数据工作表中的每一天与其在工作表中对应的列号
        ''' </summary>
        ''' <remarks>返回一个字典，以监测数据的日期key索引数据所在列号item</remarks>
        Private F_dicDateAndColumnNumber As Dictionary(Of Date, Integer)

        ''' <summary>
        ''' 记录开挖深度的直线与文本框
        ''' </summary>
        ''' <remarks>返回一个数组，数组中有两个元素，第一个为开挖深度的直线；
        ''' 第二个为写有相应文字的文本框</remarks>
        Private _ExcavationDepth_lineAndTextbox As SeriesTag_Incline

        Private F_YValues As Single()

#End Region

        ''' <summary>
        ''' 构造函数
        ''' </summary>
        ''' <param name="DataSheet">图表对应的数据工作表</param>
        ''' <param name="DrawingChart">Excel图形所在的Chart对象</param>
        ''' <param name="ParentApp">此图表所在的Excel类的实例对象</param>
        ''' <param name="DateSpan">此图表的TimeSpan跨度</param>
        ''' <param name="type">此图表的类型，则枚举DrawingType提供</param>
        '''  <param name="CanRoll">是图表是否可以滚动，即是动态图还是静态图</param>
        ''' <param name="date_ColNum">此测斜点的监测数据工作表中的每一天与其在工作表中对应的列号，
        ''' 以监测数据的日期key索引数据所在列号item</param>
        ''' <param name="usedRg">监测曲线的数据范围，此Range对象中，
        ''' 包括此工作表的UsedRange中的第一列，即深度的数据；但是不包括第一行的日期数据</param>
        ''' <param name="Info">记录数据信息的文本框</param>
        ''' <param name="FirstSeriesTag">第一条数据系列对应的Tag信息</param>
        ''' <param name="ProcessRegionData">在施工进度工作表中，每一个基坑区域相关的各种信息，比如区域名称，区域的描述，
        ''' 区域数据的Range对象，区域所属的基坑ID及其ID的数据等</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal DataSheet As Worksheet, ByVal DrawingChart As Chart,
                        ByVal ParentApp As Cls_ExcelForMonitorDrawing, ByVal DateSpan As DateSpan,
                       ByVal type As DrawingType, ByVal CanRoll As Boolean, ByVal Info As Microsoft.Office.Interop.Excel.TextFrame2,
                       ByVal DrawingTag As MonitorInfo, ByVal MonitorType As MntType,
                       ByVal date_ColNum As Dictionary(Of Date, Integer),
                       ByVal usedRg As Range,
                       ByVal FirstSeriesTag As SeriesTag_Incline,
                       Optional ByVal ProcessRegionData As clsData_ProcessRegionData = Nothing)
            '
            Call MyBase.New(DataSheet, DrawingChart, ParentApp, type, CanRoll, Info, DrawingTag, MonitorType,
                            DateSpan, New SeriesTag(FirstSeriesTag.series, FirstSeriesTag.ConstructionDate))
            ' --------------------------------------------
            Try
                With Me
                    .P_rgMntData = usedRg  ''包括第一列，但是不包括第一行的日期。
                    .F_YValues = ExcelFunction.ConvertRangeDataToVector(Of Single)(usedRg.Columns(1))
                    .Information = Info
                    ._ExcavationDepth_lineAndTextbox = FirstSeriesTag
                    .F_dicDateAndColumnNumber = date_ColNum
                    .P_ProcessRegionData = ProcessRegionData
                    ’
                    ._inclineTopElevaion = Project_Expo.InclineTopElevaion
                    ' ----- 集合数据的记录
                    .F_DicSeries_Tag.Item(LowIndexOfObjectsInExcel.SeriesInSeriesCollection) = FirstSeriesTag
                End With

                ' -----对图例进行更新---------
                'Call LegendRefresh(List_HasCurve)

            Catch ex As Exception
                MessageBox.Show("构造测斜曲线图出错。" & vbCrLf & ex.Message & vbCrLf &
                                ex.TargetSite.Name, "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End Try
        End Sub

#Region "  ---  Rolling事件"

        ''' <summary>
        ''' 图形滚动的Rolling方法
        ''' </summary>
        ''' <param name="dateThisday">施工当天的日期</param>
        ''' <remarks></remarks>
        Public Overrides Sub Rolling(ByVal dateThisday As Date)
            MyBase.F_RollingDate = dateThisday
            Dim lockobject As New Object
            SyncLock lockobject
                Dim app As Application = MyBase.Chart.Application
                app.ScreenUpdating = False
                ' ------------------- 绘制监测曲线图
                Dim Allday = Me.F_dicDateAndColumnNumber.Keys
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

                ' -------------------- 移动开挖深度的直线与文本框
                If Me.P_ProcessRegionData IsNot Nothing Then
                    With Me.P_ProcessRegionData.Date_Elevation
                        Dim excavationElevation As Single
                        Try
                            excavationElevation = .Item(dateThisday)
                            'ClsData_DataBase.GetElevation(P_rgExcavationProcess, dateThisday)
                        Catch ex As Exception
                            Debug.Print("上面的Exception已经被Try...Catch语句块捕获，不用担心。出错代码位于ClsDrawing_Mnt_Incline.vb中。")
                            Dim ClosestDate As Date = ClsData_DataBase.FindTheClosestDateInSortedList(.Keys, dateThisday)
                            excavationElevation = .Item(ClosestDate)
                        End Try
                        Call MoveExcavation(_ExcavationDepth_lineAndTextbox, excavationElevation, dateThisday, State)
                    End With
                End If
                app.ScreenUpdating = True
            End SyncLock
        End Sub

        ''' <summary>
        ''' 根据每天不同的挖深情况，移动挖深线的位置
        ''' </summary>
        ''' <param name="ExcavationLineAndTextBox">表示挖深深度的直线，文本框中记录有“挖深”二字</param>
        ''' <param name="excavElevation"> 当天开挖标高 </param>
        ''' <param name="dateThisday">施工当天的日期</param>
        ''' <remarks></remarks>
        Private Sub MoveExcavation(ByVal ExcavationLineAndTextBox As SeriesTag_Incline, ByVal excavElevation As Single,
                                   ByVal dateThisday As Date, ByVal State As TodayState)
            Dim iline As Microsoft.Office.Interop.Excel.Shape = ExcavationLineAndTextBox.DepthLine
            Dim itextbox As Microsoft.Office.Interop.Excel.Shape = ExcavationLineAndTextBox.DepthTextbox
            '
            If State = TodayState.DateMatched Or State = TodayState.DateNotFound Then
                ' 将当天的开挖标高转换为相对于测斜管顶部的深度值
                Dim relativedepth As Single = _inclineTopElevaion - excavElevation 'Project_Expo.Elevation_GroundSurface - excavElevation
                '
                Dim axisY As Axis = Me.Chart.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue)
                Dim scalemax As Single = axisY.MaximumScale
                Dim scalemin As Single = axisY.MinimumScale
                '
                Dim linetop As Single
                Dim plotA As PlotArea = Me.Chart.PlotArea
                With plotA
                    '将相对深度值转换与图表坐标系中的深度坐标值！！！！！
                    linetop = .InsideTop + .InsideHeight * (relativedepth - scalemin) / (scalemax - scalemin)
                End With
                '
                With iline
                    .Visible = MsoTriState.msoCTrue
                    .Top = linetop
                End With

                '
                If itextbox IsNot Nothing Then
                    With itextbox
                        .Visible = MsoTriState.msoCTrue
                        .Top = iline.Top - itextbox.Height

                        '指示此基坑当前的状态，为向下开挖还是向上建造。
                        With Me.P_ProcessRegionData
                            Dim strStatus As String = ""
                            If Me.P_ProcessRegionData Is Nothing OrElse (Not .HasBottomDate) Then
                                strStatus = "挖深/施工深度"
                            Else
                                Dim CompareIndex As Integer = Date.Compare(dateThisday, .BottomDate)
                                If CompareIndex < 0 Then        '说明还未开挖到底
                                    strStatus = "挖深"
                                ElseIf CompareIndex > 0 Then    '说明已经开挖到底，正在施工上部结构
                                    strStatus = "施工深度"
                                ElseIf CompareIndex = 0 Then    '说明刚好开挖到基坑底部
                                    strStatus = "开挖到底"
                                End If
                            End If
                            itextbox.TextFrame2.TextRange.Text = strStatus & relativedepth
                        End With
                    End With
                End If

            Else
                iline.Visible = MsoTriState.msoFalse
                '
                If itextbox IsNot Nothing Then
                    itextbox.Visible = False
                End If
            End If

        End Sub

        ''' <summary>
        ''' 绘制距选定的日期（包括选择的日期）最近的日期的监测数据
        ''' </summary>
        ''' <param name="dateThisDay">选定的施工日期</param>
        ''' <param name="State">指示选定的日期所处的状态</param>
        ''' <param name="Closestday">距离选定日期最近的日期，如果选定的日期在工作表中有数据，则等效于dateThisDay</param>
        ''' <remarks></remarks>
        Private Sub CurveRolling(ByVal dateThisDay As Date, ByVal State As TodayState, ByVal Closestday As Date)
            Const strDateFormat As String = AMEApplication.DateFormat
            F_DicSeries_Tag.Item(1).ConstructionDate = dateThisDay  '刷新滚动的曲线所代表的日期
            Dim blnHasCurve As Boolean
            '
            '进行监测曲线的滚动
            With Me.MovingSeries
                .Values = F_YValues           'Y轴的数据

                Select Case State
                    Case TodayState.BeforeStartDay
                        .XValues = {Nothing}     '不能设置为Series.Value=vbEmpty，因为这会将x轴标签中的某一个值设置为0.0。
                        .Name = dateThisDay.ToString(strDateFormat) & " :早于" &
                            Me.DateSpan.StartedDate.ToString(strDateFormat)
                        blnHasCurve = False

                    Case TodayState.AfterFinishedDay
                        .XValues = {Nothing}
                        .Name = dateThisDay.ToString(strDateFormat) & " :晚于" &
                            Me.DateSpan.FinishedDate.ToString(strDateFormat)
                        blnHasCurve = False

                    Case TodayState.DateNotFound
                        .Name = Closestday.ToString(strDateFormat) & "(" & dateThisDay.ToString(strDateFormat) & "无数据" & ")"
                        blnHasCurve = True

                    Case TodayState.DateMatched
                        .Name = dateThisDay.ToString(strDateFormat)
                        blnHasCurve = True
                End Select

                If blnHasCurve Then
                    '---------------------------  设置信息文本框中的信息
                    Dim max As Double
                    Dim min As Double
                    Dim StrdepthForMax As String
                    Dim StrdepthForMin As String
                    '
                    Dim strcolumn = Me.F_dicDateAndColumnNumber.Item(Closestday)
                    '当天的监测数据的Range对象
                    Dim rgDisplacement As Range = P_rgMntData.Columns(strcolumn)     '只包含对应施工日期的监测数据的那一列Range对象 
                    '当天的监测数据，这里只能定义为Object的数组，因为有可能单元格中会出现没有数据的情况。
                    Dim MonitorData As Object() = ExcelFunction.ConvertRangeDataToVector(Of Object)(rgDisplacement)
                    .XValues = MonitorData           'X轴的数据
                    '
                    With Me.Sheet_Data.Application.WorksheetFunction

                        'find the maximun/Minimum displacement
                        max = .Max(MonitorData)
                        min = .Min(MonitorData)
                        Try
                            ' find the corresponding depth of the maximun displacement
                            '如果 MATCH 函数查找匹配项不成功，在Excel中它会返回错误值 #N/A。 而在VB.NET中，它会直接报错。
                            Dim Row_Max As Integer = rgDisplacement.Cells(1, 1).row - 1 + .Match(max, MonitorData, 0)
                            Dim sngDepthForMax As Single = Me.Sheet_Data.Cells(Row_Max, Data_Drawing_Format.Mnt_Incline.ColNum_Depth).value
                            StrdepthForMax = sngDepthForMax.ToString("0.0")
                        Catch ex As Exception
                            StrdepthForMax = " Null "
                            Debug.Print("搜索最大位移所对应的深度失败！")
                        End Try
                        Try
                            ' find the corresponding depth of the mininum displacement
                            '如果 MATCH 函数查找匹配项不成功，在Excel中它会返回错误值 #N/A。 而在VB.NET中，它会直接报错。
                            Dim Row_Min As Integer = rgDisplacement.Cells(1, 1).row - 1 + .Match(min, MonitorData, 0)
                            Dim sngDepthForMin As Single = Me.Sheet_Data.Cells(Row_Min, Data_Drawing_Format.Mnt_Incline.ColNum_Depth).value
                            StrdepthForMin = sngDepthForMin.ToString("0.0")
                        Catch ex As Exception
                            StrdepthForMin = " Null "
                            Debug.Print("搜索最小位移所对应的深度失败！")
                        End Try
                        '在信息文本框中输出文本，记录当前曲线的最大与最小位移，以及出现的深度
                        Information.TextRange.Text = "Max:" & max.ToString("0.00") & "mm" & vbTab _
                            & "In depth of " & StrdepthForMax & "m" _
                            & vbCrLf & "Min:" & min.ToString("0.00") & "mm" & vbTab _
                            & "In depth of " & StrdepthForMin & "m"
                    End With
                Else
                    Information.TextRange.Text = "Max:" & vbTab _
                                & "In depth of " _
                                & vbCrLf & "Min:" & vbTab _
                                & "In depth of "
                End If
            End With
        End Sub

#End Region

#Region "  ---  数据列的锁定与删除"

        '删除
        Public Overrides Sub DeleteSeries(DeletingSeriesIndex As Integer)
            '删除数据系列以及对应的开挖深度直线与开挖深度文本框
            Dim seriesTag_incline As SeriesTag_Incline = DirectCast(Me.F_DicSeries_Tag(DeletingSeriesIndex), SeriesTag_Incline)
            With seriesTag_incline
                If .DepthLine IsNot Nothing Then .DepthLine.Delete()
                If .DepthTextbox IsNot Nothing Then .DepthTextbox.Delete()
            End With
            '
            MyBase.DeleteSeries(DeletingSeriesIndex)

        End Sub
        '添加
        Public Overrides Function CopySeries(ByVal SourceSeriesIndex As Integer) As KeyValuePair(Of Integer, SeriesTag)
            ' 调用基类中的方法
            Dim NewSeriesIndex_Tag As KeyValuePair(Of Integer, SeriesTag) = MyBase.CopySeries(SourceSeriesIndex)
            '
            ' ------------ 为新的数据列创建其对应的Tag
            ' ------------ 这是这个类与其基类所不同的地方，将下面这个函数删除，即与其基类中的这个方法相同了。  ------------
            Dim NewSeriesIndex As Integer = NewSeriesIndex_Tag.Key      '新数据列的索引下标，第一条曲线的下标为1；
            Dim newTag As SeriesTag_Incline         '新数据列的标签Tag
            With NewSeriesIndex_Tag.Value
                newTag = New SeriesTag_Incline(.series, .ConstructionDate)
            End With

            Try
                newTag = Me.ConstructNewTag(F_DicSeries_Tag.Item(SourceSeriesIndex), newTag.series)
            Catch ex As Exception
                Debug.Print(ex.Message)
                MessageBox.Show("在设置""表示开挖深度的标志线与文本框""时出错", "Error",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
            '---------------------- 
            '将基类中的F_DicSeries_Tag.Item(NewSeriesIndex)的值由SeriesTag修改为SeriesTag_Incline
            Me.F_DicSeries_Tag.Item(NewSeriesIndex) = newTag
            Return Nothing
        End Function

        ''' <summary>
        ''' 为新的数据列创建其对应的Tag，并且在Excel的绘图中复制出对应的表示开挖深度的深度线与文本框。
        ''' </summary>
        ''' <param name="SourceTag">用以参考的数据系列的Tag信息</param>
        ''' <param name="newSeries">复制出来的新的数据系列</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function ConstructNewTag(ByVal SourceTag As SeriesTag_Incline, ByVal newSeries As Series) As SeriesTag_Incline

            Dim newTag As New SeriesTag_Incline(newSeries, MyBase.F_RollingDate)

            ' ------------ 为新的数据列创建其对应的Tag
            '数据系列的颜色
            Dim seriesColor As Integer = newSeries.Format.Line.ForeColor.RGB

            '复制一个新的开挖深度线
            Try

                If SourceTag.DepthLine IsNot Nothing Then
                    newTag.DepthLine = SourceTag.DepthLine.Duplicate
                    With newTag.DepthLine
                        '线条的位置
                        .Top = SourceTag.DepthLine.Top
                        '.Left = SourceTag.DepthLine.Left
                        .Left = 0

                        '线条的颜色
                        .Line.ForeColor.RGB = seriesColor
                        '线条的阴影
                        .Shadow.Type = MsoShadowType.msoShadow21
                        '线条的缩放
                        .ScaleWidth(Factor:=0.5,
                                    RelativeToOriginalSize:=MsoTriState.msoFalse,
                                    Scale:=MsoScaleFrom.msoScaleFromTopLeft)
                    End With
                End If
            Catch ex As Exception
                Debug.Print(ex.Message)
                MessageBox.Show("在设置""表示开挖深度的标志线""时出错", "Error",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

            '复制一个新的开挖深度文本框
            Try
                If SourceTag.DepthTextbox IsNot Nothing Then
                    Dim oldTxtShp As Microsoft.Office.Interop.Excel.Shape = SourceTag.DepthTextbox
                    newTag.DepthTextbox = oldTxtShp.Duplicate

                    '设置文本框的位置
                    Dim newTxtShp As Microsoft.Office.Interop.Excel.Shape = newTag.DepthTextbox
                    With newTxtShp
                        '设置文本框的位置
                        .Top = oldTxtShp.Top
                        '.Left = oldTxtShp.Left
                        .Left = 0

                        '设置文本框中字体的格式
                        With .TextFrame2.TextRange.Font
                            .Fill.ForeColor.RGB = seriesColor
                            .Size = 10
                            .Bold = False
                        End With

                    End With
                End If
            Catch ex As Exception
                Debug.Print(ex.Message)
                MessageBox.Show("在设置""表示开挖深度的文本框""时出错", "Error",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
            Return newTag
        End Function
#End Region

    End Class
End Namespace