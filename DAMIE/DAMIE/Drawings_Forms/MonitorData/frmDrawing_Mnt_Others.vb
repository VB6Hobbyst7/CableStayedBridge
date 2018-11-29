Imports Microsoft.Office.Interop.Excel
Imports Office = Microsoft.Office.Core
Imports DAMIE.DataBase
Imports DAMIE.Miscellaneous
Imports DAMIE.Constants
Imports DAMIE.All_Drawings_In_Application
Imports DAMIE.GlobalApp_Form
Imports DAMIE.All_Drawings_In_Application.ClsDrawing_Mnt_Base
Imports DAMIE.Constants.Data_Drawing_Format.Mnt_Others
Imports DAMIE.Constants.Data_Drawing_Format.Drawing_Mnt_Others
Public Class frmDrawing_Mnt_Others

#Region "  ---  Declarations and Definitions"

    ''' <summary>
    ''' 当前进行绘图的数据工作簿发生变化时触发
    ''' </summary>
    ''' <remarks></remarks>
    Private Event DataWorkbookChanged(ByVal WorkingDataWorkbook As Workbook)

#Region "  ---  Fields"
    ''' <summary>
    ''' 当前用于绘图的数据工作簿
    ''' </summary>
    ''' <remarks></remarks>
    Dim F_wkbkData As Workbook
    Dim F_wkbkDrawing As Workbook
    Dim F_shtMonitorData As Worksheet
    Dim F_shtDrawing As Worksheet

    ''' <summary>
    ''' 进行操作的工作表的UsedRange的右下角的行号与列号
    ''' </summary>
    ''' <remarks></remarks>
    Dim F_arrBottomRightCorner(0 To 1) As Integer

    ''' <summary>
    ''' ！以每一天的日期来索引这一天的监测数据，监测数据只包含列表中选择了的监测点
    ''' </summary>
    ''' <remarks></remarks>
    Dim F_dicDate_ChosenDatum As New Dictionary(Of Date, Object())

    ''' <summary>
    ''' 用来进行绘图的Excel程序
    ''' </summary>
    ''' <remarks></remarks>
    Dim F_AppDrawing As Application

    '画布上的变量
    Dim F_Chart As Chart
    Dim F_textbox_Info As TextFrame2
    '
    ''' <summary>
    ''' 在测点列表框中所选择的所有测点标签的数组，如{SW1,SW2,SW3}
    ''' </summary>
    ''' <remarks></remarks>
    Private F_SelectedTags() As String
    ''' <summary>
    ''' 在测点列表框中所选择的每一个测点在Excel工作表中所对应的行号
    ''' </summary>
    ''' <remarks></remarks>
    Private F_RowNum_SelectedTags() As Integer
    '
    ''' <summary>
    ''' Chart中第一条监测曲线所对应的相关信息
    ''' </summary>
    ''' <remarks></remarks>
    Private F_TheFirstseriesTag As clsDrawing_Mnt_RollingBase.SeriesTag
    '
    ''' <summary>
    ''' 所绘图的监测数据类型
    ''' </summary>
    ''' <remarks></remarks>
    Private F_MonitorType As MntType
    '
    Private F_MainForm As APPLICATION_MAINFORM = APPLICATION_MAINFORM.MainForm

#End Region

#End Region

#Region "  ---  窗口的加载与关闭"

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Call SetMonitorType(Me.ComboBox_MntType)
        AddHandler ClsData_DataBase.WorkingStageChanged, AddressOf Me.RefreshCombox_WorkingStage

        ' ------------ 
        With Me.ComboBoxOpenedWorkbook
            .DisplayMember = LstbxDisplayAndItem.DisplayMember
            .ValueMember = LstbxDisplayAndItem.ValueMember
        End With

    End Sub

    '窗口加载
    Private Sub frmDrawingMonitor_Load(sender As Object, e As EventArgs) Handles Me.Load

        '设置控件的默认属性
        btnGenerate.Enabled = False
        chkBoxOpenNewExcel.Checked = True
        RbtnDynamic.Checked = True
    End Sub
    '在关闭窗口时将其隐藏
    Private Sub frmDrawing_Mnt_Others_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        '如果是子窗口自己要关闭，则将其隐藏
        '如果是mdi父窗口要关闭，则不隐藏，而由父窗口去结束整个进程
        If Not e.CloseReason = CloseReason.MdiFormClosing Then
            Me.Hide()
        End If
        e.Cancel = True
    End Sub

    Private Sub frmDrawing_Mnt_Others_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
        RemoveHandler ClsData_DataBase.WorkingStageChanged, AddressOf Me.RefreshCombox_WorkingStage
    End Sub

#End Region

    ''' <summary>
    ''' 生成监测曲线图
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnGenerate_Click(sender As Object, e As EventArgs) Handles btnGenerate.Click
        Dim SelectedTagsCount As Integer = ListBoxPointsName.SelectedIndices.Count
        If SelectedTagsCount > 0 Then
            '
            With Me.ListBoxPointsName
                ReDim F_SelectedTags(0 To SelectedTagsCount - 1)
                ReDim F_RowNum_SelectedTags(0 To SelectedTagsCount - 1)
                Dim i As Integer = 0
                For Each item As LstbxDisplayAndItem In .SelectedItems
                    F_SelectedTags(i) = DirectCast(item.DisplayedText, String)
                    F_RowNum_SelectedTags(i) = DirectCast(item.Value, Integer)
                    i += 1
                Next
            End With

            '开始绘图
            With Me.BGWK_NewDrawing
                If Not .IsBusy Then

                    '用来判断是否要创建新的Excel程序来进行绘图，以及是否要对新画布所在的Excel进行美化。
                    Dim blnNewExcelApp = GlobalApplication.Application.MntDrawing_ExcelApps.Count = 0 Or chkBoxOpenNewExcel.Checked
                    Call .RunWorkerAsync(blnNewExcelApp)
                End If
            End With
        Else
            MessageBox.Show("请选择至少一个测点。", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

#Region "  ---  后台线程进行操作"

    ''' <summary>
    ''' 生成绘图，此方法是在后台的工作者线程中执行的。
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BGW_Generate_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BGWK_NewDrawing.DoWork
        '主程序界面的进度条的UI显示
        F_MainForm.ShowProgressBar_Marquee()
        '执行具体的绘图操作
        Try
            '用来判断是否要创建新的Excel程序来进行绘图，以及是否要对新画布所在的Excel进行美化。
            Dim blnNewExcelApp As Boolean = CType(e.Argument, Boolean)
            Call Generate(blnNewExcelApp, Me.F_blnDrawDynamic, Me.F_SelectedTags, Me.F_RowNum_SelectedTags)

        Catch ex As Exception
            MessageBox.Show("绘制监测曲线图失败！" & vbCrLf & ex.Message & vbCrLf & "报错位置：" & ex.TargetSite.Name, _
        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' 当后台的工作者线程结束（即BGW_Generate_DoWork方法执行完毕）时触发，注意，此方法是在UI线程中执行的。
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BGW_Generate_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BGWK_NewDrawing.RunWorkerCompleted
        '在绘图完成后，隐藏进度条
        F_MainForm.HideProgress("Done")
    End Sub

#End Region

    ''' <summary>
    ''' 1、开始绘图
    ''' </summary>
    ''' <param name="CreateNewExcelApp">是否要新开一个Excel程序</param>
    ''' <param name="DrawDynamicDrawing">是否是要绘制动态曲线图</param>
    ''' <param name="SelectedTags">选择的测点</param>
    ''' <param name="RowNum_SelectedTags">选择的测点在Excel数据工作表中所在的行号</param>
    ''' <remarks></remarks>
    Private Sub Generate(ByVal CreateNewExcelApp As Boolean, ByVal DrawDynamicDrawing As Boolean, _
                         ByVal SelectedTags() As String, ByVal RowNum_SelectedTags() As Integer)

        ' ----------arrDateRange----------------------- 获取此工作表中的整个施工日期的数组（0-Based，数据类型为Date）
        Dim arrDateRange As Date() = GetDate()
        Dim arrDateRange_Double(0 To arrDateRange.Count - 1) As Double
        For i As Integer = 0 To arrDateRange.Count - 1
            arrDateRange_Double(i) = arrDateRange(i).ToOADate
        Next
        '-----------AllSelectedMonitorData------------ 获取所有"选择的"监测点位的监测数据的大数组。其中不包含选择的测点编号信息与施工日期的信息
        '此数组的第一个元素的下标值为0
        Dim AllSelectedMonitorData(,) As Object
        AllSelectedMonitorData = GetAllSelectedMonitorData(Me.F_shtMonitorData, RowNum_SelectedTags)

        '----------------- 设置监测曲线的时间跨度
        Array.Sort(arrDateRange)
        Dim Date_Span As DateSpan
        Date_Span.StartedDate = arrDateRange(0)
        Date_Span.FinishedDate = arrDateRange(UBound(arrDateRange))

        '----------------------------打开用来绘图的Excel程序，并将此界面加入主程序的监测曲线集合
        Dim clsExcelForMonitorDrawing As Cls_ExcelForMonitorDrawing = Nothing
        '   --------------- 获取用来绘图的Excel程序，并将此界面加入主程序的监测曲线集合 -------------------
        F_AppDrawing = GetApplication(NewExcelApp:=CreateNewExcelApp, _
                                         ExcelForMntDrawing:=clsExcelForMonitorDrawing, _
                                         MntDrawingExcelApps:=GlobalApplication.Application.MntDrawing_ExcelApps)
        F_AppDrawing.ScreenUpdating = False
        '打开工作簿以画图
        If F_AppDrawing.Workbooks.Count = 0 Then
            F_wkbkDrawing = F_AppDrawing.Workbooks.Add
        Else
            F_wkbkDrawing = F_AppDrawing.Workbooks(1) '总是定义为第一个，因为就只开了一个
        End If
        '新开一个工作表以画图
        F_shtDrawing = F_wkbkDrawing.Worksheets.Add
        F_shtDrawing.Activate()


        '-------  根据是要绘制动态的曲线图还是静态的曲线图，来执行不同的操作  ---------------------
        If DrawDynamicDrawing Then

            '-------------绘制动态的曲线图--------------
            F_dicDate_ChosenDatum = GetdicDate_Datum_ForDynamic(arrDateRange, AllSelectedMonitorData)
            '开始画图
            F_Chart = DrawDynamicChart(F_dicDate_ChosenDatum, SelectedTags, AllSelectedMonitorData)
            '设置图表的Tag属性
            Dim Tags As MonitorInfo = GetChartTags(F_shtMonitorData)
            '-------------------------------------------------------------------------
            Dim DynamicSheet As New ClsDrawing_Mnt_OtherDynamics(F_shtMonitorData, F_Chart, _
                                                    clsExcelForMonitorDrawing, Date_Span, _
                                                    DrawingType.Monitor_Dynamic, True, F_textbox_Info, Tags, Me.F_MonitorType, _
                                                    F_dicDate_ChosenDatum, Me.F_TheFirstseriesTag)

            '-------------------------------------------------------------------------
        Else

            '-------绘制静态的曲线图--------
            '开始画图
            Dim dicSeriesData As New Dictionary(Of Series, Object())
            F_Chart = DrawStaticChart(SelectedTags, arrDateRange_Double, AllSelectedMonitorData, dicSeriesData)
            '设置图表的Tag属性
            Dim Tags As MonitorInfo = GetChartTags(F_shtMonitorData)
            If Me.F_WorkingStage IsNot Nothing Then
                Call DrawWorkingStage(Me.F_Chart, Me.F_WorkingStage)
            End If
            '-------------------------------------------------------------------------
            Dim staticSheet As New ClsDrawing_Mnt_Static(F_shtMonitorData, F_Chart, _
                                                         clsExcelForMonitorDrawing, _
                                                         DrawingType.Monitor_Static, False, F_textbox_Info, Tags, Me.F_MonitorType, _
                                                         dicSeriesData, arrDateRange_Double)

            '-------------------------------------------------------------------------
        End If

        '---------------------- 界面显示与美化
        Call ExcelAppBeauty(F_shtDrawing.Application, CreateNewExcelApp)
    End Sub

#Region "  ---  动态图"

    ''' <summary>
    ''' 绘制动态曲线图
    ''' </summary>
    ''' <param name="p_dicDate_ChosenDatum"></param>
    ''' <param name="arrChosenTags"></param>
    ''' <param name="arrDataDisplacement"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function DrawDynamicChart(ByVal p_dicDate_ChosenDatum As Dictionary(Of Date, Object()), _
                                 ByVal arrChosenTags() As String, _
                                 ByVal arrDataDisplacement(,) As Object) As Chart

        'dic_Date_ChosenDatum以每一天的日期来索引这一天的监测数据，监测数据只包含列表中选择了的监测点
        Dim r_myChart As Chart
        F_shtDrawing.Activate()
        '--------------------------------------------------------------- 在工作表“标高图”中添加图表
        With F_shtDrawing
            r_myChart = .Shapes.AddChart(XlChartType.xlLineMarkers).Chart

            '---------- 选定模板
            Dim t_path As String = System.IO.Path.Combine(My.Settings.Path_Template, _
                                                          Constants.FolderOrFileName.File_Template.Chart_Horizontal_Dynamic)
            ' 如果监测曲线图所使用的"Chart模板"有问题，则在chart.ChartArea.Copy方法（或者是chartObject.Copy方法）中可能会出错。
            r_myChart.ApplyChartTemplate(t_path)
            '-------------------- 获取图表中的信息文本框
            F_textbox_Info = r_myChart.Shapes(0).TextFrame2      'Chart中的Shapes集合的第一个元素的下标值为0
            'textbox_Info.AutoSize = Microsoft.Office.Core.MsoAutoSize.msoAutoSizeShapeToFitText


            '------------------------ 设置曲线的数据
            Dim Date_theFirstCurve As Date = p_dicDate_ChosenDatum.Keys(0)
            Dim mySeriesCollection As SeriesCollection = r_myChart.SeriesCollection
            Dim series As Series = mySeriesCollection.Item(1)
            With series
                .Name = Date_theFirstCurve     '系列名称
                .XValues = arrChosenTags                     'X轴的数据
                .Values = p_dicDate_ChosenDatum.Item(Date_theFirstCurve)     'Y轴的数据
            End With
            '
            Me.F_TheFirstseriesTag = New clsDrawing_Mnt_RollingBase.SeriesTag(series, Date_theFirstCurve)
            '------------------------ 设置X、Y轴的格式——监测点位编号
            With r_myChart.Axes(XlAxisType.xlCategory)
                .AxisTitle.Text = GetAxisLabel(DrawingType.Monitor_Dynamic, Me.F_MonitorType, XlAxisType.xlCategory)
            End With

            '-设置Y轴的格式——测点位移
            Dim axes As Axis = r_myChart.Axes(XlAxisType.xlValue)
            With axes

                '由数据的最小与最大值来划分表格区间
                Dim imin As Single = F_AppDrawing.WorksheetFunction.Min(arrDataDisplacement)
                Dim imax As Single = F_AppDrawing.WorksheetFunction.Max(arrDataDisplacement)

                '主要与次要刻度单位，先确定刻度单位是为了后面将坐标轴的区间设置为主要刻度单位的倍数
                Dim unit As Single = CSng(Format((imax - imin) / ClsDrawing_Mnt_OtherDynamics.cstChartParts_Y, "0.0E+00")) '这里涉及到有效数字的处理的问题
                .MajorUnit = unit
                .MinorUnitIsAuto = True

                '坐标轴上显示的总区间
                .MinimumScale = F_AppDrawing.WorksheetFunction.Floor_Precise(imin, .MajorUnit)
                .MaximumScale = F_AppDrawing.WorksheetFunction.Ceiling_Precise(imax, .MajorUnit)


                '坐标轴标题
                .AxisTitle.Text = GetAxisLabel(DrawingType.Monitor_Dynamic, Me.F_MonitorType, XlAxisType.xlValue)
            End With
        End With
        Return r_myChart
    End Function

    ''' <summary>
    ''' 返回动态曲线图的监测数据中的字典：以每一天的日期，索引选择的测点在当天的数据。
    ''' </summary>
    ''' <param name="arrDateRange"></param>
    ''' <param name="AllSelectedMonitorData"></param>
    ''' <returns>返回动态曲线图的关键参数：dic_Date_ChosenDatum，
    ''' 字典中的值这里只能定义为Object的数组，因为有可能单元格中会出现没有数据的情况。</returns>
    ''' <remarks></remarks>
    Private Function GetdicDate_Datum_ForDynamic(ByVal arrDateRange() As Date, _
                                                  ByVal AllSelectedMonitorData(,) As Object) _
                                              As Dictionary(Of Date, Object())
        '以每一天的日期来索引这一天的监测数据，监测数据只包含列表中选择了的监测点
        Dim dic_Date_ChosenDatum As New Dictionary(Of Date, Object())
        Dim iCol As Integer

        '
        Dim blnIgnoreWarning As Boolean = False
        For iCol = 0 To UBound(AllSelectedMonitorData, 2)
            '当天的监测数据，这里只能定义为Object的数组，因为有可能单元格中会出现没有数据的情况。
            Dim arrDataForEachColumn(0 To UBound(AllSelectedMonitorData, 1)) As Object

            For irow As Byte = 0 To UBound(AllSelectedMonitorData, 1)
                arrDataForEachColumn(irow) = AllSelectedMonitorData(irow, iCol)
            Next
            Try     '将新的一天的日期以及对应的监测数据信息添加到字典中
                dic_Date_ChosenDatum.Add(arrDateRange(iCol), arrDataForEachColumn)
            Catch ex As ArgumentException        '可能的报错原因：已添加了具有相同键的项。即工作表中有相同的日期
                If Not blnIgnoreWarning Then        '提醒用户
                    Dim result As System.Windows.Forms.DialogResult = MessageBox.Show("此工作表中的数据类型不规范，可能在监测数据中出现了相同的日期" & vbCrLf _
                                             & "出错的监测数据日期是：" & arrDateRange(iCol).ToShortDateString & " . 是否忽略此提醒？" _
                                             & vbCrLf _
                                            & ex.Message & vbCrLf & ex.TargetSite.Name, "Warning", MessageBoxButtons.YesNo)
                    Select Case result
                        Case System.Windows.Forms.DialogResult.No
                            Continue For
                        Case System.Windows.Forms.DialogResult.Yes       '继续添加而不再报错
                            blnIgnoreWarning = True
                    End Select
                End If

                'btnGenerate.Enabled = False
            End Try
        Next

        '-------------------------------------
        Return dic_Date_ChosenDatum
    End Function

#End Region

#Region "  ---  静态图"

    ''' <summary>
    ''' 绘制静态曲线图
    ''' </summary>
    ''' <param name="arrChosenTags"></param>
    ''' <param name="arrDateRange"></param>
    ''' <param name="AllSelectedMonitorData"></param>
    ''' <param name="dicSeries_Data"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function DrawStaticChart(ByVal arrChosenTags() As String, _
                                ByVal arrDateRange() As Double, ByVal AllSelectedMonitorData(,) As Object, _
                                ByRef dicSeries_Data As Dictionary(Of Series, Object())) As Chart
        '
        Dim r_myChart As Chart
        F_shtDrawing.Activate()
        '--------------------------------------------------------------- 在工作表“标高图”中添加图表
        With F_shtDrawing
            r_myChart = .Shapes.AddChart().Chart

            '---------- 选定模板
            Dim t_path As String = System.IO.Path.Combine(My.Settings.Path_Template, _
                                                          Constants.FolderOrFileName.File_Template.Chart_Horizontal_Static)
            ' 如果监测曲线图所使用的"Chart模板"有问题，则在chart.ChartArea.Copy方法（或者是chartObject.Copy方法）中可能会出错。
            r_myChart.ApplyChartTemplate(t_path)

            '-------------------- 获取图表中的信息文本框
            F_textbox_Info = r_myChart.Shapes(0).TextFrame2      'Chart中的Shapes集合的第一个元素的下标值为0
            'textbox_Info.AutoSize = Microsoft.Office.Core.MsoAutoSize.msoAutoSizeShapeToFitText

        End With

        '----------------------------- 对于数据系列的集合，开始为每一行数据添加新的曲线
        Dim mySeriesCollection As SeriesCollection = r_myChart.SeriesCollection
        Dim eachSeries As Series
        With mySeriesCollection
            If mySeriesCollection.Count < arrChosenTags.Length Then
                For i As Byte = 1 To arrChosenTags.Length - mySeriesCollection.Count
                    .NewSeries()
                Next
            Else
                For i = 1 To mySeriesCollection.Count - arrChosenTags.Length    '删除集合中的元素，先锁定要删的那个元素的下标，然后删n次（因为每次删除后其后面的又填补上来了）
                    mySeriesCollection(arrChosenTags.Length).delete()
                Next
            End If
            Dim i2 As Byte = 0
            For Each tag As String In arrChosenTags
                eachSeries = mySeriesCollection(i2)         '数据系列集合中的第一条曲线的下标值为0
                Dim arrselectedDataWithDate(0 To UBound(AllSelectedMonitorData, 2)) As Object
                For i_col As Integer = 0 To UBound(AllSelectedMonitorData, 2)
                    arrselectedDataWithDate(i_col) = AllSelectedMonitorData(i2, i_col)
                Next
                With eachSeries
                    .Name = arrChosenTags(i2)        '系列名称
                    .XValues = arrDateRange                     'X轴的数据:每一天的施工日期
                    .Values = arrselectedDataWithDate                    'Y轴的数据
                End With
                dicSeries_Data.Add(eachSeries, arrselectedDataWithDate)
                i2 += 1
            Next
        End With

        '------------------------ 设置X、Y轴的格式
        '——整个施工日期跨度
        Dim axesX As Axis = r_myChart.Axes(XlAxisType.xlCategory)
        With axesX
            .CategoryType = XlCategoryType.xlTimeScale
            '绘制坐标轴的数值区间
            .MaximumScale = Max_Array(Of Double)(arrDateRange)
            .MinimumScale = Min_Array(Of Double)(arrDateRange)
            '.MaximumScaleIsAuto = True
            '.MinimumScaleIsAuto = True

            .TickLabels.NumberFormatLocal = "yy/M/d"
            .TickLabelSpacingIsAuto = True
            .AxisTitle.Text = GetAxisLabel(DrawingType.Monitor_Static, Me.F_MonitorType, XlAxisType.xlCategory)
        End With

        '-设置Y轴的格式——测点位移
        Dim axesY As Axis = r_myChart.Axes(XlAxisType.xlValue)
        With axesY
            '坐标轴标题
            .AxisTitle.Text = GetAxisLabel(DrawingType.Monitor_Static, Me.F_MonitorType, XlAxisType.xlValue)

            '由数据的最小与最大值来划分表格区间
            Dim imin As Single = F_AppDrawing.WorksheetFunction.Min(AllSelectedMonitorData)
            Dim imax As Single = F_AppDrawing.WorksheetFunction.Max(AllSelectedMonitorData)

            '主要与次要刻度单位，先确定刻度单位是为了后面将坐标轴的区间设置为主要刻度单位的倍数
            Dim unit As Single = CSng(Format((imax - imin) / ClsDrawing_Mnt_Static.cstChartParts_Y, "0.0E+00")) '这里涉及到有效数字的处理的问题
            Try
                .MajorUnit = unit           '有可能会出现unit的值为0.0的情况
            Catch ex As Exception
                .MajorUnitIsAuto = True
            End Try
            .MinorUnitIsAuto = True

            '坐标轴上显示的总区间
            .MinimumScale = F_AppDrawing.WorksheetFunction.Floor_Precise(imin, .MajorUnit)
            .MaximumScale = F_AppDrawing.WorksheetFunction.Ceiling_Precise(imax, .MajorUnit)

        End With

        Return r_myChart
    End Function

    ''' <summary>
    ''' 绘制开挖工况的位置线
    ''' </summary>
    ''' <param name="Cht"></param>
    ''' <param name="WorkingStage"></param>
    ''' <remarks></remarks>
    Private Sub DrawWorkingStage(ByVal Cht As Chart, ByVal WorkingStage As List(Of clsData_WorkingStage))
        Dim AX As Axis = Cht.Axes(XlAxisType.xlCategory)

        Dim arrLineName(0 To WorkingStage.Count - 1) As String
        Dim arrTextName(0 To WorkingStage.Count - 1) As String
        With Cht
            Try
                Dim i As Integer = 0
                For Each WS As clsData_WorkingStage In WorkingStage
                    ' -------------------------------------------------------------------------------------------

                    Dim shpLine As Shape = .Shapes.AddLine(BeginX:=0, BeginY:=Cht.PlotArea.InsideTop, _
                                                           EndX:=0, _
                                                           EndY:=75)

                    With shpLine.Line
                        .Weight = 1.2
                        '.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
                        '.EndArrowheadLength = Microsoft.Office.Core.MsoArrowheadLength.msoArrowheadLengthMedium
                        '.EndArrowheadWidth = Microsoft.Office.Core.MsoArrowheadWidth.msoArrowheadWidthMedium
                        .DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineDashDot
                        .ForeColor.RGB = RGB(255, 0, 0)
                    End With
                    '
                    ExcelFunction.setPositionInChart(shpLine, AX, WS.ConstructionDate.ToOADate)
                    ' -------------------------------------------------------------------------------------------
                    Dim TextWidth As Single = 25 : Dim textHeight As Single = 10
                    Dim shpText As Shape
                    With shpLine
                        shpText = Cht.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, _
                                                                     Left:=.Left - TextWidth / 2, Top:=.Top - textHeight, Height:=textHeight, Width:=TextWidth)
                    End With
                    Call ExcelFunction.FormatTextbox_Tag(shpText.TextFrame2, Text:=WS.Description, _
                                                         HorizontalAlignment:=Microsoft.Office.Core.MsoParagraphAlignment.msoAlignCenter)
                    ' -------------------------------------------------------------------------------------------
                    arrLineName(i) = shpLine.Name
                    arrTextName(i) = shpText.Name
                    i += 1
                Next
                With Cht.Shapes
                    Dim shp1 As Shape = .Range(arrLineName).Group()
                    Dim shp2 As Shape = .Range(arrTextName).Group
                    .Range({shp1.Name, shp2.Name}).Group()
                End With
            Catch ex As Exception
                MessageBox.Show("设置开挖工况位置出现异常。" & vbCrLf & ex.Message & vbCrLf & "报错位置：" & ex.TargetSite.Name, _
       "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End Try
        End With
    End Sub

#End Region

#Region "  ---  通用子方法"


    ''' <summary>
    ''' 获取用来绘图的Excel程序，并将此界面加入主程序的监测曲线集合
    ''' </summary>
    ''' <param name="NewExcelApp">按情况看是否要打开新的Application</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetApplication(ByVal NewExcelApp As Boolean, _
                                    ByVal MntDrawingExcelApps As Dictionary_AutoKey(Of Cls_ExcelForMonitorDrawing), _
                                    ByRef ExcelForMntDrawing As Cls_ExcelForMonitorDrawing) As Application
        Dim app As Application
        If NewExcelApp Then                '打开新的Excel程序
            app = New Application
            ExcelForMntDrawing = New Cls_ExcelForMonitorDrawing(app)
        Else                                                            '在原有的Excel程序上作图
            ExcelForMntDrawing = MntDrawingExcelApps.Last.Value
            ExcelForMntDrawing.ActiveMntDrawingSheet.RemoveFormCollection()
            app = ExcelForMntDrawing.Application
        End If
        Return app
    End Function

    ''' <summary>
    ''' 当进行绘图的数据工作簿发生变化时触发
    ''' </summary>
    ''' <param name="WorkingDataWorkbook">要进行绘图的数据工作簿</param>
    ''' <remarks></remarks>
    Private Sub frmDrawing_Mnt_Incline_DataWorkbookChanged(WorkingDataWorkbook As Workbook) Handles Me.DataWorkbookChanged
        '-------- 在列表中显示出监测数据工作簿中的所有工作表
        Dim sheetsCount As Byte = WorkingDataWorkbook.Worksheets.Count
        Dim arrWorkSheets(0 To sheetsCount - 1) As LstbxDisplayAndItem
        Dim i As Byte = 0
        For Each sht As Worksheet In WorkingDataWorkbook.Worksheets
            arrWorkSheets(i) = New LstbxDisplayAndItem(sht.Name, sht)
            i += 1
        Next
        With listSheetsName
            .DisplayMember = LstbxDisplayAndItem.DisplayMember
            .ValueMember = LstbxDisplayAndItem.ValueMember
            .DataSource = arrWorkSheets
            '.Items.Clear()
            '.Items.AddRange(sheetsNameList)
            '.SelectedItem = .Items(0)
        End With
        '
        btnGenerate.Enabled = True
    End Sub

    Private Function GetDate() As Date()
        ' ----------arrDateRange----------------------- 获取此工作表中的整个施工日期的数组（0-Based，数据类型为Date）
        Dim rg_AllDay As Range
        '注意此数组元素的类型为Object，而且其第一个元素的下标值有可能是(1,1)
        With F_shtMonitorData
            Dim startColNum As Short = ColNum_FirstData_Displacement
            Dim endColNum As Short = F_arrBottomRightCorner(1)
            rg_AllDay = .Rows(RowNumForDate) _
                        .Range(.Cells(1, startColNum),
                               .Cells(1, endColNum))
        End With
        Dim arrDateRange() As Date
        arrDateRange = ExcelFunction.ConvertRangeDataToVector(Of Date)(rg_AllDay)
        '
        Dim Dt As Date
        Try
            Dim dicDate As New Dictionary(Of Date, Integer)
            For Each Dt In arrDateRange
                dicDate.Add(Dt, 0)
            Next
        Catch ex As Exception
            MessageBox.Show("工作表中的日期字段出错，出错的日期为：" & Dt.ToString("yyyy/MM/dd") & " 。请检查工作表！" & vbCrLf & ex.Message & vbCrLf & "报错位置：" & ex.TargetSite.Name, _
       "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return Nothing
        End Try
        Return arrDateRange
    End Function

    ''' <summary>
    ''' 返回所有选择的测点的监测数组组成的大数组
    ''' </summary>
    ''' <param name="RowNum_Tags"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetAllSelectedMonitorData(ByVal sheetMonitorData As Worksheet, RowNum_Tags As Integer()) As Object(,)
        'Dim Queue_SelectedData As New Queue(Of Object())

        Dim AllSelectedMonitorData(0 To RowNum_Tags.Length - 1, _
                                   0 To F_arrBottomRightCorner(1) - ColNum_FirstData_Displacement) As Object

        Dim arrAllData(,) As Object     '所有有效的监测数据，包括没有选择的测点上的数据，不包括日期行与测点编号列
        With sheetMonitorData
            arrAllData = .Range(.Cells(RowNum_FirstData_WithoutDate, ColNum_FirstData_Displacement), _
                              .Cells(F_arrBottomRightCorner(0), F_arrBottomRightCorner(1))).Value

        End With

        '---------------------- 所有选择的测点在工作表中对应的行号,并将行号对应到上面的数组arrAllData的行号
        '先要处理Excel中的数组与VB中的数组的下标差的问题
        Dim diff_Row As Byte = RowNum_FirstData_WithoutDate - LBound(arrAllData, 1)
        Dim diff_Col As Byte = ColNum_FirstData_Displacement - LBound(arrAllData, 2)


        '选择的测点编号所在的行对应到上面的ArrAllData数组中的行，第一个元素的下标为0
        Dim Count_ChosenTag As Byte = RowNum_Tags.Length
        Dim arrChosenRownNum(0 To Count_ChosenTag - 1) As String
        Dim i2 As Byte = 0
        For Each RowNum_Tag As Integer In RowNum_Tags
            arrChosenRownNum(i2) = RowNum_Tag - diff_Row
            i2 += 1
        Next

        '------------ 为超级大数组赋值
        Dim setting As New mySettings_Application
        Dim blnCheckForEmpty As Boolean = setting.CheckForEmpty
        'Dim arrDataWithDate_EachTag(0 To UBound(arrAllData, 2) - LBound(arrAllData, 2) + 1) As Object
        If Not blnCheckForEmpty Then
            Dim irow As Byte = 0
            For Each rownum As Short In arrChosenRownNum
                For iCol As Short = LBound(arrAllData, 2) To UBound(arrAllData, 2)
                    'arrDataWithDate_EachTag(iCol - LBound(arrAllData, 2)) = arrAllData(rownum, iCol)
                    AllSelectedMonitorData(irow, iCol - LBound(arrAllData, 2)) = arrAllData(rownum, iCol)
                Next
                'Queue_SelectedData.Enqueue(arrDataWithDate_EachTag)
                irow += 1
            Next
        Else
            Dim irow As Byte = 0
            For Each rownum As Byte In arrChosenRownNum
                For iCol As Integer = LBound(arrAllData, 2) To UBound(arrAllData, 2)
                    '如果单元格中是空字符串，而不是Nothing，则将其转换为Nothing。
                    Dim v As Object = arrAllData(rownum, iCol)
                    If v IsNot Nothing Then
                        If v.GetType = GetType(String) Then
                            If v.ToString.Length = 0 Then
                                v = Nothing
                            End If
                        End If
                    End If
                    '赋值
                    'arrDataWithDate_EachTag(iCol - LBound(arrAllData, 2)) = v
                    AllSelectedMonitorData(irow, iCol - LBound(arrAllData, 2)) = v
                Next
                'Queue_SelectedData.Enqueue(arrDataWithDate_EachTag)
                irow += 1
            Next
        End If

        Return AllSelectedMonitorData
    End Function

    ''' <summary>
    ''' 设置图表的Tags属性
    ''' </summary>
    ''' <param name="MntDataSheet">监测数据所在的工作表</param>
    ''' <remarks></remarks>
    Private Function GetChartTags(ByVal MntDataSheet As Worksheet) As MonitorInfo
        Dim MonitorItem As String
        Dim ExcavationRegion As String
        '
        Dim t_wkbk As Workbook = MntDataSheet.Parent
        Dim filepathwithoutextension As String = System.IO.Path.GetFileNameWithoutExtension(t_wkbk.FullName)
        MonitorItem = filepathwithoutextension
        '
        ExcavationRegion = MntDataSheet.Name
        '
        Dim Tags As New MonitorInfo(MonitorItem, ExcavationRegion)
        Return Tags
    End Function

    ''' <summary>
    ''' 程序界面美化
    ''' </summary>
    ''' <param name="app">用来进行开头与界面设置的Excel界面</param>
    ''' <remarks>设置的元素包括Excel界面的尺寸、工具栏、滚动条等的显示方式</remarks>
    Private Sub ExcelAppBeauty(ByVal app As Application, ByVal blnCreateNewExcelApplication As Boolean)
        With app
            .DisplayStatusBar = False
            .DisplayFormulaBar = False

            With .ActiveWindow  '对标高图工作表的界面进行设置
                .DisplayGridlines = False
                .DisplayHeadings = False
                .DisplayWorkbookTabs = False
                .Zoom = 100
                .DisplayHorizontalScrollBar = False
                .DisplayVerticalScrollBar = False
                .WindowState = XlWindowState.xlMaximized
            End With
            If blnCreateNewExcelApplication Then           '如果是新创建的Excel界面，则进行美化，否则保持原样
                .Visible = True
                .WindowState = XlWindowState.xlNormal
                .ExecuteExcel4Macro("SHOW.TOOLBAR(""Ribbon"",false)")
            End If
            .ScreenUpdating = True
        End With
    End Sub

#End Region

#Region "  ---  界面操作"

    ''' <summary>
    ''' 选择监测数据的文件
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnChooseMonitorData_Click(sender As Object, e As EventArgs) Handles btnChooseMonitorData.Click
        Dim FilePath As String = ""
        With APPLICATION_MAINFORM.MainForm.OpenFileDialog1
            .Title = "选择测斜数据文件"
            .Filter = "Excel文件(*.xlsx, *.xls, *.xlsb)|*.xlsx;*.xls;*.xlsb"
            .FilterIndex = 2
            If .ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                FilePath = .FileName
            Else
                Exit Sub
            End If
        End With
        If FilePath.Length > 0 Then
            '将监测数据文件在DataBase的Excel程序中打开
            Try
                '有可能会出现选择了同样的监测数据文档
                Dim fileHasOpened As Boolean = False
                For Each item As LstbxDisplayAndItem In Me.ComboBoxOpenedWorkbook.Items
                    Dim wkbk As Workbook = DirectCast(item.Value, Workbook)
                    If String.Compare(wkbk.FullName, FilePath, True) = 0 Then
                        Me.F_wkbkData = wkbk
                        fileHasOpened = True
                        Exit For
                    End If
                Next
                ' ----------------------------
                If fileHasOpened Then
                    MessageBox.Show("选择的工作簿已经打开", "Tip", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    Me.F_wkbkData = GlobalApplication.Application.ExcelApplication_DB.Workbooks.Open( _
                        Filename:=FilePath, UpdateLinks:=False, ReadOnly:=True)
                    With Me.ComboBoxOpenedWorkbook
                        Dim lstItem As LstbxDisplayAndItem = New LstbxDisplayAndItem(Me.F_wkbkData.Name, Me.F_wkbkData)
                        .Items.Add(lstItem)
                        .SelectedItem = lstItem
                    End With
                    RaiseEvent DataWorkbookChanged(Me.F_wkbkData)
                End If
            Catch ex As Exception
                Debug.Print("打开新的数据工作簿出错！")
                Exit Sub
            End Try
        End If
    End Sub

    ''' <summary>
    ''' 在Visio中绘制相应的测点位置
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnDrawMonitorPoints_Click(sender As Object, e As EventArgs) Handles btnDrawMonitorPoints.Click
        Call GlobalApplication.Application.DrawingPointsInVisio()
    End Sub

#Region "  ---  选择列表框内容时进行赋值"

    ''' <summary>
    ''' 设置监测数据的类型
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ComboBox_MntType_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox_MntType.SelectedIndexChanged
        Dim item As LstbxDisplayAndItem = DirectCast(Me.ComboBox_MntType.SelectedItem, LstbxDisplayAndItem)
        Me.F_MonitorType = DirectCast(item.Value, MntType)
    End Sub

    ''' <summary>
    ''' 当选择的数据工作表发生变化时引发
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub listSheetsName_SelectedIndexChanged(sender As Object, e As EventArgs) Handles listSheetsName.SelectedIndexChanged
        Me.F_shtMonitorData = DirectCast(listSheetsName.SelectedValue, Worksheet)   'wkbkData.Worksheets(listSheetsName.SelectedItem)
        Dim arrPointsTag() As String       '此工作表中的测点的编号列表
        With F_shtMonitorData
            '测点编号所在的起始行与末尾行 
            Dim startRow As Integer = RowNum_FirstData_WithoutDate
            Dim endRow As Integer = .UsedRange.Rows.Count
            Dim endColumn As Integer = .UsedRange.Columns.Count
            F_arrBottomRightCorner = {endRow, endColumn}  ''进行操作的工作表的UsedRange的右下角的行号与列号
            If endRow >= startRow Then

                Dim rg_PointsTag As Range = .Columns(ColNum_PointsTag). _
                    range(.Cells(startRow, 1), .Cells(endRow, 1))
                arrPointsTag = ExcelFunction.ConvertRangeDataToVector(Of String)(rg_PointsTag)


                '------------   将编号列表的每一个值与其对应的行号添加到字典dicPointTag_RowNum中
                Dim TagsCount As Short = arrPointsTag.Length
                Dim arrPoints(0 To TagsCount - 1) As LstbxDisplayAndItem

                Dim add As Integer = 0
                Dim i As Integer = 0
                For Each tag As String In arrPointsTag
                    '在Excel数据表中，每一个监测点位的Tag所在的行号。
                    Dim RowNumToPointsTag As Short = startRow + add
                    arrPoints(i) = New LstbxDisplayAndItem(tag, startRow + add)
                    add += 1
                    i += 1
                Next

                '----------------------  将编号列表的所有值显示在窗口的测点列表中
                With ListBoxPointsName
                    .DisplayMember = LstbxDisplayAndItem.DisplayMember
                    .ValueMember = LstbxDisplayAndItem.ValueMember
                    .DataSource = arrPoints
                    '.Items.Clear()
                    '.Items.AddRange(arr)
                    '.SelectedIndex = 0          '选择列表中的第一项
                End With
                '
                btnGenerate.Enabled = True
            Else                '说明此表中的监测数据行数小于1
                ListBoxPointsName.DataSource = Nothing
                'ListPointsName.Items.Clear()
                MessageBox.Show("此工作表没有合法数据（数据行数小于1行）,请选择合适的工作表", "Warning", _
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
                F_shtMonitorData = Nothing
                btnGenerate.Enabled = False
            End If
        End With
    End Sub

    ''' <summary>
    ''' 指示是要用此窗口来绘制动态图还是静态图
    ''' </summary>
    ''' <remarks></remarks>
    Private F_blnDrawDynamic As Boolean
    Private Sub RbtnStaticWithTime_CheckedChanged(sender As Object, e As EventArgs) Handles RbtnStaticWithTime.CheckedChanged, RbtnDynamic.CheckedChanged
        If RbtnDynamic.Checked Then
            Me.Panel_Static.Visible = False
            Me.F_blnDrawDynamic = True
        Else
            Me.Panel_Static.Visible = True
            Me.F_blnDrawDynamic = False
        End If

    End Sub

    Private F_WorkingStage As List(Of clsData_WorkingStage)
    Private Sub ComboBox_WorkingStage_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_WorkingStage.SelectedIndexChanged
        Me.F_WorkingStage = Nothing
        Dim lstItem As LstbxDisplayAndItem = ComboBox_WorkingStage.SelectedItem
        If lstItem IsNot Nothing Then
            If Not lstItem.Value.Equals(LstbxDisplayAndItem.NothingInListBox.None) Then
                F_WorkingStage = DirectCast(lstItem.Value, List(Of clsData_WorkingStage))
            End If
        End If
    End Sub

    ''' <summary>
    ''' 选择进行绘图的数据工作簿
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ComboBoxOpenedWorkbook_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxOpenedWorkbook.SelectedIndexChanged
        Dim lst As LstbxDisplayAndItem = Me.ComboBoxOpenedWorkbook.SelectedItem
        Try
            Dim Wkbk As Workbook = DirectCast(lst.Value, Workbook)
            Me.F_wkbkData = Wkbk
            With APPLICATION_MAINFORM.MainForm.StatusLabel1
                .Visible = True
                .Text = Wkbk.FullName
            End With
            RaiseEvent DataWorkbookChanged(Me.F_wkbkData)
        Catch ex As Exception
            Debug.Print("选择数据工作簿出错")
        End Try
    End Sub
#End Region

#Region "  ---  关联组合列表框中的数据"

    Private Sub RefreshCombox_WorkingStage(ByVal NewWorkingStage As Dictionary(Of String, List(Of clsData_WorkingStage)))
        If NewWorkingStage IsNot Nothing Then
            With NewWorkingStage
                Try
                    Dim RegionNames = .Keys
                    Dim WorkingStages = .Values
                    Dim TagsCount As Integer = .Count
                    Dim TagsList(0 To TagsCount) As LstbxDisplayAndItem
                    TagsList(0) = New LstbxDisplayAndItem("无", LstbxDisplayAndItem.NothingInListBox.None)
                    '
                    For i1 As Integer = 0 To TagsCount - 1
                        TagsList(i1 + 1) = New LstbxDisplayAndItem(RegionNames(i1), WorkingStages(i1))
                    Next
                    Call RefreshCombobox(Me.ComboBox_WorkingStage, TagsList)
                Catch ex As Exception

                End Try
            End With
        End If
    End Sub

#End Region


#Region "  ---  窗口的激活去取消激活"

    Private Sub frmDrawing_Mnt_Incline_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        With APPLICATION_MAINFORM.MainForm.StatusLabel1
            If Me.F_wkbkData IsNot Nothing Then
                .Visible = True
                .Text = Me.F_wkbkData.FullName
            End If

        End With
    End Sub

    Private Sub frmDrawing_Mnt_Incline_Deactivate(sender As Object, e As EventArgs) Handles Me.Deactivate
        With APPLICATION_MAINFORM.MainForm.StatusLabel1
            .Visible = False
        End With
    End Sub
#End Region

#End Region

End Class
