Imports Microsoft.Office.Interop.Excel
Imports Office = Microsoft.Office.Core
Imports DAMIE.Miscellaneous
Imports DAMIE.Constants
Imports DAMIE.DataBase
Imports DAMIE.Constants.Data_Drawing_Format
Imports DAMIE.All_Drawings_In_Application
Imports DAMIE.All_Drawings_In_Application.ClsDrawing_Mnt_MaxMinDepth
Imports DAMIE.GlobalApp_Form
Imports DAMIE.Miscellaneous.ExcelFunction
''' <summary>
''' 绘制测斜数据的曲线图
''' </summary>
''' <remarks></remarks>
Public Class frmDrawing_Mnt_Incline

#Region "  ---  Declarations and Definitions"

    ''' <summary>
    ''' 当前进行绘图的数据工作簿发生变化时触发
    ''' </summary>
    ''' <remarks></remarks>
    Private Event DataWorkbookChanged(ByVal WorkingDataWorkbook As Workbook)

#Region "  ---  Fields"
    '
    Dim F_ExcelAppDrawing As Application
    ''' <summary>
    ''' 当前用于绘图的数据工作簿
    ''' </summary>
    ''' <remarks></remarks>
    Dim F_wkbkData As Workbook
    '
    Dim F_shtMonitorData As Worksheet
    Dim F_shtDrawing As Worksheet

    ''' <summary>
    ''' 在工作表中要进行处理的所有数据的范围。
    ''' 包括第一列，但是不包括第一行的日期。
    ''' </summary>
    ''' <remarks></remarks>
    Dim F_Data_UsedRange As Range

    ''' <summary>
    '''工作表中的每一天对应的数据列,
    ''' 以日期索引当天数据在工作表中的列号
    ''' </summary>
    ''' <remarks></remarks>
    Dim F_dicDate_ColNum As New Dictionary(Of Date, Integer)

    ''' <summary>
    ''' 绘图的Chart
    ''' </summary>
    ''' <remarks></remarks>
    Dim F_myChart As Chart

    Dim F_textbox_Info As TextFrame2

    ''' <summary>
    ''' Chart中第一条监测曲线所对应的相关信息
    ''' </summary>
    ''' <remarks></remarks>
    Private F_TheFirstseriesTag As ClsDrawing_Mnt_Incline.SeriesTag_Incline

    ''' <summary>
    ''' 所绘图的监测数据类型
    ''' </summary>
    ''' <remarks></remarks>
    Private F_MonitorType As MntType

    ''' <summary>
    ''' 是否是要绘制测斜数据的位移最值及对应的深度的曲线图，而不是测斜的动态图
    ''' </summary>
    ''' <remarks></remarks>
    Private F_blnMax_Depth As Boolean
    '
    Private F_GlobalApp As GlobalApplication = GlobalApplication.Application
    Private F_MainForm As APPLICATION_MAINFORM = APPLICATION_MAINFORM.MainForm
#End Region

#End Region

#Region "  ---  窗口的加载与关闭"

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        AddHandler ClsData_DataBase.dic_IDtoComponentsChanged, AddressOf Me.RefreshComobox_ExcavationID
        AddHandler ClsData_DataBase.ProcessRangeChanged, AddressOf Me.RefreshComobox_ProcessRange
        AddHandler ClsData_DataBase.WorkingStageChanged, AddressOf Me.RefreshCombox_WorkingStage
        ' ------------ 设置控件的默认属性
        btnGenerate.Enabled = False
        chkBoxOpenNewExcel.Checked = True

        ' ------------ 设置默认的监测数据类型
        Call SetMonitorType(Me.ComboBox_MntType)
        Me.F_MonitorType = MntType.Incline
        Me.ComboBox_MntType.Enabled = False
        Me.Label_MntType.Enabled = False
        ' ------------ 
        With Me.ComboBoxOpenedWorkbook
            .DisplayMember = LstbxDisplayAndItem.DisplayMember
            .ValueMember = LstbxDisplayAndItem.ValueMember
        End With

    End Sub



    '在关闭窗口时将其隐藏
    Private Sub frmDrawing_Mnt_Incline_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        '如果是子窗口自己要关闭，则将其隐藏
        '如果是mdi父窗口要关闭，则不隐藏，而由父窗口去结束整个进程
        If Not e.CloseReason = CloseReason.MdiFormClosing Then
            Me.Hide()
        End If
        e.Cancel = True
    End Sub

    Private Sub frmDrawing_Mnt_Incline_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
        RemoveHandler ClsData_DataBase.dic_IDtoComponentsChanged, AddressOf Me.RefreshComobox_ExcavationID
        RemoveHandler ClsData_DataBase.ProcessRangeChanged, AddressOf Me.RefreshComobox_ProcessRange
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
        If Me.F_shtMonitorData IsNot Nothing Then

            '开始绘图
            With Me.BGWK_NewDrawing
                If Not .IsBusy Then
                    Dim blnNewExcelApp As Boolean
                    '用来判断是否要创建新的Excel，以及是否要对新画布所在的Excel进行美化。
                    blnNewExcelApp = F_GlobalApp.MntDrawing_ExcelApps.Count = 0 Or chkBoxOpenNewExcel.Checked
                    '
                    Call .RunWorkerAsync({blnNewExcelApp})
                End If
            End With
        Else
            MessageBox.Show("请选择一个监测数据的工作表", "Warning", MessageBoxButtons.OK)
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
        Dim blnNewExcelApp As Boolean = CType(e.Argument(0), Boolean)
        '主程序界面的进度条的UI显示
        F_MainForm.ShowProgressBar_Marquee()
        '执行具体的绘图操作
        If F_blnMax_Depth Then
            Try
                Call Generate_Max_Depth(sheetMonitorData:=Me.F_shtMonitorData,
                              NewExcelApp:=blnNewExcelApp)
            Catch ex As Exception
                MessageBox.Show("绘制最大值走势图失败！" & vbCrLf & ex.Message & vbCrLf & "报错位置：" & ex.StackTrace, "Error", MessageBoxButtons.OK)
            End Try
        Else
            Try
                Call Generate_DynamicDrawing(sheetMonitorData:=Me.F_shtMonitorData,
                              Components:=Me.F_Components,
                              ProcessRegionData:=Me.F_ProcessRegionData,
                              NewExcelApp:=blnNewExcelApp)
            Catch ex As Exception
                MessageBox.Show("绘制监测曲线图失败！" & vbCrLf & ex.Message & vbCrLf & "报错位置：" & ex.StackTrace, "Error", MessageBoxButtons.OK)
            End Try
        End If
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
    ''' 1.绘制动态滚动图
    ''' </summary>
    ''' <param name="sheetMonitorData">测斜曲线的数据所在的工作表</param>
    ''' <param name="Components">与基坑ID相关的信息：结构项目与对应标高</param>
    ''' <param name="ProcessRegionData">与此监测数据所对应的基坑的施工进度</param>
    ''' <param name="NewExcelApp">用来判断是否要创建新的Excel，以及是否要对新画布所在的Excel进行美化。</param>
    ''' <remarks></remarks>
    Private Sub Generate_DynamicDrawing(ByVal sheetMonitorData As Worksheet,
                  ByVal Components As Component(),
                  ByVal ProcessRegionData As clsData_ProcessRegionData,
                  ByVal NewExcelApp As Boolean)

        Dim myExcelForMonitorDrawing As Cls_ExcelForMonitorDrawing = Nothing
        '   --------------- 获取用来绘图的Excel程序，并将此界面加入主程序的监测曲线集合 -------------------
        Call PrePare(sheetMonitorData, NewExcelApp, myExcelForMonitorDrawing)
        '   ---------------
        '对第一条监测曲线的施工日期进行初始同仁，用来绘制第一条监测曲线
        Dim Date_theFirstCurve As Date = F_dicDate_ColNum.Keys(0)
        'Data_UsedRange包括第一列，但是不包括第一行的日期。
        Me.F_myChart = DrawDynamicChart(Me.F_shtDrawing, Date_theFirstCurve, F_Data_UsedRange)

        '设置监测曲线的时间跨度
        Dim DtSp As DateSpan = GetDateSpan(Me.F_dicDate_ColNum)



        '------------------------------------ 将执行结果传递给mainform的属性中
        '设置图表的Tag属性
        Dim Tags As MonitorInfo = GetChartTags(Me.F_shtMonitorData)
        Dim moniSheet As New ClsDrawing_Mnt_Incline(F_shtMonitorData, F_myChart,
                                                    myExcelForMonitorDrawing, DtSp,
                                                    DrawingType.Monitor_Incline_Dynamic, True, F_textbox_Info, Tags, Me.F_MonitorType,
                                                    F_dicDate_ColNum, F_Data_UsedRange,
                                                    F_TheFirstseriesTag, ProcessRegionData)

        '--------------标高箭头-------------------- 在监测曲线图中绘制标高箭头
        If Components IsNot Nothing Then
            Call ComponentsAndElevatins(moniSheet.InclineTopElevaion, Components)
        End If

        ' ------------------------------------------------------------------------------------------------
        ' ----------动态开挖深度的直线与文本框-------- 根据选择确定是否要绘制动态开挖深度的直线与文本框
        If ProcessRegionData IsNot Nothing Then
            '将表示挖深的直线与文本框赋值给moniSheet的ExcavationDepth属性，以供后面在滚动时进行移动
            Call RollingDepth_lineAndtextbox(ProcessRegionData, Date_theFirstCurve, moniSheet.ShowLabelsWhileRolling, moniSheet.InclineTopElevaion)
        End If

        '扩展mainForm.TimeSpan的区间
        Call GlobalApplication.Application.refreshGlobalDateSpan(moniSheet.DateSpan)
        '-------- 界面显示与美化
        Call DrawingFinished(NewExcelApp)

    End Sub

    ''' <summary>
    ''' 2.绘制最大值的走势图
    ''' </summary>
    ''' <param name="sheetMonitorData"></param>
    ''' <param name="NewExcelApp"></param>
    ''' <remarks></remarks>
    Private Sub Generate_Max_Depth(ByVal sheetMonitorData As Worksheet, ByVal NewExcelApp As Boolean)
        Dim myExcelForMonitorDrawing As Cls_ExcelForMonitorDrawing = Nothing
        '   --------------- 获取用来绘图的Excel程序，并将此界面加入主程序的监测曲线集合 -------------------
        Call PrePare(sheetMonitorData, NewExcelApp, myExcelForMonitorDrawing)
        '   ---------------
        '以每一天的日期索引当天的测斜位移的极值与对应的深度
        Dim DMMD As DateMaxMinDepth

        DMMD = getMaxMinDepth(Me.F_dicDate_ColNum, F_Data_UsedRange)
        If DMMD Is Nothing Then
            Exit Sub
        End If
        '
        Dim cht As Chart = DrawDMMDChart(Me.F_shtDrawing, DMMD)
        ' ------------------------------------------------------------------------------------------------
        '设置图表的Tag属性
        Dim Tags As MonitorInfo = GetChartTags(Me.F_shtMonitorData)
        Dim Drawing_MMD As New ClsDrawing_Mnt_MaxMinDepth(Me.F_shtMonitorData, cht, myExcelForMonitorDrawing,
                                                DrawingType.Monitor_Incline_MaxMinDepth,
                                                False, Me.F_textbox_Info, Tags, Me.F_MonitorType,
                                                DMMD.ConstructionDate, DMMD)
        If Me.F_WorkingStage IsNot Nothing Then
            Call DrawWorkingStage(cht, Me.F_WorkingStage)
        End If
        '-------- 界面显示与美化
        Call DrawingFinished(NewExcelApp)
    End Sub

#Region "  ---  动态图"

    ''' <summary>
    ''' 工作表中的每一天对应的数据列，以日期索引当天数据在工作表中的列号。
    ''' 另外，在此函数中，对于每一个测点的工作表中的数据格式进行判断，其中第一行数据应该为日期类型，而第一列数据应该为表示开挖深度的Double类型。
    ''' 如果表头的数据格式不对，则会弹出警告，并将有效的数据限制在未出错前的范围内。
    ''' </summary>
    ''' <param name="shtMonitorData"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function getdicDate_ColNum(ByVal shtMonitorData As Worksheet) As Dictionary(Of Date, Integer)
        Dim dic_Date_ColNum As New Dictionary(Of Date, Integer)
        Dim UsedRange_shtData As Range = shtMonitorData.UsedRange
        Dim rowsCount As Integer = UsedRange_shtData.Rows.Count
        Dim colsCount As Integer = UsedRange_shtData.Columns.Count
        '
        Dim endRow As Integer, endCol As Integer
        '--------------------------------------------------------------- 找出测斜数据DataRange的范围
        '数据区域（包含x轴的深度数据）的起始单元格为“A3”
        With shtMonitorData
            Dim iRowNum As Integer = Mnt_Incline.RowNum_FirstData_WithoutDate
            '表示测点深度的ID列
            '最末尾一行取用rowsCount是基于工作表的第一行有数据的情况。
            Dim rgDepth As Range = .Range(.Cells(Mnt_Incline.RowNum_FirstData_WithoutDate, Mnt_Incline.ColNum_Depth),
                                       .Cells(rowsCount, Mnt_Incline.ColNum_Depth))
            Dim vDepth As Object(,) = rgDepth.Value
            'ReDim F_arrDepth(0 To rowsCount - Mnt_Incline.RowNum_FirstData_WithoutDate)

            For Each v As Object In vDepth
                If v IsNot Nothing Then
                    Try
                        '尝试将第一列的深度数据转换为Double
                        Dim Depth As Single = CSng(v)
                        '如果成功转换，说明此深度数据有效，否则，说明到了最后一行的深度值
                        'Me.F_arrDepth(iRowNum - Mnt_Incline.RowNum_FirstData_WithoutDate) = Depth
                        iRowNum += 1
                    Catch ex As Exception
                        MessageBox.Show("第" & Mnt_Incline.ColNum_Depth & "列的深度数据无法转换为数值，请检查第" & iRowNum.ToString & "行。",
                                        "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        Exit For
                    End Try
                Else
                    MessageBox.Show("第" & Mnt_Incline.ColNum_Depth & "列的深度数据为空，请检查第" & iRowNum.ToString & "行。",
                                      "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Exit For
                End If
            Next
            endRow = iRowNum - 1
            '------------------ 
            Dim iColNum As Integer = Mnt_Incline.ColNum_FirstData_Displacement
            '表示施工日期的表头行
            '最末尾一列取用colsCount是基于工作表的第一列有数据的情况。
            Dim rgDate As Range = .Range(.Cells(Mnt_Incline.RowNumForDate, Mnt_Incline.ColNum_FirstData_Displacement),
                                       .Cells(Mnt_Incline.RowNumForDate, colsCount))
            Dim vDate As Object(,) = rgDate.Value
            For Each v As Object In vDate
                If v IsNot Nothing Then
                    Try
                        '尝试将第一行的日期数据转换为Date
                        Dim ConstructionDate As Date = CDate(v)
                        '创建字典，以监测数据的日期key索引数据所在列号item
                        dic_Date_ColNum.Add(ConstructionDate, iColNum)
                        '如果成功转换，说明此深度数据有效，否则，说明到了最后一行的深度值
                        iColNum += 1
                    Catch ex As Exception
                        MessageBox.Show("第" & Mnt_Incline.RowNumForDate & "行的日期字段的格式不正确，请检查第" & ConvertColumnNumberToString(iColNum) & "列。",
                                        "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        Exit For
                    End Try
                Else
                    MessageBox.Show("第" & Mnt_Incline.RowNumForDate & "行的日期数据为空，请检查第" & ConvertColumnNumberToString(iColNum) & "列。",
                                      "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Exit For
                End If
            Next
            endCol = iColNum - 1
            ' 
            UsedRange_shtData = .Range(.Cells(Mnt_Incline.RowNum_FirstData_WithoutDate,
                                              Mnt_Incline.ColNum_Depth), .Cells(endRow, endCol))
            'Data_UsedRange包括第一列，但是不包括第一行的日期。
            F_Data_UsedRange = UsedRange_shtData
        End With
        '
        '-------------------------------------
        Return dic_Date_ColNum
    End Function

    ''' <summary>
    ''' 绘制测斜曲线图
    ''' </summary>
    ''' <param name="DateOfCurve">绘制的曲线所对应的施工日期</param>
    ''' <param name="dataUsedRange">包括第一列，但是不包括第一行的日期。</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function DrawDynamicChart(ByVal DrawingSheet As Worksheet, ByVal DateOfCurve As Date, ByVal dataUsedRange As Range) As Chart
        Dim TemplatePath As String = System.IO.Path.Combine(My.Settings.Path_Template, Constants.FolderOrFileName.File_Template.Chart_Incline)
        Dim cht As Chart = DrawChart(DrawingSheet, True, TemplatePath:=TemplatePath)
        '------------------------ 设置集合中的第一条曲线的数据
        Dim series As Series = cht.SeriesCollection(LowIndexOfObjectsInExcel.SeriesInSeriesCollection)

        '设置数据列的数据值
        With series
            .Name = F_shtMonitorData.Cells(Mnt_Incline.RowNumForDate, Mnt_Incline.ColNum_FirstData_Displacement).Value        '系列名称
            .XValues = dataUsedRange.Columns(Mnt_Incline.ColNum_FirstData_Displacement).Value                      'X轴的数据
            .Values = dataUsedRange.Columns(Mnt_Incline.ColNum_Depth).value                       'Y轴的数据
        End With

        '初始赋值
        Me.F_TheFirstseriesTag = New ClsDrawing_Mnt_Incline.SeriesTag_Incline(series, DateOfCurve)

        '测斜数据区的数组，即不包含dataRange中的第1列以外的区域的大数组。
        Dim arrDataDisplacement As Object
        With F_shtMonitorData
            arrDataDisplacement = .Range(.Cells(Mnt_Incline.RowNum_FirstData_WithoutDate, Mnt_Incline.ColNum_FirstData_Displacement),
                                .Cells(dataUsedRange.Rows.Count, dataUsedRange.Columns.Count)) '.Value
        End With

        '------------------------ 设置X轴的格式
        With cht.Axes(XlAxisType.xlCategory)

            '由数据的最小与最大值来划分表格区间
            Dim imax = F_ExcelAppDrawing.WorksheetFunction.Max(arrDataDisplacement)
            Dim iMin = F_ExcelAppDrawing.WorksheetFunction.Min(arrDataDisplacement)

            '主要与次要刻度单位，先确定刻度单位是为了后面将坐标轴的区间设置为主要刻度单位的倍数
            Dim unit As Single = CSng(Format((imax - iMin) / Drawing_Incline.AxisMajorUnit_Y, "0.0E+00")) '这里涉及到有效数字的处理的问题
            .MajorUnit = unit
            .MinorUnitIsAuto = True

            '坐标轴上显示的总区间
            .MinimumScale = F_ExcelAppDrawing.WorksheetFunction.Floor_Precise(iMin, .MajorUnit)
            .MaximumScale = F_ExcelAppDrawing.WorksheetFunction.Ceiling_Precise(imax, .MajorUnit)

            '坐标轴标题
            .AxisTitle.Text = GetAxisLabel(DrawingType.Monitor_Incline_Dynamic, Me.F_MonitorType, XlAxisType.xlCategory)
        End With

        '------------------------- 设置Y轴的格式
        With cht.Axes(XlAxisType.xlValue)
            .MajorUnit = Drawing_Incline.AxisMajorUnit_Y
            .MinorUnitIsAuto = True
            Dim arr = dataUsedRange.Columns(1).value
            Dim imin = F_ExcelAppDrawing.WorksheetFunction.Min(dataUsedRange.Columns(Mnt_Incline.ColNum_Depth))
            .MinimumScale = F_ExcelAppDrawing.WorksheetFunction.Floor_Precise(imin, .MajorUnit)
            '
            Dim imax = F_ExcelAppDrawing.WorksheetFunction.Max(dataUsedRange.Columns(Mnt_Incline.ColNum_Depth))
            .MaximumScale = F_ExcelAppDrawing.WorksheetFunction.Ceiling_Precise(imax, .MajorUnit)
            '
            .AxisTitle.Text = GetAxisLabel(DrawingType.Monitor_Incline_Dynamic, Me.F_MonitorType, XlAxisType.xlValue)
        End With
        Return cht
    End Function

    ''' <summary>
    ''' 由选择的基坑区域绘制对应的挖深直线与文本框
    ''' </summary>
    '''  <param name="ProcessRegion">此基坑区域所在的列的Range对象(包括前面几行的表头数据)</param>
    '''  <param name="ShowLabelsWhileRolling">指示是否要在进行滚动时指示开挖标高的标识线旁给出文字说明，比如“开挖标高”等。</param>
    '''  <param name="inclineTopElev"> 测斜管顶部的标高值 </param>
    ''' <remarks>此挖深直线与文本框用来进行后期滚动时显示每一天的开挖情况之用的。</remarks>
    Private Sub RollingDepth_lineAndtextbox(ByVal ProcessRegion As clsData_ProcessRegionData,
                                               ByVal Date_theFirstCurve As Date, ByVal ShowLabelsWhileRolling As Boolean, inclineTopElev As Single)
        ' ---------------- 绘制表示挖深的直线与文本框
        ''对第一条监测曲线的施工日期进行初始同仁，用来绘制第一条监测曲线
        'Me.F_Date_theFirstCurve = F_dicDate_ColNum.Keys(0)
        Dim excavElev As Single
        Try
            excavElev = ProcessRegion.Date_Elevation.Item(Date_theFirstCurve)
        Catch ex As KeyNotFoundException
            Dim ClosestDate As Date = ClsData_DataBase.FindTheClosestDateInSortedList(ProcessRegion.Date_Elevation.Keys, Date_theFirstCurve)
            excavElev = ProcessRegion.Date_Elevation.Item(ClosestDate)
        End Try
        '
        '根据每一个标高值，画出相应的深度线及文本框
        ' ---------------------- 绘制直线 与 设置直线格式 ----------------------------------------------------

        '直线的基本几何参数
        Dim linetop = ExcelFunction.GetPositionInChartByValue(F_myChart.Axes(XlAxisType.xlValue), inclineTopElev - excavElev)
        '
        Dim lineLeft As Single
        Dim lineWidth As Single
        Dim plotA As PlotArea = F_myChart.PlotArea
        With plotA
            lineLeft = .InsideLeft + 0.2 * .InsideWidth
            lineWidth = 0.12 * .InsideWidth     '水平线条的长度
        End With
        '
        '绘制直线与文本框
        Dim Line As Shape = F_myChart.Shapes.AddLine(BeginX:=lineLeft,
                                       BeginY:=linetop,
                                       EndX:=lineLeft + lineWidth,
                                       EndY:=linetop)
        ' -- 设置直线格式 --
        With Line.Line
            .ForeColor.RGB = RGB(255, 0, 0)
            .Weight = 1.5
            .EndArrowheadStyle = Office.MsoArrowheadStyle.msoArrowheadStealth
            .EndArrowheadLength = Office.MsoArrowheadLength.msoArrowheadLong
            .EndArrowheadWidth = Office.MsoArrowheadWidth.msoArrowheadWidthMedium
        End With
        ' ---------------------- 绘制文本框 与 设置文本框格式 ----------------------------------------------------
        Dim Textbox As Shape = Nothing
        If ShowLabelsWhileRolling Then

            Textbox = F_myChart.Shapes.AddTextbox(Orientation:=Office.MsoTextOrientation.msoTextOrientationHorizontal,
                                                Left:=plotA.InsideLeft,
                                                Top:=linetop - 10,
                                                Width:=lineLeft - plotA.InsideLeft + 100,
                                                Height:=20)

            Textbox.TextFrame2.AutoSize = Microsoft.Office.Core.MsoAutoSize.msoAutoSizeShapeToFitText
            With Textbox
                .Left = Line.Left
                .Top = Line.Top - .Height
                With Textbox.TextFrame2.TextRange
                    '文本框中的文本
                    .Text = "挖深" '& vbTab & ClassData_DataBase.cstDepthRefer
                    .Font.Size = 12
                    .Font.Name = AMEApplication.FontName_TNR
                    .Font.Bold = Microsoft.Office.Core.MsoTriState.msoCTrue
                    .Font.Fill.ForeColor.RGB = RGB(255, 0, 0)
                End With
            End With
        End If
        ' ----------------   为数据列添加信息
        With F_TheFirstseriesTag
            .DepthLine = Line
            .DepthTextbox = Textbox
        End With
    End Sub

    ''' <summary>
    ''' 标注每个构件的标高
    ''' </summary>
    ''' <param name="inclineTopElevation"> 测斜管顶端的绝对标高值 </param>
    ''' <param name="ComponentsOfExcationID">记录基坑ID中对应的每一个构件项目与其对应标高的数组,
    ''' 数组中的第一列表示构件项目的名称，第二列表示构件项目的标高值</param>
    ''' ''' <remarks></remarks>
    Private Sub ComponentsAndElevatins(inclineTopElevation As Single, ByVal ComponentsOfExcationID As Component()) ', ByVal ExcavID As String, ByVal IDtoElevationAndData As Dictionary(Of String, ClsData_DataBase.clsExcavationID)
        'ComponentsOfExcationID
        'Dim arrexcavation(,) As String = IDtoElevationAndData.Item(ExcavID).Components
        '定义标高线的长度范围
        Dim plotA As PlotArea

        Dim lst_Name_LinesAndTextBox As New List(Of String)
        plotA = F_myChart.PlotArea

        '开始对此基坑ID对应的[标高项，标高值]中的每一项，用直线标注其位置，用文本框显示其标高值或其他信息
        For i As Byte = 0 To UBound(ComponentsOfExcationID, 1)
            Dim Component As Component = ComponentsOfExcationID(i)
            Dim Elevation As Single = Component.Elevation
            '基准标高，即将标高数据换算为相对于地面标高的深度数据
            Dim relativedepth As Single = inclineTopElevation - Elevation
            If relativedepth >= 0 Then
                If Component.Type = ComponentType.Strut OrElse Component.Type = ComponentType.TopOfBottomSlab Then
                    '------------------------------------ 绘制对应标高项的标高线与文本框
                    Dim Line As Shape = Nothing
                    Dim Textbox As Shape = Nothing
                    '-------------------------------
                    ' 根据每一个标高值，画出相应的深度线及文本框

                    Call DrawDepthLineAndTextBox(myChart:=F_myChart, relativeDepth:=relativedepth, Line:=Line, ShowLabelsWhileRolling:=True, Textbox:=Textbox)

                    '文本框中的文本
                    Textbox.TextFrame2.TextRange.Text = Component.Description & " (" & Elevation.ToString & ")"
                    '-------------------------------
                    lst_Name_LinesAndTextBox.Add(Line.Name)
                    lst_Name_LinesAndTextBox.Add(Textbox.Name)
                End If
            End If
        Next        '下一个标高项

        Try     '将这些构件与标高的图形对象组合为一个组。
            '可能会由于数组中的图形小于两个，而出现不能执行Group的错误。
            F_myChart.Shapes.Range(lst_Name_LinesAndTextBox.ToArray).Group()
        Catch ex As Exception
            Debug.Print("将形状进行组合时出错。异常原因为:{0}" & vbCrLf &
                        "，报错位置：为 {1}。", ex.Message, ex.TargetSite.Name)
        End Try
    End Sub

    ''' <summary>
    ''' 根据每一个构件的标高值，画出相应的深度线及文本框
    ''' </summary>
    ''' <param name="myChart">进行绘图的图表对象</param>
    ''' <param name="relativeDepth"> 构件相对于测斜管顶部的深度值</param>
    '''  <param name="ShowLabelsWhileRolling">指示是否要在进行滚动时指示开挖标高的标识线旁给出文字说明，比如“开挖标高”等。</param>
    '''  <param name="Line">要返回的Line对象</param>
    '''  <param name="Textbox">要返回的文本框Textbox对象，如果ShowLabelsWhileRolline的值为False，则其返回Nothing</param>
    ''' <remarks></remarks>
    Private Sub DrawDepthLineAndTextBox(ByVal myChart As Chart, ByVal relativeDepth As Single, ByRef Line As Shape,
                                             Optional ByVal ShowLabelsWhileRolling As Boolean = True,
                                             Optional ByRef Textbox As Shape = Nothing)
        '直线与文本框的基本几何参数

        '
        Dim linetop = ExcelFunction.GetPositionInChartByValue(myChart.Axes(XlAxisType.xlValue), relativeDepth)

        '---------------------------- 绘制直线 设置直线格式 ----------------------------------
        With myChart.ChartArea
            Line = myChart.Shapes.AddLine(BeginX:= .Left + .Width,
                                           BeginY:=linetop,
                                           EndX:= .Left + .Width - 50,
                                           EndY:=linetop)  'EndX中的 -50 即为水平直线的长度

            ' -- 设置直线格式 --
            With Line.Line
                .ForeColor.RGB = RGB(0, 0, 255)
                .Weight = 1.5
                .EndArrowheadStyle = Office.MsoArrowheadStyle.msoArrowheadStealth
                .EndArrowheadLength = Office.MsoArrowheadLength.msoArrowheadLong
                .EndArrowheadWidth = Office.MsoArrowheadWidth.msoArrowheadWidthMedium
            End With
        End With
        '--------------------------- 绘制文本框 设置文本框格式----------------------------
        If ShowLabelsWhileRolling Then
            With myChart.ChartArea
                Textbox = myChart.Shapes.AddTextbox(Orientation:=Office.MsoTextOrientation.msoTextOrientationHorizontal,
                                                     Left:= .Left + .Width - 100,
                                                     Top:=linetop - 15,
                                                     Width:=100,
                                                     Height:=20)    '文本框的宽度为100像素
            End With
            ' ------ 设置文本框格式
            With Textbox.TextFrame2
                .AutoSize = Microsoft.Office.Core.MsoAutoSize.msoAutoSizeShapeToFitText
                '文字位于文本框的右上角
                .VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorTop
                .TextRange.ParagraphFormat.Alignment = Microsoft.Office.Core.MsoParagraphAlignment.msoAlignRight
                .MarginRight = 2    '文本距离文本框右边界的长度，单位为像素
                With .TextRange.Font
                    '文本框中的文本
                    .Size = 9
                    .Name = AMEApplication.FontName_TNR
                    .Bold = True
                    .Fill.ForeColor.RGB = RGB(0, 0, 255)
                End With
            End With
        End If
    End Sub

#End Region

#Region "  ---  最值图"

    ''' <summary>
    ''' 根据工作表中有效的日期范围以及对应的数据范围，得到每一天中，某测点的位移极值和对应的深度
    ''' </summary>
    ''' <param name="dic">以日期索引当天数据在工作表中的列号</param>
    ''' <param name="UR">包括第一列的深度，但是不包括第一行是施工日期</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function getMaxMinDepth(ByVal dic As Dictionary(Of Date, Integer), ByVal UR As Range) _
        As DateMaxMinDepth
        Dim count As Integer = dic.Count
        '
        Dim arrDate(0 To count - 1) As Double
        Dim arrMax(0 To count - 1) As Object
        Dim arrMin(0 To count - 1) As Object
        Dim arrDepth_Max(0 To count - 1) As Object
        Dim arrDepth_Min(0 To count - 1) As Object
        '
        Dim shtData As Worksheet = UR.Worksheet

        Dim arrDepth As Single() = ExcelFunction.ConvertRangeDataToVector(Of Single)(UR.Columns(1))
        '
        Dim index As Integer = 0
        Dim Day_Data As Date
        Try
            For Each Day_Data In dic.Keys
                Dim arrData As Object() = ExcelFunction.ConvertRangeDataToVector(Of Object)(UR.Columns(dic.Item(Day_Data)))
                With shtData.Application.WorksheetFunction
                    '
                    Dim max As Object = .Max(arrData)
                    Dim min As Object = .Min(arrData)


                    '
                    Dim Row_Max As Integer = .Match(max, arrData, 0)
                    Dim Row_Min As Integer = .Match(min, arrData, 0)
                    ' 

                    Dim depth_Max As Object = arrDepth(Row_Max - 1)
                    Dim depth_Min As Object = arrDepth(Row_Min - 1)
                    ' --------------- 赋值 ---------------
                    arrDate(index) = Day_Data.ToOADate
                    arrMax(index) = max
                    arrMin(index) = min
                    arrDepth_Max(index) = depth_Max
                    arrDepth_Min(index) = depth_Min
                    ' -----------------------------
                    index += 1
                End With
            Next
            Return New DateMaxMinDepth(arrDate, arrMax, arrMin, arrDepth_Max, arrDepth_Min)
        Catch ex As Exception
            MessageBox.Show("提取监测数据工作表中的位移最值及其对应的深度值出错。" &
                            vbCrLf & "出错的日期为：" & Day_Data.ToShortDateString &
                            vbCrLf & ex.Message &
                            vbCrLf & "报错位置：" & ex.TargetSite.Name,
                "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' 绘制图表
    ''' </summary>
    ''' <param name="DrawingSheet"></param>
    ''' <param name="DMMD"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function DrawDMMDChart(ByVal DrawingSheet As Worksheet, ByVal DMMD As DateMaxMinDepth) As Chart
        Dim TemplatePath As String = System.IO.Path.Combine(My.Settings.Path_Template, Constants.FolderOrFileName.File_Template.Chart_Max_Depth)
        Dim cht As Chart = DrawChart(DrawingSheet, True, TemplatePath:=TemplatePath)
        '下面这一句激活语句非常重要，如果不对ChartObject进行激活，那么下面的Chart.SetElement就会失效，
        '从而导致图表中相应的元素不存在，那么在对元素的格式进行设置的时候就会报错。
        cht.Parent.Activate()
        '
        Dim arrDate() As Double
        Dim arrMax() As Object
        Dim arrMin() As Object
        Dim arrDepth_Max() As Object
        Dim arrDepth_Min() As Object
        With DMMD
            arrDate = .ConstructionDate
            arrMax = .Max
            arrMin = .Min
            arrDepth_Max = .Depth_Max
            arrDepth_Min = .Depth_Min
        End With
        '
        '-------------------------------------------------------------------------
        Dim SC As SeriesCollection = cht.SeriesCollection
        '保证图表中有四条数据系列
        If SC.Count < 4 Then
            For i = 0 To 4 - SC.Count - 1
                SC.NewSeries()
            Next
        End If
        '-------------------------------------------------------------------------
        With cht
            .SetElement(Microsoft.Office.Core.MsoChartElementType.msoElementPrimaryCategoryAxisShow)
            .SetElement(Microsoft.Office.Core.MsoChartElementType.msoElementSecondaryValueAxisTitleRotated)
            .SetElement(Microsoft.Office.Core.MsoChartElementType.msoElementPrimaryValueAxisTitleRotated)
        End With
        '-------------------------------------------------------------------------
        '四条曲线的数据及所在的Y轴类型
        With SC.Item(1)
            .Name = Drawing_Incline_DMMD.SeriesName_Max
            .XValues = arrDate
            .Values = arrMax
            'Dim b = .XValues     'Date类型的数组传递给XValue后，便成了String类型的数组
            .AxisGroup = XlAxisGroup.xlPrimary
        End With
        With SC.Item(2)
            .Name = Drawing_Incline_DMMD.SeriesName_Min
            .XValues = arrDate
            .Values = arrMin
            .AxisGroup = XlAxisGroup.xlPrimary
        End With
        With SC.Item(3)
            .Name = Drawing_Incline_DMMD.SeriesName_Depth_Max
            .XValues = arrDate
            .Values = arrDepth_Max
            .AxisGroup = XlAxisGroup.xlSecondary
        End With
        With SC.Item(4)
            .Name = Drawing_Incline_DMMD.SeriesName_Depth_Min
            .XValues = arrDate
            .Values = arrDepth_Min
            .AxisGroup = XlAxisGroup.xlSecondary
        End With
        '-------------------------------------------------------------------------

        '------------------------ 设置X轴的格式：整个日期跨度
        Dim axisX As Axis = cht.Axes(XlAxisType.xlCategory)
        With axisX
            .CategoryType = XlCategoryType.xlTimeScale
            '设置X轴的时间跨度
            Dim maxScale As Double = Max_Array(Of Double)(arrDate)
            Dim minScale As Double = Min_Array(Of Double)(arrDate)
            .MinimumScale = minScale
            .MaximumScale = maxScale
            '.MaximumScaleIsAuto = True
            '.MinimumScaleIsAuto = True

            '设置竖向网格间距
            .MajorUnitIsAuto = True
            .MinorUnitIsAuto = True

            '设置坐标轴标签的位置
            .TickLabelPosition = XlTickLabelPosition.xlTickLabelPositionLow
            .TickLabels.NumberFormatLocal = "yy/m/d"
            .TickLabels.Orientation = 0 ' XlTickLabelOrientation.xlTickLabelOrientationHorizontal
            '
            .TickMarkSpacing = 10
            .AxisTitle.Text = GetAxisLabel(DrawingType.Monitor_Incline_MaxMinDepth, Me.F_MonitorType, XlAxisType.xlCategory)
        End With

        '------------------------- 设置Y主轴的格式：测斜位移的最值
        Dim axisY_Prime As Axis = cht.Axes(XlAxisType.xlValue, XlAxisGroup.xlPrimary)
        With axisY_Prime
            .MinorUnitIsAuto = True
            .MajorUnitIsAuto = True
            .MaximumScaleIsAuto = True
            .MinimumScaleIsAuto = True
            '
            'Dim imax1 = F_ExcelAppDrawing.WorksheetFunction.Max(arrMax)
            'Dim imax2 = F_ExcelAppDrawing.WorksheetFunction.Max(arrMin)
            '.MaximumScale = F_ExcelAppDrawing.WorksheetFunction.Ceiling_Precise(Math.Max(imax1, imax2), .MajorUnit)
            ''
            'Dim imin1 = F_ExcelAppDrawing.WorksheetFunction.Min(arrMax)
            'Dim imin2 = F_ExcelAppDrawing.WorksheetFunction.Min(arrMin)
            '.MinimumScale = F_ExcelAppDrawing.WorksheetFunction.Floor_Precise(Math.Min(imin1, imin2), .MajorUnit)

            .AxisTitle.Text = GetAxisLabel(DrawingType.Monitor_Incline_MaxMinDepth, Me.F_MonitorType,
                                           XlAxisType.xlValue, XlAxisGroup.xlPrimary)
        End With

        '------------------------- 设置Y次轴的格式：测斜位移的最值所对应的深度
        Dim ax As Axis = cht.Axes(XlAxisType.xlValue, XlAxisGroup.xlSecondary)
        With ax
            .MinorUnitIsAuto = True
            .MajorUnitIsAuto = True
            .MaximumScaleIsAuto = True
            .MinimumScaleIsAuto = True
            '
            'Dim imax1 = F_ExcelAppDrawing.WorksheetFunction.Max(arrMax)
            'Dim imax2 = F_ExcelAppDrawing.WorksheetFunction.Max(arrMin)
            '.MaximumScale = F_ExcelAppDrawing.WorksheetFunction.Ceiling_Precise(Math.Max(imax1, imax2), .MajorUnit)
            ''
            'Dim imin1 = F_ExcelAppDrawing.WorksheetFunction.Min(arrMax)
            'Dim imin2 = F_ExcelAppDrawing.WorksheetFunction.Min(arrMin)
            '.MinimumScale = F_ExcelAppDrawing.WorksheetFunction.Floor_Precise(Math.Min(imin1, imin2), .MajorUnit)
            .ReversePlotOrder = True
            .AxisTitle.Text = GetAxisLabel(DrawingType.Monitor_Incline_MaxMinDepth, Me.F_MonitorType,
                                           XlAxisType.xlValue, XlAxisGroup.xlSecondary)
        End With

        '-------------------------------------------------------------------------
        DrawingSheet.Cells(1, 1).Activate()         '以免图形显示时Chart中的对象被选中。
        Return cht
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
                Dim Depth As Single
                For Each WS As clsData_WorkingStage In WorkingStage
                    ' -------------------------------------------------------------------------------------------
                    Depth = Project_Expo.Elevation_GroundSurface - WS.Elevation
                    Dim EndY As Single = ExcelFunction.GetPositionInChartByValue(Cht.Axes(XlAxisType.xlValue, XlAxisGroup.xlSecondary), Depth)
                    '
                    Dim shpLine As Shape = .Shapes.AddLine(BeginX:=0, BeginY:=Cht.PlotArea.InsideTop,
                                                           EndX:=0, EndY:=EndY)

                    With shpLine.Line
                        .Weight = 1.5
                        .EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
                        .EndArrowheadLength = Microsoft.Office.Core.MsoArrowheadLength.msoArrowheadLengthMedium
                        .EndArrowheadWidth = Microsoft.Office.Core.MsoArrowheadWidth.msoArrowheadWidthMedium
                        .DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineLongDashDot
                        .ForeColor.RGB = RGB(0, 0, 0)
                    End With
                    '
                    ExcelFunction.setPositionInChart(shpLine, AX, WS.ConstructionDate.ToOADate)
                    ' -------------------------------------------------------------------------------------------
                    Dim TextWidth As Single = 25 : Dim textHeight As Single = 10
                    Dim shpText As Shape
                    With shpLine
                        shpText = Cht.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                                                                     Left:= .Left - TextWidth / 2, Top:= .Top - textHeight, Height:=textHeight, Width:=TextWidth)
                    End With
                    Call ExcelFunction.FormatTextbox_Tag(shpText.TextFrame2, Text:=WS.Description,
                                                         HorizontalAlignment:=Microsoft.Office.Core.MsoParagraphAlignment.msoAlignCenter)
                    ' -------------------------------------------------------------------------------------------
                    arrLineName(i) = shpLine.Name
                    arrTextName(i) = shpText.Name
                    i += 1
                Next
                With Cht.Shapes
                    Try         '可能会由于数组中的图形小于两个，而出现不能执行Group的错误。
                        Dim shp1 As Shape = .Range(arrLineName).Group()
                        Dim shp2 As Shape = .Range(arrTextName).Group
                        .Range({shp1.Name, shp2.Name}).Group()
                    Catch ex As Exception
                        Debug.Print("将形状进行组合时出错。异常原因为:{0}" & vbCrLf & "，报错位置：为 {1}。", ex.Message, ex.TargetSite.Name)
                    End Try
                End With
            Catch ex As Exception
                MessageBox.Show("设置开挖工况位置出现异常。" & vbCrLf & ex.Message & vbCrLf & "报错位置：" & ex.TargetSite.Name,
       "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End Try
        End With
    End Sub

#End Region

#Region "  ---  通用子方法"

    ''' <summary>
    ''' 在此方法中，分别获得了进行绘图所需要的ExcelApplicaion程序、进行绘图的工作表，进行绘图所需要的监测数据的Range，
    ''' 以及参与绘图的监测数据中，每一个有效的施工日期在数据工作表中的列号。
    ''' </summary>
    ''' <param name="sheetMonitorData"></param>
    ''' <param name="NewExcelApp"></param>
    ''' <param name="ExcelForMntDrawing"></param>
    ''' <remarks></remarks>
    Private Sub PrePare(ByVal sheetMonitorData As Worksheet, ByVal NewExcelApp As Boolean,
                        ByRef ExcelForMntDrawing As Cls_ExcelForMonitorDrawing)
        '   --------------- 获取用来绘图的Excel程序，并将此界面加入主程序的监测曲线集合 -------------------
        F_ExcelAppDrawing = GetApplication(NewExcelApp:=NewExcelApp,
                                         ExcelForMntDrawing:=ExcelForMntDrawing,
                                         MntDrawingExcelApps:=F_GlobalApp.MntDrawing_ExcelApps)
        '   ----------------------------------
        F_ExcelAppDrawing.ScreenUpdating = False

        '打开工作簿以画图
        Dim wkbkDrawing As Workbook
        If F_ExcelAppDrawing.Workbooks.Count = 0 Then
            wkbkDrawing = F_ExcelAppDrawing.Workbooks.Add
        Else
            wkbkDrawing = F_ExcelAppDrawing.Workbooks(1) '总是定义为第一个，因为就只开了一个
        End If
        '新开一个工作表以画图
        'wkbkDrawing.Worksheets(wkbkDrawing.Worksheets.Count).Delete()    '删除工作簿中的最后一个工作表。如果是直接在原界面上新开工作表以绘图，那么此时便可以删除上次绘图的工作表，以节省资源。
        F_shtDrawing = wkbkDrawing.Worksheets.Add

        '   ---------------------------------
        '-- 开始画图 ,并返回关键参数：dicDate_ColNum
        F_dicDate_ColNum = getdicDate_ColNum(sheetMonitorData)
        '   ---------------------------------
    End Sub

    ''' <summary>
    ''' 当进行绘图的数据工作簿发生变化时触发
    ''' </summary>
    ''' <param name="WorkingDataWorkbook">要进行绘图的数据工作簿</param>
    ''' <remarks></remarks>
    Private Sub frmDrawing_Mnt_Incline_DataWorkbookChanged(WorkingDataWorkbook As Workbook) Handles Me.DataWorkbookChanged
        '在列表中显示出监测数据工作簿中的所有工作表
        Dim sheetsCount As Byte = WorkingDataWorkbook.Worksheets.Count
        If sheetsCount > 0 Then
            Dim arrSheetsName(0 To sheetsCount - 1) As LstbxDisplayAndItem
            Dim i As Byte = 0
            For Each sht As Worksheet In WorkingDataWorkbook.Worksheets
                arrSheetsName(i) = New LstbxDisplayAndItem(sht.Name, sht)
                i += 1
            Next
            With ListBoxWorksheetsName
                .DisplayMember = LstbxDisplayAndItem.DisplayMember
                .ValueMember = LstbxDisplayAndItem.ValueMember
                .DataSource = arrSheetsName
                '.Items.Clear()
                '.Items.AddRange(arrSheetsName)
                '.SelectedItem = .Items(0)
            End With
        End If
        '
        btnGenerate.Enabled = True
    End Sub

    ''' <summary>
    ''' 获取用来绘图的Excel程序，并将此界面加入主程序的监测曲线集合
    ''' </summary>
    ''' <param name="NewExcelApp">按情况看是否要打开新的Application</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetApplication(ByVal NewExcelApp As Boolean,
                                    ByVal MntDrawingExcelApps As Dictionary_AutoKey(Of Cls_ExcelForMonitorDrawing),
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

    Private Function GetDateSpan(ByVal dic As Dictionary(Of Date, Integer)) As DateSpan
        Dim DtSp As DateSpan
        Dim arrTimeRange(0 To F_dicDate_ColNum.Count - 1) As Date
        F_dicDate_ColNum.Keys.CopyTo(arrTimeRange, 0)
        Array.Sort(arrTimeRange)
        DtSp.StartedDate = arrTimeRange(0)
        DtSp.FinishedDate = arrTimeRange(UBound(arrTimeRange))
        Return DtSp
    End Function

    ''' <summary>
    ''' 在工作表中绘制出一个Chart，但是不设置其格式
    ''' </summary>
    ''' <param name="DrawingSheet">Chart所在的工作表</param>
    ''' <param name="UserDefinedTemplate">是否使用用户自定义的Chart模板</param>
    ''' <param name="ChartType">使用Excel自带的模板时的模板类型</param>
    ''' <param name="TemplatePath">用户自定义的Chart模板文件的路径</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function DrawChart(ByVal DrawingSheet As Worksheet, ByVal UserDefinedTemplate As Boolean,
                               Optional ByVal ChartType As XlChartType = XlChartType.xlXYScatterLines, Optional ByVal TemplatePath As String = Nothing) As Chart
        Dim cht As Chart
        DrawingSheet.Activate()
        '---------- 添加图表并选定模板   
        If Not UserDefinedTemplate Then
            cht = DrawingSheet.Shapes.AddChart(ChartType).Chart
        Else
            cht = DrawingSheet.Shapes.AddChart().Chart
            cht.ApplyChartTemplate(TemplatePath)
        End If


        '---------- 设置图表尺寸及标题
        '获取图表中的信息文本框
        F_textbox_Info = cht.Shapes(0).TextFrame2      'Chart中的Shapes集合的第一个元素的下标值为0
        'textbox_Info.AutoSize = Microsoft.Office.Core.MsoAutoSize.msoAutoSizeShapeToFitText
        Return cht
    End Function

    ''' <summary>
    ''' 设置测斜曲线图的Tags属性
    ''' </summary>
    ''' <param name="MntDataSheet">监测数据所在的工作表</param>
    ''' <remarks></remarks>
    Private Function GetChartTags(ByVal MntDataSheet As Worksheet) As MonitorInfo
        Dim MonitorItem As String = DrawingItem.Mnt_Incline
        Dim PointName As String = MntDataSheet.Name
        '
        Dim ExcavationRegion As String
        Dim t_wkbk As Workbook = MntDataSheet.Parent
        Dim filepathwithoutextension As String = System.IO.Path.GetFileNameWithoutExtension(t_wkbk.FullName)
        With filepathwithoutextension
            '如果没有找到"-"，则会返回-1。
            Dim Ind As Short = .IndexOf("-")
            ExcavationRegion = .Substring(Ind + 1, .Length - Ind - 1)
        End With
        '
        Dim Tags As New MonitorInfo(MonitorItem, ExcavationRegion, PointName)
        Return Tags
    End Function

    ''' <summary>
    ''' 绘图完成后的收尾工作
    ''' </summary>
    ''' <param name="blnNewExcelApp">如果是新创建的Excel界面，则进行美化，否则保持原样</param>
    ''' <remarks></remarks>
    Private Sub DrawingFinished(ByVal blnNewExcelApp As Boolean)
        '-------- 界面显示与美化
        If blnNewExcelApp Then           '如果是新创建的Excel界面，则进行美化，否则保持原样
            Call ExcelAppBeauty(F_ExcelAppDrawing)
        End If
        Call ActiveWindowBeauty(F_ExcelAppDrawing.ActiveWindow)

        '刷新窗口
        F_ExcelAppDrawing.ScreenUpdating = True
        '启用主界面的程序滚动按钮
        Call APPLICATION_MAINFORM.MainForm.MainUI_RollingObjectCreated()
    End Sub

    ''' <summary>
    ''' 程序界面美化
    ''' </summary>
    ''' <param name="app"></param>
    ''' <remarks></remarks>
    Private Sub ExcelAppBeauty(ByVal app As Application)

        With app
            .DisplayStatusBar = False
            .DisplayFormulaBar = False
            .Visible = True
            .WindowState = XlWindowState.xlNormal
            .ExecuteExcel4Macro("SHOW.TOOLBAR(""Ribbon"",false)")
        End With
    End Sub
    Private Sub ActiveWindowBeauty(ByVal win As Window)
        With win
            .DisplayGridlines = False
            .DisplayHeadings = False
            .DisplayWorkbookTabs = False
            .Zoom = 100
            .DisplayHorizontalScrollBar = False
            .DisplayVerticalScrollBar = False
            .WindowState = XlWindowState.xlMaximized
        End With
    End Sub

#End Region

#Region "  ---  界面操作"

    ''' <summary>
    ''' 绘制测点位置图
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnDrawMonitorPoints_Click(sender As Object, e As EventArgs) Handles btnDrawMonitorPoints.Click
        Call GlobalApplication.Application.DrawingPointsInVisio()
    End Sub

    ''' <summary>
    ''' 选择文件对话框：选择监测数据的文件
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
                    Me.F_wkbkData = GlobalApplication.Application.ExcelApplication_DB.Workbooks.Open(
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

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked Then
            ComboBox1.Enabled = True
        Else
            ComboBox1.Enabled = False
        End If
    End Sub

#Region "  ---  选择列表框内容时进行赋值"

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

    Private Sub ListBoxWorksheetsName_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBoxWorksheetsName.SelectedIndexChanged
        Try
            Me.F_shtMonitorData = DirectCast(Me.ListBoxWorksheetsName.SelectedValue, Worksheet)
        Catch ex As Exception
            Me.F_shtMonitorData = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' 所选择的监测数据所对应的基坑ID，以及此基坑ID中的相关信息。
    ''' </summary>
    ''' <remarks></remarks>
    Private F_Components As Component()
    Private Sub CbBoxExcavID_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_ExcavID.SelectedIndexChanged
        Me.F_Components = Nothing
        Dim lstItem_ID As LstbxDisplayAndItem = ComboBox_ExcavID.SelectedItem
        If lstItem_ID IsNot Nothing Then
            If Not lstItem_ID.Value.Equals(LstbxDisplayAndItem.NothingInListBox.None) Then
                F_Components = DirectCast(lstItem_ID.Value, Component())
            End If
        End If
    End Sub

    ''' <summary>
    ''' 设置监测数据的类型
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ComboBox_MntType_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox_MntType.SelectedValueChanged
        Me.F_MonitorType = MntType.Incline
        '下面的为程序保留方案，以应对程序中用此窗口来绘制其他类型的监测数据的图形。
        'Dim item As LstbxDisplayAndItem = DirectCast(Me.ComboBox_MntType.SelectedItem, LstbxDisplayAndItem)
        'Me.F_MonitorType = DirectCast(item.Value, MntType)
    End Sub

    ''' <summary>
    ''' 由根据测斜数据表选择的基坑区域的标签，来得到对应的数据列的Range对象。
    ''' 作用是根据选择确定是否要绘制动态开挖深度的直线与文本框
    ''' </summary>
    ''' <remarks></remarks>
    Private F_ProcessRegionData As clsData_ProcessRegionData
    Private Sub CbBoxExcavRegion_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_ExcavRegion.SelectedIndexChanged
        F_ProcessRegionData = Nothing
        Dim lstItem_PR As LstbxDisplayAndItem = ComboBox_ExcavRegion.SelectedItem
        If lstItem_PR IsNot Nothing Then
            If Not lstItem_PR.Value.Equals(LstbxDisplayAndItem.NothingInListBox.None) Then
                F_ProcessRegionData = DirectCast(lstItem_PR.Value, clsData_ProcessRegionData)
            End If
        End If
    End Sub

    Private Sub RadioButton_Dynamic_CheckedChanged(sender As Object, e As EventArgs) _
        Handles RadioButton_Dynamic.CheckedChanged, RadioButton_Max_Depth.CheckedChanged
        If RadioButton_Dynamic.Checked Then
            Me.F_blnMax_Depth = False
            Me.Panel_Dynamic.Visible = True
            Me.Panel_Static.Visible = False
        Else
            Me.F_blnMax_Depth = True
            Me.Panel_Dynamic.Visible = False
            Me.Panel_Static.Visible = True
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

    ''' <summary>
    ''' 在列表中列出所有基坑数据的ID值
    ''' </summary>
    ''' <param name="dic_Data"></param>
    ''' <remarks></remarks>
    Private Sub RefreshComobox_ExcavationID(ByVal dic_Data As Dictionary(Of String, clsData_ExcavationID))
        If dic_Data IsNot Nothing Then
            Try
                Dim IDKeys = dic_Data.Keys
                Dim IDData = dic_Data.Values
                Dim IDcount As Integer = IDKeys.Count
                Dim IDList(0 To IDcount) As LstbxDisplayAndItem

                IDList(0) = New LstbxDisplayAndItem("无", LstbxDisplayAndItem.NothingInListBox.None)
                '
                For i As Integer = 0 To IDcount - 1
                    Dim id As String = IDKeys(i)
                    Dim Components As Component() = IDData(i).Components
                    IDList(i + 1) = New LstbxDisplayAndItem(id, Components)
                Next
                Call RefreshCombobox(Me.ComboBox_ExcavID, IDList)
                '
            Catch ex As Exception
            End Try

        End If
    End Sub

    ''' <summary>
    ''' 在列表中列出所有基坑区域的Tag值
    ''' </summary>
    ''' <param name="ProcessRange"></param>
    ''' <remarks></remarks>
    Private Sub RefreshComobox_ProcessRange(ByVal ProcessRange As List(Of clsData_ProcessRegionData))
        If ProcessRange IsNot Nothing Then
            Try
                Dim TagsCount As Integer = ProcessRange.Count
                Dim TagsList(0 To TagsCount) As LstbxDisplayAndItem

                TagsList(0) = New LstbxDisplayAndItem("无", LstbxDisplayAndItem.NothingInListBox.None)
                '
                For i1 As Integer = 0 To TagsCount - 1
                    Dim PR As clsData_ProcessRegionData = ProcessRange(i1)
                    TagsList(i1 + 1) = New LstbxDisplayAndItem(PR.description, PR)
                Next
                Call RefreshCombobox(Me.ComboBox_ExcavRegion, TagsList)
            Catch ex As Exception
            End Try
        End If
    End Sub

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
            If Me.WindowState = FormWindowState.Maximized Then
                Me.MdiParent.MaximumSize = Me.MaximumSize
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