Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Core
Imports DAMIE.Miscellaneous
Imports DAMIE.Constants
Imports DAMIE.DataBase
Imports DAMIE.All_Drawings_In_Application
Imports DAMIE.GlobalApp_Form

''' <summary>
''' 绘制剖面标高图的窗口界面
''' </summary>
''' <remarks>绘图时，标高值与Excel中的Y坐标值的对应关系：
''' 在程序中，定义了eleTop变量（以米为单位）与F_sngTopRef变量（以磅为单位），
''' 它们指示的是此基坑区域中，地下室顶板的标高值与其在Excel绘图中对应的Y坐标值，
''' 其转换关系式为：构件A在绘图中的Y坐标 = (eleTop - 构件A的标高值) * cm2pt + F_sngTopRef</remarks>
Public Class frmDrawElevation

#Region "  ---  声明与定义"

#Region "  ---  Fields"

#Region "  ---  与Excel相关的对象"

    ''' <summary>
    ''' 进行绘图的Chart对象
    ''' </summary>
    ''' <remarks></remarks>
    Private F_DrawingChart As Excel.Chart

    ''' <summary>
    ''' 绘图工作表中的文本框，用以记录施工当天的日期
    ''' </summary>
    ''' <remarks></remarks>
    Private F_Textbox_Info As Excel.TextFrame2

#End Region

    ''' <summary>
    ''' 程序的主程序对象，这里要单独再用一个变量，是为了解决在多线程调用时APPLICATION_MAINFORM可能会不能被正常调用（它的有些属性会错误地返回Nothing）
    ''' </summary>
    ''' <remarks></remarks>
    Private F_MainForm As APPLICATION_MAINFORM = APPLICATION_MAINFORM.MainForm
    Private GlobalApp As GlobalApplication = GlobalApplication.Application

#End Region

#End Region

#Region "  ---  窗体的加载与关闭"

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        AddHandler ClsData_DataBase.ProcessRangeChanged, AddressOf Me.RefreshComobox_ProcessRange
    End Sub

    '在关闭窗口时将其隐藏
    Private Sub frmDrawSectionalView_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        '如果是子窗口自己要关闭，则将其隐藏
        '如果是mdi父窗口要关闭，则不隐藏，而由父窗口去结束整个进程
        If Not e.CloseReason = CloseReason.MdiFormClosing Then
            Me.Hide()
        End If
        e.Cancel = True
    End Sub

    Private Sub frmDrawElevation_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
        RemoveHandler ClsData_DataBase.ProcessRangeChanged, AddressOf Me.RefreshComobox_ProcessRange
    End Sub

#End Region

    ''' <summary>
    ''' 点击生成按钮
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnGenerate_Click(sender As Object, e As EventArgs) Handles btnGenerate.Click
        Dim count As Short = Me.F_SelectedRegions.Count
        If count > 0 Then
            '开始绘图
            With Me.BGW_Generate
                If Not .IsBusy Then
                    '在工作线程中执行绘图操作
                    Call .RunWorkerAsync(Me.F_SelectedRegions)
                End If
            End With
        End If
    End Sub

#Region "  ---  后台线程进行操作"

    ''' <summary>
    ''' 生成绘图，此方法是在后台的工作者线程中执行的。
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BGW_Generate_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BGW_Generate.DoWork

        '主程序界面的进度条的UI显示
        F_MainForm.ShowProgressBar_Marquee()
        '执行具体的绘图操作
        Call GenerateChart(e.Argument)
        '在绘图完成后，隐藏进度条
        F_MainForm.HideProgress("Done")
    End Sub

    ''' <summary>
    ''' 当后台的工作者线程结束（即BGW_Generate_DoWork方法执行完毕）时触发，注意，此方法是在UI线程中执行的。
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BGW_Generate_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BGW_Generate.RunWorkerCompleted

    End Sub

#End Region

    ''' <summary>
    ''' 开始生成整个绘图图表
    ''' </summary>
    ''' <param name="SelectedRegion">窗口中所有选择的要进行绘图的基坑区域</param>
    ''' <remarks></remarks>
    Private Sub GenerateChart(ByVal SelectedRegion As List(Of clsData_ProcessRegionData))
        '列表中选择的基坑区域
        Dim count As Integer = SelectedRegion.Count
        If count > 0 Then
            '---------------- 打开用于绘图的Excel程序，并进行界面设计
            Dim DrawingSheet As Excel.Worksheet = Me.GetDrawingSheet()
            Dim DrawingApp As Excel.Application = DrawingSheet.Application
            Try

                '------------------- 在绘图工作表中进行绘图
                Me.F_DrawingChart = DrawChart(DrawingSheet, SelectedRegion)

                '  ----------- 绘制数据系列图 ---------------------------
                Dim src As Excel.SeriesCollection = Me.F_DrawingChart.SeriesCollection
                Dim series_DeepestExca As Series = src.Item(1)
                Dim series_Depth As Series = src.Item(2)
                Dim DataSeries(0 To 1) As Series
                DataSeries = SetDataSeries(Me.F_DrawingChart, series_DeepestExca, series_Depth, SelectedRegion)

                '-------------------------------------------------------
                Dim date_Span As DateSpan = GetDateSpan(SelectedRegion)
                '-------------------------------------------------------
                Dim shtEle As ClsDrawing_ExcavationElevation = _
                    New ClsDrawing_ExcavationElevation(series_DeepestExca, series_Depth, SelectedRegion, date_Span, _
                                                       Me.F_Textbox_Info, DrawingType.Xls_SectionalView)

            Catch ex As Exception
                MessageBox.Show("绘制基坑区域开挖标高图失败！" & vbCrLf & ex.Message & vbCrLf & "报错位置：" & ex.TargetSite.Name, _
                "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Finally
                '------- Excel的界面美化 --------------------
                Call ExcelAppBeauty(DrawingApp)
            End Try
        End If
    End Sub

#Region "  ---  绘制标高图"

    ''' <summary>
    ''' 打开用于绘图的Excel程序，并进行界面设计
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetDrawingSheet() As Excel.Worksheet
        Dim app As Excel.Application
        Dim wkbk As Excel.Workbook
        Dim sht As Excel.Worksheet
        '获取绘图的Application对象
        Dim ElevationDrawing As ClsDrawing_ExcavationElevation = GlobalApp.ElevationDrawing
        If ElevationDrawing Is Nothing Then
            app = New Excel.Application
        Else
            app = ElevationDrawing.Application
        End If
        app.Visible = False

        '获取绘图的工作簿
        Dim wkbks As Excel.Workbooks = app.Workbooks
        If wkbks.Count = 0 Then
            wkbk = wkbks.Add
        Else
            wkbk = wkbks.Item(1)
        End If

        '获取绘图的工作表
        sht = wkbk.Worksheets.Add

        sht.Activate()
        ''绘图的标题
        'Static DrawingNum As Integer = 0
        'DrawingNum += 1
        'app.Caption = "绘图" & DrawingNum.ToString
        'sht.Name = "绘图" & DrawingNum.ToString
        ''
        '
        Return sht
    End Function

    ''' <summary>
    ''' 开始绘制开挖剖面图的Chart对象
    ''' </summary>
    ''' <returns>进行绘图的Chart对象的高度值，以磅为单位，可以用来确定Excel Application的高度值</returns>
    ''' <remarks>绘图时，标高值与Excel中的Y坐标值的对应关系：
    ''' 在程序中，定义了eleTop变量（以米为单位）与F_sngTopRef变量（以磅为单位），
    ''' 它们指示的是此基坑区域中，地下室顶板的标高值与其在Excel绘图中对应的Y坐标值</remarks>
    Private Function DrawChart(ByVal DrawingSheet As Excel.Worksheet, _
                               ByVal SelectedRegion As List(Of clsData_ProcessRegionData)) As Excel.Chart
        Dim DrawingApp As Excel.Application = DrawingSheet.Application
        With DrawingApp
            .ScreenUpdating = False
            .Caption = "开挖标高图"
        End With

        '---------------- 创建一个新的，进行绘图的Chart对象 -------------------------------
        Dim DrawingChart As Excel.Chart
        DrawingChart = DrawingSheet.Shapes.AddChart(Top:=0, Left:=0).Chart
        Dim TemplatePath As String = System.IO.Path.Combine(My.Settings.Path_Template, _
                                                            Constants.FolderOrFileName.File_Template.Chart_Elevation)
        With DrawingChart
            .Parent.Activate()
            .ApplyChartTemplate(TemplatePath)
            Me.F_Textbox_Info = .Shapes.Item(1).TextFrame2
            .ChartTitle.Text = "开挖标高图"
            '
            Dim src As Excel.SeriesCollection = .SeriesCollection
            For i As Short = 0 To 1 - src.Count       '确保Chart中至少有两个数据系列
                src.NewSeries()
            Next
        End With
        ' ----------------------- 设置绘图及Excel窗口的尺寸 ----------------------------
        Dim ChartHeight As Double = 400
        Dim InsideLeft As Double = 60
        Dim InsideRight As Double = 20
        Dim LeastWidth_Chart As Double = 500
        Dim LeastWidth_Column As Double = 100
        '
        Dim ChartWidth As Double = LeastWidth_Chart
        Dim insideWidth As Double = LeastWidth_Chart - InsideLeft - InsideRight
        '
        ChartWidth = GetChartWidth(DrawingChart, SelectedRegion.Count, _
                                                 LeastWidth_Chart, LeastWidth_Column, _
                                                 InsideLeft + InsideRight, insideWidth)
        Dim Size_Chart_App As New ChartSize(ChartHeight:=ChartHeight, _
                                    ChartWidth:=ChartWidth, _
                                    MarginOut_Height:=26, _
                                    MarginOut_Width:=9)
        Call ExcelFunction.SetLocation_Size(Size_Chart_App, DrawingChart, DrawingChart.Application, True)
        'With DrawingChart.PlotArea
        '    .InsideLeft = InsideLeft
        '    .InsideWidth = insideWidth
        'End With
        ' --------------------------------------------------
        Return DrawingChart
    End Function

    ''' <summary>
    ''' 根据数据库中的数据信息，绘制两条数据系列图，并在表示基坑深度的那一条数据系列上绘制此区域所在的基坑ID的构件图
    ''' </summary>
    ''' <param name="DrawingChart"></param>
    ''' <param name="SelectedRegion"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SetDataSeries(ByVal DrawingChart As Chart, _
                                  ByVal series_DeepestExca As Series, ByVal Series_Depth As Series, _
                                  ByVal SelectedRegion As List(Of clsData_ProcessRegionData)) As Series()
        Dim RegionsCount As Integer = SelectedRegion.Count
        Dim arrDescrip(0 To RegionsCount - 1) As String                 '每一个区域的描述，作为坐标轴中的X轴数据
        Dim arrDeepest(0 To RegionsCount - 1) As Single                 '每一个区域的坑底标高
        Dim arrDepth(0 To RegionsCount - 1) As Single                   '每一个区域在当天的开挖标高，对于初始绘图，先设定这个值为地面标高
        Dim arrExcavID(0 To RegionsCount - 1) As clsData_ExcavationID   '每一个区域所对应的基坑ID对象
        Dim Elevation_Ground As Single = Project_Expo.Elevation_GroundSurface   '项目的自然地面的标高
        '所有选择的区域中的最深的标高位置，以米为单位
        Dim DeepestElevation As Single = Elevation_Ground
        For i As UShort = 0 To RegionsCount - 1
            Dim Region As clsData_ProcessRegionData = SelectedRegion.Item(i)
            With Region
                arrDescrip(i) = .description
                Dim BottomElevation As Single = .ExcavationID.ExcavationBottom
                arrDeepest(i) = BottomElevation
                arrDepth(i) = Elevation_Ground ' ClsData_DataBase.GetElevation(Region.Range, )
                DeepestElevation = Math.Min(DeepestElevation, BottomElevation)
                arrExcavID(i) = .ExcavationID
            End With
        Next
        '  ------------------------  设置Chart数据  ---------------------------
        Try
            With series_DeepestExca
                .Name = ""
                .XValues = arrDescrip
                .Values = arrDeepest
            End With
            With Series_Depth
                .Name = ""
                .XValues = arrDescrip
                .Values = arrDepth
            End With
        Catch ex As Exception
            MessageBox.Show("设置基坑区域开挖图中的开挖标高数据出错！" & vbCrLf & ex.Message &
                            vbCrLf & "报错位置：" & ex.TargetSite.Name, _
                            "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        '  ------------------------  设置坐标轴格式  ---------------------------
        Try
            Dim max As Double = Elevation_Ground

            Dim min As Double
            min = Min_Array(Of Single)(arrDeepest)
            If min > 0 Then
                min = Math.Ceiling(min)
            Else        '注意Math.Ceiling(-3.2)=-3
                min = Math.Floor(min)
            End If
            Dim axY As Excel.Axis = DrawingChart.Axes(Excel.XlAxisType.xlValue)
            With axY
                .MaximumScale = max
                .MinimumScale = min
                .AxisTitle.Text = "标高（m）"
            End With
            Dim axX As Excel.Axis = DrawingChart.Axes(Excel.XlAxisType.xlValue)
            With axX

            End With
        Catch ex As Exception
            MessageBox.Show("设置基坑区域开挖图中的坐标轴格式出错！" & vbCrLf & ex.Message & vbCrLf & "报错位置：" & ex.TargetSite.Name, _
   "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        '  --------------  绘制每一个选定区域所属的基坑ID的支撑位置  -----------
        Call DrawComponents(arrExcavID, DrawingChart, series_DeepestExca)
        '
        Return {series_DeepestExca, Series_Depth}
    End Function

    ''' <summary>
    ''' 在Chart中绘制每一个选定区域所属的基坑ID的支撑位置
    ''' </summary>
    ''' <param name="arrExcavID">所有选择的基坑区域所属的基坑ID</param>
    ''' <param name="series_DeepestExca">构件图形的位置的参考</param>
    ''' <remarks></remarks>
    Private Sub DrawComponents(ByVal arrExcavID() As clsData_ExcavationID, ByVal DrawingChart As Chart, ByVal series_DeepestExca As Series)
        Dim RegionsCount As UShort = arrExcavID.Length
        Try
            With DrawingChart
                Dim Index As UShort = 0
                Dim arrRegionName(0 To RegionsCount - 1) As String
                For Each pt As Point In series_DeepestExca.Points
                    Dim Components As Component() = arrExcavID(Index).Components
                    Dim dblLeft As Double = pt.Left
                    Dim dblWidth As Double = pt.Width
                    Dim listComponentName As New List(Of String)
                    For Each Component As Component In Components
                        If Component.Type = ComponentType.Strut Then
                            Dim dblTop As Double = ExcelFunction.GetPositionInChartByValue(.Axes(Excel.XlAxisType.xlValue), _
                                                                                                          Component.Elevation)
                            '添加标高线
                            Dim shpLine As Excel.Shape = .Shapes.AddLine(BeginX:=dblLeft, BeginY:=dblTop, _
                                                                                     EndX:=dblLeft + dblWidth, EndY:=dblTop)
                            With shpLine.Line
                                .Weight = 1.5
                                .DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineSolid
                                .ForeColor.RGB = RGB(255, 0, 0)
                            End With
                            '添加文本框
                            Dim dblTextHeight As Double = 40
                            Dim shpText As Excel.Shape = .Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, _
                                                                            dblLeft, dblTop - dblTextHeight, dblWidth, dblTextHeight)
                            With shpText
                                Call ExcelFunction.FormatTextbox_Tag(shpText.TextFrame2, TextSize:=8, Text:=Component.Description, _
                                                                     VerticalAnchor:=MsoVerticalAnchor.msoAnchorBottom)
                            End With

                            Dim shpRg_Components As Excel.Shape = .Shapes.Range({shpLine.Name, shpText.Name}).Group
                            listComponentName.Add(shpRg_Components.Name)
                        End If
                    Next            '下一组构件
                    With listComponentName
                        Dim ComponentCount As UShort = .Count
                        If ComponentCount > 2 Then
                            arrRegionName(Index) = DrawingChart.Shapes.Range(listComponentName.ToArray).Group().Name
                        ElseIf ComponentCount = 1 Then      '只有一组构件，那么不用进行Group
                            arrRegionName(Index) = .Item(0)
                        Else                                '说明此基坑区域中一个构件也没有。
                            arrRegionName(Index) = Nothing
                        End If
                    End With
                    Index += 1
                Next            '下一个基坑区域
                If arrRegionName.Count >= 2 Then
                    .Shapes.Range(arrRegionName).Group()
                End If
            End With
        Catch ex As Exception
            MessageBox.Show("绘制基坑构件出错！" & vbCrLf & ex.Message & vbCrLf & "报错位置：" & ex.TargetSite.Name, _
      "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End Try
    End Sub

    ''' <summary>
    ''' 确定Chart的宽度值，以磅为单位
    ''' </summary>
    ''' <param name="DrawingChart">进行绘图的Chart</param>
    ''' <param name="PointsCount">选择的基坑区域的数量</param>
    ''' <param name="LeastWidth_Chart">Chart对象的宽度的最小值，即使图表中只有很少的点</param>
    ''' <param name="LeastWidth_Column">柱形图中每一个柱形的最小宽度，用来进行基本的文本书写</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetChartWidth(ByVal DrawingChart As Excel.Chart, ByVal PointsCount As UShort, _
                                   ByVal LeastWidth_Chart As Double, ByVal LeastWidth_Column As Double, _
                                   ByVal Margin As Double, ByRef InsideWidth As Double) As Double
        Dim ChartWidth As Double = LeastWidth_Chart
        InsideWidth = LeastWidth_Chart - Margin
        '
        With DrawingChart
            Dim src As Excel.SeriesCollection = .SeriesCollection
            Dim n As UShort = src.Count
            Dim hh As Double = LeastWidth_Chart / PointsCount
            Dim ChtGroup As ChartGroup = .ChartGroups(1)
            ' ------------------------------------------------------------------------------------------------
            '1、在已知最小的chart宽度(即最小的PlotArea.InsideWidth)的情况下，验算柱形的宽度是否比指定的最小柱形宽度要大
            Dim H As Double = InsideWidth
            ' Dim O As Single = ChtGroup.Overlap / 100
            ChtGroup.Overlap = 0
            Dim G As Single = ChtGroup.GapWidth / 100
            Dim ColumnWidth As Double = hh / (1 + G + n)
            If ColumnWidth < LeastWidth_Column Then
                ' ------------------------------------------------------------------------------------------------
                '2、在已知柱体的最小宽度的情况下，去推算整个PlotArea.InsideWidth的值


            End If
        End With
        Return ChartWidth
    End Function

#End Region

#Region "  ---  子方法"

    ''' <summary>
    ''' 获取时间跨度
    ''' </summary>
    ''' <param name="SelectedRegion">窗口中所有选择的要进行绘图的基坑区域</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetDateSpan(ByVal SelectedRegion As List(Of clsData_ProcessRegionData)) As DateSpan
        Dim RegionCount As UShort = SelectedRegion.Count
        If RegionCount > 0 Then
            '
            Const cstColNum_DateList As Byte = Data_Drawing_Format.DB_Progress.ColNum_DateList
            Const cstRowNum_TheFirstDay As Byte = Data_Drawing_Format.DB_Progress.RowNum_TheFirstDay
            '
            Dim arrRangesChosen(0 To RegionCount - 1) As Range
            Dim i As UShort = 0
            For Each Region As clsData_ProcessRegionData In SelectedRegion
                arrRangesChosen(i) = Region.Range_Process
                i += 1
            Next
            Dim DateSpan_Old As DateSpan

            Dim dtStartDay_new As Date
            Dim dtEndDay_new As Date
            Try
                '以arrRangesChosen的第一个值来进行TimeSpan_SectionalView的初始化
                Dim sht As Worksheet = arrRangesChosen(0).Worksheet
                With sht
                    dtStartDay_new = .Cells(cstRowNum_TheFirstDay, cstColNum_DateList).value
                    dtEndDay_new = .Cells(.UsedRange.Rows.Count, cstColNum_DateList).value
                    DateSpan_Old.StartedDate = dtStartDay_new
                    DateSpan_Old.FinishedDate = dtEndDay_new
                End With
                '从第二项开始，对TimeSpan_SectionalView进行扩展
                Dim DateSpan_new As DateSpan
                For iRg As Byte = 1 To UBound(arrRangesChosen)
                    sht = arrRangesChosen(iRg).Worksheet
                    With sht
                        dtStartDay_new = .Cells(cstRowNum_TheFirstDay, cstColNum_DateList).value
                        dtEndDay_new = .Cells(.UsedRange.Rows.Count, cstColNum_DateList).value
                        DateSpan_new.StartedDate = dtStartDay_new
                        DateSpan_new.FinishedDate = dtEndDay_new
                        DateSpan_Old = ExpandDateSpan(DateSpan_Old, DateSpan_new)
                    End With
                Next
            Catch ex As Exception
                MessageBox.Show("提取基坑区域中的施工日期跨度DateSpan出错！" & vbCrLf & ex.Message & vbCrLf & "报错位置：" & ex.TargetSite.Name, _
    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

            Return DateSpan_Old
        End If
    End Function

    ''' <summary>
    ''' 程序界面美化
    ''' </summary>
    ''' <param name="ExcelApp"></param>
    ''' <remarks></remarks>
    Private Sub ExcelAppBeauty(ByVal ExcelApp As Application)
        With ExcelApp
            .Visible = False
            .DisplayStatusBar = False
            .DisplayFormulaBar = False
            With .ActiveWindow  '对标高图工作表的界面进行设置
                .DisplayGridlines = False
                .DisplayHeadings = False
                .DisplayWorkbookTabs = False
                .Zoom = 100
                .DisplayHorizontalScrollBar = False
                .DisplayVerticalScrollBar = False
            End With
            .WindowState = XlWindowState.xlNormal
            FixWindow(.Hwnd)
            .Visible = True
            '隐藏Excel的功能区。注意，在进行隐藏前，一定要确保Application.Visible属性为True，否则显示出来界面会没有“标题栏”。
            '在此情况下，只要用鼠标右击窗口中的任何区域（单元格或形状等）即可将标题栏显示出来。
            .ExecuteExcel4Macro("SHOW.TOOLBAR(""Ribbon"",false)")
            .ScreenUpdating = True
        End With
    End Sub

#End Region

#Region "  ---  一般的界面操作"

    ''' <summary>
    ''' 全选
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnChooseAll_Click(sender As Object, e As EventArgs) Handles btnChooseAll.Click
        Dim blnSelected As Boolean = True
        Dim btItemsCount As Byte = lstbxChooseRegion.Items.Count
        For i As Byte = 0 To btItemsCount - 1
            lstbxChooseRegion.SetSelected(i, blnSelected)
        Next
    End Sub

    ''' <summary>
    ''' 全不选
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnChooseNone_Click(sender As Object, e As EventArgs) Handles btnChooseNone.Click
        lstbxChooseRegion.ClearSelected()
    End Sub

    ''' <summary>
    ''' 在列表中列出所有基坑区域的Tag值
    ''' </summary>
    ''' <param name="ProcessRange"></param>
    ''' <remarks></remarks>
    Private Sub RefreshComobox_ProcessRange(ByVal ProcessRange As List(Of clsData_ProcessRegionData))
        If ProcessRange IsNot Nothing Then
            With Me.lstbxChooseRegion
                Try
                    ' list all the tags of each excavation region in the listbox
                    Dim count As Byte = ProcessRange.Count
                    Dim arrItems(0 To count - 1) As LstbxDisplayAndItem
                    Dim i As Byte = 0
                    For Each PR As clsData_ProcessRegionData In ProcessRange
                        arrItems(i) = New LstbxDisplayAndItem(DisplayedText:=PR.description, _
                                                    Value:=PR)
                        i += 1
                    Next
                    Call RefreshCombobox(Me.lstbxChooseRegion, arrItems)
                Catch ex As Exception

                End Try
            End With
        End If
    End Sub

    ''' <summary>
    ''' 窗口中所有选择的要进行绘图的基坑区域
    ''' </summary>
    ''' <remarks></remarks>
    Private F_SelectedRegions As List(Of clsData_ProcessRegionData)
    Private Sub RefreshSelectedRegion(sender As Object, e As EventArgs) Handles lstbxChooseRegion.SelectedIndexChanged
        Dim items = Me.lstbxChooseRegion.SelectedItems
        Dim SelectedRegion As New List(Of clsData_ProcessRegionData)
        For Each item As LstbxDisplayAndItem In items
            SelectedRegion.Add(DirectCast(item.Value, clsData_ProcessRegionData))
        Next
        Me.F_SelectedRegions = SelectedRegion
    End Sub

#End Region

End Class