Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Core
Imports AME.Miscellaneous
Imports AME.Constants
Imports AME.DataBase
Imports AME.All_Drawings_In_Application
Imports AME.GlobalApp_Form

''' <summary>
''' 绘制剖面标高图的窗口界面
''' </summary>
''' <remarks>绘图时，标高值与Excel中的Y坐标值的对应关系：
''' 在程序中，定义了eleTop变量（以米为单位）与F_sngTopRef变量（以磅为单位），
''' 它们指示的是此基坑区域中，地下室顶板的标高值与其在Excel绘图中对应的Y坐标值，
''' 其转换关系式为：构件A在绘图中的Y坐标 = (eleTop - 构件A的标高值) * cm2pt + F_sngTopRef</remarks>
Public Class frmDrawElevation_BackUp

#Region "  ---  声明与定义"

#Region "  --- 类型定义"

    ' ''' <summary>
    ' ''' 每一个基坑ID对应的信息
    ' ''' </summary>
    ' ''' <remarks>可以用在对每一个基坑开挖区域定义其所在的区域的剖面信息</remarks>
    'Private Structure ExcavID_Data
    '    Public ExcavationID As String
    '    Public ExcavationData As ClsData_DataBase.ExcavationID
    'End Structure

#End Region

#Region "  ---  常数值定义"

    ''' <summary>
    ''' 所有标高项目中的最高位置的标高值
    ''' </summary>
    ''' <remarks></remarks>
    Const cstEleTop As Single = Project_Expo.eleTop
    ''' <summary>
    ''' 工程项目的自然地面的标高值
    ''' </summary>
    ''' <remarks></remarks>
    Const cstEleGround As Single = Project_Expo.Elevation_GroundSurface
    ''' <summary>
    ''' 长度单位厘米到磅的转换系数
    ''' </summary>
    ''' <remarks>后面的除1.3表示将图形缩小1.3位</remarks>
    Const cm2pt As Single = AMEApplication.cm2pt
    '
    ''' <summary>
    ''' 记录施工进度日期的数据列号
    ''' </summary>
    ''' <remarks></remarks>
    Const cstColNum_DateList As Byte = Data_Drawing_Format.DB_Progress.ColNum_DateList
    ''' <summary>
    ''' 第一个日期所在的行号
    ''' </summary>
    ''' <remarks></remarks>
    Const cstRowNum_TheFirstDay As Byte = Data_Drawing_Format.DB_Progress.RowNum_TheFirstDay

    ''' <summary>
    ''' 每一个基坑区域的矩形形状的宽度，以磅为单位
    ''' </summary>
    ''' <remarks></remarks>
    Const cstWidth_Rectg As Single = 2.5 * cm2pt

    ''' <summary>
    ''' 文本框“地面标高”的宽度
    ''' </summary>
    ''' <remarks></remarks>
    Const cstWidth_Textbox As Single = cstWidth_Rectg + 5

#End Region

#Region "  ---  Fields"

#Region "  ---  与Excel相关的对象"

    ''' <summary>
    ''' 进行绘图的Chart对象
    ''' </summary>
    ''' <remarks></remarks>
    Private F_DrawingChart As Excel.Chart

    ''' <summary>
    ''' 数据库中的工作表：剖面标高
    ''' </summary>
    ''' <remarks></remarks>
    Private F_shtElevation As Excel.Worksheet

    ''' <summary>
    ''' 绘图工作表中的文本框，用以记录施工当天的日期
    ''' </summary>
    ''' <remarks></remarks>
    Private F_textbox_constructionDate As Excel.TextFrame2

#End Region

    ''' <summary>
    ''' 创建一个集合，以基坑的ID来索引[此基坑的标高项目,对应的标高]的数组。
    ''' </summary>
    ''' <remarks>其key为"基坑的ID"，value为[此基坑的标高项目,对应的标高]的数组</remarks>
    Private F_dicIDtoElevationAndData As New Dictionary(Of String, clsData_ExcavationID)(StringComparer.OrdinalIgnoreCase)

    ''' <summary>
    ''' 关键变量，以每一个选择的基坑区域的Range对象来索引其所代表的基坑区域的各种信息
    ''' </summary>
    ''' <remarks></remarks>
    Private F_dicChoseRange_Region As New Dictionary(Of Excel.Range, ClsDrawing_SectionalView.clsExcavationRegion)


    ''' <summary>
    ''' 第一个基坑区域的矩形左边界的坐标值，以磅为单位
    ''' </summary>
    ''' <remarks>sngLeftRef = shtDrawing.Columns("A:A").Width</remarks>
    Private F_sngLeftRef As Single

    ''' <summary>
    ''' 基坑区域的矩形顶部的基准高度，以磅为单位，它对应的开挖标高为矩形最顶部的标高（不一定是自然地坪标高）
    ''' </summary>
    ''' <remarks>sngTopRef = shtDrawing.Rows("1:3").Height</remarks>
    Private F_sngTopRef As Single

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

    Private Sub frmDrawElevation_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
        RemoveHandler ClsData_DataBase.ProcessRangeChanged, AddressOf Me.RefreshComobox_ProcessRange
    End Sub

    '加载窗体
    Private Sub frmGenerateSectionalView_Load(sender As Object, e As EventArgs) Handles Me.Load
        '数据库中所有的开挖区域
        With GlobalApp.DataBase
            If .Sheet_Elevation IsNot Nothing Then
                F_shtElevation = .Sheet_Elevation
            Else
                MessageBox.Show("没有发现基坑标高相关的数据，请先在项目中添加相关数据。", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End With

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

#End Region

    ''' <summary>
    ''' 点击生成按钮
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnGenerate_Click(sender As Object, e As EventArgs) Handles btnGenerate.Click
        Dim count As Short = Me.lstbxChooseRegion.SelectedItems.Count
        If count > 0 Then
            '开始绘图
            With Me.BGW_Generate
                If Not .IsBusy Then

                    '提取列表中选择的项的Item数据
                    '这里不直接用列表框的SelectedItems属性，而要将其提取到数组中，
                    '是因为后面的绘图操作是在工作者线程中完成的，而在工作者线程中不能操作UI界面中的控件和获取其属性。
                    Dim items = Me.lstbxChooseRegion.SelectedItems
                    Dim arrItems(0 To count - 1) As LstbxDisplayAndItem
                    items.CopyTo(arrItems, 0)
                    '在工作线程中执行绘图操作
                    Call .RunWorkerAsync(arrItems)
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

#Region "  !!!  进行绘图"

    ''' <summary>
    ''' 开始生成整个绘图图表
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub GenerateChart(ByVal arrSelectedItems As LstbxDisplayAndItem())

        '---------------- 打开用于绘图的Excel程序，并进行界面设计
        Dim DrawingSheet As Excel.Worksheet = Me.GetDrawingSheet()
        Dim DrawingApp As Excel.Application = DrawingSheet.Application

        '列表中选择的基坑区域
        Dim count As Integer = arrSelectedItems.Length
        '
        With GlobalApp.DataBase
            'assign dicIDtoElevationAndData here but not in the Load block, 
            'because the mainForm.DataBase.dicIDtoElevationAndData would change if the Floor is ignored
            F_dicIDtoElevationAndData = .ID_Components
        End With


        '以进度列表中的选择的"基坑区域"来索引[基坑ID,[基坑ID对应的标高项,标高项对应的标高值]]
        Dim dicProcessRegion As New Dictionary(Of Range, clsData_ProcessRegionData)

        '进度表中所有选择的基坑区域的Range对象，
        '每一个Range对象代表工作表的UsedRange中这个区域的一整列的数据(包括前面几行的表头数据)
        Dim arrRangesChosen(0 To count - 1) As Range
        '
        Dim iTag As Byte = 0
        '由列表框的SelectedValue属性来返回选择的Range对象的集合
        For Each ChosenItem As LstbxDisplayAndItem In arrSelectedItems
            Dim ChosendRegion As clsData_ProcessRegionData = DirectCast(ChosenItem.Value, clsData_ProcessRegionData)
            With ChosendRegion
                arrRangesChosen(iTag) = .Range_Process
                dicProcessRegion.Add(.Range_Process, ChosendRegion)
            End With
            iTag += 1
        Next

        '对界面形状进行设置
        F_sngLeftRef = DrawingSheet.Columns("A:A").Width
        F_sngTopRef = DrawingSheet.Rows("1:3").Height

        '------------------- 在绘图工作表中进行绘图
        Me.F_DrawingChart = DrawChart(DrawingSheet, dicProcessRegion, arrRangesChosen)

        ' ------ 设置Excel程序界面的尺寸
        With DrawingApp
            Try
                Dim sngScale As Single = .ActiveWindow.Zoom / 100
                .Width = F_DrawingChart.ChartArea.Width * sngScale + 16
                .Height = F_DrawingChart.ChartArea.Height * sngScale + 33
            Catch ex As Exception
                Debug.Print(ex.Message)
            End Try

        End With

        '------------------------------------------------------- 将绘好图的界面指针传递给主程序
        Dim ds As DateSpan = GetDateSpan(arrRangesChosen)
        Dim shtEle As ClsDrawing_SectionalView = New ClsDrawing_SectionalView( _
                                             DrawingSheet, F_DrawingChart, F_dicChoseRange_Region, ds, _
                                             F_textbox_constructionDate, DrawingType.Xls_SectionalView)
        '-------------------------------------------------------- Excel的界面美化
        Call ExcelAppBeauty(DrawingApp)
        '
        DrawingApp.Visible = True
        '隐藏Excel的功能区。注意，在进行隐藏前，一定要确保Application.Visible属性为True，否则显示出来界面会没有“标题栏”。
        '在此情况下，只要用鼠标右击窗口中的任何区域（单元格或形状等）即可将标题栏显示出来。
        DrawingApp.ExecuteExcel4Macro("SHOW.TOOLBAR(""Ribbon"",false)")

    End Sub

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
    ''' 1、开始绘制开挖剖面图的Chart对象
    ''' </summary>
    ''' <param name="dicProcessRegion">以进度列表中的选择的"基坑区域"来索引此区域的相关信息</param>
    ''' <param name="arrRangesChosen">进度表中所有选择的基坑区域的Range对象，
    ''' 每一个Range对象代表工作表的UsedRange中这个区域的一整列的数据(包括前面几行的表头数据)</param>
    ''' <returns>进行绘图的Chart对象的高度值，以磅为单位，可以用来确定Excel Application的高度值</returns>
    ''' <remarks>绘图时，标高值与Excel中的Y坐标值的对应关系：
    ''' 在程序中，定义了eleTop变量（以米为单位）与F_sngTopRef变量（以磅为单位），
    ''' 它们指示的是此基坑区域中，地下室顶板的标高值与其在Excel绘图中对应的Y坐标值</remarks>
    Private Function DrawChart(ByVal DrawingSheet As Excel.Worksheet, ByVal dicProcessRegion As Dictionary(Of Range, clsData_ProcessRegionData), _
                                ByVal arrRangesChosen() As Range) As Excel.Chart
        Dim DrawingApp As Excel.Application = DrawingSheet.Application
        DrawingApp.ScreenUpdating = False
        '
        F_dicChoseRange_Region.Clear()

        '要进行对比的基坑
        Dim intRectgColumnCount As Integer = arrRangesChosen.Count

        '---------------- 创建一个新的，进行绘图的Chart对象 -----------------------------------------------------------------------------------

        '这里的Chart的高度值要在创建时就确定下来，而不能在Chart创建后再去进行修改，
        '因为Chart的尺寸的变化会引起其中的形状的形状的缩放，而这种缩放并不是等比例的。
        Dim ChartHeight As Single
        '所有进行对比的基坑区域中，坑底标高最深的值，以米为单位。
        Dim DeepestElevation As Single
        For intExcavNum As Short = 0 To arrRangesChosen.Length - 1
            Dim rangeColumn As Range = arrRangesChosen(intExcavNum)         '每一列的列号对应一个基坑区域
            Dim ProcessRegion As clsData_ProcessRegionData = dicProcessRegion.Item(rangeColumn)           '选择的基坑区域所对应的各种数据信息
            '更新所有要进行对比的基坑区域中，坑底标高最深的值，用来定义此Excel绘图界面的窗口高度
            DeepestElevation = Math.Min(DeepestElevation, ProcessRegion.ExcavationID.ExcavationBottom)
        Next
        '确定最终的绘图Chart对象的高度值
        ChartHeight = (cstEleTop - DeepestElevation) * cm2pt + F_sngTopRef
        Dim DrawingChart As Excel.Chart
        Dim ChtObjt As Excel.ChartObjects = DrawingSheet.ChartObjects
        DrawingChart = ChtObjt.Add(Left:=0, _
                                    Top:=0, _
                                    Width:=F_sngLeftRef * 2 + intRectgColumnCount * cstWidth_Rectg + cstWidth_Textbox, _
                                    Height:=ChartHeight).Chart
        ' -----------------------------------------------------------------------------------------------------------------------------------------

        '------------------------------------------ 1、画出两条标志线：地下室顶板标高与地面标高
        Dim line1 As Excel.Shape, line2 As Excel.Shape, shprgLines As Excel.ShapeRange

        '地下室顶板标高位置在Excel绘图中对应的Y坐标
        Dim t_Y_For_Top As Single = F_sngTopRef
        line1 = DrawingChart.Shapes.AddLine(BeginX:=0, _
                                          BeginY:=t_Y_For_Top, _
                                          EndX:=2 * Me.F_sngLeftRef + cstWidth_Rectg * intRectgColumnCount, _
                                          EndY:=t_Y_For_Top)

        '地面标高位置在Excel绘图中对应的Y坐标
        Dim t_Y_For_Ground As Single = (cstEleTop - cstEleGround) * cm2pt + F_sngTopRef
        line2 = DrawingChart.Shapes.AddLine(BeginX:=0, _
                                            BeginY:=t_Y_For_Ground, _
                                            EndX:=2 * Me.F_sngLeftRef + cstWidth_Rectg * intRectgColumnCount, _
                                            EndY:=t_Y_For_Ground)
        shprgLines = DrawingChart.Shapes.Range({line1.Name, line2.Name})
        With shprgLines.Line
            .Weight = 1.5
            .DashStyle = MsoLineDashStyle.msoLineDash
            .EndArrowheadStyle = MsoArrowheadStyle.msoArrowheadStealth
            .EndArrowheadWidth = MsoArrowheadWidth.msoArrowheadWide
        End With
        '-------------------------------------------- 2、添加文本框“地面标高”
        Dim txtbox As Excel.Shape
        txtbox = DrawingChart.Shapes.AddTextbox( _
                             Orientation:=MsoTextOrientation.msoTextOrientationHorizontal, _
                             Left:=2 * Me.F_sngLeftRef + cstWidth_Rectg * intRectgColumnCount, _
                             Top:=(cstEleTop - cstEleGround) * cm2pt + F_sngTopRef - 7.5, _
                             Width:=cstWidth_Textbox, _
                             Height:=15)
        Dim RightEdge = 2 * Me.F_sngLeftRef + cstWidth_Rectg * intRectgColumnCount + cstWidth_Textbox
        '设置文本框中的文字内容与排版
        txtbox.Line.Visible = False
        With txtbox.TextFrame2
            .MarginLeft = 2
            .MarginRight = 2
            .MarginTop = 0
            .MarginBottom = 0
            With .TextRange
                .Text = "地面标高"
                .ParagraphFormat.Alignment = MsoParagraphAlignment.msoAlignCenter
                .Font.Size = 10
            End With
            .VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle
        End With


        '-------------------------------------------- 3、绘制文本框，用来记录施工当天的日期
        Dim Temp_shp As Excel.Shape
        '文本框的形状参考文本框“地面标高”
        Temp_shp = DrawingChart.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, _
                              F_sngLeftRef + cstWidth_Rectg * intRectgColumnCount, 0, cstWidth_Textbox + F_sngLeftRef, 20)

        F_textbox_constructionDate = Temp_shp.TextFrame2
        With F_textbox_constructionDate.TextRange
            .Text = "施工日期："
            .ParagraphFormat.Alignment = MsoParagraphAlignment.msoAlignLeft
            With .Font
                .Name = AMEApplication.FontName_TNR
                .Size = 8
                .Bold = False
            End With
        End With
        '------------------------------------- 4、绘制每一个基坑区域的标高图 ----------------------------

        '每一个基坑区域的图形在Excel中的左边界，以磅为单位
        Dim sngLeft As Single

        For intExcavNum As Short = 0 To arrRangesChosen.Length - 1

            '每一列的列号对应一个基坑区域
            Dim rangeColumn As Range = arrRangesChosen(intExcavNum)
            '更新下一个基坑区域的图形在Excel绘图中的左边界的值
            sngLeft = F_sngLeftRef + cstWidth_Rectg * (intExcavNum)
            '选择的基坑区域所对应的各种数据信息
            Dim ProcessRegion As clsData_ProcessRegionData = dicProcessRegion.Item(rangeColumn)

            '更新所有要进行对比的基坑区域中，坑底标高最深的值，用来定义此Excel绘图界面的窗口高度
            DeepestElevation = Math.Min(DeepestElevation, ProcessRegion.ExcavationID.ExcavationBottom)

            '---------------------------------------

            Dim regionInfo As ClsDrawing_SectionalView.clsExcavationRegion = DrawRegion(DrawingChart, sngLeft, ProcessRegion)

            '---------------------------------------
            F_dicChoseRange_Region.Add(rangeColumn, regionInfo)
            ' ------------------
        Next        '下一个基坑区域

        '刷新Excel屏幕
        DrawingApp.ScreenUpdating = True
        '
        Return DrawingChart
    End Function

    ''' <summary>
    ''' 2、对于每一个基坑区域，在Chart中去绘制与其相关的图形
    ''' </summary>
    ''' <param name="DrawingChart">进行绘图的Chart对象</param>
    ''' <param name="Left">图形在Excel中的左边界，以磅为单位</param>
    ''' <param name="ProcessRegion">基坑区域对应的各种数据信息</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function DrawRegion(ByVal DrawingChart As Excel.Chart, ByVal Left As Single, _
                                ByVal ProcessRegion As clsData_ProcessRegionData) _
                                As ClsDrawing_SectionalView.clsExcavationRegion
        '每一个基坑区域的进度的Range对象
        Dim range As Excel.Range = ProcessRegion.Range_Process
        Dim sngHeight As Single
        Dim sngTop As Single
        '-----------------1、添加文本框“每个基坑的Title”
        '
        Dim txtbox As Excel.Shape = DrawingChart.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, _
                                                                   Left:=Left, _
                                                                   Top:=0, _
                                                                   Width:=cstWidth_Rectg, _
                                                                   Height:=F_sngTopRef)
        '设置文本框中的文字内容与排版
        With txtbox.TextFrame2
            .MarginLeft = 2
            .MarginRight = 2
            .MarginTop = 0
            .MarginBottom = 0
            With .TextRange
                .Text = range.Cells(Data_Drawing_Format.DB_Progress.RowNum_ExcavTag, 1).MergeArea.Cells(1, 1).Value _
                             & range.Cells(Data_Drawing_Format.DB_Progress.RowNum_ExcavPosition, 1).Value
                .ParagraphFormat.Alignment = MsoParagraphAlignment.msoAlignCenter
                '设置字体
                .Font.Name = AMEApplication.FontName_TNR
                .Font.Size = 8
            End With
            .VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle
        End With

        '------------------- 2、画矩形

        Dim Components_ID() As Component
        Components_ID = ProcessRegion.ExcavationID.Components
        '
        '基坑ID所对应的所有构件项目中，标高位置最高的标高
        Dim excavationTop As Single = Components_ID.First.Elevation

        '矩形的顶端位置
        sngTop = (cstEleTop - excavationTop) * cm2pt + F_sngTopRef
        '矩形的高度：从这里可以看出开挖深度与Excel中的矩形是如何进行单位转换的
        sngHeight = (excavationTop - ProcessRegion.ExcavationID.ExcavationBottom) * cm2pt
        '画出矩形
        Dim rectg As Excel.Shape
        rectg = DrawingChart.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, _
                                             Left:=Left, _
                                             Top:=sngTop, _
                                             Width:=cstWidth_Rectg, _
                                             Height:=sngHeight)
        '设置矩形的格式
        With rectg
            .Placement = XlPlacement.xlFreeFloating
            .Line.Weight = 0.5
            .Fill.ForeColor.RGB = Color.Color_DiggingDown
        End With

        '------------------ 3、 在支撑的标高上画出标识直线与对应的文本框 ---------------

        Dim dicEle_lineAndTxtBx As New Dictionary(Of Single, ClsDrawing_SectionalView.clsExcavationRegion.line_Textbox)
        For itemRow As Short = LBound(Components_ID) To UBound(Components_ID) - 1
            Dim sngElevation As Single = Components_ID(itemRow).Elevation
            '此构件项目的名称
            If Components_ID(itemRow).Type = ComponentType.Strut Then
                '直线的Y坐标
                Dim sngLineY As Single = (cstEleTop - sngElevation) * cm2pt + F_sngTopRef

                '画直线
                Dim lineWithStruts As Excel.Shape
                lineWithStruts = DrawingChart.Shapes.AddLine(Left, sngLineY, Left + cstWidth_Rectg, sngLineY)

                '画文本框
                Dim textbox As Excel.Shape
                Dim t_sngHeight As Single = 25
                Dim t_sngTop As Single = sngLineY - t_sngHeight
                textbox = DrawingChart.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, _
                                                        Left, t_sngTop, cstWidth_Rectg, t_sngHeight)

                ' -------------
                Dim line_txtbx As New ClsDrawing_SectionalView.clsExcavationRegion.line_Textbox( _
                    Data_Drawing_Format.DB_Sectional.identifier_struts, sngElevation, lineWithStruts, textbox, Components_ID(itemRow).Description)
                dicEle_lineAndTxtBx.Add(sngElevation, line_txtbx)
                ' -------------

            End If

        Next    '下一个直线与文本框

        ' ------------------
        Dim regionInfo As ClsDrawing_SectionalView.clsExcavationRegion = New ClsDrawing_SectionalView.clsExcavationRegion(ProcessRegion, rectg, dicEle_lineAndTxtBx, excavationTop, F_sngTopRef)
        ' ------------------
        Return regionInfo
    End Function

#End Region

    ''' <summary>
    ''' 获取时间跨度
    ''' </summary>
    ''' <param name="arrRangesChosen">表示进度表中所有选择的基坑区域的Range对象,
    ''' 每一个Range对象代表工作表的UsedRange中这个区域的一整列的数据(包括前面几行的表头数据)</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetDateSpan(ByVal arrRangesChosen As Range()) As DateSpan

        Dim DateSpan_Old As DateSpan
        Dim sht As Worksheet
        Dim dtStartDay_new As Date
        Dim dtEndDay_new As Date

        '以arrRangesChosen的第一个值来进行TimeSpan_SectionalView的初始化
        sht = arrRangesChosen(0).Worksheet
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
        Return DateSpan_Old
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
            '隐藏Excel的功能区。注意，在进行隐藏前，一定要确保Application.Visible属性为True，否则显示出来界面会没有“标题栏”。
            '在此情况下，只要用鼠标右击窗口中的任何区域（单元格或形状等）即可将标题栏显示出来。
            ' .ExecuteExcel4Macro("SHOW.TOOLBAR(""Ribbon"",false)")
        End With
    End Sub

End Class