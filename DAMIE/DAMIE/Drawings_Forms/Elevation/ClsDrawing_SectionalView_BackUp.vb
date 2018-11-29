Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports AME.Miscellaneous
Imports AME.Constants
Imports AME.All_Drawings_In_Application
Imports AME.GlobalApp_Form
Imports AME.DataBase

Namespace All_Drawings_In_Application

    ''' <summary>
    ''' 图表类：剖面标高图
    ''' </summary>
    ''' <remarks></remarks>
    Public Class ClsDrawing_SectionalView
        Implements IAllDrawings, IRolling

#Region "  ---  定义与声明"


#Region "  ---  Types"

        ''' <summary>
        ''' 在开挖剖面图中的每一个基坑开挖区域，其中包含了此基坑区域的矩形、
        ''' 关键构件的直线与文本框、以及与此基坑区域相关的其他数据信息
        ''' </summary>
        ''' <remarks></remarks>
        Public Class clsExcavationRegion

#Region "  ---  属性值定义"

            Private _Rect As Excel.Shape
            ''' <summary>
            ''' 此基坑区域在绘图中的矩形
            ''' </summary>
            ''' <value></value>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public ReadOnly Property Rectangle As Excel.Shape
                Get
                    Return _Rect
                End Get
            End Property

            Private _dicLineAndTextbox As Dictionary(Of Single, line_Textbox)
            ''' <summary>
            ''' 在剖面标高图中，用来进行标识的每一条直线和对应的文本框
            ''' </summary>
            ''' <value></value>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public ReadOnly Property LineAndTextbox As Dictionary(Of Single, line_Textbox)
                Get
                    Return _dicLineAndTextbox
                End Get
            End Property

            Private _Depth As Single
            ''' <summary>
            ''' 基坑区域的当天的开挖深度，用其对应的深度的标高值来表示
            ''' </summary>
            ''' <value></value>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public Property Depth As Single
                Get
                    Return _Depth
                End Get
                Set(value As Single)
                    _Depth = value
                End Set
            End Property

            Private _topRef As Single
            ''' <summary>
            ''' 基坑区域的矩形顶部的基准高度，以磅为单位，它对应的开挖标高为矩形最顶部的标高（不一定是自然地坪标高）
            ''' </summary>
            ''' <value></value>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public ReadOnly Property TopReference As Single
                Get
                    Return _topRef
                End Get
            End Property

            Private _ExcavationRegion As clsData_ProcessRegionData
            ''' <summary>
            ''' 基坑区域所对应的各种数据信息
            ''' </summary>
            ''' <value></value>
            ''' <returns></returns>
            ''' <remarks></remarks>
            Public ReadOnly Property ExcavationRegion As clsData_ProcessRegionData
                Get
                    Return Me._ExcavationRegion
                End Get
            End Property

#End Region

            ''' <summary>
            ''' 构造函数
            ''' </summary>
            ''' <param name="ExcavationRegion">基坑区域所对应的各种数据信息</param>
            ''' <param name="Rectangle">此基坑区域在绘图中的矩形</param>
            ''' <param name="LineAndTextbox">在剖面标高图中，用来进行标识的每一条直线和对应的文本框</param>
            ''' <param name="depth">基坑区域的当天的开挖深度，用其对应的深度的标高值来表示</param>
            ''' <param name="topreference">基坑区域的矩形顶部的基准高度，以磅为单位，它对应的开挖标高为矩形最顶部的标高（不一定是自然地坪标高）</param>
            ''' <remarks></remarks>
            Public Sub New(ByVal ExcavationRegion As clsData_ProcessRegionData, ByVal Rectangle As Excel.Shape, _
                            ByVal LineAndTextbox As Dictionary(Of Single, line_Textbox), _
                            ByVal depth As Single, ByVal topreference As Single)
                Me._Rect = Rectangle
                Me._dicLineAndTextbox = LineAndTextbox
                Me._topRef = topreference
                Me._Depth = depth
                Me._ExcavationRegion = ExcavationRegion
            End Sub

            ''' <summary>
            ''' 在剖面标高图中，用来进行标识的每一条直线和对应的文本框
            ''' </summary>
            ''' <remarks></remarks>
            Public Structure line_Textbox

                Private _identifier As String
                ''' <summary>
                ''' 直线的标识，如：支撑、基坑底
                ''' </summary>
                ''' <value></value>
                ''' <returns></returns>
                ''' <remarks></remarks>
                Public ReadOnly Property Identifier As String
                    Get
                        Return _identifier
                    End Get
                End Property

                Private _elevation As Single
                ''' <summary>
                ''' 构件项目的标高
                ''' </summary>
                ''' <value></value>
                ''' <returns></returns>
                ''' <remarks></remarks>
                Public ReadOnly Property Elevation As Single
                    Get
                        Return _elevation
                    End Get
                End Property

                Private _line As Excel.Shape
                ''' <summary>
                ''' 直线的line对象
                ''' </summary>
                ''' <value></value>
                ''' <returns></returns>
                ''' <remarks></remarks>
                Public ReadOnly Property line As Excel.Shape
                    Get
                        Return _line
                    End Get
                End Property

                Private _textbox As Excel.Shape
                ''' <summary>
                ''' 文本框的textbox对象
                ''' </summary>
                ''' <value></value>
                ''' <returns></returns>
                ''' <remarks></remarks>
                Public ReadOnly Property TextBox As Excel.Shape
                    Get
                        Return _textbox
                    End Get
                End Property

                Private _text As String
                ''' <summary>
                ''' 文本框中的文本
                ''' </summary>
                ''' <value></value>
                ''' <returns>在设置Text属性时也同时更新了文本框中的文本</returns>
                ''' <remarks>也可以用文本框的shape对象.TextFrame2.TextRange.Text属性来表示</remarks>
                Public Property Text As String
                    Get
                        Return _text
                    End Get
                    Set(value As String)
                        _textbox.TextFrame2.TextRange.Text = value
                        _text = value
                    End Set
                End Property

                ''' <summary>
                ''' 构造函数，在创建实例时会对其中的直线与文本框的格式进行设置
                ''' </summary>
                ''' <param name="identifier">直线的标识，如：支撑、基坑底</param>
                ''' <param name="elevation">构件项目的标高</param>
                ''' <param name="line">直线的line对象</param>
                ''' <param name="textbox">文本框的textbox对象</param>
                ''' <param name="Description">对于此构件的描述</param>
                ''' <remarks></remarks>
                Public Sub New(ByVal identifier As String, ByVal elevation As Single, ByVal line As Excel.Shape, ByVal textbox As Excel.Shape, ByVal Description As String)
                    Me._identifier = identifier
                    Me._elevation = elevation
                    Me._line = line
                    Me._textbox = textbox
                    Me._text = Description

                    '设置直线格式
                    With line.Line
                        .Weight = 2
                        .ForeColor.RGB = RGB(0, 0, 155)
                        .DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineDash    '设置线型，此时支撑都还没有做，所以都以虚线显示
                    End With

                    '设置文本框的格式
                    With textbox
                        .Fill.Visible = False
                        .Line.Visible = False
                        With .TextFrame2
                            .MarginBottom = 0
                            .MarginLeft = 0
                            .MarginRight = 0
                            .MarginTop = 0
                            '文本底端对齐
                            .VerticalAnchor = Microsoft.Office.Core.MsoVerticalAnchor.msoAnchorBottom
                            With .TextRange
                                .Text = Description & vbCrLf & elevation.ToString("0.000")

                                '文本居中对齐
                                .ParagraphFormat.Alignment = Microsoft.Office.Core.MsoParagraphAlignment.msoAlignCenter

                                With .Font
                                    .Name = AMEApplication.FontName_TNR
                                    .Size = 6
                                    .Fill.ForeColor.RGB = RGB(0, 0, 0)
                                End With
                            End With
                        End With
                    End With
                End Sub

            End Structure

        End Class

#End Region

#Region "  ---  Constants"
        'Const cstRowNum_TheFirstDay As Byte = Data_Drawing_Format.DB_Progress.RowNum_TheFirstDay    '每一张进度表中，日期的第一天的数据所在的行
        'Const cstColNum_TheFirstDay As Byte = Data_Drawing_Format.DB_Progress.ColNum_DateList    '每一张进度表中，日期的第一天的数据所在的列
        Const cstDrawingTag As String = DrawingItem.SectionalView
#End Region

#Region "  ---- Properties"

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
                Dim processId As Integer
                API.APIWindows.GetWindowThreadProcessId(value.Hwnd, processId)
                F_ExcelProcess = Process.GetProcessById(processId)
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
        ''' 一个文本框，用来记录任意一天的施工日期
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property ConstructionDate As Excel.TextFrame2

        Private P_DateSpan As DateSpan
        ''' <summary>
        ''' 进行同步滚动的时间跨度，用来给出时间滚动条与日历的范围。
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property DateSpan As DateSpan Implements All_Drawings_In_Application.IRolling.DateSpan
            Get
                Return Me.P_DateSpan
            End Get
            Set(value As DateSpan)
                Call GlobalApplication.Application.refreshGlobalDateSpan(value)
                Me.P_DateSpan = value
            End Set
        End Property

        ''' <summary>
        ''' 以选择的"基坑区域的Range"为关键字，去索引此基坑区域的相关信息，包含基坑区域的矩形、关键构件的直线与文本框、以及其他的相关信息
        ''' </summary>
        ''' <value></value>
        ''' <returns>字典中的关键字key所代表的每一个Range对象，
        ''' 都代表工作表的UsedRange中这个区域的一整列的数据(包括前面几行的表头数据)</returns>
        ''' <remarks></remarks>
        Public Property dicRangeAndRegion As Dictionary(Of Range, clsExcavationRegion)

#Region "  ---  绘图的各种标签信息"

        Private _DrawingType As DrawingType
        ''' <summary>
        ''' 此图表所属的类型，由枚举DrawingType提供
        ''' </summary>
        ''' <value></value>
        ''' <returns>DrawingType枚举类型，指示此图形所属的类型</returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Type As DrawingType Implements IAllDrawings.Type
            Get
                Return _DrawingType
            End Get
        End Property

        Private P_Name As String
        ''' <summary>
        ''' 此绘图画面的标签，用来进行索引
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Name As String
            Get
                Return Me.P_Name
            End Get
        End Property

        Private _UniqueID As Long
        ''' <summary>
        ''' 图表的全局独立ID
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks>以当时时间的Tick属性来定义</remarks>
        Public ReadOnly Property UniqueID As Long Implements IAllDrawings.UniqueID
            Get
                Return _UniqueID
            End Get
        End Property

#End Region

#End Region

#Region "  ---  Fields"

        ''' <summary>
        ''' Excel程序的进程对象，用来对此进程进行相差的操作，比如，关闭进程
        ''' </summary>
        ''' <remarks></remarks>
        Private F_ExcelProcess As Process

#End Region

#End Region

        ''' <summary>
        ''' 构造函数
        ''' </summary>
        ''' <param name="Drawingsheet">用于绘制剖面图的Excel工作表</param>
        ''' <param name="DrawingChart">用于绘制剖面图的Excel的Chart对象</param>
        ''' <param name="ChosenRangeAndRectg">此绘图中所包含的矩形方块与对应的数据范围</param>
        ''' <param name="textbox">记录信息的文本框</param>
        ''' <param name="type">此图表所属的类型，由枚举DrawingType提供</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal Drawingsheet As Excel.Worksheet, _
                ByVal DrawingChart As Excel.Chart, _
                ByVal ChosenRangeAndRectg As Dictionary(Of Range, clsExcavationRegion), ByVal DateSpan As DateSpan, _
                ByVal textbox As Excel.TextFrame2, ByVal type As DrawingType)
            Me.ConstructionDate = textbox
            Me.P_Sheet_Drawing = Drawingsheet
            Me.P_Chart = DrawingChart
            Me.dicRangeAndRegion = ChosenRangeAndRectg
            Me.P_Name = cstDrawingTag
            Me._DrawingType = type
            Me.DateSpan = DateSpan
            '
            Me.Application = Drawingsheet.Application
            Me.Application.Caption = Me.P_Name
            '
            Me._UniqueID = GetUniqueID()
            ' -------------------------------------------------------------
            'GlobalApplication.Application.ElevationDrawing = Me
            ' -------------------------------------------------------------
            '刷新滚动窗口的列表框的界面显示
            APPLICATION_MAINFORM.MainForm.Form_Rolling.OnRollingDrawingsRefreshed()
        End Sub

        ''' <summary>
        ''' 进行图形滚动
        ''' </summary>
        ''' <param name="dateThisDay">进行滚动的日期</param>
        ''' <remarks></remarks>
        Public Sub Rolling(ByVal dateThisDay As Date) Implements IRolling.Rolling
            '所有进行对比的基坑区域的Range对象的集合
            Dim colleMyRanges = dicRangeAndRegion.Keys
            Dim RegionCount As Short = colleMyRanges.Count

            If RegionCount > 0 Then          '对每一个进行对比的开挖区域进行滚动
                '
                Me.P_ExcelApp.ScreenUpdating = False

                For i = 0 To RegionCount - 1
                    '进行对比的基坑区域的Range对象
                    Dim ColRange As Range = colleMyRanges(i)


                    '要进行对比的开挖区域的信息
                    Dim ExcavRegionInfo As clsExcavationRegion = dicRangeAndRegion.Item(ColRange)
                    '每一个基坑区域在当天的开挖深度值
                    Dim depth As String = Project_Expo.Elevation_GroundSurface
                    '判断此基坑区域在当天是否已经开挖到基坑底标高
                    Dim blnFilling As Boolean = False
                    If ExcavRegionInfo.ExcavationRegion.HasBottomDate And Date.Compare(dateThisDay, ExcavRegionInfo.ExcavationRegion.BottomDate) > 0 Then        '说明此基坑区域在当天已经开挖到基坑底标高
                        blnFilling = True
                    End If

                    ' ----------------
                    Call DrawingNow(ColRange, ExcavRegionInfo, depth, blnFilling)
                    ' ----------------

                Next        '下一个基坑区域

                '在文本框中显示施工日期，要放在最后显示而不放在最前面显示，是因为图像变化耗时长，而文本框变化非常快。
                ConstructionDate.TextRange.Text = "施工日期：" & dateThisDay.ToString("yyyy/MM/dd")

                '
                Me.P_ExcelApp.ScreenUpdating = True
            End If
        End Sub

        ''' <summary>
        ''' 对每一个基坑区域进行滚动绘图
        ''' </summary>
        ''' <param name="ExcavRegionInfo">要进行对比的开挖区域的信息</param>
        ''' <param name="Range">要进行对比的开挖区域的Range对象,
        ''' 它代表工作表的UsedRange中这个区域的一整列的数据(包括前面几行的表头数据)</param>
        ''' <param name="NewDepth">此开挖区域当天的开挖深度</param>
        ''' <param name="blnFilling">指示当天是否已经挖到基坑底，如果是，则即使当天的深度高于某构件高度，
        ''' 此构件也要用实线显示，因为它已经施工完成了</param>
        ''' <remarks></remarks>
        Private Sub DrawingNow(ByVal Range As Range, ByVal ExcavRegionInfo As clsExcavationRegion, _
                               ByVal NewDepth As Single, Optional ByVal blnFilling As Boolean = False)

            '-------------- 设置开挖矩形的形状：调整其顶部边界的高度
            Dim rectg As Excel.Shape = ExcavRegionInfo.Rectangle
            With rectg
                Dim oldBottom As Single = .Top + .Height
                .Top = (Project_Expo.eleTop - NewDepth) * AMEApplication.cm2pt + ExcavRegionInfo.TopReference
                .Height = oldBottom - .Top
            End With
            ExcavRegionInfo.Depth = NewDepth

            '--------------- 设置矩形形状的颜色：如果还未开挖到基底，则显示为黄色，如果已经开挖到基底（正在往上建），则显示为蓝色
            If blnFilling Then
                rectg.Fill.ForeColor.RGB = Color.Color_BuildindUp
            Else
                rectg.Fill.ForeColor.RGB = Color.Color_DiggingDown
            End If

            '--------------- 设置关键构造（支撑）的显示格式
            Dim Elevations = ExcavRegionInfo.LineAndTextbox.Keys
            Dim LTs = ExcavRegionInfo.LineAndTextbox.Values
            Dim ele As Single
            Dim LT As clsExcavationRegion.line_Textbox

            For i As Short = 0 To Elevations.Count - 1
                ele = Elevations(i)
                LT = LTs(i)
                '文本框中文本的颜色
                Dim textcolor As ColorFormat = LT.TextBox.TextFrame2.TextRange.Font.Fill.ForeColor
                If NewDepth < ele Or blnFilling Then       '说明此构件已经施工完成，应用实线显示
                    With LT.line.Line
                        .DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineSolid
                        .Weight = 2.25
                        .ForeColor.RGB = RGB(255, 0, 0)
                    End With
                    textcolor.RGB = RGB(255, 0, 0)
                Else                    '说明此构件还未施工完成，应用虚线显示
                    With LT.line.Line
                        .DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineDash
                        .Weight = 1.5
                        .ForeColor.RGB = RGB(0, 0, 255)
                    End With
                    textcolor.RGB = RGB(0, 0, 0)
                End If
            Next


        End Sub

#Region "  ---  通用子方法"
        ''' <summary>
        ''' 关闭绘图的文档以及其所在的Application程序
        ''' </summary>
        ''' <param name="SaveChanges">在关闭文档时是否保存修改的内容</param>
        ''' <remarks></remarks>
        Public Sub Close(Optional ByVal SaveChanges As Boolean = False) Implements All_Drawings_In_Application.IAllDrawings.Close
            Try
                With Me
                    Dim wkbk As Excel.Workbook = .Sheet_Drawing.Parent
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
            Try
                Wb.Close(False)
                With GlobalApplication.Application
                    .ElevationDrawing = Nothing
                End With
            Catch ex As Exception
                Debug.Print(ex.Message)
            End Try

            Try
                '刷新滚动窗口的列表框的界面显示
                APPLICATION_MAINFORM.MainForm.Form_Rolling.OnRollingDrawingsRefreshed()

                '退出进程，其操作与与用户使用系统菜单关闭应用程序主窗口的行为一样
                F_ExcelProcess.CloseMainWindow()
            Catch ex As Exception
                Debug.Print(ex.Message)
            End Try
        End Sub

#End Region

    End Class
End Namespace