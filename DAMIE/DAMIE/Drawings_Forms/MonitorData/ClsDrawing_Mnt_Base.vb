Imports Microsoft.Office.Interop.Excel
Imports Office = Microsoft.Office.Core
Imports DAMIE.Miscellaneous
Imports DAMIE.Constants
Imports DAMIE.All_Drawings_In_Application
Imports DAMIE.GlobalApp_Form

Namespace All_Drawings_In_Application
    Public MustInherit Class ClsDrawing_Mnt_Base
        Implements IAllDrawings, Dictionary_AutoKey(Of ClsDrawing_Mnt_Base).I_Dictionary_AutoKey

#Region " --- Properties"

        Private F_Application As Application
        ''' <summary>
        ''' 图表所在的Excel的Application对象
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Application As Application
            Get
                Return Me.F_Application
            End Get
        End Property

        ''' <summary>
        ''' 监测数据的工作表
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Sheet_Data As Worksheet

        ''' <summary>
        ''' 工作表“标高图”
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Sheet_Drawing As Worksheet

        Private F_myChart As Chart
        ''' <summary>
        ''' 监测曲线图
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Chart As Chart
            Get
                Return F_myChart
            End Get
            Set(value As Chart)
                F_myChart = value
            End Set
        End Property

        ''' <summary>
        ''' 绘图界面与画布的尺寸
        ''' </summary>
        ''' <value></value>
        ''' <returns>此结构有四个元素，分别代表：画布的高度、宽度；由画布边界扩展到Excel界面的尺寸的高度和宽度的增量</returns>
        ''' <remarks></remarks>
        Protected MustOverride Property ChartSize_sugested As ChartSize

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
        Public Property Information As TextFrame2
            Get
                Return F_textbox_Info
            End Get
            Set(value As TextFrame2)
                F_textbox_Info = value
            End Set
        End Property

        Private F_Class_ParentApp As Cls_ExcelForMonitorDrawing
        ''' <summary>
        ''' 此画布所有的Class_ExcelForMonitorDrawing实例
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Parent As Cls_ExcelForMonitorDrawing
            Get
                Return F_Class_ParentApp
            End Get
        End Property

        Private F_blnCanRoll As Boolean
        ''' <summary>
        ''' 指示此图表是否可以按时间进行同步滚动
        ''' </summary>
        ''' <value></value>
        ''' <returns>如果为True，则可以同步滚动，为动态图；如果为False，则为静态图</returns>
        ''' <remarks></remarks>
        Public ReadOnly Property CanRoll As Boolean
            Get
                Return F_blnCanRoll
            End Get
        End Property

#Region "  ---  绘图的各种标签信息"

        Private F_DrawingType As DrawingType
        ''' <summary>
        ''' 此图表所属的类型，由枚举DrawingType提供
        ''' </summary>
        ''' <value></value>
        ''' <returns>DrawingType枚举类型，指示此图形所属的类型</returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Type As DrawingType Implements IAllDrawings.Type
            Get
                Return F_DrawingType
            End Get
        End Property

        ''' <summary>
        ''' 监测数据的类型，比如测斜数据、立柱垂直位移数据、支撑轴力数据等
        ''' </summary>
        ''' <remarks></remarks>
        Private P_MntType As MntType
        ''' <summary>
        ''' 监测数据的类型，比如测斜数据、立柱垂直位移数据、支撑轴力数据等
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property MonitorType As MntType
            Get
                Return Me.P_MntType
            End Get
        End Property

        ''' <summary>
        ''' 绘图图表的相关标签信息
        ''' </summary>
        ''' <remarks></remarks>
        Private P_Tags As MonitorInfo
        ''' <summary>
        ''' 绘图图表的相关标签信息
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks>包括其项目名称、基坑编号、监测点位名称</remarks>
        Public Property Tags As MonitorInfo
            Get
                Return Me.P_Tags
            End Get
            Private Set(value As MonitorInfo)
                Me.P_Tags = value
                With value
                    Dim TT As String = .MonitorItem & "-" & .ExcavationRegion
                    If (Me.Type = DrawingType.Monitor_Incline_Dynamic) OrElse (Me.Type = DrawingType.Monitor_Incline_MaxMinDepth) Then
                        TT = TT & "-" & .PointName
                    End If
                    '
                    Me.Chart_App_Title = TT
                End With
            End Set
        End Property

        ''' <summary>
        ''' 此对象在它所在的集合中的键。用来在集合中索引到此对象：me=集合.item(me.key)
        ''' </summary>
        ''' <remarks></remarks>
        Private P_Key As Integer
        ''' <summary>
        ''' 此元素在其所在的集合中的键，这个键是在元素添加到集合中时自动生成的，
        ''' 所以应该在执行集合.Add函数时，用元素的Key属性接受函数的输出值。
        ''' 在集合中添加此元素：Me.Key=Me所在的集合.Add(Me)
        ''' 在集合中索引此元素：集合.item(me.key)
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Key As Integer _
            Implements Dictionary_AutoKey(Of ClsDrawing_Mnt_Base).I_Dictionary_AutoKey.Key
            Get
                Return Me.P_Key
            End Get
        End Property

        Private F_UniqueID As Long
        ''' <summary>
        ''' 图表的全局独立ID
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks>以当时时间的Tick属性来定义</remarks>
        Public ReadOnly Property UniqueID As Long Implements IAllDrawings.UniqueID
            Get
                Return F_UniqueID
            End Get
        End Property

        ''' <summary>
        ''' 这个属性值的变化会同步到监测曲线的曲线标题，以及绘图程序的窗口标题中。
        ''' </summary>
        ''' <remarks></remarks>
        Private P_Chart_App_Title As String
        ''' <summary>
        ''' 这个属性值的变化会同步到监测曲线的曲线标题，以及绘图程序的窗口标题中。
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Chart_App_Title As String
            Get
                Return Me.P_Chart_App_Title
            End Get
            Set(value As String)
                Me.P_Chart_App_Title = value
                Me.Chart.ChartTitle.Text = value
                If Me.Application IsNot Nothing Then
                    Me.Application.Caption = value
                End If
            End Set
        End Property

#End Region

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
        Public Sub New(ByVal DataSheet As Worksheet, _
                       ByVal DrawingChart As Chart, ByVal ParentApp As Cls_ExcelForMonitorDrawing, _
                       ByVal type As DrawingType, ByVal CanRoll As Boolean, ByVal Info As TextFrame2, _
                       ByVal DrawingTag As MonitorInfo, ByVal MonitorType As MntType)
            Try
                '设置Excel窗口与Chart的尺寸

                MyClass.F_Application = ParentApp.Application
                Call ExcelFunction.SetLocation_Size(Me.ChartSize_sugested, DrawingChart, Me.Application, True)
                '
                MyClass.Sheet_Data = DataSheet
                MyClass.F_myChart = DrawingChart
                MyClass.F_textbox_Info = Info
                MyClass.Sheet_Drawing = DrawingChart.Parent.Parent
                MyClass.F_blnCanRoll = CanRoll
                MyClass.F_DrawingType = type
                MyClass.P_MntType = MonitorType
                '将此对象添加进其所属的集合中
                F_Class_ParentApp = ParentApp
                '
                With F_Class_ParentApp.Mnt_Drawings
                    Me.P_Key = .Add(Me)
                End With
                Me.F_UniqueID = GetUniqueID()
                '
                Me.Tags = DrawingTag
            Catch ex As Exception
                MessageBox.Show("创建基本监测曲线图出错。" & vbCrLf & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End Try

        End Sub
        '
        ''' <summary>
        ''' 将自己从所在的集合中删除
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub RemoveFormCollection()
            Dim thd As New Thread(AddressOf RemoveMyselfFromCollection)
            With thd
                .Name = "删除工作表"
                .Start()
            End With
        End Sub
        Private Sub RemoveMyselfFromCollection()
            Try
                Me.Application.ScreenUpdating = False
                Me.Sheet_Drawing.Delete()
                Me.Parent.Mnt_Drawings.Remove(Me.Key)

            Catch ex As Exception
                MessageBox.Show("删除工作表出错。" & vbCrLf & ex.Message & vbCrLf & "报错位置：" & ex.TargetSite.Name, _
         "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)

            End Try
        End Sub
        '
        ''' <summary>
        ''' 关闭绘图的Excel文档以及其所在的Application程序
        ''' </summary>
        ''' <param name="SaveChanges">在关闭文档时是否保存修改的内容</param>
        ''' <remarks></remarks>
        Public Sub Close(Optional ByVal SaveChanges As Boolean = False) Implements All_Drawings_In_Application.IAllDrawings.Close
            Try
                With Me
                    Dim wkbk As Workbook = .Sheet_Drawing.Parent
                    wkbk.Close(SaveChanges)
                    .Application.Quit()
                End With
            Catch ex As Exception
                MessageBox.Show("关闭监测曲线图出错！" & vbCrLf & ex.Message, _
                             "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End Try
        End Sub

        ''' <summary>
        ''' 按不同的监测数据类型返回对应的坐标轴标签
        ''' </summary>
        ''' <param name="Drawing_Type"></param>
        ''' <param name="MonitorType"></param>
        ''' <param name="AxisType"></param>
        ''' <param name="AxisGroup">此选项仅对绘制双Y轴或者双X轴的Chart时有效，否则设置了也不会进行处理。</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetAxisLabel(ByVal Drawing_Type As DrawingType, ByVal MonitorType As MntType, _
                                            ByVal AxisType As XlAxisType, Optional ByVal AxisGroup As XlAxisGroup = XlAxisGroup.xlPrimary) As String
            Dim strAxisLabel As String = ""
            Select Case Drawing_Type
                ' ---------------------  动态测斜曲线图  ----------------------------------------------------
                Case DrawingType.Monitor_Incline_Dynamic
                    Select Case AxisType
                        Case XlAxisType.xlCategory
                            strAxisLabel = AxisLabels.Displacement_mm
                        Case XlAxisType.xlValue
                            strAxisLabel = AxisLabels.Depth
                    End Select
                    ' ------------------------  测斜位移最值及对应深度  -------------------------------------------------
                Case DrawingType.Monitor_Incline_MaxMinDepth
                    Select Case AxisType
                        Case XlAxisType.xlCategory
                            strAxisLabel = AxisLabels.ConstructionDate
                        Case XlAxisType.xlValue
                            Select Case AxisGroup
                                Case XlAxisGroup.xlPrimary
                                    strAxisLabel = AxisLabels.Displacement_mm

                                Case XlAxisGroup.xlSecondary
                                    strAxisLabel = AxisLabels.Depth
                            End Select
                    End Select
                    ' -----------------  测斜以外的动态监测曲线曲线图  --------------------------------------------------------
                Case DrawingType.Monitor_Dynamic
                    Select Case AxisType
                        Case XlAxisType.xlCategory
                            strAxisLabel = AxisLabels.Points
                        Case XlAxisType.xlValue
                            Select Case MonitorType
                                Case MntType.Struts
                                    strAxisLabel = AxisLabels.AxialForce
                                Case MntType.WaterLevel
                                    strAxisLabel = AxisLabels.Displacement_m
                                Case Else
                                    strAxisLabel = AxisLabels.Displacement_mm
                            End Select
                    End Select
                    ' -------------------------  测斜以外的静态曲线图  ------------------------------------------------
                Case DrawingType.Monitor_Static
                    Select Case AxisType
                        Case XlAxisType.xlCategory
                            strAxisLabel = AxisLabels.ConstructionDate
                        Case XlAxisType.xlValue
                            Select Case MonitorType
                                Case MntType.Struts
                                    strAxisLabel = AxisLabels.AxialForce
                                Case MntType.WaterLevel
                                    strAxisLabel = AxisLabels.Displacement_m
                                Case Else
                                    strAxisLabel = AxisLabels.Displacement_mm
                            End Select
                    End Select
                Case DrawingType.Xls_SectionalView
                    Select Case AxisType
                        Case XlAxisType.xlCategory
                            strAxisLabel = AxisLabels.Excavation
                        Case XlAxisType.xlValue
                            strAxisLabel = AxisLabels.Elevation
                    End Select
            End Select
            Return strAxisLabel
        End Function

    End Class
End Namespace