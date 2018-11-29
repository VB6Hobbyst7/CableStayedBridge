Imports Microsoft.Office.Interop.Excel
Imports Office = Microsoft.Office.Core
Imports DAMIE.GlobalApp_Form
Imports DAMIE.Miscellaneous
Imports DAMIE.Constants
Namespace All_Drawings_In_Application

    ''' <summary>
    ''' 用Excel绘制的，可以进行数据曲线的滚动的类型
    ''' </summary>
    ''' <remarks></remarks>
    Public MustInherit Class clsDrawing_Mnt_RollingBase
        Inherits ClsDrawing_Mnt_Base
        Implements IRolling

#Region "  ---  Types"

        ''' <summary>
        ''' 数据列series对象所对应的一些信息对象
        ''' </summary>
        ''' <remarks>它包括每一条数据系列的对象，及其对应的日期、开挖深度的标志线、记录开挖深度值的文本框</remarks>
        Public Class SeriesTag

            Private P_series As Series
            ''' <summary>
            ''' 数据系列的series对象
            ''' </summary>
            ''' 数据列对象
            Public ReadOnly Property series As Series
                Get
                    Return Me.P_series
                End Get
            End Property

            Private P_ConstructionDate As Date
            ''' <summary>
            ''' 进行滚动的施工日期
            ''' </summary>
            ''' 对应的施工日期
            Public Property ConstructionDate As Date
                Get
                    Return Me.P_ConstructionDate
                End Get
                Set(value As Date)
                    Me.P_ConstructionDate = value
                End Set
            End Property

            Public Sub New(ByVal Series As Series, ByVal ConstructionDate As Date)
                Me.P_series = Series
                Me.P_ConstructionDate = ConstructionDate
            End Sub

        End Class

        ''' <summary>
        ''' 图例框的尺寸与位置，用来在进行监测数据曲线的滚动时设置图例的大小与位置
        ''' </summary>
        ''' <remarks></remarks>
        Public Structure LegendLocation
            ''' <summary>
            ''' 图例框的高度
            ''' </summary>
            ''' <remarks></remarks>
            Public Legend_Height
            ''' <summary>
            ''' 图例框的宽度
            ''' </summary>
            ''' <remarks></remarks>
            Public Legend_Width
            ''' <summary>
            ''' 构造函数
            ''' </summary>
            ''' <param name="LegendHeight">图例框的高度</param>
            ''' <param name="LegendWidth">图例框的宽度</param>
            ''' <remarks></remarks>
            Public Sub New(LegendHeight As Single, LegendWidth As Single)
                Me.Legend_Height = LegendHeight
                Me.Legend_Width = LegendWidth
            End Sub
        End Structure

#End Region

#Region "  ---  Constants"

        ''' <summary>
        ''' 数据列在集合seriesCollection中的第一个元素的下标值
        ''' </summary>
        ''' <remarks>office2010的dll中，其值为1（2014.11.13注）</remarks>
        Private Const cst_LboundOfSeriesInCollection As Byte = LowIndexOfObjectsInExcel.SeriesInSeriesCollection

#End Region

#Region "  ---  Properties"

        ''' <summary>
        ''' 此结构有四个元素，分别代表：画布的高度、宽度；由画布边界扩展到Excel界面的尺寸的高度和宽度的增量
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Protected MustOverride Overrides Property ChartSize_sugested As ChartSize

        ''' <summary>
        ''' 图例框的尺寸与位置
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Protected Overridable ReadOnly Property Legend_Location As LegendLocation
            Get
                Return New LegendLocation(LegendHeight:=Data_Drawing_Format.Drawing_Mnt_RollingBase.Legend_Height, _
                                          LegendWidth:=Data_Drawing_Format.Drawing_Mnt_RollingBase.Legend_Width)
            End Get
        End Property

        Private P_DateSpan As DateSpan
        ''' <summary>
        ''' 进行同步滚动的时间跨度，用来给出时间滚动条与日历的范围。
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property DateSpan As DateSpan Implements IRolling.DateSpan
            Get
                Return Me.P_DateSpan
            End Get
            Private Set(value As DateSpan)
                ' 扩展MainForm.TimeSpan的区间
                Call GlobalApplication.Application.refreshGlobalDateSpan(value)
                Me.P_DateSpan = value
            End Set
        End Property

        ''' <summary>
        ''' 用来进行滚动的那一个数据列对象，此对象在复制的过程中，并不会被替换，
        ''' 而且其seriesIndex的值始终为1.
        ''' </summary>
        ''' <remarks></remarks>
        Private P_MovingSeries As Series
        ''' <summary>
        ''' 用来进行滚动的那一个数据列对象，此对象在复制的过程中，并不会被替换
        ''' 而且其seriesIndex的值始终为1.
        ''' </summary>
        ''' <remarks></remarks>
        Protected Property MovingSeries As Series
            Get
                Return Me.P_MovingSeries
            End Get
            Set(value As Series)
                Me.P_MovingSeries = value
            End Set
        End Property


        ''' <summary>
        ''' 用数据列的ID属性来索引对应的数据列，以及与之相关的Tag数据
        ''' </summary>
        Protected F_DicSeries_Tag As New Dictionary(Of Integer, SeriesTag)
        ''' <summary>
        ''' 用数据列的ID属性来索引对应的数据列，以及与之相关的Tag数据
        ''' </summary>
        ''' <value>其Key值对应是的每一条曲线的SeriesIndex的值，其中第一条曲线的下标值为1，而不是0。</value>
        ''' <returns></returns>
        ''' <remarks>作为关键字的ID值并不是以0为初始值的，
        ''' 而是在series在seriescollection中的序号值来定义的。</remarks>
        Public ReadOnly Property Dic_SeriesIndex_Tag As Dictionary(Of Integer, SeriesTag)
            Get
                Return Me.F_DicSeries_Tag
            End Get
        End Property
#End Region

#Region "  ---  Fields"

        ''' <summary>
        ''' 记录图表中所有的数据系列中的每一项是否有曲线图
        ''' </summary>
        ''' <remarks>由此来确定数据曲线是要放在哪一个系列中，
        ''' 以及图例项中要删除哪些或者显示哪些项。
        ''' 此列表中的元素的个数始终等于图表中定义的数据系列的个数，即seriesCollection.Count的值</remarks>
        Protected F_List_HasCurve As New List(Of Boolean)

        ''' <summary>
        ''' 图表中所包含的数据曲线的数量
        ''' </summary>
        ''' <remarks>由于制作模板的问题，所以SeriesCount指的是图表中的显示出来的监测曲线的数量，
        ''' 而seriesCollection.Count指的是图表中定义的数据系列的数量。
        ''' 空白系列会包含在seriesCollection.Count中，而在图表中并没有曲线图，因为没有数据。</remarks>
        Private F_CurvesCount As Integer

        ''' <summary>
        ''' 当前滚动的日期。由于进行滚动的总是seriescollection中的第一条曲线，
        ''' 所以当前滚动的日期也总是这第一条曲线所代表的日期，亦即Rolling方法中的dateThisday参数的值。
        ''' </summary>
        ''' <remarks></remarks>
        Protected F_RollingDate As Date
#End Region

        ''' <summary>
        ''' 构造函数
        ''' </summary>
        ''' <param name="DataSheet">图表对应的数据工作表</param>
        ''' <param name="DrawingChart">Excel图形所在的Chart对象</param>
        ''' <param name="ParentApp">此图表所在的Excel类的实例对象</param>
        ''' <param name="type">此图表所属的类型，由枚举drawingtype提供</param>
        ''' <param name="CanRoll">是图表是否可以滚动，即是动态图还是静态图</param>
        ''' <param name="DateSpan">此图表的TimeSpan跨度</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal DataSheet As Worksheet, _
                        ByVal DrawingChart As Chart, ByVal ParentApp As Cls_ExcelForMonitorDrawing, _
                        ByVal type As DrawingType, ByVal CanRoll As Boolean, ByVal Info As TextFrame2, _
                        ByVal DrawingTag As MonitorInfo, ByVal MonitorType As MntType, _
                        ByVal DateSpan As DateSpan, _
                        ByVal theFirstSeriesTag As SeriesTag)
            Call MyBase.New(DataSheet, DrawingChart, ParentApp, type, CanRoll, Info, DrawingTag, MonitorType)
            '
            MyClass.DateSpan = DateSpan
            '刷新滚动窗口的列表框的界面显示
            APPLICATION_MAINFORM.MainForm.Form_Rolling.OnRollingDrawingsRefreshed()

            '启用主界面的程序滚动按钮
            Call APPLICATION_MAINFORM.MainForm.MainUI_RollingObjectCreated()

            '--------------------------- 设置与数据系列的曲线相关的属性值

            With Me
                '以数据列中第一个元素作为进行滚动的那个series
                .MovingSeries = theFirstSeriesTag.series
                ' ----- 集合数据的记录
                .F_DicSeries_Tag.Add(cst_LboundOfSeriesInCollection, theFirstSeriesTag)
                '刚开始时，图表中只有一条数据曲线
                .F_CurvesCount = 1
            End With
            '
            With Me.F_List_HasCurve
                .Clear()
                .Add(True)          '第一个数据列是有曲线的，所以将其值设置为true
                For i = 1 To Me.Chart.SeriesCollection.Count - 1
                    .Add(False)
                Next
            End With
            ' -----对图例进行更新---------
            Call LegendRefresh(F_List_HasCurve)

        End Sub

        ''' <summary>
        ''' 日期滚动时的Rolling方法
        ''' </summary>
        ''' <param name="dateThisday">当天的日期</param>
        ''' <remarks>在基类中，并不给出此方法的实现，而由其子类去进行重写</remarks>
        Public MustOverride Sub Rolling(ByVal dateThisday As Date) Implements IRolling.Rolling

        ''' <summary>
        ''' 根据所有的日期集合和当天的日期，得到距当天最近的那天的日期。
        ''' </summary>
        ''' <param name="sortedlist_AllDays">有监测数据的所有的日期集合。</param>
        ''' <param name="dateThisDay">想要进行绘图的那一天的日期。</param>
        ''' <returns>在集合中距离当天最近的那一天。</returns>
        ''' <remarks></remarks>
        Protected Function GetClosestDay(ByVal sortedlist_AllDays As SortedSet(Of Date), ByVal dateThisDay As Date) As Date
            Dim day_early As Date
            Dim day_late As Date
            Dim closedDay As Date
            Dim intCompare As Integer
            If Date.Compare(dateThisDay, sortedlist_AllDays(0)) < 0 Then        '如果选定的日期比最早的监测日期还要早
                closedDay = sortedlist_AllDays(0)

            ElseIf Date.Compare(dateThisDay, sortedlist_AllDays(sortedlist_AllDays.Count - 1)) > 0 Then         '如果选定的日期比最晚的监测日期还要晚
                closedDay = sortedlist_AllDays(sortedlist_AllDays.Count - 1)

            Else                         '如果选定的日期位于监测的日期跨度之间
                For Each iDay As Date In sortedlist_AllDays
                    intCompare = Date.Compare(iDay, dateThisDay)
                    '从最早的一天开始比较，那么刚开始的intCompare的值一定小于0，直到有一天的日期值比选定日期晚，此时intCompare的值大于0
                    '不会出现intCompare等于0的情况，因为比较是在字典的键集合中没有选定日期的情况下进行的
                    If intCompare < 0 Then      '这天比选定的日期早，记录这一天，以作后面day_late的前一天的值。
                        day_early = iDay
                    Else
                        day_late = iDay             '发现了满足条件的判断点：出现了比选定日期晚的那一天
                        '接下来从选定日期的左右日期中选择距离最近的一天，如果两边距离相等，则取较早的一天。
                        If day_late.Subtract(dateThisDay) >= -day_early.Subtract(dateThisDay) Then
                            closedDay = day_early       '从选定日期的左右日期中选择距离最近的一天，如果两边距离相等，则取较早的一天。
                        Else
                            closedDay = day_late
                        End If
                        Exit For                '不用再继续找下一天了
                    End If
                Next
            End If
            Return closedDay
        End Function

#Region "  ---  数据列的锁定与删除"

        Private WithEvents F_Chart As Chart = Me.Chart
        ''' <summary>
        ''' 当在数据列上双击时执行相应的操作
        ''' </summary>
        ''' <param name="ElementID">在图表上双击击中的对象</param>
        ''' <param name="Arg1">所选数据系列在集合中的索引下标值，注意，第一第曲线的下标值为1，而不是0。</param>
        ''' <param name="Arg2"></param>
        ''' <param name="Cancel"></param>
        ''' <remarks>要么将其删除，要么将其锁定在图表中，要么什么都不做</remarks>
        Private Sub SeriesChange(ElementID As Integer, Arg1 As Integer, Arg2 As Integer, _
                                 ByRef Cancel As Boolean) Handles F_Chart.BeforeDoubleClick
            If ElementID = XlChartItem.xlSeries Then
                Debug.Print("当前选择的曲线下标值为： " & Arg1)
                '所选数据系列在集合中的索引下标值，注意，第一第曲线的下标值为1，而不是0
                Dim seriesIndex As Integer = Arg1
                '所选的数据系列
                Dim seri As Series
                seri = Me.Chart.SeriesCollection.Item(seriesIndex)


                ' -------- 打开处理对话框，并根据返回的不同结果来执行不同的操作
                Dim diafrm As New DiaFrm_LockDelete
                '判断要删除的曲线是否为当前滚动的曲线，有如下两种判断方法
                Dim blnDeletingTheRollingCurve As Boolean
                blnDeletingTheRollingCurve = (seriesIndex = cst_LboundOfSeriesInCollection)
                'blnDeletingTheRollingCurve = (seri.Name = F_movingSeries.Name)
                If F_CurvesCount = 1 Or blnDeletingTheRollingCurve Then
                    '如果图表中只有一条曲线，或者点击的正好是要正在进行滚动的曲线，那么这条曲线不能被删除
                    With diafrm
                        .btn2.Enabled = False
                        .AcceptButton = .btn1
                    End With
                Else            '只能进行删除
                    With diafrm
                        .btn1.Enabled = False
                        .AcceptButton = .btn2
                    End With
                End If

                '在界面上取消图表对象的选择
                Me.Sheet_Drawing.Range("A1").Activate()
                Me.Application.ScreenUpdating = False

                '---------------------------------- 开始执行数据系列的添加或删除
                '此数据列对应的施工日期
                Dim t_date As Date
                Dim SrTag As SeriesTag
                Try
                    SrTag = Me.F_DicSeries_Tag.Item(seriesIndex)
                    t_date = SrTag.ConstructionDate
                Catch ex As Exception
                    MessageBox.Show("提取字典中的元素出错，字典中没有此元素。" & vbCrLf & ex.Message & vbCrLf & "报错位置：" & ex.TargetSite.Name, _
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
                '----------------------------------------
                Dim result As M_DialogResult
                Dim strPrompt As String = "The date value for the selected series is:" & vbCrLf &
                              t_date.ToString(AMEApplication.DateFormat) & vbCrLf &
                              "Choose to Lock or Delete this series..."
                result = diafrm.ShowDialog(strPrompt, "Lock or Delete Series")
                '----------------------------------------
                Select Case result
                    Case M_DialogResult.Delete
                        '删除数据系列
                        Try
                            Call DeleteSeries(seriesIndex)
                        Catch ex As Exception
                            Debug.Print("删除数据系列出错。" & vbCrLf & ex.Message & vbCrLf & "报错位置：" & ex.TargetSite.Name)
                        End Try
                    Case M_DialogResult.Lock
                        '添加数据系列
                        Try
                            Call CopySeries(seriesIndex)
                        Catch ex As Exception
                            Debug.Print(ex.Message)
                            MessageBox.Show("添加数据系列出错。" & vbCrLf & ex.Message & vbCrLf & "报错位置：" & ex.TargetSite.Name, _
                                "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try
                End Select
                Me.Application.ScreenUpdating = True
            End If
            '覆盖原来的双击操作
            Cancel = True
            Me.Sheet_Drawing.Range("A1").Activate()
        End Sub

        '曲线删除
        ''' <summary>
        ''' 移除指定的数据列及其依附的对象
        ''' </summary>
        ''' <param name="DeletingSeriesIndex">要进行删除的数据列在集合中的索引下标值，注意：第一条曲线的下标值为1，而不是0。</param>
        ''' <remarks>为了保存模板中的数据系列的格式，这里的删除并不是将数据列进行了真正的删除，而是将数据列的数据设置为空。
        ''' 这样的话，后期的数据曲线应该加载在最靠前而且没有数据的数据列中。</remarks>
        Public Overridable Sub DeleteSeries(ByVal DeletingSeriesIndex As Integer)
            With F_DicSeries_Tag.Item(DeletingSeriesIndex)
                With .series
                    .XValues = {Nothing}
                    '这里不能用.XValues = Nothing
                    .Values = {Nothing}
                    '.Name = ""
                End With
            End With

            Me.F_DicSeries_Tag.Remove(DeletingSeriesIndex)
            Me.F_List_HasCurve.Item(DeletingSeriesIndex - cst_LboundOfSeriesInCollection) = False

            ' ----------------------------------- 对图例的显示进行操作
            Call LegendRefresh(F_List_HasCurve)

            Me.F_CurvesCount -= 1
        End Sub

        '曲线添加
        ''' <summary>
        ''' 由指定的数据列复制出一个新的数据列来，并在图表中显示
        ''' </summary>
        ''' <param name="SourceSeriesIndex">用来进行复制的原始数据列对象在集合中的下标值。
        ''' 注意：第一条曲线的下标值为1，而不是0。</param>
        ''' <remarks>此方法主要完成两步操作：
        ''' 1.在Excel图表中对数据列及其对应的对象进行复制；
        ''' 2.在字典索引dicSreies_Tag中新添加一项；
        ''' </remarks>
        Public Overridable Function CopySeries(ByVal SourceSeriesIndex As Integer) As Generic.KeyValuePair(Of Integer, SeriesTag)

            '生成新的数据列
            Dim newSeries As Series
            Dim NewSeriesIndex As Integer
            newSeries = AddSeries(NewSeriesIndex)

            '设置数据列格式
            Dim originalSeries As Series = F_DicSeries_Tag.Item(SourceSeriesIndex).series
            With newSeries
                '如果此处报错，则用 
                'Dim xv As Object = originalSeries.XValues, 然后将xv赋值给下面的XValue
                .XValues = originalSeries.XValues
                .Values = originalSeries.Values
                .Name = originalSeries.Name
            End With

            '设置数据列对应的Tag信息
            Me.F_DicSeries_Tag.Add(NewSeriesIndex, New SeriesTag(newSeries, Me.F_RollingDate))
            '
            Return New Generic.KeyValuePair(Of Integer, SeriesTag)(NewSeriesIndex, New SeriesTag(newSeries, Me.F_RollingDate))
        End Function

        ''' <summary>
        ''' 添加新的数据系列对象
        ''' </summary>
        ''' <param name="NewSeriesIndex">新添加的数据曲线在集合中的下标值</param>
        ''' <returns>一条新的数据曲线</returns>
        ''' <remarks>此函数只创建出对于新的数据曲线的对象索引，以及设置曲线的UI样式，
        ''' 并不涉及对于数据曲线的坐标点的赋值</remarks>
        Protected Function AddSeries(ByRef NewSeriesIndex As Integer) As Series
            Dim NewS As Series = Nothing
            If F_CurvesCount <= F_List_HasCurve.Count - 1 Then      '直接从已经定义好的模板中去提取
                Dim hasCurve As Boolean
                For NewSeriesIndex = cst_LboundOfSeriesInCollection To F_List_HasCurve.Count + cst_LboundOfSeriesInCollection - 1 Step 1
                    hasCurve = F_List_HasCurve(NewSeriesIndex - cst_LboundOfSeriesInCollection)
                    If Not hasCurve Then
                        NewS = Me.Chart.SeriesCollection.Item(NewSeriesIndex)
                        F_List_HasCurve.Item(NewSeriesIndex - cst_LboundOfSeriesInCollection) = True
                        Exit For
                    End If
                Next
            Else    '如果图表中的曲线数据已经大于当前数据系列集合中的series数量，那么就要新建一个数据系列，并设置其格式
                NewS = Me.Chart.SeriesCollection.NewSeries
                NewSeriesIndex = F_List_HasCurve.Count + cst_LboundOfSeriesInCollection
                F_List_HasCurve.Add(True)
                ''设置新数据系列的UI格式
                'With NewS

                'End With
            End If

            ' ----------------------------------- 对图例的显示进行操作
            Call LegendRefresh(F_List_HasCurve)

            '
            F_CurvesCount += 1
            Return NewS
        End Function
        ''' <summary>
        ''' 更新图例，图例的绝对尺寸的更新一定要在设置要Chart的尺寸之后，
        ''' 因为如果在设置好图例尺寸后再设置Chart尺寸，则图例尺寸会进行缩放。
        ''' </summary>
        ''' <param name="lst"></param>
        ''' <remarks></remarks>
        Protected Sub LegendRefresh(ByVal lst As List(Of Boolean))
            Dim lgd As Legend
            Dim lgdEntries As LegendEntries
            Dim lgdEnrty As LegendEntry
            With Me.Chart
                .HasLegend = False
                .HasLegend = True
                lgd = .Legend
                lgdEntries = lgd.LegendEntries
            End With

            '一定要注意，这里对图例项进行删除的时候，要从尾部开始向开头位置倒着删除:
            '因为集合的Index的索引方式, 当元素被删除后, 其后面的元素就接替了这个元素的下标值
            For LegendIndex As Integer = _
                F_List_HasCurve.Count - 1 + cst_LboundOfSeriesInCollection _
                To cst_LboundOfSeriesInCollection Step -1

                Dim hascurve As Boolean = F_List_HasCurve.Item(LegendIndex - cst_LboundOfSeriesInCollection)
                If Not hascurve Then
                    '在Visual Basic.NET中，如果用LegendEntries(Index)来索引LegendEntry对象，其第一个元素的下标值为0，
                    '而如果用LegendEntries.Item(Index)的方式来索引集合中的LegendEntry对象，则其第一个元素的下标值为1。
                    '而在VBA中，这两种方式索引集合中的LegendEntry对象，其第一个元素的下标值都是1；
                    lgdEnrty = lgdEntries.Item(LegendIndex)
                    lgdEnrty.Delete()
                End If
            Next LegendIndex
            ' 设置图例对象的位置与尺寸
            Call Me.SetLegendFormat(Me.Chart, lgd)
            Me.Sheet_Drawing.Range("A1").Activate()
        End Sub
        ''' <summary>
        ''' 设置图例对象的UI格式：图例方框的线型、背景、位置与大小等
        ''' </summary>
        ''' <param name="chart">图例对象所属的Chart对象</param>
        ''' <param name="legend">要进行UI格式设置的图例对象</param>
        ''' <remarks></remarks>
        Private Sub SetLegendFormat(ByVal chart As Chart, ByVal legend As Legend)
            '开始设置图例的格式
            '1、设置图例的格式
            With legend.Format
                With .Fill
                    .Visible = Office.MsoTriState.msoTrue
                    .ForeColor.RGB = RGB(255, 255, 255)     '图例对象的背景色
                End With
                .Shadow.Type = Microsoft.Office.Core.MsoShadowType.msoShadow21      '图例的阴影
                '.Line.ForeColor.RGB = RGB(0, 0, 0)      ' 图例对象的边框
                .TextFrame2.TextRange.Font.Name = AMEApplication.FontName_TNR
            End With
            '2、设置图例的位置与大小
            '图例方框的高度
            Dim ExpectedHeight As Single = Me.Legend_Location.Legend_Height ' Data_Drawing_Format.Drawing_Mnt_RollingBase.Legend_Height
            legend.Select()
            '对于Chart中的图例的位置与大小，其设置的原则是：
            '对于Top值，其始终不改变图例的大小，而将图例对象一直向下推进，直到图例对象的下边缘到达Chart的下边缘为止；
            '对于Height值，其始终不改变图例的Top的位置，而将图例对象的下边缘一直向下拉伸，直到图例对象的下边缘到达Chart的下边缘为止；
            '所以，如果要将图例对象成功地定位到Chart底部，应该先设置其Height值为0，然后设置其Top值，最后再回过来设置其Height值。
            With legend.Application.Selection
                .Left = 0
                .Width = Me.Legend_Location.Legend_Width
                .Height = 0     '以此来让Selection的Top值成功定位
                .Top = chart.ChartArea.Height - ExpectedHeight
                .Height = ExpectedHeight
            End With
        End Sub
#End Region
    End Class
End Namespace