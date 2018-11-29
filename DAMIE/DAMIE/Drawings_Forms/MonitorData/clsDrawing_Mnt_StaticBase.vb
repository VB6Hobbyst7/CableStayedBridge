Imports Microsoft.Office.Interop.Excel
Imports DAMIE.Miscellaneous
Imports DAMIE.Constants
Imports DAMIE.All_Drawings_In_Application
Imports DAMIE.GlobalApp_Form
Public MustInherit Class clsDrawing_Mnt_StaticBase
    Inherits ClsDrawing_Mnt_Base

#Region "  ---  Declarations & Definitions"

#Region "  ---  Properties"

    Protected MustOverride Overrides Property ChartSize_sugested As ChartSize

#End Region

#Region "  ---  Fields"

    ''' <summary>
    ''' 此工作表中的整个施工日期的数组（0-Based，数据类型为Date）
    ''' </summary>
    ''' <remarks></remarks>
    Protected F_arrAllDate() As Double

    ''' <summary>
    ''' 以每一条Series对象来索引此数据系列中的Y轴的数据，
    ''' 在表示Y轴数据的Object()中，其元素的个数必须要与F_arrAllDate中的元素个数相等。
    ''' </summary>
    ''' <remarks></remarks>
    Protected F_dicSeries As Dictionary(Of Series, Object())

#End Region

#End Region

#Region "  ---  构造函数与窗体的加载、打开与关闭"

    ''' <summary>
    ''' 构造函数，构造时一定要设置好字典F_dicFourSeries的值。
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
    Public Sub New(ByVal DataSheet As Worksheet, ByVal DrawingChart As Chart, ByVal ParentApp As Cls_ExcelForMonitorDrawing, _
                   ByVal type As DrawingType, ByVal CanRoll As Boolean, ByVal Info As TextFrame2, _
                   ByVal DrawingTag As MonitorInfo, ByVal MonitorType As MntType, _
                   ByVal Alldate As Double())
        Call MyBase.New(DataSheet, DrawingChart, ParentApp, type, CanRoll, Info, DrawingTag, MonitorType)
        '
        Me.F_arrAllDate = Alldate
        Me.currentPointsCount = Alldate.Length
    End Sub

#End Region

    ''' <summary>
    ''' Excel图表中，静态曲线图的数据系列中，每条曲线中所显示的数据点个数
    ''' </summary>
    ''' <remarks></remarks>
    Private currentPointsCount As Integer
    Private WithEvents myChart As Chart = Me.Chart
    Private Sub Chart_DoubleClick(elementID As Integer, arg1 As Integer, _arg2 As Integer, _
                          ByRef cancel As Boolean) Handles myChart.BeforeDoubleClick
        '控制界面显示
        cancel = True       '表示此事件屏蔽默认的双击事件
        Call SpeedMode(myChart, F_arrAllDate, Me.F_dicSeries)
    End Sub


    Protected Sub SpeedMode(ByVal Chart As Chart, ByVal arrAllDate As Double(), ByVal dicSeries As Dictionary(Of Series, Object()))
        Dim sht As Worksheet = Chart.Parent.Parent
        sht.Range("A1").Activate()      '取消图表上的对象的选择
        '执行SpeedMode相关的事件
        '
        Dim startday As Double = arrAllDate.First
        Dim endday As Double = arrAllDate.Last
        Dim AllDateCount As Integer = arrAllDate.Length
        Dim pointCount As Integer
        ' 
        Dim strPointsCount As String
        Dim strInputBoxTitle As String = "Speed Mode"

        strPointsCount = Me.Application.InputBox("设置曲线中显示的测点个数" & vbCrLf _
                                         & "当前记录天数为" & currentPointsCount & "天" & vbCrLf _
                                         & "最大记录天数为" & UBound(arrAllDate) + 1 & "天", _
                                        strInputBoxTitle, Type:=1)
        Try
            pointCount = CInt(strPointsCount)         '可能会出现数据类型转换异常
            If pointCount >= AllDateCount Then pointCount = AllDateCount '开始执行操作
            currentPointsCount = pointCount
            '获取按指定点数进行划分的时间区段长度
            Dim unit As Single = AllDateCount / pointCount

            '记录要绘制的日期的数组，数组中记录这些天对应的列号
            Dim slctCol(0 To pointCount - 1) As Integer
            Dim slctDate(0 To pointCount - 1) As Double

            '-------------------- 按指定的日期间隔得到对应的日期与日期在数组中的列号
            Dim referenceColumn As Double = 0
            Dim iselected As Integer = 0
            For icol As Integer = 0 To AllDateCount - 1
                '此处默认所有施工日期的数据中的日期是按从小到大的顺序进行排列的。
                '不果不是按此方法排列，则得到的结果中监测数据还是能与日期对应上，但是选择的日期可能会比较混乱
                '不具有均匀分布的特征
                If icol >= referenceColumn Then
                    slctCol(iselected) = icol
                    slctDate(iselected) = arrAllDate(icol)
                    iselected += 1          '满足条件的记录结果位置+1
                    referenceColumn = referenceColumn + unit
                End If

            Next                '下一天

            '按上面得到的列号来构造新的数组，以得到新的日期排列下的每一条曲线的数据
            Dim irow As Byte = 0
            Dim DataInSelectedDays(0 To UBound(slctCol)) As Object
            For Each curve As Series In dicSeries.Keys
                Dim idata As Integer = 0
                For Each selectedColumn As Integer In slctCol
                    DataInSelectedDays(idata) = dicSeries.Item(curve)(selectedColumn)
                    idata += 1
                Next
                With curve
                    .XValues = slctDate
                    .Values = DataInSelectedDays
                End With
                irow += 1
            Next
        Catch ex As Exception
            'MessageBox.Show("输入的格式不是合法的数值格式，请重新输入", "tip", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub

End Class
