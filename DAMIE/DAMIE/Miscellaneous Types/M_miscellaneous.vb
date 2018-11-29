Imports System.Math
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices
Imports System.Data.OleDb
Imports System.IO
Imports DAMIE.All_Drawings_In_Application
Imports eZstd.eZAPI

Imports DAMIE.Constants
Namespace Miscellaneous

#Region "  ---  全局模块:一些杂项的方法"

    ''' <summary>
    ''' 一些杂项的方法
    ''' </summary>
    ''' <remarks>记录了一些用来处理零碎问题的共享方法</remarks>
    Public Module GeneralMethods

        ''' <summary>
        ''' 进行CAD坐标与Visio坐标的坐标变换
        ''' </summary>
        ''' <param name="CAD_x1"></param>
        ''' <param name="CAD_y1"></param>
        ''' <param name="CAD_x2"></param>
        ''' <param name="CAD_y2"></param>
        ''' <param name="Visio_x1"></param>
        ''' <param name="Visio_y1"></param>
        ''' <param name="Visio_x2"></param>
        ''' <param name="Visio_y2"></param>
        ''' <returns>返回x与y方向的线性变换的斜率与截距</returns>
        ''' <remarks>其基本公式为：x_Visio=Kx*x_CAD+Cx；y_Visio=Ky*y_CAD+Cy</remarks>
        Public Function Coordinate_Conversion(CAD_x1 As Double, CAD_y1 As Double, _
                                                     CAD_x2 As Double, CAD_y2 As Double, _
                                                     Visio_x1 As Double, Visio_y1 As Double, _
                                                     Visio_x2 As Double, Visio_y2 As Double) As Cdnt_Cvsion
            Dim conversion As Cdnt_Cvsion
            With conversion
                .kx = (Visio_x1 - Visio_x2) / (CAD_x1 - CAD_x2)
                .cx = Visio_x1 - .kx * CAD_x1
                .ky = (Visio_y1 - Visio_y2) / (CAD_y1 - CAD_y2)
                .cy = Visio_y1 - .ky * CAD_y1
            End With
            Return conversion
        End Function
        ''' <summary>
        ''' CAD坐标系与Visio坐标系进行线性转换的斜率与截距
        ''' </summary>
        ''' <remarks>其基本公式为：x_Visio=Kx*x_CAD+Cx；y_Visio=Ky*y_CAD+Cy</remarks>
        Public Structure Cdnt_Cvsion
            Public kx As Double
            Public cx As Double
            Public ky As Double
            Public cy As Double
        End Structure

        ''' <summary>
        ''' 将两个TimeSpan向量进行比较，以扩展为二者区间的并集
        ''' </summary>
        ''' <param name="DateSpan_Larger">timespan1</param>
        ''' <param name="DateSpan_Shorter">timespan2</param>
        ''' <returns>返回扩展后的timespan</returns>
        ''' <remarks>将两个timespan的区间进行组合，取组合后的最小值与最大值</remarks>
        Public Function ExpandDateSpan(ByVal DateSpan_Larger As DateSpan, ByVal DateSpan_Shorter As DateSpan) As DateSpan
            If Date.Compare(DateSpan_Larger.StartedDate, DateSpan_Shorter.StartedDate) = 1 Then DateSpan_Larger.StartedDate = DateSpan_Shorter.StartedDate
            If Date.Compare(DateSpan_Larger.FinishedDate, DateSpan_Shorter.FinishedDate) = -1 Then DateSpan_Larger.FinishedDate = DateSpan_Shorter.FinishedDate
            Return DateSpan_Larger
        End Function

        ''' <summary>
        ''' 返回一个全局的唯一的ID值，用来定义每一个绘图的图表对象
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>这里先暂且就用当前时间的Tick值来定义</remarks>
        Public Function GetUniqueID() As Long
            Return Date.Now.Ticks
        End Function

        ''' <summary>
        ''' 委托：更新：基坑ID列表框、施工进度列表框
        ''' </summary>
        ''' <remarks></remarks>
        Private Delegate Sub RefreshComoboxHandler(ByVal comboBox As ListControl, ByVal DataSource As LstbxDisplayAndItem())
        ''' <summary>
        ''' 在主UI线程中更新：基坑ID列表框、施工进度列表框
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub RefreshCombobox(ByVal comboBox As ListControl, ByVal DataSource As LstbxDisplayAndItem())
            With comboBox
                If .InvokeRequired Then     '非UI线程，再次封送该方法到UI线程
                    '对于有输入参数的方法，要将其参数放置在一个数组中，统一传递给BeginInvoke方法。
                    Dim args As Object() = {comboBox, DataSource}
                    .BeginInvoke(New RefreshComoboxHandler(AddressOf RefreshCombobox), args)
                Else
                    '方式一：
                    .DataSource = DataSource
                    '方式二：
                    '.Items.Clear()
                    '.Items.AddRange(DataSource)
                    '
                    .DisplayMember = LstbxDisplayAndItem.DisplayMember
                    If DataSource.Count > 0 Then
                        .SelectedIndex = 0
                    End If
                End If
            End With
        End Sub


        ''' <summary>
        ''' 通过释放窗口的"最大化"按钮及"拖拽窗口"的功能，来达到固定应用程序窗口大小的效果
        ''' </summary>
        ''' <param name="hWnd">要释放大小的窗口的句柄</param>
        ''' <remarks></remarks>
        Public Sub FixWindow(ByVal hWnd As IntPtr)
            Dim hStyle As Integer = APIWindows.GetWindowLong(hWnd, WindowLongFlags.GWL_STYLE)
            '禁用最大化的标头及拖拽视窗
            APIWindows.SetWindowLong(hWnd, WindowLongFlags.GWL_STYLE, hStyle And Not WindowStyle.WS_MAXIMIZEBOX And Not WindowStyle.WS_EX_APPWINDOW)
        End Sub

        ''' <summary>
        ''' 通过禁用窗口的"最大化"按钮及"拖拽窗口"的功能，来达到释放应用程序窗口大小的效果
        ''' </summary>
        ''' <param name="hWnd">要固定大小的窗口的句柄</param>
        ''' <remarks></remarks>
        Public Sub unFixWindow(ByVal hWnd As IntPtr)
            Dim hStyle As Integer = APIWindows.GetWindowLong(hWnd, WindowLongFlags.GWL_STYLE)
            APIWindows.SetWindowLong(hWnd, WindowLongFlags.GWL_STYLE, hStyle Or WindowStyle.WS_MAXIMIZEBOX Or WindowStyle.WS_EX_APPWINDOW)
        End Sub

        ''' <summary>
        ''' 计算数组中的最大值，以及最大值在数组中对应的下标（第一个元素的下标为0）。
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="arr"></param>
        ''' <param name="Index">最大值在数组中的下标位置，第一个元素的下标值为0</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Max_Array(Of T)(ByVal arr() As T, Optional ByRef Index As Integer = 0) As T
            Index = 0
            Dim i As Integer = 0
            Dim max As T = arr(Index)
            Dim refer As IComparable(Of T) = arr(Index)
            For Each v As IComparable(Of T) In arr
                If v.CompareTo(refer) > 0 Then
                    max = v
                    refer = max
                    Index = i
                End If
                i += 1
            Next
            Return max
        End Function

        ''' <summary>
        ''' 计算数组中的最小值，以及最小值在数组中对应的下标（第一个元素的下标为0）。
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="arr"></param>
        ''' <param name="Index">最小值在数组中的下标位置，第一个元素的下标值为0</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Min_Array(Of T)(ByVal arr() As T, Optional ByRef Index As Integer = 0) As T
            Index = 0
            Dim min As T = arr(Index)
            Dim i As Integer = 0
            Dim refer As IComparable(Of T) = arr(Index)
            For Each v As IComparable(Of T) In arr
                If v.CompareTo(refer) < 0 Then
                    min = v
                    refer = min
                    Index = i
                End If
                i += 1
            Next
            Return min
        End Function

        ''' <summary>
        ''' 设置列表控件中显示监测数据的类型，并索引到具体的枚举项
        ''' </summary>
        ''' <param name="listControl"></param>
        ''' <remarks></remarks>
        Public Sub SetMonitorType(ByVal listControl As ListControl)
            Dim arrMntTypes() As LstbxDisplayAndItem
            arrMntTypes = { _
                New LstbxDisplayAndItem("通用", MntType.General), _
                New LstbxDisplayAndItem("墙体测斜", MntType.Incline), _
                New LstbxDisplayAndItem("立柱垂直位移", MntType.Column), _
                New LstbxDisplayAndItem("支撑轴力", MntType.Struts), _
                New LstbxDisplayAndItem("墙顶水平位移", MntType.WallTop_Horizontal), _
                New LstbxDisplayAndItem("墙顶竖直位移", MntType.WallTop_Vertical), _
                New LstbxDisplayAndItem("地表垂直位移", MntType.EarthSurface), _
                New LstbxDisplayAndItem("水位", MntType.WaterLevel)}
            With listControl
                .DisplayMember = LstbxDisplayAndItem.DisplayMember
                .DataSource = arrMntTypes
            End With
        End Sub

    End Module

#End Region

#Region "  ---  枚举值"

    ''' <summary>
    ''' 各种与RGB函数返回的值相同的颜色值
    ''' </summary>
    ''' <remarks>RGB函数返回的整数值与RGB参量的换算关系为：Color属性值=R + 256*G + 256^2*B</remarks>
    Public Enum Color As Integer

        ''' <summary>
        ''' 表示在开挖剖面图中，已经开挖到基坑底部，正在向上建时的颜色：RGB(20, 200, 230)
        ''' </summary>
        ''' <remarks>为偏暗的青色</remarks>
        Color_BuildindUp = 15124500

        ''' <summary>
        ''' 表示在开挖剖面图中，还未开挖到基坑底部标高时的颜色：RGB(220, 160, 0)
        ''' </summary>
        ''' <remarks>为偏暗的黄色</remarks>
        Color_DiggingDown = 41180

        ' ''' <summary>
        ' ''' 表示测斜曲线图中，第1条监测曲线的颜色：RGB(0, 0, 0)
        ' ''' </summary>
        ' ''' <remarks>对于Chart中的数据系列的颜色进行特别的设定，是因为调用模板中数据系列的颜色时，
        ' ''' 它返回的是“自动”，但是在赋值时却不能将颜色值赋值为“自动”。</remarks>
        'Inline1 = 0 ' RGB(0 ,0 ,0)
        'Inline2 = 65280 ' RGB(0 ,255 ,0)
        'Inline3 = 65535 ' RGB(255 ,255 ,0)
        'Inline4 = 3305658   ' RGB(186 ,112 ,50)
        'Inline5 = 3685008   ' RGB(144 ,58 ,56)
        'Inline6 = 4295796   ' RGB(116 ,140 ,65)
        'Inline7 = 4803071   ' RGB(255 ,73 ,73)
        'Inline8 = 4841471   ' RGB(255 ,223 ,73)
        'Inline9 = 4849545   ' RGB(137 ,255 ,73)
        'Inline10 = 7817053  ' RGB(93 ,71 ,119)
        'Inline11 = 9606097  ' RGB(209 ,147 ,146)
        'Inline12 = 16711680 ' RGB(0 ,0 ,255)
        'Inline13 = 16711935 ' RGB(255 ,0 ,255)
        'Inline14 = 16713537 ' RGB(65 ,7 ,255)
        'Inline15 = 16733769 ' RGB(73 ,86 ,255)
        'Inline16 = 16776960 ' RGB(0 ,255 ,255)
    End Enum

    ''' <summary>
    ''' 枚举Excel中各种对象在集合中的第一个元素的索引下标值
    ''' </summary>
    ''' <remarks>从概念上来说，这里不应该用枚举类型，而应该用常数类型。但是由于这里的所有项的值都是整数值，
    ''' 而且值的范围大概都是非0即1，所以这里用枚举值来表示也是完全可以的。
    ''' 一定要注意的是，这里的每一个枚举项都必须要给出对应的下标数值来！！！</remarks>
    Public Enum LowIndexOfObjectsInExcel As SByte
        ''' <summary>
        ''' 在Excel中用Range.Value返回的二维数组的第一个元素的下标值
        ''' </summary>
        ''' <remarks></remarks>
        ObjectArrayFromRange_Value = 1
        ''' <summary>
        ''' Chart图表中的数据列集合中，第一条曲线对应的下标值
        ''' </summary>
        ''' <remarks></remarks>
        SeriesInSeriesCollection = 1
    End Enum

    ''' <summary>
    ''' 对项目文件执行的操作方式
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum ProjectState
        ''' <summary>
        ''' 新建项目
        ''' </summary>
        ''' <remarks></remarks>
        NewProject
        ''' <summary>
        ''' 打开项目
        ''' </summary>
        ''' <remarks></remarks>
        OpenProject
        ''' <summary>
        ''' 编辑项目
        ''' </summary>
        ''' <remarks></remarks>
        EditProject

    End Enum

    ''' <summary>
    ''' .NET中的数据类型枚举
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum DataType
        ''' <summary>
        ''' 整数
        ''' </summary>
        ''' <remarks></remarks>
        Type_Integer
        ''' <summary>
        ''' 单精度
        ''' </summary>
        ''' <remarks></remarks>
        Type_Single
        ''' <summary>
        ''' 双精度
        ''' </summary>
        ''' <remarks></remarks>
        Type_Double
        ''' <summary>
        ''' 日期
        ''' </summary>
        ''' <remarks></remarks>
        Type_Date
        ''' <summary>
        ''' 字符串
        ''' </summary>
        ''' <remarks></remarks>
        Type_String
        ''' <summary>
        ''' 对象型
        ''' </summary>
        ''' <remarks></remarks>
        Type_Object

    End Enum

#End Region

#Region "  ---  类或接口"

    ''' <summary>
    ''' 用来作为ListControl类的.Add方法中的Item参数的类。通过指定ListControl类的DisplayMember属性，来设置列表框中显示的文本。 
    ''' </summary>
    ''' <remarks>
    ''' 保存数据时：
    '''  With ListBoxWorksheetsName
    '''       .DisplayMember = LstbxDisplayAndItem.DisplayMember
    '''       .ValueMember = LstbxDisplayAndItem.ValueMember
    '''       .DataSource = arrSheetsName   '  Dim arrSheetsName(0 To sheetsCount - 1) As LstbxDisplayAndItem
    '''  End With
    ''' 提取数据时：
    '''  Try
    '''      Me.F_shtMonitorData = DirectCast(Me.ListBoxWorksheetsName.SelectedValue, Worksheet)
    '''  Catch ex As Exception
    '''      Me.F_shtMonitorData = Nothing
    '''  End Try
    ''' 或者是：
    '''  Dim lst As LstbxDisplayAndItem = Me.ComboBoxOpenedWorkbook.SelectedItem
    '''  Try
    '''     Dim Wkbk As Workbook = DirectCast(lst.Value, Workbook)
    '''  Catch ex ...
    ''' </remarks>
    Public Class LstbxDisplayAndItem

        ''' <summary>
        ''' 在列表框中进行显示的文本
        ''' </summary>
        ''' <remarks>此常数的值代表此类中代表要在列表框中显示的文本的属性名，即"DisplayedText"</remarks>
        Public Const DisplayMember As String = "DisplayedText"
        ''' <summary>
        ''' 列表框中每一项对应的值（任何类型的值）
        ''' </summary>
        ''' <remarks>此常数的值代表此类中代表列表框中的每一项绑定的数据的属性名，即"Value"</remarks>
        Public Const ValueMember As String = "Value"

        Private _objValue As Object
        Public ReadOnly Property Value As Object
            Get
                Return _objValue
            End Get
        End Property

        Private _DisplayedText As String
        Public ReadOnly Property DisplayedText As String
            Get
                Return _DisplayedText
            End Get
        End Property

        ''' <summary>
        ''' 构造函数
        ''' </summary>
        ''' <param name="DisplayedText">用来显示在列表的UI界面中的文本</param>
        ''' <param name="Value">列表项对应的值</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal DisplayedText As String, ByVal Value As Object)
            Me._objValue = Value
            Me._DisplayedText = DisplayedText
        End Sub 'New 

        '表示“什么也没有”的枚举值
        ''' <summary>
        ''' 列表框中用来表示“什么也没有”。
        ''' 1、在声明时：listControl控件.Items.Add(New LstbxDisplayAndItem(" 无", NothingInListBox.None))
        ''' 2、在选择列表项时：listControl控件.SelectedValue = NothingInListBox.None
        ''' 3、在读取列表中的数据时，作出判断：If Not LstbxItem.Value.Equals(NothingInListBox.None) Then ...
        ''' </summary>
        ''' <remarks></remarks>
        Enum NothingInListBox
            ''' <summary>
            ''' 什么也没有选择
            ''' </summary>
            ''' <remarks></remarks>
            None
        End Enum

    End Class 'LstbxItem

    ''' <summary>
    ''' 此集合的最大特点就是：它的每一个元素的“键”都是一个Integer的值，
    ''' 而这个键是在元素添加到集合中时自动生成的。此集合中的值的类型中，都含有一个代表此元素在集合中的键的属性（比如ID），
    ''' 由此可以直接通过此元素来索引到它在集合中的位置。
    ''' </summary>
    ''' <typeparam name="TValue">集合中元素的值，值的类型必须要实现“I_Dictionary_AutoKey”接口</typeparam>
    ''' <remarks>  </remarks>
    Public Class Dictionary_AutoKey(Of TValue)
        Inherits System.Collections.Generic.Dictionary(Of Integer, TValue)
        ''' <summary>
        ''' 集合类Dictionary_AutoKey中的元素所必须要实现的接口。
        ''' </summary>
        ''' <remarks></remarks>
        Interface I_Dictionary_AutoKey
            'ReadOnly Property Parent As Dictionary_AutoKey(Of TValue)

            ''' <summary>
            ''' 此元素在其所在的集合中的键，这个键是在元素添加到集合中时自动生成的，
            ''' 所以应该在执行集合.Add函数时，用元素的Key属性接受函数的输出值。
            ''' 在集合中添加此元素：Me.Key=Me所在的集合.Add(Me)
            ''' 在集合中索引此元素：集合.item(me.key)
            ''' </summary>
            ''' <value></value>
            ''' <returns></returns>
            ''' <remarks></remarks>
            ReadOnly Property Key As Integer
        End Interface

        ''' <summary>
        ''' 构造函数：判断此集合中的元素类型是否实现了接口I_Dictionary_AutoKey。
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub New()
            Dim T As Type = GetType(TValue)
            If T.GetInterface(GetType(I_Dictionary_AutoKey).Name) Is Nothing Then
                '说明此集合中的元素没有实现接口：I_Dictionary_AutoKey
                MessageBox.Show("集合中的类没有实现接口，请检查!!!" & vbCrLf & "程序编写完成后将此判断删除！！！", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End Sub

        ''' <summary>
        ''' 用法：Me.Key=Me所在的集合.Add(Me)
        ''' </summary>
        ''' <param name="Value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shadows Function Add(Value As TValue) As Integer
            Dim ItemID As Integer = NewID()
            MyBase.Add(ItemID, Value)
            Return ItemID
        End Function

        ''' <summary>
        ''' 在集合中添加对象时，自动执行此方法，以返回对象元素在集合中的键。
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function NewID() As Integer
            Static myID As Integer = 0
            Try
                myID += 1
            Catch ex As OverflowException
                MessageBox.Show("集合中的元素数量溢出")
            End Try
            Return myID
        End Function

    End Class       'Collection_Generic

    ''' <summary>
    ''' 表示一段施工日期的跨度区间
    ''' </summary>
    ''' <remarks></remarks>
    Public Structure DateSpan
        ''' <summary>
        ''' 施工段的起始日期
        ''' </summary>
        ''' <remarks></remarks>
        Public Property StartedDate As Date
        ''' <summary>
        ''' 施工段的结束日期
        ''' </summary>
        ''' <remarks></remarks>
        Public Property FinishedDate As Date

        ''' <summary>
        ''' 检查DateSpan的跨度中，是否包含指定的日期
        ''' </summary>
        ''' <param name="dt"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Contains(dt As Date) As Boolean
            If dt.CompareTo(Me.StartedDate) < 0 OrElse dt.CompareTo(Me.FinishedDate) > 0 Then
                Return False
            Else
                Return True
            End If
        End Function
    End Structure       'DateSpan

    ''' <summary>
    ''' 利用ADO.NET连接Excel数据库，并执行相应的操作：  
    ''' 创建表格，读取数据，写入数据，获取工作簿中的所有工作表名称。
    ''' </summary>
    ''' <remarks></remarks>
    Public Class AdoForExcel

        ''' <summary>
        ''' 创建对Excel工作簿的连接
        ''' </summary>
        ''' <param name="ExcelWorkbookPath">要进行连接的Excel工作簿的路径</param>
        ''' <returns>一个OleDataBase的Connection连接，此连接还没有Open。</returns>
        ''' <remarks></remarks>
        Public Shared Function ConnectToExcel(ByVal ExcelWorkbookPath As String) As OleDbConnection
            Dim strConn As String = String.Empty
            If ExcelWorkbookPath.EndsWith("xls") Then

                strConn = "Provider=Microsoft.Jet.OLEDB.4.0; " +
                           "Data Source=" + ExcelWorkbookPath + "; " +
                           "Extended Properties='Excel 8.0;IMEX=1'"

            ElseIf ExcelWorkbookPath.EndsWith("xlsx") OrElse ExcelWorkbookPath.EndsWith("xlsb") Then

                strConn = "Provider=Microsoft.ACE.OLEDB.12.0;" +
                          "Data Source=" + ExcelWorkbookPath + ";" +
                          "Extended Properties=""Excel 12.0;HDR=YES"""
            End If
            Dim conn As OleDbConnection = New OleDbConnection(strConn)
            Return conn
        End Function

        ''' <summary>
        ''' 从对于Excel的数据连接中获取Excel工作簿中的所有工作表
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <returns>如果此连接不是连接到Excel数据库，则返回Nothing</returns>
        ''' <remarks></remarks>
        Public Shared Function GetSheetsName(ByVal conn As OleDbConnection) As String()
            '如果连接已经关闭，则先打开连接
            If conn.State = ConnectionState.Closed Then conn.Open()
            If ConnectionSourceValidated(conn) Then
                '获取工作簿连接中的每一个工作表，
                '注意下面的Rows属性返回的并不是Excel工作表中的每一行，而是Excel工作簿中的所有工作表。
                Dim Tables As DataRowCollection = conn.GetSchema("Tables").Rows
                '
                Dim sheetNames(0 To Tables.Count - 1) As String
                For i As Integer = 0 To Tables.Count - 1
                    '注意这里的表格Table是以DataRow的形式出现的。
                    Dim Tb As DataRow = Tables.Item(i)
                    Dim Tb_Name As Object = Tb.Item("TABLE_NAME")
                    sheetNames(i) = Tb_Name.ToString
                Next
                Return sheetNames
            Else
                MessageBox.Show("未正确连接到Excel数据库!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return Nothing
            End If
        End Function

        ''' <summary>
        ''' 创建一个新的Excel工作表，并向其中插入一条数据
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="TableName">要新创建的工作表名称</param>
        ''' <remarks></remarks>
        Public Shared Sub CreateNewTable(ByVal conn As OleDbConnection, ByVal TableName As String)
            '如果连接已经关闭，则先打开连接
            If conn.State = ConnectionState.Closed Then conn.Open()
            If ConnectionSourceValidated(conn) Then
                Using ole_cmd As OleDbCommand = conn.CreateCommand()

                    '----- 生成Excel表格 --------------------
                    '要新创建的表格不能是在Excel工作簿中已经存在的工作表。
                    ole_cmd.CommandText = "CREATE TABLE CustomerInfo ([" + TableName + "] VarChar,[Customer] VarChar)"
                    Try
                        '在工作簿中创建新表格时，Excel工作簿不能处于打开状态
                        ole_cmd.ExecuteNonQuery()
                    Catch ex As Exception
                        MessageBox.Show("创建Excel文档 " + TableName + "失败，错误信息： " + ex.Message, _
                                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End Try
                End Using
            Else
                MessageBox.Show("未正确连接到Excel数据库!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
        End Sub

        ''' <summary>
        ''' 向Excel工作表中插入一条数据，此函数并不通用，不建议使用
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <param name="TableName">要插入数据的工作表名称</param>
        ''' <param name="FieldName">要插入到的字段</param>
        ''' <param name="Value">实际插入的数据</param>
        ''' <remarks></remarks>
        Public Shared Sub InsertToTable(ByVal conn As OleDbConnection, ByVal TableName As String, _
                                       ByVal FieldName As String, ByVal Value As Object)
            '如果连接已经关闭，则先打开连接
            If conn.State = ConnectionState.Closed Then conn.Open()
            If ConnectionSourceValidated(conn) Then
                Using ole_cmd As OleDbCommand = conn.CreateCommand()
                    '在插入数据时，字段名必须是数据表中已经有的字段名，插入的数据类型也要与字段下的数据类型相符。

                    Try
                        ole_cmd.CommandText = "insert into [" & TableName & "$](" + FieldName & ") values('" & Value & "')"
                        '这种插入方式在Excel中的实时刷新的，也就是说插入时工作簿可以处于打开的状态，
                        '而且这里插入后在Excel中会立即显示出插入的值。
                        ole_cmd.ExecuteNonQuery()
                    Catch ex As Exception
                        MessageBox.Show("数据插入失败，错误信息： " & ex.Message)
                        Exit Sub
                    End Try
                End Using
            Else
                MessageBox.Show("未正确连接到Excel数据库!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
        End Sub

        ''' <summary>
        ''' 读取Excel工作簿中的数据
        ''' </summary>
        ''' <param name="conn">OleDB的数据连接</param>
        ''' <param name="SheetName">要读取的数据所在的工作表</param>
        ''' <param name="FieldName">在读取的字段</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetDataFromExcel(ByVal conn As OleDbConnection, ByVal SheetName As String, ByVal FieldName As String) As String()
            '如果连接已经关闭，则先打开连接
            If conn.State = ConnectionState.Closed Then conn.Open()
            If ConnectionSourceValidated(conn) Then
                '创建向数据库发出的指令
                Dim olecmd As OleDbCommand = conn.CreateCommand()
                '类似SQL的查询语句这个[Sheet1$对应Excel文件中的一个工作表]  
                '如果要提取Excel中的工作表中的某一个指定区域的数据，可以用："select * from [Sheet3$A1:C5]"
                olecmd.CommandText = "select * from [" & SheetName & "$]"

                '创建数据适配器——根据指定的数据库指令
                Dim Adapter As OleDbDataAdapter = New OleDbDataAdapter(olecmd)
                '创建一个数据集以保存数据
                Dim dtSet As DataSet = New DataSet()
                '将数据适配器按指令操作的数据填充到数据集中的某一工作表中（默认为“Table”工作表）
                Adapter.Fill(dtSet)
                '其中的数据都是由 "select * from [" + SheetName + "$]"得到的Excel中工作表SheetName中的数据。
                Dim intTablesCount As Integer = dtSet.Tables.Count
                '索引数据集中的第一个工作表对象
                Dim DataTable As System.Data.DataTable = dtSet.Tables(0) ' conn.GetSchema("Tables")
                '工作表中的数据有8列9行(它的范围与用Worksheet.UsedRange所得到的范围相同。
                '不一定是写有数据的单元格才算进行，对单元格的格式，如底纹，字号等进行修改的单元格也在其中。)
                Dim intRowsInTable As Integer = DataTable.Rows.Count
                Dim intColsInTable As Integer = DataTable.Columns.Count
                '提取每一行数据中的“成绩”数据
                Dim Data(0 To intRowsInTable - 1) As String
                For i As Integer = 0 To intRowsInTable - 1
                    Data(i) = DataTable.Rows(i)(FieldName).ToString()
                Next
                Return Data
            Else
                MessageBox.Show("未正确连接到Excel数据库!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return Nothing
            End If
        End Function

        '私有函数
        ''' <summary>
        ''' 验证连接的数据源是否是Excel数据库
        ''' </summary>
        ''' <param name="conn"></param>
        ''' <returns>如果是Excel数据库，则返回True，否则返回False。</returns>
        ''' <remarks></remarks>
        Private Shared Function ConnectionSourceValidated(ByVal conn As OleDbConnection) As Boolean
            '考察连接是否是针对于Excel文档的。
            Dim strDtSource As String = conn.DataSource        '"C:\Users\Administrator\Desktop\测试Visio与Excel交互\数据.xlsx"
            Dim strExt As String = Path.GetExtension(strDtSource)
            If strExt = ".xlsx" OrElse strExt = "xls" OrElse strExt = "xlsb" Then
                Return True
            Else
                Return False
            End If
        End Function

    End Class           'AdoForExcel

    Public Class ExcelFunction

#Region "  ---  Range到数组的转换"
        ''' <summary>
        ''' 将Excel的Range对象的数据转换为指定数据类型的一维向量,
        ''' Range.Value返回一个二维的表格，此函数将其数据按列行拼接为一维数组。
        ''' （即按(0,0),(0,1),(1,0),(1,1),(2,0)...的顺序排列）
        ''' </summary>
        ''' <param name="rg">用于提取数据的Range对象</param>
        ''' <returns>返回一个指定类型的一维向量，如“Single()”</returns>
        ''' <remarks>直接用Range.Value来返回的数据，其类型只能是Object，
        ''' 而其中的数据是一个元素类型为Object的二维数据（即使此Range对象只有一行或者一列）。
        ''' 所以要进行显式的转换，将其转换为指定类型的向量或者二维数组，以便于后面的数据操作。</remarks>
        Public Shared Function ConvertRangeDataToVector(Of T)(ByVal rg As Excel.Range) As T()
            'Range中的数据，这是以向量的形式给出的，其第一个元素的下标值很有可能是1，而不是0
            Dim RangeData As Object(,) = rg.Value
            Dim elementCount As Integer = RangeData.Length
            '
            Dim Value_Out(0 To elementCount - 1) As T
            '获取输入的数据类型
            Dim DestiType As Type = GetType(T)
            Dim TC As TypeCode = Type.GetTypeCode(DestiType)
            '判断此类型的值
            Select Case TC
                Case TypeCode.DateTime
                    Dim i As Integer = 0
                    For Each V As Object In RangeData
                        Try
                            Value_Out(i) = CType(V, T)
                        Catch ex As Exception
                            Debug.Print("数据：" & V.ToString & " 转换为日期出错！将其处理为日期的初始值。" _
                                            & vbCrLf & ex.Message)
                            '如果输入的数据为double类型，则将其转换为等效的Date
                            Dim O As Object = Date.FromOADate(V)
                            Value_Out(i) = CType(O, T)
                        Finally
                            i += 1
                        End Try
                    Next
                Case Else
                    Dim i As Integer = 0
                    For Each V As Object In RangeData
                        Value_Out(i) = V
                        i += 1
                    Next
            End Select
            Return Value_Out
        End Function

        ''' <summary>
        ''' 将Excel的Range对象的数据转换为指定数据类型的二维数组
        ''' </summary>
        ''' <param name="rg">用于提取数据的Range对象</param>
        ''' <returns>返回一个指定类型的二维数组，如“Single(,)”</returns>
        ''' <remarks>直接用Range.Value来返回的数据，其类型只能是Object，
        ''' 而其中的数据是一个元素类型为Object的二维数据（即使此Range对象只有一行或者一列）。
        ''' 所以要进行显式的转换，将其转换为指定类型的向量或者二维数组，以便于后面的数据操作。</remarks>
        Public Shared Function ConvertRangeDataToMatrix(Of T)(ByVal rg As Excel.Range) As T(,)

            Dim RangeData As Object(,) = rg.Value
            '由Range.Value返回的二维数组的第一个元素的下标值，这里应该为1，而不是0
            Dim LB_Range As Byte = LBound(RangeData)
            '
            Dim intRowsCount As Integer = UBound(RangeData, 1) - LB_Range + 1
            Dim intColumnsCount As Integer = UBound(RangeData, 2) - LB_Range + 1
            '
            Dim OutputMatrix(0 To intRowsCount - 1, 0 To intColumnsCount - 1) As T
            '
            Select Case Type.GetTypeCode(GetType(T))
                Case TypeCode.DateTime
                    For row As Integer = 0 To intRowsCount - 1
                        For col As Integer = 0 To intColumnsCount - 1
                            Try
                                OutputMatrix(row, col) = CType(RangeData(row + LB_Range, col + LB_Range), T)
                            Catch ex As Exception
                                MessageBox.Show("数据：" & RangeData(row + LB_Range, col + LB_Range).ToString & " 转换为日期出错！将其处理为日期的初始值。" _
                                    & vbCrLf & ex.Message)
                            End Try
                        Next
                    Next
                Case Else
                    For row As Integer = 0 To intRowsCount - 1
                        For col As Integer = 0 To intColumnsCount - 1
                            OutputMatrix(row, col) = RangeData(row + LB_Range, col + LB_Range)
                        Next
                    Next
            End Select
            Return OutputMatrix
        End Function
#End Region

        ''' <summary>
        ''' 将Excel表中的列的数值编号转换为对应的字符
        ''' </summary>
        ''' <param name="ColNum">Excel中指定列的数值序号</param>
        ''' <returns>以字母序号的形式返回指定列的列号</returns>
        ''' <remarks>1对应A；26对应Z；27对应AA</remarks>
        Public Shared Function ConvertColumnNumberToString(ByVal ColNum As Integer) As String
            ' 关键算法就是：连续除以基，直至商为0，从低到高记录余数！
            ' 其中value必须是十进制表示的数值
            'intLetterIndex的位数为digits=fix(log(value)/log(26))+1
            '本来算法很简单，但是要解决一个问题：当value为26时，其26进制数为[1 0]，这样的话，
            '以此为下标索引其下面的strAlphabetic时就会出错，因为下标0无法索引。实际上，这种特殊情况下，应该让所得的结果成为[26]，才能索引到字母Z。
            '处理的方法就是，当所得余数remain为零时，就将其改为26，然后将对应的商的值减1.

            If ColNum < 1 Then MsgBox("列数不能小于1")

            Dim intLetterIndex As New List(Of Byte)
            '
            Dim quotient As Integer          '商
            Dim remain As Byte                '余数
            '
            Dim i As Byte = 0
            Do
                quotient = Fix(ColNum / 26)          '商
                remain = ColNum Mod 26                  '余数
                If remain = 0 Then
                    intLetterIndex.Add(26)
                    quotient = quotient - 1
                Else
                    intLetterIndex.Add(remain)

                End If
                i = i + 1
                ColNum = quotient
            Loop Until quotient = 0
            Dim alphabetic As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
            Dim ans As String = ""
            For i = 0 To intLetterIndex.Count - 1
                ans = alphabetic.Chars(intLetterIndex(i) - 1) & ans
            Next
            Return ans
        End Function

        ''' <summary>
        ''' 将Excel表中的字符编号转换为对应的数值
        ''' </summary>
        ''' <param name="colString">以A1形式表示的列的字母序号，不区分大小写</param>
        ''' <returns>以整数的形式返回指定列的数值编号，A列对应数值1</returns>
        ''' <remarks></remarks>
        Public Shared Function ConvertColumnStringToNumber(ByVal colString As String) As Integer
            colString = colString.ToUpper
            Dim ASC_A As Byte = Asc("A")
            Dim ans As Integer = 0
            For i As Byte = 0 To colString.Length - 1
                Dim Chr As Char = colString.Substring(i, 1)
                ans = ans + (Asc(Chr) - ASC_A + 1) * Pow(26, colString.Length - i - 1)
            Next
            Return ans
        End Function

        ''' <summary>
        ''' 获取指定Range范围中的左下角的那一个单元格
        ''' </summary>
        ''' <param name="RangeForSearch">要进行搜索的Range区域</param>
        ''' <returns>指定Range区域中的左下角的单元格</returns>
        ''' <remarks></remarks>
        Public Shared Function GetBottomRightCell(ByVal RangeForSearch As Excel.Range) As Excel.Range
            Dim RowsCount As Integer, ColsCount As Integer
            Dim LeftTopCell As Excel.Range = RangeForSearch.Cells(1, 1)
            With RangeForSearch
                RowsCount = .Rows.Count
                ColsCount = .Columns.Count
            End With
            With RangeForSearch.Worksheet
                Return .Cells(LeftTopCell.Row + RowsCount - 1, LeftTopCell.Column + ColsCount - 1)
            End With
        End Function

#Region "  ---  工作簿或工作表的匹配"

        ''' <summary>
        ''' 比较两个工作表是否相同。
        ''' 判断的标准：先判断工作表的名称是否相同，如果相同，再判断工作表所属的工作簿的路径是否相同，
        ''' 如果二者都相同，则认为这两个工作表是同一个工作表
        ''' </summary>
        ''' <param name="sheet1">进行比较的第1个工作表对象</param>
        ''' <param name="sheet2">进行比较的第2个工作表对象</param>
        ''' <returns></returns>
        ''' <remarks>不用 blnSheetsMatched = sheet1.Equals(sheet2)，是因为这种方法并不能正确地返回True。</remarks>
        Public Shared Function SheetCompare(ByVal sheet1 As Excel.Worksheet, ByVal sheet2 As Excel.Worksheet) As Boolean
            Dim blnSheetsMatched As Boolean = False
            '先比较工作表名称
            If String.Compare(sheet1.Name, sheet2.Name) = 0 Then
                Dim wb1 As Excel.Workbook = sheet1.Parent
                Dim wb2 As Excel.Workbook = sheet2.Parent
                '再比较工作表所在工作簿的路径
                If String.Compare(wb1.FullName, wb2.FullName) = 0 Then
                    blnSheetsMatched = True
                End If
            End If
            Return blnSheetsMatched
        End Function

        ''' <summary>
        ''' 检测工作簿是否已经在指定的Application程序中打开。
        ''' 如果最后此工作簿在程序中被打开（已经打开或者后期打开），则返回对应的Workbook对象，否则返回Nothing。
        ''' 这种比较方法的好处是不会额外去打开已经打开过了的工作簿。
        ''' </summary>
        ''' <param name="wkbkPath">进行检测的工作簿</param>
        ''' <param name="Application">检测工作簿所在的Application程序</param>
        ''' <param name="blnFileHasBeenOpened">指示此Excel工作簿是否已经在此Application中被打开</param>
        ''' <param name="OpenIfNotOpened">如果此Excel工作簿并没有在此Application程序中打开，是否要将其打开。</param>
        ''' <param name="OpenByReadOnly">是否以只读方式打开</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function MatchOpenedWkbk(ByVal wkbkPath As String, _
                                               ByVal Application As Excel.Application, _
                                               Optional ByRef blnFileHasBeenOpened As Boolean = False, _
                                               Optional ByVal OpenIfNotOpened As Boolean = False, _
                                               Optional ByVal OpenByReadOnly As Boolean = True _
                                               ) As Excel.Workbook
            Dim wkbk As Excel.Workbook = Nothing
            If Application IsNot Nothing Then
                '进行具体的检测
                If File.Exists(wkbkPath) Then   '此工作簿存在
                    '如果此工作簿已经打开
                    For Each WkbkOpened As Excel.Workbook In Application.Workbooks
                        If String.Compare(WkbkOpened.FullName, wkbkPath, True) = 0 Then
                            wkbk = WkbkOpened
                            blnFileHasBeenOpened = True
                            Exit For
                        End If
                    Next

                    '如果此工作簿还没有在主程序中打开，则将此工作簿打开
                    If Not blnFileHasBeenOpened Then
                        If OpenIfNotOpened Then
                            wkbk = Application.Workbooks.Open(Filename:=wkbkPath, UpdateLinks:=False, [ReadOnly]:=OpenByReadOnly)
                        End If
                    End If
                End If
            End If
            '返回结果
            Return wkbk
        End Function

        ''' <summary>
        ''' 检测指定工作簿内是否有指定的工作表，如果存在，则返回对应的工作表对象，否则返回Nothing
        ''' </summary>
        ''' <param name="wkbk">进行检测的工作簿对象</param>
        ''' <param name="sheetName">进行检测的工作表的名称</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function MatchWorksheet(ByVal wkbk As Excel.Workbook, ByVal sheetName As String) As Excel.Worksheet
            '工作表是否存在
            Dim ValidWorksheet As Excel.Worksheet = Nothing
            For Each sht As Excel.Worksheet In wkbk.Worksheets
                If String.Compare(sht.Name, sheetName) = 0 Then
                    ValidWorksheet = sht
                    Return ValidWorksheet
                End If
            Next
            '返回检测结果
            Return ValidWorksheet
        End Function

#End Region

#Region "  ---  几何绘图"

        ''' <summary>
        ''' 设置Chart与Applicatin的尺寸，Application的尺寸默认是随着Chart的尺寸而自动变化的。
        ''' </summary>
        ''' <param name="ChartSize_sugested"></param>
        ''' <param name="Chart">进行尺寸设置的Chart对象</param>
        ''' <param name="App">进行尺寸设置的Excel程序</param>
        ''' <param name="Fixed">是否要将Excel应用程序的窗口固定</param>
        ''' <remarks></remarks>
        Public Shared Sub SetLocation_Size(ByVal ChartSize_sugested As ChartSize, _
                                           Optional ByVal Chart As Excel.Chart = Nothing, _
                                           Optional ByVal App As Excel.Application = Nothing, _
                                           Optional ByVal Fixed As Boolean = True)
            With ChartSize_sugested
                If Chart IsNot Nothing Then
                    Try
                        With Chart.ChartArea
                            .Top = 0
                            .Left = 0
                            .Height = ChartSize_sugested.ChartHeight
                            .Width = ChartSize_sugested.ChartWidth
                        End With
                    Catch ex As Exception
                        MessageBox.Show("设置Chart尺寸失败！" & vbCrLf & ex.Message & vbCrLf & "报错位置：" & ex.TargetSite.Name, "Error", MessageBoxButtons.OK)
                    End Try
                End If
                If App IsNot Nothing Then
                    Try
                        '在可以在隐藏的情况下设置程序界面的尺寸，但是在设置之前，一定要确保其WindowState不能为xlMaximized(可以为xlMinimized)
                        If App.WindowState = Excel.XlWindowState.xlMaximized Then
                            App.WindowState = Excel.XlWindowState.xlNormal
                        End If
                        App.Height = .ChartHeight + .MarginOut_Height
                        App.Width = .ChartWidth + .MarginOut_Width
                        If Fixed Then
                            FixWindow(App.Hwnd)
                        End If
                    Catch ex As Exception
                        MessageBox.Show("设置Excel的窗口尺寸失败！" & vbCrLf & ex.Message & vbCrLf & "报错位置：" & ex.TargetSite.Name, "Error", MessageBoxButtons.OK)
                    End Try
                End If
            End With
        End Sub

        ''' <summary>
        ''' 将任意形状以指定的值定位在Chart的某一坐标轴中。
        ''' </summary>
        ''' <param name="ShapeToLocate">要进行定位的形状</param>
        ''' <param name="Ax">此形状将要定位的轴</param>
        ''' <param name="Value">此形状在Chart中所处的值</param>
        ''' <param name="percent">将形状按指定的百分比的宽度或者高度的部位定位到坐标轴的指定值的位置。
        ''' 如果其值设定为0，则表示此形状的左端（或上端）定位在设定的位置处，
        ''' 如果其值为100，则表示此形状的右端（或下端）定位在设置的位置处。</param>
        ''' <remarks></remarks>
        Public Overloads Shared Sub setPositionInChart(ByVal ShapeToLocate As Excel.Shape, ByVal Ax As Excel.Axis, _
                                      ByVal Value As Double, Optional ByVal percent As Double = 0)
            Dim cht As Excel.Chart = DirectCast(Ax.Parent, Excel.Chart)
            If cht IsNot Nothing Then
                'Try          '先考察形状是否是在Chart之中

                '    ShapeToLocate = cht.Shapes.Item(ShapeToLocate.Name)
                'Catch ex As Exception           '如果形状不在Chart中，则将形状复制进Chart，并将原形状删除
                '    ShapeToLocate.Copy()
                '    cht.Paste()
                '    ShapeToLocate.Delete()
                '    ShapeToLocate = cht.Shapes.Item(cht.Shapes.Count)
                'End Try
                '
                Select Case Ax.Type

                    Case Excel.XlAxisType.xlCategory                '横向X轴
                        Dim PositionInChartByValue As Double = GetPositionInChartByValue(Ax, Value)
                        With ShapeToLocate
                            .Left = PositionInChartByValue - percent * .Width
                        End With

                    Case Excel.XlAxisType.xlValue                   '竖向Y轴
                        Dim PositionInChartByValue As Double = GetPositionInChartByValue(Ax, Value)
                        With ShapeToLocate
                            .Top = PositionInChartByValue - percent * .Height
                        End With
                    Case Excel.XlAxisType.xlSeriesAxis
                        MessageBox.Show("暂时不知道这是什么坐标轴", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End Select
            End If
        End Sub

        ''' <summary>
        ''' 将一组形状以指定的值定位在Chart的某一坐标轴中。
        ''' </summary>
        ''' <param name="ShapesToLocate">要进行定位的形状</param>
        ''' <param name="Ax">此形状将要定位的轴</param>
        ''' <param name="Values">此形状在Chart中所处的值</param>
        ''' <param name="percents">将形状按指定的百分比的宽度或者高度的部位定位到坐标轴的指定值的位置。
        ''' 如果其值设定为0，则表示此形状的左端（或上端）定位在设定的位置处，
        ''' 如果其值为100，则表示此形状的右端（或下端）定位在设置的位置处。</param>
        ''' <remarks></remarks>
        Public Overloads Shared Sub setPositionInChart(ByVal Ax As Excel.Axis, ByVal ShapesToLocate As Excel.Shape(), _
                                      ByVal Values As Double(), Optional ByVal Percents As Double() = Nothing)
            ' ------------------------------------------------------
            '检查输入的数组中的元素个数是否相同
            Dim Count As UShort = ShapesToLocate.Length
            If Values.Length <> Count Then
                MessageBox.Show("输入数组中的元素个数不相同。", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            If Percents IsNot Nothing Then
                If Percents.Count <> 1 And Percents.Length <> Count Then
                    MessageBox.Show("输入数组中的元素个数不相同。", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If
            End If
            ' ------------------------------------------------------
            Dim cht As Excel.Chart = DirectCast(Ax.Parent, Excel.Chart)
            '
            Dim max As Double = Ax.MaximumScale
            Dim min As Double = Ax.MinimumScale
            '
            Dim PlotA As Excel.PlotArea = cht.PlotArea
            ' ------------------------------------------------------
            Dim shp As Excel.Shape
            Dim Value As Double
            Dim Percent As Double = Percents(0)
            Dim PositionInChartByValue As Double
            ' ------------------------------------------------------

            Select Case Ax.Type

                Case Excel.XlAxisType.xlCategory                '横向X轴


                Case Excel.XlAxisType.xlValue                   '竖向Y轴
                    If Ax.ReversePlotOrder = False Then         '顺序刻度值，说明Y轴数据为下边小上边大
                        For i As UShort = 0 To Count - 1
                            shp = ShapesToLocate(i)
                            Value = Values(i)
                            If Percents.Count > 1 Then Percent = Percents(i)
                            With PlotA
                                PositionInChartByValue = .InsideTop + .InsideHeight * (max - Value) / (max - min)
                            End With
                            With shp
                                .Top = PositionInChartByValue - Percent * .Width
                            End With
                        Next
                    Else                                         '逆序刻度值，说明Y轴数据为上边小下边大
                        For i As UShort = 0 To Count - 1
                            shp = ShapesToLocate(i)
                            Value = Values(i)
                            If Percents.Count > 1 Then Percent = Percents(i)
                            With PlotA
                                PositionInChartByValue = .InsideTop + .InsideHeight * (Value - min) / (max - min)
                            End With
                            With shp
                                .Top = PositionInChartByValue - Percent * .Width
                            End With
                        Next
                    End If

                Case Excel.XlAxisType.xlSeriesAxis
                    MessageBox.Show("暂时不知道这是什么坐标轴", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End Select

        End Sub

        ''' <summary>
        ''' 根据在坐标轴中的值，来返回这个值在Chart中的几何位置
        ''' </summary>
        ''' <param name="Ax"></param>
        ''' <param name="Value"></param>
        ''' <returns>如果Ax是一个水平X轴，则返回的是坐标轴Ax中的值Value在Chart中的Left值；
        ''' 如果Ax是一个竖向Y轴，则返回的是坐标轴Ax中的值Value在Chart中的Top值。</returns>
        ''' <remarks></remarks>
        Public Overloads Shared Function GetPositionInChartByValue(ByVal Ax As Excel.Axis, ByVal Value As Double) As Double

            Dim PositionInChartByValue As Double
            Dim cht As Excel.Chart = DirectCast(Ax.Parent, Excel.Chart)
            '
            Dim max As Double = Ax.MaximumScale
            Dim min As Double = Ax.MinimumScale
            '
            Dim PlotA As Excel.PlotArea = cht.PlotArea
            Select Case Ax.Type

                Case Excel.XlAxisType.xlCategory                '横向X轴
                    Dim PositionInPlot As Double
                    If Ax.ReversePlotOrder = False Then         '正向分类，说明X轴数据为左边小右边大
                        PositionInPlot = PlotA.InsideWidth * (Value - min) / (max - min)
                    Else                                        '逆序类别，说明X轴数据为左边大右边小
                        PositionInPlot = PlotA.InsideWidth * (max - Value) / (max - min)
                    End If
                    PositionInChartByValue = PlotA.InsideLeft + PositionInPlot

                Case Excel.XlAxisType.xlValue                   '竖向Y轴
                    Dim PositionInPlot As Double
                    If Ax.ReversePlotOrder = False Then         '顺序刻度值，说明Y轴数据为下边小上边大
                        PositionInPlot = PlotA.InsideHeight * (max - Value) / (max - min)
                    Else                                        '逆序刻度值，说明Y轴数据为上边小下边大
                        PositionInPlot = PlotA.InsideHeight * (Value - min) / (max - min)
                    End If
                    PositionInChartByValue = PlotA.InsideTop + PositionInPlot
                Case Excel.XlAxisType.xlSeriesAxis
                    Debug.Print("暂时不知道这是什么坐标轴")
            End Select
            Return PositionInChartByValue
        End Function

        ''' <summary>
        ''' 根据一组形状在某一坐标轴中的值，来返回这些值在Chart中的几何位置
        ''' </summary>
        ''' <param name="Ax"></param>
        ''' <param name="Values">要在坐标轴中进行定位的多个形状在此坐标轴中的数值</param>
        ''' <returns>如果Ax是一个水平X轴，则返回的是坐标轴Ax中的值Value在Chart中的Left值；
        ''' 如果Ax是一个竖向Y轴，则返回的是坐标轴Ax中的值Value在Chart中的Top值。</returns>
        ''' <remarks></remarks>
        Public Overloads Shared Function GetPositionInChartByValue(ByVal Ax As Excel.Axis, ByVal Values As Double()) As Double()
            Dim Count As UShort = Values.Length
            Dim PositionInChartByValue(0 To Count - 1) As Double
            ' --------------------------------------------------
            Dim cht As Excel.Chart = DirectCast(Ax.Parent, Excel.Chart)
            '
            Dim max As Double = Ax.MaximumScale
            Dim min As Double = Ax.MinimumScale
            Dim Value As Double
            '
            Dim PlotA As Excel.PlotArea = cht.PlotArea
            With PlotA
                Select Case Ax.Type

                    Case Excel.XlAxisType.xlCategory                '横向X轴
                        If Ax.ReversePlotOrder = False Then         '正向分类，说明X轴数据为左边小右边大
                            For i As UShort = 0 To Count - 1
                                Value = Values(i)
                                PositionInChartByValue(i) = PlotA.InsideLeft + PlotA.InsideWidth * (Value - min) / (max - min)
                            Next
                        Else                                        '逆序类别，说明X轴数据为左边大右边小
                            For i As UShort = 0 To Count - 1
                                Value = Values(i)
                                PositionInChartByValue(i) = PlotA.InsideLeft + PlotA.InsideWidth * (max - Value) / (max - min)
                            Next
                        End If

                    Case Excel.XlAxisType.xlValue                   '竖向Y轴
                        If Ax.ReversePlotOrder = False Then         '顺序刻度值，说明Y轴数据为下边小上边大
                            For i As UShort = 0 To Count - 1
                                Value = Values(i)
                                PositionInChartByValue(i) = PlotA.InsideTop + PlotA.InsideHeight * (max - Value) / (max - min)
                            Next
                        Else                                        '逆序刻度值，说明Y轴数据为上边小下边大
                            For i As UShort = 0 To Count - 1
                                Value = Values(i)
                                PositionInChartByValue(i) = PlotA.InsideTop + PlotA.InsideHeight * (Value - min) / (max - min)
                            Next
                        End If
                    Case Excel.XlAxisType.xlSeriesAxis
                        Debug.Print("暂时不知道这是什么坐标轴")
                End Select
            End With
            Return PositionInChartByValue
        End Function

        ''' <summary>
        ''' 设置Excel中的文本框格式：无边距、正中排列
        ''' </summary>
        ''' <param name="TextFrame">要进行格式设置的文本框对象</param>
        ''' <param name="TextSize">字体的大小</param>
        ''' <param name="VerticalAnchor">文本的竖向排列方式</param>
        ''' <param name="HorizontalAlignment">文本的水平排列方式</param>
        ''' <param name="Text">文本框中的文本</param>
        ''' <remarks></remarks>
        Public Shared Sub FormatTextbox_Tag(ByVal TextFrame As Excel.TextFrame2, Optional TextSize As Single = 8, _
                                            Optional ByVal VerticalAnchor As MsoVerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle, _
                                            Optional ByVal HorizontalAlignment As MsoParagraphAlignment = MsoParagraphAlignment.msoAlignCenter, _
                                            Optional ByVal Text As String = Nothing)
            With TextFrame
                '文本框的边距
                .MarginLeft = 0
                .MarginRight = 0
                .MarginTop = 0
                .MarginBottom = 0
                '文本的垂直对齐
                .VerticalAnchor = VerticalAnchor
                '.WordWrap = Microsoft.Office.Core.MsoTriState.msoCTrue
                '.AutoSize = Microsoft.Office.Core.MsoAutoSize.msoAutoSizeNone
                With .TextRange
                    .Font.Size = TextSize
                    .Font.Name = AMEApplication.FontName_TNR
                    .ParagraphFormat.Alignment = HorizontalAlignment
                    .Text = Text
                End With
            End With
        End Sub

#End Region

    End Class

#End Region

End Namespace