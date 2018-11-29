Imports Microsoft.Office.Interop
Imports DAMIE.Miscellaneous
Imports DAMIE.All_Drawings_In_Application
Imports DAMIE.GlobalApp_Form

Public Class DiaFrm_PointsTreeView

#Region "  ---  Constants"

    ''' <summary>
    ''' 工作表“监测点编号与对应坐标”中，第一个测点所在的行号
    ''' </summary>
    ''' <remarks></remarks>
    Const cstRowNum_FirstPoint As Byte = 2
    ''' <summary>
    ''' 在TreeView中，表示监测点的数据在控件中的深度（即对应TreeNode的level属性）（第一级深度为0）
    ''' </summary>
    ''' <remarks></remarks>
    Const cstDepth_PointInTreeView As Byte = 2
#End Region

#Region "  ---  Fields"

    ''' <summary>
    ''' 存放监测点位信息的Excel表格
    ''' </summary>
    ''' <remarks></remarks>
    Private F_wkshtPoints As Excel.Worksheet


    ''' <summary>
    ''' 指示绘制监测点位的窗口Form是否已经加载
    ''' </summary>
    ''' <remarks></remarks>
    Private blnHasLoaded As Boolean = False

    ''' <summary>
    ''' Visio平面图中，与监测点位相关的信息（不是开挖平面），用来在Visio平面图中绘制测点。
    ''' </summary>
    ''' <remarks></remarks>
    Private F_MonitorPointsInfo As ClsDrawing_PlanView.MonitorPointsInformation

    ''' <summary>
    ''' 监测点位所要绘制的Visio图形对象
    ''' </summary>
    ''' <remarks></remarks>
    Private F_PaintingPage As ClsDrawing_PlanView

    ''' <summary>
    ''' CAD坐标系与Visio坐标系进行线性转换的斜率与截距.kx、cx、ky、cy
    ''' </summary>
    ''' <remarks>其基本公式为：x_Visio=Kx*x_CAD+Cx；y_Visio=Ky*y_CAD+Cy</remarks>
    Private ConversionParameter As Cdnt_Cvsion

    ''' <summary>
    ''' 在TreeView中选择的监测点的完整路径，索引对应的TreeNode对象
    ''' </summary>
    ''' <remarks>监测点只包括最后一级要进行绘图的测点，而不包括其父节点</remarks>
    Private F_dicChosenPoints_TreeView As New Dictionary(Of String, TreeNode)
    ''' <summary>
    ''' 选择的监测点，索引的item为传递到列表框中的文本
    ''' </summary>
    ''' <remarks></remarks>
    Private F_dicListedPoints As New Dictionary(Of String, TreeNode)

    ''' <summary>
    ''' 此visio图形中的所有监测点的形状的字典集合,
    ''' Dictionary(Of 监测点在列表框中显示的文本,监测点的形状ID)
    ''' </summary>
    ''' <remarks></remarks>
    Private F_dicVisioPoints As New Dictionary(Of String, Short)

#End Region

#Region "  ---  structures"

    ''' <summary>
    ''' TreeView中的父项目的文本与其在大数组中对应的行号区间
    ''' </summary>
    ''' <remarks></remarks>
    Structure RowsToDictionary
        ''' <summary>
        ''' 监测项目或者基坑ID的数据在大数组中对应的行号的区域。
        ''' 数组中有两个元素，第一个表示此项目中的第一个元素的行号，第二个表示此项目中的最后一个元素的行号。
        ''' </summary>
        ''' <remarks></remarks>
        Public RowsSpan() As Integer
        ''' <summary>
        ''' 一个字典对象
        ''' </summary>
        ''' <remarks></remarks>
        Public dic As Object
    End Structure


    ''' <summary>
    ''' Visio中的监测点的形状的基本信息
    ''' </summary>
    ''' <remarks></remarks>
    Structure PointInVisio
        ''' <summary>
        ''' 监测项目的名称，如“测斜”
        ''' </summary>
        ''' <remarks></remarks>
        Public strItem As String
        ''' <summary>
        ''' 监测点的编号，如“CX1”
        ''' </summary>
        ''' <remarks></remarks>
        Public strPoint As String
        ''' <summary>
        ''' 监测点在Visio中的坐标值(x,y)
        ''' </summary>
        ''' <remarks></remarks>
        Public Coordinates() As Double
    End Structure
#End Region

    ''' <summary>
    ''' 构造函数
    ''' </summary>
    ''' <param name="PaintingPage">监测点位所要绘制的Visio图形对象</param>
    ''' <param name="wkshtPoints">存放监测点位信息的Excel表格</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal wkshtPoints As Excel.Worksheet, _
                   ByVal PaintingPage As ClsDrawing_PlanView)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Me.F_PaintingPage = PaintingPage
        Me.F_MonitorPointsInfo = PaintingPage.MonitorPointsInfo
        Me.F_wkshtPoints = wkshtPoints
    End Sub

    ''' <summary>
    ''' 窗口加载
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>先提取坐标变换的斜率及截距；再从监测点编号与其对应的坐标创建树形列表</remarks>
    Private Sub FrmTreeView_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'blnHasLoaded是为了解决showdialog在每次加载窗口时都触发一次load事件的问题。
        If Not blnHasLoaded Then
            Try
                ' ------------------ 提取坐标变换的斜率及截距
                Dim CP1 As Visio.Shape
                Dim CP2 As Visio.Shape
                With Me.F_PaintingPage.Page.Shapes
                    CP1 = .ItemFromID(F_MonitorPointsInfo.pt_Visio_BottomLeft_ShapeID)
                    CP2 = .ItemFromID(F_MonitorPointsInfo.pt_Visio_UpRight_ShapeID)
                End With
                'visio中的坐标以inch为单位
                Dim xp1 As Double
                Dim yp1 As Double
                Dim xp2 As Double
                Dim yp2 As Double
                'x,y,xp,yp都是以inch为单位的
                Dim x As Double, y As Double
                With CP1
                    x = .Cells("LocPinX").Result(Visio.VisUnitCodes.visInches)
                    y = .Cells("LocPinY").Result(Visio.VisUnitCodes.visInches)
                    .XYToPage(x, y, xp1, yp1)
                End With
                With CP2
                    x = .Cells("LocPinX").Result(Visio.VisUnitCodes.visInches)
                    y = .Cells("LocPinY").Result(Visio.VisUnitCodes.visInches)
                    .XYToPage(x, y, xp2, yp2)
                End With
                With Me.F_MonitorPointsInfo
                    ConversionParameter = Coordinate_Conversion(.pt_CAD_BottomLeft.X, .pt_CAD_BottomLeft.Y, _
                                                                .pt_CAD_UpRight.X, .pt_CAD_UpRight.Y, _
                                                                xp1, yp1, xp2, yp2)
                End With

                '从监测点编号与其对应的坐标创建树形列表
                Call constructTreeView(F_wkshtPoints, TreeViewPoints)
                '标记：此窗口已经在主程序中加载过，后面就不用再加载了。因为不会对其进行close，只会将其隐藏
                blnHasLoaded = True
            Catch ex As Exception

                MessageBox.Show("不能正确地提取监测点位信息。" & vbCrLf & ex.Message & vbCrLf & "报错位置：" & ex.TargetSite.Name, _
                                "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                'Exit Sub
                'Me.Close()
                Me.Dispose()
            End Try
        End If
    End Sub

    ''' <summary>
    ''' 从监测点编号与其对应的坐标创建树形列表
    ''' </summary>
    ''' <param name="wksht">保存数据的工作表</param>
    ''' <param name="myTreeView">用于列举监测点编号的TreeView对象</param>
    ''' <remarks></remarks>
    Private Sub constructTreeView(ByVal wksht As Excel.Worksheet, myTreeView As TreeView)

        '提取数据工作表中所有测点的数据,确保其中至少有一行数据
        Dim arrData As Object
        With wksht.UsedRange
            ' ************ 看情况进行排序 ************
            ' 
            ' ***************************************
            If .Rows.Count - cstRowNum_FirstPoint < 0 Then
                MessageBox.Show("No data rows found in the specified worksheet" & vbCrLf & "please check the DataBase file or search another worksheet", _
                                "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
                Exit Sub
            Else
                arrData = .Rows(cstRowNum_FirstPoint & ":" & .Rows.Count).value
            End If
        End With
        '大数组的数据行的上、下标值
        Dim lb As Byte, ub As Integer
        lb = LBound(arrData, 1)
        ub = UBound(arrData, 1)
        '
        '-*****************************************************-
        ' ------ DicItem ------------ 第一列：监测项目
        Dim DicItems As New Dictionary(Of String, RowsToDictionary)

        '              cx1:(x,y)
        '          ID1
        '              cx2:(x,y)
        '   Item1
        '              cx1:(x,y)
        '          ID2
        '              cx2:(x,y)
        'DicItems
        '              cx1:(x,y)
        '          ID1
        '              cx2:(x,y)
        '   Item2
        '              cx1:(x,y)
        '          ID2
        '              cx2:(x,y)

        Dim rtd As RowsToDictionary
        If ub = lb Then
            rtd = New RowsToDictionary
            rtd.RowsSpan = {lb, lb}
            DicItems.Add(arrData(lb, lb), rtd)
        Else
            '
            Dim rowOldItem As Integer = lb
            For rowNewItem As Integer = lb To ub - 1
                If String.Compare(arrData(rowNewItem, lb), arrData(rowNewItem + 1, lb), True) <> 0 Then
                    rtd = New RowsToDictionary
                    rtd.RowsSpan = {rowOldItem, rowNewItem}
                    DicItems.Add(arrData(rowNewItem, lb), rtd)
                    'QueItem.Enqueue({arrData(rowNewItem, lb), rowOldItem, rowNewItem, Nothing})
                    rowOldItem = rowNewItem + 1
                End If
            Next   '下一个监测项目
            '处理最后一个监测项目，它可以与倒数第二个不一样，这时要单独拿出来处理，因为它并不包含在上面的循环中。
            If String.Compare(arrData(ub, lb), arrData(ub - 1, lb), True) = 0 Then
                rtd = New RowsToDictionary
                rtd.RowsSpan = {rowOldItem, ub}
                DicItems.Add(arrData(ub, lb), rtd)
            Else
                rtd = New RowsToDictionary
                rtd.RowsSpan = {rowOldItem, ub - 1}
                DicItems.Add(arrData(rowOldItem, lb), rtd)
                '
                rtd = New RowsToDictionary
                rtd.RowsSpan = {ub, ub}
                DicItems.Add(arrData(rowOldItem, lb), rtd)
            End If
        End If

        '-*****************************************************-
        ' ------ DicIDs ---------- 第二列：基坑ID

        Dim startrow As Integer
        Dim endrow As Integer
        For i_item As Integer = 0 To DicItems.Count - 1
            Dim struct_value As RowsToDictionary = DicItems.Values(i_item)
            startrow = struct_value.RowsSpan(0)
            endrow = struct_value.RowsSpan(1)
            '
            Dim DicIDs As New Dictionary(Of String, RowsToDictionary)
            '
            If startrow = endrow Then       '如果此监测项目下只有一行数据
                rtd = New RowsToDictionary
                rtd.RowsSpan = {startrow, startrow}
                DicIDs.Add(arrData(startrow, lb + 1), rtd)
            Else
                '
                '
                Dim rowOldID As Integer = startrow
                For rowNewID As Integer = startrow To endrow - 1
                    If String.Compare(arrData(rowNewID, lb + 1), arrData(rowNewID + 1, lb + 1), True) <> 0 Then
                        rtd = New RowsToDictionary
                        rtd.RowsSpan = {rowOldID, rowNewID}
                        DicIDs.Add(arrData(rowNewID, lb + 1), rtd)

                        rowOldID = rowNewID + 1
                    End If
                Next    '下一个基坑ID

                '处理最后一个基坑，它可以与倒数第二个不一样，这时要单独拿出来处理，因为它并不包含在上面的循环中。
                If String.Compare(arrData(endrow, lb + 1), arrData(endrow - 1, lb + 1), True) = 0 Then
                    rtd = New RowsToDictionary
                    rtd.RowsSpan = {rowOldID, endrow}
                    DicIDs.Add(arrData(rowOldID, lb + 1), rtd)

                Else
                    rtd = New RowsToDictionary
                    rtd.RowsSpan = {rowOldID, endrow - 1}
                    DicIDs.Add(arrData(rowOldID, lb + 1), rtd)
                    '
                    rtd = New RowsToDictionary
                    rtd.RowsSpan = {endrow, endrow}
                    DicIDs.Add(arrData(endrow, lb + 1), rtd)
                    '

                End If
            End If
            struct_value.dic = DicIDs
            DicItems.Item(DicItems.Keys(i_item)) = struct_value
        Next    '下一个监测项目

        '-*****************************************************-
        ' ------ TreeViewPoints ---------- 第三列：测点编号

        TreeViewPoints.Nodes.Clear()

        For i As Integer = 0 To DicItems.Keys.Count - 1

            Dim strItem As String = DicItems.Keys(i)
            Dim itemNode As New TreeNode(strItem)
            TreeViewPoints.Nodes.Add(itemNode)

            'dicid：每一个监测项目下的基坑ID的集合
            Dim dicID As Dictionary(Of String, RowsToDictionary) = DicItems.Values(i).dic

            '---------
            For j As Integer = 0 To dicID.Keys.Count - 1

                'For Each struct_ID As RowsToDictionary In dicID.Values
                Dim strID As String = dicID.Keys(j)
                Dim IDNode As New TreeNode(strID)
                itemNode.Nodes.Add(IDNode)


                'dicPoints：每一个基坑中的测点的集合
                Dim dicPoints As New Dictionary(Of String, Double())

                '坐标点赋值
                Dim rowspan(0 To 1) As Integer
                rowspan = dicID.Values(j).RowsSpan
                For row As Integer = rowspan(0) To rowspan(1)

                    '获取测点编号
                    Dim strPoint As String = arrData(row, lb + 2)
                    '获取测点对应的坐标，并进行坐标转换，将工作表中的CAD坐标系中的坐标转换成Visio中的坐标系（以inch为单位）
                    Dim Coodinates(0 To 1) As Double
                    With ConversionParameter
                        Coodinates = {arrData(row, lb + 3) * .kx + .cx, arrData(row, lb + 4) * .ky + .cy}
                    End With
                    '
                    dicPoints.Add(arrData(row, lb + 2), Coodinates)
                    Dim NdPoint As TreeNode = IDNode.Nodes.Add(strPoint)
                    '将测点编号链接对应的坐标值
                    NdPoint.Tag = Coodinates

                Next    '下一个测点
                Dim struct As RowsToDictionary = dicID.Values(j)
                struct.dic = dicPoints
                dicID.Item(dicID.Keys(j)) = struct
                ' --- Note ----
                'dicID.Item(dicID.Keys(j)).dic = dicPoints
                '上面的语句失败，因为它仅为 dicID.Item(dicID.Keys(j)) 属性返回的 RowsToDictionary 结构提供了临时的分配。 
                '这是一个值类型的结构，该语句运行后不保留临时结构。 
                '解决该问题的方法是：为 RowsToDictionary 结构的属性声明一个变量，并使用这个变量，从而为 RowsToDictionary 结构创建更为永久的分配。 
                ' \--- Note ----
            Next        '下一个基坑
        Next            '下一个监测项目
        '
    End Sub

#Region "  ---  控件操作"

    ''' <summary>
    ''' 在TreeView的Checkbox框中进行选择或者取消选择时的操作
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub TreeViewPoints_AfterCheck(sender As Object, e As TreeViewEventArgs) Handles TreeViewPoints.AfterCheck
        Dim nd As TreeNode = e.Node
        Dim blnChecked As Boolean = nd.Checked
        If nd.Level = cstDepth_PointInTreeView Then
            Try
                If blnChecked Then
                    F_dicChosenPoints_TreeView.Add(nd.FullPath, nd)
                Else
                    F_dicChosenPoints_TreeView.Remove(nd.FullPath)
                End If
            Catch ex As Exception
            End Try
        End If
        '将节点的子节点也全部选中或全部取消选择
        For Each childNode As TreeNode In nd.Nodes
            childNode.Checked = blnChecked
        Next
    End Sub

    ''' <summary>
    ''' 将TreeView中选择的测点添加进列表中
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BtnAdd_Click(sender As Object, e As EventArgs) Handles BtnAdd.Click

        Dim lstItem As New List(Of LstbxDisplayAndItem)
        Dim keys As String() = F_dicChosenPoints_TreeView.Keys.ToArray
        Dim values As TreeNode() = F_dicChosenPoints_TreeView.Values.ToArray
        '
        ListBoxChosenItems.DataSource = F_dicChosenPoints_TreeView.Keys.ToArray
        '重新构造字典dicListedPoints

        F_dicListedPoints.Clear()
        For i As Integer = 0 To F_dicChosenPoints_TreeView.Count - 1
            With F_dicChosenPoints_TreeView
                F_dicListedPoints.Add(.Keys(i), .Values(i))
            End With
        Next
    End Sub

    ''' <summary>
    ''' 移除选定的测点
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BtnRemove_Click(sender As Object, e As EventArgs) Handles BtnRemove.Click
        With ListBoxChosenItems
            Dim indices = .SelectedIndices
            '
            Dim strData As String() = .DataSource
            Dim lst As New List(Of String)
            For Each s As String In strData
                lst.Add(s)
            Next
            '
            For i = indices.Count - 1 To 0 Step -1
                Dim index As Integer = indices(i)
                lst.RemoveAt(index)
                With F_dicListedPoints
                    .Remove(.Keys(index))
                End With
            Next
            .DataSource = lst.ToArray
        End With
    End Sub

    ''' <summary>
    ''' 清空最后的测点
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BtnClear_Click(sender As Object, e As EventArgs) Handles BtnClear.Click
        ListBoxChosenItems.DataSource = Nothing
        F_dicListedPoints.Clear()
    End Sub

    ''' <summary>
    ''' CollapseAll
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BtnColpsAll_Click(sender As Object, e As EventArgs) Handles BtnColpsAll.Click
        TreeViewPoints.CollapseAll()
    End Sub

    ''' <summary>
    ''' 清空treeview中所有节点的选择
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>这里将代码限死为3层节点，这种方法有待改进</remarks>
    Private Sub ClearAllCheckedNode(sender As Object, e As EventArgs) Handles btnClearCheckedNode.Click
        RemoveHandler TreeViewPoints.AfterCheck, AddressOf TreeViewPoints_AfterCheck
        F_dicChosenPoints_TreeView.Clear()
        '取消treeview中所有节点的选择
        For Each nd0 As TreeNode In TreeViewPoints.Nodes
            nd0.Checked = False
            For Each nd1 As TreeNode In nd0.Nodes
                nd1.Checked = False
                For Each nd2 As TreeNode In nd1.Nodes
                    nd2.Checked = False
                Next
            Next
        Next
        AddHandler TreeViewPoints.AfterCheck, AddressOf TreeViewPoints_AfterCheck
    End Sub

#End Region

    ''' <summary>
    ''' 在点击“确定”时，是否还需要再执行一次绘制测点的操作，
    ''' 因为有可能从前面进行预览之后，所要绘制的点位就没有发生改变，所以此时点击确认就只需要直接将窗口关闭就可以了。
    ''' </summary>
    ''' <remarks></remarks>
    Private blnRefreshed As Boolean = False
    Private Sub btnPreview_Click(sender As Object, e As EventArgs) Handles btnPreview.Click
        Call DrawPoints()
    End Sub

    ''' <summary>
    ''' 点击确定时将窗口隐藏
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>如果已经进行过预览，则直接隐藏窗口，如果没有进行过预览，则先进行一次预览，再隐藏窗口</remarks>
    Private Sub BtnOk_Click(sender As Object, e As EventArgs) Handles BtnOk.Click
        If Not blnRefreshed Then
            Call DrawPoints()
        End If
        Me.Hide()
    End Sub
    ''' <summary>
    ''' 点击关闭窗体时将窗体隐藏
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>对于模式窗体（以showdialog显示），在点击关闭时，其实就是执行hide操作</remarks>
    Private Sub FrmPointsTreeView_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Me.Hide()
        e.Cancel = True
    End Sub

    ''' <summary>
    ''' ！进行监测点位的绘制
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub DrawPoints()
        Dim vso As ClsDrawing_PlanView = GlobalApplication.Application.PlanView_VisioWindow
        If vso IsNot Nothing Then
            Dim pg As Visio.Page = vso.Page
            Try
                Dim vsoWindow As Visio.Window = pg.Application.ActiveWindow
                vsoWindow.Page = pg
                vsoWindow.Application.ShowChanges = False
                '
                Dim arrListedPoints As String() = F_dicListedPoints.Keys.ToArray
                Dim arrExistedPoints As String() = F_dicVisioPoints.Keys.ToArray
                '
                Dim arrAddOrRemove(0 To 1) As Object
                arrAddOrRemove = GetPointsToBeProcessed(arrExistedPoints, arrListedPoints)
                Dim arrToBeAdded As String() = arrAddOrRemove(0)
                Dim arrToBeRemoved As String() = arrAddOrRemove(1)


                ' ----- arrPointsToBeAdded ------------- 处理要进行添加的图形
                Dim arrPointsToBeAdded(0 To arrToBeAdded.Length - 1) As PointInVisio
                Dim i As Integer = 0
                Dim strSeparator As String = TreeViewPoints.PathSeparator
                For Each tag_Add As String In arrToBeAdded
                    Dim nd As TreeNode = F_dicListedPoints.Item(tag_Add)
                    Dim struct_Point As New PointInVisio
                    With struct_Point
                        Dim str = tag_Add
                        .strItem = str.Substring(0, str.IndexOf(strSeparator))
                        .strPoint = nd.Text
                        .Coordinates = nd.Tag
                    End With
                    arrPointsToBeAdded(i) = struct_Point
                    i += 1
                Next

                Call AddMonitorPoints(vsoWindow, arrToBeAdded, arrPointsToBeAdded)

                ' ----- arrToBeRemoved ------------- 处理要进行删除的图形
                For Each strPointTag As String In arrToBeRemoved
                    Dim shp As Visio.Shape
                    shp = pg.Shapes.ItemFromID(F_dicVisioPoints.Item(strPointTag))
                    shp.Delete()
                    '
                    F_dicVisioPoints.Remove(strPointTag)
                Next

                ' ----------------------- 
                Me.blnRefreshed = True
                vsoWindow.Application.ShowChanges = True
            Catch ex As Exception
                Debug.Print(ex.Message)
                MessageBox.Show("出错！" & vbCrLf & ex.Message & vbCrLf & "报错位置：" & ex.TargetSite.Name, _
                                "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        Else
            MessageBox.Show("Visio绘图已经关闭，请重新打开。", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub

    ''' <summary>
    ''' 如果列表框的数据源发生改变，则在点击确定时必须要重新绘制测点
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ListBoxChosenItems_DataSourceChanged(sender As Object, e As EventArgs) Handles ListBoxChosenItems.DataSourceChanged
        Me.blnRefreshed = False
    End Sub

#Region "  ---  本地子方法"

    ''' <summary>
    ''' 根据Visio中与列表框中的测点字符的差异，得到分别要进行添加和删除的测点
    ''' </summary>
    ''' <param name="arrExistedPoints">在Visio中存在的测点的字符数据</param>
    ''' <param name="arrListedPoints">在ListBox列表中存在的测点的字符数据</param>
    ''' <returns>返回一个数组，数组中有两个元素，元素各自又都为一个数组，第一个为要在Visio进行添加的测点的字符的数组；
    ''' 第二个为要从Visio进行删除的测点图形的字符的数组</returns>
    ''' <remarks>这里的“字符数据”都是指测点在TreeView中的完整路径的字符</remarks>
    Private Function GetPointsToBeProcessed(ByVal arrExistedPoints As String(), ByVal arrListedPoints As String()) As Object()
        Dim arrAddOrRemove(0 To 1) As Object

        '将两个数组中的元素进行汇总组合并排序，且去掉其中中重复项
        Dim ssMixedPoints As New SortedSet(Of String)(StringComparer.OrdinalIgnoreCase)
        For Each ep In arrExistedPoints
            ssMixedPoints.Add(ep)
        Next
        For Each lp In arrListedPoints
            ssMixedPoints.Add(lp)
        Next
        '
        Dim listToBeAdded As New List(Of String)
        Dim listToBeRemoved As New List(Of String)
        '
        Dim PointsCount As Short = ssMixedPoints.Count
        Dim sbIn_E As SByte
        Dim sbIn_L As SByte
        Dim cmp As SByte
        For i As Integer = 0 To PointsCount - 1
            If arrExistedPoints.Contains(ssMixedPoints(i)) Then sbIn_E = 1 Else sbIn_E = 0
            If arrListedPoints.Contains(ssMixedPoints(i)) Then sbIn_L = 1 Else sbIn_L = 0
            '对上面两个数组中分别的对应元素进行比较——相减
            cmp = sbIn_E - sbIn_L
            Select Case cmp
                Case 1      '说明Visio中存在而在列表中没有列出，那么在处理时要从visio中将此图形删除
                    listToBeRemoved.Add(ssMixedPoints(i))
                Case -1     '说明Visio中没有，但是列表中有，那么要在Visio中添加进此图形
                    listToBeAdded.Add(ssMixedPoints(i))
            End Select
        Next
        Return {listToBeAdded.ToArray, listToBeRemoved.ToArray}
    End Function

    ''' <summary>
    ''' 根据指定的监测点在visio图形中绘出对应的测点图形
    ''' </summary>
    ''' <param name="vsoWindow">进行绘图的visio图形</param>
    ''' <param name="strTags">要绘制的所有监测点的在列表中对应的文本</param>
    ''' <param name="PointsToBeAdded">要绘制的所有监测点的集合</param>
    ''' <remarks></remarks>
    Public Sub AddMonitorPoints(ByVal vsoWindow As Visio.Window, ByVal strTags As String(), ByVal PointsToBeAdded As PointInVisio())
        Dim n As Integer = PointsToBeAdded.Count
        If n > 0 Then

            '
            Dim arrMaster(0 To n - 1) As Object
            Dim arrCoordinate(0 To 2 * n - 1) As Double
            Dim arrIDOut(0 To n - 1) As Short
            Dim arrPointTag(0 To n - 1) As String
            '

            Dim DocumentMasters As Visio.Masters = vsoWindow.Document.Masters

            Dim i As Integer = 0
            For Each pt In PointsToBeAdded
                Try
                    arrMaster(i) = DocumentMasters.Item(pt.strItem)
                Catch ex As Exception
                    Debug.Print("测点类型： {0} 在Visio模板文件中并不存在，此时用""OtherTypes""来表示。", pt.strItem)
                    arrMaster(i) = DocumentMasters.Item("OtherTypes")
                End Try
                arrPointTag(i) = pt.strPoint
                arrCoordinate(2 * i) = pt.Coordinates(0)
                arrCoordinate(2 * i + 1) = pt.Coordinates(1)
                i += 1
            Next

            Dim pg As Visio.Page = vsoWindow.Page
            Dim ProcessedCount As Integer = pg.DropMany(arrMaster, arrCoordinate, arrIDOut)

            '设置每一个测点形状上显示的文字
            pg.Application.ShowChanges = True
            Dim tag As String = F_MonitorPointsInfo.ShapeName_MonitorPointTag
            Try

                For i_p As Integer = 0 To n - 1
                    Dim id As Integer = arrIDOut(i_p)
                    '设置主控形状的实例对象中，“某形状”的文本，“某形状”的索引等规范化。
                    pg.Shapes.ItemFromID(id).Shapes.Item(tag).Text = arrPointTag(i_p)
                    F_dicVisioPoints.Add(strTags(i_p), arrIDOut(i_p))
                Next
            Catch ex As Exception
                MessageBox.Show("主控形状中没有找到子形状：" & "tag" & vbCrLf & ex.Message & vbCrLf & "报错位置：" & ex.TargetSite.Name, _
"Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

#End Region



End Class