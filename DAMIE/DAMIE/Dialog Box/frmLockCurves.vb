Imports DAMIE.GlobalApp_Form
Imports DAMIE.All_Drawings_In_Application
Imports System.Threading
Imports DAMIE.Miscellaneous

''' <summary>
''' 对于监测曲线进行批量操作：锁定或者删除
''' </summary>
''' <remarks></remarks>
Public Class frmLockCurves

#Region "   ---   窗口的加载与关闭"
    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.

        With Me.MyDataGridView1

            '从存储文件中提取数据
            Dim mySetting As New mySettings_Application
            Dim Date_Handle() As Object
            Date_Handle = mySetting.Curve_BatchProcessing
            If Date_Handle IsNot Nothing Then
                ' 将数据刷新到界面的列表中
                Dim rowscount As Integer = Date_Handle.Length
                .Rows.Add(rowscount)        '必须要先添加行，然后才能在后面进行赋值，否则会出现索引不在集合内的报错。
                '注意：第一行数据的行标为0，列表头所在行的下标为-1。
                For rowNum As Integer = 0 To rowscount - 1
                    .Item(0, rowNum).Value = Date_Handle(rowNum)(0)
                    .Item(1, rowNum).Value = Date_Handle(rowNum)(1)
                Next
            End If
        End With
    End Sub

    ''' <summary>
    ''' 清空列表中的所有数据
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        Me.MyDataGridView1.Rows.Clear()
    End Sub

    ''' <summary>
    ''' 关闭窗口
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btn_Cancel_Click(sender As Object, e As EventArgs) Handles btn_Cancel.Click
        Me.Close()
    End Sub


#End Region

#Region "   ---   数据验证"

    ''' <summary>
    ''' 在DataGridView中，添加新行时，将其搜索方向设置为“锁定”。
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub MyDataGridView1_RowsAdded(sender As Object, e As DataGridViewRowsAddedEventArgs) Handles MyDataGridView1.RowsAdded
        With Me.MyDataGridView1
            If e.RowIndex >= 1 Then
                Dim handle = .Item(1, e.RowIndex - 1)
                If handle.Value Is Nothing Then
                    handle.Value = "锁定"
                End If
            End If
        End With
    End Sub

    ''' <summary>
    ''' 校验单元格中的日期数据
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub MyDataGridView1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles MyDataGridView1.CellValueChanged
        If e.RowIndex >= 0 And e.ColumnIndex = 0 Then
            Dim cell As DataGridViewCell = Me.MyDataGridView1.Item(e.ColumnIndex, e.RowIndex)
            Dim v As String = cell.Value
            Dim dt As Date
            Try     '将单元格中的内容转换为日期
                dt = CType(v, Date)
                cell.Value = dt.ToShortDateString
            Catch ex1 As Exception
                Dim l As Short = v.Length
                Try
                    Select Case l
                        Case 4
                            dt = New Date(Today.Year, CInt(v.Substring(2, 2)), CInt(v.Substring(4, 2)))
                            cell.Value = dt.ToShortDateString
                        Case 6
                            dt = New Date(CInt(v.Substring(0, 2)) + 2000, CInt(v.Substring(2, 2)), CInt(v.Substring(4, 2)))
                            cell.Value = dt.ToShortDateString
                        Case 8
                            dt = New Date(CInt(v.Substring(0, 4)), CInt(v.Substring(4, 2)), CInt(v.Substring(6, 2)))
                            cell.Value = dt.ToShortDateString
                        Case Else
                            cell.Value = ""
                    End Select
                Catch ex As Exception
                    cell.Value = ""
                End Try
            End Try
        End If
    End Sub


#End Region

#Region "   ---   执行操作"
    ''' <summary>
    ''' 执行
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub btn_OK_Click(sender As Object, e As EventArgs) Handles btn_OK.Click
        Dim a = Me.MyDataGridView1.DataSource

        ' 提取frmRolling中已经选择的要进行滚动的监测曲线图对象
        Dim RollingForm As frmRolling = APPLICATION_MAINFORM.MainForm.Form_Rolling
        Dim RollingMnts As List(Of clsDrawing_Mnt_RollingBase) = RollingForm.F_SelectedDrawings.RollingMnt
        With Me.MyDataGridView1
            ' 提到表格中的数据
            Dim rowscount As Integer = .Rows.Count
            Dim Date_Handle(0 To rowscount - 2) As Object
            For rowNum As Integer = 0 To rowscount - 2
                Date_Handle(rowNum) = {.Item(0, rowNum).Value, .Item(1, rowNum).Value}
            Next
            '将结果保存到文件中
            Dim mySetting As New mySettings_Application
            With mySetting
                .Curve_BatchProcessing = Date_Handle   '将结果保存到文件中
                .Save()
            End With
            '
            '为每一个Excel图形创建一个线程
            For Each RollingMnt As clsDrawing_Mnt_RollingBase In RollingMnts
                Dim thd As Thread = New Thread(AddressOf LockOrDelete)
                thd.Name = RollingMnt.Chart_App_Title
                thd.Start({RollingMnt, Date_Handle})
            Next RollingMnt
        End With
    End Sub

    ''' <summary>
    ''' 在工作线程中进行曲线的锁定或者删除
    ''' </summary>
    ''' <param name="Arg"></param>
    ''' <remarks></remarks>
    Private Sub LockOrDelete(Arg As Object)
        Dim RollingMnt As clsDrawing_Mnt_RollingBase = Arg(0)
        Dim Date_Handle As Object() = Arg(1)
        For row As Integer = LBound(Date_Handle) To UBound(Date_Handle)
            Dim RollingDate As Date = Date_Handle(row)(0)
            Dim handle As String = Date_Handle(row)(1)
            If handle = "锁定" Then
                Call RollingAndLock(RollingMnt, RollingDate)
            Else
                Call DeleteCurve(RollingMnt, RollingDate)
            End If
        Next row
    End Sub

    ''' <summary>
    ''' 图形滚动，并锁定曲线
    ''' </summary>
    ''' <param name="RollingMnt">要进行滚动的曲线所在的监测图</param>
    ''' <param name="RollingDate">要进行滚动的的曲线所代表的日期。</param>
    ''' <remarks></remarks>
    Private Sub RollingAndLock(RollingMnt As clsDrawing_Mnt_RollingBase, RollingDate As Date)
        '首先，要绘制的日期必须在此测点的监测数据的日期跨度范围之内。
        If RollingMnt.DateSpan.Contains(RollingDate) Then
            ' 检查指定的日期是否早就已经绘制在了图表中
            Dim a = RollingMnt.Dic_SeriesIndex_Tag
            Dim blnMatched As Boolean = False   '如果图形中已经有了这条曲线，就不添加或者锁定了。
            Dim SeriesIndex As Integer
            Dim Indices = a.Keys
            For Each SeriesIndex In Indices
                If Date.Compare(a.Item(SeriesIndex).ConstructionDate, RollingDate) = 0 Then
                    blnMatched = True
                    Exit For
                End If
            Next
            ' 如果指定的日期与图表中的第一条曲线的日期相匹配，则默认为不进行绘制
            If Not blnMatched Then
                RollingMnt.Rolling(RollingDate)
                '进行滚动的曲线总是seriesCollection中的第一条曲线，其下标值为1。
                RollingMnt.CopySeries(1)
            End If
        End If
    End Sub

    ''' <summary>
    ''' 删除图形中的曲线，如果图形中的没有对应日期的曲线，就不删除
    ''' </summary>
    ''' <param name="RollingMnt">要进行滚动的曲线所在的监测图</param>
    ''' <param name="RollingDate">要进行滚动的的曲线所代表的日期。</param>
    ''' <remarks></remarks>
    Private Sub DeleteCurve(RollingMnt As clsDrawing_Mnt_RollingBase, RollingDate As Date)
        Dim a = RollingMnt.Dic_SeriesIndex_Tag
        Dim CurveCount As Short = a.Count
        Dim blnMatched As Boolean = False   '只删除图形中已经绘制的曲线
        Dim SeriesIndex As Integer
        ' 检查图形中的曲线所对应的日期是否有与要进行删除的日期相匹配的。
        Dim Indices = a.Keys
        For Each SeriesIndex In Indices
            If Date.Compare(a.Item(SeriesIndex).ConstructionDate, RollingDate) = 0 Then
                If SeriesIndex <> 1 Then
                    '删除曲线，其中第一条曲线是用来进行滚动的，不能删除，其下标值为1。
                    RollingMnt.DeleteSeries(SeriesIndex)
                    Exit For
                End If
            End If
        Next
    End Sub
#End Region

End Class