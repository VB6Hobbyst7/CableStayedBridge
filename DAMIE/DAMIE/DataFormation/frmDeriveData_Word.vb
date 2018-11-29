Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Interop
Imports System.IO
Imports System.Threading
Imports System.Text.RegularExpressions

Public Class frmDeriveData_Word
    Inherits frmDeriveData

#Region "  ---  types"

    ''' <summary>
    ''' 每一个文档中要提取的测点及数据的位置信息
    ''' </summary>
    ''' <remarks></remarks>
    Private Class PointsInfoForExport

        ''' <summary>
        ''' 文档中要进行提取的测点标签
        ''' </summary>
        ''' <remarks></remarks>
        Public PointTag As String

        ''' <summary>
        ''' 进行搜索的方向：按行或者按列，即下一个搜索单元格是按行还是按列的方向前进。
        ''' </summary>
        ''' <remarks></remarks>
        Public SearchOrder As Excel.XlSearchOrder

        ''' <summary>
        ''' 在Word文档中，测点所对应的数据距离测点单元格的水平偏移位置。
        ''' 如数据单元格是在测点标签单元格的左边且紧靠标签单元格，则Offset的值为+1。
        ''' </summary>
        ''' <remarks></remarks>
        Public Offset As Byte


        ''' <summary>
        ''' 要在excel最终保存数据的工作表中写入数据的列号，即此列前面的行都已经被写入或者是预留的空行。
        ''' </summary>
        ''' <remarks></remarks>
        Public ColNumToBeWritten As Integer
        ''' <summary>
        ''' 要在excel最终保存数据的工作表中写入数据的行号，即此行上面的行都已经被写入或者是预留的空行。
        ''' </summary>
        ''' <remarks></remarks>
        Public RowNumToBeWritten As Integer

        ''' <summary>
        ''' 每一组提取数据所占据的列数。从数据提取上来看，此字段并没有什么作用，因为一般情况下，它的值都应该是2。
        ''' 但是从表格的设计上来看，它的值可以用来腾出空的列以放置其他数据。
        ''' </summary>
        ''' <remarks></remarks>
        Public ColumnsCountToBeAdd As Byte

        ''' <summary>
        ''' 构造函数
        ''' </summary>
        ''' <param name="PointTag">文档中要进行提取的测点标签</param>
        ''' <param name="Offset">测点所对应的数据距离测点单元格的水平偏移位置。
        ''' 如数据单元格是在测点标签单元格的左边且紧靠标签单元格，则Offset的值为+1。</param>
        ''' <param name="SearchOrder">进行搜索的方向：按行或者按列，即下一个搜索单元格是按行还是按列的方向前进。</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal PointTag As String, ByVal Offset As Integer, ByVal SearchOrder As Excel.XlSearchOrder)
            Me.SearchOrder = SearchOrder
            Me.PointTag = PointTag
            Me.Offset = Offset
            '
            Me.ColNumToBeWritten = cstColNum_FirstData
            Me.RowNumToBeWritten = cstRowNum_FirstData
            Me.ColumnsCountToBeAdd = cstColumnsCountToBeAdded
        End Sub
    End Class

    ''' <summary>
    ''' 在用后台线程提取所有的工作表的数据时，进行传递的参数
    ''' </summary>
    ''' <remarks>此结构中包含了要进行数据提取的所有文档，
    ''' 以及每个文档中进行提取的测点和对应数据的位置标签信息。</remarks>
    Private Structure ExportToWorksheet

        ''' <summary>
        ''' 放置提取后的数据的工作簿
        ''' </summary>
        ''' <remarks></remarks>
        Public WorkBook_ExportedTo As Excel.Workbook

        ''' <summary>
        ''' 要进行提取的Word文档
        ''' </summary>
        ''' <remarks></remarks>
        Public arrDocsPath() As String

        ''' <summary>
        ''' 是否要分析出提取数据的工作簿中的日期数据
        ''' </summary>
        ''' <remarks></remarks>
        Public ParseDateFromFilePath As Boolean

        ''' <summary>
        ''' 用来暂时保存数据的Excel工作表对象。在提取每一个文档的数据时，
        ''' 先将文档中的表格复制到Excel中的此暂存工作表中，然后对于此工作表中的内容进行搜索。
        ''' </summary>
        ''' <remarks></remarks>
        Public BufferSheet As Excel.Worksheet

        ''' <summary>
        ''' 构造函数
        ''' </summary>
        ''' <param name="WorkBook_ExportedTo">放置提取后的数据的工作簿</param>
        ''' <param name="arrDocsPath">要进行提取的所有word文档的绝对路径</param>
        ''' <param name="ParseDateFromFilePath">是否要分析出提取数据的工作簿中的日期数据</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal WorkBook_ExportedTo As Excel.Workbook, _
                       ByVal arrDocsPath() As String, _
                       ByVal ParseDateFromFilePath As Boolean, _
                       ByVal BufferSheet As Excel.Worksheet)
            Me.WorkBook_ExportedTo = WorkBook_ExportedTo
            Me.arrDocsPath = arrDocsPath
            Me.ParseDateFromFilePath = ParseDateFromFilePath
            Me.BufferSheet = BufferSheet
        End Sub

    End Structure

#End Region

#Region "  ---  Constants"

    ''' <summary>
    ''' 每一组提取数据所占据的列数。从数据提取上来看，此字段并没有什么作用，因为一般情况下，它的值都应该是2。
    ''' 但是从表格的设计上来看，它的值可以用来腾出空的列以放置其他数据。
    ''' </summary>
    ''' <remarks></remarks>
    Const cstColumnsCountToBeAdded As Byte = 2

#End Region

#Region "  ---  Fields"

    ''' <summary>
    ''' 从Word文档中提取数据
    ''' </summary>
    ''' <remarks></remarks>
    Private WithEvents F_WordApp As Word.Application
   
    ''' <summary>
    ''' 用来暂时保存数据的Excel工作表对象。在提取每一个文档的数据时，
    ''' 先将文档中的表格复制到Excel中的此暂存工作表中，然后对于此工作表中的内容进行搜索。
    ''' </summary>
    ''' <remarks></remarks>
    Private F_BufferSheet As Excel.Worksheet

    ''' <summary>
    ''' 每一个文档中要进行提取的测点标签，和与之对应的数据的相对偏移位置。
    ''' </summary>
    ''' <remarks>集合中的Worksheet对象对应的是保存数据的工作簿中的工作表对象。</remarks>
    Private F_DicPointsInfo As Dictionary(Of Worksheet, PointsInfoForExport)

#End Region

    Public Sub New()
        InitializeComponent()
        InitializeComponent_ActivateAtRuntime()
        ' Add any initialization after the InitializeComponent() call.
        AddHandler MyDataGridView1.RowsAdded, AddressOf MyDataGridView1_RowsAdded
        Me.F_ChildType = ChildType.Word
    End Sub

#Region "  ---  窗体的加载与关闭"

    ''' <summary>
    ''' 在DataGridView中，添加新行时，将其搜索方向设置为“按行”。
    ''' </summary>
    Protected Sub MyDataGridView1_RowsAdded(sender As Object, e As DataGridViewRowsAddedEventArgs) 'Handles MyDataGridView1.RowsAdded
        With Me.MyDataGridView1
            If e.RowIndex >= 1 Then
                Dim a = .Item(2, e.RowIndex - 1)
                If a.Value Is Nothing Then
                    a.Value = "按行"
                End If
            End If
        End With
    End Sub

    ''' <summary>
    ''' 在窗口关闭前，关闭进行数据处理的Excel与Word程序
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub frmDeriveDataFromWord_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        '关闭Word程序
        If Me.F_WordApp IsNot Nothing Then
            With Me.F_WordApp
                For Each doc As Word.Document In .Documents
                    doc.Close()
                Next
                .Quit()
                Me.F_WordApp = Nothing
            End With
        End If
    End Sub

    ''' <summary>
    ''' 逻辑值，指示此时是否正在进行数据的提取操作。
    ''' 这是为了应对在程序数据提取时引发的word文档关闭与用户手动关闭Word文档时的区别对待。
    ''' </summary>
    ''' <remarks></remarks>
    Private blnIsBeingExtracting As Boolean = False
    Private Sub F_ExcelApp_WorkbookBeforeClose(Wb As Workbook, ByRef Cancel As Boolean) Handles F_ExcelApp.WorkbookBeforeClose
        If Not blnIsBeingExtracting Then
            Wb.Application.Quit()
            Me.F_ExcelApp = Nothing
        End If
    End Sub
    Private Sub F_WordApp_DocumentBeforeClose(Doc As Document, ByRef Cancel As Boolean) Handles F_WordApp.DocumentBeforeClose
        If Not blnIsBeingExtracting Then
            Doc.Application.Quit()
            Me.F_WordApp = Nothing
        End If
    End Sub
#End Region

#Region "  --- 数据提取"

    ''' <summary>
    ''' 开始输出数据
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub StartExportData()
        Me.blnIsBeingExtracting = True
        '打开进行数据提取的Word程序
        If Me.F_WordApp Is Nothing Then
            Me.F_WordApp = New Word.Application
            Me.F_WordApp.Visible = False
        End If

        Dim listPointsTag As New List(Of String)

        Me.F_BufferSheet = Me.F_WorkBook_ExportedTo.Worksheets.Add
        '


        ' ------------- 提取每一个工作表与Range范围的格式 -------------并返回DataGridView中的所有数据
        Me.F_DicPointsInfo = SearchPointsInfo(Me.F_WorkBook_ExportedTo)
        If F_DicPointsInfo Is Nothing Then Exit Sub

        ' ----------------------------------- 
        '是否要分析出提取数据的工作簿中的日期数据
        Dim blnParseDateFromFilePath As Boolean = False
        If Me.ChkboxParseDate.Checked Then blnParseDateFromFilePath = True
        '不允许再更改提取日期的正则表达式
        Me.ChkboxParseDate.Checked = False
        ' ---------------------- 开始提取数据 ---------------------
        Dim Export As New ExportToWorksheet(F_WorkBook_ExportedTo, arrDocPaths, _
                                            blnParseDateFromFilePath, Me.F_BufferSheet)
        Me.BackgroundWorker1.RunWorkerAsync(Export)

    End Sub

    Private Function SearchPointsInfo(ByVal wkbk As Excel.Workbook) As Dictionary(Of Worksheet, PointsInfoForExport)
        Dim listRangeInfo As New Dictionary(Of Worksheet, PointsInfoForExport)
        Dim strTestRange As String = ""
        '
        '记录DataGridView控件中所有数据的数组
        Try
            Dim RowsCount As Integer = Me.MyDataGridView1.Rows.Count
            For RowIndex As Integer = 0 To RowsCount - 2
                Dim RowObject As DataGridViewRow = MyDataGridView1.Rows.Item(RowIndex)

                '要进行提取的测点标签
                Dim strPointName As String = RowObject.Cells.Item(0).Value.ToString
                '设置与测点标签对应的excel工作表对象，并为其命名
                Dim sht As Worksheet = Nothing
                Try
                    sht = wkbk.Worksheets.Item(strPointName)
                Catch ex1 As Exception
                    '表示工作簿中没有这一工作表
                    sht = wkbk.Worksheets.Add
                    '为新创建的工作表命名
                    Dim blnNameOk As Boolean = False
                    Dim shtName = strPointName
                    Do
                        Try
                            sht.Name = shtName
                            blnNameOk = True
                        Catch ex2 As Exception
                            '表示此名称已经在工作簿中被使用了
                            shtName = shtName & "1"
                        End Try
                    Loop Until blnNameOk
                End Try

                '测点数据距离测点标签的偏移位置
                Dim Offset As Byte = RowObject.Cells.Item(1).Value.ToString
                '搜索的方向：按行或者是按列
                Dim SearchDirection As Excel.XlSearchOrder
                Dim comboBox As DataGridViewComboBoxCell = RowObject.Cells.Item(2)
                Select Case comboBox.Value
                    Case "按行"
                        SearchDirection = XlSearchOrder.xlByRows
                    Case "按列"
                        SearchDirection = XlSearchOrder.xlByColumns
                    Case Else
                        MessageBox.Show("请先输入搜索方向", "Error", _
                  MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return Nothing
                End Select

                Dim RangeInfo As New PointsInfoForExport(strPointName, Offset, SearchDirection)
                listRangeInfo.Add(sht, RangeInfo)
            Next
        Catch ex As Exception
            MessageBox.Show("数据的格式出错 : " & vbCrLf _
                            & strTestRange & "，请重新输入", "Error", _
                             MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return Nothing
        End Try
        Return listRangeInfo
    End Function

    '在后台线程中执行操作
    ''' <summary>
    ''' 在后台线程中执行操作
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub StartToDoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        '定义初始变量
        Dim ExportToWorksheet As ExportToWorksheet = DirectCast(e.Argument, ExportToWorksheet)
        Dim arrDocsPath As String() = ExportToWorksheet.arrDocsPath
        Dim WorkBook_ExportedTo As Excel.Workbook = ExportToWorksheet.WorkBook_ExportedTo
        Dim blnParseDateFromFilePath As Boolean = ExportToWorksheet.ParseDateFromFilePath
        Dim bufferSheet As Excel.Worksheet = ExportToWorksheet.BufferSheet

        '一共要处理的工作表数(工作簿个数*每个工作簿中提取的工作表数)，用来显示进度条的长度
        Dim Count_Documents As Integer = Me.ListBoxDocuments.Items.Count
        '
        Dim percent As Integer = 0
        '每一份数据所对应的进度条长度
        Dim unit As Single
        With Me.ProgressBar1
            unit = (.Maximum - .Minimum) / Count_Documents
        End With
        '报告进度
        Me.BackgroundWorker1.ReportProgress(percent, "")
        '开始提取数据
        For iDoc As Short = 0 To Count_Documents - 1
            Dim strDocPath As String = arrDocsPath(iDoc)
            Dim Doc As Word.Document = Nothing
            Try
                '下面有可能会出现文档打开出错
                Doc = Me.F_WordApp.Documents.Open(FileName:=strDocPath, _
                                                [ReadOnly]:=True, Visible:=False)
                '
                Dim myTable As Word.Table
                Dim CountTables As Short = Doc.Tables.Count
                If CountTables > 0 Then
                    For iTable As Short = 1 To CountTables
                        myTable = Doc.Tables.Item(iTable)
                        ' ------------- 正式开始提取数据 -------------

                        Call ExportData(DataTableInWord:=myTable)

                        ' ------------- 正式开始提取数据 -------------

                        Me.BackgroundWorker1.ReportProgress(((iDoc + iTable / CountTables)) * unit, _
                                                       "正在提取文档：" & strDocPath)
                    Next        '文档中的下一个表格Table对象
                End If
            Catch ex As Exception
                '文档打开出错
                Dim strError As String = "Document文档：" & Doc.FullName & " 打开时出错。  " & vbCrLf & ex.Message
                Me.F_ErrorList.Add(strError)
            Finally
                If Doc IsNot Nothing Then   '说明工作簿顺利打开
                    Doc.Close(SaveChanges:=False)
                End If
                Me.BackgroundWorker1.ReportProgress((iDoc + 1) * unit, "正在提取文档：" & strDocPath)
            End Try

            '更新下一个文档的数据在对应的Excel工作表中所保存的列号
            '以及表头信息
            For iSheet As Short = 0 To F_DicPointsInfo.Count - 1
                Dim sht As Excel.Worksheet = F_DicPointsInfo.Keys(iSheet)
                Dim pointinfo As PointsInfoForExport = Me.F_DicPointsInfo.Values(iSheet)
                With pointinfo
                    '此工作簿所对应的表头的数据：工作簿的名称或者是工作簿中包含的日期信息
                    Dim ColumnTitle As String = GetColumnTitle(strDocPath, blnParseDateFromFilePath)
                    sht.Cells(cstRowNum_ColumnTitle, .ColNumToBeWritten).Value = ColumnTitle
                    '
                    .ColNumToBeWritten += .ColumnsCountToBeAdd
                    .RowNumToBeWritten = cstRowNum_FirstData
                End With
            Next
        Next  'Next Document下一个文档
    End Sub
  
    ''' <summary>
    '''  !!! 正式开始提取数据
    ''' </summary>
    ''' <param name="DataTableInWord">进行数据提取的word中的表格Table对象</param>
    ''' <remarks>提取的基本思路：已有一个doc对象，并对其中的一个测点进行提取。</remarks>
    Private Sub ExportData(ByVal DataTableInWord As Word.Table)
        Try
            Dim rgTable As Word.Range = DataTableInWord.Range
            rgTable.Copy()
            Me.F_BufferSheet.UsedRange.Clear()
            With Me.F_BufferSheet
                .Activate()
                .UsedRange.Clear()
                .Cells(1, 1).select()
                .Paste()
            End With


            '此文档中的每一个要提取的测点。
            For Each sheetExportTo As Excel.Worksheet In Me.F_DicPointsInfo.Keys
                Dim PointInfo As PointsInfoForExport = Me.F_DicPointsInfo.Item(sheetExportTo)

                ' ------------ 从暂存工作表中将测点标签与对应的数据提取到目标工作表中 ----------
                '搜索得到的第一个结果的range对象，如果没有搜索到，则返回nothing。
                Dim SearchedPoint As Excel.Range
                Dim ValueCell As Excel.Range
                With Me.F_BufferSheet.UsedRange
                    SearchedPoint = .Find(What:=PointInfo.PointTag,
                                                       SearchOrder:=PointInfo.SearchOrder,
                                                       LookAt:=XlLookAt.xlPart,
                                                       LookIn:=XlFindLookIn.xlValues,
                                                       SearchDirection:=XlSearchDirection.xlNext,
                                                       MatchCase:=False)
                    If SearchedPoint IsNot Nothing Then
                        '当搜索到指定查找区域的末尾时，此方法将绕回到区域的开始位置继续搜索。
                        '发生绕回后，要停止搜索，可保存第一个找到的单元格地址，然后测试后面找到的每个单元格地址是否与其相同。
                        Dim firstAddress As String = SearchedPoint.Address
                        '提取数据并写入最终的工作表
                        Do
                            With PointInfo
                                sheetExportTo.Cells(.RowNumToBeWritten, .ColNumToBeWritten).Value = SearchedPoint.Value

                                ' 监测值所对应的单元格
                                If PointInfo.SearchOrder = XlSearchOrder.xlByRows Then
                                    ValueCell = SearchedPoint.Offset(PointInfo.Offset, 0)
                                Else
                                    ValueCell = SearchedPoint.Offset(0, PointInfo.Offset)
                                End If

                                sheetExportTo.Cells(.RowNumToBeWritten, .ColNumToBeWritten + 1).Value = ValueCell.Value
                                .RowNumToBeWritten += 1
                            End With
                            SearchedPoint = .FindNext(SearchedPoint)
                        Loop While SearchedPoint IsNot Nothing And String.Compare(SearchedPoint.Address, firstAddress) <> 0

                    End If
                End With
            Next
        Catch ex As Exception
            '数据提取出错
            Dim strError As String = ""
            Me.F_ErrorList.Add(strError)
        Finally

        End Try
    End Sub

    '操作完成
    ''' <summary>
    ''' 操作完成，关闭Excel程序，写入异常信息，并控件进度条的显示
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
                 '删除用来缓存数据的中间工作表
        Me.F_BufferSheet.Delete()
        Me.F_BufferSheet = Nothing
        Me.blnIsBeingExtracting = False
    End Sub

#End Region

End Class