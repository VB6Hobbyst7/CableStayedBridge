Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop
Imports System.IO
Imports System.Threading
Imports System.Text.RegularExpressions
Public Class frmDeriveData_Excel
    Inherits frmDeriveData

#Region "  ---  types"

    ''' <summary>
    ''' 在用后台线程提取所有的工作表的数据时，进行传递的参数
    ''' </summary>
    ''' <remarks>此结构中包含了要进行数据提取的所有文档以及工作表和Range信息</remarks>
    Private Structure ExportToWorksheet
        ''' <summary>
        ''' 放置提取后的数据的工作簿
        ''' </summary>
        ''' <remarks></remarks>
        Public WorkBook_ExportedTo As Excel.Workbook
        ''' <summary>
        ''' 要进行提取的工作簿
        ''' </summary>
        ''' <remarks></remarks>
        Public arrWkbk() As String
        ''' <summary>
        ''' 每一个工作簿中要进行提取的工作表，并用来索引此工作表中的Range范围
        ''' </summary>
        ''' <remarks>集合中的Worksheet对象对应的是保存数据的工作簿中的工作表对象。</remarks>
        Public listRangeInfo As List(Of RangeInfoForExport)

        ''' <summary>
        ''' 是否要分析出提取数据的工作簿中的日期数据
        ''' </summary>
        ''' <remarks></remarks>
        Public ParseDateFromFilePath As Boolean

        ''' <summary>
        ''' 构造函数
        ''' </summary>
        ''' <param name="WorkBook_ExportedTo">放置提取后的数据的工作簿</param>
        ''' <param name="arrWkbk">要进行提取的工作簿</param>
        ''' <param name="listRangeInfo">每一个工作簿中要进行提取的工作表，并用来索引此工作表中的Range范围</param>
        ''' <param name="ParseDateFromFilePath">是否要分析出提取数据的工作簿中的日期数据</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal WorkBook_ExportedTo As Excel.Workbook, _
                       ByVal arrWkbk() As String, _
                       ByVal listRangeInfo As List(Of RangeInfoForExport), _
                       ByVal ParseDateFromFilePath As Boolean)
            Me.WorkBook_ExportedTo = WorkBook_ExportedTo
            Me.arrWkbk = arrWkbk
            Me.listRangeInfo = listRangeInfo
            Me.ParseDateFromFilePath = ParseDateFromFilePath
        End Sub

        ''' <summary>
        ''' 每一个工作簿中要提取的Range对象的信息
        ''' </summary>
        ''' <remarks></remarks>
        Public Structure RangeInfoForExport
            ''' <summary>
            ''' 要保存到的工作表对象，也是每一个数据工作簿中要进行检索的工作表对象
            ''' </summary>
            ''' <remarks></remarks>
            Public sheet As Worksheet
            ''' <summary>
            ''' 工作表中进行提取的数据范围
            ''' </summary>
            ''' <remarks></remarks>
            Public strRange As String
            ''' <summary>
            ''' 每一种要提取的数据范围的列数，即Range.Areas集合中每一个小Area中的Columns.Count之和
            ''' </summary>
            ''' <remarks></remarks>
            Public ColumnsCount As Integer

            ''' <summary>
            ''' 构造函数
            ''' </summary>
            ''' <param name="sheet">要保存到的工作表对象，也是每一个数据工作簿中要进行检索的工作表对象</param>
            ''' <param name="strRange">工作表中进行提取的数据范围</param>
            ''' <param name="ColumnsCount">每一种要提取的数据范围的列数，即Range.Columns.Count的值</param>
            ''' <remarks></remarks>
            Public Sub New(ByVal sheet As Worksheet, ByVal strRange As String, ByVal ColumnsCount As Integer)
                Me.sheet = sheet
                Me.strRange = strRange
                Me.ColumnsCount = ColumnsCount
            End Sub
        End Structure

    End Structure

#End Region

    Public Sub New()
        InitializeComponent()
        InitializeComponent_ActivateAtRuntime()
        ' Add any initialization after the InitializeComponent() call.
        Me.F_ChildType = ChildType.Excel
    End Sub

    ''' <summary>
    ''' 开始输出数据
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub StartExportData()


        ' ------------- 提取每一个工作表与Range范围的格式 -------------
        Dim listRangeInfo As New List(Of ExportToWorksheet.RangeInfoForExport)
        '
        Dim strTestRange As String = ""
        '
        '记录DataGridView控件中所有数据的数组
        Try
            Dim RowsCount As Integer = MyDataGridView1.Rows.Count
            For RowIndex As Integer = 0 To RowsCount - 2
                Dim RowObject As DataGridViewRow = MyDataGridView1.Rows.Item(RowIndex)

                '获取对应的Worksheet对象
                Dim strSheetName As String = RowObject.Cells.Item(0).Value.ToString
                Dim ExportedSheet As Excel.Worksheet = GetExactWorksheet(F_WorkBook_ExportedTo, listSheetNameInWkbk, strSheetName)

                '检查Range对象的格式是否正确()
                strTestRange = RowObject.Cells.Item(1).Value.ToString
                Dim testRange As Range = ExportedSheet.Range(strTestRange)      '这一步可能出错：Range的格式不规范
                '
                Dim columnsCount As Integer = 0
                For Each a As Range In testRange.Areas
                    '如果想引用相交区域（公共区域)，可以在多个区域间添加空格“ ”：  如Range("B1:B10 A4:D6 ").Select()  '选中多个单元格区域的交集
                    columnsCount += a.Columns.Count
                Next
                Dim RangeInfo As New ExportToWorksheet.RangeInfoForExport(ExportedSheet, strTestRange, columnsCount)
                listRangeInfo.Add(RangeInfo)
            Next
        Catch ex As Exception
            MessageBox.Show("定义区域范围的格式出错，出错的格式为 : " & vbCrLf _
                            & strTestRange & "，请重新输入", "Error", _
                             MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
            Exit Sub
        End Try



        ' ----------------------------------- 
        '是否要分析出提取数据的工作簿中的日期数据
        Dim blnParseDateFromFilePath As Boolean = False
        If Me.ChkboxParseDate.Checked Then blnParseDateFromFilePath = True
        '不允许再更改提取日期的正则表达式
        Me.ChkboxParseDate.Checked = False
        '开始提取数据
        Dim Export As New ExportToWorksheet(F_WorkBook_ExportedTo, arrDocPaths, listRangeInfo, blnParseDateFromFilePath)
        Me.BackgroundWorker1.RunWorkerAsync(Export)

    End Sub

    '在后台线程中执行操作
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork

        '定义初始变量
        Dim ExportToWorksheet As ExportToWorksheet = DirectCast(e.Argument, ExportToWorksheet)
        Dim arrWkbk As String() = ExportToWorksheet.arrWkbk
        Dim listRangeInfo As List(Of ExportToWorksheet.RangeInfoForExport) = ExportToWorksheet.listRangeInfo
        Dim WorkBook_ExportedTo As Excel.Workbook = ExportToWorksheet.WorkBook_ExportedTo
        Dim blnParseDateFromFilePath As Boolean = ExportToWorksheet.ParseDateFromFilePath

        '一共要处理的工作表数(工作簿个数*每个工作簿中提取的工作表数)，用来显示进度条的长度
        Dim Count_Workbooks As Integer = Me.ListBoxDocuments.Items.Count
        Dim Count_RangesInOneWkbk As Integer = listRangeInfo.Count
        Dim Count_AllRanges As Integer = Count_Workbooks * Count_RangesInOneWkbk

        '
        Dim percent As Integer = 0
        Dim unit As Single
        With Me.ProgressBar1
            unit = (.Maximum - .Minimum) / Count_AllRanges
        End With
        '报告进度
        Me.BackgroundWorker1.ReportProgress(percent, "")
        '开始提取数据
        Dim blnRangeFormatValidated As Boolean = False

        For iWkbk As Short = 0 To Count_Workbooks - 1
            Dim strWkbkPath As String = arrWkbk(iWkbk)
            Dim wkbk As Workbook = Nothing
            Try
                '下面有可能会出现工作簿打开出错
                wkbk = Me.F_ExcelApp.Workbooks.Open(strWkbkPath, UpdateLinks:=False, [ReadOnly]:=True)

                '获取工作簿中的所有工作表的名称，以备后面的与要进行提取的工作表名称的比较之用
                Dim arrExistedSheetsName(0 To wkbk.Worksheets.Count - 1) As String
                Dim i As Short = 0
                For Each sht As Worksheet In wkbk.Worksheets
                    arrExistedSheetsName(i) = sht.Name
                    i += 1
                Next

                '此工作簿中的每一个要提取的Range对象在列表中的行号。
                Dim iRow_Sheet_Range As Integer
                '此工作簿所对应的表头的数据：工作簿的名称或者是工作簿中包含的日期信息
                Dim ColumnTitle As String = GetColumnTitle(strWkbkPath, blnParseDateFromFilePath)
                '
                For iRow_Sheet_Range = 0 To Count_RangesInOneWkbk - 1
                    Dim sheetExportTo As Worksheet = listRangeInfo.Item(iRow_Sheet_Range).sheet
                    Dim strRange As String = ""
                    Try
                        '有可能会出现工作表提取出错：此工作表不存在
                        Dim sheetExtractFrom As Worksheet = GetContainedWorksheet(wkbk, arrExistedSheetsName, sheetExportTo.Name)
                        If sheetExtractFrom IsNot Nothing Then
                            '
                            strRange = listRangeInfo.Item(iRow_Sheet_Range).strRange

                            '----------------------------------------

                            '更新这一组数据所放置的列号,初始的放置数据的列号
                            Dim ColumnsCount As Integer = listRangeInfo.Item(iRow_Sheet_Range).ColumnsCount
                            Dim ColNumToBeAdded As Integer = cstColNum_FirstData + ColumnsCount * iWkbk

                            ' 提取数据
                            Call ExportData(sheetExtractFrom, sheetExportTo, strRange, ColNumToBeAdded, ColumnTitle)

                            '----------------------------------------
                        Else
                            Throw New NullReferenceException
                        End If
                    Catch ex As Exception
                        '工作表提取出错：此工作表不存在
                        Dim strError As String = "工作表：" & wkbk.FullName & " ： " _
                                                 & sheetExportTo.Name & " 无法找到。"
                        Debug.Print(strError)
                        Me.F_ErrorList.Add(strError)
                    Finally
                        Me.BackgroundWorker1.ReportProgress((iWkbk * Count_RangesInOneWkbk + iRow_Sheet_Range + 1) * unit, _
                                                            strWkbkPath & ":" & sheetExportTo.Name & ":" & strRange)
                    End Try
                Next
            Catch ex As Exception
                '工作簿打开出错
                Dim strError As String = "工作簿：" & wkbk.FullName & " 打开时出错。"
                Me.F_ErrorList.Add(strError)
            Finally
                If wkbk IsNot Nothing Then   '说明工作簿顺利打开
                    wkbk.Close(SaveChanges:=False)
                End If
                Me.BackgroundWorker1.ReportProgress((iWkbk + 1) * Count_RangesInOneWkbk * unit, strWkbkPath)
            End Try
        Next
    End Sub

#Region "  --- 数据提取"

    '匹配工作表
    ''' <summary>
    ''' 获取工作簿中的工作表对象
    ''' </summary>
    ''' <param name="wkbk">工作表所在的工作簿</param>
    ''' <param name="SheetName">工作表的名称</param>
    ''' <remarks>如果此工作表已经在工作簿中出现，则返回对应的工作表，否则，创建一个新的工作表，
    ''' 并将新工作表名称添加到已经存在的工作表名称列表中</remarks>
    Private Function GetExactWorksheet(ByVal wkbk As Workbook, ByVal ExistedSheetsName As List(Of String), ByVal SheetName As String) As Worksheet
        Dim blnSheetExisted As Boolean = False
        Dim sheet As Worksheet = Nothing
        For Each ExistedSheet As String In ExistedSheetsName
            '下面的比较一定要忽略大小写，因为Excel中大小写不同的工作表名称被认为是同一个工作表
            '如果新添加的工作表的名称与已经存在的工作表名称只是大小写不同，则会报错。
            If String.Compare(ExistedSheet, SheetName, True) = 0 Then
                sheet = wkbk.Worksheets(SheetName)
                '如果检索的工作表名称与已有的工作表名称只是大小写不同，则要将工作表名称设置为进行检索的工作表名称。
                sheet.Name = SheetName
                Return sheet
            End If
        Next
        If Not blnSheetExisted Then
            sheet = wkbk.Worksheets.Add()
            sheet.Name = SheetName
            '将新工作表名称添加到已经存在的工作表名称列表中，以供下次调用
            ExistedSheetsName.Add(SheetName)
        End If
        Return sheet
    End Function
    ''' <summary>
    ''' 获取工作簿中的工作表对象，如果此工作表已经在工作簿中出现，则返回对应的工作表，否则，返回Nothing。
    ''' </summary>
    ''' <param name="wkbk">工作表所在的工作簿</param>
    ''' <param name="ExistedSheetsName">工作簿中已经存在的工作表的名称的集合</param>
    ''' <param name="SheetName">工作表的名称</param>
    ''' <remarks>比较的依据：1、忽略大小写，2、要检索的工作表的名称的字符串是包含于已经存在的工作表名称的字符串的。</remarks>
    Private Function GetContainedWorksheet(ByVal wkbk As Workbook, ByVal ExistedSheetsName As String(), ByVal SheetName As String) As Worksheet
        Dim sheet As Worksheet = Nothing
        For Each ExistedSheet As String In ExistedSheetsName
            '忽略大小写，因为Excel中大小写不同的工作表名称被认为是同一个工作表 StringComparer.OrdinalIgnoreCase
            If ExistedSheet.IndexOf(SheetName, System.StringComparison.OrdinalIgnoreCase) >= 0 Then
                sheet = wkbk.Worksheets(ExistedSheet)
                Return sheet
            End If
        Next
        Return sheet
    End Function
    '
    ''' <summary>
    ''' 正式开始提取数据
    ''' </summary>
    ''' <param name="shtExtractFrom">要进行数据提取的工作表</param>
    ''' <param name="shtExportTo">放置提取的数据的工作表</param>
    ''' <param name="strRange">提取的数据区间</param>
    ''' <param name="ColNumToBeAdded">要放置的数据Range的第一个列号</param>
    ''' <param name="ColumnTitle">每一个Range数据的表头信息</param>
    ''' <remarks>在此方法中，可以引用多个不连续的区域，即在各区域间添加逗号“,”。</remarks>
    Private Sub ExportData(ByVal shtExtractFrom As Excel.Worksheet, _
                           ByVal shtExportTo As Excel.Worksheet, _
                           ByVal strRange As String, ByVal ColNumToBeAdded As Integer, ByVal ColumnTitle As String)
        '下一个要放置的数据Range的列号
        'Dim ColNumToBeAdded As Integer = ColNum_FirstData

        '添加表头数据
        shtExportTo.Cells(cstRowNum_ColumnTitle, ColNumToBeAdded).Value = ColumnTitle

        '要提取的数据的范围
        Dim rgOut As Range = shtExtractFrom.Range(strRange)
        For Each rg As Range In rgOut.Areas
            Dim ColsCount As Integer = rg.Columns.Count
            Dim RowsCount As Integer = rg.Rows.Count
            '要放置数据的区域范围
            Dim rgIn As Range
            With shtExportTo
                '这里将每一个Area的第一个单元格都移到指定的第一个数据单元格，其实是有一点问题的。如果在一个Worksheet中，引用了两个不连续的区域，
                '而这两个区域中的最顶部的单元格并不是在同一行，那么在导出到两列数据时，这两列数据也应该不是从同一行开始的。
                rgIn = .Range(.Cells(cstRowNum_FirstData, ColNumToBeAdded), _
                              .Cells(cstRowNum_FirstData + RowsCount - 1, ColNumToBeAdded + ColsCount - 1))
            End With
            '提取数据
            rgIn.Value = rg.Value
            ColNumToBeAdded += ColsCount
        Next
    End Sub
#End Region

End Class