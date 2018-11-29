Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop
Imports System.IO
Imports System.Threading
Imports System.Text.RegularExpressions
Public Class frmDeriveDataFromExcel

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
            ''' 每一种要提取的数据范围的列数，即Range.Columns.Count的值
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

#Region "  ---  Constants"

    ''' <summary>
    ''' 记录异常信息的文本的名称
    ''' </summary>
    ''' <remarks>其文件夹路径与输出数据的Excel工作簿的路径相同</remarks>
    Private Const cstErrorInfoFileName As String = "ErrorInfo.txt"

    ''' <summary>
    ''' 每一列数据的表头信息所在的行，一般为第一行，一般为数据对应的日期
    ''' </summary>
    ''' <remarks></remarks>
    Const cstRowNum_ColumnTitle As Byte = 1

    ''' <summary>
    ''' 提取的数据中的第一行在工作表中所要放置的行号，一般为第3行。第一行一般用来放数据对应的日期，第二行一般为预留行。
    ''' </summary>
    ''' <remarks></remarks>
    Const cstRowNum_FirstData As Byte = 3

    ''' <summary>
    ''' 提取的数据中的第一列在工作表中所要放置的列号，一般为第2列。第1列用来放数据的说明
    ''' </summary>
    ''' <remarks></remarks>
    Const cstColNum_FirstData As Byte = 2

#End Region

#Region "  ---  Fields"

    ''' <summary>
    ''' 用于操作数据的Excel程序
    ''' </summary>
    ''' <remarks></remarks>
    Private F_ExcelApp As Excel.Application

    ''' <summary>
    ''' 保存提取后的数据的工作簿
    ''' </summary>
    ''' <remarks></remarks>
    Private F_WorkBook_ExportedTo As Excel.Workbook
    '
    ''' <summary>
    ''' 搜索日期的正则表达式字符串
    ''' </summary>
    ''' <remarks></remarks>
    Private F_regexPattern As String
    ''' <summary>
    ''' 利用正则表达式提取的字符中，{文件序号，年，月，日}分别在Match.Groups集合中的下标值。用值0来代表没有此项。
    ''' </summary>
    ''' <remarks>Match.Groups(0)返回的是Match结果本身，并不属于要提取的数据。</remarks>
    Private F_regexComponents(0 To 3) As Byte
    '
    ''' <summary>
    ''' 异常信息的集合
    ''' </summary>
    ''' <remarks></remarks>
    Private F_ErrorList As List(Of String)

    ''' <summary>
    ''' 定时触发器
    ''' </summary>
    ''' <remarks></remarks>
    Private WithEvents ProgressTimer As System.Windows.Forms.Timer

#End Region

#Region "  ---  窗体的加载与关闭"

    ''' <summary>
    ''' 窗体的加载
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub DeriveDataFromExcel_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '
        With Me.BackgroundWorker1
            .WorkerReportsProgress = True
            .WorkerSupportsCancellation = True
        End With
        '
        Me.txtbxSavePath.Text = Path.Combine( _
            System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory, Environment.SpecialFolderOption.None), _
            "数据提取.xlsx")
    End Sub

    ''' <summary>
    ''' 鼠标移出控件时隐藏
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub frmDeriveDataFromWord_MouseMove(sender As Object, e As MouseEventArgs) Handles Me.MouseMove
        If Not Me.AddFileOrDirectoryFiles1.Bounds.Contains(e.X, e.Y) Then
            Me.AddFileOrDirectoryFiles1.HideLabel()
        End If
    End Sub

    ''' <summary>
    ''' 点击ESC时关闭窗口
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub frmDeriveDataFromExcel_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub

    ''' <summary>
    ''' 在窗口关闭前，关闭进行数据处理的Excel与Word程序
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub frmDeriveDataFromWord_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        '关闭Excel程序
        If Me.F_ExcelApp IsNot Nothing Then
            With Me.F_ExcelApp
                For Each wkbk As Excel.Workbook In .Workbooks
                    wkbk.Close(SaveChanges:=False)
                Next
                .Quit()
                Me.F_ExcelApp = Nothing
            End With
        End If
    End Sub

#End Region


#Region "   ---   界面操作"
    ''' <summary>
    ''' 是否要提取文件名中的日期
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ChkboxParseDate_CheckedChanged(sender As Object, e As EventArgs) Handles ChkboxParseDate.CheckedChanged
        If ChkboxParseDate.Checked = True Then
            btn_DateFormat.Enabled = True
            Txtbox_DateFormat.Enabled = True
        Else
            btn_DateFormat.Enabled = False
            Txtbox_DateFormat.Enabled = False
        End If
    End Sub
    ''' <summary>
    ''' 构造提取日期的正则表达式
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btn_DateFormat_Click(sender As Object, e As EventArgs) Handles btn_DateFormat.Click
        Dim f As New frmRegexDate
        f.ShowDialog(Me.F_regexPattern, Me.F_regexComponents)
        Me.Txtbox_DateFormat.Text = Me.F_regexPattern
    End Sub
    ''' <summary>
    ''' 刷新提取日期的正则表达式
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Txtbox_DateFormat_TextChanged(sender As Object, e As EventArgs) Handles Txtbox_DateFormat.TextChanged
        Me.F_regexPattern = Txtbox_DateFormat.Text
    End Sub

    ' 拖拽操作
    Private Sub APPLICATION_MAINFORM_DragDrop(sender As Object, e As DragEventArgs) Handles ListBoxWorksheets.DragDrop
        Dim FilePaths As String() = e.Data.GetData(DataFormats.FileDrop)
        ' DoSomething with the Files or Directories that are droped in.
        Dim arrExcelFilePath As New List(Of String)
        For Each filepath As String In FilePaths
            Dim ext As String = Path.GetExtension(filepath)
            Dim str As String = ".xlsx.xls.xlsb"
            Dim blnExtensionMatched As Boolean = str.Contains(ext)
            If blnExtensionMatched Then
                Me.ListBoxWorksheets.Items.Add(filepath)
            End If
        Next
    End Sub
    Private Sub APPLICATION_MAINFORM_DragEnter(sender As Object, e As DragEventArgs) Handles ListBoxWorksheets.DragEnter
        ' See if the data includes text.
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            ' There is text. Allow copy.
            e.Effect = DragDropEffects.Copy
        Else
            ' There is no text. Prohibit drop.
            e.Effect = DragDropEffects.None
        End If

    End Sub

#End Region

#Region "  ---  获取文件或文件夹路径"

    '添加单个文件
    ''' <summary>
    ''' 以选择文件的形式向列表中添加文件
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub AddFile(sender As Object, e As EventArgs) Handles AddFileOrDirectoryFiles1.AddFile
        Dim FilePaths As String()
        With Me.OpenFileDialog1
            .Title = "选择要进行数据提取的Excel文件"
            .Filter = "Excel文件(*.xlsx, *.xls, *.xlsb)|*.xlsx;*.xls;*.xlsb"
            .FilterIndex = 2
            .Multiselect = True
            If .ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                FilePaths = .FileNames
            Else
                Exit Sub
            End If
        End With
        If FilePaths.Length > 0 Then
            Me.ListBoxWorksheets.Items.AddRange(FilePaths)
        End If
    End Sub

    '添加文件夹中的所有文件
    ''' <summary>
    ''' 以选择文件夹的形式向列表中添加文件
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub lbAddDir_Click(sender As Object, e As EventArgs) Handles AddFileOrDirectoryFiles1.AddFilesFromDirectory
        Dim strPath As String = ""
        With Me.FolderBrowserDialog1
            .ShowNewFolderButton = True
            .Description = "添加文件夹中的全部Excel文件"
            If .ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                strPath = .SelectedPath
            Else
                Exit Sub
            End If
        End With
        If strPath.Length > 0 Then
            Dim files As String() = Directory.GetFiles(strPath)
            For Each strFile As String In files
                Dim ext As String = Path.GetExtension(strFile)
                If String.Compare(ext, ".xls", True) = 0 OrElse _
                    String.Compare(ext, ".xlsx", True) = 0 OrElse _
                    String.Compare(ext, ".xlsb", True) = 0 Then
                    Me.ListBoxWorksheets.Items.Add(strFile)
                End If
            Next
        End If
    End Sub

    '保存数据的工作簿路径
    Private Sub BtnChoosePath_Click(sender As Object, e As EventArgs) Handles BtnChoosePath.Click
        Dim FilePath As String = ""
        With Me.SaveFileDialog1
            .Title = "选择将数据保存到的Excel工作簿路径"
            .Filter = "Excel文件(*.xlsx, *.xls, *.xlsb)|*.xlsx;*.xls;*.xlsb"
            .CreatePrompt = False
            .OverwritePrompt = True
            .AddExtension = True
            .DefaultExt = ".xlsx"
            .FilterIndex = 2
            If .ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                FilePath = .FileName
            Else
                Exit Sub
            End If
        End With
        If FilePath.Length > 0 Then
            Me.txtbxSavePath.Text = FilePath
        End If
    End Sub

    '从列表框中移除选择的工作簿
    ''' <summary>
    ''' 从列表框中移除选择的工作簿
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnRemove_Click(sender As Object, e As EventArgs) Handles btnRemove.Click
        With Me.ListBoxWorksheets
            Dim count As Short = .SelectedIndices.Count
            For i As Short = count - 1 To 0 Step -1
                Dim index As Short = .SelectedIndices.Item(i)
                .Items.RemoveAt(index)
            Next
        End With
    End Sub

#End Region

#Region "  --- 数据提取"

    ''' <summary>
    ''' 开始输出数据
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click
        If Not Me.BackgroundWorker1.IsBusy Then
            '打开进行数据操作的Excel程序
            If Me.F_ExcelApp Is Nothing Then
                Me.F_ExcelApp = New Excel.Application
            End If
            '初始化错误列表
            Me.F_ErrorList = New List(Of String)
            '
            Dim strWorkBook_ExportedTo As String = Me.txtbxSavePath.Text
            '打开保存数据的工作簿，并提取其中的所有工作表
            Dim listSheetNameInWkbk As New List(Of String)
            Try
                'If Me.F_WorkBook_ExportedTo Is Nothing Then
                If File.Exists(strWorkBook_ExportedTo) Then
                    F_WorkBook_ExportedTo = Me.F_ExcelApp.Workbooks.Open(strWorkBook_ExportedTo, UpdateLinks:=False, [ReadOnly]:=False)
                Else
                    F_WorkBook_ExportedTo = Me.F_ExcelApp.Workbooks.Add
                    F_WorkBook_ExportedTo.SaveAs(Filename:=strWorkBook_ExportedTo, _
                                              FileFormat:=Excel.XlFileFormat.xlOpenXMLWorkbook, _
                                              CreateBackup:=False)
                End If
                'End If
                '
                Dim AllSheets As Object = F_WorkBook_ExportedTo.Worksheets
                Dim shtCount As Integer = AllSheets.Count
                For Each shtInWorkbook As Worksheet In AllSheets
                    listSheetNameInWkbk.Add(shtInWorkbook.Name)
                Next
            Catch ex As Exception
                MessageBox.Show("保存数据的工作簿打开出错，请检查或者关闭此工作簿。", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End Try

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
                    Dim RangeInfo As New ExportToWorksheet.RangeInfoForExport(ExportedSheet, strTestRange, testRange.Columns.Count)
                    listRangeInfo.Add(RangeInfo)
                Next
            Catch ex As Exception
                MessageBox.Show("定义区域范围的格式出错，出错的格式为 : " & vbCrLf _
                                & strTestRange & "，请重新输入", "Error", _
                                 MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1)
                Exit Sub
            End Try

            ' -----------进行数据提取的工作簿对象数组------------------------ 
            Dim WkbkItems As System.Windows.Forms.ListBox.ObjectCollection = Me.ListBoxWorksheets.Items
            Dim WorkbooksCount As Integer = WkbkItems.Count
            '记录DataGridView控件中所有数据的数组
            Dim arrWkbk(0 To WorkbooksCount - 1) As String
            With Me.ListBoxWorksheets
                For i As Integer = 0 To WorkbooksCount - 1
                    Dim WkbkPath As String = WkbkItems.Item(i).ToString
                    arrWkbk(i) = WkbkPath
                Next
            End With

            ' ----------------------------------- 
            '是否要分析出提取数据的工作簿中的日期数据
            Dim blnParseDateFromFilePath As Boolean = False
            If Me.ChkboxParseDate.Checked Then blnParseDateFromFilePath = True
            '不允许再更改提取日期的正则表达式
            Me.ChkboxParseDate.Checked = False
            '开始提取数据
            Dim Export As New ExportToWorksheet(F_WorkBook_ExportedTo, arrWkbk, listRangeInfo, blnParseDateFromFilePath)
            Me.BackgroundWorker1.RunWorkerAsync(Export)
        End If
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
        Dim Count_Workbooks As Integer = Me.ListBoxWorksheets.Items.Count
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
    ''' 由工作簿的路径返回此组数据的表头信息
    ''' </summary>
    ''' <param name="FilePath">返回表头数据的依据：工作簿的路径</param>
    ''' <param name="ParseDateFromFilePath">是否要分析出提取数据的工作簿中的日期数据</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetColumnTitle(ByVal FilePath As String, ByVal ParseDateFromFilePath As Boolean) As String
        Dim filename As String = Path.GetFileNameWithoutExtension(FilePath)
        Dim ColumnTitle As String = filename
        '尝试从工作簿文件名分解出其中的日期信息
        If ParseDateFromFilePath Then
            Try
                Dim rg As New Regex(Me.F_regexPattern, RegexOptions.Singleline, New TimeSpan(10000000.0))
                Dim m As Match = rg.Match(filename)
                With m
                    If .Success Then
                        '按“年/月/日”的格式构造日期字符串
                        ColumnTitle = .Groups(Me.F_regexComponents(1)).Value & "/" &
                            .Groups(Me.F_regexComponents(2)).Value & "/" &
                            .Groups(Me.F_regexComponents(3)).Value
                    Else
                        Dim strError As String = "日期转换异常，异常的工作簿为： " & FilePath
                        Me.F_ErrorList.Add(strError)
                    End If
                End With
            Catch ex As Exception
                Dim strError As String = "日期转换异常，异常的工作簿为： " & FilePath
                Me.F_ErrorList.Add(strError)
            End Try
        End If
        Return ColumnTitle
    End Function
    ''' <summary>
    ''' 正式开始提取数据
    ''' </summary>
    ''' <param name="shtExtractFrom">要进行数据提取的工作表</param>
    ''' <param name="shtExportTo">放置提取的数据的工作表</param>
    ''' <param name="strRange">提取的数据区间</param>
    ''' <param name="ColNumToBeAdded">要放置的数据Range的第一个列号</param>
    ''' <param name="ColumnTitle">每一个Range数据的表头信息</param>
    ''' <remarks></remarks>
    Private Sub ExportData(ByVal shtExtractFrom As Excel.Worksheet, _
                           ByVal shtExportTo As Excel.Worksheet, _
                           ByVal strRange As String, ByVal ColNumToBeAdded As Integer, ByVal ColumnTitle As String)
        '下一个要放置的数据Range的列号
        'Dim ColNumToBeAdded As Integer = ColNum_FirstData

        '添加表头数据
        shtExportTo.Cells(cstRowNum_ColumnTitle, ColNumToBeAdded).Value = ColumnTitle

        '要提取的数据的范围
        Dim rgOut As Range = shtExtractFrom.Range(strRange)
        Dim ColsCount As Integer = rgOut.Columns.Count
        Dim RowsCount As Integer = rgOut.Rows.Count
        '要放置数据的区域范围
        Dim rgIn As Range
        With shtExportTo
            rgIn = .Range(.Cells(cstRowNum_FirstData, ColNumToBeAdded), _
                          .Cells(cstRowNum_FirstData + RowsCount - 1, ColNumToBeAdded + ColsCount - 1))
        End With
        '提取数据
        rgIn.Value = rgOut.Value
    End Sub

    '报告进度
    Private Sub BackgroundWorker1_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        Dim strHandlePath As String = CType(e.UserState.ToString, String)
        Me.lbSheetName.Text = strHandlePath
        Me.ProgressBar1.Value = e.ProgressPercentage
    End Sub

    ' 操作完成
    ''' <summary>
    ''' 操作完成，关闭Excel程序，写入异常信息，并控件进度条的显示
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Me.lbSheetName.Text = "Done"
        '激活更改提取日期的正则表达式
        Me.ChkboxParseDate.Checked = True
        '列举所有出错项
        If Me.F_WorkBook_ExportedTo IsNot Nothing Then
            '输出异常信息
            Dim ErrorFilePath As String = Path.Combine(Me.F_WorkBook_ExportedTo.Path, cstErrorInfoFileName)
            Dim thd As Thread = New Thread(AddressOf Me.ReportError)
            thd.Start({ErrorFilePath, Me.F_ErrorList})

            ' 保存工作簿中的数据
            Me.F_WorkBook_ExportedTo.Save()

            '关闭或者显示工作簿
            If Me.ChkBxOpenExcelWhileFinished.Checked Then
                Me.F_ExcelApp.Visible = True
                Me.F_WorkBook_ExportedTo.Worksheets.Item(1).Activate()
            Else
                Me.F_WorkBook_ExportedTo.Close(SaveChanges:=True)
                Me.F_WorkBook_ExportedTo = Nothing
                '关闭Excel程序
                Me.F_ExcelApp.Quit()
                Me.F_ExcelApp = Nothing
            End If
        End If
        '最后刷新进度条
        If Me.ProgressTimer Is Nothing Then
            Me.ProgressTimer = New System.Windows.Forms.Timer
        End If
        With Me.ProgressTimer
            .Interval = 500
            .Start()
        End With
    End Sub
    ''' <summary>
    ''' 将异常信息的集合写入文本中
    ''' </summary>
    ''' <param name="Parameters">新线程中的输入参数，它是一个有两个元素的数组，
    ''' 其中第1个元素代表异常信息文件的路径，第二个代表收集了异常信息的List集合</param>
    ''' <remarks></remarks>
    Private Sub ReportError(ByVal Parameters As Object)
        'ByVal ErrorFilePath As String, ByVal ErrorList As List(Of String)
        Dim ErrorFilePath As String = Parameters(0)
        Dim ErrorList As List(Of String) = Parameters(1)
        If ErrorList.Count > 0 Then
            Dim sw As New StreamWriter(ErrorFilePath, True)
            '上面这一步会在指定的路径上生成一个新的文件
            With sw
                .WriteLine(Date.Now.ToLongDateString & Date.Now.ToLongTimeString)
                For Each strError As String In ErrorList
                    .WriteLine(strError)
                Next
                'Close之前，文本文件中只没有数据，Close之后，数据被刷新到文本文件中。
                .Close()
            End With
        End If

    End Sub
    ''' <summary>
    ''' 在定时器触发时将进度条的值设置为0
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ProgressTimer_Tick(sender As Object, e As EventArgs) Handles ProgressTimer.Tick
        Me.ProgressBar1.Value = 0
        With Me.ProgressTimer
            .Stop()
            .Dispose()
            Me.ProgressTimer = Nothing
        End With
    End Sub

#End Region


End Class