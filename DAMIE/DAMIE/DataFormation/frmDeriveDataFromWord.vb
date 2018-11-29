Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Interop
Imports System.IO
Imports System.Threading
Imports System.Text.RegularExpressions

Public Class frmDeriveDataFromWord

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
    ''' 记录异常信息的文本的名称
    ''' </summary>
    ''' <remarks>其文件夹路径与输出数据的Excel工作簿的路径相同</remarks>
    Const cstErrorInfoFileName As String = "ErrorInfo.txt"

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

    ''' <summary>
    ''' 每一组提取数据所占据的列数。从数据提取上来看，此字段并没有什么作用，因为一般情况下，它的值都应该是2。
    ''' 但是从表格的设计上来看，它的值可以用来腾出空的列以放置其他数据。
    ''' </summary>
    ''' <remarks></remarks>
    Const cstColumnsCountToBeAdded As Byte=2

#End Region

#Region "  ---  Fields"

    ''' <summary>
    ''' 保存数据的Excel程序
    ''' </summary>
    ''' <remarks></remarks>
    Private WithEvents F_ExcelApp As Excel.Application
    ''' <summary>
    ''' 从Word文档中提取数据
    ''' </summary>
    ''' <remarks></remarks>
    Private WithEvents F_WordApp As Word.Application
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
    ''' 定时触发器，用来将进度滚动条的值设置为0
    ''' </summary>
    ''' <remarks></remarks>
    Private WithEvents ProgressTimer As System.Windows.Forms.Timer

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

#Region "  ---  窗体的加载与关闭"

    ''' <summary>
    ''' 窗体的加载
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub DeriveDataFromWord_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '设置后台线程的属性
        With Me.BkgWk_Extract
            .WorkerReportsProgress = True
            .WorkerSupportsCancellation = True
        End With
        '
        Me.txtbxSavePath.Text = Path.Combine( _
            System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory, Environment.SpecialFolderOption.None), _
            "数据提取.xlsx")
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
    '
    ' 拖拽操作
    Private Sub APPLICATION_MAINFORM_DragDrop(sender As Object, e As DragEventArgs) Handles ListBoxDocuments.DragDrop
        Dim FilePaths As String() = e.Data.GetData(DataFormats.FileDrop)
        ' DoSomething with the Files or Directories that are droped in.
        Dim arrExcelFilePath As New List(Of String)
        For Each filepath As String In FilePaths
            Dim ext As String = Path.GetExtension(filepath)
            Dim str As String = ".docx.doc.docm"
            Dim blnExtensionMatched As Boolean = str.Contains(ext)
            If blnExtensionMatched Then
                Me.ListBoxDocuments.Items.Add(filepath)
            End If
        Next
    End Sub
    Private Sub APPLICATION_MAINFORM_DragEnter(sender As Object, e As DragEventArgs) Handles ListBoxDocuments.DragEnter
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
            .Title = "选择要进行数据提取的Word文件"
            .Filter = "Word文档(*.docx, *.doc, *.docm)|*.docx;*.doc;*.docm"
            .FilterIndex = 2
            .Multiselect = True
            If .ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                FilePaths = .FileNames
            Else
                Exit Sub
            End If
        End With
        If FilePaths.Length > 0 Then
            Me.ListBoxDocuments.Items.AddRange(FilePaths)
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
            .Description = "添加文件夹中的全部Word文档"
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
                If String.Compare(ext, ".doc", True) = 0 OrElse _
                    String.Compare(ext, ".docx", True) = 0 OrElse _
                    String.Compare(ext, ".doxm", True) = 0 Then
                    Me.ListBoxDocuments.Items.Add(strFile)
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
        With Me.ListBoxDocuments
            Dim count As Short = .SelectedIndices.Count
            For i As Short = count - 1 To 0 Step -1
                Dim index As Short = .SelectedIndices.Item(i)
                .Items.RemoveAt(index)
            Next
        End With
    End Sub

    ''' <summary>
    ''' 在DataGridView中，添加新行时，将其搜索方向设置为“按行”。
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub MyDataGridView1_RowsAdded(sender As Object, e As DataGridViewRowsAddedEventArgs) Handles MyDataGridView1.RowsAdded
        With Me.MyDataGridView1
            If e.RowIndex >= 1 Then
                Dim a = .Item(2, e.RowIndex - 1)
                If a.Value Is Nothing Then
                    a.Value = "按行"
                End If
            End If
        End With
    End Sub

#End Region

#Region "  --- 数据提取"

    ''' <summary>
    ''' 开始输出数据
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub btnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click

        If Not Me.BkgWk_Extract.IsBusy Then
            '
            Me.blnIsBeingExtracting = True
            '打开进行数据操作的Excel程序
            If Me.F_ExcelApp Is Nothing Then
                Me.F_ExcelApp = New Excel.Application
                With Me.F_ExcelApp
                    .DisplayAlerts = False
                    '一般情况下，默认是隐藏的，如果原来是打开的，则手动隐藏
                    .Visible = False
                End With
            End If

            '打开进行数据提取的Word程序
            If Me.F_WordApp Is Nothing Then
                Me.F_WordApp = New Word.Application
                Me.F_WordApp.Visible = False
            End If


            '初始化错误列表
            Me.F_ErrorList = New List(Of String)
            '
            Dim strWorkBook_ExportedTo As String = Me.txtbxSavePath.Text


            '---------- 打开保存数据的工作簿，并提取其中的所有工作表 ----------------
            Dim listPointsTag As New List(Of String)
            Dim listSheetNameInWkbk As New List(Of String)
            Try
                If File.Exists(strWorkBook_ExportedTo) Then
                    F_WorkBook_ExportedTo = Me.F_ExcelApp.Workbooks.Open(strWorkBook_ExportedTo, UpdateLinks:=False, [ReadOnly]:=False)
                Else
                    F_WorkBook_ExportedTo = Me.F_ExcelApp.Workbooks.Add
                    F_WorkBook_ExportedTo.SaveAs(Filename:=strWorkBook_ExportedTo, _
                                              FileFormat:=Excel.XlFileFormat.xlOpenXMLWorkbook, _
                                              CreateBackup:=False)
                End If
                Me.F_BufferSheet = Me.F_WorkBook_ExportedTo.Worksheets.Add
                '
                Dim AllSheets As Object = F_WorkBook_ExportedTo.Worksheets
                For Each shtInWorkbook As Worksheet In AllSheets
                    listSheetNameInWkbk.Add(shtInWorkbook.Name)
                Next
            Catch ex As Exception
                MessageBox.Show("保存数据的Word文档打开出错，请检查或者关闭此文档。", "Error", _
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End Try

            ' ------------- 提取每一个工作表与Range范围的格式 -------------并返回DataGridView中的所有数据
            Me.F_DicPointsInfo = SearchPointsInfo(Me.F_WorkBook_ExportedTo)
            If F_DicPointsInfo Is Nothing Then Exit Sub


            ' -----------进行数据提取的Document对象数组------------------------ 
            Dim DocItems As System.Windows.Forms.ListBox.ObjectCollection = Me.ListBoxDocuments.Items
            Dim DocsCount As Integer = DocItems.Count
            '记录DataGridView控件中所有数据的数组
            Dim arrDocsPath(0 To DocsCount - 1) As String
            With Me.ListBoxDocuments
                For i As Integer = 0 To DocsCount - 1
                    arrDocsPath(i) = DocItems.Item(i).ToString
                Next
            End With

            ' ----------------------------------- 
            '是否要分析出提取数据的工作簿中的日期数据
            Dim blnParseDateFromFilePath As Boolean = False
            If Me.ChkboxParseDate.Checked Then blnParseDateFromFilePath = True
            '不允许再更改提取日期的正则表达式
            Me.ChkboxParseDate.Checked = False
            ' ---------------------- 开始提取数据 ---------------------
            Dim Export As New ExportToWorksheet(F_WorkBook_ExportedTo, arrDocsPath, _
                                                blnParseDateFromFilePath, Me.F_BufferSheet)
            Me.BkgWk_Extract.RunWorkerAsync(Export)
        End If
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
    Private Sub StartToDoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BkgWk_Extract.DoWork
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
        Me.BkgWk_Extract.ReportProgress(percent, "")
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

                        Me.BkgWk_Extract.ReportProgress(((iDoc + iTable / CountTables)) * unit, _
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
                Me.BkgWk_Extract.ReportProgress((iDoc + 1) * unit, "正在提取文档：" & strDocPath)
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
                Dim rgNextPoint As Excel.Range
                With Me.F_BufferSheet.UsedRange
                    rgNextPoint = .Find(What:=PointInfo.PointTag, _
                                                       SearchOrder:=PointInfo.SearchOrder, _
                                                       LookAt:=XlLookAt.xlPart, _
                                                       LookIn:=XlFindLookIn.xlValues, _
                                                       SearchDirection:=XlSearchDirection.xlNext, _
                                                       MatchCase:=False)
                    If rgNextPoint IsNot Nothing Then
                        '当搜索到指定查找区域的末尾时，此方法将绕回到区域的开始位置继续搜索。
                        '发生绕回后，要停止搜索，可保存第一个找到的单元格地址，然后测试后面找到的每个单元格地址是否与其相同。
                        Dim firstAddress As String = rgNextPoint.Address
                        '提取数据并写入最终的工作表
                        Do
                            With PointInfo
                                sheetExportTo.Cells(.RowNumToBeWritten, .ColNumToBeWritten).Value = rgNextPoint.Value
                                sheetExportTo.Cells(.RowNumToBeWritten, .ColNumToBeWritten + 1).Value = rgNextPoint _
                                    .Offset(0, PointInfo.Offset).Value
                                .RowNumToBeWritten += 1
                            End With
                            rgNextPoint = .FindNext(rgNextPoint)
                        Loop While rgNextPoint IsNot Nothing And String.Compare(rgNextPoint.Address, firstAddress) <> 0

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

    '报告进度
    ''' <summary>
    ''' 报告进度
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BackgroundWorker1_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles BkgWk_Extract.ProgressChanged
        Dim strHandlePath As String = CType(e.UserState.ToString, String)
        Me.lbSheetName.Text = strHandlePath
        Me.ProgressBar1.Value = e.ProgressPercentage
    End Sub

    '操作完成
    ''' <summary>
    ''' 操作完成，关闭Excel程序，写入异常信息，并控件进度条的显示
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BkgWk_Extract.RunWorkerCompleted
        Me.lbSheetName.Text = "Done"
        '激活更改提取日期的正则表达式
        Me.ChkboxParseDate.Checked = True
        '列举所有出错项
        If Me.F_WorkBook_ExportedTo IsNot Nothing Then
            '输出异常信息
            Dim ErrorFilePath As String = Path.Combine(Me.F_WorkBook_ExportedTo.Path, cstErrorInfoFileName)
            Dim thd As Thread = New Thread(AddressOf Me.ReportError)
            thd.Start({ErrorFilePath, Me.F_ErrorList})
            'Call ReportError(ErrorFilePath, Me.F_ErrorList)

            '删除用来缓存数据的中间工作表
            Me.F_BufferSheet.Delete()
            Me.F_BufferSheet = Nothing

            '保存最终结果数据
            Me.F_WorkBook_ExportedTo.Save()

            '关闭或者显示工作簿
            If Me.ChkBxOpenExcelWhileFinished.Checked Then
                Me.F_ExcelApp.Visible = True
                Me.F_WorkBook_ExportedTo.Worksheets.Item(1).Activate()
            Else
                Me.F_WorkBook_ExportedTo.Close(SaveChanges:=True)
                Me.F_WorkBook_ExportedTo = Nothing
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
        '
        Me.blnIsBeingExtracting = False
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
        Me.lbSheetName.Text = ""
        With Me.ProgressTimer
            .Stop()
            .Dispose()
            Me.ProgressTimer = Nothing
        End With
    End Sub

#End Region

End Class