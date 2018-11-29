Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop
Imports System.IO
Imports System.Threading
Imports System.Text.RegularExpressions
Public MustInherit Class frmDeriveData

    ''' <summary>
    ''' 此界面所处理的数据类型：Excel中的监测数据还是Word中的监测数据
    ''' </summary>
    ''' <remarks></remarks>
    Protected Enum ChildType
        Word
        Excel
    End Enum

#Region "  ---  Constants"

    ''' <summary>
    ''' 记录异常信息的文本的名称
    ''' </summary>
    ''' <remarks>其文件夹路径与输出数据的Excel工作簿的路径相同</remarks>
    Protected Const cstErrorInfoFileName As String = "ErrorInfo.txt"

    ''' <summary>
    ''' 每一列数据的表头信息所在的行，一般为第一行，一般为数据对应的日期
    ''' </summary>
    ''' <remarks></remarks>
    Protected Const cstRowNum_ColumnTitle As Byte = 1

    ''' <summary>
    ''' 提取的数据中的第一行在工作表中所要放置的行号，一般为第2行。第一行一般用来放数据对应的日期，第二行一般为预留行。
    ''' </summary>
    ''' <remarks></remarks>
    Protected Const cstRowNum_FirstData As Byte = 2

    ''' <summary>
    ''' 提取的数据中的第一列在工作表中所要放置的列号，一般为第2列。第1列用来放数据的说明
    ''' </summary>
    ''' <remarks></remarks>
    Protected Const cstColNum_FirstData As Byte = 2

#End Region

#Region "  ---  Fields"
    ''' <summary>
    ''' 此界面所处理的数据类型：Excel中的监测数据还是Word中的监测数据
    ''' </summary>
    ''' <remarks></remarks>
    Protected F_ChildType As ChildType

    ''' <summary>
    ''' 用于操作数据的Excel程序
    ''' </summary>
    ''' <remarks></remarks>
    Protected WithEvents F_ExcelApp As Excel.Application

    ''' <summary>
    ''' 保存提取后的数据的工作簿
    ''' </summary>
    ''' <remarks></remarks>
    Protected F_WorkBook_ExportedTo As Excel.Workbook
    '
    ''' <summary>
    ''' 搜索日期的正则表达式字符串
    ''' </summary>
    ''' <remarks></remarks>
    Protected F_regexPattern As String
    ''' <summary>
    ''' 利用正则表达式提取的字符中，{文件序号，年，月，日}分别在Match.Groups集合中的下标值。用值0来代表没有此项。
    ''' </summary>
    ''' <remarks>Match.Groups(0)返回的是Match结果本身，并不属于要提取的数据。</remarks>
    Protected F_regexComponents(0 To 3) As Byte
    '
    ''' <summary>
    ''' 异常信息的集合
    ''' </summary>
    ''' <remarks></remarks>
    Protected F_ErrorList As List(Of String)

    ''' <summary>
    ''' 定时触发器
    ''' </summary>
    ''' <remarks></remarks>
    Protected WithEvents ProgressTimer As System.Windows.Forms.Timer

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Protected listSheetNameInWkbk As List(Of String)

    ''' <summary>
    ''' 列表框中所记录的所有要进行数据提取的Excel或者Word文档的路径
    ''' </summary>
    ''' <remarks></remarks>
    Protected arrDocPaths As String()
#End Region

#Region "  ---  窗体的加载与关闭"

    ''' <summary>
    ''' 构造函数
    ''' </summary>
    ''' <remarks></remarks>
    Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        '
        With Me.BackgroundWorker1
            .WorkerReportsProgress = True
            .WorkerSupportsCancellation = True
        End With
        '
        Me.txtbxSavePath.Text = Path.Combine( _
            System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory, Environment.SpecialFolderOption.None), _
            "数据提取.xlsx")
        Me.F_ChildType = ChildType.Excel
        Me.F_regexComponents = {1, 2, 3, 4}
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
                If .Visible = False Then
                    For Each wkbk As Excel.Workbook In .Workbooks
                        wkbk.Close(SaveChanges:=False)
                    Next
                    .Quit()
                    Me.F_ExcelApp = Nothing
                End If
            End With
        End If
    End Sub

#End Region

#Region "  ---  界面操作"
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
    Private Sub APPLICATION_MAINFORM_DragDrop(sender As Object, e As DragEventArgs) Handles ListBoxDocuments.DragDrop
        Dim FilePaths As String() = e.Data.GetData(DataFormats.FileDrop)
        ' DoSomething with the Files or Directories that are droped in.
        Dim arrExcelFilePath As New List(Of String)
        For Each filepath As String In FilePaths
            Dim ext As String = Path.GetExtension(filepath)
            Dim str As String = ".xlsx.xls.xlsb"
            If Me.F_ChildType = ChildType.Word Then
                str = ".docx.doc.docm"
            End If
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
    '
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
            If Me.F_ChildType = ChildType.Word Then
                .Title = "选择要进行数据提取的Word文件"
                .Filter = "Word文档(*.docx, *.doc, *.docm)|*.docx;*.doc;*.docm"
            End If
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
        If Me.F_ChildType = ChildType.Excel Then
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
                        Me.ListBoxDocuments.Items.Add(strFile)
                    End If
                Next
            End If
        Else
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
        End If

    End Sub

    '保存数据的Excel工作簿路径
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
                ' Exit Sub
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

#End Region

#Region "  --- 数据提取"

    ''' <summary>
    ''' 开始输出数据
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub btnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click
        If Not Me.BackgroundWorker1.IsBusy Then
            '打开进行数据操作的Excel程序
            If Me.F_ExcelApp Is Nothing Then
                Me.F_ExcelApp = New Excel.Application
            End If
            With Me.F_ExcelApp
                .DisplayAlerts = False
                '一般情况下，默认是隐藏的，如果原来是打开的，则手动隐藏
                .Visible = False
            End With
            '初始化错误列表
            Me.F_ErrorList = New List(Of String)

            '---------- 打开保存数据的工作簿，并提取其中的所有工作表 ----------------
            Dim strWorkBook_ExportedTo As String = Me.txtbxSavePath.Text
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
            Catch ex As Exception
                MessageBox.Show("保存数据的工作簿打开出错，请检查或者关闭此工作簿。", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End Try
            listSheetNameInWkbk = New List(Of String)
            For Each shtInWorkbook As Worksheet In F_WorkBook_ExportedTo.Worksheets
                listSheetNameInWkbk.Add(shtInWorkbook.Name)
            Next
            ' -----------进行数据提取的工作簿对象数组------------------------ 
            Dim WkbkItems As System.Windows.Forms.ListBox.ObjectCollection = Me.ListBoxDocuments.Items
            Dim WorkbooksCount As Integer = WkbkItems.Count
            ReDim arrDocPaths(0 To WorkbooksCount - 1)
            With Me.ListBoxDocuments
                For i As Integer = 0 To WorkbooksCount - 1
                    arrDocPaths(i) = WkbkItems.Item(i).ToString
                Next
            End With
            StartExportData()
        End If
    End Sub

    ''' <summary>
    ''' 开始输出数据
    ''' </summary>
    ''' <remarks></remarks>
    Protected MustOverride Sub StartExportData()

    ''' <summary>
    ''' 由工作簿的路径返回此组数据的表头信息
    ''' </summary>
    ''' <param name="FilePath">返回表头数据的依据：工作簿的路径</param>
    ''' <param name="ParseDateFromFilePath">是否要分析出提取数据的工作簿中的日期数据</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Function GetColumnTitle(ByVal FilePath As String, ByVal ParseDateFromFilePath As Boolean) As String
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
#End Region
#Region "  --- 操作进度与对应处理"


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