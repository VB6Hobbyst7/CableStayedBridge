Public Class frmRegexDate

#Region "  ---  Declarations & Definitions"

#Region "  ---  Types"

#End Region

#Region "  ---  Constants"

#End Region

#Region "  ---  Properties"

#End Region

#Region "  ---  Fields"
    ''' <summary>
    ''' 搜索日期的正则表达式字符串
    ''' </summary>
    ''' <remarks></remarks>
    Private F_Pattern As String
    ''' <summary>
    ''' 要提取的字符中，{文件序号，年，月，日}分别在Match.Groups集合中的下标值。用值0来代表没有此项。
    ''' </summary>
    ''' <remarks>Match.Groups(0)返回的是Match结果本身，并不属于要提取的数据。</remarks>
    Private F_Components(0 To 3) As Byte

#End Region

#End Region

#Region "  ---  构造函数与窗体的加载、打开与关闭"
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Me.Button1.Tag = 1
        Me.Button2.Tag = 2
        Me.Button3.Tag = 3
        Me.Button4.Tag = 4

    End Sub

    ''' <summary>
    ''' 弹出窗口，开始执行操作
    ''' </summary>
    ''' <param name="pattern"></param>
    ''' <param name="Components"></param>
    ''' <remarks></remarks>
    Public Overloads Sub ShowDialog(ByRef pattern As String, ByRef Components As Byte())
        Dim result = MyBase.ShowDialog()
        If result = Windows.Forms.DialogResult.OK Then
            Components = Me.F_Components
            pattern = Me.F_Pattern
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles BtnCancel.Click
        Me.Close()
    End Sub
#End Region

#Region "  ---  界面操作"
    ''' <summary>
    ''' 通过点击按钮来设置对应的数据类型
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub SetComponent(sender As Object, e As EventArgs) Handles Button1.Click, Button2.Click, Button3.Click, Button4.Click
        Dim bt As Button = DirectCast(sender, Button)
        With bt
            Select Case .Tag
                Case 1
                    .Tag = 2
                    .Text = "年"
                Case 2
                    .Tag = 3
                    .Text = "月"
                Case 3
                    .Tag = 4
                    .Text = "日"
                Case 4
                    .Tag = 1
                    .Text = "序号"
            End Select
        End With
    End Sub

    Private Sub selectText(sender As Object, e As EventArgs) Handles _
        Textbox_btn1.MouseClick, Textbox_btn2.MouseClick, Textbox_btn3.MouseClick, Textbox_btn4.MouseClick
        Static i As Long
        i += 1
        Dim tb As TextBox = DirectCast(sender, TextBox)
        tb.SelectAll()
    End Sub

    ''' <summary>
    ''' 实时刷新正则表达式
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub GenerateRegex() Handles _
        TextBox1.TextChanged, TextBox2.TextChanged, TextBox3.TextChanged, TextBox4.TextChanged, TextBox5.TextChanged, _
        Textbox_btn1.TextChanged, Textbox_btn2.TextChanged, Textbox_btn3.TextChanged, Textbox_btn4.TextChanged
        '下面格式字符串中的0~8分别代表：  前缀字符、              序号数字的个数、
        '                               序号与年的分隔字符、    表示年的数值的数字个数、
        '                               年与月的分隔字符、      表示月的数值的数字个数、
        '                               月与日的分隔字符、      表示日的数值的数字个数、
        '                               后缀字符()
        Dim strRegexp As String = String.Format("{0}\s*(\d{{0,{1}}})" &
                                                "\s*{2}\s*(\d{{{3}}})" &
                                                "\s*{4}\s*(\d{{{5}}})" &
                                                "\s*{6}\s*(\d{{{7}}})\s*{8}",
                TextBox1.Text, Textbox_btn1.Text, TextBox2.Text, Textbox_btn2.Text, _
                 TextBox3.Text, Textbox_btn3.Text, TextBox4.Text, Textbox_btn4.Text, TextBox5.Text)
        Me.F_Pattern = strRegexp
        Me.Label_Regex.Text = strRegexp
    End Sub

#End Region

    ''' <summary>
    ''' 执行操作，确认最终的正则表达式
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BtnOk_Click(sender As Object, e As EventArgs) Handles BtnOk.Click
        '要提取的字符中，{文件序号，年，月，日}分别在Match.Groups集合中的下标值。用值0来代表没有此项。
        Dim component(0 To 3) As Byte
        Dim btnTag(0 To 3) As Byte          '每一个按钮所代表的数据类型，其Tag值为1~4
        ' button.Tag指的是每一个button所代表的数据类型
        btnTag = {Button1.Tag, Button2.Tag, Button3.Tag, Button4.Tag}
        For i As Byte = 0 To 3
            component(i) = btnTag(i)
        Next
        ' 查看数组中是否有相同的元素，而不是有且只有"1、2、3、4"这四个元素
        For i As SByte = 0 To 2
            For j As SByte = i + 1 To 3
                If component(i) = component(j) Then
                    Return
                End If
            Next
        Next
        '验证成功，所有的数据合法
        Me.F_Components = component
        Me.DialogResult = Windows.Forms.DialogResult.OK
        Me.Close()    '这一句是必须的，用以返回showDialog方法。
    End Sub
End Class