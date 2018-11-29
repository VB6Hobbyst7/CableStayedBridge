Imports Microsoft.Office.Interop.Excel
Imports DAMIE.Miscellaneous
Imports DAMIE.Constants
Imports DAMIE.GlobalApp_Form
Imports eZstd.eZAPI
Namespace All_Drawings_In_Application
    Public Class Cls_ExcelForMonitorDrawing
        Implements Dictionary_AutoKey(Of Cls_ExcelForMonitorDrawing).I_Dictionary_AutoKey

#Region "  ---  Properties"

        Private WithEvents P_Excelapp As Application
        Public Property Application As Application
            Get
                Return Me.P_Excelapp
            End Get
            Private Set(value As Application)
                If value IsNot Nothing Then
                    '不弹出警告对话框
                    value.DisplayAlerts = False
                    '获取Excel的进程
                    Dim processId As Integer
                    APIWindows.GetWindowThreadProcessId(value.Hwnd, processId)
                    F_ExcelProcess = Process.GetProcessById(processId)
                Else
                    F_ExcelProcess = Nothing
                End If
                Me.P_Excelapp = value
            End Set
        End Property

        ''' <summary>
        ''' 此Excel监测曲线绘图窗口在主程序的集合中的关键字，用来在集合键值对中对此窗口进行索引
        ''' </summary>
        ''' <remarks></remarks>
        Private P_Key As Integer
        ''' <summary>
        ''' 此元素在其所在的集合中的键，这个键是在元素添加到集合中时自动生成的，
        ''' 所以应该在执行集合.Add函数时，用元素的Key属性接受函数的输出值。
        ''' 在集合中添加此元素：Me.Key=Me所在的集合.Add(Me)
        ''' 在集合中索引此元素：集合.item(me.key)
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Key As Integer _
            Implements Dictionary_AutoKey(Of Cls_ExcelForMonitorDrawing).I_Dictionary_AutoKey.Key
            Get
                Return P_Key
            End Get
        End Property

        ''' <summary>
        ''' 在一个监测曲线绘图的Application的一个工作簿中，当前活动的那一个绘图工作表。
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property ActiveMntDrawingSheet As ClsDrawing_Mnt_Base
            Get
                Return P_Mnt_Drawings.Last.Value
            End Get
        End Property

        '主界面中的所有监测曲线图的集合
        Private P_Mnt_Drawings As New Dictionary_AutoKey(Of ClsDrawing_Mnt_Base)
        Public Property Mnt_Drawings As Dictionary_AutoKey(Of ClsDrawing_Mnt_Base)
            Get
                Return P_Mnt_Drawings
            End Get
            Set(value As Dictionary_AutoKey(Of ClsDrawing_Mnt_Base))
                P_Mnt_Drawings = value
            End Set
        End Property

#End Region

#Region "  ---  Fields"

        ''' <summary>
        ''' Excel程序的进程对象，用来对此进程进行相差的操作，比如，关闭进程
        ''' </summary>
        ''' <remarks></remarks>
        Private F_ExcelProcess As Process

#End Region

        ''' <summary>
        ''' 构造函数
        ''' </summary>
        ''' <param name="Application">进行绘图的Excel绘图程序</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal Application As Application)
            '在集合中以此对象的ID值来索引此对象
            Me.Application = Application
            Me.P_Key = GlobalApplication.Application.MntDrawing_ExcelApps.Add(Me)
        End Sub


        ''' <summary>
        ''' Excel程序关闭时触发的事件
        ''' </summary>
        ''' <param name="wkbk"></param>
        ''' <param name="cancel"></param>
        ''' <remarks>在处理过程中不能再执行wkbk.Close方法，不然会多次执行文档的关闭，从而出错。</remarks>
        Private Sub AppQuit(wkbk As Workbook, ByRef cancel As Boolean) Handles P_Excelapp.WorkbookBeforeClose
            Try
                '方法一：利用进程：退出进程，其操作与与用户使用系统菜单关闭应用程序主窗口的行为一样
                'F_ExcelProcess.CloseMainWindow()
                '方法二:Application.Quit()
                unFixWindow(wkbk.Application.Hwnd)
                wkbk.Application.Quit()
                '但是还有一个问题，上面两种方法，都会出现是否要保存文档更改的弹窗，
                '为了不出现此弹窗而直接关闭程序，可以设置Application.DisplayAlerts = False。
                Me.Application = Nothing
            Catch ex As Exception
                MessageBox.Show("关闭Excel程序出错！" & vbCrLf & ex.Message & vbCrLf & "报错位置：" & ex.TargetSite.Name, _
     "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Finally
                Me.RemoveFormCollection()
                '刷新滚动窗口的列表框的界面显示
                APPLICATION_MAINFORM.MainForm.Form_Rolling.OnRollingDrawingsRefreshed()
            End Try
        End Sub

        ''' <summary>
        ''' 将自己从所在的集合中删除
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function RemoveFormCollection() As Boolean
            Try
                '！！这这里可能会出现执行Excel事件时主程序被初始化——主程序的调用问题
                '所以采用了如下全局共享属性的解决方法
                With GlobalApplication.Application
                    .MntDrawing_ExcelApps.Remove(Me.Key)
                End With
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function

    End Class
End Namespace