Imports System.ComponentModel

Namespace AME_UserControl
    ''' <summary>
    ''' 自定义控件，用来对于PictureBox中的图片进行中心旋转
    ''' </summary>
    ''' <remarks></remarks>
    Public Class PictureBox_AutoRoll
        Inherits PictureBox

#Region "  ---  定义与声明"

#Region "  ---  属性值定义"

        ''' <summary>
        ''' 要进行滚动旋转的Image对象
        ''' </summary>
        ''' <remarks></remarks>
        Private _RollingImage As Image
        ''' <summary>
        ''' 要进行滚动旋转的Image对象
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property RollingImage As Image
            Get
                Return Me._RollingImage
            End Get
        End Property

        ''' <summary>
        ''' Gets or sets the time, in milliseconds, 
        ''' before the Tick event is raised relative to the last occurrence of the Tick event.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <Browsable(True), Description("Gets or sets the time, in milliseconds, before the Tick event is raised relative to the last occurrence of the Tick event.")>
        Public Property Interval As Integer

        ''' <summary>
        ''' 每一次旋转的增量角，以度来表示。
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <Browsable(True), Description("The angle, in degrees, to be added  each time the image is rotated.")>
        Public Property RollingAngle As Single

#End Region

#Region "  ---  字段值定义"

        ''' <summary>
        ''' 对图形进行滚动旋转的定时触发器
        ''' </summary>
        ''' <remarks></remarks>
        Private WithEvents Timer_ProcessRing As System.Windows.Forms.Timer

#End Region

#End Region

#Region "  ---  构造函数与窗体的加载、打开与关闭"
        Public Sub New()
            Call Me.InitializeComponent()
        End Sub

        Private Sub InitializeComponent()

        End Sub

#End Region

        ''' <summary>
        ''' 开始将图形进行旋转，在应用此方法前，
        ''' 请务必先为控件的Image属性赋值，即指定要进行旋转的图形对象。
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub StartRolling()
            If Me.Image Is Nothing Then
                MessageBox.Show("请先在Image属性中指定要进行滚动旋转的图形", "Error", _
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            '为只读属性RollingImage赋值
            Me._RollingImage = Me.Image
            '为定时器赋值
            If Me.Timer_ProcessRing Is Nothing Then
                Me.Timer_ProcessRing = New System.Windows.Forms.Timer
            End If
            '
            With Me.Timer_ProcessRing
                .Interval = Me.Interval
                .Start()
            End With
        End Sub
        Private Sub Rolling() Handles Timer_ProcessRing.Tick
            Static ang As Integer = 0
            With Me
                Dim img As Image = Me._RollingImage
                Dim bitSize As Single = Math.Sqrt(img.Width ^ 2 + img.Height ^ 2)
                bitSize = img.Width
                '定义一张新画纸
                Dim bmp As Bitmap = New Bitmap(CInt(bitSize), CInt(bitSize))
                '为画纸创建一个画板，用来在画纸上进行相关的绘画
                '此时画板的坐标系与画纸的坐标系相同Coord1=Coord_map
                Using g As Graphics = Graphics.FromImage(bmp)
                    '将画板坐标系的原点移动到画纸中心——>坐标系Coord2
                    g.TranslateTransform(bitSize / 2, bitSize / 2)
                    '将画板坐标系在Coord2的基础上旋转一定的角度——>坐标系Coord3
                    g.RotateTransform(ang)
                    '将画板坐标系在Coord3的基础上将原点平移到新坐标，
                    '使得所画的内容的中心点位于画纸的中心——>坐标系Coord4
                    g.TranslateTransform(-bitSize / 2, -bitSize / 2)
                    '在画纸bmp上画图：所绘图形的定位是以画板的坐标系Coord4为基准，
                    '并通过一个矩形来定义图形在坐标系Coord4下的位置与尺寸，如果目标矩形的大小
                    '与原始图像的大小不同，原始图像将进行缩放，以适应目标矩形。
                    g.DrawImage(img, New System.Drawing.Rectangle(0, 0, img.Width, img.Width))
                End Using
                '将画纸（及其上面的图形）赋值给PictureBox控件的Image属性，以在控件上进行显示。
                .Image = bmp
            End With
            Try
                '增加旋转的角度
                ang += Me.RollingAngle
            Catch ex As OverflowException
                '如果角度的值溢出，则将其重置为0
                ang = 0
            End Try
        End Sub

        ''' <summary>
        ''' 停止图形的旋转
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub StopRolling()
            If Me.Timer_ProcessRing IsNot Nothing Then
                With Me.Timer_ProcessRing
                    .Stop()
                    .Dispose()
                    Me.Timer_ProcessRing = Nothing
                End With
            End If
        End Sub
    End Class

End Namespace