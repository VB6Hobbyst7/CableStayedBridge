// VBConversions Note: VB project level imports
using System.Data;
using System.Diagnostics;
using System.Xml.Linq;
using System.Drawing;
using System.Collections.Generic;
using Microsoft.VisualBasic;
using System.Collections;
using System;
using System.Windows.Forms;
using System.Threading;
using System.Threading.Tasks;
using System.Linq;
using System.IO;
// End of VB project level imports

using System.ComponentModel;


namespace CableStayedBridge
{
	namespace AME_UserControl
	{
		/// <summary>
		/// 自定义控件，用来对于PictureBox中的图片进行中心旋转
		/// </summary>
		/// <remarks></remarks>
		public class PictureBox_AutoRoll : PictureBox
		{
			
#region   ---  定义与声明
			
#region   ---  属性值定义
			
			/// <summary>
			/// 要进行滚动旋转的Image对象
			/// </summary>
			/// <remarks></remarks>
			private Image _RollingImage;
			/// <summary>
			/// 要进行滚动旋转的Image对象
			/// </summary>
			/// <value></value>
			/// <returns></returns>
			/// <remarks></remarks>
public Image RollingImage
			{
				get
				{
					return this._RollingImage;
				}
			}
			
			/// <summary>
			/// Gets or sets the time, in milliseconds,
			/// before the Tick event is raised relative to the last occurrence of the Tick event.
			/// </summary>
			/// <value></value>
			/// <returns></returns>
			/// <remarks></remarks>
			[Browsable(true), Description("Gets or sets the time, in milliseconds, before the Tick event is raised relative to the last occurrence of the Tick event.")]public int Interval {get; set;}
			
			/// <summary>
			/// 每一次旋转的增量角，以度来表示。
			/// </summary>
			/// <value></value>
			/// <returns></returns>
			/// <remarks></remarks>
			[Browsable(true), Description("The angle, in degrees, to be added  each time the image is rotated.")]public float RollingAngle {get; set;}
			
#endregion
			
#region   ---  字段值定义
			
			/// <summary>
			/// 对图形进行滚动旋转的定时触发器
			/// </summary>
			/// <remarks></remarks>
			private System.Windows.Forms.Timer Timer_ProcessRing;
			
#endregion
			
#endregion
			
#region   ---  构造函数与窗体的加载、打开与关闭
			public PictureBox_AutoRoll()
			{
				this.InitializeComponent();
			}
			
			private void InitializeComponent()
			{
				
			}
			
#endregion
			
			/// <summary>
			/// 开始将图形进行旋转，在应用此方法前，
			/// 请务必先为控件的Image属性赋值，即指定要进行旋转的图形对象。
			/// </summary>
			/// <remarks></remarks>
			public void StartRolling()
			{
				if (this.Image == null)
				{
					MessageBox.Show("请先在Image属性中指定要进行滚动旋转的图形", "Error", 
						MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}
				//为只读属性RollingImage赋值
				this._RollingImage = this.Image;
				//为定时器赋值
				if (this.Timer_ProcessRing == null)
				{
					this.Timer_ProcessRing = new System.Windows.Forms.Timer();
					this.Timer_ProcessRing.Tick += this.Rolling;
				}
				//
				this.Timer_ProcessRing.Interval = this.Interval;
				this.Timer_ProcessRing.Start();
			}
			// VBConversions Note: Former VB static variables moved to class level because they aren't supported in C#.
			private int Rolling_ang = 0;
			
			private void Rolling(System.Object sender, System.EventArgs e)
			{
				// static int ang = 0; VBConversions Note: Static variable moved to class level and renamed Rolling_ang. Local static variables are not supported in C#.
				PictureBox_AutoRoll with_1 = this;
				Image img = this._RollingImage;
				float bitSize = (float) (Math.Sqrt(Math.Pow(img.Width, 2) + Math.Pow(img.Height, 2)));
				bitSize = img.Width;
				//定义一张新画纸
				Bitmap bmp = new Bitmap((int) bitSize, (int) bitSize);
				//为画纸创建一个画板，用来在画纸上进行相关的绘画
				//此时画板的坐标系与画纸的坐标系相同Coord1=Coord_map
				using (Graphics g = Graphics.FromImage(bmp))
				{
					//将画板坐标系的原点移动到画纸中心——>坐标系Coord2
					g.TranslateTransform(bitSize / 2, bitSize / 2);
					//将画板坐标系在Coord2的基础上旋转一定的角度——>坐标系Coord3
					g.RotateTransform(Rolling_ang);
					//将画板坐标系在Coord3的基础上将原点平移到新坐标，
					//使得所画的内容的中心点位于画纸的中心——>坐标系Coord4
					g.TranslateTransform(- bitSize / 2, - bitSize / 2);
					//在画纸bmp上画图：所绘图形的定位是以画板的坐标系Coord4为基准，
					//并通过一个矩形来定义图形在坐标系Coord4下的位置与尺寸，如果目标矩形的大小
					//与原始图像的大小不同，原始图像将进行缩放，以适应目标矩形。
					g.DrawImage(img, new System.Drawing.Rectangle(0, 0, img.Width, img.Width));
				}
				
				//将画纸（及其上面的图形）赋值给PictureBox控件的Image属性，以在控件上进行显示。
				with_1.Image = bmp;
				try
				{
					//增加旋转的角度
					Rolling_ang += (int) this.RollingAngle;
				}
				catch (OverflowException)
				{
					//如果角度的值溢出，则将其重置为0
					Rolling_ang = 0;
				}
			}
			
			/// <summary>
			/// 停止图形的旋转
			/// </summary>
			/// <remarks></remarks>
			public void StopRolling()
			{
				if (this.Timer_ProcessRing != null)
				{
					this.Timer_ProcessRing.Stop();
					this.Timer_ProcessRing.Dispose();
					this.Timer_ProcessRing = null;
					this.Timer_ProcessRing.Tick += this.Rolling;
				}
			}
		}
		
	}
}
