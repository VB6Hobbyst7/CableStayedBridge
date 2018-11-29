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


namespace CableStayedBridge
{
	namespace UI
	{
		public sealed partial class UI_BackGround
		{
			public UI_BackGround()
			{
				
				// This call is required by the designer.
				InitializeComponent();
				
				// Add any initialization after the InitializeComponent() call.
				
				//设置图片的相对透明的父控件与相对位置
				this.PictureBoxAME.Parent = this.PictureBoxBackGround;
				this.PictureBoxAME.Top = this.PictureBoxAME.Top - this.PictureBoxAME.Parent.Top;
				this.PictureBoxAME.Left = this.PictureBoxAME.Left - this.PictureBoxAME.Parent.Left;
				
				//设置背景图片的背景色
				this.BackColor = System.Drawing.Color.FromArgb(195, 195, 195);
			}
			
		}
		
	}
}
