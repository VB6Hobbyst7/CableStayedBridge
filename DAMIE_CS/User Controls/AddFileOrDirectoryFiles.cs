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
	namespace AME_UserControl
	{
		/// <summary>
		/// 自定义控件：用来添加文件或者批量添加文件夹中的指定文件
		/// </summary>
		/// <remarks></remarks>
		public partial class AddFileOrDirectoryFiles
		{
			public AddFileOrDirectoryFiles()
			{
				InitializeComponent();
			}
			
			/// <summary>
			/// 添加文件
			/// </summary>
			/// <remarks></remarks>
			private EventHandler AddFileEvent;
			public event EventHandler AddFile
			{
				add
				{
					AddFileEvent = (EventHandler) System.Delegate.Combine(AddFileEvent, value);
				}
				remove
				{
					AddFileEvent = (EventHandler) System.Delegate.Remove(AddFileEvent, value);
				}
			}
			
			/// <summary>
			/// 批量添加文件夹中的指定文件
			/// </summary>
			/// <remarks></remarks>
			private EventHandler AddFilesFromDirectoryEvent;
			public event EventHandler AddFilesFromDirectory
			{
				add
				{
					AddFilesFromDirectoryEvent = (EventHandler) System.Delegate.Combine(AddFilesFromDirectoryEvent, value);
				}
				remove
				{
					AddFilesFromDirectoryEvent = (EventHandler) System.Delegate.Remove(AddFilesFromDirectoryEvent, value);
				}
			}
			
			
			/// <summary>
			/// 提供给外部调用，用来从外部隐藏“添加文件”与“添加文件夹”两个标签
			/// </summary>
			/// <remarks></remarks>
			public void HideLabel()
			{
				this.PanelAddFileOrDir.Visible = false;
			}
			
			/// <summary>
			/// 添加文件
			/// </summary>
			/// <param name="sender"></param>
			/// <param name="e"></param>
			/// <remarks></remarks>
			public void _AddFile(object sender, EventArgs e)
			{
				if (AddFileEvent != null)
					AddFileEvent(sender, e);
			}
			/// <summary>
			/// 添加文件夹中的文件
			/// </summary>
			/// <param name="sender"></param>
			/// <param name="e"></param>
			/// <remarks></remarks>
			public void _AddDire(object sender, EventArgs e)
			{
				if (AddFilesFromDirectoryEvent != null)
					AddFilesFromDirectoryEvent(sender, e);
			}
			
#region   ---  UI界面显示
			
			public void PanelAdd_MouseLeave(object sender, EventArgs e)
			{
				this.HideLabel();
			}
			
			public void btnAdd_MouseEnter(object sender, EventArgs e)
			{
				this.PanelAddFileOrDir.Visible = true;
			}
			public void colorFocused(object sender, EventArgs e)
			{
				System.Windows.Forms.Label lb = (System.Windows.Forms.Label) sender;
				lb.BackColor = Color.LightBlue;
			}
			public void colorLostFocus(object sender, EventArgs e)
			{
				System.Windows.Forms.Label lb = (System.Windows.Forms.Label) sender;
				lb.BackColor = Color.White;
			}
			
#endregion
		}
	}
}
