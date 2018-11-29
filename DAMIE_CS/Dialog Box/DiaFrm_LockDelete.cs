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
	namespace Miscellaneous
	{
		
		public enum M_DialogResult
		{
			Lock,
			Delete,
			Cancel
		}
		
		
		public partial class DiaFrm_LockDelete
		{
			
#region   ---  Declarations & Definitions
			
#region   ---  Types
			
#endregion
			
#region   ---  Events
			
#endregion
			
#region   ---  Constants
			
#endregion
			
#region   ---  Properties
			
public string Prompt
			{
				set
				{
					this.LabelPrompt.Text = value;
				}
			}
public string Title
			{
				set
				{
					this.Text = value;
				}
			}
public string Text_Button1
			{
				set
				{
					this.btn1.Text = value;
				}
			}
public string Text_Button2
			{
				set
				{
					this.btn2.Text = value;
				}
			}
public string Text_Button3
			{
				set
				{
					this.btn3.Text = value;
				}
			}
			
#endregion
			
#region   ---  Fields
			
			/// <summary>
			/// 最后的返回值
			/// </summary>
			/// <remarks></remarks>
			private M_DialogResult result; //= M_DialogResult.ignore
			
#endregion
			
#endregion
			
#region   ---  构造函数与窗体的加载、打开与关闭
			public DiaFrm_LockDelete()
			{
				
				// This call is required by the designer.
				InitializeComponent();
				
				// Add any initialization after the InitializeComponent() call.
				this.TopMost = true;
				this.StartPosition = FormStartPosition.CenterParent;
			}
#endregion
			
			
			private delegate void ShowDialogHandler(string Prompt, string Title);
			/// <summary>
			///  打开对话框
			/// </summary>
			/// <param name="Prompt">提示</param>
			/// <param name="Title">窗口标题</param>
			/// <param name="Text_Button1">第一个按钮的文字</param>
			/// <param name="Text_Button2">第二个按钮的文字</param>
			/// <param name="Text_Button3">第三个按钮的文字</param>
			/// <returns>选择的操作方式：lock、delete、Cancel</returns>
			/// <remarks></remarks>
			public new M_DialogResult ShowDialog(string Prompt, string Title = "TIP", string Text_Button1 = "Lock", string Text_Button2 = "Delete", string Text_Button3 = "Cancel")
			{
				if (this.InvokeRequired)
				{
					//非UI线程，再次封送该方法到UI线程
					Debug.Print("非UI线程，再次封送该方法到UI线程");
					this.BeginInvoke(new ShowDialogHandler[this.ShowDialog], new[] {Prompt, Title});
				}
				else
				{
					DiaFrm_LockDelete with_1 = this;
					with_1.Title = Title;
					with_1.Prompt = Prompt;
					with_1.Text_Button1 = Text_Button1;
					with_1.Text_Button2 = Text_Button2;
					with_1.Text_Button3 = Text_Button3;
					//
					with_1.btn1.Focus();
					//
					base.ShowDialog();
					this.Dispose();
				}
				return result;
			}
			
#region   ---  界面操作
			//点击按钮
			public void btnLock_Click(object sender, EventArgs e)
			{
				result = Miscellaneous.M_DialogResult.Lock;
				this.Close();
			}
			public void btnDelete_Click(object sender, EventArgs e)
			{
				result = Miscellaneous.M_DialogResult.Delete;
				this.Close();
			}
			public void btnIgnore_Click(object sender, EventArgs e)
			{
				result = Miscellaneous.M_DialogResult.Cancel;
				this.Close();
			}
			
			//控制窗口的尺寸
			public void LabelPrompt_SizeChanged(object sender, EventArgs e)
			{
				this.Height = this.LabelPrompt.Height + 96;
			}
#endregion
		}
		
		
	}
}
