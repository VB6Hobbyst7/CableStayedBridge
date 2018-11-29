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

//using eZstd.eZAPI.APIWindows;
using eZstd.eZAPI;

namespace CableStayedBridge
{
	
	public partial class Test_WindowProc
	{
		public Test_WindowProc()
		{
			InitializeComponent();
		}
		// VBConversions Note: Former VB static variables moved to class level because they aren't supported in C#.
		private decimal WndProc_i = 0;
		private decimal WndProc_j = 0;
		
		protected override void WndProc(ref System.Windows.Forms.Message m)
		{
			// static decimal i = 0; VBConversions Note: Static variable moved to class level and renamed WndProc_i. Local static variables are not supported in C#.
			WndProc_i += 0.01M;
			Debug.Print(System.Convert.ToString(WndProc_i));
			// 当点击窗口的“关闭”按钮时，将窗口最小化
			if (m.Msg == (int) WindowsMessages.WM_SYSCOMMAND & m.WParam == SysCommands.SC_CLOSE)
			{
				// static decimal j = 0; VBConversions Note: Static variable moved to class level and renamed WndProc_j. Local static variables are not supported in C#.
				WndProc_j += 0.01M;
				Debug.Print("进入" + System.Convert.ToString(WndProc_j));
				// 屏蔽传入的消息事件()
				this.DefWndProc(ref m);
				this.WindowState = FormWindowState.Minimized;
				return ;
			}
			base.WndProc(ref m);
		}
	}
}
