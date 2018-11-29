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
using eZstd.eZAPI;


namespace CableStayedBridge
{
	namespace AME_UserControl
	{
		
		/// <summary>
		/// 列表框控件，但是屏蔽了对于列表框的键盘事件。
		/// 除非按下的键与列表框中的元素的第一个字符相同，否则此键按下不生效。
		/// </summary>
		/// <remarks></remarks>
		internal class ListBox_NoReplyForKeyDown : System.Windows.Forms.ListBox
		{
			
#region   ---  Properties
			
			/// <summary>
			/// 当此列表框拥有焦点，并响应键盘按下的事件时，自动将焦点转移到此属性所指定的控件上，并向其发送此键盘消息。
			/// </summary>
			/// <value></value>
			/// <returns></returns>
			/// <remarks></remarks>
			[Browsable(true), Description("当此列表框拥有焦点，并响应键盘按下的事件时，自动将焦点转移到此属性所指定的控件上，并向其发送此键盘消息。")]public Control ParentControl {get; set;}
			
#endregion
			
			protected override void WndProc(ref Message m)
			{
				if (m.Msg == (int) WindowsMessages.WM_KEYDOWN)
				{
					if (ParentControl != null)
					{
						ParentControl.Focus();
						APIMessage.SendMessage(ParentControl.FindForm().Handle, 
							m.Msg, m.WParam, m.LParam);
					}
				}
				else
				{
					base.WndProc(ref m);
				}
				
			}
			
		}
	}
}
