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
	public partial class frmRegexDate
	{
		
#region   ---  Declarations & Definitions
		
#region   ---  Types
		
#endregion
		
#region   ---  Constants
		
#endregion
		
#region   ---  Properties
		
#endregion
		
#region   ---  Fields
		/// <summary>
		/// 搜索日期的正则表达式字符串
		/// </summary>
		/// <remarks></remarks>
		private string F_Pattern;
		/// <summary>
		/// 要提取的字符中，{文件序号，年，月，日}分别在Match.Groups集合中的下标值。用值0来代表没有此项。
		/// </summary>
		/// <remarks>Match.Groups(0)返回的是Match结果本身，并不属于要提取的数据。</remarks>
		private byte[] F_Components = new byte[4];
		
#endregion
		
#endregion
		
#region   ---  构造函数与窗体的加载、打开与关闭
		public frmRegexDate()
		{
			
			// This call is required by the designer.
			InitializeComponent();
			
			// Add any initialization after the InitializeComponent() call.
			this.Button1.Tag = 1;
			this.Button2.Tag = 2;
			this.Button3.Tag = 3;
			this.Button4.Tag = 4;
			
		}
		
		/// <summary>
		/// 弹出窗口，开始执行操作
		/// </summary>
		/// <param name="pattern"></param>
		/// <param name="Components"></param>
		/// <remarks></remarks>
		public void ShowDialog(ref string pattern, ref byte[] Components)
		{
			var result = base.ShowDialog();
			if (result == Windows.Forms.DialogResult.OK)
			{
				Components = this.F_Components;
				pattern = this.F_Pattern;
			}
		}
		
		public void Button6_Click(object sender, EventArgs e)
		{
			this.Close();
		}
#endregion
		
#region   ---  界面操作
		/// <summary>
		/// 通过点击按钮来设置对应的数据类型
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		/// <remarks></remarks>
		public void SetComponent(object sender, EventArgs e)
		{
			Button bt = (Button) sender;
			if ((int) bt.Tag == 1)
			{
				bt.Tag = 2;
				bt.Text = "年";
			}
			else if ((int) bt.Tag == 2)
			{
				bt.Tag = 3;
				bt.Text = "月";
			}
			else if ((int) bt.Tag == 3)
			{
				bt.Tag = 4;
				bt.Text = "日";
			}
			else if ((int) bt.Tag == 4)
			{
				bt.Tag = 1;
				bt.Text = "序号";
			}
		}
		
		// VBConversions Note: Former VB static variables moved to class level because they aren't supported in C#.
		private long selectText_i = 0;
		
		public void selectText(object sender, EventArgs e)
		{
			// static long i = 0; VBConversions Note: Static variable moved to class level and renamed selectText_i. Local static variables are not supported in C#.
			selectText_i++;
			TextBox tb = (TextBox) sender;
			tb.SelectAll();
		}
		
		/// <summary>
		/// 实时刷新正则表达式
		/// </summary>
		/// <remarks></remarks>
		public void GenerateRegex(System.Object sender, System.EventArgs e)
		{
			//下面格式字符串中的0~8分别代表：  前缀字符、              序号数字的个数、
			//                               序号与年的分隔字符、    表示年的数值的数字个数、
			//                               年与月的分隔字符、      表示月的数值的数字个数、
			//                               月与日的分隔字符、      表示日的数值的数字个数、
			//                               后缀字符()
			string strRegexp = string.Format("{0}\\s*(\\d{{0,{1}}})" +
				"\\s*{2}\\s*(\\d{{{3}}})" +
				"\\s*{4}\\s*(\\d{{{5}}})" +
				"\\s*{6}\\s*(\\d{{{7}}})\\s*{8}",
				TextBox1.Text, Textbox_btn1.Text, TextBox2.Text, Textbox_btn2.Text, TextBox3.Text, Textbox_btn3.Text, TextBox4.Text, Textbox_btn4.Text, TextBox5.Text);
			this.F_Pattern = strRegexp;
			this.Label_Regex.Text = strRegexp;
		}
		
#endregion
		
		/// <summary>
		/// 执行操作，确认最终的正则表达式
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		/// <remarks></remarks>
		public void BtnOk_Click(object sender, EventArgs e)
		{
			//要提取的字符中，{文件序号，年，月，日}分别在Match.Groups集合中的下标值。用值0来代表没有此项。
			byte[] component = new byte[4];
			byte[] btnTag = new byte[4]; //每一个按钮所代表的数据类型，其Tag值为1~4
			// button.Tag指的是每一个button所代表的数据类型
			btnTag = new[] {Button1.Tag, Button2.Tag, Button3.Tag, Button4.Tag};
			for (byte i = 0; i <= 3; i++)
			{
				component[i] = btnTag[i];
			}
			// 查看数组中是否有相同的元素，而不是有且只有"1、2、3、4"这四个元素
			for (SByte i = 0; i <= 2; i++)
			{
				for (SByte j = i + 1; j <= 3; j++)
				{
					if (component[i] == component[j])
					{
						return ;
					}
				}
			}
			//验证成功，所有的数据合法
			this.F_Components = component;
			this.DialogResult = Windows.Forms.DialogResult.OK;
			this.Close(); //这一句是必须的，用以返回showDialog方法。
		}
	}
}
