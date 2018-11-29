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

using CableStayedBridge.Miscellaneous;

namespace CableStayedBridge
{
	public class TemplateCodes : System.Windows.Forms.Form
	{
		public TemplateCodes()
		{
			InitializeComponent();
		}
		
		private void TemplateCodes_Load(object sender, EventArgs e)
		{
			int i = 0;
			foreach (int tempLoopVar_i in test())
			{
				i = tempLoopVar_i;
				Debug.Print(System.Convert.ToString(i));
			}
		}
		// VBConversions Note: Former VB static variables moved to class level because they aren't supported in C#.
		private int test_i = 0;
		
		public int[] test()
		{
			// static int i = 0; VBConversions Note: Static variable moved to class level and renamed test_i. Local static variables are not supported in C#.
			test_i++;
			return new[] {test_i, test_i + 1, test_i + 2};
			
		}
		private void ModelCode()
		{
			
			// -------- Tye 语句 --------------------
			try
			{
			}
			catch (Exception ex)
			{
				MessageBox.Show("" + "\r\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				
				
				MessageBox.Show("" + "\r\n" + ex.Message + "\r\n" + "报错位置：" +
					ex.TargetSite.Name, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				
			}
			
			
			
		}
		
		
	}
	
	
	
}
