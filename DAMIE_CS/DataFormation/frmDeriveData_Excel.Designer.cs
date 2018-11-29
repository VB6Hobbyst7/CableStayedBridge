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
	[global::Microsoft.VisualBasic.CompilerServices.DesignerGenerated()]public 
	partial class frmDeriveData_Excel : frmDeriveData
	{
		//Form overrides dispose to clean up the component list.
		[System.Diagnostics.DebuggerNonUserCode()]protected override void Dispose(bool disposing)
		{
			try
			{
				if (disposing && components != null)
				{
					components.Dispose();
				}
			}
			finally
			{
				base.Dispose(disposing);
			}
		}
		
		//Required by the Windows Form Designer
		private System.ComponentModel.Container components = null;
		
		//NOTE: The following procedure is required by the Windows Form Designer
		//It can be modified using the Windows Form Designer.
		//Do not modify it using the code editor.
		[System.Diagnostics.DebuggerStepThrough()]private void InitializeComponent()
		{
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmDeriveData_Excel));
			this.SuspendLayout();
			//
			//btnExport
			//
			//
			//frmDeriveData_Excel
			//
			this.AutoScaleDimensions = new System.Drawing.SizeF((float) (6.0F), (float) (12.0F));
			BackgroundWorker1.DoWork += new System.ComponentModel.DoWorkEventHandler(BackgroundWorker1_DoWork);
			this.ClientSize = new System.Drawing.Size(539, 396);
			this.Icon = (System.Drawing.Icon) (resources.GetObject("$this.Icon"));
			this.Name = "frmDeriveData_Excel";
			this.Text = "从Excel中提取数据";
			this.ResumeLayout(false);
			this.PerformLayout();
			
		}
		
		/// <summary>
		/// 只有在程序运行时才能显示出来的界面更新效果
		/// </summary>
		private void InitializeComponent_ActivateAtRuntime()
		{
			var SheetName = new System.Windows.Forms.DataGridViewTextBoxColumn();
			var RangeName = new System.Windows.Forms.DataGridViewTextBoxColumn();
			//SheetName
			SheetName.HeaderText = "工作表名称";
			SheetName.Name = "SheetName";
			SheetName.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
			SheetName.ToolTipText = "提取的工作表名称包含于要进行检索的工作表名称，比如输入\"CX\"，则会提取工作簿中第一个名称中含有\"CX\"的工作表。" + "\r\n" +
				"每一个工作表名称都会在用来保存数据的工作簿中创建一个对应的工作表。";
			SheetName.Width = 183;
			//
			//RangeName
			System.Windows.Forms.DataGridViewCellStyle DataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
			DataGridViewCellStyle1.ForeColor = System.Drawing.Color.Blue;
			RangeName.DefaultCellStyle = DataGridViewCellStyle1;
			RangeName.HeaderText = "区域范围";
			RangeName.Name = "RangeName";
			RangeName.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
			RangeName.ToolTipText = "示例： A1:C3,如果要引用一张表中不连续的两个区域，可以使用\"A1:A3,C1:C3\"";
			//
			System.Windows.Forms.DataGridViewCellStyle DataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
			DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
			DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Control;
			DataGridViewCellStyle2.Font = new System.Drawing.Font("SimSun", (float) (9.0F), System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, System.Convert.ToByte(134));
			DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText;
			DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
			DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
			DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
			this.MyDataGridView1.RowTemplate.Height = 23;
			this.MyDataGridView1.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle2;
			this.MyDataGridView1.ColumnHeadersHeight = 25;
			this.MyDataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
			this.MyDataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {SheetName, RangeName});
			
		}
		
	}
	
}
