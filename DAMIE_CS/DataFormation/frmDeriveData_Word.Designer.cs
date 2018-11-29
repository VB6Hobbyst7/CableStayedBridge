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
	partial class frmDeriveData_Word : frmDeriveData
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmDeriveData_Word));
			this.SuspendLayout();
			//
			//btnExport
			//
			//
			//frmDeriveData_Word
			//
			this.AutoScaleDimensions = new System.Drawing.SizeF((float) (6.0F), (float) (12.0F));
			this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(frmDeriveDataFromWord_FormClosing);
			BackgroundWorker1.DoWork += new System.ComponentModel.DoWorkEventHandler(StartToDoWork);
			BackgroundWorker1.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(BackgroundWorker1_RunWorkerCompleted);
			this.ClientSize = new System.Drawing.Size(539, 396);
			this.Icon = (System.Drawing.Icon) (resources.GetObject("$this.Icon"));
			this.Name = "frmDeriveData_Word";
			this.Text = "从Word中提取数据";
			this.ResumeLayout(false);
			this.PerformLayout();
			
		}
		
		/// <summary>
		/// 只有在程序运行时才能显示出来的界面更新效果
		/// </summary>
		private void InitializeComponent_ActivateAtRuntime()
		{
			
			System.Windows.Forms.DataGridViewTextBoxColumn PointName = default(System.Windows.Forms.DataGridViewTextBoxColumn);
			System.Windows.Forms.DataGridViewTextBoxColumn DataOffset = default(System.Windows.Forms.DataGridViewTextBoxColumn);
			System.Windows.Forms.DataGridViewComboBoxColumn SearchDirection = default(System.Windows.Forms.DataGridViewComboBoxColumn);
			
			
			PointName = new System.Windows.Forms.DataGridViewTextBoxColumn();
			DataOffset = new System.Windows.Forms.DataGridViewTextBoxColumn();
			SearchDirection = new System.Windows.Forms.DataGridViewComboBoxColumn();
			
			//PointName
			//
			PointName.HeaderText = "点位特征名";
			PointName.Name = "PointName";
			PointName.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
			PointName.ToolTipText = "特征名是包含于实际的监测点位的，比如：特征名CX会在Word文档中搜索包含有CX的所有测点，如TCX01。";
			//
			
			
			//SearchDirection
			//
			SearchDirection.HeaderText = "搜索";
			SearchDirection.Items.AddRange(new object[] {"按行", "按列"});
			SearchDirection.Name = "SearchDirection";
			//
			
			//DataOffset
			//
			System.Windows.Forms.DataGridViewCellStyle DataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
			DataGridViewCellStyle2.ForeColor = System.Drawing.Color.Blue;
			DataOffset.DefaultCellStyle = DataGridViewCellStyle2;
			DataOffset.HeaderText = "数据偏移";
			DataOffset.Name = "DataOffset";
			DataOffset.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
			DataOffset.ToolTipText = "如果点位的数据在点位单元格的右侧且与之相邻，则为1";
			DataOffset.Width = 80;
			
			
			//MyDataGridView1
			//
			System.Windows.Forms.DataGridViewCellStyle DataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
			DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
			DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
			DataGridViewCellStyle1.Font = new System.Drawing.Font("SimSun", (float) (9.0F), System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, System.Convert.ToByte(134));
			DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
			DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
			DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
			DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
			this.MyDataGridView1.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1;
			this.MyDataGridView1.ColumnHeadersHeight = 25;
			this.MyDataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
			this.MyDataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {PointName, DataOffset, SearchDirection});
			this.MyDataGridView1.Location = new System.Drawing.Point(11, 181);
			this.MyDataGridView1.RowTemplate.Height = 23;
			this.MyDataGridView1.RowTemplate.Resizable = System.Windows.Forms.DataGridViewTriState.False;
			this.MyDataGridView1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
			this.MyDataGridView1.Size = new System.Drawing.Size(346, 110);
			this.MyDataGridView1.TabIndex = 14;
		}
		
		
		
	}
	
}
