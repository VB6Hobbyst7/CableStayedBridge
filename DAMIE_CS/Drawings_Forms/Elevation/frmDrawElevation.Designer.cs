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
	partial class frmDrawElevation : System.Windows.Forms.Form
	{
		
		//Form 重写 Dispose，以清理组件列表。
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
		
		//Windows 窗体设计器所必需的
		private System.ComponentModel.Container components = null;
		
		//注意: 以下过程是 Windows 窗体设计器所必需的
		//可以使用 Windows 窗体设计器修改它。
		//不要使用代码编辑器修改它。
		[System.Diagnostics.DebuggerStepThrough()]private void InitializeComponent()
		{
			this.components = new System.ComponentModel.Container();
			this.lbChooseExcav = new System.Windows.Forms.Label();
			this.lstbxChooseRegion = new System.Windows.Forms.ListBox();
			this.lstbxChooseRegion.SelectedIndexChanged += new System.EventHandler(this.RefreshSelectedRegion);
			this.btnGenerate = new System.Windows.Forms.Button();
			this.btnGenerate.Click += new System.EventHandler(this.btnGenerate_Click);
			this.OpenFileDlg_DataExcel = new System.Windows.Forms.OpenFileDialog();
			this.btnChooseAll = new System.Windows.Forms.Button();
			this.btnChooseAll.Click += new System.EventHandler(this.btnChooseAll_Click);
			this.btnChooseNone = new System.Windows.Forms.Button();
			this.btnChooseNone.Click += new System.EventHandler(this.btnChooseNone_Click);
			this.BGW_Generate = new System.ComponentModel.BackgroundWorker();
			this.BGW_Generate.DoWork += new System.ComponentModel.DoWorkEventHandler(this.BGW_Generate_DoWork);
			this.BGW_Generate.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.BGW_Generate_RunWorkerCompleted);
			this.ContextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
			this.ToolStripMenuItemRefresh = new System.Windows.Forms.ToolStripMenuItem();
			this.ContextMenuStrip1.SuspendLayout();
			this.SuspendLayout();
			//
			//lbChooseExcav
			//
			this.lbChooseExcav.AutoSize = true;
			this.lbChooseExcav.Location = new System.Drawing.Point(11, 12);
			this.lbChooseExcav.Name = "lbChooseExcav";
			this.lbChooseExcav.Size = new System.Drawing.Size(89, 12);
			this.lbChooseExcav.TabIndex = 2;
			this.lbChooseExcav.Text = "进行对比的区域";
			//
			//lstbxChooseRegion
			//
			this.lstbxChooseRegion.Anchor = (System.Windows.Forms.AnchorStyles) (((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right);
			this.lstbxChooseRegion.FormattingEnabled = true;
			this.lstbxChooseRegion.HorizontalScrollbar = true;
			this.lstbxChooseRegion.ItemHeight = 12;
			this.lstbxChooseRegion.Location = new System.Drawing.Point(13, 32);
			this.lstbxChooseRegion.Name = "lstbxChooseRegion";
			this.lstbxChooseRegion.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple;
			this.lstbxChooseRegion.Size = new System.Drawing.Size(237, 220);
			this.lstbxChooseRegion.TabIndex = 3;
			//
			//btnGenerate
			//
			this.btnGenerate.Anchor = (System.Windows.Forms.AnchorStyles) (System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right);
			this.btnGenerate.Location = new System.Drawing.Point(175, 264);
			this.btnGenerate.Name = "btnGenerate";
			this.btnGenerate.Size = new System.Drawing.Size(75, 25);
			this.btnGenerate.TabIndex = 4;
			this.btnGenerate.Text = "Generate";
			this.btnGenerate.UseVisualStyleBackColor = true;
			//
			//OpenFileDlg_DataExcel
			//
			this.OpenFileDlg_DataExcel.FileName = "OpenFileDialog1";
			//
			//btnChooseAll
			//
			this.btnChooseAll.Anchor = (System.Windows.Forms.AnchorStyles) (System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left);
			this.btnChooseAll.Location = new System.Drawing.Point(12, 264);
			this.btnChooseAll.Name = "btnChooseAll";
			this.btnChooseAll.Size = new System.Drawing.Size(60, 25);
			this.btnChooseAll.TabIndex = 5;
			this.btnChooseAll.Text = "全选";
			this.btnChooseAll.UseVisualStyleBackColor = true;
			//
			//btnChooseNone
			//
			this.btnChooseNone.Anchor = (System.Windows.Forms.AnchorStyles) (System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left);
			this.btnChooseNone.Location = new System.Drawing.Point(78, 264);
			this.btnChooseNone.Name = "btnChooseNone";
			this.btnChooseNone.Size = new System.Drawing.Size(60, 25);
			this.btnChooseNone.TabIndex = 5;
			this.btnChooseNone.Text = "清空";
			this.btnChooseNone.UseVisualStyleBackColor = true;
			//
			//BGW_Generate
			//
			this.BGW_Generate.WorkerReportsProgress = true;
			this.BGW_Generate.WorkerSupportsCancellation = true;
			//
			//ContextMenuStrip1
			//
			this.ContextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {this.ToolStripMenuItemRefresh});
			this.ContextMenuStrip1.Name = "ContextMenuStrip1";
			this.ContextMenuStrip1.Size = new System.Drawing.Size(117, 26);
			//
			//ToolStripMenuItemRefresh
			//
			this.ToolStripMenuItemRefresh.Name = "ToolStripMenuItemRefresh";
			this.ToolStripMenuItemRefresh.Size = new System.Drawing.Size(116, 22);
			this.ToolStripMenuItemRefresh.Text = "刷新(&R)";
			//
			//frmDrawElevation
			//
			this.AcceptButton = this.btnGenerate;
			this.AutoScaleDimensions = new System.Drawing.SizeF((float) (6.0F), (float) (12.0F));
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(262, 301);
			this.Controls.Add(this.btnChooseNone);
			this.Controls.Add(this.btnChooseAll);
			this.Controls.Add(this.btnGenerate);
			this.Controls.Add(this.lstbxChooseRegion);
			this.Controls.Add(this.lbChooseExcav);
			this.MinimumSize = new System.Drawing.Size(278, 339);
			this.Name = "frmDrawElevation";
			this.Text = "生成剖面图";
			this.ContextMenuStrip1.ResumeLayout(false);
			this.ResumeLayout(false);
			this.PerformLayout();
			
		}
		internal System.Windows.Forms.Label lbChooseExcav;
		internal System.Windows.Forms.ListBox lstbxChooseRegion;
		internal System.Windows.Forms.Button btnGenerate;
		internal System.Windows.Forms.OpenFileDialog OpenFileDlg_DataExcel;
		internal System.Windows.Forms.Button btnChooseAll;
		internal System.Windows.Forms.Button btnChooseNone;
		internal System.ComponentModel.BackgroundWorker BGW_Generate;
		internal System.Windows.Forms.ContextMenuStrip ContextMenuStrip1;
		internal System.Windows.Forms.ToolStripMenuItem ToolStripMenuItemRefresh;
	}
	
}
