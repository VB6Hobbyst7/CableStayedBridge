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
	partial class Diafrm_Output_Word : System.Windows.Forms.Form
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
			this.GroupBox1 = new System.Windows.Forms.GroupBox();
			this.CheckBox_PlanView = new System.Windows.Forms.CheckBox();
			this.CheckBox_PlanView.CheckedChanged += new System.EventHandler(this.CheckBox_PlanView_CheckedChanged);
			this.ProgressBar_PlanView = new System.Windows.Forms.ProgressBar();
			this.ProgressBar_PlanView.Click += new System.EventHandler(this.ProgressBar_PlanView_Click);
			this.Label5 = new System.Windows.Forms.Label();
			this.CheckBox_SectionalView = new System.Windows.Forms.CheckBox();
			this.CheckBox_SectionalView.CheckedChanged += new System.EventHandler(this.CheckBox_SectionalView_CheckedChanged);
			this.ProgressBar_SectionalView = new System.Windows.Forms.ProgressBar();
			this.ProgressBar_SectionalView.Click += new System.EventHandler(this.ProgressBar_SectionalView_Click);
			this.Label3 = new System.Windows.Forms.Label();
			this.ListBoxMonitor_Static = new System.Windows.Forms.ListBox();
			this.ListBoxMonitor_Static.SelectedIndexChanged += new System.EventHandler(this.ListBoxMonitor_Dynamic_SelectedIndexChanged);
			this.ListBoxMonitor_Dynamic = new System.Windows.Forms.ListBox();
			this.ListBoxMonitor_Dynamic.SelectedIndexChanged += new System.EventHandler(this.ListBoxMonitor_Dynamic_SelectedIndexChanged);
			this.LabelDate = new System.Windows.Forms.Label();
			this.Label4 = new System.Windows.Forms.Label();
			this.btnExport = new System.Windows.Forms.Button();
			this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
			this.ChkBxSelect = new System.Windows.Forms.CheckBox();
			this.ChkBxSelect.Click += new System.EventHandler(this.ChkBxSelect_Click);
			this.BackgroundWorker1 = new System.ComponentModel.BackgroundWorker();
			this.BackgroundWorker1.DoWork += new System.ComponentModel.DoWorkEventHandler(this.BackgroundWorker1_DoWork);
			this.BackgroundWorker1.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.BackgroundWorker1_ProgressChanged);
			this.BackgroundWorker1.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.BackgroundWorker1_RunWorkerCompleted);
			this.GroupBox1.SuspendLayout();
			this.SuspendLayout();
			//
			//GroupBox1
			//
			this.GroupBox1.Anchor = (System.Windows.Forms.AnchorStyles) (((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right);
			this.GroupBox1.Controls.Add(this.CheckBox_PlanView);
			this.GroupBox1.Controls.Add(this.ProgressBar_PlanView);
			this.GroupBox1.Controls.Add(this.Label5);
			this.GroupBox1.Controls.Add(this.CheckBox_SectionalView);
			this.GroupBox1.Controls.Add(this.ProgressBar_SectionalView);
			this.GroupBox1.Controls.Add(this.Label3);
			this.GroupBox1.Controls.Add(this.ListBoxMonitor_Static);
			this.GroupBox1.Controls.Add(this.ListBoxMonitor_Dynamic);
			this.GroupBox1.Location = new System.Drawing.Point(12, 12);
			this.GroupBox1.Name = "GroupBox1";
			this.GroupBox1.Size = new System.Drawing.Size(454, 245);
			this.GroupBox1.TabIndex = 20;
			this.GroupBox1.TabStop = false;
			this.GroupBox1.Text = "选择要进行同步滚动和结果输出的对象";
			//
			//CheckBox_PlanView
			//
			this.CheckBox_PlanView.AutoSize = true;
			this.CheckBox_PlanView.Location = new System.Drawing.Point(9, 30);
			this.CheckBox_PlanView.Name = "CheckBox_PlanView";
			this.CheckBox_PlanView.Size = new System.Drawing.Size(84, 16);
			this.CheckBox_PlanView.TabIndex = 23;
			this.CheckBox_PlanView.Text = "开挖平面图";
			this.CheckBox_PlanView.UseVisualStyleBackColor = true;
			//
			//ProgressBar_PlanView
			//
			this.ProgressBar_PlanView.Location = new System.Drawing.Point(100, 29);
			this.ProgressBar_PlanView.Name = "ProgressBar_PlanView";
			this.ProgressBar_PlanView.Size = new System.Drawing.Size(100, 16);
			this.ProgressBar_PlanView.TabIndex = 24;
			//
			//Label5
			//
			this.Label5.Anchor = (System.Windows.Forms.AnchorStyles) (System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right);
			this.Label5.AutoSize = true;
			this.Label5.Location = new System.Drawing.Point(234, 59);
			this.Label5.Name = "Label5";
			this.Label5.Size = new System.Drawing.Size(95, 12);
			this.Label5.TabIndex = 18;
			this.Label5.Text = "监测曲线 - 静态";
			//
			//CheckBox_SectionalView
			//
			this.CheckBox_SectionalView.AutoSize = true;
			this.CheckBox_SectionalView.Location = new System.Drawing.Point(236, 30);
			this.CheckBox_SectionalView.Name = "CheckBox_SectionalView";
			this.CheckBox_SectionalView.Size = new System.Drawing.Size(84, 16);
			this.CheckBox_SectionalView.TabIndex = 23;
			this.CheckBox_SectionalView.Text = "开挖剖面图";
			this.CheckBox_SectionalView.UseVisualStyleBackColor = true;
			//
			//ProgressBar_SectionalView
			//
			this.ProgressBar_SectionalView.Location = new System.Drawing.Point(327, 29);
			this.ProgressBar_SectionalView.Name = "ProgressBar_SectionalView";
			this.ProgressBar_SectionalView.Size = new System.Drawing.Size(100, 16);
			this.ProgressBar_SectionalView.TabIndex = 24;
			//
			//Label3
			//
			this.Label3.AutoSize = true;
			this.Label3.Location = new System.Drawing.Point(7, 59);
			this.Label3.Name = "Label3";
			this.Label3.Size = new System.Drawing.Size(95, 12);
			this.Label3.TabIndex = 18;
			this.Label3.Text = "监测曲线 - 动态";
			//
			//ListBoxMonitor_Static
			//
			this.ListBoxMonitor_Static.Anchor = (System.Windows.Forms.AnchorStyles) ((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Right);
			this.ListBoxMonitor_Static.FormattingEnabled = true;
			this.ListBoxMonitor_Static.HorizontalScrollbar = true;
			this.ListBoxMonitor_Static.ItemHeight = 12;
			this.ListBoxMonitor_Static.Location = new System.Drawing.Point(236, 74);
			this.ListBoxMonitor_Static.Name = "ListBoxMonitor_Static";
			this.ListBoxMonitor_Static.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
			this.ListBoxMonitor_Static.Size = new System.Drawing.Size(191, 160);
			this.ListBoxMonitor_Static.TabIndex = 17;
			//
			//ListBoxMonitor_Dynamic
			//
			this.ListBoxMonitor_Dynamic.Anchor = (System.Windows.Forms.AnchorStyles) ((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left);
			this.ListBoxMonitor_Dynamic.FormattingEnabled = true;
			this.ListBoxMonitor_Dynamic.HorizontalScrollbar = true;
			this.ListBoxMonitor_Dynamic.ItemHeight = 12;
			this.ListBoxMonitor_Dynamic.Location = new System.Drawing.Point(9, 74);
			this.ListBoxMonitor_Dynamic.Name = "ListBoxMonitor_Dynamic";
			this.ListBoxMonitor_Dynamic.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
			this.ListBoxMonitor_Dynamic.Size = new System.Drawing.Size(191, 160);
			this.ListBoxMonitor_Dynamic.TabIndex = 17;
			//
			//LabelDate
			//
			this.LabelDate.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.LabelDate.Font = new System.Drawing.Font("Times New Roman", (float) (10.5F), System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, System.Convert.ToByte(0));
			this.LabelDate.Location = new System.Drawing.Point(69, 269);
			this.LabelDate.Name = "LabelDate";
			this.LabelDate.Size = new System.Drawing.Size(93, 21);
			this.LabelDate.TabIndex = 14;
			this.LabelDate.Text = "2014/09/28";
			this.LabelDate.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			//
			//Label4
			//
			this.Label4.AutoSize = true;
			this.Label4.Location = new System.Drawing.Point(10, 274);
			this.Label4.Name = "Label4";
			this.Label4.Size = new System.Drawing.Size(53, 12);
			this.Label4.TabIndex = 18;
			this.Label4.Text = "施工日期";
			//
			//btnExport
			//
			this.btnExport.Location = new System.Drawing.Point(391, 269);
			this.btnExport.Name = "btnExport";
			this.btnExport.Size = new System.Drawing.Size(75, 23);
			this.btnExport.TabIndex = 19;
			this.btnExport.Text = "结果输出";
			this.btnExport.UseVisualStyleBackColor = true;
			//
			//ChkBxSelect
			//
			this.ChkBxSelect.AutoSize = true;
			this.ChkBxSelect.Location = new System.Drawing.Point(183, 273);
			this.ChkBxSelect.Name = "ChkBxSelect";
			this.ChkBxSelect.Size = new System.Drawing.Size(138, 16);
			this.ChkBxSelect.TabIndex = 21;
			this.ChkBxSelect.Text = "Select/DeSelect All";
			this.ChkBxSelect.ThreeState = true;
			this.ChkBxSelect.UseVisualStyleBackColor = true;
			//
			//BackgroundWorker1
			//
			this.BackgroundWorker1.WorkerReportsProgress = true;
			this.BackgroundWorker1.WorkerSupportsCancellation = true;
			//
			//Diafrm_Output_Word
			//
			this.AutoScaleDimensions = new System.Drawing.SizeF((float) (6.0F), (float) (12.0F));
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(478, 301);
			this.Controls.Add(this.ChkBxSelect);
			this.Controls.Add(this.GroupBox1);
			this.Controls.Add(this.btnExport);
			this.Controls.Add(this.Label4);
			this.Controls.Add(this.LabelDate);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "Diafrm_Output_Word";
			this.Text = "输出";
			this.GroupBox1.ResumeLayout(false);
			this.GroupBox1.PerformLayout();
			this.ResumeLayout(false);
			this.PerformLayout();
			
		}
		internal System.Windows.Forms.GroupBox GroupBox1;
		internal System.Windows.Forms.Label Label5;
		internal System.Windows.Forms.Label Label3;
		internal System.Windows.Forms.ListBox ListBoxMonitor_Static;
		internal System.Windows.Forms.ListBox ListBoxMonitor_Dynamic;
		internal System.Windows.Forms.Label LabelDate;
		internal System.Windows.Forms.Label Label4;
		internal System.Windows.Forms.Button btnExport;
		internal System.Windows.Forms.CheckBox ChkBxSelect;
		internal System.ComponentModel.BackgroundWorker BackgroundWorker1;
		internal System.Windows.Forms.CheckBox CheckBox_PlanView;
		internal System.Windows.Forms.ProgressBar ProgressBar_PlanView;
		internal System.Windows.Forms.CheckBox CheckBox_SectionalView;
		internal System.Windows.Forms.ProgressBar ProgressBar_SectionalView;
	}
	
}
