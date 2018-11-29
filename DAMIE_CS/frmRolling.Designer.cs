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
using CableStayedBridge.AME_UserControl;
// End of VB project level imports

namespace CableStayedBridge
{
	
	[global::Microsoft.VisualBasic.CompilerServices.DesignerGenerated()]public 
	partial class frmRolling : System.Windows.Forms.Form
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
			this.Label4 = new System.Windows.Forms.Label();
			this.Load += new System.EventHandler(frmRolling_Load);
			this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(frmRolling_FormClosing);
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(frmRolling_KeyDown);
			this.Label3 = new System.Windows.Forms.Label();
			this.LabelDate = new System.Windows.Forms.Label();
			this.btnRefresh = new System.Windows.Forms.Button();
			this.btnRefresh.Click += new System.EventHandler(this.OnRollingDrawingsRefreshed);
			this.btnOutPut = new System.Windows.Forms.Button();
			this.btnOutPut.Click += new System.EventHandler(this.btnOutPut_Click);
			this.GroupBox1 = new System.Windows.Forms.GroupBox();
			this.btn_GroupHandle = new System.Windows.Forms.Button();
			this.btn_GroupHandle.Click += new System.EventHandler(this.btn_GroupHandle_Click);
			this.Panel2 = new System.Windows.Forms.Panel();
			this.CheckBox_PlanView = new System.Windows.Forms.CheckBox();
			this.CheckBox_PlanView.CheckedChanged += new System.EventHandler(this.CheckBox_PlanView_CheckedChanged);
			this.ProgressBar_PlanView = new System.Windows.Forms.ProgressBar();
			this.ProgressBar_PlanView.Click += new System.EventHandler(this.ProgressBar_PlanView_Click);
			this.CheckBox_SectionalView = new System.Windows.Forms.CheckBox();
			this.CheckBox_SectionalView.CheckedChanged += new System.EventHandler(this.CheckBox_SectionalView_CheckedChanged);
			this.ProgressBar_SectionalView = new System.Windows.Forms.ProgressBar();
			this.ProgressBar_SectionalView.Click += new System.EventHandler(this.ProgressBar_SectionalView_Click);
			this.Panel_Roll = new System.Windows.Forms.Panel();
			this.btnRoll = new System.Windows.Forms.Button();
			this.NumChanging = new UsrCtrl_NumberChanging();
			this.NumChanging.ValueAdded += new UsrCtrl_NumberChanging.ValueAddedEventHandler(this.NumChanging_ValueAdded);
			this.NumChanging.ValueMinused += new UsrCtrl_NumberChanging.ValueMinusedEventHandler(this.NumChanging_ValueMinused);
			this.Calendar_Construction = new System.Windows.Forms.MonthCalendar();
			this.Calendar_Construction.DateSelected += new System.Windows.Forms.DateRangeEventHandler(this.Calendar_Construction_DateSelected);
			this.Label5 = new System.Windows.Forms.Label();
			this.ListBoxMonitorData = new ListBox_NoReplyForKeyDown();
			this.ListBoxMonitorData.SelectedIndexChanged += new System.EventHandler(this.ListBoxMonitorData_SelectedIndexChanged);
			this.GroupBox1.SuspendLayout();
			this.Panel2.SuspendLayout();
			this.Panel_Roll.SuspendLayout();
			this.SuspendLayout();
			//
			//Label4
			//
			this.Label4.AutoSize = true;
			this.Label4.Location = new System.Drawing.Point(3, 44);
			this.Label4.Name = "Label4";
			this.Label4.Size = new System.Drawing.Size(53, 12);
			this.Label4.TabIndex = 18;
			this.Label4.Text = "施工日期";
			//
			//Label3
			//
			this.Label3.AutoSize = true;
			this.Label3.Location = new System.Drawing.Point(276, 21);
			this.Label3.Name = "Label3";
			this.Label3.Size = new System.Drawing.Size(77, 12);
			this.Label3.TabIndex = 18;
			this.Label3.Text = "监测数据曲线";
			//
			//LabelDate
			//
			this.LabelDate.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.LabelDate.Font = new System.Drawing.Font("Times New Roman", (float) (10.5F), System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, System.Convert.ToByte(0));
			this.LabelDate.Location = new System.Drawing.Point(156, 40);
			this.LabelDate.Name = "LabelDate";
			this.LabelDate.Size = new System.Drawing.Size(93, 21);
			this.LabelDate.TabIndex = 14;
			this.LabelDate.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			//
			//btnRefresh
			//
			this.btnRefresh.Anchor = (System.Windows.Forms.AnchorStyles) (System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right);
			this.btnRefresh.Location = new System.Drawing.Point(371, 316);
			this.btnRefresh.Name = "btnRefresh";
			this.btnRefresh.Size = new System.Drawing.Size(75, 23);
			this.btnRefresh.TabIndex = 19;
			this.btnRefresh.Text = "刷新(&R)";
			this.btnRefresh.UseVisualStyleBackColor = true;
			//
			//btnOutPut
			//
			this.btnOutPut.Anchor = (System.Windows.Forms.AnchorStyles) (System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right);
			this.btnOutPut.Location = new System.Drawing.Point(371, 345);
			this.btnOutPut.Name = "btnOutPut";
			this.btnOutPut.Size = new System.Drawing.Size(75, 23);
			this.btnOutPut.TabIndex = 20;
			this.btnOutPut.Text = "输出...";
			this.btnOutPut.UseVisualStyleBackColor = true;
			//
			//GroupBox1
			//
			this.GroupBox1.Anchor = (System.Windows.Forms.AnchorStyles) (((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right);
			this.GroupBox1.Controls.Add(this.btn_GroupHandle);
			this.GroupBox1.Controls.Add(this.Panel2);
			this.GroupBox1.Controls.Add(this.Panel_Roll);
			this.GroupBox1.Controls.Add(this.btnOutPut);
			this.GroupBox1.Controls.Add(this.btnRefresh);
			this.GroupBox1.Controls.Add(this.Label3);
			this.GroupBox1.Controls.Add(this.ListBoxMonitorData);
			this.GroupBox1.Location = new System.Drawing.Point(12, 12);
			this.GroupBox1.MinimumSize = new System.Drawing.Size(459, 380);
			this.GroupBox1.Name = "GroupBox1";
			this.GroupBox1.Size = new System.Drawing.Size(459, 380);
			this.GroupBox1.TabIndex = 19;
			this.GroupBox1.TabStop = false;
			this.GroupBox1.Text = "选择要进行同步滚动和结果输出的对象";
			//
			//btn_GroupHandle
			//
			this.btn_GroupHandle.Anchor = (System.Windows.Forms.AnchorStyles) (System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right);
			this.btn_GroupHandle.Location = new System.Drawing.Point(290, 345);
			this.btn_GroupHandle.Name = "btn_GroupHandle";
			this.btn_GroupHandle.Size = new System.Drawing.Size(75, 23);
			this.btn_GroupHandle.TabIndex = 26;
			this.btn_GroupHandle.Text = "批量操作";
			this.btn_GroupHandle.UseVisualStyleBackColor = true;
			//
			//Panel2
			//
			this.Panel2.Controls.Add(this.CheckBox_PlanView);
			this.Panel2.Controls.Add(this.ProgressBar_PlanView);
			this.Panel2.Controls.Add(this.CheckBox_SectionalView);
			this.Panel2.Controls.Add(this.ProgressBar_SectionalView);
			this.Panel2.Location = new System.Drawing.Point(12, 36);
			this.Panel2.Name = "Panel2";
			this.Panel2.Size = new System.Drawing.Size(258, 69);
			this.Panel2.TabIndex = 25;
			//
			//CheckBox_PlanView
			//
			this.CheckBox_PlanView.AutoSize = true;
			this.CheckBox_PlanView.Location = new System.Drawing.Point(6, 11);
			this.CheckBox_PlanView.Name = "CheckBox_PlanView";
			this.CheckBox_PlanView.Size = new System.Drawing.Size(84, 16);
			this.CheckBox_PlanView.TabIndex = 23;
			this.CheckBox_PlanView.Text = "开挖平面图";
			this.CheckBox_PlanView.UseVisualStyleBackColor = true;
			//
			//ProgressBar_PlanView
			//
			this.ProgressBar_PlanView.Location = new System.Drawing.Point(97, 10);
			this.ProgressBar_PlanView.Name = "ProgressBar_PlanView";
			this.ProgressBar_PlanView.Size = new System.Drawing.Size(143, 16);
			this.ProgressBar_PlanView.TabIndex = 24;
			//
			//CheckBox_SectionalView
			//
			this.CheckBox_SectionalView.AutoSize = true;
			this.CheckBox_SectionalView.Location = new System.Drawing.Point(6, 44);
			this.CheckBox_SectionalView.Name = "CheckBox_SectionalView";
			this.CheckBox_SectionalView.Size = new System.Drawing.Size(84, 16);
			this.CheckBox_SectionalView.TabIndex = 23;
			this.CheckBox_SectionalView.Text = "开挖剖面图";
			this.CheckBox_SectionalView.UseVisualStyleBackColor = true;
			//
			//ProgressBar_SectionalView
			//
			this.ProgressBar_SectionalView.Location = new System.Drawing.Point(97, 43);
			this.ProgressBar_SectionalView.Name = "ProgressBar_SectionalView";
			this.ProgressBar_SectionalView.Size = new System.Drawing.Size(143, 16);
			this.ProgressBar_SectionalView.TabIndex = 24;
			//
			//Panel_Roll
			//
			this.Panel_Roll.Controls.Add(this.btnRoll);
			this.Panel_Roll.Controls.Add(this.LabelDate);
			this.Panel_Roll.Controls.Add(this.NumChanging);
			this.Panel_Roll.Controls.Add(this.Calendar_Construction);
			this.Panel_Roll.Controls.Add(this.Label4);
			this.Panel_Roll.Controls.Add(this.Label5);
			this.Panel_Roll.Location = new System.Drawing.Point(12, 117);
			this.Panel_Roll.Name = "Panel_Roll";
			this.Panel_Roll.Size = new System.Drawing.Size(258, 251);
			this.Panel_Roll.TabIndex = 22;
			//
			//btnRoll
			//
			this.btnRoll.Location = new System.Drawing.Point(63, 36);
			this.btnRoll.Name = "btnRoll";
			this.btnRoll.Size = new System.Drawing.Size(75, 23);
			this.btnRoll.TabIndex = 22;
			this.btnRoll.Text = "滚动";
			this.btnRoll.UseVisualStyleBackColor = true;
			//
			//NumChanging
			//
			this.NumChanging.BackColor = System.Drawing.Color.Transparent;
			this.NumChanging.Location = new System.Drawing.Point(61, 7);
			this.NumChanging.Name = "NumChanging";
			this.NumChanging.Size = new System.Drawing.Size(190, 21);
			this.NumChanging.TabIndex = 21;
			this.NumChanging.unit = UsrCtrl_NumberChanging.YearMonthDay.Days;
			//
			//Calendar_Construction
			//
			this.Calendar_Construction.Location = new System.Drawing.Point(5, 65);
			this.Calendar_Construction.MaxDate = new DateTime(2014, 10, 5, 0, 0, 0, 0);
			this.Calendar_Construction.MinDate = new DateTime(2013, 1, 1, 0, 0, 0, 0);
			this.Calendar_Construction.Name = "Calendar_Construction";
			this.Calendar_Construction.ShowTodayCircle = false;
			this.Calendar_Construction.ShowWeekNumbers = true;
			this.Calendar_Construction.TabIndex = 13;
			//
			//Label5
			//
			this.Label5.AutoSize = true;
			this.Label5.Location = new System.Drawing.Point(3, 10);
			this.Label5.Name = "Label5";
			this.Label5.Size = new System.Drawing.Size(53, 12);
			this.Label5.TabIndex = 18;
			this.Label5.Text = "增减日期";
			//
			//ListBoxMonitorData
			//
			this.ListBoxMonitorData.Anchor = (System.Windows.Forms.AnchorStyles) (((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right);
			this.ListBoxMonitorData.FormattingEnabled = true;
			this.ListBoxMonitorData.HorizontalScrollbar = true;
			this.ListBoxMonitorData.ItemHeight = 12;
			this.ListBoxMonitorData.Location = new System.Drawing.Point(276, 39);
			this.ListBoxMonitorData.MinimumSize = new System.Drawing.Size(4, 184);
			this.ListBoxMonitorData.Name = "ListBoxMonitorData";
			this.ListBoxMonitorData.ParentControl = this.NumChanging;
			this.ListBoxMonitorData.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
			this.ListBoxMonitorData.Size = new System.Drawing.Size(170, 256);
			this.ListBoxMonitorData.TabIndex = 17;
			//
			//frmRolling
			//
			this.AutoScaleDimensions = new System.Drawing.SizeF(6, 12);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(483, 404);
			this.Controls.Add(this.GroupBox1);
			this.MinimumSize = new System.Drawing.Size(457, 432);
			this.Name = "frmRolling";
			this.Text = "动态同步控制";
			this.GroupBox1.ResumeLayout(false);
			this.GroupBox1.PerformLayout();
			this.Panel2.ResumeLayout(false);
			this.Panel2.PerformLayout();
			this.Panel_Roll.ResumeLayout(false);
			this.Panel_Roll.PerformLayout();
			this.ResumeLayout(false);
			
		}
		
		
		internal ListBox_NoReplyForKeyDown ListBoxMonitorData;
		
		internal System.Windows.Forms.Label Label4;
		internal System.Windows.Forms.Label Label3;
		internal System.Windows.Forms.Label LabelDate;
		internal System.Windows.Forms.Button btnRefresh;
		internal System.Windows.Forms.Button btnOutPut;
		internal System.Windows.Forms.GroupBox GroupBox1;
		internal System.Windows.Forms.MonthCalendar Calendar_Construction;
		internal UsrCtrl_NumberChanging NumChanging;
		internal System.Windows.Forms.Label Label5;
		internal System.Windows.Forms.Panel Panel2;
		internal System.Windows.Forms.CheckBox CheckBox_PlanView;
		internal System.Windows.Forms.ProgressBar ProgressBar_PlanView;
		internal System.Windows.Forms.CheckBox CheckBox_SectionalView;
		internal System.Windows.Forms.ProgressBar ProgressBar_SectionalView;
		internal System.Windows.Forms.Panel Panel_Roll;
		internal System.Windows.Forms.Button btn_GroupHandle;
		internal System.Windows.Forms.Button btnRoll;
	}
	
}
