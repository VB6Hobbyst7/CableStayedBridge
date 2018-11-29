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
	partial class frmDrawing_Mnt_Others : System.Windows.Forms.Form
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
			this.components = new System.ComponentModel.Container();
			this.btnChooseMonitorData = new System.Windows.Forms.Button();
			this.btnChooseMonitorData.Click += new System.EventHandler(this.btnChooseMonitorData_Click);
			this.btnGenerate = new System.Windows.Forms.Button();
			this.btnGenerate.Click += new System.EventHandler(this.btnGenerate_Click);
			this.GroupBox1 = new System.Windows.Forms.GroupBox();
			this.RbtnStaticWithTime = new System.Windows.Forms.RadioButton();
			this.RbtnStaticWithTime.CheckedChanged += new System.EventHandler(this.RbtnStaticWithTime_CheckedChanged);
			this.RbtnDynamic = new System.Windows.Forms.RadioButton();
			this.RbtnDynamic.CheckedChanged += new System.EventHandler(this.RbtnStaticWithTime_CheckedChanged);
			this.chkBoxOpenNewExcel = new System.Windows.Forms.CheckBox();
			this.ListBoxPointsName = new System.Windows.Forms.ListBox();
			this.Label2 = new System.Windows.Forms.Label();
			this.listSheetsName = new System.Windows.Forms.ComboBox();
			this.listSheetsName.SelectedIndexChanged += new System.EventHandler(this.listSheetsName_SelectedIndexChanged);
			this.Label3 = new System.Windows.Forms.Label();
			this.ToolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.ComboBox_WorkingStage = new System.Windows.Forms.ComboBox();
			this.ComboBox_WorkingStage.SelectedIndexChanged += new System.EventHandler(this.ComboBox_WorkingStage_SelectedIndexChanged);
			this.btnDrawMonitorPoints = new System.Windows.Forms.Button();
			this.btnDrawMonitorPoints.Click += new System.EventHandler(this.btnDrawMonitorPoints_Click);
			this.BGWK_NewDrawing = new System.ComponentModel.BackgroundWorker();
			this.BGWK_NewDrawing.DoWork += new System.ComponentModel.DoWorkEventHandler(this.BGW_Generate_DoWork);
			this.BGWK_NewDrawing.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.BGW_Generate_RunWorkerCompleted);
			this.Label1 = new System.Windows.Forms.Label();
			this.ComboBox_MntType = new System.Windows.Forms.ComboBox();
			this.ComboBox_MntType.SelectedIndexChanged += new System.EventHandler(this.ComboBox_MntType_SelectedValueChanged);
			this.Panel_Static = new System.Windows.Forms.Panel();
			this.Label4 = new System.Windows.Forms.Label();
			this.ComboBoxOpenedWorkbook = new System.Windows.Forms.ComboBox();
			this.ComboBoxOpenedWorkbook.SelectedIndexChanged += new System.EventHandler(this.ComboBoxOpenedWorkbook_SelectedIndexChanged);
			this.GroupBox1.SuspendLayout();
			this.Panel_Static.SuspendLayout();
			this.SuspendLayout();
			//
			//btnChooseMonitorData
			//
			this.btnChooseMonitorData.Location = new System.Drawing.Point(13, 13);
			this.btnChooseMonitorData.Name = "btnChooseMonitorData";
			this.btnChooseMonitorData.Size = new System.Drawing.Size(75, 23);
			this.btnChooseMonitorData.TabIndex = 0;
			this.btnChooseMonitorData.Text = "监测数据";
			this.btnChooseMonitorData.UseVisualStyleBackColor = true;
			//
			//btnGenerate
			//
			this.btnGenerate.Anchor = (System.Windows.Forms.AnchorStyles) (System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right);
			this.btnGenerate.Location = new System.Drawing.Point(336, 355);
			this.btnGenerate.Name = "btnGenerate";
			this.btnGenerate.Size = new System.Drawing.Size(75, 23);
			this.btnGenerate.TabIndex = 4;
			this.btnGenerate.Text = "Generate";
			this.btnGenerate.UseVisualStyleBackColor = true;
			//
			//GroupBox1
			//
			this.GroupBox1.Anchor = (System.Windows.Forms.AnchorStyles) (System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right);
			this.GroupBox1.Controls.Add(this.RbtnStaticWithTime);
			this.GroupBox1.Controls.Add(this.RbtnDynamic);
			this.GroupBox1.Location = new System.Drawing.Point(312, 114);
			this.GroupBox1.Name = "GroupBox1";
			this.GroupBox1.Size = new System.Drawing.Size(99, 73);
			this.GroupBox1.TabIndex = 5;
			this.GroupBox1.TabStop = false;
			this.GroupBox1.Text = "曲线图的类型";
			//
			//RbtnStaticWithTime
			//
			this.RbtnStaticWithTime.AutoSize = true;
			this.RbtnStaticWithTime.Location = new System.Drawing.Point(7, 22);
			this.RbtnStaticWithTime.Name = "RbtnStaticWithTime";
			this.RbtnStaticWithTime.Size = new System.Drawing.Size(47, 16);
			this.RbtnStaticWithTime.TabIndex = 1;
			this.RbtnStaticWithTime.TabStop = true;
			this.RbtnStaticWithTime.Text = "静态";
			this.ToolTip1.SetToolTip(this.RbtnStaticWithTime, "以时间为X轴，查看每一个测点在整个施工过程中的变化");
			this.RbtnStaticWithTime.UseVisualStyleBackColor = true;
			//
			//RbtnDynamic
			//
			this.RbtnDynamic.AutoSize = true;
			this.RbtnDynamic.Location = new System.Drawing.Point(7, 45);
			this.RbtnDynamic.Name = "RbtnDynamic";
			this.RbtnDynamic.Size = new System.Drawing.Size(47, 16);
			this.RbtnDynamic.TabIndex = 0;
			this.RbtnDynamic.TabStop = true;
			this.RbtnDynamic.Text = "动态";
			this.ToolTip1.SetToolTip(this.RbtnDynamic, "以测点为X轴，动态查看每一天的变化情况");
			this.RbtnDynamic.UseVisualStyleBackColor = true;
			//
			//chkBoxOpenNewExcel
			//
			this.chkBoxOpenNewExcel.Anchor = (System.Windows.Forms.AnchorStyles) (System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right);
			this.chkBoxOpenNewExcel.AutoSize = true;
			this.chkBoxOpenNewExcel.Location = new System.Drawing.Point(321, 328);
			this.chkBoxOpenNewExcel.Name = "chkBoxOpenNewExcel";
			this.chkBoxOpenNewExcel.Size = new System.Drawing.Size(90, 16);
			this.chkBoxOpenNewExcel.TabIndex = 6;
			this.chkBoxOpenNewExcel.Text = "打开新Excel";
			this.chkBoxOpenNewExcel.UseVisualStyleBackColor = true;
			//
			//ListBoxPointsName
			//
			this.ListBoxPointsName.Anchor = (System.Windows.Forms.AnchorStyles) (((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right);
			this.ListBoxPointsName.FormattingEnabled = true;
			this.ListBoxPointsName.HorizontalScrollbar = true;
			this.ListBoxPointsName.ItemHeight = 12;
			this.ListBoxPointsName.Location = new System.Drawing.Point(106, 94);
			this.ListBoxPointsName.Name = "ListBoxPointsName";
			this.ListBoxPointsName.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
			this.ListBoxPointsName.Size = new System.Drawing.Size(185, 280);
			this.ListBoxPointsName.TabIndex = 7;
			//
			//Label2
			//
			this.Label2.AutoSize = true;
			this.Label2.Location = new System.Drawing.Point(11, 59);
			this.Label2.Name = "Label2";
			this.Label2.Size = new System.Drawing.Size(89, 12);
			this.Label2.TabIndex = 8;
			this.Label2.Text = "选择数据工作表";
			//
			//listSheetsName
			//
			this.listSheetsName.Anchor = (System.Windows.Forms.AnchorStyles) ((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right);
			this.listSheetsName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.listSheetsName.FormattingEnabled = true;
			this.listSheetsName.Location = new System.Drawing.Point(106, 56);
			this.listSheetsName.Name = "listSheetsName";
			this.listSheetsName.Size = new System.Drawing.Size(188, 20);
			this.listSheetsName.TabIndex = 9;
			//
			//Label3
			//
			this.Label3.AutoSize = true;
			this.Label3.Location = new System.Drawing.Point(11, 94);
			this.Label3.Name = "Label3";
			this.Label3.Size = new System.Drawing.Size(77, 12);
			this.Label3.TabIndex = 8;
			this.Label3.Text = "选择相应测点";
			//
			//ComboBox_WorkingStage
			//
			this.ComboBox_WorkingStage.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.ComboBox_WorkingStage.FormattingEnabled = true;
			this.ComboBox_WorkingStage.Location = new System.Drawing.Point(3, 20);
			this.ComboBox_WorkingStage.Name = "ComboBox_WorkingStage";
			this.ComboBox_WorkingStage.Size = new System.Drawing.Size(90, 20);
			this.ComboBox_WorkingStage.TabIndex = 1;
			this.ToolTip1.SetToolTip(this.ComboBox_WorkingStage, "用来在绘制测斜位移的最值走势图时，在图表中标开挖工况信息");
			//
			//btnDrawMonitorPoints
			//
			this.btnDrawMonitorPoints.Anchor = (System.Windows.Forms.AnchorStyles) (System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left);
			this.btnDrawMonitorPoints.Location = new System.Drawing.Point(12, 355);
			this.btnDrawMonitorPoints.Name = "btnDrawMonitorPoints";
			this.btnDrawMonitorPoints.Size = new System.Drawing.Size(75, 23);
			this.btnDrawMonitorPoints.TabIndex = 10;
			this.btnDrawMonitorPoints.Text = "绘制测点";
			this.btnDrawMonitorPoints.UseVisualStyleBackColor = true;
			//
			//BGWK_NewDrawing
			//
			this.BGWK_NewDrawing.WorkerReportsProgress = true;
			this.BGWK_NewDrawing.WorkerSupportsCancellation = true;
			//
			//Label1
			//
			this.Label1.Anchor = (System.Windows.Forms.AnchorStyles) (System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right);
			this.Label1.AutoSize = true;
			this.Label1.Location = new System.Drawing.Point(311, 56);
			this.Label1.Name = "Label1";
			this.Label1.Size = new System.Drawing.Size(77, 12);
			this.Label1.TabIndex = 12;
			this.Label1.Text = "监测数据类型";
			//
			//ComboBox_MntType
			//
			this.ComboBox_MntType.Anchor = (System.Windows.Forms.AnchorStyles) (System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right);
			this.ComboBox_MntType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.ComboBox_MntType.FormattingEnabled = true;
			this.ComboBox_MntType.IntegralHeight = false;
			this.ComboBox_MntType.Location = new System.Drawing.Point(313, 76);
			this.ComboBox_MntType.Name = "ComboBox_MntType";
			this.ComboBox_MntType.Size = new System.Drawing.Size(90, 20);
			this.ComboBox_MntType.TabIndex = 11;
			//
			//Panel_Static
			//
			this.Panel_Static.Anchor = (System.Windows.Forms.AnchorStyles) (System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right);
			this.Panel_Static.Controls.Add(this.Label4);
			this.Panel_Static.Controls.Add(this.ComboBox_WorkingStage);
			this.Panel_Static.Location = new System.Drawing.Point(312, 206);
			this.Panel_Static.Name = "Panel_Static";
			this.Panel_Static.Size = new System.Drawing.Size(99, 46);
			this.Panel_Static.TabIndex = 13;
			//
			//Label4
			//
			this.Label4.AutoSize = true;
			this.Label4.Location = new System.Drawing.Point(1, 0);
			this.Label4.Name = "Label4";
			this.Label4.Size = new System.Drawing.Size(53, 12);
			this.Label4.TabIndex = 3;
			this.Label4.Text = "开挖工况";
			//
			//ComboBoxOpenedWorkbook
			//
			this.ComboBoxOpenedWorkbook.Anchor = (System.Windows.Forms.AnchorStyles) ((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right);
			this.ComboBoxOpenedWorkbook.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.ComboBoxOpenedWorkbook.FormattingEnabled = true;
			this.ComboBoxOpenedWorkbook.Location = new System.Drawing.Point(95, 15);
			this.ComboBoxOpenedWorkbook.Name = "ComboBoxOpenedWorkbook";
			this.ComboBoxOpenedWorkbook.Size = new System.Drawing.Size(324, 20);
			this.ComboBoxOpenedWorkbook.TabIndex = 14;
			//
			//frmDrawing_Mnt_Others
			//
			this.AutoScaleDimensions = new System.Drawing.SizeF((float) (6.0F), (float) (12.0F));
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(431, 395);
			this.Controls.Add(this.ComboBoxOpenedWorkbook);
			this.Controls.Add(this.Panel_Static);
			this.Controls.Add(this.Label1);
			this.Controls.Add(this.ComboBox_MntType);
			this.Controls.Add(this.chkBoxOpenNewExcel);
			this.Controls.Add(this.btnGenerate);
			this.Controls.Add(this.btnDrawMonitorPoints);
			this.Controls.Add(this.listSheetsName);
			this.Controls.Add(this.Label3);
			this.Controls.Add(this.Label2);
			this.Controls.Add(this.ListBoxPointsName);
			this.Controls.Add(this.GroupBox1);
			this.Controls.Add(this.btnChooseMonitorData);
			this.MinimumSize = new System.Drawing.Size(340, 300);
			this.Name = "frmDrawing_Mnt_Others";
			this.Text = "其他监测曲线";
			this.GroupBox1.ResumeLayout(false);
			this.GroupBox1.PerformLayout();
			this.Panel_Static.ResumeLayout(false);
			this.Panel_Static.PerformLayout();
			this.ResumeLayout(false);
			this.PerformLayout();
			
		}
		internal System.Windows.Forms.Button btnChooseMonitorData;
		internal System.Windows.Forms.Button btnGenerate;
		internal System.Windows.Forms.GroupBox GroupBox1;
		internal System.Windows.Forms.RadioButton RbtnStaticWithTime;
		internal System.Windows.Forms.RadioButton RbtnDynamic;
		internal System.Windows.Forms.CheckBox chkBoxOpenNewExcel;
		internal System.Windows.Forms.ListBox ListBoxPointsName;
		internal System.Windows.Forms.Label Label2;
		internal System.Windows.Forms.ComboBox listSheetsName;
		internal System.Windows.Forms.Label Label3;
		internal System.Windows.Forms.ToolTip ToolTip1;
		internal System.Windows.Forms.Button btnDrawMonitorPoints;
		internal System.ComponentModel.BackgroundWorker BGWK_NewDrawing;
		internal System.Windows.Forms.Label Label1;
		internal System.Windows.Forms.ComboBox ComboBox_MntType;
		internal System.Windows.Forms.Panel Panel_Static;
		internal System.Windows.Forms.Label Label4;
		internal System.Windows.Forms.ComboBox ComboBox_WorkingStage;
		internal System.Windows.Forms.ComboBox ComboBoxOpenedWorkbook;
	}
	
}
