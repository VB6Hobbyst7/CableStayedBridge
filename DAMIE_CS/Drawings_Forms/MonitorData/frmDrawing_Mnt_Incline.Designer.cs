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
	partial class frmDrawing_Mnt_Incline : System.Windows.Forms.Form
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
			this.ComboBox_ExcavID = new System.Windows.Forms.ComboBox();
			this.ComboBox_ExcavID.SelectedIndexChanged += new System.EventHandler(this.CbBoxExcavID_SelectedIndexChanged);
			this.Label_Component_Elevation = new System.Windows.Forms.Label();
			this.btnGenerate = new System.Windows.Forms.Button();
			this.btnGenerate.Click += new System.EventHandler(this.btnGenerate_Click);
			this.chkBoxOpenNewExcel = new System.Windows.Forms.CheckBox();
			this.ListBoxWorksheetsName = new System.Windows.Forms.ListBox();
			this.ListBoxWorksheetsName.SelectedIndexChanged += new System.EventHandler(this.ListBoxWorksheetsName_SelectedIndexChanged);
			this.ComboBox_ExcavRegion = new System.Windows.Forms.ComboBox();
			this.ComboBox_ExcavRegion.SelectedIndexChanged += new System.EventHandler(this.CbBoxExcavRegion_SelectedIndexChanged);
			this.Label2 = new System.Windows.Forms.Label();
			this.btnDrawMonitorPoints = new System.Windows.Forms.Button();
			this.btnDrawMonitorPoints.Click += new System.EventHandler(this.btnDrawMonitorPoints_Click);
			this.BGWK_NewDrawing = new System.ComponentModel.BackgroundWorker();
			this.BGWK_NewDrawing.DoWork += new System.ComponentModel.DoWorkEventHandler(this.BGW_Generate_DoWork);
			this.BGWK_NewDrawing.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.BGW_Generate_RunWorkerCompleted);
			this.GroupBox1 = new System.Windows.Forms.GroupBox();
			this.RadioButton_Dynamic = new System.Windows.Forms.RadioButton();
			this.RadioButton_Dynamic.CheckedChanged += new System.EventHandler(this.RadioButton_Dynamic_CheckedChanged);
			this.RadioButton_Max_Depth = new System.Windows.Forms.RadioButton();
			this.RadioButton_Max_Depth.CheckedChanged += new System.EventHandler(this.RadioButton_Dynamic_CheckedChanged);
			this.Panel_Dynamic = new System.Windows.Forms.Panel();
			this.ComboBox_MntType = new System.Windows.Forms.ComboBox();
			this.ComboBox_MntType.SelectedValueChanged += new System.EventHandler(this.ComboBox_MntType_SelectedValueChanged);
			this.Label_MntType = new System.Windows.Forms.Label();
			this.ComboBox_WorkingStage = new System.Windows.Forms.ComboBox();
			this.ComboBox_WorkingStage.SelectedIndexChanged += new System.EventHandler(this.ComboBox_WorkingStage_SelectedIndexChanged);
			this.Label3 = new System.Windows.Forms.Label();
			this.Panel_Static = new System.Windows.Forms.Panel();
			this.ToolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.ComboBoxOpenedWorkbook = new System.Windows.Forms.ComboBox();
			this.ComboBoxOpenedWorkbook.SelectedIndexChanged += new System.EventHandler(this.ComboBoxOpenedWorkbook_SelectedIndexChanged);
			this.CheckBox1 = new System.Windows.Forms.CheckBox();
			this.CheckBox1.CheckedChanged += new System.EventHandler(this.CheckBox1_CheckedChanged);
			this.ComboBox1 = new System.Windows.Forms.ComboBox();
			this.GroupBox1.SuspendLayout();
			this.Panel_Dynamic.SuspendLayout();
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
			//ComboBox_ExcavID
			//
			this.ComboBox_ExcavID.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.ComboBox_ExcavID.FormattingEnabled = true;
			this.ComboBox_ExcavID.Location = new System.Drawing.Point(5, 38);
			this.ComboBox_ExcavID.Name = "ComboBox_ExcavID";
			this.ComboBox_ExcavID.Size = new System.Drawing.Size(90, 20);
			this.ComboBox_ExcavID.TabIndex = 1;
			//
			//Label_Component_Elevation
			//
			this.Label_Component_Elevation.AutoSize = true;
			this.Label_Component_Elevation.Location = new System.Drawing.Point(3, 18);
			this.Label_Component_Elevation.Name = "Label_Component_Elevation";
			this.Label_Component_Elevation.Size = new System.Drawing.Size(59, 12);
			this.Label_Component_Elevation.TabIndex = 3;
			this.Label_Component_Elevation.Text = "构件-标高";
			//
			//btnGenerate
			//
			this.btnGenerate.Anchor = (System.Windows.Forms.AnchorStyles) (System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right);
			this.btnGenerate.Location = new System.Drawing.Point(354, 290);
			this.btnGenerate.Name = "btnGenerate";
			this.btnGenerate.Size = new System.Drawing.Size(75, 23);
			this.btnGenerate.TabIndex = 4;
			this.btnGenerate.Text = "生成";
			this.btnGenerate.UseVisualStyleBackColor = true;
			//
			//chkBoxOpenNewExcel
			//
			this.chkBoxOpenNewExcel.Anchor = (System.Windows.Forms.AnchorStyles) (System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right);
			this.chkBoxOpenNewExcel.AutoSize = true;
			this.chkBoxOpenNewExcel.Location = new System.Drawing.Point(319, 69);
			this.chkBoxOpenNewExcel.Name = "chkBoxOpenNewExcel";
			this.chkBoxOpenNewExcel.Size = new System.Drawing.Size(90, 16);
			this.chkBoxOpenNewExcel.TabIndex = 6;
			this.chkBoxOpenNewExcel.Text = "打开新Excel";
			this.chkBoxOpenNewExcel.UseVisualStyleBackColor = true;
			//
			//ListBoxWorksheetsName
			//
			this.ListBoxWorksheetsName.Anchor = (System.Windows.Forms.AnchorStyles) (((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right);
			this.ListBoxWorksheetsName.FormattingEnabled = true;
			this.ListBoxWorksheetsName.ItemHeight = 12;
			this.ListBoxWorksheetsName.Location = new System.Drawing.Point(12, 51);
			this.ListBoxWorksheetsName.Name = "ListBoxWorksheetsName";
			this.ListBoxWorksheetsName.Size = new System.Drawing.Size(161, 256);
			this.ListBoxWorksheetsName.TabIndex = 7;
			//
			//ComboBox_ExcavRegion
			//
			this.ComboBox_ExcavRegion.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.ComboBox_ExcavRegion.FormattingEnabled = true;
			this.ComboBox_ExcavRegion.Location = new System.Drawing.Point(5, 141);
			this.ComboBox_ExcavRegion.Name = "ComboBox_ExcavRegion";
			this.ComboBox_ExcavRegion.Size = new System.Drawing.Size(90, 20);
			this.ComboBox_ExcavRegion.TabIndex = 1;
			//
			//Label2
			//
			this.Label2.AutoSize = true;
			this.Label2.Location = new System.Drawing.Point(3, 119);
			this.Label2.Name = "Label2";
			this.Label2.Size = new System.Drawing.Size(53, 12);
			this.Label2.TabIndex = 3;
			this.Label2.Text = "施工进度";
			//
			//btnDrawMonitorPoints
			//
			this.btnDrawMonitorPoints.Anchor = (System.Windows.Forms.AnchorStyles) (System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right);
			this.btnDrawMonitorPoints.Location = new System.Drawing.Point(354, 261);
			this.btnDrawMonitorPoints.Name = "btnDrawMonitorPoints";
			this.btnDrawMonitorPoints.Size = new System.Drawing.Size(75, 23);
			this.btnDrawMonitorPoints.TabIndex = 8;
			this.btnDrawMonitorPoints.Text = "绘制测点";
			this.btnDrawMonitorPoints.UseVisualStyleBackColor = true;
			//
			//BGWK_NewDrawing
			//
			this.BGWK_NewDrawing.WorkerReportsProgress = true;
			this.BGWK_NewDrawing.WorkerSupportsCancellation = true;
			//
			//GroupBox1
			//
			this.GroupBox1.Anchor = (System.Windows.Forms.AnchorStyles) (System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right);
			this.GroupBox1.Controls.Add(this.ComboBox1);
			this.GroupBox1.Controls.Add(this.CheckBox1);
			this.GroupBox1.Controls.Add(this.RadioButton_Dynamic);
			this.GroupBox1.Controls.Add(this.RadioButton_Max_Depth);
			this.GroupBox1.Location = new System.Drawing.Point(313, 119);
			this.GroupBox1.Name = "GroupBox1";
			this.GroupBox1.Size = new System.Drawing.Size(117, 131);
			this.GroupBox1.TabIndex = 9;
			this.GroupBox1.TabStop = false;
			this.GroupBox1.Text = "绘图类型";
			//
			//RadioButton_Dynamic
			//
			this.RadioButton_Dynamic.AutoSize = true;
			this.RadioButton_Dynamic.Checked = true;
			this.RadioButton_Dynamic.Location = new System.Drawing.Point(6, 46);
			this.RadioButton_Dynamic.Name = "RadioButton_Dynamic";
			this.RadioButton_Dynamic.Size = new System.Drawing.Size(83, 16);
			this.RadioButton_Dynamic.TabIndex = 0;
			this.RadioButton_Dynamic.TabStop = true;
			this.RadioButton_Dynamic.Text = "形状动态图";
			this.RadioButton_Dynamic.UseVisualStyleBackColor = true;
			//
			//RadioButton_Max_Depth
			//
			this.RadioButton_Max_Depth.AutoSize = true;
			this.RadioButton_Max_Depth.Location = new System.Drawing.Point(7, 21);
			this.RadioButton_Max_Depth.Name = "RadioButton_Max_Depth";
			this.RadioButton_Max_Depth.Size = new System.Drawing.Size(83, 16);
			this.RadioButton_Max_Depth.TabIndex = 0;
			this.RadioButton_Max_Depth.Text = "最值走势图";
			this.RadioButton_Max_Depth.UseVisualStyleBackColor = true;
			//
			//Panel_Dynamic
			//
			this.Panel_Dynamic.Anchor = (System.Windows.Forms.AnchorStyles) (System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right);
			this.Panel_Dynamic.Controls.Add(this.Label_Component_Elevation);
			this.Panel_Dynamic.Controls.Add(this.ComboBox_ExcavID);
			this.Panel_Dynamic.Controls.Add(this.ComboBox_ExcavRegion);
			this.Panel_Dynamic.Controls.Add(this.Label2);
			this.Panel_Dynamic.Location = new System.Drawing.Point(186, 101);
			this.Panel_Dynamic.Name = "Panel_Dynamic";
			this.Panel_Dynamic.Size = new System.Drawing.Size(106, 194);
			this.Panel_Dynamic.TabIndex = 10;
			//
			//ComboBox_MntType
			//
			this.ComboBox_MntType.Anchor = (System.Windows.Forms.AnchorStyles) (System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right);
			this.ComboBox_MntType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.ComboBox_MntType.DropDownWidth = 90;
			this.ComboBox_MntType.FormattingEnabled = true;
			this.ComboBox_MntType.Location = new System.Drawing.Point(191, 71);
			this.ComboBox_MntType.Name = "ComboBox_MntType";
			this.ComboBox_MntType.Size = new System.Drawing.Size(90, 20);
			this.ComboBox_MntType.TabIndex = 1;
			//
			//Label_MntType
			//
			this.Label_MntType.Anchor = (System.Windows.Forms.AnchorStyles) (System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right);
			this.Label_MntType.AutoSize = true;
			this.Label_MntType.Location = new System.Drawing.Point(189, 51);
			this.Label_MntType.Name = "Label_MntType";
			this.Label_MntType.Size = new System.Drawing.Size(77, 12);
			this.Label_MntType.TabIndex = 3;
			this.Label_MntType.Text = "监测数据类型";
			//
			//ComboBox_WorkingStage
			//
			this.ComboBox_WorkingStage.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.ComboBox_WorkingStage.DropDownWidth = 90;
			this.ComboBox_WorkingStage.FormattingEnabled = true;
			this.ComboBox_WorkingStage.Location = new System.Drawing.Point(5, 41);
			this.ComboBox_WorkingStage.Name = "ComboBox_WorkingStage";
			this.ComboBox_WorkingStage.Size = new System.Drawing.Size(90, 20);
			this.ComboBox_WorkingStage.TabIndex = 1;
			this.ToolTip1.SetToolTip(this.ComboBox_WorkingStage, "用来在绘制测斜位移的最值走势图时，在图表中标开挖工况信息");
			//
			//Label3
			//
			this.Label3.AutoSize = true;
			this.Label3.Location = new System.Drawing.Point(3, 21);
			this.Label3.Name = "Label3";
			this.Label3.Size = new System.Drawing.Size(53, 12);
			this.Label3.TabIndex = 3;
			this.Label3.Text = "开挖工况";
			//
			//Panel_Static
			//
			this.Panel_Static.Anchor = (System.Windows.Forms.AnchorStyles) (System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right);
			this.Panel_Static.Controls.Add(this.Label3);
			this.Panel_Static.Controls.Add(this.ComboBox_WorkingStage);
			this.Panel_Static.Location = new System.Drawing.Point(186, 98);
			this.Panel_Static.Name = "Panel_Static";
			this.Panel_Static.Size = new System.Drawing.Size(106, 73);
			this.Panel_Static.TabIndex = 11;
			//
			//ComboBoxOpenedWorkbook
			//
			this.ComboBoxOpenedWorkbook.Anchor = (System.Windows.Forms.AnchorStyles) ((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right);
			this.ComboBoxOpenedWorkbook.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.ComboBoxOpenedWorkbook.FormattingEnabled = true;
			this.ComboBoxOpenedWorkbook.Location = new System.Drawing.Point(94, 15);
			this.ComboBoxOpenedWorkbook.Name = "ComboBoxOpenedWorkbook";
			this.ComboBoxOpenedWorkbook.Size = new System.Drawing.Size(335, 20);
			this.ComboBoxOpenedWorkbook.TabIndex = 12;
			//
			//CheckBox1
			//
			this.CheckBox1.AutoSize = true;
			this.CheckBox1.Location = new System.Drawing.Point(6, 77);
			this.CheckBox1.Name = "CheckBox1";
			this.CheckBox1.Size = new System.Drawing.Size(60, 16);
			this.CheckBox1.TabIndex = 1;
			this.CheckBox1.Text = "自定义";
			this.CheckBox1.UseVisualStyleBackColor = true;
			//
			//ComboBox1
			//
			this.ComboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.ComboBox1.FormattingEnabled = true;
			this.ComboBox1.Items.AddRange(new object[] {"空间分布1"});
			this.ComboBox1.Location = new System.Drawing.Point(6, 100);
			this.ComboBox1.Name = "ComboBox1";
			this.ComboBox1.Size = new System.Drawing.Size(90, 20);
			this.ComboBox1.TabIndex = 2;
			//
			//frmDrawing_Mnt_Incline
			//
			this.AutoScaleDimensions = new System.Drawing.SizeF((float) (6.0F), (float) (12.0F));
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(442, 318);
			this.Controls.Add(this.ComboBoxOpenedWorkbook);
			this.Controls.Add(this.Panel_Static);
			this.Controls.Add(this.Label_MntType);
			this.Controls.Add(this.ComboBox_MntType);
			this.Controls.Add(this.Panel_Dynamic);
			this.Controls.Add(this.GroupBox1);
			this.Controls.Add(this.btnDrawMonitorPoints);
			this.Controls.Add(this.ListBoxWorksheetsName);
			this.Controls.Add(this.chkBoxOpenNewExcel);
			this.Controls.Add(this.btnGenerate);
			this.Controls.Add(this.btnChooseMonitorData);
			this.MinimumSize = new System.Drawing.Size(333, 303);
			this.Name = "frmDrawing_Mnt_Incline";
			this.Text = "测斜曲线绘制";
			this.GroupBox1.ResumeLayout(false);
			this.GroupBox1.PerformLayout();
			this.Panel_Dynamic.ResumeLayout(false);
			this.Panel_Dynamic.PerformLayout();
			this.Panel_Static.ResumeLayout(false);
			this.Panel_Static.PerformLayout();
			this.ResumeLayout(false);
			this.PerformLayout();
			
		}
		internal System.Windows.Forms.Button btnChooseMonitorData;
		internal System.Windows.Forms.ComboBox ComboBox_ExcavID;
		internal System.Windows.Forms.Label Label_Component_Elevation;
		internal System.Windows.Forms.Button btnGenerate;
		internal System.Windows.Forms.CheckBox chkBoxOpenNewExcel;
		internal System.Windows.Forms.ListBox ListBoxWorksheetsName;
		internal System.Windows.Forms.ComboBox ComboBox_ExcavRegion;
		internal System.Windows.Forms.Label Label2;
		internal System.Windows.Forms.Button btnDrawMonitorPoints;
		internal System.ComponentModel.BackgroundWorker BGWK_NewDrawing;
		internal System.Windows.Forms.GroupBox GroupBox1;
		internal System.Windows.Forms.RadioButton RadioButton_Dynamic;
		internal System.Windows.Forms.RadioButton RadioButton_Max_Depth;
		internal System.Windows.Forms.Panel Panel_Dynamic;
		internal System.Windows.Forms.ComboBox ComboBox_MntType;
		internal System.Windows.Forms.Label Label_MntType;
		internal System.Windows.Forms.ComboBox ComboBox_WorkingStage;
		internal System.Windows.Forms.Label Label3;
		internal System.Windows.Forms.Panel Panel_Static;
		internal System.Windows.Forms.ToolTip ToolTip1;
		internal System.Windows.Forms.ComboBox ComboBoxOpenedWorkbook;
		internal System.Windows.Forms.ComboBox ComboBox1;
		internal System.Windows.Forms.CheckBox CheckBox1;
	}
	
}
