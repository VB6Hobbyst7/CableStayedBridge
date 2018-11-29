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
	partial class frmDrawingPlan : System.Windows.Forms.Form
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
			this.Label1 = new System.Windows.Forms.Label();
			this.TextBoxInfoBoxID = new System.Windows.Forms.TextBox();
			this.GroupBox1 = new System.Windows.Forms.GroupBox();
			this.TextBoxAllRegions = new System.Windows.Forms.TextBox();
			this.Label2 = new System.Windows.Forms.Label();
			this.ToolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.Label4 = new System.Windows.Forms.Label();
			this.Label7 = new System.Windows.Forms.Label();
			this.Label8 = new System.Windows.Forms.Label();
			this.Label9 = new System.Windows.Forms.Label();
			this.Label10 = new System.Windows.Forms.Label();
			this.TextBoxPageName = new System.Windows.Forms.TextBox();
			this.BtnGenerate = new System.Windows.Forms.Button();
			this.BtnGenerate.Click += new System.EventHandler(this.ConstructVisioPlanView);
			this.TextBoxFilePath = new System.Windows.Forms.TextBox();
			this.btnChooseVisioPlanView = new System.Windows.Forms.Button();
			this.btnChooseVisioPlanView.Click += new System.EventHandler(this.btnChooseVisioPlanView_Click);
			this.ChkBx_PointInfo = new System.Windows.Forms.CheckBox();
			this.ChkBx_PointInfo.CheckedChanged += new System.EventHandler(this.ChkBx_PointInfo_CheckedChanged);
			this.txtbx_ShapeName_MonitorPointTag = new System.Windows.Forms.TextBox();
			this.Panel1 = new System.Windows.Forms.Panel();
			this.TableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
			this.txtbx_Pt_UR_CAD_Y = new System.Windows.Forms.TextBox();
			this.txtbx_Pt_UR_CAD_Y.Validating += new System.ComponentModel.CancelEventHandler(this.ValidateForSingle);
			this.txtbx_Pt_BL_CAD_Y = new System.Windows.Forms.TextBox();
			this.txtbx_Pt_BL_CAD_Y.Validating += new System.ComponentModel.CancelEventHandler(this.ValidateForSingle);
			this.txtbx_Pt_BL_CAD_X = new System.Windows.Forms.TextBox();
			this.txtbx_Pt_BL_CAD_X.Validating += new System.ComponentModel.CancelEventHandler(this.ValidateForSingle);
			this.txtbx_Pt_UR_ShapeID = new System.Windows.Forms.TextBox();
			this.txtbx_Pt_UR_ShapeID.Validating += new System.ComponentModel.CancelEventHandler(this.ValidateForInteger);
			this.txtbx_Pt_UR_CAD_X = new System.Windows.Forms.TextBox();
			this.txtbx_Pt_UR_CAD_X.Validating += new System.ComponentModel.CancelEventHandler(this.ValidateForSingle);
			this.Label12 = new System.Windows.Forms.Label();
			this.txtbx_Pt_BL_ShapeID = new System.Windows.Forms.TextBox();
			this.txtbx_Pt_BL_ShapeID.Validating += new System.ComponentModel.CancelEventHandler(this.ValidateForInteger);
			this.Label11 = new System.Windows.Forms.Label();
			this.Btn_Import = new System.Windows.Forms.Button();
			this.Btn_Import.Click += new System.EventHandler(this.Btn_Import_Click);
			this.Btn_Export = new System.Windows.Forms.Button();
			this.Btn_Export.Click += new System.EventHandler(this.Btn_Export_Click);
			this.btnCancel = new System.Windows.Forms.Button();
			this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
			this.GroupBox1.SuspendLayout();
			this.Panel1.SuspendLayout();
			this.TableLayoutPanel1.SuspendLayout();
			this.SuspendLayout();
			//
			//Label1
			//
			this.Label1.AutoSize = true;
			this.Label1.Location = new System.Drawing.Point(6, 23);
			this.Label1.Name = "Label1";
			this.Label1.Size = new System.Drawing.Size(65, 12);
			this.Label1.TabIndex = 0;
			this.Label1.Text = "信息文本框";
			this.ToolTip1.SetToolTip(this.Label1, "记录开挖信息的文本框的ID值");
			//
			//TextBoxInfoBoxID
			//
			this.TextBoxInfoBoxID.Location = new System.Drawing.Point(100, 20);
			this.TextBoxInfoBoxID.Name = "TextBoxInfoBoxID";
			this.TextBoxInfoBoxID.Size = new System.Drawing.Size(72, 21);
			this.TextBoxInfoBoxID.TabIndex = 1;
			this.TextBoxInfoBoxID.Text = "5079";
			//
			//GroupBox1
			//
			this.GroupBox1.Controls.Add(this.TextBoxAllRegions);
			this.GroupBox1.Controls.Add(this.TextBoxInfoBoxID);
			this.GroupBox1.Controls.Add(this.Label2);
			this.GroupBox1.Controls.Add(this.Label1);
			this.GroupBox1.Location = new System.Drawing.Point(6, 84);
			this.GroupBox1.Name = "GroupBox1";
			this.GroupBox1.Size = new System.Drawing.Size(356, 56);
			this.GroupBox1.TabIndex = 2;
			this.GroupBox1.TabStop = false;
			this.GroupBox1.Text = "特征形状ID值";
			//
			//TextBoxAllRegions
			//
			this.TextBoxAllRegions.Location = new System.Drawing.Point(272, 20);
			this.TextBoxAllRegions.Name = "TextBoxAllRegions";
			this.TextBoxAllRegions.Size = new System.Drawing.Size(72, 21);
			this.TextBoxAllRegions.TabIndex = 1;
			this.TextBoxAllRegions.Text = "5078";
			//
			//Label2
			//
			this.Label2.AutoSize = true;
			this.Label2.Location = new System.Drawing.Point(178, 23);
			this.Label2.Name = "Label2";
			this.Label2.Size = new System.Drawing.Size(53, 12);
			this.Label2.TabIndex = 0;
			this.Label2.Text = "所有分区";
			this.ToolTip1.SetToolTip(this.Label2, "所有分区的组合形状的ID值");
			//
			//Label4
			//
			this.Label4.AutoSize = true;
			this.Label4.Location = new System.Drawing.Point(12, 48);
			this.Label4.Name = "Label4";
			this.Label4.Size = new System.Drawing.Size(53, 12);
			this.Label4.TabIndex = 0;
			this.Label4.Text = "页面名称";
			this.ToolTip1.SetToolTip(this.Label4, "开挖平面图在Visio中的页面名称");
			//
			//Label7
			//
			this.Label7.AutoSize = true;
			this.Label7.Location = new System.Drawing.Point(5, 9);
			this.Label7.Name = "Label7";
			this.Label7.Size = new System.Drawing.Size(209, 12);
			this.Label7.TabIndex = 0;
			this.Label7.Text = "测点主控形状中表示编号的形状的Name";
			this.ToolTip1.SetToolTip(this.Label7, "visio中在监测点的主控形状中，用来显示测点编号的形状的Name属性");
			//
			//Label8
			//
			this.Label8.AutoSize = true;
			this.Label8.Location = new System.Drawing.Point(3, 20);
			this.Label8.Name = "Label8";
			this.Label8.Size = new System.Drawing.Size(95, 12);
			this.Label8.TabIndex = 0;
			this.Label8.Text = "Visio中的形状ID";
			this.ToolTip1.SetToolTip(this.Label8, "Visio平面图中用于坐标变换的两个定位点的形状ID，" + System.Convert.ToString(global::Microsoft.VisualBasic.Strings.ChrW(13)) + System.Convert.ToString(global::Microsoft.VisualBasic.Strings.ChrW(10)) + "这两个点分别代表ABCD基坑群的左下角与右上角。");
			//
			//Label9
			//
			this.Label9.AutoSize = true;
			this.Label9.Location = new System.Drawing.Point(3, 40);
			this.Label9.Name = "Label9";
			this.Label9.Size = new System.Drawing.Size(101, 12);
			this.Label9.TabIndex = 1;
			this.Label9.Text = "CAD中点的坐标(X)";
			this.ToolTip1.SetToolTip(this.Label9, "CAD平面图中用于坐标变换的两个定位点的坐标，这两个点分别代表ABCD基坑群的左下角与右上角。");
			//
			//Label10
			//
			this.Label10.AutoSize = true;
			this.Label10.Location = new System.Drawing.Point(3, 60);
			this.Label10.Name = "Label10";
			this.Label10.Size = new System.Drawing.Size(101, 12);
			this.Label10.TabIndex = 2;
			this.Label10.Text = "CAD中点的坐标(Y)";
			this.ToolTip1.SetToolTip(this.Label10, "CAD平面图中用于坐标变换的两个定位点的坐标，这两个点分别代表ABCD基坑群的左下角与右上角。");
			//
			//TextBoxPageName
			//
			this.TextBoxPageName.Location = new System.Drawing.Point(106, 45);
			this.TextBoxPageName.Name = "TextBoxPageName";
			this.TextBoxPageName.Size = new System.Drawing.Size(72, 21);
			this.TextBoxPageName.TabIndex = 1;
			this.TextBoxPageName.Text = "开挖平面";
			//
			//BtnGenerate
			//
			this.BtnGenerate.Location = new System.Drawing.Point(293, 320);
			this.BtnGenerate.Name = "BtnGenerate";
			this.BtnGenerate.Size = new System.Drawing.Size(75, 23);
			this.BtnGenerate.TabIndex = 3;
			this.BtnGenerate.Text = "确定";
			this.BtnGenerate.UseVisualStyleBackColor = true;
			//
			//TextBoxFilePath
			//
			this.TextBoxFilePath.Anchor = (System.Windows.Forms.AnchorStyles) ((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right);
			this.TextBoxFilePath.Location = new System.Drawing.Point(88, 13);
			this.TextBoxFilePath.MinimumSize = new System.Drawing.Size(210, 21);
			this.TextBoxFilePath.Name = "TextBoxFilePath";
			this.TextBoxFilePath.Size = new System.Drawing.Size(274, 21);
			this.TextBoxFilePath.TabIndex = 5;
			this.TextBoxFilePath.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
			//
			//btnChooseVisioPlanView
			//
			this.btnChooseVisioPlanView.Location = new System.Drawing.Point(6, 12);
			this.btnChooseVisioPlanView.Name = "btnChooseVisioPlanView";
			this.btnChooseVisioPlanView.Size = new System.Drawing.Size(75, 23);
			this.btnChooseVisioPlanView.TabIndex = 4;
			this.btnChooseVisioPlanView.Text = "Visio文档";
			this.btnChooseVisioPlanView.UseVisualStyleBackColor = true;
			//
			//ChkBx_PointInfo
			//
			this.ChkBx_PointInfo.AutoSize = true;
			this.ChkBx_PointInfo.Location = new System.Drawing.Point(6, 158);
			this.ChkBx_PointInfo.Name = "ChkBx_PointInfo";
			this.ChkBx_PointInfo.Size = new System.Drawing.Size(72, 16);
			this.ChkBx_PointInfo.TabIndex = 7;
			this.ChkBx_PointInfo.Text = "测点信息";
			this.ChkBx_PointInfo.UseVisualStyleBackColor = true;
			//
			//txtbx_ShapeName_MonitorPointTag
			//
			this.txtbx_ShapeName_MonitorPointTag.Location = new System.Drawing.Point(220, 6);
			this.txtbx_ShapeName_MonitorPointTag.Name = "txtbx_ShapeName_MonitorPointTag";
			this.txtbx_ShapeName_MonitorPointTag.Size = new System.Drawing.Size(75, 21);
			this.txtbx_ShapeName_MonitorPointTag.TabIndex = 1;
			this.txtbx_ShapeName_MonitorPointTag.Text = "Tag";
			//
			//Panel1
			//
			this.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.Panel1.Controls.Add(this.TableLayoutPanel1);
			this.Panel1.Controls.Add(this.Label7);
			this.Panel1.Controls.Add(this.txtbx_ShapeName_MonitorPointTag);
			this.Panel1.Location = new System.Drawing.Point(6, 180);
			this.Panel1.Name = "Panel1";
			this.Panel1.Size = new System.Drawing.Size(362, 120);
			this.Panel1.TabIndex = 8;
			//
			//TableLayoutPanel1
			//
			this.TableLayoutPanel1.ColumnCount = 3;
			this.TableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, (float) (33.33333F)));
			this.TableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, (float) (33.33334F)));
			this.TableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, (float) (33.33334F)));
			this.TableLayoutPanel1.Controls.Add(this.txtbx_Pt_UR_CAD_Y, 2, 3);
			this.TableLayoutPanel1.Controls.Add(this.txtbx_Pt_BL_CAD_Y, 1, 3);
			this.TableLayoutPanel1.Controls.Add(this.txtbx_Pt_BL_CAD_X, 1, 2);
			this.TableLayoutPanel1.Controls.Add(this.txtbx_Pt_UR_ShapeID, 2, 1);
			this.TableLayoutPanel1.Controls.Add(this.txtbx_Pt_UR_CAD_X, 2, 2);
			this.TableLayoutPanel1.Controls.Add(this.Label9, 0, 2);
			this.TableLayoutPanel1.Controls.Add(this.Label10, 0, 3);
			this.TableLayoutPanel1.Controls.Add(this.Label12, 2, 0);
			this.TableLayoutPanel1.Controls.Add(this.txtbx_Pt_BL_ShapeID, 1, 1);
			this.TableLayoutPanel1.Controls.Add(this.Label8, 0, 1);
			this.TableLayoutPanel1.Controls.Add(this.Label11, 1, 0);
			this.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Bottom;
			this.TableLayoutPanel1.Location = new System.Drawing.Point(0, 35);
			this.TableLayoutPanel1.Name = "TableLayoutPanel1";
			this.TableLayoutPanel1.RowCount = 4;
			this.TableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 25));
			this.TableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 25));
			this.TableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 25));
			this.TableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 25));
			this.TableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20));
			this.TableLayoutPanel1.Size = new System.Drawing.Size(360, 83);
			this.TableLayoutPanel1.TabIndex = 3;
			//
			//txtbx_Pt_UR_CAD_Y
			//
			this.txtbx_Pt_UR_CAD_Y.Dock = System.Windows.Forms.DockStyle.Fill;
			this.txtbx_Pt_UR_CAD_Y.Location = new System.Drawing.Point(240, 61);
			this.txtbx_Pt_UR_CAD_Y.Margin = new System.Windows.Forms.Padding(1);
			this.txtbx_Pt_UR_CAD_Y.Name = "txtbx_Pt_UR_CAD_Y";
			this.txtbx_Pt_UR_CAD_Y.Size = new System.Drawing.Size(119, 21);
			this.txtbx_Pt_UR_CAD_Y.TabIndex = 9;
			this.txtbx_Pt_UR_CAD_Y.Text = "201852.14";
			//
			//txtbx_Pt_BL_CAD_Y
			//
			this.txtbx_Pt_BL_CAD_Y.Dock = System.Windows.Forms.DockStyle.Fill;
			this.txtbx_Pt_BL_CAD_Y.Location = new System.Drawing.Point(120, 61);
			this.txtbx_Pt_BL_CAD_Y.Margin = new System.Windows.Forms.Padding(1);
			this.txtbx_Pt_BL_CAD_Y.Name = "txtbx_Pt_BL_CAD_Y";
			this.txtbx_Pt_BL_CAD_Y.Size = new System.Drawing.Size(118, 21);
			this.txtbx_Pt_BL_CAD_Y.TabIndex = 9;
			this.txtbx_Pt_BL_CAD_Y.Text = "-119668.436";
			//
			//txtbx_Pt_BL_CAD_X
			//
			this.txtbx_Pt_BL_CAD_X.Dock = System.Windows.Forms.DockStyle.Fill;
			this.txtbx_Pt_BL_CAD_X.Location = new System.Drawing.Point(120, 41);
			this.txtbx_Pt_BL_CAD_X.Margin = new System.Windows.Forms.Padding(1);
			this.txtbx_Pt_BL_CAD_X.Name = "txtbx_Pt_BL_CAD_X";
			this.txtbx_Pt_BL_CAD_X.Size = new System.Drawing.Size(118, 21);
			this.txtbx_Pt_BL_CAD_X.TabIndex = 9;
			this.txtbx_Pt_BL_CAD_X.Text = "309598.527";
			//
			//txtbx_Pt_UR_ShapeID
			//
			this.txtbx_Pt_UR_ShapeID.Dock = System.Windows.Forms.DockStyle.Fill;
			this.txtbx_Pt_UR_ShapeID.Location = new System.Drawing.Point(240, 21);
			this.txtbx_Pt_UR_ShapeID.Margin = new System.Windows.Forms.Padding(1);
			this.txtbx_Pt_UR_ShapeID.Name = "txtbx_Pt_UR_ShapeID";
			this.txtbx_Pt_UR_ShapeID.Size = new System.Drawing.Size(119, 21);
			this.txtbx_Pt_UR_ShapeID.TabIndex = 9;
			this.txtbx_Pt_UR_ShapeID.Text = "217";
			//
			//txtbx_Pt_UR_CAD_X
			//
			this.txtbx_Pt_UR_CAD_X.Dock = System.Windows.Forms.DockStyle.Fill;
			this.txtbx_Pt_UR_CAD_X.Location = new System.Drawing.Point(240, 41);
			this.txtbx_Pt_UR_CAD_X.Margin = new System.Windows.Forms.Padding(1);
			this.txtbx_Pt_UR_CAD_X.Name = "txtbx_Pt_UR_CAD_X";
			this.txtbx_Pt_UR_CAD_X.Size = new System.Drawing.Size(119, 21);
			this.txtbx_Pt_UR_CAD_X.TabIndex = 9;
			this.txtbx_Pt_UR_CAD_X.Text = "536642.644";
			//
			//Label12
			//
			this.Label12.AutoSize = true;
			this.Label12.Location = new System.Drawing.Point(242, 0);
			this.Label12.Name = "Label12";
			this.Label12.Size = new System.Drawing.Size(53, 12);
			this.Label12.TabIndex = 4;
			this.Label12.Text = "右上角点";
			//
			//txtbx_Pt_BL_ShapeID
			//
			this.txtbx_Pt_BL_ShapeID.Dock = System.Windows.Forms.DockStyle.Fill;
			this.txtbx_Pt_BL_ShapeID.Location = new System.Drawing.Point(120, 21);
			this.txtbx_Pt_BL_ShapeID.Margin = new System.Windows.Forms.Padding(1);
			this.txtbx_Pt_BL_ShapeID.Name = "txtbx_Pt_BL_ShapeID";
			this.txtbx_Pt_BL_ShapeID.Size = new System.Drawing.Size(118, 21);
			this.txtbx_Pt_BL_ShapeID.TabIndex = 5;
			this.txtbx_Pt_BL_ShapeID.Text = "197";
			//
			//Label11
			//
			this.Label11.AutoSize = true;
			this.Label11.Location = new System.Drawing.Point(122, 0);
			this.Label11.Name = "Label11";
			this.Label11.Size = new System.Drawing.Size(53, 12);
			this.Label11.TabIndex = 3;
			this.Label11.Text = "左下角点";
			//
			//Btn_Import
			//
			this.Btn_Import.Location = new System.Drawing.Point(6, 320);
			this.Btn_Import.Name = "Btn_Import";
			this.Btn_Import.Size = new System.Drawing.Size(63, 23);
			this.Btn_Import.TabIndex = 9;
			this.Btn_Import.Text = "导入";
			this.Btn_Import.UseVisualStyleBackColor = true;
			//
			//Btn_Export
			//
			this.Btn_Export.Location = new System.Drawing.Point(75, 320);
			this.Btn_Export.Name = "Btn_Export";
			this.Btn_Export.Size = new System.Drawing.Size(63, 23);
			this.Btn_Export.TabIndex = 9;
			this.Btn_Export.Text = "导出";
			this.Btn_Export.UseVisualStyleBackColor = true;
			//
			//btnCancel
			//
			this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btnCancel.Location = new System.Drawing.Point(212, 320);
			this.btnCancel.Name = "btnCancel";
			this.btnCancel.Size = new System.Drawing.Size(75, 23);
			this.btnCancel.TabIndex = 10;
			this.btnCancel.Text = "取消";
			this.btnCancel.UseVisualStyleBackColor = true;
			//
			//frmDrawingPlan
			//
			this.AcceptButton = this.BtnGenerate;
			this.AutoScaleDimensions = new System.Drawing.SizeF(6, 12);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.CancelButton = this.btnCancel;
			this.ClientSize = new System.Drawing.Size(377, 355);
			this.Controls.Add(this.btnCancel);
			this.Controls.Add(this.Btn_Export);
			this.Controls.Add(this.Btn_Import);
			this.Controls.Add(this.Panel1);
			this.Controls.Add(this.ChkBx_PointInfo);
			this.Controls.Add(this.TextBoxFilePath);
			this.Controls.Add(this.btnChooseVisioPlanView);
			this.Controls.Add(this.BtnGenerate);
			this.Controls.Add(this.GroupBox1);
			this.Controls.Add(this.TextBoxPageName);
			this.Controls.Add(this.Label4);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.MaximizeBox = false;
			this.Name = "frmDrawingPlan";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "开挖平面图";
			this.GroupBox1.ResumeLayout(false);
			this.GroupBox1.PerformLayout();
			this.Panel1.ResumeLayout(false);
			this.Panel1.PerformLayout();
			this.TableLayoutPanel1.ResumeLayout(false);
			this.TableLayoutPanel1.PerformLayout();
			this.ResumeLayout(false);
			this.PerformLayout();
			
		}
		internal System.Windows.Forms.Label Label1;
		internal System.Windows.Forms.TextBox TextBoxInfoBoxID;
		internal System.Windows.Forms.GroupBox GroupBox1;
		internal System.Windows.Forms.TextBox TextBoxAllRegions;
		internal System.Windows.Forms.Label Label2;
		internal System.Windows.Forms.ToolTip ToolTip1;
		internal System.Windows.Forms.Label Label4;
		internal System.Windows.Forms.TextBox TextBoxPageName;
		internal System.Windows.Forms.Button BtnGenerate;
		internal System.Windows.Forms.TextBox TextBoxFilePath;
		internal System.Windows.Forms.Button btnChooseVisioPlanView;
		internal System.Windows.Forms.CheckBox ChkBx_PointInfo;
		internal System.Windows.Forms.TextBox txtbx_ShapeName_MonitorPointTag;
		internal System.Windows.Forms.Label Label7;
		internal System.Windows.Forms.Panel Panel1;
		internal System.Windows.Forms.TableLayoutPanel TableLayoutPanel1;
		internal System.Windows.Forms.Label Label8;
		internal System.Windows.Forms.Label Label9;
		internal System.Windows.Forms.Label Label10;
		internal System.Windows.Forms.Label Label11;
		internal System.Windows.Forms.Label Label12;
		internal System.Windows.Forms.TextBox txtbx_Pt_BL_ShapeID;
		internal System.Windows.Forms.TextBox txtbx_Pt_UR_CAD_Y;
		internal System.Windows.Forms.TextBox txtbx_Pt_BL_CAD_Y;
		internal System.Windows.Forms.TextBox txtbx_Pt_BL_CAD_X;
		internal System.Windows.Forms.TextBox txtbx_Pt_UR_ShapeID;
		internal System.Windows.Forms.TextBox txtbx_Pt_UR_CAD_X;
		internal System.Windows.Forms.Button Btn_Import;
		internal System.Windows.Forms.Button Btn_Export;
		internal System.Windows.Forms.Button btnCancel;
	}
	
}
