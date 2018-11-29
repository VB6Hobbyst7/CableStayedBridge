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
	namespace DataBase
	{
		
		[global::Microsoft.VisualBasic.CompilerServices.DesignerGenerated()]public 
		partial class frmProjectFile : System.Windows.Forms.Form
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
				this.LstbxSheetsProgressInWkbk = new System.Windows.Forms.ListBox();
				base.Load += new System.EventHandler(frmProjectContents_Load);
				this.LstbxSheetsProgressInWkbk.Click += new System.EventHandler(this.PanelFather_MouseEnter);
				this.LstbxSheetsProgressInProject = new System.Windows.Forms.ListBox();
				this.LstbxSheetsProgressInProject.Click += new System.EventHandler(this.PanelFather_MouseEnter);
				this.LstbxSheetsProgressInProject.DataSourceChanged += new System.EventHandler(this.LstbxSheetsProgressInProject_DataSourceChanged);
				this.Label1 = new System.Windows.Forms.Label();
				this.Label2 = new System.Windows.Forms.Label();
				this.BtnAddSheet = new System.Windows.Forms.Button();
				this.BtnAddSheet.Click += new System.EventHandler(this.BtnAddSheet_Click);
				this.BtnRemoveSheet = new System.Windows.Forms.Button();
				this.BtnRemoveSheet.Click += new System.EventHandler(this.BtnRemoveSheet_Click);
				this.Label3 = new System.Windows.Forms.Label();
				this.Label5 = new System.Windows.Forms.Label();
				this.CmbbxSectional = new System.Windows.Forms.ComboBox();
				this.CmbbxSectional.SelectedIndexChanged += new System.EventHandler(this.CmbbxSectional_SelectedIndexChanged);
				this.Label4 = new System.Windows.Forms.Label();
				this.LineShape2 = new Microsoft.VisualBasic.PowerPacks.LineShape();
				this.LineShape1 = new Microsoft.VisualBasic.PowerPacks.LineShape();
				this.PanelGeneral = new System.Windows.Forms.Panel();
				this.CmbbxWorkingStage = new System.Windows.Forms.ComboBox();
				this.CmbbxWorkingStage.SelectedIndexChanged += new System.EventHandler(this.CmbbxWorkingStage_SelectedIndexChanged);
				this.Label9 = new System.Windows.Forms.Label();
				this.CmbbxPointCoordinates = new System.Windows.Forms.ComboBox();
				this.CmbbxPointCoordinates.SelectedIndexChanged += new System.EventHandler(this.CmbbxPointCoordinates_SelectedIndexChanged);
				this.CmbbxPlan = new System.Windows.Forms.ComboBox();
				this.CmbbxPlan.SelectedIndexChanged += new System.EventHandler(this.CmbbxPlan_SelectedIndexChanged);
				this.ShapeContainer2 = new Microsoft.VisualBasic.PowerPacks.ShapeContainer();
				this.LineShape3 = new Microsoft.VisualBasic.PowerPacks.LineShape();
				this.PanelFather = new System.Windows.Forms.Panel();
				this.PanelFather.MouseEnter += new System.EventHandler(this.PanelFather_MouseEnter);
				this.GroupBox1 = new System.Windows.Forms.GroupBox();
				this.CmbbxProgressWkbk = new System.Windows.Forms.ComboBox();
				this.CmbbxProgressWkbk.SelectedValueChanged += new System.EventHandler(this.CmbbxProgressWkbk_SelectedValueChanged);
				this.Label6 = new System.Windows.Forms.Label();
				this.btnAddWorkbook = new System.Windows.Forms.Button();
				this.btnAddWorkbook.Click += new System.EventHandler(this.btnAddWorkbook_Click);
				this.LstBxWorkbooks = new System.Windows.Forms.ListBox();
				this.LstBxWorkbooks.DataSourceChanged += new System.EventHandler(this.LstBxWorkbooks_DataSourceChanged);
				this.Label7 = new System.Windows.Forms.Label();
				this.BtnRemoveWorkbook = new System.Windows.Forms.Button();
				this.BtnRemoveWorkbook.Click += new System.EventHandler(this.BtnRemoveWorkbook_Click);
				this.btnOk = new System.Windows.Forms.Button();
				this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
				this.btnCancel = new System.Windows.Forms.Button();
				this.Label8 = new System.Windows.Forms.Label();
				this.Panel1 = new System.Windows.Forms.Panel();
				this.LabelProjectFilePath = new System.Windows.Forms.Label();
				this.PanelGeneral.SuspendLayout();
				this.PanelFather.SuspendLayout();
				this.GroupBox1.SuspendLayout();
				this.Panel1.SuspendLayout();
				this.SuspendLayout();
				//
				//LstbxSheetsProgressInWkbk
				//
				this.LstbxSheetsProgressInWkbk.FormattingEnabled = true;
				this.LstbxSheetsProgressInWkbk.HorizontalScrollbar = true;
				this.LstbxSheetsProgressInWkbk.ItemHeight = 12;
				this.LstbxSheetsProgressInWkbk.Location = new System.Drawing.Point(12, 80);
				this.LstbxSheetsProgressInWkbk.Name = "LstbxSheetsProgressInWkbk";
				this.LstbxSheetsProgressInWkbk.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
				this.LstbxSheetsProgressInWkbk.Size = new System.Drawing.Size(129, 172);
				this.LstbxSheetsProgressInWkbk.TabIndex = 0;
				//
				//LstbxSheetsProgressInProject
				//
				this.LstbxSheetsProgressInProject.FormattingEnabled = true;
				this.LstbxSheetsProgressInProject.HorizontalScrollbar = true;
				this.LstbxSheetsProgressInProject.ItemHeight = 12;
				this.LstbxSheetsProgressInProject.Location = new System.Drawing.Point(149, 80);
				this.LstbxSheetsProgressInProject.Name = "LstbxSheetsProgressInProject";
				this.LstbxSheetsProgressInProject.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
				this.LstbxSheetsProgressInProject.Size = new System.Drawing.Size(129, 172);
				this.LstbxSheetsProgressInProject.TabIndex = 0;
				//
				//Label1
				//
				this.Label1.AutoSize = true;
				this.Label1.Location = new System.Drawing.Point(10, 58);
				this.Label1.Name = "Label1";
				this.Label1.Size = new System.Drawing.Size(89, 12);
				this.Label1.TabIndex = 3;
				this.Label1.Text = "文件中的工作表";
				//
				//Label2
				//
				this.Label2.AutoSize = true;
				this.Label2.Location = new System.Drawing.Point(147, 58);
				this.Label2.Name = "Label2";
				this.Label2.Size = new System.Drawing.Size(89, 12);
				this.Label2.TabIndex = 3;
				this.Label2.Text = "项目中的工作表";
				//
				//BtnAddSheet
				//
				this.BtnAddSheet.Location = new System.Drawing.Point(284, 132);
				this.BtnAddSheet.Name = "BtnAddSheet";
				this.BtnAddSheet.Size = new System.Drawing.Size(61, 23);
				this.BtnAddSheet.TabIndex = 4;
				this.BtnAddSheet.Text = "Add";
				this.BtnAddSheet.UseVisualStyleBackColor = true;
				//
				//BtnRemoveSheet
				//
				this.BtnRemoveSheet.Location = new System.Drawing.Point(284, 176);
				this.BtnRemoveSheet.Name = "BtnRemoveSheet";
				this.BtnRemoveSheet.Size = new System.Drawing.Size(61, 23);
				this.BtnRemoveSheet.TabIndex = 4;
				this.BtnRemoveSheet.Text = "Remove";
				this.BtnRemoveSheet.UseVisualStyleBackColor = true;
				//
				//Label3
				//
				this.Label3.AutoSize = true;
				this.Label3.Location = new System.Drawing.Point(3, 11);
				this.Label3.Name = "Label3";
				this.Label3.Size = new System.Drawing.Size(53, 12);
				this.Label3.TabIndex = 1;
				this.Label3.Text = "剖面标高";
				//
				//Label5
				//
				this.Label5.AutoSize = true;
				this.Label5.Location = new System.Drawing.Point(3, 122);
				this.Label5.Name = "Label5";
				this.Label5.Size = new System.Drawing.Size(53, 12);
				this.Label5.TabIndex = 1;
				this.Label5.Text = "测点坐标";
				//
				//CmbbxSectional
				//
				this.CmbbxSectional.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
				this.CmbbxSectional.FormattingEnabled = true;
				this.CmbbxSectional.Location = new System.Drawing.Point(62, 8);
				this.CmbbxSectional.Name = "CmbbxSectional";
				this.CmbbxSectional.Size = new System.Drawing.Size(271, 20);
				this.CmbbxSectional.TabIndex = 2;
				//
				//Label4
				//
				this.Label4.AutoSize = true;
				this.Label4.Location = new System.Drawing.Point(3, 68);
				this.Label4.Name = "Label4";
				this.Label4.Size = new System.Drawing.Size(53, 12);
				this.Label4.TabIndex = 1;
				this.Label4.Text = "开挖分块";
				//
				//LineShape2
				//
				this.LineShape2.BorderColor = System.Drawing.SystemColors.ControlDark;
				this.LineShape2.Cursor = System.Windows.Forms.Cursors.Default;
				this.LineShape2.Enabled = false;
				this.LineShape2.Name = "LineShape1";
				this.LineShape2.X1 = 12;
				this.LineShape2.X2 = 344;
				this.LineShape2.Y1 = 48;
				this.LineShape2.Y2 = 48;
				//
				//LineShape1
				//
				this.LineShape1.BorderColor = System.Drawing.SystemColors.ControlDark;
				this.LineShape1.Enabled = false;
				this.LineShape1.Name = "LineShape1";
				this.LineShape1.X1 = 12;
				this.LineShape1.X2 = 344;
				this.LineShape1.Y1 = 104;
				this.LineShape1.Y2 = 104;
				//
				//PanelGeneral
				//
				this.PanelGeneral.Controls.Add(this.CmbbxWorkingStage);
				this.PanelGeneral.Controls.Add(this.Label9);
				this.PanelGeneral.Controls.Add(this.CmbbxPointCoordinates);
				this.PanelGeneral.Controls.Add(this.CmbbxPlan);
				this.PanelGeneral.Controls.Add(this.Label3);
				this.PanelGeneral.Controls.Add(this.Label5);
				this.PanelGeneral.Controls.Add(this.Label4);
				this.PanelGeneral.Controls.Add(this.CmbbxSectional);
				this.PanelGeneral.Controls.Add(this.ShapeContainer2);
				this.PanelGeneral.Location = new System.Drawing.Point(3, 3);
				this.PanelGeneral.Name = "PanelGeneral";
				this.PanelGeneral.Size = new System.Drawing.Size(358, 215);
				this.PanelGeneral.TabIndex = 8;
				//
				//CmbbxWorkingStage
				//
				this.CmbbxWorkingStage.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
				this.CmbbxWorkingStage.FormattingEnabled = true;
				this.CmbbxWorkingStage.Location = new System.Drawing.Point(65, 175);
				this.CmbbxWorkingStage.Name = "CmbbxWorkingStage";
				this.CmbbxWorkingStage.Size = new System.Drawing.Size(271, 20);
				this.CmbbxWorkingStage.TabIndex = 2;
				//
				//Label9
				//
				this.Label9.AutoSize = true;
				this.Label9.Location = new System.Drawing.Point(3, 178);
				this.Label9.Name = "Label9";
				this.Label9.Size = new System.Drawing.Size(53, 12);
				this.Label9.TabIndex = 1;
				this.Label9.Text = "开挖工况";
				//
				//CmbbxPointCoordinates
				//
				this.CmbbxPointCoordinates.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
				this.CmbbxPointCoordinates.FormattingEnabled = true;
				this.CmbbxPointCoordinates.Location = new System.Drawing.Point(65, 119);
				this.CmbbxPointCoordinates.Name = "CmbbxPointCoordinates";
				this.CmbbxPointCoordinates.Size = new System.Drawing.Size(271, 20);
				this.CmbbxPointCoordinates.TabIndex = 2;
				//
				//CmbbxPlan
				//
				this.CmbbxPlan.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
				this.CmbbxPlan.FormattingEnabled = true;
				this.CmbbxPlan.Location = new System.Drawing.Point(65, 65);
				this.CmbbxPlan.Name = "CmbbxPlan";
				this.CmbbxPlan.Size = new System.Drawing.Size(271, 20);
				this.CmbbxPlan.TabIndex = 2;
				//
				//ShapeContainer2
				//
				this.ShapeContainer2.Location = new System.Drawing.Point(0, 0);
				this.ShapeContainer2.Margin = new System.Windows.Forms.Padding(0);
				this.ShapeContainer2.Name = "ShapeContainer2";
				this.ShapeContainer2.Shapes.AddRange(new Microsoft.VisualBasic.PowerPacks.Shape[] {this.LineShape3, this.LineShape2, this.LineShape1});
				this.ShapeContainer2.Size = new System.Drawing.Size(358, 215);
				this.ShapeContainer2.TabIndex = 4;
				this.ShapeContainer2.TabStop = false;
				//
				//LineShape3
				//
				this.LineShape3.BorderColor = System.Drawing.SystemColors.ControlDark;
				this.LineShape3.Cursor = System.Windows.Forms.Cursors.Default;
				this.LineShape3.Enabled = false;
				this.LineShape3.Name = "LineShape1";
				this.LineShape3.X1 = 12;
				this.LineShape3.X2 = 344;
				this.LineShape3.Y1 = 160;
				this.LineShape3.Y2 = 160;
				//
				//PanelFather
				//
				this.PanelFather.AutoScroll = true;
				this.PanelFather.AutoScrollMargin = new System.Drawing.Size(0, 10);
				this.PanelFather.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
				this.PanelFather.Controls.Add(this.GroupBox1);
				this.PanelFather.Controls.Add(this.PanelGeneral);
				this.PanelFather.Location = new System.Drawing.Point(12, 146);
				this.PanelFather.Name = "PanelFather";
				this.PanelFather.Size = new System.Drawing.Size(389, 260);
				this.PanelFather.TabIndex = 10;
				//
				//GroupBox1
				//
				this.GroupBox1.Controls.Add(this.CmbbxProgressWkbk);
				this.GroupBox1.Controls.Add(this.Label6);
				this.GroupBox1.Controls.Add(this.Label1);
				this.GroupBox1.Controls.Add(this.Label2);
				this.GroupBox1.Controls.Add(this.LstbxSheetsProgressInWkbk);
				this.GroupBox1.Controls.Add(this.BtnRemoveSheet);
				this.GroupBox1.Controls.Add(this.BtnAddSheet);
				this.GroupBox1.Controls.Add(this.LstbxSheetsProgressInProject);
				this.GroupBox1.Location = new System.Drawing.Point(3, 224);
				this.GroupBox1.Name = "GroupBox1";
				this.GroupBox1.Size = new System.Drawing.Size(358, 260);
				this.GroupBox1.TabIndex = 9;
				this.GroupBox1.TabStop = false;
				this.GroupBox1.Text = "施工进度";
				//
				//CmbbxProgressWkbk
				//
				this.CmbbxProgressWkbk.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
				this.CmbbxProgressWkbk.FormattingEnabled = true;
				this.CmbbxProgressWkbk.Location = new System.Drawing.Point(107, 26);
				this.CmbbxProgressWkbk.Name = "CmbbxProgressWkbk";
				this.CmbbxProgressWkbk.Size = new System.Drawing.Size(229, 20);
				this.CmbbxProgressWkbk.TabIndex = 6;
				//
				//Label6
				//
				this.Label6.AutoSize = true;
				this.Label6.Location = new System.Drawing.Point(12, 30);
				this.Label6.Name = "Label6";
				this.Label6.Size = new System.Drawing.Size(89, 12);
				this.Label6.TabIndex = 5;
				this.Label6.Text = "选择工作簿文件";
				//
				//btnAddWorkbook
				//
				this.btnAddWorkbook.Location = new System.Drawing.Point(305, 27);
				this.btnAddWorkbook.Name = "btnAddWorkbook";
				this.btnAddWorkbook.Size = new System.Drawing.Size(75, 23);
				this.btnAddWorkbook.TabIndex = 11;
				this.btnAddWorkbook.Text = "Add";
				this.btnAddWorkbook.UseVisualStyleBackColor = true;
				//
				//LstBxWorkbooks
				//
				this.LstBxWorkbooks.FormattingEnabled = true;
				this.LstBxWorkbooks.HorizontalScrollbar = true;
				this.LstBxWorkbooks.ItemHeight = 12;
				this.LstBxWorkbooks.Location = new System.Drawing.Point(6, 27);
				this.LstBxWorkbooks.Name = "LstBxWorkbooks";
				this.LstBxWorkbooks.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple;
				this.LstBxWorkbooks.Size = new System.Drawing.Size(293, 64);
				this.LstBxWorkbooks.TabIndex = 12;
				//
				//Label7
				//
				this.Label7.AutoSize = true;
				this.Label7.Location = new System.Drawing.Point(6, 6);
				this.Label7.Name = "Label7";
				this.Label7.Size = new System.Drawing.Size(89, 12);
				this.Label7.TabIndex = 13;
				this.Label7.Text = "数据工作簿列表";
				//
				//BtnRemoveWorkbook
				//
				this.BtnRemoveWorkbook.Location = new System.Drawing.Point(305, 68);
				this.BtnRemoveWorkbook.Name = "BtnRemoveWorkbook";
				this.BtnRemoveWorkbook.Size = new System.Drawing.Size(75, 23);
				this.BtnRemoveWorkbook.TabIndex = 11;
				this.BtnRemoveWorkbook.Text = "Remove";
				this.BtnRemoveWorkbook.UseVisualStyleBackColor = true;
				//
				//btnOk
				//
				this.btnOk.Location = new System.Drawing.Point(326, 421);
				this.btnOk.Name = "btnOk";
				this.btnOk.Size = new System.Drawing.Size(75, 23);
				this.btnOk.TabIndex = 14;
				this.btnOk.Text = "确定";
				this.btnOk.UseVisualStyleBackColor = true;
				//
				//btnCancel
				//
				this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
				this.btnCancel.Location = new System.Drawing.Point(243, 421);
				this.btnCancel.Name = "btnCancel";
				this.btnCancel.Size = new System.Drawing.Size(75, 23);
				this.btnCancel.TabIndex = 14;
				this.btnCancel.Text = "取消(&C)";
				this.btnCancel.UseVisualStyleBackColor = true;
				//
				//Label8
				//
				this.Label8.AutoSize = true;
				this.Label8.Location = new System.Drawing.Point(14, 15);
				this.Label8.Name = "Label8";
				this.Label8.Size = new System.Drawing.Size(71, 12);
				this.Label8.TabIndex = 16;
				this.Label8.Text = "项目文件 : ";
				//
				//Panel1
				//
				this.Panel1.Controls.Add(this.Label7);
				this.Panel1.Controls.Add(this.btnAddWorkbook);
				this.Panel1.Controls.Add(this.BtnRemoveWorkbook);
				this.Panel1.Controls.Add(this.LstBxWorkbooks);
				this.Panel1.Location = new System.Drawing.Point(12, 43);
				this.Panel1.Name = "Panel1";
				this.Panel1.Size = new System.Drawing.Size(389, 94);
				this.Panel1.TabIndex = 17;
				//
				//LabelProjectFilePath
				//
				this.LabelProjectFilePath.AutoSize = true;
				this.LabelProjectFilePath.Location = new System.Drawing.Point(79, 15);
				this.LabelProjectFilePath.Name = "LabelProjectFilePath";
				this.LabelProjectFilePath.Size = new System.Drawing.Size(53, 12);
				this.LabelProjectFilePath.TabIndex = 18;
				this.LabelProjectFilePath.Text = "FilePath";
				//
				//frmProjectFile
				//
				this.AcceptButton = this.btnOk;
				this.AutoScaleDimensions = new System.Drawing.SizeF((float) (6.0F), (float) (12.0F));
				this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
				this.CancelButton = this.btnCancel;
				this.ClientSize = new System.Drawing.Size(419, 456);
				this.Controls.Add(this.LabelProjectFilePath);
				this.Controls.Add(this.Label8);
				this.Controls.Add(this.Panel1);
				this.Controls.Add(this.btnCancel);
				this.Controls.Add(this.btnOk);
				this.Controls.Add(this.PanelFather);
				this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
				this.HelpButton = true;
				this.MaximizeBox = false;
				this.MinimizeBox = false;
				this.Name = "frmProjectFile";
				this.Text = "New Project";
				this.PanelGeneral.ResumeLayout(false);
				this.PanelGeneral.PerformLayout();
				this.PanelFather.ResumeLayout(false);
				this.GroupBox1.ResumeLayout(false);
				this.GroupBox1.PerformLayout();
				this.Panel1.ResumeLayout(false);
				this.Panel1.PerformLayout();
				this.ResumeLayout(false);
				this.PerformLayout();
				
			}
			internal System.Windows.Forms.ListBox LstbxSheetsProgressInWkbk;
			internal System.Windows.Forms.ListBox LstbxSheetsProgressInProject;
			internal System.Windows.Forms.Label Label1;
			internal System.Windows.Forms.Label Label2;
			internal System.Windows.Forms.Button BtnAddSheet;
			internal System.Windows.Forms.Button BtnRemoveSheet;
			internal System.Windows.Forms.ComboBox CmbbxSectional;
			internal System.Windows.Forms.Label Label5;
			internal System.Windows.Forms.Label Label4;
			internal System.Windows.Forms.Label Label3;
			internal Microsoft.VisualBasic.PowerPacks.LineShape LineShape2;
			internal Microsoft.VisualBasic.PowerPacks.LineShape LineShape1;
			internal System.Windows.Forms.Panel PanelGeneral;
			internal System.Windows.Forms.ComboBox CmbbxPointCoordinates;
			internal System.Windows.Forms.ComboBox CmbbxPlan;
			internal Microsoft.VisualBasic.PowerPacks.ShapeContainer ShapeContainer2;
			internal System.Windows.Forms.Panel PanelFather;
			internal System.Windows.Forms.GroupBox GroupBox1;
			internal System.Windows.Forms.ComboBox CmbbxProgressWkbk;
			internal System.Windows.Forms.Label Label6;
			internal System.Windows.Forms.Button btnAddWorkbook;
			internal System.Windows.Forms.ListBox LstBxWorkbooks;
			internal System.Windows.Forms.Label Label7;
			internal System.Windows.Forms.Button BtnRemoveWorkbook;
			internal System.Windows.Forms.Button btnOk;
			internal System.Windows.Forms.Button btnCancel;
			internal System.Windows.Forms.Label Label8;
			internal System.Windows.Forms.Panel Panel1;
			internal System.Windows.Forms.Label LabelProjectFilePath;
			internal System.Windows.Forms.ComboBox CmbbxWorkingStage;
			internal System.Windows.Forms.Label Label9;
			internal Microsoft.VisualBasic.PowerPacks.LineShape LineShape3;
		}
	}
}
