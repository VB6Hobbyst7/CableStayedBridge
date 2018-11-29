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
	partial class frmDeriveDataFromExcel : System.Windows.Forms.Form
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
			System.Windows.Forms.DataGridViewCellStyle DataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
			System.Windows.Forms.DataGridViewCellStyle DataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmDeriveDataFromExcel));
			this.btnExport = new System.Windows.Forms.Button();
			base.Load += new System.EventHandler(DeriveDataFromExcel_Load);
			this.MouseMove += new System.Windows.Forms.MouseEventHandler(frmDeriveDataFromWord_MouseMove);
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(frmDeriveDataFromExcel_KeyDown);
			this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(frmDeriveDataFromWord_FormClosing);
			this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
			this.SaveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.OpenFileDialog1 = new System.Windows.Forms.OpenFileDialog();
			this.BtnChoosePath = new System.Windows.Forms.Button();
			this.BtnChoosePath.Click += new System.EventHandler(this.BtnChoosePath_Click);
			this.ListBoxWorksheets = new System.Windows.Forms.ListBox();
			this.ListBoxWorksheets.DragDrop += new System.Windows.Forms.DragEventHandler(this.APPLICATION_MAINFORM_DragDrop);
			this.ListBoxWorksheets.DragEnter += new System.Windows.Forms.DragEventHandler(this.APPLICATION_MAINFORM_DragEnter);
			this.btnRemove = new System.Windows.Forms.Button();
			this.btnRemove.Click += new System.EventHandler(this.btnRemove_Click);
			this.FolderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
			this.Label3 = new System.Windows.Forms.Label();
			this.Label4 = new System.Windows.Forms.Label();
			this.txtbxSavePath = new System.Windows.Forms.TextBox();
			this.ProgressBar1 = new System.Windows.Forms.ProgressBar();
			this.BackgroundWorker1 = new System.ComponentModel.BackgroundWorker();
			this.BackgroundWorker1.DoWork += new System.ComponentModel.DoWorkEventHandler(this.BackgroundWorker1_DoWork);
			this.BackgroundWorker1.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.BackgroundWorker1_ProgressChanged);
			this.BackgroundWorker1.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.BackgroundWorker1_RunWorkerCompleted);
			this.lbSheetName = new System.Windows.Forms.Label();
			this.AddFileOrDirectoryFiles1 = new AddFileOrDirectoryFiles();
			this.AddFileOrDirectoryFiles1.AddFile += new AddFileOrDirectoryFiles.EventHandler(this.AddFile);
			this.AddFileOrDirectoryFiles1.AddFilesFromDirectory += new AddFileOrDirectoryFiles.EventHandler(this.lbAddDir_Click);
			this.MyDataGridView1 = new myDataGridView();
			this.SheetName = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.RangeName = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.Panel1 = new System.Windows.Forms.Panel();
			this.Txtbox_DateFormat = new System.Windows.Forms.TextBox();
			this.Txtbox_DateFormat.TextChanged += new System.EventHandler(this.Txtbox_DateFormat_TextChanged);
			this.btn_DateFormat = new System.Windows.Forms.Button();
			this.btn_DateFormat.Click += new System.EventHandler(this.btn_DateFormat_Click);
			this.ChkboxParseDate = new System.Windows.Forms.CheckBox();
			this.ChkboxParseDate.CheckedChanged += new System.EventHandler(this.ChkboxParseDate_CheckedChanged);
			this.ChkBxOpenExcelWhileFinished = new System.Windows.Forms.CheckBox();
			((System.ComponentModel.ISupportInitialize) this.MyDataGridView1).BeginInit();
			this.Panel1.SuspendLayout();
			this.SuspendLayout();
			//
			//btnExport
			//
			this.btnExport.Location = new System.Drawing.Point(452, 351);
			this.btnExport.Name = "btnExport";
			this.btnExport.Size = new System.Drawing.Size(75, 23);
			this.btnExport.TabIndex = 2;
			this.btnExport.Text = "输出";
			this.btnExport.UseVisualStyleBackColor = true;
			//
			//OpenFileDialog1
			//
			this.OpenFileDialog1.FileName = "OpenFileDialog1";
			//
			//BtnChoosePath
			//
			this.BtnChoosePath.BackColor = System.Drawing.SystemColors.Control;
			this.BtnChoosePath.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
			this.BtnChoosePath.ForeColor = System.Drawing.SystemColors.InfoText;
			this.BtnChoosePath.Location = new System.Drawing.Point(453, 319);
			this.BtnChoosePath.Name = "BtnChoosePath";
			this.BtnChoosePath.Size = new System.Drawing.Size(74, 23);
			this.BtnChoosePath.TabIndex = 3;
			this.BtnChoosePath.Text = "选择...";
			this.BtnChoosePath.UseVisualStyleBackColor = false;
			//
			//ListBoxWorksheets
			//
			this.ListBoxWorksheets.AllowDrop = true;
			this.ListBoxWorksheets.FormattingEnabled = true;
			this.ListBoxWorksheets.HorizontalScrollbar = true;
			this.ListBoxWorksheets.ItemHeight = 12;
			this.ListBoxWorksheets.Location = new System.Drawing.Point(13, 38);
			this.ListBoxWorksheets.Name = "ListBoxWorksheets";
			this.ListBoxWorksheets.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
			this.ListBoxWorksheets.Size = new System.Drawing.Size(407, 136);
			this.ListBoxWorksheets.TabIndex = 6;
			//
			//btnRemove
			//
			this.btnRemove.Location = new System.Drawing.Point(426, 114);
			this.btnRemove.Name = "btnRemove";
			this.btnRemove.Size = new System.Drawing.Size(100, 24);
			this.btnRemove.TabIndex = 6;
			this.btnRemove.Text = "移除";
			this.btnRemove.UseVisualStyleBackColor = true;
			//
			//Label3
			//
			this.Label3.AutoSize = true;
			this.Label3.Location = new System.Drawing.Point(14, 13);
			this.Label3.Name = "Label3";
			this.Label3.Size = new System.Drawing.Size(155, 12);
			this.Label3.TabIndex = 8;
			this.Label3.Text = "进行数据提取的Excel工作簿";
			//
			//Label4
			//
			this.Label4.AutoSize = true;
			this.Label4.Location = new System.Drawing.Point(10, 303);
			this.Label4.Name = "Label4";
			this.Label4.Size = new System.Drawing.Size(41, 12);
			this.Label4.TabIndex = 0;
			this.Label4.Text = "保存至";
			//
			//txtbxSavePath
			//
			this.txtbxSavePath.BackColor = System.Drawing.Color.White;
			this.txtbxSavePath.Location = new System.Drawing.Point(11, 321);
			this.txtbxSavePath.Margin = new System.Windows.Forms.Padding(0);
			this.txtbxSavePath.Name = "txtbxSavePath";
			this.txtbxSavePath.Size = new System.Drawing.Size(427, 21);
			this.txtbxSavePath.TabIndex = 1;
			//
			//ProgressBar1
			//
			this.ProgressBar1.Anchor = (System.Windows.Forms.AnchorStyles) ((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right);
			this.ProgressBar1.BackColor = System.Drawing.SystemColors.Control;
			this.ProgressBar1.Location = new System.Drawing.Point(0, 384);
			this.ProgressBar1.Name = "ProgressBar1";
			this.ProgressBar1.Size = new System.Drawing.Size(539, 8);
			this.ProgressBar1.TabIndex = 9;
			//
			//BackgroundWorker1
			//
			//
			//lbSheetName
			//
			this.lbSheetName.Anchor = (System.Windows.Forms.AnchorStyles) (System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left);
			this.lbSheetName.Location = new System.Drawing.Point(11, 356);
			this.lbSheetName.Name = "lbSheetName";
			this.lbSheetName.Size = new System.Drawing.Size(427, 25);
			this.lbSheetName.TabIndex = 10;
			this.lbSheetName.Text = ".";
			//
			//AddFileOrDirectoryFiles1
			//
			this.AddFileOrDirectoryFiles1.BackColor = System.Drawing.Color.Transparent;
			this.AddFileOrDirectoryFiles1.Location = new System.Drawing.Point(427, 38);
			this.AddFileOrDirectoryFiles1.Margin = new System.Windows.Forms.Padding(0);
			this.AddFileOrDirectoryFiles1.Name = "AddFileOrDirectoryFiles1";
			this.AddFileOrDirectoryFiles1.Size = new System.Drawing.Size(100, 68);
			this.AddFileOrDirectoryFiles1.TabIndex = 17;
			//
			//MyDataGridView1
			//
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
			this.MyDataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {this.SheetName, this.RangeName});
			this.MyDataGridView1.Location = new System.Drawing.Point(11, 181);
			this.MyDataGridView1.Name = "MyDataGridView1";
			this.MyDataGridView1.RowTemplate.Height = 23;
			this.MyDataGridView1.RowTemplate.Resizable = System.Windows.Forms.DataGridViewTriState.False;
			this.MyDataGridView1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
			this.MyDataGridView1.Size = new System.Drawing.Size(346, 110);
			this.MyDataGridView1.TabIndex = 14;
			//
			//SheetName
			//
			this.SheetName.HeaderText = "工作表名称";
			this.SheetName.Name = "SheetName";
			this.SheetName.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
			this.SheetName.ToolTipText = "提取的工作表名称包含于要进行检索的工作表名称，比如输入\"CX\"，则会提取工作簿中第一个名称中含有\"CX\"的工作表。";
			this.SheetName.Width = 183;
			//
			//RangeName
			//
			DataGridViewCellStyle2.ForeColor = System.Drawing.Color.Blue;
			this.RangeName.DefaultCellStyle = DataGridViewCellStyle2;
			this.RangeName.HeaderText = "区域范围";
			this.RangeName.Name = "RangeName";
			this.RangeName.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
			this.RangeName.ToolTipText = "示例： A1:C3";
			//
			//Panel1
			//
			this.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.Panel1.Controls.Add(this.Txtbox_DateFormat);
			this.Panel1.Controls.Add(this.btn_DateFormat);
			this.Panel1.Controls.Add(this.ChkboxParseDate);
			this.Panel1.Location = new System.Drawing.Point(366, 181);
			this.Panel1.Name = "Panel1";
			this.Panel1.Size = new System.Drawing.Size(161, 88);
			this.Panel1.TabIndex = 18;
			//
			//Txtbox_DateFormat
			//
			this.Txtbox_DateFormat.Enabled = false;
			this.Txtbox_DateFormat.Location = new System.Drawing.Point(13, 33);
			this.Txtbox_DateFormat.Name = "Txtbox_DateFormat";
			this.Txtbox_DateFormat.Size = new System.Drawing.Size(143, 21);
			this.Txtbox_DateFormat.TabIndex = 3;
			//
			//btn_DateFormat
			//
			this.btn_DateFormat.Enabled = false;
			this.btn_DateFormat.Location = new System.Drawing.Point(13, 60);
			this.btn_DateFormat.Name = "btn_DateFormat";
			this.btn_DateFormat.Size = new System.Drawing.Size(75, 23);
			this.btn_DateFormat.TabIndex = 2;
			this.btn_DateFormat.Text = "日期格式";
			this.btn_DateFormat.UseVisualStyleBackColor = true;
			//
			//ChkboxParseDate
			//
			this.ChkboxParseDate.AutoSize = true;
			this.ChkboxParseDate.Location = new System.Drawing.Point(3, 3);
			this.ChkboxParseDate.Name = "ChkboxParseDate";
			this.ChkboxParseDate.Size = new System.Drawing.Size(132, 16);
			this.ChkboxParseDate.TabIndex = 0;
			this.ChkboxParseDate.Text = "提取文件名中的日期";
			this.ChkboxParseDate.UseVisualStyleBackColor = true;
			//
			//ChkBxOpenExcelWhileFinished
			//
			this.ChkBxOpenExcelWhileFinished.AutoSize = true;
			this.ChkBxOpenExcelWhileFinished.Location = new System.Drawing.Point(366, 275);
			this.ChkBxOpenExcelWhileFinished.Name = "ChkBxOpenExcelWhileFinished";
			this.ChkBxOpenExcelWhileFinished.Size = new System.Drawing.Size(138, 16);
			this.ChkBxOpenExcelWhileFinished.TabIndex = 22;
			this.ChkBxOpenExcelWhileFinished.Text = "操作完成后打开Excel";
			this.ChkBxOpenExcelWhileFinished.UseVisualStyleBackColor = true;
			//
			//frmDeriveDataFromExcel
			//
			this.AutoScaleDimensions = new System.Drawing.SizeF((float) (6.0F), (float) (12.0F));
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(539, 396);
			this.Controls.Add(this.ChkBxOpenExcelWhileFinished);
			this.Controls.Add(this.Panel1);
			this.Controls.Add(this.AddFileOrDirectoryFiles1);
			this.Controls.Add(this.MyDataGridView1);
			this.Controls.Add(this.lbSheetName);
			this.Controls.Add(this.ProgressBar1);
			this.Controls.Add(this.Label3);
			this.Controls.Add(this.btnRemove);
			this.Controls.Add(this.ListBoxWorksheets);
			this.Controls.Add(this.BtnChoosePath);
			this.Controls.Add(this.btnExport);
			this.Controls.Add(this.txtbxSavePath);
			this.Controls.Add(this.Label4);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.Icon = (System.Drawing.Icon) (resources.GetObject("$this.Icon"));
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.Name = "frmDeriveDataFromExcel";
			this.Text = "从Excel提取数据";
			((System.ComponentModel.ISupportInitialize) this.MyDataGridView1).EndInit();
			this.Panel1.ResumeLayout(false);
			this.Panel1.PerformLayout();
			this.ResumeLayout(false);
			this.PerformLayout();
			
		}
		internal System.Windows.Forms.Button btnExport;
		internal System.Windows.Forms.SaveFileDialog SaveFileDialog1;
		internal System.Windows.Forms.OpenFileDialog OpenFileDialog1;
		internal System.Windows.Forms.Button BtnChoosePath;
		internal System.Windows.Forms.ListBox ListBoxWorksheets;
		internal System.Windows.Forms.Button btnRemove;
		internal System.Windows.Forms.FolderBrowserDialog FolderBrowserDialog1;
		internal System.Windows.Forms.Label Label3;
		internal System.Windows.Forms.Label Label4;
		internal System.Windows.Forms.TextBox txtbxSavePath;
		internal System.Windows.Forms.ProgressBar ProgressBar1;
		internal System.ComponentModel.BackgroundWorker BackgroundWorker1;
		internal System.Windows.Forms.Label lbSheetName;
		internal AME_UserControl.myDataGridView MyDataGridView1;
		internal AME_UserControl.AddFileOrDirectoryFiles AddFileOrDirectoryFiles1;
		internal System.Windows.Forms.DataGridViewTextBoxColumn SheetName;
		internal System.Windows.Forms.DataGridViewTextBoxColumn RangeName;
		internal System.Windows.Forms.Panel Panel1;
		internal System.Windows.Forms.TextBox Txtbox_DateFormat;
		internal System.Windows.Forms.Button btn_DateFormat;
		internal System.Windows.Forms.CheckBox ChkboxParseDate;
		internal System.Windows.Forms.CheckBox ChkBxOpenExcelWhileFinished;
		
	}
	
}
