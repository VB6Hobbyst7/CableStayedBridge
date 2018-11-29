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
	partial class frmLockCurves : System.Windows.Forms.Form
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
			this.btn_OK = new System.Windows.Forms.Button();
			this.btn_OK.Click += new System.EventHandler(this.btn_OK_Click);
			this.btn_Cancel = new System.Windows.Forms.Button();
			this.btn_Cancel.Click += new System.EventHandler(this.btn_Cancel_Click);
			this.btnClear = new System.Windows.Forms.Button();
			this.btnClear.Click += new System.EventHandler(this.btnClear_Click);
			this.MyDataGridView1 = new myDataGridView();
			this.MyDataGridView1.RowsAdded += new System.Windows.Forms.DataGridViewRowsAddedEventHandler(this.MyDataGridView1_RowsAdded);
			this.MyDataGridView1.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.MyDataGridView1_CellValueChanged);
			this.CurveDate = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.CurveHandle = new System.Windows.Forms.DataGridViewComboBoxColumn();
			((System.ComponentModel.ISupportInitialize) this.MyDataGridView1).BeginInit();
			this.SuspendLayout();
			//
			//btn_OK
			//
			this.btn_OK.Anchor = (System.Windows.Forms.AnchorStyles) (System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right);
			this.btn_OK.Location = new System.Drawing.Point(298, 12);
			this.btn_OK.Name = "btn_OK";
			this.btn_OK.Size = new System.Drawing.Size(75, 23);
			this.btn_OK.TabIndex = 1;
			this.btn_OK.Text = "确定";
			this.btn_OK.UseVisualStyleBackColor = true;
			//
			//btn_Cancel
			//
			this.btn_Cancel.Anchor = (System.Windows.Forms.AnchorStyles) (System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right);
			this.btn_Cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btn_Cancel.Location = new System.Drawing.Point(298, 70);
			this.btn_Cancel.Name = "btn_Cancel";
			this.btn_Cancel.Size = new System.Drawing.Size(75, 23);
			this.btn_Cancel.TabIndex = 1;
			this.btn_Cancel.Text = "取消";
			this.btn_Cancel.UseVisualStyleBackColor = true;
			//
			//btnClear
			//
			this.btnClear.Anchor = (System.Windows.Forms.AnchorStyles) (System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right);
			this.btnClear.Location = new System.Drawing.Point(298, 41);
			this.btnClear.Name = "btnClear";
			this.btnClear.Size = new System.Drawing.Size(75, 23);
			this.btnClear.TabIndex = 2;
			this.btnClear.Text = "清空";
			this.btnClear.UseVisualStyleBackColor = true;
			//
			//MyDataGridView1
			//
			this.MyDataGridView1.Anchor = (System.Windows.Forms.AnchorStyles) (((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right);
			this.MyDataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
			this.MyDataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {this.CurveDate, this.CurveHandle});
			this.MyDataGridView1.Location = new System.Drawing.Point(12, 12);
			this.MyDataGridView1.Name = "MyDataGridView1";
			this.MyDataGridView1.RowTemplate.Height = 23;
			this.MyDataGridView1.RowTemplate.Resizable = System.Windows.Forms.DataGridViewTriState.False;
			this.MyDataGridView1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
			this.MyDataGridView1.Size = new System.Drawing.Size(270, 235);
			this.MyDataGridView1.TabIndex = 0;
			//
			//CurveDate
			//
			this.CurveDate.HeaderText = "     日期";
			this.CurveDate.Name = "CurveDate";
			//
			//CurveHandle
			//
			this.CurveHandle.HeaderText = "     操作";
			this.CurveHandle.Items.AddRange(new object[] {"锁定", "删除"});
			this.CurveHandle.Name = "CurveHandle";
			//
			//frmLockCurves
			//
			this.AutoScaleDimensions = new System.Drawing.SizeF((float) (6.0F), (float) (12.0F));
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.CancelButton = this.btn_Cancel;
			this.ClientSize = new System.Drawing.Size(385, 259);
			this.Controls.Add(this.btnClear);
			this.Controls.Add(this.btn_Cancel);
			this.Controls.Add(this.btn_OK);
			this.Controls.Add(this.MyDataGridView1);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmLockCurves";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "批量操作绘图曲线";
			((System.ComponentModel.ISupportInitialize) this.MyDataGridView1).EndInit();
			this.ResumeLayout(false);
			
		}
		internal myDataGridView MyDataGridView1;
		internal System.Windows.Forms.Button btn_OK;
		internal System.Windows.Forms.Button btn_Cancel;
		internal System.Windows.Forms.DataGridViewTextBoxColumn CurveDate;
		internal System.Windows.Forms.DataGridViewComboBoxColumn CurveHandle;
		internal System.Windows.Forms.Button btnClear;
	}
	
}
