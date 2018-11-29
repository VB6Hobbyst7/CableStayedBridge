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
	partial class DiaFrm_PointsTreeView : System.Windows.Forms.Form
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
			this.TreeViewPoints = new System.Windows.Forms.TreeView();
			base.Load += new System.EventHandler(FrmTreeView_Load);
			this.TreeViewPoints.AfterCheck += new System.Windows.Forms.TreeViewEventHandler(this.TreeViewPoints_AfterCheck);
			this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(FrmPointsTreeView_FormClosing);
			this.ListBoxChosenItems = new System.Windows.Forms.ListBox();
			this.ListBoxChosenItems.DataSourceChanged += new System.EventHandler(this.ListBoxChosenItems_DataSourceChanged);
			this.BtnOk = new System.Windows.Forms.Button();
			this.BtnOk.Click += new System.EventHandler(this.BtnOk_Click);
			this.BtnRemove = new System.Windows.Forms.Button();
			this.BtnRemove.Click += new System.EventHandler(this.BtnRemove_Click);
			this.BtnClear = new System.Windows.Forms.Button();
			this.BtnClear.Click += new System.EventHandler(this.BtnClear_Click);
			this.Label1 = new System.Windows.Forms.Label();
			this.Label2 = new System.Windows.Forms.Label();
			this.btnPreview = new System.Windows.Forms.Button();
			this.btnPreview.Click += new System.EventHandler(this.btnPreview_Click);
			this.BtnAdd = new System.Windows.Forms.Button();
			this.BtnAdd.Click += new System.EventHandler(this.BtnAdd_Click);
			this.BtnColpsAll = new System.Windows.Forms.Button();
			this.BtnColpsAll.Click += new System.EventHandler(this.BtnColpsAll_Click);
			this.btnClearCheckedNode = new System.Windows.Forms.Button();
			this.btnClearCheckedNode.Click += new System.EventHandler(this.ClearAllCheckedNode);
			this.SuspendLayout();
			//
			//TreeViewPoints
			//
			this.TreeViewPoints.CheckBoxes = true;
			this.TreeViewPoints.FullRowSelect = true;
			this.TreeViewPoints.Indent = 20;
			this.TreeViewPoints.ItemHeight = 14;
			this.TreeViewPoints.Location = new System.Drawing.Point(12, 37);
			this.TreeViewPoints.Name = "TreeViewPoints";
			this.TreeViewPoints.ShowLines = false;
			this.TreeViewPoints.Size = new System.Drawing.Size(210, 340);
			this.TreeViewPoints.TabIndex = 0;
			//
			//ListBoxChosenItems
			//
			this.ListBoxChosenItems.FormattingEnabled = true;
			this.ListBoxChosenItems.HorizontalScrollbar = true;
			this.ListBoxChosenItems.ItemHeight = 12;
			this.ListBoxChosenItems.Location = new System.Drawing.Point(309, 37);
			this.ListBoxChosenItems.Name = "ListBoxChosenItems";
			this.ListBoxChosenItems.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
			this.ListBoxChosenItems.Size = new System.Drawing.Size(214, 340);
			this.ListBoxChosenItems.TabIndex = 1;
			//
			//BtnOk
			//
			this.BtnOk.Location = new System.Drawing.Point(448, 392);
			this.BtnOk.Name = "BtnOk";
			this.BtnOk.Size = new System.Drawing.Size(75, 23);
			this.BtnOk.TabIndex = 0;
			this.BtnOk.Text = "确定";
			this.BtnOk.UseVisualStyleBackColor = true;
			//
			//BtnRemove
			//
			this.BtnRemove.Location = new System.Drawing.Point(228, 207);
			this.BtnRemove.Name = "BtnRemove";
			this.BtnRemove.Size = new System.Drawing.Size(75, 23);
			this.BtnRemove.TabIndex = 3;
			this.BtnRemove.Text = "<== 移除";
			this.BtnRemove.UseVisualStyleBackColor = true;
			//
			//BtnClear
			//
			this.BtnClear.Location = new System.Drawing.Point(228, 273);
			this.BtnClear.Name = "BtnClear";
			this.BtnClear.Size = new System.Drawing.Size(75, 23);
			this.BtnClear.TabIndex = 3;
			this.BtnClear.Text = "清空(&C)";
			this.BtnClear.UseVisualStyleBackColor = true;
			//
			//Label1
			//
			this.Label1.AutoSize = true;
			this.Label1.Location = new System.Drawing.Point(12, 13);
			this.Label1.Name = "Label1";
			this.Label1.Size = new System.Drawing.Size(113, 12);
			this.Label1.TabIndex = 4;
			this.Label1.Text = "选择单个或多个测点";
			//
			//Label2
			//
			this.Label2.AutoSize = true;
			this.Label2.Location = new System.Drawing.Point(307, 13);
			this.Label2.Name = "Label2";
			this.Label2.Size = new System.Drawing.Size(53, 12);
			this.Label2.TabIndex = 4;
			this.Label2.Text = "选择结果";
			//
			//btnPreview
			//
			this.btnPreview.Location = new System.Drawing.Point(313, 392);
			this.btnPreview.Name = "btnPreview";
			this.btnPreview.Size = new System.Drawing.Size(75, 23);
			this.btnPreview.TabIndex = 3;
			this.btnPreview.Text = "应用(&P)";
			this.btnPreview.UseVisualStyleBackColor = true;
			//
			//BtnAdd
			//
			this.BtnAdd.Location = new System.Drawing.Point(228, 159);
			this.BtnAdd.Name = "BtnAdd";
			this.BtnAdd.Size = new System.Drawing.Size(75, 23);
			this.BtnAdd.TabIndex = 3;
			this.BtnAdd.Text = "==> 更新";
			this.BtnAdd.UseVisualStyleBackColor = true;
			//
			//BtnColpsAll
			//
			this.BtnColpsAll.Location = new System.Drawing.Point(14, 392);
			this.BtnColpsAll.Name = "BtnColpsAll";
			this.BtnColpsAll.Size = new System.Drawing.Size(117, 23);
			this.BtnColpsAll.TabIndex = 3;
			this.BtnColpsAll.Text = "Collapse All(&C)";
			this.BtnColpsAll.UseVisualStyleBackColor = true;
			//
			//btnClearCheckedNode
			//
			this.btnClearCheckedNode.Location = new System.Drawing.Point(147, 392);
			this.btnClearCheckedNode.Name = "btnClearCheckedNode";
			this.btnClearCheckedNode.Size = new System.Drawing.Size(75, 23);
			this.btnClearCheckedNode.TabIndex = 5;
			this.btnClearCheckedNode.Text = "清空选择";
			this.btnClearCheckedNode.UseVisualStyleBackColor = true;
			//
			//DiaFrm_PointsTreeView
			//
			this.AcceptButton = this.BtnOk;
			this.AutoScaleDimensions = new System.Drawing.SizeF((float) (6.0F), (float) (12.0F));
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(539, 427);
			this.Controls.Add(this.btnClearCheckedNode);
			this.Controls.Add(this.Label2);
			this.Controls.Add(this.Label1);
			this.Controls.Add(this.BtnColpsAll);
			this.Controls.Add(this.btnPreview);
			this.Controls.Add(this.BtnClear);
			this.Controls.Add(this.BtnAdd);
			this.Controls.Add(this.BtnRemove);
			this.Controls.Add(this.BtnOk);
			this.Controls.Add(this.ListBoxChosenItems);
			this.Controls.Add(this.TreeViewPoints);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.Name = "DiaFrm_PointsTreeView";
			this.Text = "选择测点标志";
			this.ResumeLayout(false);
			this.PerformLayout();
			
		}
		internal System.Windows.Forms.TreeView TreeViewPoints;
		internal System.Windows.Forms.ListBox ListBoxChosenItems;
		internal System.Windows.Forms.Button BtnOk;
		internal System.Windows.Forms.Button BtnRemove;
		internal System.Windows.Forms.Button BtnClear;
		internal System.Windows.Forms.Label Label1;
		internal System.Windows.Forms.Label Label2;
		internal System.Windows.Forms.Button btnPreview;
		internal System.Windows.Forms.Button BtnAdd;
		internal System.Windows.Forms.Button BtnColpsAll;
		internal System.Windows.Forms.Button btnClearCheckedNode;
	}
	
}
