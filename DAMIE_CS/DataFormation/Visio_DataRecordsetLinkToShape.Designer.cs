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
	partial class Visio_DataRecordsetLinkToShape : System.Windows.Forms.Form
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
			base.Load += new System.EventHandler(Visio_DataRecordsetLinkToShape_Load);
			this.vsoDocumentChanged += new Visio_DataRecordsetLinkToShape.Action(DocumentChanged);
			this.ShapeIDValidated += new Visio_DataRecordsetLinkToShape.Action(Visio_DataRecordsetLinkToShape_ShapeIDValidated);
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Visio_DataRecordsetLinkToShape));
			this.BtnChooseVsoDoc = new System.Windows.Forms.Button();
			this.BtnChooseVsoDoc.Click += new System.EventHandler(this.BtnChooseVsoDoc_Click);
			this.txtbxVsoDoc = new System.Windows.Forms.TextBox();
			this.Label3 = new System.Windows.Forms.Label();
			this.OpenFileDialog1 = new System.Windows.Forms.OpenFileDialog();
			this.Label5 = new System.Windows.Forms.Label();
			this.Label6 = new System.Windows.Forms.Label();
			this.btnLink = new System.Windows.Forms.Button();
			this.btnLink.Click += new System.EventHandler(this.btnLink_Click);
			this.ComboBox_Page = new System.Windows.Forms.ComboBox();
			this.ComboBox_Page.SelectedIndexChanged += new System.EventHandler(this.ComboBox_Page_SelectedIndexChanged);
			this.ComboBox_DataRs = new System.Windows.Forms.ComboBox();
			this.ComboBox_DataRs.SelectedIndexChanged += new System.EventHandler(this.ComboBox_DataRs_SelectedIndexChanged);
			this.btnValidate = new System.Windows.Forms.Button();
			this.btnValidate.Click += new System.EventHandler(this.btnValidate_Click);
			this.ToolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.Label1 = new System.Windows.Forms.Label();
			this.ComboBox_Column_ShapeID = new System.Windows.Forms.ComboBox();
			this.ComboBox_Column_ShapeID.SelectedIndexChanged += new System.EventHandler(this.ComboBox_Column_ShapeID_SelectedIndexChanged);
			this.SuspendLayout();
			//
			//BtnChooseVsoDoc
			//
			this.BtnChooseVsoDoc.Anchor = (System.Windows.Forms.AnchorStyles) (System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right);
			this.BtnChooseVsoDoc.BackColor = System.Drawing.SystemColors.Control;
			this.BtnChooseVsoDoc.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.BtnChooseVsoDoc.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
			this.BtnChooseVsoDoc.ForeColor = System.Drawing.SystemColors.InfoText;
			this.BtnChooseVsoDoc.Location = new System.Drawing.Point(319, 9);
			this.BtnChooseVsoDoc.Name = "BtnChooseVsoDoc";
			this.BtnChooseVsoDoc.Size = new System.Drawing.Size(74, 23);
			this.BtnChooseVsoDoc.TabIndex = 5;
			this.BtnChooseVsoDoc.Text = "选择...";
			this.BtnChooseVsoDoc.UseVisualStyleBackColor = false;
			//
			//txtbxVsoDoc
			//
			this.txtbxVsoDoc.Anchor = (System.Windows.Forms.AnchorStyles) ((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right);
			this.txtbxVsoDoc.BackColor = System.Drawing.Color.White;
			this.txtbxVsoDoc.Location = new System.Drawing.Point(86, 9);
			this.txtbxVsoDoc.Margin = new System.Windows.Forms.Padding(0);
			this.txtbxVsoDoc.Name = "txtbxVsoDoc";
			this.txtbxVsoDoc.Size = new System.Drawing.Size(222, 21);
			this.txtbxVsoDoc.TabIndex = 4;
			//
			//Label3
			//
			this.Label3.AutoSize = true;
			this.Label3.Location = new System.Drawing.Point(11, 12);
			this.Label3.Name = "Label3";
			this.Label3.Size = new System.Drawing.Size(59, 12);
			this.Label3.TabIndex = 6;
			this.Label3.Text = "Visio绘图";
			//
			//OpenFileDialog1
			//
			this.OpenFileDialog1.FileName = "OpenFileDialog1";
			//
			//Label5
			//
			this.Label5.AutoSize = true;
			this.Label5.Location = new System.Drawing.Point(5, 51);
			this.Label5.Name = "Label5";
			this.Label5.Size = new System.Drawing.Size(65, 12);
			this.Label5.TabIndex = 13;
			this.Label5.Text = "数据记录集";
			//
			//Label6
			//
			this.Label6.AutoSize = true;
			this.Label6.Location = new System.Drawing.Point(213, 51);
			this.Label6.Name = "Label6";
			this.Label6.Size = new System.Drawing.Size(53, 12);
			this.Label6.TabIndex = 13;
			this.Label6.Text = "绘图页面";
			//
			//btnLink
			//
			this.btnLink.Location = new System.Drawing.Point(318, 91);
			this.btnLink.Name = "btnLink";
			this.btnLink.Size = new System.Drawing.Size(75, 23);
			this.btnLink.TabIndex = 0;
			this.btnLink.Text = "链接";
			this.ToolTip1.SetToolTip(this.btnLink, "以数据记录集中的形状ID为指标，将其链接到页面的对应ID的形状上。");
			this.btnLink.UseVisualStyleBackColor = true;
			//
			//ComboBox_Page
			//
			this.ComboBox_Page.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.ComboBox_Page.FormattingEnabled = true;
			this.ComboBox_Page.Location = new System.Drawing.Point(272, 48);
			this.ComboBox_Page.Name = "ComboBox_Page";
			this.ComboBox_Page.Size = new System.Drawing.Size(121, 20);
			this.ComboBox_Page.TabIndex = 12;
			//
			//ComboBox_DataRs
			//
			this.ComboBox_DataRs.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.ComboBox_DataRs.FormattingEnabled = true;
			this.ComboBox_DataRs.Location = new System.Drawing.Point(75, 48);
			this.ComboBox_DataRs.Name = "ComboBox_DataRs";
			this.ComboBox_DataRs.Size = new System.Drawing.Size(121, 20);
			this.ComboBox_DataRs.TabIndex = 12;
			//
			//btnValidate
			//
			this.btnValidate.Location = new System.Drawing.Point(232, 91);
			this.btnValidate.Name = "btnValidate";
			this.btnValidate.Size = new System.Drawing.Size(75, 23);
			this.btnValidate.TabIndex = 14;
			this.btnValidate.Text = "验证";
			this.ToolTip1.SetToolTip(this.btnValidate, "验证页面的形状中是否有对应的形状ID。");
			this.btnValidate.UseVisualStyleBackColor = true;
			//
			//Label1
			//
			this.Label1.Location = new System.Drawing.Point(5, 91);
			this.Label1.Name = "Label1";
			this.Label1.Size = new System.Drawing.Size(65, 29);
			this.Label1.TabIndex = 15;
			this.Label1.Text = "形状ID" + System.Convert.ToString(global::Microsoft.VisualBasic.Strings.ChrW(13)) + System.Convert.ToString(global::Microsoft.VisualBasic.Strings.ChrW(10)) + "所在的字段";
			//
			//ComboBox_Column_ShapeID
			//
			this.ComboBox_Column_ShapeID.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.ComboBox_Column_ShapeID.FormattingEnabled = true;
			this.ComboBox_Column_ShapeID.Location = new System.Drawing.Point(75, 91);
			this.ComboBox_Column_ShapeID.Name = "ComboBox_Column_ShapeID";
			this.ComboBox_Column_ShapeID.Size = new System.Drawing.Size(121, 20);
			this.ComboBox_Column_ShapeID.TabIndex = 12;
			//
			//Visio_DataRecordsetLinkToShape
			//
			this.AutoScaleDimensions = new System.Drawing.SizeF((float) (6.0F), (float) (12.0F));
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(406, 125);
			this.Controls.Add(this.Label1);
			this.Controls.Add(this.btnValidate);
			this.Controls.Add(this.Label5);
			this.Controls.Add(this.Label6);
			this.Controls.Add(this.Label3);
			this.Controls.Add(this.btnLink);
			this.Controls.Add(this.BtnChooseVsoDoc);
			this.Controls.Add(this.ComboBox_Column_ShapeID);
			this.Controls.Add(this.ComboBox_Page);
			this.Controls.Add(this.txtbxVsoDoc);
			this.Controls.Add(this.ComboBox_DataRs);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.Icon = (System.Drawing.Icon) (resources.GetObject("$this.Icon"));
			this.Name = "Visio_DataRecordsetLinkToShape";
			this.Text = "数据记录集链接到形状";
			this.ResumeLayout(false);
			this.PerformLayout();
			
		}
		internal System.Windows.Forms.Button BtnChooseVsoDoc;
		internal System.Windows.Forms.TextBox txtbxVsoDoc;
		internal System.Windows.Forms.Label Label3;
		internal System.Windows.Forms.OpenFileDialog OpenFileDialog1;
		internal System.Windows.Forms.Label Label5;
		internal System.Windows.Forms.Label Label6;
		internal System.Windows.Forms.Button btnLink;
		internal System.Windows.Forms.ComboBox ComboBox_Page;
		internal System.Windows.Forms.ComboBox ComboBox_DataRs;
		internal System.Windows.Forms.ToolTip ToolTip1;
		internal System.Windows.Forms.Button btnValidate;
		internal System.Windows.Forms.Label Label1;
		internal System.Windows.Forms.ComboBox ComboBox_Column_ShapeID;
	}
	
}
