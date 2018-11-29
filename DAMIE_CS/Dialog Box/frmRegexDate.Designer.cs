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
	partial class frmRegexDate : System.Windows.Forms.Form
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
			this.TextBox1 = new System.Windows.Forms.TextBox();
			this.TextBox1.TextChanged += new System.EventHandler(this.GenerateRegex);
			this.Button1 = new System.Windows.Forms.Button();
			this.Button1.Click += new System.EventHandler(this.SetComponent);
			this.Button2 = new System.Windows.Forms.Button();
			this.Button2.Click += new System.EventHandler(this.SetComponent);
			this.Button3 = new System.Windows.Forms.Button();
			this.Button3.Click += new System.EventHandler(this.SetComponent);
			this.TextBox2 = new System.Windows.Forms.TextBox();
			this.TextBox2.TextChanged += new System.EventHandler(this.GenerateRegex);
			this.Button4 = new System.Windows.Forms.Button();
			this.Button4.Click += new System.EventHandler(this.SetComponent);
			this.TextBox3 = new System.Windows.Forms.TextBox();
			this.TextBox3.TextChanged += new System.EventHandler(this.GenerateRegex);
			this.BtnOk = new System.Windows.Forms.Button();
			this.BtnOk.Click += new System.EventHandler(this.BtnOk_Click);
			this.TextBox4 = new System.Windows.Forms.TextBox();
			this.TextBox4.TextChanged += new System.EventHandler(this.GenerateRegex);
			this.BtnCancel = new System.Windows.Forms.Button();
			this.BtnCancel.Click += new System.EventHandler(this.Button6_Click);
			this.TextBox5 = new System.Windows.Forms.TextBox();
			this.TextBox5.TextChanged += new System.EventHandler(this.GenerateRegex);
			this.Panel1 = new System.Windows.Forms.Panel();
			this.Label2 = new System.Windows.Forms.Label();
			this.Textbox_btn4 = new System.Windows.Forms.TextBox();
			this.Textbox_btn4.MouseClick += new System.Windows.Forms.MouseEventHandler(this.selectText);
			this.Textbox_btn4.TextChanged += new System.EventHandler(this.GenerateRegex);
			this.Textbox_btn3 = new System.Windows.Forms.TextBox();
			this.Textbox_btn3.MouseClick += new System.Windows.Forms.MouseEventHandler(this.selectText);
			this.Textbox_btn3.TextChanged += new System.EventHandler(this.GenerateRegex);
			this.Textbox_btn2 = new System.Windows.Forms.TextBox();
			this.Textbox_btn2.MouseClick += new System.Windows.Forms.MouseEventHandler(this.selectText);
			this.Textbox_btn2.TextChanged += new System.EventHandler(this.GenerateRegex);
			this.Textbox_btn1 = new System.Windows.Forms.TextBox();
			this.Textbox_btn1.MouseClick += new System.Windows.Forms.MouseEventHandler(this.selectText);
			this.Textbox_btn1.TextChanged += new System.EventHandler(this.GenerateRegex);
			this.Label_Regex = new System.Windows.Forms.Label();
			this.ToolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.Panel1.SuspendLayout();
			this.SuspendLayout();
			//
			//TextBox1
			//
			this.TextBox1.Location = new System.Drawing.Point(3, 13);
			this.TextBox1.Name = "TextBox1";
			this.TextBox1.Size = new System.Drawing.Size(102, 21);
			this.TextBox1.TabIndex = 0;
			//
			//Button1
			//
			this.Button1.Location = new System.Drawing.Point(113, 35);
			this.Button1.Name = "Button1";
			this.Button1.Size = new System.Drawing.Size(37, 23);
			this.Button1.TabIndex = 1;
			this.Button1.Text = "序号";
			this.Button1.UseVisualStyleBackColor = true;
			//
			//Button2
			//
			this.Button2.Location = new System.Drawing.Point(113, 125);
			this.Button2.Name = "Button2";
			this.Button2.Size = new System.Drawing.Size(37, 23);
			this.Button2.TabIndex = 1;
			this.Button2.Text = "月";
			this.Button2.UseVisualStyleBackColor = true;
			//
			//Button3
			//
			this.Button3.Location = new System.Drawing.Point(113, 170);
			this.Button3.Name = "Button3";
			this.Button3.Size = new System.Drawing.Size(37, 23);
			this.Button3.TabIndex = 1;
			this.Button3.Text = "日";
			this.Button3.UseVisualStyleBackColor = true;
			//
			//TextBox2
			//
			this.TextBox2.Location = new System.Drawing.Point(3, 58);
			this.TextBox2.Name = "TextBox2";
			this.TextBox2.Size = new System.Drawing.Size(102, 21);
			this.TextBox2.TabIndex = 0;
			this.TextBox2.Text = "-";
			//
			//Button4
			//
			this.Button4.Location = new System.Drawing.Point(113, 80);
			this.Button4.Name = "Button4";
			this.Button4.Size = new System.Drawing.Size(37, 23);
			this.Button4.TabIndex = 1;
			this.Button4.Text = "年";
			this.Button4.UseVisualStyleBackColor = true;
			//
			//TextBox3
			//
			this.TextBox3.Location = new System.Drawing.Point(3, 103);
			this.TextBox3.Name = "TextBox3";
			this.TextBox3.Size = new System.Drawing.Size(102, 21);
			this.TextBox3.TabIndex = 0;
			this.TextBox3.Text = "-";
			//
			//BtnOk
			//
			this.BtnOk.Anchor = (System.Windows.Forms.AnchorStyles) (System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right);
			this.BtnOk.Location = new System.Drawing.Point(240, 12);
			this.BtnOk.Name = "BtnOk";
			this.BtnOk.Size = new System.Drawing.Size(75, 23);
			this.BtnOk.TabIndex = 1;
			this.BtnOk.Text = "确定";
			this.BtnOk.UseVisualStyleBackColor = true;
			//
			//TextBox4
			//
			this.TextBox4.Location = new System.Drawing.Point(3, 148);
			this.TextBox4.Name = "TextBox4";
			this.TextBox4.Size = new System.Drawing.Size(102, 21);
			this.TextBox4.TabIndex = 0;
			this.TextBox4.Text = "-";
			//
			//BtnCancel
			//
			this.BtnCancel.Anchor = (System.Windows.Forms.AnchorStyles) (System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right);
			this.BtnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.BtnCancel.Location = new System.Drawing.Point(240, 41);
			this.BtnCancel.Name = "BtnCancel";
			this.BtnCancel.Size = new System.Drawing.Size(75, 23);
			this.BtnCancel.TabIndex = 1;
			this.BtnCancel.Text = "取消";
			this.BtnCancel.UseVisualStyleBackColor = true;
			//
			//TextBox5
			//
			this.TextBox5.Location = new System.Drawing.Point(3, 193);
			this.TextBox5.Name = "TextBox5";
			this.TextBox5.Size = new System.Drawing.Size(102, 21);
			this.TextBox5.TabIndex = 0;
			//
			//Panel1
			//
			this.Panel1.Controls.Add(this.Label2);
			this.Panel1.Controls.Add(this.TextBox1);
			this.Panel1.Controls.Add(this.Button3);
			this.Panel1.Controls.Add(this.TextBox2);
			this.Panel1.Controls.Add(this.Button2);
			this.Panel1.Controls.Add(this.TextBox3);
			this.Panel1.Controls.Add(this.Textbox_btn4);
			this.Panel1.Controls.Add(this.TextBox4);
			this.Panel1.Controls.Add(this.Textbox_btn3);
			this.Panel1.Controls.Add(this.TextBox5);
			this.Panel1.Controls.Add(this.Textbox_btn2);
			this.Panel1.Controls.Add(this.Button4);
			this.Panel1.Controls.Add(this.Textbox_btn1);
			this.Panel1.Controls.Add(this.Button1);
			this.Panel1.Location = new System.Drawing.Point(12, 12);
			this.Panel1.Name = "Panel1";
			this.Panel1.Size = new System.Drawing.Size(214, 223);
			this.Panel1.TabIndex = 2;
			//
			//Label2
			//
			this.Label2.AutoSize = true;
			this.Label2.Location = new System.Drawing.Point(149, 11);
			this.Label2.Name = "Label2";
			this.Label2.Size = new System.Drawing.Size(53, 12);
			this.Label2.TabIndex = 2;
			this.Label2.Text = "数值个数";
			this.ToolTip1.SetToolTip(this.Label2, "\"序号\"的数值指的是最多的数值个数，\"年、月、日\"的数值指的是精确的数值个数。");
			//
			//Textbox_btn4
			//
			this.Textbox_btn4.Location = new System.Drawing.Point(156, 170);
			this.Textbox_btn4.Name = "Textbox_btn4";
			this.Textbox_btn4.Size = new System.Drawing.Size(40, 21);
			this.Textbox_btn4.TabIndex = 104;
			this.Textbox_btn4.Text = "2";
			this.Textbox_btn4.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			//
			//Textbox_btn3
			//
			this.Textbox_btn3.Location = new System.Drawing.Point(156, 125);
			this.Textbox_btn3.Name = "Textbox_btn3";
			this.Textbox_btn3.Size = new System.Drawing.Size(40, 21);
			this.Textbox_btn3.TabIndex = 103;
			this.Textbox_btn3.Text = "2";
			this.Textbox_btn3.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			//
			//Textbox_btn2
			//
			this.Textbox_btn2.Location = new System.Drawing.Point(156, 80);
			this.Textbox_btn2.Name = "Textbox_btn2";
			this.Textbox_btn2.Size = new System.Drawing.Size(40, 21);
			this.Textbox_btn2.TabIndex = 102;
			this.Textbox_btn2.Text = "4";
			this.Textbox_btn2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			//
			//Textbox_btn1
			//
			this.Textbox_btn1.Location = new System.Drawing.Point(156, 35);
			this.Textbox_btn1.Name = "Textbox_btn1";
			this.Textbox_btn1.Size = new System.Drawing.Size(40, 21);
			this.Textbox_btn1.TabIndex = 101;
			this.Textbox_btn1.Text = "3";
			this.Textbox_btn1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			//
			//Label_Regex
			//
			this.Label_Regex.Anchor = (System.Windows.Forms.AnchorStyles) (((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right);
			this.Label_Regex.BackColor = System.Drawing.Color.FromArgb(System.Convert.ToInt32(System.Convert.ToByte(224)), System.Convert.ToInt32(System.Convert.ToByte(224)), System.Convert.ToInt32(System.Convert.ToByte(224)));
			this.Label_Regex.Location = new System.Drawing.Point(12, 238);
			this.Label_Regex.Name = "Label_Regex";
			this.Label_Regex.Size = new System.Drawing.Size(301, 63);
			this.Label_Regex.TabIndex = 3;
			//
			//frmRegexDate
			//
			this.AutoScaleDimensions = new System.Drawing.SizeF((float) (6.0F), (float) (12.0F));
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.CancelButton = this.BtnCancel;
			this.ClientSize = new System.Drawing.Size(324, 310);
			this.Controls.Add(this.Label_Regex);
			this.Controls.Add(this.Panel1);
			this.Controls.Add(this.BtnOk);
			this.Controls.Add(this.BtnCancel);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmRegexDate";
			this.Text = "frmRegexDate";
			this.Panel1.ResumeLayout(false);
			this.Panel1.PerformLayout();
			this.ResumeLayout(false);
			
		}
		internal System.Windows.Forms.TextBox TextBox1;
		internal System.Windows.Forms.Button Button1;
		internal System.Windows.Forms.Button Button2;
		internal System.Windows.Forms.Button Button3;
		internal System.Windows.Forms.TextBox TextBox2;
		internal System.Windows.Forms.Button Button4;
		internal System.Windows.Forms.TextBox TextBox3;
		internal System.Windows.Forms.Button BtnOk;
		internal System.Windows.Forms.TextBox TextBox4;
		internal System.Windows.Forms.Button BtnCancel;
		internal System.Windows.Forms.TextBox TextBox5;
		internal System.Windows.Forms.Panel Panel1;
		internal System.Windows.Forms.TextBox Textbox_btn4;
		internal System.Windows.Forms.TextBox Textbox_btn3;
		internal System.Windows.Forms.TextBox Textbox_btn2;
		internal System.Windows.Forms.TextBox Textbox_btn1;
		internal System.Windows.Forms.Label Label_Regex;
		internal System.Windows.Forms.Label Label2;
		internal System.Windows.Forms.ToolTip ToolTip1;
	}
	
}
