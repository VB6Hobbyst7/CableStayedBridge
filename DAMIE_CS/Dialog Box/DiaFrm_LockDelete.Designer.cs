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
	namespace Miscellaneous
	{
		
		[global::Microsoft.VisualBasic.CompilerServices.DesignerGenerated()]public 
		partial class DiaFrm_LockDelete : System.Windows.Forms.Form
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
				this.btn1 = new System.Windows.Forms.Button();
				this.btn1.Click += new System.EventHandler(this.btnLock_Click);
				this.btn2 = new System.Windows.Forms.Button();
				this.btn2.Click += new System.EventHandler(this.btnDelete_Click);
				this.LabelPrompt = new System.Windows.Forms.Label();
				this.LabelPrompt.SizeChanged += new System.EventHandler(this.LabelPrompt_SizeChanged);
				this.btn3 = new System.Windows.Forms.Button();
				this.btn3.Click += new System.EventHandler(this.btnIgnore_Click);
				this.SuspendLayout();
				//
				//btn1
				//
				this.btn1.Anchor = (System.Windows.Forms.AnchorStyles) (System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left);
				this.btn1.Location = new System.Drawing.Point(12, 38);
				this.btn1.Name = "btn1";
				this.btn1.Size = new System.Drawing.Size(75, 23);
				this.btn1.TabIndex = 0;
				this.btn1.Text = "Lock";
				this.btn1.UseVisualStyleBackColor = true;
				//
				//btn2
				//
				this.btn2.Anchor = (System.Windows.Forms.AnchorStyles) (System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left);
				this.btn2.Location = new System.Drawing.Point(116, 38);
				this.btn2.Name = "btn2";
				this.btn2.Size = new System.Drawing.Size(75, 23);
				this.btn2.TabIndex = 1;
				this.btn2.Text = "Delete";
				this.btn2.UseVisualStyleBackColor = true;
				//
				//LabelPrompt
				//
				this.LabelPrompt.AutoEllipsis = true;
				this.LabelPrompt.AutoSize = true;
				this.LabelPrompt.Location = new System.Drawing.Point(12, 9);
				this.LabelPrompt.Name = "LabelPrompt";
				this.LabelPrompt.Size = new System.Drawing.Size(41, 12);
				this.LabelPrompt.TabIndex = 3;
				this.LabelPrompt.Text = "PROMPT";
				this.LabelPrompt.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
				//
				//btn3
				//
				this.btn3.Anchor = (System.Windows.Forms.AnchorStyles) (System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left);
				this.btn3.DialogResult = System.Windows.Forms.DialogResult.Cancel;
				this.btn3.Location = new System.Drawing.Point(220, 38);
				this.btn3.Name = "btn3";
				this.btn3.Size = new System.Drawing.Size(75, 23);
				this.btn3.TabIndex = 2;
				this.btn3.Text = "Ignore";
				this.btn3.UseVisualStyleBackColor = true;
				//
				//DiaFrm_LockDelete
				//
				this.AutoScaleDimensions = new System.Drawing.SizeF((float) (6.0F), (float) (12.0F));
				this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
				this.CancelButton = this.btn3;
				this.ClientSize = new System.Drawing.Size(307, 70);
				this.Controls.Add(this.LabelPrompt);
				this.Controls.Add(this.btn3);
				this.Controls.Add(this.btn2);
				this.Controls.Add(this.btn1);
				this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
				this.MaximizeBox = false;
				this.MinimizeBox = false;
				this.Name = "DiaFrm_LockDelete";
				this.Text = "Operate on series";
				this.TopMost = true;
				this.ResumeLayout(false);
				this.PerformLayout();
				
			}
			internal System.Windows.Forms.Button btn1;
			internal System.Windows.Forms.Button btn2;
			internal System.Windows.Forms.Label LabelPrompt;
			internal System.Windows.Forms.Button btn3;
		}
		
		
	}
	
}
