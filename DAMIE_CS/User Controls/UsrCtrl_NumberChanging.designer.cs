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
	namespace AME_UserControl
	{
		[global::Microsoft.VisualBasic.CompilerServices.DesignerGenerated()]public 
		partial class UsrCtrl_NumberChanging : System.Windows.Forms.UserControl
		{
			
			//UserControl overrides dispose to clean up the component list.
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
				this.btnNext = new System.Windows.Forms.Button();
				this.Load += new System.EventHandler(NumberChanging_Load);
				this.btnNext.Click += new System.EventHandler(this.btnNext_Click);
				this.TextBoxNumber = new System.Windows.Forms.TextBox();
				this.TextBoxNumber.KeyUp += new System.Windows.Forms.KeyEventHandler(this.TextBoxNumber_KeyDown);
				this.btnPrevious = new System.Windows.Forms.Button();
				this.btnPrevious.Click += new System.EventHandler(this.btnPrevious_Click);
				this.cbUnit = new System.Windows.Forms.ComboBox();
				this.cbUnit.SelectedIndexChanged += new System.EventHandler(this.btnUnit_SelectedIndexChanged);
				this.SuspendLayout();
				//
				//btnNext
				//
				this.btnNext.BackColor = System.Drawing.SystemColors.ButtonFace;
				this.btnNext.Location = new System.Drawing.Point(148, -1);
				this.btnNext.Name = "btnNext";
				this.btnNext.Size = new System.Drawing.Size(40, 21);
				this.btnNext.TabIndex = 2;
				this.btnNext.Text = "->";
				this.btnNext.UseVisualStyleBackColor = false;
				//
				//TextBoxNumber
				//
				this.TextBoxNumber.BackColor = System.Drawing.Color.White;
				this.TextBoxNumber.BorderStyle = System.Windows.Forms.BorderStyle.None;
				this.TextBoxNumber.Font = new System.Drawing.Font("SimSun", (float) (10.5F), System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, System.Convert.ToByte(134));
				this.TextBoxNumber.Location = new System.Drawing.Point(44, 2);
				this.TextBoxNumber.Name = "TextBoxNumber";
				this.TextBoxNumber.Size = new System.Drawing.Size(35, 16);
				this.TextBoxNumber.TabIndex = 0;
				//
				//btnPrevious
				//
				this.btnPrevious.BackColor = System.Drawing.SystemColors.ButtonFace;
				this.btnPrevious.Location = new System.Drawing.Point(-1, -1);
				this.btnPrevious.Name = "btnPrevious";
				this.btnPrevious.Size = new System.Drawing.Size(40, 21);
				this.btnPrevious.TabIndex = 1;
				this.btnPrevious.Text = "<-";
				this.btnPrevious.UseVisualStyleBackColor = false;
				//
				//cbUnit
				//
				this.cbUnit.BackColor = System.Drawing.SystemColors.Control;
				this.cbUnit.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
				this.cbUnit.FormattingEnabled = true;
				this.cbUnit.Location = new System.Drawing.Point(84, 0);
				this.cbUnit.Name = "cbUnit";
				this.cbUnit.Size = new System.Drawing.Size(60, 20);
				this.cbUnit.TabIndex = 3;
				//
				//UsrCtrl_NumberChanging
				//
				this.AutoScaleDimensions = new System.Drawing.SizeF((float) (6.0F), (float) (12.0F));
				this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
				this.BackColor = System.Drawing.Color.Transparent;
				this.Controls.Add(this.cbUnit);
				this.Controls.Add(this.btnPrevious);
				this.Controls.Add(this.btnNext);
				this.Controls.Add(this.TextBoxNumber);
				this.Name = "UsrCtrl_NumberChanging";
				this.Size = new System.Drawing.Size(190, 21);
				this.ResumeLayout(false);
				this.PerformLayout();
				
			}
			private System.Windows.Forms.Button btnNext;
			private System.Windows.Forms.Button btnPrevious;
			private System.Windows.Forms.TextBox TextBoxNumber;
			private System.Windows.Forms.ComboBox cbUnit;
			
		}
	}
}
