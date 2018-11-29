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
	//Namespace UI
	public sealed partial class UI_SplashScreen
	{
		
#region Default Instance
		
		private static UI_SplashScreen defaultInstance;
		
		/// <summary>
		/// Added by the VB.Net to C# Converter to support default instance behavour in C#
		/// </summary>
public static UI_SplashScreen Default
		{
			get
			{
				if (defaultInstance == null)
				{
					defaultInstance = new UI_SplashScreen();
					defaultInstance.FormClosed += new FormClosedEventHandler(defaultInstance_FormClosed);
				}
				
				return defaultInstance;
			}
			set
			{
				defaultInstance = value;
			}
		}
		
		static void defaultInstance_FormClosed(object sender, FormClosedEventArgs e)
		{
			defaultInstance = null;
		}
		
#endregion
		
		//TODO: This form can easily be set as the splash screen for the application by going to the "Application" tab
		//  of the Project Designer ("Properties" under the "Project" menu).
		
		public UI_SplashScreen()
		{
			
			// This call is required by the designer.
			InitializeComponent();
			
			//Added to support default instance behavour in C#
			if (defaultInstance == null)
				defaultInstance = this;
			
			// Add any initialization after the InitializeComponent() call.
			
			//With Panel1
			//    .BackColor = Color.FromArgb(150, 0, 128, 128)
			//End With
		}
		
		public void SplashScreen1_Load(object sender, System.EventArgs e)
		{
			//Set up the dialog text at runtime according to the application's assembly information.
			
			//TODO: Customize the application's assembly information in the "Application" pane of the project
			//  properties dialog (under the "Project" menu).
			
			//Application title
			//If My.Application.Info.Title <> "" Then
			//    ApplicationTitle.Text = My.Application.Info.Title
			//Else
			//    'If the application title is missing, use the application name, without the extension
			//    ApplicationTitle.Text = System.IO.Path.GetFileNameWithoutExtension(My.Application.Info.AssemblyName)
			//End If
			
			// Version.Text = System.String.Format(Version.Text, My.Application.Info.Version.Major, My.Application.Info.Version.Minor, My.Application.Info.Version.Build, My.Application.Info.Version.Revision)
			
			Version.Text = System.String.Format(Version.Text, (new Microsoft.VisualBasic.ApplicationServices.WindowsFormsApplicationBase()).Info.Version.Major, (new Microsoft.VisualBasic.ApplicationServices.WindowsFormsApplicationBase()).Info.Version.Minor);
			
			//Copyright info
			Copyright.Text = (new Microsoft.VisualBasic.ApplicationServices.WindowsFormsApplicationBase()).Info.Copyright;
			//
			
		}
		
		
		
		
	}
	//End Namespace
}
