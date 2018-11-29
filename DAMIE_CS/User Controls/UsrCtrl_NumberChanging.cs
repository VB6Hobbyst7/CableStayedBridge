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
		
		/// <summary>
		/// 用户控件，用来增加或减少指定的日期值。
		/// </summary>
		/// <remarks></remarks>
		public partial class UsrCtrl_NumberChanging
		{
			public UsrCtrl_NumberChanging()
			{
				InitializeComponent();
			}
			
#region   ---  Types
			
			public enum YearMonthDay
			{
				Days = 0,
				Months = 1,
				Years = 2
			}
			
#endregion
			
#region   ---  Properties
			private YearMonthDay _unit;
public YearMonthDay unit
			{
				get
				{
					return _unit;
				}
				set
				{
					cbUnit.SelectedText = System.Convert.ToString(value.ToString());
					_unit = value;
				}
			}
			
			//Private P_CanGetFocus As Boolean
			//Public ReadOnly Property CanGetFocus As Boolean
			//    Get
			//        Return Me.P_CanGetFocus
			//    End Get
			//End Property
#endregion
			
#region   ---  Fields
			public delegate void ValueAddedEventHandler();
			private ValueAddedEventHandler ValueAddedEvent;
			
			public event ValueAddedEventHandler ValueAdded
			{
				add
				{
					ValueAddedEvent = (ValueAddedEventHandler) System.Delegate.Combine(ValueAddedEvent, value);
				}
				remove
				{
					ValueAddedEvent = (ValueAddedEventHandler) System.Delegate.Remove(ValueAddedEvent, value);
				}
			}
			
			public delegate void ValueMinusedEventHandler();
			private ValueMinusedEventHandler ValueMinusedEvent;
			
			public event ValueMinusedEventHandler ValueMinused
			{
				add
				{
					ValueMinusedEvent = (ValueMinusedEventHandler) System.Delegate.Combine(ValueMinusedEvent, value);
				}
				remove
				{
					ValueMinusedEvent = (ValueMinusedEventHandler) System.Delegate.Remove(ValueMinusedEvent, value);
				}
			}
			
			public new delegate void TextChangedEventHandler();
			private TextChangedEventHandler TextChangedEvent;
			
			public new event TextChangedEventHandler TextChanged
			{
				add
				{
					TextChangedEvent = (TextChangedEventHandler) System.Delegate.Combine(TextChangedEvent, value);
				}
				remove
				{
					TextChangedEvent = (TextChangedEventHandler) System.Delegate.Remove(TextChangedEvent, value);
				}
			}
			
#endregion
			
			private float _Value_TimeSpan;
			/// <summary>
			/// 日期文本框上显示的用来进行日期值增减的数量
			/// </summary>
			/// <value></value>
			/// <returns></returns>
			/// <remarks></remarks>
public float Value_TimeSpan
			{
				get
				{
					return _Value_TimeSpan;
				}
			}
			
			public void NumberChanging_Load(object sender, EventArgs e)
			{
				string[] names = Enum.GetNames(typeof(YearMonthDay));
				cbUnit.Items.Clear();
				cbUnit.Items.AddRange(names);
				cbUnit.SelectedIndex = 0;
				TextBoxNumber.Text = System.Convert.ToString(1);
				this._Value_TimeSpan = 1;
			}
			
			public void btnNext_Click(object sender, EventArgs e)
			{
				if (ValueAddedEvent != null)
					ValueAddedEvent();
			}
			public void btnPrevious_Click(object sender, EventArgs e)
			{
				if (ValueMinusedEvent != null)
					ValueMinusedEvent();
			}
			
			public void btnUnit_SelectedIndexChanged(object sender, EventArgs e)
			{
				this._unit = Enum.Parse(typeof(YearMonthDay), System.Convert.ToString(cbUnit.SelectedItem), true);
			}
			
			public void TextBoxNumber_KeyDown(object sender, KeyEventArgs e)
			{
				float v;
				try
				{
					v = float.Parse(TextBoxNumber.Text);
					this._Value_TimeSpan = float.Parse(TextBoxNumber.Text);
					if (TextChangedEvent != null)
						TextChangedEvent();
				}
				catch (Exception)
				{
					TextBoxNumber.Text = "";
					this._Value_TimeSpan = 0;
				}
			}
			//
		}
		
	}
	
}
