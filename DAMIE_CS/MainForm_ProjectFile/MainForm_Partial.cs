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

using System.ComponentModel;
using CableStayedBridge.Miscellaneous;

namespace CableStayedBridge
{
    namespace GlobalApp_Form
    {

        public partial class APPLICATION_MAINFORM
        {
            #region   --- 进度条的显示与隐藏

            /// <summary>
            /// 委托：在主程序界面上显示进度条与进度信息
            /// </summary>
            /// <remarks></remarks>
            private delegate void ShowProgressBar_MarqueeHandler();
            /// <summary>
            /// 在主程序界面上显示进度条与进度信息
            /// </summary>
            /// <remarks></remarks>
            public void ShowProgressBar_Marquee()
            {
                APPLICATION_MAINFORM with_1 = this;
                if (with_1.InvokeRequired)
                {
                    //非UI线程，再次封送该方法到UI线程
                    this.BeginInvoke(new ShowProgressBar_MarqueeHandler(this.ShowProgressBar_Marquee));

                }
                else
                {
                    //UI线程，进度更新
                    //主程序界面的进度条的UI显示

                    with_1.StatusLabel1.Text = "Please Wait...";
                    with_1.StatusLabel1.Visible = true;

                    with_1.ProgressBar1.Visible = true;
                    with_1.ProgressBar1.Style = ProgressBarStyle.Marquee;
                    //获取或设置进度块在进度栏内滚动所用的时间段，以毫秒为单位，其值越小，图形滚动得越快。
                    //只对于Style属性为Marquee的ProgressBar控件有效。
                    with_1.ProgressBar1.MarqueeAnimationSpeed = 30;
                }
            }

            private delegate void ShowProgressBar_ContinueHander(int Value);
            public void ShowProgressBar_Continue(int Value)
            {
                APPLICATION_MAINFORM with_1 = this;
                if (with_1.InvokeRequired)
                {
                    //非UI线程，再次封送该方法到UI线程
                    this.BeginInvoke(new ShowProgressBar_ContinueHander[this.ShowProgressBar_Continue], new[] { Value });
                }
                else
                {
                    //UI线程，进度更新
                    //主程序界面的进度条的UI显示

                    with_1.StatusLabel1.Text = "Please Wait...";
                    with_1.StatusLabel1.Visible = true;

                    with_1.ProgressBar1.Visible = true;
                    with_1.ProgressBar1.Style = ProgressBarStyle.Continuous;
                    //获取或设置进度块在进度栏内滚动所用的时间段，以毫秒为单位，其值越小，图形滚动得越快。
                    //只对于Style属性为Marquee的ProgressBar控件有效。
                    with_1.ProgressBar1.Value = Value;
                }
            }

            /// <summary>
            /// 委托：隐藏主程序界面中的进度条与进度信息
            /// </summary>
            /// <remarks></remarks>
            private delegate void HideProgressHandler(string state);
            /// <summary>
            /// 隐藏主程序界面中的进度条与进度信息
            /// </summary>
            /// <param name="state">要显示在进度信息标签中的文本</param>
            /// <remarks></remarks>
            public void HideProgress(string state)
            {
                APPLICATION_MAINFORM with_1 = this;
                if (with_1.InvokeRequired)
                {
                    //非UI线程，再次封送该方法到UI线程
                    this.BeginInvoke(new HideProgressHandler(this.HideProgress), state);
                }
                else
                {
                    //UI线程，进度更新
                    //主程序界面的进度条的UI显示
                    this.ProgressBar1.Value = 100;
                    this.ProgressBar1.Visible = false;
                    this.StatusLabel1.Text = state;
                    myTimer.Interval = 500;
                    myTimer.Start();
                }
            }

            // This is the method to run when the timer is raised.
            /// <summary>
            /// 闹钟触发的次数
            /// </summary>
            /// <remarks></remarks>
            private int alarmCounter = 1;
            private System.Windows.Forms.Timer myTimer = new System.Windows.Forms.Timer();
            /// <summary>
            /// 在指定的时间间隔内触发闹钟事件
            /// </summary>
            /// <param name="myObject"></param>
            /// <param name="myEventArgs"></param>
            /// <remarks></remarks>
            public void TimerEventProcessor(object myObject, EventArgs myEventArgs)
            {
                if (alarmCounter == 2)
                {
                    APPLICATION_MAINFORM with_1 = this;
                    with_1.ProgressBar1.Visible = false;
                    with_1.StatusLabel1.Visible = false;
                    myTimer.Stop();
                    alarmCounter = 1;
                }
                else
                {
                    alarmCounter++;
                }
            }

            #endregion

            #region   --- 控制主程序界面的显示

            private delegate void MainUI_ProjectNotOpenedHandler();
            /// <summary>
            /// 设置主程序中还没有任何项目文件打开时的UI界面
            /// </summary>
            /// <remarks></remarks>
            private void MainUI_ProjectNotOpened()
            {
                if (this.InvokeRequired)
                {
                    this.BeginInvoke(new MainUI_ProjectNotOpenedHandler(this.MainUI_ProjectNotOpened));
                }
                else
                {
                    APPLICATION_MAINFORM with_1 = this;
                    //Add What you want to do here.
                    //禁用菜单项 ------ 文件
                    with_1.MenuItem_EditProject.Enabled = false;
                    with_1.MenuItem_SaveProject.Enabled = false;
                    with_1.MenuItem_SaveAsProject.Enabled = false;
                    with_1.MenuItemExport.Enabled = false;
                    //禁用菜单项 ------ 编辑
                    with_1.MenuItemDrawingPoints.Enabled = false;
                    //禁用菜单项 ------ 绘图
                    with_1.MenuItemSectionalView.Enabled = false;
                    with_1.MenuItemPlanView.Enabled = false;
                    //禁用工具栏 ------ 同步滚动
                    with_1.TlStrpBtn_Roll.Enabled = false;
                }
            }

            private delegate void MainUI_ProjectOpenedHandler();
            /// <summary>
            /// 设置项目文档打开后的主程序的UI界面
            /// </summary>
            /// <remarks></remarks>
            internal void MainUI_ProjectOpened()
            {
                if (this.InvokeRequired)
                {
                    this.BeginInvoke(new MainUI_ProjectOpenedHandler(this.MainUI_ProjectOpened));
                }
                else
                {
                    APPLICATION_MAINFORM with_1 = this;
                    //Add What you want to do here.
                    //启用菜单项 ------ 文件
                    with_1.MenuItem_EditProject.Enabled = true;
                    with_1.MenuItem_SaveProject.Enabled = true;
                    with_1.MenuItem_SaveAsProject.Enabled = true;
                    with_1.MenuItemExport.Enabled = true;
                    //启用菜单项 ------ 编辑
                    with_1.MenuItemDrawingPoints.Enabled = true;
                    //启用菜单项 ------ 绘图
                    with_1.MenuItemSectionalView.Enabled = true;
                    with_1.MenuItemPlanView.Enabled = true;
                    //启用工具栏 ------ 同步滚动
                    with_1.TlStrpBtn_Roll.Enabled = true;
                }
            }

            private delegate void MainUI_RollingObjectCreatedHandler();
            internal void MainUI_RollingObjectCreated()
            {
                if (this.InvokeRequired)
                {
                    this.BeginInvoke(new MainUI_RollingObjectCreatedHandler(this.MainUI_RollingObjectCreated));
                }
                else
                {
                    APPLICATION_MAINFORM with_1 = this;
                    //Add What you want to do here.
                    //启用工具栏 ------ 同步滚动
                    with_1.TlStrpBtn_Roll.Enabled = true;
                }
            }
            #endregion

        }
    }
}
