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
using CableStayedBridge.All_Drawings_In_Application;
using CableStayedBridge.AME_UserControl;
using CableStayedBridge.GlobalApp_Form;
using CableStayedBridge.Miscellaneous;
// End of VB project level imports

namespace CableStayedBridge
{

    /// <summary>
    /// 图形滚动窗口
    /// </summary>
    /// <remarks>最关键的触发图形滚动的事件是</remarks>
    public partial class frmRolling
    {

        #region   ---  Declarations & Definitions

        #region   ---  Types
        /// <summary>
        /// 用来进行滚动显示的绘图对象
        /// </summary>
        /// <remarks></remarks>
        public struct Drawings_For_Rolling
        {
            public ClsDrawing_PlanView PlanView;
            public ClsDrawing_ExcavationElevation SectionalView;
            public List<clsDrawing_Mnt_RollingBase> RollingMnt;

            public short Count()
            {
                short SumUp = System.Convert.ToInt16(RollingMnt.Count);
                if (PlanView != null)
                {
                    SumUp++;
                }
                if (SectionalView != null)
                {
                    SumUp++;
                }
                return SumUp;
            }

            public Drawings_For_Rolling(frmRolling Sender)
            {
                RollingMnt = new List<clsDrawing_Mnt_RollingBase>();
                SectionalView = null;
                PlanView = null;
            }
        }

        #endregion

        /// <summary>
        /// 日期的显示格式：2013/11/2
        /// </summary>
        /// <remarks></remarks>
        const string DateFormat_yyyyMd = "yyyy/M/d";

        #region   ---  Properties
        private DateSpan F_FocusOn_DateSpan;
        /// <summary>
        /// 同步设置滚动条与日历的时间跨度值
        /// </summary>
        /// <value></value>
        /// <remarks>只写属性,代表滚动界面中选择的图表对象的时间跨度的最大并集</remarks>
        public DateSpan FocusOn_DateSpan
        {
            get
            {
                return this.F_FocusOn_DateSpan;
            }
            private set
            {
                //日历的最小日期不能早于1753年1月1日。
                if (DateTime.Compare(value.FinishedDate, new DateTime(1753, 1, 1)) > 0)
                {
                    try
                    {
                        if (DateTime.Compare(value.StartedDate, Calendar_Construction.MaxDate) < 0)
                        {
                            Calendar_Construction.MinDate = value.StartedDate;
                            Calendar_Construction.MaxDate = value.FinishedDate;
                        }
                        else //如果新的最小值比旧的最大值还要大，则先设置最大值，以避免出现先设置最小值时，它比最大值还要大的异常
                        {
                            Calendar_Construction.MaxDate = value.FinishedDate;
                            Calendar_Construction.MinDate = value.StartedDate;
                        }
                    }
                    catch (ArgumentOutOfRangeException ex)
                    {
                        Debug.Print("\r\n" + ex.Message +
                            "\r\n" + "报错位置：" + ex.TargetSite.Name +
                            "日历范围设置出错，但不用进行修正处理。");
                    }
                    finally
                    {
                        this.F_FocusOn_DateSpan = value;
                    }
                }
            }
        }

        /// <summary>
        /// 在滚动时记录的“当天”的日期值。
        /// </summary>
        /// <remarks></remarks>
        DateTime F_Rollingday;
        /// <summary>
        /// !!!触发图形滚动事件。
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public DateTime Rollingday
        {
            get
            {
                return F_Rollingday;
            }
            set // 改变RollingDay属性，以触发图形滚动事件。
            {
                if (DateTime.Compare(value, this.F_Rollingday) != 0)
                {
                    //触发图形滚动事件。
                    Rolling(value, this.F_arrThreadDelegete);
                    //触发DateChanged事件，在事件中会更新日期label的值。
                    this.LabelDate.Text = value.ToString(DateFormat_yyyyMd);
                }
                this.F_Rollingday = value;
            }
        }

        #endregion

        #region   ---  Fields

        /// <summary>
        /// 窗口中是否选择了用来进行滚动的图形
        /// </summary>
        /// <remarks></remarks>
        private bool F_blnHasDrawingToRoll;

        /// <summary>
        /// 列表中选择的用来进行滚动显示的绘图对象
        /// </summary>
        /// <remarks></remarks>
        public Drawings_For_Rolling F_SelectedDrawings;

        /// <summary>
        /// 批量操作绘图曲线的对话框。用来对多条曲线进行锁定或者删除
        /// </summary>
        /// <remarks></remarks>
        private frmLockCurves F_frmHandleMultipleCurve = new frmLockCurves();

        #endregion

        #endregion

        #region   ---  构造函数与窗体的加载、打开与关闭

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <remarks></remarks>
        public frmRolling()
        {

            // This call is required by the designer.
            InitializeComponent();

            // Add any initialization after the InitializeComponent() call.
            //
            this.FocusOn_DateSpan = null;
            this.F_Rollingday = DateTime.Today;
            CheckBox_PlanView.Enabled = false;
            CheckBox_SectionalView.Enabled = false;
            //
            this.F_SelectedDrawings = new Drawings_For_Rolling(this);
        }

        /// <summary>
        /// 窗体加载
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <remarks></remarks>
        public void frmRolling_Load(object sender, EventArgs e)
        {
            //窗口于所有控件之前先接收键盘事件
            //如果此属性设置为True，当键盘按下时，此窗口会接收 KeyPress, KeyDown和KeyUp事件，
            //当窗口完全执行完这些键盘事件的方法后，键盘事件才会接着传递给拥有焦点的那个控件。
            this.KeyPreview = true;
            //在窗口第一次出现时，根据是否有可滚动的图形设置初始界面。
            this.OnRollingDrawingsRefreshed();
        }

        /// <summary>
        /// 窗体关闭之前的事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <remarks>在关闭窗口时将其隐藏</remarks>
        public void frmRolling_FormClosing(object sender, FormClosingEventArgs e)
        {
            //如果是子窗口自己要关闭，则将其隐藏
            //如果是mdi父窗口要关闭，则不隐藏，而由父窗口去结束整个进程
            if (!(e.CloseReason == CloseReason.MdiFormClosing))
            {
                this.Hide();
                e.Cancel = true;
            }
        }

        #endregion

        #region   ---  刷新UI界面——所有滚动窗口

        /// <summary>
        /// 委托：在非UI线程中刷新窗口控件时，用Me.BeginInvoke来转到创建此窗口的线程中来执行。
        /// </summary>
        /// <remarks></remarks>
        private delegate void RollingDrawingsRefreshedHandler();
        /// <summary>
        /// 在此方法中，触发了滚动窗口的RollingDrawingsRefreshed事件：
        /// 刷新滚动窗口中的列表框的数据与其UI显示
        /// </summary>
        /// <remarks></remarks>
        public void OnRollingDrawingsRefreshed()
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new RollingDrawingsRefreshedHandler(this.OnRollingDrawingsRefreshed));
            }
            else
            {
                frmRolling with_1 = this;
                try
                {
                    this.F_SelectedDrawings = RefreshUI_RollingDrawings();
                    SelectedRollingDrawingsChanged(this.F_SelectedDrawings);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("在添加或者删除程序中的图表时，滚动窗口的界面刷新出错。" + "\r\n" + ex.Message + "\r\n" + "报错位置：" + ex.TargetSite.Name,
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        /// <summary>
        /// 在整个程序中的可以滚动的图表发生增加或者减少时触发的事件：
        /// 刷新窗口中的列表框的数据与其UI显示
        /// </summary>
        /// <remarks></remarks>
        private Drawings_For_Rolling RefreshUI_RollingDrawings()
        {
            Drawings_For_Rolling SelectedDrawings = new Drawings_For_Rolling(this);
            // 主程序中所有可以滚动的图形的汇总()
            RollingEnabledDrawings RollingMethods = GlobalApplication.Application.ExposeRollingDrawings();

            //
            ClsDrawing_PlanView plan = RollingMethods.PlanView;
            SelectedDrawings.PlanView = plan;
            CheckBox_PlanView.Tag = plan;
            if (plan != null)
            {
                CheckBox_PlanView.Checked = true;
                CheckBox_PlanView.Enabled = true;
            }
            else
            {
                CheckBox_PlanView.Checked = false;
                CheckBox_PlanView.Enabled = false;
            }
            //
            ClsDrawing_ExcavationElevation Sectional = RollingMethods.SectionalView;
            SelectedDrawings.SectionalView = Sectional;
            CheckBox_SectionalView.Tag = Sectional;
            if (Sectional != null)
            {
                CheckBox_SectionalView.Checked = true;
                CheckBox_SectionalView.Enabled = true;
            }
            else
            {
                CheckBox_SectionalView.Checked = false;
                CheckBox_SectionalView.Enabled = false;
            }

            // --------------  为窗口中的控件赋值  ------------------------
            //
            List<LstbxDisplayAndItem> listMnt = new List<LstbxDisplayAndItem>();

            foreach (clsDrawing_Mnt_RollingBase M in RollingMethods.MonitorData)
            {
                listMnt.Add(new LstbxDisplayAndItem(DisplayedText: M.Chart_App_Title, Value:
                    M));
            }
            this.ListBoxMonitorData.DisplayMember = LstbxDisplayAndItem.DisplayMember;
            this.ListBoxMonitorData.DataSource = listMnt;
            SelectedDrawings.RollingMnt.Clear();
            foreach (LstbxDisplayAndItem item in this.ListBoxMonitorData.SelectedItems)
            {
                SelectedDrawings.RollingMnt.Add((clsDrawing_Mnt_RollingBase)item.Value);
            }

            return SelectedDrawings;
        }

        #endregion

        #region   ---  刷新进行滚动的线程

        //选择复选框或者列表项——更新选择的图表对象
        public void CheckBox_PlanView_CheckedChanged(object sender, EventArgs e)
        {
            if (CheckBox_PlanView.Checked)
            {
                this.F_SelectedDrawings.PlanView = (ClsDrawing_PlanView)CheckBox_PlanView.Tag;
            }
            else
            {
                this.F_SelectedDrawings.PlanView = null;
            }
            //
            SelectedRollingDrawingsChanged(this.F_SelectedDrawings);
        }
        public void CheckBox_SectionalView_CheckedChanged(object sender, EventArgs e)
        {
            if (CheckBox_SectionalView.Checked)
            {
                this.F_SelectedDrawings.SectionalView = (ClsDrawing_ExcavationElevation)CheckBox_SectionalView.Tag;
            }
            else
            {
                this.F_SelectedDrawings.SectionalView = null;
            }
            //
            SelectedRollingDrawingsChanged(this.F_SelectedDrawings);
        }
        public void ListBoxMonitorData_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.F_SelectedDrawings.RollingMnt.Clear();
            //
            var Items = ListBoxMonitorData.SelectedItems;
            clsDrawing_Mnt_RollingBase Drawing = default(clsDrawing_Mnt_RollingBase);
            foreach (LstbxDisplayAndItem lstboxItem in Items)
            {
                Drawing = (clsDrawing_Mnt_RollingBase)lstboxItem.Value;
                this.F_SelectedDrawings.RollingMnt.Add(Drawing);
            }
            //
            SelectedRollingDrawingsChanged(this.F_SelectedDrawings);
        }

        //选择复选框或者列表项——更新滚动线程与窗口界面
        /// <summary>
        /// ！选择的图形发生改变时，更新滚动线程与窗口界面。
        /// </summary>
        /// <param name="Selected_Drawings">更新后的要进行滚动的图形</param>
        /// <remarks>此方法不能直接Handle复选框的CheckedChanged或者列表框的SelectedIndexChanged事件，
        /// 因为此方法必须是在更新了Me.F_SelectedDrawings属性之后，才能去更新窗口界面。</remarks>
        private void SelectedRollingDrawingsChanged(Drawings_For_Rolling Selected_Drawings)
        {
            //
            ConstructRollingThreads(Selected_Drawings);
            this.FocusOn_DateSpan = Refesh_FocusOn_DateSpan(Selected_Drawings);
            //
            if (Selected_Drawings.Count() > 0)
            {
                F_blnHasDrawingToRoll = true;
            }
            else
            {
                F_blnHasDrawingToRoll = false;
            }
            RefreshUI_SelectedDrawings(F_blnHasDrawingToRoll, this.F_FocusOn_DateSpan);
            //
        }

        #region   ---  子方法

        /// <summary>
        /// 为每一个图形滚动设置一个线程
        /// </summary>
        /// <remarks></remarks>
        private Thread[] F_arrThread;
        /// <summary>
        /// 每一个图形滚动的线程所指向的方法
        /// </summary>
        /// <remarks></remarks>
        private ParameterizedThreadStart[] F_arrThreadDelegete;
        /// <summary>
        /// 为每一个选择的要进行滚动的图形构造一个线程
        /// </summary>
        /// <param name="Selected_Drawings"></param>
        /// <remarks></remarks>
        private void ConstructRollingThreads(Drawings_For_Rolling Selected_Drawings)
        {
            AbortAllThread(F_arrThread);

            // -------------------------------------  包含所有要进行滚动的线程的数组
            short SelectedDrawingsCount = Selected_Drawings.Count();
            //
            F_arrThread = new Thread[SelectedDrawingsCount - 1 + 1];
            F_arrThreadDelegete = new ParameterizedThreadStart[SelectedDrawingsCount - 1 + 1];
            //'
            short btThread = (short)0;
            //
            //
            // ------------------------------------------ 标高剖面图
            ClsDrawing_ExcavationElevation ele = Selected_Drawings.SectionalView;
            if (ele != null)
            {
                //为线程数组中的元素线程赋值
                F_arrThreadDelegete[btThread] = new ParameterizedThreadStart(ele.Rolling);
                Thread thd = new Thread(F_arrThreadDelegete[btThread]);
                thd.Name = "滚动标高剖面图";
                F_arrThread[btThread] = thd;
                F_arrThread[btThread].IsBackground = true;
                btThread++;
            }

            // ------------------------------------------ 开挖平面图
            ClsDrawing_PlanView Plan = Selected_Drawings.PlanView;
            if (Plan != null)
            {
                //为线程数组中的元素线程赋值
                F_arrThreadDelegete[btThread] = new ParameterizedThreadStart(Plan.Rolling);
                Thread thd = new Thread(F_arrThreadDelegete[btThread]);
                thd.Name = "滚动开挖平面图";
                F_arrThread[btThread] = thd;
                F_arrThread[btThread].IsBackground = true;
                btThread++;
            }

            // ------------------------------------------ 监测曲线图
            foreach (clsDrawing_Mnt_RollingBase Moni in Selected_Drawings.RollingMnt)
            {
                //为线程数组中的元素线程赋值
                F_arrThreadDelegete[btThread] = new ParameterizedThreadStart(Moni.Rolling);
                Thread thd = new Thread(F_arrThreadDelegete[btThread]);
                thd.Name = "滚动监测曲线图";
                F_arrThread[btThread] = thd;
                F_arrThread[btThread].IsBackground = true;
                btThread++;
            }
        }

        /// <summary>
        /// !!! 在每一次更改选择项时刷新滚动条和日历的时间跨度，以及终结旧线程，创建新的线程
        /// </summary>
        /// <remarks></remarks>
        private DateSpan Refesh_FocusOn_DateSpan(Drawings_For_Rolling Selected_Drawings)
        {
            // -------------------------------------  包含所有要进行滚动的线程的数组
            //
            bool blnDateSpanInitialized = false;
            DateSpan DateSpan = new DateSpan();
            //
            // ------------------------------------------ 标高剖面图
            ClsDrawing_ExcavationElevation ele = Selected_Drawings.SectionalView;
            if (ele != null)
            {
                //更新滚动界面的时间跨度
                DateSpan = RenewTimeSpan(ref blnDateSpanInitialized, DateSpan, ele.DateSpan);
            }

            // ------------------------------------------ 开挖平面图
            ClsDrawing_PlanView Plan = Selected_Drawings.PlanView;
            if (Plan != null)
            {
                //更新滚动界面的时间跨度
                DateSpan = RenewTimeSpan(ref blnDateSpanInitialized, DateSpan, Plan.DateSpan);
            }

            // ------------------------------------------ 监测曲线图
            foreach (clsDrawing_Mnt_RollingBase Moni in Selected_Drawings.RollingMnt)
            {
                //更新滚动界面的时间跨度
                DateSpan = RenewTimeSpan(ref blnDateSpanInitialized, DateSpan, Moni.DateSpan);
            }
            return DateSpan;
        }
        /// <summary>
        /// 更新滚动窗口中的日期跨度
        /// </summary>
        /// <param name="blnDateSpanInitialized">引用类型，指示滚动窗口中是否已经有了最初的作为基准日期跨度的值</param>
        /// <param name="OldDateSpan">已经有的日期跨度的值</param>
        /// <param name="ObjectDateSpan">进行扩充的日期跨度</param>
        /// <returns></returns>
        /// <remarks></remarks>
        private DateSpan RenewTimeSpan(ref bool blnDateSpanInitialized, DateSpan
            OldDateSpan, DateSpan ObjectDateSpan)
        {
            DateSpan NewDateSpan = new DateSpan();
            if (!blnDateSpanInitialized)
            {
                NewDateSpan = ObjectDateSpan;
                blnDateSpanInitialized = true;
            }
            else
            {
                NewDateSpan = GeneralMethods.ExpandDateSpan(OldDateSpan, ObjectDateSpan);
            }
            return NewDateSpan;
        }

        #endregion

        #endregion

        #region   ---  触发图形滚动的操作

        /// <summary>
        /// 1.通过键盘的方向键来进行图形的滚动。按下键盘时列表框不要处于激活状态。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <remarks>要注意的是，在按下键盘时，如果窗口中的列表框是处于激活状态，
        /// 则程序会同时执行ListBox.SelectedIndexChanged与此处的KeyDown，那么，
        /// 其结果就是先启动了一个图形的滚动线程，然后第二个事件又马上将此线程关闭了。</remarks>
        public void frmRolling_KeyDown(object sender, KeyEventArgs e)
        {
            if (F_blnHasDrawingToRoll)
            {
                if (e.KeyCode == Keys.Left)
                {
                    this.NumChanging_ValueMinused();
                }
                else if (e.KeyCode == Keys.Right)
                {
                    this.NumChanging_ValueAdded();
                }
            }
        }

        /// <summary>
        /// 2.日期的增加
        /// </summary>
        /// <remarks></remarks>
        public void NumChanging_ValueAdded()
        {
            DateTime NewDay = default(DateTime);
            switch (NumChanging.unit)
            {
                case UsrCtrl_NumberChanging.YearMonthDay.Days:
                    NewDay = this.F_Rollingday.AddDays(NumChanging.Value_TimeSpan);
                    break;
                case UsrCtrl_NumberChanging.YearMonthDay.Months:
                    NewDay = this.F_Rollingday.AddMonths((int)NumChanging.Value_TimeSpan);
                    break;
                case UsrCtrl_NumberChanging.YearMonthDay.Years:
                    NewDay = this.F_Rollingday.AddYears((int)NumChanging.Value_TimeSpan);
                    break;
            }

            //限制日期的跨度在最大跨度范围之内
            if (DateTime.Compare(NewDay, this.FocusOn_DateSpan.FinishedDate) > 0)
            {
                //这里没有判断与最早日期的比较，是因为这里的操作是将日期增加
                NewDay = this.FocusOn_DateSpan.FinishedDate;
            }
            this.Rollingday = NewDay;

        }
        /// <summary>
        /// 2.日期的后退
        /// </summary>
        /// <remarks></remarks>
        public void NumChanging_ValueMinused()
        {
            DateTime Newday = default(DateTime);
            switch (NumChanging.unit)
            {
                case UsrCtrl_NumberChanging.YearMonthDay.Days:
                    Newday = this.F_Rollingday.AddDays(System.Convert.ToDouble(-NumChanging.Value_TimeSpan));
                    break;
                case UsrCtrl_NumberChanging.YearMonthDay.Months:
                    Newday = this.F_Rollingday.AddMonths(System.Convert.ToInt32(-NumChanging.Value_TimeSpan));
                    break;
                case UsrCtrl_NumberChanging.YearMonthDay.Years:
                    Newday = this.F_Rollingday.AddYears(System.Convert.ToInt32(-NumChanging.Value_TimeSpan));
                    break;
            }

            //限制日期的跨度在最大跨度范围之内
            if (DateTime.Compare(Newday, this.FocusOn_DateSpan.StartedDate) < 0)
            {
                //这里没有判断与最晚日期的比较，是因为这里的操作是将日期减小
                Newday = this.FocusOn_DateSpan.StartedDate;
            }
            //改变RollingDay属性，以触发图形滚动事件。
            this.Rollingday = Newday;
        }

        /// <summary>
        /// 3.从日历中选择指定日期
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <remarks></remarks>
        public void Calendar_Construction_DateSelected(object sender, DateRangeEventArgs e)
        {
            DateTime NewDate = this.Calendar_Construction.SelectionEnd.Date;
            //改变RollingDay属性，以触发图形滚动事件。
            this.Rollingday = NewDate;
        }

        /// <summary>
        /// 4.在窗口中选择的进行滚动的窗口发生改变时刷新UI界面：日历、日期前进或后退
        /// </summary>
        /// <param name="blnHasDrawingToRoll">指示是否有要进行滚动的图形</param>
        /// <param name="DateSpan">当前选择的滚动图形的日期跨度的并集。
        /// 如果窗口中没有要进行滚动的图形，则此参数值无效</param>
        /// <remarks></remarks>
        private void RefreshUI_SelectedDrawings(bool blnHasDrawingToRoll, DateSpan DateSpan)
        {
            if (!blnHasDrawingToRoll) //如果窗口中没有选择任何一项
            {
                //窗口的界面刷新
                this.Panel_Roll.Enabled = false;
                this.Calendar_Construction.Visible = false;
                //设置当天的日期
                this.LabelDate.Text = this.F_Rollingday.ToString(DateFormat_yyyyMd);
            }
            else
            {
                //窗口的界面刷新
                this.Panel_Roll.Enabled = true;
                this.Calendar_Construction.Visible = true;
                //设置当天的日期
                DateTime NewDay = this.F_Rollingday;
                //如果超出日期跨度的范围，则设置其为跨度的边界
                if (DateTime.Compare(this.F_Rollingday, DateSpan.StartedDate) < 0)
                {
                    NewDay = DateSpan.StartedDate;
                }
                else if (DateTime.Compare(this.F_Rollingday, DateSpan.FinishedDate) > 0)
                {
                    NewDay = DateSpan.FinishedDate;
                }
                //如果文本框前后的字符不变，那么此时是否会触发其TextChanged事件呢——
                //会，因为程序只会只会判断是否进行了重新赋值，而不会判断前后的值是否相同。
                this.LabelDate.Text = NewDay.ToString(DateFormat_yyyyMd);
            }
        }

        #endregion

        #region   ---  图形滚动

        /// <summary>
        /// 【关键】滚动时的执行方法
        /// </summary>
        /// <param name="ThisDay">进行滚动的当天的日期</param>
        /// <param name="arrThreadDelegete">每一个图形滚动的线程所指向的方法</param>
        /// <remarks>在进行滚动时，先判断是否还有线程没有执行完，如果有，则要先将未结束的线程取消掉，然后再重新启动线程。
        /// 而线程的数量与调用的方法是不变的。</remarks>
        private void Rolling(DateTime ThisDay, System.Threading.ParameterizedThreadStart[] arrThreadDelegete)
        {

            //取消未结束的线程
            AbortAllThread(this.F_arrThread);
            try
            {
                //创建新线程并重新启动
                if (arrThreadDelegete != null)
                {
                    for (int i = 0; i <= arrThreadDelegete.Count() - 1; i++)
                    {
                        this.F_arrThread[i] = new Thread(arrThreadDelegete[i]);
                        this.F_arrThread[i].Start(ThisDay);
                    }
                }
            }
            catch (Exception)
            {
                Debug.Print("图形滚动出错！——StartRolling");
            }
        }

        /// <summary>
        /// 取消所有当前正在运行的线程
        /// </summary>
        /// <param name="Threads"></param>
        /// <remarks></remarks>
        private void AbortAllThread(Thread[] Threads)
        {
            if (Threads != null)
            {
                foreach (Thread thd in Threads)
                {
                    if (thd.IsAlive)
                    {
                        //在Abort方法中会自动抛出异常：ThreadAbortException
                        thd.Abort();
                        Debug.Print("线程成功终止。");
                    }
                }
            }
        }

        #endregion

        #region   --- 一般界面操作

        public void ProgressBar_PlanView_Click(object sender, EventArgs e)
        {
            if (CheckBox_PlanView.Enabled)
            {
                if (CheckBox_PlanView.Checked)
                {
                    CheckBox_PlanView.Checked = false;
                }
                else
                {
                    CheckBox_PlanView.Checked = true;
                }
            }
        }
        public void ProgressBar_SectionalView_Click(object sender, EventArgs e)
        {
            if (CheckBox_SectionalView.Enabled)
            {
                if (CheckBox_SectionalView.Checked)
                {
                    CheckBox_SectionalView.Checked = false;
                }
                else
                {
                    CheckBox_SectionalView.Checked = true;
                }
            }
        }

        /// <summary>
        /// 输出到Word
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <remarks></remarks>
        public void btnOutPut_Click(object sender, EventArgs e)
        {
            APPLICATION_MAINFORM.MainForm.ExportToWord(null, null);
        }

        /// <summary>
        /// 进行批量操作
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <remarks></remarks>
        public void btn_GroupHandle_Click(object sender, EventArgs e)
        {
            this.F_frmHandleMultipleCurve.ShowDialog();
        }
        #endregion

    }

}
