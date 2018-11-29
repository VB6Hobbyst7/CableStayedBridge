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
using CableStayedBridge.Constants;
using CableStayedBridge.GlobalApp_Form;
using CableStayedBridge.Miscellaneous;
// End of VB project level imports

using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace CableStayedBridge
{
    namespace All_Drawings_In_Application
    {

        /// <summary>
        /// 用Excel绘制的，可以进行数据曲线的滚动的类型
        /// </summary>
        /// <remarks></remarks>
        public abstract class clsDrawing_Mnt_RollingBase : ClsDrawing_Mnt_Base, IRolling
        {

            #region   ---  Types

            /// <summary>
            /// 数据列series对象所对应的一些信息对象
            /// </summary>
            /// <remarks>它包括每一条数据系列的对象，及其对应的日期、开挖深度的标志线、记录开挖深度值的文本框</remarks>
            public class SeriesTag
            {

                private Excel.Series P_series;
                /// <summary>
                /// 数据系列的series对象
                /// </summary>
                /// 数据列对象
                public Excel.Series series
                {
                    get
                    {
                        return this.P_series;
                    }
                }

                private DateTime P_ConstructionDate;
                /// <summary>
                /// 进行滚动的施工日期
                /// </summary>
                /// 对应的施工日期
                public DateTime ConstructionDate
                {
                    get
                    {
                        return this.P_ConstructionDate;
                    }
                    set
                    {
                        this.P_ConstructionDate = value;
                    }
                }

                public SeriesTag(Excel.Series Series, DateTime ConstructionDate)
                {
                    this.P_series = Series;
                    this.P_ConstructionDate = ConstructionDate;
                }

            }

            /// <summary>
            /// 图例框的尺寸与位置，用来在进行监测数据曲线的滚动时设置图例的大小与位置
            /// </summary>
            /// <remarks></remarks>
            public struct LegendLocation
            {
                /// <summary>
                /// 图例框的高度
                /// </summary>
                /// <remarks></remarks>
                public object Legend_Height;
                /// <summary>
                /// 图例框的宽度
                /// </summary>
                /// <remarks></remarks>
                public object Legend_Width;
                /// <summary>
                /// 构造函数
                /// </summary>
                /// <param name="LegendHeight">图例框的高度</param>
                /// <param name="LegendWidth">图例框的宽度</param>
                /// <remarks></remarks>
                public LegendLocation(float LegendHeight, float LegendWidth)
                {
                    this.Legend_Height = LegendHeight;
                    this.Legend_Width = LegendWidth;
                }
            }

            #endregion

            #region   ---  Constants

            /// <summary>
            /// 数据列在集合seriesCollection中的第一个元素的下标值
            /// </summary>
            /// <remarks>office2010的dll中，其值为1（2014.11.13注）</remarks>
            private const byte cst_LboundOfSeriesInCollection = (byte)LowIndexOfObjectsInExcel.SeriesInSeriesCollection;

            #endregion

            #region   ---  Properties

            /// <summary>
            /// 此结构有四个元素，分别代表：画布的高度、宽度；由画布边界扩展到Excel界面的尺寸的高度和宽度的增量
            /// </summary>
            /// <value></value>
            /// <returns></returns>
            /// <remarks></remarks>
            protected abstract override ChartSize ChartSize_sugested { get; set; }

            /// <summary>
            /// 图例框的尺寸与位置
            /// </summary>
            /// <value></value>
            /// <returns></returns>
            /// <remarks></remarks>
            protected virtual LegendLocation Legend_Location
            {
                get
                {
                    return new LegendLocation(Data_Drawing_Format.Drawing_Mnt_RollingBase.Legend_Height,
                        Data_Drawing_Format.Drawing_Mnt_RollingBase.Legend_Width);
                }
            }

            private DateSpan P_DateSpan;
            /// <summary>
            /// 进行同步滚动的时间跨度，用来给出时间滚动条与日历的范围。
            /// </summary>
            /// <value></value>
            /// <returns></returns>
            /// <remarks></remarks>
            public DateSpan DateSpan
            {
                get
                {
                    return this.P_DateSpan;
                }
                private set
                {
                    // 扩展MainForm.TimeSpan的区间
                    GlobalApplication.Application.refreshGlobalDateSpan(value);
                    this.P_DateSpan = value;
                }
            }

            /// <summary>
            /// 用来进行滚动的那一个数据列对象，此对象在复制的过程中，并不会被替换，
            /// 而且其seriesIndex的值始终为1.
            /// </summary>
            /// <remarks></remarks>
            private Excel.Series P_MovingSeries;
            /// <summary>
            /// 用来进行滚动的那一个数据列对象，此对象在复制的过程中，并不会被替换
            /// 而且其seriesIndex的值始终为1.
            /// </summary>
            /// <remarks></remarks>
            protected Excel.Series MovingSeries
            {
                get
                {
                    return this.P_MovingSeries;
                }
                set
                {
                    this.P_MovingSeries = value;
                }
            }


            /// <summary>
            /// 用数据列的ID属性来索引对应的数据列，以及与之相关的Tag数据
            /// </summary>
            protected Dictionary<int, SeriesTag> F_DicSeries_Tag = new Dictionary<int, SeriesTag>();
            /// <summary>
            /// 用数据列的ID属性来索引对应的数据列，以及与之相关的Tag数据
            /// </summary>
            /// <value>其Key值对应是的每一条曲线的SeriesIndex的值，其中第一条曲线的下标值为1，而不是0。</value>
            /// <returns></returns>
            /// <remarks>作为关键字的ID值并不是以0为初始值的，
            /// 而是在series在seriescollection中的序号值来定义的。</remarks>
            public Dictionary<int, SeriesTag> Dic_SeriesIndex_Tag
            {
                get
                {
                    return this.F_DicSeries_Tag;
                }
            }
            #endregion

            #region   ---  Fields

            /// <summary>
            /// 记录图表中所有的数据系列中的每一项是否有曲线图
            /// </summary>
            /// <remarks>由此来确定数据曲线是要放在哪一个系列中，
            /// 以及图例项中要删除哪些或者显示哪些项。
            /// 此列表中的元素的个数始终等于图表中定义的数据系列的个数，即seriesCollection.Count的值</remarks>
            protected List<bool> F_List_HasCurve = new List<bool>();

            /// <summary>
            /// 图表中所包含的数据曲线的数量
            /// </summary>
            /// <remarks>由于制作模板的问题，所以SeriesCount指的是图表中的显示出来的监测曲线的数量，
            /// 而seriesCollection.Count指的是图表中定义的数据系列的数量。
            /// 空白系列会包含在seriesCollection.Count中，而在图表中并没有曲线图，因为没有数据。</remarks>
            private int F_CurvesCount;

            /// <summary>
            /// 当前滚动的日期。由于进行滚动的总是seriescollection中的第一条曲线，
            /// 所以当前滚动的日期也总是这第一条曲线所代表的日期，亦即Rolling方法中的dateThisday参数的值。
            /// </summary>
            /// <remarks></remarks>
            protected DateTime F_RollingDate;
            #endregion

            /// <summary>
            /// 构造函数
            /// </summary>
            /// <param name="DataSheet">图表对应的数据工作表</param>
            /// <param name="DrawingChart">Excel图形所在的Chart对象</param>
            /// <param name="ParentApp">此图表所在的Excel类的实例对象</param>
            /// <param name="type">此图表所属的类型，由枚举drawingtype提供</param>
            /// <param name="CanRoll">是图表是否可以滚动，即是动态图还是静态图</param>
            /// <param name="DateSpan">此图表的TimeSpan跨度</param>
            /// <remarks></remarks>
            public clsDrawing_Mnt_RollingBase(Excel.Worksheet DataSheet, Excel.Chart
                DrawingChart, Cls_ExcelForMonitorDrawing ParentApp,
                DrawingType type, bool CanRoll, Excel.TextFrame2 Info,
                MonitorInfo DrawingTag, MntType MonitorType,
                DateSpan DateSpan, SeriesTag
                theFirstSeriesTag) : base(DataSheet, DrawingChart, ParentApp, type, CanRoll, Info, DrawingTag, MonitorType)
            {
                // VBConversions Note: Non-static class variable initialization is below.  Class variables cannot be initially assigned non-static values in C#.
                F_Chart = this.Chart;

                //
                this.DateSpan = DateSpan;
                //刷新滚动窗口的列表框的界面显示
                APPLICATION_MAINFORM.MainForm.Form_Rolling.OnRollingDrawingsRefreshed();

                //启用主界面的程序滚动按钮
                APPLICATION_MAINFORM.MainForm.MainUI_RollingObjectCreated();

                //--------------------------- 设置与数据系列的曲线相关的属性值

                clsDrawing_Mnt_RollingBase with_1 = this;
                //以数据列中第一个元素作为进行滚动的那个series
                with_1.MovingSeries = theFirstSeriesTag.series;
                // ----- 集合数据的记录
                with_1.F_DicSeries_Tag.Add(cst_LboundOfSeriesInCollection, theFirstSeriesTag);
                //刚开始时，图表中只有一条数据曲线
                with_1.F_CurvesCount = 1;
                //
                this.F_List_HasCurve.Clear();
                this.F_List_HasCurve.Add(true); //第一个数据列是有曲线的，所以将其值设置为true
                Excel.SeriesCollection seriesColl = Chart.SeriesCollection() as Excel.SeriesCollection;
                for (var i = 1; i <= seriesColl.Count - 1; i++)
                {
                    this.F_List_HasCurve.Add(false);
                }
                // -----对图例进行更新---------
                LegendRefresh(F_List_HasCurve);

            }

            /// <summary>
            /// 日期滚动时的Rolling方法
            /// </summary>
            /// <param name="dateThisday">当天的日期</param>
            /// <remarks>在基类中，并不给出此方法的实现，而由其子类去进行重写</remarks>
            void IRolling.Rolling(DateTime dateThisday)
            {
                this.Rolling(dateThisday);
            }

            public abstract void Rolling(DateTime dateThisday);

            /// <summary>
            /// 根据所有的日期集合和当天的日期，得到距当天最近的那天的日期。
            /// </summary>
            /// <param name="sortedlist_AllDays">有监测数据的所有的日期集合。</param>
            /// <param name="dateThisDay">想要进行绘图的那一天的日期。</param>
            /// <returns>在集合中距离当天最近的那一天。</returns>
            /// <remarks></remarks>
            protected DateTime GetClosestDay(SortedSet<DateTime> sortedlist_AllDays, DateTime dateThisDay)
            {
                DateTime day_early = default(DateTime);
                DateTime day_late = default(DateTime);
                DateTime closedDay = default(DateTime);
                int intCompare = 0;
                if (DateTime.Compare(dateThisDay, sortedlist_AllDays.First()) < 0) //如果选定的日期比最早的监测日期还要早
                {
                    closedDay = sortedlist_AllDays.First();

                }
                else if (DateTime.Compare(dateThisDay, sortedlist_AllDays[sortedlist_AllDays.Count - 1]) > 0) //如果选定的日期比最晚的监测日期还要晚
                {
                    closedDay = sortedlist_AllDays[sortedlist_AllDays.Count - 1];

                }
                else //如果选定的日期位于监测的日期跨度之间
                {
                    foreach (DateTime iDay in sortedlist_AllDays)
                    {
                        intCompare = DateTime.Compare(iDay, dateThisDay);
                        //从最早的一天开始比较，那么刚开始的intCompare的值一定小于0，直到有一天的日期值比选定日期晚，此时intCompare的值大于0
                        //不会出现intCompare等于0的情况，因为比较是在字典的键集合中没有选定日期的情况下进行的
                        if (intCompare < 0) //这天比选定的日期早，记录这一天，以作后面day_late的前一天的值。
                        {
                            day_early = iDay;
                        }
                        else
                        {
                            day_late = iDay; //发现了满足条件的判断点：出现了比选定日期晚的那一天
                                             //接下来从选定日期的左右日期中选择距离最近的一天，如果两边距离相等，则取较早的一天。
                            if (day_late.Subtract(dateThisDay) >= -day_early.Subtract(dateThisDay))
                            {
                                closedDay = day_early; //从选定日期的左右日期中选择距离最近的一天，如果两边距离相等，则取较早的一天。
                            }
                            else
                            {
                                closedDay = day_late;
                            }
                            break; //不用再继续找下一天了
                        }
                    }
                }
                return closedDay;
            }

            #region   ---  数据列的锁定与删除

            private Excel.Chart F_Chart; // VBConversions Note: Initial value cannot be assigned here since it is non-static.  Assignment has been moved to the class constructors.
                                         /// <summary>
                                         /// 当在数据列上双击时执行相应的操作
                                         /// </summary>
                                         /// <param name="ElementID">在图表上双击击中的对象</param>
                                         /// <param name="Arg1">所选数据系列在集合中的索引下标值，注意，第一第曲线的下标值为1，而不是0。</param>
                                         /// <param name="Arg2"></param>
                                         /// <param name="Cancel"></param>
                                         /// <remarks>要么将其删除，要么将其锁定在图表中，要么什么都不做</remarks>
            private void SeriesChange(int ElementID, int Arg1, int Arg2, ref
                bool Cancel)
            {
                if (ElementID == (int)Excel.XlChartItem.xlSeries)
                {
                    Debug.Print("当前选择的曲线下标值为： " + System.Convert.ToString(Arg1));
                    //所选数据系列在集合中的索引下标值，注意，第一第曲线的下标值为1，而不是0
                    int seriesIndex = Arg1;
                    //所选的数据系列
                    Excel.Series seri;
                    Excel.SeriesCollection seriColl = this.Chart.SeriesCollection() as Excel.SeriesCollection;
                    seri = seriColl.Item(seriesIndex);


                    // -------- 打开处理对话框，并根据返回的不同结果来执行不同的操作
                    DiaFrm_LockDelete diafrm = new DiaFrm_LockDelete();
                    //判断要删除的曲线是否为当前滚动的曲线，有如下两种判断方法
                    bool blnDeletingTheRollingCurve = false;
                    blnDeletingTheRollingCurve = seriesIndex == cst_LboundOfSeriesInCollection;
                    //blnDeletingTheRollingCurve = (seri.Name = F_movingSeries.Name)
                    if (F_CurvesCount == 1 || blnDeletingTheRollingCurve)
                    {
                        //如果图表中只有一条曲线，或者点击的正好是要正在进行滚动的曲线，那么这条曲线不能被删除
                        diafrm.btn2.Enabled = false;
                        diafrm.AcceptButton = diafrm.btn1;
                    }
                    else //只能进行删除
                    {
                        diafrm.btn1.Enabled = false;
                        diafrm.AcceptButton = diafrm.btn2;
                    }

                    //在界面上取消图表对象的选择
                    this.Sheet_Drawing.Range["A1"].Activate();
                    this.Application.ScreenUpdating = false;

                    //---------------------------------- 开始执行数据系列的添加或删除
                    //此数据列对应的施工日期
                    DateTime t_date = default(DateTime);
                    SeriesTag SrTag = default(SeriesTag);
                    try
                    {
                        SrTag = this.F_DicSeries_Tag[seriesIndex];
                        t_date = SrTag.ConstructionDate;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("提取字典中的元素出错，字典中没有此元素。" + "\r\n" + ex.Message + "\r\n" + "报错位置：" + ex.TargetSite.Name,
                            "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    //----------------------------------------
                    M_DialogResult result = default(M_DialogResult);
                    string strPrompt = "The date value for the selected series is:" + "\r\n" +
                        t_date.ToString(AMEApplication.DateFormat) + "\r\n" +
                        "Choose to Lock or Delete this series...";
                    result = diafrm.ShowDialog(strPrompt, "Lock or Delete Series");
                    //----------------------------------------
                    switch (result)
                    {
                        case M_DialogResult.Delete:
                            //删除数据系列
                            try
                            {
                                DeleteSeries(seriesIndex);
                            }
                            catch (Exception ex)
                            {
                                Debug.Print("删除数据系列出错。" + "\r\n" + ex.Message + "\r\n" + "报错位置：" + ex.TargetSite.Name);
                            }
                            break;
                        case M_DialogResult.Lock:
                            //添加数据系列
                            try
                            {
                                CopySeries(seriesIndex);
                            }
                            catch (Exception ex)
                            {
                                Debug.Print(ex.Message);
                                MessageBox.Show("添加数据系列出错。" + "\r\n" + ex.Message + "\r\n" + "报错位置：" + ex.TargetSite.Name,
                                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                            break;
                    }
                    this.Application.ScreenUpdating = true;
                }
                //覆盖原来的双击操作
                Cancel = true;
                this.Sheet_Drawing.Range["A1"].Activate();
            }

            //曲线删除
            /// <summary>
            /// 移除指定的数据列及其依附的对象
            /// </summary>
            /// <param name="DeletingSeriesIndex">要进行删除的数据列在集合中的索引下标值，注意：第一条曲线的下标值为1，而不是0。</param>
            /// <remarks>为了保存模板中的数据系列的格式，这里的删除并不是将数据列进行了真正的删除，而是将数据列的数据设置为空。
            /// 这样的话，后期的数据曲线应该加载在最靠前而且没有数据的数据列中。</remarks>
            public virtual void DeleteSeries(int DeletingSeriesIndex)
            {
                SeriesTag with_1 = F_DicSeries_Tag[DeletingSeriesIndex];
                Excel.Series with_2 = with_1.series;
                with_2.XValues = new object[] {null}; // in VB: with_2.XValues = {Nothing}
                //这里不能用.XValues = Nothing
                with_2.Values = new object[] { null };
                //.Name = ""

                this.F_DicSeries_Tag.Remove(DeletingSeriesIndex);
                this.F_List_HasCurve[DeletingSeriesIndex - cst_LboundOfSeriesInCollection] = false;

                // ----------------------------------- 对图例的显示进行操作
                LegendRefresh(F_List_HasCurve);

                this.F_CurvesCount--;
            }

            //曲线添加
            /// <summary>
            /// 由指定的数据列复制出一个新的数据列来，并在图表中显示
            /// </summary>
            /// <param name="SourceSeriesIndex">用来进行复制的原始数据列对象在集合中的下标值。
            /// 注意：第一条曲线的下标值为1，而不是0。</param>
            /// <remarks>此方法主要完成两步操作：
            /// 1.在Excel图表中对数据列及其对应的对象进行复制；
            /// 2.在字典索引dicSreies_Tag中新添加一项；
            /// </remarks>
            public virtual System.Collections.Generic.KeyValuePair<int, SeriesTag> CopySeries(int SourceSeriesIndex)
            {

                //生成新的数据列
                Excel.Series newSeries = default(Excel.Series);
                int NewSeriesIndex = 0;
                newSeries = AddSeries(ref NewSeriesIndex);

                //设置数据列格式
                Excel.Series originalSeries = F_DicSeries_Tag[SourceSeriesIndex].series;
                //如果此处报错，则用
                //Dim xv As Object = originalSeries.XValues, 然后将xv赋值给下面的XValue
                newSeries.XValues = originalSeries.XValues;
                newSeries.Values = originalSeries.Values;
                newSeries.Name = originalSeries.Name;

                //设置数据列对应的Tag信息
                this.F_DicSeries_Tag.Add(NewSeriesIndex, new SeriesTag(newSeries, this.F_RollingDate));
                //
                return new System.Collections.Generic.KeyValuePair<int, SeriesTag>(NewSeriesIndex, new SeriesTag(newSeries, this.F_RollingDate));
            }

            /// <summary>
            /// 添加新的数据系列对象
            /// </summary>
            /// <param name="NewSeriesIndex">新添加的数据曲线在集合中的下标值</param>
            /// <returns>一条新的数据曲线</returns>
            /// <remarks>此函数只创建出对于新的数据曲线的对象索引，以及设置曲线的UI样式，
            /// 并不涉及对于数据曲线的坐标点的赋值</remarks>
            protected Excel.Series AddSeries(ref int NewSeriesIndex)
            {
                Excel.Series NewS = null;
                var seriColl = this.Chart.SeriesCollection() as Excel.SeriesCollection;
                if (F_CurvesCount <= F_List_HasCurve.Count - 1) //直接从已经定义好的模板中去提取
                {
                    bool hasCurve = false;
                    for (NewSeriesIndex = cst_LboundOfSeriesInCollection; NewSeriesIndex <= System.Convert.ToInt32(F_List_HasCurve.Count + cst_LboundOfSeriesInCollection) - 1; NewSeriesIndex++)
                    {
                        hasCurve = F_List_HasCurve[NewSeriesIndex - cst_LboundOfSeriesInCollection];
                        if (!hasCurve)
                        {

                            NewS = seriColl.Item(NewSeriesIndex);
                            F_List_HasCurve[NewSeriesIndex - cst_LboundOfSeriesInCollection] = true;
                            break;
                        }
                    }
                }
                else //如果图表中的曲线数据已经大于当前数据系列集合中的series数量，那么就要新建一个数据系列，并设置其格式
                {
                    NewS = seriColl.NewSeries();
                    NewSeriesIndex = System.Convert.ToInt32(F_List_HasCurve.Count + cst_LboundOfSeriesInCollection);
                    F_List_HasCurve.Add(true);
                    //'设置新数据系列的UI格式
                    //With NewS

                    //End With
                }

                // ----------------------------------- 对图例的显示进行操作
                LegendRefresh(F_List_HasCurve);

                //
                F_CurvesCount++;
                return NewS;
            }
            /// <summary>
            /// 更新图例，图例的绝对尺寸的更新一定要在设置要Chart的尺寸之后，
            /// 因为如果在设置好图例尺寸后再设置Chart尺寸，则图例尺寸会进行缩放。
            /// </summary>
            /// <param name="lst"></param>
            /// <remarks></remarks>
            protected void LegendRefresh(List<bool> lst)
            {
                Excel.Legend lgd = default(Excel.Legend);
                Excel.LegendEntries lgdEntries = default(Excel.LegendEntries);
                Excel.LegendEntry lgdEnrty = default(Excel.LegendEntry);
                this.Chart.HasLegend = false;
                this.Chart.HasLegend = true;
                lgd = this.Chart.Legend;
                lgdEntries = lgd.LegendEntries() as Excel.LegendEntries;

                //一定要注意，这里对图例项进行删除的时候，要从尾部开始向开头位置倒着删除:
                //因为集合的Index的索引方式, 当元素被删除后, 其后面的元素就接替了这个元素的下标值
                for (int LegendIndex =
                    F_List_HasCurve.Count - 1 + cst_LboundOfSeriesInCollection; LegendIndex >= cst_LboundOfSeriesInCollection; LegendIndex--)
                {

                    bool hascurve = System.Convert.ToBoolean(F_List_HasCurve[LegendIndex - cst_LboundOfSeriesInCollection]);
                    if (!hascurve)
                    {
                        //在Visual Basic.NET中，如果用LegendEntries(Index)来索引LegendEntry对象，其第一个元素的下标值为0，
                        //而如果用LegendEntries.Item(Index)的方式来索引集合中的LegendEntry对象，则其第一个元素的下标值为1。
                        //而在VBA中，这两种方式索引集合中的LegendEntry对象，其第一个元素的下标值都是1；
                        lgdEnrty = lgdEntries.Item(LegendIndex);
                        lgdEnrty.Delete();
                    }
                }
                // 设置图例对象的位置与尺寸
                this.SetLegendFormat(this.Chart, lgd);
                this.Sheet_Drawing.Range("A1").Activate();
            }
            /// <summary>
            /// 设置图例对象的UI格式：图例方框的线型、背景、位置与大小等
            /// </summary>
            /// <param name="chart">图例对象所属的Chart对象</param>
            /// <param name="legend">要进行UI格式设置的图例对象</param>
            /// <remarks></remarks>
            private void SetLegendFormat(Excel.Chart chart, Excel.Legend legend)
            {
                //开始设置图例的格式
                //1、设置图例的格式
                legend.Format.Fill.Visible = Office.MsoTriState.msoTrue;
                legend.Format.Fill.ForeColor.RGB = Information.RGB(255, 255, 255); //图例对象的背景色
                legend.Format.Shadow.Type = Microsoft.Office.Core.MsoShadowType.msoShadow21; //图例的阴影
                                                                                             //.Line.ForeColor.RGB = RGB(0, 0, 0)      ' 图例对象的边框
                legend.Format.TextFrame2.TextRange.Font.Name = AMEApplication.FontName_TNR;
                //2、设置图例的位置与大小
                //图例方框的高度
                float ExpectedHeight = System.Convert.ToSingle(this.Legend_Location.Legend_Height); // Data_Drawing_Format.Drawing_Mnt_RollingBase.Legend_Height
                legend.Select();
                //对于Chart中的图例的位置与大小，其设置的原则是：
                //对于Top值，其始终不改变图例的大小，而将图例对象一直向下推进，直到图例对象的下边缘到达Chart的下边缘为止；
                //对于Height值，其始终不改变图例的Top的位置，而将图例对象的下边缘一直向下拉伸，直到图例对象的下边缘到达Chart的下边缘为止；
                //所以，如果要将图例对象成功地定位到Chart底部，应该先设置其Height值为0，然后设置其Top值，最后再回过来设置其Height值。
                var selection = legend.Application.Selection  as Excel.Range;
                selection.Left = 0;
                selection.Width = this.Legend_Location.Legend_Width;
                selection.Height = 0; //以此来让Selection的Top值成功定位
                selection.Top = chart.ChartArea.Height - ExpectedHeight;
                selection.Height = ExpectedHeight;
            }
            #endregion
        }
    }
}
