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
using CableStayedBridge.DataBase;
using CableStayedBridge.GlobalApp_Form;
using CableStayedBridge.Miscellaneous;
// End of VB project level imports

using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using CableStayedBridge.Constants;

//using DAMIE.All_Drawings_In_Application.ClsDrawing_Mnt_Base;
//using DAMIE.Constants.Data_Drawing_Format.Mnt_Others;
//using DAMIE.Constants.Data_Drawing_Format.Drawing_Mnt_Others;

namespace CableStayedBridge
{
    public partial class frmDrawing_Mnt_Others
    {

        #region   ---  Declarations and Definitions

        /// <summary>
        /// 当前进行绘图的数据工作簿发生变化时触发
        /// </summary>
        /// <remarks></remarks>
        private delegate void DataWorkbookChangedEventHandler(Excel.Workbook WorkingDataWorkbook);
        private DataWorkbookChangedEventHandler DataWorkbookChangedEvent;

        private event DataWorkbookChangedEventHandler DataWorkbookChanged
        {
            add
            {
                DataWorkbookChangedEvent = (DataWorkbookChangedEventHandler)System.Delegate.Combine(DataWorkbookChangedEvent, value);
            }
            remove
            {
                DataWorkbookChangedEvent = (DataWorkbookChangedEventHandler)System.Delegate.Remove(DataWorkbookChangedEvent, value);
            }
        }


        #region   ---  Fields
        /// <summary>
        /// 当前用于绘图的数据工作簿
        /// </summary>
        /// <remarks></remarks>
        Excel.Workbook F_wkbkData;
        Excel.Workbook F_wkbkDrawing;
        Excel.Worksheet F_shtMonitorData;
        Excel.Worksheet F_shtDrawing;

        /// <summary>
        /// 进行操作的工作表的UsedRange的右下角的行号与列号
        /// </summary>
        /// <remarks></remarks>
        int[] F_arrBottomRightCorner = new int[2];

        /// <summary>
        /// ！以每一天的日期来索引这一天的监测数据，监测数据只包含列表中选择了的监测点
        /// </summary>
        /// <remarks></remarks>
        Dictionary<DateTime, object[]> F_dicDate_ChosenDatum = new Dictionary<DateTime, object[]>();

        /// <summary>
        /// 用来进行绘图的Excel程序
        /// </summary>
        /// <remarks></remarks>
        Excel.Application F_AppDrawing;

        //画布上的变量
        Excel.Chart F_Chart;
        Excel.TextFrame2 F_textbox_Info;
        //
        /// <summary>
        /// 在测点列表框中所选择的所有测点标签的数组，如{SW1,SW2,SW3}
        /// </summary>
        /// <remarks></remarks>
        private string[] F_SelectedTags;
        /// <summary>
        /// 在测点列表框中所选择的每一个测点在Excel工作表中所对应的行号
        /// </summary>
        /// <remarks></remarks>
        private int[] F_RowNum_SelectedTags;
        //
        /// <summary>
        /// Chart中第一条监测曲线所对应的相关信息
        /// </summary>
        /// <remarks></remarks>
        private clsDrawing_Mnt_RollingBase.SeriesTag F_TheFirstseriesTag;
        //
        /// <summary>
        /// 所绘图的监测数据类型
        /// </summary>
        /// <remarks></remarks>
        private MntType F_MonitorType;
        //
        private APPLICATION_MAINFORM F_MainForm; // VBConversions Note: Initial value cannot be assigned here since it is non-static.  Assignment has been moved to the class constructors.

        #endregion

        #endregion

        #region   ---  窗口的加载与关闭

        public frmDrawing_Mnt_Others()
        {

            // This call is required by the designer.
            InitializeComponent();

            // Add any initialization after the InitializeComponent() call.
            GeneralMethods.SetMonitorType(this.ComboBox_MntType);
            ClsData_DataBase.WorkingStageChanged += this.RefreshCombox_WorkingStage;

            // ------------
            this.ComboBoxOpenedWorkbook.DisplayMember = LstbxDisplayAndItem.DisplayMember;
            this.ComboBoxOpenedWorkbook.ValueMember = LstbxDisplayAndItem.ValueMember;

        }

        //窗口加载
        public void frmDrawingMonitor_Load(object sender, EventArgs e)
        {

            //设置控件的默认属性
            btnGenerate.Enabled = false;
            chkBoxOpenNewExcel.Checked = true;
            RbtnDynamic.Checked = true;
        }
        //在关闭窗口时将其隐藏
        public void frmDrawing_Mnt_Others_FormClosing(object sender, FormClosingEventArgs e)
        {
            //如果是子窗口自己要关闭，则将其隐藏
            //如果是mdi父窗口要关闭，则不隐藏，而由父窗口去结束整个进程
            if (!(e.CloseReason == CloseReason.MdiFormClosing))
            {
                this.Hide();
            }
            e.Cancel = true;
        }

        public void frmDrawing_Mnt_Others_Disposed(object sender, EventArgs e)
        {
            ClsData_DataBase.WorkingStageChanged -= this.RefreshCombox_WorkingStage;
        }

        #endregion

        /// <summary>
        /// 生成监测曲线图
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <remarks></remarks>
        public void btnGenerate_Click(object sender, EventArgs e)
        {
            int SelectedTagsCount = ListBoxPointsName.SelectedIndices.Count;
            if (SelectedTagsCount > 0)
            {
                //
                F_SelectedTags = new string[SelectedTagsCount - 1 + 1];
                F_RowNum_SelectedTags = new int[SelectedTagsCount - 1 + 1];
                int i = 0;
                foreach (LstbxDisplayAndItem item in this.ListBoxPointsName.SelectedItems)
                {
                    F_SelectedTags[i] = item.DisplayedText;
                    F_RowNum_SelectedTags[i] = (int)item.Value;
                    i++;
                }

                //开始绘图
                if (!this.BGWK_NewDrawing.IsBusy)
                {

                    //用来判断是否要创建新的Excel程序来进行绘图，以及是否要对新画布所在的Excel进行美化。
                    var blnNewExcelApp = GlobalApplication.Application.MntDrawing_ExcelApps.Count == 0 || chkBoxOpenNewExcel.Checked;
                    this.BGWK_NewDrawing.RunWorkerAsync(blnNewExcelApp);
                }
            }
            else
            {
                MessageBox.Show("请选择至少一个测点。", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #region   ---  后台线程进行操作

        /// <summary>
        /// 生成绘图，此方法是在后台的工作者线程中执行的。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <remarks></remarks>
        public void BGW_Generate_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            //主程序界面的进度条的UI显示
            F_MainForm.ShowProgressBar_Marquee();
            //执行具体的绘图操作
            try
            {
                //用来判断是否要创建新的Excel程序来进行绘图，以及是否要对新画布所在的Excel进行美化。
                bool blnNewExcelApp = System.Convert.ToBoolean(e.Argument);
                Generate(blnNewExcelApp, this.F_blnDrawDynamic, this.F_SelectedTags, this.F_RowNum_SelectedTags);

            }
            catch (Exception ex)
            {
                MessageBox.Show("绘制监测曲线图失败！" + "\r\n" + ex.Message + "\r\n" + "报错位置：" + ex.TargetSite.Name,
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 当后台的工作者线程结束（即BGW_Generate_DoWork方法执行完毕）时触发，注意，此方法是在UI线程中执行的。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <remarks></remarks>
        public void BGW_Generate_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {
            //在绘图完成后，隐藏进度条
            F_MainForm.HideProgress("Done");
        }

        #endregion

        /// <summary>
        /// 1、开始绘图
        /// </summary>
        /// <param name="CreateNewExcelApp">是否要新开一个Excel程序</param>
        /// <param name="DrawDynamicDrawing">是否是要绘制动态曲线图</param>
        /// <param name="SelectedTags">选择的测点</param>
        /// <param name="RowNum_SelectedTags">选择的测点在Excel数据工作表中所在的行号</param>
        /// <remarks></remarks>
        private void Generate(bool CreateNewExcelApp, bool DrawDynamicDrawing,
            string[] SelectedTags, int[] RowNum_SelectedTags)
        {

            // ----------arrDateRange----------------------- 获取此工作表中的整个施工日期的数组（0-Based，数据类型为Date）
            DateTime[] arrDateRange = GetDate();
            double[] arrDateRange_Double = new double[arrDateRange.Count() - 1 + 1];
            for (int i = 0; i <= arrDateRange.Count() - 1; i++)
            {
                arrDateRange_Double[i] = arrDateRange[i].ToOADate();
            }
            //-----------AllSelectedMonitorData------------ 获取所有"选择的"监测点位的监测数据的大数组。其中不包含选择的测点编号信息与施工日期的信息
            //此数组的第一个元素的下标值为0
            object[,] AllSelectedMonitorData = null;
            AllSelectedMonitorData = GetAllSelectedMonitorData(this.F_shtMonitorData, RowNum_SelectedTags);

            //----------------- 设置监测曲线的时间跨度
            Array.Sort(arrDateRange);
            DateSpan Date_Span = new DateSpan();
            Date_Span.StartedDate = arrDateRange[0];
            Date_Span.FinishedDate = arrDateRange[arrDateRange.Length - 1];

            //----------------------------打开用来绘图的Excel程序，并将此界面加入主程序的监测曲线集合
            Cls_ExcelForMonitorDrawing clsExcelForMonitorDrawing = null;
            //   --------------- 获取用来绘图的Excel程序，并将此界面加入主程序的监测曲线集合 -------------------
            F_AppDrawing = GetApplication(NewExcelApp: CreateNewExcelApp, ExcelForMntDrawing: clsExcelForMonitorDrawing, MntDrawingExcelApps:
                GlobalApplication.Application.MntDrawing_ExcelApps);
            F_AppDrawing.ScreenUpdating = false;
            //打开工作簿以画图
            if (F_AppDrawing.Workbooks.Count == 0)
            {
                F_wkbkDrawing = F_AppDrawing.Workbooks.Add();
            }
            else
            {
                F_wkbkDrawing = F_AppDrawing.Workbooks[1]; //总是定义为第一个，因为就只开了一个
            }
            //新开一个工作表以画图
            F_shtDrawing = F_wkbkDrawing.Worksheets.Add();
            F_shtDrawing.Activate();


            //-------  根据是要绘制动态的曲线图还是静态的曲线图，来执行不同的操作  ---------------------
            if (DrawDynamicDrawing)
            {

                //-------------绘制动态的曲线图--------------
                F_dicDate_ChosenDatum = GetdicDate_Datum_ForDynamic(arrDateRange, AllSelectedMonitorData);
                //开始画图
                F_Chart = DrawDynamicChart(F_dicDate_ChosenDatum, SelectedTags, AllSelectedMonitorData);
                //设置图表的Tag属性
                MonitorInfo Tags = GetChartTags(F_shtMonitorData);
                //-------------------------------------------------------------------------
                ClsDrawing_Mnt_OtherDynamics DynamicSheet = new ClsDrawing_Mnt_OtherDynamics(F_shtMonitorData, F_Chart,
                    clsExcelForMonitorDrawing, Date_Span,
                    DrawingType.Monitor_Dynamic, true, F_textbox_Info, Tags, this.F_MonitorType,
                    F_dicDate_ChosenDatum, this.F_TheFirstseriesTag);

                //-------------------------------------------------------------------------
            }
            else
            {

                //-------绘制静态的曲线图--------
                //开始画图
                Dictionary<Excel.Series, object[]> dicSeriesData = new Dictionary<Excel.Series, object[]>();
                F_Chart = DrawStaticChart(SelectedTags, arrDateRange_Double, AllSelectedMonitorData, dicSeriesData);
                //设置图表的Tag属性
                MonitorInfo Tags = GetChartTags(F_shtMonitorData);
                if (this.F_WorkingStage != null)
                {
                    DrawWorkingStage(this.F_Chart, this.F_WorkingStage);
                }
                //-------------------------------------------------------------------------
                ClsDrawing_Mnt_Static staticSheet = new ClsDrawing_Mnt_Static(F_shtMonitorData, F_Chart,
                    clsExcelForMonitorDrawing,
                    DrawingType.Monitor_Static, false, F_textbox_Info, Tags, this.F_MonitorType,
                    dicSeriesData, arrDateRange_Double);

                //-------------------------------------------------------------------------
            }

            //---------------------- 界面显示与美化
            ExcelAppBeauty(F_shtDrawing.Application, CreateNewExcelApp);
        }

        #region   ---  动态图

        /// <summary>
        /// 绘制动态曲线图
        /// </summary>
        /// <param name="p_dicDate_ChosenDatum"></param>
        /// <param name="arrChosenTags"></param>
        /// <param name="arrDataDisplacement"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        private Chart DrawDynamicChart(Dictionary<DateTime, object[]> p_dicDate_ChosenDatum,
            string[] arrChosenTags,
            object[,] arrDataDisplacement)
        {

            //dic_Date_ChosenDatum以每一天的日期来索引这一天的监测数据，监测数据只包含列表中选择了的监测点
            Chart r_myChart = default(Chart);
            F_shtDrawing.Activate();
            //--------------------------------------------------------------- 在工作表“标高图”中添加图表
            r_myChart = F_shtDrawing.Shapes.AddChart(XlChartType.xlLineMarkers).Chart;

            //---------- 选定模板
            string t_path = System.IO.Path.Combine(System.Convert.ToString(My.Settings.Default.Path_Template),
                Constants.FolderOrFileName.File_Template.Chart_Horizontal_Dynamic);
            // 如果监测曲线图所使用的"Chart模板"有问题，则在chart.ChartArea.Copy方法（或者是chartObject.Copy方法）中可能会出错。
            r_myChart.ApplyChartTemplate(t_path);
            //-------------------- 获取图表中的信息文本框
            F_textbox_Info = r_myChart.Shapes[0].TextFrame2; //Chart中的Shapes集合的第一个元素的下标值为0
                                                             //textbox_Info.AutoSize = Microsoft.Office.Core.MsoAutoSize.msoAutoSizeShapeToFitText


            //------------------------ 设置曲线的数据
            DateTime Date_theFirstCurve = System.Convert.ToDateTime(p_dicDate_ChosenDatum.Keys(0));
            SeriesCollection mySeriesCollection = r_myChart.SeriesCollection();
            Series series = mySeriesCollection.Item(1);
            series.Name = Date_theFirstCurve.ToString(); //系列名称
            series.XValues = arrChosenTags; //X轴的数据
            series.Values = p_dicDate_ChosenDatum.Item(Date_theFirstCurve); //Y轴的数据
                                                                            //
            this.F_TheFirstseriesTag = new clsDrawing_Mnt_RollingBase.SeriesTag(series, Date_theFirstCurve);
            //------------------------ 设置X、Y轴的格式——监测点位编号
            dynamic with_3 = r_myChart.Axes(XlAxisType.xlCategory);
            with_3.AxisTitle.Text = GetAxisLabel(DrawingType.Monitor_Dynamic, this.F_MonitorType, XlAxisType.xlCategory);

            //-设置Y轴的格式——测点位移
            Axis axes = r_myChart.Axes(XlAxisType.xlValue);

            //由数据的最小与最大值来划分表格区间
            float imin = (float)(F_AppDrawing.WorksheetFunction.Min(arrDataDisplacement));
            float imax = (float)(F_AppDrawing.WorksheetFunction.Max(arrDataDisplacement));

            //主要与次要刻度单位，先确定刻度单位是为了后面将坐标轴的区间设置为主要刻度单位的倍数
            float unit = float.Parse(Strings.Format((imax - imin) / ClsDrawing_Mnt_OtherDynamics.cstChartParts_Y, "0.0E+00")); //这里涉及到有效数字的处理的问题
            axes.MajorUnit = unit;
            axes.MinorUnitIsAuto = true;

            //坐标轴上显示的总区间
            axes.MinimumScale = F_AppDrawing.WorksheetFunction.Floor_Precise(imin, axes.MajorUnit);
            axes.MaximumScale = F_AppDrawing.WorksheetFunction.Ceiling_Precise(imax, axes.MajorUnit);


            //坐标轴标题
            axes.AxisTitle.Text = GetAxisLabel(DrawingType.Monitor_Dynamic, this.F_MonitorType, XlAxisType.xlValue);
            return r_myChart;
        }

        /// <summary>
        /// 返回动态曲线图的监测数据中的字典：以每一天的日期，索引选择的测点在当天的数据。
        /// </summary>
        /// <param name="arrDateRange"></param>
        /// <param name="AllSelectedMonitorData"></param>
        /// <returns>返回动态曲线图的关键参数：dic_Date_ChosenDatum，
        /// 字典中的值这里只能定义为Object的数组，因为有可能单元格中会出现没有数据的情况。</returns>
        /// <remarks></remarks>
        private Dictionary<DateTime, object[]> GetdicDate_Datum_ForDynamic(DateTime[] arrDateRange,
            object[,] AllSelectedMonitorData)
        {
            //以每一天的日期来索引这一天的监测数据，监测数据只包含列表中选择了的监测点
            Dictionary<DateTime, object[]> dic_Date_ChosenDatum = new Dictionary<DateTime, object[]>();
            int iCol = 0;

            //
            bool blnIgnoreWarning = false;
            for (iCol = 0; iCol <= Information.UBound((System.Array)AllSelectedMonitorData, 2); iCol++)
            {
                //当天的监测数据，这里只能定义为Object的数组，因为有可能单元格中会出现没有数据的情况。
                object[] arrDataForEachColumn = new object[(AllSelectedMonitorData.Length - 1) + 1];

                for (byte irow = 0; irow <= (AllSelectedMonitorData.Length - 1); irow++)
                {
                    arrDataForEachColumn[irow] = AllSelectedMonitorData[irow, iCol];
                }
                try //将新的一天的日期以及对应的监测数据信息添加到字典中
                {
                    dic_Date_ChosenDatum.Add(arrDateRange[iCol], arrDataForEachColumn);
                }
                catch (ArgumentException ex) //可能的报错原因：已添加了具有相同键的项。即工作表中有相同的日期
                {
                    if (!blnIgnoreWarning) //提醒用户
                    {
                        System.Windows.Forms.DialogResult result = MessageBox.Show("此工作表中的数据类型不规范，可能在监测数据中出现了相同的日期" + "\r\n"
                            + "出错的监测数据日期是：" + arrDateRange[iCol].ToShortDateString() + " . 是否忽略此提醒？"
                            + "\r\n"
                            + ex.Message + "\r\n" + ex.TargetSite.Name, "Warning", MessageBoxButtons.YesNo);
                        switch (result)
                        {
                            case System.Windows.Forms.DialogResult.No:
                                continue;
                                break;
                            case System.Windows.Forms.DialogResult.Yes: //继续添加而不再报错
                                blnIgnoreWarning = true;
                                break;
                        }
                    }

                    //btnGenerate.Enabled = False
                }
            }

            //-------------------------------------
            return dic_Date_ChosenDatum;
        }

        #endregion

        #region   ---  静态图

        /// <summary>
        /// 绘制静态曲线图
        /// </summary>
        /// <param name="arrChosenTags"></param>
        /// <param name="arrDateRange"></param>
        /// <param name="AllSelectedMonitorData"></param>
        /// <param name="dicSeries_Data"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        private Chart DrawStaticChart(string[] arrChosenTags,
            double[] arrDateRange, object[,] AllSelectedMonitorData, Dictionary<Series,
            object[]> dicSeries_Data)
        {
            //
            Chart r_myChart = default(Chart);
            F_shtDrawing.Activate();
            //--------------------------------------------------------------- 在工作表“标高图”中添加图表
            r_myChart = F_shtDrawing.Shapes.AddChart().Chart;

            //---------- 选定模板
            string t_path = System.IO.Path.Combine(System.Convert.ToString(My.Settings.Default.Path_Template),
                Constants.FolderOrFileName.File_Template.Chart_Horizontal_Static);
            // 如果监测曲线图所使用的"Chart模板"有问题，则在chart.ChartArea.Copy方法（或者是chartObject.Copy方法）中可能会出错。
            r_myChart.ApplyChartTemplate(t_path);

            //-------------------- 获取图表中的信息文本框
            F_textbox_Info = r_myChart.Shapes[0].TextFrame2; //Chart中的Shapes集合的第一个元素的下标值为0
                                                             //textbox_Info.AutoSize = Microsoft.Office.Core.MsoAutoSize.msoAutoSizeShapeToFitText


            //----------------------------- 对于数据系列的集合，开始为每一行数据添加新的曲线
            SeriesCollection mySeriesCollection = r_myChart.SeriesCollection();
            Series eachSeries = default(Series);
            if (mySeriesCollection.Count < arrChosenTags.Length)
            {
                for (byte i = 1; i <= arrChosenTags.Length - mySeriesCollection.Count; i++)
                {
                    mySeriesCollection.NewSeries();
                }
            }
            else
            {
                for (var i = 1; i <= mySeriesCollection.Count - arrChosenTags.Length; i++) //删除集合中的元素，先锁定要删的那个元素的下标，然后删n次（因为每次删除后其后面的又填补上来了）
                {
                    mySeriesCollection[arrChosenTags.Length].delete();
                }
            }
            byte i2 = (byte)0;
            foreach (string tag in arrChosenTags)
            {
                eachSeries = mySeriesCollection[i2]; //数据系列集合中的第一条曲线的下标值为0
                object[] arrselectedDataWithDate = new object[Information.UBound((System.Array)AllSelectedMonitorData, 2) + 1];
                for (int i_col = 0; i_col <= Information.UBound((System.Array)AllSelectedMonitorData, 2); i_col++)
                {
                    arrselectedDataWithDate[i_col] = AllSelectedMonitorData[i2, i_col];
                }
                eachSeries.Name = arrChosenTags[i2]; //系列名称
                eachSeries.XValues = arrDateRange; //X轴的数据:每一天的施工日期
                eachSeries.Values = arrselectedDataWithDate; //Y轴的数据
                dicSeries_Data.Add(eachSeries, arrselectedDataWithDate);
                i2++;
            }

            //------------------------ 设置X、Y轴的格式
            //——整个施工日期跨度
            Axis axesX = r_myChart.Axes(XlAxisType.xlCategory);
            axesX.CategoryType = XlCategoryType.xlTimeScale;
            //绘制坐标轴的数值区间
            axesX.MaximumScale = System.Convert.ToDouble(Max_Array<double>(arrDateRange));
            axesX.MinimumScale = System.Convert.ToDouble(Min_Array<double>(arrDateRange));
            //.MaximumScaleIsAuto = True
            //.MinimumScaleIsAuto = True

            axesX.TickLabels.NumberFormatLocal = "yy/M/d";
            axesX.TickLabelSpacingIsAuto = true;
            axesX.AxisTitle.Text = GetAxisLabel(DrawingType.Monitor_Static, this.F_MonitorType, XlAxisType.xlCategory);

            //-设置Y轴的格式——测点位移
            Axis axesY = r_myChart.Axes(XlAxisType.xlValue);
            //坐标轴标题
            axesY.AxisTitle.Text = GetAxisLabel(DrawingType.Monitor_Static, this.F_MonitorType, XlAxisType.xlValue);

            //由数据的最小与最大值来划分表格区间
            float imin = (float)(F_AppDrawing.WorksheetFunction.Min(AllSelectedMonitorData));
            float imax = (float)(F_AppDrawing.WorksheetFunction.Max(AllSelectedMonitorData));

            //主要与次要刻度单位，先确定刻度单位是为了后面将坐标轴的区间设置为主要刻度单位的倍数
            float unit = float.Parse(Strings.Format((imax - imin) / ClsDrawing_Mnt_Static.cstChartParts_Y, "0.0E+00")); //这里涉及到有效数字的处理的问题
            try
            {
                axesY.MajorUnit = unit; //有可能会出现unit的值为0.0的情况
            }
            catch (Exception)
            {
                axesY.MajorUnitIsAuto = true;
            }
            axesY.MinorUnitIsAuto = true;

            //坐标轴上显示的总区间
            axesY.MinimumScale = F_AppDrawing.WorksheetFunction.Floor_Precise(imin, axesY.MajorUnit);
            axesY.MaximumScale = F_AppDrawing.WorksheetFunction.Ceiling_Precise(imax, axesY.MajorUnit);


            return r_myChart;
        }

        /// <summary>
        /// 绘制开挖工况的位置线
        /// </summary>
        /// <param name="Cht"></param>
        /// <param name="WorkingStage"></param>
        /// <remarks></remarks>
        private void DrawWorkingStage(Excel.Chart Cht, List<clsData_WorkingStage> WorkingStage)
        {
            Excel.Axis AX = Cht.Axes(Excel.XlAxisType.xlCategory) as Excel.Axis;

            string[] arrLineName = new string[WorkingStage.Count - 1 + 1];
            string[] arrTextName = new string[WorkingStage.Count - 1 + 1];
            Excel.Chart with_1 = Cht;
            try
            {
                int i = 0;
                foreach (clsData_WorkingStage WS in WorkingStage)
                {
                    // -------------------------------------------------------------------------------------------

                    Excel.Shape shpLine = with_1.Shapes.AddLine(BeginX: 0, BeginY: Cht.PlotArea.InsideTop, EndX:
                        0, EndY:
                        75);

                    shpLine.Line.Weight = (float)(1.2F);
                    //.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadStealth
                    //.EndArrowheadLength = Microsoft.Office.Core.MsoArrowheadLength.msoArrowheadLengthMedium
                    //.EndArrowheadWidth = Microsoft.Office.Core.MsoArrowheadWidth.msoArrowheadWidthMedium
                    shpLine.Line.DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineDashDot;
                    shpLine.Line.ForeColor.RGB = Information.RGB(255, 0, 0);
                    //
                    ExcelFunction.setPositionInChart(shpLine, AX, WS.ConstructionDate.ToOADate());
                    // -------------------------------------------------------------------------------------------
                    float TextWidth = 25;
                    float textHeight = 10;
                    Excel.Shape shpText = default(Excel.Shape);
                    shpText = Cht.Shapes.AddTextbox(Orientation: Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Left: shpLine.Left - TextWidth / 2, Top: shpLine.Top - textHeight, Height: ref textHeight, Width: ref TextWidth);
                    ExcelFunction.FormatTextbox_Tag(TextFrame: shpText.TextFrame2, Text: WS.Description, HorizontalAlignment: Microsoft.Office.Core.MsoParagraphAlignment.msoAlignCenter);
                    // -------------------------------------------------------------------------------------------
                    arrLineName[i] = shpLine.Name;
                    arrTextName[i] = shpText.Name;
                    i++;
                }
                Excel.Shape shp1 = Cht.Shapes.Range[arrLineName].Group();
                Excel.Shape shp2 = Cht.Shapes.Range[arrTextName].Group();
                Cht.Shapes.Range[new[] { shp1.Name, shp2.Name }].Group();
            }
            catch (Exception ex)
            {
                MessageBox.Show("设置开挖工况位置出现异常。" + "\r\n" + ex.Message + "\r\n" + "报错位置：" + ex.TargetSite.Name, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        #endregion

        #region   ---  通用子方法


        /// <summary>
        /// 获取用来绘图的Excel程序，并将此界面加入主程序的监测曲线集合
        /// </summary>
        /// <param name="NewExcelApp">按情况看是否要打开新的Application</param>
        /// <returns></returns>
        /// <remarks></remarks>
        private Application GetApplication(bool NewExcelApp, Dictionary_AutoKey<Cls_ExcelForMonitorDrawing> MntDrawingExcelApps, ref Cls_ExcelForMonitorDrawing ExcelForMntDrawing)
        {
            Application app = default(Application);
            if (NewExcelApp) //打开新的Excel程序
            {
                app = new Application();
                ExcelForMntDrawing = new Cls_ExcelForMonitorDrawing(app);
            }
            else //在原有的Excel程序上作图
            {
                ExcelForMntDrawing = MntDrawingExcelApps.Last.Value;
                ExcelForMntDrawing.ActiveMntDrawingSheet.RemoveFormCollection();
                app = ExcelForMntDrawing.Application;
            }
            return app;
        }

        /// <summary>
        /// 当进行绘图的数据工作簿发生变化时触发
        /// </summary>
        /// <param name="WorkingDataWorkbook">要进行绘图的数据工作簿</param>
        /// <remarks></remarks>
        public void frmDrawing_Mnt_Incline_DataWorkbookChanged(Excel.Workbook WorkingDataWorkbook)
        {
            //-------- 在列表中显示出监测数据工作簿中的所有工作表
            byte sheetsCount = (byte)WorkingDataWorkbook.Worksheets.Count;
            LstbxDisplayAndItem[] arrWorkSheets = new LstbxDisplayAndItem[sheetsCount - 1 + 1];
            byte i = (byte)0;
            foreach (Excel.Worksheet sht in WorkingDataWorkbook.Worksheets)
            {
                arrWorkSheets[i] = new LstbxDisplayAndItem(sht.Name, sht);
                i++;
            }
            listSheetsName.DisplayMember = LstbxDisplayAndItem.DisplayMember;
            listSheetsName.ValueMember = LstbxDisplayAndItem.ValueMember;
            listSheetsName.DataSource = arrWorkSheets;
            //.Items.Clear()
            //.Items.AddRange(sheetsNameList)
            //.SelectedItem = .Items(0)
            //
            btnGenerate.Enabled = true;
        }

        private DateTime[] GetDate()
        {
            // ----------arrDateRange----------------------- 获取此工作表中的整个施工日期的数组（0-Based，数据类型为Date）
            Excel.Range rg_AllDay = default(Excel.Range);
            //注意此数组元素的类型为Object，而且其第一个元素的下标值有可能是(1,1)
            short startColNum = ColNum_FirstData_Displacement;
            short endColNum = (short)(F_arrBottomRightCorner[1]);
            rg_AllDay = F_shtMonitorData.Rows[RowNumForDate].Range(F_shtMonitorData.Cells[1, startColNum],
                F_shtMonitorData.Cells[1, endColNum]);
            DateTime[] arrDateRange = null;
            arrDateRange = ExcelFunction.ConvertRangeDataToVector<DateTime>(rg_AllDay);
            //
            DateTime Dt = default(DateTime);
            try
            {
                Dictionary<DateTime, int> dicDate = new Dictionary<DateTime, int>();
                foreach (DateTime tempLoopVar_Dt in arrDateRange)
                {
                    Dt = tempLoopVar_Dt;
                    dicDate.Add(Dt, 0);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("工作表中的日期字段出错，出错的日期为：" + Dt.ToString("yyyy/MM/dd") + " 。请检查工作表！" + "\r\n" + ex.Message + "\r\n" + "报错位置：" + ex.TargetSite.Name, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
            return arrDateRange;
        }

        /// <summary>
        /// 返回所有选择的测点的监测数组组成的大数组
        /// </summary>
        /// <param name="RowNum_Tags"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        private object[,] GetAllSelectedMonitorData(Excel.Worksheet sheetMonitorData, int[] RowNum_Tags)
        {
            //Dim Queue_SelectedData As New Queue(Of Object())

            object[,] AllSelectedMonitorData = new object[RowNum_Tags.Length - 1 + 1, F_arrBottomRightCorner[1] - ColNum_FirstData_Displacement + 1];

            object[,] arrAllData = null; //所有有效的监测数据，包括没有选择的测点上的数据，不包括日期行与测点编号列
            Excel.Worksheet with_1 = sheetMonitorData;
            arrAllData = with_1.Range[with_1.Cells[RowNum_FirstData_WithoutDate, ColNum_FirstData_Displacement], with_1.Cells[F_arrBottomRightCorner[0], F_arrBottomRightCorner[1]]].Value;


            //---------------------- 所有选择的测点在工作表中对应的行号,并将行号对应到上面的数组arrAllData的行号
            //先要处理Excel中的数组与VB中的数组的下标差的问题
            byte diff_Row = RowNum_FirstData_WithoutDate - 0;
            byte diff_Col = ColNum_FirstData_Displacement - Information.LBound((System.Array)arrAllData, 2);


            //选择的测点编号所在的行对应到上面的ArrAllData数组中的行，第一个元素的下标为0
            byte Count_ChosenTag = (byte)RowNum_Tags.Length;
            string[] arrChosenRownNum = new string[Count_ChosenTag - 1 + 1];
            byte i2 = (byte)0;
            foreach (int RowNum_Tag in RowNum_Tags)
            {
                arrChosenRownNum[i2] = System.Convert.ToString(RowNum_Tag - diff_Row);
                i2++;
            }

            //------------ 为超级大数组赋值
            mySettings_Application setting = new mySettings_Application();
            bool blnCheckForEmpty = setting.CheckForEmpty;
            //Dim arrDataWithDate_EachTag(0 To UBound(arrAllData, 2) - LBound(arrAllData, 2) + 1) As Object
            if (!blnCheckForEmpty)
            {
                byte irow = (byte)0;
                foreach (short rownum in arrChosenRownNum)
                {
                    for (short iCol = Information.LBound((System.Array)arrAllData, 2); iCol <= Information.UBound((System.Array)arrAllData, 2); iCol++)
                    {
                        //arrDataWithDate_EachTag(iCol - LBound(arrAllData, 2)) = arrAllData(rownum, iCol)
                        AllSelectedMonitorData[irow, iCol - Information.LBound((System.Array)arrAllData, 2)] = arrAllData[rownum, iCol];
                    }
                    //Queue_SelectedData.Enqueue(arrDataWithDate_EachTag)
                    irow++;
                }
            }
            else
            {
                byte irow = (byte)0;
                foreach (byte rownum in arrChosenRownNum)
                {
                    for (int iCol = Information.LBound((System.Array)arrAllData, 2); iCol <= Information.UBound((System.Array)arrAllData, 2); iCol++)
                    {
                        //如果单元格中是空字符串，而不是Nothing，则将其转换为Nothing。
                        object v = arrAllData[rownum, iCol];
                        if (v != null)
                        {
                            if (v.GetType() == typeof(string))
                            {
                                if (v.ToString().Length == 0)
                                {
                                    v = null;
                                }
                            }
                        }
                        //赋值
                        //arrDataWithDate_EachTag(iCol - LBound(arrAllData, 2)) = v
                        AllSelectedMonitorData[irow, iCol - Information.LBound((System.Array)arrAllData, 2)] = v;
                    }
                    //Queue_SelectedData.Enqueue(arrDataWithDate_EachTag)
                    irow++;
                }
            }

            return AllSelectedMonitorData;
        }

        /// <summary>
        /// 设置图表的Tags属性
        /// </summary>
        /// <param name="MntDataSheet">监测数据所在的工作表</param>
        /// <remarks></remarks>
        private MonitorInfo GetChartTags(Excel.Worksheet MntDataSheet)
        {
            string MonitorItem = "";
            string ExcavationRegion = "";
            //
            Excel.Workbook t_wkbk = MntDataSheet.Parent as Excel.Workbook;
            string filepathwithoutextension = System.IO.Path.GetFileNameWithoutExtension(t_wkbk.FullName);
            MonitorItem = filepathwithoutextension;
            //
            ExcavationRegion = MntDataSheet.Name;
            //
            MonitorInfo Tags = new MonitorInfo(MonitorItem, ExcavationRegion);
            return Tags;
        }

        /// <summary>
        /// 程序界面美化
        /// </summary>
        /// <param name="app">用来进行开头与界面设置的Excel界面</param>
        /// <remarks>设置的元素包括Excel界面的尺寸、工具栏、滚动条等的显示方式</remarks>
        private void ExcelAppBeauty(Excel.Application app, bool blnCreateNewExcelApplication)
        {
            app.DisplayStatusBar = false;
            app.DisplayFormulaBar = false;

            app.ActiveWindow.DisplayGridlines = false;
            app.ActiveWindow.DisplayHeadings = false;
            app.ActiveWindow.DisplayWorkbookTabs = false;
            app.ActiveWindow.Zoom = 100;
            app.ActiveWindow.DisplayHorizontalScrollBar = false;
            app.ActiveWindow.DisplayVerticalScrollBar = false;
            app.ActiveWindow.WindowState = Excel.XlWindowState.xlMaximized;
            if (blnCreateNewExcelApplication) //如果是新创建的Excel界面，则进行美化，否则保持原样
            {
                app.Visible = true;
                app.WindowState = Excel.XlWindowState.xlNormal;
                app.ExecuteExcel4Macro("SHOW.TOOLBAR(\"Ribbon\",false)");
            }
            app.ScreenUpdating = true;
        }

        #endregion

        #region   ---  界面操作

        /// <summary>
        /// 选择监测数据的文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <remarks></remarks>
        public void btnChooseMonitorData_Click(object sender, EventArgs e)
        {
            string FilePath = "";
            APPLICATION_MAINFORM.MainForm.OpenFileDialog1.Title = "选择测斜数据文件";
            APPLICATION_MAINFORM.MainForm.OpenFileDialog1.Filter = "Excel文件(*.xlsx, *.xls, *.xlsb)|*.xlsx;*.xls;*.xlsb";
            APPLICATION_MAINFORM.MainForm.OpenFileDialog1.FilterIndex = 2;
            if (APPLICATION_MAINFORM.MainForm.OpenFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                FilePath = APPLICATION_MAINFORM.MainForm.OpenFileDialog1.FileName;
            }
            else
            {
                return;
            }
            if (FilePath.Length > 0)
            {
                //将监测数据文件在DataBase的Excel程序中打开
                try
                {
                    //有可能会出现选择了同样的监测数据文档
                    bool fileHasOpened = false;
                    foreach (LstbxDisplayAndItem item in this.ComboBoxOpenedWorkbook.Items)
                    {
                        Excel. Workbook wkbk = (Excel.Workbook)item.Value;
                        if (string.Compare(wkbk.FullName, FilePath, true) == 0)
                        {
                            this.F_wkbkData = wkbk;
                            fileHasOpened = true;
                            break;
                        }
                    }
                    // ----------------------------
                    if (fileHasOpened)
                    {
                        MessageBox.Show("选择的工作簿已经打开", "Tip", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        this.F_wkbkData = GlobalApplication.Application.ExcelApplication_DB.Workbooks.Open(Filename:  FilePath, UpdateLinks: false, ReadOnly: true);
                        LstbxDisplayAndItem lstItem = new LstbxDisplayAndItem(this.F_wkbkData.Name, this.F_wkbkData);
                        this.ComboBoxOpenedWorkbook.Items.Add(lstItem);
                        this.ComboBoxOpenedWorkbook.SelectedItem = lstItem;
                        if (DataWorkbookChangedEvent != null)
                            DataWorkbookChangedEvent(this.F_wkbkData);
                    }
                }
                catch (Exception)
                {
                    Debug.Print("打开新的数据工作簿出错！");
                    return;
                }
            }
        }

        /// <summary>
        /// 在Visio中绘制相应的测点位置
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <remarks></remarks>
        public void btnDrawMonitorPoints_Click(object sender, EventArgs e)
        {
            GlobalApplication.Application.DrawingPointsInVisio();
        }

        #region   ---  选择列表框内容时进行赋值

        /// <summary>
        /// 设置监测数据的类型
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <remarks></remarks>
        public void ComboBox_MntType_SelectedValueChanged(object sender, EventArgs e)
        {
            LstbxDisplayAndItem item = (LstbxDisplayAndItem)this.ComboBox_MntType.SelectedItem;
            this.F_MonitorType = (MntType)item.Value;
        }

        /// <summary>
        /// 当选择的数据工作表发生变化时引发
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <remarks></remarks>
        public void listSheetsName_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.F_shtMonitorData = (Excel.Worksheet)listSheetsName.SelectedValue; //wkbkData.Worksheets(listSheetsName.SelectedItem)
            string[] arrPointsTag = null; //此工作表中的测点的编号列表
                                          //测点编号所在的起始行与末尾行
            int startRow = RowNum_FirstData_WithoutDate;
            int endRow = F_shtMonitorData.UsedRange.Rows.Count;
            int endColumn = F_shtMonitorData.UsedRange.Columns.Count;
            F_arrBottomRightCorner = new[] { endRow, endColumn }; //'进行操作的工作表的UsedRange的右下角的行号与列号
            if (endRow >= startRow)
            {

                Excel.Range rg_PointsTag = F_shtMonitorData.Columns[ColNum_PointsTag].Range(F_shtMonitorData.Cells[startRow, 1], F_shtMonitorData.Cells[endRow, 1]);
                arrPointsTag = ExcelFunction.ConvertRangeDataToVector<string>(rg_PointsTag);


                //------------   将编号列表的每一个值与其对应的行号添加到字典dicPointTag_RowNum中
                short TagsCount = (short)arrPointsTag.Length;
                LstbxDisplayAndItem[] arrPoints = new LstbxDisplayAndItem[TagsCount - 1 + 1];

                int add = 0;
                int i = 0;
                foreach (string tag in arrPointsTag)
                {
                    //在Excel数据表中，每一个监测点位的Tag所在的行号。
                    short RowNumToPointsTag = (short)(startRow + add);
                    arrPoints[i] = new LstbxDisplayAndItem(tag, startRow + add);
                    add++;
                    i++;
                }

                //----------------------  将编号列表的所有值显示在窗口的测点列表中
                ListBoxPointsName.DisplayMember = LstbxDisplayAndItem.DisplayMember;
                ListBoxPointsName.ValueMember = LstbxDisplayAndItem.ValueMember;
                ListBoxPointsName.DataSource = arrPoints;
                //.Items.Clear()
                //.Items.AddRange(arr)
                //.SelectedIndex = 0          '选择列表中的第一项
                //
                btnGenerate.Enabled = true;
            }
            else //说明此表中的监测数据行数小于1
            {
                ListBoxPointsName.DataSource = null;
                //ListPointsName.Items.Clear()
                MessageBox.Show("此工作表没有合法数据（数据行数小于1行）,请选择合适的工作表", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error);
                F_shtMonitorData = null;
                btnGenerate.Enabled = false;
            }
        }

        /// <summary>
        /// 指示是要用此窗口来绘制动态图还是静态图
        /// </summary>
        /// <remarks></remarks>
        private bool F_blnDrawDynamic;
        public void RbtnStaticWithTime_CheckedChanged(object sender, EventArgs e)
        {
            if (RbtnDynamic.Checked)
            {
                this.Panel_Static.Visible = false;
                this.F_blnDrawDynamic = true;
            }
            else
            {
                this.Panel_Static.Visible = true;
                this.F_blnDrawDynamic = false;
            }

        }

        private List<clsData_WorkingStage> F_WorkingStage;
        public void ComboBox_WorkingStage_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.F_WorkingStage = null;
            LstbxDisplayAndItem lstItem = ComboBox_WorkingStage.SelectedItem as LstbxDisplayAndItem; ;
            if (lstItem != null)
            {
                if (!lstItem.Value.Equals(LstbxDisplayAndItem.NothingInListBox.None))
                {
                    F_WorkingStage = (List<clsData_WorkingStage>)lstItem.Value;
                }
            }
        }

        /// <summary>
        /// 选择进行绘图的数据工作簿
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <remarks></remarks>
        public void ComboBoxOpenedWorkbook_SelectedIndexChanged(object sender, EventArgs e)
        {
            LstbxDisplayAndItem lst = this.ComboBoxOpenedWorkbook.SelectedItem as LstbxDisplayAndItem;
            try
            {
              Excel.  Workbook Wkbk = (Excel.Workbook)lst.Value;
                this.F_wkbkData = Wkbk;
                APPLICATION_MAINFORM.MainForm.StatusLabel1.Visible = true;
                APPLICATION_MAINFORM.MainForm.StatusLabel1.Text = Wkbk.FullName;
                if (DataWorkbookChangedEvent != null)
                    DataWorkbookChangedEvent(this.F_wkbkData);
            }
            catch (Exception)
            {
                Debug.Print("选择数据工作簿出错");
            }
        }
        #endregion

        #region   ---  关联组合列表框中的数据

        private void RefreshCombox_WorkingStage(Dictionary<string, List<clsData_WorkingStage>> NewWorkingStage)
        {
            if (NewWorkingStage != null)
            {
                Dictionary<,> with_1 = NewWorkingStage;
                try
                {
                    var RegionNames = with_1.Keys;
                    var WorkingStages = with_1.Values;
                    int TagsCount = System.Convert.ToInt32(with_1.Count);
                    LstbxDisplayAndItem[] TagsList = new LstbxDisplayAndItem[TagsCount + 1];
                    TagsList[0] = new LstbxDisplayAndItem("无", LstbxDisplayAndItem.NothingInListBox.None);
                    //
                    for (int i1 = 0; i1 <= TagsCount - 1; i1++)
                    {
                        TagsList[i1 + 1] = new LstbxDisplayAndItem(System.Convert.ToString(RegionNames(i1)), WorkingStages(i1));
                    }
                    GeneralMethods.RefreshCombobox(this.ComboBox_WorkingStage, TagsList);
                }
                catch (Exception)
                {

                }
            }
        }

        #endregion


        #region   ---  窗口的激活去取消激活

        public void frmDrawing_Mnt_Incline_Activated(object sender, EventArgs e)
        {
            if (this.F_wkbkData != null)
            {
                APPLICATION_MAINFORM.MainForm.StatusLabel1.Visible = true;
                APPLICATION_MAINFORM.MainForm.StatusLabel1.Text = this.F_wkbkData.FullName;
            }

        }

        public void frmDrawing_Mnt_Incline_Deactivate(object sender, EventArgs e)
        {
            APPLICATION_MAINFORM.MainForm.StatusLabel1.Visible = false;
        }
        #endregion

        #endregion

    }

}
