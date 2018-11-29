// VBConversions Note: VB project level imports

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows.Forms;
using CableStayedBridge.All_Drawings_In_Application;
using CableStayedBridge.Constants;
using CableStayedBridge.GlobalApp_Form;
using CableStayedBridge.Miscellaneous;
using CableStayedBridge.My;
using eZstd.eZAPI;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using Visio = Microsoft.Office.Interop.Visio;
using Word = Microsoft.Office.Interop.Word;
using Application = Microsoft.Office.Interop.Word.Application;
using Chart = Microsoft.Office.Interop.Excel.Chart;
using Document = Microsoft.Office.Interop.Word.Document;
using ListBox = System.Windows.Forms.ListBox;
using Page = Microsoft.Office.Interop.Visio.Page;
using PageSetup = Microsoft.Office.Interop.Word.PageSetup;
using Path = System.IO.Path;
using Shape = Microsoft.Office.Interop.Excel.Shape;
using ShapeRange = Microsoft.Office.Interop.Word.ShapeRange;
using Window = Microsoft.Office.Interop.Visio.Window;
// End of VB project level imports

//using eZstd.eZAPI.APIWindows;

namespace CableStayedBridge
{
    public partial class Diafrm_Output_Word
    {
        #region   ---  Declarations & Definitions

        #region   ---  Types

        /// <summary>
        /// 所有选择的要进行输出的图形
        /// </summary>
        /// <remarks></remarks>
        private struct Drawings_For_Output
        {
            public ClsDrawing_PlanView PlanView;
            public ClsDrawing_ExcavationElevation SectionalView;
            public readonly List<ClsDrawing_Mnt_Base> MntDrawings;

            public short Count()
            {
                short SumUp = Convert.ToInt16(MntDrawings.Count);
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

            public Drawings_For_Output(Diafrm_Output_Word Sender)
            {
                MntDrawings = new List<ClsDrawing_Mnt_Base>();
                PlanView = null;
                SectionalView = null;
            }
        }

        #endregion

        #region   ---  Field定义

        //word页面中正文区域的宽度，用来限定图片的宽度
        private float ContentWidth;

        //“当天”的日期值
        private DateTime dateThisday;

        /// <summary>
        /// 窗口中的所有列表框listbox对象
        /// </summary>
        /// <remarks>此数组是为了便于后面的统一操作：清空内容、全部选择，取消全选</remarks>
        private readonly ListBox[] F_arrListBoxes = new ListBox[2];

        /// <summary>
        /// 所有选择的要进行输出的图形
        /// </summary>
        /// <remarks></remarks>
        private Drawings_For_Output F_SelectedDrawings;

        /// <summary>
        /// 程序中所有图表的窗口的句柄值，用来对窗口进行禁用或者启用窗口
        /// </summary>
        /// <remarks></remarks>
        private IntPtr[] WindowHandles;

        #endregion

        #region   ---  Properties

        /// <summary>
        /// Word的Application对象
        /// </summary>
        /// <remarks></remarks>
        private Application WdApp;

        /// <summary>
        /// Word的Application对象
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public Application Application
        {
            get { return WdApp; }
        }

        private Document WdDoc;

        /// <summary>
        /// 输出到的word.document对象
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public Document Document
        {
            get { return WdDoc; }
        }

        #endregion

        #endregion

        #region   ---  窗口的加载与关闭

        /// <summary>
        /// Showdialog式加载
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <remarks></remarks>
        public void frm_Output_Word_Load(object sender, EventArgs e)
        {
            //刷新时间
            dateThisday = APPLICATION_MAINFORM.MainForm.Form_Rolling.Rollingday;
            //设置初始界面
            LabelDate.Text = dateThisday.ToString("yyyy/MM/dd");
            ChkBxSelect.CheckState = CheckState.Unchecked;
            CheckBox_PlanView.Checked = false;
            CheckBox_SectionalView.Checked = false;
            btnExport.Enabled = false;
            //为数组中的每一个元素赋值，以便于后面的统一操作：清空内容、全部选择，取消全选
            F_arrListBoxes[0] = ListBoxMonitor_Dynamic;
            F_arrListBoxes[1] = ListBoxMonitor_Static;
            //
            F_SelectedDrawings = new Drawings_For_Output(this);
            //刷新主程序与界面
            AmeDrawings AllDrawing = GlobalApplication.Application.ExposeAllDrawings();
            // ---------- 禁用所有绘图窗口
            WindowHandles = GlobalApplication.GetWindwosHandles(AllDrawing);
            foreach (IntPtr H in WindowHandles)
            {
                APIWindows.EnableWindow(H, false);
            }
            //
            RefreshUI(AllDrawing);
        }

        /// <summary>
        /// 从主程序中提取所有的图表，以进行输出之用。同时刷新窗口中的可供输出的图形
        /// </summary>
        /// <param name="AllDrawing">主程序对象中所有的图表</param>
        /// <remarks>在此方法中，将提取主程序中的所有图表对象，而且将其显示在输入窗口的列表框中</remarks>
        private void RefreshUI(AmeDrawings AllDrawing)
        {
            //-------------------------1、剖面图---------------------------------------------
            ClsDrawing_ExcavationElevation Sectional = AllDrawing.SectionalView;
            CheckBox_SectionalView.Tag = Sectional;
            if (Sectional != null)
            {
                CheckBox_SectionalView.Enabled = true;
            }
            else
            {
                CheckBox_SectionalView.Checked = false;
                CheckBox_SectionalView.Enabled = false;
            }
            //--------------------------2、开挖平面图----------------------------------------------
            ClsDrawing_PlanView Plan = AllDrawing.PlanView;
            CheckBox_PlanView.Tag = Plan;
            if (Plan != null)
            {
                CheckBox_PlanView.Enabled = true;
            }
            else
            {
                CheckBox_PlanView.Checked = false;
                CheckBox_PlanView.Enabled = false;
            }

            //---------------------------------3、监测曲线图---------------------------------------

            //清空两个列表框中的所有项目
            foreach (ListBox lstbox in F_arrListBoxes)
            {
                lstbox.Items.Clear();
                lstbox.DisplayMember = LstbxDisplayAndItem.DisplayMember;
            }
            foreach (ClsDrawing_Mnt_Base sht in AllDrawing.MonitorData)
            {
                switch (sht.Type)
                {
                    case DrawingType.Monitor_Incline_Dynamic:
                    case DrawingType.Monitor_Dynamic:
                        ListBoxMonitor_Dynamic.Items.Add(new LstbxDisplayAndItem(sht.Chart_App_Title, sht));
                        break;

                    case DrawingType.Monitor_Static:
                    case DrawingType.Monitor_Incline_MaxMinDepth:
                        ListBoxMonitor_Static.Items.Add(new LstbxDisplayAndItem(sht.Chart_App_Title, sht));
                        break;
                }
            }
        }

        /// <summary>
        /// Word退出时，将对应的文档与word程序的变量设置为nothing
        /// </summary>
        /// <param name="docBeingClosed"></param>
        /// <param name="cancel"></param>
        /// <remarks></remarks>
        private void word_Quit(Document docBeingClosed, ref bool cancel)
        {
            float a = 0;
            if (docBeingClosed.Name == WdDoc.Name)
            {
                WdDoc = null;
                WdApp = null;
                WdApp.DocumentBeforeClose += word_Quit;
            }
        }

        public void Diafrm_Output_Word_FormClosed(object sender, FormClosedEventArgs e)
        {
            foreach (IntPtr H in WindowHandles)
            {
                APIWindows.EnableWindow(H, true);
            }
        }

        #endregion

        /// <summary>
        /// 将结果输出到word中
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <remarks></remarks>
        public void btnExport_Click(object sender, EventArgs e)
        {
            if (F_SelectedDrawings.Count() > 0)
            {
                if (BackgroundWorker1.IsBusy != true)
                {
                    // Start the asynchronous operation.
                    BackgroundWorker1.RunWorkerAsync(F_SelectedDrawings);
                }
            }
            else
            {
                return;
            }
        }

        #region   ---  后台线程的操作

        public void BackgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            Drawings_For_Output selectedDrawings = (Drawings_For_Output)e.Argument;

            //打开Word程序
            if (WdApp == null)
            {
                WdApp = new Application();
                WdApp.DocumentBeforeClose += word_Quit;
            }
            //指定进行输入的Word文档
            if (WdDoc == null)
            {
                //以模板文件打开新文档
                string t_path = Path.Combine(Convert.ToString(Settings.Default.Path_Template),
                    FolderOrFileName.File_Template.Word_Output);
                WdDoc = WdApp.Documents.Add(Template: t_path);
                //文档的正文宽度，用以限制图形的宽度
                ContentWidth = GetContentWidth(WdDoc);
            }
            //设置界面的可见性
            if (WdApp.Visible == true)
            {
                WdApp.Visible = true; //即保持原来的可见性
            }
            WdApp.ScreenUpdating = false;

            //  --------------------------- 输出 ------------------------------------

            ExportToWord(WdApp, selectedDrawings);

            //  --------------------------- 输出 ------------------------------------
            WdApp.Visible = true;
            WdApp.ScreenUpdating = true;
        }

        private void ExportToWord(Application WdApp, Drawings_For_Output selectedDrawings)
        {
            Word.Range rg = WdDoc.Range(Start: 0);
            //在写入标题部分内容时所占的进度
            int intProgressForStartPart = 10;
            //一共要导出的元素个数
            int intElementsCount = selectedDrawings.Count();
            //每一个导出的元素所占的进度
            float sngUnit = (float)((double)(100 - intProgressForStartPart) / intElementsCount);
            //实时的进度值
            int intProgress = intProgressForStartPart;
            try
            {
                //写入标题项
                Export_OverView(ref rg);
            }
            catch (Exception)
            {
                MessageBox.Show("写入概述部分出错，但可以继续工作。", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally
            {
                APPLICATION_MAINFORM.MainForm.ShowProgressBar_Continue(intProgressForStartPart);
            }

            // ------------- 取消绘图窗口的禁用 ------------------
            //一定要在将绘图窗口中的图形导出到Word之前取消窗口的禁用，
            //否则的话，当调用这些窗口的Application属性时，就会出现报错：应用程序正在使用中。
            foreach (IntPtr H in WindowHandles)
            {
                APIWindows.EnableWindow(H, true);
            }

            //输出每一个选定的图形
            // ------------- 开挖平面图 ------------------
            try
            {
                ClsDrawing_PlanView D = selectedDrawings.PlanView;
                if (D != null)
                {
                    Page page = D.Page;
                    //
                    NewLine(rg, ParagraphStyle.Title_2);
                    rg.InsertAfter("开挖平面图：");
                    //
                    Export_VisioPlanview(page, ref rg);
                    //
                    intProgress += (int)sngUnit;
                    APPLICATION_MAINFORM.MainForm.ShowProgressBar_Continue(intProgress);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "导出Visio开挖平面图出错，但可以继续工作。" + "\r\n" + ex.Message + "\r\n" + "报错位置：" + ex.TargetSite.Name, "Warning",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            // ------------- 剖面标高图 -------------------------
            try
            {
                ClsDrawing_ExcavationElevation D = selectedDrawings.SectionalView;
                if (D != null)
                {
                    //
                    NewLine(rg, ParagraphStyle.Title_2);
                    rg.InsertAfter("开挖剖面图：");
                    //
                    Export_ExcelChart(D.Chart, ref rg);
                    //
                    intProgress += (int)sngUnit;
                    APPLICATION_MAINFORM.MainForm.ShowProgressBar_Continue(intProgress);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "导出Excel开挖剖面图出错，但可以继续工作。" + "\r\n" + ex.Message + "\r\n" + "报错位置：" + ex.TargetSite.Name, "Warning",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            // ---------------------- 监测曲线图 --------------------
            Chart cht = default(Chart);
            foreach (ClsDrawing_Mnt_Base Drawing in selectedDrawings.MntDrawings)
            {
                try
                {
                    switch (Drawing.Type)
                    {
                        // ------------- 测斜曲线图 ---------------------------------------------------
                        case DrawingType.Monitor_Incline_Dynamic:
                            ClsDrawing_Mnt_Incline D_1 = (ClsDrawing_Mnt_Incline)Drawing;
                            cht = D_1.Chart;
                            //
                            NewLine(rg, ParagraphStyle.Title_2);
                            rg.InsertAfter(D_1.Chart_App_Title);
                            //
                            Export_ExcelChart(cht, ref rg);
                            break;

                        // ------------- 动态监测曲线图 ---------------------------------------------
                        case DrawingType.Monitor_Dynamic:
                            ClsDrawing_Mnt_OtherDynamics D_2 = (ClsDrawing_Mnt_OtherDynamics)Drawing;
                            cht = D_2.Chart;

                            //
                            NewLine(rg, ParagraphStyle.Title_2);
                            rg.InsertAfter(D_2.Chart_App_Title);

                            Export_ExcelChart(cht, ref rg);
                            break;

                        // ------------- 静态监测曲线图 ---------------------------------------------
                        case DrawingType.Monitor_Static:
                            ClsDrawing_Mnt_Static D_3 = (ClsDrawing_Mnt_Static)Drawing;
                            cht = D_3.Chart;
                            //
                            NewLine(rg, ParagraphStyle.Title_2);
                            rg.InsertAfter(D_3.Chart_App_Title);

                            Export_ExcelChart(cht, ref rg);
                            break;
                        // ------------- 静态监测曲线图 ---------------------------------------------
                        case DrawingType.Monitor_Incline_MaxMinDepth:
                            ClsDrawing_Mnt_MaxMinDepth D = (ClsDrawing_Mnt_MaxMinDepth)Drawing;
                            cht = D.Chart;
                            //
                            NewLine(rg, ParagraphStyle.Title_2);
                            rg.InsertAfter(D.Chart_App_Title);

                            Export_ExcelChart(cht, ref rg);
                            break;
                        default:
                            break;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("导出监测曲线图\"" + Drawing.Chart_App_Title.ToString() + "\"出错，但可以继续工作。" +
                                    "\r\n" + ex.Message + "\r\n" + "报错位置：" +
                                    ex.TargetSite.Name, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                finally
                {
                    intProgress += (int)sngUnit;
                    APPLICATION_MAINFORM.MainForm.ShowProgressBar_Continue(intProgress);
                }
            }
        }

        /// <summary>
        /// 输出开头的一些概况信息
        /// </summary>
        /// <param name="Range"></param>
        /// <remarks>包括标题、施工日期，施工工况等</remarks>
        private void Export_OverView(ref Word.Range Range)
        {
            Word.Range with_1 = Range;
            if (with_1.End > 1) //说明当前range不是在文档的开头，那么就要新起一行
            {
                NewLine(Range, ParagraphStyle.Title_1);
            }
            else //说明当前range就是在文档的开头，那么就直接设置段落样式为“标题”就可以了。
            {
                Range.ParagraphFormat.set_Style(ParagraphStyle.Title_1);
            }
            with_1.InsertAfter(Settings.Default.ProjectName + " 实测数据动态分析：" + dateThisday.ToShortDateString());
            //
            NewLine(Range, ParagraphStyle.Content);
            with_1.InsertAfter("施工日期： " + (dateThisday.ToLongDateString() + '\r' + "施工工况： （***）"));
        }

        public void BackgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
        }

        public void BackgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            Close();
            //
            APPLICATION_MAINFORM.MainForm.HideProgress("图形导出完成！");
        }

        #endregion

        #region   ---  图表输出操作

        /// <summary>
        /// 导出excel中的chart对象到word中
        /// </summary>
        /// <param name="cht">excel中的chart对象</param>
        /// <param name="range">此时word文档中的全局range的位置或者范围</param>
        /// <remarks>由局部安全的原则，在进行绘图前，将另起一行，并将段落样式设置为“图片”样式</remarks>
        private void Export_ExcelChart(Chart cht, ref Word.Range range)
        {
            cht.Application.ScreenUpdating = false;
            try
            {
                // 下面复制Chart的操作中，如果监测曲线图所使用的Chart模板有问题，则可能会出错。
                Excel.ChartObject chtObj = cht.Parent as Excel.ChartObject;
                chtObj.Activate();
                chtObj.Copy(); // 或者用  cht.ChartArea.Copy()都可以。
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "导出监测曲线图\"" + cht.Application.Caption.ToString() + "\"出错（请检查是否是用户使用的Chart模板有问题），跳过此图的导出。" +
                    "\r\n" + ex.Message + "\r\n" + "报错位置：" +
                    ex.TargetSite.Name, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                //刷新excel屏幕
                cht.Application.ScreenUpdating = true;
                return;
            }
            //设置word.range的格式
            Word.Range with_1 = range;
            //新起一行，并设置新段落的段落样式为图片
            NewLine(range, ParagraphStyle.picture);

            //进行粘贴，下面也可以用：DataType:=23
            with_1.PasteSpecial(DataType: Word.WdPasteDataType.wdPasteOLEObject, Placement: Word.WdOLEPlacement.wdInLine);


            Word.InlineShape shp = default(Word.InlineShape);
            range.Select();
            range.Application.Selection.MoveLeft(Unit: Word.WdUnits.wdCharacter, Count: 1,
                Extend: Word.WdMovementType.wdExtend);
            shp = range.Application.Selection.InlineShapes[1];
            //约束图形的宽度，将其限制在word页面的正文宽度之内
            WidthRestrain(shp, ContentWidth);

            //刷新excel屏幕
            cht.Application.ScreenUpdating = true;
        }

        /// <summary>
        /// 未使用。导出excel的工作表中的所有形状到word中
        /// </summary>
        /// <param name="sheet">要进行粘贴的Excel中的工作表对象</param>
        /// <param name="range">此时word文档中的全局range的位置或者范围</param>
        /// <remarks>在此方法中，会将Excel的工作表中的所有形状进行选择，然后进行组合，最后将其输出到word中；
        /// 由局部安全的原则，在进行绘图前，将另起一行，并将段落样式设置为“图片”样式</remarks>
        private void Export_ExcelSheet(Excel.Worksheet sheet, ref Word.Range range)
        {
            Excel.Worksheet sht = sheet;
            sht.Application.ScreenUpdating = false;

            //在excel中将工作表里的所有形状进行复制并组合
            sht.Shapes.SelectAll();
            Excel.ShapeRange shprg = sht.Application.Selection as Excel.ShapeRange;
            if (shprg.Type == MsoShapeType.msoGroup)
            {
                Excel.Shape excelShp = default(Excel.Shape);
                excelShp = shprg.Item(0); //ShapeRange中的第一个形状的下标值为0
                excelShp.Copy();
            }
            else
            {
                shprg.Group().Copy();
            }


            Word.Range with_2 = range;
            //新起一行，并设置新段落的段落样式为图片
            NewLine(range, ParagraphStyle.picture);

            // ----------------------------  进行粘贴。这里的DataType不能指定为wdPasteOLEObject，
            //因为从Excel中复制过来的图片，它不是一个OLE对象。
            with_2.PasteSpecial(DataType: 23, Placement: Word. WdOLEPlacement.wdInLine);

            Microsoft.Office.Interop.Word.Shape shp = default(Microsoft.Office.Interop.Word.Shape);
            // 获取刚刚粘贴过来的图片，此时图片很可能不是以嵌入式粘贴进来的。
            shp = with_2.ShapeRange[1];
            //约束图形的宽度
            WidthRestrain(shp, ContentWidth);
            //将shape转换为inlineshape
            shp.ConvertToInlineShape();
            //刷新excel屏幕
            sheet.Application.ScreenUpdating = true;
        }

        /// <summary>
        /// 导出visio的Page中的所有形状到word中
        /// </summary>
        /// <param name="Page">要进行粘贴的Visio中的页面</param>
        /// <param name="range">此时word文档中的全局range的位置或者范围</param>
        /// <remarks>在此方法中，会将visio指定页面中的所有形状进行选择，然后进行组合，最后将其输出到word中；
        /// 由局部安全的原则，在进行绘图前，将另起一行，并将段落样式设置为“图片”样式</remarks>
        private void Export_VisioPlanview(Page Page, ref Word. Range range)
        {
            Microsoft.Office.Interop.Visio.Application app = Page.Application;
            //
            Window wnd = default(Window);
            wnd = app.ActiveWindow;
            wnd.Page = Page;


            wnd.Activate();
            //这里要将ShowChanges设置为True，否则下面的SelectAll()方法会被禁止。
            app.ShowChanges = true;
            wnd.SelectAll();
            //  ---------------------- 耗时代码1：复制Visio的Page中的所有形状
            //而且在这一步的时候Visio的窗口中可能会弹出子窗口
            wnd.Selection.Copy();
            //这一步也可能会导致Visio的窗口中弹出子窗口
            wnd.DeselectAll();

            //关闭所有的子窗口
            //Debug.Print(app.ActiveWindow.Windows.Count)     ‘即使只显示出一个子窗口，这里也会返回10
            //For Each subWnd As Visio.Window In app.ActiveWindow.Windows
            //    subWnd.Visible = False
            //Next

            //根据实际情况：每次都只弹出“外部数据”这一子窗口，所以在这里就只对其进行单独隐藏。
            wnd.Windows.ItemFromID((Int32)Visio.VisWinTypes.visWinIDExternalData).Visible = false;

            //让窗口的显示适应页面
            wnd.ViewFit = (Int32)Visio.VisWindowFit.visFitPage;
            Word.Range with_2 = range;
            //新起一行，并设置新段落的段落样式为图片
            NewLine(range, ParagraphStyle.picture);

            //  ---------------------- 耗时代码2：将Visio的Page中的所有形状粘贴到Word中（不可以用：DataType:=23）
            with_2.PasteSpecial(DataType:  Word.WdPasteDataType.wdPasteOLEObject, Placement: Word. WdOLEPlacement.wdInLine);


            Word.InlineShape shp = default(Word.InlineShape);
            range.Select();
            range.Application.Selection.MoveLeft(Unit: Word. WdUnits.wdCharacter, Count: 1,
                Extend: Word. WdMovementType.wdExtend);
            shp = range.Application.Selection.InlineShapes[1];
            //约束图形的宽度，将其限制在word页面的正文宽度之内
            WidthRestrain(shp, ContentWidth);

            //刷新visio屏幕
            app.ShowChanges = true;
        }

        #endregion

        #region   ---  零碎方法

        /// <summary>
        /// 新起一段
        /// </summary>
        /// <param name="range">Range对象，可以为选区范围或者光标插入点</param>
        /// <param name="PrphStyle">新起一段的段落样式</param>
        /// <remarks></remarks>
        private void NewLine(Word.Range range, string PrphStyle)
        {
            Word.Range with_1 = range;
            with_1.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            with_1.InsertParagraphAfter();
            with_1.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            with_1.ParagraphFormat.set_Style(PrphStyle);
        }

        /// <summary>
        /// 限制粘贴到word中的图形的宽度
        /// </summary>
        /// <param name="shape">粘贴过来的图形，可能为shape对象或者inlineshape对象</param>
        /// <param name="PageWidth_Content">用来限制图片宽度的值，一般取页面的正文版面的宽度值</param>
        /// <remarks></remarks>
        private void WidthRestrain(object shape, float PageWidth_Content)
        {
            dynamic with_1 = shape;
            float W = Convert.ToSingle(with_1.width);
            if (W > PageWidth_Content)
            {
                float H = Convert.ToSingle(with_1.height);
                double AspectRatio = H / W;
                //
                with_1.width = PageWidth_Content;
                with_1.height = with_1.width * AspectRatio;
            }
        }

        /// <summary>
        /// 获取word页面的正文范围的宽度，用来限定图片的宽度值
        /// </summary>
        /// <param name="doc"></param>
        /// <returns>word页面的正文范围的宽度，以磅为单位</returns>
        /// <remarks></remarks>
        private float GetContentWidth(Document doc)
        {
            //正文的区域的宽度CW
            float CW = 0;
            PageSetup ps = doc.PageSetup;
            float W = ps.PageWidth;
            float Margin = ps.LeftMargin + ps.RightMargin;
            CW = W - Margin;
            return CW;
        }

        #endregion

        #region   ---  一般界面操作

        /// <summary>
        /// 对列表中的项目进行全选或者取消全部选择
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <remarks></remarks>
        public void ChkBxSelect_Click(object sender, EventArgs e)
        {
            switch (ChkBxSelect.CheckState)
            {
                case CheckState.Checked:
                    //执行全选操作
                    if (CheckBox_PlanView.Enabled)
                    {
                        CheckBox_PlanView.Checked = true;
                    }
                    if (CheckBox_SectionalView.Enabled)
                    {
                        CheckBox_SectionalView.Checked = true;
                    }
                    foreach (ListBox lstbox in F_arrListBoxes)
                    {
                        short n = (short)lstbox.Items.Count;
                        for (short index = 0; index <= n - 1; index++)
                        {
                            lstbox.SetSelected(index, true);
                        }
                    }
                    break;

                case CheckState.Unchecked:
                case CheckState.Indeterminate:
                    //在UI上跳过中间状态，直接进入“取消全部选择”，并执行取消选择的操作。
                    if (ChkBxSelect.CheckState == CheckState.Indeterminate)
                    {
                        ChkBxSelect.CheckState = CheckState.Unchecked;
                    }
                    //执行取消全部选择操作
                    CheckBox_PlanView.Checked = false;
                    CheckBox_SectionalView.Checked = false;
                    foreach (ListBox lstbox in F_arrListBoxes)
                    {
                        lstbox.ClearSelected();
                    }
                    break;
            }
        }

        //选择复选框或者列表项——更新选择的图表对象
        public void CheckBox_PlanView_CheckedChanged(object sender, EventArgs e)
        {
            if (CheckBox_PlanView.Checked)
            {
                F_SelectedDrawings.PlanView = (ClsDrawing_PlanView)CheckBox_PlanView.Tag;
            }
            else
            {
                F_SelectedDrawings.PlanView = null;
            }
            //
            SelectedDrawingsChanged(F_SelectedDrawings);
        }

        public void CheckBox_SectionalView_CheckedChanged(object sender, EventArgs e)
        {
            if (CheckBox_SectionalView.Checked)
            {
                F_SelectedDrawings.SectionalView = (ClsDrawing_ExcavationElevation)CheckBox_SectionalView.Tag;
            }
            else
            {
                F_SelectedDrawings.SectionalView = null;
            }
            //
            SelectedDrawingsChanged(F_SelectedDrawings);
        }

        public void ListBoxMonitor_Dynamic_SelectedIndexChanged(object sender, EventArgs e)
        {
            F_SelectedDrawings.MntDrawings.Clear();
            //
            ClsDrawing_Mnt_Base Drawing = default(ClsDrawing_Mnt_Base);
            var Items = ListBoxMonitor_Dynamic.SelectedItems;
            foreach (LstbxDisplayAndItem lstboxItem in Items)
            {
                Drawing = (ClsDrawing_Mnt_Base)lstboxItem.Value;
                F_SelectedDrawings.MntDrawings.Add(Drawing);
            }
            Items = ListBoxMonitor_Static.SelectedItems;
            foreach (LstbxDisplayAndItem lstboxItem in Items)
            {
                Drawing = (ClsDrawing_Mnt_Base)lstboxItem.Value;
                F_SelectedDrawings.MntDrawings.Add(Drawing);
            }
            //
            SelectedDrawingsChanged(F_SelectedDrawings);
        }

        //选择复选框或者列表项——更新滚动线程与窗口界面
        /// <summary>
        /// ！选择的图形发生改变时，更新滚动线程与窗口界面。
        /// </summary>
        /// <param name="Selected_Drawings">更新后的要进行滚动的图形</param>
        /// <remarks>此方法不能直接Handle复选框的CheckedChanged或者列表框的SelectedIndexChanged事件，
        /// 因为此方法必须是在更新了Me.F_SelectedDrawings属性之后，才能去更新窗口界面。</remarks>
        private void SelectedDrawingsChanged(Drawings_For_Output Selected_Drawings)
        {
            if (Selected_Drawings.Count() > 0)
            {
                btnExport.Enabled = true;
            }
            else
            {
                btnExport.Enabled = false;
            }
        }

        //
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

        #endregion
    }
}