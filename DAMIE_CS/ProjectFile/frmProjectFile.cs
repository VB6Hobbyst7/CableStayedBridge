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
using CableStayedBridge.GlobalApp_Form;
using CableStayedBridge.Miscellaneous;
// End of VB project level imports

using Microsoft.Office.Interop.Excel;
using CableStayedBridge.Constants;


namespace CableStayedBridge
{
    namespace DataBase
    {

        /// <summary>
        /// 对项目文件进行操作的窗口，比如新建项目、打开项目、编辑项目等
        /// </summary>
        /// <remarks></remarks>
        public partial class frmProjectFile
        {
            public frmProjectFile()
            {
                InitializeComponent();
            }

            #region   ---  定义与声明

            #region   ---  Events

            /// <summary>
            /// 工作簿列表框中的工作簿对象发生变化时触发的事件
            /// </summary>
            /// <param name="Sender"></param>
            /// <param name="FileContents"></param>
            /// <param name="clearListBox_Progress">是否要清空项目文件中的施工进度列表框中的内容</param>
            /// <remarks></remarks>
            private delegate void WorkBookInProjectChangedEventHandler(object Sender, clsData_FileContents FileContents, bool clearListBox_Progress);
            private WorkBookInProjectChangedEventHandler WorkBookInProjectChangedEvent;

            private event WorkBookInProjectChangedEventHandler WorkBookInProjectChanged
            {
                add
                {
                    WorkBookInProjectChangedEvent = (WorkBookInProjectChangedEventHandler)System.Delegate.Combine(WorkBookInProjectChangedEvent, value);
                }
                remove
                {
                    WorkBookInProjectChangedEvent = (WorkBookInProjectChangedEventHandler)System.Delegate.Remove(WorkBookInProjectChangedEvent, value);
                }
            }


            #endregion

            #region   ---  Constants

            /// <summary>
            /// 列表框的DataSource中用来表示Value的属性名
            /// </summary>
            /// <remarks></remarks>
            private const string cstValueMember = LstbxDisplayAndItem.ValueMember;
            /// <summary>
            /// 列表框的DataSource中用来显示在UI界面中的DisplayMember的属性名
            /// </summary>
            /// <remarks></remarks>
            private const string cstDisplayMember = LstbxDisplayAndItem.DisplayMember;

            #endregion

            #region   ---  Properties

            /// <summary>
            /// 窗口打开时的作用状态
            /// </summary>
            /// <remarks></remarks>
            private ProjectState P_ProjectState;
            /// <summary>
            /// 窗口打开时的作用状态
            /// </summary>
            /// <value></value>
            /// <returns></returns>
            /// <remarks></remarks>
            public ProjectState ProjectState
            {
                get
                {
                    return this.P_ProjectState;
                }
                set
                {
                    this.P_ProjectState = value;
                }
            }

            #endregion

            #region   ---  Fields
            private GlobalApplication GlbApp; // VBConversions Note: Initial value cannot be assigned here since it is non-static.  Assignment has been moved to the class constructors.

            /// <summary>
            /// 窗口中的界面所反映出的项目文件内容。
            /// 此变量是为了在窗口中点击确认的时候赋值给主程序，
            /// 而如果窗口不点击确定的话，主程序的变量就不会被更新。
            /// </summary>
            /// <remarks></remarks>
            private clsData_FileContents F_NewFileContents;
            /// <summary>
            /// 在打开此窗口时，程序中的数据库文件对象，此对象在此窗口中是只读的。
            /// </summary>
            /// <remarks></remarks>
            private clsData_FileContents F_OldFileContents = new clsData_FileContents();
            #endregion

            #endregion

            #region   ---  构造函数与窗体的加载、打开与关闭

            /// <summary>
            /// 窗口加载
            /// </summary>
            public void frmProjectContents_Load(object sender, EventArgs e)
            {
                //这一步的New非常重要，因为在每一次编辑完成后都会将其传址给主程序。
                F_NewFileContents = new clsData_FileContents();
                //
                APPLICATION_MAINFORM.MainForm.AllowDrop = false;
                clsProjectFile GlobalProjectfile = GlbApp.ProjectFile;
                if (GlobalProjectfile != null)
                {
                    this.F_OldFileContents = GlobalProjectfile.Contents;
                    FileContentsToUI(this.F_OldFileContents);
                    //在Label中更新此项目文件的绝对路径
                    this.LabelProjectFilePath.Text = GlobalProjectfile.FilePath;

                }

                //
                this.BindProperty();

                //根据不同的窗口状态设置窗口的样式
                switch (this.P_ProjectState)
                {
                    case Miscellaneous.ProjectState.NewProject:
                        frmProjectFile with_1 = this;
                        with_1.Text = "New Project";
                        break;

                    case Miscellaneous.ProjectState.EditProject:
                        //将项目文件中的内容更新的窗口中
                        frmProjectFile with_2 = this;
                        with_2.Text = "Edit Project";
                        break;
                }

            }
            /// <summary>
            /// 为列表框绑定文本显示与数据的属性值
            /// </summary>
            /// <remarks></remarks>
            private void BindProperty()
            {
                //设置列表框的显示文本的属性
                List<ListControl> AllListControl = new List<ListControl>();
                AllListControl.Clear();
                AllListControl.Add(this.CmbbxPlan);
                AllListControl.Add(this.CmbbxPointCoordinates);
                AllListControl.Add(this.CmbbxProgressWkbk);
                AllListControl.Add(this.CmbbxSectional);
                AllListControl.Add(this.CmbbxWorkingStage);
                AllListControl.Add(this.LstbxSheetsProgressInProject);
                AllListControl.Add(this.LstbxSheetsProgressInWkbk);
                AllListControl.Add(this.LstBxWorkbooks);
                foreach (ListControl lstControl in AllListControl)
                {
                    lstControl.DisplayMember = cstDisplayMember;
                    lstControl.ValueMember = cstValueMember;
                }
            }

            public void frmProjectFile_FormClosed(object sender, FormClosedEventArgs e)
            {
                APPLICATION_MAINFORM.MainForm.AllowDrop = true;
            }
            #endregion

            //项目文件刷新到窗口UI
            /// <summary>
            /// 在打开窗口时将主程序中的ProjectFile对象的信息反映到窗口的控件中
            /// </summary>
            /// <param name="FileContents"></param>
            /// <remarks></remarks>
            public void FileContentsToUI(clsData_FileContents FileContents)
            {
                if (FileContents != null)
                {
                    clsData_FileContents with_1 = FileContents;

                    //显示出项目中的工作簿
                    List<LstbxDisplayAndItem> listWkbks = new List<LstbxDisplayAndItem>();
                    foreach (Workbook wkbk in with_1.lstWkbks)
                    {
                        listWkbks.Add(new LstbxDisplayAndItem(wkbk.FullName, wkbk));
                    }
                    this.LstBxWorkbooks.DataSource = listWkbks;

                    //显示出项目文件中的施工进度工作表
                    List<LstbxDisplayAndItem> DataSource_SheetsProgressInProject = new List<LstbxDisplayAndItem>();
                    foreach (Worksheet shtProgress in with_1.lstSheets_Progress)
                    {
                        DataSource_SheetsProgressInProject.Add(new LstbxDisplayAndItem(DisplayedText: shtProgress.Name, Value:
                            shtProgress));
                    }
                    this.LstbxSheetsProgressInProject.DataSource = DataSource_SheetsProgressInProject;
                    //将数据源更新到所有的列表框
                    if (WorkBookInProjectChangedEvent != null)
                        WorkBookInProjectChangedEvent(null, FileContents, false);
                }
            }

            #region   ---  !项目内容与窗口显示的交互

            //项目中的数据库工作簿发生变化时的事件——窗口控件中列表框的刷新
            /// <summary>
            /// 项目中的数据库工作簿发生变化时的事件
            /// </summary>
            /// <param name="sender"></param>
            /// <param name="clearListBox_Progress">是否要清空项目文件中的施工进度列表框中的内容</param>
            /// <remarks></remarks>
            public void _WorkBookInProjectChanged(object Sender, clsData_FileContents FileContents,
                bool clearListBox_Progress)
            {

                //' 整个项目文件中所有的工作簿中的所有工作表的集合，用来作为列表控件中DataSource。
                List<LstbxDisplayAndItem> lstAllWorksheets = new List<LstbxDisplayAndItem>();
                //下面这一项“无”很重要，因为当数据库中对于某一个项目（如开挖剖面，测点坐标），
                //如果没有相应的数据，就应该选择“无”。
                lstAllWorksheets.Add(new LstbxDisplayAndItem(" 无", LstbxDisplayAndItem.NothingInListBox.None));

                //
                this.CmbbxProgressWkbk.Items.Clear();
                //
                foreach (LstbxDisplayAndItem lstItem in (this.LstBxWorkbooks.DataSource as List<LstbxDisplayAndItem>))
                {
                    Workbook wkbk = (Workbook)lstItem.Value;

                    //更新项目中的工作簿的组合框列表
                    this.CmbbxProgressWkbk.Items.Add(new LstbxDisplayAndItem(wkbk.Name, wkbk));

                    //提取工作簿中的所有工作表
                    foreach (Worksheet sht in wkbk.Worksheets)
                    {
                        lstAllWorksheets.Add(new LstbxDisplayAndItem(DisplayedText: wkbk.Name + " : " + sht.Name, Value:
                            sht));
                    }
                }

                //更新： 将数据源更新到所有的列表框
                ListControlRefresh(lstAllWorksheets, FileContents);

                if (clearListBox_Progress)
                {
                    // 项目文件中的施工进度工作表的列表框的刷新
                    //这里应该判断项目文件中保存的施工进度工作表中，有哪些是位于被删除的那个工作簿的，然后将属于那个工作簿的工作表进行移除。
                    //现在为了简单起见，先直接将其清空。
                    this.LstbxSheetsProgressInProject.DataSource = null;
                    //.Items.Clear()
                    //.DisplayMember = LstbxDisplayAndItem.DisplayMember
                }

            }
            /// <summary>
            /// 项目中的数据库工作簿发生变化时，更新窗口中的相关列表框中的数据对象
            /// </summary>
            /// <param name="lstAllsheet"></param>
            /// <param name="FileContents"></param>
            /// <remarks></remarks>
            private void ListControlRefresh(List<LstbxDisplayAndItem> lstAllsheet,
                clsData_FileContents FileContents)
            {
                short intSheetsCount = System.Convert.ToInt16(lstAllsheet.Count);
                LstbxDisplayAndItem[] arrAllSheets1 = new LstbxDisplayAndItem[intSheetsCount - 1 + 1];
                LstbxDisplayAndItem[] arrAllSheets2 = new LstbxDisplayAndItem[intSheetsCount - 1 + 1];
                LstbxDisplayAndItem[] arrAllSheets3 = new LstbxDisplayAndItem[intSheetsCount - 1 + 1];
                LstbxDisplayAndItem[] arrAllSheets4 = new LstbxDisplayAndItem[intSheetsCount - 1 + 1];
                //这里一定要生成副本，因为如果是同一个引用变量，那么设置到三个控件的DataSource属性中后，
                //如果一个列表组合框的选择项发生变化,那个另外两个控件的选择项也会同步变化。
                arrAllSheets1 = lstAllsheet.ToArray() as LstbxDisplayAndItem[];
                arrAllSheets2 = arrAllSheets1.Clone() as LstbxDisplayAndItem[];
                arrAllSheets3 = arrAllSheets1.Clone() as LstbxDisplayAndItem[];
                arrAllSheets4 = arrAllSheets1.Clone() as LstbxDisplayAndItem[];

                //设置各种列表框中的数据以及选择的项lstbxWorksheetsInProjectFileChanged
                frmProjectFile with_1 = this;

                //开挖平面工作表
                CmbbxPlan.DataSource = arrAllSheets1;
                CmbbxPlan.ValueMember = cstValueMember;
                if (FileContents.Sheet_PlanView != null)
                {
                    this.SheetToComboBox(CmbbxPlan, FileContents.Sheet_PlanView);
                }
                else
                {
                    CmbbxPlan.SelectedValue = LstbxDisplayAndItem.NothingInListBox.None;
                }

                //测点坐标工作表
                CmbbxPointCoordinates.DataSource = arrAllSheets2;
                CmbbxPointCoordinates.ValueMember = cstValueMember;
                if (FileContents.Sheet_PointCoordinates != null)
                {
                    this.SheetToComboBox(CmbbxPointCoordinates, FileContents.Sheet_PointCoordinates);
                }
                else
                {
                    CmbbxPointCoordinates.SelectedValue = LstbxDisplayAndItem.NothingInListBox.None;
                }

                //开挖剖面工作表
                CmbbxSectional.DataSource = arrAllSheets3;
                CmbbxSectional.ValueMember = cstValueMember;
                if (FileContents.Sheet_Elevation != null)
                {
                    this.SheetToComboBox(CmbbxSectional, FileContents.Sheet_Elevation);
                }
                else
                {
                    CmbbxSectional.SelectedValue = LstbxDisplayAndItem.NothingInListBox.None;
                }

                //开挖工况工作表
                CmbbxWorkingStage.DataSource = arrAllSheets4;
                CmbbxWorkingStage.ValueMember = cstValueMember;
                if (FileContents.Sheet_WorkingStage != null)
                {
                    this.SheetToComboBox(CmbbxWorkingStage, FileContents.Sheet_WorkingStage);
                }
                else
                {
                    CmbbxWorkingStage.SelectedValue = LstbxDisplayAndItem.NothingInListBox.None;
                }


                //为施工进度列表服务的组合列表框
                if (CmbbxProgressWkbk.Items.Count == 0)
                {
                    //工作簿中的工作表列表框
                    this.LstbxSheetsProgressInWkbk.DataSource = null;
                    //上面将DataSource设置为Nothing会清空DisplayMember属性的值，那么下次再向列表框中添加成员时，
                    //其DisplayMember就为Nothing了，所以必须在下面设置好其DisplayMember的值。
                    this.LstbxSheetsProgressInWkbk.DisplayMember = LstbxDisplayAndItem.DisplayMember;
                }
                else
                {
                    CmbbxProgressWkbk.SelectedIndex = 0;
                }

            }
            /// <summary>
            /// 将项目文件中的工作表添加到组合列表框中
            /// </summary>
            /// <param name="cmbx">进行添加的组合列表框</param>
            /// <param name="destinationSheet">要添加到组合列表框中的工作表对象</param>
            /// <remarks></remarks>
            private void SheetToComboBox(ComboBox cmbx, Worksheet destinationSheet)
            {
                LstbxDisplayAndItem lstbxItem = default(LstbxDisplayAndItem);
                Worksheet sht = default(Worksheet);
                ComboBox with_1 = cmbx;
                foreach (LstbxDisplayAndItem tempLoopVar_lstbxItem in with_1.Items)
                {
                    lstbxItem = tempLoopVar_lstbxItem;
                    if (!lstbxItem.Value.Equals(LstbxDisplayAndItem.NothingInListBox.None))
                    {
                        //有可能会出现列表中的项目不能转换为工作表对象的情况，比如第一项""
                        sht = (Worksheet)lstbxItem.Value;
                        if (ExcelFunction.SheetCompare(sht, destinationSheet))
                        {
                            with_1.SelectedItem = lstbxItem;
                            break;
                        }
                    }
                }
            }

            #endregion

            #region   ---  一般界面操作

            //添加或者移除工作簿
            /// <summary>
            /// 在项目中添加数据库的工作簿
            /// </summary>
            /// <param name="sender"></param>
            /// <param name="e"></param>
            /// <remarks></remarks>
            public void btnAddWorkbook_Click(object sender, EventArgs e)
            {
                string FilePath = "";
                APPLICATION_MAINFORM.MainForm.OpenFileDialog1.Title = "选择Excel数据工作簿";
                APPLICATION_MAINFORM.MainForm.OpenFileDialog1.Filter = "Excel文件(*.xlsx, *.xls, *.xlsb)|*.xlsx;*.xls;*.xlsb";
                APPLICATION_MAINFORM.MainForm.OpenFileDialog1.FilterIndex = 1;
                if (APPLICATION_MAINFORM.MainForm.OpenFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    FilePath = APPLICATION_MAINFORM.MainForm.OpenFileDialog1.FileName;
                }
                if (FilePath.Length > 0)
                {
                    Workbook wkbk = null;
                    //先看选择的工作簿是否已经在数据库的Excel程序中打开
                    bool blnOpenedInApp = false;
                    foreach (Workbook wkbkOpened in GlbApp.ExcelApplication_DB.Workbooks)
                    {
                        if (string.Compare(wkbkOpened.FullName, FilePath) == 0)
                        {
                            wkbk = wkbkOpened;
                            blnOpenedInApp = true;
                            break;
                        }
                    }

                    //
                    if (!blnOpenedInApp) //如果此工作簿还没有在Excel程序中打开
                    {
                        //则将其打开，并添加到列表框中
                        wkbk = GlbApp.ExcelApplication_DB.Workbooks.Open(Filename: FilePath, UpdateLinks:
                            false, ReadOnly: true);
                        //为列表框中添加新元素
                        List<LstbxDisplayAndItem> DataSource = new List<LstbxDisplayAndItem>();
                        if (LstBxWorkbooks.DataSource != null)
                        {
                            foreach (LstbxDisplayAndItem i in (LstBxWorkbooks.DataSource as List<LstbxDisplayAndItem>))
                            {
                                DataSource.Add(i);
                            }
                        }
                        DataSource.Add(new LstbxDisplayAndItem(FilePath, wkbk));
                        LstBxWorkbooks.DataSource = DataSource;
                        //
                        if (WorkBookInProjectChangedEvent != null)
                            WorkBookInProjectChangedEvent(btnAddWorkbook, this.F_NewFileContents, false);

                    }
                    else //说明此工作簿已经在Excel中打开
                    {
                        //则先检查它是否已经添加到了列表框中
                        LstbxDisplayAndItem lstbxItem = default(LstbxDisplayAndItem);
                        bool blnShownInListbox = false;
                        foreach (LstbxDisplayAndItem tempLoopVar_lstbxItem in LstBxWorkbooks.Items)
                        {
                            lstbxItem = tempLoopVar_lstbxItem;
                            Workbook wkbkInproject = (Workbook)lstbxItem.Value;
                            if (string.Compare(wkbkInproject.FullName, FilePath) == 0)
                            {
                                blnShownInListbox = true;
                            }
                        }
                        if (!blnShownInListbox)
                        {
                            //为列表框中添加新元素
                            List<LstbxDisplayAndItem> DataSource = new List<LstbxDisplayAndItem>();
                            var controlSource = LstBxWorkbooks.DataSource as List<LstbxDisplayAndItem>;
                            if (controlSource != null)
                            {
                                foreach (LstbxDisplayAndItem i in controlSource)
                                {
                                    DataSource.Add(i);
                                }
                            }
                            DataSource.Add(new LstbxDisplayAndItem(FilePath, wkbk));
                            LstBxWorkbooks.DataSource = DataSource;
                            //
                            if (WorkBookInProjectChangedEvent != null)
                                WorkBookInProjectChangedEvent(btnAddWorkbook, this.F_NewFileContents, false);

                        }
                    }

                    //再看将工作簿对象添加到列表中
                }
            }
            /// <summary>
            /// 在项目中移除数据库工作簿
            /// </summary>
            /// <param name="sender"></param>
            /// <param name="e"></param>
            /// <remarks>在从列表框中移除项目时，一定要注意，由于集合的数据结构，所以删除时要从后面的开始删除，
            /// 否则可能会出现索引不到的情况。</remarks>
            public void BtnRemoveWorkbook_Click(object sender, EventArgs e)
            {
                byte count = (byte)LstBxWorkbooks.SelectedIndices.Count;
                if (count > 0)
                {

                    List<LstbxDisplayAndItem> DataSource = new List<LstbxDisplayAndItem>();
                    if (LstBxWorkbooks.DataSource != null)
                    {
                        foreach (LstbxDisplayAndItem i in (LstBxWorkbooks.DataSource as List<LstbxDisplayAndItem>))
                        {
                            DataSource.Add(i);
                        }
                    }
                    //
                    Workbook wkbk = default(Workbook);
                    LstbxDisplayAndItem lstbxItem = default(LstbxDisplayAndItem);
                    int index = 0;
                    for (int i = (count - 1); i >= 0; i--)
                    {
                        index = LstBxWorkbooks.SelectedIndices[i];
                        lstbxItem = DataSource[index];
                        //
                        wkbk = (Workbook)lstbxItem.Value;
                        wkbk.Close(false);
                        DataSource.RemoveAt(index);
                    }
                    LstBxWorkbooks.DataSource = DataSource;
                    if (WorkBookInProjectChangedEvent != null)
                        WorkBookInProjectChangedEvent(BtnRemoveWorkbook, this.F_NewFileContents, true);
                }
            }

            //对施工进度表的列表项进行操作
            /// <summary>
            /// 在项目中添加施工进度工作表
            /// </summary>
            /// <param name="sender"></param>
            /// <param name="e"></param>
            /// <remarks></remarks>
            public void BtnAddSheet_Click(object sender, EventArgs e)
            {

                switch (this.P_ProjectState)
                {

                    case Miscellaneous.ProjectState.NewProject: //直接用LstbxDisplayAndItem对象来进行比较

                        //项目文件中的施工进度工作表的列表框中的所有工作表项
                        object OldDataSource = this.LstbxSheetsProgressInProject.DataSource;
                        List<LstbxDisplayAndItem> NewDataSource = new List<LstbxDisplayAndItem>();
                        if (this.LstbxSheetsProgressInProject.DataSource != null)
                        {
                            foreach (LstbxDisplayAndItem i in (IEnumerable)OldDataSource)
                            {
                                NewDataSource.Add(i);
                            }
                        }
                        //
                        foreach (LstbxDisplayAndItem lstItem_Wkbk in this.LstbxSheetsProgressInWkbk.SelectedItems)
                        {
                            //看选择的工作表是否已经包含在项目工作表中
                            if (!NewDataSource.Contains(lstItem_Wkbk))
                            {
                                NewDataSource.Add(lstItem_Wkbk);
                            }
                        }
                        this.LstbxSheetsProgressInProject.DataSource = NewDataSource;
                        break;

                    case Miscellaneous.ProjectState.EditProject: //用工作表的路径来进行比较
                                                                 //工作簿中选择的工作表
                        Worksheet sht_InWkbk = default(Worksheet);
                        //项目文件中已经存在的工作表
                        Worksheet sht_InProject = default(Worksheet);
                        //
                        List<LstbxDisplayAndItem> ItemsAddToProject = new List<LstbxDisplayAndItem>();
                        object OldDataSource_SheetsInProject = this.LstbxSheetsProgressInProject.DataSource;
                        List<LstbxDisplayAndItem> NewDataSource_SheetsInProject = new List<LstbxDisplayAndItem>();
                        //
                        if (OldDataSource_SheetsInProject == null)
                        {
                            foreach (LstbxDisplayAndItem lstbxItem_Wkbk in this.LstbxSheetsProgressInWkbk.SelectedItems)
                            {
                                NewDataSource_SheetsInProject.Add(lstbxItem_Wkbk);
                            }
                            this.LstbxSheetsProgressInProject.DataSource = NewDataSource_SheetsInProject;

                        }
                        else
                        {

                            foreach (LstbxDisplayAndItem lstbxItem_Wkbk in (IEnumerable)OldDataSource_SheetsInProject)
                            {
                                NewDataSource_SheetsInProject.Add(lstbxItem_Wkbk);
                            }
                            //
                            foreach (LstbxDisplayAndItem lstbxItem_Wkbk in this.LstbxSheetsProgressInWkbk.SelectedItems)
                            {
                                //判断两个工作表是否相等
                                bool blnSheetsMatched = false;
                                //
                                sht_InWkbk = (Worksheet)lstbxItem_Wkbk.Value;
                                foreach (LstbxDisplayAndItem lstbxItem_Project in (IEnumerable)OldDataSource_SheetsInProject)
                                {
                                    sht_InProject = (Worksheet)lstbxItem_Project.Value;

                                    if (ExcelFunction.SheetCompare(sht_InProject, sht_InWkbk) == true)
                                    {
                                        blnSheetsMatched = true;
                                        break;
                                    }
                                } //lstbxItem_Project

                                //如果两个工作表不匹配，则添加到项目文件中。
                                if (!blnSheetsMatched)
                                {
                                    NewDataSource_SheetsInProject.Add(lstbxItem_Wkbk);
                                }

                            } //lstbxItem_Wkbk
                            this.LstbxSheetsProgressInProject.DataSource = NewDataSource_SheetsInProject;
                        }
                        break;

                }
            }
            /// <summary>
            /// 在项目中移除施工进度工作表
            /// </summary>
            /// <param name="sender"></param>
            /// <param name="e"></param>
            /// <remarks></remarks>
            public void BtnRemoveSheet_Click(object sender, EventArgs e)
            {
                List<LstbxDisplayAndItem> DataSource = new List<LstbxDisplayAndItem>();
                if (LstbxSheetsProgressInProject.DataSource != null)
                {
                    foreach (LstbxDisplayAndItem i in (LstbxSheetsProgressInProject.DataSource as List<LstbxDisplayAndItem>))
                    {
                        DataSource.Add(i);
                    }
                }
                byte count = (byte)LstbxSheetsProgressInProject.SelectedIndices.Count;
                for (int i = count - 1; i >= 0; i--)
                {
                    byte index = (byte)(LstbxSheetsProgressInProject.SelectedIndices[i]);
                    DataSource.RemoveAt(index);
                }
                LstbxSheetsProgressInProject.DataSource = DataSource;
            }
            /// <summary>
            /// 工作簿的组合列表框的选择项发生变化时，更新施工进度工作表的列表框
            /// </summary>
            /// <param name="sender"></param>
            /// <param name="e"></param>
            /// <remarks></remarks>
            public void CmbbxProgressWkbk_SelectedValueChanged(object sender, EventArgs e)
            {
                Workbook wkbk = default(Workbook);
                LstbxDisplayAndItem lstItem = (LstbxDisplayAndItem)CmbbxProgressWkbk.SelectedItem;
                if (lstItem == null)
                {
                    this.LstbxSheetsProgressInProject.DataSource = null;
                    this.LstbxSheetsProgressInWkbk.DataSource = null;
                }
                else
                {
                    wkbk = (Workbook)lstItem.Value;
                    //'-- 这里的shtsObject只能声明为Object，而不能声明为Worksheets接口，否则会出现异常：
                    //无法将类型为“System.__ComObject”的 COM 对象强制转换为接口类型
                    //“Microsoft.Office.Interop.Excel.Worksheets”。此操作失败的原因是对 IID
                    //为“{000208B1-0000-0000-C000-000000000046}”的接口的 COM 组件调用 QueryInterface
                    //因以下错误而失败: 不支持此接口 (异常来自 HRESULT:0x80004002 (E_NOINTERFACE))。
                    //'-- 但是这里一定要如此调用，而不用dim sht as worksheet=wkbk.Worksheets(i)去一次一次地
                    //调用每一个单个的工作表，是为了避免多次调用Worksheets接口。
                    dynamic shtsObject = wkbk.Worksheets;
                    short shtCount = System.Convert.ToInt16(shtsObject.Count);
                    LstbxDisplayAndItem[] arrSheets = new LstbxDisplayAndItem[shtCount - 1 + 1];
                    for (var i = 0; i <= shtCount - 1; i++)
                    {
                        Worksheet sht = shtsObject[i + 1]; //Worksheets接口的集合中的第一个元素的下标值为1
                        arrSheets[(int)i] = new LstbxDisplayAndItem(sht.Name, sht);
                    }
                    LstbxSheetsProgressInWkbk.DataSource = arrSheets;
                }
            }

            //窗口的最终处理：确定，取消等
            /// <summary>
            /// 将界面中的内容保存到XML文档中
            /// </summary>
            /// <param name="sender"></param>
            /// <param name="e"></param>
            /// <remarks></remarks>
            public void btnOk_Click(object sender, EventArgs e)
            {

                //'提取界面中绑定的数据
                //Me.F_FileContents = Me.UIToFileContents()

                //新开一个线程以将FileContents中的内容更新到程序的DataBase中。
                Thread thd = new Thread(new System.Threading.ThreadStart(this.RefreshDataBase));
                thd.Name = "在项目文件窗口关闭时刷新程序中的数据库";
                thd.Start(this.F_NewFileContents);
                //
                this.Close();
            }
            private void RefreshDataBase(clsData_FileContents FileContents)
            {
                GlbApp.ProjectFile.Contents = FileContents;
                if (this.ProjectState == Miscellaneous.ProjectState.NewProject)
                {
                    GlbApp.ProjectFile.FilePath = null;
                }
                GlbApp.DataBase = new ClsData_DataBase(FileContents);
                //将主程序的界面刷新为打开了文件后的界面
                APPLICATION_MAINFORM.MainForm.MainUI_ProjectOpened();
            }

            /// <summary>
            /// 鼠标移动进Panel时引发的事件。
            /// </summary>
            /// <remarks>此时将Panel设置为获得焦点。</remarks>
            public void PanelFather_MouseEnter(object sender, EventArgs e)
            {
                this.PanelFather.Focus();
            }

            #endregion

            #region   ---  选择列表框内容时进行赋值

            //开挖平面分块工作表
            public void CmbbxPlan_SelectedIndexChanged(object sender, EventArgs e)
            {
                LstbxDisplayAndItem LstbxItem = CmbbxPlan.SelectedItem as LstbxDisplayAndItem;
                if (LstbxItem != null)
                {
                    if (!LstbxItem.Value.Equals(LstbxDisplayAndItem.NothingInListBox.None))
                    {
                        this.F_NewFileContents.Sheet_PlanView = (Worksheet)LstbxItem.Value;
                    }
                    else
                    {
                        this.F_NewFileContents.Sheet_PlanView = null;
                    }
                }
            }
            //监测点位坐标的数据工作表
            public void CmbbxPointCoordinates_SelectedIndexChanged(object sender, EventArgs e)
            {
                LstbxDisplayAndItem LstbxItem = CmbbxPointCoordinates.SelectedItem as LstbxDisplayAndItem;
                if (LstbxItem != null)
                {
                    if (!LstbxItem.Value.Equals(LstbxDisplayAndItem.NothingInListBox.None))
                    {
                        this.F_NewFileContents.Sheet_PointCoordinates = (Worksheet)LstbxItem.Value;
                    }
                    else
                    {
                        this.F_NewFileContents.Sheet_PointCoordinates = null;
                    }
                }
            }
            //开挖剖面标高图的数据工作表
            public void CmbbxSectional_SelectedIndexChanged(object sender, EventArgs e)
            {
                LstbxDisplayAndItem LstbxItem = CmbbxSectional.SelectedItem as LstbxDisplayAndItem;
                if (LstbxItem != null)
                {
                    if (!LstbxItem.Value.Equals(LstbxDisplayAndItem.NothingInListBox.None))
                    {
                        this.F_NewFileContents.Sheet_Elevation = (Worksheet)LstbxItem.Value;
                    }
                    else
                    {
                        this.F_NewFileContents.Sheet_Elevation = null;
                    }
                }
            }
            //开挖工况的数据工作表
            public void CmbbxWorkingStage_SelectedIndexChanged(object sender, EventArgs e)
            {
                LstbxDisplayAndItem LstbxItem = CmbbxWorkingStage.SelectedItem as LstbxDisplayAndItem;
                if (LstbxItem != null)
                {
                    if (!LstbxItem.Value.Equals(LstbxDisplayAndItem.NothingInListBox.None))
                    {
                        this.F_NewFileContents.Sheet_WorkingStage = (Worksheet)LstbxItem.Value;
                    }
                    else
                    {
                        this.F_NewFileContents.Sheet_WorkingStage = null;
                    }
                }
            }
            //施工进度工作表
            public void LstbxSheetsProgressInProject_DataSourceChanged(object sender, EventArgs e)
            {
                List<Worksheet> lstSheetProgress = new List<Worksheet>();
                if (LstbxSheetsProgressInProject.DataSource != null)
                {
                    foreach (LstbxDisplayAndItem LstbxItem in (LstbxSheetsProgressInProject.DataSource as List<LstbxDisplayAndItem>))
                    {
                        Worksheet sht = (Worksheet)LstbxItem.Value;
                        lstSheetProgress.Add(sht);
                    }
                }

                this.F_NewFileContents.lstSheets_Progress = lstSheetProgress;
            }
            //所有代表项目数据库的工作簿文件
            public void LstBxWorkbooks_DataSourceChanged(object sender, EventArgs e)
            {
                //项目中所有的工作簿
                List<Workbook> lstWkbk = new List<Workbook>();
                if (LstBxWorkbooks.DataSource != null)
                {
                    foreach (LstbxDisplayAndItem LstbxItem in (LstBxWorkbooks.DataSource as List<LstbxDisplayAndItem>))
                    {
                        lstWkbk.Add((Workbook)LstbxItem.Value);
                    }
                }
                this.F_NewFileContents.lstWkbks = lstWkbk;
            }

            #endregion

        }
    }

}
