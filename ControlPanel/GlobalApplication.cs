using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace CableStayedBridge
{
    public class GlobalApplication // GlobalApplication
    {
        #region   ---  属性值的定义

        /// <summary> 整个程序中用来放置各种隐藏的Excel数据文档的Application对象 </summary>
        private Excel.Application m_ExcelApplication_DB;

        /// <summary> 读取或设置整个程序中用来放置各种隐藏的Excel数据文档的Application对象 </summary>
        /// <returns>一个Excel.Application对象，用来装载整个程序中的所有隐藏的后台数据的Excel文档</returns>
        /// <remarks></remarks>
        public Excel.Application ExcelApplication_DB
        {
            get
            {
                //如果此时还没有打开装载Excel数据库工作簿的Excel程序，则先创建一个Excel程序
                if (m_ExcelApplication_DB == null)
                {
                    m_ExcelApplication_DB = new Excel.Application();
                    m_ExcelApplication_DB.Visible = false;
                }
                //
                return m_ExcelApplication_DB;
            }
            set { m_ExcelApplication_DB = value; }
        }
        
        /// <summary> 整个程序当前的操作状态 </summary>
        private ApplicationState m_applicationState;
        public ApplicationState ApplicationState { get; }
        /// <summary> 设置整个程序当前的操作状态 </summary>
        public void SetApplicationState(ApplicationState appState) { m_applicationState = appState; }
        
        #endregion

        #region   ---  单例构造

        private static readonly object locker = new object();
        private static GlobalApplication _instance = null;
        /// <summary>
        /// 用来索引用来保存全局数据的类实例
        /// </summary>
        /// <returns></returns>
        public static GlobalApplication UniqueInstance
        {
            get
            {
                //只有当第一次创建实例的时候再执行lock语句
                if (_instance == null)
                {
                    //当一个线程执行的时候，会先检测locker对象是否是锁定状态，
                    //如果不是则会对locker对象锁定，如果是则该线程就会挂起等待
                    //locker对象解锁,lock语句执行运行完之后会对该locker对象解锁
                    lock (locker)
                    {
                        if (_instance == null)
                        {
                            _instance = new GlobalApplication();
                        }
                    }
                }
                return _instance;
            }
        }

        private GlobalApplication()
        {
        }

        #endregion

    }
}