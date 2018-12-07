using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CableStayedBridge.ControlPanel
{
    class LogInfoHandler
    {

        #region   ---  单例构造

        private static readonly object locker = new object();
        private static LogInfoHandler _instance = null;
        /// <summary>
        /// 用来索引用来保存全局数据的类实例
        /// </summary>
        /// <returns></returns>
        public static LogInfoHandler UniqueInstance
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
                            _instance = new LogInfoHandler();
                        }
                    }
                }
                return _instance;
            }
        }

        private LogInfoHandler()
        {
        }

        #endregion
    }
}
