using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PTool
{

    public class StartOrStopArgs : EventArgs
    {
        private bool m_IsStart = false;

        public bool IsStart
        {
            get { return m_IsStart; }
            set { m_IsStart = value; }
        }

        public StartOrStopArgs(bool isStart)
        {
            m_IsStart = isStart;
        }
    }

    /// <summary>
    /// 当测量结束后，将数据发给主界面，由主界面统一调度
    /// </summary>
    public class DoublePumpDataArgs : EventArgs
    {
        private List<SampleData> m_SampleDataList = new List<SampleData>();

        public List<SampleData> SampleDataList
        {
            get { return m_SampleDataList; }
            set { m_SampleDataList = value; }
        }

        public DoublePumpDataArgs(List<SampleData> dataList)
        {
            foreach (var data in dataList)
            {
                SampleData obj = new SampleData();
                obj.Copy(data);
                m_SampleDataList.Add(obj);
            }
        }
    }




}
 
