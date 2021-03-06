﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Collections.Specialized;
using System.ComponentModel;
using System.IO;
using System.Data;
using System.Drawing;
using System.Threading;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using System.IO.Ports;
using ClosedXML.Excel;
using CCWin;
using ApplicationClient;
using Misc = ComunicationProtocol.Misc;
using SerialDevice;

namespace PTool
{
    public partial class PressureForm : Form
    {
        private bool moving = false;
        private Point oldMousePosition;
        private PumpID m_LocalPid = PumpID.GrasebyC6;//默认显示的是C6
        private int m_SampleInterval = 500;//采样频率：毫秒
        //private List<List<SampleData>> m_SampleDataList = new List<List<SampleData>>();//存放双道泵上传的数据，等第二道泵结束后，一起存在一张表中

        private Hashtable hashSampleData = new Hashtable();//[Key=通道编号（0，1）, Value=List<PressureCalibrationParameter>]



        public static int RangeMinP = 170;
        public static int RangeMaxP = 210;
        public static int PressureCalibrationMax = 418;
        public static int SerialNumberCount = 28;               //在指定时间内连续输入字符数量不低于28个时方可认为是由条码枪输入
        public static List<double> SamplingPoints1 = new List<double>();//采样点大概5个，当工装读数在某个值时，自动停止，等待5秒，再读三次工装和P值，比较是否稳定。不稳定，再读
        public static List<double> SamplingPoints2 = new List<double>();//采样点大概5个，当工装读数在某个值时，自动停止，等待5秒，再读三次工装和P值，比较是否稳定。不稳定，再读

        public static List<double> SingleSamplingPoints1 = new List<double>();//采样点大概5个，当工装读数在某个值时，自动停止，等待5秒，再读三次工装和P值，比较是否稳定。不稳定，再读
        public static List<double> SingleSamplingPoints2 = new List<double>();//采样点大概5个，当工装读数在某个值时，自动停止，等待5秒，再读三次工装和P值，比较是否稳定。不稳定，再读

        public static List<double> SamplingPoints = new List<double>();//上面三个数组的总和
        public static double m_StandardError = 0.05; 
        public static double m_SamplingError = 0.1; //采样误差范围


        private const int INPUTSPEED = 50;//条码枪输入字符速率小于50毫秒
        

        private DateTime m_CharInputTimestamp = DateTime.Now;  //定义一个成员函数用于保存每次的时间点
        private DateTime m_FirstCharInputTimestamp = DateTime.Now;  //定义一个成员函数用于保存每次的时间点
        private DateTime m_SecondCharInputTimestamp = DateTime.Now;  //定义一个成员函数用于保存每次的时间点
        private int m_PressCount = 0;

        public PressureForm()
        {
            InitializeComponent();
            InitUI();
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == 0x1EE1)
            {
                ClearPumpNoWhenCompleteTest();
            }
            base.WndProc(ref m);
        }


        private void PressureForm_Load(object sender, EventArgs e)
        {
            InitPumpType();
            LoadSettings();
            LoadConfig();
        }

        /// <summary>
        /// 加载配置文件中的参数
        /// </summary>
        private void LoadConfig()
        {
            try
            {
                System.Configuration.Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                string strInterval = ConfigurationManager.AppSettings.Get("SampleInterval");
                if (!int.TryParse(strInterval, out m_SampleInterval))
                    m_SampleInterval = 500;
                chart1.SampleInterval = m_SampleInterval;
                chart2.SampleInterval = m_SampleInterval;
                string strTool1 = ConfigurationManager.AppSettings.Get("Tool1");
                string strTool2 = ConfigurationManager.AppSettings.Get("Tool2");
                tbToolingNo.Text = strTool1;
                tbToolingNo2.Text = strTool2;

                RangeMinP = Int32.Parse(ConfigurationManager.AppSettings.Get("RangeMinP"));
                RangeMaxP = Int32.Parse(ConfigurationManager.AppSettings.Get("RangeMaxP"));
                PressureCalibrationMax = Int32.Parse(ConfigurationManager.AppSettings.Get("PressureCalibrationMax"));
                SerialNumberCount = Int32.Parse(ConfigurationManager.AppSettings.Get("SerialNumberCount"));
                var samplingPoint = ConfigurationManager.AppSettings.Get("SamplingPoint1");

                //0~3.5kg区间
                string[] strSamplingPoints = samplingPoint.Trim().Split(',');
                SamplingPoints1.Clear();
                foreach (string s in strSamplingPoints)
                {
                    SamplingPoints1.Add(double.Parse(s));
                }

                //4~7kg区间
                samplingPoint = ConfigurationManager.AppSettings.Get("SamplingPoint2");
                strSamplingPoints = samplingPoint.Trim().Split(',');
                SamplingPoints2.Clear();
                foreach (string s in strSamplingPoints)
                {
                    SamplingPoints2.Add(double.Parse(s));
                }

                //单道泵
                samplingPoint = ConfigurationManager.AppSettings.Get("SingleSamplingPoint1");
                //0~3.5kg区间
                strSamplingPoints = samplingPoint.Trim().Split(',');
                SingleSamplingPoints1.Clear();
                foreach (string s in strSamplingPoints)
                {
                    SingleSamplingPoints1.Add(double.Parse(s));
                }

                //4~7kg区间
                samplingPoint = ConfigurationManager.AppSettings.Get("SingleSamplingPoint2");
                strSamplingPoints = samplingPoint.Trim().Split(',');
                SingleSamplingPoints2.Clear();
                foreach (string s in strSamplingPoints)
                {
                    SingleSamplingPoints2.Add(double.Parse(s));
                }


                SamplingPoints.Clear();
                SamplingPoints.AddRange(SamplingPoints1);
                SamplingPoints.AddRange(SamplingPoints2);
                var standardError = ConfigurationManager.AppSettings.Get("StandardError");
                m_StandardError = double.Parse(standardError);

                var samplingError = ConfigurationManager.AppSettings.Get("SamplingError");
                m_SamplingError = double.Parse(samplingError);
                
                #region 不要从config文件读取压力参数
                /*
                #region 读GrasebyC6压力范围
                ConfigurationSectionGroup group = config.GetSectionGroup("GrasebyC6");
                string scetionGroupName = string.Empty;
                PumpID pid = PumpID.GrasebyC6;
                NameValueCollection pressureCollection = null;

                for (int iLoop = 0; iLoop < group.Sections.Count; iLoop++)
                {
                    string scetionName = "GrasebyC6/" + group.Sections.GetKey(iLoop);
                    string strLevel = group.Sections.GetKey(iLoop);
                    Misc.OcclusionLevel level = Misc.OcclusionLevel.None;
                    if (Enum.IsDefined(typeof(Misc.OcclusionLevel), strLevel))
                    {
                        level = (Misc.OcclusionLevel)Enum.Parse(typeof(Misc.OcclusionLevel), strLevel);
                    }
                    pressureCollection = (NameValueCollection)ConfigurationManager.GetSection(scetionName);
                    string key = string.Empty;
                    string pressureValue = string.Empty;
                    int iCount = pressureCollection.Count;
                    for (int k = 0; k < iCount; k++)
                    {
                        key = pressureCollection.GetKey(k);
                        pressureValue = pressureCollection[k].ToString();
                        string[] splitPressure = pressureValue.Split(',');
                        //PressureManager.Instance().Add(pid, level, int.Parse(key), float.Parse(splitPressure[0]), float.Parse(splitPressure[1]), float.Parse(splitPressure[2]));
                    }
                }
                #endregion

                #region 读WZ50C6压力范围
                group = config.GetSectionGroup("WZ50C6");
                pid = PumpID.GrasebyC6;
                for (int iLoop = 0; iLoop < group.Sections.Count; iLoop++)
                {
                    string scetionName = "WZ50C6/" + group.Sections.GetKey(iLoop);
                    string strLevel = group.Sections.GetKey(iLoop);
                    Misc.OcclusionLevel level = Misc.OcclusionLevel.None;
                    if (Enum.IsDefined(typeof(Misc.OcclusionLevel), strLevel))
                    {
                        level = (Misc.OcclusionLevel)Enum.Parse(typeof(Misc.OcclusionLevel), strLevel);
                    }
                    pressureCollection = (NameValueCollection)ConfigurationManager.GetSection(scetionName);
                    string key = string.Empty;
                    string pressureValue = string.Empty;
                    int iCount = pressureCollection.Count;
                    for (int k = 0; k < iCount; k++)
                    {
                        key = pressureCollection.GetKey(k);
                        pressureValue = pressureCollection[k].ToString();
                        string[] splitPressure = pressureValue.Split(',');
                        //PressureManager.Instance().Add(pid, level, int.Parse(key), float.Parse(splitPressure[0]), float.Parse(splitPressure[1]), float.Parse(splitPressure[2]));
                    }
                }
                #endregion

                #region 读GrasebyF6压力范围
                group = config.GetSectionGroup("GrasebyF6");
                pid = PumpID.GrasebyF6;
                for (int iLoop = 0; iLoop < group.Sections.Count; iLoop++)
                {
                    string scetionName = "GrasebyF6/" + group.Sections.GetKey(iLoop);
                    string strLevel = group.Sections.GetKey(iLoop);
                    Misc.OcclusionLevel level = Misc.OcclusionLevel.None;
                    if (Enum.IsDefined(typeof(Misc.OcclusionLevel), strLevel))
                    {
                        level = (Misc.OcclusionLevel)Enum.Parse(typeof(Misc.OcclusionLevel), strLevel);
                    }
                    pressureCollection = (NameValueCollection)ConfigurationManager.GetSection(scetionName);
                    string key = string.Empty;
                    string pressureValue = string.Empty;
                    int iCount = pressureCollection.Count;
                    for (int k = 0; k < iCount; k++)
                    {
                        key = pressureCollection.GetKey(k);
                        pressureValue = pressureCollection[k].ToString();
                        string[] splitPressure = pressureValue.Split(',');
                        //PressureManager.Instance().Add(pid, level, int.Parse(key), float.Parse(splitPressure[0]), float.Parse(splitPressure[1]), float.Parse(splitPressure[2]));
                    }
                }
                #endregion

                #region 读GrasebyC6T压力范围
                group = config.GetSectionGroup("GrasebyC6T");
                pid = PumpID.GrasebyC6T;
                for (int iLoop = 0; iLoop < group.Sections.Count; iLoop++)
                {
                    string scetionName = "GrasebyC6T/" + group.Sections.GetKey(iLoop);
                    string strLevel = group.Sections.GetKey(iLoop);
                    Misc.OcclusionLevel level = Misc.OcclusionLevel.None;
                    if (Enum.IsDefined(typeof(Misc.OcclusionLevel), strLevel))
                    {
                        level = (Misc.OcclusionLevel)Enum.Parse(typeof(Misc.OcclusionLevel), strLevel);
                    }
                    pressureCollection = (NameValueCollection)ConfigurationManager.GetSection(scetionName);
                    string key = string.Empty;
                    string pressureValue = string.Empty;
                    int iCount = pressureCollection.Count;
                    for (int k = 0; k < iCount; k++)
                    {
                        key = pressureCollection.GetKey(k);
                        pressureValue = pressureCollection[k].ToString();
                        string[] splitPressure = pressureValue.Split(',');
                        //PressureManager.Instance().Add(pid, level, int.Parse(key), float.Parse(splitPressure[0]), float.Parse(splitPressure[1]), float.Parse(splitPressure[2]));
                    }
                }
                #endregion

                #region 读Graseby2000压力范围
                group = config.GetSectionGroup("Graseby2000");
                pid = PumpID.Graseby2000;
                for (int iLoop = 0; iLoop < group.Sections.Count; iLoop++)
                {
                    string scetionName = "Graseby2000/" + group.Sections.GetKey(iLoop);
                    string strLevel = group.Sections.GetKey(iLoop);
                    Misc.OcclusionLevel level = Misc.OcclusionLevel.None;
                    if (Enum.IsDefined(typeof(Misc.OcclusionLevel), strLevel))
                    {
                        level = (Misc.OcclusionLevel)Enum.Parse(typeof(Misc.OcclusionLevel), strLevel);
                    }
                    pressureCollection = (NameValueCollection)ConfigurationManager.GetSection(scetionName);
                    string key = string.Empty;
                    string pressureValue = string.Empty;
                    int iCount = pressureCollection.Count;
                    for (int k = 0; k < iCount; k++)
                    {
                        key = pressureCollection.GetKey(k);
                        pressureValue = pressureCollection[k].ToString();
                        string[] splitPressure = pressureValue.Split(',');
                        //PressureManager.Instance().Add(pid, level, int.Parse(key), float.Parse(splitPressure[0]), float.Parse(splitPressure[1]), float.Parse(splitPressure[2]));
                    }
                }
                #endregion

                #region 读Graseby2100压力范围
                group = config.GetSectionGroup("Graseby2100");
                pid = PumpID.Graseby2100;
                for (int iLoop = 0; iLoop < group.Sections.Count; iLoop++)
                {
                    string scetionName = "Graseby2100/" + group.Sections.GetKey(iLoop);
                    string strLevel = group.Sections.GetKey(iLoop);
                    Misc.OcclusionLevel level = Misc.OcclusionLevel.None;
                    if (Enum.IsDefined(typeof(Misc.OcclusionLevel), strLevel))
                    {
                        level = (Misc.OcclusionLevel)Enum.Parse(typeof(Misc.OcclusionLevel), strLevel);
                    }
                    pressureCollection = (NameValueCollection)ConfigurationManager.GetSection(scetionName);
                    string key = string.Empty;
                    string pressureValue = string.Empty;
                    int iCount = pressureCollection.Count;
                    for (int k = 0; k < iCount; k++)
                    {
                        key = pressureCollection.GetKey(k);
                        pressureValue = pressureCollection[k].ToString();
                        string[] splitPressure = pressureValue.Split(',');
                        //PressureManager.Instance().Add(pid, level, int.Parse(key), float.Parse(splitPressure[0]), float.Parse(splitPressure[1]), float.Parse(splitPressure[2]));
                    }
                }
                #endregion

                #region 读WZ50F6压力范围
                group = config.GetSectionGroup("WZS50F6");
                pid = PumpID.WZS50F6;
                for (int iLoop = 0; iLoop < group.Sections.Count; iLoop++)
                {
                    string scetionName = "WZS50F6/" + group.Sections.GetKey(iLoop);
                    string strLevel = group.Sections.GetKey(iLoop);
                    Misc.OcclusionLevel level = Misc.OcclusionLevel.None;
                    if (Enum.IsDefined(typeof(Misc.OcclusionLevel), strLevel))
                    {
                        level = (Misc.OcclusionLevel)Enum.Parse(typeof(Misc.OcclusionLevel), strLevel);
                    }
                    pressureCollection = (NameValueCollection)ConfigurationManager.GetSection(scetionName);
                    string key = string.Empty;
                    string pressureValue = string.Empty;
                    int iCount = pressureCollection.Count;
                    for (int k = 0; k < iCount; k++)
                    {
                        key = pressureCollection.GetKey(k);
                        pressureValue = pressureCollection[k].ToString();
                        string[] splitPressure = pressureValue.Split(',');
                        //PressureManager.Instance().Add(pid, level, int.Parse(key), float.Parse(splitPressure[0]), float.Parse(splitPressure[1]), float.Parse(splitPressure[2]));
                    }
                }
                #endregion

                #region 读WZ50C6T压力范围
                group = config.GetSectionGroup("WZ50C6T");
                pid = PumpID.WZ50C6T;
                for (int iLoop = 0; iLoop < group.Sections.Count; iLoop++)
                {
                    string scetionName = "WZ50C6T/" + group.Sections.GetKey(iLoop);
                    string strLevel = group.Sections.GetKey(iLoop);
                    Misc.OcclusionLevel level = Misc.OcclusionLevel.None;
                    if (Enum.IsDefined(typeof(Misc.OcclusionLevel), strLevel))
                    {
                        level = (Misc.OcclusionLevel)Enum.Parse(typeof(Misc.OcclusionLevel), strLevel);
                    }
                    pressureCollection = (NameValueCollection)ConfigurationManager.GetSection(scetionName);
                    string key = string.Empty;
                    string pressureValue = string.Empty;
                    int iCount = pressureCollection.Count;
                    for (int k = 0; k < iCount; k++)
                    {
                        key = pressureCollection.GetKey(k);
                        pressureValue = pressureCollection[k].ToString();
                        string[] splitPressure = pressureValue.Split(',');
                        //PressureManager.Instance().Add(pid, level, int.Parse(key), float.Parse(splitPressure[0]), float.Parse(splitPressure[1]), float.Parse(splitPressure[2]));
                    }
                }
                #endregion
                */
                #endregion

            }
            catch (Exception ex)
            {
                MessageBox.Show("PTool.config文件参数配置错误，请先检查该文件后再重新启动程序!" + ex.Message);
            }
        }

        private void LoadSettings()
        {
            string currentPath = Assembly.GetExecutingAssembly().Location;
            currentPath = currentPath.Substring(0, currentPath.LastIndexOf('\\'));  //删除文件名
            string iniPath = currentPath + "\\ptool.ini";
            IniReader reader = new IniReader(iniPath);
            reader.ReadSettings();
        }

        private void SaveLastToolingNo()
        {
            Configuration cfa = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            cfa.AppSettings.Settings["Tool1"].Value = tbToolingNo.Text;
            cfa.AppSettings.Settings["Tool2"].Value = tbToolingNo2.Text;
            cfa.Save();
        }

        private void InitPumpType()
        {
            cbPumpType.Items.Clear();
            cbPumpType.Items.AddRange(ProductIDConvertor.GetAllPumpIDString().ToArray());
            cbPumpType.SelectedIndex = 0;
            m_LocalPid = ProductIDConvertor.String2PumpID(cbPumpType.Items[cbPumpType.SelectedIndex].ToString());
            chart1.SetPid(m_LocalPid);
            chart2.SetPid(m_LocalPid);
            chart1.SetChannel(1);
            chart2.SetChannel(2);
        }

        private void InitUI()
        {
            lbTitle.ForeColor = Color.FromArgb(3, 116, 214);
            tlpParameter.BackColor = Color.FromArgb(19, 113, 185);
            cbPumpType.BackColor = Color.FromArgb(19, 113, 185);
            tbPumpNo.BackColor = Color.FromArgb(19, 113, 185);
            tbToolingNo.BackColor = Color.FromArgb(19, 113, 185);
            tbToolingNo2.BackColor = Color.FromArgb(19, 113, 185);
            chart1.Channel = 1;
            chart2.Channel = 2;
            chart2.Enabled = false;
            chart1.SamplingStartOrStop += OnSamplingStartOrStop;
            chart2.SamplingStartOrStop += OnSamplingStartOrStop;
            chart1.OnSamplingComplete += OnChartSamplingComplete;
            chart2.OnSamplingComplete += OnChartSamplingComplete;
            chart1.OpratorNumberInput += OnOpratorNumberInput;
            hashSampleData.Clear();
        }

        /// <summary>
        /// 双道泵测量数据统一放进m_SampleDataList中，第一道数据索引为0，第二道为1
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnChartSamplingComplete(object sender, DoublePumpDataArgs e)
        {
            Chart chart = sender as Chart;
            if (e.SampleDataList != null)
            {
                if (chart.Name == "chart1")
                {
                    if (hashSampleData.ContainsKey(1))
                        hashSampleData[1] = e.SampleDataList;
                    else
                        hashSampleData.Add(1, e.SampleDataList);
                }
                else
                {
                    if (hashSampleData.ContainsKey(2))
                        hashSampleData[2] = e.SampleDataList;
                    else
                        hashSampleData.Add(2, e.SampleDataList);
                }
            }
            if(hashSampleData.Count>=2)
            {
                //写入excel,调用chart类中函数
                string path = Path.GetDirectoryName(Assembly.GetAssembly(typeof(PressureForm)).Location) + "\\数据导出";
                PumpID pid = PumpID.None;
                switch (m_LocalPid)
                {
                    case PumpID.GrasebyF6_2:
                        pid = PumpID.GrasebyF6;
                        break;
                    case PumpID.WZS50F6_2:
                        pid = PumpID.WZS50F6;
                        break;
                    default:
                        pid = m_LocalPid;
                        break;
                }
                string fileName = string.Format("{0}_{1}_{2}", pid.ToString(), tbPumpNo.Text, DateTime.Now.ToString("yyyy-MM-dd HH_mm_ss"));
                if (!System.IO.Directory.Exists(path))
                    System.IO.Directory.CreateDirectory(path);
                string saveFileName = path + "\\" + fileName + ".xlsx";


                string path2 = Path.GetDirectoryName(Assembly.GetAssembly(typeof(PressureForm)).Location) + "\\数据导出备份";
                string fileName2 = string.Format("{0}_{1}_{2}", pid.ToString(), tbPumpNo.Text, DateTime.Now.ToString("yyyy-MM-dd HH_mm_ss"));
                if (!System.IO.Directory.Exists(path2))
                    System.IO.Directory.CreateDirectory(path2);
                string saveFileName2 = path2 + "\\" + fileName2 + ".xlsx";

                List<List<PressureCalibrationParameter>> sampleDataList = new List<List<PressureCalibrationParameter>>();
                if(hashSampleData.ContainsKey(1))
                    sampleDataList.Add(hashSampleData[1] as List<PressureCalibrationParameter>);
                if (hashSampleData.ContainsKey(2))
                    sampleDataList.Add(hashSampleData[2] as List<PressureCalibrationParameter>);
                if (sampleDataList.Count == 2)
                    chart1.GenDoublePunmpReport(saveFileName, sampleDataList, tbToolingNo2.Text, saveFileName2);
                else
                    Logger.Instance().ErrorFormat("双道泵保存数据异常，由于结果数量不等于2，无法保存Count={0}", sampleDataList.Count);
                Thread.Sleep(2000);
                hashSampleData.Clear();
            }
        }

        private void OnSamplingStartOrStop(object sender, EventArgs e)
        {
            StartOrStopArgs args = e as StartOrStopArgs;
            cbPumpType.Enabled = args.IsStart;
            chart1.ToolingNo = tbToolingNo.Text;
            chart2.ToolingNo = tbToolingNo2.Text;
            chart1.PumpNo = tbPumpNo.Text;
            chart2.PumpNo = tbPumpNo.Text;
        }

        private void OnOpratorNumberInput(object sender, OpratorNumberArgs e)
        {
            Chart chart1 = sender as Chart;
            if(chart1.Channel==1)
            {
                chart2.InputOpratorNumber(e.Number);
            }
          
           
        }

        private void tlpTitle_MouseDown(object sender, MouseEventArgs e)
        {
            if (this.WindowState == FormWindowState.Maximized)
            {
                return;
            }
            oldMousePosition = e.Location;
            moving = true; 
        }

        private void tlpTitle_MouseUp(object sender, MouseEventArgs e)
        {
            moving = false;
        }

        private void tlpTitle_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left && moving)
            {
                Point newPosition = new Point(e.Location.X - oldMousePosition.X, e.Location.Y - oldMousePosition.Y);
                this.Location += new Size(newPosition);
            }
        }

        private void cbPumpType_SelectedIndexChanged(object sender, EventArgs e)
        {
            m_LocalPid = ProductIDConvertor.String2PumpID(cbPumpType.Items[cbPumpType.SelectedIndex].ToString());
#if DEBUG
            chart2.Enabled = true;

#else
            if (m_LocalPid == PumpID.GrasebyF6 || m_LocalPid == PumpID.WZS50F6 || m_LocalPid == PumpID.GrasebyF6_2 || m_LocalPid == PumpID.WZS50F6_2)
            {
                chart2.Enabled = true;
            }
            else
            {
                chart2.Enabled = false;
            }
#endif
            chart1.SetPid(m_LocalPid);
            chart2.SetPid(m_LocalPid);
            chart1.SetChannel(1);
            chart2.SetChannel(2);
        }

        private void picCloseWindow_Click(object sender, EventArgs e)
        {
            chart1.SamplingStartOrStop -= OnSamplingStartOrStop;
            chart2.SamplingStartOrStop -= OnSamplingStartOrStop;
            chart1.OnSamplingComplete -= OnChartSamplingComplete;
            chart2.OnSamplingComplete -= OnChartSamplingComplete;
            SaveLastToolingNo();
            this.Close();
        }

        /// <summary>
        /// 采样结束，清空产品序号 20180820
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ClearPumpNoWhenCompleteTest()
        {
            tbPumpNo.Clear();
        }

        private void tbPumpNo_KeyPress(object sender, KeyPressEventArgs e)
        {
            TimeSpan ts;
            m_SecondCharInputTimestamp = DateTime.Now;
            ts = m_SecondCharInputTimestamp.Subtract(m_FirstCharInputTimestamp);     //获取时间间隔
            if (ts.Milliseconds < INPUTSPEED)
                m_PressCount++;
            else
            {
                m_PressCount = 0;
            }

            if (m_PressCount == SerialNumberCount)
            {
                if (tbPumpNo.Text.Length >= SerialNumberCount)
                {
                    if (tbPumpNo.SelectionStart < tbPumpNo.Text.Length)
                        tbPumpNo.Text = tbPumpNo.Text.Remove(tbPumpNo.SelectionStart);
                    try
                    {
                        tbPumpNo.Text = tbPumpNo.Text.Substring(tbPumpNo.Text.Length - SerialNumberCount, SerialNumberCount);
                        tbPumpNo.SelectionStart = tbPumpNo.Text.Length;
                    }
                    catch
                    {
                        tbPumpNo.Text = "";
                    }
                }
                m_PressCount = 0;
            }
            m_FirstCharInputTimestamp = m_SecondCharInputTimestamp;
        }
    }

    public class SampleData
    {
        public DateTime m_SampleTime = DateTime.Now;
        public float m_PressureValue;
        public float m_Weight;

        public SampleData()
        {
        }

        public SampleData(DateTime sampleTime, float pressureVale, float weight)
        {
            m_SampleTime = sampleTime;
            m_PressureValue = pressureVale;
            m_Weight = weight;
        }

        public void Copy(SampleData other)
        {
            this.m_SampleTime = other.m_SampleTime;
            this.m_PressureValue = other.m_PressureValue;
            this.m_Weight = other.m_Weight;
        }
    }
}
