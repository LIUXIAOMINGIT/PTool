using System;
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
using  Misc = ComunicationProtocol.Misc;
using SerialDevice;

namespace PTool
{
    public partial class MainForm : CCSkinMain
    {
        private readonly int BAUDRATE = 9600;
        private const string VOL = "P值";
        private const int LEFTBORDEROFFSET = 30;
        private const int RIGHTBORDEROFFSET = 10;
        private const int BOTTOMBORDEROFFSET = 30;     //X坐标与下边距，一般是绘图区域的一半高度
        private const int TOPBOTTOMFFSET = 5;      //坐标上下边距
        private const int CIRCLEDIAMETER = 5;      //曲线图上的圆点直径0
        private const int TRYCOUNTSAMPLINGTIMEOUT = 5;      //采样超时次数为5.超时5次就停止 
        private Graphics m_gh = null;
        private Graphics m_gh2 = null;
        private System.Drawing.Rectangle m_Rect;
        private System.Drawing.Rectangle m_Rect2;

        private Pen m_WaveLinePen = new Pen(Color.FromArgb(128, 208, 255));


        private float m_XCoordinateMaxValue = 10;
        private int m_YCoordinateMaxValue = 5;
        private int m_XSectionCount = 10;
        private int m_YSectionCount = 5;
        private float m_CoordinateIntervalX = 0;  //X轴上的区间实际长度，单位为像素
        private float m_CoordinateIntervalY = 0;  //Y轴上的区间实际长度，单位为像素
        private float m_ValueInervalX = 0;  //X轴上的坐标值，根据实际放大倍数和量程决定
        private float m_ValueInervalY = 0;
        protected GlobalResponse m_ConnResponseCh1 = null;//单道泵
        protected GlobalResponse m_ConnResponseCh2 = null;//双道泵
        private PTooling m_PToolCh1 = null;//单道泵
        private PTooling m_PToolCh2 = null;//双道泵
        private PTooling m_DetectPToolCh1 = null;//单道泵,只是用来检测
        private PTooling m_DetectPToolCh2 = null;//双道泵只是用来检测
        private Graseby9600 m_GrasebyDevice = new Graseby9600();//只用于串口刷新
        private Graseby9600 m_GrasebyDevice2 = new Graseby9600();//只用于串口刷新
        private ProductID m_LocalPid = ProductID.GrasebyC6;//默认显示的是C6
        private int m_SampleInterval = 500;//采样频率：毫秒
        private System.Timers.Timer m_Ch1Timer = new System.Timers.Timer();
        private System.Timers.Timer m_Ch2Timer = new System.Timers.Timer();

        //两个通道采样数据
        private List<SampleData> m_Ch1SampleDataList = new List<SampleData>();
        private List<SampleData> m_Ch2SampleDataList = new List<SampleData>();

        public delegate void DelegateSetWeightValue(float weight, bool isDetect);
        public delegate void DelegateSetPValue(float p);
        public delegate void DelegateEnableContols(int channel, bool bEnabled);
        public delegate void DelegateAlertMessageWhenComplete(string msg);
        
        public MainForm()
        {
            InitializeComponent();
            m_PToolCh1 = new PTooling();
            m_PToolCh2 = new PTooling();
            m_PToolCh1.DeviceDataRecerived += OnPToolCh1_DeviceDataRecerived;
            m_PToolCh2.DeviceDataRecerived += OnPToolCh2_DeviceDataRecerived;

            m_DetectPToolCh1 = new PTooling();
            m_DetectPToolCh2 = new PTooling();
            m_DetectPToolCh1.DeviceDataRecerived += OnPToolCh1_DetectDeviceDataRecerived;
            m_DetectPToolCh2.DeviceDataRecerived += OnPToolCh2_DetectDeviceDataRecerived;

            m_GrasebyDevice.DeviceDataRecerived += OnGrasebyDeviceDataRecerived;
            m_GrasebyDevice2.DeviceDataRecerived += OnGrasebyDevice2DataRecerived;
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            cbPumpPort.Items.AddRange(SerialPort.GetPortNames());
            cbToolingPort.Items.AddRange(SerialPort.GetPortNames());
            cbPumpPort2.Items.AddRange(SerialPort.GetPortNames());
            cbToolingPort2.Items.AddRange(SerialPort.GetPortNames());
            InitPumpType();
            m_gh = WavelinePanel.CreateGraphics();
            m_gh2 = WavelinePanel2.CreateGraphics();
            m_Rect = WavelinePanel.ClientRectangle;
            m_Rect2 = WavelinePanel2.ClientRectangle;
            LoadConfig();
            m_Ch1Timer.Interval = m_SampleInterval;
            m_Ch1Timer.Elapsed += OnChannel1Timer_Elapsed;
            m_Ch2Timer.Interval = m_SampleInterval;
            m_Ch2Timer.Elapsed += OnChannel2Timer_Elapsed;
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

                #region 读GrasebyC6压力范围
                ConfigurationSectionGroup group = config.GetSectionGroup("GrasebyC6");
                string scetionGroupName = string.Empty;
                ProductID pid = ProductID.GrasebyC6;
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
                        PressureManager.Instance().Add(pid, level, int.Parse(key), float.Parse(splitPressure[0]), float.Parse(splitPressure[1]), float.Parse(splitPressure[2]));
                    }
                }
                #endregion

                #region 读WZ50C6压力范围
                group = config.GetSectionGroup("WZ50C6");
                pid = ProductID.WZ50C6;
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
                        PressureManager.Instance().Add(pid, level, int.Parse(key), float.Parse(splitPressure[0]), float.Parse(splitPressure[1]), float.Parse(splitPressure[2]));
                    }
                }
                #endregion

                #region 读GrasebyF6压力范围
                group = config.GetSectionGroup("GrasebyF6");
                pid = ProductID.GrasebyF6;
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
                        PressureManager.Instance().Add(pid, level, int.Parse(key), float.Parse(splitPressure[0]), float.Parse(splitPressure[1]), float.Parse(splitPressure[2]));
                    }
                }
                #endregion
            }
            catch (Exception ex)
            {
                MessageBox.Show("PTool.config文件参数配置错误，请先检查该文件后再重新启动程序!");
            }
        }

        public void AddHandler(int channel = 1)
        {
            if (channel == 1)
            {
                if (m_ConnResponseCh1 != null)
                {
                    m_ConnResponseCh1.SetVTBIParameterResponse += new EventHandler<ResponseEventArgs<String>>(SetInfusionParas);
                    m_ConnResponseCh1.SetStartControlResponse += new EventHandler<ResponseEventArgs<String>>(SetStartControl);
                    m_ConnResponseCh1.SetStopControlResponse += new EventHandler<ResponseEventArgs<String>>(SetStopControl);
                    m_ConnResponseCh1.GetPressureSensorResponse += new EventHandler<ResponseEventArgs<Misc.PressureSensorInfo>>(GetPressureSensor);
                    m_ConnResponseCh1.SetPressureCalibrationParameterResponse += new EventHandler<ResponseEventArgs<String>>(SetPressureCalibrationParameter);
                }
            }
            else
            {
                if (m_ConnResponseCh2 != null)
                {
                    m_ConnResponseCh2.SetVTBIParameterResponse += new EventHandler<ResponseEventArgs<String>>(SetInfusionParas2);
                    m_ConnResponseCh2.SetStartControlResponse += new EventHandler<ResponseEventArgs<String>>(SetStartControl2);
                    m_ConnResponseCh2.SetStopControlResponse += new EventHandler<ResponseEventArgs<String>>(SetStopControl2);
                    m_ConnResponseCh2.GetPressureSensorResponse += new EventHandler<ResponseEventArgs<Misc.PressureSensorInfo>>(GetPressureSensor2);
                    m_ConnResponseCh2.SetPressureCalibrationParameterResponse += new EventHandler<ResponseEventArgs<String>>(SetPressureCalibrationParameter2);
                }
            }
        }

        public void RemoveHandler(int channel = 1)
        {
            if (channel == 1)
            {
                if (m_ConnResponseCh1 != null)
                {
                    m_ConnResponseCh1.SetVTBIParameterResponse -= new EventHandler<ResponseEventArgs<String>>(SetInfusionParas);
                    m_ConnResponseCh1.SetStartControlResponse -= new EventHandler<ResponseEventArgs<String>>(SetStartControl);
                    m_ConnResponseCh1.SetStopControlResponse -= new EventHandler<ResponseEventArgs<String>>(SetStopControl);
                    m_ConnResponseCh1.GetPressureSensorResponse -= new EventHandler<ResponseEventArgs<Misc.PressureSensorInfo>>(GetPressureSensor);
                    m_ConnResponseCh1.SetPressureCalibrationParameterResponse -= new EventHandler<ResponseEventArgs<String>>(SetPressureCalibrationParameter2);
                }
            }
            else
            {
                if (m_ConnResponseCh2 != null)
                {
                    m_ConnResponseCh2.SetVTBIParameterResponse -= new EventHandler<ResponseEventArgs<String>>(SetInfusionParas2);
                    m_ConnResponseCh2.SetStartControlResponse -= new EventHandler<ResponseEventArgs<String>>(SetStartControl2);
                    m_ConnResponseCh2.SetStopControlResponse -= new EventHandler<ResponseEventArgs<String>>(SetStopControl2);
                    m_ConnResponseCh2.GetPressureSensorResponse -= new EventHandler<ResponseEventArgs<Misc.PressureSensorInfo>>(GetPressureSensor2);
                    m_ConnResponseCh2.SetPressureCalibrationParameterResponse -= new EventHandler<ResponseEventArgs<String>>(SetPressureCalibrationParameter2);
                }
            }
        }

        private void OnPToolCh1_DeviceDataRecerived(object sender, EventArgs e)
        {
            PToolingDataEventArgs args = e as PToolingDataEventArgs;
            lock (m_Ch1SampleDataList)
            {
                if (m_Ch1SampleDataList.Count > 0)
                {
                    SampleData sp = m_Ch1SampleDataList[m_Ch1SampleDataList.Count - 1];
                    sp.m_Weight = args.Weight;
                }
            }
            SetWeightValue(args.Weight, false);
            DrawSingleAccuracyMap();
            //当采集到的重量大于配置参数时，可以停止采集，并计算相关数据写入到泵中
            float max = PressureManager.Instance().GetMaxBySizeLevel(m_LocalPid, 50, Misc.OcclusionLevel.H);
            if (max <= args.Weight)
            {
                //是否要自动停止？？？？？？？？？？？？？？？
                Complete(1);
                AlertMessageWhenComplete(string.Format("压力值已超出最大范围{0},调试结束!", max));
            }
        }

        private void OnPToolCh2_DeviceDataRecerived(object sender, EventArgs e)
        {
            PToolingDataEventArgs args = e as PToolingDataEventArgs;
            lock (m_Ch2SampleDataList)
            {
                if (m_Ch2SampleDataList.Count > 0)
                {
                    SampleData sp = m_Ch2SampleDataList[m_Ch2SampleDataList.Count - 1];
                    sp.m_Weight = args.Weight;
                }
            }
            SetWeightValue2(args.Weight, false);
            DrawSingleAccuracyMap2();
            //当采集到的重量大于配置参数时，可以停止采集，并计算相关数据写入到泵中
            float max = PressureManager.Instance().GetMaxBySizeLevel(m_LocalPid, 50, Misc.OcclusionLevel.H);
            if (max <= args.Weight)
            {
                //是否要自动停止？？？？？？？？？？？？？？？
                Complete(2);
                AlertMessageWhenComplete(string.Format("压力值已超出最大范围{0},调试结束!", max));
            }
        }

        private void AlertMessageWhenComplete(string msg)
        {
            if(this.InvokeRequired)
            {
                this.Invoke(new DelegateAlertMessageWhenComplete(AlertMessageWhenComplete), new object[] { msg });
                return;
            }
            MessageBox.Show(msg);
        }

        /// <summary>
        ///仅检测串口使用
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnPToolCh1_DetectDeviceDataRecerived(object sender, EventArgs e)
        {
            PToolingDataEventArgs args = e as PToolingDataEventArgs;
            SetWeightValue(args.Weight, true);
        }

        /// <summary>
        ///仅检测串口使用
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnPToolCh2_DetectDeviceDataRecerived(object sender, EventArgs e)
        {
            PToolingDataEventArgs args = e as PToolingDataEventArgs;
            SetWeightValue2(args.Weight, true);
        }

        /// <summary>
        /// 仅检测串口使用
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnGrasebyDeviceDataRecerived(object sender, EventArgs e)
        {
            Graseby9600DataEventArgs args = e as Graseby9600DataEventArgs;
            SetPValue(args.SensorValue);
        }

        /// <summary>
        /// 仅检测串口使用
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnGrasebyDevice2DataRecerived(object sender, EventArgs e)
        {
            Graseby9600DataEventArgs args = e as Graseby9600DataEventArgs;
            SetPValue2(args.SensorValue);
        }

        private void SetWeightValue(float weight, bool isDetect = false)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new DelegateSetWeightValue(SetWeightValue), new object[] { weight, isDetect });
            }
            lbWeightChannel1.Text = weight.ToString("F3");
            if (isDetect)
                m_DetectPToolCh1.Close();
        }

        private void SetWeightValue2(float weight, bool isDetect = false)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new DelegateSetWeightValue(SetWeightValue2), new object[] { weight, isDetect });
            }
            lbWeightChannel2.Text = weight.ToString("F3");
            if (isDetect)
                m_DetectPToolCh2.Close();
        }

        private void SetPValue(float p)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new DelegateSetPValue(SetPValue), new object[] { p });
                return;
            }
            lbPChannel1.Text = p.ToString("F2");
            m_GrasebyDevice.Close();
        }

        private void SetPValue2(float p)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new DelegateSetPValue(SetPValue2), new object[] { p });
                return;
            }
            lbPChannel2.Text = p.ToString("F2");
            m_GrasebyDevice2.Close();
        }

        private void InitPumpType()
        {
            cbPumpType.Items.Clear();
            cbPumpType.Items.AddRange(Enum.GetNames(typeof(ProductID)));
            cbPumpType.SelectedIndex = 0;
            m_LocalPid = (ProductID)Enum.Parse(typeof(ProductID), cbPumpType.Items[cbPumpType.SelectedIndex].ToString(), true);
        }

        #region 时钟
        private void StartCh1Timer()
        {
            StopCh1Timer();
            m_Ch1Timer.Start();
        }

        private void StopCh1Timer()
        {
            m_Ch1Timer.Stop();
        }

        private void OnChannel1Timer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            if (m_ConnResponseCh1 != null && m_ConnResponseCh1.IsOpen())
            {
                m_ConnResponseCh1.GetPressureSensor();
            }
        }

        private void StartCh2Timer()
        {
            StopCh2Timer();
            m_Ch2Timer.Start();
        }

        private void StopCh2Timer()
        {
            m_Ch2Timer.Stop();
        }

        private void OnChannel2Timer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            if (m_ConnResponseCh2 != null && m_ConnResponseCh2.IsOpen())
            {
                m_ConnResponseCh2.GetPressureSensor();
            }
        }


        #endregion

        #region 绘图

        private void WavelinePanel_Paint(object sender, PaintEventArgs e)
        {
            DrawCoordinate(m_XCoordinateMaxValue, m_XSectionCount, m_YCoordinateMaxValue, m_YSectionCount);
            DrawAccuracyMap(m_XSectionCount, m_YSectionCount);
        }

        private void WavelinePanel2_Paint(object sender, PaintEventArgs e)
        {
            DrawCoordinate2(m_XCoordinateMaxValue, m_XSectionCount, m_YCoordinateMaxValue, m_YSectionCount);
            DrawAccuracyMap2(m_XSectionCount, m_YSectionCount);
        }

        /// <summary>
        /// 画坐标轴
        /// </summary>
        /// <param name="xMax">X坐标最大值</param>
        /// <param name="xSectionCount">X坐标分成几段</param>
        /// <param name="yMax">Y坐标最大值</param>
        /// <param name="ySectionCount">Y坐标分成几段</param>
        private void DrawCoordinate(float xMax, int xSectionCount, int yMax, int ySectionCount)
        {
            try
            {
                Rectangle rect = m_Rect;
                Font xValuefont = new Font("宋体", 7);
                Font fontTitle = new Font("宋体", 8);
                //画X轴
                PointF originalpoint = new PointF((float)rect.Left + LEFTBORDEROFFSET, rect.Bottom - BOTTOMBORDEROFFSET);
                PointF xEndPoint = new PointF((float)rect.Right - RIGHTBORDEROFFSET, rect.Bottom - BOTTOMBORDEROFFSET);
                m_gh.DrawLine(Pens.Black, originalpoint, xEndPoint);
                //画X坐标箭头
                PointF arrowpointUp = new PointF(xEndPoint.X - 12, xEndPoint.Y - 6);
                PointF arrowpointDwon = new PointF(xEndPoint.X - 12, xEndPoint.Y + 6);
                m_gh.DrawLine(Pens.Black, arrowpointUp, xEndPoint);
                m_gh.DrawLine(Pens.Black, arrowpointDwon, xEndPoint);

                //画X轴坐标,SECTIONCOUNT个点
                float intervalX = (xEndPoint.X - originalpoint.X) / xSectionCount;
                m_CoordinateIntervalX = intervalX;
                float lineSegmentHeight = 8f;
                for (int i = 1; i <= xSectionCount - 1; i++)
                {
                    m_gh.DrawLine(Pens.Black, new PointF(originalpoint.X + intervalX * i, originalpoint.Y), new PointF(originalpoint.X + intervalX * i, originalpoint.Y - lineSegmentHeight));
                }
                //画X坐标文字
                m_gh.DrawString("重量(kg)", fontTitle, Brushes.Black, new PointF(originalpoint.X + intervalX * xSectionCount - 65, originalpoint.Y + 18));
                //画X坐标值
                float xValueInerval = xMax / xSectionCount;
                m_ValueInervalX = xValueInerval;
                for (int i = 0; i <= xSectionCount; i++)
                {
                    if (i == 0)
                    {
                        // m_gh.DrawString("0", xValuefont, Brushes.Black, new PointF(originalpoint.X - 5, originalpoint.Y + 5));
                    }
                    else
                        m_gh.DrawString((i * xValueInerval).ToString(), xValuefont, Brushes.Black, new PointF(originalpoint.X + intervalX * i - 8, originalpoint.Y + 5));
                }
                //画Y轴
                PointF yEndPoint = new PointF((float)rect.Left + LEFTBORDEROFFSET, (float)rect.Top + TOPBOTTOMFFSET);
                //y轴的起始点，从底部往上
                PointF yOriginalPoint = originalpoint;//new PointF((float)rect.Left + LEFTBORDEROFFSET, rect.Bottom - TOPBOTTOMFFSET);
                m_gh.DrawLine(Pens.Black, yOriginalPoint, yEndPoint);
                //画Y坐标箭头
                PointF arrowpointLeft = new PointF(yEndPoint.X - 6, yEndPoint.Y + 12);
                PointF arrowpointRight = new PointF(yEndPoint.X + 6, yEndPoint.Y + 12);
                m_gh.DrawLine(Pens.Black, arrowpointLeft, yEndPoint);
                m_gh.DrawLine(Pens.Black, arrowpointRight, yEndPoint);
                //画Y坐标文字
                m_gh.DrawString("压力值(V)", fontTitle, Brushes.Black, new PointF(yEndPoint.X + 10, yEndPoint.Y));
                //画Y轴坐标,每个区间的实际坐标长度
                float intervalY = Math.Abs(yEndPoint.Y - yOriginalPoint.Y) / ySectionCount;
                m_CoordinateIntervalY = intervalY;
                for (int i = 0; i < ySectionCount; i++)
                {
                    m_gh.DrawLine(Pens.Black, new PointF(yOriginalPoint.X, yOriginalPoint.Y - intervalY * i), new PointF(yOriginalPoint.X + lineSegmentHeight, yOriginalPoint.Y - intervalY * i));
                }
                float yValueInerval = (float)yMax / ySectionCount;
                m_ValueInervalY = yValueInerval;//Y轴上的坐标值，根据实际放大倍数和量程决定
                for (int i = 0; i <= ySectionCount; i++)
                {
                    m_gh.DrawString(i.ToString(), xValuefont, Brushes.Black, new PointF(yOriginalPoint.X - 24, yOriginalPoint.Y - intervalY * i - 6));
                }
                //画legend
                m_gh.DrawString(VOL, fontTitle, Brushes.Black, new PointF(xEndPoint.X - 40, 10));
                SizeF fontSize = m_gh.MeasureString(VOL, fontTitle);

                m_gh.DrawLine(new Pen(Color.Red), new PointF(xEndPoint.X - 60, 10 + fontSize.Height / 2), new PointF(xEndPoint.X - 40, 10 + fontSize.Height / 2));
            }
            catch (Exception e)
            {
                MessageBox.Show("DrawHomeostasisMap Error:" + e.Message);
            }
        }

        /// <summary>
        /// 画坐标轴
        /// </summary>
        /// <param name="xMax">X坐标最大值</param>
        /// <param name="xSectionCount">X坐标分成几段</param>
        /// <param name="yMax">Y坐标最大值</param>
        /// <param name="ySectionCount">Y坐标分成几段</param>
        private void DrawCoordinate2(float xMax, int xSectionCount, int yMax, int ySectionCount)
        {
            try
            {
                Rectangle rect = m_Rect2;
                Font xValuefont = new Font("宋体", 7);
                Font fontTitle = new Font("宋体", 8);
                //画X轴
                PointF originalpoint = new PointF((float)rect.Left + LEFTBORDEROFFSET, rect.Bottom - BOTTOMBORDEROFFSET);
                PointF xEndPoint = new PointF((float)rect.Right - RIGHTBORDEROFFSET, rect.Bottom - BOTTOMBORDEROFFSET);
                m_gh2.DrawLine(Pens.Black, originalpoint, xEndPoint);
                //画X坐标箭头
                PointF arrowpointUp = new PointF(xEndPoint.X - 12, xEndPoint.Y - 6);
                PointF arrowpointDwon = new PointF(xEndPoint.X - 12, xEndPoint.Y + 6);
                m_gh2.DrawLine(Pens.Black, arrowpointUp, xEndPoint);
                m_gh2.DrawLine(Pens.Black, arrowpointDwon, xEndPoint);

                //画X轴坐标,SECTIONCOUNT个点
                float intervalX = (xEndPoint.X - originalpoint.X) / xSectionCount;
                m_CoordinateIntervalX = intervalX;
                float lineSegmentHeight = 8f;
                for (int i = 1; i <= xSectionCount - 1; i++)
                {
                    m_gh2.DrawLine(Pens.Black, new PointF(originalpoint.X + intervalX * i, originalpoint.Y), new PointF(originalpoint.X + intervalX * i, originalpoint.Y - lineSegmentHeight));
                }
                //画X坐标文字
                m_gh2.DrawString("重量(kg)", fontTitle, Brushes.Black, new PointF(originalpoint.X + intervalX * xSectionCount - 65, originalpoint.Y + 18));
                //画X坐标值
                float xValueInerval = xMax / xSectionCount;
                m_ValueInervalX = xValueInerval;
                for (int i = 0; i <= xSectionCount; i++)
                {
                    if (i == 0)
                    {
                        // m_gh2.DrawString("0", xValuefont, Brushes.Black, new PointF(originalpoint.X - 5, originalpoint.Y + 5));
                    }
                    else
                        m_gh2.DrawString((i * xValueInerval).ToString(), xValuefont, Brushes.Black, new PointF(originalpoint.X + intervalX * i - 8, originalpoint.Y + 5));
                }
                //画Y轴
                PointF yEndPoint = new PointF((float)rect.Left + LEFTBORDEROFFSET, (float)rect.Top + TOPBOTTOMFFSET);
                //y轴的起始点，从底部往上
                PointF yOriginalPoint = originalpoint;//new PointF((float)rect.Left + LEFTBORDEROFFSET, rect.Bottom - TOPBOTTOMFFSET);
                m_gh2.DrawLine(Pens.Black, yOriginalPoint, yEndPoint);
                //画Y坐标箭头
                PointF arrowpointLeft = new PointF(yEndPoint.X - 6, yEndPoint.Y + 12);
                PointF arrowpointRight = new PointF(yEndPoint.X + 6, yEndPoint.Y + 12);
                m_gh2.DrawLine(Pens.Black, arrowpointLeft, yEndPoint);
                m_gh2.DrawLine(Pens.Black, arrowpointRight, yEndPoint);
                //画Y坐标文字
                m_gh2.DrawString("压力值(V)", fontTitle, Brushes.Black, new PointF(yEndPoint.X + 10, yEndPoint.Y));
                //画Y轴坐标,每个区间的实际坐标长度
                float intervalY = Math.Abs(yEndPoint.Y - yOriginalPoint.Y) / ySectionCount;
                m_CoordinateIntervalY = intervalY;
                for (int i = 0; i < ySectionCount; i++)
                {
                    m_gh2.DrawLine(Pens.Black, new PointF(yOriginalPoint.X, yOriginalPoint.Y - intervalY * i), new PointF(yOriginalPoint.X + lineSegmentHeight, yOriginalPoint.Y - intervalY * i));
                }
                float yValueInerval = (float)yMax / ySectionCount;
                m_ValueInervalY = yValueInerval;//Y轴上的坐标值，根据实际放大倍数和量程决定
                for (int i = 0; i <= ySectionCount; i++)
                {
                    m_gh2.DrawString(i.ToString(), xValuefont, Brushes.Black, new PointF(yOriginalPoint.X - 24, yOriginalPoint.Y - intervalY * i - 6));
                }
                //画legend
                m_gh2.DrawString(VOL, fontTitle, Brushes.Black, new PointF(xEndPoint.X - 40, 10));
                SizeF fontSize = m_gh2.MeasureString(VOL, fontTitle);

                m_gh2.DrawLine(new Pen(Color.Red), new PointF(xEndPoint.X - 60, 10 + fontSize.Height / 2), new PointF(xEndPoint.X - 40, 10 + fontSize.Height / 2));
            }
            catch (Exception e)
            {
                MessageBox.Show("DrawHomeostasisMap Error:" + e.Message);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="xSectionCount">X轴的坐标数量</param>
        /// <param name="ySectionCount">Y轴坐标数量</param>
        private void DrawSingleAccuracyMap(int xSectionCount = 10, int ySectionCount = 5)
        {
            if (m_Ch1SampleDataList.Count <= 1)
                return;
            int interval = 1;// (int)m_Timer.Interval / 1000;
            Rectangle rect = m_Rect;
            Font xValuefont = new Font("宋体", 7);
            Font fontTitle = new Font("宋体", 8);
            //X轴原点
            PointF xOriginalPoint = new PointF((float)rect.Left + LEFTBORDEROFFSET, rect.Bottom - BOTTOMBORDEROFFSET);
            //X轴终点
            PointF xEndPoint = new PointF((float)rect.Right - RIGHTBORDEROFFSET, rect.Bottom - BOTTOMBORDEROFFSET);
            //Y轴最下面的点位置
            PointF yOriginalPoint = xOriginalPoint;
            //Y轴终点（由下向上）
            PointF yEndPoint = new PointF((float)rect.Left + LEFTBORDEROFFSET, (float)rect.Top + TOPBOTTOMFFSET);
            float y0 = 0, y1 = 0, x0 = 0, x1 = 0;
            int i = m_Ch1SampleDataList.Count - 1;
            y0 = xOriginalPoint.Y - ((yOriginalPoint.Y - yEndPoint.Y) / ySectionCount * ((m_Ch1SampleDataList[i - 1].m_PressureValue / m_ValueInervalY)));
            y1 = xOriginalPoint.Y - ((yOriginalPoint.Y - yEndPoint.Y) / ySectionCount * ((m_Ch1SampleDataList[i].m_PressureValue / m_ValueInervalY)));
            x0 = (xEndPoint.X - xOriginalPoint.X) / xSectionCount * m_Ch1SampleDataList[i - 1].m_Weight + xOriginalPoint.X;
            x1 = (xEndPoint.X - xOriginalPoint.X) / xSectionCount * m_Ch1SampleDataList[i].m_Weight + xOriginalPoint.X;
            m_gh.DrawLine(Pens.Red, new PointF(x0, y0), new PointF(x1, y1));
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="xSectionCount">X轴的坐标数量</param>
        /// <param name="ySectionCount">Y轴坐标数量</param>
        private void DrawSingleAccuracyMap2(int xSectionCount = 10, int ySectionCount = 5)
        {
            if (m_Ch2SampleDataList.Count <= 1)
                return;
            //int interval = 1;// (int)m_Timer.Interval / 1000;
            Rectangle rect = m_Rect;
            Font xValuefont = new Font("宋体", 7);
            Font fontTitle = new Font("宋体", 8);
            //X轴原点
            PointF xOriginalPoint = new PointF((float)rect.Left + LEFTBORDEROFFSET, rect.Bottom - BOTTOMBORDEROFFSET);
            //X轴终点
            PointF xEndPoint = new PointF((float)rect.Right - RIGHTBORDEROFFSET, rect.Bottom - BOTTOMBORDEROFFSET);
            //Y轴最下面的点位置
            PointF yOriginalPoint = xOriginalPoint;
            //Y轴终点（由下向上）
            PointF yEndPoint = new PointF((float)rect.Left + LEFTBORDEROFFSET, (float)rect.Top + TOPBOTTOMFFSET);
            float y0 = 0, y1 = 0, x0 = 0, x1 = 0;
            int i = m_Ch2SampleDataList.Count - 1;
            y0 = xOriginalPoint.Y - ((yOriginalPoint.Y - yEndPoint.Y) / ySectionCount * ((m_Ch2SampleDataList[i - 1].m_PressureValue / m_ValueInervalY)));
            y1 = xOriginalPoint.Y - ((yOriginalPoint.Y - yEndPoint.Y) / ySectionCount * ((m_Ch2SampleDataList[i].m_PressureValue / m_ValueInervalY)));
            x0 = (xEndPoint.X - xOriginalPoint.X) / xSectionCount * m_Ch2SampleDataList[i - 1].m_Weight + xOriginalPoint.X;
            x1 = (xEndPoint.X - xOriginalPoint.X) / xSectionCount * m_Ch2SampleDataList[i].m_Weight + xOriginalPoint.X;
            m_gh2.DrawLine(Pens.Red, new PointF(x0, y0), new PointF(x1, y1));
        }

        /// <summary>
        /// 界面移动或变化时需要重绘所有点
        /// </summary>
        /// <param name="xSectionCount"></param>
        /// <param name="ySectionCount"></param>
        private void DrawAccuracyMap(int xSectionCount = 10, int ySectionCount = 5)
        {
            if (m_Ch1SampleDataList.Count <= 1)
                return;
            int interval = 1;// (int)m_Timer.Interval / 1000;
            Rectangle rect = m_Rect;
            Font xValuefont = new Font("宋体", 7);
            Font fontTitle = new Font("宋体", 8);
            //画X轴
            //X轴原点
            PointF xOriginalPoint = new PointF((float)rect.Left + LEFTBORDEROFFSET, rect.Bottom - BOTTOMBORDEROFFSET);
            //X轴终点
            PointF xEndPoint = new PointF((float)rect.Right - RIGHTBORDEROFFSET, rect.Bottom - BOTTOMBORDEROFFSET);
            //Y轴最下面的点位置
            PointF yOriginalPoint = xOriginalPoint;
            //Y轴终点（由下向上）
            PointF yEndPoint = new PointF((float)rect.Left + LEFTBORDEROFFSET, (float)rect.Top + TOPBOTTOMFFSET);
            string strMsg = string.Empty;
            float y0 = 0, y1 = 0, x0 = 0, x1 = 0;
            int count = m_Ch1SampleDataList.Count;
            for (int iLoop = 1; iLoop < count; iLoop++)
            {
                y0 = xOriginalPoint.Y - ((yOriginalPoint.Y - yEndPoint.Y) / ySectionCount * ((m_Ch1SampleDataList[iLoop - 1].m_PressureValue / m_ValueInervalY)));
                y1 = xOriginalPoint.Y - ((yOriginalPoint.Y - yEndPoint.Y) / ySectionCount * ((m_Ch1SampleDataList[iLoop].m_PressureValue / m_ValueInervalY)));
                x0 = (xEndPoint.X - xOriginalPoint.X) / xSectionCount * m_Ch1SampleDataList[iLoop - 1].m_Weight + xOriginalPoint.X;
                x1 = (xEndPoint.X - xOriginalPoint.X) / xSectionCount * m_Ch1SampleDataList[iLoop].m_Weight + xOriginalPoint.X;
                m_gh.DrawLine(Pens.Red, new PointF(x0, y0), new PointF(x1, y1));
            }
        }

        /// <summary>
        /// 界面移动或变化时需要重绘所有点
        /// </summary>
        /// <param name="xSectionCount"></param>
        /// <param name="ySectionCount"></param>
        private void DrawAccuracyMap2(int xSectionCount = 10, int ySectionCount = 5)
        {
            if (m_Ch2SampleDataList.Count <= 1)
                return;
            //int interval = 1;// (int)m_Timer.Interval / 1000;
            Rectangle rect = m_Rect;
            Font xValuefont = new Font("宋体", 7);
            Font fontTitle = new Font("宋体", 8);
            //画X轴
            //X轴原点
            PointF xOriginalPoint = new PointF((float)rect.Left + LEFTBORDEROFFSET, rect.Bottom - BOTTOMBORDEROFFSET);
            //X轴终点
            PointF xEndPoint = new PointF((float)rect.Right - RIGHTBORDEROFFSET, rect.Bottom - BOTTOMBORDEROFFSET);
            //Y轴最下面的点位置
            PointF yOriginalPoint = xOriginalPoint;
            //Y轴终点（由下向上）
            PointF yEndPoint = new PointF((float)rect.Left + LEFTBORDEROFFSET, (float)rect.Top + TOPBOTTOMFFSET);
            string strMsg = string.Empty;
            float y0 = 0, y1 = 0, x0 = 0, x1 = 0;
            int count = m_Ch2SampleDataList.Count;
            for (int iLoop = 1; iLoop < count; iLoop++)
            {
                y0 = xOriginalPoint.Y - ((yOriginalPoint.Y - yEndPoint.Y) / ySectionCount * ((m_Ch2SampleDataList[iLoop - 1].m_PressureValue / m_ValueInervalY)));
                y1 = xOriginalPoint.Y - ((yOriginalPoint.Y - yEndPoint.Y) / ySectionCount * ((m_Ch2SampleDataList[iLoop].m_PressureValue / m_ValueInervalY)));
                x0 = (xEndPoint.X - xOriginalPoint.X) / xSectionCount * m_Ch2SampleDataList[iLoop - 1].m_Weight + xOriginalPoint.X;
                x1 = (xEndPoint.X - xOriginalPoint.X) / xSectionCount * m_Ch2SampleDataList[iLoop].m_Weight + xOriginalPoint.X;
                m_gh2.DrawLine(Pens.Red, new PointF(x0, y0), new PointF(x1, y1));
            }
        }
 
        #endregion

        #region 控件事件

        /// <summary>
        /// 泵类型选择
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cbPumpType_SelectedIndexChanged(object sender, EventArgs e)
        {
            m_LocalPid = (ProductID)Enum.Parse(typeof(ProductID), cbPumpType.Items[cbPumpType.SelectedIndex].ToString(), true);
            if(m_LocalPid== ProductID.GrasebyF6 || m_LocalPid == ProductID.WZS50F6)
            {
                gbChannel2.Enabled = true;
            }
            else
            {
                gbChannel2.Enabled = false;
            }
        }

        /// <summary>
        /// 1通道工装串口号选择
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cbToolingPort_SelectedIndexChanged(object sender, EventArgs e)
        {
            lbWeightChannel1.Text = "-----";
            if (m_DetectPToolCh1 == null)
                m_DetectPToolCh1 = new PTooling();
            if (m_DetectPToolCh1.IsOpen())
                m_DetectPToolCh1.Close();
            m_DetectPToolCh1.Init(cbToolingPort.Items[cbToolingPort.SelectedIndex].ToString());
            m_DetectPToolCh1.Open();
            Thread.Sleep(500);
            m_DetectPToolCh1.ReadWeight();
        }

        private void cbToolingPort2_SelectedIndexChanged(object sender, EventArgs e)
        {
            lbWeightChannel2.Text = "-----";
            if (m_DetectPToolCh2 == null)
                m_DetectPToolCh2 = new PTooling();
            if (m_DetectPToolCh2.IsOpen())
                m_DetectPToolCh2.Close();
            m_DetectPToolCh2.Init(cbToolingPort2.Items[cbToolingPort2.SelectedIndex].ToString());
            m_DetectPToolCh2.Open();
            m_DetectPToolCh2.ReadWeight();
        }

        /// <summary>
        /// 1通道泵串口号选择
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cbPumpPort_SelectedIndexChanged(object sender, EventArgs e)
        {
            lbPChannel1.Text = "-----";
            if (m_GrasebyDevice == null)
                m_GrasebyDevice = new Graseby9600();
            if (m_GrasebyDevice.IsOpen())
                m_GrasebyDevice.Close();

            switch (m_LocalPid)
            {
                case ProductID.GrasebyC6:
                    m_GrasebyDevice.SetDeviceType(DeviceType.GrasebyC6);
                    break;
                case ProductID.GrasebyF6:
                    m_GrasebyDevice.SetDeviceType(DeviceType.GrasebyF6);
                    break;
                case ProductID.GrasebyC6T:
                    m_GrasebyDevice.SetDeviceType(DeviceType.GrasebyC6T);
                    break;
                case ProductID.Graseby2000:
                    m_GrasebyDevice.SetDeviceType(DeviceType.Graseby2000);
                    break;
                case ProductID.Graseby2100:
                    m_GrasebyDevice.SetDeviceType(DeviceType.Graseby2100);
                    break;
                case ProductID.WZ50C6:
                    m_GrasebyDevice.SetDeviceType(DeviceType.WZ50C6);
                    break;
                case ProductID.WZS50F6:
                    m_GrasebyDevice.SetDeviceType(DeviceType.WZS50F6);
                    break;
                case ProductID.WZ50C6T:
                    m_GrasebyDevice.SetDeviceType(DeviceType.WZ50C6T);
                    break;
            }
            m_GrasebyDevice.Init(cbPumpPort.Items[cbPumpPort.SelectedIndex].ToString());
            m_GrasebyDevice.Open();
            m_GrasebyDevice.Get();
        }

        private void cbPumpPort2_SelectedIndexChanged(object sender, EventArgs e)
        {
            lbPChannel2.Text = "-----";
            if (m_GrasebyDevice2 == null)
                m_GrasebyDevice2 = new Graseby9600();
            if (m_GrasebyDevice2.IsOpen())
                m_GrasebyDevice2.Close();

            switch (m_LocalPid)
            {
                case ProductID.GrasebyC6:
                    m_GrasebyDevice2.SetDeviceType(DeviceType.GrasebyC6);
                    break;
                case ProductID.GrasebyF6:
                    m_GrasebyDevice2.SetDeviceType(DeviceType.GrasebyF6);
                    break;
                case ProductID.GrasebyC6T:
                    m_GrasebyDevice2.SetDeviceType(DeviceType.GrasebyC6T);
                    break;
                case ProductID.Graseby2000:
                    m_GrasebyDevice2.SetDeviceType(DeviceType.Graseby2000);
                    break;
                case ProductID.Graseby2100:
                    m_GrasebyDevice2.SetDeviceType(DeviceType.Graseby2100);
                    break;
                case ProductID.WZ50C6:
                    m_GrasebyDevice2.SetDeviceType(DeviceType.WZ50C6);
                    break;
                case ProductID.WZS50F6:
                    m_GrasebyDevice2.SetDeviceType(DeviceType.WZS50F6);
                    break;
                case ProductID.WZ50C6T:
                    m_GrasebyDevice2.SetDeviceType(DeviceType.WZ50C6T);
                    break;
            }
            m_GrasebyDevice2.Init(cbPumpPort2.Items[cbPumpPort2.SelectedIndex].ToString());
            m_GrasebyDevice2.Open();
            m_GrasebyDevice2.Get();
        }
            

        /// <summary>
        /// 开始调试1通道
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnStartPumpChannel1_Click(object sender, EventArgs e)
        {
            m_Ch1SampleDataList.Clear();
            WavelinePanel.Invalidate();

            #region 参数输入检查
            if (cbToolingPort.SelectedIndex < 0)
            {
                MessageBox.Show("请选择工装串口");
                return;
            }
            float weight = 0;
            if (!float.TryParse(lbWeightChannel1.Text, out weight))
            {
                MessageBox.Show("工装串口连接错误，请正确选择串口！");
                return;
            }

            if (cbPumpPort.SelectedIndex < 0)
            {
                MessageBox.Show("请选择泵串口");
                return;
            }
            if (!float.TryParse(lbPChannel1.Text, out weight))
            {
                MessageBox.Show("泵串口连接错误，请正确选择串口！");
                return;
            }
            if (string.IsNullOrEmpty(tbRateChannel1.Text))
            {
                MessageBox.Show("请输入速率！");
                return;
            }
            float rate = 0;
            if (!float.TryParse(tbRateChannel1.Text, out rate))
            {
                MessageBox.Show("请正确输入速率！");
                return;
            }
            #endregion

            #region 泵型号选择
            Misc.ProductID pid = Misc.ProductID.None;
            switch (m_LocalPid)
            {
                case ProductID.WZ50C6:
                    pid = Misc.ProductID.GrasebyC6;
                    break;
                case ProductID.WZ50C6T:
                    pid = Misc.ProductID.GrasebyC6T;
                    break;
                case ProductID.WZS50F6:
                    pid = Misc.ProductID.GrasebyF6;
                    break;
                case ProductID.Graseby2000:
                    pid = Misc.ProductID.Graseby2000;
                    break;
                case ProductID.Graseby2100:
                    pid = Misc.ProductID.Graseby2100;
                    break;
                case ProductID.GrasebyC6:
                    pid = Misc.ProductID.GrasebyC6;
                    break;
                case ProductID.GrasebyF6:
                    pid = Misc.ProductID.GrasebyF6;
                    break;
                default:
                    pid = Misc.ProductID.None;
                    break;
            }
            #endregion

            if (pid == Misc.ProductID.None)
            {
                MessageBox.Show("选择的泵类型错误，请联系管理员!");
                return;
            }

            if (m_ConnResponseCh1 == null)
                m_ConnResponseCh1 = new GlobalResponse(pid, Misc.CommunicationProtocolType.General);
            if (m_ConnResponseCh1.IsOpen())
            {
                m_ConnResponseCh1.CloseConnection();
            }
            m_ConnResponseCh1.Initialize(cbPumpPort.Items[cbPumpPort.SelectedIndex].ToString(), BAUDRATE);
            RemoveHandler();
            AddHandler();
            if (m_PToolCh1 != null)
            {
                if (m_PToolCh1.IsOpen())
                {

                }
                else
                {
                    m_PToolCh1.Init(cbToolingPort.Items[cbToolingPort.SelectedIndex].ToString());
                    m_PToolCh1.Open();
                }
            }
            else
            {
                m_PToolCh1 = new PTooling();
                m_PToolCh1.Init(cbToolingPort.Items[cbToolingPort.SelectedIndex].ToString());
                m_PToolCh1.Open();
            }
            m_PToolCh1.Tare();
            Thread.Sleep(500);
            m_ConnResponseCh1.SetVTBIParameter(0, rate);
            StartCh1Timer();
            EnableContols(1, false);
        }

        private void btnStartPumpChannel2_Click(object sender, EventArgs e)
        {
            m_Ch2SampleDataList.Clear();
            WavelinePanel2.Invalidate();

            #region 参数输入检查
            if (cbToolingPort2.SelectedIndex < 0)
            {
                MessageBox.Show("请选择工装串口");
                return;
            }
            float weight = 0;
            if (!float.TryParse(lbWeightChannel2.Text, out weight))
            {
                MessageBox.Show("工装串口连接错误，请正确选择串口！");
                return;
            }

            if (cbPumpPort2.SelectedIndex < 0)
            {
                MessageBox.Show("请选择泵串口");
                return;
            }
            if (!float.TryParse(lbPChannel2.Text, out weight))
            {
                MessageBox.Show("泵串口连接错误，请正确选择串口！");
                return;
            }
            if (string.IsNullOrEmpty(tbRateChannel2.Text))
            {
                MessageBox.Show("请输入速率！");
                return;
            }
            float rate = 0;
            if (!float.TryParse(tbRateChannel2.Text, out rate))
            {
                MessageBox.Show("请正确输入速率！");
                return;
            }
            #endregion

            #region 泵型号选择
            Misc.ProductID pid = Misc.ProductID.None;
            switch (m_LocalPid)
            {
                case ProductID.WZ50C6:
                    pid = Misc.ProductID.GrasebyC6;
                    break;
                case ProductID.WZ50C6T:
                    pid = Misc.ProductID.GrasebyC6T;
                    break;
                case ProductID.WZS50F6:
                    pid = Misc.ProductID.GrasebyF6;
                    break;
                case ProductID.Graseby2000:
                    pid = Misc.ProductID.Graseby2000;
                    break;
                case ProductID.Graseby2100:
                    pid = Misc.ProductID.Graseby2100;
                    break;
                case ProductID.GrasebyC6:
                    pid = Misc.ProductID.GrasebyC6;
                    break;
                case ProductID.GrasebyF6:
                    pid = Misc.ProductID.GrasebyF6;
                    break;
                default:
                    pid = Misc.ProductID.None;
                    break;
            }
            #endregion

            if (pid == Misc.ProductID.None)
            {
                MessageBox.Show("选择的泵类型错误，请联系管理员!");
                return;
            }

            if (m_ConnResponseCh2 == null)
                m_ConnResponseCh2 = new GlobalResponse(pid, Misc.CommunicationProtocolType.General);
            if (m_ConnResponseCh2.IsOpen())
            {
                m_ConnResponseCh2.CloseConnection();
            }
            m_ConnResponseCh2.Initialize(cbPumpPort2.Items[cbPumpPort2.SelectedIndex].ToString(), BAUDRATE);
            RemoveHandler(2);
            AddHandler(2);
            if (m_PToolCh2 != null)
            {
                if (m_PToolCh2.IsOpen())
                {

                }
                else
                {
                    m_PToolCh2.Init(cbToolingPort2.Items[cbToolingPort2.SelectedIndex].ToString());
                    m_PToolCh2.Open();
                }
            }
            else
            {
                m_PToolCh2 = new PTooling();
                m_PToolCh2.Init(cbToolingPort2.Items[cbToolingPort2.SelectedIndex].ToString());
                m_PToolCh2.Open();
            }
            m_PToolCh2.Tare();
            Thread.Sleep(500);
            m_ConnResponseCh2.SetVTBIParameter(0, rate);
            StartCh2Timer();
            EnableContols(2, false);
        }

        /// <summary>
        /// 停止调试1通道，并关闭串口
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnStopPumpChannel1_Click(object sender, EventArgs e)
        {
            Complete();
            EnableContols(1, true);
        }

        /// <summary>
        /// 停止调试1通道，并关闭串口
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnStopPumpChannel2_Click(object sender, EventArgs e)
        {
            Complete(2);
            EnableContols(2, true);
        }

        #endregion

        #region 单通道命令响应

        /// <summary>
        /// this function is invoked by GlobalResponse class event
        /// when m_ConnResponse.SetOcclusionLevel() is called; 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        private void SetInfusionParas(object sender, ResponseEventArgs<String> args)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new EventHandler<ResponseEventArgs<String>>(SetInfusionParas), new object[] { sender, args });
                return;
            }
            if (string.IsNullOrEmpty(args.ErrorMessage))
            {
                if (m_ConnResponseCh1 != null && m_ConnResponseCh1.IsOpen())
                    m_ConnResponseCh1.SetStartControl();
                else
                    MessageBox.Show("泵端串口不可用，请检查串口是否已连接!");
            }
            else
            {
                MessageBox.Show(args.ErrorMessage);
            }
        }

        private void SetStartControl(object sender, ResponseEventArgs<String> args)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new EventHandler<ResponseEventArgs<String>>(SetStartControl), new object[] { sender, args });
                return;
            }
            if (String.Empty != args.ErrorMessage)
            {
                MessageBox.Show(args.ErrorMessage);
            }
        }

        /// <summary>
        /// Invoked by GlobalResponse class event, when m_ConnResponse.SetStopControl() is called; 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args">ErrorMessage or Empty</param>
        private void SetStopControl(object sender, ResponseEventArgs<String> args)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new EventHandler<ResponseEventArgs<String>>(SetStopControl), new object[] { sender, args });
                return;
            }
            if (String.Empty != args.ErrorMessage)
            {
                MessageBox.Show("停止泵失败，请手动操作停止！");
            }
        }

        private void GetPressureSensor(object sender, ResponseEventArgs<Misc.PressureSensorInfo> args)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new EventHandler<ResponseEventArgs<Misc.PressureSensorInfo>>(GetPressureSensor), new object[] { sender, args });
                return;
            }
            if (String.Empty != args.ErrorMessage)
            {
                lbPChannel1.Text = "";
                Complete();
                EnableContols(1, true);
                MessageBox.Show("读取压力值失败，串口连接失败！");
            }
            else
            {
                Misc.PressureSensorInfo paras = args.EventData;
                lbPChannel1.Text = paras.pressureVoltage.ToString("F2");
                lock (m_Ch1SampleDataList)
                {
                    m_Ch1SampleDataList.Add(new SampleData(DateTime.Now, paras.pressureVoltage, -1000f));
                }
                if (m_PToolCh1 != null && m_PToolCh1.IsOpen())
                {
                    m_PToolCh1.ReadWeight();
                }
                else
                {
                    Complete();
                    EnableContols(1, true);
                    MessageBox.Show("工装串口关闭，请检查设备");
                }
            }
        }

        public void SetPressureCalibrationParameter(object sender, ResponseEventArgs<String> args)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new EventHandler<ResponseEventArgs<string>>(SetPressureCalibrationParameter), new object[] { sender, args });
                return;
            }
            
            if (String.Empty != args.ErrorMessage)
            {
            }
            else
            {
             
            }
        }

        #endregion

        #region 双通道命令响应

        /// <summary>
        /// this function is invoked by GlobalResponse class event
        /// when m_ConnResponse.SetOcclusionLevel() is called; 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        private void SetInfusionParas2(object sender, ResponseEventArgs<String> args)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new EventHandler<ResponseEventArgs<String>>(SetInfusionParas2), new object[] { sender, args });
                return;
            }
            if (string.IsNullOrEmpty(args.ErrorMessage))
            {
                if (m_ConnResponseCh2 != null && m_ConnResponseCh2.IsOpen())
                    m_ConnResponseCh2.SetStartControl();
                else
                    MessageBox.Show("泵端串口不可用，请检查串口是否已连接!");
            }
            else
            {
                MessageBox.Show(args.ErrorMessage);
            }
        }

        private void SetStartControl2(object sender, ResponseEventArgs<String> args)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new EventHandler<ResponseEventArgs<String>>(SetStartControl2), new object[] { sender, args });
                return;
            }
            if (String.Empty != args.ErrorMessage)
            {
                MessageBox.Show(args.ErrorMessage);
            }
        }

        /// <summary>
        /// Invoked by GlobalResponse class event, when m_ConnResponse.SetStopControl() is called; 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args">ErrorMessage or Empty</param>
        private void SetStopControl2(object sender, ResponseEventArgs<String> args)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new EventHandler<ResponseEventArgs<String>>(SetStopControl2), new object[] { sender, args });
                return;
            }
            if (String.Empty != args.ErrorMessage)
            {
                MessageBox.Show("停止泵失败，请手动操作停止！");
            }
        }

        private void GetPressureSensor2(object sender, ResponseEventArgs<Misc.PressureSensorInfo> args)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new EventHandler<ResponseEventArgs<Misc.PressureSensorInfo>>(GetPressureSensor2), new object[] { sender, args });
                return;
            }
            if (String.Empty != args.ErrorMessage)
            {
                lbPChannel2.Text = "";
                Complete(2);
                EnableContols(2, true);
                MessageBox.Show("读取压力值失败，串口连接失败！");
            }
            else
            {
                Misc.PressureSensorInfo paras = args.EventData;
                lbPChannel2.Text = paras.pressureVoltage.ToString("F2");
                lock (m_Ch2SampleDataList)
                {
                    m_Ch2SampleDataList.Add(new SampleData(DateTime.Now, paras.pressureVoltage, -1000f));
                }
                if (m_PToolCh2 != null && m_PToolCh2.IsOpen())
                {
                    m_PToolCh2.ReadWeight();
                }
                else
                {
                    Complete(2);
                    EnableContols(2, true);
                    MessageBox.Show("工装串口关闭，请检查设备");
                }
            }
        }

        public void SetPressureCalibrationParameter2(object sender, ResponseEventArgs<String> args)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new EventHandler<ResponseEventArgs<string>>(SetPressureCalibrationParameter2), new object[] { sender, args });
                return;
            }

            if (String.Empty != args.ErrorMessage)
            {
            }
            else
            {

            }
        }

        #endregion

        private void OnRateKeyPress(object sender, KeyPressEventArgs e)
        {
            TextBox txt = sender as TextBox;
            if (char.IsNumber(e.KeyChar) || e.KeyChar == (char)Keys.Back)
            {
                e.Handled = false;                         //让操作生效
                if (txt.Text.Length == 0)
                {
                    if (e.KeyChar == '0')
                        e.Handled = true;                  //让操作失效，第一个字符不能输入0
                }
                else if (txt.Text.Length >= 3)
                {

                    if (e.KeyChar == (char)Keys.Back)
                        e.Handled = false;             //让操作生效
                    else
                        e.Handled = true;              //让操作失效，如果第一个字符是2以上，不能输入其他字符
                }
                else
                {
                    e.Handled = false;                 //让操作生效
                }
            }
            else
            {
                e.Handled = true;
            }
        }

        #region 菜单操作
        private void menuInitTooling_Click(object sender, EventArgs e)
        {
            if (m_PToolCh1 != null && m_PToolCh1.IsOpen())
            {
                m_PToolCh1.SetChannel();
                Thread.Sleep(500);
                //m_PToolCh1.SetScale();
                //Thread.Sleep(500);
                m_PToolCh1.SetDecimalPlace();
                Thread.Sleep(500);
                m_PToolCh1.Tare();
            }
        }

        private void menuTareTooling_Click(object sender, EventArgs e)
        {
            if (m_PToolCh1 != null && m_PToolCh1.IsOpen())
            {
                m_PToolCh1.Tare();
            }
        }
        #endregion

        /// <summary>
        /// 调试结束，保存相关数据，停止泵，停止时钟
        /// </summary>
        private void Complete(int channel = 1)
        {
            if (channel == 1)
            {
                StopCh1Timer();
                if (m_ConnResponseCh1 != null)
                {
                    if (m_ConnResponseCh1.IsOpen())
                    {
                        m_ConnResponseCh1.SetStopControl(GlobalResponse.CommandPriority.High);
                        RemoveHandler();
                        Thread.Sleep(500);
                        CalcuatePressure(channel, m_LocalPid, m_Ch1SampleDataList);
                        Thread.Sleep(500);
                        m_ConnResponseCh1.CloseConnection();
                    }
                }
                if (m_PToolCh1 != null && m_PToolCh1.IsOpen())
                    m_PToolCh1.Close();

                string fileName = tbPumpNo.Text;
                if (m_LocalPid == ProductID.GrasebyF6 || m_LocalPid == ProductID.WZS50F6)
                    fileName = string.Format("{0}{1}道{2}{3}", m_LocalPid.ToString(), channel, tbPumpNo.Text, DateTime.Now.ToString("yyyy-MM-dd HH_mm_ss"));
                else
                    fileName = string.Format("{0}{1}{2}", m_LocalPid.ToString(), tbPumpNo.Text, DateTime.Now.ToString("yyyy-MM-dd HH_mm_ss"));
                string path = Path.GetDirectoryName(Assembly.GetAssembly(typeof(MainForm)).Location) + "\\压力调试数据";
                if (!System.IO.Directory.Exists(path))
                    System.IO.Directory.CreateDirectory(path);
                string saveFileName = path + "\\" + fileName + ".xlsx";
                Export(1, saveFileName);
                EnableContols(1, true);
            }
            else
            {
                StopCh2Timer();
                if (m_ConnResponseCh2 != null)
                {
                    if (m_ConnResponseCh2.IsOpen())
                    {
                        m_ConnResponseCh2.SetStopControl(GlobalResponse.CommandPriority.High);
                        Thread.Sleep(500);
                        CalcuatePressure(channel, m_LocalPid, m_Ch2SampleDataList);
                        Thread.Sleep(500);
                        m_ConnResponseCh2.CloseConnection();
                    }
                }
                if (m_PToolCh2 != null && m_PToolCh2.IsOpen())
                    m_PToolCh2.Close();

                string fileName = tbPumpNo.Text;
                if (m_LocalPid == ProductID.GrasebyF6 || m_LocalPid == ProductID.WZS50F6)
                    fileName = string.Format("{0}{1}道{2}{3}", m_LocalPid.ToString(), channel, tbPumpNo.Text, DateTime.Now.ToString("yyyy-MM-dd HH_mm_ss"));
                else
                    fileName = string.Format("{0}{1}{2}", m_LocalPid.ToString(), tbPumpNo.Text, DateTime.Now.ToString("yyyy-MM-dd HH_mm_ss"));
                string path = Path.GetDirectoryName(Assembly.GetAssembly(typeof(MainForm)).Location) + "\\压力调试数据";
                if (!System.IO.Directory.Exists(path))
                    System.IO.Directory.CreateDirectory(path);
                string saveFileName = path + "\\" + fileName + ".xlsx";
                Export(channel, saveFileName);
                EnableContols(2, true);
            }
        }

        private void EnableContols(int channel, bool bEnabled = true)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new DelegateEnableContols(EnableContols), new object[] { channel, bEnabled });
                return;
            }
            if (channel == 1)
            {
                cbPumpType.Enabled = bEnabled;
                cbToolingPort.Enabled = bEnabled;
                cbPumpPort.Enabled = bEnabled;
                tbRateChannel1.Enabled = bEnabled;
                btnStartPumpChannel1.Enabled = bEnabled;
                btnStopPumpChannel1.Enabled = !bEnabled;
            }
            else if (channel == 2)
            {
                cbPumpType.Enabled = bEnabled;
                cbToolingPort2.Enabled = bEnabled;
                cbPumpPort2.Enabled = bEnabled;
                tbRateChannel2.Enabled = bEnabled;
                btnStartPumpChannel2.Enabled = bEnabled;
                btnStopPumpChannel2.Enabled = !bEnabled;
            }
        }

        private void Export(int channel, string name)
        {
            List<SampleData> sampleDataList = null;
            if (channel == 1)
                sampleDataList = m_Ch1SampleDataList;
            else
                sampleDataList = m_Ch2SampleDataList;
            if (sampleDataList == null || sampleDataList.Count == 0)
                return;
            string title = string.Empty;
            if (m_LocalPid == ProductID.GrasebyF6 || m_LocalPid == ProductID.WZS50F6)
            {
                title = string.Format("泵型号:{0}{1}道 产品序号:{2} 工装编号:{3}", m_LocalPid.ToString(), channel, tbPumpNo.Text, tbToolingNo.Text);
            }
            else
            {
                title = string.Format("泵型号：{0}", m_LocalPid.ToString());
            }
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("压力调试数据");

            ws.Cell(1, 1).Value = title;
            ws.Cell(1, 1).Style.Font.Bold = true;
            ws.Range(1, 1, 1, 3).Merge();
            ws.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell(1, 1).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            ws.Cell(2, 1).Value = "采样时间";
            ws.Cell(2, 2).Value = "重量(kg)";
            ws.Cell(2, 3).Value = "压力(V)";

            int count = sampleDataList.Count;
            int index = 3;
            for (int i = 0; i < count; i++)
            {
                ws.Cell(index, 1).Value = sampleDataList[i].m_SampleTime.ToString("yyyy-MM-dd HH_mm_ss");
                ws.Cell(index, 2).Value = sampleDataList[i].m_Weight;
                ws.Cell(index, 3).Value = sampleDataList[i].m_PressureValue*100;
                index++;
            }
            wb.SaveAs(name);
        }

        private void CalcuatePressure(int channel, ProductID pid, List<SampleData> sampleDataList)
        {
            if (sampleDataList == null || sampleDataList.Count == 0)
                return;
            List<PressureParameter> parameters = new List<PressureParameter>();
            ProductPressure pp = PressureManager.Instance().GetPressureByProductID(pid);
            if (pp == null)
                return;
            List<LevelPressure> lps = pp.GetLevelPressureList();
            List<float> midWeights = new List<float>();
            List<int> sizes = new List<int>();
            if (pid == ProductID.WZS50F6 || pid == ProductID.GrasebyF6)
                sizes.Add(10);
            sizes.Add(20);
            sizes.Add(30);
            sizes.Add(50);

            foreach (var size in sizes)
            {
                foreach (Misc.OcclusionLevel level in Enum.GetValues(typeof(Misc.OcclusionLevel)))
                {
                    LevelPressure lp = lps.Find((x) => { return x.m_Level == level; });
                    if (lp != null)
                    {
                        SizePressure sp = lp.Find(size);
                        if (sp != null)
                        {
                            PressureParameter para = new PressureParameter(size, level, sp.m_Mid, 0);
                            parameters.Add(para);
                        }
                    }
                }
            }
            //找到相关的值后，需要写入到泵中
            FindNearestPValue(ref parameters, sampleDataList);
            List<PressureCalibrationParameter> caliParameters = new List<PressureCalibrationParameter>();
            foreach (var size in sizes)
            {
                PressureCalibrationParameter p = new PressureCalibrationParameter();
                p.m_SyringeSize = size;
                List<PressureParameter> findobjs = parameters.FindAll((x) => { return x.m_SyringeSize == size; });
                foreach(var obj in findobjs)
                {
                    switch(obj.m_Level)
                    {
                        case Misc.OcclusionLevel.L:
                            p.m_PressureL = obj.m_Pressure;
                            break;
                        case Misc.OcclusionLevel.C:
                            p.m_PressureC = obj.m_Pressure;
                            break;
                        case Misc.OcclusionLevel.H:
                            p.m_PressureH = obj.m_Pressure;
                            break;
                        default: break;
                    }
                }
                caliParameters.Add(p);
            }
            WritePressureCaliParameter2Pump(channel, caliParameters);
        }

        /// <summary>
        /// 给一组标准重量值(kg)，从采样的结果中查找与它最相近的值所对应的压力值（V）
        /// </summary>
        /// <param name="weight"></param>
        /// <param name="sampleDataList"></param>
        /// <returns></returns>
        private void FindNearestPValue(ref List<PressureParameter> parameters, List<SampleData> sampleDataList)
        {
            if (sampleDataList == null || sampleDataList.Count == 0)
                return;
            List<float> absList = new List<float>();
            List<int> indexs = new List<int>();
            for (int i = 0; i < parameters.Count; i++)
            {
                absList.Add(10000f);
                indexs.Add(0);
                for (int iLoop = 0; iLoop < sampleDataList.Count; iLoop++)
                {
                    float abs = Math.Abs(parameters[i].m_MidWeight - sampleDataList[iLoop].m_Weight);
                    if (absList[i] > abs)
                    {
                        absList[i] = abs;
                        indexs[i] = iLoop;
                        parameters[i].SetPressure(sampleDataList[iLoop].m_PressureValue);
                    }
                }
            }
        }

        private void WritePressureCaliParameter2Pump(int channel, List<PressureCalibrationParameter> caliParas)
        {
            if (caliParas == null || caliParas.Count == 0)
                return;
            for(int i=0;i<caliParas.Count;i++)
            {
                if (channel == 1)
                {
                    m_ConnResponseCh1.SetPressureCalibrationParameter((byte)(caliParas[i].m_SyringeSize),caliParas[i].m_PressureL, caliParas[i].m_PressureC,caliParas[i].m_PressureH);
                    Thread.Sleep(500);
                }
                else if(channel == 2)
                {
                    m_ConnResponseCh2.SetPressureCalibrationParameter((byte)(caliParas[i].m_SyringeSize), caliParas[i].m_PressureL, caliParas[i].m_PressureC, caliParas[i].m_PressureH);
                    Thread.Sleep(500);
                }
            }
        }

       

       
    }

   
}
