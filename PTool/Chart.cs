﻿using System;
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
    public partial class Chart : UserControl
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
        private System.Drawing.Rectangle m_Rect;
        private Pen m_WaveLinePen = new Pen(Color.FromArgb(19, 113, 185));
        private SolidBrush m_WaveLineBrush = new SolidBrush(Color.FromArgb(19, 113, 185));
        private float m_XCoordinateMaxValue = 10;
        private int m_YCoordinateMaxValue = 5;
        private int m_XSectionCount = 10;
        private int m_YSectionCount = 5;
        private float m_CoordinateIntervalX = 0;  //X轴上的区间实际长度，单位为像素
        private float m_CoordinateIntervalY = 0;  //Y轴上的区间实际长度，单位为像素
        private float m_ValueInervalX = 0;  //X轴上的坐标值，根据实际放大倍数和量程决定
        private float m_ValueInervalY = 0;
        private List<SampleData> m_Ch1SampleDataList = new List<SampleData>();

        protected GlobalResponse m_ConnResponse = null;
        private PTooling m_PTool = null;
        private PTooling m_DetectPTool = null;
        private Graseby9600 m_GrasebyDevice = new Graseby9600();//只用于串口刷新
        private PumpID m_LocalPid = PumpID.GrasebyC6;//默认显示的是C6
        private System.Timers.Timer m_Ch1Timer = new System.Timers.Timer();
        private bool m_bStopClick = false;//是否人工点击停止
        private int m_SampleInterval = 400;//采样频率：毫秒

        private int m_Channel = 1;//1号通道，默认值
        private string m_PumpNo = string.Empty;//产品序号
        private string m_ToolingNo = string.Empty;//工装编号

        public delegate void DelegateSetWeightValue(float weight, bool isDetect);
        public delegate void DelegateSetPValue(float p);
        public delegate void DelegateEnableContols(bool bEnabled);
        public delegate void DelegateAlertMessageWhenComplete(string msg);
        public delegate void DelegateInputOpratorNumber(string number);

        /// <summary>
        /// 当启动或停止时通知主界面
        /// </summary>
        public event EventHandler<EventArgs> SamplingStartOrStop;

        /// <summary>
        /// 当双道泵，测量结束后通知主界面，把数据传入
        /// </summary>
        public event EventHandler<DoublePumpDataArgs> OnSamplingComplete;
        public event EventHandler<OpratorNumberArgs> OpratorNumberInput;

        private int mCurrentSamplingIndex = 0;//当前采样点

        private System.Timers.Timer mSamplingPointTimer = new System.Timers.Timer();
        private bool mSamplingPointStart = false;//是否开始采样了
        private int mSamplingCount = 0;//采样了几次，三次即可
        private List<SampleData> mSamplingPointList = new List<SampleData>();//9个采样点的数据

        

        /// <summary>
        /// 采样间隔
        /// </summary>
        public int SampleInterval
        {
            get { return m_SampleInterval; }
            set { m_SampleInterval = value; }
        }

        /// <summary>
        /// 设置通道号1 or 2
        /// </summary>
        public int Channel
        {
            get { return m_Channel; }
            set
            {
                m_Channel = value;

            }
        }

        /// <summary>
        /// 产品序号
        /// </summary>
        public string PumpNo
        {
            get { return m_PumpNo; }
            set { m_PumpNo = value; }
        }

        /// <summary>
        /// 工装编号
        /// </summary>
        public string ToolingNo
        {
            get { return m_ToolingNo; }
            set { m_ToolingNo = value; }
        }

        public Chart()
        {
            InitializeComponent();
            m_Channel = 1;
            m_gh = WavelinePanel.CreateGraphics();
            m_Rect = WavelinePanel.ClientRectangle;
            m_PTool = new PTooling();
            m_PTool.DeviceDataRecerived += OnPTool_DeviceDataRecerived;
            m_DetectPTool = new PTooling();
            m_DetectPTool.DeviceDataRecerived += OnPTool_DetectDeviceDataRecerived;
            m_GrasebyDevice.DeviceDataRecerived += OnGrasebyDeviceDataRecerived;
        }

        public Chart(int channel = 1)
        {
            InitializeComponent();
            m_Channel = channel;
            m_gh = WavelinePanel.CreateGraphics();
            m_Rect = WavelinePanel.ClientRectangle;
            m_PTool = new PTooling();
            m_PTool.DeviceDataRecerived += OnPTool_DeviceDataRecerived;
            m_DetectPTool = new PTooling();
            m_DetectPTool.DeviceDataRecerived += OnPTool_DetectDeviceDataRecerived;
            m_GrasebyDevice.DeviceDataRecerived += OnGrasebyDeviceDataRecerived;
        }

        private void Chart_Load(object sender, EventArgs e)
        {
            cbPumpPort.Items.AddRange(SerialPort.GetPortNames());
            cbToolingPort.Items.AddRange(SerialPort.GetPortNames());
            m_Ch1Timer.Interval = m_SampleInterval;
            m_Ch1Timer.Elapsed += OnChannel1Timer_Elapsed;

            mSamplingPointTimer.Interval = 500;
            mSamplingPointTimer.Elapsed += OnSamplingPointTimer_Elapsed; 
        }


        public void AddHandler(int channel = 1)
        {
            if (m_ConnResponse != null)
            {
                m_ConnResponse.SetVTBIParameterResponse += new EventHandler<ResponseEventArgs<String>>(SetInfusionParas);
                m_ConnResponse.SetStartControlResponse += new EventHandler<ResponseEventArgs<String>>(SetStartControl);
                m_ConnResponse.SetStopControlResponse += new EventHandler<ResponseEventArgs<String>>(SetStopControl);
                m_ConnResponse.GetPressureSensorResponse += new EventHandler<ResponseEventArgs<Misc.PressureSensorInfo>>(GetPressureSensor);
                m_ConnResponse.SetPressureCalibrationParameterResponse += new EventHandler<ResponseEventArgs<String>>(SetPressureCalibrationParameter);
                m_ConnResponse.SetPressureCalibrationPValueResponse += new EventHandler<ResponseEventArgs<String>>(SetPressureCalibrationPValue);
            }
        }

        public void RemoveHandler(int channel = 1)
        {
            if (m_ConnResponse != null)
            {
                m_ConnResponse.SetVTBIParameterResponse -= new EventHandler<ResponseEventArgs<String>>(SetInfusionParas);
                m_ConnResponse.SetStartControlResponse -= new EventHandler<ResponseEventArgs<String>>(SetStartControl);
                m_ConnResponse.SetStopControlResponse -= new EventHandler<ResponseEventArgs<String>>(SetStopControl);
                m_ConnResponse.GetPressureSensorResponse -= new EventHandler<ResponseEventArgs<Misc.PressureSensorInfo>>(GetPressureSensor);
                m_ConnResponse.SetPressureCalibrationParameterResponse -= new EventHandler<ResponseEventArgs<String>>(SetPressureCalibrationParameter);
                m_ConnResponse.SetPressureCalibrationPValueResponse -= new EventHandler<ResponseEventArgs<String>>(SetPressureCalibrationPValue);
            }
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
            if (m_bStopClick==false && m_ConnResponse != null && m_ConnResponse.IsOpen())
            {
                m_ConnResponse.GetPressureSensor();
            }
        }

        private void StartSamplingPointTimer()
        {
            //时钟开启后5秒钟才能
            StopSamplingPointTimer();
            mSamplingPointStart = true;
            mSamplingPointTimer.Start();
        }

        private void StopSamplingPointTimer()
        {
            mSamplingCount = 0;
            mSamplingPointStart = false;
            mSamplingPointTimer.Stop();
        }
        private void OnSamplingPointTimer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            if(m_bStopClick==false && mSamplingCount >= 3)
            {
                StopSamplingPointTimer();
                StartCh1Timer();
                m_ConnResponse.SetStartControl();
                return;
            }
            if(m_bStopClick==false && mSamplingPointStart)
            {
                m_ConnResponse.GetPressureSensor();
                mSamplingCount++;
            }
        }

        #endregion

        public void SetPid(PumpID pid)
        {
            detail.Pid = pid;
            m_LocalPid = pid;
            switch(pid)
            {
                case PumpID.Graseby2000:
                case PumpID.Graseby2100:
                case PumpID.GrasebyC6:
                case PumpID.GrasebyC6T:
                case PumpID.WZ50C6:
                case PumpID.WZ50C6T:
                    PressureForm.SamplingPoints.Clear();
                    PressureForm.SamplingPoints.AddRange(PressureForm.SingleSamplingPoints1);
                    PressureForm.SamplingPoints.AddRange(PressureForm.SingleSamplingPoints2);
                    break;
                case PumpID.GrasebyF6:
                case PumpID.GrasebyF6_2:
                case PumpID.WZS50F6:
                case PumpID.WZS50F6_2:
                    PressureForm.SamplingPoints.Clear();
                    PressureForm.SamplingPoints.AddRange(PressureForm.SamplingPoints1);
                    PressureForm.SamplingPoints.AddRange(PressureForm.SamplingPoints2);
                    break;
            }
        }

        /// <summary>
        /// 设置详细信息通道号
        /// </summary>
        /// <param name="channel"></param>
        public void SetChannel(int channel)
        {
            detail.Channel = channel;
        }

        /// <summary>
        ///仅检测串口使用
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnPTool_DetectDeviceDataRecerived(object sender, EventArgs e)
        {
            PToolingDataEventArgs args = e as PToolingDataEventArgs;
            SetWeightValue(args.Weight, true);
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

        private void OnPTool_DeviceDataRecerived(object sender, EventArgs e)
        {
            PToolingDataEventArgs args = e as PToolingDataEventArgs;

            //采样数据，输出
            if (mSamplingPointStart)
            {
                System.Diagnostics.Debug.Write("工装的公斤值：");
                System.Diagnostics.Debug.WriteLine(args.Weight.ToString("F3"));
                //先取最后一次采样数据
                if (mSamplingCount == 2)
                {
                    lock (mSamplingPointList)
                    {
                        if (mSamplingPointList.Count > 0)
                        {
                            SampleData sp = mSamplingPointList[mSamplingPointList.Count - 1];
                            sp.m_Weight = args.Weight;
                        }
                    }
                }
            }

            lock (m_Ch1SampleDataList)
            {
                if (m_Ch1SampleDataList.Count > 0)
                {
                    SampleData sp = m_Ch1SampleDataList[m_Ch1SampleDataList.Count - 1];
                    sp.m_Weight = args.Weight;
                }
            }

            if (mCurrentSamplingIndex < PressureForm.SamplingPoints.Count && mCurrentSamplingIndex >= 0)
            {
                if (args.Weight >= PressureForm.SamplingPoints[mCurrentSamplingIndex]
                    || Math.Abs(args.Weight - PressureForm.SamplingPoints[mCurrentSamplingIndex])<= PressureForm.m_SamplingError
                    )
                {
                    StopCh1Timer();
                    m_ConnResponse.SetStopControl(GlobalResponse.CommandPriority.High);
                    Thread.Sleep(4000);
                    StartSamplingPointTimer();
                    mCurrentSamplingIndex++;
                }
            }

            SetWeightValue(args.Weight, false);
            DrawSingleAccuracyMap();
            //当采集到的重量大于配置参数时，可以停止采集，并计算相关数据写入到泵中
            PumpID pid = PumpID.None;
            if (this.m_Channel == 1)
            {
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
            }
            else
            {
                pid = m_LocalPid;
            }

            //if (mCurrentSamplingIndex >= PressureForm.SamplingPoints.Count)
            //{
            //    Complete(1);
            //    AlertMessageWhenComplete(string.Format("采样结束，共{0}个采样点！", mCurrentSamplingIndex));
            //    return;
            //}
         

            float max = PressureManager.Instance().GetMaxBySizeLevel(pid, 50, Misc.OcclusionLevel.H);

            int pointsCount = PressureForm.SamplingPoints.Count;
            if (pointsCount > 0)
                max = (float)Math.Max(PressureForm.SamplingPoints[pointsCount - 1], max);
            if (max < args.Weight)
            {
                Complete();
                AlertMessageWhenComplete(string.Format("压力值已超出最大范围{0},调试结束!", max));
            }
        }

        private void AlertMessageWhenComplete(string msg)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new DelegateAlertMessageWhenComplete(AlertMessageWhenComplete), new object[] { msg });
                return;
            }
            MessageBox.Show(msg);
        }

        private void SetPValue(float p)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new DelegateSetPValue(SetPValue), new object[] { p });
                return;
            }
            lbPValue.Text = (p * 100).ToString("F0");
            m_GrasebyDevice.Close();
        }

        private void SetWeightValue(float weight, bool isDetect = false)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new DelegateSetWeightValue(SetWeightValue), new object[] { weight, isDetect });
            }
            lbWeight.Text = weight.ToString("F3");
            if (isDetect)
                m_DetectPTool.Close();
        }

        /// <summary>
        /// 调试结束，保存相关数据，停止泵，停止时钟
        /// </summary>
        private void Complete()
        {
            StopSamplingPointTimer();
            StopCh1Timer();
            if (m_ConnResponse != null)
            {
                if (m_ConnResponse.IsOpen())
                {
                    m_ConnResponse.SetStopControl(GlobalResponse.CommandPriority.High);
                    //RemoveHandler();
                    Thread.Sleep(500);
                    //CalcuatePressure(m_LocalPid, m_Ch1SampleDataList);以后采用一次方程计算
                    //List<PressureCalibrationParameter> caliParameters = CalcuatePressureBySlope(m_LocalPid, m_Ch1SampleDataList);

                    List<PressureCalibrationParameter> caliParameters = CalcuatePressureByPoly(m_LocalPid, m_Ch1SampleDataList);

                    if (m_LocalPid == PumpID.GrasebyF6_2 || m_LocalPid == PumpID.WZS50F6_2)
                    {
                        if (OnSamplingComplete != null && caliParameters != null && caliParameters.Count > 0)
                            OnSamplingComplete(this, new DoublePumpDataArgs(caliParameters));
                    }
                    Thread.Sleep(500);
                    m_ConnResponse.CloseConnection();
                }
            }
            if (m_PTool != null && m_PTool.IsOpen())
                m_PTool.Close();

            string fileName = m_PumpNo;
            if (m_LocalPid == PumpID.GrasebyF6 || m_LocalPid == PumpID.WZS50F6)
                fileName = string.Format("{0}{1}道{2}{3}", m_LocalPid.ToString(), m_Channel, m_PumpNo, DateTime.Now.ToString("yyyy-MM-dd HH_mm_ss"));
            else
                fileName = string.Format("{0}{1}{2}", m_LocalPid.ToString(), m_PumpNo, DateTime.Now.ToString("yyyy-MM-dd HH_mm_ss"));
            string path = Path.GetDirectoryName(Assembly.GetAssembly(typeof(PressureForm)).Location) + "\\压力调试数据";
            if (!System.IO.Directory.Exists(path))
                System.IO.Directory.CreateDirectory(path);
            string saveFileName = path + "\\" + fileName + ".xlsx";
            Export(m_Channel, saveFileName);
            EnableContols(true);
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
            m_gh.DrawLine(m_WaveLinePen, new PointF(x0, y0), new PointF(x1, y1));
        }

        private void WavelinePanel_Paint(object sender, PaintEventArgs e)
        {
            DrawCoordinate(m_XCoordinateMaxValue, m_XSectionCount, m_YCoordinateMaxValue, m_YSectionCount);
            //DrawAccuracyMap(m_XSectionCount, m_YSectionCount);
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
                Font fontChartDes = new Font("Noto Sans CJK SC Bold", 12);
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
                //写图形描述字符
                m_gh.DrawString("P值压力值关系图", fontChartDes, m_WaveLineBrush, new PointF(yEndPoint.X + 180, yEndPoint.Y));
                //y轴的起始点，从底部往上
                PointF yOriginalPoint = originalpoint;//new PointF((float)rect.Left + LEFTBORDEROFFSET, rect.Bottom - TOPBOTTOMFFSET);
                m_gh.DrawLine(Pens.Black, yOriginalPoint, yEndPoint);
                //画Y坐标箭头
                PointF arrowpointLeft = new PointF(yEndPoint.X - 6, yEndPoint.Y + 12);
                PointF arrowpointRight = new PointF(yEndPoint.X + 6, yEndPoint.Y + 12);
                m_gh.DrawLine(Pens.Black, arrowpointLeft, yEndPoint);
                m_gh.DrawLine(Pens.Black, arrowpointRight, yEndPoint);
                //画Y坐标文字
                //m_gh.DrawString("压力值(V)", fontTitle, Brushes.Black, new PointF(yEndPoint.X + 10, yEndPoint.Y));
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
                m_gh.DrawString(VOL, fontTitle, m_WaveLineBrush, new PointF(xEndPoint.X - 80, 10));
                SizeF fontSize = m_gh.MeasureString(VOL, fontTitle);

                m_gh.DrawLine(m_WaveLinePen, new PointF(xEndPoint.X - 100, 10 + fontSize.Height / 2), new PointF(xEndPoint.X - 80, 10 + fontSize.Height / 2));
            }
            catch (Exception e)
            {
                MessageBox.Show("DrawHomeostasisMap Error:" + e.Message);
            }
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
                m_gh.DrawLine(m_WaveLinePen, new PointF(x0, y0), new PointF(x1, y1));
            }
        }

        /// <summary>
        /// 当不可用时，将按钮图标变灰
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Chart_EnabledChanged(object sender, EventArgs e)
        {
            if (this.Enabled)
            {
                picStart.Image = global::PTool.Properties.Resources.icon_start_Blue;
                picStop.Image = global::PTool.Properties.Resources.icon_stop_blue;
                picDetail.Image = global::PTool.Properties.Resources.icon_tablelist_blue;
                if (m_Channel == 2)
                {
                    picChannel.Image = global::PTool.Properties.Resources.icon_2_blue;
                }
            }
            else
            {
                picStart.Image = global::PTool.Properties.Resources.icon_start_gray;
                picStop.Image = global::PTool.Properties.Resources.icon_stop_gray;
                picDetail.Image = global::PTool.Properties.Resources.icon_tablelist_gray;
                if (m_Channel == 2)
                {
                    picChannel.Image = global::PTool.Properties.Resources.icon_2_gray;
                }
            }
        }

        /// <summary>
        /// 串口选择时发送命令
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cbToolingPort_SelectedIndexChanged(object sender, EventArgs e)
        {
            lbWeight.Text = "-----";
            if (m_DetectPTool == null)
                m_DetectPTool = new PTooling();
            if (m_DetectPTool.IsOpen())
                m_DetectPTool.Close();
            m_DetectPTool.Init(cbToolingPort.Items[cbToolingPort.SelectedIndex].ToString());
            m_DetectPTool.Open();
            Thread.Sleep(500);
            m_DetectPTool.ReadWeight();
        }

        private void cbPumpPort_SelectedIndexChanged(object sender, EventArgs e)
        {
            lbPValue.Text = "-----";
            if (m_GrasebyDevice == null)
                m_GrasebyDevice = new Graseby9600();
            if (m_GrasebyDevice.IsOpen())
                m_GrasebyDevice.Close();

            switch (m_LocalPid)
            {
                case PumpID.GrasebyC6:
                    m_GrasebyDevice.SetDeviceType(DeviceType.GrasebyC6);
                    break;
                case PumpID.GrasebyF6:
                    m_GrasebyDevice.SetDeviceType(DeviceType.GrasebyF6);
                    break;
                case PumpID.GrasebyC6T:
                    m_GrasebyDevice.SetDeviceType(DeviceType.GrasebyC6T);
                    break;
                case PumpID.Graseby2000:
                    m_GrasebyDevice.SetDeviceType(DeviceType.Graseby2000);
                    break;
                case PumpID.Graseby2100:
                    m_GrasebyDevice.SetDeviceType(DeviceType.Graseby2100);
                    break;
                case PumpID.WZ50C6:
                    m_GrasebyDevice.SetDeviceType(DeviceType.WZ50C6);
                    break;
                case PumpID.WZS50F6:
                    m_GrasebyDevice.SetDeviceType(DeviceType.WZS50F6);
                    break;
                case PumpID.WZ50C6T:
                    m_GrasebyDevice.SetDeviceType(DeviceType.WZ50C6T);
                    break;
            }
            m_GrasebyDevice.Init(cbPumpPort.Items[cbPumpPort.SelectedIndex].ToString());
            m_GrasebyDevice.Open();
            m_GrasebyDevice.Get();
        }

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
                if (m_ConnResponse != null && m_ConnResponse.IsOpen())
                    m_ConnResponse.SetStartControl();
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
                lbPValue.Text = "";
                Complete();
                EnableContols(true);
                MessageBox.Show("读取压力值失败，串口连接失败！");
            }
            else
            {
                Misc.PressureSensorInfo paras = args.EventData;
                lbPValue.Text = (paras.pressureVoltage * 100).ToString("F0");

                //采样数据，输出
                if(mSamplingPointStart)
                {
                    System.Diagnostics.Debug.Write("泵的P值：");
                    System.Diagnostics.Debug.WriteLine((paras.pressureVoltage * 100).ToString("F0"));
                    if (mSamplingCount == 2)
                    {
                        lock (mSamplingPointList)
                        {
                            mSamplingPointList.Add(new SampleData(DateTime.Now, paras.pressureVoltage, -1000f));
                        }
                    }
                }


                lock (m_Ch1SampleDataList)
                {
                    m_Ch1SampleDataList.Add(new SampleData(DateTime.Now, paras.pressureVoltage, -1000f));
                }
                if (m_PTool != null && m_PTool.IsOpen())
                {
                    m_PTool.ReadWeight();
                }
                else
                {
                    Complete();
                    EnableContols(true);
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
                Logger.Instance().Error("写入压力数据失败");
            }
            else
            {
                Logger.Instance().Info("写入压力数据成功");
            }
        }

        /// <summary>
        /// P值写入响应函数
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        private void SetPressureCalibrationPValue(object sender, ResponseEventArgs<string> args)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new EventHandler<ResponseEventArgs<string>>(SetPressureCalibrationPValue), new object[] { sender, args });
                return;
            }
            if (String.Empty != args.ErrorMessage)
            {
                Logger.Instance().Error("写入P值失败");
            }
            else
            {
                Logger.Instance().Info("写入P值成功");
            }
        }

        #endregion

        /// <summary>
        /// 速率不能输入非法值
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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

        private void EnableContols(bool bEnabled = true)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new DelegateEnableContols(EnableContols), new object[] { bEnabled });
                return;
            }
            cbToolingPort.Enabled = bEnabled;
            cbPumpPort.Enabled = bEnabled;
            tbRate.Enabled = bEnabled;
            tbOprator.Enabled = bEnabled;
            picStart.Enabled = bEnabled;
            picStop.Enabled = !bEnabled;

            if (!bEnabled)
            {
                picStart.Image = global::PTool.Properties.Resources.icon_start_gray;
                picStop.Image = global::PTool.Properties.Resources.icon_stop_blue;
            }
            else
            {
                picStart.Image = global::PTool.Properties.Resources.icon_start_Blue;
                picStop.Image = global::PTool.Properties.Resources.icon_stop_gray;
            }
            if (SamplingStartOrStop != null)
            {
                SamplingStartOrStop(this, new StartOrStopArgs(bEnabled));
            }
        }

        private void Export(int channel, string name)
        {
            List<SampleData> sampleDataList = null;
            sampleDataList = m_Ch1SampleDataList;
            if (sampleDataList == null || sampleDataList.Count == 0)
                return;
            string title = string.Empty;
            if (m_LocalPid == PumpID.GrasebyF6 || m_LocalPid == PumpID.WZS50F6)
            {
                title = string.Format("泵型号:{0}{1}道 产品序号:{2} 工装编号:{3}", m_LocalPid.ToString(), channel, m_PumpNo, m_ToolingNo);
            }
            else
            {
                title = string.Format("泵型号：{0} 产品序号:{1} 工装编号:{2}", m_LocalPid.ToString(), m_PumpNo, m_ToolingNo);
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
                ws.Cell(index, 3).Value = sampleDataList[i].m_PressureValue * 100;
                index++;
            }
            wb.SaveAs(name);
        }

        /// <summary>
        /// 生成第三方公司需要的表格
        /// </summary>
        /// <param name="name"></param>
        /// <param name="caliParameters">已经生成好的数据，直接写到表格中</param>
        private void GenReport(string name, List<PressureCalibrationParameter> caliParameters, string nameBackup="")
        {
            if (caliParameters == null || caliParameters.Count == 0)
                return;
            string title = string.Empty;
            if (m_LocalPid == PumpID.GrasebyF6 || m_LocalPid == PumpID.WZS50F6)
            {
                title = string.Format("泵型号:{0}{1}道 产品序号:{2} 工装编号:{3}", m_LocalPid.ToString(), m_Channel, m_PumpNo, m_ToolingNo);
            }
            else
            {
                title = string.Format("泵型号：{0} 产品序号:{1} 工装编号:{2}", m_LocalPid.ToString(), m_PumpNo, m_ToolingNo);
            }
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("压力调试数据");
            int columnIndex = 0;
            ws.Cell(1, ++columnIndex).Value = "机器编号";
            ws.Cell(1, ++columnIndex).Value = "机器型号";
            ws.Cell(1, ++columnIndex).Value = "道数";
            ws.Cell(1, ++columnIndex).Value = "工装编号";
            ws.Cell(1, ++columnIndex).Value = "P0值";
            ws.Cell(1, ++columnIndex).Value = "10mlL预设值";
            ws.Cell(1, ++columnIndex).Value = "10mlC预设值";
            ws.Cell(1, ++columnIndex).Value = "10mlH预设值";
            ws.Cell(1, ++columnIndex).Value = "20mlL预设值";
            ws.Cell(1, ++columnIndex).Value = "20mlC预设值";
            ws.Cell(1, ++columnIndex).Value = "20mlH预设值";
            ws.Cell(1, ++columnIndex).Value = "30mlL预设值";
            ws.Cell(1, ++columnIndex).Value = "30mlC预设值";
            ws.Cell(1, ++columnIndex).Value = "30mlH预设值";
            ws.Cell(1, ++columnIndex).Value = "50mlL预设值";
            ws.Cell(1, ++columnIndex).Value = "50mlC预设值";
            ws.Cell(1, ++columnIndex).Value = "50mlH预设值";
            ws.Cell(1, ++columnIndex).Value = "10ml低压";
            ws.Cell(1, ++columnIndex).Value = "10ml中压";
            ws.Cell(1, ++columnIndex).Value = "10ml高压";
            ws.Cell(1, ++columnIndex).Value = "20ml低压";
            ws.Cell(1, ++columnIndex).Value = "20ml中压";
            ws.Cell(1, ++columnIndex).Value = "20ml高压";
            ws.Cell(1, ++columnIndex).Value = "30ml低压";
            ws.Cell(1, ++columnIndex).Value = "30ml中压";
            ws.Cell(1, ++columnIndex).Value = "30ml高压";
            ws.Cell(1, ++columnIndex).Value = "50ml低压";
            ws.Cell(1, ++columnIndex).Value = "50ml中压";
            ws.Cell(1, ++columnIndex).Value = "50ml高压";
            ws.Cell(1, ++columnIndex).Value = "操作员";

            columnIndex = 0;
            ws.Cell(2, ++columnIndex).Value = m_PumpNo;
            ws.Cell(2, ++columnIndex).Value = m_LocalPid.ToString();
            ws.Cell(2, ++columnIndex).Value = m_Channel;
            ws.Cell(2, ++columnIndex).Value = m_ToolingNo;
            ws.Cell(2, ++columnIndex).Value = m_Ch1SampleDataList.Min(x => x.m_PressureValue) * 100;
            float mid = PressureManager.Instance().GetMidBySizeLevel(m_LocalPid, 10, Misc.OcclusionLevel.L);
            ws.Cell(2, ++columnIndex).Value = mid == 0 ? "" : (mid).ToString("F2");
            mid = PressureManager.Instance().GetMidBySizeLevel(m_LocalPid, 10, Misc.OcclusionLevel.C);
            ws.Cell(2, ++columnIndex).Value = mid == 0 ? "" : (mid).ToString("F2");
            mid = PressureManager.Instance().GetMidBySizeLevel(m_LocalPid, 10, Misc.OcclusionLevel.H);
            ws.Cell(2, ++columnIndex).Value = mid == 0 ? "" : (mid).ToString("F2");
            ws.Cell(2, ++columnIndex).Value = PressureManager.Instance().GetMidBySizeLevel(m_LocalPid, 20, Misc.OcclusionLevel.L);
            ws.Cell(2, ++columnIndex).Value = PressureManager.Instance().GetMidBySizeLevel(m_LocalPid, 20, Misc.OcclusionLevel.C);
            ws.Cell(2, ++columnIndex).Value = PressureManager.Instance().GetMidBySizeLevel(m_LocalPid, 20, Misc.OcclusionLevel.H);
            ws.Cell(2, ++columnIndex).Value = PressureManager.Instance().GetMidBySizeLevel(m_LocalPid, 30, Misc.OcclusionLevel.L);
            ws.Cell(2, ++columnIndex).Value = PressureManager.Instance().GetMidBySizeLevel(m_LocalPid, 30, Misc.OcclusionLevel.C);
            ws.Cell(2, ++columnIndex).Value = PressureManager.Instance().GetMidBySizeLevel(m_LocalPid, 30, Misc.OcclusionLevel.H);
            ws.Cell(2, ++columnIndex).Value = PressureManager.Instance().GetMidBySizeLevel(m_LocalPid, 50, Misc.OcclusionLevel.L);
            ws.Cell(2, ++columnIndex).Value = PressureManager.Instance().GetMidBySizeLevel(m_LocalPid, 50, Misc.OcclusionLevel.C);
            ws.Cell(2, ++columnIndex).Value = PressureManager.Instance().GetMidBySizeLevel(m_LocalPid, 50, Misc.OcclusionLevel.H);

            PressureCalibrationParameter para = null;
            para = caliParameters.Find((x) => { return x.m_SyringeSize == 10; });
            if (para != null)
            {
                columnIndex = 17;
                ws.Cell(2, ++columnIndex).Value = para.m_PressureL * 100;
                ws.Cell(2, ++columnIndex).Value = para.m_PressureC * 100;
                ws.Cell(2, ++columnIndex).Value = para.m_PressureH * 100;
            }
            para = caliParameters.Find((x) => { return x.m_SyringeSize == 20; });
            if (para != null)
            {
                columnIndex = 20;
                ws.Cell(2, ++columnIndex).Value = para.m_PressureL * 100;
                ws.Cell(2, ++columnIndex).Value = para.m_PressureC * 100;
                ws.Cell(2, ++columnIndex).Value = para.m_PressureH * 100;
            }

            para = caliParameters.Find((x) => { return x.m_SyringeSize == 30; });
            if (para != null)
            {
                columnIndex = 23;
                ws.Cell(2, ++columnIndex).Value = para.m_PressureL * 100;
                ws.Cell(2, ++columnIndex).Value = para.m_PressureC * 100;
                ws.Cell(2, ++columnIndex).Value = para.m_PressureH * 100;
            }

            para = caliParameters.Find((x) => { return x.m_SyringeSize == 50; });
            if (para != null)
            {
                columnIndex = 26;
                ws.Cell(2, ++columnIndex).Value = para.m_PressureL * 100;
                ws.Cell(2, ++columnIndex).Value = para.m_PressureC * 100;
                ws.Cell(2, ++columnIndex).Value = para.m_PressureH * 100;
            }
            ws.Cell(2, ++columnIndex).Value = tbOprator.Text;

            ws.Range(1, 1, 2, 1).SetDataType(XLCellValues.Text);
            ws.Range(1, 4, 2, 4).SetDataType(XLCellValues.Text);
            wb.SaveAs(name);
            Thread.Sleep(1000);
            File.Copy(name, nameBackup, true);
            IntPtr handle = UserMessageHelper.FindWindow(null, "压力测试工具");
            UserMessageHelper.SendMessage(handle, 0x1EE1, 0, 0);
        }

        /// <summary>
        /// 双道泵生成报告,两道数据放在一张表格中
        /// </summary>
        /// <param name="name"></param>
        /// <param name="sampleDataList"></param>
        public void GenDoublePunmpReport(string name, List<List<PressureCalibrationParameter>> sampleDataList, string tool2No, string nameBackup = "")
        {
            if (sampleDataList == null || sampleDataList.Count < 2)
                return;
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
PumpID pid2 = m_LocalPid;


            /*
            List<PressureParameter>[] parameters = new List<PressureParameter>[2];
            parameters[0] = new List<PressureParameter>();
            parameters[1] = new List<PressureParameter>();
            
            
            //第一道泵的压力参数
            ProductPressure pp = PressureManager.Instance().GetPressureByProductID(pid);
            //第二道泵的压力参数（根据需求，两道泵传感器有差别，参数需要单独配置）
            ProductPressure pp2 = PressureManager.Instance().GetPressureByProductID(pid2);
            if (pp == null || pp2 == null)
                return;
            List<LevelPressure> lps = pp.GetLevelPressureList();
            List<LevelPressure> lps2 = pp2.GetLevelPressureList();
            List<float> midWeights = new List<float>();
            List<int> sizes = new List<int>();
            if (pid == PumpID.WZS50F6 || pid == PumpID.GrasebyF6)
                sizes.Add(10);
            sizes.Add(20);
            sizes.Add(30);
            sizes.Add(50);

            //两道泵分开
            foreach (var size in sizes)
            {
                foreach (Misc.OcclusionLevel level in Enum.GetValues(typeof(Misc.OcclusionLevel)))
                {
                    LevelPressure lp = lps.Find((x) => { return x.m_Level == level; });
                    LevelPressure lp2 = lps2.Find((x) => { return x.m_Level == level; });
                    if (lp != null)
                    {
                        SizePressure sp = lp.Find(size);
                        SizePressure sp2 = lp2.Find(size);
                        if (sp != null && sp2 != null)
                        {
                            PressureParameter para = new PressureParameter(size, level, sp.m_Mid, 0);
                            parameters[0].Add(para);
                            PressureParameter para2 = new PressureParameter(size, level, sp2.m_Mid, 0);
                            parameters[1].Add(para2);
                        }

                    }
                }
            }
            //找到相关的值
            FindNearestPValue(ref parameters[0], sampleDataList[0]);
            FindNearestPValue(ref parameters[1], sampleDataList[1]);

            List<PressureCalibrationParameter>[] caliParameters = new List<PressureCalibrationParameter>[2];
            caliParameters[0] = new List<PressureCalibrationParameter>();
            caliParameters[1] = new List<PressureCalibrationParameter>();
            foreach (var size in sizes)
            {
                PressureCalibrationParameter p = new PressureCalibrationParameter();
                PressureCalibrationParameter p2 = new PressureCalibrationParameter();
                p.m_SyringeSize = size;
                p2.m_SyringeSize = size;
                List<PressureParameter> findobjs = parameters[0].FindAll((x) => { return x.m_SyringeSize == size; });
                List<PressureParameter> findobjs2 = parameters[1].FindAll((x) => { return x.m_SyringeSize == size; });
                foreach (var obj in findobjs)
                {
                    switch (obj.m_Level)
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
                foreach (var obj in findobjs2)
                {
                    switch (obj.m_Level)
                    {
                        case Misc.OcclusionLevel.L:
                            p2.m_PressureL = obj.m_Pressure;
                            break;
                        case Misc.OcclusionLevel.C:
                            p2.m_PressureC = obj.m_Pressure;
                            break;
                        case Misc.OcclusionLevel.H:
                            p2.m_PressureH = obj.m_Pressure;
                            break;
                        default: break;
                    }
                }
                caliParameters[0].Add(p);
                caliParameters[1].Add(p2);
            }

            */
            List<PressureCalibrationParameter>[] caliParameters = new List<PressureCalibrationParameter>[2];
            caliParameters[0] = sampleDataList[0];
            caliParameters[1] = sampleDataList[1];
            if (caliParameters == null || caliParameters.Length == 0)
                return;
            string title = string.Empty;
            if (m_LocalPid == PumpID.GrasebyF6 || m_LocalPid == PumpID.WZS50F6) //if (m_LocalPid == PumpID.GrasebyF6_2 || m_LocalPid == PumpID.WZS50F6_2)
            {
                title = string.Format("泵型号:{0}{1}道 产品序号:{2} 工装编号:{3}", m_LocalPid.ToString(), m_Channel, m_PumpNo, m_ToolingNo);
            }
            else
            {
                title = string.Format("泵型号：{0} 产品序号:{1} 工装编号:{2}", m_LocalPid.ToString(), m_PumpNo, m_ToolingNo);
            }
            #region 定义表头
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("压力调试数据");
            int columnIndex = 0, rowIndex = 2;
            ws.Cell(1, ++columnIndex).Value = "机器编号";
            ws.Cell(1, ++columnIndex).Value = "机器型号";
            ws.Cell(1, ++columnIndex).Value = "道数";
            ws.Cell(1, ++columnIndex).Value = "工装编号";
            ws.Cell(1, ++columnIndex).Value = "P0值";
            ws.Cell(1, ++columnIndex).Value = "10mlL预设值";
            ws.Cell(1, ++columnIndex).Value = "10mlC预设值";
            ws.Cell(1, ++columnIndex).Value = "10mlH预设值";
            ws.Cell(1, ++columnIndex).Value = "20mlL预设值";
            ws.Cell(1, ++columnIndex).Value = "20mlC预设值";
            ws.Cell(1, ++columnIndex).Value = "20mlH预设值";
            ws.Cell(1, ++columnIndex).Value = "30mlL预设值";
            ws.Cell(1, ++columnIndex).Value = "30mlC预设值";
            ws.Cell(1, ++columnIndex).Value = "30mlH预设值";
            ws.Cell(1, ++columnIndex).Value = "50mlL预设值";
            ws.Cell(1, ++columnIndex).Value = "50mlC预设值";
            ws.Cell(1, ++columnIndex).Value = "50mlH预设值";
            ws.Cell(1, ++columnIndex).Value = "10ml低压";
            ws.Cell(1, ++columnIndex).Value = "10ml中压";
            ws.Cell(1, ++columnIndex).Value = "10ml高压";
            ws.Cell(1, ++columnIndex).Value = "20ml低压";
            ws.Cell(1, ++columnIndex).Value = "20ml中压";
            ws.Cell(1, ++columnIndex).Value = "20ml高压";
            ws.Cell(1, ++columnIndex).Value = "30ml低压";
            ws.Cell(1, ++columnIndex).Value = "30ml中压";
            ws.Cell(1, ++columnIndex).Value = "30ml高压";
            ws.Cell(1, ++columnIndex).Value = "50ml低压";
            ws.Cell(1, ++columnIndex).Value = "50ml中压";
            ws.Cell(1, ++columnIndex).Value = "50ml高压";
            ws.Cell(1, ++columnIndex).Value = "操作员";

            #endregion

            #region 给第一道泵进行表格赋值
            for (int iLoop = 0; iLoop < sampleDataList.Count - 1 && sampleDataList.Count == 2; iLoop++)
            {
                columnIndex = 0;
                ws.Cell(rowIndex, ++columnIndex).Value = m_PumpNo;
                ws.Cell(rowIndex, ++columnIndex).Value = pid.ToString();
                ws.Cell(rowIndex, ++columnIndex).Value = iLoop + 1;
                if (iLoop == 0)
                    ws.Cell(rowIndex, ++columnIndex).Value = m_ToolingNo;
                else if (iLoop == 1)
                    ws.Cell(rowIndex, ++columnIndex).Value = tool2No;
                else
                    ws.Cell(rowIndex, ++columnIndex).Value = m_ToolingNo;
                ws.Cell(rowIndex, ++columnIndex).Value = caliParameters[iLoop][0].m_P0 * 100;
                float mid = PressureManager.Instance().GetMidBySizeLevel(pid, 10, Misc.OcclusionLevel.L);
                ws.Cell(rowIndex, ++columnIndex).Value = mid == 0 ? "" : (mid).ToString("F2");
                mid = PressureManager.Instance().GetMidBySizeLevel(pid, 10, Misc.OcclusionLevel.C);
                ws.Cell(rowIndex, ++columnIndex).Value = mid == 0 ? "" : (mid).ToString("F2");
                mid = PressureManager.Instance().GetMidBySizeLevel(pid, 10, Misc.OcclusionLevel.H);
                ws.Cell(rowIndex, ++columnIndex).Value = mid == 0 ? "" : (mid).ToString("F2");
                ws.Cell(rowIndex, ++columnIndex).Value = PressureManager.Instance().GetMidBySizeLevel(pid, 20, Misc.OcclusionLevel.L);
                ws.Cell(rowIndex, ++columnIndex).Value = PressureManager.Instance().GetMidBySizeLevel(pid, 20, Misc.OcclusionLevel.C);
                ws.Cell(rowIndex, ++columnIndex).Value = PressureManager.Instance().GetMidBySizeLevel(pid, 20, Misc.OcclusionLevel.H);
                ws.Cell(rowIndex, ++columnIndex).Value = PressureManager.Instance().GetMidBySizeLevel(pid, 30, Misc.OcclusionLevel.L);
                ws.Cell(rowIndex, ++columnIndex).Value = PressureManager.Instance().GetMidBySizeLevel(pid, 30, Misc.OcclusionLevel.C);
                ws.Cell(rowIndex, ++columnIndex).Value = PressureManager.Instance().GetMidBySizeLevel(pid, 30, Misc.OcclusionLevel.H);
                ws.Cell(rowIndex, ++columnIndex).Value = PressureManager.Instance().GetMidBySizeLevel(pid, 50, Misc.OcclusionLevel.L);
                ws.Cell(rowIndex, ++columnIndex).Value = PressureManager.Instance().GetMidBySizeLevel(pid, 50, Misc.OcclusionLevel.C);
                ws.Cell(rowIndex, ++columnIndex).Value = PressureManager.Instance().GetMidBySizeLevel(pid, 50, Misc.OcclusionLevel.H);

                PressureCalibrationParameter para = null;
                para = caliParameters[iLoop].Find((x) => { return x.m_SyringeSize == 10; });
                if (para != null)
                {
                    columnIndex = 17;
                    ws.Cell(rowIndex, ++columnIndex).Value = para.m_PressureL * 100;
                    ws.Cell(rowIndex, ++columnIndex).Value = para.m_PressureC * 100;
                    ws.Cell(rowIndex, ++columnIndex).Value = para.m_PressureH * 100;
                }
                para = caliParameters[iLoop].Find((x) => { return x.m_SyringeSize == 20; });
                if (para != null)
                {
                    columnIndex = 20;
                    ws.Cell(rowIndex, ++columnIndex).Value = para.m_PressureL * 100;
                    ws.Cell(rowIndex, ++columnIndex).Value = para.m_PressureC * 100;
                    ws.Cell(rowIndex, ++columnIndex).Value = para.m_PressureH * 100;
                }

                para = caliParameters[iLoop].Find((x) => { return x.m_SyringeSize == 30; });
                if (para != null)
                {
                    columnIndex = 23;
                    ws.Cell(rowIndex, ++columnIndex).Value = para.m_PressureL * 100;
                    ws.Cell(rowIndex, ++columnIndex).Value = para.m_PressureC * 100;
                    ws.Cell(rowIndex, ++columnIndex).Value = para.m_PressureH * 100;
                }

                para = caliParameters[iLoop].Find((x) => { return x.m_SyringeSize == 50; });
                if (para != null)
                {
                    columnIndex = 26;
                    ws.Cell(rowIndex, ++columnIndex).Value = para.m_PressureL * 100;
                    ws.Cell(rowIndex, ++columnIndex).Value = para.m_PressureC * 100;
                    ws.Cell(rowIndex, ++columnIndex).Value = para.m_PressureH * 100;
                }
                ws.Cell(rowIndex, ++columnIndex).Value = tbOprator.Text;
                rowIndex++;
            }
            #endregion

            #region 给第二道泵进行表格赋值
            for (int iLoop = 1; iLoop < sampleDataList.Count && sampleDataList.Count == 2; iLoop++)
            {
                columnIndex = 0;
                ws.Cell(rowIndex, ++columnIndex).Value = m_PumpNo;
                ws.Cell(rowIndex, ++columnIndex).Value = pid2.ToString();
                ws.Cell(rowIndex, ++columnIndex).Value = iLoop + 1;
                if (iLoop == 0)
                    ws.Cell(rowIndex, ++columnIndex).Value = m_ToolingNo;
                else if (iLoop == 1)
                    ws.Cell(rowIndex, ++columnIndex).Value = tool2No;
                else
                    ws.Cell(rowIndex, ++columnIndex).Value = m_ToolingNo;
                ws.Cell(rowIndex, ++columnIndex).Value = caliParameters[iLoop][0].m_P0 * 100;
                float mid = PressureManager.Instance().GetMidBySizeLevel(pid2, 10, Misc.OcclusionLevel.L);
                ws.Cell(rowIndex, ++columnIndex).Value = mid == 0 ? "" : (mid).ToString("F2");
                mid = PressureManager.Instance().GetMidBySizeLevel(pid2, 10, Misc.OcclusionLevel.C);
                ws.Cell(rowIndex, ++columnIndex).Value = mid == 0 ? "" : (mid).ToString("F2");
                mid = PressureManager.Instance().GetMidBySizeLevel(pid2, 10, Misc.OcclusionLevel.H);
                ws.Cell(rowIndex, ++columnIndex).Value = mid == 0 ? "" : (mid).ToString("F2");
                ws.Cell(rowIndex, ++columnIndex).Value = PressureManager.Instance().GetMidBySizeLevel(pid2, 20, Misc.OcclusionLevel.L);
                ws.Cell(rowIndex, ++columnIndex).Value = PressureManager.Instance().GetMidBySizeLevel(pid2, 20, Misc.OcclusionLevel.C);
                ws.Cell(rowIndex, ++columnIndex).Value = PressureManager.Instance().GetMidBySizeLevel(pid2, 20, Misc.OcclusionLevel.H);
                ws.Cell(rowIndex, ++columnIndex).Value = PressureManager.Instance().GetMidBySizeLevel(pid2, 30, Misc.OcclusionLevel.L);
                ws.Cell(rowIndex, ++columnIndex).Value = PressureManager.Instance().GetMidBySizeLevel(pid2, 30, Misc.OcclusionLevel.C);
                ws.Cell(rowIndex, ++columnIndex).Value = PressureManager.Instance().GetMidBySizeLevel(pid2, 30, Misc.OcclusionLevel.H);
                ws.Cell(rowIndex, ++columnIndex).Value = PressureManager.Instance().GetMidBySizeLevel(pid2, 50, Misc.OcclusionLevel.L);
                ws.Cell(rowIndex, ++columnIndex).Value = PressureManager.Instance().GetMidBySizeLevel(pid2, 50, Misc.OcclusionLevel.C);
                ws.Cell(rowIndex, ++columnIndex).Value = PressureManager.Instance().GetMidBySizeLevel(pid2, 50, Misc.OcclusionLevel.H);

                PressureCalibrationParameter para = null;
                para = caliParameters[iLoop].Find((x) => { return x.m_SyringeSize == 10; });
                if (para != null)
                {
                    columnIndex = 17;
                    ws.Cell(rowIndex, ++columnIndex).Value = para.m_PressureL * 100;
                    ws.Cell(rowIndex, ++columnIndex).Value = para.m_PressureC * 100;
                    ws.Cell(rowIndex, ++columnIndex).Value = para.m_PressureH * 100;
                }
                para = caliParameters[iLoop].Find((x) => { return x.m_SyringeSize == 20; });
                if (para != null)
                {
                    columnIndex = 20;
                    ws.Cell(rowIndex, ++columnIndex).Value = para.m_PressureL * 100;
                    ws.Cell(rowIndex, ++columnIndex).Value = para.m_PressureC * 100;
                    ws.Cell(rowIndex, ++columnIndex).Value = para.m_PressureH * 100;
                }

                para = caliParameters[iLoop].Find((x) => { return x.m_SyringeSize == 30; });
                if (para != null)
                {
                    columnIndex = 23;
                    ws.Cell(rowIndex, ++columnIndex).Value = para.m_PressureL * 100;
                    ws.Cell(rowIndex, ++columnIndex).Value = para.m_PressureC * 100;
                    ws.Cell(rowIndex, ++columnIndex).Value = para.m_PressureH * 100;
                }

                para = caliParameters[iLoop].Find((x) => { return x.m_SyringeSize == 50; });
                if (para != null)
                {
                    columnIndex = 26;
                    ws.Cell(rowIndex, ++columnIndex).Value = para.m_PressureL * 100;
                    ws.Cell(rowIndex, ++columnIndex).Value = para.m_PressureC * 100;
                    ws.Cell(rowIndex, ++columnIndex).Value = para.m_PressureH * 100;
                }

                ws.Cell(rowIndex, ++columnIndex).Value = tbOprator.Text;
                rowIndex++;
            }
            #endregion

            ws.Range(1, 1, rowIndex, 1).SetDataType(XLCellValues.Text);
            ws.Range(1, 4, rowIndex, 4).SetDataType(XLCellValues.Text);
            wb.SaveAs(name);
            Thread.Sleep(1000);
            File.Copy(name, nameBackup, true);
            sampleDataList.Clear();
            IntPtr handle = UserMessageHelper.FindWindow(null, "压力测试工具");
            UserMessageHelper.SendMessage(handle, 0x1EE1, 0, 0);
        }

        private void CalcuatePressure(PumpID pid, List<SampleData> sampleDataList)
        {
            if (sampleDataList == null || sampleDataList.Count == 0)
                return;
            if (m_Channel == 1)
            {
                switch (pid)
                {
                    case PumpID.GrasebyF6_2:
                        pid = PumpID.GrasebyF6;
                        break;
                    case PumpID.WZS50F6_2:
                        pid = PumpID.WZS50F6;
                        break;
                    default:
                        break;
                }
            }
            List<PressureParameter> parameters = new List<PressureParameter>();
            ProductPressure pp = PressureManager.Instance().GetPressureByProductID(pid);
            if (pp == null)
                return;
            List<LevelPressure> lps = pp.GetLevelPressureList();
            List<float> midWeights = new List<float>();
            List<int> sizes = new List<int>();
            if (pid == PumpID.WZS50F6 || pid == PumpID.GrasebyF6 || pid == PumpID.GrasebyF6_2 || pid == PumpID.WZS50F6_2)
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
            //if (!IsValidEx(sampleDataList))
            //{
            //    sampleDataList.Clear();
            //    MessageBox.Show("测量数据异常，请重试！");
            //    return;
            //}

            float pValue = FindZeroPValue(sampleDataList);

            if (pValue * 100 >= PressureForm.RangeMaxP || pValue * 100 <= PressureForm.RangeMinP)
            {
                Logger.Instance().ErrorFormat("P值超范围，请重试！P值={0},最小值={1},最大值={2}", pValue, PressureForm.RangeMinP, PressureForm.RangeMaxP);
                sampleDataList.Clear();
                MessageBox.Show("P值超范围，请重试！");
                return;
            }

            WritePValue2Pump(pValue);
            Logger.Instance().InfoFormat("测量结束，P值为{0}", pValue);

            List<PressureCalibrationParameter> caliParameters = new List<PressureCalibrationParameter>();
            foreach (var size in sizes)
            {
                PressureCalibrationParameter p = new PressureCalibrationParameter();
                p.m_P0 = pValue;
                p.m_SyringeSize = size;
                List<PressureParameter> findobjs = parameters.FindAll((x) => { return x.m_SyringeSize == size; });
                foreach (var obj in findobjs)
                {
                    switch (obj.m_Level)
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
#if DEBUG
            //if (this.Channel==1)
            //{
            //    if (IsOutOfRange(caliParameters))
            //    {
            //        sampleDataList.Clear();
            //        MessageBox.Show("P值变化大，请重试！");
            //        return;
            //    }
            //}
#else
            if (IsOutOfRange(caliParameters))
            {
                sampleDataList.Clear();
                MessageBox.Show("P值变化大，请重试！");
                return;
            }
#endif
          

            WritePressureCaliParameter2Pump(caliParameters);
            detail.P0 = m_Ch1SampleDataList.Min(x => x.m_PressureValue) * 100;
            detail.CaliParameters = caliParameters;
            //如果是单泵，则调用下列代码
            if (m_LocalPid != PumpID.GrasebyF6_2 && m_LocalPid != PumpID.WZS50F6_2)
            {
                string path = Path.GetDirectoryName(Assembly.GetAssembly(typeof(PressureForm)).Location) + "\\数据导出";
                string fileName = m_PumpNo;
                if (m_LocalPid == PumpID.GrasebyF6 || m_LocalPid == PumpID.WZS50F6)
                    fileName = string.Format("{0}{1}道{2}{3}", m_LocalPid.ToString(), m_Channel, m_PumpNo, DateTime.Now.ToString("yyyy-MM-dd HH_mm_ss"));
                else
                    fileName = string.Format("{0}{1}{2}", m_LocalPid.ToString(), m_PumpNo, DateTime.Now.ToString("yyyy-MM-dd HH_mm_ss"));
                if (!System.IO.Directory.Exists(path))
                    System.IO.Directory.CreateDirectory(path);
                string saveFileName = path + "\\" + fileName + ".xlsx";

                string path2 = Path.GetDirectoryName(Assembly.GetAssembly(typeof(PressureForm)).Location) + "\\数据导出备份";
                string fileName2 = m_PumpNo;
                if (m_LocalPid == PumpID.GrasebyF6 || m_LocalPid == PumpID.WZS50F6)
                    fileName2 = string.Format("{0}{1}道{2}{3}", m_LocalPid.ToString(), m_Channel, m_PumpNo, DateTime.Now.ToString("yyyy-MM-dd HH_mm_ss"));
                else
                    fileName2 = string.Format("{0}{1}{2}", m_LocalPid.ToString(), m_PumpNo, DateTime.Now.ToString("yyyy-MM-dd HH_mm_ss"));
                if (!System.IO.Directory.Exists(path2))
                    System.IO.Directory.CreateDirectory(path2);
                string saveFileName2 = path2 + "\\" + fileName2 + ".xlsx";

                GenReport(saveFileName, caliParameters,saveFileName2);
            }
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

        /// <summary>
        /// 找到P值最小的
        /// </summary>
        /// <param name="sampleDataList"></param>
        private float FindZeroPValue(List<SampleData> sampleDataList)
        {
            if (sampleDataList == null || sampleDataList.Count == 0)
            {
                Logger.Instance().Error("测量数据为空，无法确定P值大小!");
                return 0;
            }
            float minP = m_Ch1SampleDataList.Min(x => x.m_PressureValue);
            return minP;
        }

        /// <summary>
        /// 写入LCH压力档值
        /// </summary>
        /// <param name="caliParas"></param>
        private void WritePressureCaliParameter2Pump(List<PressureCalibrationParameter> caliParas)
        {
            if (caliParas == null || caliParas.Count == 0)
                return;
            for (int i = 0; i < caliParas.Count; i++)
            {
                Logger.Instance().InfoFormat("准备写入压力数据Size={0},L={1},C={2},H={3}", (byte)(caliParas[i].m_SyringeSize), caliParas[i].m_PressureL, caliParas[i].m_PressureC, caliParas[i].m_PressureH);
                m_ConnResponse.SetPressureCalibrationParameter((byte)(caliParas[i].m_SyringeSize), caliParas[i].m_PressureL, caliParas[i].m_PressureC, caliParas[i].m_PressureH);
                Logger.Instance().Info("写入压力数据完毕，等待泵回应......");
                Thread.Sleep(1000);
            }
        }

        /// <summary>
        /// 在写入泵之前判断数据是否超范围
        /// </summary>
        /// <param name="caliParas"></param>
        /// <returns></returns>
        private bool IsOutOfRange(List<PressureCalibrationParameter> caliParas)
        {
            bool bRet = false;
            for (int i = 0; i < caliParas.Count; i++)
            {
                if (caliParas[i].m_PressureL * 100 >= PressureForm.PressureCalibrationMax)
                {
                    Logger.Instance().ErrorFormat("P值变化大，请重试！L ={0}", caliParas[i].m_PressureL);
                    bRet = true;
                    break;
                }
                if (caliParas[i].m_PressureC * 100 >= PressureForm.PressureCalibrationMax)
                {
                    Logger.Instance().ErrorFormat("P值变化大，请重试！C ={0}", caliParas[i].m_PressureC);
                    bRet = true;
                    break;
                }
                if (caliParas[i].m_PressureH * 100 >= PressureForm.PressureCalibrationMax)
                {
                    Logger.Instance().ErrorFormat("P值变化大，请重试！H ={0}", caliParas[i].m_PressureH);
                    bRet = true;
                    break;
                }

                if (caliParas[i].m_PressureH <= caliParas[i].m_PressureC || caliParas[i].m_PressureC <= caliParas[i].m_PressureL || caliParas[i].m_PressureH <= caliParas[i].m_PressureL)
                {
                    Logger.Instance().ErrorFormat("P值异常高档位值小于低档位，请重试！{0}，{1}，{2}", caliParas[i].m_PressureL, caliParas[i].m_PressureC, caliParas[i].m_PressureH);
                    bRet = true;
                    break;
                }
            }
            return bRet;
        }

        /// <summary>
        /// 写入P值
        /// </summary>
        /// <param name="caliParas"></param>
        private void WritePValue2Pump(float p)
        {
            Logger.Instance().InfoFormat("准备写入P值 {0}", p);
            m_ConnResponse.SetPressureCalibrationPValue(p);
            Logger.Instance().Info("写入P值完毕，等待泵回应......");
            Thread.Sleep(1000);
        }

        /// <summary>
        /// 计算得出的数据是否合法，此函数弃用
        /// </summary>
        /// <param name="parameters"></param>
        /// <returns></returns>
        private bool IsValid(List<PressureParameter> parameters)
        {
            bool bRet = true;
            foreach (var p in parameters)
            {
                if (p.m_Pressure >= p.m_MidWeight * 2.5)
                {
                    bRet = false;
                    break;
                }
            }
            return bRet;
        }

        private bool IsValidEx(List<SampleData> sampleDataList)
        {
            bool bRet = true;
            int iLength = sampleDataList.Count;
            if (iLength > 0 && sampleDataList[0].m_PressureValue > 2.5f)
            {
                bRet = false;
                return bRet;
            }
            //当采集到的重量大于配置参数时，可以停止采集，并计算相关数据写入到泵中
            //PumpID pid = PumpID.None;
            //switch (m_LocalPid)
            //{
            //    case PumpID.GrasebyF6_2:
            //        pid = PumpID.GrasebyF6;
            //        break;
            //    case PumpID.WZS50F6_2:
            //        pid = PumpID.WZS50F6;
            //        break;
            //    default:
            //        pid = m_LocalPid;
            //        break;
            //}
            //float max = PressureManager.Instance().GetMaxBySizeLevel(pid, 50, Misc.OcclusionLevel.H);
            //if (iLength > 0 && sampleDataList[iLength - 1].m_Weight < max)
            //{
            //    bRet = false;
            //}
            return bRet;
        }

        private void picStart_Click(object sender, EventArgs e)
        {
            m_bStopClick = false;
            mCurrentSamplingIndex = 0;
            detail.P0 = 0f;
            detail.CaliParameters.Clear();
            detail.ClearLabelValue();
            m_Ch1SampleDataList.Clear();
            mSamplingPointList.Clear();
            WavelinePanel.Invalidate();

            switch (m_LocalPid)
            {
                case PumpID.Graseby2000:
                case PumpID.Graseby2100:
                case PumpID.GrasebyC6:
                case PumpID.GrasebyC6T:
                case PumpID.WZ50C6:
                case PumpID.WZ50C6T:
                    PressureForm.SamplingPoints.Clear();
                    PressureForm.SamplingPoints.AddRange(PressureForm.SingleSamplingPoints1);
                    PressureForm.SamplingPoints.AddRange(PressureForm.SingleSamplingPoints2);
                    break;
                case PumpID.GrasebyF6:
                case PumpID.GrasebyF6_2:
                case PumpID.WZS50F6:
                case PumpID.WZS50F6_2:
                    PressureForm.SamplingPoints.Clear();
                    PressureForm.SamplingPoints.AddRange(PressureForm.SamplingPoints1);
                    PressureForm.SamplingPoints.AddRange(PressureForm.SamplingPoints2);
                    break;
            }

            #region 参数输入检查

            if (SamplingStartOrStop != null)
            {
                SamplingStartOrStop(this, new StartOrStopArgs(true));
            }

            if (string.IsNullOrEmpty(tbOprator.Text))
            {
                MessageBox.Show("请输入操作员工号");
                return;
            }
            if (tbOprator.Text.Length != 8)
            {
                MessageBox.Show("请输入正确操作员工号");
                return;
            }

            if (string.IsNullOrEmpty(PumpNo))
            {
                MessageBox.Show("请输入产品序号");
                return;
            }

            if (PumpNo.Length != PressureForm.SerialNumberCount)
            {
                string message = string.Format("产品序号长度不等于{0}位", PressureForm.SerialNumberCount);
                MessageBox.Show(message);
                return;
            }

            float weight = 0;
            float rate = 0;

            if (cbToolingPort.SelectedIndex < 0)
            {
                MessageBox.Show("请选择工装串口");
                return;
            }
            if (!float.TryParse(lbWeight.Text, out weight))
            {
                MessageBox.Show("工装串口连接错误，请正确选择串口！");
                return;
            }

            if (cbPumpPort.SelectedIndex < 0)
            {
                MessageBox.Show("请选择泵串口");
                return;
            }
            if (!float.TryParse(lbPValue.Text, out weight))
            {
                MessageBox.Show("泵串口连接错误，请正确选择串口！");
                return;
            }
            if (string.IsNullOrEmpty(tbRate.Text))
            {
                MessageBox.Show("请输入速率！");
                return;
            }
            if (!float.TryParse(tbRate.Text, out rate))
            {
                MessageBox.Show("请正确输入速率！");
                return;
            }
            #endregion

            #region 泵型号选择
            Misc.ProductID pid = Misc.ProductID.None;
            switch (m_LocalPid)
            {
                case PumpID.GrasebyC6:
                    pid = Misc.ProductID.GrasebyC6;
                    break;
                case PumpID.WZ50C6:
                    pid = Misc.ProductID.GrasebyC6;
                    break;
                case PumpID.GrasebyC6T:
                case PumpID.WZ50C6T:
                    pid = Misc.ProductID.GrasebyC6T;
                    break;
                case PumpID.Graseby2000:
                    pid = Misc.ProductID.Graseby2000;
                    break;
                case PumpID.Graseby2100:
                    pid = Misc.ProductID.Graseby2100;
                    break;
                case PumpID.WZS50F6:
                    pid = Misc.ProductID.GrasebyF6;
                    break;
                case PumpID.GrasebyF6:
                    pid = Misc.ProductID.GrasebyF6;
                    break;
                case PumpID.GrasebyF6_2:
                    pid = Misc.ProductID.GrasebyF6;
                    break;
                case PumpID.WZS50F6_2:
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

            if (m_ConnResponse != null && m_ConnResponse.IsOpen())
            {
                m_ConnResponse.CloseConnection();
            }
#if DEBUG
            m_ConnResponse = new GlobalResponse(Misc.ProductID.GrasebyC6, Misc.CommunicationProtocolType.General);

#else
            m_ConnResponse = new GlobalResponse(pid, Misc.CommunicationProtocolType.General);

#endif
            m_ConnResponse.Initialize(cbPumpPort.Items[cbPumpPort.SelectedIndex].ToString(), BAUDRATE);
            RemoveHandler();
            AddHandler();
            if (m_PTool != null)
            {
                if (m_PTool.IsOpen())
                {

                }
                else
                {
                    m_PTool.Init(cbToolingPort.Items[cbToolingPort.SelectedIndex].ToString());
                    m_PTool.Open();
                }
            }
            else
            {
                m_PTool = new PTooling();
                m_PTool.Init(cbToolingPort.Items[cbToolingPort.SelectedIndex].ToString());
                m_PTool.Open();
            }
            m_PTool.Tare();
            Thread.Sleep(500);
            m_ConnResponse.SetVTBIParameter(0, rate);
            StartCh1Timer();
            EnableContols(false);
        }

        private void picStop_Click(object sender, EventArgs e)
        {
            m_bStopClick = true;//人工停止了
            mCurrentSamplingIndex = 0;
            Complete();
            EnableContols(true);
        }

        private void picDetail_Click(object sender, EventArgs e)
        {
            this.detail.Show();
        }

        private void tbOprator_KeyPress(object sender, KeyPressEventArgs e)
        {
            TextBox txt = sender as TextBox;
            if (char.IsNumber(e.KeyChar) || e.KeyChar == (char)Keys.Back)
            {
                e.Handled = false;                         //让操作生效
                if (txt.Text.Length >= 8)
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

        public void InputOpratorNumber(string number)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new DelegateInputOpratorNumber(InputOpratorNumber), new object[] { number });
            }
            else
            {
                this.tbOprator.Text = number;
            }
        }

        private void tbOprator_TextChanged(object sender, EventArgs e)
        {
            if (this.Channel == 1)
            {
                if (OpratorNumberInput != null)
                {
                    OpratorNumberInput(this, new OpratorNumberArgs(tbOprator.Text));
                }
            }
        }


        #region 一次方程计算压力P值

        /// <summary>
        /// 计算一次方程的a和b，需要两个点
        /// </summary>
        /// <param name="x1"></param>
        /// <param name="y1"></param>
        /// <param name="x2"></param>
        /// <param name="y2"></param>
        /// <returns></returns>
        private Tuple<double, double> CalculateSlope(double x1, double y1, double x2, double y2)
        {
            double a = (y2 - y1) / (x2 - x1);
            double b = y2 - (a * x2);
            Tuple<double, double> slope = new Tuple<double, double>(a ,b);
            return slope;
        }

        /// <summary>
        /// 根据一次方程计算泵的P值
        /// </summary>
        /// <param name="pid"></param>
        /// <param name="sampleDataList"></param>
        private List<PressureCalibrationParameter>  CalcuatePressureBySlope(PumpID pid, List<SampleData> sampleDataList)
        {
            if (sampleDataList == null || sampleDataList.Count == 0)
                return null;
            if (m_Channel == 1)
            {
                switch (pid)
                {
                    case PumpID.GrasebyF6_2:
                        pid = PumpID.GrasebyF6;
                        break;
                    case PumpID.WZS50F6_2:
                        pid = PumpID.WZS50F6;
                        break;
                    default:
                        break;
                }
            }
            List<PressureParameter> parameters = new List<PressureParameter>();
            ProductPressure pp = PressureManager.Instance().GetPressureByProductID(pid);
            if (pp == null)
                return null;
            List<LevelPressure> lps = pp.GetLevelPressureList();
            List<float> midWeights = new List<float>();
            List<int> sizes = new List<int>();
            if (pid == PumpID.WZS50F6 || pid == PumpID.GrasebyF6 || pid == PumpID.GrasebyF6_2 || pid == PumpID.WZS50F6_2)
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

            //计算相关的值，一次方程
            CalcuatePValue(ref parameters);

            if (!IsValidEx(sampleDataList))
            {
                sampleDataList.Clear();
                MessageBox.Show("测量数据异常，请重试！");
                return null;
            }

            float pValue = FindZeroPValue(sampleDataList);

            if (pValue * 100 >= PressureForm.RangeMaxP || pValue * 100 <= PressureForm.RangeMinP)
            {
                Logger.Instance().ErrorFormat("P值超范围，请重试！P值={0},最小值={1},最大值={2}", pValue, PressureForm.RangeMinP, PressureForm.RangeMaxP);
                sampleDataList.Clear();
                MessageBox.Show("P值超范围，请重试！");
                return null;
            }

            WritePValue2Pump(pValue);
            Logger.Instance().InfoFormat("测量结束，P值为{0}", pValue);

            List<PressureCalibrationParameter> caliParameters = new List<PressureCalibrationParameter>();
            foreach (var size in sizes)
            {
                PressureCalibrationParameter p = new PressureCalibrationParameter();
                p.m_P0 = pValue;
                p.m_SyringeSize = size;
                List<PressureParameter> findobjs = parameters.FindAll((x) => { return x.m_SyringeSize == size; });
                foreach (var obj in findobjs)
                {
                    switch (obj.m_Level)
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
#if DEBUG
            //if (this.Channel==1)
            //{
            //    if (IsOutOfRange(caliParameters))
            //    {
            //        sampleDataList.Clear();
            //        MessageBox.Show("P值变化大，请重试！");
            //        return;
            //    }
            //}
#else
            if (IsOutOfRange(caliParameters))
            {
                sampleDataList.Clear();
                MessageBox.Show("P值变化大，请重试！");
                return null;
            }
#endif


            WritePressureCaliParameter2Pump(caliParameters);
            detail.P0 = m_Ch1SampleDataList.Min(x => x.m_PressureValue) * 100;
            detail.CaliParameters = caliParameters;
            //如果是单泵，则调用下列代码
            if (m_LocalPid != PumpID.GrasebyF6_2 && m_LocalPid != PumpID.WZS50F6_2)
            {
                string path = Path.GetDirectoryName(Assembly.GetAssembly(typeof(PressureForm)).Location) + "\\数据导出";
                string fileName = m_PumpNo;
                if (m_LocalPid == PumpID.GrasebyF6 || m_LocalPid == PumpID.WZS50F6)
                    fileName = string.Format("{0}{1}道{2}{3}", m_LocalPid.ToString(), m_Channel, m_PumpNo, DateTime.Now.ToString("yyyy-MM-dd HH_mm_ss"));
                else
                    fileName = string.Format("{0}{1}{2}", m_LocalPid.ToString(), m_PumpNo, DateTime.Now.ToString("yyyy-MM-dd HH_mm_ss"));
                if (!System.IO.Directory.Exists(path))
                    System.IO.Directory.CreateDirectory(path);
                string saveFileName = path + "\\" + fileName + ".xlsx";

                string path2 = Path.GetDirectoryName(Assembly.GetAssembly(typeof(PressureForm)).Location) + "\\数据导出备份";
                string fileName2 = m_PumpNo;
                if (m_LocalPid == PumpID.GrasebyF6 || m_LocalPid == PumpID.WZS50F6)
                    fileName2 = string.Format("{0}{1}道{2}{3}", m_LocalPid.ToString(), m_Channel, m_PumpNo, DateTime.Now.ToString("yyyy-MM-dd HH_mm_ss"));
                else
                    fileName2 = string.Format("{0}{1}{2}", m_LocalPid.ToString(), m_PumpNo, DateTime.Now.ToString("yyyy-MM-dd HH_mm_ss"));
                if (!System.IO.Directory.Exists(path2))
                    System.IO.Directory.CreateDirectory(path2);
                string saveFileName2 = path2 + "\\" + fileName2 + ".xlsx";

                GenReport(saveFileName, caliParameters, saveFileName2);
            }
            //返回计算好的值
            return caliParameters;
        }

        private void CalcuatePValue(ref List<PressureParameter> parameters)
        {
            int pointCount = mSamplingPointList.Count;
            if (pointCount < 8)
            {
                Logger.Instance().WarnFormat("采样点{0}个点，曲线精度可能受影响！",pointCount );
            }

            StringBuilder sb = new StringBuilder("采样的点[重量：压力]：");
            for (int i = 0; i < pointCount; i++)
            {
                sb.Append(mSamplingPointList[i].m_Weight.ToString());
                sb.Append(" : ");
                sb.Append(mSamplingPointList[i].m_PressureValue.ToString());
                if (i != pointCount - 1)
                    sb.Append(",  ");
            }
            Logger.Instance().Info(sb.ToString());


            double[]a = CalculateSection1Poly();
            //double[]b = CalculateSection2Poly();//第二段直接用直线，不要拟合了

            Tuple<double, double> slope = CalculateSlope(mSamplingPointList[pointCount -3].m_Weight, mSamplingPointList[pointCount - 3].m_PressureValue, mSamplingPointList[pointCount - 1].m_Weight, mSamplingPointList[pointCount - 1].m_PressureValue);
            double[] b = new double[2];
            b[0] = slope.Item2;
            b[1] = slope.Item1;

            double temp = 0;
            temp = b[0] + b[1] * mSamplingPointList[pointCount - 2].m_Weight;
            if(Math.Abs(temp- mSamplingPointList[pointCount - 2].m_PressureValue) >= PressureForm.m_StandardError)
            {
                Logger.Instance().InfoFormat("直线方程计算结果与实际采样有较大误差,采样值={0}，计算值={1}", mSamplingPointList[pointCount - 2].m_PressureValue, temp);
                MessageBox.Show("采样误差大，请重试！");
                return;
            }
         
            int position = -1;
            foreach(var item in parameters)
            {
                position = GetPositionOfWeight(item.m_MidWeight);
                switch(position)
                {
                    case 0:
                        temp = a[0] + a[1] * item.m_MidWeight + a[2] * Math.Pow(item.m_MidWeight, 2);
                        break;
                    case 1:
                        temp = b[0] + b[1] * item.m_MidWeight;
                        break;
                    //case 2:
                    //    temp = c[0] + c[1] * item.m_MidWeight;
                    //    break;
                    default:
                        break;
                }
                item.m_Pressure = (float)Math.Round(temp, 2);
            }
        }
        #endregion

        #region 2次方程计算压力P值,11采样点分2段进行

        private  int GetPositionOfWeight(double weight)
        {
            int count1 = PressureForm.SamplingPoints1.Count;
            if (count1 < 3)
            {
                MessageBox.Show("采样点太少，无法计算拟合曲线！");
                return -1;
            }

            int count2 = PressureForm.SamplingPoints2.Count;
            if (count2 < 2)
            {
                MessageBox.Show("采样点太少，无法计算拟合曲线！");
                return -1;
            }

            //int count3 = PressureForm.SamplingPoints3.Count;
            //if (count3 <= 0)
            //{
            //    MessageBox.Show("采样点太少，无法计算拟合曲线！");
            //    return -1;
            //}

            if (weight <= PressureForm.SamplingPoints1[count1 - 1])
            {
                return 0;
            }
            else
                return 1;

        }

        /// <summary>
        ///y =  a0 + a1 * x + a2 * x *x + a3 * x *x*x
        /// </summary>
        /// <param name="arrX"></param>
        /// <param name="arrY"></param>
        /// <param name="length"></param>
        /// <param name="dimension"></param>
        /// <returns>a0,a1,a2,a3</returns>
        private double[] CalculatePoly(double[] arrX, double[] arrY, int dimension = 2)
        {
            return Polynomial.MultiLine(arrX, arrY, arrX.Length, dimension);
        }

        private double[] CalculateSection1Poly()
        {
            if (mSamplingPointList.Count < PressureForm.SamplingPoints1.Count)
            {
                MessageBox.Show("第一段曲线采样点太少，无法计算拟合曲线！");
                return null;
            }
            double[] arrX = new double[PressureForm.SamplingPoints1.Count];
            double[] arrY = new double[PressureForm.SamplingPoints1.Count];
            for (int i=0;i< PressureForm.SamplingPoints1.Count;i++)
            {
                arrX[i] = mSamplingPointList[i].m_Weight;
                arrY[i] = mSamplingPointList[i].m_PressureValue;
            }
            return CalculatePoly(arrX, arrY,2);
        }

        private double[] CalculateSection2Poly()
        {
            if (mSamplingPointList.Count < PressureForm.SamplingPoints1.Count  + 2  || PressureForm.SamplingPoints2.Count<2)
            {
                MessageBox.Show("第二段曲线采样点太少，无法计算拟合曲线！");
                return null;
            }
            int index = PressureForm.SamplingPoints1.Count;
             
            int count = mSamplingPointList.Count - index;

            count = Math.Min(count, PressureForm.SamplingPoints2.Count);
             
            double[] arrX = new double[count];
            double[] arrY = new double[count];

            for (int i = index; i < index + count; i++)
            {
                arrX[i- index] = mSamplingPointList[i].m_Weight;
                arrY[i- index] = mSamplingPointList[i].m_PressureValue;
            }
            return CalculatePoly(arrX, arrY);
        }

        //private double[] CalculateSection3Poly()
        //{
        //    if (mSamplingPointList.Count <= PressureForm.SamplingPoints1.Count + PressureForm.SamplingPoints2.Count)
        //    {
        //        MessageBox.Show("第三段曲线采样点太少，无法计算拟合曲线！");
        //        return null;
        //    }
        //    int index = PressureForm.SamplingPoints1.Count + PressureForm.SamplingPoints2.Count;

        //    if(mSamplingPointList.Count - index<3)
        //    {
        //        MessageBox.Show("第三段曲线采样点太少，无法计算拟合曲线！");
        //        return null;
        //    }

        //    double[] arrX = new double[mSamplingPointList.Count - index];
        //    double[] arrY = new double[mSamplingPointList.Count - index];

        //    for (int i = index; i < mSamplingPointList.Count; i++)
        //    {
        //        arrX[i - index] = mSamplingPointList[i].m_Weight;
        //        arrY[i - index] = mSamplingPointList[i].m_PressureValue;
        //    }
        //    return CalculatePoly(arrX, arrY, 1);
        //}

        /// <summary>
        /// 多项式拟合
        /// </summary>
        /// <param name="pid"></param>
        /// <param name="sampleDataList"></param>
        /// <returns></returns>
        private List<PressureCalibrationParameter> CalcuatePressureByPoly(PumpID pid, List<SampleData> sampleDataList)
        {
            if (sampleDataList == null || sampleDataList.Count == 0)
                return null;
            if (m_Channel == 1)
            {
                switch (pid)
                {
                    case PumpID.GrasebyF6_2:
                        pid = PumpID.GrasebyF6;
                        break;
                    case PumpID.WZS50F6_2:
                        pid = PumpID.WZS50F6;
                        break;
                    default:
                        break;
                }
            }
            List<PressureParameter> parameters = new List<PressureParameter>();
            ProductPressure pp = PressureManager.Instance().GetPressureByProductID(pid);
            if (pp == null)
                return null;
            List<LevelPressure> lps = pp.GetLevelPressureList();
            List<float> midWeights = new List<float>();
            List<int> sizes = new List<int>();
            if (pid == PumpID.WZS50F6 || pid == PumpID.GrasebyF6 || pid == PumpID.GrasebyF6_2 || pid == PumpID.WZS50F6_2)
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

            //计算相关的值，一次方程
            CalcuatePValue(ref parameters);

            if (!IsValidEx(sampleDataList))
            {
                sampleDataList.Clear();
                MessageBox.Show("测量数据异常，请重试！");
                return null;
            }

            float pValue = FindZeroPValue(sampleDataList);

            if (pValue * 100 >= PressureForm.RangeMaxP || pValue * 100 <= PressureForm.RangeMinP)
            {
                Logger.Instance().ErrorFormat("P值超范围，请重试！P值={0},最小值={1},最大值={2}", pValue, PressureForm.RangeMinP, PressureForm.RangeMaxP);
                sampleDataList.Clear();
                MessageBox.Show("P值超范围，请重试！");
                return null;
            }

            WritePValue2Pump(pValue);
            Logger.Instance().InfoFormat("测量结束，P值为{0}", pValue);

            List<PressureCalibrationParameter> caliParameters = new List<PressureCalibrationParameter>();
            foreach (var size in sizes)
            {
                PressureCalibrationParameter p = new PressureCalibrationParameter();
                p.m_P0 = pValue;
                p.m_SyringeSize = size;
                List<PressureParameter> findobjs = parameters.FindAll((x) => { return x.m_SyringeSize == size; });
                foreach (var obj in findobjs)
                {
                    switch (obj.m_Level)
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

            if (IsOutOfRange(caliParameters))
            {
                sampleDataList.Clear();
                MessageBox.Show("P值变化大，请重试！");
                return null;
            }

            WritePressureCaliParameter2Pump(caliParameters);
            detail.P0 = m_Ch1SampleDataList.Min(x => x.m_PressureValue) * 100;
            detail.CaliParameters = caliParameters;
            //如果是单泵，则调用下列代码
            if (m_LocalPid != PumpID.GrasebyF6_2 && m_LocalPid != PumpID.WZS50F6_2)
            {
                string path = Path.GetDirectoryName(Assembly.GetAssembly(typeof(PressureForm)).Location) + "\\数据导出";
                string fileName = m_PumpNo;
                if (m_LocalPid == PumpID.GrasebyF6 || m_LocalPid == PumpID.WZS50F6)
                    fileName = string.Format("{0}{1}道{2}{3}", m_LocalPid.ToString(), m_Channel, m_PumpNo, DateTime.Now.ToString("yyyy-MM-dd HH_mm_ss"));
                else
                    fileName = string.Format("{0}{1}{2}", m_LocalPid.ToString(), m_PumpNo, DateTime.Now.ToString("yyyy-MM-dd HH_mm_ss"));
                if (!System.IO.Directory.Exists(path))
                    System.IO.Directory.CreateDirectory(path);
                string saveFileName = path + "\\" + fileName + ".xlsx";

                string path2 = Path.GetDirectoryName(Assembly.GetAssembly(typeof(PressureForm)).Location) + "\\数据导出备份";
                string fileName2 = m_PumpNo;
                if (m_LocalPid == PumpID.GrasebyF6 || m_LocalPid == PumpID.WZS50F6)
                    fileName2 = string.Format("{0}{1}道{2}{3}", m_LocalPid.ToString(), m_Channel, m_PumpNo, DateTime.Now.ToString("yyyy-MM-dd HH_mm_ss"));
                else
                    fileName2 = string.Format("{0}{1}{2}", m_LocalPid.ToString(), m_PumpNo, DateTime.Now.ToString("yyyy-MM-dd HH_mm_ss"));
                if (!System.IO.Directory.Exists(path2))
                    System.IO.Directory.CreateDirectory(path2);
                string saveFileName2 = path2 + "\\" + fileName2 + ".xlsx";

                GenReport(saveFileName, caliParameters, saveFileName2);
            }
            //返回计算好的值
            return caliParameters;
        }

        #endregion



    }
}
