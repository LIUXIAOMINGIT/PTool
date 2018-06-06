using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO.Ports;
using System.Threading;

namespace SerialDevice
{
    public class DeviceBase : IDevice
    {
        protected DeviceType       _deviceType = DeviceType.None;
        protected string           _portName;
        protected int              _baudRate;
        protected int              _dataBits;
        protected StopBits         _stopBits;
        protected Parity           _parity;
        protected string           _detectedPortName  = string.Empty;               //检测到的串口名字存放在此处
        protected byte             _detectedProductID = 0;                          //检测到的产品ID存放在此处
        protected byte[]           _detectCommandBytes;
        protected AutoResetEvent   _detectEvent       = new AutoResetEvent(false);
        protected Hashtable        _bufferByCom       = new Hashtable();            //不同串口，不同的buffer
        protected int              _detectByteLength  = 0;                          //刷新命令的回应只有多少个字节
        protected const int WAITFOREVENTTIMEOUT       = 2000;                       //2秒
        protected SerialPortCommunication _communicateDevice = new SerialPortCommunication();     //用于串口通信的串口类

        public event EventHandler<EventArgs> DeviceDataRecerived;

        /// <summary>
        /// 初始化串口参数
        /// </summary>
        /// <param name="detectBytes">用于检测串口的命令字节数组</param>
        /// <param name="baudRate">波特率</param>
        /// <param name="dataBits">数据位</param>
        /// <param name="stopBits">停止位</param>
        /// <param name="parity">奇偶检验方式</param>
        public void Init(byte[] detectBytes, int baudRate, int dataBits, StopBits stopBits, Parity parity,string portName="")
        {
            _portName           = portName;
            _detectCommandBytes = detectBytes;
            _baudRate           = baudRate;
            _dataBits           = dataBits;
            _stopBits           = stopBits;
            _parity             = parity;
        }

        /// <summary>
        /// 初始化串口参数
        /// </summary>
        /// <param name="baudRate">波特率</param>
        /// <param name="dataBits">数据位</param>
        /// <param name="stopBits">停止位</param>
        /// <param name="parity">奇偶检验方式</param>
        public void Init(int baudRate, int dataBits, StopBits stopBits, Parity parity,string portName="")
        {
            _portName = portName;
            _baudRate = baudRate;
            _dataBits = dataBits;
            _stopBits = stopBits;
            _parity = parity;
        }

        public void Init(string portName)
        {
            _portName = portName;
        }

        /// <summary>
        /// 打开串口
        /// </summary>
        /// <returns></returns>
        public bool Open()
        {
            if (null == _communicateDevice)
                _communicateDevice = new SerialPortCommunication(); 
            if (_communicateDevice.IsOpen())
                _communicateDevice.Close();
            _communicateDevice.InitializeDevice(_portName, _baudRate, _dataBits, _stopBits, _parity, Handshake.None);
            _communicateDevice.DataReceived += ReceiveData;
            try
            {
                _communicateDevice.Open();
            }
            catch
            {
                return false;
            }
            return true;
        }

        /// <summary>
        /// 关闭串口
        /// </summary>
        public void Close()
        {
            if (null != _communicateDevice)
            {
                _communicateDevice.DataReceived -= ReceiveData;
                _communicateDevice.Close();
            }
        }

        /// <summary>
        /// 返回串口设备连接状态
        /// </summary>
        /// <returns></returns>
        public bool IsOpen()
        {
            return _communicateDevice.IsOpen();
        }

        /// <summary>
        /// 设置检测命令字节
        /// </summary>
        /// <param name="detectBytes"></param>
        public void SetDetectBytes(byte[] detectBytes)
        {
            _detectCommandBytes = detectBytes;
        }

        /// <summary>
        /// 检测线程函数，用于多线程检测
        /// </summary>
        /// <param name="com"></param>
        private void ProcDetectting(object com)
        {
            SerialPortCommunication serialPort = com as SerialPortCommunication;
            bool bOpen = serialPort.IsOpen();
            if (!bOpen)
                return;
            serialPort.SendData(_detectCommandBytes);
        }

        /// <summary>
        /// 检测执行函数
        /// </summary>
        /// <returns></returns>
        public string DoDetect()
        {
            _detectEvent.Reset();
            string connectedCom = string.Empty;
            string[] portNames = SerialPort.GetPortNames();
            List<Thread> threadPool = new List<Thread>();
            List<SerialPortCommunication> serialPortPool = new List<SerialPortCommunication>();
            _bufferByCom.Clear();
            foreach (string port in portNames)
            {
                //开启多线程，每个串口开一个
                _bufferByCom.Add(port, new List<byte>());
                Thread detectThread = new Thread(new ParameterizedThreadStart(ProcDetectting));
                SerialPortCommunication serialPort = new SerialPortCommunication();
                serialPort.InitializeDevice(port, _baudRate, _dataBits, _stopBits, _parity);
                serialPort.DataReceived += OnDetectDataReceived;
                try
                {
                    serialPort.Open();
                }
                catch
                {
                    continue;
                }
                serialPortPool.Add(serialPort);
                threadPool.Add(detectThread);
                detectThread.Start(serialPort);
            }
            for (int i = 0; i < threadPool.Count; i++)
            {
                threadPool[i].Join();
            }
            if (_detectEvent.WaitOne(WAITFOREVENTTIMEOUT))
            {
            }
            for (int i = 0; i < serialPortPool.Count; i++)
            {
                serialPortPool[i].Close();
            }
            for (int i = 0; i < threadPool.Count; i++)
            {
                threadPool[i].Abort();
            }
            return connectedCom = _detectedPortName;
        }

        public Tuple<string, byte> DoDetectEx()
        {
            _detectEvent.Reset();
            string connectedCom = string.Empty;
            string[] portNames = SerialPort.GetPortNames();//new string[] { "COM1" };//
            List<Thread> threadPool = new List<Thread>();
            List<SerialPortCommunication> serialPortPool = new List<SerialPortCommunication>();
            _bufferByCom.Clear();
            foreach (string port in portNames)
            {
                //开启多线程，每个串口开一个
                _bufferByCom.Add(port, new List<byte>());
                Thread detectThread = new Thread(new ParameterizedThreadStart(ProcDetectting));
                SerialPortCommunication serialPort = new SerialPortCommunication();
                serialPort.InitializeDevice(port, _baudRate, _dataBits, _stopBits, _parity);
                serialPort.DataReceived += OnDetectDataReceived;
                try
                {
                    serialPort.Open();
                }
                catch
                {
                    continue;
                }
                serialPortPool.Add(serialPort);
                threadPool.Add(detectThread);
                detectThread.Start(serialPort);
                //Thread.Sleep(500);
            }
            for (int i = 0; i < threadPool.Count; i++)
            {
                threadPool[i].Join();
            }
            if (_detectEvent.WaitOne(WAITFOREVENTTIMEOUT))
            {
            }
            for (int i = 0; i < serialPortPool.Count; i++)
            {
                serialPortPool[i].Close();
            }
            for (int i = 0; i < threadPool.Count; i++)
            {
                threadPool[i].Abort();
            }
            return new Tuple<string, byte>(_detectedPortName, _detectedProductID);
        }

        public string DoDetect(string portName)
        {
            _detectEvent.Reset();
            string connectedCom = string.Empty;
            string[] portNames = new string[]{portName};
            List<Thread> threadPool = new List<Thread>();
            List<SerialPortCommunication> serialPortPool = new List<SerialPortCommunication>();
            _bufferByCom.Clear();
            foreach (string port in portNames)
            {
                //开启多线程，每个串口开一个
                _bufferByCom.Add(port, new List<byte>());
                Thread detectThread = new Thread(new ParameterizedThreadStart(ProcDetectting));
                SerialPortCommunication serialPort = new SerialPortCommunication();
                serialPort.InitializeDevice(port, _baudRate, _dataBits, _stopBits, _parity);
                serialPort.DataReceived += OnDetectDataReceived;
                try
                {
                    serialPort.Open();
                }
                catch
                {
                    continue;
                }
                serialPortPool.Add(serialPort);
                threadPool.Add(detectThread);
                detectThread.Start(serialPort);
            }
            for (int i = 0; i < threadPool.Count; i++)
            {
                threadPool[i].Join();
            }
            if (_detectEvent.WaitOne(WAITFOREVENTTIMEOUT))
            {
            }
            for (int i = 0; i < serialPortPool.Count; i++)
            {
                serialPortPool[i].Close();
            }
            for (int i = 0; i < threadPool.Count; i++)
            {
                threadPool[i].Abort();
            }
            return connectedCom = _detectedPortName;
        }

        public virtual void Get()
        {
        }

        /// <summary>
        /// 设置某些参数
        /// </summary>
        /// <param name="buffer"></param>
        public virtual void Set(byte[] buffer)
        {

        }

        /// <summary>
        /// 只能用于检测串口，收到串口设备数据包
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public virtual void OnDetectDataReceived(object sender, DataTransmissionEventArgs args)
        {
           
        }

        /// <summary>
        /// 用于正式接收串口设备的数据
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        public virtual void ReceiveData(object sender, DataTransmissionEventArgs args)
        {
            if (DeviceDataRecerived != null)
            {
                DeviceDataRecerived(this, args);
            }
        }

    }
}
