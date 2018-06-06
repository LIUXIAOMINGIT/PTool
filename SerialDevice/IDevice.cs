using System;
using System.IO.Ports;

namespace SerialDevice
{
    /// <summary>
    /// 串口设备统一接口
    /// </summary>
    public interface IDevice
    {
        event EventHandler<EventArgs> DeviceDataRecerived;

        /// <summary>
        /// 初始化串口参数
        /// </summary>
        /// <param name="detectBytes">用于检测串口的命令字节数组</param>
        /// <param name="baudRate">波特率</param>
        /// <param name="dataBits">数据位</param>
        /// <param name="stopBits">停止位</param>
        /// <param name="parity">奇偶检验方式</param>
        void Init(byte[] detectBytes, int baudRate, int dataBits, StopBits stopBits, Parity parity,string portName);

        /// <summary>
        /// 初始化串口参数
        /// </summary>
        /// <param name="baudRate"></param>
        /// <param name="dataBits"></param>
        /// <param name="stopBits"></param>
        /// <param name="parity"></param>
        /// <param name="portName"></param>
        void Init(int baudRate, int dataBits, StopBits stopBits, Parity parity, string portName);

        /// <summary>
        /// 初始化串口名字
        /// </summary>
        /// <param name="portName"></param>
        void Init(string portName);

        /// <summary>
        /// 打开串口
        /// </summary>
        /// <returns></returns>
        bool Open();

        /// <summary>
        /// 关闭串口
        /// </summary>
        void Close();

        /// <summary>
        /// 返回串口设备连接状态
        /// </summary>
        /// <returns></returns>
        bool IsOpen();

        /// <summary>
        /// 设置检测命令字节
        /// </summary>
        /// <param name="detectBytes"></param>
        void SetDetectBytes(byte[] detectBytes);

        /// <summary>
        /// 检测执行函数
        /// </summary>
        /// <returns></returns>
        string DoDetect();

        /// <summary>
        /// 检测执行函数
        /// </summary>
        /// <returns></returns>
        Tuple<string, byte> DoDetectEx();

        /// <summary>
        /// 只检测指定的串口
        /// </summary>
        /// <param name="portName"></param>
        /// <returns></returns>
        string DoDetect(string portName);

        /// <summary>
        /// 发送GET命令，仪表类设备都应该具有类似的命令
        /// </summary>
        void Get();
        
        /// <summary>
        /// 设置某些参数
        /// </summary>
        /// <param name="buffer"></param>
        void Set(byte[] buffer);


    }
}
