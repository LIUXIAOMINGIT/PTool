using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO.Ports;

namespace SerialDevice
{
    public class HHSTooling : DeviceBase
    {
        private List<byte> m_ReadBuffer = new List<byte>(); //存放数据缓存，如果数据到达数量少于指定长度，等待下次接受
        public HHSTooling()
        {
            this._deviceType = DeviceType.HHS;
            _detectByteLength = 12;                                     //数据长度
            SetDetectBytes(new byte[] { 0x0B, 0x1C, 0x02, 0x00 });      //设置检测命令
            Init(115200, 8, StopBits.One, Parity.None, "");
        }

        public override void Get()
        {
            try
            {
                this._communicateDevice.SendData(this._detectCommandBytes);
            }
            catch(Exception ex)
            {

            }
        }

        public override void Set(byte[] buffer)
        {
        }

        /// <summary>
        /// 只能用于检测串口，收到串口设备数据包
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public override void OnDetectDataReceived(object sender, DataTransmissionEventArgs args)
        {
            List<byte> buffer = (List<byte>)_bufferByCom[args.PortName];
            if (buffer == null)
                return;
            buffer.AddRange(args.EventData);
            if (buffer.Count >= _detectByteLength && buffer[0] == 0x0B && buffer[1] == 0x1C)
            {
                _bufferByCom.Clear();//找到串口，清除缓存
                base.OnDetectDataReceived(sender, args);
                System.Diagnostics.Debug.WriteLine("OnDetectDataReceived Invoked : " + args.PortName);
            }
        }
        
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        public override void ReceiveData(object sender, DataTransmissionEventArgs args)
        {
           
        }
    }
}
