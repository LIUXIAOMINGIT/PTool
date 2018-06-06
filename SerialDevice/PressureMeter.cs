using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO.Ports;

namespace SerialDevice
{
    /// <summary>
    /// 压力表类:目前只定义安森 ACD-1F
    /// </summary>
    public class PressureMeter : DeviceBase
    {
        private List<byte> m_ReadBuffer = new List<byte>(); //存放数据缓存，如果数据到达数量少于指定长度，等待下次接受

        public PressureMeter()
        {
            this._deviceType = DeviceType.ACD_1F;
            _detectByteLength = 9;                                          //ACD-1F的数据长度为9个字节
            SetDetectBytes(new byte[5] { 0x40, 0x30, 0x30, 0x31, 0x21 });   //设置检测命令
            Init(9600, 8, StopBits.One, Parity.None, "");
        }

        public override void Get()
        {
            this._communicateDevice.SendData(this._detectCommandBytes);
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
            if (buffer.Count >= _detectByteLength && buffer[0] == 0x7B && buffer[_detectByteLength - 1] == 0x7D)
            {
                _bufferByCom.Clear();//找到串口，清除缓存
                base.OnDetectDataReceived(sender, args);
                System.Diagnostics.Debug.WriteLine("OnDetectDataReceived Invoked : " + args.PortName);
            }
        }

        /// <summary>
        /// 用于正式接收串口设备的数据
        /// -32768 ~ 32767在此范围的值，当传过来的值超过32767时一定是负值，此时要异或0xFFFF后再加1补码是为负数的绝对值
        /// 数据格式：以一对大括号开始结束，共9个字节
        /// { DP D5 D4 D3 D2 D1 UNIT }\{ E R R O R 0 1 }\{ E R R O R 0 2 }\{ E R R O R 0 3 }
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        public override void ReceiveData(object sender, DataTransmissionEventArgs args)
        {
            byte[] temp = new byte[_detectByteLength];
            try
            {
                lock (m_ReadBuffer)
                {
                    m_ReadBuffer.AddRange(args.EventData);
                }
                PressureMeterArgs para = Analyze(m_ReadBuffer);
                if (para != null)
                    base.ReceiveData(sender, para);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(ex.Message);
                throw ex;
            }
        }

        /// <summary>
        /// 对压力表数据进行解析,分别生成数据和单位
        /// </summary>
        /// <param name="eventData"></param>
        /// <returns></returns>
        private PressureMeterArgs Analyze(List<byte> eventData)
        {
            PressureMeterArgs args = null;
            if (eventData!=null && eventData.Count < _detectByteLength)
                return null;
            byte[] buffer = new byte[_detectByteLength];
            bool bFind = false;
            lock (m_ReadBuffer)
            {
                while (eventData.Count >= _detectByteLength)
                {
                    if (eventData[0] != 0x7B)
                    {
                        eventData.RemoveAt(0);
                        continue;
                    }
                    else
                    {
                        if (eventData[_detectByteLength - 1] != 0x7D)
                        {
                            eventData.RemoveAt(0);
                            continue;
                        }
                        else
                        {
                            bFind = true;
                            eventData.CopyTo(0, buffer, 0, _detectByteLength);
                            eventData.RemoveRange(0, _detectByteLength);
                        }
                    }
                }
            }

            if (bFind)
            {
                //小数点错误
                if (buffer[1] > 0x33 || buffer[1] < 0x30)
                    return null;
                //单位错误
                if (buffer[7] > 0x37 || buffer[1] < 0x30)
                    return null;

                for (int iLoop = 2; iLoop <= 6; iLoop++)
                {
                    if (buffer[iLoop] < 0x30 || buffer[iLoop] > 0x39)
                    {
                        return null;
                    }
                }
                int D5 = (buffer[2] - 0x30) * 10000;
                int D4 = (buffer[3] - 0x30) * 1000;
                int D3 = (buffer[4] - 0x30) * 100;
                int D2 = (buffer[5] - 0x30) * 10;
                int D1 = (buffer[6] - 0x30);
                int total = D1 + D2 + D3 + D4 + D5;
                if (total > 0xFFFF) //错误数字
                    return null;
                ushort num = 0;
                float sum = 0;
                if (total > 32767) //负数
                {
                    num = (ushort)total;
                    num = (ushort)(num ^ 0xFFFF + 0x0001);
                    sum = -1 * num;
                }
                else
                {
                    num = (ushort)total;
                    sum = num;
                }
                switch (buffer[1])
                {
                    case 0x31:
                        sum *= 0.1f;
                        break;
                    case 0x32:
                        sum *= 0.01f;
                        break;
                    case 0x33:
                        sum *= 0.001f;
                        break;
                }
                PressureUnit unit = (PressureUnit)(buffer[7] - 0x30);

                //单位换算，统一为KPa
                switch(unit)
                {
                    case PressureUnit.MPa:
                        sum = sum * 1000;
                        break;
                    default:
                        break;
                }
                args = new PressureMeterArgs(unit, sum);
                return args;
            }
            else
            {
                return null;
            }

        }


    }
}
