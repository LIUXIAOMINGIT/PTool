using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO.Ports;

namespace SerialDevice
{
    public class PTooling : DeviceBase
    {
        private List<byte> m_ReadBuffer = new List<byte>(); //存放数据缓存，如果数据到达数量少于指定长度，等待下次接受
        public PTooling()
        {
            this._deviceType = DeviceType.DL101;
            _detectByteLength = 11;                                     //数据长度
            SetDetectBytes(new byte[] { 0x02, 0x52, 0x44, 0x53, 0x41, 0x2A, 0x0D });      //设置检测命令,读当前值，返回0x02 0x41 0x44 0x32 0x3C 0x35 0x32 0x30 0x38 0x72 0x0D 
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
        /// 读数据
        /// </summary>
        public void ReadWeight()
        {
            this._communicateDevice.SendData(this._detectCommandBytes);
        }

        /// <summary>
        /// 设置模块编号（设置1号模块通道）
        /// </summary>
        public void SetChannel()
        {
            this._communicateDevice.SendData(new byte[] { 0x02, 0x53, 0x4E, 0x4D, 0x41, 0x2F, 0x0D });
        }

        /// <summary>
        /// 模块置零02 5A 45 52 41 32 0D 
        /// </summary>
        public void Tare()
        {
            this._communicateDevice.SendData(new byte[] { 0x02, 0x5A, 0x45, 0x52, 0x41, 0x32, 0x0D });
        }

        /// <summary>
        /// 设置分度值和零点跟踪开关  02 44 54 32 41 0B 0D 
        /// </summary>
        public void SetScale()
        {
            this._communicateDevice.SendData(new byte[] { 0x02, 0x44, 0x54, 0x32, 0x41, 0x0B, 0x0D });
        }

        /// <summary>
        /// 设置3位小数 02 5A 50 37 41 22 0D
        /// </summary>
        public void SetDecimalPlace()
        {
            this._communicateDevice.SendData(new byte[] { 0x02, 0x5A, 0x50, 0x37, 0x41, 0x22, 0x0D });
        }

        /// <summary>
        /// 只能用于检测串口，收到串口设备数据包
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public override void OnDetectDataReceived(object sender, DataTransmissionEventArgs args)
        {
            byte[] buffer = new byte[_detectByteLength];
            lock (m_ReadBuffer)
            {
                m_ReadBuffer.AddRange(args.EventData);
                if (m_ReadBuffer.Count >= _detectByteLength)
                {
                    int endIndex = m_ReadBuffer.FindIndex(0, (x) => { return x == 0x0D; });
                    if (endIndex < 0)
                    {
                        return;
                    }
                    else if (endIndex > 0 && endIndex != _detectByteLength - 1)
                    {
                        m_ReadBuffer.RemoveRange(0, endIndex + 1);
                        return;
                    }
                    else
                    {
                        m_ReadBuffer.CopyTo(0, buffer, 0, _detectByteLength);
                        m_ReadBuffer.RemoveRange(0, _detectByteLength);
                    }
                }
                else
                {
                    return;
                }
            }
            int weight = 0;
            for (int iLoop = 7; iLoop >= 3; iLoop--)
            {
                weight += (buffer[iLoop] - 0x30) << (iLoop - 3) * 4;
            }

            int decimal_place = buffer[8] & 0x03;
            float fweightKiloGrams = 0f;
            switch (decimal_place)
            {
                case 0:
                    fweightKiloGrams = weight * 1.0f;
                    break;
                case 1:
                    fweightKiloGrams = weight * 0.1f;
                    break;
                case 2:
                    fweightKiloGrams = weight * 0.01f;
                    break;
                case 3:
                    fweightKiloGrams = weight * 0.001f;
                    break;
                default:
                    fweightKiloGrams = weight * 1.0f;
                    break;
            }
            PToolingDataEventArgs toolingData = new PToolingDataEventArgs(fweightKiloGrams, true);
            base.OnDetectDataReceived(sender, toolingData);
        }
        
        /// <summary>
        ///STX   n    D    X1   X2   X3   X4   X5   X6  BCC  CR 
        /// 格式：0x02 0x41 0x44 0x32 0x3C 0x35 0x32 0x30 0x38 0x72 0x0D => 1#模块 9666 16 进制为 025C2H=9666
        /// X1   X2   X3   X4   X5是数据位，存储的是16进制的字符，X1 为低位，X5 为高位
        /// 设置3位小数
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        public override void ReceiveData(object sender, DataTransmissionEventArgs args)
        {
            byte[] buffer = new byte[_detectByteLength];
            lock (m_ReadBuffer)
            {
                m_ReadBuffer.AddRange(args.EventData);
                if(m_ReadBuffer.Count>=_detectByteLength)
                {
                    int endIndex = m_ReadBuffer.FindIndex(0, (x) => { return x == 0x0D; });
                    if (endIndex<0)
                    {
                        return;
                    }
                    else if (endIndex > 0 && endIndex != _detectByteLength-1)
                    {
                        m_ReadBuffer.RemoveRange(0, endIndex + 1);
                        return;
                    }
                    else
                    {
                        m_ReadBuffer.CopyTo(0, buffer, 0, _detectByteLength);
                        m_ReadBuffer.RemoveRange(0, _detectByteLength);
                    }
                }
                else
                {
                    return;
                }
            }
            int weight = 0;
            for(int iLoop=7;iLoop>=3;iLoop--)
            {
                weight += (buffer[iLoop] - 0x30) << (iLoop-3)*4;
            }

            int decimal_place = buffer[8] & 0x03;
            float fweightKiloGrams = 0f;
            switch (decimal_place)
            {
                case 0:
                    fweightKiloGrams = weight * 1.0f;
                    break;
                case 1:
                    fweightKiloGrams = weight * 0.1f;
                    break;
                case 2:
                    fweightKiloGrams = weight * 0.01f;
                    break;
                case 3:
                    fweightKiloGrams = weight * 0.001f;
                    break;
                default:
                    fweightKiloGrams = weight * 1.0f;
                    break;
            }
            PToolingDataEventArgs toolingData = new PToolingDataEventArgs(fweightKiloGrams, true);
            base.ReceiveData(sender, toolingData);
        }
    }
}
