using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO.Ports;

namespace SerialDevice
{
    public class Graseby9600 : DeviceBase
    {
        private List<byte> m_ReadBuffer = new List<byte>(); //存放数据缓存，如果数据到达数量少于指定长度，等待下次接受
        public Graseby9600()
        {
            this._deviceType = DeviceType.GrasebyC6;
            _detectByteLength = 11;
            Init(9600, 8, StopBits.One, Parity.None, "");
        }

        public void SetDeviceType(DeviceType devType)
        {
            _deviceType = devType;
        }

        public override void Get()
        {
            switch (_deviceType)
            {
                case DeviceType.GrasebyC6:
                    _detectCommandBytes = new byte[] { 0x55, 0xAA, 0x05, 0x04, 0x01, 0x70, 0x00, 0x86 };
                    break;
                case DeviceType.GrasebyF6:
                    _detectCommandBytes = new byte[] { 0x55, 0xAA, 0x05, 0x05, 0x01, 0x70, 0x00, 0x85 };
                    break;
                case DeviceType.GrasebyC6T:
                    _detectCommandBytes = new byte[] { 0x55, 0xAA, 0x05, 0x09, 0x01, 0x70, 0x00, 0x81 };
                    break;
                case DeviceType.Graseby2000:
                    _detectCommandBytes = new byte[] { 0x55, 0xAA, 0x05, 0x0A, 0x01, 0x70, 0x00, 0x80 };
                    break;
                case DeviceType.Graseby2100:
                    _detectCommandBytes = new byte[] { 0x55, 0xAA, 0x05, 0x0B, 0x01, 0x70, 0x00, 0x7F };
                    break;
                case DeviceType.WZ50C6:
                    _detectCommandBytes = new byte[] { 0x55, 0xAA, 0x05, 0x04, 0x01, 0x70, 0x00, 0x86 };
                    break;
                case DeviceType.WZS50F6:
                    _detectCommandBytes = new byte[] { 0x55, 0xAA, 0x05, 0x05, 0x01, 0x70, 0x00, 0x85 };
                    break;
                case DeviceType.WZ50C6T:
                    _detectCommandBytes = new byte[] { 0x55, 0xAA, 0x05, 0x09, 0x01, 0x70, 0x00, 0x81 };
                    break;
                default:
                    _detectCommandBytes = new byte[] { 0x55, 0xAA, 0x05, 0x04, 0x01, 0x70, 0x00, 0x86 };
                    break;
            }
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
            
        }
        
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        public override void ReceiveData(object sender, DataTransmissionEventArgs args)
        {
            byte[] buffer = new byte[_detectByteLength];
            lock (m_ReadBuffer)
            {
                m_ReadBuffer.AddRange(args.EventData);
                if (m_ReadBuffer.Count >= _detectByteLength)
                {
                    int headIndex = m_ReadBuffer.FindIndex(0, (x) => { return x == 0x55; });
                    if (headIndex < 0)
                    {
                        m_ReadBuffer.Clear();
                        return;
                    }
                    else if (headIndex > 0)
                    {
                        m_ReadBuffer.RemoveRange(0, headIndex);
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
            ushort sensorValue = buffer[7];
            sensorValue += (ushort)(buffer[8] << 8);
            byte decimal_place = buffer[9];
            float fSensorValue = 0f;
            switch (decimal_place)
            {
                case 0:
                    fSensorValue = sensorValue * 1.0f;
                    break;
                case 1:
                    fSensorValue = sensorValue * 0.1f;
                    break;
                case 2:
                    fSensorValue = sensorValue * 0.01f;
                    break;
                case 3:
                    fSensorValue = sensorValue * 0.001f;
                    break;
                default:
                    fSensorValue = sensorValue * 1.0f;
                    break;
            }
            Graseby9600DataEventArgs toolingData = new Graseby9600DataEventArgs(fSensorValue, true);
            base.ReceiveData(sender, toolingData);
        }
    }
}
