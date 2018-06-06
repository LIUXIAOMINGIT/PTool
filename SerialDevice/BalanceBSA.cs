using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO.Ports;

namespace SerialDevice
{
    public class BalanceBSA : DeviceBase
    {
        private List<byte> m_ReadBuffer = new List<byte>(); //存放数据缓存，如果数据到达数量少于指定长度，等待下次接受
        public BalanceBSA()
        {
            this._deviceType = DeviceType.Balance_BSA;
            _detectByteLength = 16;                                     //数据长度
            SetDetectBytes(new byte[] { 0x1B, 0x50, 0x0D, 0x0A });      //设置检测命令
            Init(1200, 7, StopBits.One, Parity.None, "");
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
            if (buffer.Count >= _detectByteLength && buffer[0] >= 0x20 && buffer[_detectByteLength - 1] == 0x0A)
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
            byte[] buffer = new byte[_detectByteLength];
            lock (m_ReadBuffer)
            {
                m_ReadBuffer.AddRange(args.EventData);
                if(m_ReadBuffer.Count>=_detectByteLength)
                {
                    int endIndex = m_ReadBuffer.FindIndex(0, (x) => { return x == 0x0A; });
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

            string weight = System.Text.Encoding.ASCII.GetString(buffer);
            if (weight.Length < 14)
                return;

            //get the balance unit
            string unitString = weight.Substring(11, 3);
            string unit = string.Empty;
            int iLoop = 0;
            while (iLoop < 3)
            {
                if (unitString[iLoop] != 0x20)
                    unit += unitString[iLoop];
                iLoop++;
            }
            WeightUint balanceUnit = WeightUint.g;
            foreach (string s in Enum.GetNames(typeof(WeightUint)))
            {
                if (s.Equals(unit))
                {
                    balanceUnit = (WeightUint)Enum.Parse(typeof(WeightUint), unit);
                    break;
                }
            }
            int negative = weight[0] == '-' ? -1 : 1;
            float weightValue = 0f;
            bool bRet = false;
            bRet = float.TryParse(weight.Substring(2, 8), out weightValue);
            float weightGrams = 0f;
            switch (balanceUnit)
            {
                case WeightUint.kg:
                    weightGrams = weightValue * 1000 * negative;
                    break;
                case WeightUint.mg:
                    weightGrams = weightValue / 1000 * negative;
                    break;
                case WeightUint.g:
                    weightGrams = weightValue * negative;
                    break;
                default:
                    weightGrams = weightValue * negative;
                    break;
            }
            BalanceDataEventArgs balanceData = new BalanceDataEventArgs(weightGrams, bRet);
            base.ReceiveData(sender, balanceData);
        }
    }
}
