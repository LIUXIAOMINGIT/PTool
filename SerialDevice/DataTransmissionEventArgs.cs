using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SerialDevice
{
    /// <summary>
    /// Event arguments containing event data.
    /// </summary>
    public class DataTransmissionEventArgs : EventArgs
    {
        protected byte[] data;

        public byte[] EventData
        {
            get { return this.data; }
            set { this.data = value; }
        }

        private string m_PortName = string.Empty;
        /// <summary>
        /// 返回当前端口号
        /// </summary>
        public string PortName
        {
            get { return m_PortName; }
            set { m_PortName = value; }
        }

        /// <summary>
        ///Copies the data starting at the specified index and paste them to the inner array
        /// </summary>
        /// <param name="result">Data raised in the event.</param>
        /// <param name="index">the index in the sourceArray at which copying begins.</param>
        /// <param name="length"> the number of elements to copy</param>
        public DataTransmissionEventArgs(byte[] result, int index, int length, string portName = "")
        {
            m_PortName = portName;
            if (result != null)
            {
                data = new byte[length];
                Array.Copy(result, index, data, 0, length);
            }
        }

        /// <summary>
        /// Copies the data to the inner array 
        /// </summary>
        /// <param name="result">Data raised in the event.</param>
        public DataTransmissionEventArgs(byte[] result, string portName = "")
        {
            if (result != null)
            {
                data = new byte[result.Length];
                Array.Copy(result, data, result.Length);
            }
            m_PortName = portName;
        }
         
    }
}
