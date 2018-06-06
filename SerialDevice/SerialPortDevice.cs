using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Ports;

namespace SerialDevice
{
    public class SerialPortDevice
    {
        private readonly int        BAUDRATE = 9600;
        private readonly int        DATABITS = 8;               
        private readonly StopBits   STOPBITS = StopBits.One;
        private readonly Parity     PARITY = Parity.None;
        private readonly Handshake  HANDSHAKE=Handshake.None; 
        private SerialPort          _connectedSerialPort = null;
        
        public event EventHandler<DataTransmissionEventArgs> DataReceived;

        /// <summary>
        /// Default constructor
        /// </summary>
        public SerialPortDevice()
        {
        }

        /// <summary>
        /// use serial port and set the port name
        /// </summary>
        /// <param name="portName">serial port name</param>
        public void InitializeDevice(string portName)
        {
            InitializeDevice(portName, BAUDRATE, DATABITS, STOPBITS, PARITY, HANDSHAKE);
        }

        /// <summary>
        /// use serial port, set the port name and baud rate
        /// </summary>
        /// <param name="portName">serial port name</param>
        /// <param name="baudRate">serial port baud rate</param>
        public void InitializeDevice(string portName, int baudRate)
        {
            InitializeDevice(portName, baudRate, DATABITS, STOPBITS, PARITY, HANDSHAKE);
        }

        /// <summary>
        /// use serial port, set the port name,baud rate,data bits,stopbits,parity
        /// </summary>
        /// <param name="portName">serial port name</param>
        /// <param name="baudRate">serial port baud rate</param>
        /// <param name="dataBits">serial port data bits</param>
        /// <param name="StopBits">serial port stopbits</param>
        /// <param name="Parity">serial port parity</param>
        public void InitializeDevice(string portName, int baudRate, int dataBits, StopBits stopBits, Parity parity)
        {
            InitializeDevice(portName, baudRate, dataBits, stopBits, parity, HANDSHAKE);
        }

        /// <summary>
        /// use serial port, set the port name,baud rate,data bits,stopbits,parity
        /// </summary>
        /// <param name="portName">serial port name</param>
        /// <param name="baudRate">serial port baud rate</param>
        /// <param name="dataBits">serial port data bits</param>
        /// <param name="StopBits">serial port stopbits</param>
        /// <param name="Parity">serial port parity</param>
        /// <param name="handshake">handshake type(hardware )</param>
        public void InitializeDevice(string portName, int baudRate, int dataBits, StopBits stopBits, Parity parity, Handshake handshake)
        {
            ResetDevice();
            if (null != _connectedSerialPort)
            {
                _connectedSerialPort.PortName = portName;
                _connectedSerialPort.BaudRate = baudRate;
                _connectedSerialPort.DataBits = dataBits;
                _connectedSerialPort.StopBits = stopBits;
                _connectedSerialPort.Parity = parity;
                _connectedSerialPort.Handshake = handshake;
            }
        }

        /// <summary>
        /// use serial port and set port information
        /// </summary>
        /// <param name="devInfo">serial port to use</param>
        public void InitializeDevice(SerialPort devInfo)
        {
            ResetDevice();
            if (null != _connectedSerialPort)
            {
                _connectedSerialPort = devInfo;
            }
        }

        /// <summary>
        /// set device type with SerialPort and close other devices
        /// </summary>
        private void ResetDevice()
        {
            try
            {
                Close();
            }
            catch (Exception)
            {
            }
            if (null == _connectedSerialPort)
            {
                _connectedSerialPort = new SerialPort();
            }
        }

        /// <summary>
        /// The number of bytes in the internal input buffer
        /// before a System.IO.Ports.SerialPort.DataReceived event is fired
        /// </summary>
        /// <param name="len"></param>
        public void SetReceivedBytesThreshold(int len = 1)
        {
            if (_connectedSerialPort != null)
                _connectedSerialPort.ReceivedBytesThreshold = len;
        }

        public void Open()
        {
            if (null != _connectedSerialPort)
            {
                if (!_connectedSerialPort.IsOpen)
                {
                    try
                    { 
                        _connectedSerialPort.DataReceived += SerialPortDataReceived;
                        _connectedSerialPort.Open();
                    }
                    catch (Exception e)
                    {
                        throw new Exception("Open - ", e);
                    }
                }
            }
        }

        public void Close()
        {
            if (null != _connectedSerialPort)
            {
                _connectedSerialPort.DataReceived -= SerialPortDataReceived;
                if (_connectedSerialPort.IsOpen)
                {
                    try
                    {
                        _connectedSerialPort.DiscardInBuffer();
                        _connectedSerialPort.DiscardOutBuffer();
                        _connectedSerialPort.Close();
                    }
                    catch (Exception e)
                    {
                        throw new Exception("Close - ", e);
                    }
                }
            }
        }

        /// <summary>
        /// Gets a value that indicates whether the connected device is open 
        /// </summary>
        /// <returns>true if the device is open; otherwise,false</returns>
        public bool IsOpen()
        {
            bool isOpen = false;
            if (null != _connectedSerialPort)
            {
                isOpen = _connectedSerialPort.IsOpen;
            }
            return isOpen;
        }

        public void SendData(byte[] data)
        {
            if (_connectedSerialPort.IsOpen)
            {
                try
                {
                    // Writer raw data
                    this._connectedSerialPort.Write(data, 0, data.Length);
                }
                catch (Exception e)
                {
                    throw new Exception("Write - ", e);
                }
            }
            else
            {
                throw new Exception("Write - Port is not open");
            }
        }

        private void SerialPortDataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            try
            {
                //read from stream
                int n = _connectedSerialPort.BytesToRead;
                byte[] buff = new byte[n];
                _connectedSerialPort.Read(buff, 0, n);
                //save to the buffer
                if (DataReceived != null)
                {
                    DataReceived(this, new DataTransmissionEventArgs(buff, _connectedSerialPort.PortName));
                }

            }
            catch (Exception exception)
            {
                throw new Exception(exception.Message, exception);
            }
        }       
    }

}
