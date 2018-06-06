using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SerialDevice
{
    public class BalanceDataEventArgs : DataTransmissionEventArgs
    {
        public bool IsValid { get; set; }
        public float Weight { get; set; }

        public BalanceDataEventArgs(BalanceDataEventArgs data)
            : base(null)
        {
            Weight = data.Weight;
            IsValid = data.IsValid;
            PortName = data.PortName;
        }

        public BalanceDataEventArgs(float weightValue, bool valid, string portName = "")
            : base(null)
        {
            Weight = weightValue;
            IsValid = valid;
            PortName = portName;
        }

    }


    public class PToolingDataEventArgs : DataTransmissionEventArgs
    {
        public bool IsValid { get; set; }
        public float Weight { get; set; }

        public PToolingDataEventArgs(PToolingDataEventArgs data)
            : base(null)
        {
            Weight = data.Weight;
            IsValid = data.IsValid;
            PortName = data.PortName;
        }

        public PToolingDataEventArgs(float weightValue, bool valid, string portName = "")
            : base(null)
        {
            Weight = weightValue;
            IsValid = valid;
            PortName = portName;
        }

    }

    public class Graseby9600DataEventArgs : DataTransmissionEventArgs
    {
        public bool IsValid { get; set; }
        public float SensorValue { get; set; }

        public Graseby9600DataEventArgs(Graseby9600DataEventArgs data)
            : base(null)
        {
            SensorValue = data.SensorValue;
            IsValid = data.IsValid;
            PortName = data.PortName;
        }

        public Graseby9600DataEventArgs(float sensorValue, bool valid, string portName = "")
            : base(null)
        {
            SensorValue = sensorValue;
            IsValid = valid;
            PortName = portName;
        }

    }
}
