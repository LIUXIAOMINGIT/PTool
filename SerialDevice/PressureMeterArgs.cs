using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SerialDevice
{
    public class PressureMeterArgs : DataTransmissionEventArgs
    {
        public PressureUnit Unit { get; set; }
        public float PressureValue { get; set; }
        public string ErrorMessage { get; set; }

        public PressureMeterArgs(PressureMeterArgs data)
            :base(null)
        {
            this.Unit = data.Unit;
            this.PressureValue = data.PressureValue;
            this.ErrorMessage = data.ErrorMessage;
        }

        public PressureMeterArgs(PressureUnit unit, float pressureValue, string errorMessage = "")
            : base(null)
        {
            this.Unit = unit;
            this.PressureValue = pressureValue;
            this.ErrorMessage = errorMessage;
        }

    }
}
