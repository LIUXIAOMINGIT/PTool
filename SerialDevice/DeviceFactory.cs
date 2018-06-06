using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SerialDevice
{
    public class DeviceFactory
    {
        private DeviceFactory()
        { }

        public static IDevice CreateDevice(DeviceType deviceType)
        {
            IDevice device = null;
            switch(deviceType)
            {
                case DeviceType.ACD_1F:
                    device = new PressureMeter();
                    break;
                case DeviceType.Balance_BSA:
                    device = new BalanceBSA();
                    break;
                case DeviceType.Balance_XS205:
                    device = new BalanceXS205();
                    break;
                case DeviceType.HHS:
                    device = new HHSTooling();
                    break;
                case DeviceType.C9:
                    device = new  GrasebyC9();
                    break;
                default:
                    device = null;
                    break;
            }
            return device;
        }
    }
}
