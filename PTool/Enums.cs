using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PTool
{
    public enum ProductID : byte
    {
        GrasebyC6   = 0x04,
        GrasebyF6   = 0x05,
        GrasebyC6T  = 0x09,
        Graseby2000 = 0x0A,
        Graseby2100 = 0x0B,
        //WZ50C6      = 0x04,
        //WZS50F6     = 0x05,
        //WZ50C6T     = 0x09,
        None = 0xFF,
    }

    public enum PumpID : byte
    {
        GrasebyC6,
        GrasebyF6,
        GrasebyC6T,
        Graseby2000,
        Graseby2100,
        WZ50C6,
        WZS50F6,
        WZ50C6T,
        GrasebyF6_2, //F6双道
        WZS50F6_2, //WZS50F6双道
        None = 0xFF,
    }

    public class ProductIDConvertor
    {
        /// <summary>
        /// pumpID 转 成 ProductID
        /// </summary>
        /// <param name="pumpID"></param>
        /// <returns></returns>
        public static ProductID PumpID2ProductID(PumpID pumpID)
        {
            ProductID pid = ProductID.None;
            switch(pumpID)
            {
                case PumpID.GrasebyC6:
                    pid = ProductID.GrasebyC6;
                    break;
                case PumpID.GrasebyF6:
                    pid = ProductID.GrasebyF6;
                    break;
                case PumpID.GrasebyC6T:
                    pid = ProductID.GrasebyC6T;
                    break;
                case PumpID.Graseby2100:
                    pid = ProductID.Graseby2100;
                    break;
                case PumpID.WZ50C6:
                    pid = ProductID.GrasebyC6;
                    break;
                case PumpID.WZS50F6:
                    pid = ProductID.GrasebyF6;
                    break;
                case PumpID.WZ50C6T:
                    pid = ProductID.GrasebyC6T;
                    break;
                case PumpID.GrasebyF6_2:
                    pid = ProductID.GrasebyF6;
                    break;
                case PumpID.WZS50F6_2:
                    pid = ProductID.GrasebyF6;
                    break;
                default:
                    pid = ProductID.None;
                    break;
            }
            return pid;
        }

        /// <summary>
        /// 将泵类型字符串换成枚举
        /// </summary>
        /// <param name="strPumpID"></param>
        /// <returns></returns>
        public static PumpID String2PumpID(string strPumpID)
        {
            PumpID pid = PumpID.None;
            switch (strPumpID)
            {
                case "GrasebyC6":
                    pid = PumpID.GrasebyC6;
                    break;
                case "GrasebyF6单道":
                    pid = PumpID.GrasebyF6;
                    break;
                case "GrasebyC6T":
                    pid = PumpID.GrasebyC6T;
                    break;
                case "Graseby2100":
                    pid = PumpID.Graseby2100;
                    break;
                case "Graseby2000":
                    pid = PumpID.Graseby2000;
                    break;
                case "WZ50C6":
                    pid = PumpID.GrasebyC6;
                    break;
                case "WZS50F6单道":
                    pid = PumpID.GrasebyF6;
                    break;
                case "WZ50C6T":
                    pid = PumpID.GrasebyC6T;
                    break;
                case "GrasebyF6双道":
                    pid = PumpID.GrasebyF6_2;
                    break;
                case "WZS50F6双道":
                    pid = PumpID.WZS50F6_2;
                    break;
                default:
                    pid = PumpID.None;
                    break;
            }
            return pid;
        }


        /// <summary>
        /// 读所有自定义型号字符串
        /// </summary>
        /// <returns></returns>
        public static List<string> GetAllPumpIDString()
        {
            List<string> pumps = new List<string>();
            pumps.Add("GrasebyC6");
            pumps.Add("GrasebyC6T");
            pumps.Add("Graseby2000");
            pumps.Add("Graseby2100");
            pumps.Add("WZ50C6");
            pumps.Add("WZ50C6T");
            pumps.Add("GrasebyF6单道");
            pumps.Add("GrasebyF6双道");
            pumps.Add("WZS50F6单道");
            pumps.Add("WZS50F6双道");
            return pumps;
        }





    }

//    C6 读取当前压力传感器命令
//55 aa 05 04 01 70 00 86

//F6 读取当前压力传感器命令
//55 aa 05 05 01 70 00 85

//C6T 读取当前压力传感器命令
//55 aa 05 09 01 70 00 81

//2000 读取当前压力传感器命令
//55 aa 05 0A 01 70 00 80

//2100 读取当前压力传感器命令
//55 aa 05 0B 01 70 00 7F


    //返回11个字节，取ID为7，8，9三个字节
}
