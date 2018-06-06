using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SerialDevice
{
    /// <summary>
    /// 定义各种串口设备类型
    /// </summary>
    public enum DeviceType
    {
        None = 0,
        Balance_BSA = 1,    //+ * D D D D D D D D * U U U CR LF 16位,22位不支持,
        Balance_XS205,      //实验室及检测所天平
        Balance_YP,         //越平天平，主动发送数据，间隔有30s,60s不等，还有连续发送，具体时间不详
        ACD_1F,             //安森压力表ACD-1F
        HHS,                //HHS工装
        C9,                 //C9
        DL101,              //DL101
        GrasebyC6,  
        GrasebyF6,  
        GrasebyC6T, 
        Graseby2000,
        Graseby2100,
        WZ50C6,     
        WZS50F6,    
        WZ50C6T,    
    }

    /// <summary>
    /// 暂时只支持g\mg\kg三种单位
    /// </summary>
    public enum WeightUint
    {
        none = 0,
        g    = 1,
        mg   = 2,
        kg   = 3,
        ct   = 4,//克拉
        oz   = 5,//盎司 1盎司=31.1035克
        ozt  = 6,//金盎司 1金盎司=31.1034768克
    }

    public enum PressureUnit
    {
        m = 0,
        KPa = 1,
        MPa = 2,
        C = 3,      //℃
        mA = 4,
        A = 5,
        V = 6,
        Other = 7,
    }
}
