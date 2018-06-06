using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Runtime.InteropServices;
using Misc = ComunicationProtocol.Misc;

namespace PTool
{
    public class IniReader
    {
        private string _fileName;

        [DllImport("kernel32")]
        private static extern int GetPrivateProfileInt(
          string lpAppName,// 指向包含 Section 名称的字符串地址
          string lpKeyName,// 指向包含 Key 名称的字符串地址
          int nDefault,// 如果 Key 值没有找到，则返回缺省的值是多少
          string lpFileName
          );

        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(
          string lpAppName,// 指向包含 Section 名称的字符串地址
          string lpKeyName,// 指向包含 Key 名称的字符串地址
          string lpDefault,// 如果 Key 值没有找到，则返回缺省的字符串的地址
          StringBuilder lpReturnedString,// 返回字符串的缓冲区地址
          int nSize,// 缓冲区的长度
          string lpFileName
          );

        [DllImport("kernel32")]
        private static extern bool WritePrivateProfileString(
          string lpAppName,// 指向包含 Section 名称的字符串地址
          string lpKeyName,// 指向包含 Key 名称的字符串地址
          string lpString,// 要写的字符串地址
          string lpFileName
          );

        public IniReader(string filename)
        {
            _fileName = filename;
        }

        public int GetInt(string section, string key, int def)
        {
            return GetPrivateProfileInt(section, key, def, _fileName);
        }

        public string GetString(string section, string key, string def = "")
        {
            StringBuilder temp = new StringBuilder(1024);
            GetPrivateProfileString(section, key, def, temp, 1024, _fileName);
            return temp.ToString();
        }

        public void ReadSettings()
        {
            ReadGrasebyPumpPressureSettings(PumpID.GrasebyC6, Misc.OcclusionLevel.L);
            ReadGrasebyPumpPressureSettings(PumpID.GrasebyC6, Misc.OcclusionLevel.C);
            ReadGrasebyPumpPressureSettings(PumpID.GrasebyC6, Misc.OcclusionLevel.H);

            ReadGrasebyPumpPressureSettings(PumpID.Graseby2000, Misc.OcclusionLevel.L);
            ReadGrasebyPumpPressureSettings(PumpID.Graseby2000, Misc.OcclusionLevel.C);
            ReadGrasebyPumpPressureSettings(PumpID.Graseby2000, Misc.OcclusionLevel.H);

            ReadGrasebyPumpPressureSettings(PumpID.Graseby2100, Misc.OcclusionLevel.L);
            ReadGrasebyPumpPressureSettings(PumpID.Graseby2100, Misc.OcclusionLevel.C);
            ReadGrasebyPumpPressureSettings(PumpID.Graseby2100, Misc.OcclusionLevel.H);

            ReadGrasebyPumpPressureSettings(PumpID.GrasebyC6T, Misc.OcclusionLevel.L);
            ReadGrasebyPumpPressureSettings(PumpID.GrasebyC6T, Misc.OcclusionLevel.C);
            ReadGrasebyPumpPressureSettings(PumpID.GrasebyC6T, Misc.OcclusionLevel.H);

            ReadGrasebyPumpPressureSettings(PumpID.WZ50C6, Misc.OcclusionLevel.L);
            ReadGrasebyPumpPressureSettings(PumpID.WZ50C6, Misc.OcclusionLevel.C);
            ReadGrasebyPumpPressureSettings(PumpID.WZ50C6, Misc.OcclusionLevel.H);

            ReadGrasebyPumpPressureSettings(PumpID.WZ50C6T, Misc.OcclusionLevel.L);
            ReadGrasebyPumpPressureSettings(PumpID.WZ50C6T, Misc.OcclusionLevel.C);
            ReadGrasebyPumpPressureSettings(PumpID.WZ50C6T, Misc.OcclusionLevel.H);

            ReadGrasebyPumpPressureSettings(PumpID.GrasebyF6, Misc.OcclusionLevel.L);
            ReadGrasebyPumpPressureSettings(PumpID.GrasebyF6, Misc.OcclusionLevel.C);
            ReadGrasebyPumpPressureSettings(PumpID.GrasebyF6, Misc.OcclusionLevel.H);

            ReadGrasebyPumpPressureSettings(PumpID.GrasebyF6_2, Misc.OcclusionLevel.L);
            ReadGrasebyPumpPressureSettings(PumpID.GrasebyF6_2, Misc.OcclusionLevel.C);
            ReadGrasebyPumpPressureSettings(PumpID.GrasebyF6_2, Misc.OcclusionLevel.H);

            ReadGrasebyPumpPressureSettings(PumpID.WZS50F6, Misc.OcclusionLevel.L);
            ReadGrasebyPumpPressureSettings(PumpID.WZS50F6, Misc.OcclusionLevel.C);
            ReadGrasebyPumpPressureSettings(PumpID.WZS50F6, Misc.OcclusionLevel.H);

            ReadGrasebyPumpPressureSettings(PumpID.WZS50F6_2, Misc.OcclusionLevel.L);
            ReadGrasebyPumpPressureSettings(PumpID.WZS50F6_2, Misc.OcclusionLevel.C);
            ReadGrasebyPumpPressureSettings(PumpID.WZS50F6_2, Misc.OcclusionLevel.H);
        }

        public void ReadGrasebyPumpPressureSettings(PumpID pid, Misc.OcclusionLevel level)
        {
            string section = string.Empty;

            #region section
            switch (pid)
            {
                case PumpID.GrasebyC6:
                    section = "GrasebyC6";
                    break;
                case PumpID.Graseby2000:
                    section = "Graseby2000";
                    break;
                case PumpID.Graseby2100:
                    section = "Graseby2100";
                    break;
                case PumpID.GrasebyC6T:
                    section = "GrasebyC6T";
                    break;
                case PumpID.WZ50C6:
                    section = "WZ50C6";
                    break;
                case PumpID.WZ50C6T:
                    section = "WZ50C6T";
                    break;
                case PumpID.GrasebyF6:
                    section = "GrasebyF6_1";
                    break;
                case PumpID.GrasebyF6_2:
                    section = "GrasebyF6_2";
                    break;
                case PumpID.WZS50F6:
                    section = "WZS50F6_1";
                    break;
                case PumpID.WZS50F6_2:
                    section = "WZS50F6_2";
                    break;
                default:
                    section = "GrasebyC6";
                    break;
            }
            #endregion

            string val = GetString(section, level.ToString()).Trim();
            string[] arrSyringeSizePValue = val.Split('|');
            if (arrSyringeSizePValue.Length == 0)
                return;
            int size = 0;
            float p = 0;
            foreach (string pair in arrSyringeSizePValue)
            {
                string[] arrSizePValuePair = pair.Trim().Split(',');
                if (arrSizePValuePair.Length != 2)
                    continue;
                if (int.TryParse(arrSizePValuePair[0].Trim(), out size) && float.TryParse(arrSizePValuePair[1].Trim(), out p))
                {
                    PressureManager.Instance().Add(pid, level, size, p);
                }
            }

        }



    }
}