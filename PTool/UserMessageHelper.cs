using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace PTool
{
    public class UserMessageHelper
    {
        [DllImport("user32.dll", EntryPoint = "SendMessage")]
        public static extern int SendMessage(IntPtr hwnd, int wMsg, int wParam, int lParam);

        /// <summary>
        /// windowapi 找到指定窗体的句柄函数
        /// </summary>
        /// <param name="lpClassName">窗体类名</param>
        /// <param name="lpWindowName">窗体标题名</param>
        /// <returns>返回窗体句柄（IntPtr）</returns>
        [DllImport("user32.dll", EntryPoint = "FindWindow", SetLastError = true)]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        public static void SendClearPumpNoMessage()
        {
            IntPtr handle = FindWindow(null, "压力测试工具");
            if (handle != IntPtr.Zero)
                UserMessageHelper.SendMessage(handle, 0x1EE1, 0, 0);
        }
    }
}
