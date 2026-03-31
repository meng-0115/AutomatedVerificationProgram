using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace AutomatedVerificationProgram
{
    internal class ClickFnctionClass
    {
        public static class MouseWin32
        {
            [DllImport("user32.dll")]
            static extern void mouse_event(int dwFlags, int dx, int dy, int dwData, int dwExtraInfo);

            [DllImport("user32.dll")]
            static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, uint dwExtraInfo);

            const int MOUSEEVENTF_LEFTDOWN = 0x02;
            const int MOUSEEVENTF_LEFTUP = 0x04;
            const int MOUSEEVENTF_RIGHTDOWN = 0x08;
            const int MOUSEEVENTF_RIGHTUP = 0x10;
            const int KEYEVENTF_KEYDOWN = 0x00;
            const int KEYEVENTF_KEYUP = 0x02;
            const byte VK_CONTROL = 0x11;
            const byte VK_V = 0x56;
            const byte vKeyA = 0x41;
           

            public static void Click(string type, int x = 0, int y = 0)
            {
                // 1. 移動滑鼠到座標
                System.Windows.Forms.Cursor.Position = new Point(x, y);

                // 2. 根據類型觸發點擊
                if (type == "左鍵單下點擊")
                {
                    mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
                }                  
                else if (type == "左鍵單下點擊")
                {
                    mouse_event(MOUSEEVENTF_RIGHTDOWN | MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0);
                }
                else if(type == "左鍵雙下點擊")
                {
                    mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
                    mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
                }
                else if( type == "右鍵雙下點擊")
                {
                    mouse_event(MOUSEEVENTF_RIGHTDOWN | MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0);
                    mouse_event(MOUSEEVENTF_RIGHTDOWN | MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0);
                }                
            }
            public static void Click(string type, int param )
            {
                if (type == "等待時間")
                {
                    System.Threading.Thread.Sleep(param);
                }
            }
            public static void Click(string type)
            {
                if (type == "貼上文字")
                {
                    keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYDOWN, 0); // 按下 Ctrl
                    keybd_event(vKeyA, 0, KEYEVENTF_KEYDOWN, 0);      // 按下 A
                    keybd_event(vKeyA, 0, KEYEVENTF_KEYUP, 0);        // 放開 A
                    keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0);   // 放開 Ctrl

                    keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYDOWN, 0); // 按下 Ctrl
                    keybd_event(VK_V, 0, KEYEVENTF_KEYDOWN, 0);       // 按下 V
                    keybd_event(VK_V, 0, KEYEVENTF_KEYUP, 0);         // 放開 V
                    keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0);   // 放開 Ctrl
                }
                else if(type == "全選文字")
                {
                   
                }
            }
            
        }
    }
}
