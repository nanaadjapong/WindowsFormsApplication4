using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Automation;
using System.Windows.Input;
using WindowsInput.Native;
using WatiN.Core.Native.Windows;
using WatiN.Core;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using mshtml;
using SHDocVw;


namespace WindowsFormsApplication4
{
    public static class Utility
    {
        public static IE Browser { get; set; }

        // Wait specified number of seconds
        public static void Wait(int seconds)
        {
            System.Threading.Thread.Sleep(seconds * 1000);
        }

        // Wait for condition to evaluate true, timeout after 30 seconds
        public static void Wait(Func<bool> condition)
        {
            int count = 0;

            while (!condition() && count < 30)
            {
                System.Threading.Thread.Sleep(1000);
                count++;
            }
        }

        //Send tab key press to browser
        public static void PressTab()
        {
            System.Windows.Forms.SendKeys.SendWait("{TAB}");
            System.Threading.Thread.Sleep(300);
        }

        //Send specified key press to browser
        public static void PressKey(string keyname)
        {
            System.Windows.Forms.SendKeys.SendWait("{" + keyname.ToUpper() + "}");
            System.Threading.Thread.Sleep(300);
        }

    }
}
