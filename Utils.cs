using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using FlaUI.Core;
using FlaUI.Core.WindowsAPI;
using FlaUI.UIA3;
namespace StaffTester
{
    internal static class Utils
    {
        internal static string smartPath = @"D:\smartconnect\Compro 7\CP7Clnt.exe";
        internal static string attendancePath = @"\\SYS532\Smart_Staff_Att\Ams\AMS.exe";
        internal static string tempPath = @"c:\mm\staffattendance\temp\";
        internal static Application? Smart = null;
        internal static Application? Att = null;

        internal static void OpenSF()
        {
            if (Smart == null)
            {
                Smart = Application.Launch(smartPath);
                Smart.WaitWhileBusy();
            }
        }

        internal static void OpenStaffAtt() 
        {
            if(Att == null) 
            {
                Att = Application.Launch(attendancePath);
                Att.WaitWhileBusy();
            }

        }

        internal static RetrySettings RetrySettings()
        {
            var retrySettings = new RetrySettings();
            retrySettings.Timeout = TimeSpan.FromSeconds(5);
            retrySettings.Interval = TimeSpan.FromSeconds(0.2);
            retrySettings.ThrowOnTimeout = true;
            return retrySettings;
        }
                
        internal static void CloseSmart()
        {
            if (Smart != null)
            {
                Smart.Kill(); Smart = null;
            }
        }
                    
    }
    
}
