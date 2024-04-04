using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Flit;
using FlaUI.Core;
using FlaUI.UIA3;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Input;
using FlaUI.Core.WindowsAPI;
using System.Runtime.Serialization;
using System.Timers;
using Microsoft.Win32;

namespace StaffTester
{   
    [TestFixture (1 , "AttendanceTracker")]
    internal class StaffAttendance
    {

        [Test (1, "Launch SMART FINGER & Scan")]

        internal static void schedular () {
            var aTimer = new System.Timers.Timer (60 * 1000);
            aTimer.Elapsed += new ElapsedEventHandler (Scan);
            aTimer.Start ();
        }

        public static void Scan (Object sender, ElapsedEventArgs e) {
            var auto = new UIA3Automation ();
            Utils.OpenSF ();
            Assert.IsNotNull (Utils.Smart, "SMART FINGER did not open");
            var window = Utils.Smart.GetMainWindow (auto);
            Thread.Sleep (5000);
            Keyboard.Type (VirtualKeyShort.KEY_S);
            var selectAll = FlaUI.Core.Tools.Retry.Find (() => { return window.FindFirstDescendant (cf => cf.ByName ("Select All")).AsButton (); }, Utils.RetrySettings ());
            selectAll.Click ();
            var startScan = FlaUI.Core.Tools.Retry.Find (() => { return window.FindFirstDescendant (cf => cf.ByName ("Start Scan")).AsButton (); }, Utils.RetrySettings ());
            startScan.Click ();
            Thread.Sleep (3000);
            Keyboard.Type (VirtualKeyShort.KEY_C);
            selectAll.Click ();
            startScan.Click ();
            Thread.Sleep (3000);
        }

        [Test (2, "Launch Staff Attendance")]

        public void AttReport () {
            var auto = new UIA3Automation ();
            Utils.OpenStaffAtt ();
            Assert.IsNotNull (Utils.Att, "Staff Attendance did not open");
            var window = Utils.Att.GetMainWindow (auto);
            Thread.Sleep (5000);
            var selectCompany = FlaUI.Core.Tools.Retry.Find (() => { return window.FindFirstDescendant (cf => cf.ByName ("TRUMPF METAMATION VISITOR AND CONTRACTORS").And (cf.ByControlType (FlaUI.Core.Definitions.ControlType.ListItem))); }, Utils.RetrySettings ());    
            selectCompany.Click ();
            var clickNext = FlaUI.Core.Tools.Retry.Find(()=> { return window.FindFirstDescendant (cf => cf.ByName ("Next")); },Utils.RetrySettings ()); 
            clickNext.Click ();
            Keyboard.Type ("smart");
            Keyboard.Press (VirtualKeyShort.TAB);
            Keyboard.Type ("smart");
            var okButton = FlaUI.Core.Tools.Retry.Find (() => { return window.FindFirstDescendant (cf => cf.ByAutomationId("butok")).AsButton(); }, Utils.RetrySettings ());
            okButton.Click ();
            Thread.Sleep (5000);
            Keyboard.TypeSimultaneously (VirtualKeyShort.ALT, VirtualKeyShort.KEY_M);
            Keyboard.Type (VirtualKeyShort.KEY_T);
            var AMSDialoge = FlaUI.Core.Tools.Retry.Find (() => { return window.FindFirstDescendant (cf => cf.ByName ("AMS")); }, Utils.RetrySettings()).Parent;
            AMSDialoge.FindFirstChild(cf => cf.ByName("OK")).Click ();
            var invalidDialoge = FlaUI.Core.Tools.Retry.Find (() => { return window.FindFirstDescendant (cf => cf.ByName ("Invalid Time File")); }, Utils.RetrySettings ()).Parent;
            invalidDialoge.FindFirstChild (cf => cf.ByName ("OK")).Click ();

            //Report
            Keyboard.TypeSimultaneously (VirtualKeyShort.ALT, VirtualKeyShort.KEY_R);
            var reportTree = FlaUI.Core.Tools.Retry.Find (() => { return window.FindFirstDescendant (cf => cf.ByName ("Staff Reports").And(cf.ByLocalizedControlType("tree item"))); }, Utils.RetrySettings ()).Parent;
            reportTree.FindFirstChild (cf => cf.ByName ("tree")).Click ();

            var clickDaily = FlaUI.Core.Tools.Retry.Find (() => { return window.FindFirstDescendant (cf => cf.ByName ("Daily").And (cf.ByLocalizedControlType ("tree item"))); }, Utils.RetrySettings ()).Parent;
            clickDaily.Click ();

            var clickAttendance = FlaUI.Core.Tools.Retry.Find (() => { return window.FindFirstDescendant (cf => cf.ByName ("Attendance").And (cf.ByLocalizedControlType ("tree item"))); }, Utils.RetrySettings ()).Parent;
            clickAttendance.Click ();

            var clickToday = FlaUI.Core.Tools.Retry.Find (() => { return window.FindFirstDescendant (cf => cf.ByName ("Today").And (cf.ByLocalizedControlType ("radio button"))); }, Utils.RetrySettings ()).Parent;
            clickAttendance.Click ();

            var clickAll = FlaUI.Core.Tools.Retry.Find (() => { return window.FindFirstDescendant (cf => cf.ByName ("All").And (cf.ByLocalizedControlType ("radio button"))); }, Utils.RetrySettings ()).Parent;
            clickAttendance.Click ();

            AMSDialoge.FindFirstChild (cf => cf.ByName ("OK")).Click ();
            Thread.Sleep (5000);

            var expReport = FlaUI.Core.Tools.Retry.Find (() => { return window.FindFirstDescendant (cf => cf.ByName ("Export Report").And (cf.ByLocalizedControlType ("button"))); }, Utils.RetrySettings ()).Parent;
            clickAttendance.Click ();

            var saveReport = FlaUI.Core.Tools.Retry.Find (() => { return window.FindFirstDescendant (cf => cf.ByName ("Export Report").And(cf.ByLocalizedControlType("dialog"))); }, Utils.RetrySettings ()).Parent;
            saveReport.FindFirstChild (cf => cf.ByName ("Address").And(cf.ByLocalizedControlType("edit"))).Click();
            Keyboard.Type (@"D:\OneDrive - Trumpf Metamation Pvt Ltd\Work");
            Keyboard.TypeSimultaneously (VirtualKeyShort.ALT, VirtualKeyShort.KEY_N);
            Keyboard.Type ("AttendanceReport" + "_" + DateTime.Now.ToString ("dd.MM.yyyy HH:mm:ss"));
            Keyboard.TypeSimultaneously (VirtualKeyShort.ALT, VirtualKeyShort.KEY_T);
            Keyboard.Type ("Adobe");
            Keyboard.TypeSimultaneously (VirtualKeyShort.ALT, VirtualKeyShort.KEY_S);

            var expDialog = FlaUI.Core.Tools.Retry.Find (() => { return window.FindFirstDescendant (cf => cf.ByName ("Export Report").And (cf.ByLocalizedControlType ("dialog"))); }, Utils.RetrySettings ()).Parent;
            expDialog.FindFirstChild (cf => cf.ByName ("OK")).Click ();
        }

    }
}
