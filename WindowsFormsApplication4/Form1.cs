using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
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
using System.Timers;


namespace WindowsFormsApplication4
{
    public partial class Form1 : System.Windows.Forms.Form
    {
        public static System.Timers.Timer aTimer;

        IE browser;

        [DllImport("user32.dll", EntryPoint = "GetWindowText",
        ExactSpelling = false, CharSet = CharSet.Auto, SetLastError = true)]
        
        static extern IntPtr FindWindowByCaption(IntPtr parent, string strWindowTitle);



        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        public Form1()
        {
            InitializeComponent();

            
            
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //STEP 1
            webbMain.Navigate("http://wgp-selma-ap1:9080/LogAnalyzer/posEntry.jsp"); // THIS WEBSITE WONT WORK ON YOURS BECAUSE ITS A PRIVATE WEBSITE

            DateTime end = DateTime.Now;

            var start = end.AddDays(-1);

            textBox1.Text = start.ToString("MM/dd/yy");
            textBox2.Text = end.ToString("MM/dd/yy");
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            StringBuilder popUp = new StringBuilder();
            popUp.Append("http://wgp-selma-ap1:9080/LogAnalyzer/posReport.jsp?dcnt=2&schema=prod&sdate=10/15/2018&edate=10/16/2018&cycle=0&version=0&bu=1&rtype=1&pda=0&lcnt=46&loc=0:3:207.8600:1:2:@:0:1:258.4400:1:2:@:0:1:327.2600:1:2:@:0:1:389.2000:1:2:@:0:1:467.9700:1:2:@:0:1:517.2700:1:2:@:0:1:588.6200:1:2:@:0:1:626.5900:1:2:@:0:1:661.6600:1:2:@:0:1:740.5000:1:2:@:0:1:784.6700:1:2:@:0:1:811.1300:1:2:@:0:1:850.9700:1:2:@:0:1:890.6100:1:2:@:0:1:926.1000:1:2:@:0:1:969.5000:1:2:@:0:1:1007.7500:1:2:@:0:1:1048.3400:1:2:@:0:1:1085.9700:1:2:@:0:1:1124.7100:1:2:@:0:1:1164.8500:1:2:@:0:1:1205.8900:1:2:@:0:1:1247.0900:1:2:@:0:1:1287.1100:1:2:@:0:1:1326.1100:1:2:@:0:1:1369.4400:1:2:@:0:1:1413.0000:1:2:@:0:1:1457.9700:1:2:@:0:1:1499.3600:1:2:@:0:1:1540.3700:1:2:@:0:1:1583.3700:1:2:@:0:1:1628.7800:1:2:@:0:1:1674.5500:1:2:@:0:1:1722.2000:1:2:@:0:1:1773.4000:1:2:@:0:2:205.0000:1:2:@:0:2:170.3900:1:2:@:0:2:127.9500:1:2:@:0:2:81.7100:1:2:@:0:2:12.7600:1:2:@:0:1:1776.8000:1:2:@:0:550:1.0000:1:2:@:0:551:1.0000:1:2:@:0:10:122.6800:1:2:@:0:10:68.4000:1:2:@:0:10:1.0000:1:2");

            var p = popUp.ToString();

            //Step 2 clicks on load

            var buttonElem = webbMain.Document.GetElementsByTagName("input");
            foreach (HtmlElement b in buttonElem)
            {
                if (b.GetAttribute("value").Equals("Load"))
                {
                    b.InvokeMember("click");
                }
            }


            // STEP 3 CHANGES THE DATES..........
            var inputElements = webbMain.Document.GetElementsByTagName("input");
            foreach (HtmlElement i in inputElements)
            {
                if (i.GetAttribute("name").Equals("sflowdate"))
                {
                    i.InnerText = textBox1.Text;

                }
                if (i.GetAttribute("name").Equals("eflowdate"))
                {
                    i.InnerText = textBox2.Text;
                }
            }



            //Step 4 clicks on " Dislay  Available Information"

            var buttonElem2 = webbMain.Document.GetElementsByTagName("input");
            foreach (HtmlElement c in buttonElem2)
            {
                if (c.GetAttribute("value").Equals("Display Available Information"))
                {
                    c.InvokeMember("click");
                }
            }

            ////Step 5  New Pop Up
            popup.Navigate(p); // using string builder to open another webbrowser
        }

        
       public  static void DownLoadFile(IE browser)
        {
            browser.Link(Find.ByText("File Download")).ClickNoWait();

            Thread.Sleep(1000);
            AutomationElementCollection dialogElements = AutomationElement.FromHandle(FindWindow(null, "Internet Explorer")).FindAll(TreeScope.Children, Condition.TrueCondition);
            foreach (AutomationElement element in dialogElements)
            {
                if (element.Current.Name.Equals("Save"))
                {
                    var invokePattern = element.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                    invokePattern.Invoke();

                }
            }
        }

        private async void popup_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            if (this.popup.ReadyState != WebBrowserReadyState.Complete)
                return;
            else
            {
                var buttonElem5 = popup.Document.GetElementsByTagName("input");
                foreach (HtmlElement d in buttonElem5)
                {
                    if (d.GetAttribute("value").Equals("Export All to Excel"))
                    {

                        d.InvokeMember("click"); // THIS OPENS A FILE DOWNLOAD POPUP WINDOW WHICH HAS THE "OPEN" "SAVE" "CANCEL"
                                                 //  I AM STUCK HERE !!

                        //hitKey();
                        //SendKeys.SendWait("{LEFT}");

                        //SendKeys.SendWait("{ENTER}");
                        //var ss = "File Download";
                        //DownLoadFile(ss);
                        //WindowsHelper.DownloadIEFile("Save", "D", "kpkp");
                        //SetTimer();
                    }
                }
            }
            
           // button2.PerformClick();
           //WindowsHelper.DownloadIEFile("Save", "D", "kpkp");
            //SetTimer();
        }

       /


       // 
       //   THE METHOD BELOW IS THE FINAL STAGE OF THE PROGRAM, I WANT IT TO AUTOMATICALLY CLICK THE SAVE BUTTON
       //    OF THE FILE DOWNLOAD WINDOW

        public void DownLoadFile(string strWindowTitle)
        {
            

            
            
            
            IntPtr TargetHandle = FindWindowByCaption(IntPtr.Zero , strWindowTitle);
            AutomationElementCollection ParentElements = AutomationElement.FromHandle(TargetHandle).FindAll(TreeScope.Children, Condition.TrueCondition);
            foreach (AutomationElement ParentElement in ParentElements)
            {
                // Identidfy Download Manager Window in Internet Explorer
                if (ParentElement.Current.ClassName == "Frame Notification Bar")
                {
                    AutomationElementCollection ChildElements = ParentElement.FindAll(TreeScope.Children, Condition.TrueCondition);
                    // Idenfify child window with the name Notification Bar or class name as DirectUIHWND 
                    foreach (AutomationElement ChildElement in ChildElements)
                    {
                        if (ChildElement.Current.Name == "Notification bar" || ChildElement.Current.ClassName == "DirectUIHWND")
                        {

                            AutomationElementCollection DownloadCtrls = ChildElement.FindAll(TreeScope.Children, Condition.TrueCondition);
                            foreach (AutomationElement ctrlButton in DownloadCtrls)
                            {
                                //Now invoke the button click whichever you wish
                                if (ctrlButton.Current.Name.ToLower() == "save")
                                {
                                    var invokePattern = ctrlButton.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                                    invokePattern.Invoke();
                                }

                            }
                        }
                    }


                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {   
            //SetTimer();
            WindowsHelper.DownloadIEFile("Save", @"C:\Users\nadjapon\Desktop", "kpkp");
            
        }

    }
}
