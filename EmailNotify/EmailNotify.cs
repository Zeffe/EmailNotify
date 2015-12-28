using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Management;
using Microsoft.Win32;

namespace EmailNotify
{
    public partial class EmailNotify : ServiceBase
    {
        public EmailNotify()
        {
            InitializeComponent();
        }

        public enum ServiceState
        {
            SERVICE_STOPPED = 0x00000001,
            SERVICE_START_PENDING = 0x00000002,
            SERVICE_STOP_PENDING = 0x00000003,
            SERVICE_RUNNING = 0x00000004,
            SERVICE_CONTINUE_PENDING = 0x00000005,
            SERVICE_PAUSE_PENDING = 0x00000006,
            SERVICE_PAUSED = 0x00000007,
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct ServiceStatus
        {
            public long dwServiceType;
            public ServiceState dwCurrentState;
            public long dwControlsAccepted;
            public long dwWin32ExitCode;
            public long dwServiceSpecificExitCode;
            public long dwCheckPoint;
            public long dwWaitHint;
        };

        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern bool SetServiceStatus(IntPtr handle, ref ServiceStatus serviceStatus);

        bool GO = true;
        string nuUser;
        DateTime now = DateTime.Now;

        void sendMail()
        {
            var client = new SmtpClient("smtp.gmail.com", 587)
            {
                Credentials = new System.Net.NetworkCredential("SMTP5241@gmail.com", "livnutgd"),
                EnableSsl = true
            };
            string subject = "Logged into Seth on: " + now.ToShortDateString() + " - " + now.ToShortTimeString();
            client.Send("SMTP5241@gmail.com", "zeffer5789@gmail.com", subject, "");
        }


        protected override void OnStart(string[] args)
        {
            ServiceStatus serviceStatus = new ServiceStatus();
            serviceStatus.dwCurrentState = ServiceState.SERVICE_START_PENDING;
            serviceStatus.dwWaitHint = 100000;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);

            System.Timers.Timer timer = new System.Timers.Timer();
            timer.Interval = 60000; // 60 seconds
            timer.Elapsed += new System.Timers.ElapsedEventHandler(this.OnTimer);
            timer.Start();

            serviceStatus.dwCurrentState = ServiceState.SERVICE_RUNNING;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);
        }

        protected override bool OnPowerEvent(PowerBroadcastStatus powerStatus)
        {
            GO = true;
            return true;
        }

        public void OnTimer(object sender, System.Timers.ElapsedEventArgs args)
        {
            now = DateTime.Now;
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT UserName FROM Win32_ComputerSystem");
            ManagementObjectCollection collection = searcher.Get();
            string userName = (string)collection.Cast<ManagementBaseObject>().First()["UserName"];
            nuUser = userName.Split('\\')[1];
            if (nuUser == "Seth" && GO)
            {
                sendMail();
                GO = false;
            } else if (nuUser != "Seth")
            {
                GO = true;
            }
        }

        protected override void OnStop()
        {
        }
    }
}
