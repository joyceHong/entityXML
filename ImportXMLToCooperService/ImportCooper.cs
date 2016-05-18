using entityXML;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;

namespace ImportCooper
{
   
    public partial class ImportCooper : ServiceBase
    {
        private Timer MyTimer;

        public ImportCooper()
        {
            InitializeComponent();
        }

        Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
        WriteEvent.WrittingEventLog writeObj = new WriteEvent.WrittingEventLog();
        string cooperPath, xmlFromPath, xmlToPath, isWriteLog, timeStart, timeEnd = "";
        protected override void OnStart(string[] args)
        {
            AppSettingsSection appSettings = (AppSettingsSection)config.GetSection("appSettings");

            cooperPath = (appSettings.Settings["COOPER"].Value == null) ? "" : appSettings.Settings["COOPER"].Value;
            xmlFromPath = (appSettings.Settings["XMLFromPath"].Value == null) ? "" : appSettings.Settings["XMLFromPath"].Value;
            xmlToPath = (appSettings.Settings["XMLToPath"].Value == null) ? "" : appSettings.Settings["XMLToPath"].Value;
            isWriteLog = (appSettings.Settings["WriteLog"].Value == null) ? "" : appSettings.Settings["WriteLog"].Value;

            timeStart = (appSettings.Settings["StartTime"].Value == null) ? "" : appSettings.Settings["StartTime"].Value;
            timeEnd = (appSettings.Settings["EndTime"].Value == null) ? "" : appSettings.Settings["EndTime"].Value;
            string strInterVal = appSettings.Settings["Interval"].Value;
            int intInterval = 0;
            int.TryParse(strInterVal, out intInterval);

            if (isWriteLog.ToUpper() == "Y")
                writeObj.writeToFile("服務啟動中--------");

            /**
             * 1 每間隔多久才執行一次             
             * 2 檢查時間是否落在區間才執行程式
             */
            MyTimer = new Timer();
            MyTimer.Elapsed += new ElapsedEventHandler(MyTimer_Elapsed);
            MyTimer.Interval = intInterval * (1000 * 60);
            MyTimer.Start();
        }

        private void MyTimer_Elapsed(object sender, ElapsedEventArgs e)
        {
            try
            {
                if (isWriteLog.ToUpper() == "Y")
                {
                    writeObj.writeToFile(DateTime.Now.ToString("yyyyMMdd HH:mm:ss"));
                }

                string strStartTime = DateTime.Today.ToString("yyyyMMdd") + timeStart;
                string strEndTime = DateTime.Today.ToString("yyyyMMdd") + timeEnd;

               
                DateTime startTime = DateTime.ParseExact(strStartTime, "yyyyMMddHHmm",null);

                DateTime endTime = DateTime.ParseExact(strEndTime, "yyyyMMddHHmm", null);

                if (endTime < startTime)
                {
                    endTime.AddDays(1);
                }

                controllerXMLtoFoxpro controlObj = new controllerXMLtoFoxpro();
                controlObj._cooperPath = cooperPath;
                controlObj._xmlFromPath = xmlFromPath;
                controlObj._xmlToPath = xmlToPath;
                controlObj._isWriteLog = isWriteLog;

                if (DateTime.Now >= startTime && DateTime.Now <= endTime)
                {
                    if (isWriteLog.ToUpper() == "Y")
                    {
                        writeObj.writeToFile("起迄時間:" + startTime + "~ " + endTime + "路徑:" + xmlFromPath);
                    }
                    controlObj.start();
                    if (isWriteLog.ToUpper() == "Y")
                    {
                        writeObj.writeToFile("已執行 start:" + startTime + "~ " + endTime);
                    }
                }
            }
            catch (Exception ex)
            {
                WriteEvent.WrittingEventLog writeObj = new WriteEvent.WrittingEventLog();
                writeObj.writeToFile(DateTime.Now.ToString("yyyyMMdd") + "_ImoprtCooperServiceErrorLog", Directory.GetCurrentDirectory(), ex.Message.ToString());
                //throw new Exception(ex.Message);
            }
        }

        protected override void OnStop()
        {
            MyTimer.Stop();
            MyTimer = null;
            if (isWriteLog.ToUpper() == "Y")
                writeObj.writeToFile("服務結束中--------");
        }
    }
}
