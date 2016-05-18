using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace settingService
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void butSetService_Click(object sender, EventArgs e)
        {
            commonExecuteCommand("setupCooperService.bat"); 
        }

        private void button2_Click(object sender, EventArgs e)
        {
            commonExecuteCommand("unsetupCooperService.bat");   
        }

        private void button3_Click(object sender, EventArgs e)
        {
            commonExecuteCommand("stop.bat");            
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                commonExecuteCommand("start.bat");
               
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private string checkServiceStatus(string servicename)
        {
            try
            {
                ServiceController sc = new ServiceController(servicename);
                
                switch (sc.Status)
                {
                    case ServiceControllerStatus.Running:
                        return "Running";
                    case ServiceControllerStatus.Stopped:
                        return "Stopped";
                    case ServiceControllerStatus.Paused:
                        return "Paused";
                    case ServiceControllerStatus.StopPending:
                        return "Stopping";
                    case ServiceControllerStatus.StartPending:
                        return "Starting";
                    default:
                        return "Status Changing";
                }
            }
            catch (Exception ex)
            {
                return "服務不存在";
            }
        }

        private void commonExecuteCommand(string batFileName)
        {
            try
            {
                string currentPath = Directory.GetCurrentDirectory();
                string targetDir = string.Format(currentPath);//this is where mybatch.bat lies
                Process proc = new Process();
                proc.StartInfo.WorkingDirectory = targetDir;
                proc.StartInfo.FileName = batFileName;
                proc.StartInfo.CreateNoWindow = false;
                proc.Start();
                proc.WaitForExit();
                lbServiceStatus.Text = checkServiceStatus("importXMLToCooper");
            }
            catch (Exception ex)
            {   
                throw new Exception(ex.Message);
            }
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            try
            {
                lbServiceStatus.Text = checkServiceStatus("importXMLToCooper");
            }
            catch (Exception ex)
            {
                lbServiceStatus.Text = "未安裝";
                //throw new Exception(ex.Message);
            }
             
        }
    }
}
