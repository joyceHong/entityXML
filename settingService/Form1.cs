using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace settingService
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }



        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {

                string applicationName =
                Environment.GetCommandLineArgs()[0];
                string exePath = Path.Combine(Environment.CurrentDirectory, "WindowsService1.exe");

                Configuration config =
       ConfigurationManager.OpenExeConfiguration(exePath);

                txtCooperPath.Text = config.AppSettings.Settings["COOPER"].Value;
                txtXmlFromPath.Text = config.AppSettings.Settings["XMLFromPath"].Value;
                txtXmlToPath.Text = config.AppSettings.Settings["XMLToPath"].Value;
                txtStartTime.Text = config.AppSettings.Settings["StartTime"].Value;
                txtEndTime.Text = config.AppSettings.Settings["EndTime"].Value;
                txtInterVal.Text = config.AppSettings.Settings["InterVal"].Value;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }

        }

        private void butSave_Click(object sender, EventArgs e)
        {
            try
            {
                Configuration config = ConfigurationManager.OpenExeConfiguration("WindowsService1.exe");
                config.AppSettings.Settings["COOPER"].Value = txtCooperPath.Text.Trim();
                config.AppSettings.Settings["XMLFromPath"].Value = txtXmlFromPath.Text.Trim();
                config.AppSettings.Settings["XMLToPath"].Value = txtXmlToPath.Text.Trim();
                config.AppSettings.Settings["StartTime"].Value = txtStartTime.Text.Trim();
                config.AppSettings.Settings["EndTime"].Value = txtEndTime.Text.Trim();
                config.AppSettings.Settings["InterVal"].Value = txtInterVal.Text.Trim();
                config.Save(ConfigurationSaveMode.Modified);
                MessageBox.Show("儲存完畢");
            }
            catch (Exception ex)
            {   
                throw new Exception(ex.Message);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog selectedFilePath = new FolderBrowserDialog();
            
            if (selectedFilePath.ShowDialog() == DialogResult.OK)
            {
                txtCooperPath.Text = selectedFilePath.SelectedPath;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog selectedFilePath = new FolderBrowserDialog();

            if (selectedFilePath.ShowDialog() == DialogResult.OK)
            {
                txtXmlFromPath.Text = selectedFilePath.SelectedPath;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog selectedFilePath = new FolderBrowserDialog();

            if (selectedFilePath.ShowDialog() == DialogResult.OK)
            {
                txtXmlToPath.Text = selectedFilePath.SelectedPath;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Form2 settingService = new Form2();
            settingService.ShowDialog();
        }
    }
        
}
