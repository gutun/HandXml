using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace HandXml2
{
    public partial class Form1 : Form
    {
        public BackgroundWorker bw = new BackgroundWorker();

        public Form1()
        {
            InitializeComponent();
            bw.WorkerReportsProgress = true;
            bw.WorkerSupportsCancellation = true;
            bw.DoWork += new DoWorkEventHandler(bw_DoWork);

            //ecds
            string ecdssheet = ConfigurationManager.AppSettings["ECDS_sheet"];
            if (!string.IsNullOrEmpty(ecdssheet))
            {
                string[] sheets = ecdssheet.Split(new string[] { ",", "，" }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim()).ToArray();
                this.cklb_ECDS.Items.Clear();
                foreach (var item in sheets)
                {
                    this.cklb_ECDS.Items.Add(item, true);
                }
            }
            //ibps
            string ibpssheet = ConfigurationManager.AppSettings["IBPS_sheet"];
            if (!string.IsNullOrEmpty(ibpssheet))
            {
                string[] sheets = ibpssheet.Split(new string[] { ",", "，" }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim()).ToArray();
                this.cklb_IBPS.Items.Clear();
                foreach (var item in sheets)
                {
                    this.cklb_IBPS.Items.Add(item, true);
                }
            }
            //cfxps
            string cfxpssheet = ConfigurationManager.AppSettings["CFXPS_sheet"];
            if (!string.IsNullOrEmpty(cfxpssheet))
            {
                string[] sheets = cfxpssheet.Split(new string[] { ",", "，" }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim()).ToArray();
                this.cklb_CFXPS.Items.Clear();
                foreach (var item in sheets)
                {
                    this.cklb_CFXPS.Items.Add(item, true);
                }
            }

            //beps
            string bepsssheet = ConfigurationManager.AppSettings["BEPS_sheet"];
            if (!string.IsNullOrEmpty(bepsssheet))
            {
                string[] sheets = bepsssheet.Split(new string[] { ",", "，" }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim()).ToArray();
                this.cklb_BEPS.Items.Clear();
                foreach (var item in sheets)
                {
                    this.cklb_BEPS.Items.Add(item, true);
                }
            }

            //hvps
            string hvpsssheet = ConfigurationManager.AppSettings["HVPS_sheet"];
            if (!string.IsNullOrEmpty(hvpsssheet))
            {
                string[] sheets = hvpsssheet.Split(new string[] { ",", "，" }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim()).ToArray();
                this.cklb_HVPS.Items.Clear();
                foreach (var item in sheets)
                {
                    this.cklb_HVPS.Items.Add(item, true);
                }
            }

            //cis
            string cisssheet = ConfigurationManager.AppSettings["CIS_sheet"];
            if (!string.IsNullOrEmpty(cisssheet))
            {
                string[] sheets = cisssheet.Split(new string[] { ",", "，" }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim()).ToArray();
                this.cklb_CIS.Items.Clear();
                foreach (var item in sheets)
                {
                    this.cklb_CIS.Items.Add(item, true);
                }
            }

            //cips
            string cipsssheet = ConfigurationManager.AppSettings["CIPS_sheet"];
            if (!string.IsNullOrEmpty(cipsssheet))
            {
                string[] sheets = cipsssheet.Split(new string[] { ",", "，" }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim()).ToArray();
                this.cklb_CIPS.Items.Clear();
                foreach (var item in sheets)
                {
                    this.cklb_CIPS.Items.Add(item, true);
                }
            }
        }

        private void btn_ECDS_Click(object sender, EventArgs e)
        {
            bool ignoreflag = this.cb_ECDS.Checked;
            RichTextBox txtbox = this.rtb_ECDS;
            string floder = AppDomain.CurrentDomain.BaseDirectory + "Logs//";
            if (Directory.Exists(floder))
            {
                Directory.Delete(floder, true);
            }

            if (!string.IsNullOrEmpty(this.textBox1.Text) || !string.IsNullOrEmpty(this.textBox2.Text) || !string.IsNullOrEmpty(this.textBox3.Text) || !string.IsNullOrEmpty(this.textBox4.Text) || !string.IsNullOrEmpty(this.textBox5.Text) || !string.IsNullOrEmpty(this.textBox6.Text))
            {
                Dictionary<string, string[]> dic = new Dictionary<string, string[]>();
                if (!string.IsNullOrEmpty(this.textBox1.Text))
                    dic.Add(this.label2.Text.Trim(), this.textBox1.Text.Split(new string[] { ",", "，" }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim()).ToArray());
                if (!string.IsNullOrEmpty(this.textBox2.Text))
                    dic.Add(this.label3.Text.Trim(), this.textBox2.Text.Split(new string[] { ",", "，" }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim()).ToArray());
                if (!string.IsNullOrEmpty(this.textBox3.Text))
                    dic.Add(this.label4.Text.Trim(), this.textBox3.Text.Split(new string[] { ",", "，" }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim()).ToArray());
                if (!string.IsNullOrEmpty(this.textBox4.Text))
                    dic.Add(this.label5.Text.Trim(), this.textBox4.Text.Split(new string[] { ",", "，" }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim()).ToArray());
                if (!string.IsNullOrEmpty(this.textBox5.Text))
                    dic.Add(this.label6.Text.Trim(), this.textBox5.Text.Split(new string[] { ",", "，" }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim()).ToArray());
                if (!string.IsNullOrEmpty(this.textBox6.Text))
                    dic.Add(this.label7.Text.Trim(), this.textBox6.Text.Split(new string[] { ",", "，" }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim()).ToArray());
                OpenFileDialog fileDialog = new OpenFileDialog();
                //筛选
                fileDialog.Multiselect = true;
                if (fileDialog.ShowDialog().Equals(DialogResult.OK) && fileDialog.FileNames.Length == 2)
                {
                    string txtFile = "";
                    string excelFile = "";
                    if (fileDialog.FileNames[0].Trim().EndsWith(".txt", StringComparison.OrdinalIgnoreCase))
                    {
                        txtFile = fileDialog.FileNames[0];
                        excelFile = fileDialog.FileNames[1];
                    }
                    else
                    {
                        txtFile = fileDialog.FileNames[1];
                        excelFile = fileDialog.FileNames[0];
                    }
                    //显示选择文件
                    HandECDS.HandFile(txtbox, excelFile, txtFile, dic, ignoreflag);
                }
                else
                {
                    MessageBox.Show("请选择文件!");
                }
            }
            else
            {
                MessageBox.Show("请填入变量值!");
            }
        }

        private void btn_IBPS_Click(object sender, EventArgs e)
        {
            //bool ignoreflag = this.cb_IBPS.Checked;
            //RichTextBox txtbox = this.rtb_IBPS;
            //string floder = AppDomain.CurrentDomain.BaseDirectory + "Logs//";
            //if (Directory.Exists(floder))
            //{
            //    Directory.Delete(floder, true);
            //}
            //OpenFileDialog fileDialog = new OpenFileDialog();
            ////筛选
            //fileDialog.Multiselect = true;
            //if (fileDialog.ShowDialog().Equals(DialogResult.OK) && fileDialog.FileNames.Length == 2)
            //{
            //    string txtFile = "";
            //    string excelFile = "";
            //    if (fileDialog.FileNames[0].Trim().EndsWith(".txt", StringComparison.OrdinalIgnoreCase))
            //    {
            //        txtFile = fileDialog.FileNames[0];
            //        excelFile = fileDialog.FileNames[1];
            //    }
            //    else
            //    {
            //        txtFile = fileDialog.FileNames[1];
            //        excelFile = fileDialog.FileNames[0];
            //    }
            //    //显示选择文件
            //    HandIBPS.HandFile(txtbox, excelFile, txtFile, ignoreflag);

            //}
            //else
            //{
            //    MessageBox.Show("请选择文件!");
            //}

            if (bw.IsBusy != true)
            {
                bool ignoreflag = this.cb_IBPS.Checked;
                RichTextBox txtbox = this.rtb_IBPS;
                string floder = AppDomain.CurrentDomain.BaseDirectory + "Logs//";
                if (Directory.Exists(floder))
                {
                    Directory.Delete(floder, true);
                }
                OpenFileDialog fileDialog = new OpenFileDialog();
                //筛选
                fileDialog.Multiselect = true;
                //文件大于等于1
                if (fileDialog.ShowDialog().Equals(DialogResult.OK) && fileDialog.FileNames.Length >= 1)
                {
                    string[] files = fileDialog.FileNames;
                    string excelfile = files.FirstOrDefault(x => x.Trim().EndsWith(".xls", StringComparison.OrdinalIgnoreCase) || x.Trim().EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase));
                    List<string> txtfiles = files.Where(x => x.Trim().EndsWith(".txt", StringComparison.OrdinalIgnoreCase)).ToList();
                    List<string> sheets = new List<string>();
                    foreach (var item in this.cklb_IBPS.CheckedItems)
                    {
                        sheets.Add(item.ToString());
                    }
                    this.rtb_IBPS.Clear();
                    bw.RunWorkerAsync(new FileData() { excelfile = excelfile, txtfiles = txtfiles, sheets = sheets, ignoreflag = ignoreflag });

                }
                else
                {
                    MessageBox.Show("请选择文件!");
                }

            }

        }

        private void btn_CFXPS_Click(object sender, EventArgs e)
        {
            bool ignoreflag = this.cb_CFXPS.Checked;
            RichTextBox txtbox = this.rtb_CFXPS;
            string floder = AppDomain.CurrentDomain.BaseDirectory + "Logs//";
            if (Directory.Exists(floder))
            {
                Directory.Delete(floder, true);
            }
            OpenFileDialog fileDialog = new OpenFileDialog();
            //筛选
            fileDialog.Multiselect = true;
            if (fileDialog.ShowDialog().Equals(DialogResult.OK) && fileDialog.FileNames.Length == 2)
            {
                List<string> sheets = new List<string>();
                foreach (var item in this.cklb_CFXPS.CheckedItems)
                {
                    sheets.Add(item.ToString());
                }
                string txtFile = "";
                string excelFile = "";
                if (fileDialog.FileNames[0].Trim().EndsWith(".txt", StringComparison.OrdinalIgnoreCase))
                {
                    txtFile = fileDialog.FileNames[0];
                    excelFile = fileDialog.FileNames[1];
                }
                else
                {
                    txtFile = fileDialog.FileNames[1];
                    excelFile = fileDialog.FileNames[0];
                }
                //显示选择文件
                HandCFXPS.HandFile(txtbox, excelFile, txtFile, sheets, ignoreflag);

            }
            else
            {
                MessageBox.Show("请选择文件!");
            }
        }

        private void btn_BEPS_Click(object sender, EventArgs e)
        {
            bool ignoreflag = this.cb_BEPS.Checked;
            RichTextBox txtbox = this.rtb_BEPS;
            string floder = AppDomain.CurrentDomain.BaseDirectory + "Logs//";
            if (Directory.Exists(floder))
            {
                Directory.Delete(floder, true);
            }
            OpenFileDialog fileDialog = new OpenFileDialog();
            //筛选
            fileDialog.Multiselect = true;
            //文件大于等于1
            if (fileDialog.ShowDialog().Equals(DialogResult.OK) && fileDialog.FileNames.Length >= 1)
            {
                string[] files = fileDialog.FileNames;
                string excelfile = files.FirstOrDefault(x => x.Trim().EndsWith(".xls", StringComparison.OrdinalIgnoreCase) || x.Trim().EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase));
                List<string> txtfiles = files.Where(x => x.Trim().EndsWith(".txt", StringComparison.OrdinalIgnoreCase)).ToList();
                List<string> sheets = new List<string>();
                foreach (var item in this.cklb_BEPS.CheckedItems)
                {
                    sheets.Add(item.ToString());
                }
                //显示选择文件
                HandBEPS.HandFile(txtbox, excelfile, txtfiles, sheets, ignoreflag);

            }
            else
            {
                MessageBox.Show("请选择文件!");
            }
        }

        private void btn_HVPS_Click(object sender, EventArgs e)
        {
            bool ignoreflag = this.cb_HVPS.Checked;
            RichTextBox txtbox = this.rtb_HVPS;
            string floder = AppDomain.CurrentDomain.BaseDirectory + "Logs//";
            if (Directory.Exists(floder))
            {
                Directory.Delete(floder, true);
            }
            OpenFileDialog fileDialog = new OpenFileDialog();
            //筛选
            fileDialog.Multiselect = true;
            //文件大于等于1
            if (fileDialog.ShowDialog().Equals(DialogResult.OK) && fileDialog.FileNames.Length >= 1)
            {
                string[] files = fileDialog.FileNames;
                string excelfile = files.FirstOrDefault(x => x.Trim().EndsWith(".xls", StringComparison.OrdinalIgnoreCase) || x.Trim().EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase));
                List<string> txtfiles = files.Where(x => x.Trim().EndsWith(".txt", StringComparison.OrdinalIgnoreCase)).ToList();
                List<string> sheets = new List<string>();
                foreach (var item in this.cklb_HVPS.CheckedItems)
                {
                    sheets.Add(item.ToString());
                }
                //显示选择文件
                HandHVPS.HandFile(txtbox, excelfile, txtfiles, sheets, ignoreflag);

            }
            else
            {
                MessageBox.Show("请选择文件!");
            }
        }

        private void btn_CIS_Click(object sender, EventArgs e)
        {
            bool ignoreflag = this.cb_CIS.Checked;
            RichTextBox txtbox = this.rtb_CIS;
            string floder = AppDomain.CurrentDomain.BaseDirectory + "Logs//";
            if (Directory.Exists(floder))
            {
                Directory.Delete(floder, true);
            }
            OpenFileDialog fileDialog = new OpenFileDialog();
            //筛选
            fileDialog.Multiselect = true;
            //文件大于等于1
            if (fileDialog.ShowDialog().Equals(DialogResult.OK) && fileDialog.FileNames.Length >= 1)
            {
                string[] files = fileDialog.FileNames;
                string excelfile = files.FirstOrDefault(x => x.Trim().EndsWith(".xls", StringComparison.OrdinalIgnoreCase) || x.Trim().EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase));
                List<string> txtfiles = files.Where(x => x.Trim().EndsWith(".txt", StringComparison.OrdinalIgnoreCase)).ToList();
                List<string> sheets = new List<string>();
                foreach (var item in this.cklb_CIS.CheckedItems)
                {
                    sheets.Add(item.ToString());
                }
                //显示选择文件
                HandCIS.HandFile(txtbox, excelfile, txtfiles, sheets, ignoreflag);

            }
            else
            {
                MessageBox.Show("请选择文件!");
            }
        }

        private void btn_CIPS_Click(object sender, EventArgs e)
        {

        }

        public delegate void RichTextBoxUpdateEventHandler(string message, Color? color = null);

        private void UpdateIBPS(string message, Color? color = null)
        {
            if (this.rtb_IBPS.InvokeRequired)
            {
                // this means we’re on the wrong thread!  
                // use BeginInvoke or Invoke to call back on the 
                // correct thread.
                this.rtb_IBPS.Invoke(
                    new RichTextBoxUpdateEventHandler(UpdateIBPS), // the method to call back on
                    new object[] { message, color });                              // the list of arguments to pass
            }
            else
            {
                if (null != color)
                {
                    var defcolor = this.rtb_IBPS.ForeColor;
                    this.rtb_IBPS.ForeColor = (Color)color;
                    this.rtb_IBPS.WriteLine(message);
                    this.rtb_IBPS.ForeColor = defcolor;
                }
                else
                {
                    this.rtb_IBPS.WriteLine(message);
                }

            }
        }

        private void bw_DoWork(object sender, DoWorkEventArgs e)
        {
            FileData data = e.Argument as FileData;
            //显示选择文件
            HandIBPSNEW.HandFile(UpdateIBPS, data.excelfile, data.txtfiles, data.sheets, data.ignoreflag);
        }
    }

    public class FileData
    {
        public string excelfile { get; set; }
        public List<string> txtfiles { get; set; }
        public List<string> sheets { get; set; }
        public bool ignoreflag { get; set; }
    }
}
