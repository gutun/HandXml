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

namespace HandXml2 {
    public partial class Form1 : Form {
        public Form1() {
            InitializeComponent();
            string sheet = ConfigurationManager.AppSettings["sheet"];
            if (!string.IsNullOrEmpty(sheet)) {
                string[] sheets = sheet.Split(new string[] { ",", "，" }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim()).ToArray();
                this.checkedListBox1.Items.Clear();
                foreach (var item in sheets) {
                    this.checkedListBox1.Items.Add(item,true);
                }
            }
        }

        private void btnHand_Click_1(object sender, EventArgs e) {
            bool ignoreflag = this.checkBox1.Checked;
            RichTextBox txtbox = this.richTextBox1;
            string floder = AppDomain.CurrentDomain.BaseDirectory + "Logs//";
            if (Directory.Exists(floder)) {
                Directory.Delete(floder, true);
            }

            if (!string.IsNullOrEmpty(this.textBox1.Text) || !string.IsNullOrEmpty(this.textBox2.Text) || !string.IsNullOrEmpty(this.textBox3.Text) || !string.IsNullOrEmpty(this.textBox4.Text) || !string.IsNullOrEmpty(this.textBox5.Text) || !string.IsNullOrEmpty(this.textBox6.Text)) {
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
                if (fileDialog.ShowDialog().Equals(DialogResult.OK) && fileDialog.FileNames.Length == 2) {
                    string txtFile = "";
                    string excelFile = "";
                    if (fileDialog.FileNames[0].Trim().EndsWith(".txt", StringComparison.OrdinalIgnoreCase)) {
                        txtFile = fileDialog.FileNames[0];
                        excelFile = fileDialog.FileNames[1];
                    } else {
                        txtFile = fileDialog.FileNames[1];
                        excelFile = fileDialog.FileNames[0];
                    }
                    //显示选择文件
                    HandECDS.HandFile(txtbox, excelFile, txtFile, dic, ignoreflag);
                } else {
                    MessageBox.Show("请选择文件!");
                }
            } else {
                MessageBox.Show("请填入变量值!");
            }
        }

        private void button1_Click(object sender, EventArgs e) {
            bool ignoreflag = this.checkBox2.Checked;
            RichTextBox txtbox = this.richTextBox2;
            string floder = AppDomain.CurrentDomain.BaseDirectory + "Logs//";
            if (Directory.Exists(floder)) {
                Directory.Delete(floder, true);
            }
            OpenFileDialog fileDialog = new OpenFileDialog();
            //筛选
            fileDialog.Multiselect = true;
            if (fileDialog.ShowDialog().Equals(DialogResult.OK) && fileDialog.FileNames.Length == 2) {
                string txtFile = "";
                string excelFile = "";
                if (fileDialog.FileNames[0].Trim().EndsWith(".txt", StringComparison.OrdinalIgnoreCase)) {
                    txtFile = fileDialog.FileNames[0];
                    excelFile = fileDialog.FileNames[1];
                } else {
                    txtFile = fileDialog.FileNames[1];
                    excelFile = fileDialog.FileNames[0];
                }
                //显示选择文件
                HandIBPS.HandFile(txtbox, excelFile, txtFile, ignoreflag);

            } else {
                MessageBox.Show("请选择文件!");
            }
        }

        private void button2_Click(object sender, EventArgs e) {
            bool ignoreflag = this.checkBox3.Checked;
            RichTextBox txtbox = this.richTextBox3;
            string floder = AppDomain.CurrentDomain.BaseDirectory + "Logs//";
            if (Directory.Exists(floder)) {
                Directory.Delete(floder, true);
            }
            OpenFileDialog fileDialog = new OpenFileDialog();
            //筛选
            fileDialog.Multiselect = true;
            if (fileDialog.ShowDialog().Equals(DialogResult.OK) && fileDialog.FileNames.Length == 2) {
                List<string> sheets = new List<string>();
                foreach (var item in this.checkedListBox1.CheckedItems) {
                    sheets.Add(item.ToString());
                }
                string txtFile = "";
                string excelFile = "";
                if (fileDialog.FileNames[0].Trim().EndsWith(".txt", StringComparison.OrdinalIgnoreCase)) {
                    txtFile = fileDialog.FileNames[0];
                    excelFile = fileDialog.FileNames[1];
                } else {
                    txtFile = fileDialog.FileNames[1];
                    excelFile = fileDialog.FileNames[0];
                }
                //显示选择文件
                HandCFXPS.HandFile(txtbox, excelFile, txtFile,sheets,ignoreflag);

            } else {
                MessageBox.Show("请选择文件!");
            }
        }
    }
}
