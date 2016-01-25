using Aspose.Cells;
using FluentData;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace HandXml2
{
    /// <summary>
    /// 公共的帮助类
    /// </summary>
    public static class CommonHelper
    {
        /// <summary>
        /// 日志文件名称
        /// </summary>
        private static string fileName;

        /// <summary>
        /// 设置或者读取日志文件名称
        /// </summary>
        public static string FileName
        {
            get
            {
                if (string.IsNullOrEmpty(fileName))
                {
                    return DateTime.Now.ToString("yyyy-MM-dd");
                }
                return fileName;
            }
            set { fileName = value; }
        }

        #region 读取Excel的Sheet到虚拟表 Datatable
        public static DataTable ReadExcel(string file, string sheetName, ref int titleRowIndex)
        {
            Workbook wkBook = new Workbook();
            wkBook.Open(file);

            Worksheet wkSheet = wkBook.Worksheets[sheetName];
            if (null != wkSheet)
            {
                //声明DataTable存放sheet
                DataTable dtTemp = new DataTable();
                //设置Table名为sheet的名称
                dtTemp.TableName = wkSheet.Name;

                //遍历行
                for (int x = 0; x < wkSheet.Cells.MaxDataRow + 1; x++)
                {
                    bool firstRow = true;
                    for (int y = 0; y < wkSheet.Cells.MaxDataColumn + 1; y++)
                    {
                        string value = wkSheet.Cells[x, y].StringValue.Trim();
                        firstRow = firstRow && !string.IsNullOrEmpty(value);
                    }
                    if (!firstRow)
                    {
                        continue;
                    }
                    else
                    {
                        titleRowIndex = x;
                        break;
                    }
                }
                //遍历行
                for (int x = titleRowIndex; x < wkSheet.Cells.MaxDataRow + 1; x++)
                {
                    //声明DataRow存放sheet的数据行
                    DataRow dRow = null;

                    //遍历列
                    for (int y = 0; y < wkSheet.Cells.MaxDataColumn + 1; y++)
                    {
                        //获取单元格的值
                        string value = wkSheet.Cells[x, y].StringValue.Trim();

                        //如果是第一行，则当作表头
                        if (x == titleRowIndex)
                        {
                            //设置表头
                            DataColumn dCol = new DataColumn(value);
                            dtTemp.Columns.Add(dCol);
                        }

                        //非第一行，则为数据行
                        else
                        {
                            //每次循环到第一列时，实例DataRow
                            if (y == 0)
                            {
                                dRow = dtTemp.NewRow();
                            }
                            //给第Y列赋值
                            dRow[y] = value;
                        }
                    }

                    if (dRow != null)
                    {
                        dtTemp.Rows.Add(dRow);
                    }
                }
                //释放对象
                wkSheet = null;
                wkBook = null;
                return dtTemp;
            }
            return null;
        }
        #endregion

        #region 数据库连接对象
        /// <summary>
        /// 数据库连接对象
        /// </summary>
        /// <returns></returns>
        public static IDbContext Context(string name)
        {
            string connStr = System.Configuration.ConfigurationManager.ConnectionStrings[name].ToString();
            var provider = new DB2Provider();
            var dbcontext = new DbContext().ConnectionString(connStr, provider);
            dbcontext.CommandTimeout(90);
            return dbcontext;
        }
        #endregion

        #region 日志相关的方法
        /// <summary>
        /// 输出Log
        /// </summary>
        /// <param name="msg"></param>
        public static void WriteLog(string msg)
        {
            try
            {
                string fullSaveDir = AppDomain.CurrentDomain.BaseDirectory + "Logs//";
                if (!Directory.Exists(fullSaveDir))
                    Directory.CreateDirectory(fullSaveDir);
                string filepath = string.Format("{0}{1}.txt", fullSaveDir, FileName);
                System.IO.File.AppendAllText(filepath, msg);
            }
            catch (Exception)
            {
            }
        }
        /// <summary>
        /// 输出Error
        /// </summary>
        /// <param name="msg"></param>
        public static void WriteError(string msg)
        {
            string fullSaveDir = AppDomain.CurrentDomain.BaseDirectory + "Logs//";
            if (!Directory.Exists(fullSaveDir))
                Directory.CreateDirectory(fullSaveDir);
            string filepath = string.Format("{0}{1}.txt", fullSaveDir, FileName + "ErrorLog.txt");
            System.IO.File.AppendAllText(filepath, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss ") + msg + Environment.NewLine);
        }

        #endregion

        #region 界面文本框扩展方法
        public static void WriteLine(this RichTextBox textbox, string msg)
        {
            textbox.AppendText(msg);
            textbox.AppendText(Environment.NewLine);
            textbox.SelectionStart = textbox.Text.Length;
            textbox.ScrollToCaret();

        }
        public static void WriteLine(this RichTextBox textbox, string msg, bool flag)
        {
            if (!flag)
            {
                var defforColor = textbox.SelectionColor;
                textbox.SelectionColor = Color.Blue;
                textbox.AppendText(msg);
                textbox.AppendText(Environment.NewLine);
                textbox.SelectionStart = textbox.Text.Length;
                textbox.ScrollToCaret();
                textbox.SelectionColor = defforColor;
            }
            else
            {
                textbox.AppendText(msg);
                textbox.AppendText(Environment.NewLine);
                textbox.SelectionStart = textbox.Text.Length;
                textbox.ScrollToCaret();
            }

        }
        #endregion

        public static string HandString(string source, string key, string value)
        {
            do
            {
                int keyInde = source.IndexOf(key);
                if (keyInde > -1)
                {
                    string subSource = source.Substring(keyInde);
                    int endKeyIndex = subSource.IndexOf("'");
                    string dest = subSource.Substring(0, endKeyIndex);
                    string[] keyArray = dest.Split(new string[] { "+", "(", ",", ")" }, StringSplitOptions.RemoveEmptyEntries);
                    if (keyArray.Length == 1)
                    {
                        string keyStr = source.Substring(keyInde, endKeyIndex);
                        source = source.Replace(keyStr, value);
                    }
                    else
                    {
                        var index = 0;
                        do
                        {
                            index++;
                            if (keyArray[index].ToLower().Equals("substr"))
                            {
                                if (int.Parse(keyArray[index + 1]) > 0 && int.Parse(keyArray[index + 2]) > int.Parse(keyArray[index + 1]))
                                {
                                    string subValue = value.Substring(int.Parse(keyArray[index + 1]) - 1, int.Parse(keyArray[index + 2]) - int.Parse(keyArray[index + 1]) + 1);
                                    string keyStr = source.Substring(keyInde, endKeyIndex);
                                    source = source.Replace(keyStr, subValue);
                                    index += 2;
                                }
                            }
                            else if (keyArray[index].ToLower().Equals("math"))
                            {
                                switch (keyArray[index + 1])
                                {
                                    case "+":
                                        long longvalue1 = long.Parse(value) + int.Parse(keyArray[index + 2]);
                                        value = String.Format("{0:D" + value.Length + "}", longvalue1);
                                        index += 2;
                                        break;
                                    case "-":
                                        long longvalue2 = long.Parse(value) - int.Parse(keyArray[index + 2]);
                                        value = String.Format("{0:D" + value.Length + "}", longvalue2);
                                        index += 2;
                                        break;
                                    default: break;
                                }
                            }
                        } while (index < keyArray.Length - 1);
                    }
                }
            } while (source.IndexOf(key) > -1);
            return source;
        }
    }

    public static class StringExtensions
    {
        public static string SafeReplace(this string input, string find, string replace, bool matchWholeWord)
        {
            string textToFind = matchWholeWord ? string.Format(@"\b{0}\b", find) : find;
            return Regex.Replace(input, textToFind, replace);
        }
    }
}
