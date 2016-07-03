using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;

namespace HandXml2 {
    public static class HandIBPS {
        public static void HandFile(System.Windows.Forms.RichTextBox textbox, string excelfile, string txtfile, bool ignoreflag) {
            textbox.Clear();
            try {
                //读取Excel到内存中
                int titleRowIndex = 0;
                DataTable table = CommonHelper.ReadExcel(excelfile, "IBPS验收案例检测标准", ref titleRowIndex);
                Dictionary<string, string[]> dbresult = ReadResult(txtfile);
                DataTable resultTable = table.Copy();
                resultTable.Columns.Add("index", typeof(int));
                resultTable.Columns.Add("sql", typeof(string));
                resultTable.Columns.Add("result", typeof(string));
                resultTable.Columns.Add("filetype", typeof(string));
                resultTable.Columns.Add("baowenbiaoshihao", typeof(string));
                resultTable.Columns.Add("log", typeof(string));
                //匹配Excel每一行
                for (int i = 0; i < table.Rows.Count; i++) {
                    DataRow row = table.Rows[i];
                    if (ignoreflag && ((table.Columns.Contains("是否通过") && row["是否通过"].ToString().Equals("通过")) || (table.Columns.Contains("是否" + Environment.NewLine + "通过") && row["是否" + Environment.NewLine + "通过"].ToString().Equals("通过")))) {
                        continue;
                    }
                    //验收项目编号
                    string ysalbh = row["验收案例编号"].ToString();
                    //提交内容项
                    string tjnrx = row["提交内容项"].ToString();
                    //提交内容	
                    string tjnr = row["提交内容"].ToString();

                    if (!string.IsNullOrEmpty(ysalbh) && !ysalbh.Trim().Equals("验收案例编号") && !string.IsNullOrEmpty(tjnrx) && !tjnrx.Trim().Equals("提交内容项") && !string.IsNullOrEmpty(tjnr) && !tjnr.Trim().Equals("提交内容")) {
                        try {
                            textbox.WriteLine("验收案例编号:" + ysalbh);

                            if (!dbresult.ContainsKey(ysalbh)) {
                                DataRow resultTableRow = resultTable.Rows[i];
                                resultTableRow["index"] = i + titleRowIndex + 1;
                                resultTableRow["result"] = "不通过";
                                //不做处理
                                textbox.WriteLine("找不到相关的sql");
                                resultTableRow["log"] = "找不到相关的sql";
                            } else {
                                DataRow resultTableRow = resultTable.Rows[i];
                                resultTableRow["index"] = i + titleRowIndex + 1;
                                if (dbresult[ysalbh].Length == 4) {
                                    string[] sqlInfo = null;
                                    string sql = string.Empty;
                                    string filetypeName = string.Empty;
                                    bool sqlresult = dbresult.TryGetValue(ysalbh, out sqlInfo);
                                    int itemIndex = 0;
                                    if (sqlresult && null != sqlInfo && sqlInfo.Length > 0) {
                                        sql = sqlInfo[3];
                                        filetypeName = sqlInfo[1];
                                        resultTableRow["filetype"] = filetypeName;
                                        do {
                                            DataRow itemRow = table.Rows[i + itemIndex];
                                            //验收项目编号
                                            string itemysalbh = itemRow["验收案例编号"].ToString();
                                            string[] itemtjnrxlines = itemRow["提交内容项"].ToString().Replace("\r", "").Split('\n').Where(x => !string.IsNullOrEmpty(x.Trim())).ToArray();
                                            string[] itemtjnrlines = itemRow["提交内容"].ToString().Replace("\r", "").Split('\n').Where(x => !string.IsNullOrEmpty(x.Trim())).ToArray();
                                            for (int lineindex = 0; lineindex <= itemtjnrxlines.Length - 1; lineindex++) {
                                                //提交内容项
                                                string itemtjnrx = itemtjnrxlines[lineindex].ToString();
                                                //提交内容	
                                                string itemtjnr = itemtjnrlines[lineindex].ToString();
                                                if (!string.IsNullOrEmpty(itemtjnrx) && !string.IsNullOrEmpty(itemtjnr)) {
                                                    itemtjnrx = itemtjnrx.Substring(itemtjnrx.IndexOf(')') + 1);
                                                    itemtjnrx = itemtjnrx.Replace("：", "").Trim();
                                                    sql = CommonHelper.HandString(sql, itemtjnrx, itemtjnr);
                                                    if (itemIndex == 0) {
                                                        resultTableRow["baowenbiaoshihao"] = itemtjnr;
                                                    }
                                                }
                                            }
                                            itemIndex++;
                                        } while (i + itemIndex < table.Rows.Count && string.IsNullOrEmpty(table.Rows[i + itemIndex]["验收案例编号"].ToString()));
                                    }
                                    textbox.WriteLine("sql语句:" + sql);
                                    resultTableRow["sql"] = sql;
                                    try {
                                        bool sqlResult = GetDbResult(sqlInfo[2], sql);
                                        resultTableRow["result"] = sqlResult ? "通过" : "不通过";
                                        resultTableRow["log"] = "sql执行" + sqlResult;
                                        textbox.WriteLine("sql执行结果:" + sqlResult);
                                    } catch (Exception ex) {
                                        resultTableRow["log"] = ex.Message;
                                        resultTableRow["result"] = "不通过";
                                        textbox.WriteLine("sql执行结果:" + ex.Message);
                                    }
                                } else {
                                    if (dbresult[ysalbh].Length >= 2) {
                                        string[] sqlInfo = null;
                                        string filetypeName = string.Empty;
                                        bool sqlresult = dbresult.TryGetValue(ysalbh, out sqlInfo);
                                        int itemIndex = 0;
                                        if (sqlresult && null != sqlInfo && sqlInfo.Length > 0) {
                                            filetypeName = sqlInfo[1];
                                            resultTableRow["filetype"] = filetypeName;
                                        }
                                        do {
                                            DataRow itemRow = table.Rows[i + itemIndex];
                                            //验收项目编号
                                            string itemysalbh = itemRow["验收案例编号"].ToString();
                                            string[] itemtjnrxlines = itemRow["提交内容项"].ToString().Replace("\r", "").Split('\n').Where(x => !string.IsNullOrEmpty(x.Trim())).ToArray();
                                            string[] itemtjnrlines = itemRow["提交内容"].ToString().Replace("\r", "").Split('\n').Where(x => !string.IsNullOrEmpty(x.Trim())).ToArray();
                                            for (int lineindex = 0; lineindex <= itemtjnrxlines.Length - 1; lineindex++) {
                                                //提交内容项
                                                string itemtjnrx = itemtjnrxlines[lineindex].ToString();
                                                //提交内容	
                                                string itemtjnr = itemtjnrlines[lineindex].ToString();
                                                if (!string.IsNullOrEmpty(itemtjnrx) && !string.IsNullOrEmpty(itemtjnr)) {
                                                    itemtjnrx = itemtjnrx.Substring(itemtjnrx.IndexOf(')') + 1);
                                                    itemtjnrx = itemtjnrx.Replace("：", "").Trim();
                                                    if (itemIndex == 0) {
                                                        resultTableRow["baowenbiaoshihao"] = itemtjnr;
                                                    }
                                                }
                                            }
                                            itemIndex++;
                                        } while (i + itemIndex < table.Rows.Count && string.IsNullOrEmpty(table.Rows[i + itemIndex]["验收案例编号"].ToString()));
                                    }
                                    resultTableRow["result"] = "人工处理";
                                    textbox.WriteLine("人工处理");
                                }
                                textbox.WriteLine("");
                            }
                        } catch (Exception exx) {
                            var defforColor = textbox.SelectionColor;
                            textbox.SelectionColor = Color.Red;
                            textbox.WriteLine("第" + (i + 1) + "行出错" + exx.Message);
                            textbox.SelectionColor = defforColor;
                        }

                    }
                }
                List<string> fileTypes = new List<string>();
                foreach (var item in dbresult.Keys) {
                    if (dbresult[item].Length >= 2) {
                        string filetypeName = dbresult[item][1];
                        if (!fileTypes.Contains(filetypeName)) {
                            fileTypes.Add(filetypeName);
                        }
                    }
                }
                foreach (var fileType in fileTypes) {
                    CommonHelper.FileName = fileType;
                    var rengonglist = (from myRow in resultTable.AsEnumerable()
                                       where !string.IsNullOrEmpty(myRow.Field<string>("filetype")) && myRow.Field<string>("filetype").Trim().Equals(fileType)
                                       select myRow).ToList();
                    foreach (var item in rengonglist) {
                        CommonHelper.WriteLog(item.Field<string>("验收案例编号") + " " + item.Field<string>("baowenbiaoshihao") + Environment.NewLine);
                    }
                    CommonHelper.WriteLog(fileType + ":" + rengonglist.Count);
                    textbox.WriteLine(fileType + ":总数:" + rengonglist.Count);
                }
                //-----------------------
                CommonHelper.FileName = "错误案例";
                var cwlist = (from myRow in resultTable.AsEnumerable()
                              where !string.IsNullOrEmpty(myRow.Field<string>("result")) && myRow.Field<string>("result").Trim().Equals("不通过")
                              select myRow).ToList();
                foreach (var item in cwlist) {
                    CommonHelper.WriteLog(item.Field<string>("验收案例编号") + Environment.NewLine);
                    CommonHelper.WriteLog(item.Field<string>("sql") + Environment.NewLine);
                    CommonHelper.WriteLog(item.Field<string>("log") + Environment.NewLine);
                }
                CommonHelper.WriteLog("错误案例:" + cwlist.Count);
                //----------------
                CommonHelper.FileName = "正确案例";
                var zqlist = (from myRow in resultTable.AsEnumerable()
                              where !string.IsNullOrEmpty(myRow.Field<string>("result")) && myRow.Field<string>("result").Trim().Equals("通过")
                              select myRow).ToList();
                textbox.WriteLine("正确案例:" + zqlist.Count);
                textbox.WriteLine("失败案例总数:" + cwlist.Count);
                textbox.WriteLine("正确案例总数:" + zqlist.Count);
                //保存结果到Excel中
                SaveExcel(excelfile, "IBPS验收案例检测标准", resultTable);
                textbox.WriteLine("执行结束");
            } catch (Exception ex) {
                textbox.WriteLine(ex.Message);
            }
        }


        /// <summary>
        /// 读取Txt结果
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        private static Dictionary<string, string> ReadTxt(string file) {
            string[] lines = File.ReadAllLines(file, Encoding.Default);
            Dictionary<string, string> data = new Dictionary<string, string>();
            foreach (var line in lines) {
                string lineStr = line.Trim();
                if (!string.IsNullOrEmpty(lineStr)) {
                    string[] lineArray = lineStr.Split(new string[] { "=" }, StringSplitOptions.RemoveEmptyEntries);
                    if (!data.ContainsKey(lineArray[0].Trim())) {
                        data.Add(lineArray[0].Trim(), lineArray[1].Trim());
                    } else {
                        data[lineArray[0].Trim()] = lineArray[1].Trim();
                    }
                }
            }
            return data;
        }

        private static Dictionary<string, string[]> ReadResult(string file) {
            string[] lines = File.ReadAllLines(file, Encoding.UTF8);
            Dictionary<string, string[]> data = new Dictionary<string, string[]>();
            foreach (var line in lines) {
                string lineStr = line.Trim();
                if (!string.IsNullOrEmpty(lineStr)) {
                    string[] lineArray = lineStr.Split(new string[] { "#" }, StringSplitOptions.RemoveEmptyEntries);
                    if (!data.ContainsKey(lineArray[0].Trim())) {
                        data.Add(lineArray[0].Trim(), lineArray);
                    }
                }
            }
            return data;
        }

        private static bool GetDbResult(string dbName, string sql) {
            bool result = CommonHelper.Context(dbName.ToUpper()).Sql(sql).QueryMany<dynamic>().Count > 0;
            return result;
            return true;
        }



        public static void SaveExcel(string file, string sheetName, DataTable table) {
            Workbook wkBook = new Workbook();
            wkBook.Open(file);
            Worksheet wkSheet = wkBook.Worksheets[sheetName];

            foreach (DataRow row in table.Rows) {
                if (!DBNull.Value.Equals(row["index"]) && !DBNull.Value.Equals(row["result"])) {
                    int x = Convert.ToInt32(row["index"]);
                    Cell celldb = null;
                    if (table.Columns.Contains("是否通过")) {
                        celldb = wkSheet.Cells[x, table.Columns["是否通过"].Ordinal];
                    } else if (table.Columns.Contains("是否" + Environment.NewLine + "通过")) {
                        celldb = wkSheet.Cells[x, table.Columns["是否" + Environment.NewLine + "通过"].Ordinal];
                    }
                    string db = row["result"].ToString();
                    if (!string.IsNullOrEmpty(db)) {
                        celldb.PutValue(db);
                        Style styleanli = celldb.GetStyle();
                        styleanli.Pattern = BackgroundType.Solid;
                        if (db.Equals("不通过")) {
                            styleanli.ForegroundColor = Color.Red;
                            celldb.SetStyle(styleanli);
                        } else if (db.Equals("通过")) {
                            styleanli.ForegroundColor = Color.Green;
                            celldb.SetStyle(styleanli);
                        } else { }
                    }
                }
            }
            wkBook.Save(file);
            //释放对象
            wkSheet = null;
            wkBook = null;
        }
    }
}
