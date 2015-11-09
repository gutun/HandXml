using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;

namespace HandXml2 {
    public static class HandCFXPS {
        public static void HandFile(System.Windows.Forms.RichTextBox textbox, string excelfile, string txtfile, List<string> sheets) {
            textbox.Clear();
            try {
                DataSet resultSet = new DataSet();
                Dictionary<string, string[]> dbresult = ReadResult(txtfile);
                Dictionary<string, int> titleRowIndexs = new Dictionary<string, int>();
                foreach (var sheet in sheets) {
                    //读取Excel到内存中
                    int titleRowIndex = 0;
                    DataTable table = CommonHelper.ReadExcel(excelfile, sheet, ref titleRowIndex);
                    if (null != table) {
                        DataTable resultTable = table.Copy();
                        resultTable.TableName = sheet;
                        resultTable.Columns.Add("index", typeof(int));
                        resultTable.Columns.Add("sql", typeof(string));
                        resultTable.Columns.Add("result", typeof(string));
                        titleRowIndexs.Add(sheet, titleRowIndex);
                        resultSet.Tables.Add(resultTable);
                    }
                }
                foreach (DataTable resultTable in resultSet.Tables) {
                    int titleRowIndex = titleRowIndexs[resultTable.TableName];
                    //匹配Excel每一行
                    for (int i = 0; i < resultTable.Rows.Count; i++) {
                        DataRow row = resultTable.Rows[i];
                        //验收项目编号
                        string ysalbh = row["验收案例编号"].ToString();
                        //提交项目
                        string tjnrx = row["提交项目"].ToString();
                        //提交内容	
                        string tjnr = row["提交内容"].ToString();

                        if (!string.IsNullOrEmpty(ysalbh) && !ysalbh.Trim().Equals("验收案例编号") && !string.IsNullOrEmpty(tjnrx) && !tjnrx.Trim().Equals("提交项目") && !string.IsNullOrEmpty(tjnr) && !tjnr.Trim().Equals("提交内容")) {
                            try {
                                textbox.WriteLine("验收案例编号:" + ysalbh);

                                if (!dbresult.ContainsKey(ysalbh)) {
                                    DataRow resultTableRow = resultTable.Rows[i];
                                    resultTableRow["index"] = i + titleRowIndex + 1;
                                    resultTableRow["result"] = "不通过";
                                    //不做处理
                                    textbox.WriteLine("找不到相关的sql");
                                } else {
                                    if (dbresult[ysalbh].Length == 3) {
                                        DataRow resultTableRow = resultTable.Rows[i];
                                        resultTableRow["index"] = i + titleRowIndex + 1;
                                        string[] sqlInfo = null;
                                        string sql = string.Empty;
                                        bool sqlresult = dbresult.TryGetValue(ysalbh, out sqlInfo);
                                        int itemIndex = 0;
                                        if (sqlresult && null != sqlInfo && sqlInfo.Length > 0) {
                                            sql = sqlInfo[2];
                                            do {
                                                DataRow itemRow = resultTable.Rows[i + itemIndex];
                                                //验收项目编号
                                                string itemysalbh = itemRow["验收案例编号"].ToString();
                                                string[] itemtjnrxlines = itemRow["提交项目"].ToString().Replace("\r", "").Split('\n').Where(x => !string.IsNullOrEmpty(x.Trim())).ToArray();
                                                string[] itemtjnrlines = itemRow["提交内容"].ToString().Replace("\r", "").Split('\n').Where(x => !string.IsNullOrEmpty(x.Trim())).ToArray();
                                                for (int lineindex = 0; lineindex <= itemtjnrxlines.Length - 1; lineindex++) {
                                                    //提交项目
                                                    string itemtjnrx = itemtjnrxlines[lineindex].ToString();
                                                    //提交内容	
                                                    string itemtjnr = itemtjnrlines[lineindex].ToString();
                                                    if (!string.IsNullOrEmpty(itemtjnrx) && !string.IsNullOrEmpty(itemtjnr)) {
                                                        itemtjnrx = itemtjnrx.Substring(itemtjnrx.IndexOf(')') + 1);
                                                        itemtjnrx = itemtjnrx.Replace("：", "").Trim();
                                                        if (sql.IndexOf("TXID='报文标识号'") > -1 && itemtjnrx.Equals("报文标识号")) {
                                                            sql = sql.Replace("TXID='报文标识号'", "TXID='" + itemtjnr.Trim().Substring(itemtjnr.Trim().Length - 8, 8) + "'");
                                                        } else if (sql.IndexOf("CURCODE = '外币种类'") > -1) {
                                                            sql = sql.Replace("CURCODE = '外币种类'", "CURCODE = '" + resultTable.TableName.Replace("案例","").Trim() + "'");
                                                        } else {
                                                            sql = sql.Replace("'" + itemtjnrx + "'", "'" + itemtjnr + "'");
                                                        }
                                                    }
                                                }
                                                itemIndex++;
                                            } while (i + itemIndex < resultTable.Rows.Count && string.IsNullOrEmpty(resultTable.Rows[i + itemIndex]["验收案例编号"].ToString()));
                                        }
                                        textbox.WriteLine("sql语句:" + sql);
                                        resultTableRow["sql"] = sql;
                                        bool sqlResult = GetDbResult(sqlInfo[1], sql);
                                        resultTableRow["result"] = sqlResult ? "通过" : "不通过";
                                        textbox.WriteLine("检查结果:" + sqlResult);
                                    } else {
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
                    CommonHelper.FileName = resultTable.TableName + "人工处理";
                    var ckcslist = (from myRow in resultTable.AsEnumerable()
                                    where !string.IsNullOrEmpty(myRow.Field<string>("result")) && !myRow.Field<string>("result").Trim().Equals("通过")
                                    select myRow).ToList();
                    foreach (var item in ckcslist) {
                        CommonHelper.WriteLog(item.Field<string>("验收案例编号") + Environment.NewLine);
                    }
                    CommonHelper.WriteLog(resultTable.TableName + "人工处理:" + ckcslist.Count);
                    //保存结果到Excel中
                    SaveExcel(excelfile, "CFXPS验收案例检测标准", resultTable);
                }
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
                    Cell celldb = wkSheet.Cells[x, table.Columns["是否"+Environment.NewLine+"通过"].Ordinal];
                    string db = row["result"].ToString();
                    if (!string.IsNullOrEmpty(db)) {
                        celldb.PutValue(db);
                        Style styleanli = celldb.GetStyle();
                        styleanli.Pattern = BackgroundType.Solid;
                        styleanli.ForegroundColor = db.Equals("不通过") ? Color.Red : Color.Green;
                        celldb.SetStyle(styleanli);
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
