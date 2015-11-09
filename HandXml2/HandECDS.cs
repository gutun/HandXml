using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using System.Xml;
using Aspose.Cells;
using System.Text.RegularExpressions;
using System.IO;
using System.Xml.Linq;
using System.Text;

namespace HandXml2 {
    public static class HandECDS {
        /// <summary>
        /// 处理Excel文件,ECDS处理程序的主入口
        /// </summary>
        /// <param name="textbox">程序界面的文本框,用于输出日志</param>
        /// <param name="file">要处理的Excel文件</param>
        public static void HandFile(System.Windows.Forms.RichTextBox textbox, string file, string txtfile, Dictionary<string, string[]> dic) {
            textbox.Clear();
            try {
                //读取Excel到内存中
                int titleRowIndex = 0;
                DataTable table = CommonHelper.ReadExcel(file, "ECDS案例验收规则", ref titleRowIndex);
                Dictionary<string, string> dbresult = ReadResult(txtfile);
                table.Columns.Add("节点路由日志", typeof(string));
                table.Columns.Add("接发路由日志", typeof(string));
                table.Columns.Add("报文类型规则日志", typeof(string));
                table.Columns.Add("报文规则日志", typeof(string));
                table.Columns.Add("MsgId", typeof(string));
                table.Columns.Add("MsgId日志", typeof(string));
                table.Columns.Add("数据库日志", typeof(string));
                //匹配Excel每一行
                for (int i = 0; i < table.Rows.Count; i++) {
                    try {
                        DataRow row = table.Rows[i];
                        //验收项目编号
                        string ysxmbh = row["验收项目编号"].ToString();
                        //节点路由
                        string jdly = row["节点路由"].ToString();
                        //接发路由
                        string jfly = row["接发路由"].ToString();
                        //报文类型规则
                        string bwlxgz = row["报文类型规则"].ToString();
                        //报文规则
                        string bwgz = row["报文规则"].ToString();
                        //提交内容	
                        string tjnr = row["提交内容"].ToString();
                        //入库
                        string rk = row["入库"].ToString();
                        //
                        string sxcl = row["上行处理"].ToString();

                        if (!string.IsNullOrEmpty(ysxmbh) && !string.IsNullOrEmpty(tjnr)) {
                            textbox.WriteLine("验收项目编号:" + ysxmbh);
                            if (ysxmbh == "ZR_ECDS123_001") {

                            }
                            //报文规则
                            string[] nodeRules = bwgz.Split(new string[] { ",", "，" }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim()).ToArray();
                            List<string> headers = new List<string>();
                            List<string> xmls = new List<string>();

                            #region  解析多个报文和xml

                            //多个xml
                            bool multiXml = false;
                            do {
                                tjnr = tjnr.Trim();
                                string header = tjnr.Substring(0, tjnr.IndexOf("}") + 1).Trim();
                                headers.Add(header);
                                tjnr = tjnr.Substring(tjnr.IndexOf("}") + 1).Trim();
                                if (tjnr.LastIndexOf("?xml") > 0 && tjnr.IndexOf("?xml") != tjnr.LastIndexOf("?xml")) {
                                    multiXml = true;
                                    string xml = tjnr.Substring(tjnr.IndexOf("?xml"), tjnr.LastIndexOf("?xml") - 1).Trim();
                                    xml = xml.Substring(0, xml.LastIndexOf(">") + 1);
                                    xmls.Add("<" + xml);
                                    tjnr = tjnr.Substring((xml.LastIndexOf(">") + 1)).Trim();
                                } else {
                                    multiXml = false;
                                    string xml = tjnr.Substring(tjnr.IndexOf("?xml")).Trim();
                                    xmls.Add("<" + xml);
                                }
                            } while (multiXml);
                            #endregion

                            //案例结果
                            bool anliresult = true;

                            #region 节点路由
                            StringBuilder sbjdly = new StringBuilder();
                            try {
                                //节点路由
                                if (!string.IsNullOrEmpty(jdly)) {
                                    string[] jdlyRules = jdly.Split(new string[] { ",", "，" }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim()).ToArray(); ;
                                    bool jdlyresult = true;

                                    if (headers.Count > 0) {
                                        //foreach (var headitem in headers)
                                        //{
                                        var headitem = headers.FirstOrDefault();
                                        string header = headitem.Substring(headitem.IndexOf("H:01") + 4);
                                        if (header.LastIndexOf("}") > -1) {
                                            header = header.Substring(0, header.LastIndexOf("}"));
                                        }
                                        var index = 0;
                                        var headerArray = header.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries).ToArray();
                                        foreach (var item in jdlyRules) {
                                            if (dic.ContainsKey(item)) {
                                                string[] itemValues = dic[item];
                                                if (headerArray.Length > 0 && index == 0) {
                                                    var result = itemValues.Contains(headerArray[0]);
                                                    jdlyresult = jdlyresult && result;
                                                    textbox.WriteLine("节点路由:" + item + "匹配结果是" + result);
                                                    sbjdly.AppendLine("节点路由:" + item + "匹配结果是" + result);
                                                    index++;
                                                } else if (headerArray.Length > 0 && index == 1) {
                                                    var result = itemValues.Contains(headerArray[2]);
                                                    jdlyresult = jdlyresult && result;
                                                    textbox.WriteLine("节点路由:" + item + "匹配结果是" + result);
                                                    sbjdly.AppendLine("节点路由:" + item + "匹配结果是" + result);
                                                    index++;
                                                } else {
                                                    jdlyresult = jdlyresult && false;
                                                    textbox.WriteLine("节点路由:" + item + "匹配结果是" + false);
                                                    sbjdly.AppendLine("验收项目编号:" + ysxmbh + " " + "节点路由:" + item + "匹配结果是" + false);
                                                    index++;
                                                }
                                            } else {
                                                jdlyresult = jdlyresult && false;
                                                textbox.WriteLine("节点路由:" + item + "匹配结果是" + false);
                                                sbjdly.AppendLine("验收项目编号:" + ysxmbh + " " + "节点路由:" + item + "匹配结果是" + false);
                                                index++;
                                            }
                                        }
                                        //}
                                    } else {
                                        jdlyresult = false;
                                    }
                                    anliresult = anliresult && jdlyresult;
                                    row["节点路由检查结果"] = jdlyresult ? "通过" : "不通过";
                                    textbox.WriteLine("节点路由匹配结果是" + jdlyresult, jdlyresult);
                                }
                            } catch (Exception exx) {
                                var defforColor = textbox.SelectionColor;
                                textbox.SelectionColor = Color.Red;
                                textbox.WriteLine("第" + (i + 2) + ":" + ysxmbh + "行出错   节点路由错误消息:" + exx.Message);
                                textbox.SelectionColor = defforColor;
                                sbjdly.AppendLine("第" + (i + 2) + ":" + ysxmbh + "行出错   节点路由错误消息:" + exx.Message);
                                anliresult = anliresult && false;
                                row["节点路由检查结果"] = "不通过";
                                textbox.WriteLine("节点路由匹配结果是" + false, false);
                            }
                            row["节点路由日志"] = sbjdly.ToString();
                            #endregion

                            #region 接发路由
                            StringBuilder sbjfly = new StringBuilder();
                            try {
                                //接发路由

                                if (!string.IsNullOrEmpty(jfly)) {
                                    string[] jflyRules = jfly.Split(new string[] { ",", "，" }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim()).ToArray(); ;
                                    bool jflyresult = true;
                                    if (headers.Count > 0) {
                                        //foreach (var headitem in headers)
                                        //{
                                        var headitem = headers.FirstOrDefault();
                                        string header = headitem.Substring(headitem.IndexOf("H:01") + 4);
                                        if (header.LastIndexOf("}") > -1) {
                                            header = header.Substring(0, header.LastIndexOf("}"));
                                        }
                                        var index = 0;
                                        var headerArray = header.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries).ToArray();
                                        foreach (var item in jflyRules) {
                                            if (dic.ContainsKey(item)) {
                                                string[] itemValues = dic[item];
                                                if (headerArray.Length > 0 && index == 0) {
                                                    var result = itemValues.Contains(headerArray[4]);
                                                    jflyresult = jflyresult && result;
                                                    textbox.WriteLine("接发路由:" + item + "匹配结果是" + result);
                                                    sbjfly.AppendLine("接发路由:" + item + "匹配结果是" + result);
                                                    index++;
                                                } else if (headerArray.Length > 0 && index == 1) {
                                                    var result = itemValues.Contains(headerArray[5]);
                                                    jflyresult = jflyresult && result;
                                                    textbox.WriteLine("接发路由:" + item + "匹配结果是" + result);
                                                    sbjfly.AppendLine("接发路由:" + item + "匹配结果是" + result);
                                                    index++;
                                                } else {
                                                    jflyresult = jflyresult && false;
                                                    textbox.WriteLine("接发路由:" + item + "匹配结果是" + false);
                                                    sbjfly.AppendLine("接发路由:" + item + "匹配结果是" + false);
                                                    index++;
                                                }
                                            } else {
                                                jflyresult = jflyresult && false;
                                                textbox.WriteLine("接发路由:" + item + "匹配结果是" + false);
                                                sbjfly.AppendLine("接发路由:" + item + "匹配结果是" + false);
                                                index++;
                                            }
                                        }
                                        //}
                                    } else {
                                        jflyresult = false;
                                    }
                                    anliresult = anliresult && jflyresult;
                                    row["接发路由检查结果"] = jflyresult ? "通过" : "不通过";
                                    textbox.WriteLine("接发路由匹配结果是" + jflyresult, jflyresult);
                                }
                            } catch (Exception exx) {
                                var defforColor = textbox.SelectionColor;
                                textbox.SelectionColor = Color.Red;
                                textbox.WriteLine("第" + (i + 2) + ":" + ysxmbh + "行出错   错误消息:" + exx.Message);
                                textbox.SelectionColor = defforColor;
                                sbjfly.AppendLine("第" + (i + 2) + ":" + ysxmbh + "行出错   错误消息:" + exx.Message);
                                anliresult = anliresult && false;
                                row["接发路由检查结果"] = "不通过";
                                textbox.WriteLine("接发路由匹配结果是" + false, false);
                            }
                            row["接发路由日志"] = sbjfly.ToString();
                            #endregion

                            #region 报文类型检查
                            StringBuilder sbbwlxgz = new StringBuilder();
                            try {
                                //报文类型检查
                                if (!string.IsNullOrEmpty(bwlxgz)) {
                                    bool bwlxgzresult = true;
                                    //foreach (var headitem in headers) 
                                    var headitem = headers.FirstOrDefault();
                                    if (null != headers) {
                                        string header = headitem.Substring(headitem.IndexOf("H:01") + 4);
                                        if (header.LastIndexOf("}") > -1) {
                                            header = header.Substring(0, header.LastIndexOf("}"));
                                        }
                                        List<string> listitem = header.Split(' ').Where(x => !string.IsNullOrEmpty(x)).ToList();
                                        var result = !string.IsNullOrEmpty(listitem[6]) && listitem[6].EndsWith(bwlxgz);
                                        textbox.WriteLine("报文类型:" + bwlxgz + "匹配结果是" + result);
                                        sbbwlxgz.AppendLine("报文类型:" + bwlxgz + "匹配结果是" + result);
                                        bwlxgzresult = bwlxgzresult && result;
                                    } else {
                                        bwlxgzresult = false;
                                    }
                                    //}
                                    anliresult = anliresult && bwlxgzresult;
                                    row["报文类型检查结果"] = bwlxgzresult ? "通过" : "不通过";
                                    textbox.WriteLine("报文类型检查结果" + bwlxgzresult, bwlxgzresult);
                                }
                            } catch (Exception exx) {
                                var defforColor = textbox.SelectionColor;
                                textbox.SelectionColor = Color.Red;
                                textbox.WriteLine("第" + (i + 2) + ":" + ysxmbh + "行出错   报文类型检查错误消息:" + exx.Message);
                                textbox.SelectionColor = defforColor;
                                sbbwlxgz.AppendLine("第" + (i + 2) + ":" + ysxmbh + "行出错   报文类型检查错误消息:" + exx.Message);
                                anliresult = anliresult && false;
                                row["报文类型检查结果"] = "不通过";
                                textbox.WriteLine("报文类型检查结果" + false, false);
                            }
                            row["报文类型规则日志"] = sbbwlxgz.ToString();
                            #endregion

                            #region 报文规则
                            StringBuilder sbbwgz = new StringBuilder();
                            try {
                                //报文规则
                                if (!string.IsNullOrEmpty(bwgz)) {
                                    //msgId
                                    string[] bwgzRule = bwgz.Split(new string[] { ",", "，" }, StringSplitOptions.RemoveEmptyEntries).Where(x => x.Contains("=")).Select(x => x.Trim()).ToArray();
                                    bool bwgzresult = true;
                                    foreach (var item in bwgzRule) {
                                        if (item.IndexOf("!=") > 0) {
                                            string[] itemDict = item.Split(new string[] { "!=", "！=" }, StringSplitOptions.RemoveEmptyEntries);
                                            string key = itemDict[0].Trim();
                                            string value = itemDict[1].Trim();
                                            string keyPath = "";
                                            string valuePath = "";
                                            int keyIndex = 0;
                                            int valueIndex = 0;
                                            if (key.StartsWith("[")) {
                                                keyIndex = int.Parse(key.Substring(key.IndexOf("[") + 1, key.IndexOf("]") - 1));
                                                keyPath = key.Substring(key.IndexOf("]") + 1);
                                            }
                                            if (value.StartsWith("[")) {
                                                valueIndex = int.Parse(value.Substring(value.IndexOf("[") + 1, value.IndexOf("]") - 1));
                                                valuePath = value.Substring(value.IndexOf("]") + 1);
                                            }
                                            if (string.IsNullOrEmpty(valuePath)) {
                                                bool result = GetXmlElement(xmls, key, value, false);
                                                bwgzresult = bwgzresult && result;
                                                textbox.WriteLine("规则:" + item + " 的结果是" + result);
                                                sbbwgz.AppendLine("规则:" + item + " 的结果是" + result);
                                            } else {
                                                if (keyIndex >= 0 && keyIndex < xmls.Count && valueIndex >= 0 && valueIndex < xmls.Count) {
                                                    bool result = GetXmlElement(xmls[keyIndex], keyPath, xmls[valueIndex], valuePath);
                                                    bwgzresult = bwgzresult && result;
                                                    textbox.WriteLine("规则:" + item + " 的结果是" + result);
                                                    sbbwgz.AppendLine("规则:" + item + " 的结果是" + result);
                                                } else {
                                                    bwgzresult = bwgzresult && false;
                                                    textbox.WriteLine("规则:" + item + "格式有误,找不到相应的xml");
                                                    sbbwgz.AppendLine("规则:" + item + "格式有误,找不到相应的xml");
                                                }
                                            }
                                        } else {
                                            string[] itemDict = item.Split(new string[] { "=" }, StringSplitOptions.RemoveEmptyEntries);
                                            string key = itemDict[0].Trim();
                                            string value = itemDict[1].Trim();
                                            string keyPath = "";
                                            string valuePath = "";
                                            int keyIndex = 0;
                                            int valueIndex = 0;
                                            if (key.StartsWith("[")) {
                                                keyIndex = int.Parse(key.Substring(key.IndexOf("[") + 1, key.IndexOf("]") - 1));
                                                keyPath = key.Substring(key.IndexOf("]") + 1);
                                            }
                                            if (value.StartsWith("[")) {
                                                valueIndex = int.Parse(value.Substring(value.IndexOf("[") + 1, value.IndexOf("]") - 1));
                                                valuePath = value.Substring(value.IndexOf("]") + 1);
                                            }
                                            if (string.IsNullOrEmpty(valuePath)) {
                                                bool result = GetXmlElement(xmls, key, value);
                                                bwgzresult = bwgzresult && result;
                                                textbox.WriteLine("规则:" + item + " 的结果是" + result);
                                                sbbwgz.AppendLine("规则:" + item + " 的结果是" + result);
                                            } else {
                                                if (keyIndex >= 0 && keyIndex < xmls.Count && valueIndex >= 0 && valueIndex < xmls.Count) {
                                                    bool result = GetXmlElement(xmls[keyIndex], keyPath, xmls[valueIndex], valuePath);
                                                    bwgzresult = bwgzresult && result;
                                                    textbox.WriteLine("规则:" + item + " 的结果是" + result);
                                                    sbbwgz.AppendLine("规则:" + item + " 的结果是" + result);
                                                } else {
                                                    bwgzresult = bwgzresult && false;
                                                    textbox.WriteLine("规则:" + item + "格式有误,找不到相应的xml");
                                                    sbbwgz.AppendLine("规则:" + item + "格式有误,找不到相应的xml");
                                                }
                                            }
                                        }

                                    }
                                    anliresult = anliresult && bwgzresult;
                                    row["报文xml检测结果"] = bwgzresult ? "通过" : "不通过";
                                    textbox.WriteLine("报文xml检测结果:" + bwgzresult, bwgzresult);
                                }
                            } catch (Exception exx) {
                                var defforColor = textbox.SelectionColor;
                                textbox.SelectionColor = Color.Red;
                                textbox.WriteLine("第" + (i + 2) + ":" + ysxmbh + "行出错   报文xml检测错误消息:" + exx.Message);
                                textbox.SelectionColor = defforColor;
                                sbbwgz.AppendLine("第" + (i + 2) + ":" + ysxmbh + "行出错   报文xml检测错误消息:" + exx.Message);
                                anliresult = anliresult && false;
                                row["报文xml检测结果"] = "不通过";
                                textbox.WriteLine("报文xml检测结果:" + false, false);
                            }
                            row["报文规则日志"] = sbbwgz.ToString();
                            #endregion

                            if (xmls.Count > 0) {
                                List<string> msgIdList = new List<string>();
                                foreach (var item in xmls) {
                                    string msgid = GetXmlMsgId(item);
                                    if (!string.IsNullOrEmpty(msgid)) {
                                        msgIdList.Add(msgid);
                                    }
                                }
                                row["MsgId"] = string.Join(",", msgIdList.ToArray());
                            }
                            #region 执行sql检查
                            if (rk.Trim().Equals("是") && !string.IsNullOrEmpty(sxcl.Trim())) {
                                StringBuilder sbmsgidlog = new StringBuilder();
                                List<string> msgIdList = new List<string>();
                                StringBuilder sbdblog = new StringBuilder();
                                if (sxcl.Trim().StartsWith("上行")) {
                                    string sql = string.Empty;
                                    bool sqlresult = dbresult.TryGetValue(ysxmbh, out sql);
                                    if (sqlresult && !string.IsNullOrEmpty(sql)) {
                                        #region 解析MsgId
                                        try {
                                            string msgidStr = "报文标示号";
                                            //解析MsgId
                                            bool msgidresult = true;
                                            if (xmls.Count > 0) {
                                                //foreach (var item in xmls)
                                                //{
                                                string msgid = GetXmlMsgId(xmls.FirstOrDefault());
                                                if (sql.IndexOf("原报文标示号") > -1) {
                                                    msgidStr = "原报文标示号";
                                                    msgid = GetXmlOrgnlMsgIdId(xmls.FirstOrDefault());
                                                } else if (sql.IndexOf("票号") > -1) {
                                                    msgidStr = "票号";
                                                    msgid = GetXmlIdNb(xmls.FirstOrDefault());
                                                }

                                                if (sql.IndexOf("原报文标示号") > -1) {
                                                    sql = sql.Replace("原报文标示号", msgid);
                                                } else if (sql.IndexOf("报文标示号") > -1) {
                                                    sql = sql.Replace("报文标示号", msgid);
                                                } else if (sql.IndexOf("票号") > -1) {
                                                    sql = sql.Replace("票号", msgid);
                                                }
                                                if (!string.IsNullOrEmpty(msgid)) {
                                                    try {
                                                        msgIdList.Add(msgid);
                                                        textbox.WriteLine(ysxmbh + "的sql:" + sql);
                                                        sbmsgidlog.AppendLine(ysxmbh + "的sql:" + sql);
                                                        var result = CheckMsgId(msgid, sql);
                                                        msgidresult = msgidresult && result;
                                                        textbox.WriteLine(msgid + "的sql查询结果是：" + result);
                                                        sbdblog.AppendLine(msgid + "的sql查询结果是：" + result);

                                                    } catch (Exception ex) {
                                                        msgidresult = false;
                                                        sbdblog.AppendLine(msgid + ",sql查询异常" + ex.Message);
                                                        textbox.WriteLine(msgid + ",sql查询异常" + ex.Message);
                                                    }
                                                } else {
                                                    msgidresult = false;
                                                    textbox.WriteLine("xml未找到" + msgidStr);
                                                    sbmsgidlog.AppendLine("xml未找到" + msgidStr);
                                                    anliresult = anliresult && false;
                                                    row["MsgId日志"] = "xml未找到" + msgidStr;
                                                }
                                                //}
                                                row["数据库检测结果"] = msgidresult ? "通过" : "不通过";
                                                textbox.WriteLine("数据库检测结果:" + msgidresult, msgidresult);
                                            } else {
                                                throw new Exception("提交内容xml有误!");
                                            }
                                            anliresult = anliresult && msgidresult;
                                        } catch (Exception exx) {
                                            var defforColor = textbox.SelectionColor;
                                            textbox.SelectionColor = Color.Red;
                                            textbox.WriteLine("第" + (i + 2) + ":" + ysxmbh + "行出错   数据库检测错误消息:" + exx.Message);
                                            textbox.SelectionColor = defforColor;
                                            sbmsgidlog.AppendLine("第" + (i + 2) + ":" + ysxmbh + "行出错   数据库检测错误消息:" + exx.Message);
                                            row["数据库检测结果"] = "不通过";
                                            textbox.WriteLine("数据库检测结果:" + false, false);
                                        }
                                        row["MsgId日志"] = sbmsgidlog.ToString();
                                        row["数据库日志"] = sbdblog.ToString();
                                        row["MsgId"] = string.Join(",", msgIdList.ToArray());
                                        #endregion
                                    } else {
                                        row["数据库检测结果"] = false ? "通过" : "不通过";
                                        anliresult = anliresult && false;
                                        row["MsgId日志"] = ysxmbh + "未找到sql";
                                        row["数据库日志"] = ysxmbh + "未找到sql";
                                        textbox.WriteLine(ysxmbh + "未找到sql");
                                        sbmsgidlog.AppendLine(ysxmbh + "未找到sql");
                                        textbox.WriteLine("数据库检测结果:" + false, false);
                                    }
                                } else if (sxcl.Trim().StartsWith("下行")) {
                                    #region 解析MsgId
                                    try {
                                        string msgidStr = "报文标示号";
                                        //解析MsgId
                                        bool msgidresult = true;
                                        if (xmls.Count > 0) {
                                            foreach (var item in xmls) {
                                                string msgid = GetXmlMsgId(item);
                                                if (!string.IsNullOrEmpty(msgid)) {
                                                    try {
                                                        msgIdList.Add(msgid);
                                                        var result = CheckMsgId(msgid);
                                                        msgidresult = msgidresult && result;
                                                        textbox.WriteLine(msgidStr + ":" + msgid + "的数据库匹配结果是：" + result);
                                                        sbdblog.AppendLine(msgidStr + ":" + msgid + "的数据库匹配结果是：" + result);
                                                    } catch (Exception ex) {
                                                        msgidresult = false;
                                                        sbdblog.AppendLine(msgidStr + ":" + msgid + ",数据库查询异常" + ex.Message);
                                                    }
                                                } else {
                                                    msgidresult = false;
                                                    textbox.WriteLine("xml未找到" + msgidStr);
                                                    sbmsgidlog.AppendLine("xml未找到" + msgidStr);
                                                }
                                            }
                                            row["数据库检测结果"] = msgidresult ? "通过" : "不通过";
                                            textbox.WriteLine("数据库检测结果:" + msgidresult, msgidresult);
                                        } else {
                                            throw new Exception("提交内容xml有误!");
                                        }
                                        anliresult = anliresult && msgidresult;
                                    } catch (Exception exx) {
                                        var defforColor = textbox.SelectionColor;
                                        textbox.SelectionColor = Color.Red;
                                        textbox.WriteLine("第" + (i + 2) + ":" + ysxmbh + "行出错   数据库检测错误消息:" + exx.Message);
                                        textbox.SelectionColor = defforColor;
                                        sbmsgidlog.AppendLine("第" + (i + 2) + ":" + ysxmbh + "行出错   数据库检测错误消息:" + exx.Message);
                                        row["数据库检测结果"] = "不通过";
                                        textbox.WriteLine("数据库检测结果:" + false, false);
                                    }
                                    row["MsgId日志"] = sbmsgidlog.ToString();
                                    row["数据库日志"] = sbdblog.ToString();
                                    row["MsgId"] = string.Join(",", msgIdList.ToArray());
                                    #endregion
                                } else { }
                            }
                            #endregion

                            //#region 解析MsgId
                            //StringBuilder sbmsgid = new StringBuilder();
                            //StringBuilder sbdb = new StringBuilder();
                            //try
                            //{
                            //    //解析MsgId
                            //    bool msgidresult = true;
                            //    if (xmls.Count > 0)
                            //    {
                            //        //foreach (var item in xmls)
                            //        //{
                            //        string msgid = GetXmlMsgId(xmls.FirstOrDefault());
                            //        if (!string.IsNullOrEmpty(msgid))
                            //        {
                            //            try
                            //            {
                            //                var result = CheckMsgId(msgid);
                            //                msgidresult = msgidresult && result;
                            //                textbox.WriteLine("MsgId:" + msgid + "的数据库匹配结果是：" + result);
                            //                sbmsgid.AppendLine("MsgId:" + msgid + "的数据库匹配结果是：" + result);

                            //            }
                            //            catch (Exception ex)
                            //            {
                            //                sbdb.AppendLine("MsgId:" + msgid + ",数据库查询异常" + ex.Message);
                            //            }
                            //        }
                            //        else
                            //        {
                            //            msgidresult = false;
                            //            textbox.WriteLine("xml未找到msgId");
                            //            sbmsgid.AppendLine("xml未找到msgId");
                            //        }
                            //        //}
                            //        row["数据库检测结果"] = msgidresult ? "通过" : "不通过";
                            //        textbox.WriteLine("数据库检测结果:" + msgidresult, msgidresult);
                            //    }
                            //    else
                            //    {
                            //        throw new Exception("提交内容xml有误!");
                            //    }
                            //}
                            //catch (Exception exx)
                            //{
                            //    var defforColor = textbox.SelectionColor;
                            //    textbox.SelectionColor = Color.Red;
                            //    textbox.WriteLine("第" + (i + 2) + ":" + ysxmbh + "行出错   数据库检测错误消息:" + exx.Message);
                            //    textbox.SelectionColor = defforColor;
                            //    sbmsgid.AppendLine("第" + (i + 2) + ":" + ysxmbh + "行出错   数据库检测错误消息:" + exx.Message);
                            //    row["数据库检测结果"] = "不通过";
                            //    textbox.WriteLine("数据库检测结果:" + false, false);
                            //}
                            //row["MsgId日志"] = sbmsgid.ToString();
                            //row["数据库日志"] = sbdb.ToString();
                            //#endregion

                            #region 案例检测结果
                            row["案例检测结果"] = anliresult ? "通过" : "不通过";
                            textbox.WriteLine("案例检测结果:" + anliresult, anliresult);
                            #endregion
                            //anliresult = anliresult && msgidresult;

                            textbox.WriteLine("");
                            CommonHelper.WriteLog("");
                        }
                    } catch (Exception exx) {
                        var defforColor = textbox.SelectionColor;
                        textbox.SelectionColor = Color.Red;
                        textbox.WriteLine("第" + (i + 2) + "行出错   错误消息:" + exx.Message);
                        textbox.SelectionColor = defforColor;
                    }

                }
                #region  统计虚拟表并输出日志到文件
                //保存结果到Excel中
                int okCount = (from myRow in table.AsEnumerable()
                               where !string.IsNullOrEmpty(myRow.Field<string>("验收项目编号")) && myRow.Field<string>("案例检测结果").Equals("通过")
                               select myRow).Count();
                int errorCount = (from myRow in table.AsEnumerable()
                                  where !string.IsNullOrEmpty(myRow.Field<string>("验收项目编号")) && myRow.Field<string>("案例检测结果").Equals("不通过")
                                  select myRow).Count();
                CommonHelper.FileName = "案例检测失败结果集";
                var erroranliList = (from myRow in table.AsEnumerable()
                                     where !string.IsNullOrEmpty(myRow.Field<string>("验收项目编号")) && myRow.Field<string>("案例检测结果").Equals("不通过")
                                     select myRow).ToList();
                foreach (var row in erroranliList) {
                    string xmbh = row["验收项目编号"].ToString();
                    string jdly = row["节点路由检查结果"].ToString();
                    CommonHelper.WriteLog("验收项目编号" + xmbh + Environment.NewLine);
                    CommonHelper.WriteLog(row["节点路由日志"].ToString());
                    CommonHelper.WriteLog(row["接发路由日志"].ToString());
                    CommonHelper.WriteLog(row["报文类型规则日志"].ToString());
                    CommonHelper.WriteLog(row["报文规则日志"].ToString());
                    CommonHelper.WriteLog(row["MsgId日志"].ToString());
                    CommonHelper.WriteLog(row["数据库日志"].ToString());
                    CommonHelper.WriteLog(Environment.NewLine);
                }

                textbox.WriteLine("通过案例数:" + okCount);
                CommonHelper.WriteLog("通过案例数:" + okCount);
                textbox.WriteLine("不通过案例数:" + errorCount);
                CommonHelper.WriteLog("不通过案例数:" + errorCount);
                int okdbCount = (from myRow in table.AsEnumerable()
                                 where !string.IsNullOrEmpty(myRow.Field<string>("验收项目编号")) && myRow.Field<string>("数据库检测结果").Equals("通过")
                                 select myRow).Count();
                int errordbCount = (from myRow in table.AsEnumerable()
                                    where !string.IsNullOrEmpty(myRow.Field<string>("验收项目编号")) && myRow.Field<string>("数据库检测结果").Equals("不通过")
                                    select myRow).Count();
                //--------------------------------------------------------------
                CommonHelper.FileName = "查看cics日志";
                var ckcslist = (from myRow in table.AsEnumerable()
                                where !string.IsNullOrEmpty(myRow.Field<string>("提交内容")) && myRow.Field<string>("入库").Trim().Equals("否") && !string.IsNullOrEmpty(myRow.Field<string>("上行处理"))
                                select myRow).ToList();
                foreach (var item in ckcslist) {
                    CommonHelper.WriteLog(item.Field<string>("验收项目编号") + ":" + item.Field<string>("MsgId") + Environment.NewLine);
                }
                CommonHelper.WriteLog("查看cics日志总数:" + ckcslist.Count);
                //----------------------------------------------------------------------
                CommonHelper.FileName = "查看cics日志+行内日志";
                var ckcshnlist = (from myRow in table.AsEnumerable()
                                  where !string.IsNullOrEmpty(myRow.Field<string>("提交内容")) && myRow.Field<string>("入库").Trim().Equals("否") && myRow.Field<string>("上行处理").Trim().Equals("上行_查看行内日志")
                                  select myRow).ToList();
                foreach (var item in ckcshnlist) {
                    CommonHelper.WriteLog(item.Field<string>("验收项目编号") + ":" + item.Field<string>("MsgId") + Environment.NewLine);
                }
                CommonHelper.WriteLog("查看cics日志+行内日志总数:" + ckcshnlist.Count);
                //----------------------------------------------------------------------
                CommonHelper.FileName = "查看行内日志";
                var ckhnlist = (from myRow in table.AsEnumerable()
                                where !string.IsNullOrEmpty(myRow.Field<string>("提交内容")) && myRow.Field<string>("入库").Trim().Equals("是") && (myRow.Field<string>("上行处理").Trim().Equals("上行_查看行内日志") || myRow.Field<string>("上行处理").Trim().Equals("下行"))
                                select myRow).ToList();
                foreach (var item in ckhnlist) {
                    CommonHelper.WriteLog(item.Field<string>("验收项目编号") + ":" + item.Field<string>("MsgId") + Environment.NewLine);
                }
                CommonHelper.WriteLog("查看行内日志总数:" + ckhnlist.Count);
                //----------------------------------------------------------------------
                CommonHelper.FileName = "人工处理";
                var rglist = (from myRow in table.AsEnumerable()
                              where !string.IsNullOrEmpty(myRow.Field<string>("提交内容")) && (myRow.Field<string>("入库").Trim().Equals("人工处理") || string.IsNullOrEmpty(myRow.Field<string>("入库")))
                              select myRow).ToList();
                foreach (var item in rglist) {
                    CommonHelper.WriteLog(item.Field<string>("验收项目编号") + ":" + item.Field<string>("MsgId") + Environment.NewLine);
                }
                CommonHelper.WriteLog("人工处理总数:" + rglist.Count);
                #endregion
                SaveExcel(file, table);
                textbox.WriteLine("执行结束");
                CommonHelper.WriteLog("执行结束");

            } catch (Exception ex) {
                textbox.WriteLine(ex.Message);
                CommonHelper.WriteError(ex.Message);
            }
        }


        /// <summary>
        /// 正则方式获取xml的内容
        /// </summary>
        /// <param name="xmls">xml文本</param>
        /// <param name="nodePath">读取的节点规则</param>
        /// <param name="value">对比的值</param>
        /// <param name="flag">true为等于,false为不等于</param>
        /// <returns></returns>
        private static bool GetXmlElement(List<string> xmls, string nodePath, string value, bool flag = true) {
            try {
                foreach (var xml in xmls) {
                    //使用xPath选择需要的节点
                    string[] nodeArray = nodePath.Split(new string[] { "\\", "\\\\", "/", "//" }, StringSplitOptions.RemoveEmptyEntries);
                    List<string> result = new List<string>();
                    int nodeIndex = 0;
                    string subNode = Regex.Replace(xml, "^[^<]", "");
                    subNode = Regex.Replace(subNode, @"\s", "");
                    do {
                        if (nodeIndex == 0) {
                            GetMatchCollection(nodeArray[nodeIndex], subNode, result);
                        } else {
                            List<string> noderesult = new List<string>();
                            foreach (var item in result) {
                                List<string> subNodeResult = new List<string>();
                                GetMatchCollection(nodeArray[nodeIndex], item, subNodeResult);
                                noderesult.AddRange(subNodeResult);
                            }
                            result = noderesult;
                        }
                        nodeIndex++;
                    } while (nodeIndex < nodeArray.Length);
                    if (result.Count > 0) {
                        if (flag) {
                            foreach (var item in result) {
                                if (value.Trim().Equals("NULL") && string.IsNullOrEmpty(item.Trim())) {
                                    return true;
                                } else if (item.Trim().Equals(value.Trim())) {
                                    return true;
                                }

                            }
                        } else {
                            foreach (var item in result) {
                                if (value.Trim().Equals("NULL") && !string.IsNullOrEmpty(item.Trim())) {
                                    return true;
                                } else if (!item.Trim().Equals(value.Trim())) {
                                    return true;
                                }
                            }
                        }
                    }
                }
                return false;
            } catch (Exception ex) {
                throw ex;
            }

        }

        /// <summary>
        /// 正则方式对比两个xmls节点的值
        /// </summary>
        /// <param name="xml1">第一个xml文本</param>
        /// <param name="nodePath1">第一个xml的节点</param>
        /// <param name="xml2">第二个xml文本</param>
        /// <param name="nodePath2">第二个xml的节点</param>
        /// <param name="flag">true为等于,false为不等于</param>
        /// <returns></returns>
        private static bool GetXmlElement(string xml1, string nodePath1, string xml2, string nodePath2, bool flag = true) {
            try {
                List<string> result1 = new List<string>();
                List<string> result2 = new List<string>();
                if (!string.IsNullOrEmpty(xml1)) {
                    //使用xPath选择需要的节点
                    string[] nodeArray = nodePath1.Split(new string[] { "\\", "\\\\", "/", "//" }, StringSplitOptions.RemoveEmptyEntries);

                    int nodeIndex = 0;
                    string subNode = Regex.Replace(xml1, "^[^<]", "");
                    subNode = Regex.Replace(subNode, @"\s", "");
                    do {
                        if (nodeIndex == 0) {
                            GetMatchCollection(nodeArray[nodeIndex], subNode, result1);
                        } else {
                            List<string> noderesult = new List<string>();
                            foreach (var item in result1) {
                                List<string> subNodeResult = new List<string>();
                                GetMatchCollection(nodeArray[nodeIndex], item, subNodeResult);
                                noderesult.AddRange(subNodeResult);
                            }
                            result1 = noderesult;
                        }
                        nodeIndex++;
                    } while (nodeIndex < nodeArray.Length);
                }
                if (!string.IsNullOrEmpty(xml2)) {
                    //使用xPath选择需要的节点
                    string[] nodeArray = nodePath2.Split(new string[] { "\\", "\\\\", "/", "//" }, StringSplitOptions.RemoveEmptyEntries);

                    int nodeIndex = 0;
                    string subNode = Regex.Replace(xml2, "^[^<]", "");
                    subNode = Regex.Replace(subNode, @"\s", "");
                    do {
                        if (nodeIndex == 0) {
                            GetMatchCollection(nodeArray[nodeIndex], subNode, result2);
                        } else {
                            List<string> noderesult = new List<string>();
                            foreach (var item in result2) {
                                List<string> subNodeResult = new List<string>();
                                GetMatchCollection(nodeArray[nodeIndex], item, subNodeResult);
                                noderesult.AddRange(subNodeResult);
                            }
                            result2 = noderesult;
                        }
                        nodeIndex++;
                    } while (nodeIndex < nodeArray.Length);
                }
                if (result1.Count > 0 && result2.Count > 0) {
                    if (flag) {
                        return result1.First().Trim().Equals(result2.First().Trim());
                    } else {
                        return !result1.First().Trim().Equals(result2.First().Trim());
                    }
                }
                return false;
            } catch (Exception ex) {
                throw ex;
            }

        }

        /// <summary>
        /// 正则循环递归读取某个节点,获得多个节点值的列表
        /// </summary>
        /// <param name="node">节点名称</param>
        /// <param name="xml">xml文本</param>
        /// <param name="result">返回结果值列表</param>
        private static void GetMatchCollection(string node, string xml, List<string> result) {
            Regex re = new Regex(string.Format("<{0}[^>]*>(?:(?!<{0}>).)*?</{0}>", node), RegexOptions.None);
            MatchCollection mc = re.Matches(xml);
            if (mc.Count > 0) {
                foreach (Match item in mc) {
                    string value = item.Value.Trim().Substring(node.Length + 2);
                    value = value.Substring(0, value.Length - (node.Length + 3));
                    //if (!string.IsNullOrEmpty(value)) {
                    result.Add(value);
                    //}
                }
            }

        }

        /// <summary>
        /// 正则方式读取xml的MsgId
        /// </summary>
        /// <param name="xml"></param>
        /// <returns></returns>
        private static string GetXmlMsgId(string xml) {
            try {
                //使用xPath选择需要的节点
                if (!string.IsNullOrEmpty(xml)) {
                    List<string> result = new List<string>();
                    string[] nodeArray = new string[] { "MsgId", "Id" };
                    int nodeIndex = 0;
                    string subNode = Regex.Replace(xml, "^[^<]", "");
                    subNode = Regex.Replace(subNode, @"\s", "");
                    do {
                        if (nodeIndex == 0) {
                            GetMatchCollection(nodeArray[nodeIndex], subNode, result);
                        } else {
                            List<string> noderesult = new List<string>();
                            foreach (var item in result) {
                                List<string> subNodeResult = new List<string>();
                                GetMatchCollection(nodeArray[nodeIndex], item, subNodeResult);
                                noderesult.AddRange(subNodeResult);
                            }
                            result = noderesult;
                        }
                        nodeIndex++;
                    } while (nodeIndex < nodeArray.Length);
                    if (result.Count > 0) {
                        return result.First();
                    }
                }
                return null;
            } catch (Exception ex) {

                throw ex;
            }
        }

        private static string GetXmlOrgnlMsgIdId(string xml) {
            try {
                //使用xPath选择需要的节点
                if (!string.IsNullOrEmpty(xml)) {
                    List<string> result = new List<string>();
                    string[] nodeArray = new string[] { "OrgnlMsgId", "Id" };
                    int nodeIndex = 0;
                    string subNode = Regex.Replace(xml, "^[^<]", "");
                    subNode = Regex.Replace(subNode, @"\s", "");
                    do {
                        if (nodeIndex == 0) {
                            GetMatchCollection(nodeArray[nodeIndex], subNode, result);
                        } else {
                            List<string> noderesult = new List<string>();
                            foreach (var item in result) {
                                List<string> subNodeResult = new List<string>();
                                GetMatchCollection(nodeArray[nodeIndex], item, subNodeResult);
                                noderesult.AddRange(subNodeResult);
                            }
                            result = noderesult;
                        }
                        nodeIndex++;
                    } while (nodeIndex < nodeArray.Length);
                    if (result.Count > 0) {
                        return result.First();
                    }
                }
                return null;
            } catch (Exception ex) {

                throw ex;
            }
        }

        private static string GetXmlIdNb(string xml) {
            try {
                //使用xPath选择需要的节点
                if (!string.IsNullOrEmpty(xml)) {
                    List<string> result = new List<string>();
                    string[] nodeArray = new string[] { "IdNb" };
                    int nodeIndex = 0;
                    string subNode = Regex.Replace(xml, "^[^<]", "");
                    subNode = Regex.Replace(subNode, @"\s", "");
                    do {
                        if (nodeIndex == 0) {
                            GetMatchCollection(nodeArray[nodeIndex], subNode, result);
                        } else {
                            List<string> noderesult = new List<string>();
                            foreach (var item in result) {
                                List<string> subNodeResult = new List<string>();
                                GetMatchCollection(nodeArray[nodeIndex], item, subNodeResult);
                                noderesult.AddRange(subNodeResult);
                            }
                            result = noderesult;
                        }
                        nodeIndex++;
                    } while (nodeIndex < nodeArray.Length);
                    if (result.Count > 0) {
                        return result.First();
                    }
                }
                return null;
            } catch (Exception ex) {

                throw ex;
            }
        }
        /// <summary>
        /// 检测数据库的MsgId
        /// </summary>
        /// <param name="msgId"></param>
        /// <returns></returns>
        private static bool CheckMsgId(string msgId) {
            //string sqlStr = System.Configuration.ConfigurationManager.AppSettings["sql"].ToString();
            //string sql = string.Format(sqlStr, msgId);
            //string result = CommonHelper.Context("ECDS").Sql(sql).QuerySingle<string>();
            //return !string.IsNullOrEmpty(result);
            return true;
        }

        private static bool CheckMsgId(string msgId, string sql) {
            //return CommonHelper.Context("ECDS").Sql(sql).QueryMany<dynamic>().Count > 0;
            return true;
        }

        private static Dictionary<string, string> ReadResult(string file) {
            string[] lines = File.ReadAllLines(file, Encoding.UTF8);
            Dictionary<string, string> data = new Dictionary<string, string>();
            foreach (var line in lines) {
                string lineStr = line.Trim();
                if (!string.IsNullOrEmpty(lineStr)) {
                    string[] lineArray = lineStr.Split(new string[] { "#" }, StringSplitOptions.RemoveEmptyEntries);
                    if (!data.ContainsKey(lineArray[0].Trim())) {
                        data.Add(lineArray[0].Trim(), lineArray[1]);
                    }
                }
            }
            return data;
        }

        #region XPath方式处理xml

        /// <summary>
        /// XPath方式获取xml的内容
        /// </summary>
        /// <param name="xmls">xml文本</param>
        /// <param name="nodePath">读取的节点规则</param>
        /// <param name="value">对比的值</param>
        /// <param name="flag">true为等于,false为不等于</param>
        /// <returns></returns>
        private static bool ReadXmlElement(List<string> xmls, string nodePath, string value, bool flag = true) {
            try {
                foreach (var xml in xmls) {
                    XmlDocument doc = new XmlDocument();
                    doc.LoadXml(xml);
                    //使用xPath选择需要的节点
                    nodePath = "//" + string.Join("/", nodePath.Split(new string[] { "\\", "\\\\", "/", "//" }, StringSplitOptions.RemoveEmptyEntries));
                    if (flag) {

                        XmlNodeList nodes = doc.SelectNodes(nodePath + "[text()='" + value + "']");
                        if (nodes.Count > 0) {
                            return true;
                        }
                    } else {
                        XmlNodeList nodes = doc.SelectNodes(nodePath + "[text()!='" + value + "']");
                        if (nodes.Count > 0) {
                            return true;
                        }
                    }
                }
                return false;
            } catch (Exception ex) {
                throw ex;
            }
        }

        /// <summary>
        /// XPath方式对比两个xmls节点的值
        /// </summary>
        /// <param name="xml1">第一个xml文本</param>
        /// <param name="nodePath1">第一个xml的节点</param>
        /// <param name="xml2">第二个xml文本</param>
        /// <param name="nodePath2">第二个xml的节点</param>
        /// <param name="flag">true为等于,false为不等于</param>
        /// <returns></returns>
        private static bool ReadXmlElement(string xml1, string nodePath1, string xml2, string nodePath2, bool flag = true) {
            try {
                XmlDocument doc1 = new XmlDocument();
                xml1 = Regex.Replace(xml1, "^[^<]", "");
                doc1.LoadXml(xml1);
                XmlDocument doc2 = new XmlDocument();
                xml2 = Regex.Replace(xml2, "^[^<]", "");
                doc2.LoadXml(xml2);
                //使用xPath选择需要的节点
                nodePath1 = "//" + string.Join("/", nodePath1.Split(new string[] { "\\", "\\\\", "/", "//" }, StringSplitOptions.RemoveEmptyEntries));
                nodePath2 = "//" + string.Join("/", nodePath2.Split(new string[] { "\\", "\\\\", "/", "//" }, StringSplitOptions.RemoveEmptyEntries));
                XmlNodeList nodes1 = doc1.SelectNodes(nodePath1);
                XmlNodeList nodes2 = doc2.SelectNodes(nodePath2);
                if (nodes1.Count > 0 && nodes2.Count > 0) {
                    XmlNode node1 = nodes1[0];
                    XmlNode node2 = nodes2[0];
                    if (flag) {
                        if (node1.InnerText.Trim().Equals(node2.InnerText.Trim())) {
                            return true;
                        }
                    } else {
                        if (!node1.InnerText.Trim().Equals(node2.InnerText.Trim())) {
                            return true;
                        }
                    }
                }
                return false;
            } catch (Exception ex) {
                throw ex;
            }

        }
        /// <summary>
        /// XPath方式读取xml的MsgId
        /// </summary>
        /// <param name="xml"></param>
        /// <returns></returns>
        private static string ReadXmlMsgId(string xml) {
            try {
                XmlDocument doc = new XmlDocument();
                xml = Regex.Replace(xml, "^[^<]", "");
                doc.LoadXml(xml);
                //使用xPath选择需要的节点
                XmlNodeList nodes = doc.SelectNodes("//MsgId/Id");
                if (nodes.Count > 0) {
                    return nodes[0].InnerText.Trim();
                }
                return null;
            } catch (Exception ex) {

                throw ex;
            }
        }

        #endregion

        /// <summary>
        /// 虚拟的Datatable保存到Excel
        /// </summary>
        /// <param name="file">文件名称</param>
        /// <param name="table">虚拟表</param>
        public static void SaveExcel(string file, DataTable table) {
            Workbook wkBook = new Workbook();
            wkBook.Open(file);
            Worksheet wkSheet = wkBook.Worksheets["ECDS案例验收规则"];

            //遍历行
            for (int x = 0; x < wkSheet.Cells.MaxDataRow + 1; x++) {
                if (x != 0) {
                    var row = table.Rows[x - 1];
                    string xmbh = row["验收项目编号"].ToString();
                    if (!string.IsNullOrEmpty(xmbh)) {
                        Cell celljdly = wkSheet.Cells[x, table.Columns["节点路由检查结果"].Ordinal];
                        string jdly = row["节点路由检查结果"].ToString();
                        if (!string.IsNullOrEmpty(jdly)) {
                            celljdly.PutValue(jdly);
                            //Style stylejdly = celljdly.GetStyle();
                            //stylejdly.Pattern = BackgroundType.Solid;
                            //stylejdly.ForegroundColor = jdly.Equals("不通过") ? Color.Red : Color.Green;
                            //celljdly.SetStyle(stylejdly);
                        }

                        Cell celljfly = wkSheet.Cells[x, table.Columns["接发路由检查结果"].Ordinal];
                        string jfly = row["接发路由检查结果"].ToString();
                        if (!string.IsNullOrEmpty(jfly)) {
                            celljfly.PutValue(jfly);
                            //Style stylejfly = celljfly.GetStyle();
                            //stylejfly.Pattern = BackgroundType.Solid;
                            //stylejfly.ForegroundColor = jfly.Equals("不通过") ? Color.Red : Color.Green;
                            //celljfly.SetStyle(stylejfly);
                        }

                        Cell cellbwlx = wkSheet.Cells[x, table.Columns["报文类型检查结果"].Ordinal];
                        string bwlx = row["报文类型检查结果"].ToString();
                        if (!string.IsNullOrEmpty(bwlx)) {
                            cellbwlx.PutValue(bwlx);
                            //Style stylebwlx = cellbwlx.GetStyle();
                            //stylebwlx.Pattern = BackgroundType.Solid;
                            //stylebwlx.ForegroundColor = bwlx.Equals("不通过") ? Color.Red : Color.Green;
                            //cellbwlx.SetStyle(stylebwlx);
                        }

                        Cell cellxml = wkSheet.Cells[x, table.Columns["报文xml检测结果"].Ordinal];
                        string xml = row["报文xml检测结果"].ToString();
                        if (!string.IsNullOrEmpty(xml)) {
                            cellxml.PutValue(xml);
                            //Style stylexml = cellxml.GetStyle();
                            //stylexml.Pattern = BackgroundType.Solid;
                            //stylexml.ForegroundColor = xml.Equals("不通过") ? Color.Red : Color.Green;
                            //cellxml.SetStyle(stylexml);
                        }

                        Cell celldb = wkSheet.Cells[x, table.Columns["数据库检测结果"].Ordinal];
                        string db = row["数据库检测结果"].ToString();
                        if (!string.IsNullOrEmpty(db)) {
                            celldb.PutValue(db);
                            //Style styledb = celldb.GetStyle();
                            //styledb.Pattern = BackgroundType.Solid;
                            //styledb.ForegroundColor = db.Equals("不通过") ? Color.Red : Color.Green;
                            //celldb.SetStyle(styledb);
                        }

                        Cell cellanli = wkSheet.Cells[x, table.Columns["案例检测结果"].Ordinal];
                        string anli = row["案例检测结果"].ToString();
                        if (!string.IsNullOrEmpty(anli)) {
                            cellanli.PutValue(anli);
                            Style styleanli = cellanli.GetStyle();
                            styleanli.Pattern = BackgroundType.Solid;
                            styleanli.ForegroundColor = anli.Equals("不通过") ? Color.Red : Color.Green;
                            cellanli.SetStyle(styleanli);
                        }
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
