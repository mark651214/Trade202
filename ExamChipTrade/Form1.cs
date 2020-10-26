using HtmlAgilityPack;
//using HtmlAgilityPack;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.IO.MemoryMappedFiles;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExamChipTrade
{
    public partial class Form1 : Form
    {


        HashSet<string> TseStock = new HashSet<string>();
        HashSet<string> OtcStock = new HashSet<string>();

        HashSet<string> CandidateFinMarStock = new HashSet<string>();
        HashSet<string> CandidateParty3Stock = new HashSet<string>();
        //HashSet<string> CandidateWeekChipSet = new HashSet<string>();
        HashSet<DateTime> workingdays = new HashSet<DateTime>();


        Dictionary<string, Dictionary<DateTime, int>> dicStockFinancing = new Dictionary<string, Dictionary<DateTime, int>>();
        Dictionary<string, Dictionary<DateTime, int>> dicStockMarriage = new Dictionary<string, Dictionary<DateTime, int>>();

        Dictionary<string, Dictionary<DateTime, double>> dicFinanceRate = new Dictionary<string, Dictionary<DateTime, double>>();
        Dictionary<string, Dictionary<DateTime, double>> dicMarriageRate = new Dictionary<string, Dictionary<DateTime, double>>();

        Dictionary<string, Dictionary<DateTime, PartyParam>> dicStock3Party = new Dictionary<string, Dictionary<DateTime, PartyParam>>();
        Dictionary<string, Dictionary<DateTime, PartyParam>> dicStock3Party2 = new Dictionary<string, Dictionary<DateTime, PartyParam>>();
        //Dictionary<string, Dictionary<DateTime, PartyParamLong>> dicStock3PartyLong = new Dictionary<string, Dictionary<DateTime, PartyParamLong>>();
        
        Dictionary<string, Dictionary<DateTime, double[]>> dbEvenydayOLHCDic = new Dictionary<string, Dictionary<DateTime, double[]>>();
        Dictionary<string, Dictionary<DateTime, long>> dbStockVolDic = new Dictionary<string, Dictionary<DateTime, long>>();

        Dictionary<string, int[]> dicFinancingSum = new Dictionary<string, int[]>();
        Dictionary<string, int[]> dicMarriageSum = new Dictionary<string, int[]>();
        Dictionary<string, string> dicStockNoMapping = new Dictionary<string, string>();

        string strMarginPath = "D:\\Chips\\Margin\\";
        string strDayPricePath = "D:\\Chips\\eachdaystock\\";
        string str3PartyPath = "D:\\Chips\\3Party\\";

        public Form1()
        {
            InitializeComponent();
        }


        private void button1_Click(object sender, EventArgs e)
        {
            int iMaxDays = 2;
            string strmsg = string.Format("最近{0}天融資融券新高",iMaxDays);
            Console.WriteLine(strmsg);
            // Console
            DateTime dtStart = monthCalendar1.SelectionStart;
            DateTime dtLoopStart = dtStart;

            HashSet<string> hs = new HashSet<string>();
            foreach (string keys in dicStockFinancing.Keys)
            {
                if (dicStockNoMapping.ContainsKey(keys) && dicStockFinancing[keys].Values.Max() > 0 && (dicStockFinancing[keys].Values.Max() == dicStockFinancing[keys].Values.ElementAt(0) || dicStockFinancing[keys].Values.Max() == dicStockFinancing[keys].Values.ElementAt(1)))
                {
                    hs.Add(keys);
                    strmsg = string.Format("{0},{1},FinMax {2:F4},Today {3:F4}", keys, dicStockNoMapping[keys], dicStockFinancing[keys].Values.Max(), dicStockFinancing[keys].Values.ElementAt(0));
                    Console.WriteLine(strmsg);
                }               
            }

            foreach (string keys in dicStockMarriage.Keys)
            {
                if (dicStockNoMapping.ContainsKey(keys) && dicStockMarriage[keys].Values.Max() > 0 && (dicStockMarriage[keys].Values.Max() == dicStockMarriage[keys].Values.ElementAt(0) || dicStockMarriage[keys].Values.Max() == dicStockMarriage[keys].Values.ElementAt(1)))
                {
                    string stradd = "";
                    if (hs.Contains(keys))
                        stradd = "*";
                    if (dicStockMarriage[keys].Values.ElementAt(0) > dicStockMarriage[keys].Values.ElementAt(1)*2.9)
                        stradd += "*3";
                    strmsg = string.Format("{0},{1},MarMax {2:F4},Today {3:F4},{4}", keys, dicStockNoMapping[keys], dicStockMarriage[keys].Values.Max(), dicStockMarriage[keys].Values.ElementAt(0), stradd);
                    Console.WriteLine(strmsg);
                }
            }

            //foreach (string keys in dicFinancing.Keys)
            //{
            //    //if (dicFinancing[keys].ElementAt(0) > (dicFinancingAvg[keys] + dicFinancingDev[keys] * 3))// && dicMarriage[keys].ElementAt(0) > (dicMarriageAvg[keys] + dicMarriageDev[keys] * 3))
            //    double dbmax = dicFinancing[keys].Max();
            //    if (dicFinancing[keys].ElementAt(2) == dbmax || dicFinancing[keys].ElementAt(1) == dbmax || dicFinancing[keys].ElementAt(0) == dbmax)
            //    {
            //        Console.WriteLine(keys);
            //        CandidateFinMarStock.Add(keys);
            //    }
            //}
            

        }

        public static Byte[] ReadMMFAllBytes(string fileName)
        {
            using (var mmf = MemoryMappedFile.OpenExisting(fileName))
            {
                using (var stream = mmf.CreateViewStream())
                {
                    using (BinaryReader binReader = new BinaryReader(stream))
                    {
                        return binReader.ReadBytes((int)stream.Length);
                    }
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            /*
            string[] lines = File.ReadAllLines("tse.txt");
            foreach (string line in lines)
            {

                string[] strSp = Regex.Split(line, @"[\s]+");
                if(strSp.Length == 8)
                {
                    string[] strSp2 = strSp[0].Split('\r');
                    if(strSp[0].Length == 4)
                    {
                        TseStock.Add(strSp[0]);
                    }
                }
                int m = 0;
            }
            string[] linesotc = File.ReadAllLines("otc.txt");
            foreach (string line in linesotc)
            {
                string[] strSp = Regex.Split(line, @"[\s]+");
                if (strSp.Length == 8)
                {
                    if (strSp[0].Length == 4)
                    {
                        OtcStock.Add(strSp[0]);
                    }
                }
                int m = 0;
            }    */
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;            
            ServicePointManager.DefaultConnectionLimit = 10;
            button4_Click(null, null);
        
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DateTime dtStartDay =  new DateTime(DateTime.Now.Year,DateTime.Now.Month,DateTime.Now.Day);
            DateTime dtEnd = dtStartDay.AddYears(-1);
            for (int i = 0; dtEnd < dtStartDay; i++)
            {
                bool bisworkday = false;
                string strFileNameMargin = string.Format("D:\\Chips\\Margin\\{0:0000}{1:00}{2:00}.json", dtStartDay.Year, dtStartDay.Month, dtStartDay.Day);
                if (File.Exists(strFileNameMargin))
                {
                    string json = File.ReadAllText(strFileNameMargin);
                    Dictionary<string, object> json_Dictionary = JsonConvert.DeserializeObject<Dictionary<string, object>>(json);
                    if(json_Dictionary.ContainsKey("data"))
                    {

                        List<List<string>> jsondata_ListList = JsonConvert.DeserializeObject<List<List<string>>>(json_Dictionary["data"].ToString());
                        foreach (List<string> marginelist in jsondata_ListList)
                        {
                            string strno = marginelist.ElementAt(0);
                            string strFin = marginelist.ElementAt(6);
                            string strFinLimit = marginelist.ElementAt(7);
                            string strMar = marginelist.ElementAt(12);
                            string strMarLimit = marginelist.ElementAt(13);

                            int iFin = int.Parse(strFin,NumberStyles.AllowThousands);
                            int iMar = int.Parse(strMar, NumberStyles.AllowThousands);
                            int iFinLimit = int.Parse(strFinLimit, NumberStyles.AllowThousands);
                            int iMarLimit = int.Parse(strMarLimit, NumberStyles.AllowThousands);

                            if (!dicStockFinancing.ContainsKey(strno))
                            {
                                dicStockFinancing.Add(strno, new Dictionary<DateTime, int>());
                            }
                            dicStockFinancing[strno].Add(new DateTime(dtStartDay.Year, dtStartDay.Month, dtStartDay.Day), iFin);

                            if (!dicStockMarriage.ContainsKey(strno))
                            {
                                dicStockMarriage.Add(strno, new Dictionary<DateTime, int>());
                            }
                            dicStockMarriage[strno].Add(new DateTime(dtStartDay.Year, dtStartDay.Month, dtStartDay.Day), iMar);


                            if (!dicFinanceRate.ContainsKey(strno))
                            {
                                dicFinanceRate.Add(strno, new Dictionary<DateTime, double>());
                            }
                            dicFinanceRate[strno].Add(new DateTime(dtStartDay.Year, dtStartDay.Month, dtStartDay.Day), (double)iFin / (double)iFinLimit);
                            if (!dicMarriageRate.ContainsKey(strno))
                            {
                                dicMarriageRate.Add(strno, new Dictionary<DateTime, double>());
                            }
                            dicMarriageRate[strno].Add(new DateTime(dtStartDay.Year, dtStartDay.Month, dtStartDay.Day), (double)iMar / (double)iMarLimit);

                            bisworkday = true;
                        }
                    }

                   
                }
                string strFileNameMarginTpex = string.Format("D:\\Chips\\Margin\\{0:0000}{1:00}{2:00}_tpex.json", dtStartDay.Year, dtStartDay.Month, dtStartDay.Day);
                if (File.Exists(strFileNameMarginTpex))
                {
                    string jsonTpex = File.ReadAllText(strFileNameMarginTpex);
                    Dictionary<string, object> jsonTpex_Dictionary = JsonConvert.DeserializeObject<Dictionary<string, object>>(jsonTpex);
                    if (jsonTpex_Dictionary.ContainsKey("aaData"))
                    {
                        List<List<string>> jsonaadata_ListList = JsonConvert.DeserializeObject<List<List<string>>>(jsonTpex_Dictionary["aaData"].ToString());
                        foreach (List<string> marginelist in jsonaadata_ListList)
                        {
                            string strno = marginelist.ElementAt(0);
                            string strFin = marginelist.ElementAt(6);
                            string strMar = marginelist.ElementAt(14);
                            string strFinUsingRate = marginelist.ElementAt(8);
                            string strMarUsingRate = marginelist.ElementAt(16);

                            int iFin = int.Parse(strFin, NumberStyles.AllowThousands);
                            int iMar = int.Parse(strMar, NumberStyles.AllowThousands);
                            double dbFinUsingRate = double.Parse(strFinUsingRate, NumberStyles.Float);
                            double dbMarUsingRate = double.Parse(strMarUsingRate, NumberStyles.Float);

                            if (!dicStockFinancing.ContainsKey(strno))
                            {
                                dicStockFinancing.Add(strno, new Dictionary<DateTime, int>());
                            }
                            dicStockFinancing[strno].Add(new DateTime(dtStartDay.Year, dtStartDay.Month, dtStartDay.Day), iFin);

                            if (!dicStockMarriage.ContainsKey(strno))
                            {
                                dicStockMarriage.Add(strno, new Dictionary<DateTime, int>());
                            }
                            dicStockMarriage[strno].Add(new DateTime(dtStartDay.Year, dtStartDay.Month, dtStartDay.Day), iMar);


                            if (!dicFinanceRate.ContainsKey(strno))
                            {
                                dicFinanceRate.Add(strno, new Dictionary<DateTime, double>());
                            }
                            dicFinanceRate[strno].Add(new DateTime(dtStartDay.Year, dtStartDay.Month, dtStartDay.Day), dbFinUsingRate);
                            if (!dicMarriageRate.ContainsKey(strno))
                            {
                                dicMarriageRate.Add(strno, new Dictionary<DateTime, double>());
                            }
                            dicMarriageRate[strno].Add(new DateTime(dtStartDay.Year, dtStartDay.Month, dtStartDay.Day), dbMarUsingRate);
                        }
                    }

                }

                
                dtStartDay = dtStartDay.AddDays(-1);
            }

            string jsonFinancing = JsonConvert.SerializeObject(dicStockFinancing);
            string jsonMarriage = JsonConvert.SerializeObject(dicStockMarriage);
            string jsonFinanceRate = JsonConvert.SerializeObject(dicFinanceRate);
            string jsonMarriageRate = JsonConvert.SerializeObject(dicMarriageRate);

            File.WriteAllText("Y:\\Fin" + DateTime.Now.ToString("MMdd") + ".json", jsonFinancing);
            File.WriteAllText("Y:\\Mar" + DateTime.Now.ToString("MMdd") + ".json", jsonMarriage);
            File.WriteAllText("Y:\\FinRate" + DateTime.Now.ToString("MMdd") + ".json", jsonFinanceRate);
            File.WriteAllText("Y:\\MarRate" + DateTime.Now.ToString("MMdd") + ".json", jsonMarriageRate);

            if(checkBoxAll.Checked)
            {
                button5_Click(null, null);
                button15_Click(null, null);
            }
        }   

        private void button5_Click(object sender, EventArgs e)
        {
            dicStock3Party.Clear();
            //dicStock3PartyLong.Clear();
            //test 3party
            DateTime dtStartDay = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);
            //DateTime dtEnd = dtStartDay.AddMonths(-6);
            DateTime dtEnd = dtStartDay.AddMonths(-4);// .AddYears(-1);
            DateTime dtEnd2 = dtEnd.AddMonths(-4);
            DateTime dtEnd3 = dtEnd.AddMonths(-4);
            for (int i = 0; dtEnd < dtStartDay; i++)
            {
                bool bisworkday = false;
                string strFileName3Party = string.Format("D:\\Chips\\3Party\\{0:0000}{1:00}{2:00}.json", dtStartDay.Year, dtStartDay.Month, dtStartDay.Day);
                if (File.Exists(strFileName3Party))
                {
                    string json = File.ReadAllText(strFileName3Party);
                    Dictionary<string, object> json_Dictionary = JsonConvert.DeserializeObject<Dictionary<string, object>>(json);
                    if (json_Dictionary.ContainsKey("data"))
                    {
                        List<List<string>> jsondata_ListList = JsonConvert.DeserializeObject<List<List<string>>>(json_Dictionary["data"].ToString());
                        foreach (List<string> marginelist in jsondata_ListList)
                        {
                            if(marginelist.ElementAt(0).Length == 4)
                            {
                                PartyParam Assign = new PartyParam();
                                Assign.Assign(marginelist);
                                if (Assign.name != null)
                                {
                                    //PartyParamLong AssignLong = new PartyParamLong();
                                    //AssignLong.Assign(marginelist);

                                    if (!dicStock3Party.ContainsKey(Assign.number))
                                    {
                                        dicStock3Party.Add(Assign.number, new Dictionary<DateTime, PartyParam>());
                                    }
                                    if (!dicStock3Party[Assign.number].ContainsKey(dtStartDay))
                                    {
                                        dicStock3Party[Assign.number].Add(dtStartDay, Assign);
                                    }

                                    /*if (!dicStock3PartyLong.ContainsKey(AssignLong.number))
                                    {
                                        dicStock3PartyLong.Add(AssignLong.number, new Dictionary<DateTime, PartyParamLong>());
                                    }
                                    if (!dicStock3PartyLong[AssignLong.number].ContainsKey(dtStartDay))
                                    {
                                        dicStock3PartyLong[AssignLong.number].Add(dtStartDay, AssignLong);
                                    }*/
                                }
                            }                            
                        }
                    }
                }
                string strFileName3PartyTpex = string.Format("D:\\Chips\\3Party\\{0:0000}{1:00}{2:00}_tpex.json", dtStartDay.Year, dtStartDay.Month, dtStartDay.Day);
                if (File.Exists(strFileName3PartyTpex))
                {
                    string json = File.ReadAllText(strFileName3PartyTpex);
                    Dictionary<string, object> json_Dictionary = JsonConvert.DeserializeObject<Dictionary<string, object>>(json);
                    if (json_Dictionary.ContainsKey("aaData"))
                    {
                        List<List<string>> jsondata_ListList = JsonConvert.DeserializeObject<List<List<string>>>(json_Dictionary["aaData"].ToString());
                        foreach (List<string> marginelist in jsondata_ListList)
                        {
                            if(marginelist.ElementAt(0).Length == 4)
                            {
                                PartyParam Assign = new PartyParam();
                                Assign.AssignTPEX(marginelist);

                                //PartyParamLong AssignLong = new PartyParamLong();
                                //AssignLong.AssignTPEX(marginelist);

                                if (!dicStock3Party.ContainsKey(Assign.number))
                                {
                                    dicStock3Party.Add(Assign.number, new Dictionary<DateTime, PartyParam>());
                                }
                                if (!dicStock3Party[Assign.number].ContainsKey(dtStartDay))
                                {
                                    dicStock3Party[Assign.number].Add(dtStartDay, Assign);
                                }


                                /*if (!dicStock3PartyLong.ContainsKey(AssignLong.number))
                                {
                                    dicStock3PartyLong.Add(AssignLong.number, new Dictionary<DateTime, PartyParamLong>());
                                }
                                if (!dicStock3PartyLong[AssignLong.number].ContainsKey(dtStartDay))
                                {
                                    dicStock3PartyLong[AssignLong.number].Add(dtStartDay, AssignLong);
                                }*/
                            }
                            
                        }
                    }
                }
                dtStartDay = dtStartDay.AddDays(-1);
            }
            string jsonStock3Party = JsonConvert.SerializeObject(dicStock3Party);
            File.WriteAllText("Y:\\3Party" + DateTime.Now.ToString("MMdd") + ".json", jsonStock3Party);

            //string jsonStock3PartyLong = JsonConvert.SerializeObject(dicStock3PartyLong);
            //File.WriteAllText("Y:\\3PartyLong" + DateTime.Now.ToString("MMdd") + ".json", jsonStock3PartyLong);

            Dictionary<string, Dictionary<DateTime, PartyParam>> dicStock3Party2 = new Dictionary<string, Dictionary<DateTime, PartyParam>>();
            for (int i = 0; dtEnd2 < dtStartDay; i++)
            {
                string strFileName3Party = string.Format("D:\\Chips\\3Party\\{0:0000}{1:00}{2:00}.json", dtStartDay.Year, dtStartDay.Month, dtStartDay.Day);
                if (File.Exists(strFileName3Party))
                {
                    string json = File.ReadAllText(strFileName3Party);
                    Dictionary<string, object> json_Dictionary = JsonConvert.DeserializeObject<Dictionary<string, object>>(json);
                    if (json_Dictionary.ContainsKey("data"))
                    {
                        List<List<string>> jsondata_ListList = JsonConvert.DeserializeObject<List<List<string>>>(json_Dictionary["data"].ToString());
                        foreach (List<string> marginelist in jsondata_ListList)
                        {
                            if (marginelist.ElementAt(0).Length == 4)
                            {
                                PartyParam Assign = new PartyParam();
                                Assign.Assign(marginelist);
                                if (marginelist.ElementAt(0) == "1701" && dtStartDay == new DateTime(2020, 6, 4))
                                {
                                    int sdvm = 0;
                                }
                                else
                                {
                                    if (!dicStock3Party2.ContainsKey(Assign.number))
                                    {
                                        dicStock3Party2.Add(Assign.number, new Dictionary<DateTime, PartyParam>());
                                    }
                                    if (!dicStock3Party2[Assign.number].ContainsKey(dtStartDay))
                                    {
                                        dicStock3Party2[Assign.number].Add(dtStartDay, Assign);
                                    }
                                }


                            }
                        }
                    }
                }
                string strFileName3PartyTpex = string.Format("D:\\Chips\\3Party\\{0:0000}{1:00}{2:00}_tpex.json", dtStartDay.Year, dtStartDay.Month, dtStartDay.Day);
                if (File.Exists(strFileName3PartyTpex))
                {
                    string json = File.ReadAllText(strFileName3PartyTpex);
                    Dictionary<string, object> json_Dictionary = JsonConvert.DeserializeObject<Dictionary<string, object>>(json);
                    if (json_Dictionary.ContainsKey("aaData"))
                    {
                        List<List<string>> jsondata_ListList = JsonConvert.DeserializeObject<List<List<string>>>(json_Dictionary["aaData"].ToString());
                        foreach (List<string> marginelist in jsondata_ListList)
                        {
                            if (marginelist.ElementAt(0).Length == 4)
                            {
                                PartyParam Assign = new PartyParam();
                                Assign.AssignTPEX(marginelist);


                                if (!dicStock3Party2.ContainsKey(Assign.number))
                                {
                                    dicStock3Party2.Add(Assign.number, new Dictionary<DateTime, PartyParam>());
                                }
                                if (!dicStock3Party2[Assign.number].ContainsKey(dtStartDay))
                                {
                                    dicStock3Party2[Assign.number].Add(dtStartDay, Assign);
                                }
                            }

                        }
                    }
                }
                dtStartDay = dtStartDay.AddDays(-1);
            }
            string jsonStock3Party2 = JsonConvert.SerializeObject(dicStock3Party2);
            File.WriteAllText("Y:\\3Party2" + DateTime.Now.ToString("MMdd") + ".json", jsonStock3Party2);

        }

        private void button4_Click(object sender, EventArgs e)
        {
            //return;
            DateTime dtStartDay = DateTime.Now;
            string strFile = "Y:\\Fin" + dtStartDay.ToString("MMdd") + ".json";
            if(File.Exists(strFile))
            {
                string jsonFinancing = File.ReadAllText("Y:\\Fin" + dtStartDay.ToString("MMdd") + ".json");
                dicStockFinancing = JsonConvert.DeserializeObject<Dictionary<string, Dictionary<DateTime, int>>>(jsonFinancing);

                string jsonMarriage = File.ReadAllText("Y:\\Mar" + dtStartDay.ToString("MMdd") + ".json");
                dicStockMarriage = JsonConvert.DeserializeObject<Dictionary<string, Dictionary<DateTime, int>>>(jsonMarriage);

                string jsonFinanceRate = File.ReadAllText("Y:\\FinRate" + dtStartDay.ToString("MMdd") + ".json");
                dicFinanceRate = JsonConvert.DeserializeObject<Dictionary<string, Dictionary<DateTime, double>>>(jsonFinanceRate);

                string jsonMarriageRate = File.ReadAllText("Y:\\MarRate" + dtStartDay.ToString("MMdd") + ".json");
                dicMarriageRate = JsonConvert.DeserializeObject<Dictionary<string, Dictionary<DateTime, double>>>(jsonMarriageRate);       

                string jsonWorking = File.ReadAllText("Y:\\Working" + dtStartDay.ToString("MMdd") + ".json");
                workingdays = JsonConvert.DeserializeObject<HashSet<DateTime>>(jsonWorking);

                string json3Party = File.ReadAllText("Y:\\3Party" + dtStartDay.ToString("MMdd") + ".json");
                dicStock3Party = JsonConvert.DeserializeObject<Dictionary<string, Dictionary<DateTime, PartyParam>>>(json3Party);

                string json3Party2 = File.ReadAllText("Y:\\3Party2" + dtStartDay.ToString("MMdd") + ".json");
                dicStock3Party2 = JsonConvert.DeserializeObject<Dictionary<string, Dictionary<DateTime, PartyParam>>>(json3Party2);

                string jsonOLHC = File.ReadAllText("Y:\\OLHC" + dtStartDay.ToString("MMdd") + ".json");
                dbEvenydayOLHCDic = JsonConvert.DeserializeObject<Dictionary<string, Dictionary<DateTime, double[]>>>(jsonOLHC);

                string jsonVOL = File.ReadAllText("Y:\\VOL" + dtStartDay.ToString("MMdd") + ".json");
                dbStockVolDic = JsonConvert.DeserializeObject<Dictionary<string, Dictionary<DateTime, long>>>(jsonVOL);
            }

           /* string json3PartyLong = "Y:\\3PartyLong" + dtStartDay.ToString("MMdd") + ".json";
            if (File.Exists(json3PartyLong))
            {
                byte[] docBytes = File.ReadAllBytes(json3PartyLong);

                string jsonString = Encoding.UTF32.GetString(docBytes);

                dicStock3PartyLong = JsonConvert.DeserializeObject<Dictionary<string, Dictionary<DateTime, PartyParamLong>>>(jsonString);
            }*/

            
            foreach(string strno in dicStock3Party.Keys)
            {
                dicStockNoMapping.Add(strno, dicStock3Party[strno].ElementAt(0).Value.name.Trim());
                if (dicStock3Party2.ContainsKey(strno))
                {
                    foreach (var item in dicStock3Party2[strno])
                    {
                        dicStock3Party[strno].Add(item.Key, item.Value);
                    }
                   
                }
            }
            dicStockNoMapping.Add("006205", "富邦上証");
        }

        private void button_GetPrice_Click(object sender, EventArgs e)
        {
            DateTime dt = DateTime.Now;
            // dt = new DateTime(2018, 8, 27);
            string strSavePath = "D:\\Chips\\eachdaystock\\";
            for (int i = 0; i < 90; i++)
            {
                StreamWriter myWriter;

                if (!checkBoxToday.Checked)
                    dt = dt.AddDays(-1);
                string strStockMon = string.Format("{0}.json", dt.ToString("yyyyMMdd"));
                string strfile = strSavePath + strStockMon;
                if (!File.Exists(strfile))
                {

                    string url = string.Format("https://www.twse.com.tw/exchangeReport/MI_INDEX?response=json&date={0}{1:00}{2:00}&type=ALLBUT0999", dt.Year, dt.Month, dt.Day);

                    HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
                    httpWebRequest.ContentType = "application/json";
                    httpWebRequest.Method = WebRequestMethods.Http.Get;
                    httpWebRequest.Accept = "application/json";

                    string text;
                    var response = (HttpWebResponse)httpWebRequest.GetResponse();

                    using (var sr = new StreamReader(response.GetResponseStream()))
                    {
                        text = sr.ReadToEnd();
                    }

                    File.WriteAllText(strSavePath + strStockMon, text);

                    response.Close();
                    httpWebRequest.Abort();
                    Thread.Sleep(1000);
                }

                string strStockMontpex = string.Format("{0}_tpex.json", dt.ToString("yyyyMMdd"));
                string strfiletpex = strSavePath + strStockMontpex;
                if (!File.Exists(strfiletpex))
                {
                    string url = string.Format("https://www.tpex.org.tw/web/stock/aftertrading/otc_quotes_no1430/stk_wn1430_result.php?l=zh-tw&se=EW&o=JSON&d={0}/{1:00}/{2:00}", dt.Year - 1911, dt.Month, dt.Day);

                    HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
                    httpWebRequest.ContentType = "application/json";
                    httpWebRequest.Method = WebRequestMethods.Http.Get;
                    httpWebRequest.Accept = "application/json";
                    string text;
                    var response = (HttpWebResponse)httpWebRequest.GetResponse();

                    using (var sr = new StreamReader(response.GetResponseStream()))
                    {
                        text = sr.ReadToEnd();
                    }

                    File.WriteAllText(strfiletpex, text);

                    response.Close();
                    httpWebRequest.Abort();
                    Thread.Sleep(1000);
                }
            }
        }

        private void button_GetMargin_Click(object sender, EventArgs e)
        {
              https://www.twse.com.tw/exchangeReport/MI_MARGN?
            //https://www.twse.com.tw/exchangeReport/MI_MARGN?response=json&date=20180328&selectType=ALL
            //https://www.tpex.org.tw/web/stock/margin_trading/margin_balance/margin_bal_result.php?l=zh-tw&o=JSON&d=108/10/01
            string strSavePath = "D:\\Chips\\Margin\\";
            DateTime dt = DateTime.Now;


            for (int i = 0; i < 90; i++)
            {
                if(!checkBoxToday.Checked)
                    dt = dt.AddDays(-1);
                string strStockMon = string.Format("{0:0000}{1:00}{2:00}.json", dt.Year, dt.Month, dt.Day);
                string strLocalFile = strSavePath + strStockMon;
                if (!File.Exists(strLocalFile))
                {

                    string url = string.Format("https://www.twse.com.tw/exchangeReport/MI_MARGN?response=json&date={0:0000}{1:00}{2:00}&selectType=ALL", dt.Year, dt.Month, dt.Day);

                    HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
                    httpWebRequest.ContentType = "application/json";
                    httpWebRequest.Method = WebRequestMethods.Http.Get;
                    httpWebRequest.Accept = "application/json";

                    string text;
                    var response = (HttpWebResponse)httpWebRequest.GetResponse();

                    using (var sr = new StreamReader(response.GetResponseStream()))
                    {
                        text = sr.ReadToEnd();
                    }

                    File.WriteAllText(strLocalFile, text);

                    response.Close();
                    httpWebRequest.Abort();
                }

                string strTpexStockMon = string.Format("{0:0000}{1:00}{2:00}_tpex.json", dt.Year, dt.Month, dt.Day);
                string strTpexLocalFile = strSavePath + strTpexStockMon;
                if (!File.Exists(strTpexLocalFile))
                {
                    string url = string.Format("https://www.tpex.org.tw/web/stock/margin_trading/margin_balance/margin_bal_result.php?l=zh-tw&o=JSON&d={0:000}/{1:00}/{2:00}", dt.Year - 1911, dt.Month, dt.Day);

                    HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
                    httpWebRequest.ContentType = "application/json";
                    httpWebRequest.Method = WebRequestMethods.Http.Get;
                    httpWebRequest.Accept = "application/json";
                    string text;
                    var response = (HttpWebResponse)httpWebRequest.GetResponse();

                    using (var sr = new StreamReader(response.GetResponseStream()))
                    {
                        text = sr.ReadToEnd();
                    }

                    File.WriteAllText(strTpexLocalFile, text);

                    response.Close();
                    httpWebRequest.Abort();
                }
            }
        }

        private void button_Get3Party_Click(object sender, EventArgs e)
        {
            //https://www.twse.com.tw/fund/T86?response=JSON&date={0}{1:00}{2:00}&selectType=ALL
            //https://www.tpex.org.tw/web/stock/margin_trading/margin_balance/margin_bal_result.php?l=zh-tw&o=JSON&d=108/10/01
            string strSavePath = "D:\\Chips\\3Party\\";
            DateTime dt = DateTime.Now;
    

            for (int i = 0; i < 120; i++)
            {
                if (!checkBoxToday.Checked)
                    dt = dt.AddDays(-1);
                string strStockMon = string.Format("{0:0000}{1:00}{2:00}.json", dt.Year, dt.Month, dt.Day);
                string strLocalFile = strSavePath + strStockMon;
                if (!File.Exists(strLocalFile))
                {

                    string url = string.Format("https://www.twse.com.tw/fund/T86?response=JSON&date={0}{1:00}{2:00}&selectType=ALL", dt.Year, dt.Month, dt.Day);

                    HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
                    httpWebRequest.ContentType = "application/json";
                    httpWebRequest.Method = WebRequestMethods.Http.Get;
                    httpWebRequest.Accept = "application/json";
                    string text;
                    var response = (HttpWebResponse)httpWebRequest.GetResponse();

                    using (var sr = new StreamReader(response.GetResponseStream()))
                    {
                        text = sr.ReadToEnd();
                    }

                    File.WriteAllText(strLocalFile, text);

                    response.Close();
                    httpWebRequest.Abort();
                }

                string strTpexStockMon = string.Format("{0:0000}{1:00}{2:00}_tpex.json", dt.Year, dt.Month, dt.Day);
                string strTpexLocalFile = strSavePath + strTpexStockMon;
                if (!File.Exists(strTpexLocalFile))
                {
                    string url = string.Format("https://www.tpex.org.tw/web/stock/3insti/DAILY_TradE/3itrade_hedge_result.php?l=zh-tw&se=EW&t=D&d={0:000}/{1:00}/{2:00}", dt.Year - 1911, dt.Month, dt.Day);

                    HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
                    httpWebRequest.ContentType = "application/json";
                    httpWebRequest.Method = WebRequestMethods.Http.Get;
                    httpWebRequest.Accept = "application/json";
                    string text;
                    var response = (HttpWebResponse)httpWebRequest.GetResponse();

                    using (var sr = new StreamReader(response.GetResponseStream()))
                    {
                        text = sr.ReadToEnd();
                    }

                    File.WriteAllText(strTpexLocalFile, text);

                    response.Close();
                    httpWebRequest.Abort();
                }
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            // 外資 & 投信同買
            DateTime dtEnd = new DateTime(2018, 10, 20);
            int iUseCount = 120;
            DateTime dtStart = monthCalendar1.SelectionStart;
            DateTime dtLoopStart =dtStart;
            DateTime dtYesterday = dtStart.AddDays(-1);
            while (!workingdays.Contains(dtYesterday))
            {
                dtYesterday = dtYesterday.AddDays(-1);
            }
            CandidateParty3Stock.Clear();

            Dictionary<string, double[]> stock3PartyTotalDo = new Dictionary<string, double[]>();
            Dictionary<string, double> stock3PartyAvg = new Dictionary<string, double>();
            Dictionary<string, double> stock3PartyStd = new Dictionary<string, double>();

            foreach(string strStockNo in dicStock3Party.Keys)
            {
                if (!stock3PartyTotalDo.ContainsKey(strStockNo))
                {
                    stock3PartyTotalDo.Add(strStockNo, new double[300]);
                }
                int iLoopDay = 0;
                dtLoopStart = dtStart;
                while (dtLoopStart > dtEnd)
                {
                    if (dicStock3Party[strStockNo].ContainsKey(dtLoopStart))
                    {
                        int iTotal = int.Parse(dicStock3Party[strStockNo][dtLoopStart].Party3Total, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                        int iHedge = int.Parse(dicStock3Party[strStockNo][dtLoopStart].SelfHedgeTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                        stock3PartyTotalDo[strStockNo][iLoopDay] = iTotal - iHedge;
                        iLoopDay++;
                    }
                    else{
                        if(workingdays.Contains(dtLoopStart))
                        {
                            stock3PartyTotalDo[strStockNo][iLoopDay] = 0;
                            iLoopDay++;
                        }
                    }

                    dtLoopStart = dtLoopStart.AddDays(-1);
                }

                double dbavg=0;
                double dbstd = StandardDeviation(stock3PartyTotalDo[strStockNo], out dbavg);
                stock3PartyAvg.Add(strStockNo, dbavg);
                stock3PartyStd.Add(strStockNo, dbstd);
            }

            foreach(string strNo in stock3PartyAvg.Keys)
            {
                //if (stock3PartyTotalDo[strNo][0] > stock3PartyAvg[strNo] + (stock3PartyStd[strNo] * 3))
                
                if (stock3PartyTotalDo[strNo][0] > stock3PartyAvg[strNo] + (stock3PartyStd[strNo] * 3))
                {
                    CandidateParty3Stock.Add(strNo);
                    Console.WriteLine(strNo);
                }
            }
            //dicStock3Party
        }

        private void button7_Click(object sender, EventArgs e)
        {
            DateTime dtEnd = new DateTime(2018, 10, 20);
            DateTime dtStart = monthCalendar1.SelectionStart;
            DateTime dtLoopStart = monthCalendar1.SelectionStart;
            //CandidateWeekChipSet.Clear();

            Queue<string> qcsvfile = new Queue<string>();
            string strBasePath = "D:\\Chips\\stock\\TDCC_OD_1-5_";

            while (dtLoopStart > dtEnd)
            {
                string strcsv = strBasePath + dtLoopStart.ToString("yyyyMMdd") + ".csv";
                if (File.Exists(strcsv))
                {
                    qcsvfile.Enqueue(strcsv);
                }
                dtLoopStart = dtLoopStart.AddDays(-1);
            }

            int iCount = 16;
            Dictionary<string, StockLevelCount>[] weekLevelCount = new Dictionary<string, StockLevelCount>[iCount];
            for (int i = 0; i < iCount; i++)
            {
                string strLat = qcsvfile.ElementAt(i);
                weekLevelCount[i] = new Dictionary<string, StockLevelCount>();

                int iReadCount = 0;
                IEnumerable<string> lines = File.ReadLines(strLat);
                foreach (string line in lines)
                {
                    if (iReadCount > 0)
                    {
                        string[] strSplitLine = line.Split(',');
                        if (weekLevelCount[i].ContainsKey(strSplitLine[1]))
                        {
                            int iLevel = int.Parse(strSplitLine[2]) - 1;
                            weekLevelCount[i][strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                            weekLevelCount[i][strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                            weekLevelCount[i][strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                        }
                        else
                        {
                            weekLevelCount[i].Add(strSplitLine[1], new StockLevelCount());
                            int iLevel = int.Parse(strSplitLine[2]) - 1;
                            weekLevelCount[i][strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                            weekLevelCount[i][strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                            weekLevelCount[i][strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                        }
                    }
                    iReadCount++;
                }
            }
            //Console.WriteLine("CandidateWeekChipSet");
            foreach (KeyValuePair<string, StockLevelCount> item in weekLevelCount[0])
            {
                if (weekLevelCount[1].ContainsKey(item.Key) && weekLevelCount[2].ContainsKey(item.Key) && weekLevelCount[3].ContainsKey(item.Key) && weekLevelCount[4].ContainsKey(item.Key) && weekLevelCount[5].ContainsKey(item.Key) && weekLevelCount[6].ContainsKey(item.Key) && weekLevelCount[7].ContainsKey(item.Key) &&
                    weekLevelCount[8].ContainsKey(item.Key) && weekLevelCount[9].ContainsKey(item.Key) && weekLevelCount[10].ContainsKey(item.Key) && weekLevelCount[11].ContainsKey(item.Key) && weekLevelCount[12].ContainsKey(item.Key) && weekLevelCount[13].ContainsKey(item.Key) && weekLevelCount[14].ContainsKey(item.Key) && weekLevelCount[15].ContainsKey(item.Key))
                {
                    // LevelRate[15] 是差異數調整，[14]:1,000,001以上，[13]:800,001-1,000,000，[12]:600,001-800,000，[11]:400,001-600,000，[10]200,001-400,000   ，[9]:100,001-200,000，[8]:50,001-100,000
                    double dbrate0 = weekLevelCount[0][item.Key].LevelRate[10] + weekLevelCount[0][item.Key].LevelRate[11] + weekLevelCount[0][item.Key].LevelRate[12] + weekLevelCount[0][item.Key].LevelRate[13] + weekLevelCount[0][item.Key].LevelRate[14];// +weekLevelCount[0][item.Key].LevelRate[15];
                    double dbrate1 = weekLevelCount[1][item.Key].LevelRate[10] + weekLevelCount[1][item.Key].LevelRate[11] + weekLevelCount[1][item.Key].LevelRate[12] + weekLevelCount[1][item.Key].LevelRate[13] + weekLevelCount[1][item.Key].LevelRate[14];// + weekLevelCount[1][item.Key].LevelRate[15];
                    double dbrate2 = weekLevelCount[2][item.Key].LevelRate[10] + weekLevelCount[2][item.Key].LevelRate[11] + weekLevelCount[2][item.Key].LevelRate[12] + weekLevelCount[2][item.Key].LevelRate[13] + weekLevelCount[2][item.Key].LevelRate[14];// + weekLevelCount[2][item.Key].LevelRate[15];
                    double dbrate3 = weekLevelCount[3][item.Key].LevelRate[10] + weekLevelCount[3][item.Key].LevelRate[11] + weekLevelCount[3][item.Key].LevelRate[12] + weekLevelCount[3][item.Key].LevelRate[13] + weekLevelCount[3][item.Key].LevelRate[14];// + weekLevelCount[3][item.Key].LevelRate[15];
                    double dbrate4 = weekLevelCount[4][item.Key].LevelRate[10] + weekLevelCount[4][item.Key].LevelRate[11] + weekLevelCount[4][item.Key].LevelRate[12] + weekLevelCount[4][item.Key].LevelRate[13] + weekLevelCount[4][item.Key].LevelRate[14];// + weekLevelCount[4][item.Key].LevelRate[15];
                    double dbrate5 = weekLevelCount[5][item.Key].LevelRate[10] + weekLevelCount[5][item.Key].LevelRate[11] + weekLevelCount[5][item.Key].LevelRate[12] + weekLevelCount[5][item.Key].LevelRate[13] + weekLevelCount[5][item.Key].LevelRate[14];// + weekLevelCount[5][item.Key].LevelRate[15];
                    double dbrate6 = weekLevelCount[6][item.Key].LevelRate[10] + weekLevelCount[6][item.Key].LevelRate[11] + weekLevelCount[6][item.Key].LevelRate[12] + weekLevelCount[6][item.Key].LevelRate[13] + weekLevelCount[6][item.Key].LevelRate[14];// + weekLevelCount[6][item.Key].LevelRate[15];
                    double dbrate7 = weekLevelCount[7][item.Key].LevelRate[10] + weekLevelCount[7][item.Key].LevelRate[11] + weekLevelCount[7][item.Key].LevelRate[12] + weekLevelCount[7][item.Key].LevelRate[13] + weekLevelCount[7][item.Key].LevelRate[14];// + weekLevelCount[7][item.Key].LevelRate[15];
                    double dbrate08 = weekLevelCount[08][item.Key].LevelRate[10] + weekLevelCount[08][item.Key].LevelRate[11] + weekLevelCount[08][item.Key].LevelRate[12] + weekLevelCount[08][item.Key].LevelRate[13] + weekLevelCount[08][item.Key].LevelRate[14];// + weekLevelCount[08][item.Key].LevelRate[15];
                    double dbrate09 = weekLevelCount[09][item.Key].LevelRate[10] + weekLevelCount[09][item.Key].LevelRate[11] + weekLevelCount[09][item.Key].LevelRate[12] + weekLevelCount[09][item.Key].LevelRate[13] + weekLevelCount[09][item.Key].LevelRate[14];// + weekLevelCount[09][item.Key].LevelRate[15];
                    double dbrate10 = weekLevelCount[10][item.Key].LevelRate[10] + weekLevelCount[10][item.Key].LevelRate[11] + weekLevelCount[10][item.Key].LevelRate[12] + weekLevelCount[10][item.Key].LevelRate[13] + weekLevelCount[10][item.Key].LevelRate[14];// + weekLevelCount[10][item.Key].LevelRate[15];
                    double dbrate11 = weekLevelCount[11][item.Key].LevelRate[10] + weekLevelCount[11][item.Key].LevelRate[11] + weekLevelCount[11][item.Key].LevelRate[12] + weekLevelCount[11][item.Key].LevelRate[13] + weekLevelCount[11][item.Key].LevelRate[14];// + weekLevelCount[11][item.Key].LevelRate[15];
                    double dbrate12 = weekLevelCount[12][item.Key].LevelRate[10] + weekLevelCount[12][item.Key].LevelRate[11] + weekLevelCount[12][item.Key].LevelRate[12] + weekLevelCount[12][item.Key].LevelRate[13] + weekLevelCount[12][item.Key].LevelRate[14];// + weekLevelCount[12][item.Key].LevelRate[15];
                    double dbrate13 = weekLevelCount[13][item.Key].LevelRate[10] + weekLevelCount[13][item.Key].LevelRate[11] + weekLevelCount[13][item.Key].LevelRate[12] + weekLevelCount[13][item.Key].LevelRate[13] + weekLevelCount[13][item.Key].LevelRate[14];// + weekLevelCount[13][item.Key].LevelRate[15];
                    double dbrate14 = weekLevelCount[14][item.Key].LevelRate[10] + weekLevelCount[14][item.Key].LevelRate[11] + weekLevelCount[14][item.Key].LevelRate[12] + weekLevelCount[14][item.Key].LevelRate[13] + weekLevelCount[14][item.Key].LevelRate[14];// + weekLevelCount[14][item.Key].LevelRate[15];
                    double dbrate15 = weekLevelCount[15][item.Key].LevelRate[10] + weekLevelCount[15][item.Key].LevelRate[11] + weekLevelCount[15][item.Key].LevelRate[12] + weekLevelCount[15][item.Key].LevelRate[13] + weekLevelCount[15][item.Key].LevelRate[14];//+ weekLevelCount[15][item.Key].LevelRate[15];

                    // LevelPeople[16] 總人數 - [0]1-999股 - [1]:1-5張 -[2]:5-10張
                    double dbpeople0 = weekLevelCount[0][item.Key].LevelPeople[16] - weekLevelCount[0][item.Key].LevelPeople[0] - weekLevelCount[0][item.Key].LevelPeople[1] - weekLevelCount[0][item.Key].LevelPeople[2];
                    double dbpeople1 = weekLevelCount[1][item.Key].LevelPeople[16] - weekLevelCount[1][item.Key].LevelPeople[0] - weekLevelCount[1][item.Key].LevelPeople[1] - weekLevelCount[1][item.Key].LevelPeople[2];
                    double dbpeople2 = weekLevelCount[2][item.Key].LevelPeople[16] - weekLevelCount[2][item.Key].LevelPeople[0] - weekLevelCount[2][item.Key].LevelPeople[1] - weekLevelCount[2][item.Key].LevelPeople[2];
                    double dbpeople3 = weekLevelCount[3][item.Key].LevelPeople[16] - weekLevelCount[3][item.Key].LevelPeople[0] - weekLevelCount[3][item.Key].LevelPeople[1] - weekLevelCount[3][item.Key].LevelPeople[2];
                    double dbpeople4 = weekLevelCount[4][item.Key].LevelPeople[16] - weekLevelCount[4][item.Key].LevelPeople[0] - weekLevelCount[4][item.Key].LevelPeople[1] - weekLevelCount[4][item.Key].LevelPeople[2];
                    double dbpeople5 = weekLevelCount[5][item.Key].LevelPeople[16] - weekLevelCount[5][item.Key].LevelPeople[0] - weekLevelCount[5][item.Key].LevelPeople[1] - weekLevelCount[5][item.Key].LevelPeople[2];
                    double dbpeople6 = weekLevelCount[6][item.Key].LevelPeople[16] - weekLevelCount[6][item.Key].LevelPeople[0] - weekLevelCount[6][item.Key].LevelPeople[1] - weekLevelCount[6][item.Key].LevelPeople[2];
                    double dbpeople7 = weekLevelCount[7][item.Key].LevelPeople[16] - weekLevelCount[7][item.Key].LevelPeople[0] - weekLevelCount[7][item.Key].LevelPeople[1] - weekLevelCount[7][item.Key].LevelPeople[2];
                    double dbpeople08 = weekLevelCount[08][item.Key].LevelPeople[16] - weekLevelCount[08][item.Key].LevelPeople[0] - weekLevelCount[08][item.Key].LevelPeople[1] - weekLevelCount[08][item.Key].LevelPeople[2];
                    double dbpeople09 = weekLevelCount[09][item.Key].LevelPeople[16] - weekLevelCount[09][item.Key].LevelPeople[0] - weekLevelCount[09][item.Key].LevelPeople[1] - weekLevelCount[09][item.Key].LevelPeople[2];
                    double dbpeople10 = weekLevelCount[10][item.Key].LevelPeople[16] - weekLevelCount[10][item.Key].LevelPeople[0] - weekLevelCount[10][item.Key].LevelPeople[1] - weekLevelCount[10][item.Key].LevelPeople[2];
                    double dbpeople11 = weekLevelCount[11][item.Key].LevelPeople[16] - weekLevelCount[11][item.Key].LevelPeople[0] - weekLevelCount[11][item.Key].LevelPeople[1] - weekLevelCount[11][item.Key].LevelPeople[2];
                    double dbpeople12 = weekLevelCount[12][item.Key].LevelPeople[16] - weekLevelCount[12][item.Key].LevelPeople[0] - weekLevelCount[12][item.Key].LevelPeople[1] - weekLevelCount[12][item.Key].LevelPeople[2];
                    double dbpeople13 = weekLevelCount[13][item.Key].LevelPeople[16] - weekLevelCount[13][item.Key].LevelPeople[0] - weekLevelCount[13][item.Key].LevelPeople[1] - weekLevelCount[13][item.Key].LevelPeople[2];
                    double dbpeople14 = weekLevelCount[14][item.Key].LevelPeople[16] - weekLevelCount[14][item.Key].LevelPeople[0] - weekLevelCount[14][item.Key].LevelPeople[1] - weekLevelCount[14][item.Key].LevelPeople[2];
                    double dbpeople15 = weekLevelCount[15][item.Key].LevelPeople[16] - weekLevelCount[15][item.Key].LevelPeople[0] - weekLevelCount[15][item.Key].LevelPeople[1] - weekLevelCount[15][item.Key].LevelPeople[2];

                    //

                    //long lpeople0 = weekLevelCount[0][item.Key].LevelPeople[16] - weekLevelCount[0][item.Key].LevelPeople[0];
                    //long lpeople7 = weekLevelCount[7][item.Key].LevelPeople[16] - weekLevelCount[7][item.Key].LevelPeople[0];
                    //抓上週大戶翻多　人數變少
                    //if (dbrate0 > dbrate1 && dbrate1 <= dbrate2 && dbrate2 <= dbrate3 && dbrate3 <= dbrate4 && 
                    //dbpeople0 < dbpeople1 && dbpeople1 >= dbpeople2 && dbpeople2 >= dbpeople3 && dbpeople3 >= dbpeople4 )
                    //if (dbrate0 > dbrate1 && dbrate1 <= dbrate2 && dbrate2 <= dbrate3 && dbrate3 <= dbrate4 && 
                    //    dbpeople0 < dbpeople1 && dbpeople1 >= dbpeople2 && dbpeople2 >= dbpeople3 && dbpeople3 >= dbpeople4 )
                    // if (dbrate0 > dbrate1 && dbrate1 > dbrate2 && dbrate2 > dbrate3)
                    //if (dbrate0 > dbrate1 && dbrate1 > dbrate2 && dbrate2 > dbrate3 && dbpeople0 < dbpeople1 && dbpeople1 <= dbpeople2 && dbpeople2 <= dbpeople3 && dbpeople3 <= dbpeople4)// 34/93
              
                    if (dbrate0 - dbrate3 > 3)
                    {
                        //string strf = string.Format("{0},{1:0.##},{2:0.##},{3:0.##},{4:0.##},{5:0.##},{6:0.##},{6:0.##},{7:0.##},{8:0.##},{9},{10}", item.Key, dbrate0, dbrate1, dbrate2, dbrate3, dbrate4, dbrate5, dbrate6, dbrate7, lpeople0, lpeople7);
                        string strf = string.Format("{0},{1:0.##},{2:0.##},{3:0.##},{4:0.##},{5:0.##},{6:0.##},{6:0.##},{7:0.##},{8:0.##},{9:0.##},{10:0.##},{11:0.##},{12:0.##},{13:0.##},{14:0.##},{14:0.##},{15:0.##}", item.Key, dbrate0, dbrate1, dbrate2, dbrate3, dbrate4, dbrate5, dbrate6, dbrate7, dbrate08, dbrate09, dbrate10, dbrate11, dbrate12, dbrate13, dbrate14, dbrate15);
                        if (strf == "")
                        {
                            int mmm = 0;
                        }
                        //CandidateWeekChipSet.Add(item.Key);
                        Console.WriteLine(strf);
                    }
                }
            }


        }

        private void button2_Click(object sender, EventArgs e)
        {
            int iCheckDays = 20;
            int iCountDataDays = 0;
            DateTime dtStart = monthCalendar1.SelectionStart;

            Dictionary<string, List<double>> dbStockCloseDic = new Dictionary<string, List<double>>();
            Dictionary<string, List<double>> dbStockOpenDic = new Dictionary<string, List<double>>();
            Dictionary<string, List<double>> dbStockHighDic = new Dictionary<string, List<double>>();
            Dictionary<string, List<double>> dbStockLowDic = new Dictionary<string, List<double>>();
            

            while (iCountDataDays < iCheckDays)
            {
                dtStart = dtStart.AddDays(1);

                string jsonFilePath = string.Format("D:\\Chips\\eachdaystock\\{0}.json", dtStart.ToString("yyyyMMdd"));
                string jsonFilePathTpex = string.Format("D:\\Chips\\eachdaystock\\{0}_tpex.json", dtStart.ToString("yyyyMMdd"));

                string json = File.ReadAllText(jsonFilePath);
                Dictionary<string, object> json_Dictionary = JsonConvert.DeserializeObject<Dictionary<string, object>>(json);
                if (json_Dictionary.Count > 2)
                {
                    iCountDataDays++;

                    List<List<string>> json_data9 = JsonConvert.DeserializeObject<List<List<string>>>(json_Dictionary["data9"].ToString());
                    foreach (List<string> pricelst in json_data9)
                    {
                        string strstockno = pricelst.ElementAt(0);
                        if (!dbStockCloseDic.ContainsKey(strstockno))
                        {
                            dbStockCloseDic.Add(strstockno, new List<double>());
                            dbStockOpenDic.Add(strstockno, new List<double>());
                            dbStockHighDic.Add(strstockno, new List<double>());
                            dbStockLowDic.Add(strstockno, new List<double>());
                        }

                        string strcloseprice = pricelst.ElementAt(8);
                        string stropenprice = pricelst.ElementAt(5);
                        string strhighprice = pricelst.ElementAt(6);
                        string strlowprice = pricelst.ElementAt(7);
                        if (strcloseprice != "--")
                        {
                            double dbopen = double.Parse(stropenprice);
                            double dbclose = double.Parse(strcloseprice);
                            double dbhigh = double.Parse(strhighprice);
                            double dblow = double.Parse(strlowprice);
                            dbStockCloseDic[strstockno].Add(dbclose);
                            dbStockOpenDic[strstockno].Add(dbopen);
                            dbStockHighDic[strstockno].Add(dbhigh);
                            dbStockLowDic[strstockno].Add(dblow);
                        }
                        else
                        {
                            dbStockCloseDic[strstockno].Add(0);
                            dbStockOpenDic[strstockno].Add(0);
                            dbStockHighDic[strstockno].Add(0);
                            dbStockLowDic[strstockno].Add(0);
                        }
                    }

                    string jsonTpex = File.ReadAllText(jsonFilePathTpex);
                    Dictionary<string, object> jsonTpex_Dictionary = JsonConvert.DeserializeObject<Dictionary<string, object>>(jsonTpex);
                    List<List<string>> json_aaData = JsonConvert.DeserializeObject<List<List<string>>>(jsonTpex_Dictionary["aaData"].ToString());
                    foreach (List<string> pricelst in json_aaData)
                    {
                        string strstockno = pricelst.ElementAt(0);
                        string strcloseprice = pricelst.ElementAt(2);
                        string strhighprice = pricelst.ElementAt(5);
                        string strlowprice = pricelst.ElementAt(6);
                        string stropenprice = pricelst.ElementAt(4);

                        if (!dbStockCloseDic.ContainsKey(strstockno))
                        {
                            dbStockCloseDic.Add(strstockno, new List<double>());
                            dbStockOpenDic.Add(strstockno, new List<double>());
                            dbStockHighDic.Add(strstockno, new List<double>());
                            dbStockLowDic.Add(strstockno, new List<double>());
                        }

                        if (strcloseprice != "---" && strcloseprice != "----")
                        {
                            double dbclose = double.Parse(strcloseprice);
                            double dbhigh = double.Parse(strhighprice);
                            double dblow = double.Parse(strlowprice);
                            double dbopen = double.Parse(stropenprice);
                            dbStockCloseDic[strstockno].Add(dbclose);
                            dbStockOpenDic[strstockno].Add(dbopen);
                            dbStockHighDic[strstockno].Add(dbhigh);
                            dbStockLowDic[strstockno].Add(dblow);
                        }
                        else
                        {
                            dbStockCloseDic[strstockno].Add(0);
                            dbStockOpenDic[strstockno].Add(0);
                            dbStockHighDic[strstockno].Add(0);
                            dbStockLowDic[strstockno].Add(0);
                        }
                    }
                }
            }

            foreach (string stno in CandidateParty3Stock)
            {
                //if (dbStockCloseDic.ContainsKey(stno) && dbStockHighDic.ContainsKey(stno) && CandidateFinMarStock.Contains(stno) && CandidateParty3Stock.Contains(stno))
                //if (CandidateWeekChipSet.Contains(stno))
                if (dbStockCloseDic[stno].Count >= iCheckDays)
                {
                    double dbmax = dbStockHighDic[stno].Max();
                    double dblow = dbStockLowDic[stno].Min();

                    double dbMaxEarn = (dbmax - dbStockCloseDic[stno].ElementAt(0)) * 100 / dbStockCloseDic[stno].ElementAt(0);
                    double dbMaxloss = (dblow - dbStockCloseDic[stno].ElementAt(0)) * 100 / dbStockCloseDic[stno].ElementAt(0);

                    double dbEarn = (dbStockCloseDic[stno].ElementAt(iCheckDays-1) - dbStockCloseDic[stno].ElementAt(0)) * 100 / dbStockCloseDic[stno].ElementAt(0);

                    string strmsg = string.Format("{0},{1:F2},{2:F2},{3:F2}", stno, dbEarn, dbMaxEarn, dbMaxloss);
                    Console.WriteLine(strmsg);
                }
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            CandidateParty3Stock.Clear();

            foreach(string sn in dicStock3Party.Keys)
            {
                int icountday = 0;
                int iBuyday = 0;
                int iTrustBuyday = 0;
                DateTime dtStart = monthCalendar1.SelectionStart;


                while (icountday < 5 )
                {
                    if(workingdays.Contains(dtStart))
                    {
                        icountday++;
                        if (dicStock3Party[sn].ContainsKey(dtStart))
                        {
                            int iFS = int.Parse(dicStock3Party[sn][dtStart].ForeigenTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                            if (iFS > 0)
                            {
                                iBuyday++;
                            }
                            int iT = int.Parse(dicStock3Party[sn][dtStart].TrustTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                            if (iT > 0)
                            {
                                iTrustBuyday++;
                            }
                        }
                    }

                    dtStart = dtStart.AddDays(-1);
                }
                if (iBuyday >= 5 && iTrustBuyday>=5)
                {
                    Console.WriteLine(sn);
                    CandidateParty3Stock.Add(sn);
                }
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            int iRange = 3;  //iRange天內
            double dbPercent = 0.015; // 成長數大於總股數
            DateTime dtStart = monthCalendar1.SelectionStart;
            DateTime dtLoopStart = monthCalendar1.SelectionStart;

            string strMsg = string.Format("找{0}前{1}天內三大買超過總股數{2:F4} , 佔成交比重", dtStart.ToString("MM/dd"), iRange, dbPercent);
            Console.WriteLine(strMsg);

            CandidateParty3Stock.Clear();

            //CandidateWeekChipSet.Clear();            
            string strBasePath = "D:\\Chips\\stock\\TDCC_OD_1-5_";
            string strReadFile = "";
            int icountchip = 0;
            while (true)
            {
                string strcsv = strBasePath + dtLoopStart.ToString("yyyyMMdd") + ".csv";
                if (File.Exists(strcsv))
                {
                    strReadFile = strcsv;
                    break;
                }
                dtLoopStart = dtLoopStart.AddDays(-1);
            }
            Dictionary<string, StockLevelCount> weekLevelCount = new Dictionary<string, StockLevelCount>();
            int iReadCount = 0;
            IEnumerable<string> lines = File.ReadLines(strReadFile);
            foreach (string line in lines)  // 取流通在外數
            {
                if (iReadCount > 0)
                {
                    string[] strSplitLine = line.Split(',');
                    if (weekLevelCount.ContainsKey(strSplitLine[1]))
                    {
                        int iLevel = int.Parse(strSplitLine[2]) - 1;
                        weekLevelCount[strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                        weekLevelCount[strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                        weekLevelCount[strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                    }
                    else
                    {
                        weekLevelCount.Add(strSplitLine[1], new StockLevelCount());
                        int iLevel = int.Parse(strSplitLine[2]) - 1;
                        weekLevelCount[strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                        weekLevelCount[strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                        weekLevelCount[strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                    }
                }
                iReadCount++;
            }

            Dictionary<string, long> stock3PartyVol = new Dictionary<string, long>();
            foreach (string strStockNo in weekLevelCount.Keys)
            {
                long l3PTotal = 0;
                long lHedgeTotal = 0;
                long lTrust = 0;
                long lSelf = 0;
                long lTotal = 0;
                long lFor = 0;
                long lVol = 0;
                long ladd = 0;
                int iLoopDay = 0;
                dtLoopStart = dtStart;
                while (iLoopDay < iRange)  // 計算最近幾天
                {
                    if (workingdays.Contains(dtLoopStart))
                    {
                        if(dicStock3Party.ContainsKey(strStockNo))
                        {
                            if (dicStock3Party[strStockNo].ContainsKey(dtLoopStart) && dbStockVolDic[strStockNo].ContainsKey(dtLoopStart))
                            {
                                int iTotal = int.Parse(dicStock3Party[strStockNo][dtLoopStart].Party3Total, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                int iHedge = int.Parse(dicStock3Party[strStockNo][dtLoopStart].SelfHedgeTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                int iHedgeBuy = int.Parse(dicStock3Party[strStockNo][dtLoopStart].SelfHedgeBuy, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                int iHedgeSell = int.Parse(dicStock3Party[strStockNo][dtLoopStart].SelfHedgeSell, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                int iForeigen = int.Parse(dicStock3Party[strStockNo][dtLoopStart].ForeigenTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                int iTrust = int.Parse(dicStock3Party[strStockNo][dtLoopStart].TrustTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                int iSelf = int.Parse(dicStock3Party[strStockNo][dtLoopStart].SelfSelfTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                ladd += (iTotal - iHedge); // 計算不含避險3大買量
                                //ladd += iTotal; // 全算
                                //ladd += (iTrust + iForeigen); //外 信
                                //ladd += iHedge;

                                lTotal += (iTotal - iHedge);
                                l3PTotal += iTotal;
                                lHedgeTotal += iHedge;
                                lTrust += iTrust;
                                lSelf += iSelf;
                                lFor += iForeigen;

                                lVol += dbStockVolDic[strStockNo][dtLoopStart];
                            }
                        }
                        iLoopDay++;
                    }
                    
                    dtLoopStart = dtLoopStart.AddDays(-1);
                }

                stock3PartyVol.Add(strStockNo,ladd);

                if(weekLevelCount.ContainsKey(strStockNo))
                {
                    double dbper = (double)stock3PartyVol[strStockNo] / (double)weekLevelCount[strStockNo].LevelVol[16];
                    double dbpervol = (double)stock3PartyVol[strStockNo] / (double)lVol;
                    if (dbper > dbPercent) // 成長數
                    {
                        CandidateParty3Stock.Add(strStockNo);
                        string strmsg = string.Format("{0},{1:F4},{2:F4},{3},{4},外:{5},信:{6}", strStockNo, dbper, dbpervol, dicStockNoMapping[strStockNo], dbEvenydayOLHCDic[strStockNo][dtStart][3], lFor, lTrust);
                        Console.WriteLine(strmsg);
                    }
                }
            }    
            /*foreach (string strStockNo in weekLevelCount.Keys)
            {
                //if (stock3PartyVol[strStockNo] >　0)
                {
                    double dbper = (double)stock3PartyVol[strStockNo] / (double)weekLevelCount[strStockNo].LevelVol[16]; 
                    if (dbper > dbPercent) // 成長數
                    {
                        CandidateParty3Stock.Add(strStockNo);
                        Console.WriteLine(strStockNo);
                    }
                    int m = 0;
                }
            }*/
        }

        private void button10_Click(object sender, EventArgs e)
        {
            DateTime dtEnd = new DateTime(2018, 10, 20);
            int iUseCount = 120;
            Dictionary<string, double> dicStockFin = new Dictionary<string, double>();
            Dictionary<string, double> dicStockMar = new Dictionary<string, double>();

            CandidateFinMarStock.Clear();
            DateTime dtStart = monthCalendar1.SelectionStart;
            DateTime dtLoopStart = monthCalendar1.SelectionStart;
            //CandidateWeekChipSet.Clear();
            string strBasePath = "D:\\Chips\\stock\\TDCC_OD_1-5_";
            string strReadFile = "";
            while (true)
            {
                string strcsv = strBasePath + dtLoopStart.ToString("yyyyMMdd") + ".csv";
                if (File.Exists(strcsv))
                {
                    strReadFile = strcsv;
                    break;
                }
                dtLoopStart = dtLoopStart.AddDays(-1);
            }
            Dictionary<string, StockLevelCount> weekLevelCount = new Dictionary<string, StockLevelCount>();
            int iReadCount = 0;
            IEnumerable<string> lines = File.ReadLines(strReadFile);
            foreach (string line in lines)
            {
                if (iReadCount > 0)
                {
                    string[] strSplitLine = line.Split(',');
                    if (weekLevelCount.ContainsKey(strSplitLine[1]))
                    {
                        int iLevel = int.Parse(strSplitLine[2]) - 1;
                        weekLevelCount[strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                        weekLevelCount[strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                        weekLevelCount[strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                    }
                    else
                    {
                        weekLevelCount.Add(strSplitLine[1], new StockLevelCount());
                        int iLevel = int.Parse(strSplitLine[2]) - 1;
                        weekLevelCount[strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                        weekLevelCount[strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                        weekLevelCount[strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                    }
                }
                iReadCount++;
            }


            Dictionary<string, List<double>> dicFinancing = new Dictionary<string, List<double>>();
            Dictionary<string, List<double>> dicMarriage = new Dictionary<string, List<double>>();
            Dictionary<string, double> dicFinancingAvg = new Dictionary<string, double>();
            Dictionary<string, double> dicFinancingDev = new Dictionary<string, double>();
            Dictionary<string, double> dicMarriageAvg = new Dictionary<string, double>();
            Dictionary<string, double> dicMarriageDev = new Dictionary<string, double>();
            Dictionary<string, double> dicFinancingEnd = new Dictionary<string, double>();
            Dictionary<string, double> dicMarriageEnd = new Dictionary<string, double>();

            foreach (string keys in dicStockFinancing.Keys)
            {
                DateTime dtStartFirst = new DateTime(1970, 1, 1);
                List<double> FinList = new List<double>();
                List<double> MarList = new List<double>();
                DateTime dtIterator = dtStart;

                int iCount = 0;
                while (dtIterator >= dtEnd && iCount < iUseCount)
                {
                    if (dicStockFinancing[keys].ContainsKey(dtIterator))
                    {
                        int iFin = dicStockFinancing[keys][dtIterator];
                        int iMar = dicStockMarriage[keys][dtIterator];
                        FinList.Add(iFin);
                        MarList.Add(iMar);
                        iCount++;
                        if (dtStartFirst.Year == 1970)
                            dtStartFirst = dtIterator;
                    }
                    dtIterator = dtIterator.AddDays(-1);
                }
                if (FinList.Count == iUseCount)
                {
                    double dbFinAvg;
                    double dbFinSDev = StandardDeviation(FinList, out dbFinAvg);
                    double dbMarAvg;
                    double dbFMarSDev = StandardDeviation(MarList, out dbMarAvg);
                    dicFinancingAvg.Add(keys, dbFinAvg);
                    dicFinancingDev.Add(keys, dbFinSDev);
                    dicMarriageAvg.Add(keys, dbMarAvg);
                    dicMarriageDev.Add(keys, dbFMarSDev);
                    dicFinancingEnd.Add(keys, dicStockFinancing[keys][dtStartFirst]);
                    dicMarriageEnd.Add(keys, dicStockMarriage[keys][dtStartFirst]);
                    dicFinancing.Add(keys, FinList);
                    dicMarriage.Add(keys, MarList);
                }
                else
                {
                }
            }

            foreach (string keys in dicFinancing.Keys)
            {
                //if (dicFinancing[keys].ElementAt(0) > (dicFinancingAvg[keys] + dicFinancingDev[keys] * 3))// && dicMarriage[keys].ElementAt(0) > (dicMarriageAvg[keys] + dicMarriageDev[keys] * 3))
                double dbFinmax = dicFinancing[keys].Max();
                double dbMarmax = dicMarriage[keys].Max();
                if (dicFinancing[keys].ElementAt(2) == dbFinmax || dicFinancing[keys].ElementAt(1) == dbFinmax || dicFinancing[keys].ElementAt(0) == dbFinmax)
                {
                    //Console.WriteLine(keys);
                    dicStockFin.Add(keys, dbFinmax);
                }
                if (dicMarriage[keys].ElementAt(2) == dbMarmax || dicMarriage[keys].ElementAt(1) == dbMarmax || dicMarriage[keys].ElementAt(0) == dbMarmax)
                {
                    //Console.WriteLine(keys);
                    dicStockMar.Add(keys, dbMarmax);
                }
            }
            

            foreach (string strStockNo in weekLevelCount.Keys)
            {
                if (dicStockFin.ContainsKey(strStockNo) && dicStockFin[strStockNo] > 0)
                {
                    double dbper = (double)dicStockFin[strStockNo] * 100000 / (double)weekLevelCount[strStockNo].LevelVol[16];
                    //Console.WriteLine(dbper.ToString());
                    if (dbper > 1)
                    {
                        CandidateParty3Stock.Add(strStockNo);
                        Console.WriteLine(strStockNo);
                    }
                }
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            int iRange = 3;  //iRange天內
            double dbPercent = 0.003; // 成長數
            CandidateParty3Stock.Clear();
            DateTime dtStart = monthCalendar1.SelectionStart;
            DateTime dtLoopStart = monthCalendar1.SelectionStart;
            //CandidateWeekChipSet.Clear();
            string strBasePath = "D:\\Chips\\stock\\TDCC_OD_1-5_";
            string strReadFile = "";
            int icountchip = 0;
            while (true)
            {
                string strcsv = strBasePath + dtLoopStart.ToString("yyyyMMdd") + ".csv";
                if (File.Exists(strcsv))
                {
                    strReadFile = strcsv;
                    break;
                }
                dtLoopStart = dtLoopStart.AddDays(-1);
            }
            Dictionary<string, StockLevelCount> weekLevelCount = new Dictionary<string, StockLevelCount>();
            int iReadCount = 0;
            IEnumerable<string> lines = File.ReadLines(strReadFile);
            foreach (string line in lines)  // 取流通在外數
            {
                if (iReadCount > 0)
                {
                    string[] strSplitLine = line.Split(',');
                    if (weekLevelCount.ContainsKey(strSplitLine[1]))
                    {
                        int iLevel = int.Parse(strSplitLine[2]) - 1;
                        weekLevelCount[strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                        weekLevelCount[strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                        weekLevelCount[strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                    }
                    else
                    {
                        weekLevelCount.Add(strSplitLine[1], new StockLevelCount());
                        int iLevel = int.Parse(strSplitLine[2]) - 1;
                        weekLevelCount[strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                        weekLevelCount[strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                        weekLevelCount[strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                    }
                }
                iReadCount++;
            }

            Dictionary<string, long> stock3PartyVol = new Dictionary<string, long>();
            foreach (string strStockNo in weekLevelCount.Keys)
            {
                long ladd = 0;
                int iLoopDay = 0;
                dtLoopStart = dtStart;
                while (iLoopDay < iRange)  // 計算最近幾天
                {
                    if (workingdays.Contains(dtLoopStart))
                    {
                        if (dicStock3Party.ContainsKey(strStockNo))
                        {
                            if (dicStock3Party[strStockNo].ContainsKey(dtLoopStart))
                            {
                                int iTotal = int.Parse(dicStock3Party[strStockNo][dtLoopStart].Party3Total, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                int iHedge = int.Parse(dicStock3Party[strStockNo][dtLoopStart].SelfHedgeTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                int iHedgeBuy = int.Parse(dicStock3Party[strStockNo][dtLoopStart].SelfHedgeBuy, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                int iHedgeSell = int.Parse(dicStock3Party[strStockNo][dtLoopStart].SelfHedgeSell, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                int iForeigen = int.Parse(dicStock3Party[strStockNo][dtLoopStart].ForeigenTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                int iTrust = int.Parse(dicStock3Party[strStockNo][dtLoopStart].TrustTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                int iSelf = int.Parse(dicStock3Party[strStockNo][dtLoopStart].SelfSelfTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                //ladd += (iTotal - iHedge); // 計算不含避險3大買量
                                //ladd += iTotal; // 全算
                                //ladd += (iTrust + iForeigen); //外 信
                                ladd += iHedge;
                            }
                        }
                        iLoopDay++;
                    }

                    dtLoopStart = dtLoopStart.AddDays(-1);
                }

                stock3PartyVol.Add(strStockNo, ladd);
            }
            foreach (string strStockNo in weekLevelCount.Keys)
            {
                //if (stock3PartyVol[strStockNo] >　0)
                {
                    double dbper = (double)stock3PartyVol[strStockNo] / (double)weekLevelCount[strStockNo].LevelVol[16];
                    if (dbper > dbPercent) // 成長數
                    {
                        CandidateParty3Stock.Add(strStockNo);
                        Console.WriteLine(strStockNo);
                    }
                    int m = 0;
                }
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            int iRange = 10;  //iRange天內
            double dbPercent = 0.01; // 成長數大於總股數
            CandidateParty3Stock.Clear();
            DateTime dtStart = monthCalendar1.SelectionStart;
            DateTime dtLoopStart = monthCalendar1.SelectionStart;
            //CandidateWeekChipSet.Clear();
            string strBasePath = "D:\\Chips\\stock\\TDCC_OD_1-5_";
            string strReadFile = "";
            int icountchip = 0;
            while (true)
            {
                string strcsv = strBasePath + dtLoopStart.ToString("yyyyMMdd") + ".csv";
                if (File.Exists(strcsv))
                {
                    strReadFile = strcsv;
                    break;
                }
                dtLoopStart = dtLoopStart.AddDays(-1);
            }
            Dictionary<string, StockLevelCount> weekLevelCount = new Dictionary<string, StockLevelCount>();
            int iReadCount = 0;
            IEnumerable<string> lines = File.ReadLines(strReadFile);
            foreach (string line in lines)  // 取流通在外數
            {
                if (iReadCount > 0)
                {
                    string[] strSplitLine = line.Split(',');
                    if (weekLevelCount.ContainsKey(strSplitLine[1]))
                    {
                        int iLevel = int.Parse(strSplitLine[2]) - 1;
                        weekLevelCount[strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                        weekLevelCount[strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                        weekLevelCount[strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                    }
                    else
                    {
                        weekLevelCount.Add(strSplitLine[1], new StockLevelCount());
                        int iLevel = int.Parse(strSplitLine[2]) - 1;
                        weekLevelCount[strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                        weekLevelCount[strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                        weekLevelCount[strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                    }
                }
                iReadCount++;
            }

            Dictionary<string, long> stock3PartyVol = new Dictionary<string, long>();
            foreach (string strStockNo in weekLevelCount.Keys)
            {
                long ladd = 0;
                int iLoopDay = 0;
                dtLoopStart = dtStart;
                while (iLoopDay < iRange)  // 計算最近幾天
                {
                    if (workingdays.Contains(dtLoopStart))
                    {
                        if (dicStock3Party.ContainsKey(strStockNo))
                        {
                            if (dicStock3Party[strStockNo].ContainsKey(dtLoopStart))
                            {
                                int iTotal = int.Parse(dicStock3Party[strStockNo][dtLoopStart].Party3Total, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                int iHedge = int.Parse(dicStock3Party[strStockNo][dtLoopStart].SelfHedgeTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                int iHedgeBuy = int.Parse(dicStock3Party[strStockNo][dtLoopStart].SelfHedgeBuy, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                int iHedgeSell = int.Parse(dicStock3Party[strStockNo][dtLoopStart].SelfHedgeSell, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                int iForeigen = int.Parse(dicStock3Party[strStockNo][dtLoopStart].ForeigenTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                int iTrust = int.Parse(dicStock3Party[strStockNo][dtLoopStart].TrustTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                int iSelf = int.Parse(dicStock3Party[strStockNo][dtLoopStart].SelfSelfTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                //ladd += (iTotal - iHedge); // 計算不含避險3大買量
                                //ladd += iTotal; // 全算
                                //ladd += (iTrust + iForeigen); //外 信
                                ladd += iHedge;
                            }
                        }
                        iLoopDay++;
                    }

                    dtLoopStart = dtLoopStart.AddDays(-1);
                }

                stock3PartyVol.Add(strStockNo, ladd);
            }
            foreach (string strStockNo in weekLevelCount.Keys)
            {
                //if (stock3PartyVol[strStockNo] >　0)
                {
                    double dbper = (double)stock3PartyVol[strStockNo] / (double)weekLevelCount[strStockNo].LevelVol[16];
                    if (dbper > dbPercent) // 成長數
                    {
                        CandidateParty3Stock.Add(strStockNo);
                        Console.WriteLine(strStockNo);
                    }
                    int m = 0;
                }
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            int iRange = 1;  //iRange天內
            double dbPercent = 0.004; // 成長數大於總股數
            CandidateParty3Stock.Clear();
            DateTime dtStart = monthCalendar1.SelectionStart;
            DateTime dtLoopStart = monthCalendar1.SelectionStart;
            //CandidateWeekChipSet.Clear();
            string strBasePath = "D:\\Chips\\stock\\TDCC_OD_1-5_";
            string strReadFile = "";
            int icountchip = 0;
            while (true)
            {
                string strcsv = strBasePath + dtLoopStart.ToString("yyyyMMdd") + ".csv";
                if (File.Exists(strcsv))
                {
                    strReadFile = strcsv;
                    break;
                }
                dtLoopStart = dtLoopStart.AddDays(-1);
            }
            Dictionary<string, StockLevelCount> weekLevelCount = new Dictionary<string, StockLevelCount>();
            int iReadCount = 0;
            IEnumerable<string> lines = File.ReadLines(strReadFile);
            foreach (string line in lines)  // 取流通在外數
            {
                if (iReadCount > 0)
                {
                    string[] strSplitLine = line.Split(',');
                    if (weekLevelCount.ContainsKey(strSplitLine[1]))
                    {
                        int iLevel = int.Parse(strSplitLine[2]) - 1;
                        weekLevelCount[strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                        weekLevelCount[strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                        weekLevelCount[strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                    }
                    else
                    {
                        weekLevelCount.Add(strSplitLine[1], new StockLevelCount());
                        int iLevel = int.Parse(strSplitLine[2]) - 1;
                        weekLevelCount[strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                        weekLevelCount[strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                        weekLevelCount[strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                    }
                }
                iReadCount++;
            }

            Dictionary<string, long> stock3PartyVol = new Dictionary<string, long>();
            foreach (string strStockNo in weekLevelCount.Keys)
            {
                int iHedgeTotal = 0;
                long ladd = 0;
                int iLoopDay = 0;
                dtLoopStart = dtStart;
                while (iLoopDay < iRange)  // 計算最近幾天
                {
                    if (workingdays.Contains(dtLoopStart))
                    {
                        if (dicStock3Party.ContainsKey(strStockNo))
                        {
                            if (dicStock3Party[strStockNo].ContainsKey(dtLoopStart))
                            {
                                int iTotal = int.Parse(dicStock3Party[strStockNo][dtLoopStart].Party3Total, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                int iHedge = int.Parse(dicStock3Party[strStockNo][dtLoopStart].SelfHedgeTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                int iHedgeBuy = int.Parse(dicStock3Party[strStockNo][dtLoopStart].SelfHedgeBuy, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                int iHedgeSell = int.Parse(dicStock3Party[strStockNo][dtLoopStart].SelfHedgeSell, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                int iForeigen = int.Parse(dicStock3Party[strStockNo][dtLoopStart].ForeigenTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                int iTrust = int.Parse(dicStock3Party[strStockNo][dtLoopStart].TrustTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                int iSelf = int.Parse(dicStock3Party[strStockNo][dtLoopStart].SelfSelfTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                ladd += (iTotal - iHedge); // 計算不含避險3大買量
                                iHedgeTotal += iHedge;
                                //ladd += iTotal; // 全算
                                //ladd += (iTrust + iForeigen); //外 信
                                //ladd += iHedge;
                            }
                        }
                        iLoopDay++;
                    }

                    dtLoopStart = dtLoopStart.AddDays(-1);
                }
                if (iHedgeTotal < 0)
                stock3PartyVol.Add(strStockNo, ladd);
            }
            foreach (string strStockNo in weekLevelCount.Keys)
            {
                if (stock3PartyVol.ContainsKey(strStockNo))
                {
                    double dbper = (double)stock3PartyVol[strStockNo] / (double)weekLevelCount[strStockNo].LevelVol[16];
                    if (dbper > dbPercent) // 成長數
                    {
                        CandidateParty3Stock.Add(strStockNo);
                        Console.WriteLine(strStockNo);
                    }
                    int m = 0;
                }
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            string strsn = textBoxStockPeroid.Text;
            int istartyear = int.Parse(textBoxFromYear.Text);
            int istartmonth = int.Parse(textBoxFromMonth.Text);
            int istartday = int.Parse(textBoxFromDay.Text);
            int itoyear = int.Parse(textBoxToYear.Text);
            int itomonth = int.Parse(textBoxToMonth.Text);
            int itoday = int.Parse(textBoxToDay.Text);

            DateTime dtStart = new DateTime(istartyear, istartmonth, istartday);
            DateTime dtEnd = new DateTime(itoyear, itomonth, itoday);
            DateTime dtloop = dtStart;
            string strMsg = string.Format("{0} {1} 自{2} 至{3}", strsn, dicStockNoMapping[strsn], dtStart.ToString("MM/dd"), dtEnd.ToString("MM/dd"));
            Console.WriteLine(strMsg);
            DateTime dtLoopStart = DateTime.Now;
            string strBasePath = "D:\\Chips\\stock\\TDCC_OD_1-5_";
            string strReadFile = "";
            int icountchip = 0;
            while (true)
            {
                string strcsv = strBasePath + dtLoopStart.ToString("yyyyMMdd") + ".csv";
                if (File.Exists(strcsv))
                {
                    strReadFile = strcsv;
                    break;
                }
                dtLoopStart = dtLoopStart.AddDays(-1);
            }
            Dictionary<string, StockLevelCount> weekLevelCount = new Dictionary<string, StockLevelCount>();
            int iReadCount = 0;
            IEnumerable<string> lines = File.ReadLines(strReadFile);
            foreach (string line in lines)  // 取流通在外數
            {
                if (iReadCount > 0)
                {
                    string[] strSplitLine = line.Split(',');
                    if (weekLevelCount.ContainsKey(strSplitLine[1]))
                    {
                        int iLevel = int.Parse(strSplitLine[2]) - 1;
                        weekLevelCount[strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                        weekLevelCount[strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                        weekLevelCount[strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                    }
                    else
                    {
                        weekLevelCount.Add(strSplitLine[1], new StockLevelCount());
                        int iLevel = int.Parse(strSplitLine[2]) - 1;
                        weekLevelCount[strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                        weekLevelCount[strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                        weekLevelCount[strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                    }
                }
                iReadCount++;
            }

            long lTrust = 0;
            long lForign = 0;
            long l3PTotal=0;
            long lHedgeTotal = 0;
            long lSelf = 0;
            long lTotal = 0;
            while(dtloop <= dtEnd)
            {
                if(dicStock3Party[strsn].ContainsKey(dtloop))
                {
                    int iTotal = int.Parse(dicStock3Party[strsn][dtloop].Party3Total, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                    int iHedge = int.Parse(dicStock3Party[strsn][dtloop].SelfHedgeTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                    int iTrust = int.Parse(dicStock3Party[strsn][dtloop].TrustTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                    int iForign = int.Parse(dicStock3Party[strsn][dtloop].ForeigenTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                    int iself = int.Parse(dicStock3Party[strsn][dtloop].SelfSelfTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);

                    lTotal += (iTotal - iHedge);
                    l3PTotal += iTotal;
                    lHedgeTotal += iHedge;
                    lTrust += iTrust;
                    lSelf += iself;
                    lForign += iForign;
      
                }
                dtloop = dtloop.AddDays(1);
            }

            double dbPercentTotal = (double)lTotal / (double)weekLevelCount[strsn].LevelVol[16];
            double dbPercent3PTotal = (double)l3PTotal / (double)weekLevelCount[strsn].LevelVol[16];
            double dbPercentHedgeTotal = (double)lHedgeTotal / (double)weekLevelCount[strsn].LevelVol[16];
            double dbPercentTrust = (double)lTrust / (double)weekLevelCount[strsn].LevelVol[16];
            double dbPercentForign = (double)lForign / (double)weekLevelCount[strsn].LevelVol[16];
            double dbPercentSelf = (double)lSelf / (double)weekLevelCount[strsn].LevelVol[16];



            // 主力1 5 10 20 60 120
            //https://fubon-ebrokerdj.fbs.com.tw/z/zc/zco/zco_8155_6.djhtm
            int ibuy0=0;
            int iSell0=0;
            bool b0 = GetKeyBuy(strsn, 1, out ibuy0, out iSell0);

            int ibuy5 = 0;
            int iSell5 = 0;
            bool b5 = GetKeyBuy(strsn, 2, out ibuy5, out iSell5);

            int ibuy10 = 0;
            int iSell10 = 0;
            bool b10 = GetKeyBuy(strsn, 3, out ibuy10, out iSell10);

            int ibuy20 = 0;
            int iSell20 = 0;
            bool b20 = GetKeyBuy(strsn, 4, out ibuy20, out iSell20);

            int ibuy60 = 0;
            int iSell60 = 0;
            bool b60 = GetKeyBuy(strsn, 6, out ibuy60, out iSell60);

            strMsg = string.Format("3大{0:F4}，外{1:F4}，信{2:F4}，自{3:F4}，避{4:F4},主力1D:{5},主力5D:{6},主力10D:{7},主力20D:{8},主力60D:{9}", 
                dbPercent3PTotal, dbPercentForign, dbPercentTrust, dbPercentSelf, dbPercentHedgeTotal,
                ibuy0 - iSell0,ibuy5 - iSell5,ibuy10 - iSell0,ibuy20 - iSell20,ibuy60 - iSell60);
            Console.WriteLine(strMsg);
            int m = 0;
        }

        private void button15_Click(object sender, EventArgs e)
        {

            dbStockVolDic.Clear();



            string strSavePath = "D:\\Chips\\eachdaystock\\";

            Dictionary<string, Dictionary<DateTime, double>> dbStockOpenDic = new Dictionary<string, Dictionary<DateTime, double>>();
            Dictionary<string, Dictionary<DateTime, double>> dbStockLowDic = new Dictionary<string, Dictionary<DateTime, double>>();
            Dictionary<string, Dictionary<DateTime, double>> dbStockHighDic = new Dictionary<string, Dictionary<DateTime, double>>();
            Dictionary<string, Dictionary<DateTime, double>> dbStockCloseDic = new Dictionary<string, Dictionary<DateTime, double>>();

            Dictionary<string, Dictionary<DateTime, double[]>> dbStockOLHCDic = new Dictionary<string, Dictionary<DateTime, double[]>>();
            HashSet<DateTime> workingdayssave = new HashSet<DateTime>();

            int iCheckDays = 360;
            int iCountDataDays = 0;
            DateTime dtStart = new DateTime(DateTime.Now.Year,DateTime.Now.Month,DateTime.Now.Day);
            while (iCountDataDays < iCheckDays)
            {
                bool bisworkday = false;

                string jsonFilePath = string.Format("D:\\Chips\\eachdaystock\\{0}.json", dtStart.ToString("yyyyMMdd"));
                string jsonFilePathTpex = string.Format("D:\\Chips\\eachdaystock\\{0}_tpex.json", dtStart.ToString("yyyyMMdd"));
                if(File.Exists(jsonFilePath))
                {
                    string json = File.ReadAllText(jsonFilePath);
                    Dictionary<string, object> json_Dictionary = JsonConvert.DeserializeObject<Dictionary<string, object>>(json);
                    if (json_Dictionary.Count > 2)
                    {
                        iCountDataDays++;

                        List<List<string>> json_data9 = JsonConvert.DeserializeObject<List<List<string>>>(json_Dictionary["data9"].ToString());
                        foreach (List<string> pricelst in json_data9)
                        {
                            string strstockno = pricelst.ElementAt(0);
                            if (!dbStockCloseDic.ContainsKey(strstockno))
                            {
                                bisworkday = true;
                                dbStockCloseDic.Add(strstockno, new Dictionary<DateTime, double>());
                                dbStockOpenDic.Add(strstockno, new Dictionary<DateTime, double>());
                                dbStockHighDic.Add(strstockno, new Dictionary<DateTime, double>());
                                dbStockLowDic.Add(strstockno, new Dictionary<DateTime, double>());
                                dbStockOLHCDic.Add(strstockno, new Dictionary<DateTime, double[]>());
                                dbStockVolDic.Add(strstockno, new Dictionary<DateTime, long>());
                            }
                            string strVol = pricelst.ElementAt(2);
                            string strcloseprice = pricelst.ElementAt(8);
                            string stropenprice = pricelst.ElementAt(5);
                            string strhighprice = pricelst.ElementAt(6);
                            string strlowprice = pricelst.ElementAt(7);
                            if (strcloseprice != "--")
                            {
                                double dbopen = double.Parse(stropenprice);
                                double dbclose = double.Parse(strcloseprice);
                                double dbhigh = double.Parse(strhighprice);
                                double dblow = double.Parse(strlowprice);
                                long lvol = long.Parse(strVol,NumberStyles.AllowThousands);
                                dbStockCloseDic[strstockno].Add(dtStart, dbclose);
                                dbStockOpenDic[strstockno].Add(dtStart, dbopen);
                                dbStockHighDic[strstockno].Add(dtStart, dbhigh);
                                dbStockLowDic[strstockno].Add(dtStart, dblow);
                                dbStockOLHCDic[strstockno].Add(dtStart, new double[4]);
                                dbStockOLHCDic[strstockno][dtStart][0] = dbopen;
                                dbStockOLHCDic[strstockno][dtStart][1] = dblow;
                                dbStockOLHCDic[strstockno][dtStart][2] = dbhigh;
                                dbStockOLHCDic[strstockno][dtStart][3] = dbclose;
                                dbStockVolDic[strstockno].Add(dtStart, lvol);
                                if (strstockno == "3691")
                                {
                                    int m = 0;
                                }
                                bisworkday = true;
                            }

                        }

                        string jsonTpex = File.ReadAllText(jsonFilePathTpex);
                        Dictionary<string, object> jsonTpex_Dictionary = JsonConvert.DeserializeObject<Dictionary<string, object>>(jsonTpex);
                        List<List<string>> json_aaData = JsonConvert.DeserializeObject<List<List<string>>>(jsonTpex_Dictionary["aaData"].ToString());
                        foreach (List<string> pricelst in json_aaData)
                        {
                            string strstockno = pricelst.ElementAt(0);
                            string strcloseprice = pricelst.ElementAt(2);
                            string strhighprice = pricelst.ElementAt(5);
                            string strlowprice = pricelst.ElementAt(6);
                            string stropenprice = pricelst.ElementAt(4);
                            string strvol = pricelst.ElementAt(7);
                            if (!dbStockCloseDic.ContainsKey(strstockno))
                            {
                                dbStockCloseDic.Add(strstockno, new Dictionary<DateTime, double>());
                                dbStockOpenDic.Add(strstockno, new Dictionary<DateTime, double>());
                                dbStockHighDic.Add(strstockno, new Dictionary<DateTime, double>());
                                dbStockLowDic.Add(strstockno, new Dictionary<DateTime, double>());
                                dbStockOLHCDic.Add(strstockno, new Dictionary<DateTime, double[]>());
                                dbStockVolDic.Add(strstockno, new Dictionary<DateTime, long>());
                            }

                            if (strcloseprice != "---" && strcloseprice != "----")
                            {
                                double dbclose = double.Parse(strcloseprice);
                                double dbhigh = double.Parse(strhighprice);
                                double dblow = double.Parse(strlowprice);
                                double dbopen = double.Parse(stropenprice);
                                long lvol = long.Parse(strvol,NumberStyles.AllowThousands);
                                dbStockCloseDic[strstockno].Add(dtStart, dbclose);
                                dbStockOpenDic[strstockno].Add(dtStart, dbopen);
                                dbStockHighDic[strstockno].Add(dtStart, dbhigh);
                                dbStockLowDic[strstockno].Add(dtStart, dblow);
                                dbStockOLHCDic[strstockno].Add(dtStart, new double[4]);
                                dbStockOLHCDic[strstockno][dtStart][0] = dbopen;
                                dbStockOLHCDic[strstockno][dtStart][1] = dblow;
                                dbStockOLHCDic[strstockno][dtStart][2] = dbhigh;
                                dbStockOLHCDic[strstockno][dtStart][3] = dbclose;
                                dbStockVolDic[strstockno].Add(dtStart, lvol);
                                if (strstockno == "3691")
                                {
                                    int m = 0;
                                }
                            }
                        }
                    }

                    if (bisworkday)
                    {
                        workingdayssave.Add(new DateTime(dtStart.Year, dtStart.Month, dtStart.Day));
                    }
                }

                dtStart = dtStart.AddDays(-1);
            }

            string jsonClose = JsonConvert.SerializeObject(dbStockCloseDic);
            string jsonOpen = JsonConvert.SerializeObject(dbStockOpenDic);
            string jsonHigh = JsonConvert.SerializeObject(dbStockHighDic);
            string jsonLow = JsonConvert.SerializeObject(dbStockLowDic);
            string jsonOLHC = JsonConvert.SerializeObject(dbStockOLHCDic);
            string jsonWorking = JsonConvert.SerializeObject(workingdayssave);
            string jsonVol = JsonConvert.SerializeObject(dbStockVolDic);
            File.WriteAllText("Y:\\Working" + DateTime.Now.ToString("MMdd") + ".json", jsonWorking);
            File.WriteAllText("Y:\\Close" + DateTime.Now.ToString("MMdd") + ".json", jsonClose);
            File.WriteAllText("Y:\\Open" + DateTime.Now.ToString("MMdd") + ".json", jsonOpen);
            File.WriteAllText("Y:\\High" + DateTime.Now.ToString("MMdd") + ".json", jsonHigh);
            File.WriteAllText("Y:\\Low" + DateTime.Now.ToString("MMdd") + ".json", jsonLow);
            File.WriteAllText("Y:\\OLHC" + DateTime.Now.ToString("MMdd") + ".json", jsonOLHC);
            File.WriteAllText("Y:\\VOL" + DateTime.Now.ToString("MMdd") + ".json", jsonVol);
        }


        //Rattlesnake
        int iRattlesnakeNDays = 60;
        int iLowNDays = 20;
        double dbRattlesnakeOver = 0.02;

        DateTime dtRattlesnakeBaseDate = new DateTime(2020,10, 23);   //
        DateTime dtRattlesnakeExamday = new DateTime(2019, 10, 31);    // 


        int iCheckForWeeksHolder = 8;
        List<string> StockHolderList = new List<string>();
        private void button69_Click(object sender, EventArgs e)
        {
            // 同下 button16_Click 但是歷史資料，直接讀檔
            //int iRattlesnakeNDays = 60;

            int iPreparedays = 120;
            Dictionary<string, List<double>> dbStockCloseDic = new Dictionary<string, List<double>>();
            Dictionary<string, List<double>> dbStockOpenDic = new Dictionary<string, List<double>>();
            Dictionary<string, List<double>> dbStockHighDic = new Dictionary<string, List<double>>();
            Dictionary<string, List<double>> dbStockLowDic = new Dictionary<string, List<double>>();

            Dictionary<string, List<int>> dbTrustTotalBuy = new Dictionary<string, List<int>>();
            Dictionary<string, List<int>> dbSelfTotalBuy = new Dictionary<string, List<int>>();
            Dictionary<string, List<int>> dbForTotalBuy = new Dictionary<string, List<int>>();

            string strBasePath = "D:\\Chips\\eachdaystock\\";
            DateTime dtLoopStart = dtRattlesnakeBaseDate;
            for (int icountday = 0; icountday < iPreparedays; )
            {
                string strFile = strBasePath + dtLoopStart.ToString("yyyyMMdd") + ".json";
                string strFileTpex = strBasePath + dtLoopStart.ToString("yyyyMMdd") + "_tpex.json";
                dtLoopStart = dtLoopStart.AddDays(-1);
                if (File.Exists(strFile) && File.Exists(strFileTpex))
                {
                    string json = File.ReadAllText(strFile);
                    Dictionary<string, object> json_Dictionary = JsonConvert.DeserializeObject<Dictionary<string, object>>(json);
                    if (json_Dictionary.Count > 2)
                    {
                        icountday++;

                        List<List<string>> json_data9 = JsonConvert.DeserializeObject<List<List<string>>>(json_Dictionary["data9"].ToString());
                        foreach (List<string> pricelst in json_data9)
                        {
                            string strstockno = pricelst.ElementAt(0);
                            if (!dbStockCloseDic.ContainsKey(strstockno))
                            {
                                dbStockCloseDic.Add(strstockno, new List<double>());
                                dbStockOpenDic.Add(strstockno, new List<double>());
                                dbStockHighDic.Add(strstockno, new List<double>());
                                dbStockLowDic.Add(strstockno, new List<double>());
                            }

                            string strcloseprice = pricelst.ElementAt(8);
                            string stropenprice = pricelst.ElementAt(5);
                            string strhighprice = pricelst.ElementAt(6);
                            string strlowprice = pricelst.ElementAt(7);
                            if (strcloseprice != "--")
                            {
                                double dbopen = double.Parse(stropenprice);
                                double dbclose = double.Parse(strcloseprice);
                                double dbhigh = double.Parse(strhighprice);
                                double dblow = double.Parse(strlowprice);
                                dbStockCloseDic[strstockno].Add(dbclose);
                                dbStockOpenDic[strstockno].Add(dbopen);
                                dbStockHighDic[strstockno].Add(dbhigh);
                                dbStockLowDic[strstockno].Add(dblow);
                            }
                            else
                            {
                                dbStockCloseDic[strstockno].Add(0);
                                dbStockOpenDic[strstockno].Add(0);
                                dbStockHighDic[strstockno].Add(0);
                                dbStockLowDic[strstockno].Add(0);
                            }
                        }

                        string jsonTpex = File.ReadAllText(strFileTpex);
                        Dictionary<string, object> jsonTpex_Dictionary = JsonConvert.DeserializeObject<Dictionary<string, object>>(jsonTpex);
                        List<List<string>> json_aaData = JsonConvert.DeserializeObject<List<List<string>>>(jsonTpex_Dictionary["aaData"].ToString());
                        foreach (List<string> pricelst in json_aaData)
                        {
                            string strstockno = pricelst.ElementAt(0);
                            string strcloseprice = pricelst.ElementAt(2);
                            string strhighprice = pricelst.ElementAt(5);
                            string strlowprice = pricelst.ElementAt(6);
                            string stropenprice = pricelst.ElementAt(4);

                            if (!dbStockCloseDic.ContainsKey(strstockno))
                            {
                                dbStockCloseDic.Add(strstockno, new List<double>());
                                dbStockOpenDic.Add(strstockno, new List<double>());
                                dbStockHighDic.Add(strstockno, new List<double>());
                                dbStockLowDic.Add(strstockno, new List<double>());
                            }

                            if (strcloseprice != "---" && strcloseprice != "----")
                            {
                                double dbclose = double.Parse(strcloseprice);
                                double dbhigh = double.Parse(strhighprice);
                                double dblow = double.Parse(strlowprice);
                                double dbopen = double.Parse(stropenprice);
                                dbStockCloseDic[strstockno].Add(dbclose);
                                dbStockOpenDic[strstockno].Add(dbopen);
                                dbStockHighDic[strstockno].Add(dbhigh);
                                dbStockLowDic[strstockno].Add(dblow);
                            }
                            else
                            {
                                dbStockCloseDic[strstockno].Add(0);
                                dbStockOpenDic[strstockno].Add(0);
                                dbStockHighDic[strstockno].Add(0);
                                dbStockLowDic[strstockno].Add(0);
                            }
                        }
                    }
                }
                else
                {
                    MessageBox.Show("loss file" + strFile + " or " + strFileTpex);
                }
            }

            string strBase3ptyPath = "D:\\Chips\\3Party\\";
            dtLoopStart = dtRattlesnakeBaseDate;
            for (int icountday = 0; icountday < iRattlesnakeNDays; )
            {
                string strFile = strBase3ptyPath + dtLoopStart.ToString("yyyyMMdd") + ".json";
                string strFileTpex = strBase3ptyPath + dtLoopStart.ToString("yyyyMMdd") + "_tpex.json";
                dtLoopStart = dtLoopStart.AddDays(-1);
                if (File.Exists(strFile) && File.Exists(strFileTpex))
                {
                    int mmmm = 0;
                    string json = File.ReadAllText(strFile);
                    Dictionary<string, object> json_Dictionary = JsonConvert.DeserializeObject<Dictionary<string, object>>(json);
                    if (json_Dictionary.ContainsKey("data"))
                    {
                        icountday++;
                        List<List<string>> jsondata_ListList = JsonConvert.DeserializeObject<List<List<string>>>(json_Dictionary["data"].ToString());
                        foreach (List<string> marginelist in jsondata_ListList)
                        {
                            mmmm++;
                            if (marginelist.ElementAt(0).Length == 4)
                            {
                                if (marginelist.ElementAt(0) == "1701" && dtLoopStart == new DateTime(2020,6,3))
                                {
                                    int sdvm = 0;
                                }
                                else
                                {
                                    PartyParam Assign = new PartyParam();
                                    Assign.Assign(marginelist);

                                    if (!dbTrustTotalBuy.ContainsKey(Assign.number))
                                    {
                                        dbTrustTotalBuy.Add(Assign.number, new List<int>());
                                    }
                                    if (!dbSelfTotalBuy.ContainsKey(Assign.number))
                                    {
                                        dbSelfTotalBuy.Add(Assign.number, new List<int>());
                                    }
                                    if (!dbForTotalBuy.ContainsKey(Assign.number))
                                    {
                                        dbForTotalBuy.Add(Assign.number, new List<int>());
                                    }
                                    dbTrustTotalBuy[Assign.number].Add(int.Parse(Assign.TrustTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign));
                                    dbSelfTotalBuy[Assign.number].Add(int.Parse(Assign.SelfTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign));
                                    dbForTotalBuy[Assign.number].Add(int.Parse(Assign.ForeigenTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign));
                                }
                                
                            }

                        }
                    }

                    string jsontpex = File.ReadAllText(strFileTpex);
                    Dictionary<string, object> jsontpex_Dictionary = JsonConvert.DeserializeObject<Dictionary<string, object>>(jsontpex);
                    if (jsontpex_Dictionary.ContainsKey("aaData"))
                    {
                        List<List<string>> jsondata_ListList = JsonConvert.DeserializeObject<List<List<string>>>(jsontpex_Dictionary["aaData"].ToString());
                        foreach (List<string> marginelist in jsondata_ListList)
                        {
                            if (marginelist.ElementAt(0).Length == 4)
                            {
                                PartyParam Assign = new PartyParam();
                                Assign.AssignTPEX(marginelist);

                                if (!dbTrustTotalBuy.ContainsKey(Assign.number))
                                {
                                    dbTrustTotalBuy.Add(Assign.number, new List<int>());
                                }
                                if (!dbSelfTotalBuy.ContainsKey(Assign.number))
                                {
                                    dbSelfTotalBuy.Add(Assign.number, new List<int>());
                                }
                                if (!dbForTotalBuy.ContainsKey(Assign.number))
                                {
                                    dbForTotalBuy.Add(Assign.number, new List<int>());
                                }
                                dbTrustTotalBuy[Assign.number].Add(int.Parse(Assign.TrustTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign));
                                dbSelfTotalBuy[Assign.number].Add(int.Parse(Assign.SelfTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign));
                                dbForTotalBuy[Assign.number].Add(int.Parse(Assign.ForeigenTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign));
                            }
                        }
                    }

                }
                else
                {
                    MessageBox.Show("loss file" + strFile + " or " + strFileTpex);
                }
            }

            string jsonTrust = JsonConvert.SerializeObject(dbTrustTotalBuy);
            string jsonSelf = JsonConvert.SerializeObject(dbSelfTotalBuy);
            string jsonFor = JsonConvert.SerializeObject(dbForTotalBuy);


            string jsonClose = JsonConvert.SerializeObject(dbStockCloseDic);
            string jsonOpen = JsonConvert.SerializeObject(dbStockOpenDic);
            string jsonHigh = JsonConvert.SerializeObject(dbStockHighDic);
            string jsonLow = JsonConvert.SerializeObject(dbStockLowDic);



            File.WriteAllText("Y:\\tempclose.json", jsonClose);
            File.WriteAllText("Y:\\tempopen.json", jsonOpen);
            File.WriteAllText("Y:\\temphi.json", jsonHigh);
            File.WriteAllText("Y:\\templo.json", jsonLow);
            File.WriteAllText("Y:\\tempfor.json", jsonFor);
            File.WriteAllText("Y:\\tempself.json", jsonSelf);
            File.WriteAllText("Y:\\temptrust.json", jsonTrust);

        }
        //int iRattlesnakeNDays = 60;
        //double dbRattlesnakeOver = 0.02;
        //DateTime dtRattlesnakeBaseDate = new DateTime(2019, 9, 10);
        private void button70_Click(object sender, EventArgs e)
        {
            // Load Y:\\temp...........

            Dictionary<string, List<double>> dbStockCloseDic = new Dictionary<string, List<double>>();
            Dictionary<string, List<double>> dbStockOpenDic = new Dictionary<string, List<double>>();
            Dictionary<string, List<double>> dbStockHighDic = new Dictionary<string, List<double>>();
            Dictionary<string, List<double>> dbStockLowDic = new Dictionary<string, List<double>>();

            Dictionary<string, List<int>> dbTrustTotalBuy = new Dictionary<string, List<int>>();
            Dictionary<string, List<int>> dbSelfTotalBuy = new Dictionary<string, List<int>>();
            Dictionary<string, List<int>> dbForTotalBuy = new Dictionary<string, List<int>>();

            Dictionary<string, double> dbExamDayDic = new Dictionary<string, double>();

            string jsondbStockCloseDic = File.ReadAllText("Y:\\tempclose.json");
            dbStockCloseDic = JsonConvert.DeserializeObject<Dictionary<string, List<double>>>(jsondbStockCloseDic);

            string jsondbStockopenDic = File.ReadAllText("Y:\\tempopen.json");
            dbStockOpenDic = JsonConvert.DeserializeObject<Dictionary<string, List<double>>>(jsondbStockopenDic);

            string jsondbStockhiDic = File.ReadAllText("Y:\\temphi.json");
            dbStockHighDic = JsonConvert.DeserializeObject<Dictionary<string, List<double>>>(jsondbStockhiDic);

            string jsondbStockloDic = File.ReadAllText("Y:\\templo.json");
            dbStockLowDic = JsonConvert.DeserializeObject<Dictionary<string, List<double>>>(jsondbStockloDic);

            string jsondbStockforDic = File.ReadAllText("Y:\\tempfor.json");
            dbForTotalBuy = JsonConvert.DeserializeObject<Dictionary<string, List<int>>>(jsondbStockforDic);

            string jsondbStockseldfeDic = File.ReadAllText("Y:\\tempself.json");
            dbSelfTotalBuy = JsonConvert.DeserializeObject<Dictionary<string, List<int>>>(jsondbStockseldfeDic);

            string jsondbStocktrudtDic = File.ReadAllText("Y:\\temptrust.json");
            dbTrustTotalBuy = JsonConvert.DeserializeObject<Dictionary<string, List<int>>>(jsondbStocktrudtDic);

            Dictionary<string,  int>   dFinancing     = new Dictionary<string,  int>();
            Dictionary<string,  int>   dStockMarriage = new Dictionary<string,  int>();
            Dictionary<string, double> dFinanceRate   = new Dictionary<string, double>();
            Dictionary<string, double> dMarriageRate = new Dictionary<string, double>();

            DateTime dtStartDay = dtRattlesnakeBaseDate;
            string strFileNameMargin = string.Format("D:\\Chips\\Margin\\{0:0000}{1:00}{2:00}.json", dtStartDay.Year, dtStartDay.Month, dtStartDay.Day);
            if (File.Exists(strFileNameMargin))
            {
                string json = File.ReadAllText(strFileNameMargin);
                Dictionary<string, object> json_Dictionary = JsonConvert.DeserializeObject<Dictionary<string, object>>(json);
                if (json_Dictionary.ContainsKey("data"))
                {

                    List<List<string>> jsondata_ListList = JsonConvert.DeserializeObject<List<List<string>>>(json_Dictionary["data"].ToString());
                    foreach (List<string> marginelist in jsondata_ListList)
                    {
                        string strno = marginelist.ElementAt(0);
                        string strFin = marginelist.ElementAt(6);
                        string strFinLimit = marginelist.ElementAt(7);
                        string strMar = marginelist.ElementAt(12);
                        string strMarLimit = marginelist.ElementAt(13);

                        int iFin = int.Parse(strFin, NumberStyles.AllowThousands);
                        int iMar = int.Parse(strMar, NumberStyles.AllowThousands);
                        int iFinLimit = int.Parse(strFinLimit, NumberStyles.AllowThousands);
                        int iMarLimit = int.Parse(strMarLimit, NumberStyles.AllowThousands);

                        dFinancing.Add(strno, iFin);
                        dStockMarriage.Add(strno, iMar);
                        dFinanceRate.Add(strno, (double)iFin / (double)iFinLimit);
                        dMarriageRate.Add(strno, (double)iMar / (double)iMarLimit);
                    }
                }


            }
            string strFileNameMarginTpex = string.Format("D:\\Chips\\Margin\\{0:0000}{1:00}{2:00}_tpex.json", dtStartDay.Year, dtStartDay.Month, dtStartDay.Day);
            if (File.Exists(strFileNameMarginTpex))
            {
                string jsonTpex = File.ReadAllText(strFileNameMarginTpex);
                Dictionary<string, object> jsonTpex_Dictionary = JsonConvert.DeserializeObject<Dictionary<string, object>>(jsonTpex);
                if (jsonTpex_Dictionary.ContainsKey("aaData"))
                {
                    List<List<string>> jsonaadata_ListList = JsonConvert.DeserializeObject<List<List<string>>>(jsonTpex_Dictionary["aaData"].ToString());
                    foreach (List<string> marginelist in jsonaadata_ListList)
                    {
                        string strno = marginelist.ElementAt(0);
                        string strFin = marginelist.ElementAt(6);
                        string strMar = marginelist.ElementAt(14);
                        string strFinUsingRate = marginelist.ElementAt(8);
                        string strMarUsingRate = marginelist.ElementAt(16);

                        int iFin = int.Parse(strFin, NumberStyles.AllowThousands);
                        int iMar = int.Parse(strMar, NumberStyles.AllowThousands);
                        double dbFinUsingRate = double.Parse(strFinUsingRate, NumberStyles.Float);
                        double dbMarUsingRate = double.Parse(strMarUsingRate, NumberStyles.Float);

                        dFinancing.Add(strno, iFin);
                        dStockMarriage.Add(strno, iMar);
                        dFinanceRate.Add(strno, dbFinUsingRate);
                        dMarriageRate.Add(strno, dbMarUsingRate);
                    }
                }

            }


            string strBasePath = "D:\\Chips\\eachdaystock\\";
            string strFiletwst = strBasePath + dtRattlesnakeExamday.ToString("yyyyMMdd") + ".json";
            string strFileTpex = strBasePath + dtRattlesnakeExamday.ToString("yyyyMMdd") + "_tpex.json";

            if (File.Exists(strFiletwst) && File.Exists(strFileTpex))
            {
                string json = File.ReadAllText(strFiletwst);
                Dictionary<string, object> json_Dictionary = JsonConvert.DeserializeObject<Dictionary<string, object>>(json);
                if (json_Dictionary.Count > 2)
                {
                    List<List<string>> json_data9 = JsonConvert.DeserializeObject<List<List<string>>>(json_Dictionary["data9"].ToString());
                    foreach (List<string> pricelst in json_data9)
                    {
                        string strstockno = pricelst.ElementAt(0);
                        string strcloseprice = pricelst.ElementAt(8);
                        if (strcloseprice != "--")
                        dbExamDayDic.Add(strstockno,double.Parse(strcloseprice));
                    }
                }
                string jsonTpex = File.ReadAllText(strFileTpex);
                Dictionary<string, object> jsonTpex_Dictionary = JsonConvert.DeserializeObject<Dictionary<string, object>>(jsonTpex);
                List<List<string>> json_aaData = JsonConvert.DeserializeObject<List<List<string>>>(jsonTpex_Dictionary["aaData"].ToString());
                foreach (List<string> pricelst in json_aaData)
                {
                    string strstockno = pricelst.ElementAt(0);
                    string strcloseprice = pricelst.ElementAt(2);
                    if (strcloseprice != "----")
                    dbExamDayDic.Add(strstockno,double.Parse(strcloseprice));
                }
            }

            DateTime dtLastMonth = dtRattlesnakeBaseDate.AddMonths(-1);
            DateTime dtChipStartDay = new DateTime(2020, 10, 25);
            string strRevenueBasePath = "D:\\Chips\\Revenue\\";
            string strTWSEcsv = strRevenueBasePath + string.Format("t21sc03_{0}_{1}.csv", dtLastMonth.Year - 1911, dtLastMonth.Month);
            string strTPEXcsv = strRevenueBasePath + string.Format("t21sc03_{0}_{1}_otc.csv", dtLastMonth.Year - 1911, dtLastMonth.Month);
            // 放棄找KY的公司 string strTPEXFor = strRevenueBasePath + string.Format("OTC_{0}_FOR.csv", dtLastMonth.ToString("yyyyMM"));

            Dictionary<string, RevenueClass> revenueMap = new Dictionary<string, RevenueClass>();
            if (File.Exists(strTWSEcsv))
            {
                IEnumerable<string> lines = File.ReadLines(strTWSEcsv);
                foreach (string line in lines)  // 取營收
                {
                    RevenueClass rc = new RevenueClass();
                    rc.AssignTWSE(line);
                    if (rc.number != null)
                        revenueMap.Add(rc.number, rc);
                }
            }
            if (File.Exists(strTPEXcsv))
            {
                IEnumerable<string> lines = File.ReadLines(strTPEXcsv, Encoding.UTF8);
                foreach (string line in lines)  // 取營收
                {
                    RevenueClass rc = new RevenueClass();
                    rc.AssignTWSE(line);
                    if (rc.number != null)
                        revenueMap.Add(rc.number, rc);
                }
            }

            List<Dictionary<string, StockLevelCount>> weekkeveklist = new List<Dictionary<string, StockLevelCount>>();
            string strHolderBasePath = "D:\\Chips\\stock\\";
            //DateTime dtLoopstrt = dtRattlesnakeBaseDate;
            DateTime dtLoopstrt = dtChipStartDay;
            for (int iholder = 0; iholder < iCheckForWeeksHolder; )
            {
              
                string strFile = strHolderBasePath + "TDCC_OD_1-5_" +dtLoopstrt.ToString("yyyyMMdd") + ".csv";
                dtLoopstrt = dtLoopstrt.AddDays(-1);
                 if (File.Exists(strFile))
                 {
                     iholder++;
                     int iReadCount = 0;
                     Dictionary<string, StockLevelCount> weekLevelCount = new Dictionary<string, StockLevelCount>();
                     IEnumerable<string> lines = File.ReadLines(strFile);
                     foreach (string line in lines)  // 取流通在外數
                     {
                         if (iReadCount > 0)
                         {
                             string[] strSplitLine = line.Split(',');
                             if (weekLevelCount.ContainsKey(strSplitLine[1]))
                             {
                                 int iLevel = int.Parse(strSplitLine[2]) - 1;
                                 weekLevelCount[strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                                 weekLevelCount[strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                                 weekLevelCount[strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                             }
                             else
                             {
                                 weekLevelCount.Add(strSplitLine[1], new StockLevelCount());
                                 int iLevel = int.Parse(strSplitLine[2]) - 1;
                                 weekLevelCount[strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                                 weekLevelCount[strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                                 weekLevelCount[strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                             }
                         }
                         iReadCount++;
                     }
                     weekkeveklist.Add(weekLevelCount);
                 }
            }

            string strwritefilename = "RattlesnakeDayChip.csv";
            string strTitle = "no,pricechangepercent,Revenuegrow,Below60times,Above60times,for60buy,trust60buy,self60buy,for30buy,trust30buy,self30buy,buysamedays,sellsamedays,FTBDays,TSBDays,SFBDays,FTSDays,TSSDays,SFSDays,P0RateOf0_1,P1RateOf0_1,P2-5RateOf0_1,P6-9RateOf0_1,P10-16RateOf0_1,H0RateOf0_1,H1RateOf0_1,H2-5RateOf0_1,H6-9RateOf0_1,H10-16RateOf0_1,P0RateOf1_2,P1RateOf1_2,P2-5RateOf1_2,P6-9RateOf1_2,P10-16RateOf1_2,H0RateOf1_2,H1RateOf1_2,H2-5RateOf1_2,H6-9RateOf1_2,H10-16RateOf1_2,P0RateO41_4,P1RateOf4_4,P2-5RateOf4_4,P6-9RateOf4_4,P10-16RateOf4_4,H0RateOf4_4,H1RateOf4_4,H2-5RateOf4_4,H6-9RateOf4_4,H10-16RateOf4_4,Fin,Mar,Finrate,Marrate\r\n";
            File.AppendAllText(strwritefilename, strTitle, Encoding.UTF8);
            foreach (string strno in dbStockCloseDic.Keys)
            {
                if (dbStockCloseDic[strno].Count > iRattlesnakeNDays + iLowNDays)
                {
                    double dbAvgAll = dbStockCloseDic[strno].Average();// .Sum / dbStockCloseDic[strno].Count;
                    double dbAvg60 = dbStockCloseDic[strno].Take(iRattlesnakeNDays).Average();
                    double dbAvg60_1 = dbStockCloseDic[strno].Skip(1).Take(iRattlesnakeNDays).Average();
                    if (dbAvg60 > dbAvgAll && dbAvg60 > dbAvg60_1)  // 60日多頭排列
                    {
                        int iCountHiDays = 0;
                        int iCountLoDays = 0;
                        double dbHiof60 = dbStockHighDic[strno].ElementAt(0);
                        double dbLoof60 = dbStockLowDic[strno].ElementAt(0);
                        double dbclsoe = dbStockCloseDic[strno].ElementAt(0);
                        for (int iCnt60 = 0; iCnt60 < iRattlesnakeNDays; iCnt60++)
                        {
                            if (dbHiof60 < dbStockHighDic[strno].ElementAt(iCnt60))
                            {
                                dbHiof60 = dbStockHighDic[strno].ElementAt(iCnt60);
                                iCountHiDays = iCnt60;
                            }
                            if (dbLoof60 > dbStockLowDic[strno].ElementAt(iCnt60))
                            {
                                dbLoof60 = dbStockLowDic[strno].ElementAt(iCnt60);
                                iCountLoDays = iCnt60;
                            }
                        }
                        if (dbclsoe <= dbHiof60 * 0.98 && dbclsoe >= dbLoof60 * 1.01 )
                        {

                            if (dbForTotalBuy.ContainsKey(strno) && dbForTotalBuy[strno].Count >= 40)
                            {
                                int iFor60Buy = dbForTotalBuy[strno].Take(iRattlesnakeNDays).Sum();
                                int iFor30Buy = dbForTotalBuy[strno].Take(iRattlesnakeNDays / 2).Sum();

                                int itrust60Buy = dbTrustTotalBuy[strno].Take(iRattlesnakeNDays).Sum();
                                int itrust30Buy = dbTrustTotalBuy[strno].Take(iRattlesnakeNDays / 2).Sum();

                                int iself60Buy = dbSelfTotalBuy[strno].Take(iRattlesnakeNDays).Sum();
                                int iself30Buy = dbSelfTotalBuy[strno].Take(iRattlesnakeNDays / 2).Sum();

                                int iCountLowBelowAvg60 = 0;  //最近N日最低點比季線小
                                int iCountHiAboveAvg60 = 0;  //最近N日最高點比季線高
                                int i3PtyBuySameDay = 0;   //3 大同買
                                int i3PtySellSameDay = 0;  // 3 大同賣
                                
                                int iForTrustSameDay = 0; //外資投信同買
                                int iTrustSelfSameDay = 0;//投信自營同買
                                int iSelfForSameDay = 0;  //自營外資同買

                                int iForTrustSameSellDay = 0; //外資投信同賣
                                int iTrustSelfSameSellDay = 0;//投信自營同賣
                                int iSelfForSameSellDay = 0;  //自營外資同賣

                                for (int iCountloday = 0; iCountloday < iLowNDays; iCountloday++)
                                {
                                    double dbAvg60forlow = dbStockCloseDic[strno].Skip(iCountloday).Take(iRattlesnakeNDays).Average();
                                    double dblowaday = dbStockLowDic[strno].ElementAt(iCountloday);
                                    double dbhiaday = dbStockHighDic[strno].ElementAt(iCountloday);
                                    if (dblowaday <= dbAvg60forlow)
                                    {
                                        iCountLowBelowAvg60++;
                                    }
                                    if (dbhiaday >= dbAvg60forlow)
                                    {
                                        iCountHiAboveAvg60++;
                                    }
                                    if (dbForTotalBuy[strno].ElementAt(iCountloday) > 0 && dbTrustTotalBuy[strno].ElementAt(iCountloday) > 0 && dbSelfTotalBuy[strno].ElementAt(iCountloday) > 0)
                                    {
                                        i3PtyBuySameDay++;
                                    }
                                    if (dbForTotalBuy[strno].ElementAt(iCountloday) < 0 && dbTrustTotalBuy[strno].ElementAt(iCountloday) < 0 && dbSelfTotalBuy[strno].ElementAt(iCountloday) < 0)
                                    {
                                        i3PtySellSameDay++;
                                    }
                                    if (dbForTotalBuy[strno].ElementAt(iCountloday) > 0 && dbTrustTotalBuy[strno].ElementAt(iCountloday) > 0)
                                    {
                                        iForTrustSameDay++;
                                    }
                                    if (dbForTotalBuy[strno].ElementAt(iCountloday) > 0 && dbSelfTotalBuy[strno].ElementAt(iCountloday) > 0)
                                    {
                                        iSelfForSameDay++;
                                    }
                                    if (dbTrustTotalBuy[strno].ElementAt(iCountloday) > 0 && dbSelfTotalBuy[strno].ElementAt(iCountloday) > 0)
                                    {
                                        iTrustSelfSameDay++;
                                    }

                                    if (dbForTotalBuy[strno].ElementAt(iCountloday) < 0 && dbTrustTotalBuy[strno].ElementAt(iCountloday) < 0)
                                    {
                                        iForTrustSameSellDay++;
                                    }
                                    if (dbForTotalBuy[strno].ElementAt(iCountloday)< 0 && dbSelfTotalBuy[strno].ElementAt(iCountloday) <0)
                                    {
                                        iSelfForSameSellDay++;
                                    }
                                    if (dbTrustTotalBuy[strno].ElementAt(iCountloday) < 0 && dbSelfTotalBuy[strno].ElementAt(iCountloday) < 0)
                                    {
                                        iTrustSelfSameSellDay++;
                                    }
                                }

                                if (revenueMap.ContainsKey(strno) && weekkeveklist.ElementAt(0).ContainsKey(strno) && weekkeveklist.ElementAt(2).ContainsKey(strno) && weekkeveklist.ElementAt(4).ContainsKey(strno) && dbExamDayDic.ContainsKey(strno) && dFinancing.ContainsKey(strno))
                                {
                                    double dbperP0_0 = (double)weekkeveklist.ElementAt(0)[strno].LevelPeople[0] / (double)weekkeveklist.ElementAt(0)[strno].LevelPeople[16];
                                    double dbperP1_0 = (double)weekkeveklist.ElementAt(0)[strno].LevelPeople[1] / (double)weekkeveklist.ElementAt(0)[strno].LevelPeople[16];
                                    double dbperP2_0 = ((double)weekkeveklist.ElementAt(0)[strno].LevelPeople[2] + (double)weekkeveklist.ElementAt(0)[strno].LevelPeople[3] + (double)weekkeveklist.ElementAt(0)[strno].LevelPeople[4] + (double)weekkeveklist.ElementAt(0)[strno].LevelPeople[5]) / (double)weekkeveklist.ElementAt(0)[strno].LevelPeople[16];
                                    double dbperP6_0 = ((double)weekkeveklist.ElementAt(0)[strno].LevelPeople[6] + (double)weekkeveklist.ElementAt(0)[strno].LevelPeople[7] + (double)weekkeveklist.ElementAt(0)[strno].LevelPeople[8] + (double)weekkeveklist.ElementAt(0)[strno].LevelPeople[9]) / (double)weekkeveklist.ElementAt(0)[strno].LevelPeople[16];
                                    double dbperP10_0 = ((double)weekkeveklist.ElementAt(0)[strno].LevelPeople[10] + (double)weekkeveklist.ElementAt(0)[strno].LevelPeople[11] + (double)weekkeveklist.ElementAt(0)[strno].LevelPeople[12] + (double)weekkeveklist.ElementAt(0)[strno].LevelPeople[13] + (double)weekkeveklist.ElementAt(0)[strno].LevelPeople[14] + (double)weekkeveklist.ElementAt(0)[strno].LevelPeople[15]) / (double)weekkeveklist.ElementAt(0)[strno].LevelPeople[16];
                                    double dbperH0_0 = weekkeveklist.ElementAt(0)[strno].LevelRate[0];
                                    double dbperH1_0 = weekkeveklist.ElementAt(0)[strno].LevelRate[1];
                                    double dbperH2_0 = weekkeveklist.ElementAt(0)[strno].LevelRate[2] + weekkeveklist.ElementAt(0)[strno].LevelRate[3] + weekkeveklist.ElementAt(0)[strno].LevelRate[4] + weekkeveklist.ElementAt(0)[strno].LevelRate[5];
                                    double dbperH6_0 = weekkeveklist.ElementAt(0)[strno].LevelRate[6] + weekkeveklist.ElementAt(0)[strno].LevelRate[7] + weekkeveklist.ElementAt(0)[strno].LevelRate[8] + weekkeveklist.ElementAt(0)[strno].LevelRate[9];
                                    double dbperH10_0 = weekkeveklist.ElementAt(0)[strno].LevelRate[10] + weekkeveklist.ElementAt(0)[strno].LevelRate[11] + weekkeveklist.ElementAt(0)[strno].LevelRate[12] + weekkeveklist.ElementAt(0)[strno].LevelRate[13] + weekkeveklist.ElementAt(0)[strno].LevelRate[14] + weekkeveklist.ElementAt(0)[strno].LevelRate[15];

                                    double dbperP0_1 = (double)weekkeveklist.ElementAt(1)[strno].LevelPeople[0] / (double)weekkeveklist.ElementAt(1)[strno].LevelPeople[16];
                                    double dbperP1_1 = (double)weekkeveklist.ElementAt(1)[strno].LevelPeople[1] / (double)weekkeveklist.ElementAt(1)[strno].LevelPeople[16];
                                    double dbperP2_1 = ((double)weekkeveklist.ElementAt(1)[strno].LevelPeople[2] + (double)weekkeveklist.ElementAt(1)[strno].LevelPeople[3] + (double)weekkeveklist.ElementAt(1)[strno].LevelPeople[4] + (double)weekkeveklist.ElementAt(1)[strno].LevelPeople[5]) / (double)weekkeveklist.ElementAt(1)[strno].LevelPeople[16];
                                    double dbperP6_1 = ((double)weekkeveklist.ElementAt(1)[strno].LevelPeople[6] + (double)weekkeveklist.ElementAt(1)[strno].LevelPeople[7] + (double)weekkeveklist.ElementAt(1)[strno].LevelPeople[8] + (double)weekkeveklist.ElementAt(1)[strno].LevelPeople[9]) / (double)weekkeveklist.ElementAt(1)[strno].LevelPeople[16];
                                    double dbperP10_1 = ((double)weekkeveklist.ElementAt(1)[strno].LevelPeople[10] + (double)weekkeveklist.ElementAt(1)[strno].LevelPeople[11] + (double)weekkeveklist.ElementAt(1)[strno].LevelPeople[12] + (double)weekkeveklist.ElementAt(1)[strno].LevelPeople[13] + (double)weekkeveklist.ElementAt(1)[strno].LevelPeople[14] + (double)weekkeveklist.ElementAt(1)[strno].LevelPeople[15]) / (double)weekkeveklist.ElementAt(1)[strno].LevelPeople[16];
                                    double dbperH0_1 = weekkeveklist.ElementAt(1)[strno].LevelRate[0];
                                    double dbperH1_1 = weekkeveklist.ElementAt(1)[strno].LevelRate[1];
                                    double dbperH2_1 = weekkeveklist.ElementAt(1)[strno].LevelRate[2] + weekkeveklist.ElementAt(1)[strno].LevelRate[3] + weekkeveklist.ElementAt(1)[strno].LevelRate[4] + weekkeveklist.ElementAt(1)[strno].LevelRate[5];
                                    double dbperH6_1 = weekkeveklist.ElementAt(1)[strno].LevelRate[6] + weekkeveklist.ElementAt(1)[strno].LevelRate[7] + weekkeveklist.ElementAt(1)[strno].LevelRate[8] + weekkeveklist.ElementAt(1)[strno].LevelRate[9];
                                    double dbperH10_1 = weekkeveklist.ElementAt(1)[strno].LevelRate[10] + weekkeveklist.ElementAt(1)[strno].LevelRate[11] + weekkeveklist.ElementAt(1)[strno].LevelRate[12] + weekkeveklist.ElementAt(1)[strno].LevelRate[13] + weekkeveklist.ElementAt(1)[strno].LevelRate[14] + weekkeveklist.ElementAt(1)[strno].LevelRate[15];

                                    double dbperP0_2 = (double)weekkeveklist.ElementAt(2)[strno].LevelPeople[0] / (double)weekkeveklist.ElementAt(2)[strno].LevelPeople[16];
                                    double dbperP1_2 = (double)weekkeveklist.ElementAt(2)[strno].LevelPeople[1] / (double)weekkeveklist.ElementAt(2)[strno].LevelPeople[16];
                                    double dbperP2_2 = ((double)weekkeveklist.ElementAt(2)[strno].LevelPeople[2] + (double)weekkeveklist.ElementAt(2)[strno].LevelPeople[3] + (double)weekkeveklist.ElementAt(2)[strno].LevelPeople[4] + (double)weekkeveklist.ElementAt(2)[strno].LevelPeople[5]) / (double)weekkeveklist.ElementAt(2)[strno].LevelPeople[16];
                                    double dbperP6_2 = ((double)weekkeveklist.ElementAt(2)[strno].LevelPeople[6] + (double)weekkeveklist.ElementAt(2)[strno].LevelPeople[7] + (double)weekkeveklist.ElementAt(2)[strno].LevelPeople[8] + (double)weekkeveklist.ElementAt(2)[strno].LevelPeople[9]) / (double)weekkeveklist.ElementAt(2)[strno].LevelPeople[16];
                                    double dbperP10_2 = ((double)weekkeveklist.ElementAt(2)[strno].LevelPeople[10] + (double)weekkeveklist.ElementAt(2)[strno].LevelPeople[11] + (double)weekkeveklist.ElementAt(2)[strno].LevelPeople[12] + (double)weekkeveklist.ElementAt(2)[strno].LevelPeople[13] + (double)weekkeveklist.ElementAt(2)[strno].LevelPeople[14] + (double)weekkeveklist.ElementAt(2)[strno].LevelPeople[15]) / (double)weekkeveklist.ElementAt(2)[strno].LevelPeople[16];
                                    double dbperH0_2 = weekkeveklist.ElementAt(2)[strno].LevelRate[0];
                                    double dbperH1_2 = weekkeveklist.ElementAt(2)[strno].LevelRate[1];
                                    double dbperH2_2 = weekkeveklist.ElementAt(2)[strno].LevelRate[2] + weekkeveklist.ElementAt(2)[strno].LevelRate[3] + weekkeveklist.ElementAt(2)[strno].LevelRate[4] + weekkeveklist.ElementAt(2)[strno].LevelRate[5];
                                    double dbperH6_2 = weekkeveklist.ElementAt(2)[strno].LevelRate[6] + weekkeveklist.ElementAt(2)[strno].LevelRate[7] + weekkeveklist.ElementAt(2)[strno].LevelRate[8] + weekkeveklist.ElementAt(2)[strno].LevelRate[9];
                                    double dbperH10_2 = weekkeveklist.ElementAt(2)[strno].LevelRate[10] + weekkeveklist.ElementAt(2)[strno].LevelRate[11] + weekkeveklist.ElementAt(2)[strno].LevelRate[12] + weekkeveklist.ElementAt(2)[strno].LevelRate[13] + weekkeveklist.ElementAt(2)[strno].LevelRate[14] + weekkeveklist.ElementAt(2)[strno].LevelRate[15];

                                    double dbperP0_4 = (double)weekkeveklist.ElementAt(4)[strno].LevelPeople[0] / (double)weekkeveklist.ElementAt(4)[strno].LevelPeople[16];
                                    double dbperP1_4 = (double)weekkeveklist.ElementAt(4)[strno].LevelPeople[1] / (double)weekkeveklist.ElementAt(4)[strno].LevelPeople[16];
                                    double dbperP2_4 = ((double)weekkeveklist.ElementAt(4)[strno].LevelPeople[2] + (double)weekkeveklist.ElementAt(4)[strno].LevelPeople[3] + (double)weekkeveklist.ElementAt(4)[strno].LevelPeople[4] + (double)weekkeveklist.ElementAt(4)[strno].LevelPeople[5]) / (double)weekkeveklist.ElementAt(4)[strno].LevelPeople[16];
                                    double dbperP6_4 = ((double)weekkeveklist.ElementAt(4)[strno].LevelPeople[6] + (double)weekkeveklist.ElementAt(4)[strno].LevelPeople[7] + (double)weekkeveklist.ElementAt(4)[strno].LevelPeople[8] + (double)weekkeveklist.ElementAt(4)[strno].LevelPeople[9]) / (double)weekkeveklist.ElementAt(4)[strno].LevelPeople[16];
                                    double dbperP10_4 = ((double)weekkeveklist.ElementAt(4)[strno].LevelPeople[10] + (double)weekkeveklist.ElementAt(4)[strno].LevelPeople[11] + (double)weekkeveklist.ElementAt(4)[strno].LevelPeople[12] + (double)weekkeveklist.ElementAt(4)[strno].LevelPeople[13] + (double)weekkeveklist.ElementAt(4)[strno].LevelPeople[14] + (double)weekkeveklist.ElementAt(4)[strno].LevelPeople[15]) / (double)weekkeveklist.ElementAt(4)[strno].LevelPeople[16];
                                    double dbperH0_4 = weekkeveklist.ElementAt(4)[strno].LevelRate[0];
                                    double dbperH1_4 = weekkeveklist.ElementAt(4)[strno].LevelRate[1];
                                    double dbperH2_4 = weekkeveklist.ElementAt(4)[strno].LevelRate[2] + weekkeveklist.ElementAt(4)[strno].LevelRate[3] + weekkeveklist.ElementAt(4)[strno].LevelRate[4] + weekkeveklist.ElementAt(4)[strno].LevelRate[5];
                                    double dbperH6_4 = weekkeveklist.ElementAt(4)[strno].LevelRate[6] + weekkeveklist.ElementAt(4)[strno].LevelRate[7] + weekkeveklist.ElementAt(4)[strno].LevelRate[8] + weekkeveklist.ElementAt(4)[strno].LevelRate[9];
                                    double dbperH10_4 = weekkeveklist.ElementAt(4)[strno].LevelRate[10] + weekkeveklist.ElementAt(4)[strno].LevelRate[11] + weekkeveklist.ElementAt(4)[strno].LevelRate[12] + weekkeveklist.ElementAt(4)[strno].LevelRate[13] + weekkeveklist.ElementAt(4)[strno].LevelRate[14] + weekkeveklist.ElementAt(4)[strno].LevelRate[15];

                                    double dbP0_0_1 = dbperP0_0  -dbperP0_1 ;
                                    double dbP1_0_1 = dbperP1_0  -dbperP1_1 ;
                                    double dbP2_0_1 = dbperP2_0  -dbperP2_1 ;
                                    double dbP6_0_1 = dbperP6_0  -dbperP6_1 ;
                                    double dbP10_0_1 = dbperP10_0 -dbperP10_1;
                                    double dbH0_0_1 = dbperH0_0  -dbperH0_1 ;
                                    double dbH1_0_1 = dbperH1_0  -dbperH1_1 ;
                                    double dbH2_0_1 = dbperH2_0  -dbperH2_1 ;
                                    double dbH6_0_1 = dbperH6_0  -dbperH6_1 ;
                                    double dbH10_0_1 = dbperH10_0 - dbperH10_1;

                                    double dbP0_0_2 = dbperP0_0 - dbperP0_2;
                                    double dbP1_0_2 = dbperP1_0 - dbperP1_2;
                                    double dbP2_0_2 = dbperP2_0 - dbperP2_2;
                                    double dbP6_0_2 = dbperP6_0 - dbperP6_2;
                                    double dbP10_0_2 = dbperP10_0 - dbperP10_2;
                                    double dbH0_0_2 = dbperH0_0 - dbperH0_2;
                                    double dbH1_0_2 = dbperH1_0 - dbperH1_2;
                                    double dbH2_0_2 = dbperH2_0 - dbperH2_2;
                                    double dbH6_0_2 = dbperH6_0 - dbperH6_2;
                                    double dbH10_0_2 = dbperH10_0 - dbperH10_2;

                                    double dbP0_0_4 = dbperP0_0 - dbperP0_4;
                                    double dbP1_0_4 = dbperP1_0 - dbperP1_4;
                                    double dbP2_0_4 = dbperP2_0 - dbperP2_4;
                                    double dbP6_0_4 = dbperP6_0 - dbperP6_4;
                                    double dbP10_0_4 = dbperP10_0 - dbperP10_4;
                                    double dbH0_0_4 = dbperH0_0 - dbperH0_4;
                                    double dbH1_0_4 = dbperH1_0 - dbperH1_4;
                                    double dbH2_0_4 = dbperH2_0 - dbperH2_4;
                                    double dbH6_0_4 = dbperH6_0 - dbperH6_4;
                                    double dbH10_0_4 = dbperH10_0 - dbperH10_4;

                                    double dbraisepercent = (dbExamDayDic[strno] - dbStockCloseDic[strno].ElementAt(0)) / dbStockCloseDic[strno].ElementAt(0);
                                    if (dtRattlesnakeExamday < dtRattlesnakeBaseDate)
                                        dbraisepercent = 0;
                                    double dbPercentiFor60Buy = (double)iFor60Buy / (double)weekkeveklist.ElementAt(0)[strno].LevelVol[16];
                                    double dbPercentitrust60Buy = (double)itrust60Buy / (double)weekkeveklist.ElementAt(0)[strno].LevelVol[16];
                                    double dbPercentiself60Buy = (double)iself60Buy / (double)weekkeveklist.ElementAt(0)[strno].LevelVol[16];
                                    double dbPercentiFor30Buy = (double)iFor30Buy / (double)weekkeveklist.ElementAt(0)[strno].LevelVol[16];
                                    double dbPercentitrust30Buy = (double)itrust30Buy / (double)weekkeveklist.ElementAt(0)[strno].LevelVol[16];
                                    double dbPercentiself30Buy = (double)iself30Buy / (double)weekkeveklist.ElementAt(0)[strno].LevelVol[16];

                                    string strApp1 = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18},"
                                        , strno
                                        , dbraisepercent
                                        , revenueMap[strno].acccomparelastyear
                                        , iCountLowBelowAvg60
                                        , iCountHiAboveAvg60
                                        , dbPercentiFor60Buy
                                        , dbPercentitrust60Buy
                                        , dbPercentiself60Buy
                                        , dbPercentiFor30Buy
                                        , dbPercentitrust30Buy
                                        , dbPercentiself30Buy
                                        , i3PtyBuySameDay
                                        , i3PtySellSameDay
                                        , iForTrustSameDay
                                        , iTrustSelfSameDay
                                        , iSelfForSameDay
                                        , iForTrustSameSellDay
                                        , iTrustSelfSameSellDay
                                        , iSelfForSameSellDay
                                        );


                                    string strApp2 = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},"
                                        , dbP0_0_1
                                        , dbP1_0_1
                                        , dbP2_0_1
                                        , dbP6_0_1
                                        , dbP10_0_1
                                        , dbH0_0_1
                                        , dbH1_0_1
                                        , dbH2_0_1
                                        , dbH6_0_1
                                        , dbH10_0_1
                                        );
                                    string strApp3 = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},"
                                        , dbP0_0_2
                                        , dbP1_0_2
                                        , dbP2_0_2
                                        , dbP6_0_2
                                        , dbP10_0_2
                                        , dbH0_0_2
                                        , dbH1_0_2
                                        , dbH2_0_2
                                        , dbH6_0_2
                                        , dbH10_0_2
                                       );
                                    string strApp4 = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},"
                                        , dbP0_0_4
                                        , dbP1_0_4
                                        , dbP2_0_4
                                        , dbP6_0_4
                                        , dbP10_0_4
                                        , dbH0_0_4
                                        , dbH1_0_4
                                        , dbH2_0_4
                                        , dbH6_0_4
                                        , dbH10_0_4
                                        );

                                    
                                    string strApp5 = string.Format("{0},{1},{2},{3}\r\n"
                                            , (double)(dFinancing[strno]*1000) / (double)weekkeveklist.ElementAt(0)[strno].LevelVol[16]
                                            , (double)(dStockMarriage[strno]*1000) / (double)weekkeveklist.ElementAt(0)[strno].LevelVol[16]
                                            , dFinanceRate[strno]
                                            , dMarriageRate[strno]
                                            );


                                    File.AppendAllText(strwritefilename, strApp1 + strApp2 + strApp3 + strApp4 + strApp5, Encoding.UTF8);
                                }
                            }
                        }
                    }
                }
            }    
        }

        private void button16_Click(object sender, EventArgs e)
        {
            CandidateParty3Stock.Clear();
            DateTime dtStart = new DateTime(DateTime.Now.Year,DateTime.Now.Month,DateTime.Now.Day);
            dtStart = monthCalendar1.SelectionStart;
            int iNDays = 60;
            double dbOver = 0.02;
            string strMsg = string.Format("找比{0}前{1}天低的股，但3大買過{2:F4}:3大布局股", dtStart.ToString("MM/dd"), iNDays, dbOver);
            Console.WriteLine(strMsg);

            int icountday = iNDays;
            DateTime dtLoopStart = dtStart;
            
            while(icountday!=0)
            {
                if(workingdays.Contains(dtLoopStart))
                {
                    icountday--;
                }
                dtLoopStart = dtLoopStart.AddDays(-1);
            }
            DateTime dtEnd = dtLoopStart;

            string strBasePath = "D:\\Chips\\stock\\TDCC_OD_1-5_";
            string strReadFile = "";
            int icountchip = 0;
            while (true)
            {
                string strcsv = strBasePath + dtLoopStart.ToString("yyyyMMdd") + ".csv";
                if (File.Exists(strcsv))
                {
                    strReadFile = strcsv;
                    break;
                }
                dtLoopStart = dtLoopStart.AddDays(-1);
            }
            Dictionary<string, StockLevelCount> weekLevelCount = new Dictionary<string, StockLevelCount>();
            int iReadCount = 0;
            IEnumerable<string> lines = File.ReadLines(strReadFile);
            foreach (string line in lines)  // 取流通在外數
            {
                if (iReadCount > 0)
                {
                    string[] strSplitLine = line.Split(',');
                    if (weekLevelCount.ContainsKey(strSplitLine[1]))
                    {
                        int iLevel = int.Parse(strSplitLine[2]) - 1;
                        weekLevelCount[strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                        weekLevelCount[strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                        weekLevelCount[strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                    }
                    else
                    {
                        weekLevelCount.Add(strSplitLine[1], new StockLevelCount());
                        int iLevel = int.Parse(strSplitLine[2]) - 1;
                        weekLevelCount[strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                        weekLevelCount[strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                        weekLevelCount[strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                    }
                }
                iReadCount++;
            }
            string strmsg = string.Format("號,全,不含避,外,信,避,價減率");
            Console.WriteLine(strmsg);
            HashSet<string> Candidate = new HashSet<string>();
            foreach (string strno in dbEvenydayOLHCDic.Keys)
            {
                Dictionary<double, DateTime> PriceMap = new Dictionary<double, DateTime>();
                List<double> dbCloseList = new List<double>();  // 近N日的CLOSE
                for (DateTime dti = dtEnd; dti <= dtStart; dti = dti.AddDays(1))
                {
                    if (dbEvenydayOLHCDic[strno].ContainsKey(dti))
                    {
                        dbCloseList.Add(dbEvenydayOLHCDic[strno][dti][3]);
                        if (!PriceMap.ContainsKey(dbEvenydayOLHCDic[strno][dti][3]))
                            PriceMap.Add(dbEvenydayOLHCDic[strno][dti][3], dti);
                    }
                }
                if(dbCloseList.Count >= iNDays)
                {
                    double dbmax = dbCloseList.Max();
                    int idex = dbCloseList.LastIndexOf(dbmax);
                    if (dbEvenydayOLHCDic[strno].ContainsKey(dtStart) && dbmax > dbEvenydayOLHCDic[strno][dtStart][3])
                    {
                        if (PriceMap.Count > 0)
                        {
                            DateTime dtHightCloseDay = PriceMap[dbCloseList.Max()];
                            double lossrate = (dbCloseList.Max() - dbCloseList.ElementAt(0)) / dbCloseList.Max();
                            DateTime dtloop = dtStart;
                            long l3PTotal = 0;
                            long lHedgeTotal = 0;
                            long lTrust = 0;
                            long lSelf = 0;
                            long lTotal = 0;
                            long lFor = 0;

                            if (dicStock3Party.ContainsKey(strno) && weekLevelCount.ContainsKey(strno))
                            {
                                while (dtloop > dtHightCloseDay)
                                {
                                    if (dicStock3Party[strno].ContainsKey(dtloop))
                                    {
                                        int iTotal = int.Parse(dicStock3Party[strno][dtloop].Party3Total, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                        int iHedge = int.Parse(dicStock3Party[strno][dtloop].SelfHedgeTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                        int itruet = int.Parse(dicStock3Party[strno][dtloop].TrustTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                        int iself = int.Parse(dicStock3Party[strno][dtloop].SelfSelfTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                        int ifor = int.Parse(dicStock3Party[strno][dtloop].ForeigenTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);

                                        lTotal += (iTotal - iHedge);
                                        l3PTotal += iTotal;
                                        lHedgeTotal += iHedge;
                                        lTrust += itruet;
                                        lSelf += iself;
                                        lFor += ifor;
                                    }
                                    dtloop = dtloop.AddDays(-1);
                                }

                                double dbPercent3PTotal = (double)l3PTotal / (double)weekLevelCount[strno].LevelVol[16]; // 全
                                double dbPercentTotal = (double)lTotal / (double)weekLevelCount[strno].LevelVol[16]; //不含避險
                                double dbPercentlFor= (double)lFor / (double)weekLevelCount[strno].LevelVol[16]; //外
                                double dbPercentlTrust= (double)lTrust / (double)weekLevelCount[strno].LevelVol[16]; //信
                                double dbPercentHedgeTotal = (double)lHedgeTotal / (double)weekLevelCount[strno].LevelVol[16]; //避險

                                if (dbPercentTotal > dbOver && lTrust>0)
                                {
                                    CandidateParty3Stock.Add(strno);
                                    strmsg = string.Format("{0},{1:F4},{2:F4},{3:F4},{4:F4},{5:F4},{6:F4},{7}", strno, dbPercent3PTotal, dbPercentTotal, dbPercentlFor, dbPercentlTrust, dbPercentHedgeTotal, lossrate, dicStockNoMapping[strno]);
                                    Console.WriteLine(strmsg);
                                }
                            }
                        }
                    }
                }
            }
        }

        private void button17_Click(object sender, EventArgs e)
        {
            int iLowRange = 3;  //維持很少的天數
            double dbPercent = 0.001; // 成長數大於總股數

            CandidateParty3Stock.Clear();
            DateTime dtStart = monthCalendar1.SelectionStart;
            DateTime dtLoopStart = monthCalendar1.SelectionStart;
            //CandidateWeekChipSet.Clear();
            string strBasePath = "D:\\Chips\\stock\\TDCC_OD_1-5_";
            string strReadFile = "";
            int icountchip = 0;
            while (true)
            {
                string strcsv = strBasePath + dtLoopStart.ToString("yyyyMMdd") + ".csv";
                if (File.Exists(strcsv))
                {
                    strReadFile = strcsv;
                    break;
                }
                dtLoopStart = dtLoopStart.AddDays(-1);
            }
            Dictionary<string, StockLevelCount> weekLevelCount = new Dictionary<string, StockLevelCount>();
            int iReadCount = 0;
            IEnumerable<string> lines = File.ReadLines(strReadFile);
            foreach (string line in lines)  // 取流通在外數
            {
                if (iReadCount > 0)
                {
                    string[] strSplitLine = line.Split(',');
                    if (weekLevelCount.ContainsKey(strSplitLine[1]))
                    {
                        int iLevel = int.Parse(strSplitLine[2]) - 1;
                        weekLevelCount[strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                        weekLevelCount[strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                        weekLevelCount[strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                    }
                    else
                    {
                        weekLevelCount.Add(strSplitLine[1], new StockLevelCount());
                        int iLevel = int.Parse(strSplitLine[2]) - 1;
                        weekLevelCount[strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                        weekLevelCount[strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                        weekLevelCount[strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                    }
                }
                iReadCount++;
            }

            Dictionary<string, long> stock3PartyBustVol = new Dictionary<string, long>();
            Dictionary<string, long> stock3PartyVol = new Dictionary<string, long>();
            Dictionary<string, double> stock3PartyBust = new Dictionary<string, double>();
            Dictionary<string, double> stock3PartyNormal = new Dictionary<string, double>();
            foreach (string strStockNo in weekLevelCount.Keys)
            {
                dtLoopStart = dtStart;
                if (dicStock3Party.ContainsKey(strStockNo))
                {
                    if (dicStock3Party[strStockNo].ContainsKey(dtLoopStart))
                    {
                        int iHedge = int.Parse(dicStock3Party[strStockNo][dtLoopStart].SelfHedgeTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                        stock3PartyBustVol.Add(strStockNo, iHedge);
                        double dbper = (double)iHedge / (double)weekLevelCount[strStockNo].LevelVol[16];

                        dtLoopStart = dtLoopStart.AddDays(-1);
                        int iLoopDay = 0;
                        long ladd = 0;
                        while (iLoopDay < iLowRange)  // 計算最近幾天
                        {
                            if (workingdays.Contains(dtLoopStart))
                            {
                                iHedge = 0;
                                if (dicStock3Party[strStockNo].ContainsKey(dtLoopStart))
                                    iHedge = int.Parse(dicStock3Party[strStockNo][dtLoopStart].SelfHedgeTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                ladd += iHedge;
                                iLoopDay++;
                            }
                            dtLoopStart = dtLoopStart.AddDays(-1);
                        }
                        stock3PartyBust.Add(strStockNo, dbper);
                        stock3PartyVol.Add(strStockNo, ladd);
                        double dbNormal = (double)ladd / (double)weekLevelCount[strStockNo].LevelVol[16];
                        stock3PartyNormal.Add(strStockNo, dbNormal);
                    }
                }
            }
            foreach (string strStockNo in stock3PartyBust.Keys)
            {
               double dbBust = Math.Abs(stock3PartyBust[strStockNo]);
               double dbNormal = Math.Abs(stock3PartyNormal[strStockNo]);
               
               long lbust = Math.Abs(stock3PartyBustVol[strStockNo]);
               long lnormal = Math.Abs(stock3PartyVol[strStockNo]);

               if (lnormal == 0 && dbBust > 0.001)
               {
                   string strmsg = string.Format("{0},{1:F4},{2:F4},{3:F4},{4:F4}", strStockNo, dbBust, dbNormal, lbust, lnormal);
                   Console.WriteLine(strmsg);
               }
            }
        }

        private void button18_Click(object sender, EventArgs e)
        {
            DateTime dtLatest = dbEvenydayOLHCDic["2330"].Keys.Max();

            DateTime dtLatest_1 = dtLatest.AddDays(-1);
            while (!dbEvenydayOLHCDic["2330"].ContainsKey(dtLatest_1))
            {
                dtLatest_1 = dtLatest_1.AddDays(-1);
            }

            DateTime dtLatest_2 = dtLatest_1.AddDays(-1);
            while (!dbEvenydayOLHCDic["2330"].ContainsKey(dtLatest_1))
            {
                dtLatest_2 = dtLatest_2.AddDays(-1);
            }

            DateTime dtLatest_3 = dtLatest_2.AddDays(-1);
            while (!dbEvenydayOLHCDic["2330"].ContainsKey(dtLatest_1))
            {
                dtLatest_3 = dtLatest_3.AddDays(-1);
            }

            DateTime dtLatest_4 = dtLatest_3.AddDays(-1);
            while (!dbEvenydayOLHCDic["2330"].ContainsKey(dtLatest_1))
            {
                dtLatest_4 = dtLatest_4.AddDays(-1);
            }

            DateTime dtLatest_5 = dtLatest_4.AddDays(-1);
            while (!dbEvenydayOLHCDic["2330"].ContainsKey(dtLatest_1))
            {
                dtLatest_5 = dtLatest_5.AddDays(-1);
            }
            foreach(string strno in dbEvenydayOLHCDic.Keys)
            {
                double sma20 = SMA(dbEvenydayOLHCDic[strno], dtLatest, 20, 3);
                double sma20_1 = SMA(dbEvenydayOLHCDic[strno], dtLatest_1, 20, 3);
                double sma60 = SMA(dbEvenydayOLHCDic[strno], dtLatest, 60, 3);
                double sma60_1 = SMA(dbEvenydayOLHCDic[strno], dtLatest_1, 60, 3);
                double sma120 = SMA(dbEvenydayOLHCDic[strno], dtLatest, 120, 3);
                double sma120_1 = SMA(dbEvenydayOLHCDic[strno], dtLatest_1, 120, 3);
                double sma240 = SMA(dbEvenydayOLHCDic[strno], dtLatest, 240, 3);
                double sma240_1 = SMA(dbEvenydayOLHCDic[strno], dtLatest_1, 240, 3);
                if ( sma60 >= sma60_1 && sma120 >= sma120 && sma240 >= sma240 && sma20 >= sma60 && sma60 >= sma120 && sma120 >= sma240)
                {
                    double avgout = 0, avgout_1 = 0, avgout_2 = 0, avgout_3 = 0, avgout_4 = 0, avgout_5 = 0;
                    double dbStdev = STDev(dbEvenydayOLHCDic[strno], dtLatest, 20, out avgout);
                    double dbStdev_1 = STDev(dbEvenydayOLHCDic[strno], dtLatest_1, 20, out avgout_1);
                    double dbStdev_2 = STDev(dbEvenydayOLHCDic[strno], dtLatest_2, 20, out avgout_2);
                    double dbStdev_3 = STDev(dbEvenydayOLHCDic[strno], dtLatest_3, 20, out avgout_3);
                    double dbStdev_4 = STDev(dbEvenydayOLHCDic[strno], dtLatest_4, 20, out avgout_4);
                    double dbStdev_5 = STDev(dbEvenydayOLHCDic[strno], dtLatest_5, 20, out avgout_5);

                    if (dbStdev_5 >= dbStdev_4 &&dbStdev_4 >= dbStdev_3 && dbStdev_3 >= dbStdev_2 && dbStdev_2 >= dbStdev_1 && dbStdev_1 >= dbStdev)
                    {
                        string strmsg = string.Format("{0}", strno);

                        Console.WriteLine(strmsg);
                    }

                }
            }
        }

        private void button19_Click(object sender, EventArgs e)
        {
            CandidateParty3Stock.Clear();
            DateTime dtStart = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);
            dtStart = monthCalendar1.SelectionStart;
            int iNDays_Long = 180;
            int iNDays_Mid = 60;
            int iNDays_Short = 20;
            string strMsg = string.Format("找比{0}前 {1}，{2}，{3} 三角收斂", dtStart.ToString("MM/dd"), iNDays_Long, iNDays_Mid, iNDays_Short);
            Console.WriteLine(strMsg);

            int icountday = iNDays_Long;
            DateTime dtLoopStart = dtStart;

            while (icountday != 0)
            {
                if (workingdays.Contains(dtLoopStart))
                {
                    icountday--;
                }
                dtLoopStart = dtLoopStart.AddDays(-1);
            }
            DateTime dtEnd = dtLoopStart;

            string strBasePath = "D:\\Chips\\stock\\TDCC_OD_1-5_";
            string strReadFile = "";
            int icountchip = 0;
            while (true)
            {
                string strcsv = strBasePath + dtLoopStart.ToString("yyyyMMdd") + ".csv";
                if (File.Exists(strcsv))
                {
                    strReadFile = strcsv;
                    break;
                }
                dtLoopStart = dtLoopStart.AddDays(-1);
            }
            Dictionary<string, StockLevelCount> weekLevelCount = new Dictionary<string, StockLevelCount>();
            int iReadCount = 0;
            IEnumerable<string> lines = File.ReadLines(strReadFile);
            foreach (string line in lines)  // 取流通在外數
            {
                if (iReadCount > 0)
                {
                    string[] strSplitLine = line.Split(',');
                    if (weekLevelCount.ContainsKey(strSplitLine[1]))
                    {
                        int iLevel = int.Parse(strSplitLine[2]) - 1;
                        weekLevelCount[strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                        weekLevelCount[strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                        weekLevelCount[strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                    }
                    else
                    {
                        weekLevelCount.Add(strSplitLine[1], new StockLevelCount());
                        int iLevel = int.Parse(strSplitLine[2]) - 1;
                        weekLevelCount[strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                        weekLevelCount[strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                        weekLevelCount[strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                    }
                }
                iReadCount++;
            }

            HashSet<string> Candidate = new HashSet<string>();
            foreach (string strno in dbEvenydayOLHCDic.Keys)
            {
                if (dbEvenydayOLHCDic[strno].Count > 200 && strno.Length == 4 && dbEvenydayOLHCDic[strno].ContainsKey(dtStart))
                {
                    List<double> dbCloseListShort = new List<double>();
                    List<double> dbCloseListMid = new List<double>();
                    List<double> dbCloseListLong = new List<double>();

                    List<double> dbHighListShort = new List<double>();
                    List<double> dbHighListMid = new List<double>();
                    List<double> dbHighListLong = new List<double>();

                    List<double> dbLowListShort = new List<double>();
                    List<double> dbLowListMid = new List<double>();
                    List<double> dbLowListLong = new List<double>();

                    DateTime dtLoop = dtStart;
                    int iCountAdded = 0;

                    while (iCountAdded < iNDays_Short && dtLoop > dtEnd)
                    {
                        if (dbEvenydayOLHCDic[strno].ContainsKey(dtLoop))
                        {
                            dbCloseListShort.Add(dbEvenydayOLHCDic[strno][dtLoop][3]);
                            dbHighListShort.Add(dbEvenydayOLHCDic[strno][dtLoop][2]);
                            dbLowListShort.Add(dbEvenydayOLHCDic[strno][dtLoop][1]);
                            iCountAdded++;
                        }
                        dtLoop = dtLoop.AddDays(-1);
                    }
                    dtLoop = dtStart;
                    iCountAdded = 0;
                    while (iCountAdded < iNDays_Mid && dtLoop > dtEnd)
                    {
                        if (dbEvenydayOLHCDic[strno].ContainsKey(dtLoop))
                        {
                            dbCloseListMid.Add(dbEvenydayOLHCDic[strno][dtLoop][3]);
                            dbHighListMid.Add(dbEvenydayOLHCDic[strno][dtLoop][2]);
                            dbLowListMid.Add(dbEvenydayOLHCDic[strno][dtLoop][1]);
                            iCountAdded++;
                        }
                        dtLoop = dtLoop.AddDays(-1);
                    }
                    dtLoop = dtStart;
                    iCountAdded = 0;
                    while (iCountAdded < iNDays_Long && dtLoop > dtEnd)
                    {
                        if (dbEvenydayOLHCDic[strno].ContainsKey(dtLoop))
                        {
                            dbCloseListLong.Add(dbEvenydayOLHCDic[strno][dtLoop][3]);
                            dbHighListLong.Add(dbEvenydayOLHCDic[strno][dtLoop][2]);
                            dbLowListLong.Add(dbEvenydayOLHCDic[strno][dtLoop][1]);
                            iCountAdded++;
                        }
                        dtLoop = dtLoop.AddDays(-1);
                    }

                    double dbshotmax = dbHighListShort.Max();
                    double dbmidmax = dbHighListMid.Max();
                    double dblongmax = dbHighListLong.Max();

                    double dbshotmin = dbLowListShort.Min();
                    double dbmidmin = dbLowListMid.Min();
                    double dblongmin = dbLowListLong.Min();


                    if (dbshotmax < dbmidmax && dbmidmax < dblongmax && dbshotmin > dbmidmin && dbmidmin > dblongmin)
                    {

                        long l3PTotal = 0;
                        long lHedgeTotal = 0;
                        long lTrust = 0;
                        long lSelf = 0;
                        long lTotal = 0;
                        long lFor = 0;

                        long ladd = 0;
                        int iLoopDay = 0;
                        dtLoopStart = dtStart;
                        while (iLoopDay < iNDays_Short)  // 計算最近幾天
                        {
                            if (workingdays.Contains(dtLoopStart))
                            {
                                if (dicStock3Party.ContainsKey(strno))
                                {
                                    if (dicStock3Party[strno].ContainsKey(dtLoopStart))
                                    {
                                        int iTotal = int.Parse(dicStock3Party[strno][dtLoopStart].Party3Total, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                        int iHedge = int.Parse(dicStock3Party[strno][dtLoopStart].SelfHedgeTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                        int iHedgeBuy = int.Parse(dicStock3Party[strno][dtLoopStart].SelfHedgeBuy, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                        int iHedgeSell = int.Parse(dicStock3Party[strno][dtLoopStart].SelfHedgeSell, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                        int iForeigen = int.Parse(dicStock3Party[strno][dtLoopStart].ForeigenTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                        int iTrust = int.Parse(dicStock3Party[strno][dtLoopStart].TrustTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                        int iSelf = int.Parse(dicStock3Party[strno][dtLoopStart].SelfSelfTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                        ladd += (iTotal - iHedge); // 計算不含避險3大買量

                                        lTotal += (iTotal - iHedge);
                                        l3PTotal += iTotal;
                                        lHedgeTotal += iHedge;
                                        lTrust += iTrust;
                                        lSelf += iSelf;
                                        lFor += iForeigen;
                                    }
                                }
                                iLoopDay++;
                            }

                            dtLoopStart = dtLoopStart.AddDays(-1);
                        }

                        if(weekLevelCount.ContainsKey(strno))
                        {
                            double dbper = (double)ladd / (double)weekLevelCount[strno].LevelVol[16];

                            if (dbper > 0.03 && lHedgeTotal > 0)
                            {
                                CandidateParty3Stock.Add(strno);
                                string strmsg = string.Format("{0}", strno);
                                Console.WriteLine(strmsg);
                            }
                        }

                    }
                }
                
            }
        }

        private void button20_Click(object sender, EventArgs e)
        {
           
            double dbPercent = 0.007; // 成長數大於總股數
            DateTime dtStart = monthCalendar1.SelectionStart;
            DateTime dtLoopStart = monthCalendar1.SelectionStart;

            string strMsg = string.Format("找{0}投信或外資買超過總股數{1:F4}", dtStart.ToString("MM/dd"),dbPercent);
            Console.WriteLine(strMsg);

            CandidateParty3Stock.Clear();

            //CandidateWeekChipSet.Clear();
            string strBasePath = "D:\\Chips\\stock\\TDCC_OD_1-5_";
            string strReadFile = "";
            int icountchip = 0;
            while (true)
            {
                string strcsv = strBasePath + dtLoopStart.ToString("yyyyMMdd") + ".csv";
                if (File.Exists(strcsv))
                {
                    strReadFile = strcsv;
                    break;
                }
                dtLoopStart = dtLoopStart.AddDays(-1);
            }
            Dictionary<string, StockLevelCount> weekLevelCount = new Dictionary<string, StockLevelCount>();
            int iReadCount = 0;
            IEnumerable<string> lines = File.ReadLines(strReadFile);
            foreach (string line in lines)  // 取流通在外數
            {
                if (iReadCount > 0)
                {
                    string[] strSplitLine = line.Split(',');
                    if (weekLevelCount.ContainsKey(strSplitLine[1]))
                    {
                        int iLevel = int.Parse(strSplitLine[2]) - 1;
                        weekLevelCount[strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                        weekLevelCount[strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                        weekLevelCount[strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                    }
                    else
                    {
                        weekLevelCount.Add(strSplitLine[1], new StockLevelCount());
                        int iLevel = int.Parse(strSplitLine[2]) - 1;
                        weekLevelCount[strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                        weekLevelCount[strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                        weekLevelCount[strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                    }
                }
                iReadCount++;
            }

            Dictionary<string, long> stock3PartyVol = new Dictionary<string, long>();
            foreach (string strStockNo in weekLevelCount.Keys)
            {
                long l3PTotal = 0;
                long lHedgeTotal = 0;
                long lTrust = 0;
                long lSelf = 0;
                long lTotal = 0;
                long lFor = 0;

                long ladd = 0;

                dtLoopStart = dtStart;

                if (workingdays.Contains(dtLoopStart))
                {
                    if (dicStock3Party.ContainsKey(strStockNo))
                    {
                        if (dicStock3Party[strStockNo].ContainsKey(dtLoopStart))
                        {
                            int iTotal = int.Parse(dicStock3Party[strStockNo][dtLoopStart].Party3Total, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                            int iHedge = int.Parse(dicStock3Party[strStockNo][dtLoopStart].SelfHedgeTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                            int iHedgeBuy = int.Parse(dicStock3Party[strStockNo][dtLoopStart].SelfHedgeBuy, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                            int iHedgeSell = int.Parse(dicStock3Party[strStockNo][dtLoopStart].SelfHedgeSell, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                            int iForeigen = int.Parse(dicStock3Party[strStockNo][dtLoopStart].ForeigenTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                            int iTrust = int.Parse(dicStock3Party[strStockNo][dtLoopStart].TrustTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                            int iSelf = int.Parse(dicStock3Party[strStockNo][dtLoopStart].SelfSelfTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                            ladd += (iTotal - iHedge); // 計算不含避險3大買量
                            //ladd += iTotal; // 全算
                            //ladd += (iTrust + iForeigen); //外 信
                            //ladd += iHedge;

                            lTotal += (iTotal - iHedge);
                            l3PTotal += iTotal;
                            lHedgeTotal += iHedge;
                            lTrust += iTrust;
                            lSelf += iSelf;
                            lFor += iForeigen;



                            if (weekLevelCount.ContainsKey(strStockNo) && dbStockVolDic[strStockNo].ContainsKey(dtLoopStart))
                            {
                                double dbtrust = (double)lTrust / (double)weekLevelCount[strStockNo].LevelVol[16];
                                double dbfor = (double)lFor / (double)weekLevelCount[strStockNo].LevelVol[16];
                                double dbHedge = (double)lHedgeTotal / (double)weekLevelCount[strStockNo].LevelVol[16];

                                double dbvolfor = (double)lFor / (double)dbStockVolDic[strStockNo][dtLoopStart];
                                double dbvoltrust = (double)lTrust / (double)dbStockVolDic[strStockNo][dtLoopStart];
                                double dbvolhedge = (double)lHedgeTotal / (double)dbStockVolDic[strStockNo][dtLoopStart];

                                if (dbtrust >= dbPercent || dbtrust <= -dbPercent)
                                {
                                    CandidateParty3Stock.Add(strStockNo);
                                    Console.WriteLine("信 ," + strStockNo + "," + dicStock3Party[strStockNo][dtLoopStart].name + "," + dbtrust + "," + dbEvenydayOLHCDic[strStockNo][dtLoopStart][3] + "," + dbvoltrust);
                                }
                                if (dbfor >= dbPercent || dbfor <= -dbPercent)
                                {
                                    CandidateParty3Stock.Add(strStockNo);
                                    Console.WriteLine("外 ," + strStockNo + "," + dicStock3Party[strStockNo][dtLoopStart].name + "," + dbfor + "," + dbEvenydayOLHCDic[strStockNo][dtLoopStart][3] + "," + dbvolfor);
                                }
                                if (dbHedge >= dbPercent || dbHedge <= -dbPercent)
                                {
                                    CandidateParty3Stock.Add(strStockNo);
                                    Console.WriteLine("避 ," + strStockNo + "," + dicStock3Party[strStockNo][dtLoopStart].name + "," + dbHedge + "," + dbEvenydayOLHCDic[strStockNo][dtLoopStart][3] + "," + dbvolhedge);
                                }
                            }
                        }
                    }

                }





                
            }    
        }

        private void button21_Click(object sender, EventArgs e)
        {
            CandidateParty3Stock.Clear();
            DateTime dtStart = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);
            dtStart = monthCalendar1.SelectionStart;
            int iNDays_Long = 180;
            int iNDays_Mid = 60;
            int iNDays_Short = 20;
            string strMsg = string.Format("找{0}為止 {1}，{2}，{3} W底", dtStart.ToString("MM/dd"), iNDays_Long, iNDays_Mid, iNDays_Short);
            Console.WriteLine(strMsg);

            int icountday = iNDays_Long;
            DateTime dtLoopStart = dtStart;

            while (icountday != 0)
            {
                if (workingdays.Contains(dtLoopStart))
                {
                    icountday--;
                }
                dtLoopStart = dtLoopStart.AddDays(-1);
            }
            DateTime dtEnd = dtLoopStart;

            string strBasePath = "D:\\Chips\\stock\\TDCC_OD_1-5_";
            string strReadFile = "";
            int icountchip = 0;
            while (true)
            {
                string strcsv = strBasePath + dtLoopStart.ToString("yyyyMMdd") + ".csv";
                if (File.Exists(strcsv))
                {
                    strReadFile = strcsv;
                    break;
                }
                dtLoopStart = dtLoopStart.AddDays(-1);
            }
            Dictionary<string, StockLevelCount> weekLevelCount = new Dictionary<string, StockLevelCount>();
            int iReadCount = 0;
            IEnumerable<string> lines = File.ReadLines(strReadFile);
            foreach (string line in lines)  // 取流通在外數
            {
                if (iReadCount > 0)
                {
                    string[] strSplitLine = line.Split(',');
                    if (weekLevelCount.ContainsKey(strSplitLine[1]))
                    {
                        int iLevel = int.Parse(strSplitLine[2]) - 1;
                        weekLevelCount[strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                        weekLevelCount[strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                        weekLevelCount[strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                    }
                    else
                    {
                        weekLevelCount.Add(strSplitLine[1], new StockLevelCount());
                        int iLevel = int.Parse(strSplitLine[2]) - 1;
                        weekLevelCount[strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                        weekLevelCount[strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                        weekLevelCount[strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                    }
                }
                iReadCount++;
            }

            HashSet<string> Candidate = new HashSet<string>();
            foreach (string strno in dbEvenydayOLHCDic.Keys)
            {
                if (dbEvenydayOLHCDic[strno].Count > 200 && strno.Length == 4 && dbEvenydayOLHCDic[strno].ContainsKey(dtStart))
                {
                    List<double> dbCloseListShort = new List<double>();
                    List<double> dbCloseListMid = new List<double>();
                    List<double> dbCloseListLong = new List<double>();

                    List<double> dbHighListShort = new List<double>();
                    List<double> dbHighListMid = new List<double>();
                    List<double> dbHighListLong = new List<double>();

                    List<double> dbLowListShort = new List<double>();
                    List<double> dbLowListMid = new List<double>();
                    List<double> dbLowListLong = new List<double>();

                    DateTime[] dtWorkingArray = new DateTime[iNDays_Long];

                    DateTime dtLoop = dtStart;
                    int iCountAdded = 0;

                    while (iCountAdded < iNDays_Short && dtLoop > dtEnd)
                    {
                        if (dbEvenydayOLHCDic[strno].ContainsKey(dtLoop))
                        {
                            dbCloseListShort.Add(dbEvenydayOLHCDic[strno][dtLoop][3]);
                            dbHighListShort.Add(dbEvenydayOLHCDic[strno][dtLoop][2]);
                            dbLowListShort.Add(dbEvenydayOLHCDic[strno][dtLoop][1]);
                            iCountAdded++;
                        }
                        dtLoop = dtLoop.AddDays(-1);
                    }
                    dtLoop = dtStart;
                    iCountAdded = 0;
                    while (iCountAdded < iNDays_Mid && dtLoop > dtEnd)
                    {
                        if (dbEvenydayOLHCDic[strno].ContainsKey(dtLoop))
                        {
                            dbCloseListMid.Add(dbEvenydayOLHCDic[strno][dtLoop][3]);
                            dbHighListMid.Add(dbEvenydayOLHCDic[strno][dtLoop][2]);
                            dbLowListMid.Add(dbEvenydayOLHCDic[strno][dtLoop][1]);
                            iCountAdded++;
                        }
                        dtLoop = dtLoop.AddDays(-1);
                    }
                    dtLoop = dtStart;
                    iCountAdded = 0;
                    while (iCountAdded < iNDays_Long && dtLoop > dtEnd)
                    {
                        if (dbEvenydayOLHCDic[strno].ContainsKey(dtLoop))
                        {
                            dbCloseListLong.Add(dbEvenydayOLHCDic[strno][dtLoop][3]);
                            dbHighListLong.Add(dbEvenydayOLHCDic[strno][dtLoop][2]);
                            dbLowListLong.Add(dbEvenydayOLHCDic[strno][dtLoop][1]);
                            iCountAdded++;
                        }
                        dtLoop = dtLoop.AddDays(-1);
                    }

                    double dbshortmax = dbHighListShort.Max();
                    double dbmidmax = dbHighListMid.Max();
                    double dblongmax = dbHighListLong.Max();

                    double dbshotmin = dbLowListShort.Min();
                    double dbmidmin = dbLowListMid.Min();
                    double dblongmin = dbLowListLong.Min();


                    //if (dbshotmax < dbmidmax && dbmidmax < dblongmax && dbshotmin > dbmidmin && dbmidmin > dblongmin)
                    if (dbshortmax > dbmidmax && dbmidmax < dblongmax && dbshotmin > dbmidmin && dbmidmin == dblongmin)
                    {

                        long l3PTotal = 0;
                        long lHedgeTotal = 0;
                        long lTrust = 0;
                        long lSelf = 0;
                        long lTotal = 0;
                        long lFor = 0;

                        long ladd = 0;
                        int iLoopDay = 0;
                        dtLoopStart = dtStart;
                        while (iLoopDay < iNDays_Short)  // 計算最近幾天
                        {
                            if (workingdays.Contains(dtLoopStart))
                            {
                                if (dicStock3Party.ContainsKey(strno))
                                {
                                    if (dicStock3Party[strno].ContainsKey(dtLoopStart))
                                    {
                                        int iTotal = int.Parse(dicStock3Party[strno][dtLoopStart].Party3Total, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                        int iHedge = int.Parse(dicStock3Party[strno][dtLoopStart].SelfHedgeTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                        int iHedgeBuy = int.Parse(dicStock3Party[strno][dtLoopStart].SelfHedgeBuy, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                        int iHedgeSell = int.Parse(dicStock3Party[strno][dtLoopStart].SelfHedgeSell, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                        int iForeigen = int.Parse(dicStock3Party[strno][dtLoopStart].ForeigenTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                        int iTrust = int.Parse(dicStock3Party[strno][dtLoopStart].TrustTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                        int iSelf = int.Parse(dicStock3Party[strno][dtLoopStart].SelfSelfTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                        ladd += (iTotal - iHedge); // 計算不含避險3大買量

                                        lTotal += (iTotal - iHedge);
                                        l3PTotal += iTotal;
                                        lHedgeTotal += iHedge;
                                        lTrust += iTrust;
                                        lSelf += iSelf;
                                        lFor += iForeigen;
                                    }
                                }
                                iLoopDay++;
                            }

                            dtLoopStart = dtLoopStart.AddDays(-1);
                        }

                        if (weekLevelCount.ContainsKey(strno))
                        {
                            double dbper = (double)ladd / (double)weekLevelCount[strno].LevelVol[16];

                            if (dbper > 0.03 && lHedgeTotal > 0)
                            {
                                CandidateParty3Stock.Add(strno);
                                string strmsg = string.Format("{0}", strno);
                                Console.WriteLine(strmsg);
                            }
                        }

                    }
                }

            }
        }

        private void button22_Click(object sender, EventArgs e)
        {
            Dictionary<string, string> stockname = new Dictionary<string, string>();

            Dictionary<string, long> dicTrustHold = new Dictionary<string, long>();
            Dictionary<string, long> dic1MonthagoTrustHold = new Dictionary<string, long>();
            Dictionary<string, long>  dicMonthVol  = new Dictionary<string, long>();


            DateTime dtnow = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);
            DateTime dtLoopStart = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);
            DateTime dt1Monthago = dtLoopStart.AddMonths(-1);
            DateTime dt12Month = dtLoopStart.AddMonths(-2);
            string strmsg = string.Format("投信自{0}/{1}/{2}至今日{3}/{4}/{5}買賣比例，及本月買賣比例，本月成交比重", dt12Month.Year, dt12Month.Month, dt12Month.Day, DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);
            Console.WriteLine(strmsg);

            for (; dtLoopStart > dt12Month; dtLoopStart = dtLoopStart.AddDays(-1))
            {
                bool bisworkday = false;
                string strFileName3Party = string.Format("D:\\Chips\\3Party\\{0:0000}{1:00}{2:00}.json", dtLoopStart.Year, dtLoopStart.Month, dtLoopStart.Day);
                if (File.Exists(strFileName3Party))
                {
                    string json = File.ReadAllText(strFileName3Party);
                    Dictionary<string, object> json_Dictionary = JsonConvert.DeserializeObject<Dictionary<string, object>>(json);
                    if (json_Dictionary.ContainsKey("data"))
                    {
                        List<List<string>> jsondata_ListList = JsonConvert.DeserializeObject<List<List<string>>>(json_Dictionary["data"].ToString());
                        foreach (List<string> marginelist in jsondata_ListList)
                        {
                            if (marginelist.ElementAt(0).Length == 4)
                            {
                                PartyParam Assign = new PartyParam();
                                Assign.Assign(marginelist);

                                if (!dicTrustHold.ContainsKey(Assign.number))
                                {
                                    dicTrustHold.Add(Assign.number, 0);
                                    dic1MonthagoTrustHold.Add(Assign.number, 0);
                                    dicMonthVol.Add(Assign.number, 0);
                                }
                                dicTrustHold[Assign.number] += int.Parse(Assign.TrustTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                if (dtLoopStart > dt1Monthago)
                                {
                                    dic1MonthagoTrustHold[Assign.number] += int.Parse(Assign.TrustTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                    if (dbStockVolDic.ContainsKey(Assign.number) && dbStockVolDic[Assign.number].ContainsKey(dtLoopStart))
                                        dicMonthVol[Assign.number] += dbStockVolDic[Assign.number][dtLoopStart];
                                }
                                if (!stockname.ContainsKey(Assign.number))
                                    stockname.Add(Assign.number, Assign.name);
                            }

                        }
                    }
                }
                string strFileName3PartyTpex = string.Format("D:\\Chips\\3Party\\{0:0000}{1:00}{2:00}_tpex.json", dtLoopStart.Year, dtLoopStart.Month, dtLoopStart.Day);
                if (File.Exists(strFileName3PartyTpex))
                {
                    string json = File.ReadAllText(strFileName3PartyTpex);
                    Dictionary<string, object> json_Dictionary = JsonConvert.DeserializeObject<Dictionary<string, object>>(json);
                    if (json_Dictionary.ContainsKey("aaData"))
                    {
                        List<List<string>> jsondata_ListList = JsonConvert.DeserializeObject<List<List<string>>>(json_Dictionary["aaData"].ToString());
                        foreach (List<string> marginelist in jsondata_ListList)
                        {
                            if (marginelist.ElementAt(0).Length == 4)
                            {
                                PartyParam Assign = new PartyParam();
                                Assign.AssignTPEX(marginelist);

                                if (!dicTrustHold.ContainsKey(Assign.number))
                                {
                                    dicTrustHold.Add(Assign.number, 0);
                                    dic1MonthagoTrustHold.Add(Assign.number, 0);
                                    dicMonthVol.Add(Assign.number, 0);
                                }
                                dicTrustHold[Assign.number] += int.Parse(Assign.TrustTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                if (dtLoopStart > dt1Monthago)
                                {
                                    dic1MonthagoTrustHold[Assign.number] += int.Parse(Assign.TrustTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                    if (dbStockVolDic.ContainsKey(Assign.number) && dbStockVolDic[Assign.number].ContainsKey(dtLoopStart))
                                        dicMonthVol[Assign.number] += dbStockVolDic[Assign.number][dtLoopStart];
                                }

                                if (!stockname.ContainsKey(Assign.number))
                                    stockname.Add(Assign.number, Assign.name);
                            }
                        }
                    }
                }
            }


            string strBasePath = "D:\\Chips\\stock\\TDCC_OD_1-5_";
            string strReadFile = "";
            dtLoopStart = dtnow;
            while (true)
            {
                string strcsv = strBasePath + dtLoopStart.ToString("yyyyMMdd") + ".csv";
                if (File.Exists(strcsv))
                {
                    strReadFile = strcsv;
                    break;
                }
                dtLoopStart = dtLoopStart.AddDays(-1);
            }
            Dictionary<string, StockLevelCount> weekLevelCount = new Dictionary<string, StockLevelCount>();
            int iReadCount = 0;
            IEnumerable<string> lines = File.ReadLines(strReadFile);
            foreach (string line in lines)  // 取流通在外數
            {
                if (iReadCount > 0)
                {
                    string[] strSplitLine = line.Split(',');
                    if (weekLevelCount.ContainsKey(strSplitLine[1]))
                    {
                        int iLevel = int.Parse(strSplitLine[2]) - 1;
                        weekLevelCount[strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                        weekLevelCount[strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                        weekLevelCount[strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                    }
                    else
                    {
                        weekLevelCount.Add(strSplitLine[1], new StockLevelCount());
                        int iLevel = int.Parse(strSplitLine[2]) - 1;
                        weekLevelCount[strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                        weekLevelCount[strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                        weekLevelCount[strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                    }
                }
                iReadCount++;
            }

            foreach(string sno in dicTrustHold.Keys)
            {
                if (weekLevelCount.ContainsKey(sno))
                {
                    double TotalChip = (double)weekLevelCount[sno].LevelVol[16];
                    double dbper = (double)dicTrustHold[sno] / (double)TotalChip;
                    double dbthismonthper = (double)dic1MonthagoTrustHold[sno] / (double)TotalChip;
                    double dbthismonthvol = (double)dic1MonthagoTrustHold[sno] / (double)dicMonthVol[sno];

                    if ( dbper < -0.03)
                    {
                        string smsg = string.Format("{0},{1},{2:F5},{3:F5},--,{4:F5}", sno, stockname[sno], dbper, dbthismonthper, dbthismonthvol);
                        Console.WriteLine(smsg);
                    }
                    else if (dbper <= 0 && dbper>=-0.03)
                    {
                        string smsg = string.Format("{0},{1},{2:F5},{3:F5},-,{4:F5}", sno, stockname[sno], dbper, dbthismonthper, dbthismonthvol);
                        Console.WriteLine(smsg);
                    }
                    else if (dbper > 0 && dbper <= 0.03)
                    {
                        string smsg = string.Format("{0},{1},{2:F5},{3:F5},*,{4:F5}", sno, stockname[sno], dbper, dbthismonthper, dbthismonthvol);
                        Console.WriteLine(smsg);
                    }
                    else if (dbper > 0.03 && dbper < 0.1)
                    {
                        string smsg = string.Format("{0},{1},{2:F5},{3:F5},**,{4:F5}", sno, stockname[sno], dbper, dbthismonthper, dbthismonthvol);
                        Console.WriteLine(smsg);
                    }
                    else if (dbper >= 0.1)
                    {
                        string smsg = string.Format("{0},{1},{2:F5},{3:F5},***,{4:F5}", sno, stockname[sno], dbper, dbthismonthper, dbthismonthvol);
                        Console.WriteLine(smsg);
                    }
                }
            }
            int m = 0;
        }

        private void button23_Click(object sender, EventArgs e)
        {
            Console.WriteLine("投信進4周買比");
            Dictionary<string, long> dicTrustHold = new Dictionary<string, long>();
            Dictionary<string, long> dicLastweekTrustHold = new Dictionary<string, long>();
            Dictionary<string, long> dicLast2weekTrustHold = new Dictionary<string, long>();
            Dictionary<string, long> dicLast3weekTrustHold = new Dictionary<string, long>();
            Dictionary<string, long> dicLast4weekTrustHold = new Dictionary<string, long>();
            DateTime dtLastweek = new DateTime(2019,11,16);
            DateTime dtLast2week = dtLastweek.AddDays(-7);
            DateTime dtLast3week = dtLastweek.AddDays(-14);
            DateTime dtLast4week = dtLastweek.AddDays(-21);

            DateTime dtnow = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);
            DateTime dtLoopStart = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);
            DateTime dt12Month = dtLoopStart.AddMonths(-12);
            dt12Month = new DateTime(2018, 7, 28);
            for (; dtLoopStart > dt12Month; dtLoopStart = dtLoopStart.AddDays(-1))
            {
                bool bisworkday = false;
                string strFileName3Party = string.Format("D:\\Chips\\3Party\\{0:0000}{1:00}{2:00}.json", dtLoopStart.Year, dtLoopStart.Month, dtLoopStart.Day);
                if (File.Exists(strFileName3Party))
                {
                    string json = File.ReadAllText(strFileName3Party);
                    Dictionary<string, object> json_Dictionary = JsonConvert.DeserializeObject<Dictionary<string, object>>(json);
                    if (json_Dictionary.ContainsKey("data"))
                    {
                        List<List<string>> jsondata_ListList = JsonConvert.DeserializeObject<List<List<string>>>(json_Dictionary["data"].ToString());
                        foreach (List<string> marginelist in jsondata_ListList)
                        {
                            if (marginelist.ElementAt(0).Length == 4)
                            {
                                PartyParam Assign = new PartyParam();
                                Assign.Assign(marginelist);

                                if (!dicTrustHold.ContainsKey(Assign.number))
                                {
                                    dicTrustHold.Add(Assign.number, 0);
                                }
                                dicTrustHold[Assign.number] += int.Parse(Assign.TrustTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);

                                if(dtLoopStart<dtLastweek)
                                {
                                    if (!dicLastweekTrustHold.ContainsKey(Assign.number))
                                    {
                                        dicLastweekTrustHold.Add(Assign.number, 0);
                                    }
                                    dicLastweekTrustHold[Assign.number] += int.Parse(Assign.TrustTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                }
                                if(dtLoopStart<dtLast2week)
                                {
                                    if (!dicLast2weekTrustHold.ContainsKey(Assign.number))
                                    {
                                        dicLast2weekTrustHold.Add(Assign.number, 0);
                                    }
                                    dicLast2weekTrustHold[Assign.number] += int.Parse(Assign.TrustTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                }
                                if (dtLoopStart < dtLast3week)
                                {
                                    if (!dicLast3weekTrustHold.ContainsKey(Assign.number))
                                    {
                                        dicLast3weekTrustHold.Add(Assign.number, 0);
                                    }
                                    dicLast3weekTrustHold[Assign.number] += int.Parse(Assign.TrustTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                }
                                if (dtLoopStart < dtLast4week)
                                {
                                    if (!dicLast4weekTrustHold.ContainsKey(Assign.number))
                                    {
                                        dicLast4weekTrustHold.Add(Assign.number, 0);
                                    }
                                    dicLast4weekTrustHold[Assign.number] += int.Parse(Assign.TrustTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                }
                            }

                        }
                    }
                }
                string strFileName3PartyTpex = string.Format("D:\\Chips\\3Party\\{0:0000}{1:00}{2:00}_tpex.json", dtLoopStart.Year, dtLoopStart.Month, dtLoopStart.Day);
                if (File.Exists(strFileName3PartyTpex))
                {
                    string json = File.ReadAllText(strFileName3PartyTpex);
                    Dictionary<string, object> json_Dictionary = JsonConvert.DeserializeObject<Dictionary<string, object>>(json);
                    if (json_Dictionary.ContainsKey("aaData"))
                    {
                        List<List<string>> jsondata_ListList = JsonConvert.DeserializeObject<List<List<string>>>(json_Dictionary["aaData"].ToString());
                        foreach (List<string> marginelist in jsondata_ListList)
                        {
                            if (marginelist.ElementAt(0).Length == 4)
                            {
                                PartyParam Assign = new PartyParam();
                                Assign.AssignTPEX(marginelist);

                                if (!dicTrustHold.ContainsKey(Assign.number))
                                {
                                    dicTrustHold.Add(Assign.number, 0);
                                }
                                dicTrustHold[Assign.number] += int.Parse(Assign.TrustTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);

                                if (dtLoopStart < dtLastweek)
                                {
                                    if (!dicLastweekTrustHold.ContainsKey(Assign.number))
                                    {
                                        dicLastweekTrustHold.Add(Assign.number, 0);
                                    }
                                    dicLastweekTrustHold[Assign.number] += int.Parse(Assign.TrustTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                }
                                if (dtLoopStart < dtLast2week)
                                {
                                    if (!dicLast2weekTrustHold.ContainsKey(Assign.number))
                                    {
                                        dicLast2weekTrustHold.Add(Assign.number, 0);
                                    }
                                    dicLast2weekTrustHold[Assign.number] += int.Parse(Assign.TrustTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                }
                                if (dtLoopStart < dtLast3week)
                                {
                                    if (!dicLast3weekTrustHold.ContainsKey(Assign.number))
                                    {
                                        dicLast3weekTrustHold.Add(Assign.number, 0);
                                    }
                                    dicLast3weekTrustHold[Assign.number] += int.Parse(Assign.TrustTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                }
                                if (dtLoopStart < dtLast4week)
                                {
                                    if (!dicLast4weekTrustHold.ContainsKey(Assign.number))
                                    {
                                        dicLast4weekTrustHold.Add(Assign.number, 0);
                                    }
                                    dicLast4weekTrustHold[Assign.number] += int.Parse(Assign.TrustTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                }
                            }

                        }
                    }
                }
            }


            string strBasePath = "D:\\Chips\\stock\\TDCC_OD_1-5_";
            string strReadFile = "";
            dtLoopStart = dtnow;
            while (true)
            {
                string strcsv = strBasePath + dtLoopStart.ToString("yyyyMMdd") + ".csv";
                if (File.Exists(strcsv))
                {
                    strReadFile = strcsv;
                    break;
                }
                dtLoopStart = dtLoopStart.AddDays(-1);
            }
            Dictionary<string, StockLevelCount> weekLevelCount = new Dictionary<string, StockLevelCount>();
            int iReadCount = 0;
            IEnumerable<string> lines = File.ReadLines(strReadFile);
            foreach (string line in lines)  // 取流通在外數
            {
                if (iReadCount > 0)
                {
                    string[] strSplitLine = line.Split(',');
                    if (weekLevelCount.ContainsKey(strSplitLine[1]))
                    {
                        int iLevel = int.Parse(strSplitLine[2]) - 1;
                        weekLevelCount[strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                        weekLevelCount[strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                        weekLevelCount[strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                    }
                    else
                    {
                        weekLevelCount.Add(strSplitLine[1], new StockLevelCount());
                        int iLevel = int.Parse(strSplitLine[2]) - 1;
                        weekLevelCount[strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                        weekLevelCount[strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                        weekLevelCount[strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                    }
                }
                iReadCount++;
            }

            foreach (string sno in dicTrustHold.Keys)
            {
                if (weekLevelCount.ContainsKey(sno) && dicTrustHold.ContainsKey(sno) && dicLastweekTrustHold.ContainsKey(sno) && dicLast2weekTrustHold.ContainsKey(sno) && dicLast3weekTrustHold.ContainsKey(sno) && dicLast4weekTrustHold.ContainsKey(sno))
                {
                    double TotalChip = (double)weekLevelCount[sno].LevelVol[16];
                    double dbper = (double)dicTrustHold[sno] / (double)TotalChip;

                    double dbLastweekper = (double)dicLastweekTrustHold[sno] / (double)TotalChip;
                    double dbLast2weekper = (double)dicLast2weekTrustHold[sno] / (double)TotalChip;
                    double dbLast3weekper = (double)dicLast3weekTrustHold[sno] / (double)TotalChip;
                    double dbLast4weekper = (double)dicLast4weekTrustHold[sno] / (double)TotalChip;

                    if (dbper < -0.03)
                    {
                        string smsg = string.Format("{0},{1:F5},--,{2:F5},{3:F5},{4:F5},{5:F5}", sno, dbper, (dbper - dbLastweekper), (dbLastweekper - dbLast2weekper), (dbLast2weekper - dbLast3weekper), (dbLast3weekper - dbLast4weekper));
                        Console.WriteLine(smsg);
                    }
                    else if (dbper <= 0 && dbper >= -0.03)
                    {
                        string smsg = string.Format("{0},{1:F5},-,{2:F5},{3:F5},{4:F5},{5:F5}", sno, dbper, (dbper - dbLastweekper), (dbLastweekper - dbLast2weekper), (dbLast2weekper - dbLast3weekper), (dbLast3weekper - dbLast4weekper));
                        Console.WriteLine(smsg);
                    }
                    else if (dbper > 0 && dbper <= 0.03)
                    {
                        string smsg = string.Format("{0},{1:F5},*,{2:F5},{3:F5},{4:F5},{5:F5}", sno, dbper, (dbper - dbLastweekper), (dbLastweekper - dbLast2weekper), (dbLast2weekper - dbLast3weekper), (dbLast3weekper - dbLast4weekper));
                        Console.WriteLine(smsg);
                    }
                    else if (dbper > 0.03 && dbper < 0.1)
                    {
                        string smsg = string.Format("{0},{1:F5},**,{2:F5},{3:F5},{4:F5},{5:F5}", sno, dbper, (dbper - dbLastweekper), (dbLastweekper - dbLast2weekper), (dbLast2weekper - dbLast3weekper), (dbLast3weekper - dbLast4weekper));
                        Console.WriteLine(smsg);
                    }
                    else if (dbper >= 0.1)
                    {
                        string smsg = string.Format("{0},{1:F5},***,{2:F5},{3:F5},{4:F5},{5:F5}", sno, dbper, (dbper - dbLastweekper), (dbLastweekper - dbLast2weekper), (dbLast2weekper - dbLast3weekper), (dbLast3weekper - dbLast4weekper));
                        Console.WriteLine(smsg);
                    }
                }
            }
            int m = 0;
        }

        private void button24_Click(object sender, EventArgs e)
        {
            //https://fubon-ebrokerdj.fbs.com.tw/z/zc/zc0/zc07/zc07_3653.djhtm
            string url = string.Format("https://fubon-ebrokerdj.fbs.com.tw/z/zc/zc0/zc07/zc07_3653.djhtm");

            HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
            httpWebRequest.ContentType = "application/json";
            httpWebRequest.Method = WebRequestMethods.Http.Get;
            httpWebRequest.Accept = "application/json";
            string text;
            var response = (HttpWebResponse)httpWebRequest.GetResponse();

            using (var sr = new StreamReader(response.GetResponseStream(), Encoding.GetEncoding("big5")))
            {
                text = sr.ReadToEnd();
            }

            File.WriteAllText("D:\\a.htm", text, Encoding.GetEncoding("big5"));
            response.Close();
            httpWebRequest.Abort();
        }

        private void button25_Click(object sender, EventArgs e)
        {
            // 法卷比 dicStock3Party / dicStockMarriage
            int iRange = 5; //iRange天內
            int iMarCount = 800;
            DateTime dtStart = monthCalendar1.SelectionStart;
            DateTime dtLoopStart = monthCalendar1.SelectionStart;

            string strMsg = string.Format("找{0}前{1}天投信外資買比融券數", dtStart.ToString("MM/dd"), iRange);
            Console.WriteLine(strMsg);
            strMsg = string.Format("號,法卷比,名,外買,信買,收價,卷,資");
            Console.WriteLine(strMsg);


            CandidateParty3Stock.Clear();

            //CandidateWeekChipSet.Clear();
            Dictionary<string, long> stock3PartyVol = new Dictionary<string, long>();
            foreach (string strStockNo in dicStock3Party.Keys)
            {
                long l3PTotal = 0;
                long lHedgeTotal = 0;
                long lTrust = 0;
                long lSelf = 0;
                long lTotal = 0;
                long lFor = 0;

                long ladd = 0;
                int iLoopDay = 0;
                dtLoopStart = dtStart;
                while (iLoopDay < iRange)  // 計算最近幾天
                {
                    if (workingdays.Contains(dtLoopStart))
                    {
                        if (dicStock3Party.ContainsKey(strStockNo))
                        {
                            if (dicStock3Party[strStockNo].ContainsKey(dtLoopStart))
                            {
                                int iTotal = int.Parse(dicStock3Party[strStockNo][dtLoopStart].Party3Total, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                int iHedge = int.Parse(dicStock3Party[strStockNo][dtLoopStart].SelfHedgeTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                int iHedgeBuy = int.Parse(dicStock3Party[strStockNo][dtLoopStart].SelfHedgeBuy, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                int iHedgeSell = int.Parse(dicStock3Party[strStockNo][dtLoopStart].SelfHedgeSell, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                int iForeigen = int.Parse(dicStock3Party[strStockNo][dtLoopStart].ForeigenTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                int iTrust = int.Parse(dicStock3Party[strStockNo][dtLoopStart].TrustTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                int iSelf = int.Parse(dicStock3Party[strStockNo][dtLoopStart].SelfSelfTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                ladd += (iTotal - iHedge); // 計算不含避險3大買量
                                //ladd += iTotal; // 全算
                                //ladd += (iTrust + iForeigen); //外 信
                                //ladd += iHedge;

                                lTotal += (iTotal - iHedge);
                                l3PTotal += iTotal;
                                lHedgeTotal += iHedge;
                                lTrust += iTrust;
                                lSelf += iSelf;
                                lFor += iForeigen;
                            }
                        }
                        iLoopDay++;
                    }

                    dtLoopStart = dtLoopStart.AddDays(-1);
                }

                stock3PartyVol.Add(strStockNo, lFor + lTrust);

                if (dicStockMarriage.ContainsKey(strStockNo) && dicStockMarriage[strStockNo].ContainsKey(dtStart))
                {
                    double dbper = (double)((lFor + lTrust)/1000) / (double)dicStockMarriage[strStockNo][dtStart];
                    if (dbper > 1.5 && dicStockMarriage[strStockNo][dtStart] > iMarCount)
                    {
                        CandidateParty3Stock.Add(strStockNo);
                        string strmsg;
                        if(dicStock3Party[strStockNo].ContainsKey(dtStart))
                        {
                            strmsg = string.Format("{0},{1},{2},{3},{4},{5},{6},{7}", strStockNo, dbper, dicStockNoMapping[strStockNo], lFor, lTrust, dbEvenydayOLHCDic[strStockNo][dtStart][3], dicStockMarriage[strStockNo][dtStart], dicStockFinancing[strStockNo][dtStart]);
                        }                            
                        else
                        {
                            strmsg = string.Format("{0},{1},{2},{3},{4},{5},{6},{7}", strStockNo, dbper, dicStockNoMapping[strStockNo], lFor, lTrust, dbEvenydayOLHCDic[strStockNo][dtStart][3], dicStockMarriage[strStockNo][dtStart], dicStockFinancing[strStockNo][dtStart]);
                        }
                        Console.WriteLine(strmsg);
                    }
                }
            }    
        }

        private void button26_Click(object sender, EventArgs e)
        {
            DateTime dtStart = monthCalendar1.SelectionStart;
            DateTime dtLoopStart = monthCalendar1.SelectionStart;

            string strMsg = string.Format("查{0} 當天創60 120 240 新高,60 120 240 新低數", dtStart.ToString("MM/dd"));
            Console.WriteLine(strMsg);


            int iBreakHigh60 = 0;
            int iBreakLow60 = 0;
            int iBreakHigh120 = 0;
            int iBreakLow120 = 0;
            int iBreakHigh240 = 0;
            int iBreakLow240 = 0;

            foreach(string strno in dbEvenydayOLHCDic.Keys)
            {
                if (dbEvenydayOLHCDic[strno].Count > 280 && strno.Length == 4 && dbEvenydayOLHCDic[strno].ContainsKey(dtStart))
                {
                    List<double> dbCloseList60 = new List<double>();
                    List<double> dbCloseList120 = new List<double>();
                    List<double> dbCloseList240 = new List<double>();

                    DateTime dtLoop = dtStart;
                    int iCountAdded = 0;

                    while (iCountAdded < 240)
                    {
                        if (dbEvenydayOLHCDic[strno].ContainsKey(dtLoop))
                        {
                            if (dbCloseList60.Count<60)
                                dbCloseList60.Add(dbEvenydayOLHCDic[strno][dtLoop][3]);
                            if (dbCloseList120.Count < 120)
                                dbCloseList120.Add(dbEvenydayOLHCDic[strno][dtLoop][3]);
                            if (dbCloseList240.Count < 240)
                                dbCloseList240.Add(dbEvenydayOLHCDic[strno][dtLoop][3]);

                            iCountAdded++;
                        }
                        dtLoop = dtLoop.AddDays(-1);
                    }
                    dtLoop = dtStart;
                    int m = 0;
                    if(dbEvenydayOLHCDic[strno][dtStart][3]>=dbCloseList60.Max())
                    {
                        iBreakHigh60++;
                    }
                    if (dbEvenydayOLHCDic[strno][dtStart][3] >= dbCloseList120.Max())
                    {
                        iBreakHigh120++;
                    }
                    if (dbEvenydayOLHCDic[strno][dtStart][3] >= dbCloseList240.Max())
                    {
                        iBreakHigh240++;
                    }

                    if (dbEvenydayOLHCDic[strno][dtStart][3] <= dbCloseList60.Min())
                    {
                        iBreakLow60++;
                    }
                    if (dbEvenydayOLHCDic[strno][dtStart][3] <= dbCloseList120.Min())
                    {
                        iBreakLow120++;
                    }
                    if (dbEvenydayOLHCDic[strno][dtStart][3] <= dbCloseList240.Min())
                    {
                        iBreakLow240++;
                    }
                }

            }
            string strMsgout = string.Format("{0},{1},{2},{3},{4},{5},{6}", dtStart.ToString("MM/dd"), iBreakHigh60, iBreakHigh120, iBreakHigh240, iBreakLow60, iBreakLow120, iBreakLow240);
            Console.WriteLine(strMsgout);
        }

        private void button27_Click(object sender, EventArgs e)
        {
            DateTime dtEnd = new DateTime(2018, 10, 20);
            DateTime dtStart = DateTime.Now;
            DateTime dtLoopEnd = dtStart.AddMonths(-7);
            DateTime dtLoopStart = dtStart;

            string strfilename = string.Format("{0}People.csv", dtStart.ToString("yyyyMMdd"));
            File.AppendAllText(strfilename, "\r\n", Encoding.UTF8);


            string strfilenameLV = string.Format("{0}PeopleLV.csv", dtStart.ToString("yyyyMMdd"));
            File.AppendAllText(strfilenameLV, "\r\n", Encoding.UTF8);

            string strfilename600 = string.Format("{0}holder600.csv", dtStart.ToString("yyyyMMdd"));
            File.AppendAllText(strfilename600, "\r\n", Encoding.UTF8);


            //Queue<string> qcsvfile = new Queue<string>();
            string strBasePath = "D:\\Chips\\stock\\TDCC_OD_1-5_";
            int iCount = 32;
            Dictionary<string, StockLevelCount>[] weekLevelCount = new Dictionary<string, StockLevelCount>[iCount];
            int iCountWeek = 0;
            while (dtLoopStart > dtLoopEnd)
            {
                string strcsv = strBasePath + dtLoopStart.ToString("yyyyMMdd") + ".csv";
                if (File.Exists(strcsv))
                {
                    weekLevelCount[iCountWeek] = new Dictionary<string, StockLevelCount>();
                    int iReadCount = 0;
                    IEnumerable<string> lines = File.ReadLines(strcsv);
                    foreach (string line in lines)
                    {
                        if (iReadCount > 0)
                        {
                            string[] strSplitLine = line.Split(',');
                            if (weekLevelCount[iCountWeek].ContainsKey(strSplitLine[1]))
                            {
                                int iLevel = int.Parse(strSplitLine[2]) - 1;
                                weekLevelCount[iCountWeek][strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                                weekLevelCount[iCountWeek][strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                                weekLevelCount[iCountWeek][strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                            }
                            else
                            {
                                weekLevelCount[iCountWeek].Add(strSplitLine[1], new StockLevelCount());
                                int iLevel = int.Parse(strSplitLine[2]) - 1;
                                weekLevelCount[iCountWeek][strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                                weekLevelCount[iCountWeek][strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                                weekLevelCount[iCountWeek][strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                            }
                        }
                        iReadCount++;
                    }
                    iCountWeek++;
                }
                dtLoopStart = dtLoopStart.AddDays(-1);
            }
    
        
            foreach (KeyValuePair<string, StockLevelCount> item in weekLevelCount[0])
            {
                if (weekLevelCount[1].ContainsKey(item.Key) && weekLevelCount[2].ContainsKey(item.Key) && weekLevelCount[3].ContainsKey(item.Key) && weekLevelCount[4].ContainsKey(item.Key) && weekLevelCount[5].ContainsKey(item.Key) && weekLevelCount[6].ContainsKey(item.Key) && weekLevelCount[7].ContainsKey(item.Key) &&
                    weekLevelCount[8].ContainsKey(item.Key) && weekLevelCount[9].ContainsKey(item.Key) && weekLevelCount[10].ContainsKey(item.Key) && weekLevelCount[11].ContainsKey(item.Key) && weekLevelCount[12].ContainsKey(item.Key) && weekLevelCount[13].ContainsKey(item.Key) && weekLevelCount[14].ContainsKey(item.Key)
                    && weekLevelCount[15].ContainsKey(item.Key) && weekLevelCount[16].ContainsKey(item.Key) && weekLevelCount[17].ContainsKey(item.Key) && weekLevelCount[18].ContainsKey(item.Key) && weekLevelCount[19].ContainsKey(item.Key) && weekLevelCount[20].ContainsKey(item.Key) && weekLevelCount[21].ContainsKey(item.Key)
                    && weekLevelCount[22].ContainsKey(item.Key) && weekLevelCount[23].ContainsKey(item.Key) && weekLevelCount[24].ContainsKey(item.Key))
                {
                    // LevelRate[15] 是差異數調整，[14]:1,000,001以上，[13]:800,001-1,000,000，[12]:600,001-800,000，[11]:400,001-600,000，[10]200,001-400,000   ，[9]:100,001-200,000，[8]:50,001-100,000
                    double dbrate0 = weekLevelCount[0][item.Key].LevelRate[12] + weekLevelCount[0][item.Key].LevelRate[13] + weekLevelCount[0][item.Key].LevelRate[14];// +weekLevelCount[0][item.Key].LevelRate[15];
                    double dbrate1 = weekLevelCount[1][item.Key].LevelRate[12] + weekLevelCount[1][item.Key].LevelRate[13] + weekLevelCount[1][item.Key].LevelRate[14];// + weekLevelCount[1][item.Key].LevelRate[15];
                    double dbrate2 = weekLevelCount[2][item.Key].LevelRate[12] + weekLevelCount[2][item.Key].LevelRate[13] + weekLevelCount[2][item.Key].LevelRate[14];// + weekLevelCount[2][item.Key].LevelRate[15];
                    double dbrate3 = weekLevelCount[3][item.Key].LevelRate[12] + weekLevelCount[3][item.Key].LevelRate[13] + weekLevelCount[3][item.Key].LevelRate[14];// + weekLevelCount[3][item.Key].LevelRate[15];
                    double dbrate4 = weekLevelCount[4][item.Key].LevelRate[12] + weekLevelCount[4][item.Key].LevelRate[13] + weekLevelCount[4][item.Key].LevelRate[14];// + weekLevelCount[4][item.Key].LevelRate[15];
                    double dbrate5 = weekLevelCount[5][item.Key].LevelRate[12] + weekLevelCount[5][item.Key].LevelRate[13] + weekLevelCount[5][item.Key].LevelRate[14];// + weekLevelCount[5][item.Key].LevelRate[15];
                    double dbrate6 = weekLevelCount[6][item.Key].LevelRate[12] + weekLevelCount[6][item.Key].LevelRate[13] + weekLevelCount[6][item.Key].LevelRate[14];// + weekLevelCount[6][item.Key].LevelRate[15];
                    double dbrate7 = weekLevelCount[7][item.Key].LevelRate[12] + weekLevelCount[7][item.Key].LevelRate[13] + weekLevelCount[7][item.Key].LevelRate[14];// + weekLevelCount[7][item.Key].LevelRate[15];
                    double dbrate08 = weekLevelCount[08][item.Key].LevelRate[12] + weekLevelCount[08][item.Key].LevelRate[13] + weekLevelCount[08][item.Key].LevelRate[14];// + weekLevelCount[08][item.Key].LevelRate[15];
                    double dbrate09 = weekLevelCount[09][item.Key].LevelRate[12] + weekLevelCount[09][item.Key].LevelRate[13] + weekLevelCount[09][item.Key].LevelRate[14];// + weekLevelCount[09][item.Key].LevelRate[15];
                    double dbrate10 = weekLevelCount[10][item.Key].LevelRate[12] + weekLevelCount[10][item.Key].LevelRate[13] + weekLevelCount[10][item.Key].LevelRate[14];// + weekLevelCount[10][item.Key].LevelRate[15];
                    double dbrate11 = weekLevelCount[11][item.Key].LevelRate[12] + weekLevelCount[11][item.Key].LevelRate[13] + weekLevelCount[11][item.Key].LevelRate[14];// + weekLevelCount[11][item.Key].LevelRate[15];
                    double dbrate12 = weekLevelCount[12][item.Key].LevelRate[12] + weekLevelCount[12][item.Key].LevelRate[13] + weekLevelCount[12][item.Key].LevelRate[14];// + weekLevelCount[12][item.Key].LevelRate[15];
                    double dbrate13 = weekLevelCount[13][item.Key].LevelRate[12] + weekLevelCount[13][item.Key].LevelRate[13] + weekLevelCount[13][item.Key].LevelRate[14];// + weekLevelCount[13][item.Key].LevelRate[15];
                    double dbrate14 = weekLevelCount[14][item.Key].LevelRate[12] + weekLevelCount[14][item.Key].LevelRate[13] + weekLevelCount[14][item.Key].LevelRate[14];// + weekLevelCount[14][item.Key].LevelRate[15];
                    double dbrate15 = weekLevelCount[15][item.Key].LevelRate[12] + weekLevelCount[15][item.Key].LevelRate[13] + weekLevelCount[15][item.Key].LevelRate[14];//+ weekLevelCount[15][item.Key].LevelRate[15];

                    // LevelPeople[16] 總人數 - [0]1-999股 - [1]:1-5張 -[2]:5-10張
                    double dbpeople0 = weekLevelCount[0][item.Key].LevelPeople[16] - weekLevelCount[0][item.Key].LevelPeople[0] - weekLevelCount[0][item.Key].LevelPeople[1];// - weekLevelCount[0][item.Key].LevelPeople[2];
                    double dbpeople1 = weekLevelCount[1][item.Key].LevelPeople[16] - weekLevelCount[1][item.Key].LevelPeople[0] - weekLevelCount[1][item.Key].LevelPeople[1];// - weekLevelCount[1][item.Key].LevelPeople[2];
                    double dbpeople2 = weekLevelCount[2][item.Key].LevelPeople[16] - weekLevelCount[2][item.Key].LevelPeople[0] - weekLevelCount[2][item.Key].LevelPeople[1];// - weekLevelCount[2][item.Key].LevelPeople[2];
                    double dbpeople3 = weekLevelCount[3][item.Key].LevelPeople[16] - weekLevelCount[3][item.Key].LevelPeople[0] - weekLevelCount[3][item.Key].LevelPeople[1];// - weekLevelCount[3][item.Key].LevelPeople[2];
                    double dbpeople4 = weekLevelCount[4][item.Key].LevelPeople[16] - weekLevelCount[4][item.Key].LevelPeople[0] - weekLevelCount[4][item.Key].LevelPeople[1];// - weekLevelCount[4][item.Key].LevelPeople[2];
                    double dbpeople5 = weekLevelCount[5][item.Key].LevelPeople[16] - weekLevelCount[5][item.Key].LevelPeople[0] - weekLevelCount[5][item.Key].LevelPeople[1];// - weekLevelCount[5][item.Key].LevelPeople[2];
                    double dbpeople6 = weekLevelCount[6][item.Key].LevelPeople[16] - weekLevelCount[6][item.Key].LevelPeople[0] - weekLevelCount[6][item.Key].LevelPeople[1];// - weekLevelCount[6][item.Key].LevelPeople[2];
                    double dbpeople7 = weekLevelCount[7][item.Key].LevelPeople[16] - weekLevelCount[7][item.Key].LevelPeople[0] - weekLevelCount[7][item.Key].LevelPeople[1];// -weekLevelCount[7][item.Key].LevelPeople[2];
                    double dbpeople08 = weekLevelCount[08][item.Key].LevelPeople[16] - weekLevelCount[08][item.Key].LevelPeople[0] - weekLevelCount[08][item.Key].LevelPeople[1];// - weekLevelCount[08][item.Key].LevelPeople[2];
                    double dbpeople09 = weekLevelCount[09][item.Key].LevelPeople[16] - weekLevelCount[09][item.Key].LevelPeople[0] - weekLevelCount[09][item.Key].LevelPeople[1];// - weekLevelCount[09][item.Key].LevelPeople[2];
                    double dbpeople10 = weekLevelCount[10][item.Key].LevelPeople[16] - weekLevelCount[10][item.Key].LevelPeople[0] - weekLevelCount[10][item.Key].LevelPeople[1];// - weekLevelCount[10][item.Key].LevelPeople[2];
                    double dbpeople11 = weekLevelCount[11][item.Key].LevelPeople[16] - weekLevelCount[11][item.Key].LevelPeople[0] - weekLevelCount[11][item.Key].LevelPeople[1];// - weekLevelCount[11][item.Key].LevelPeople[2];
                    double dbpeople12 = weekLevelCount[12][item.Key].LevelPeople[16] - weekLevelCount[12][item.Key].LevelPeople[0] - weekLevelCount[12][item.Key].LevelPeople[1];// - weekLevelCount[12][item.Key].LevelPeople[2];
                    double dbpeople13 = weekLevelCount[13][item.Key].LevelPeople[16] - weekLevelCount[13][item.Key].LevelPeople[0] - weekLevelCount[13][item.Key].LevelPeople[1];// - weekLevelCount[13][item.Key].LevelPeople[2];
                    double dbpeople14 = weekLevelCount[14][item.Key].LevelPeople[16] - weekLevelCount[14][item.Key].LevelPeople[0] - weekLevelCount[14][item.Key].LevelPeople[1];// - weekLevelCount[14][item.Key].LevelPeople[2];
                    double dbpeople15 = weekLevelCount[15][item.Key].LevelPeople[16] - weekLevelCount[15][item.Key].LevelPeople[0] - weekLevelCount[15][item.Key].LevelPeople[1];// -weekLevelCount[15][item.Key].LevelPeople[2];
                    double dbpeople16 = weekLevelCount[16][item.Key].LevelPeople[16] - weekLevelCount[16][item.Key].LevelPeople[0] - weekLevelCount[16][item.Key].LevelPeople[1];
                    double dbpeople17 = weekLevelCount[17][item.Key].LevelPeople[16] - weekLevelCount[17][item.Key].LevelPeople[0] - weekLevelCount[17][item.Key].LevelPeople[1];
                    double dbpeople18 = weekLevelCount[18][item.Key].LevelPeople[16] - weekLevelCount[18][item.Key].LevelPeople[0] - weekLevelCount[18][item.Key].LevelPeople[1];
                    double dbpeople19 = weekLevelCount[19][item.Key].LevelPeople[16] - weekLevelCount[19][item.Key].LevelPeople[0] - weekLevelCount[19][item.Key].LevelPeople[1];
                    double dbpeople20 = weekLevelCount[20][item.Key].LevelPeople[16] - weekLevelCount[20][item.Key].LevelPeople[0] - weekLevelCount[20][item.Key].LevelPeople[1];
                    double dbpeople21 = weekLevelCount[21][item.Key].LevelPeople[16] - weekLevelCount[21][item.Key].LevelPeople[0] - weekLevelCount[21][item.Key].LevelPeople[1];
                    double dbpeople22 = weekLevelCount[22][item.Key].LevelPeople[16] - weekLevelCount[22][item.Key].LevelPeople[0] - weekLevelCount[22][item.Key].LevelPeople[1];
                    double dbpeople23 = weekLevelCount[23][item.Key].LevelPeople[16] - weekLevelCount[23][item.Key].LevelPeople[0] - weekLevelCount[23][item.Key].LevelPeople[1];
                    double dbpeople24 = weekLevelCount[24][item.Key].LevelPeople[16] - weekLevelCount[24][item.Key].LevelPeople[0] - weekLevelCount[24][item.Key].LevelPeople[1];

                    string strf = item.Key;
                    for (int i = 0; i < 24;i++ )
                    {
                        if(weekLevelCount[i]!= null)
                        {
                            double dbpeople = weekLevelCount[i][item.Key].LevelPeople[16] - weekLevelCount[i][item.Key].LevelPeople[0] - weekLevelCount[i][item.Key].LevelPeople[1];
                            string stradd = string.Format(",{0:0.##}", dbpeople);
                            strf += stradd;
                        }
                    }
                    strf += "\r\n";
                    File.AppendAllText(strfilename, strf);


                    string strfLV = item.Key;
                    for (int i = 0; i < 24; i++)
                    {
                        if (weekLevelCount[i] != null)
                        {
                            //double dbpeople = weekLevelCount[i][item.Key].LevelPeople[16] - weekLevelCount[i][item.Key].LevelPeople[0] - weekLevelCount[i][item.Key].LevelPeople[1];
                            string straddLV = string.Format(",[0],{0:0.##},[1],{1:0.##}", weekLevelCount[i][item.Key].LevelPeople[0], weekLevelCount[i][item.Key].LevelPeople[1]);
                            strfLV += straddLV;
                        }
                    }

                    strfLV += "\r\n";
                    File.AppendAllText(strfilenameLV, strfLV);


                    string strf600 = item.Key;
                    for (int i = 0; i < 24; i++)
                    {
                        if (weekLevelCount[i] != null)
                        {
                            //double dbpeople = weekLevelCount[i][item.Key].LevelPeople[16] - weekLevelCount[i][item.Key].LevelPeople[0] - weekLevelCount[i][item.Key].LevelPeople[1];
                            string stradd600 = string.Format(",{0:0.##}", weekLevelCount[i][item.Key].LevelRate[11] + weekLevelCount[i][item.Key].LevelRate[12] + weekLevelCount[i][item.Key].LevelRate[13] + weekLevelCount[i][item.Key].LevelRate[14]);
                            strf600 += stradd600;
                        }
                    }

                    strf600 += "\r\n";
                    File.AppendAllText(strfilename600, strf600);

                }
            }
        }

        private void button28_Click(object sender, EventArgs e)
        {
            
            double dbPercent = 500*100000; // 成長數大於總股數
            DateTime dtStart = monthCalendar1.SelectionStart;
            DateTime dtLoopStart = monthCalendar1.SelectionStart;

            string strMsg = string.Format("找{0}投信或外資買超過市值{1:F4}", dtStart.ToString("MM/dd"), dbPercent);
            Console.WriteLine(strMsg);

            CandidateParty3Stock.Clear();

            //CandidateWeekChipSet.Clear();
            string strBasePath = "D:\\Chips\\stock\\TDCC_OD_1-5_";
            string strReadFile = "";
            int icountchip = 0;
            while (true)
            {
                string strcsv = strBasePath + dtLoopStart.ToString("yyyyMMdd") + ".csv";
                if (File.Exists(strcsv))
                {
                    strReadFile = strcsv;
                    break;
                }
                dtLoopStart = dtLoopStart.AddDays(-1);
            }
            Dictionary<string, StockLevelCount> weekLevelCount = new Dictionary<string, StockLevelCount>();
            int iReadCount = 0;
            IEnumerable<string> lines = File.ReadLines(strReadFile);
            foreach (string line in lines)  // 取流通在外數
            {
                if (iReadCount > 0)
                {
                    string[] strSplitLine = line.Split(',');
                    if (weekLevelCount.ContainsKey(strSplitLine[1]))
                    {
                        int iLevel = int.Parse(strSplitLine[2]) - 1;
                        weekLevelCount[strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                        weekLevelCount[strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                        weekLevelCount[strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                    }
                    else
                    {
                        weekLevelCount.Add(strSplitLine[1], new StockLevelCount());
                        int iLevel = int.Parse(strSplitLine[2]) - 1;
                        weekLevelCount[strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                        weekLevelCount[strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                        weekLevelCount[strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                    }
                }
                iReadCount++;
            }

            Dictionary<string, long> stock3PartyVol = new Dictionary<string, long>();
            foreach (string strStockNo in weekLevelCount.Keys)
            {
                long l3PTotal = 0;
                long lHedgeTotal = 0;
                long lTrust = 0;
                long lSelf = 0;
                long lTotal = 0;
                long lFor = 0;

                long ladd = 0;

                dtLoopStart = dtStart;

                if (workingdays.Contains(dtLoopStart))
                {
                    if (dicStock3Party.ContainsKey(strStockNo))
                    {
                        if (dicStock3Party[strStockNo].ContainsKey(dtLoopStart))
                        {
                            int iTotal = int.Parse(dicStock3Party[strStockNo][dtLoopStart].Party3Total, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                            int iHedge = int.Parse(dicStock3Party[strStockNo][dtLoopStart].SelfHedgeTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                            int iHedgeBuy = int.Parse(dicStock3Party[strStockNo][dtLoopStart].SelfHedgeBuy, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                            int iHedgeSell = int.Parse(dicStock3Party[strStockNo][dtLoopStart].SelfHedgeSell, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                            int iForeigen = int.Parse(dicStock3Party[strStockNo][dtLoopStart].ForeigenTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                            int iTrust = int.Parse(dicStock3Party[strStockNo][dtLoopStart].TrustTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                            int iSelf = int.Parse(dicStock3Party[strStockNo][dtLoopStart].SelfSelfTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                            ladd += (iTotal - iHedge); // 計算不含避險3大買量
                            //ladd += iTotal; // 全算
                            //ladd += (iTrust + iForeigen); //外 信
                            //ladd += iHedge;

                            lTotal += (iTotal - iHedge);
                            l3PTotal += iTotal;
                            lHedgeTotal += iHedge;
                            lTrust += iTrust;
                            lSelf += iSelf;
                            lFor += iForeigen;

                            double dbclose = dbEvenydayOLHCDic[strStockNo][dtStart][3];
                            if (dbEvenydayOLHCDic.ContainsKey(strStockNo) && dbEvenydayOLHCDic[strStockNo].ContainsKey(dtLoopStart) && dbclose<200)
                            {
                                double dbForVal = lFor * dbEvenydayOLHCDic[strStockNo][dtLoopStart][3];
                                double dbTrustVal = lTrust * dbEvenydayOLHCDic[strStockNo][dtLoopStart][3];
                                double dbHedgeVal = lHedgeTotal * dbEvenydayOLHCDic[strStockNo][dtLoopStart][3];
                                double dbSelfVal = lSelf * dbEvenydayOLHCDic[strStockNo][dtLoopStart][3];
                                
                                if (dbTrustVal >= dbPercent || dbTrustVal <= -dbPercent)
                                {
                                    CandidateParty3Stock.Add(strStockNo);
                                    Console.WriteLine("信 ," + strStockNo + "," + dicStock3Party[strStockNo][dtLoopStart].name + "," + dbTrustVal + "," + dbclose);
                                }
                                if (dbForVal >= dbPercent || dbForVal <= -dbPercent)
                                {
                                    CandidateParty3Stock.Add(strStockNo);
                                    Console.WriteLine("外 ," + strStockNo + "," + dicStock3Party[strStockNo][dtLoopStart].name + "," + dbForVal + "," + dbclose);
                                }
                                if (dbHedgeVal >= dbPercent || dbHedgeVal <= -dbPercent)
                                {
                                    CandidateParty3Stock.Add(strStockNo);
                                    Console.WriteLine("避 ," + strStockNo + "," + dicStock3Party[strStockNo][dtLoopStart].name + "," + dbHedgeVal + "," + dbclose);
                                }
                                if (lSelf >= dbPercent || lSelf <= -dbPercent)
                                {
                                    CandidateParty3Stock.Add(strStockNo);
                                    Console.WriteLine("自 ," + strStockNo + "," + dicStock3Party[strStockNo][dtLoopStart].name + "," + lSelf + "," + dbclose);
                                }

                               // string strtotal = string.Format("{0},{1},{2},{3},{4},{5}", strStockNo, dicStockNoMapping[strStockNo], dbForVal, dbTrustVal, dbSelfVal, dbHedgeVal);
                            }
                        }
                    }
                }
            }    
        }

        private void button29_Click(object sender, EventArgs e)
        {
            int iRange = 5;  //iRange天內
            double dbPercent = 500000000; // 成長數大於市值
            DateTime dtStart = monthCalendar1.SelectionStart;
            DateTime dtLoopStart = monthCalendar1.SelectionStart;

            string strMsg = string.Format("找{0}前{1}天內三大買超市值{2:F4}", dtStart.ToString("MM/dd"), iRange, dbPercent);
            Console.WriteLine(strMsg);

            CandidateParty3Stock.Clear();

            //CandidateWeekChipSet.Clear();
            string strBasePath = "D:\\Chips\\stock\\TDCC_OD_1-5_";
            string strReadFile = "";
            int icountchip = 0;
            while (true)
            {
                string strcsv = strBasePath + dtLoopStart.ToString("yyyyMMdd") + ".csv";
                if (File.Exists(strcsv))
                {
                    strReadFile = strcsv;
                    break;
                }
                dtLoopStart = dtLoopStart.AddDays(-1);
            }
            Dictionary<string, StockLevelCount> weekLevelCount = new Dictionary<string, StockLevelCount>();
            int iReadCount = 0;
            IEnumerable<string> lines = File.ReadLines(strReadFile);
            foreach (string line in lines)  // 取流通在外數
            {
                if (iReadCount > 0)
                {
                    string[] strSplitLine = line.Split(',');
                    if (weekLevelCount.ContainsKey(strSplitLine[1]))
                    {
                        int iLevel = int.Parse(strSplitLine[2]) - 1;
                        weekLevelCount[strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                        weekLevelCount[strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                        weekLevelCount[strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                    }
                    else
                    {
                        weekLevelCount.Add(strSplitLine[1], new StockLevelCount());
                        int iLevel = int.Parse(strSplitLine[2]) - 1;
                        weekLevelCount[strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                        weekLevelCount[strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                        weekLevelCount[strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                    }
                }
                iReadCount++;
            }

            Dictionary<string, long> stock3PartyVol = new Dictionary<string, long>();
            foreach (string strStockNo in weekLevelCount.Keys)
            {
                double l3PTotal = 0;
                double lHedgeTotal = 0;
                double lTrust = 0;
                double lSelf = 0;
                double lTotal = 0;
                double lFor = 0;

                long ladd = 0;
                int iLoopDay = 0;
                dtLoopStart = dtStart;
                while (iLoopDay < iRange)  // 計算最近幾天
                {
                    if (workingdays.Contains(dtLoopStart))
                    {
                        if (dicStock3Party.ContainsKey(strStockNo))
                        {
                            if (dicStock3Party[strStockNo].ContainsKey(dtLoopStart))
                            {
                                if(dbEvenydayOLHCDic.ContainsKey(strStockNo) && dbEvenydayOLHCDic[strStockNo].ContainsKey(dtLoopStart))
                                {
                                    int iTotal = int.Parse(dicStock3Party[strStockNo][dtLoopStart].Party3Total, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                    int iHedge = int.Parse(dicStock3Party[strStockNo][dtLoopStart].SelfHedgeTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                    int iHedgeBuy = int.Parse(dicStock3Party[strStockNo][dtLoopStart].SelfHedgeBuy, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                    int iHedgeSell = int.Parse(dicStock3Party[strStockNo][dtLoopStart].SelfHedgeSell, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                    int iForeigen = int.Parse(dicStock3Party[strStockNo][dtLoopStart].ForeigenTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                    int iTrust = int.Parse(dicStock3Party[strStockNo][dtLoopStart].TrustTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                    int iSelf = int.Parse(dicStock3Party[strStockNo][dtLoopStart].SelfSelfTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                    ladd += (iTotal - iHedge); // 計算不含避險3大買量
                                    //ladd += iTotal; // 全算
                                    //ladd += (iTrust + iForeigen); //外 信
                                    //ladd += iHedge;

                                    lTotal += (iTotal - iHedge) * dbEvenydayOLHCDic[strStockNo][dtLoopStart][3];
                                    l3PTotal += iTotal * dbEvenydayOLHCDic[strStockNo][dtLoopStart][3];
                                    lHedgeTotal += iHedge * dbEvenydayOLHCDic[strStockNo][dtLoopStart][3];
                                    lTrust += iTrust * dbEvenydayOLHCDic[strStockNo][dtLoopStart][3];
                                    lSelf += iSelf * dbEvenydayOLHCDic[strStockNo][dtLoopStart][3];
                                    lFor += iForeigen * dbEvenydayOLHCDic[strStockNo][dtLoopStart][3];
                                }

                            }
                        }
                        iLoopDay++;
                    }

                    dtLoopStart = dtLoopStart.AddDays(-1);
                }

                stock3PartyVol.Add(strStockNo, ladd);

                //if (weekLevelCount.ContainsKey(strStockNo))
                {
                    //double dbper = (double)stock3PartyVol[strStockNo] / (double)weekLevelCount[strStockNo].LevelVol[16];
                    if (lTrust + lFor > dbPercent && weekLevelCount[strStockNo].LevelPeople[16]<10000) // 成長數
                    {
                        CandidateParty3Stock.Add(strStockNo);
                        Console.WriteLine(strStockNo);
                    }
                }
            }    
        }

        private void button30_Click(object sender, EventArgs e)
        {
            // 券新高，人數減
            int iWeekCount = 3;
            DateTime dtStart = monthCalendar1.SelectionStart;
            DateTime dtLoopStart = monthCalendar1.SelectionStart;

            string strBasePath = "D:\\Chips\\stock\\TDCC_OD_1-5_";
            string[] strReadFile = new string[iWeekCount];

            string strMsg = string.Format("人數連{0}周減少 AND 券高", iWeekCount);
            Console.WriteLine(strMsg);
            int icountchip = 0;
            while (true)
            {
                string strcsv = strBasePath + dtLoopStart.ToString("yyyyMMdd") + ".csv";
                if (File.Exists(strcsv))
                {
                    strReadFile[icountchip] = strcsv;
                    icountchip++;
                    if (icountchip >= iWeekCount)
                        break;
                }
                dtLoopStart = dtLoopStart.AddDays(-1);
            }

            Dictionary<string, StockLevelCount>[] weekLevelCount = new Dictionary<string, StockLevelCount>[3];
            for (int i = 0; i < iWeekCount; i++)
            {
                weekLevelCount[i] = new Dictionary<string, StockLevelCount>();
                int iReadCount = 0;
                IEnumerable<string> lines = File.ReadLines(strReadFile[i]);
                foreach (string line in lines)  // 取流通在外數
                {
                    if (iReadCount > 0)
                    {
                        string[] strSplitLine = line.Split(',');
                        if (weekLevelCount[i].ContainsKey(strSplitLine[1]))
                        {
                            int iLevel = int.Parse(strSplitLine[2]) - 1;
                            weekLevelCount[i][strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                            weekLevelCount[i][strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                            weekLevelCount[i][strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                        }
                        else
                        {
                            weekLevelCount[i].Add(strSplitLine[1], new StockLevelCount());
                            int iLevel = int.Parse(strSplitLine[2]) - 1;
                            weekLevelCount[i][strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                            weekLevelCount[i][strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                            weekLevelCount[i][strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                        }
                    }
                    iReadCount++;
                }
            }
            

            foreach(string strno in weekLevelCount[0].Keys)
            {
                if(weekLevelCount[1].ContainsKey(strno) &&weekLevelCount[2].ContainsKey(strno) )
                {
                    long lP0 = weekLevelCount[0][strno].LevelPeople[16] - weekLevelCount[0][strno].LevelPeople[0] - weekLevelCount[0][strno].LevelPeople[1];
                    long lP1 = weekLevelCount[1][strno].LevelPeople[16] - weekLevelCount[1][strno].LevelPeople[0] - weekLevelCount[1][strno].LevelPeople[1];
                    long lP2 = weekLevelCount[2][strno].LevelPeople[16] - weekLevelCount[2][strno].LevelPeople[0] - weekLevelCount[2][strno].LevelPeople[1];
                    if(lP0 <= lP1 && lP1<=lP2)
                    {
                        if(dicStockMarriage.ContainsKey(strno))
                        {
                            if (dicStockMarriage[strno].Values.Max() != 0 && 
                                (dicStockMarriage[strno].Values.Max() == dicStockMarriage[strno].Values.ElementAt(0)
                                //|| dicStockMarriage[strno].Values.Max() == dicStockMarriage[strno].Values.ElementAt(1)
                                //|| dicStockMarriage[strno].Values.Max() == dicStockMarriage[strno].Values.ElementAt(2)
                                ))
                            {
                                strMsg = string.Format("{0},{1},{2}", strno,  dicStockNoMapping[strno],dicStockMarriage[strno].Values.ElementAt(0));
                                Console.WriteLine(strMsg);
                            }
                        }
                    }
                }
            }
        }

        private void button31_Click(object sender, EventArgs e)
        {
            // 季報  https://mops.twse.com.tw/mops/web/ajax_t163sb04
            string strBasePath = "d:\\Chips\\Season\\";
            int year = 108;
            for (int i = 1; i <= 4; i++)
            {
                string strGolfFile = strBasePath + string.Format("ajax_t163sb04_{0}_{1}.html", year, i);
                string Url = "https://mops.twse.com.tw/mops/web/ajax_t163sb04";
                string param = string.Format("encodeURIComponent=1&step=1&firstin=1&off=1&isQuery=Y&TYPEK=sii&year={0}&season=0{1}", year,i);
                HttpWebRequest request = HttpWebRequest.Create(Url) as HttpWebRequest;
                string result = null;
                request.Method = "POST";    // 方法
                request.KeepAlive = true; //是否保持連線
                request.ContentType = "application/x-www-form-urlencoded";

                byte[] bs = Encoding.ASCII.GetBytes(param);

                using (Stream reqStream = request.GetRequestStream())
                {
                    reqStream.Write(bs, 0, bs.Length);
                }

                using (WebResponse response = request.GetResponse())
                {
                    StreamReader sr = new StreamReader(response.GetResponseStream());
                    result = sr.ReadToEnd();
                    sr.Close();
                }
                File.WriteAllText(strGolfFile, result, Encoding.UTF8);
            }

            /*
            //  上櫃每月營收 https://mops.twse.com.tw/nas/t21/otc/t21sc03_109_9_0.html 
            string strBasePath = "d:\\Chips\\Revenue\\";
            DateTime dtStatMon = new DateTime(2014, 10, 1);
            for (int i = 0; i < 1; i++)
            {
                string strGolfFile = strBasePath + string.Format("t21sc03_{0}_{1}_otc.csv", dtStatMon.Year - 1911, dtStatMon.Month);
                string param = string.Format("step=9&functionName=show_file&filePath=/home/html/nas/t21/otc/&fileName=t21sc03_{0}_{1}.csv", dtStatMon.Year - 1911, dtStatMon.Month);
                dtStatMon = dtStatMon.AddMonths(-1);
                if (!File.Exists(strGolfFile))
                {
                    string Url = "https://mops.twse.com.tw/server-java/FileDownLoad";
                    HttpWebRequest request = HttpWebRequest.Create(Url) as HttpWebRequest;
                    string result = null;
                    request.Method = "POST";    // 方法
                    request.KeepAlive = true; //是否保持連線
                    request.ContentType = "application/x-www-form-urlencoded";



                    byte[] bs = Encoding.ASCII.GetBytes(param);

                    using (Stream reqStream = request.GetRequestStream())
                    {
                        reqStream.Write(bs, 0, bs.Length);
                    }

                    using (WebResponse response = request.GetResponse())
                    {
                        StreamReader sr = new StreamReader(response.GetResponseStream());
                        result = sr.ReadToEnd();
                        sr.Close();
                    }
                    File.WriteAllText(strGolfFile, result, Encoding.UTF8);

                    request.Abort();
                    Thread.Sleep(10000);
                }
            }*/

            /*
            int worksheetNumber = 100;
            string excelFilePath = "D:\\Chips\\Revenue\\O_202008.xls";
            var cnnStr = String.Format("Provider=Microsoft.ACE.OLEDB.16.0;Data Source={0}; Extended Properties='Excel 16.0; HDR=Yes'", excelFilePath);
            var cnn = new OleDbConnection(cnnStr);

            var dt = new DataTable();
            try
            {
                cnn.Open();
                var schemaTable = cnn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                if (schemaTable.Rows.Count < worksheetNumber) throw new ArgumentException("The worksheet number provided cannot be found in the spreadsheet");
                string worksheet = schemaTable.Rows[worksheetNumber - 1]["table_name"].ToString().Replace("'", "");
                string sql = String.Format("select * from [{0}]", worksheet);
                var da = new OleDbDataAdapter(sql, cnn);
                da.Fill(dt);
            }
            catch (Exception exx)
            {
                // ???
                throw exx;
            }
            finally
            {
                // free resources
                cnn.Close();
            }*/


            /* Revenue every month*/
            /*
            //https://mops.twse.com.tw/server-java/FileDownLoad
            //filename="t21sc03_106_1.csv"
            string strBasePath = "d:\\Chips\\Revenue\\";
            DateTime dtStatMon = new DateTime(2018, 1, 1);
            for(int i=0;i<48;i++)
            {
                string strGolfFile = strBasePath + string.Format("t21sc03_{0}_{1}.csv", dtStatMon.Year - 1911, dtStatMon.Month);
                string param = string.Format("step=9&functionName=show_file&filePath=/home/html/nas/t21/sii/&fileName=t21sc03_{0}_{1}.csv", dtStatMon.Year - 1911, dtStatMon.Month);
                dtStatMon = dtStatMon.AddMonths(-1);
                if(!File.Exists(strGolfFile))
                {
                    string Url = "https://mops.twse.com.tw/server-java/FileDownLoad";
                    HttpWebRequest request = HttpWebRequest.Create(Url) as HttpWebRequest;
                    string result = null;
                    request.Method = "POST";    // 方法
                    request.KeepAlive = true; //是否保持連線
                    request.ContentType = "application/x-www-form-urlencoded";


         
                    byte[] bs = Encoding.ASCII.GetBytes(param);

                    using (Stream reqStream = request.GetRequestStream())
                    {
                        reqStream.Write(bs, 0, bs.Length);
                    }

                    using (WebResponse response = request.GetResponse())
                    {
                        StreamReader sr = new StreamReader(response.GetResponseStream());
                        result = sr.ReadToEnd();
                        sr.Close();
                    }
                    File.WriteAllText(strGolfFile, result);

                    request.Abort();
                    Thread.Sleep(10000);
                }
                
            }*/
            

            /*HttpWebRequest request = (HttpWebRequest)WebRequest.Create("https://mops.twse.com.tw/server-java/FileDownLoad?step=9&functionName=show_file&filePath=/home/html/nas/t21/sii/&fileName=t21sc03_106_8.csv");

            HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create("https://mops.twse.com.tw/server-java/FileDownLoad?step=9&functionName=show_file&filePath=/home/html/nas/t21/sii/&fileName=t21sc03_106_8.csv");
            httpWebRequest.ContentType = "application/json";
            httpWebRequest.Method = WebRequestMethods.Http.Post;
            httpWebRequest.Accept = "application/json";


            string text;
            var response = (HttpWebResponse)httpWebRequest.GetResponse();

            using (var sr = new StreamReader(response.GetResponseStream()))
            {
                text = sr.ReadToEnd();
            }

            File.WriteAllText("a.csv", text);



            request.Abort();*/


            //期權二類 https://www.taifex.com.tw/cht/3/futAndOptDateView
    /*
                 // 3大
                 //https://www.twse.com.tw/fund/BFI82U?response=json&dayDate=20200831&type=day
                 int iCount = 0;
                 List<string> lstrdays = new List<string>();
                 DateTime dtLoop = new DateTime(2020, 8, 1);
                 for (; dtLoop < new DateTime(2020, 8, 31); dtLoop = dtLoop.AddDays(1))
                 {
                     lstrdays.Add(dtLoop.ToString("yyyyMMdd"));
                 }
                 string strSavePath = "D:\\Chips\\BFI82U\\";
                 DateTime dt = DateTime.Now;
                 foreach (string strdayinfo in lstrdays)
                 {
                     string strLocalFile = string.Format("{0}BFI82U_day_{1}.json", strSavePath, strdayinfo);
                     if (!File.Exists(strLocalFile))
                     {
                         string url = string.Format("https://www.twse.com.tw/fund/BFI82U?response=json&dayDate={0}&type=day", strdayinfo);

                         HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
                         httpWebRequest.ContentType = "application/json";
                         httpWebRequest.Method = WebRequestMethods.Http.Get;
                         httpWebRequest.Accept = "application/json";


                         string text;
                         var response = (HttpWebResponse)httpWebRequest.GetResponse();

                         using (var sr = new StreamReader(response.GetResponseStream()))
                         {
                             text = sr.ReadToEnd();
                         }

                         File.WriteAllText(strLocalFile, text);

                         response.Close();
                         httpWebRequest.Abort();
                         Thread.Sleep(2000);
                         if (iCount++ % 8 == 0)
                         {
                             Thread.Sleep(20000);
                         }
                     }


                 }*/
            /*
                    //信用交易統計
                    //https://www.twse.com.tw/exchangeReport/MI_MARGN?response=json&date=20200803&selectType=MS
                    int iCount = 0;
                    List<string> lstrdays = new List<string>();
                    DateTime dtLoop = new DateTime(2020, 7, 1);
                    for (; dtLoop < new DateTime(2020, 8,29); dtLoop = dtLoop.AddDays(1))
                    {
                        lstrdays.Add(dtLoop.ToString("yyyyMMdd"));
                    }
                    string strSavePath = "D:\\Chips\\MARGNMS\\";
                    DateTime dt = DateTime.Now;
                    foreach (string strdayinfo in lstrdays)
                    {
                        string strLocalFile = string.Format("{0}MI_MARGN_MS_{1}.json", strSavePath, strdayinfo);
                        if (!File.Exists(strLocalFile))
                        {
                            string url = string.Format("https://www.twse.com.tw/exchangeReport/MI_MARGN?response=json&date={0}&selectType=MS", strdayinfo);

                            HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
                            httpWebRequest.ContentType = "application/json";
                            httpWebRequest.Method = WebRequestMethods.Http.Get;
                            httpWebRequest.Accept = "application/json";


                            string text;
                            var response = (HttpWebResponse)httpWebRequest.GetResponse();

                            using (var sr = new StreamReader(response.GetResponseStream()))
                            {
                                text = sr.ReadToEnd();
                            }

                            File.WriteAllText(strLocalFile, text);

                            response.Close();
                            httpWebRequest.Abort();
                            Thread.Sleep(2100);
                            if(iCount++%8 ==0)
                            {
                                Thread.Sleep(7105);
                            }
                        }
      

                    }
         */
    /*
            // 每日5秒紀錄
            //https://www.twse.com.tw/exchangeReport/MI_5MINS?response=csv&date=20200731
            int iCount = 0;
            List<string> lstrdays = new List<string>();
            DateTime dtLoop = new DateTime(2020, 7, 1);
            for (; dtLoop < new DateTime(2020, 8, 29); dtLoop = dtLoop.AddDays(1))
            {
                lstrdays.Add(dtLoop.ToString("yyyyMMdd"));
            }
            string strSavePath = "D:\\Chips\\TWSE5MINS\\";
            DateTime dt = DateTime.Now;
            foreach (string strdayinfo in lstrdays)
            {
                string strLocalFile = string.Format("{0}MI_5MINS_{1}.json", strSavePath, strdayinfo);
                if (!File.Exists(strLocalFile))
                {
                    string url = string.Format("https://www.twse.com.tw/exchangeReport/MI_5MINS?response=json&date={0}", strdayinfo);

                    HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
                    httpWebRequest.ContentType = "application/json";
                    httpWebRequest.Method = WebRequestMethods.Http.Get;
                    httpWebRequest.Accept = "application/json";


                    string text;
                    var response = (HttpWebResponse)httpWebRequest.GetResponse();

                    using (var sr = new StreamReader(response.GetResponseStream()))
                    {
                        text = sr.ReadToEnd();
                    }

                    File.WriteAllText(strLocalFile, text);

                    response.Close();
                    httpWebRequest.Abort();
                    Thread.Sleep(2101);
                    if (iCount++ % 8 == 0)
                    {
                        Thread.Sleep(7105);
                    }
                }
      

            }*/
     /*
           // 這是抓權指__*
           //https://www.twse.com.tw/indicesReport/MI_5MINS_HIST?response=csv&date=20160501
            int iCount = 0;
           List<string> lstrmonth = new List<string>();
           DateTime dtLoop = new DateTime(2020, 7, 1);
           for(;dtLoop < new DateTime(2020, 8, 29);dtLoop = dtLoop.AddMonths(1))
           {
               lstrmonth.Add(dtLoop.ToString("yyyyMMdd"));
           }

           string strSavePath = "D:\\Chips\\TWSE\\";
           DateTime dt = DateTime.Now;
           foreach (string strMonthinfo in lstrmonth)
           {
               string strLocalFile = string.Format("{0}MI_5MINS_HIST_{1}.json", strSavePath, strMonthinfo);
               if (!File.Exists(strLocalFile))
               {
                   string url = string.Format("https://www.twse.com.tw/indicesReport/MI_5MINS_HIST?response=json&date={0}", strMonthinfo);

                   HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
                   httpWebRequest.ContentType = "application/json";
                   httpWebRequest.Method = WebRequestMethods.Http.Get;
                   httpWebRequest.Accept = "application/json";
    

                   string text;
                   var response = (HttpWebResponse)httpWebRequest.GetResponse();

                   using (var sr = new StreamReader(response.GetResponseStream()))
                   {
                       text = sr.ReadToEnd();
                   }

                   File.WriteAllText(strLocalFile, text);

                   response.Close();
                   httpWebRequest.Abort();

                   Thread.Sleep(2101);
                   if (iCount++ % 8 == 0)
                   {
                       Thread.Sleep(7105);
                   }

               }
           }*/
            

        }

        private void button32_Click(object sender, EventArgs e)
        {
            int iRange = 30;  //iRange天內
            int iBuydays = 10;  //iRange天內

            DateTime dtStart = monthCalendar1.SelectionStart;
            DateTime dtLoopStart = monthCalendar1.SelectionStart;

            string strMsg = string.Format("找{0}前{1}天內投信偷買{2}天 ", dtStart.ToString("MM/dd"), iRange, iBuydays);
            Console.WriteLine(strMsg);

            CandidateParty3Stock.Clear();
            //CandidateWeekChipSet.Clear();
            string strBasePath = "D:\\Chips\\stock\\TDCC_OD_1-5_";
            string strReadFile = "";
            int icountchip = 0;
            while (true)
            {
                string strcsv = strBasePath + dtLoopStart.ToString("yyyyMMdd") + ".csv";
                if (File.Exists(strcsv))
                {
                    strReadFile = strcsv;
                    break;
                }
                dtLoopStart = dtLoopStart.AddDays(-1);
            }
            Dictionary<string, StockLevelCount> weekLevelCount = new Dictionary<string, StockLevelCount>();
            int iReadCount = 0;
            IEnumerable<string> lines = File.ReadLines(strReadFile);
            foreach (string line in lines)  // 取流通在外數
            {
                if (iReadCount > 0)
                {
                    string[] strSplitLine = line.Split(',');
                    if (weekLevelCount.ContainsKey(strSplitLine[1]))
                    {
                        int iLevel = int.Parse(strSplitLine[2]) - 1;
                        weekLevelCount[strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                        weekLevelCount[strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                        weekLevelCount[strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                    }
                    else
                    {
                        weekLevelCount.Add(strSplitLine[1], new StockLevelCount());
                        int iLevel = int.Parse(strSplitLine[2]) - 1;
                        weekLevelCount[strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                        weekLevelCount[strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                        weekLevelCount[strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                    }
                }
                iReadCount++;
            }

            Dictionary<string, long> stock3PartyVol = new Dictionary<string, long>();
            foreach (string strStockNo in weekLevelCount.Keys)
            {
                double l3PTotal = 0;
                double lHedgeTotal = 0;
                double lTrust = 0;
                double lSelf = 0;
                double lTotal = 0;
                double lFor = 0;
                int iCountBuy = 0;

                int iLoopDay = 0;
                dtLoopStart = dtStart;
                while (iLoopDay < iRange)  // 計算最近幾天
                {
                    if (workingdays.Contains(dtLoopStart))
                    {
                        if (dicStock3Party.ContainsKey(strStockNo))
                        {
                            if (dicStock3Party[strStockNo].ContainsKey(dtLoopStart))
                            {
                                if (dbEvenydayOLHCDic.ContainsKey(strStockNo) && dbEvenydayOLHCDic[strStockNo].ContainsKey(dtLoopStart))
                                {
                                    int iTotal = int.Parse(dicStock3Party[strStockNo][dtLoopStart].Party3Total, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                    int iHedge = int.Parse(dicStock3Party[strStockNo][dtLoopStart].SelfHedgeTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                    int iHedgeBuy = int.Parse(dicStock3Party[strStockNo][dtLoopStart].SelfHedgeBuy, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                    int iHedgeSell = int.Parse(dicStock3Party[strStockNo][dtLoopStart].SelfHedgeSell, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                    int iForeigen = int.Parse(dicStock3Party[strStockNo][dtLoopStart].ForeigenTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                    int iTrust = int.Parse(dicStock3Party[strStockNo][dtLoopStart].TrustTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                    int iSelf = int.Parse(dicStock3Party[strStockNo][dtLoopStart].SelfSelfTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);

                                    lTotal += (iTotal - iHedge);// *dbEvenydayOLHCDic[strStockNo][dtLoopStart][3];
                                    l3PTotal += iTotal;// * dbEvenydayOLHCDic[strStockNo][dtLoopStart][3];
                                    lHedgeTotal += iHedge;// * dbEvenydayOLHCDic[strStockNo][dtLoopStart][3];
                                    lTrust += iTrust;// * dbEvenydayOLHCDic[strStockNo][dtLoopStart][3];
                                    lSelf += iSelf;// * dbEvenydayOLHCDic[strStockNo][dtLoopStart][3];
                                    lFor += iForeigen;// * dbEvenydayOLHCDic[strStockNo][dtLoopStart][3];

                                    if (iTrust > 0)
                                        iCountBuy++;
                                }

                            }
                        }
                        iLoopDay++;
                    }

                    dtLoopStart = dtLoopStart.AddDays(-1);
                }

                if (weekLevelCount.ContainsKey(strStockNo))
                {
                    double dbtrust = (double)lTrust / (double)weekLevelCount[strStockNo].LevelVol[16];
                    List<double> closelst = new List<double>();
                    if (dbEvenydayOLHCDic.ContainsKey(strStockNo) && dbEvenydayOLHCDic[strStockNo].Values.Count > 20)
                    {
                        for (int i = 0; i < 20; i++)
                        {
                            closelst.Add(dbEvenydayOLHCDic[strStockNo].Values.ElementAt(i)[3]);
                        }
                        double dbavg = closelst.Average();
                        double dbmax = closelst.Max();
                        double dbCurrent = dbEvenydayOLHCDic[strStockNo].Values.ElementAt(0)[3];
                        double dLow0 = dbEvenydayOLHCDic[strStockNo].Values.ElementAt(0)[1];
                        double dLow1 = dbEvenydayOLHCDic[strStockNo].Values.ElementAt(1)[1];
                        if (iCountBuy >= iBuydays && lTrust >= 0 && lFor >= 0 && dbtrust > 0.003 && dbCurrent >= dbmax) // 成長數
                        {
                            strMsg = string.Format("{0},{1},{2},{3:F5}", strStockNo, iCountBuy, lTrust, dbtrust);
                            Console.WriteLine(strMsg);
                            CandidateParty3Stock.Add(strStockNo);
                        }
                    }

                }



            }    
        }

        private void button33_Click(object sender, EventArgs e)
        {
            int iRange = 10;  //iRange天內
            double dbPercent = 0.005; 

            DateTime dtStart = monthCalendar1.SelectionStart;
            DateTime dtLoopStart = monthCalendar1.SelectionStart;

            string strMsg = string.Format("找前{0}天投信都沒買或賣超,{1}第一天投信買,外資&自營已買 {2:F4}", iRange, dtStart.ToString("MM/dd"), dbPercent);
            Console.WriteLine(strMsg);

            CandidateParty3Stock.Clear();
            //CandidateWeekChipSet.Clear();
            string strBasePath = "D:\\Chips\\stock\\TDCC_OD_1-5_";
            string strReadFile = "";
            int icountchip = 0;
            while (true)
            {
                string strcsv = strBasePath + dtLoopStart.ToString("yyyyMMdd") + ".csv";
                if (File.Exists(strcsv))
                {
                    strReadFile = strcsv;
                    break;
                }
                dtLoopStart = dtLoopStart.AddDays(-1);
            }
            Dictionary<string, StockLevelCount> weekLevelCount = new Dictionary<string, StockLevelCount>();
            int iReadCount = 0;
            IEnumerable<string> lines = File.ReadLines(strReadFile);
            foreach (string line in lines)  // 取流通在外數
            {
                if (iReadCount > 0)
                {
                    string[] strSplitLine = line.Split(',');
                    if (weekLevelCount.ContainsKey(strSplitLine[1]))
                    {
                        int iLevel = int.Parse(strSplitLine[2]) - 1;
                        weekLevelCount[strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                        weekLevelCount[strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                        weekLevelCount[strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                    }
                    else
                    {
                        weekLevelCount.Add(strSplitLine[1], new StockLevelCount());
                        int iLevel = int.Parse(strSplitLine[2]) - 1;
                        weekLevelCount[strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                        weekLevelCount[strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                        weekLevelCount[strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                    }
                }
                iReadCount++;
            }

            Dictionary<string, long> stock3PartyVol = new Dictionary<string, long>();
            foreach (string strStockNo in weekLevelCount.Keys)
            {
                double l3PTotal = 0;
                double lHedgeTotal = 0;
                double lTrustDay0 = 0;
                double lTrust = 0;
                double lSelf = 0;
                double lTotal = 0;
                double lFor = 0;
                int iCountBuy = 0;

                int iLoopDay = 0;
                dtLoopStart = dtStart;
                while (iLoopDay < iRange)  // 計算最近幾天
                {
                    if (workingdays.Contains(dtLoopStart))
                    {
                        if (dicStock3Party.ContainsKey(strStockNo))
                        {
                            if (dicStock3Party[strStockNo].ContainsKey(dtLoopStart))
                            {
                                if (dbEvenydayOLHCDic.ContainsKey(strStockNo) && dbEvenydayOLHCDic[strStockNo].ContainsKey(dtLoopStart))
                                {
                                    int iTotal = int.Parse(dicStock3Party[strStockNo][dtLoopStart].Party3Total, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                    int iHedge = int.Parse(dicStock3Party[strStockNo][dtLoopStart].SelfHedgeTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                    int iHedgeBuy = int.Parse(dicStock3Party[strStockNo][dtLoopStart].SelfHedgeBuy, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                    int iHedgeSell = int.Parse(dicStock3Party[strStockNo][dtLoopStart].SelfHedgeSell, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                    int iForeigen = int.Parse(dicStock3Party[strStockNo][dtLoopStart].ForeigenTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                    int iTrust = int.Parse(dicStock3Party[strStockNo][dtLoopStart].TrustTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                    int iSelf = int.Parse(dicStock3Party[strStockNo][dtLoopStart].SelfSelfTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);

                                    lTotal += (iTotal - iHedge);// *dbEvenydayOLHCDic[strStockNo][dtLoopStart][3];
                                    l3PTotal += iTotal;// * dbEvenydayOLHCDic[strStockNo][dtLoopStart][3];
                                    lHedgeTotal += iHedge;// * dbEvenydayOLHCDic[strStockNo][dtLoopStart][3];        
                                    lSelf += iSelf;// * dbEvenydayOLHCDic[strStockNo][dtLoopStart][3];
                                    lFor += iForeigen;// * dbEvenydayOLHCDic[strStockNo][dtLoopStart][3]
                                    if (iLoopDay == 0)
                                        lTrustDay0 += iTrust; 
                                    else
                                        lTrust += iTrust;// * dbEvenydayOLHCDic[strStockNo][dtLoopStart][3];
                                }
                            }
                        }
                        iLoopDay++;
                    }

                    dtLoopStart = dtLoopStart.AddDays(-1);
                }

                if (weekLevelCount.ContainsKey(strStockNo))
                {
                    double dbselffor = (double)(lSelf + lFor ) / (double)weekLevelCount[strStockNo].LevelVol[16];

                    if (dbselffor > dbPercent && lTrust <= 0 && lTrustDay0>0) // 成長數
                        {
                            strMsg = string.Format("{0},{1:F5}", strStockNo, dbselffor);
                            Console.WriteLine(strMsg);
                            CandidateParty3Stock.Add(strStockNo);
                        }
          

                }



            }    
        }

        private void button34_Click(object sender, EventArgs e)
        {

            DateTime dtStart = monthCalendar1.SelectionStart;
            DateTime dtLoopStart = monthCalendar1.SelectionStart;

            string strBasePath = "D:\\Chips\\stock\\TDCC_OD_1-5_";
            string strReadFile = "";
            int icountchip = 0;
            int iCount = 5;

            Queue<string> qcsvfile = new Queue<string>();

            while (icountchip < iCount)
            {
                string strcsv = strBasePath + dtLoopStart.ToString("yyyyMMdd") + ".csv";
                if (File.Exists(strcsv))
                {
                    qcsvfile.Enqueue(strcsv);
                    icountchip++;
                }
                dtLoopStart = dtLoopStart.AddDays(-1);
            }



     
            Dictionary<string, StockLevelCount>[] weekLevelCount = new Dictionary<string, StockLevelCount>[iCount];
            for (int i = 0; i < iCount; i++)
            {
                string strLat = qcsvfile.ElementAt(i);
                weekLevelCount[i] = new Dictionary<string, StockLevelCount>();
                int iReadCount = 0;
                IEnumerable<string> lines = File.ReadLines(strLat);
                foreach (string line in lines)
                {
                    if (iReadCount > 0)
                    {
                        string[] strSplitLine = line.Split(',');
                        if (weekLevelCount[i].ContainsKey(strSplitLine[1]))
                        {
                            int iLevel = int.Parse(strSplitLine[2]) - 1;
                            weekLevelCount[i][strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                            weekLevelCount[i][strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                            weekLevelCount[i][strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                        }
                        else
                        {
                            weekLevelCount[i].Add(strSplitLine[1], new StockLevelCount());
                            int iLevel = int.Parse(strSplitLine[2]) - 1;
                            weekLevelCount[i][strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                            weekLevelCount[i][strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                            weekLevelCount[i][strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                        }
                    }
                    iReadCount++;
                }
            }
            foreach(string strno in weekLevelCount[0].Keys)
            {
                if(weekLevelCount[1].ContainsKey(strno) && weekLevelCount[2].ContainsKey(strno) &&weekLevelCount[3].ContainsKey(strno) &&weekLevelCount[4].ContainsKey(strno) )
                {
                    double dbpeoplerate0 = (double)weekLevelCount[0][strno].LevelPeople[0] / (double)weekLevelCount[0][strno].LevelPeople[1];
                    double dbpeoplerate1 = (double)weekLevelCount[1][strno].LevelPeople[0] / (double)weekLevelCount[1][strno].LevelPeople[1];
                    double dbpeoplerate2 = (double)weekLevelCount[2][strno].LevelPeople[0] / (double)weekLevelCount[2][strno].LevelPeople[1];
                    double dbpeoplerate3 = (double)weekLevelCount[3][strno].LevelPeople[0] / (double)weekLevelCount[3][strno].LevelPeople[1];
                    double dbpeoplerate4 = (double)weekLevelCount[4][strno].LevelPeople[0] / (double)weekLevelCount[4][strno].LevelPeople[1];
                    string strmsg = string.Format("{0},{1:F4},{2:F4},{3:F4},{4:F4},{5:F4}", strno, dbpeoplerate0, dbpeoplerate1, dbpeoplerate2, dbpeoplerate3, dbpeoplerate4);
                    Console.WriteLine(strmsg);
                }

            }
        }

        private void button35_Click(object sender, EventArgs e)
        {
            // 60 20 5
            string strmsg = string.Format("號,名,外60,信60,自60,避60,Sum60,外20,信20,自20,避20,Sum20,外5,信5,自5,避5,Sum5,資,卷,卷/資,法/卷,\r\n");
            Console.WriteLine(strmsg);


            DateTime dtStart = monthCalendar1.SelectionStart;
            string strfilename=string.Format("{0}MarFin3Pty.csv",dtStart.ToString("yyyyMMdd"));
            File.AppendAllText(strfilename, strmsg, Encoding.UTF8);

            foreach (string strno in dicStock3Party.Keys)
            {
                long lFor = 0;
                long lTrust = 0;
                long lSelf = 0;
                long lHedge = 0;

                long lFor20 = 0;
                long lTrust20 = 0;
                long lSelf20 = 0;
                long lHedge20 = 0;

                long lFor5 = 0;
                long lTrust5 = 0;
                long lSelf5 = 0;
                long lHedge5 = 0;

                int iMar =0;
                int iFin =0;
                if (dicStockFinancing.ContainsKey(strno) && dicStockFinancing[strno].ContainsKey(dtStart))
                    iFin = dicStockFinancing[strno][dtStart];
                if (dicStockMarriage.ContainsKey(strno) && dicStockMarriage[strno].ContainsKey(dtStart))
                    iMar = dicStockMarriage[strno][dtStart];

                DateTime dtloopEnd = dtStart.AddDays(-(60));
                DateTime dtloop = dtStart;
                for (; dtloop>=dtloopEnd; )
                {
                    if (dicStock3Party[strno].ContainsKey(dtloop))
                    {
                        int iTotal = int.Parse(dicStock3Party[strno][dtloop].Party3Total, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                        int iHedge = int.Parse(dicStock3Party[strno][dtloop].SelfHedgeTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                        int iHedgeBuy = int.Parse(dicStock3Party[strno][dtloop].SelfHedgeBuy, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                        int iHedgeSell = int.Parse(dicStock3Party[strno][dtloop].SelfHedgeSell, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                        int iForeigen = int.Parse(dicStock3Party[strno][dtloop].ForeigenTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                        int iTrust = int.Parse(dicStock3Party[strno][dtloop].TrustTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                        int iSelf = int.Parse(dicStock3Party[strno][dtloop].SelfSelfTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);

                        lFor += iForeigen;
                        lTrust += iTrust;
                        lSelf += iSelf;
                        lHedge += iHedge;
                    }
                    dtloop = dtloop.AddDays(-1);
                }

                dtloopEnd = dtStart.AddDays(-(20));
                dtloop = dtStart;
                for (; dtloop >= dtloopEnd; )
                {
                    if (dicStock3Party[strno].ContainsKey(dtloop))
                    {
                        int iTotal = int.Parse(dicStock3Party[strno][dtloop].Party3Total, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                        int iHedge = int.Parse(dicStock3Party[strno][dtloop].SelfHedgeTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                        int iHedgeBuy = int.Parse(dicStock3Party[strno][dtloop].SelfHedgeBuy, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                        int iHedgeSell = int.Parse(dicStock3Party[strno][dtloop].SelfHedgeSell, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                        int iForeigen = int.Parse(dicStock3Party[strno][dtloop].ForeigenTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                        int iTrust = int.Parse(dicStock3Party[strno][dtloop].TrustTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                        int iSelf = int.Parse(dicStock3Party[strno][dtloop].SelfSelfTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);

                        lFor20 += iForeigen;
                        lTrust20 += iTrust;
                        lSelf20 += iSelf;
                        lHedge20 += iHedge;
                    }
                    dtloop = dtloop.AddDays(-1);
                }


                dtloopEnd = dtStart.AddDays(-(5));
                dtloop = dtStart;
                for (; dtloop >= dtloopEnd; )
                {
                    if (dicStock3Party[strno].ContainsKey(dtloop))
                    {
                        int iTotal = int.Parse(dicStock3Party[strno][dtloop].Party3Total, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                        int iHedge = int.Parse(dicStock3Party[strno][dtloop].SelfHedgeTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                        int iHedgeBuy = int.Parse(dicStock3Party[strno][dtloop].SelfHedgeBuy, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                        int iHedgeSell = int.Parse(dicStock3Party[strno][dtloop].SelfHedgeSell, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                        int iForeigen = int.Parse(dicStock3Party[strno][dtloop].ForeigenTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                        int iTrust = int.Parse(dicStock3Party[strno][dtloop].TrustTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                        int iSelf = int.Parse(dicStock3Party[strno][dtloop].SelfSelfTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);

                        lFor5 += iForeigen;
                        lTrust5 += iTrust;
                        lSelf5 += iSelf;
                        lHedge5 += iHedge;
                    }
                    dtloop = dtloop.AddDays(-1);
                }

                strmsg = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18},{19:F2},{20}\r\n", 
                    strno, dicStockNoMapping[strno], lFor, lTrust, lSelf, lHedge, lFor + lTrust + lSelf, 
                    lFor20, lTrust20, lSelf20, lHedge20, lFor20+lTrust20+lSelf20,
                    lFor5, lTrust5, lSelf5, lHedge5, lFor5+ lTrust5+ lSelf5,
                    iFin, iMar, (double)iMar / (double)iFin, (double)(lFor5 + lTrust5 + lSelf5) / (double)iMar);
                //Console.WriteLine(strmsg);
                File.AppendAllText(strfilename, strmsg);
            }
        }

        private void button36_Click(object sender, EventArgs e)
        {
            string strFile = string.Format("Y:\\CandidatePrice{0}{1}.json", DateTime.Now.Month, DateTime.Now.Day);
            if (File.Exists(strFile))
            {
                string json = File.ReadAllText(strFile);
                Dictionary<string, object> json_Dictionary = JsonConvert.DeserializeObject<Dictionary<string, object>>(json);
                foreach(string strno in json_Dictionary.Keys)
                {
                    Dictionary<string, object> json_Msgs = JsonConvert.DeserializeObject<Dictionary<string, object>>(json_Dictionary[strno].ToString());
                    if(json_Msgs.ContainsKey("msgArray"))
                    {
                        string strjsonarray = json_Msgs["msgArray"].ToString();
                        List<object> lstObj = JsonConvert.DeserializeObject<List<object>>(strjsonarray);
                        Dictionary<string, object> json_FinalMsgs = JsonConvert.DeserializeObject<Dictionary<string, object>>(lstObj[0].ToString());
                        double dbopen = JsonConvert.DeserializeObject<double>( json_FinalMsgs["o"].ToString());
                        double dbclose = JsonConvert.DeserializeObject<double>(json_FinalMsgs["c"].ToString());

                        if(dbEvenydayOLHCDic.ContainsKey(strno))
                        {
                            List<double> db60now = new List<double>();
                            db60now.Add(dbclose);
                            for (int i = 0; i < 59; i++)
                            {
                                db60now.Add(dbEvenydayOLHCDic[strno].Values.ElementAt(i)[3]);
                            }
                            if(dbclose >= db60now.Average())
                            {
                                Console.WriteLine(strno);
                            }
                        }
                    }
                    

                }
            }
        }

        private void button37_Click(object sender, EventArgs e)
        {
            // 找人數連N禮拜減價也減
            int iCount = 4;
            string strWeek = string.Format("找人數連{0}禮拜減價也減", iCount);
            Console.WriteLine(strWeek);
            DateTime dtStart = monthCalendar1.SelectionStart;
            DateTime dtLoopStart = dtStart;
            DateTime dtLoopEnd = dtStart;

            string strBasePath = "D:\\Chips\\stock\\TDCC_OD_1-5_";
            string[] strReadFile = new string[iCount];

            int icountchip = 0;
            while (true)
            {
                string strcsv = strBasePath + dtLoopStart.ToString("yyyyMMdd") + ".csv";
                if (File.Exists(strcsv))
                {
                    strReadFile[icountchip] = strcsv;
                    icountchip++;
                    if (icountchip >= iCount)
                        break;
                }
                dtLoopStart = dtLoopStart.AddDays(-1);
            }

            Dictionary<string, StockLevelCount>[] weekLevelCount = new Dictionary<string, StockLevelCount>[iCount];
            for (int i = 0; i < iCount; i++)
            {
                string strLat = strReadFile[i];
                weekLevelCount[i] = new Dictionary<string, StockLevelCount>();
                int iReadCount = 0;
                IEnumerable<string> lines = File.ReadLines(strLat);
                foreach (string line in lines)
                {
                    if (iReadCount > 0)
                    {
                        string[] strSplitLine = line.Split(',');
                        if (weekLevelCount[i].ContainsKey(strSplitLine[1]))
                        {
                            int iLevel = int.Parse(strSplitLine[2]) - 1;
                            weekLevelCount[i][strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                            weekLevelCount[i][strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                            weekLevelCount[i][strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                        }
                        else
                        {
                            weekLevelCount[i].Add(strSplitLine[1], new StockLevelCount());
                            int iLevel = int.Parse(strSplitLine[2]) - 1;
                            weekLevelCount[i][strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                            weekLevelCount[i][strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                            weekLevelCount[i][strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                        }
                    }
                    iReadCount++;
                }
            }
            foreach(string strno in weekLevelCount[0].Keys)
            {
                long[] icheckPeopleLoss = new long[iCount - 1];

                for(int iChekexist=1 ; iChekexist<iCount;iChekexist++)
                {
                    icheckPeopleLoss[iChekexist - 1] = -90000000;
                    if(weekLevelCount[iChekexist].ContainsKey(strno))
                    {
                        long lastweekpeople = weekLevelCount[iChekexist][strno].LevelPeople[16] - weekLevelCount[iChekexist][strno].LevelPeople[0];
                        long thisweekpeople = weekLevelCount[iChekexist-1][strno].LevelPeople[16] - weekLevelCount[iChekexist-1][strno].LevelPeople[0];
                        icheckPeopleLoss[iChekexist - 1] = lastweekpeople - thisweekpeople;
                    }
                }
                int iCheckLoss = 0;
                for (int ic = 0; ic < iCount - 1;ic++ )
                {
                    if (icheckPeopleLoss[ic]>0)
                    {
                        iCheckLoss++;
                    }
                }
                if(iCheckLoss == iCount - 1)
                {                    
                    if(dbEvenydayOLHCDic.ContainsKey(strno))
                    {
                        if(dbEvenydayOLHCDic[strno].ContainsKey(dtLoopStart) && dbEvenydayOLHCDic[strno].ContainsKey(dtStart))
                        {
                            if(dbEvenydayOLHCDic[strno][dtLoopStart][3] >= dbEvenydayOLHCDic[strno][dtStart][3])
                            {
                                if(dicStock3Party.ContainsKey(strno))
                                {
                                    string strMsg = strno;

                                    DateTime dtloop = dtLoopStart;
                                    long lTrust = 0;
                                    long lForign = 0;
                                    long l3PTotal = 0;
                                    long lHedgeTotal = 0;
                                    long lTotal = 0;
                                    while (dtloop <= dtStart)
                                    {
                                        if (dicStock3Party[strno].ContainsKey(dtloop))
                                        {
                                            int iTotal = int.Parse(dicStock3Party[strno][dtloop].Party3Total, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                            int iHedge = int.Parse(dicStock3Party[strno][dtloop].SelfHedgeTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                            int iTrust = int.Parse(dicStock3Party[strno][dtloop].TrustTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                            int iForign = int.Parse(dicStock3Party[strno][dtloop].ForeigenTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);

                                            lTotal += (iTotal - iHedge);
                                            l3PTotal += iTotal;
                                            lHedgeTotal += iHedge;
                                            lTrust += iTrust;
                                            lForign += iForign;
                                        }
                                        dtloop = dtloop.AddDays(1);
                                    }

                                    if (lForign > 0)
                                    {
                                        strMsg += ",F";
                                    }
                                    if (lTrust > 0)
                                    {
                                        strMsg += ",T";
                                    }
                                    for (int i = 0; i < iCount - 1;i++ )
                                    {
                                         strMsg += string.Format(",{0}",icheckPeopleLoss[i]);
                                    }                                       
                                    Console.WriteLine(strMsg);
                                }

                            }
                        }
                    }
                }
            }

            int m = 0;
        }

        private void button38_Click(object sender, EventArgs e)
        {
            string strsn = textBoxStockPeroid.Text;
            int istartyear = int.Parse(textBoxFromYear.Text);
            int istartmonth = int.Parse(textBoxFromMonth.Text);
            int istartday = int.Parse(textBoxFromDay.Text);
            int itoyear = int.Parse(textBoxToYear.Text);
            int itomonth = int.Parse(textBoxToMonth.Text);
            int itoday = int.Parse(textBoxToDay.Text);
            int iMonth = int.Parse(comboBox1.Text);

            DateTime dtStart = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);
            DateTime dtEnd = dtStart.AddMonths(-iMonth);
            DateTime dtloop = dtEnd;

            DateTime dtLoopStart = DateTime.Now;
            string strBasePath = "D:\\Chips\\stock\\TDCC_OD_1-5_";
            string strReadFile = "";
            int icountchip = 0;
            while (true)
            {
                string strcsv = strBasePath + dtLoopStart.ToString("yyyyMMdd") + ".csv";
                if (File.Exists(strcsv))
                {
                    strReadFile = strcsv;
                    break;
                }
                dtLoopStart = dtLoopStart.AddDays(-1);
            }
            Dictionary<string, StockLevelCount> weekLevelCount = new Dictionary<string, StockLevelCount>();
            int iReadCount = 0;
            IEnumerable<string> lines = File.ReadLines(strReadFile);
            foreach (string line in lines)  // 取流通在外數
            {
                if (iReadCount > 0)
                {
                    string[] strSplitLine = line.Split(',');
                    if (weekLevelCount.ContainsKey(strSplitLine[1]))
                    {
                        int iLevel = int.Parse(strSplitLine[2]) - 1;
                        weekLevelCount[strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                        weekLevelCount[strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                        weekLevelCount[strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                    }
                    else
                    {
                        weekLevelCount.Add(strSplitLine[1], new StockLevelCount());
                        int iLevel = int.Parse(strSplitLine[2]) - 1;
                        weekLevelCount[strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                        weekLevelCount[strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                        weekLevelCount[strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                    }
                }
                iReadCount++;
            }

            long lTrust = 0;
            long lForign = 0;
            long l3PTotal = 0;
            long lHedgeTotal = 0;
            long lTotal = 0;
            while (dtloop <= dtStart)
            {
                if (dicStock3Party[strsn].ContainsKey(dtloop))
                {
                    int iTotal = int.Parse(dicStock3Party[strsn][dtloop].Party3Total, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                    int iHedge = int.Parse(dicStock3Party[strsn][dtloop].SelfHedgeTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                    int iTrust = int.Parse(dicStock3Party[strsn][dtloop].TrustTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                    int iForign = int.Parse(dicStock3Party[strsn][dtloop].ForeigenTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);

                    lTotal += (iTotal - iHedge);
                    l3PTotal += iTotal;
                    lHedgeTotal += iHedge;
                    lTrust += iTrust;
                    lForign += iForign;
                }
                dtloop = dtloop.AddDays(1);
            }

            double dbPercentTotal = (double)lTotal / (double)weekLevelCount[strsn].LevelVol[16];
            double dbPercent3PTotal = (double)l3PTotal / (double)weekLevelCount[strsn].LevelVol[16];
            double dbPercentHedgeTotal = (double)lHedgeTotal / (double)weekLevelCount[strsn].LevelVol[16];
            double dbPercentTrust = (double)lTrust / (double)weekLevelCount[strsn].LevelVol[16];
            double dbPercentForign = (double)lForign / (double)weekLevelCount[strsn].LevelVol[16];
            string strMsg = string.Format("3大{0:F4}，外{1:F4}，信{2:F4}", dbPercent3PTotal, dbPercentForign, dbPercentTrust);
            Console.WriteLine(strMsg);
            int m = 0;
        }

        private void button39_Click(object sender, EventArgs e)
        {
            Dictionary<string, string> stockname = new Dictionary<string, string>();

            Dictionary<string, long> dicForHold = new Dictionary<string, long>();
            Dictionary<string, long> dic1MonthagoForHold = new Dictionary<string, long>();
            Dictionary<string, long> dicMonthVol = new Dictionary<string, long>();


            DateTime dtnow = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);
            DateTime dtLoopStart = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);
            DateTime dt1Monthago = dtLoopStart.AddMonths(-1);
            DateTime dt12Month = dtLoopStart.AddMonths(-2);
            string strmsg = string.Format("外資自{0}/{1}/{2}至今日{3}/{4}/{5}買賣比例，及本月買賣比例，本月成交比重", dt12Month.Year, dt12Month.Month, dt12Month.Day, DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);
            Console.WriteLine(strmsg);

            for (; dtLoopStart > dt12Month; dtLoopStart = dtLoopStart.AddDays(-1))
            {
                bool bisworkday = false;
                string strFileName3Party = string.Format("D:\\Chips\\3Party\\{0:0000}{1:00}{2:00}.json", dtLoopStart.Year, dtLoopStart.Month, dtLoopStart.Day);
                if (File.Exists(strFileName3Party))
                {
                    string json = File.ReadAllText(strFileName3Party);
                    Dictionary<string, object> json_Dictionary = JsonConvert.DeserializeObject<Dictionary<string, object>>(json);
                    if (json_Dictionary.ContainsKey("data"))
                    {
                        List<List<string>> jsondata_ListList = JsonConvert.DeserializeObject<List<List<string>>>(json_Dictionary["data"].ToString());
                        foreach (List<string> marginelist in jsondata_ListList)
                        {
                            if (marginelist.ElementAt(0).Length == 4)
                            {
                                PartyParam Assign = new PartyParam();
                                Assign.Assign(marginelist);

                                if (!dicForHold.ContainsKey(Assign.number))
                                {
                                    dicForHold.Add(Assign.number, 0);
                                    dic1MonthagoForHold.Add(Assign.number, 0);
                                    dicMonthVol.Add(Assign.number, 0);
                                }
                                dicForHold[Assign.number] += int.Parse(Assign.ForeigenTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                if (dtLoopStart > dt1Monthago)
                                {
                                    dic1MonthagoForHold[Assign.number] += int.Parse(Assign.ForeigenTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                    if (dbStockVolDic.ContainsKey(Assign.number) && dbStockVolDic[Assign.number].ContainsKey(dtLoopStart))
                                        dicMonthVol[Assign.number] += dbStockVolDic[Assign.number][dtLoopStart];
                                }
                                if (!stockname.ContainsKey(Assign.number))
                                    stockname.Add(Assign.number, Assign.name);
                            }

                        }
                    }
                }
                string strFileName3PartyTpex = string.Format("D:\\Chips\\3Party\\{0:0000}{1:00}{2:00}_tpex.json", dtLoopStart.Year, dtLoopStart.Month, dtLoopStart.Day);
                if (File.Exists(strFileName3PartyTpex))
                {
                    string json = File.ReadAllText(strFileName3PartyTpex);
                    Dictionary<string, object> json_Dictionary = JsonConvert.DeserializeObject<Dictionary<string, object>>(json);
                    if (json_Dictionary.ContainsKey("aaData"))
                    {
                        List<List<string>> jsondata_ListList = JsonConvert.DeserializeObject<List<List<string>>>(json_Dictionary["aaData"].ToString());
                        foreach (List<string> marginelist in jsondata_ListList)
                        {
                            if (marginelist.ElementAt(0).Length == 4)
                            {
                                PartyParam Assign = new PartyParam();
                                Assign.AssignTPEX(marginelist);

                                if (!dicForHold.ContainsKey(Assign.number))
                                {
                                    dicForHold.Add(Assign.number, 0);
                                    dic1MonthagoForHold.Add(Assign.number, 0);
                                    dicMonthVol.Add(Assign.number, 0);
                                }
                                dicForHold[Assign.number] += int.Parse(Assign.ForeigenTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                if (dtLoopStart > dt1Monthago)
                                {
                                    dic1MonthagoForHold[Assign.number] += int.Parse(Assign.ForeigenTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                    if (dbStockVolDic.ContainsKey(Assign.number) && dbStockVolDic[Assign.number].ContainsKey(dtLoopStart))
                                        dicMonthVol[Assign.number] += dbStockVolDic[Assign.number][dtLoopStart];
                                }

                                if (!stockname.ContainsKey(Assign.number))
                                    stockname.Add(Assign.number, Assign.name);
                            }
                        }
                    }
                }
            }


            string strBasePath = "D:\\Chips\\stock\\TDCC_OD_1-5_";
            string strReadFile = "";
            dtLoopStart = dtnow;
            while (true)
            {
                string strcsv = strBasePath + dtLoopStart.ToString("yyyyMMdd") + ".csv";
                if (File.Exists(strcsv))
                {
                    strReadFile = strcsv;
                    break;
                }
                dtLoopStart = dtLoopStart.AddDays(-1);
            }
            Dictionary<string, StockLevelCount> weekLevelCount = new Dictionary<string, StockLevelCount>();
            int iReadCount = 0;
            IEnumerable<string> lines = File.ReadLines(strReadFile);
            foreach (string line in lines)  // 取流通在外數
            {
                if (iReadCount > 0)
                {
                    string[] strSplitLine = line.Split(',');
                    if (weekLevelCount.ContainsKey(strSplitLine[1]))
                    {
                        int iLevel = int.Parse(strSplitLine[2]) - 1;
                        weekLevelCount[strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                        weekLevelCount[strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                        weekLevelCount[strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                    }
                    else
                    {
                        weekLevelCount.Add(strSplitLine[1], new StockLevelCount());
                        int iLevel = int.Parse(strSplitLine[2]) - 1;
                        weekLevelCount[strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                        weekLevelCount[strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                        weekLevelCount[strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                    }
                }
                iReadCount++;
            }

            foreach (string sno in dicForHold.Keys)
            {
                if (weekLevelCount.ContainsKey(sno))
                {
                    double TotalChip = (double)weekLevelCount[sno].LevelVol[16];
                    double dbper = (double)dicForHold[sno] / (double)TotalChip;
                    double dbthismonthper = (double)dic1MonthagoForHold[sno] / (double)TotalChip;
                    double dbthismonthvol = (double)dic1MonthagoForHold[sno] / (double)dicMonthVol[sno];
                    if (sno == "2618")
                    {
                        int mm = 0;
                    }
                        
                    if (dbper < -0.03)
                    {
                        string smsg = string.Format("{0},{1},{2:F5},{3:F5},--,{4:F5}", sno, stockname[sno], dbper, dbthismonthper, dbthismonthvol);
                        Console.WriteLine(smsg);
                    }
                    else if (dbper <= 0 && dbper >= -0.03)
                    {
                        string smsg = string.Format("{0},{1},{2:F5},{3:F5},-,{4:F5}", sno, stockname[sno], dbper, dbthismonthper, dbthismonthvol);
                        Console.WriteLine(smsg);
                    }
                    else if (dbper > 0 && dbper <= 0.03)
                    {
                        string smsg = string.Format("{0},{1},{2:F5},{3:F5},*,{4:F5}", sno, stockname[sno], dbper, dbthismonthper, dbthismonthvol);
                        Console.WriteLine(smsg);
                    }
                    else if (dbper > 0.03 && dbper < 0.1)
                    {
                        string smsg = string.Format("{0},{1},{2:F5},{3:F5},**,{4:F5}", sno, stockname[sno], dbper, dbthismonthper, dbthismonthvol);
                        Console.WriteLine(smsg);
                    }
                    else if (dbper >= 0.1)
                    {
                        string smsg = string.Format("{0},{1},{2:F5},{3:F5},***,{4:F5}", sno, stockname[sno], dbper, dbthismonthper, dbthismonthvol);
                        Console.WriteLine(smsg);
                    }
                }
            }
            int m = 0;
        }

        private void button40_Click(object sender, EventArgs e)
        {
            Dictionary<string, string> stockname = new Dictionary<string, string>();

            Dictionary<string, long> dicSelfHold = new Dictionary<string, long>();
            Dictionary<string, long> dic1MonthagoSelfHold = new Dictionary<string, long>();
            Dictionary<string, long> dicMonthVol = new Dictionary<string, long>();


            DateTime dtnow = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);
            DateTime dtLoopStart = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);
            DateTime dt1Monthago = dtLoopStart.AddMonths(-1);
            DateTime dt12Month = dtLoopStart.AddMonths(-2);
            string strmsg = string.Format("自營自{0}/{1}/{2}至今日{3}/{4}/{5}買賣比例，及本月買賣比例，本月成交比重", dt12Month.Year, dt12Month.Month, dt12Month.Day, DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);
            Console.WriteLine(strmsg);

            for (; dtLoopStart > dt12Month; dtLoopStart = dtLoopStart.AddDays(-1))
            {
                bool bisworkday = false;
                string strFileName3Party = string.Format("D:\\Chips\\3Party\\{0:0000}{1:00}{2:00}.json", dtLoopStart.Year, dtLoopStart.Month, dtLoopStart.Day);
                if (File.Exists(strFileName3Party))
                {
                    string json = File.ReadAllText(strFileName3Party);
                    Dictionary<string, object> json_Dictionary = JsonConvert.DeserializeObject<Dictionary<string, object>>(json);
                    if (json_Dictionary.ContainsKey("data"))
                    {
                        List<List<string>> jsondata_ListList = JsonConvert.DeserializeObject<List<List<string>>>(json_Dictionary["data"].ToString());
                        foreach (List<string> marginelist in jsondata_ListList)
                        {
                            if (marginelist.ElementAt(0).Length == 4)
                            {
                                PartyParam Assign = new PartyParam();
                                Assign.Assign(marginelist);

                                if (!dicSelfHold.ContainsKey(Assign.number))
                                {
                                    dicSelfHold.Add(Assign.number, 0);
                                    dic1MonthagoSelfHold.Add(Assign.number, 0);
                                    dicMonthVol.Add(Assign.number, 0);
                                }
                                dicSelfHold[Assign.number] += int.Parse(Assign.SelfSelfTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                if (dtLoopStart > dt1Monthago)
                                {
                                    dic1MonthagoSelfHold[Assign.number] += int.Parse(Assign.SelfSelfTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                    if (dbStockVolDic.ContainsKey(Assign.number) && dbStockVolDic[Assign.number].ContainsKey(dtLoopStart))
                                        dicMonthVol[Assign.number] += dbStockVolDic[Assign.number][dtLoopStart];
                                }
                                if (!stockname.ContainsKey(Assign.number))
                                    stockname.Add(Assign.number, Assign.name);
                            }

                        }
                    }
                }
                string strFileName3PartyTpex = string.Format("D:\\Chips\\3Party\\{0:0000}{1:00}{2:00}_tpex.json", dtLoopStart.Year, dtLoopStart.Month, dtLoopStart.Day);
                if (File.Exists(strFileName3PartyTpex))
                {
                    string json = File.ReadAllText(strFileName3PartyTpex);
                    Dictionary<string, object> json_Dictionary = JsonConvert.DeserializeObject<Dictionary<string, object>>(json);
                    if (json_Dictionary.ContainsKey("aaData"))
                    {
                        List<List<string>> jsondata_ListList = JsonConvert.DeserializeObject<List<List<string>>>(json_Dictionary["aaData"].ToString());
                        foreach (List<string> marginelist in jsondata_ListList)
                        {
                            if (marginelist.ElementAt(0).Length == 4)
                            {
                                PartyParam Assign = new PartyParam();
                                Assign.AssignTPEX(marginelist);

                                if (!dicSelfHold.ContainsKey(Assign.number))
                                {
                                    dicSelfHold.Add(Assign.number, 0);
                                    dic1MonthagoSelfHold.Add(Assign.number, 0);
                                    dicMonthVol.Add(Assign.number, 0);
                                }
                                dicSelfHold[Assign.number] += int.Parse(Assign.SelfSelfTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                if (dtLoopStart > dt1Monthago)
                                {
                                    dic1MonthagoSelfHold[Assign.number] += int.Parse(Assign.SelfSelfTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                    if (dbStockVolDic.ContainsKey(Assign.number) && dbStockVolDic[Assign.number].ContainsKey(dtLoopStart))
                                        dicMonthVol[Assign.number] += dbStockVolDic[Assign.number][dtLoopStart];
                                }

                                if (!stockname.ContainsKey(Assign.number))
                                    stockname.Add(Assign.number, Assign.name);
                            }
                        }
                    }
                }
            }


            string strBasePath = "D:\\Chips\\stock\\TDCC_OD_1-5_";
            string strReadFile = "";
            dtLoopStart = dtnow;
            while (true)
            {
                string strcsv = strBasePath + dtLoopStart.ToString("yyyyMMdd") + ".csv";
                if (File.Exists(strcsv))
                {
                    strReadFile = strcsv;
                    break;
                }
                dtLoopStart = dtLoopStart.AddDays(-1);
            }
            Dictionary<string, StockLevelCount> weekLevelCount = new Dictionary<string, StockLevelCount>();
            int iReadCount = 0;
            IEnumerable<string> lines = File.ReadLines(strReadFile);
            foreach (string line in lines)  // 取流通在外數
            {
                if (iReadCount > 0)
                {
                    string[] strSplitLine = line.Split(',');
                    if (weekLevelCount.ContainsKey(strSplitLine[1]))
                    {
                        int iLevel = int.Parse(strSplitLine[2]) - 1;
                        weekLevelCount[strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                        weekLevelCount[strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                        weekLevelCount[strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                    }
                    else
                    {
                        weekLevelCount.Add(strSplitLine[1], new StockLevelCount());
                        int iLevel = int.Parse(strSplitLine[2]) - 1;
                        weekLevelCount[strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                        weekLevelCount[strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                        weekLevelCount[strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                    }
                }
                iReadCount++;
            }

            foreach (string sno in dicSelfHold.Keys)
            {
                if (weekLevelCount.ContainsKey(sno))
                {
                    double TotalChip = (double)weekLevelCount[sno].LevelVol[16];
                    double dbper = (double)dicSelfHold[sno] / (double)TotalChip;
                    double dbthismonthper = (double)dic1MonthagoSelfHold[sno] / (double)TotalChip;
                    double dbthismonthvol = (double)dic1MonthagoSelfHold[sno] / (double)dicMonthVol[sno];
                    if (sno == "2618")
                    {
                        int mm = 0;
                    }

                    if (dbper < -0.03)
                    {
                        string smsg = string.Format("{0},{1},{2:F5},{3:F5},--,{4:F5}", sno, stockname[sno], dbper, dbthismonthper, dbthismonthvol);
                        Console.WriteLine(smsg);
                    }
                    else if (dbper <= 0 && dbper >= -0.03)
                    {
                        string smsg = string.Format("{0},{1},{2:F5},{3:F5},-,{4:F5}", sno, stockname[sno], dbper, dbthismonthper, dbthismonthvol);
                        Console.WriteLine(smsg);
                    }
                    else if (dbper > 0 && dbper <= 0.03)
                    {
                        string smsg = string.Format("{0},{1},{2:F5},{3:F5},*,{4:F5}", sno, stockname[sno], dbper, dbthismonthper, dbthismonthvol);
                        Console.WriteLine(smsg);
                    }
                    else if (dbper > 0.03 && dbper < 0.1)
                    {
                        string smsg = string.Format("{0},{1},{2:F5},{3:F5},**,{4:F5}", sno, stockname[sno], dbper, dbthismonthper, dbthismonthvol);
                        Console.WriteLine(smsg);
                    }
                    else if (dbper >= 0.1)
                    {
                        string smsg = string.Format("{0},{1},{2:F5},{3:F5},***,{4:F5}", sno, stockname[sno], dbper, dbthismonthper, dbthismonthvol);
                        Console.WriteLine(smsg);
                    }
                }
            }
            int m = 0;
        }

        private void button41_Click(object sender, EventArgs e)
        {
            // rsi 高檔鈍化3日數
            int irsidays = 5;
            int idelday = 3;
            int iDevDays=20;
            int iTempforClose = iDevDays + 20;
            string strMsg = string.Format("計算{0}RSI，{1}日鈍化股數。 {2}日BBand上下數", irsidays, idelday, iDevDays);
            Console.WriteLine(strMsg);
            DateTime dtStart = monthCalendar1.SelectionStart;
           

            for(int itotalcount=0;itotalcount<100;itotalcount++)
            {
                DateTime dtLoopStart = dtStart;
                DateTime dtStop = dtStart.AddDays(-iTempforClose);
                int i80count = 0;
                int i20count = 0;
                int iupcount = 0;
                int ibocount = 0;
                foreach (string strno in dbEvenydayOLHCDic.Keys)
                {
                    dtLoopStart = dtStart;

                    double[] dbclose = new double[iDevDays];
                    int iFill = 0;
                    for (; iFill < iDevDays && dtLoopStart>dtStop; )
                    {
                        if (dbEvenydayOLHCDic[strno].ContainsKey(dtLoopStart))
                        {
                            dbclose[iFill] = dbEvenydayOLHCDic[strno][dtLoopStart][3];
                            iFill++;
                        }
                        dtLoopStart = dtLoopStart.AddDays(-1);
                    }
                    if (iFill == iDevDays)
                    {
                        int iOver80 = 0;
                        int iUnder20 = 0;
                        for (int iloopidle = 0; iloopidle < idelday; iloopidle++)
                        {
                            double dbup = 0;
                            double dbdn = 0;
                            for (int iloop = iloopidle; iloop < irsidays; iloop++)
                            {
                                double dbdiff = dbclose[iloop] - dbclose[iloop + 1];
                                if (dbdiff > 0)
                                    dbup += dbdiff;
                                if (dbdiff < 0)
                                    dbdn += dbdiff;
                            }
                            dbup /= irsidays;
                            dbdn /= -irsidays;
                            double rs = dbup / dbdn;
                            double rsi = 100 - (100 / (1 + rs));
                            if (rsi > 80)
                                iOver80++;
                            if (rsi < 20)
                                iUnder20++;
                        }

                        if (iOver80 == idelday)
                            i80count++;
                        if (iUnder20 == idelday)
                            i20count++;

                        double dbavg = dbclose.Average();
                        double dbdev = Math.Sqrt((double)dbclose.Average(v => Math.Pow(v - dbavg, 2)));
                        double dbupper = dbavg + (2 * dbdev);
                        double dbbotton = dbavg - (2 * dbdev);
                        if (dbclose[0] >= dbupper)
                            iupcount++;
                        if (dbclose[0] <= dbbotton)
                            ibocount++;
                    }
                }
                strMsg = string.Format("{0} RSI80:20 = {1}:{2} = {3:F2} , BB up:bo =  {4}:{5}", dtStart.ToString("MM/dd"), i80count, i20count, (double)i80count / (double)i20count, iupcount, ibocount);
                Console.WriteLine(strMsg);

                dtStart = dtStart.AddDays(-1);
            }
            
        }

        private void button42_Click(object sender, EventArgs e)
        {
            //投信連買 MA10 leg

            int iAvgDays = 20;
            int iSeasion = 60;
            int iMaxdays = 80;
            DateTime dtStart = monthCalendar1.SelectionStart;

            string strmsg = string.Format("投信{0}近2日在{1}MA打腳，{2}多排", dtStart.ToString("MM/dd"), iAvgDays, iSeasion);
            Console.WriteLine(strmsg);

            strmsg = string.Format("no,name,連買日,總買日,總賣日,總買量,月下日,月上日");
            Console.WriteLine(strmsg);

            foreach (string strno in dicStock3Party.Keys)
            {
                int iKeepBuyDays = 0;
                int iBuyday = 0;
                int iSellday = 0;
                int iBuyVol = 0;
                int iUnderAvg = 0;
                int iAboveAvg = 0;

                DateTime dtloop = dtStart;

                // Make 10 * 10MA
                double[] db10X10MA = new double[iAvgDays];

                double[] db2X60MA = new double[2];


                Queue<DateTime> tmpDays = new Queue<DateTime>();
                if (dbEvenydayOLHCDic[strno].Values.Count > iMaxdays)
                {
                    while (tmpDays.Count < iMaxdays)
                    {
                        if (dbEvenydayOLHCDic[strno].ContainsKey(dtloop))
                        {
                            tmpDays.Enqueue(dtloop);
                        }
                        dtloop = dtloop.AddDays(-1);
                    }
                    for (int icday = 0; icday < iAvgDays; icday++)
                    {
                        db10X10MA[icday] = 0;
                        for(int iloop=0;iloop< iAvgDays; iloop++)
                        {
                            db10X10MA[icday] += dbEvenydayOLHCDic[strno][tmpDays.ElementAt(icday + iloop)][3];                            
                        }
                        db10X10MA[icday] /= iAvgDays;

                        if (dbEvenydayOLHCDic[strno][tmpDays.ElementAt(icday)][3] < db10X10MA[icday])
                            iUnderAvg++;
                        if (dbEvenydayOLHCDic[strno][tmpDays.ElementAt(icday)][3] > db10X10MA[icday])
                            iAboveAvg++;
                    }
                    for (int icday = 0; icday < 2; icday++)
                    {
                        db2X60MA[icday] = 0;
                        for (int iloop = 0; iloop < iSeasion; iloop++)
                        {
                            db2X60MA[icday] += dbEvenydayOLHCDic[strno][tmpDays.ElementAt(icday + iloop)][3];
                        }
                        db2X60MA[icday] /= iSeasion;
                    }
                    for (int icday = 0; icday < iAvgDays; icday++)
                    {
                        if(dicStock3Party[strno].ContainsKey(tmpDays.ElementAt(icday)))
                        {
                            int iTrust = int.Parse(dicStock3Party[strno][tmpDays.ElementAt(icday)].TrustTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                            if (iTrust > 0)
                                iBuyday++;
                            if (iTrust < 0)
                                iSellday++;
                            if (icday == 0 && iTrust > 0)
                                iKeepBuyDays++;
                            if (iKeepBuyDays == icday && iTrust > 0)
                                iKeepBuyDays++;
                            iBuyVol += iTrust;
                        }

                    }

                    if (db2X60MA[0] > db2X60MA[1] && db10X10MA[0] >= db2X60MA[0] &&
                        (dbEvenydayOLHCDic[strno][tmpDays.ElementAt(0)][1] <= db10X10MA[0] || dbEvenydayOLHCDic[strno][tmpDays.ElementAt(1)][1] <= db10X10MA[1]))
                    {
                        int m = 0;
                        strmsg = string.Format("{0},{1},{2},{3},{4},{5},{6},{7}", strno, dicStockNoMapping[strno], iKeepBuyDays, iBuyday, iSellday, iBuyVol, iUnderAvg, iAboveAvg);
                        Console.WriteLine(strmsg);
                    }
                    
                }



            }
        }

        private void button43_Click(object sender, EventArgs e)
        {
            // 法融比 dicStock3Party / dicStockFinancing
            int iRange = 5; //iRange天內
            int iFinCount = 800;
            DateTime dtStart = monthCalendar1.SelectionStart;
            DateTime dtLoopStart = monthCalendar1.SelectionStart;

            string strMsg = string.Format("找{0}前{1}天投信外資賣比融資數", dtStart.ToString("MM/dd"), iRange);
            Console.WriteLine(strMsg);

            CandidateParty3Stock.Clear();

            //CandidateWeekChipSet.Clear();
            Dictionary<string, long> stock3PartyVol = new Dictionary<string, long>();
            foreach (string strStockNo in dicStock3Party.Keys)
            {
                long l3PTotal = 0;
                long lHedgeTotal = 0;
                long lTrust = 0;
                long lSelf = 0;
                long lTotal = 0;
                long lFor = 0;

                long ladd = 0;
                int iLoopDay = 0;
                dtLoopStart = dtStart;
                while (iLoopDay < iRange)  // 計算最近幾天
                {
                    if (workingdays.Contains(dtLoopStart))
                    {
                        if (dicStock3Party.ContainsKey(strStockNo))
                        {
                            if (dicStock3Party[strStockNo].ContainsKey(dtLoopStart))
                            {
                                int iTotal = int.Parse(dicStock3Party[strStockNo][dtLoopStart].Party3Total, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                int iHedge = int.Parse(dicStock3Party[strStockNo][dtLoopStart].SelfHedgeTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                int iHedgeBuy = int.Parse(dicStock3Party[strStockNo][dtLoopStart].SelfHedgeBuy, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                int iHedgeSell = int.Parse(dicStock3Party[strStockNo][dtLoopStart].SelfHedgeSell, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                int iForeigen = int.Parse(dicStock3Party[strStockNo][dtLoopStart].ForeigenTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                int iTrust = int.Parse(dicStock3Party[strStockNo][dtLoopStart].TrustTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                int iSelf = int.Parse(dicStock3Party[strStockNo][dtLoopStart].SelfSelfTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                ladd += (iTotal - iHedge); // 計算不含避險3大買量
                                //ladd += iTotal; // 全算
                                //ladd += (iTrust + iForeigen); //外 信
                                //ladd += iHedge;

                                lTotal += (iTotal - iHedge);
                                l3PTotal += iTotal;
                                lHedgeTotal += iHedge;
                                lTrust += iTrust;
                                lSelf += iSelf;
                                lFor += iForeigen;
                            }
                        }
                        iLoopDay++;
                    }

                    dtLoopStart = dtLoopStart.AddDays(-1);
                }

                stock3PartyVol.Add(strStockNo, lFor + lTrust);

                if (dicStockFinancing.ContainsKey(strStockNo) && dicStockFinancing[strStockNo].ContainsKey(dtStart))
                {
                    double dbper = (double)((lFor + lTrust) / 1000) / (double)dicStockFinancing[strStockNo][dtStart];
                    if (dbper < -1.5 && dicStockFinancing[strStockNo][dtStart] > iFinCount)
                    {
                        CandidateParty3Stock.Add(strStockNo);
                        string strmsg;
                        if (dicStock3Party[strStockNo].ContainsKey(dtStart))
                        {
                            strmsg = string.Format("{0},{1},{2}", strStockNo, dbper, dicStock3Party[strStockNo][dtStart].name);
                        }
                        else
                        {
                            strmsg = string.Format("{0},{1},{2}", strStockNo, dbper, dicStock3Party[strStockNo][dtStart.AddDays(-1)].name);
                        }
                        Console.WriteLine(strmsg);
                    }
                }
            }    
        }

        private void button44_Click(object sender, EventArgs e)
        {
            int iRange = 60; // 60天內有跳空上漲形成關鍵價
            DateTime dtStart = monthCalendar1.SelectionStart;
            int iMaxContainDays = 300;
            DateTime dtfail = DateTime.Now;
            foreach (string strno in dbEvenydayOLHCDic.Keys)
            {
          
                Dictionary<int, DateTime> backDaysmapDatetime = new Dictionary<int, DateTime>();
                double[] dbOpenArray = new double[iMaxContainDays];
                double[] dbLowArray = new double[iMaxContainDays];
                double[] dbHighArray = new double[iMaxContainDays];
                double[] dbCloseArray = new double[iMaxContainDays];
                long[] lvol = new long[iMaxContainDays];
                for (int i = 0; i < iMaxContainDays; i++)
                {
                    dbOpenArray[i] = 0;
                    dbLowArray[i] = 0;
                    dbHighArray[i] = 0;
                    dbCloseArray[i] = 0;
                    lvol[i] = 0;
                }
                int iloop = 0;
                DateTime dtchk = dtfail;
                foreach (DateTime dtloop in dbEvenydayOLHCDic[strno].Keys)
                {
                    if (dtchk != dtfail)
                    {
                        // 確認時間是照順序排
                        if (dtchk <= dtloop)
                        {
                            MessageBox.Show("False");
                        }
                    }
                    dtchk = dtloop;
                    if (dtchk <= dtStart)
                    {
                        backDaysmapDatetime.Add(iloop, dtloop);
                        lvol[iloop] = dbStockVolDic[strno][dtchk];
                        dbOpenArray[iloop] = dbEvenydayOLHCDic[strno][dtchk][0];
                        dbLowArray[iloop] = dbEvenydayOLHCDic[strno][dtchk][1];
                        dbHighArray[iloop] = dbEvenydayOLHCDic[strno][dtchk][2];
                        dbCloseArray[iloop] = dbEvenydayOLHCDic[strno][dtchk][3];
                        iloop++;
                        if (iloop == iMaxContainDays)
                            break;
                    }
                }
                double[] db5ary = new double[50];
                double[] db10ary = new double[50];
                double[] db20ary = new double[50];
                double[] db60ary = new double[50];
                double[] db120ary = new double[50];
                double[] db200ary = new double[50];
                double[] db240ary = new double[50];

                for (int icrossdays = 0; icrossdays < 50; icrossdays++)
                {
                    db5ary[icrossdays] = dbCloseArray.Skip(icrossdays).Take(5).Average();
                    db10ary[icrossdays] = dbCloseArray.Skip(icrossdays).Take(10).Average();
                    db20ary[icrossdays] = dbCloseArray.Skip(icrossdays).Take(20).Average();
                    db60ary[icrossdays] = dbCloseArray.Skip(icrossdays).Take(60).Average();
                    db120ary[icrossdays] = dbCloseArray.Skip(icrossdays).Take(120).Average();
                    db200ary[icrossdays] = dbCloseArray.Skip(icrossdays).Take(200).Average();
                    db240ary[icrossdays] = dbCloseArray.Skip(icrossdays).Take(240).Average();
                }

                if (dicStock3Party.ContainsKey(strno) && dicStock3Party[strno].ContainsKey(dtStart))
                {
                    int iTotal = int.Parse(dicStock3Party[strno][dtStart].Party3Total, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                    int iHedge = int.Parse(dicStock3Party[strno][dtStart].SelfHedgeTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                    int itruet = int.Parse(dicStock3Party[strno][dtStart].TrustTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                    int iself = int.Parse(dicStock3Party[strno][dtStart].SelfSelfTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                    int ifor = int.Parse(dicStock3Party[strno][dtStart].ForeigenTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);

                    if (dicStockNoMapping.ContainsKey(strno) &&
                        db60ary[0] > db60ary[1] && dbCloseArray[0] > db60ary[0] &&
                       dbCloseArray[1] < db60ary[1])
                    {
                        string strmsg = string.Format("{0},{1},{2},{3},{4},{5},{6}", strno, dicStockNoMapping[strno], ifor, itruet, iHedge, iself, iTotal);
                        Console.WriteLine(strmsg);
                    }
                }

                
            }
        }

        private void button45_Click(object sender, EventArgs e)
        {
            // 黑K信買
            DateTime dtCheckNow = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day - 1);
            DateTime dtStart = monthCalendar1.SelectionStart;
            DateTime dtLoopStart = dtStart.AddDays(-1);
            DateTime dtLoopEnd = dtLoopStart.AddMonths(-1);

            string strBasePath = "D:\\Chips\\stock\\TDCC_OD_1-5_";
            int iCount = 10;
            Dictionary<string, StockLevelCount>[] weekLevelCount = new Dictionary<string, StockLevelCount>[iCount];
            int iCountWeek = 0;
            while (dtLoopStart > dtLoopEnd)
            {
                string strcsv = strBasePath + dtLoopStart.ToString("yyyyMMdd") + ".csv";
                if (File.Exists(strcsv))
                {
                    weekLevelCount[iCountWeek] = new Dictionary<string, StockLevelCount>();
                    int iReadCount = 0;
                    IEnumerable<string> lines = File.ReadLines(strcsv);
                    foreach (string line in lines)
                    {
                        if (iReadCount > 0)
                        {
                            string[] strSplitLine = line.Split(',');
                            if (weekLevelCount[iCountWeek].ContainsKey(strSplitLine[1]))
                            {
                                int iLevel = int.Parse(strSplitLine[2]) - 1;
                                weekLevelCount[iCountWeek][strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                                weekLevelCount[iCountWeek][strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                                weekLevelCount[iCountWeek][strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                            }
                            else
                            {
                                weekLevelCount[iCountWeek].Add(strSplitLine[1], new StockLevelCount());
                                int iLevel = int.Parse(strSplitLine[2]) - 1;
                                weekLevelCount[iCountWeek][strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                                weekLevelCount[iCountWeek][strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                                weekLevelCount[iCountWeek][strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                            }
                        }
                        iReadCount++;
                    }
                    iCountWeek++;
                }
                dtLoopStart = dtLoopStart.AddDays(-1);
            }

            foreach (string strno in dbEvenydayOLHCDic.Keys)
            {
                dtLoopStart = dtStart;
                if (dbEvenydayOLHCDic[strno].ContainsKey(dtLoopStart) && dbEvenydayOLHCDic[strno].Count > 2)
                {
                    DateTime dtnext = dtLoopStart.AddDays(-1);
                    while (!dbEvenydayOLHCDic[strno].ContainsKey(dtnext) && dtnext>dbEvenydayOLHCDic[strno].Keys.Min())
                    {
                        dtnext = dtnext.AddDays(-1);
                    }




                    if (strno == "6491")
                    {
                        int m = 9;
                    }
                    if (dicStock3Party.ContainsKey(strno) && dicStock3Party[strno].ContainsKey(dtLoopStart) &&
                        dbEvenydayOLHCDic[strno][dtLoopStart][3] > dbEvenydayOLHCDic[strno][dtLoopStart][1] && //今收>低(下影線)
                        dbEvenydayOLHCDic[strno][dtLoopStart][0]>=dbEvenydayOLHCDic[strno][dtLoopStart][3] &&  //今開>=收(黑K)
                        //dbEvenydayOLHCDic[strno][dtnext][3] >= dbEvenydayOLHCDic[strno][dtLoopStart][3] &&     //昨收>=今收
                        dbEvenydayOLHCDic[strno].Values.Count > 60
                       // weekLevelCount[0].ContainsKey(strno) && weekLevelCount[1].ContainsKey(strno)// &&
                        //(weekLevelCount[0][strno].LevelPeople[16] - weekLevelCount[0][strno].LevelPeople[0] - weekLevelCount[0][strno].LevelPeople[1]) < (weekLevelCount[1][strno].LevelPeople[16] - weekLevelCount[1][strno].LevelPeople[0] - weekLevelCount[1][strno].LevelPeople[1])
                       )      //跌
                    {

                        double[] closearys = new double[60];
                        double[] lowarys = new double[60];
                        for (int i = 0; i < 60; i++)
                        {
                            closearys[i] = dbEvenydayOLHCDic[strno].Values.ElementAt(i)[3];
                            lowarys[i] = dbEvenydayOLHCDic[strno].Values.ElementAt(i)[1];
                        }
                        double db20ma = closearys.Take(20).Average();
                        double db60ma = closearys.Take(60).Average();
                        double db60ma1 = closearys.Skip(1).Take(60).Average();

                        int iTotal = int.Parse(dicStock3Party[strno][dtLoopStart].Party3Total, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                        int iHedge = int.Parse(dicStock3Party[strno][dtLoopStart].SelfHedgeTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                        int itruet = int.Parse(dicStock3Party[strno][dtLoopStart].TrustTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                        int iself = int.Parse(dicStock3Party[strno][dtLoopStart].SelfSelfTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                        int ifor = int.Parse(dicStock3Party[strno][dtLoopStart].ForeigenTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);

                        if (itruet > 0 && ifor > 0)// && lowarys[0] <= db20ma && closearys[0] >= db20ma && closearys[0] >= db60ma && db60ma >= db60ma1)
                        {
                            string strmsg = string.Format("{0},{1},{2},{3},{4}    now:{5}", strno, dicStockNoMapping[strno], dbEvenydayOLHCDic[strno][dtStart][3], ifor, itruet, dbEvenydayOLHCDic[strno][dtCheckNow][3]);
                            Console.WriteLine(strmsg);
                        }
                    }
                }

            }
        }

        private void button46_Click(object sender, EventArgs e)
        {
            DateTime dtStart = monthCalendar1.SelectionStart;
            long sumTotal = 0;
            long sumHedge = 0;
            long sumTrust = 0;
            long sumSelf = 0;
            long sumFor = 0;

            foreach (string strno in dicStock3Party.Keys)
            {
                if(dicStock3Party[strno].ContainsKey(dtStart))
                {
                    int iTotal = int.Parse(dicStock3Party[strno][dtStart].Party3Total, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                    int iHedge = int.Parse(dicStock3Party[strno][dtStart].SelfHedgeTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                    int iTrust = int.Parse(dicStock3Party[strno][dtStart].TrustTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                    int iSelf = int.Parse(dicStock3Party[strno][dtStart].SelfSelfTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                    int iFor = int.Parse(dicStock3Party[strno][dtStart].ForeigenTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                    sumTotal += iTotal;
                    sumHedge += iHedge;
                    sumTrust += iTrust;
                    sumSelf += iSelf;
                    sumFor += iFor;
                }

            }

            string strMsg = string.Format("{0:n0} : {1:n0} : {2:n0} : {3:n0} : {4:n0}", sumTotal, sumFor, sumTrust, sumSelf, sumHedge);
            Console.WriteLine(strMsg);
        }

        private void button47_Click(object sender, EventArgs e)
        {
            DateTime dtStart = monthCalendar1.SelectionStart;


            int iMaxloopDays = 201;
            DateTime dtfail = DateTime.Now;
            foreach(string strno in dbEvenydayOLHCDic.Keys)
            {
                Dictionary<int, DateTime> backDaysmapDatetime = new Dictionary<int, DateTime>();
                if (dbEvenydayOLHCDic[strno].Count > iMaxloopDays)
                {
                    double[] dbOpenArray = new double[iMaxloopDays];
                    double[] dbLowArray = new double[iMaxloopDays];
                    double[] dbHighArray = new double[iMaxloopDays];
                    double[] dbCloseArray = new double[iMaxloopDays];
                    long[] lvol = new long[iMaxloopDays];
                    for(int i=0;i<iMaxloopDays;i++)
                    {
                        dbOpenArray[i] = 0;
                        dbLowArray[i] = 0;
                        dbHighArray[i] = 0;
                        dbCloseArray[i] = 0;
                        lvol[i] = 0;
                    }
                    int iloop = 0;
                    DateTime dtchk = dtfail;
                    foreach (DateTime dtloop in dbEvenydayOLHCDic[strno].Keys)
                    {
                        if (dtchk != dtfail)
                        {
                            // 確認時間是照順序排
                            if (dtchk <= dtloop)
                            {
                                MessageBox.Show("False");
                            }
                        }
                        dtchk = dtloop;
                        if (dtchk <= dtStart)
                        {
                            backDaysmapDatetime.Add(iloop, dtloop);
                            lvol[iloop] = dbStockVolDic[strno][dtchk];
                            dbOpenArray[iloop] = dbEvenydayOLHCDic[strno][dtchk][0];
                            dbLowArray[iloop] = dbEvenydayOLHCDic[strno][dtchk][1];
                            dbHighArray[iloop] = dbEvenydayOLHCDic[strno][dtchk][2];
                            dbCloseArray[iloop] = dbEvenydayOLHCDic[strno][dtchk][3];
                            iloop++;
                            if (iloop == iMaxloopDays)
                                break;
                        }
                    }


                    if(iloop == iMaxloopDays)
                    {
                        double db120_0 = dbCloseArray.Skip(0).Take(120).Average();
                        double db120_1 = dbCloseArray.Skip(1).Take(120).Average();
                        if (db120_0 > db120_1 && dicStockNoMapping.ContainsKey(strno))
                        {

                            // 前跳空紅支撐
                            for (int i = 1; i < 120; i++)
                            {
                                if (dbCloseArray[i] > dbOpenArray[i] && dbOpenArray[i] > dbHighArray[i + 1] &&// 跳空
                                        dbCloseArray[i] == dbHighArray[i])// 實紅
                                {

                                    if (dbLowArray[0] < dbOpenArray[i] && dbCloseArray[0] > dbOpenArray[i])
                                    {
                                        string strmsg = string.Format("{0},{1},{2}", strno, dicStockNoMapping[strno], backDaysmapDatetime[i]);
                                        Console.WriteLine(strmsg);
                                    }

                                }
                            }
                        }



                        /* 翻多回測近三天
                        double db200_0 = dbCloseArray.Skip(0).Take(200).Average();
                        double db200_1 = dbCloseArray.Skip(1).Take(200).Average();
                        long lminvol = lvol.Min();
                        double[] dbavg60 = new double[20];
                        for (int i = 0; i < 20; i++)
                        {
                            dbavg60[i] = dbCloseArray.Skip(i).Take(60).Average();
                        }
                        if (lminvol>1000 && dbavg60[19] > dbavg60[18] && dbavg60[0] > dbavg60[1] && dbCloseArray[0] > dbavg60[0] &&
                            (dbLowArray[0] <= dbavg60[0] || dbLowArray[1] <= dbavg60[1] || dbLowArray[2] <= dbavg60[2]) &&
                            db200_0 >= db200_1)
                        {
                            if (dicStockNoMapping.ContainsKey(strno))
                            {

                                string strmsg = string.Format("{0},{1}", strno, dicStockNoMapping[strno]);
                                Console.WriteLine(strmsg);
                            }
                        }*/
                    }
                }
            }
        }

        private void button48_Click(object sender, EventArgs e)
        {
            DateTime dtStart = DateTime.Now;
            DateTime dtLoopEnd = dtStart.AddMonths(12);
            DateTime dtLoopStart = dtStart;


            //Queue<string> qcsvfile = new Queue<string>();
            string strBasePathNew = "D:\\Chips\\stock\\TDCC_OD_1-5_20191220.csv";
            string strBasePathOld = "D:\\Chips\\stock\\TDCC_OD_1-5_20181214.csv";


            Dictionary<string, StockLevelCount> weekLevelCountNew = new Dictionary<string, StockLevelCount>();
            Dictionary<string, StockLevelCount> weekLevelCountOld = new Dictionary<string, StockLevelCount>();

            if (File.Exists(strBasePathNew))
            {
                int iReadCount = 0;
                IEnumerable<string> lines = File.ReadLines(strBasePathNew);
                foreach (string line in lines)
                {
                    if (iReadCount > 0)
                    {
                        string[] strSplitLine = line.Split(',');
                        if (weekLevelCountNew.ContainsKey(strSplitLine[1]))
                        {
                            int iLevel = int.Parse(strSplitLine[2]) - 1;
                            weekLevelCountNew[strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                            weekLevelCountNew[strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                            weekLevelCountNew[strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                        }
                        else
                        {
                            weekLevelCountNew.Add(strSplitLine[1], new StockLevelCount());
                            int iLevel = int.Parse(strSplitLine[2]) - 1;
                            weekLevelCountNew[strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                            weekLevelCountNew[strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                            weekLevelCountNew[strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                        }
                    }
                    iReadCount++;
                }
            }


            if (File.Exists(strBasePathOld))
            {
                int iReadCount = 0;
                IEnumerable<string> lines = File.ReadLines(strBasePathOld);
                foreach (string line in lines)
                {
                    if (iReadCount > 0)
                    {
                        string[] strSplitLine = line.Split(',');
                        if (weekLevelCountOld.ContainsKey(strSplitLine[1]))
                        {
                            int iLevel = int.Parse(strSplitLine[2]) - 1;
                            weekLevelCountOld[strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                            weekLevelCountOld[strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                            weekLevelCountOld[strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                        }
                        else
                        {
                            weekLevelCountOld.Add(strSplitLine[1], new StockLevelCount());
                            int iLevel = int.Parse(strSplitLine[2]) - 1;
                            weekLevelCountOld[strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                            weekLevelCountOld[strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                            weekLevelCountOld[strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                        }
                    }
                    iReadCount++;
                }
            }

            int iMaxloopDays = 200;
            int iMaxContainDays = 201;
            DateTime dtfail = DateTime.Now;

            foreach(string strno in weekLevelCountNew.Keys)
            {
                if (weekLevelCountOld.ContainsKey(strno) && dbEvenydayOLHCDic.ContainsKey(strno) && dbEvenydayOLHCDic[strno].Count > iMaxContainDays)
                {
                    Dictionary<int, DateTime> backDaysmapDatetime = new Dictionary<int, DateTime>();
                    double[] dbOpenArray = new double[iMaxContainDays];
                    double[] dbLowArray = new double[iMaxContainDays];
                    double[] dbHighArray = new double[iMaxContainDays];
                    double[] dbCloseArray = new double[iMaxContainDays];
                    long[] lvol = new long[iMaxContainDays];
                    for (int i = 0; i < iMaxContainDays; i++)
                    {
                        dbOpenArray[i] = 0;
                        dbLowArray[i] = 0;
                        dbHighArray[i] = 0;
                        dbCloseArray[i] = 0;
                        lvol[i] = 0;
                    }
                    int iloop = 0;
                    DateTime dtchk = dtfail;
                    foreach (DateTime dtloop in dbEvenydayOLHCDic[strno].Keys)
                    {
                        if (dtchk != dtfail)
                        {
                            // 確認時間是照順序排
                            if (dtchk <= dtloop)
                            {
                                MessageBox.Show("False");
                            }
                        }
                        dtchk = dtloop;
                        if (dtchk <= dtStart)
                        {
                            backDaysmapDatetime.Add(iloop, dtloop);
                            lvol[iloop] = dbStockVolDic[strno][dtchk];
                            dbOpenArray[iloop] = dbEvenydayOLHCDic[strno][dtchk][0];
                            dbLowArray[iloop] = dbEvenydayOLHCDic[strno][dtchk][1];
                            dbHighArray[iloop] = dbEvenydayOLHCDic[strno][dtchk][2];
                            dbCloseArray[iloop] = dbEvenydayOLHCDic[strno][dtchk][3];
                            iloop++;
                            if (iloop == iMaxloopDays)
                                break;
                        }
                    }

                    double vol20 = lvol.Take(20).Average();

                    double db20 = dbCloseArray.Take(20).Average();
                    double db20_1 = dbCloseArray.Skip(1).Take(20).Average();

                    double db60 = dbCloseArray.Take(60).Average();
                    double db60_1 = dbCloseArray.Skip(1).Take(60).Average();

                    double db120 = dbCloseArray.Take(120).Average();
                    double db120_1 = dbCloseArray.Skip(1).Take(120).Average();

                    double db200 = dbCloseArray.Take(200).Average();
                    double db200_1 = dbCloseArray.Skip(1).Take(200).Average();

                    if(db200 >= db200_1 && db120>=db120_1 && db60>=db60_1 && db20 >=db20_1 &&
                        db20 >= db60 && db60 >= db120 && db120 >= db200 && dbCloseArray[0] > db20 )
                    {
                        // LevelRate[15] 是差異數調整，[14]:1,000,001以上，[13]:800,001-1,000,000，[12]:600,001-800,000，[11]:400,001-600,000，[10]200,001-400,000   ，[9]:100,001-200,000，[8]:50,001-100,000
                        //double dbrate0 = weekLevelCount[0][item.Key].LevelRate[10] + weekLevelCount[0][item.Key].LevelRate[11] + weekLevelCount[0][item.Key].LevelRate[12] + weekLevelCount[0][item.Key].LevelRate[13] + weekLevelCount[0][item.Key].LevelRate[14];// +weekLevelCount[0][item.Key].LevelRate[15];
                        double dbnew = weekLevelCountNew[strno].LevelRate[14] + weekLevelCountNew[strno].LevelRate[13] + weekLevelCountNew[strno].LevelRate[12] + weekLevelCountNew[strno].LevelRate[11];
                        double dbold = weekLevelCountOld[strno].LevelRate[14] + weekLevelCountOld[strno].LevelRate[13] + weekLevelCountOld[strno].LevelRate[12] + weekLevelCountOld[strno].LevelRate[11];




                        double dbdiff = (dbnew-dbold)/dbold;
                        if(dbdiff > 0.08)
                        {

                            if(dicStockNoMapping.ContainsKey(strno))
                            {
                                string strmsg = string.Format("{0},{1},{2:F2},{3:F4}", strno, dicStockNoMapping[strno], dbCloseArray[0], dbdiff);
                                Console.WriteLine(strmsg);
                            }

                        }
                    }

                }
            }

        }

        private void button49_Click(object sender, EventArgs e)
        {



            DateTime dtStart = DateTime.Now;
            DateTime dtLoopEnd = dtStart.AddMonths(12);
            DateTime dtLoopStart = dtStart;


            //Queue<string> qcsvfile = new Queue<string>();
            string strBasePathNew = "D:\\Chips\\stock\\TDCC_OD_1-5_20191220.csv";
            string strBasePathOld = "D:\\Chips\\stock\\TDCC_OD_1-5_20181214.csv";


            Dictionary<string, StockLevelCount> weekLevelCountNew = new Dictionary<string, StockLevelCount>();
            Dictionary<string, StockLevelCount> weekLevelCountOld = new Dictionary<string, StockLevelCount>();

            if (File.Exists(strBasePathNew))
            {
                int iReadCount = 0;
                IEnumerable<string> lines = File.ReadLines(strBasePathNew);
                foreach (string line in lines)
                {
                    if (iReadCount > 0)
                    {
                        string[] strSplitLine = line.Split(',');
                        if (weekLevelCountNew.ContainsKey(strSplitLine[1]))
                        {
                            int iLevel = int.Parse(strSplitLine[2]) - 1;
                            weekLevelCountNew[strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                            weekLevelCountNew[strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                            weekLevelCountNew[strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                        }
                        else
                        {
                            weekLevelCountNew.Add(strSplitLine[1], new StockLevelCount());
                            int iLevel = int.Parse(strSplitLine[2]) - 1;
                            weekLevelCountNew[strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                            weekLevelCountNew[strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                            weekLevelCountNew[strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                        }
                    }
                    iReadCount++;
                }
            }


            if (File.Exists(strBasePathOld))
            {
                int iReadCount = 0;
                IEnumerable<string> lines = File.ReadLines(strBasePathOld);
                foreach (string line in lines)
                {
                    if (iReadCount > 0)
                    {
                        string[] strSplitLine = line.Split(',');
                        if (weekLevelCountOld.ContainsKey(strSplitLine[1]))
                        {
                            int iLevel = int.Parse(strSplitLine[2]) - 1;
                            weekLevelCountOld[strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                            weekLevelCountOld[strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                            weekLevelCountOld[strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                        }
                        else
                        {
                            weekLevelCountOld.Add(strSplitLine[1], new StockLevelCount());
                            int iLevel = int.Parse(strSplitLine[2]) - 1;
                            weekLevelCountOld[strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                            weekLevelCountOld[strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                            weekLevelCountOld[strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                        }
                    }
                    iReadCount++;
                }
            }

            int iMaxloopDays = 200;
            int iMaxContainDays = 201;
            DateTime dtfail = DateTime.Now;

            foreach (string strno in weekLevelCountNew.Keys)
            {
                if (weekLevelCountOld.ContainsKey(strno) && dbEvenydayOLHCDic.ContainsKey(strno) && dbEvenydayOLHCDic[strno].Count > iMaxContainDays)
                {
                    Dictionary<int, DateTime> backDaysmapDatetime = new Dictionary<int, DateTime>();
                    double[] dbOpenArray = new double[iMaxContainDays];
                    double[] dbLowArray = new double[iMaxContainDays];
                    double[] dbHighArray = new double[iMaxContainDays];
                    double[] dbCloseArray = new double[iMaxContainDays];
                    long[] lvol = new long[iMaxContainDays];
                    for (int i = 0; i < iMaxContainDays; i++)
                    {
                        dbOpenArray[i] = 0;
                        dbLowArray[i] = 0;
                        dbHighArray[i] = 0;
                        dbCloseArray[i] = 0;
                        lvol[i] = 0;
                    }
                    int iloop = 0;
                    DateTime dtchk = dtfail;
                    foreach (DateTime dtloop in dbEvenydayOLHCDic[strno].Keys)
                    {
                        if (dtchk != dtfail)
                        {
                            // 確認時間是照順序排
                            if (dtchk <= dtloop)
                            {
                                MessageBox.Show("False");
                            }
                        }
                        dtchk = dtloop;
                        if (dtchk <= dtStart)
                        {
                            backDaysmapDatetime.Add(iloop, dtloop);
                            lvol[iloop] = dbStockVolDic[strno][dtchk];
                            dbOpenArray[iloop] = dbEvenydayOLHCDic[strno][dtchk][0];
                            dbLowArray[iloop] = dbEvenydayOLHCDic[strno][dtchk][1];
                            dbHighArray[iloop] = dbEvenydayOLHCDic[strno][dtchk][2];
                            dbCloseArray[iloop] = dbEvenydayOLHCDic[strno][dtchk][3];
                            iloop++;
                            if (iloop == iMaxloopDays)
                                break;
                        }
                    }

                    double vol20 = lvol.Take(20).Average();

                    double db20 = dbCloseArray.Take(20).Average();
                    double db20_1 = dbCloseArray.Skip(1).Take(20).Average();

                    double db60 = dbCloseArray.Take(60).Average();
                    double db60_1 = dbCloseArray.Skip(1).Take(60).Average();

                    double db120 = dbCloseArray.Take(120).Average();
                    double db120_1 = dbCloseArray.Skip(1).Take(120).Average();

                    double db200 = dbCloseArray.Take(200).Average();
                    double db200_1 = dbCloseArray.Skip(1).Take(200).Average();

                    if (db200 >= db200_1 && db120 >= db120_1 && db60 >= db60_1 && db20 >= db20_1 &&
                        db20 >= db60 && db60 >= db120 && db120 >= db200 )
                    {
                        // LevelPeople[16] 總人數 - [0]1-999股 - [1]:1-5張 -[2]:5-10張
                        double dbpeoplenew = weekLevelCountNew[strno].LevelPeople[16] - weekLevelCountNew[strno].LevelPeople[0];
                        double dbpeopleold = weekLevelCountOld[strno].LevelPeople[16] - weekLevelCountOld[strno].LevelPeople[0];

                        double dbnew = weekLevelCountNew[strno].LevelRate[14] + weekLevelCountNew[strno].LevelRate[13] + weekLevelCountNew[strno].LevelRate[12] + weekLevelCountNew[strno].LevelRate[11];
                        double dbold = weekLevelCountOld[strno].LevelRate[14] + weekLevelCountOld[strno].LevelRate[13] + weekLevelCountOld[strno].LevelRate[12] + weekLevelCountOld[strno].LevelRate[11];

                        double dbdiff = (dbpeopleold - dbpeoplenew) / dbpeopleold;
                        double dbdiff400 = (dbold - dbnew) / dbold;

                        if (dbdiff > 0.1 && dbdiff400 < -0.1 && vol20>1000*1000)
                        {
                            if ("3629" == strno)
                            {
                                int m = 0;
                            }
                            if (dicStockNoMapping.ContainsKey(strno))
                            {
                                string strmsg = string.Format("{0},{1},{2:F2},{3:F4},{4:F4},{5}", strno, dicStockNoMapping[strno], dbdiff, dbdiff400, dbCloseArray[0], weekLevelCountNew[strno].LevelPeople[1]);
                                Console.WriteLine(strmsg);
                            }

                        }
                    }

                }
            }
        }

        private void button50_Click(object sender, EventArgs e)
        {
            DateTime dtStart = DateTime.Now;
            DateTime dtLoopEnd = dtStart.AddMonths(-12);
            DateTime dtLoopStart = dtStart;
            int iMaxContainDays = 300;


            string strBasePath = "D:\\Chips\\stock\\TDCC_OD_1-5_";
            int iCount = 8;
            Dictionary<string, StockLevelCount>[] weekLevelCount = new Dictionary<string, StockLevelCount>[iCount];
            int iCountWeek = 0;
            while (dtLoopStart > dtLoopEnd && iCountWeek < iCount)
            {
                string strcsv = strBasePath + dtLoopStart.ToString("yyyyMMdd") + ".csv";
                if (File.Exists(strcsv))
                {
                    weekLevelCount[iCountWeek] = new Dictionary<string, StockLevelCount>();
                    int iReadCount = 0;
                    IEnumerable<string> lines = File.ReadLines(strcsv);
                    foreach (string line in lines)
                    {
                        if (iReadCount > 0)
                        {
                            string[] strSplitLine = line.Split(',');
                            if (weekLevelCount[iCountWeek].ContainsKey(strSplitLine[1]))
                            {
                                int iLevel = int.Parse(strSplitLine[2]) - 1;
                                weekLevelCount[iCountWeek][strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                                weekLevelCount[iCountWeek][strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                                weekLevelCount[iCountWeek][strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                            }
                            else
                            {
                                weekLevelCount[iCountWeek].Add(strSplitLine[1], new StockLevelCount());
                                int iLevel = int.Parse(strSplitLine[2]) - 1;
                                weekLevelCount[iCountWeek][strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                                weekLevelCount[iCountWeek][strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                                weekLevelCount[iCountWeek][strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                            }
                        }
                        iReadCount++;
                    }
                    iCountWeek++;
                }
                dtLoopStart = dtLoopStart.AddDays(-1);
            }
    


            DateTime dtfail = DateTime.Now;
            foreach (string strno in dbEvenydayOLHCDic.Keys)
            {
                Dictionary<int, DateTime> backDaysmapDatetime = new Dictionary<int, DateTime>();
                double[] dbOpenArray = new double[iMaxContainDays];
                double[] dbLowArray = new double[iMaxContainDays];
                double[] dbHighArray = new double[iMaxContainDays];
                double[] dbCloseArray = new double[iMaxContainDays];
                long[] lvol = new long[iMaxContainDays];
                for (int i = 0; i < iMaxContainDays; i++)
                {
                    dbOpenArray[i] = 0;
                    dbLowArray[i] = 0;
                    dbHighArray[i] = 0;
                    dbCloseArray[i] = 0;
                    lvol[i] = 0;
                }
                int iloop = 0;
                DateTime dtchk = dtfail;
                foreach (DateTime dtloop in dbEvenydayOLHCDic[strno].Keys)
                {
                    if (dtchk != dtfail)
                    {
                        // 確認時間是照順序排
                        if (dtchk <= dtloop)
                        {
                            MessageBox.Show("False");
                        }
                    }
                    dtchk = dtloop;
                    if (dtchk <= dtStart)
                    {
                        backDaysmapDatetime.Add(iloop, dtloop);
                        lvol[iloop] = dbStockVolDic[strno][dtchk];
                        dbOpenArray[iloop] = dbEvenydayOLHCDic[strno][dtchk][0];
                        dbLowArray[iloop] = dbEvenydayOLHCDic[strno][dtchk][1];
                        dbHighArray[iloop] = dbEvenydayOLHCDic[strno][dtchk][2];
                        dbCloseArray[iloop] = dbEvenydayOLHCDic[strno][dtchk][3];
                        iloop++;
                        if (iloop == iMaxContainDays)
                            break;
                    }
                }
                double[] db5ary = new double[50];
                double[] db10ary = new double[50];
                double[] db20ary = new double[50];
                double[] db60ary = new double[50];
                double[] db120ary = new double[50];
                double[] db200ary = new double[50];
                double[] db240ary = new double[50];

                for (int icrossdays = 0; icrossdays < 50; icrossdays++)
                {
                    db5ary[icrossdays] = dbCloseArray.Skip(icrossdays).Take(5).Average();
                    db10ary[icrossdays] = dbCloseArray.Skip(icrossdays).Take(10).Average();
                    db20ary[icrossdays] = dbCloseArray.Skip(icrossdays).Take(20).Average();
                    db60ary[icrossdays] = dbCloseArray.Skip(icrossdays).Take(60).Average();
                    db120ary[icrossdays] = dbCloseArray.Skip(icrossdays).Take(120).Average();
                    db200ary[icrossdays] = dbCloseArray.Skip(icrossdays).Take(200).Average();
                    db240ary[icrossdays] = dbCloseArray.Skip(icrossdays).Take(240).Average();
                }

                double[] dbAllLineMaxMinPercentAry = new double[50];
                for (int icrossdays = 0; icrossdays < 50; icrossdays++)
                {
                    double[] dbAllLineAry = new double[7];
                    dbAllLineAry[0] = dbHighArray[icrossdays];
                    dbAllLineAry[1] = dbLowArray[icrossdays];
                    dbAllLineAry[2] = db5ary[icrossdays];
                    dbAllLineAry[3] = db10ary[icrossdays];
                    dbAllLineAry[4] = db20ary[icrossdays];
                    dbAllLineAry[5] = db60ary[icrossdays];
                    dbAllLineAry[6] = db120ary[icrossdays];
                    //dbAllLineAry[7] = db200ary[icrossdays];
                    dbAllLineMaxMinPercentAry[icrossdays] = (dbAllLineAry.Max() - dbAllLineAry.Min()) / dbAllLineAry.Min();

                }
                int iCount0_3 = 0;
                for (int i = 0; i < 20; i++)
                {
                    if (dbAllLineMaxMinPercentAry[i] <= 0.05 && dicStockNoMapping.ContainsKey(strno))
                    {
                        iCount0_3++;


                    }
                }

                double dbopenmax = dbOpenArray.Max();
                double dbopenmin = dbOpenArray.Min();
                double dbclosemax = dbOpenArray.Max();
                double dbclosemin = dbOpenArray.Min();
                double[] dbrealKhilow = new double[] { dbopenmax, dbopenmin, dbclosemax, dbclosemin };

                double db60openmax = dbOpenArray.Take(60).Max();
                double db60openmin = dbOpenArray.Take(60).Min();
                double db60closemax = dbOpenArray.Take(60).Max();
                double db60closemin = dbOpenArray.Take(60).Min();
                double[] dbrealK60hilow = new double[] { db60openmax, db60openmin, db60closemax, db60closemin };

                double db20openmax = dbOpenArray.Take(20).Max();
                double db20openmin = dbOpenArray.Take(20).Min();
                double db20closemax = dbOpenArray.Take(20).Max();
                double db20closemin = dbOpenArray.Take(20).Min();
                double[] dbrealK20hilow = new double[] { db20openmax, db20openmin, db20closemax, db20closemin };

                double[] db4avghilow = new double[] { db5ary[0], db10ary[0], db20ary[0], db60ary[0],db120ary[0],db240ary[0] };

                double dballhi = dbHighArray.Max();
                double dballlow = dbLowArray.Min();

                double dball60hi = dbHighArray.Take(60).Max();
                double dball60low = dbLowArray.Take(60).Min();

                double dball20hi = dbHighArray.Take(20).Max();
                double dball20low = dbLowArray.Take(20).Min();


                double dbperc = (dball60hi - dball60low) / (dballhi - dballlow);
                double dbperc20 = (dball20hi - dball20low) / (dball60hi - dball60low);

                double dbpercrealK = (dbrealK60hilow.Max() - dbrealK60hilow.Min()) / (dbrealKhilow.Max() - dbrealKhilow.Min());
                double dbperc20realK = (dbrealK20hilow.Max() - dbrealK20hilow.Min()) / (dbrealK60hilow.Max() - dbrealK60hilow.Min());

                double dbavghilowpercent = (db4avghilow.Max() - db4avghilow.Min()) / db4avghilow.Min();

                if (dicStockNoMapping.ContainsKey(strno) && lvol.Average() > 500000 &&
                    dbpercrealK < 0.7 && dbperc20realK < 0.36 && dbavghilowpercent<0.041 &&
                    weekLevelCount[0].ContainsKey(strno) && weekLevelCount[7].ContainsKey(strno) &&
                    weekLevelCount[0][strno].LevelPeople[1] <= weekLevelCount[7][strno].LevelPeople[1]
                    )
                {
                    string trmsg = string.Format("{0},{1},{2:F2},{3:F5},{4:F5}", strno, dicStockNoMapping[strno], dbCloseArray[0], dbperc20realK, dbavghilowpercent);
                    Console.WriteLine(trmsg);
                }

            }
        }
		

        private void button51_Click(object sender, EventArgs e)
        {
            DateTime dtEnd = new DateTime(2018, 10, 20);
            DateTime dtStart = DateTime.Now;
            DateTime dtLoopEnd = dtStart.AddMonths(-7);
            DateTime dtLoopStart = dtStart;

            string strfilename = string.Format("{0}LVPeople.csv", dtStart.ToString("yyyyMMdd"));
            File.AppendAllText(strfilename, "\r\n", Encoding.UTF8);

            //Queue<string> qcsvfile = new Queue<string>();
            string strBasePath = "D:\\Chips\\stock\\TDCC_OD_1-5_";
            int iCount = 32;
            Dictionary<string, StockLevelCount>[] weekLevelCount = new Dictionary<string, StockLevelCount>[iCount];
            int iCountWeek = 0;
            while (dtLoopStart > dtLoopEnd)
            {
                string strcsv = strBasePath + dtLoopStart.ToString("yyyyMMdd") + ".csv";
                if (File.Exists(strcsv))
                {
                    weekLevelCount[iCountWeek] = new Dictionary<string, StockLevelCount>();
                    int iReadCount = 0;
                    IEnumerable<string> lines = File.ReadLines(strcsv);
                    foreach (string line in lines)
                    {
                        if (iReadCount > 0)
                        {
                            string[] strSplitLine = line.Split(',');
                            if (weekLevelCount[iCountWeek].ContainsKey(strSplitLine[1]))
                            {
                                int iLevel = int.Parse(strSplitLine[2]) - 1;
                                weekLevelCount[iCountWeek][strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                                weekLevelCount[iCountWeek][strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                                weekLevelCount[iCountWeek][strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                            }
                            else
                            {
                                weekLevelCount[iCountWeek].Add(strSplitLine[1], new StockLevelCount());
                                int iLevel = int.Parse(strSplitLine[2]) - 1;
                                weekLevelCount[iCountWeek][strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                                weekLevelCount[iCountWeek][strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                                weekLevelCount[iCountWeek][strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                            }
                        }
                        iReadCount++;
                    }
                    iCountWeek++;
                }
                dtLoopStart = dtLoopStart.AddDays(-1);
            }


            foreach (KeyValuePair<string, StockLevelCount> item in weekLevelCount[0])
            {
                if (weekLevelCount[1].ContainsKey(item.Key) && weekLevelCount[2].ContainsKey(item.Key) && weekLevelCount[3].ContainsKey(item.Key) && weekLevelCount[4].ContainsKey(item.Key) && weekLevelCount[5].ContainsKey(item.Key) && weekLevelCount[6].ContainsKey(item.Key) && weekLevelCount[7].ContainsKey(item.Key) &&
                    weekLevelCount[8].ContainsKey(item.Key) && weekLevelCount[9].ContainsKey(item.Key) && weekLevelCount[10].ContainsKey(item.Key) && weekLevelCount[11].ContainsKey(item.Key) && weekLevelCount[12].ContainsKey(item.Key) && weekLevelCount[13].ContainsKey(item.Key) && weekLevelCount[14].ContainsKey(item.Key)
                    && weekLevelCount[15].ContainsKey(item.Key) && weekLevelCount[16].ContainsKey(item.Key) && weekLevelCount[17].ContainsKey(item.Key) && weekLevelCount[18].ContainsKey(item.Key) && weekLevelCount[19].ContainsKey(item.Key) && weekLevelCount[20].ContainsKey(item.Key) && weekLevelCount[21].ContainsKey(item.Key)
                    && weekLevelCount[22].ContainsKey(item.Key) && weekLevelCount[23].ContainsKey(item.Key) && weekLevelCount[24].ContainsKey(item.Key))
                {
                    // LevelRate[15] 是差異數調整，[14]:1,000,001以上，[13]:800,001-1,000,000，[12]:600,001-800,000，[11]:400,001-600,000，[10]200,001-400,000   ，[9]:100,001-200,000，[8]:50,001-100,000
                    //double dbrate0 = weekLevelCount[0][item.Key].LevelRate[10] + weekLevelCount[0][item.Key].LevelRate[11] + weekLevelCount[0][item.Key].LevelRate[12] + weekLevelCount[0][item.Key].LevelRate[13] + weekLevelCount[0][item.Key].LevelRate[14];// +weekLevelCount[0][item.Key].LevelRate[15];
                    //double dbrate1 = weekLevelCount[1][item.Key].LevelRate[10] + weekLevelCount[1][item.Key].LevelRate[11] + weekLevelCount[1][item.Key].LevelRate[12] + weekLevelCount[1][item.Key].LevelRate[13] + weekLevelCount[1][item.Key].LevelRate[14];// + weekLevelCount[1][item.Key].LevelRate[15];
                    //double dbrate2 = weekLevelCount[2][item.Key].LevelRate[10] + weekLevelCount[2][item.Key].LevelRate[11] + weekLevelCount[2][item.Key].LevelRate[12] + weekLevelCount[2][item.Key].LevelRate[13] + weekLevelCount[2][item.Key].LevelRate[14];// + weekLevelCount[2][item.Key].LevelRate[15];
                    //double dbrate3 = weekLevelCount[3][item.Key].LevelRate[10] + weekLevelCount[3][item.Key].LevelRate[11] + weekLevelCount[3][item.Key].LevelRate[12] + weekLevelCount[3][item.Key].LevelRate[13] + weekLevelCount[3][item.Key].LevelRate[14];// + weekLevelCount[3][item.Key].LevelRate[15];
                    //double dbrate4 = weekLevelCount[4][item.Key].LevelRate[10] + weekLevelCount[4][item.Key].LevelRate[11] + weekLevelCount[4][item.Key].LevelRate[12] + weekLevelCount[4][item.Key].LevelRate[13] + weekLevelCount[4][item.Key].LevelRate[14];// + weekLevelCount[4][item.Key].LevelRate[15];
                    //double dbrate5 = weekLevelCount[5][item.Key].LevelRate[10] + weekLevelCount[5][item.Key].LevelRate[11] + weekLevelCount[5][item.Key].LevelRate[12] + weekLevelCount[5][item.Key].LevelRate[13] + weekLevelCount[5][item.Key].LevelRate[14];// + weekLevelCount[5][item.Key].LevelRate[15];
                    //double dbrate6 = weekLevelCount[6][item.Key].LevelRate[10] + weekLevelCount[6][item.Key].LevelRate[11] + weekLevelCount[6][item.Key].LevelRate[12] + weekLevelCount[6][item.Key].LevelRate[13] + weekLevelCount[6][item.Key].LevelRate[14];// + weekLevelCount[6][item.Key].LevelRate[15];
                    //double dbrate7 = weekLevelCount[7][item.Key].LevelRate[10] + weekLevelCount[7][item.Key].LevelRate[11] + weekLevelCount[7][item.Key].LevelRate[12] + weekLevelCount[7][item.Key].LevelRate[13] + weekLevelCount[7][item.Key].LevelRate[14];// + weekLevelCount[7][item.Key].LevelRate[15];
                    //double dbrate08 = weekLevelCount[08][item.Key].LevelRate[10] + weekLevelCount[08][item.Key].LevelRate[11] + weekLevelCount[08][item.Key].LevelRate[12] + weekLevelCount[08][item.Key].LevelRate[13] + weekLevelCount[08][item.Key].LevelRate[14];// + weekLevelCount[08][item.Key].LevelRate[15];
                    //double dbrate09 = weekLevelCount[09][item.Key].LevelRate[10] + weekLevelCount[09][item.Key].LevelRate[11] + weekLevelCount[09][item.Key].LevelRate[12] + weekLevelCount[09][item.Key].LevelRate[13] + weekLevelCount[09][item.Key].LevelRate[14];// + weekLevelCount[09][item.Key].LevelRate[15];
                    //double dbrate10 = weekLevelCount[10][item.Key].LevelRate[10] + weekLevelCount[10][item.Key].LevelRate[11] + weekLevelCount[10][item.Key].LevelRate[12] + weekLevelCount[10][item.Key].LevelRate[13] + weekLevelCount[10][item.Key].LevelRate[14];// + weekLevelCount[10][item.Key].LevelRate[15];
                    //double dbrate11 = weekLevelCount[11][item.Key].LevelRate[10] + weekLevelCount[11][item.Key].LevelRate[11] + weekLevelCount[11][item.Key].LevelRate[12] + weekLevelCount[11][item.Key].LevelRate[13] + weekLevelCount[11][item.Key].LevelRate[14];// + weekLevelCount[11][item.Key].LevelRate[15];
                    //double dbrate12 = weekLevelCount[12][item.Key].LevelRate[10] + weekLevelCount[12][item.Key].LevelRate[11] + weekLevelCount[12][item.Key].LevelRate[12] + weekLevelCount[12][item.Key].LevelRate[13] + weekLevelCount[12][item.Key].LevelRate[14];// + weekLevelCount[12][item.Key].LevelRate[15];
                    //double dbrate13 = weekLevelCount[13][item.Key].LevelRate[10] + weekLevelCount[13][item.Key].LevelRate[11] + weekLevelCount[13][item.Key].LevelRate[12] + weekLevelCount[13][item.Key].LevelRate[13] + weekLevelCount[13][item.Key].LevelRate[14];// + weekLevelCount[13][item.Key].LevelRate[15];
                    //double dbrate14 = weekLevelCount[14][item.Key].LevelRate[10] + weekLevelCount[14][item.Key].LevelRate[11] + weekLevelCount[14][item.Key].LevelRate[12] + weekLevelCount[14][item.Key].LevelRate[13] + weekLevelCount[14][item.Key].LevelRate[14];// + weekLevelCount[14][item.Key].LevelRate[15];
                    //double dbrate15 = weekLevelCount[15][item.Key].LevelRate[10] + weekLevelCount[15][item.Key].LevelRate[11] + weekLevelCount[15][item.Key].LevelRate[12] + weekLevelCount[15][item.Key].LevelRate[13] + weekLevelCount[15][item.Key].LevelRate[14];//+ weekLevelCount[15][item.Key].LevelRate[15];

                    // LevelPeople[16] 總人數 - [0]1-999股 - [1]:1-5張 -[2]:5-10張
                    double dbpeople0 = weekLevelCount[0][item.Key].LevelPeople[16] - weekLevelCount[0][item.Key].LevelPeople[0] - weekLevelCount[0][item.Key].LevelPeople[1];// - weekLevelCount[0][item.Key].LevelPeople[2];
                    double dbpeople1 = weekLevelCount[1][item.Key].LevelPeople[16] - weekLevelCount[1][item.Key].LevelPeople[0] - weekLevelCount[1][item.Key].LevelPeople[1];// - weekLevelCount[1][item.Key].LevelPeople[2];
                    double dbpeople2 = weekLevelCount[2][item.Key].LevelPeople[16] - weekLevelCount[2][item.Key].LevelPeople[0] - weekLevelCount[2][item.Key].LevelPeople[1];// - weekLevelCount[2][item.Key].LevelPeople[2];
                    double dbpeople3 = weekLevelCount[3][item.Key].LevelPeople[16] - weekLevelCount[3][item.Key].LevelPeople[0] - weekLevelCount[3][item.Key].LevelPeople[1];// - weekLevelCount[3][item.Key].LevelPeople[2];
                    double dbpeople4 = weekLevelCount[4][item.Key].LevelPeople[16] - weekLevelCount[4][item.Key].LevelPeople[0] - weekLevelCount[4][item.Key].LevelPeople[1];// - weekLevelCount[4][item.Key].LevelPeople[2];
                    double dbpeople5 = weekLevelCount[5][item.Key].LevelPeople[16] - weekLevelCount[5][item.Key].LevelPeople[0] - weekLevelCount[5][item.Key].LevelPeople[1];// - weekLevelCount[5][item.Key].LevelPeople[2];
                    double dbpeople6 = weekLevelCount[6][item.Key].LevelPeople[16] - weekLevelCount[6][item.Key].LevelPeople[0] - weekLevelCount[6][item.Key].LevelPeople[1];// - weekLevelCount[6][item.Key].LevelPeople[2];
                    double dbpeople7 = weekLevelCount[7][item.Key].LevelPeople[16] - weekLevelCount[7][item.Key].LevelPeople[0] - weekLevelCount[7][item.Key].LevelPeople[1];// -weekLevelCount[7][item.Key].LevelPeople[2];
                    double dbpeople08 = weekLevelCount[08][item.Key].LevelPeople[16] - weekLevelCount[08][item.Key].LevelPeople[0] - weekLevelCount[08][item.Key].LevelPeople[1];// - weekLevelCount[08][item.Key].LevelPeople[2];
                    double dbpeople09 = weekLevelCount[09][item.Key].LevelPeople[16] - weekLevelCount[09][item.Key].LevelPeople[0] - weekLevelCount[09][item.Key].LevelPeople[1];// - weekLevelCount[09][item.Key].LevelPeople[2];
                    double dbpeople10 = weekLevelCount[10][item.Key].LevelPeople[16] - weekLevelCount[10][item.Key].LevelPeople[0] - weekLevelCount[10][item.Key].LevelPeople[1];// - weekLevelCount[10][item.Key].LevelPeople[2];
                    double dbpeople11 = weekLevelCount[11][item.Key].LevelPeople[16] - weekLevelCount[11][item.Key].LevelPeople[0] - weekLevelCount[11][item.Key].LevelPeople[1];// - weekLevelCount[11][item.Key].LevelPeople[2];
                    double dbpeople12 = weekLevelCount[12][item.Key].LevelPeople[16] - weekLevelCount[12][item.Key].LevelPeople[0] - weekLevelCount[12][item.Key].LevelPeople[1];// - weekLevelCount[12][item.Key].LevelPeople[2];
                    double dbpeople13 = weekLevelCount[13][item.Key].LevelPeople[16] - weekLevelCount[13][item.Key].LevelPeople[0] - weekLevelCount[13][item.Key].LevelPeople[1];// - weekLevelCount[13][item.Key].LevelPeople[2];
                    double dbpeople14 = weekLevelCount[14][item.Key].LevelPeople[16] - weekLevelCount[14][item.Key].LevelPeople[0] - weekLevelCount[14][item.Key].LevelPeople[1];// - weekLevelCount[14][item.Key].LevelPeople[2];
                    double dbpeople15 = weekLevelCount[15][item.Key].LevelPeople[16] - weekLevelCount[15][item.Key].LevelPeople[0] - weekLevelCount[15][item.Key].LevelPeople[1];// -weekLevelCount[15][item.Key].LevelPeople[2];
                    double dbpeople16 = weekLevelCount[16][item.Key].LevelPeople[16] - weekLevelCount[16][item.Key].LevelPeople[0] - weekLevelCount[16][item.Key].LevelPeople[1];
                    double dbpeople17 = weekLevelCount[17][item.Key].LevelPeople[16] - weekLevelCount[17][item.Key].LevelPeople[0] - weekLevelCount[17][item.Key].LevelPeople[1];
                    double dbpeople18 = weekLevelCount[18][item.Key].LevelPeople[16] - weekLevelCount[18][item.Key].LevelPeople[0] - weekLevelCount[18][item.Key].LevelPeople[1];
                    double dbpeople19 = weekLevelCount[19][item.Key].LevelPeople[16] - weekLevelCount[19][item.Key].LevelPeople[0] - weekLevelCount[19][item.Key].LevelPeople[1];
                    double dbpeople20 = weekLevelCount[20][item.Key].LevelPeople[16] - weekLevelCount[20][item.Key].LevelPeople[0] - weekLevelCount[20][item.Key].LevelPeople[1];
                    double dbpeople21 = weekLevelCount[21][item.Key].LevelPeople[16] - weekLevelCount[21][item.Key].LevelPeople[0] - weekLevelCount[21][item.Key].LevelPeople[1];
                    double dbpeople22 = weekLevelCount[22][item.Key].LevelPeople[16] - weekLevelCount[22][item.Key].LevelPeople[0] - weekLevelCount[22][item.Key].LevelPeople[1];
                    double dbpeople23 = weekLevelCount[23][item.Key].LevelPeople[16] - weekLevelCount[23][item.Key].LevelPeople[0] - weekLevelCount[23][item.Key].LevelPeople[1];
                    double dbpeople24 = weekLevelCount[24][item.Key].LevelPeople[16] - weekLevelCount[24][item.Key].LevelPeople[0] - weekLevelCount[24][item.Key].LevelPeople[1];

                    string strf = item.Key;
                    for (int i = 0; i < 24; i++)
                    {
                        if (weekLevelCount[i] != null)
                        {
                            //double dbpeople = weekLevelCount[i][item.Key].LevelPeople[16] - weekLevelCount[i][item.Key].LevelPeople[0] - weekLevelCount[i][item.Key].LevelPeople[1];
                            string stradd = string.Format(",[0],{0:0.##},[1],{1:0.##}", weekLevelCount[i][item.Key].LevelPeople[0], weekLevelCount[i][item.Key].LevelPeople[1]);
                            strf += stradd;
                        }
                    }

                    strf += "\r\n";
                    File.AppendAllText(strfilename, strf);
                    //Console.WriteLine(strf);

                }
            }
        }

        private void button52_Click(object sender, EventArgs e)
        {

            DateTime dtStart = new DateTime(DateTime.Now.Year,DateTime.Now.Month,DateTime.Now.Day);
            DateTime dtLoopEnd = dtStart.AddMonths(-12);
            DateTime dtLoopStart = dtStart;
            Dictionary<int, DateTime> weekmapdate = new Dictionary<int, DateTime>();

            //Queue<string> qcsvfile = new Queue<string>();
            string strBasePath = "D:\\Chips\\stock\\TDCC_OD_1-5_";
            int iCount = 50;
            Dictionary<string, StockLevelCount>[] weekLevelCount = new Dictionary<string, StockLevelCount>[iCount];
            int iCountWeek = 0;
            while (dtLoopStart > dtLoopEnd)
            {
                string strcsv = strBasePath + dtLoopStart.ToString("yyyyMMdd") + ".csv";
                if (File.Exists(strcsv))
                {
                    weekmapdate.Add(iCountWeek, dtLoopStart);
                    weekLevelCount[iCountWeek] = new Dictionary<string, StockLevelCount>();
                    int iReadCount = 0;
                    IEnumerable<string> lines = File.ReadLines(strcsv);
                    foreach (string line in lines)
                    {
                        if (iReadCount > 0)
                        {
                            string[] strSplitLine = line.Split(',');
                            if (weekLevelCount[iCountWeek].ContainsKey(strSplitLine[1]))
                            {
                                int iLevel = int.Parse(strSplitLine[2]) - 1;
                                weekLevelCount[iCountWeek][strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                                weekLevelCount[iCountWeek][strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                                weekLevelCount[iCountWeek][strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                            }
                            else
                            {
                                weekLevelCount[iCountWeek].Add(strSplitLine[1], new StockLevelCount());
                                int iLevel = int.Parse(strSplitLine[2]) - 1;
                                weekLevelCount[iCountWeek][strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                                weekLevelCount[iCountWeek][strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                                weekLevelCount[iCountWeek][strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                            }
                        }
                        iReadCount++;
                    }
                    iCountWeek++;
                }
                dtLoopStart = dtLoopStart.AddDays(-1);
            }

            string strnono = "3016";
            for (int iweeks = 0; iweeks < iCountWeek-1; iweeks++)
            {
                if (weekLevelCount[iweeks].ContainsKey(strnono) && weekLevelCount[iweeks + 1].ContainsKey(strnono))
                {
                    double dbdownrate = (double)(weekLevelCount[iweeks + 1][strnono].LevelPeople[1] - weekLevelCount[iweeks][strnono].LevelPeople[1]) / (double)weekLevelCount[iweeks+1][strnono].LevelPeople[1];
                    DateTime dtCatchday = weekmapdate[iweeks];
                    while (!dbEvenydayOLHCDic[strnono].ContainsKey(dtCatchday))
                    {
                        dtCatchday = dtCatchday.AddDays(-1);
                    }
                    string strmsg = string.Format("{0} , {1} , {2} , {3} , {4}, {5}, {6}", iweeks, strnono, dbdownrate, weekLevelCount[iweeks][strnono].LevelPeople[0], weekLevelCount[iweeks][strnono].LevelPeople[1], weekLevelCount[iweeks][strnono].LevelPeople[16] - weekLevelCount[iweeks][strnono].LevelPeople[0] - weekLevelCount[iweeks][strnono].LevelPeople[1], dbEvenydayOLHCDic[strnono][dtCatchday][3]);
                    Console.WriteLine(strmsg);
                }
            }

            foreach (KeyValuePair<string, StockLevelCount> item in weekLevelCount[0])
            {
                if(weekLevelCount[0][item.Key].LevelPeople[0] > 10000 && weekLevelCount[0][item.Key].LevelPeople[0] < 20000 && 
                   weekLevelCount[0][item.Key].LevelPeople[1] > 12000 && weekLevelCount[0][item.Key].LevelPeople[1] < 22000 && 
                   weekLevelCount[0][item.Key].LevelPeople[16]-weekLevelCount[0][item.Key].LevelPeople[0]-weekLevelCount[0][item.Key].LevelPeople[1] > 2000 && 
                   weekLevelCount[0][item.Key].LevelPeople[16]-weekLevelCount[0][item.Key].LevelPeople[0]-weekLevelCount[0][item.Key].LevelPeople[1] < 4000)
                {
                    string strmsg = string.Format("{0},{1}", item.Key, dicStockNoMapping[item.Key]);

                    Console.WriteLine(strmsg);
                }
           
            }
                /*for (int i = 0; i < 23; i++)
                {
                    int i20000down = 0;
                    foreach (KeyValuePair<string, StockLevelCount> item in weekLevelCount[i])
                    {
                        if (weekLevelCount[i + 1].ContainsKey(item.Key))
                        {
                            if (weekLevelCount[i + 1][item.Key].LevelPeople[1] > iwater && weekLevelCount[i][item.Key].LevelPeople[1] < iwater)
                            {
                                i20000down++;
                            }
                        }
                    }
                    string strf = string.Format("{0}=>{1} , ", i, i20000down);
                    strmsg += strf;
                }*/
    

            /*
            foreach (KeyValuePair<string, StockLevelCount> item in weekLevelCount[0])
            {
                if (weekLevelCount[1].ContainsKey(item.Key) && weekLevelCount[2].ContainsKey(item.Key) && weekLevelCount[3].ContainsKey(item.Key) && weekLevelCount[4].ContainsKey(item.Key) && weekLevelCount[5].ContainsKey(item.Key) && weekLevelCount[6].ContainsKey(item.Key) && weekLevelCount[7].ContainsKey(item.Key) &&
                    weekLevelCount[8].ContainsKey(item.Key) && weekLevelCount[9].ContainsKey(item.Key) && weekLevelCount[10].ContainsKey(item.Key) && weekLevelCount[11].ContainsKey(item.Key) && weekLevelCount[12].ContainsKey(item.Key) && weekLevelCount[13].ContainsKey(item.Key) && weekLevelCount[14].ContainsKey(item.Key)
                    && weekLevelCount[15].ContainsKey(item.Key) && weekLevelCount[16].ContainsKey(item.Key) && weekLevelCount[17].ContainsKey(item.Key) && weekLevelCount[18].ContainsKey(item.Key) && weekLevelCount[19].ContainsKey(item.Key) && weekLevelCount[20].ContainsKey(item.Key) && weekLevelCount[21].ContainsKey(item.Key)
                    && weekLevelCount[22].ContainsKey(item.Key) && weekLevelCount[23].ContainsKey(item.Key) && weekLevelCount[24].ContainsKey(item.Key))
                {
                    // LevelRate[15] 是差異數調整，[14]:1,000,001以上，[13]:800,001-1,000,000，[12]:600,001-800,000，[11]:400,001-600,000，[10]200,001-400,000   ，[9]:100,001-200,000，[8]:50,001-100,000
                    //double dbrate0 = weekLevelCount[0][item.Key].LevelRate[10] + weekLevelCount[0][item.Key].LevelRate[11] + weekLevelCount[0][item.Key].LevelRate[12] + weekLevelCount[0][item.Key].LevelRate[13] + weekLevelCount[0][item.Key].LevelRate[14];// +weekLevelCount[0][item.Key].LevelRate[15];
                    //double dbrate1 = weekLevelCount[1][item.Key].LevelRate[10] + weekLevelCount[1][item.Key].LevelRate[11] + weekLevelCount[1][item.Key].LevelRate[12] + weekLevelCount[1][item.Key].LevelRate[13] + weekLevelCount[1][item.Key].LevelRate[14];// + weekLevelCount[1][item.Key].LevelRate[15];
                    //double dbrate2 = weekLevelCount[2][item.Key].LevelRate[10] + weekLevelCount[2][item.Key].LevelRate[11] + weekLevelCount[2][item.Key].LevelRate[12] + weekLevelCount[2][item.Key].LevelRate[13] + weekLevelCount[2][item.Key].LevelRate[14];// + weekLevelCount[2][item.Key].LevelRate[15];
                    //double dbrate3 = weekLevelCount[3][item.Key].LevelRate[10] + weekLevelCount[3][item.Key].LevelRate[11] + weekLevelCount[3][item.Key].LevelRate[12] + weekLevelCount[3][item.Key].LevelRate[13] + weekLevelCount[3][item.Key].LevelRate[14];// + weekLevelCount[3][item.Key].LevelRate[15];
                    //double dbrate4 = weekLevelCount[4][item.Key].LevelRate[10] + weekLevelCount[4][item.Key].LevelRate[11] + weekLevelCount[4][item.Key].LevelRate[12] + weekLevelCount[4][item.Key].LevelRate[13] + weekLevelCount[4][item.Key].LevelRate[14];// + weekLevelCount[4][item.Key].LevelRate[15];
                    //double dbrate5 = weekLevelCount[5][item.Key].LevelRate[10] + weekLevelCount[5][item.Key].LevelRate[11] + weekLevelCount[5][item.Key].LevelRate[12] + weekLevelCount[5][item.Key].LevelRate[13] + weekLevelCount[5][item.Key].LevelRate[14];// + weekLevelCount[5][item.Key].LevelRate[15];
                    //double dbrate6 = weekLevelCount[6][item.Key].LevelRate[10] + weekLevelCount[6][item.Key].LevelRate[11] + weekLevelCount[6][item.Key].LevelRate[12] + weekLevelCount[6][item.Key].LevelRate[13] + weekLevelCount[6][item.Key].LevelRate[14];// + weekLevelCount[6][item.Key].LevelRate[15];
                    //double dbrate7 = weekLevelCount[7][item.Key].LevelRate[10] + weekLevelCount[7][item.Key].LevelRate[11] + weekLevelCount[7][item.Key].LevelRate[12] + weekLevelCount[7][item.Key].LevelRate[13] + weekLevelCount[7][item.Key].LevelRate[14];// + weekLevelCount[7][item.Key].LevelRate[15];
                    //double dbrate08 = weekLevelCount[08][item.Key].LevelRate[10] + weekLevelCount[08][item.Key].LevelRate[11] + weekLevelCount[08][item.Key].LevelRate[12] + weekLevelCount[08][item.Key].LevelRate[13] + weekLevelCount[08][item.Key].LevelRate[14];// + weekLevelCount[08][item.Key].LevelRate[15];
                    //double dbrate09 = weekLevelCount[09][item.Key].LevelRate[10] + weekLevelCount[09][item.Key].LevelRate[11] + weekLevelCount[09][item.Key].LevelRate[12] + weekLevelCount[09][item.Key].LevelRate[13] + weekLevelCount[09][item.Key].LevelRate[14];// + weekLevelCount[09][item.Key].LevelRate[15];
                    //double dbrate10 = weekLevelCount[10][item.Key].LevelRate[10] + weekLevelCount[10][item.Key].LevelRate[11] + weekLevelCount[10][item.Key].LevelRate[12] + weekLevelCount[10][item.Key].LevelRate[13] + weekLevelCount[10][item.Key].LevelRate[14];// + weekLevelCount[10][item.Key].LevelRate[15];
                    //double dbrate11 = weekLevelCount[11][item.Key].LevelRate[10] + weekLevelCount[11][item.Key].LevelRate[11] + weekLevelCount[11][item.Key].LevelRate[12] + weekLevelCount[11][item.Key].LevelRate[13] + weekLevelCount[11][item.Key].LevelRate[14];// + weekLevelCount[11][item.Key].LevelRate[15];
                    //double dbrate12 = weekLevelCount[12][item.Key].LevelRate[10] + weekLevelCount[12][item.Key].LevelRate[11] + weekLevelCount[12][item.Key].LevelRate[12] + weekLevelCount[12][item.Key].LevelRate[13] + weekLevelCount[12][item.Key].LevelRate[14];// + weekLevelCount[12][item.Key].LevelRate[15];
                    //double dbrate13 = weekLevelCount[13][item.Key].LevelRate[10] + weekLevelCount[13][item.Key].LevelRate[11] + weekLevelCount[13][item.Key].LevelRate[12] + weekLevelCount[13][item.Key].LevelRate[13] + weekLevelCount[13][item.Key].LevelRate[14];// + weekLevelCount[13][item.Key].LevelRate[15];
                    //double dbrate14 = weekLevelCount[14][item.Key].LevelRate[10] + weekLevelCount[14][item.Key].LevelRate[11] + weekLevelCount[14][item.Key].LevelRate[12] + weekLevelCount[14][item.Key].LevelRate[13] + weekLevelCount[14][item.Key].LevelRate[14];// + weekLevelCount[14][item.Key].LevelRate[15];
                    //double dbrate15 = weekLevelCount[15][item.Key].LevelRate[10] + weekLevelCount[15][item.Key].LevelRate[11] + weekLevelCount[15][item.Key].LevelRate[12] + weekLevelCount[15][item.Key].LevelRate[13] + weekLevelCount[15][item.Key].LevelRate[14];//+ weekLevelCount[15][item.Key].LevelRate[15];

                    // LevelPeople[16] 總人數 - [0]1-999股 - [1]:1-5張 -[2]:5-10張
                    double dbpeople0 = weekLevelCount[0][item.Key].LevelPeople[16] - weekLevelCount[0][item.Key].LevelPeople[0] - weekLevelCount[0][item.Key].LevelPeople[1];// - weekLevelCount[0][item.Key].LevelPeople[2];
                    double dbpeople1 = weekLevelCount[1][item.Key].LevelPeople[16] - weekLevelCount[1][item.Key].LevelPeople[0] - weekLevelCount[1][item.Key].LevelPeople[1];// - weekLevelCount[1][item.Key].LevelPeople[2];
                    double dbpeople2 = weekLevelCount[2][item.Key].LevelPeople[16] - weekLevelCount[2][item.Key].LevelPeople[0] - weekLevelCount[2][item.Key].LevelPeople[1];// - weekLevelCount[2][item.Key].LevelPeople[2];
                    double dbpeople3 = weekLevelCount[3][item.Key].LevelPeople[16] - weekLevelCount[3][item.Key].LevelPeople[0] - weekLevelCount[3][item.Key].LevelPeople[1];// - weekLevelCount[3][item.Key].LevelPeople[2];
                    double dbpeople4 = weekLevelCount[4][item.Key].LevelPeople[16] - weekLevelCount[4][item.Key].LevelPeople[0] - weekLevelCount[4][item.Key].LevelPeople[1];// - weekLevelCount[4][item.Key].LevelPeople[2];
                    double dbpeople5 = weekLevelCount[5][item.Key].LevelPeople[16] - weekLevelCount[5][item.Key].LevelPeople[0] - weekLevelCount[5][item.Key].LevelPeople[1];// - weekLevelCount[5][item.Key].LevelPeople[2];
                    double dbpeople6 = weekLevelCount[6][item.Key].LevelPeople[16] - weekLevelCount[6][item.Key].LevelPeople[0] - weekLevelCount[6][item.Key].LevelPeople[1];// - weekLevelCount[6][item.Key].LevelPeople[2];
                    double dbpeople7 = weekLevelCount[7][item.Key].LevelPeople[16] - weekLevelCount[7][item.Key].LevelPeople[0] - weekLevelCount[7][item.Key].LevelPeople[1];// -weekLevelCount[7][item.Key].LevelPeople[2];
                    double dbpeople08 = weekLevelCount[08][item.Key].LevelPeople[16] - weekLevelCount[08][item.Key].LevelPeople[0] - weekLevelCount[08][item.Key].LevelPeople[1];// - weekLevelCount[08][item.Key].LevelPeople[2];
                    double dbpeople09 = weekLevelCount[09][item.Key].LevelPeople[16] - weekLevelCount[09][item.Key].LevelPeople[0] - weekLevelCount[09][item.Key].LevelPeople[1];// - weekLevelCount[09][item.Key].LevelPeople[2];
                    double dbpeople10 = weekLevelCount[10][item.Key].LevelPeople[16] - weekLevelCount[10][item.Key].LevelPeople[0] - weekLevelCount[10][item.Key].LevelPeople[1];// - weekLevelCount[10][item.Key].LevelPeople[2];
                    double dbpeople11 = weekLevelCount[11][item.Key].LevelPeople[16] - weekLevelCount[11][item.Key].LevelPeople[0] - weekLevelCount[11][item.Key].LevelPeople[1];// - weekLevelCount[11][item.Key].LevelPeople[2];
                    double dbpeople12 = weekLevelCount[12][item.Key].LevelPeople[16] - weekLevelCount[12][item.Key].LevelPeople[0] - weekLevelCount[12][item.Key].LevelPeople[1];// - weekLevelCount[12][item.Key].LevelPeople[2];
                    double dbpeople13 = weekLevelCount[13][item.Key].LevelPeople[16] - weekLevelCount[13][item.Key].LevelPeople[0] - weekLevelCount[13][item.Key].LevelPeople[1];// - weekLevelCount[13][item.Key].LevelPeople[2];
                    double dbpeople14 = weekLevelCount[14][item.Key].LevelPeople[16] - weekLevelCount[14][item.Key].LevelPeople[0] - weekLevelCount[14][item.Key].LevelPeople[1];// - weekLevelCount[14][item.Key].LevelPeople[2];
                    double dbpeople15 = weekLevelCount[15][item.Key].LevelPeople[16] - weekLevelCount[15][item.Key].LevelPeople[0] - weekLevelCount[15][item.Key].LevelPeople[1];// -weekLevelCount[15][item.Key].LevelPeople[2];
                    double dbpeople16 = weekLevelCount[16][item.Key].LevelPeople[16] - weekLevelCount[16][item.Key].LevelPeople[0] - weekLevelCount[16][item.Key].LevelPeople[1];
                    double dbpeople17 = weekLevelCount[17][item.Key].LevelPeople[16] - weekLevelCount[17][item.Key].LevelPeople[0] - weekLevelCount[17][item.Key].LevelPeople[1];
                    double dbpeople18 = weekLevelCount[18][item.Key].LevelPeople[16] - weekLevelCount[18][item.Key].LevelPeople[0] - weekLevelCount[18][item.Key].LevelPeople[1];
                    double dbpeople19 = weekLevelCount[19][item.Key].LevelPeople[16] - weekLevelCount[19][item.Key].LevelPeople[0] - weekLevelCount[19][item.Key].LevelPeople[1];
                    double dbpeople20 = weekLevelCount[20][item.Key].LevelPeople[16] - weekLevelCount[20][item.Key].LevelPeople[0] - weekLevelCount[20][item.Key].LevelPeople[1];
                    double dbpeople21 = weekLevelCount[21][item.Key].LevelPeople[16] - weekLevelCount[21][item.Key].LevelPeople[0] - weekLevelCount[21][item.Key].LevelPeople[1];
                    double dbpeople22 = weekLevelCount[22][item.Key].LevelPeople[16] - weekLevelCount[22][item.Key].LevelPeople[0] - weekLevelCount[22][item.Key].LevelPeople[1];
                    double dbpeople23 = weekLevelCount[23][item.Key].LevelPeople[16] - weekLevelCount[23][item.Key].LevelPeople[0] - weekLevelCount[23][item.Key].LevelPeople[1];
                    double dbpeople24 = weekLevelCount[24][item.Key].LevelPeople[16] - weekLevelCount[24][item.Key].LevelPeople[0] - weekLevelCount[24][item.Key].LevelPeople[1];
                }

            }*/
        }

        private void button53_Click(object sender, EventArgs e)
        {
            DateTime dtStart = DateTime.Now;
            DateTime dtLoopEnd = dtStart.AddMonths(-12);
            DateTime dtLoopStart = dtStart;
            int iMaxContainDays = 250;

            DateTime dtfail = DateTime.Now;
            foreach (string strno in dbEvenydayOLHCDic.Keys)
            {
                Dictionary<int, DateTime> backDaysmapDatetime = new Dictionary<int, DateTime>();
                double[] dbOpenArray = new double[iMaxContainDays];
                double[] dbLowArray = new double[iMaxContainDays];
                double[] dbHighArray = new double[iMaxContainDays];
                double[] dbCloseArray = new double[iMaxContainDays];
                long[] lvol = new long[iMaxContainDays];
                for (int i = 0; i < iMaxContainDays; i++)
                {
                    dbOpenArray[i] = 0;
                    dbLowArray[i] = 0;
                    dbHighArray[i] = 0;
                    dbCloseArray[i] = 0;
                    lvol[i] = 0;
                }
                int iloop = 0;
                DateTime dtchk = dtfail;
                foreach (DateTime dtloop in dbEvenydayOLHCDic[strno].Keys)
                {
                    if (dtchk != dtfail)
                    {
                        // 確認時間是照順序排
                        if (dtchk <= dtloop)
                        {
                            MessageBox.Show("False");
                        }
                    }
                    dtchk = dtloop;
                    if (dtchk <= dtStart)
                    {
                        backDaysmapDatetime.Add(iloop, dtloop);
                        lvol[iloop] = dbStockVolDic[strno][dtchk];
                        dbOpenArray[iloop] = dbEvenydayOLHCDic[strno][dtchk][0];
                        dbLowArray[iloop] = dbEvenydayOLHCDic[strno][dtchk][1];
                        dbHighArray[iloop] = dbEvenydayOLHCDic[strno][dtchk][2];
                        dbCloseArray[iloop] = dbEvenydayOLHCDic[strno][dtchk][3];
                        iloop++;
                        if (iloop == iMaxContainDays)
                            break;
                    }
                }
                double[] db5ary = new double[50];
                double[] db10ary = new double[50];
                double[] db20ary = new double[50];
                double[] db60ary = new double[50];
                double[] db120ary = new double[50];
                double[] db200ary = new double[50];
                double[] dbstdev = new double[50];

                for (int icrossdays = 0; icrossdays < 50; icrossdays++)
                {
                    db5ary[icrossdays] = dbCloseArray.Skip(icrossdays).Take(5).Average();
                    db10ary[icrossdays] = dbCloseArray.Skip(icrossdays).Take(10).Average();
                    db20ary[icrossdays] = dbCloseArray.Skip(icrossdays).Take(20).Average();
                    db60ary[icrossdays] = dbCloseArray.Skip(icrossdays).Take(60).Average();
                    db120ary[icrossdays] = dbCloseArray.Skip(icrossdays).Take(120).Average();
                    db200ary[icrossdays] = dbCloseArray.Skip(icrossdays).Take(200).Average();

                    double[] stdevprep = new double[20];
                    Array.Copy(dbCloseArray, icrossdays, stdevprep, 0, 20);
                    double dbavg=0;
                    dbstdev[icrossdays] = StandardDeviation(stdevprep, out dbavg);

                }


                int iCountub = 0;
                int iCountdb = 0;
                int iCountContinue = 0;
                if (strno == "0050")
                {
                    int mmm = 0;
                }
                for (int i = 0; i < 50; i++)
                {
                    if (dicStockNoMapping.ContainsKey(strno))
                    {
                        double ub = db20ary[i] + (dbstdev[i] * 2);
                        double db = db20ary[i] - (dbstdev[i] * 2);

                        if (dbHighArray[i] >= ub)
                        {
                            iCountub++;
                        }
                        else if (dbLowArray[i] <= db )
                        {
                            iCountdb++;
                        }
                        else
                        {
                            if (i == iCountContinue)
                            {
                                iCountContinue++;
                            }

                        }
                            
                    }
                }
                if (strno == "3711")
                {
                    int mmm = 0;
                }
                if(dicStockNoMapping.ContainsKey(strno))
                {
                    string trmsg = string.Format("{0},{1},{2:F2},{3},{4},{5}", strno, dicStockNoMapping[strno], dbCloseArray[0], iCountub, iCountdb, iCountContinue);
                    Console.WriteLine(trmsg);
                }
                


            }
        }

        private void button54_Click(object sender, EventArgs e)
        {
            //DateTime dtStart = DateTime.Now;
            DateTime dtStart = monthCalendar1.SelectionStart;
            DateTime dtLoopEnd = dtStart.AddMonths(-12);
            DateTime dtLoopStart = dtStart;
            int iMaxContainDays = 300;

            DateTime dtfail = DateTime.Now;
            foreach (string strno in dbEvenydayOLHCDic.Keys)
            {

                Dictionary<int, DateTime> backDaysmapDatetime = new Dictionary<int, DateTime>();
                double[] dbOpenArray = new double[iMaxContainDays];
                double[] dbLowArray = new double[iMaxContainDays];
                double[] dbHighArray = new double[iMaxContainDays];
                double[] dbCloseArray = new double[iMaxContainDays];
                long[] lvol = new long[iMaxContainDays];
                for (int i = 0; i < iMaxContainDays; i++)
                {
                    dbOpenArray[i] = 0;
                    dbLowArray[i] = 0;
                    dbHighArray[i] = 0;
                    dbCloseArray[i] = 0;
                    lvol[i] = 0;
                }
                int iloop = 0;
                DateTime dtchk = dtfail;
                foreach (DateTime dtloop in dbEvenydayOLHCDic[strno].Keys)
                {
                    if (dtchk != dtfail)
                    {
                        // 確認時間是照順序排
                        if (dtchk <= dtloop)
                        {
                            MessageBox.Show("False");
                        }
                    }
                    dtchk = dtloop;
                    if (dtchk <= dtStart)
                    {
                        backDaysmapDatetime.Add(iloop, dtloop);
                        lvol[iloop] = dbStockVolDic[strno][dtchk];
                        dbOpenArray[iloop] = dbEvenydayOLHCDic[strno][dtchk][0];
                        dbLowArray[iloop] = dbEvenydayOLHCDic[strno][dtchk][1];
                        dbHighArray[iloop] = dbEvenydayOLHCDic[strno][dtchk][2];
                        dbCloseArray[iloop] = dbEvenydayOLHCDic[strno][dtchk][3];
                        iloop++;
                        if (iloop == iMaxContainDays)
                            break;
                    }
                }
                double[] db5ary = new double[50];
                double[] db10ary = new double[50];
                double[] db20ary = new double[50];
                double[] db60ary = new double[50];
                double[] db120ary = new double[50];
                double[] db200ary = new double[50];

                for (int icrossdays = 0; icrossdays < 50; icrossdays++)
                {
                    db5ary[icrossdays] = dbCloseArray.Skip(icrossdays).Take(5).Average();
                    db10ary[icrossdays] = dbCloseArray.Skip(icrossdays).Take(10).Average();
                    db20ary[icrossdays] = dbCloseArray.Skip(icrossdays).Take(20).Average();
                    db60ary[icrossdays] = dbCloseArray.Skip(icrossdays).Take(60).Average();
                    db120ary[icrossdays] = dbCloseArray.Skip(icrossdays).Take(120).Average();
                    db200ary[icrossdays] = dbCloseArray.Skip(icrossdays).Take(200).Average();
                }

                double[] dbAllLineMaxMinPercentAry = new double[50];
                for (int icrossdays = 0; icrossdays < 50; icrossdays++)
                {
                    double[] dbAllLineAry = new double[7];
                    dbAllLineAry[0] = dbHighArray[icrossdays];
                    dbAllLineAry[1] = dbLowArray[icrossdays];
                    dbAllLineAry[2] = db5ary[icrossdays];
                    dbAllLineAry[3] = db10ary[icrossdays];
                    dbAllLineAry[4] = db20ary[icrossdays];
                    dbAllLineAry[5] = db60ary[icrossdays];
                    dbAllLineAry[6] = db120ary[icrossdays];
                    //dbAllLineAry[7] = db200ary[icrossdays];
                    dbAllLineMaxMinPercentAry[icrossdays] = (dbAllLineAry.Max() - dbAllLineAry.Min()) / dbAllLineAry.Min();

                }

                double db300hi = dbHighArray.Max();
                double db300low = dbLowArray.Min();



                double db60hi = dbHighArray.Take(60).Max();
                double db60low = dbLowArray.Take(60).Min();

                double dbrange = (db60hi-db60low)/(db300hi-db300low);
                if (dicStockNoMapping.ContainsKey(strno) &&
                    db60ary[0] >= db60ary[1] &&
                    db20ary[0] >= db20ary[1] &&
                    db20ary[0] >=  db60ary[0] &&
                    db300hi >=  db60hi && db300low <= db60low &&
                    lvol.Average() > 500000 &&
                    dbrange < 0.2)
                {
                    string trmsg = string.Format("{0},{1},{2:F2},{3:F2}", strno, dicStockNoMapping[strno], dbCloseArray[0], dbrange);
                    Console.WriteLine(trmsg);
                }

            }
        }

        private void button55_Click(object sender, EventArgs e)
        {
            DateTime dtStart = monthCalendar1.SelectionStart;
            DateTime dtLoopEnd = dtStart.AddMonths(-12);
            DateTime dtLoopStart = dtStart;
            int iMaxContainDays = 360;

            string strBasePath = "D:\\Chips\\stock\\TDCC_OD_1-5_";
            int iCount = 4;
            Dictionary<string, StockLevelCount>[] weekLevelCount = new Dictionary<string, StockLevelCount>[iCount];
            int iCountWeek = 0;
            while (dtLoopStart > dtLoopEnd && iCountWeek < iCount)
            {
                string strcsv = strBasePath + dtLoopStart.ToString("yyyyMMdd") + ".csv";
                if (File.Exists(strcsv))
                {
                    weekLevelCount[iCountWeek] = new Dictionary<string, StockLevelCount>();
                    int iReadCount = 0;
                    IEnumerable<string> lines = File.ReadLines(strcsv);
                    foreach (string line in lines)
                    {
                        if (iReadCount > 0)
                        {
                            string[] strSplitLine = line.Split(',');
                            if (weekLevelCount[iCountWeek].ContainsKey(strSplitLine[1]))
                            {
                                int iLevel = int.Parse(strSplitLine[2]) - 1;
                                weekLevelCount[iCountWeek][strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                                weekLevelCount[iCountWeek][strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                                weekLevelCount[iCountWeek][strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                            }
                            else
                            {
                                weekLevelCount[iCountWeek].Add(strSplitLine[1], new StockLevelCount());
                                int iLevel = int.Parse(strSplitLine[2]) - 1;
                                weekLevelCount[iCountWeek][strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                                weekLevelCount[iCountWeek][strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                                weekLevelCount[iCountWeek][strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                            }
                        }
                        iReadCount++;
                    }
                    iCountWeek++;
                }
                dtLoopStart = dtLoopStart.AddDays(-1);
            }

            DateTime dtfail = DateTime.Now;
            foreach (string strno in dbEvenydayOLHCDic.Keys)
            {

                Dictionary<int, DateTime> backDaysmapDatetime = new Dictionary<int, DateTime>();
                double[] dbOpenArray = new double[iMaxContainDays];
                double[] dbLowArray = new double[iMaxContainDays];
                double[] dbHighArray = new double[iMaxContainDays];
                double[] dbCloseArray = new double[iMaxContainDays];
                long[] lvol = new long[iMaxContainDays];
                for (int i = 0; i < iMaxContainDays; i++)
                {
                    dbOpenArray[i] = 0;
                    dbLowArray[i] = 0;
                    dbHighArray[i] = 0;
                    dbCloseArray[i] = 0;
                    lvol[i] = 0;
                }
                int iloop = 0;
                DateTime dtchk = dtfail;
                foreach (DateTime dtloop in dbEvenydayOLHCDic[strno].Keys)
                {
                    if (dtchk != dtfail)
                    {
                        // 確認時間是照順序排
                        if (dtchk <= dtloop)
                        {
                            MessageBox.Show("False");
                        }
                    }
                    dtchk = dtloop;
                    if (dtchk <= dtStart)
                    {
                        backDaysmapDatetime.Add(iloop, dtloop);
                        lvol[iloop] = dbStockVolDic[strno][dtchk];
                        dbOpenArray[iloop] = dbEvenydayOLHCDic[strno][dtchk][0];
                        dbLowArray[iloop] = dbEvenydayOLHCDic[strno][dtchk][1];
                        dbHighArray[iloop] = dbEvenydayOLHCDic[strno][dtchk][2];
                        dbCloseArray[iloop] = dbEvenydayOLHCDic[strno][dtchk][3];
                        iloop++;
                        if (iloop == iMaxContainDays)
                            break;
                    }
                }
                double[] db5ary = new double[50];
                double[] db10ary = new double[50];
                double[] db20ary = new double[50];
                double[] db60ary = new double[50];
                double[] db120ary = new double[50];
                double[] db200ary = new double[50];

                for (int icrossdays = 0; icrossdays < 50; icrossdays++)
                {
                    db5ary[icrossdays] = dbCloseArray.Skip(icrossdays).Take(5).Average();
                    db10ary[icrossdays] = dbCloseArray.Skip(icrossdays).Take(10).Average();
                    db20ary[icrossdays] = dbCloseArray.Skip(icrossdays).Take(20).Average();
                    db60ary[icrossdays] = dbCloseArray.Skip(icrossdays).Take(60).Average();
                    db120ary[icrossdays] = dbCloseArray.Skip(icrossdays).Take(120).Average();
                    db200ary[icrossdays] = dbCloseArray.Skip(icrossdays).Take(200).Average();
                }

                double[] dbAllLineMaxMinPercentAry = new double[50];
                for (int icrossdays = 0; icrossdays < 50; icrossdays++)
                {
                    double[] dbAllLineAry = new double[7];
                    dbAllLineAry[0] = dbHighArray[icrossdays];
                    dbAllLineAry[1] = dbLowArray[icrossdays];
                    dbAllLineAry[2] = db5ary[icrossdays];
                    dbAllLineAry[3] = db10ary[icrossdays];
                    dbAllLineAry[4] = db20ary[icrossdays];
                    dbAllLineAry[5] = db60ary[icrossdays];
                    dbAllLineAry[6] = db120ary[icrossdays];
                    //dbAllLineAry[7] = db200ary[icrossdays];
                    dbAllLineMaxMinPercentAry[icrossdays] = (dbAllLineAry.Max() - dbAllLineAry.Min()) / dbAllLineAry.Min();

                }

                double db300hi = dbHighArray.Max();
                double db300low = dbLowArray.Min();



                double db60hi = dbHighArray.Take(60).Max();
                double db60low = dbLowArray.Take(60).Min();

                double dbrange = (db60hi - db60low) / (db300hi - db300low);


                double[] dbbb20ary = new double[60];
                double[] dbbb60ary = new double[60];
                double[] dbbb120ary = new double[60];
                double[] dbbb240ary = new double[60];
                for (int ibbdays = 0; ibbdays < 60; ibbdays++)
                {
                    double dbavg = 0;
                    dbbb20ary[ibbdays] = StandardDeviation(dbCloseArray.Skip(ibbdays).Take(20), out dbavg);
                    dbbb60ary[ibbdays] = StandardDeviation(dbCloseArray.Skip(ibbdays).Take(60), out dbavg);
                    dbbb120ary[ibbdays] = StandardDeviation(dbCloseArray.Skip(ibbdays).Take(120), out dbavg);
                    dbbb240ary[ibbdays] = StandardDeviation(dbCloseArray.Skip(ibbdays).Take(240), out dbavg);


                }

                if (weekLevelCount[0].ContainsKey(strno) && weekLevelCount[3].ContainsKey(strno) && dicStockNoMapping.ContainsKey(strno))
                {
                    long l0 = weekLevelCount[0][strno].LevelPeople[16] - weekLevelCount[0][strno].LevelPeople[0];// -weekLevelCount[0][strno].LevelPeople[1];
                    long l3 = weekLevelCount[3][strno].LevelPeople[16] - weekLevelCount[3][strno].LevelPeople[0];// -weekLevelCount[3][strno].LevelPeople[1];
                    double big0 = weekLevelCount[0][strno].LevelRate[10] + weekLevelCount[0][strno].LevelRate[11] + weekLevelCount[0][strno].LevelRate[12] + weekLevelCount[0][strno].LevelRate[13] + weekLevelCount[0][strno].LevelRate[14];
                    double big3 = weekLevelCount[3][strno].LevelRate[10] + weekLevelCount[3][strno].LevelRate[11] + weekLevelCount[3][strno].LevelRate[12] + weekLevelCount[3][strno].LevelRate[13] + weekLevelCount[3][strno].LevelRate[14];

                    long l3PTotal = 0;
                    long lHedgeTotal = 0;
                    long lTrust = 0;
                    long lSelf = 0;
                    long lTotal = 0;
                    long lFor = 0;

                    long ladd = 0;
                    int iLoopDay = 0;
                    dtLoopStart = dtStart;
                    while (iLoopDay < 20)  // 計算最近幾天
                    {
                        if (workingdays.Contains(dtLoopStart))
                        {
                            if (dicStock3Party.ContainsKey(strno))
                            {
                                if (dicStock3Party[strno].ContainsKey(dtLoopStart))
                                {
                                    int iTotal = int.Parse(dicStock3Party[strno][dtLoopStart].Party3Total, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                    int iHedge = int.Parse(dicStock3Party[strno][dtLoopStart].SelfHedgeTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                    int iHedgeBuy = int.Parse(dicStock3Party[strno][dtLoopStart].SelfHedgeBuy, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                    int iHedgeSell = int.Parse(dicStock3Party[strno][dtLoopStart].SelfHedgeSell, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                    int iForeigen = int.Parse(dicStock3Party[strno][dtLoopStart].ForeigenTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                    int iTrust = int.Parse(dicStock3Party[strno][dtLoopStart].TrustTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                    int iSelf = int.Parse(dicStock3Party[strno][dtLoopStart].SelfSelfTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                    ladd += (iTotal - iHedge); // 計算不含避險3大買量
                                    //ladd += iTotal; // 全算
                                    //ladd += (iTrust + iForeigen); //外 信
                                    //ladd += iHedge;

                                    lTotal += (iTotal - iHedge);
                                    l3PTotal += iTotal;
                                    lHedgeTotal += iHedge;
                                    lTrust += iTrust;
                                    lSelf += iSelf;
                                    lFor += iForeigen;
                                }
                            }
                            iLoopDay++;
                        }

                        dtLoopStart = dtLoopStart.AddDays(-1);
                    }

                    
                    /*三天包20MAif (((dbOpenArray[0] < db20ary[0] && dbCloseArray[0] > db20ary[0]) || (dbOpenArray[0] > db20ary[0] && dbCloseArray[0] < db20ary[0])) &&
                        ((dbOpenArray[1] < db20ary[1] && dbCloseArray[1] > db20ary[1]) || (dbOpenArray[1] > db20ary[1] && dbCloseArray[1] < db20ary[1])) &&
                        ((dbOpenArray[2] < db20ary[2] && dbCloseArray[2] > db20ary[2]) || (dbOpenArray[2] > db20ary[2] && dbCloseArray[2] < db20ary[2]))
                        )*/
                    if (lTrust > 0 && lFor>0)
                    {
                        double dbbandwidthpercent = ((db20ary[0] + (dbbb20ary[0] * 2.1)) - (db20ary[0] - (dbbb20ary[0] * 2.1))) / db20ary[0];
                        string trmsg = string.Format("{0},{1},{2:F2},{3:F5}", strno, dicStockNoMapping[strno], dbCloseArray[0], dbbandwidthpercent);
                        Console.WriteLine(trmsg);
                    }
                    //if (dicStockNoMapping.ContainsKey(strno) &&
                    /*if (dicStockNoMapping.ContainsKey(strno) &&
                        db120ary[0] > db120ary[49] &&
                        db60ary[0] < db60ary[49] &&
                        db60ary[0] > db120ary[0] &&
                        db60ary[49] - db120ary[49] > db60ary[0] - db120ary[0] &&
                        dbCloseArray[0] > db20ary[0] &&
                        db20ary[0] < db120ary[0] &&
                        (l0 < l3 || big0 > big3) &&
                        lvol.Average() > 100000 //&&                    dbrange < 0.25
                        )
                    {
                        // 2019/12/26 前有找到8028
                        string trmsg = string.Format("{0},{1},{2:F2},{3:F2}", strno, dicStockNoMapping[strno], dbCloseArray[0], dbrange);
                        Console.WriteLine(trmsg);
                    }*/
                }


            }
        }

        private void button56_Click(object sender, EventArgs e)
        {
            DateTime dtStart = monthCalendar1.SelectionStart;
            DateTime dtend = dtStart.AddDays(-300);
            int iMaxContainDays = 120;
            foreach (string strno in dbEvenydayOLHCDic.Keys)
            {
                if (dbStockVolDic.ContainsKey(strno) && dbStockVolDic[strno].Values.Count>60)
                {

                    double[] dbOpenArray = new double[iMaxContainDays];
                    double[] dbLowArray = new double[iMaxContainDays];
                    double[] dbHighArray = new double[iMaxContainDays];
                    double[] dbCloseArray = new double[iMaxContainDays];
                    long[] volArray = new long[iMaxContainDays];
                    for (int i = 0; i < iMaxContainDays; i++)
                    {
                        dbOpenArray[i] = 0;
                        dbLowArray[i] = 0;
                        dbHighArray[i] = 0;
                        dbCloseArray[i] = 0;
                        volArray[i] = 0;
                    }
                    DateTime dtloop = dtStart;
                    bool bnotenoughvol = true;
                    for (int iCount = 0; iCount < iMaxContainDays; )
                    {
                        if (dbStockVolDic[strno].ContainsKey(dtloop))
                        {
                            volArray[iCount] = dbStockVolDic[strno][dtloop];
                            dbOpenArray[iCount] = dbEvenydayOLHCDic[strno][dtloop][0];
                            dbLowArray[iCount] = dbEvenydayOLHCDic[strno][dtloop][1];
                            dbHighArray[iCount] = dbEvenydayOLHCDic[strno][dtloop][2];
                            dbCloseArray[iCount] = dbEvenydayOLHCDic[strno][dtloop][3];
                            if (dbStockVolDic[strno][dtloop] < 500000)
                                bnotenoughvol = false;

                            iCount++;
                        }
                        dtloop = dtloop.AddDays(-1);
                        if(dtloop < dtend)
                        {
                            break;
                        }
                    }
                    
                    long lMax = volArray.Max();
                    double dbavg = volArray.Average();
                    int iMaxupday = -1;
                    for (int iCount = 0; iCount < iMaxContainDays; iCount++)
                    {
                        if (lMax == volArray[iCount] && bnotenoughvol)
                        {
                            if(dbOpenArray[iCount] < dbCloseArray[iCount])
                            {
                                if(iCount != iMaxContainDays-1)
                                {
                                    if (iCount==0)
                                    {
                                        iMaxupday = iCount;
                                        break;
                                    }
                                    if(dbCloseArray[iCount] > dbCloseArray[iCount-1] )
                                    {
                                        iMaxupday = iCount;
                                        break;
                                    }
                                }

                            }
                        }
                    }

                    if(iMaxupday != -1)
                    {
                        DateTime dtMaxvolday = new DateTime(2000, 1, 1);
                        DateTime dtGotMaxday = dtMaxvolday;
                        foreach (DateTime dtkey in dbStockVolDic[strno].Keys)
                        {
                            if (dbStockVolDic[strno][dtkey] == lMax)
                            {
                                dtGotMaxday = dtkey;
                                break;
                            }
                        }
                        string strdesc="";
                        
                        if (dbCloseArray[0] > dbHighArray[iMaxupday]*1.1)
                        {
                            strdesc = "too high";
                        }
                        else if (dbCloseArray[0] > dbHighArray[iMaxupday])
                        {
                            strdesc = "high";
                        }
                        else if (dbCloseArray[0] <= dbHighArray[iMaxupday] && dbCloseArray[0] >= dbLowArray[iMaxupday])
                        {
                            strdesc = "mid";
                        }
                        else if (dbCloseArray[0] <= dbLowArray[iMaxupday] * 0.9)
                        {
                            strdesc = "too low";
                        }
                        else if (dbCloseArray[0] <= dbLowArray[iMaxupday] )
                        {
                            strdesc = "low";
                        }


                        if (dtGotMaxday != dtMaxvolday && dicStockNoMapping.ContainsKey(strno))
                        {
                            string strMsg = string.Format("{0},{1},{2},{3}", strno, dicStockNoMapping[strno], dtGotMaxday.ToString("MM/dd"), strdesc);
                            Console.WriteLine(strMsg);
                        }
                    }
                   
                    

                }
            }
        }

        private void button57_Click(object sender, EventArgs e)
        {
            double dbPercent = 0.007; // 成長數大於總股數
            DateTime dtStart = monthCalendar1.SelectionStart;
            DateTime dtLoopStart = monthCalendar1.SelectionStart;

            string strMsg = string.Format("找{0}投信或外資買超過總股數{1:F4}", dtStart.ToString("MM/dd"), dbPercent);
            Console.WriteLine(strMsg);

            CandidateParty3Stock.Clear();

            //CandidateWeekChipSet.Clear();
            string strBasePath = "D:\\Chips\\stock\\TDCC_OD_1-5_";
            string strReadFile = "";
            int icountchip = 0;
            while (true)
            {
                string strcsv = strBasePath + dtLoopStart.ToString("yyyyMMdd") + ".csv";
                if (File.Exists(strcsv))
                {
                    strReadFile = strcsv;
                    break;
                }
                dtLoopStart = dtLoopStart.AddDays(-1);
            }
            Dictionary<string, StockLevelCount> weekLevelCount = new Dictionary<string, StockLevelCount>();
            int iReadCount = 0;
            IEnumerable<string> lines = File.ReadLines(strReadFile);
            foreach (string line in lines)  // 取流通在外數
            {
                if (iReadCount > 0)
                {
                    string[] strSplitLine = line.Split(',');
                    if (weekLevelCount.ContainsKey(strSplitLine[1]))
                    {
                        int iLevel = int.Parse(strSplitLine[2]) - 1;
                        weekLevelCount[strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                        weekLevelCount[strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                        weekLevelCount[strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                    }
                    else
                    {
                        weekLevelCount.Add(strSplitLine[1], new StockLevelCount());
                        int iLevel = int.Parse(strSplitLine[2]) - 1;
                        weekLevelCount[strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                        weekLevelCount[strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                        weekLevelCount[strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                    }
                }
                iReadCount++;
            }

            Dictionary<string, long> stock3PartyVol = new Dictionary<string, long>();
            foreach (string strStockNo in weekLevelCount.Keys)
            {
                long l3PTotal = 0;
                long lHedgeTotal = 0;
                long lTrust = 0;
                long lSelf = 0;
                long lTotal = 0;
                long lFor = 0;

                long ladd = 0;

                dtLoopStart = dtStart;

                if (workingdays.Contains(dtLoopStart))
                {
                    if (dicStock3Party.ContainsKey(strStockNo))
                    {
                        if (dicStock3Party[strStockNo].ContainsKey(dtLoopStart))
                        {
                            int iTotal = int.Parse(dicStock3Party[strStockNo][dtLoopStart].Party3Total, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                            int iHedge = int.Parse(dicStock3Party[strStockNo][dtLoopStart].SelfHedgeTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                            int iHedgeBuy = int.Parse(dicStock3Party[strStockNo][dtLoopStart].SelfHedgeBuy, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                            int iHedgeSell = int.Parse(dicStock3Party[strStockNo][dtLoopStart].SelfHedgeSell, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                            int iForeigen = int.Parse(dicStock3Party[strStockNo][dtLoopStart].ForeigenTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                            int iTrust = int.Parse(dicStock3Party[strStockNo][dtLoopStart].TrustTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                            int iSelf = int.Parse(dicStock3Party[strStockNo][dtLoopStart].SelfSelfTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                            ladd += (iTotal - iHedge); // 計算不含避險3大買量
                            //ladd += iTotal; // 全算
                            //ladd += (iTrust + iForeigen); //外 信
                            //ladd += iHedge;

                            lTotal += (iTotal - iHedge);
                            l3PTotal += iTotal;
                            lHedgeTotal += iHedge;
                            lTrust += iTrust;
                            lSelf += iSelf;
                            lFor += iForeigen;

                            int iTrustBuy = int.Parse(dicStock3Party[strStockNo][dtLoopStart].TrustBuy, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                            int iTrustSell = int.Parse(dicStock3Party[strStockNo][dtLoopStart].TrustSell, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);

                            int iForBuy = int.Parse(dicStock3Party[strStockNo][dtLoopStart].ForeigenBuy, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                            int iForSell = int.Parse(dicStock3Party[strStockNo][dtLoopStart].ForeigenSell, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);

                            if (weekLevelCount.ContainsKey(strStockNo))
                            {
                                double dbtrust = (double)lTrust / (double)weekLevelCount[strStockNo].LevelVol[16];
                                double dbfor = (double)lFor / (double)weekLevelCount[strStockNo].LevelVol[16];
                                double dbHedge = (double)lHedgeTotal / (double)weekLevelCount[strStockNo].LevelVol[16];
                                if (iTrustSell > 0 && iTrustBuy>0 && ((iTrustBuy / iTrustSell) > 2.2))// || (iTrustSell / iTrustBuy) > 2.2))
                                {
                                    CandidateParty3Stock.Add(strStockNo);
                                    Console.WriteLine("信 ," + strStockNo + "," + dicStock3Party[strStockNo][dtLoopStart].name + "," + dbtrust + "," + dbEvenydayOLHCDic[strStockNo][dtLoopStart][3]);
                                }
                                if (iForBuy > 0 && iForSell>0 &&((iForBuy / iForSell) > 2.2))// || (iForSell / iForBuy) > 2.2))
                                {
                                    CandidateParty3Stock.Add(strStockNo);
                                    Console.WriteLine("外 ," + strStockNo + "," + dicStock3Party[strStockNo][dtLoopStart].name + "," + dbfor + "," + dbEvenydayOLHCDic[strStockNo][dtLoopStart][3]);
                                }
                                /*if (dbHedge >= dbPercent || dbHedge <= -dbPercent)
                                {
                                    CandidateParty3Stock.Add(strStockNo);
                                    Console.WriteLine("避 ," + strStockNo + "," + dicStock3Party[strStockNo][dtLoopStart].name + "," + dbHedge + "," + dbEvenydayOLHCDic[strStockNo][dtLoopStart][3]);
                                }*/
                            }
                        }
                    }

                }






            }    
        }

        private void button58_Click(object sender, EventArgs e)
        {
            DateTime dtStart = DateTime.Now;
            DateTime dtLoopEnd = dtStart.AddMonths(-12);
            DateTime dtLoopStart = dtStart;
            int iMaxContainDays = 300;


            string strBasePath = "D:\\Chips\\stock\\TDCC_OD_1-5_";
            int iCount = 8;
            Dictionary<string, StockLevelCount>[] weekLevelCount = new Dictionary<string, StockLevelCount>[iCount];
            int iCountWeek = 0;
            while (dtLoopStart > dtLoopEnd && iCountWeek < iCount)
            {
                string strcsv = strBasePath + dtLoopStart.ToString("yyyyMMdd") + ".csv";
                if (File.Exists(strcsv))
                {
                    weekLevelCount[iCountWeek] = new Dictionary<string, StockLevelCount>();
                    int iReadCount = 0;
                    IEnumerable<string> lines = File.ReadLines(strcsv);
                    foreach (string line in lines)
                    {
                        if (iReadCount > 0)
                        {
                            string[] strSplitLine = line.Split(',');
                            if (weekLevelCount[iCountWeek].ContainsKey(strSplitLine[1]))
                            {
                                int iLevel = int.Parse(strSplitLine[2]) - 1;
                                weekLevelCount[iCountWeek][strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                                weekLevelCount[iCountWeek][strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                                weekLevelCount[iCountWeek][strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                            }
                            else
                            {
                                weekLevelCount[iCountWeek].Add(strSplitLine[1], new StockLevelCount());
                                int iLevel = int.Parse(strSplitLine[2]) - 1;
                                weekLevelCount[iCountWeek][strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                                weekLevelCount[iCountWeek][strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                                weekLevelCount[iCountWeek][strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                            }
                        }
                        iReadCount++;
                    }
                    iCountWeek++;
                }
                dtLoopStart = dtLoopStart.AddDays(-1);
            }



            DateTime dtfail = DateTime.Now;
            foreach (string strno in dbEvenydayOLHCDic.Keys)
            {
                Dictionary<int, DateTime> backDaysmapDatetime = new Dictionary<int, DateTime>();
                double[] dbOpenArray = new double[iMaxContainDays];
                double[] dbLowArray = new double[iMaxContainDays];
                double[] dbHighArray = new double[iMaxContainDays];
                double[] dbCloseArray = new double[iMaxContainDays];
                long[] lvol = new long[iMaxContainDays];
                for (int i = 0; i < iMaxContainDays; i++)
                {
                    dbOpenArray[i] = 0;
                    dbLowArray[i] = 0;
                    dbHighArray[i] = 0;
                    dbCloseArray[i] = 0;
                    lvol[i] = 0;
                }
                int iloop = 0;
                DateTime dtchk = dtfail;
                foreach (DateTime dtloop in dbEvenydayOLHCDic[strno].Keys)
                {
                    if (dtchk != dtfail)
                    {
                        // 確認時間是照順序排
                        if (dtchk <= dtloop)
                        {
                            MessageBox.Show("False");
                        }
                    }
                    dtchk = dtloop;
                    if (dtchk <= dtStart)
                    {
                        backDaysmapDatetime.Add(iloop, dtloop);
                        lvol[iloop] = dbStockVolDic[strno][dtchk];
                        dbOpenArray[iloop] = dbEvenydayOLHCDic[strno][dtchk][0];
                        dbLowArray[iloop] = dbEvenydayOLHCDic[strno][dtchk][1];
                        dbHighArray[iloop] = dbEvenydayOLHCDic[strno][dtchk][2];
                        dbCloseArray[iloop] = dbEvenydayOLHCDic[strno][dtchk][3];
                        iloop++;
                        if (iloop == iMaxContainDays)
                            break;
                    }
                }
                double[] db5ary = new double[50];
                double[] db10ary = new double[50];
                double[] db20ary = new double[50];
                double[] db60ary = new double[50];
                double[] db120ary = new double[50];
                double[] db200ary = new double[50];
                double[] db240ary = new double[50];

                for (int icrossdays = 0; icrossdays < 50; icrossdays++)
                {
                    db5ary[icrossdays] = dbCloseArray.Skip(icrossdays).Take(5).Average();
                    db10ary[icrossdays] = dbCloseArray.Skip(icrossdays).Take(10).Average();
                    db20ary[icrossdays] = dbCloseArray.Skip(icrossdays).Take(20).Average();
                    db60ary[icrossdays] = dbCloseArray.Skip(icrossdays).Take(60).Average();
                    db120ary[icrossdays] = dbCloseArray.Skip(icrossdays).Take(120).Average();
                    db200ary[icrossdays] = dbCloseArray.Skip(icrossdays).Take(200).Average();
                    db240ary[icrossdays] = dbCloseArray.Skip(icrossdays).Take(240).Average();
                }

                double[] dbAllLineMaxMinPercentAry = new double[50];
                for (int icrossdays = 0; icrossdays < 50; icrossdays++)
                {
                    double[] dbAllLineAry = new double[7];
                    dbAllLineAry[0] = dbHighArray[icrossdays];
                    dbAllLineAry[1] = dbLowArray[icrossdays];
                    dbAllLineAry[2] = db5ary[icrossdays];
                    dbAllLineAry[3] = db10ary[icrossdays];
                    dbAllLineAry[4] = db20ary[icrossdays];
                    dbAllLineAry[5] = db60ary[icrossdays];
                    dbAllLineAry[6] = db120ary[icrossdays];
                    //dbAllLineAry[7] = db200ary[icrossdays];
                    dbAllLineMaxMinPercentAry[icrossdays] = (dbAllLineAry.Max() - dbAllLineAry.Min()) / dbAllLineAry.Min();

                }
                int iCount0_3 = 0;
                for (int i = 0; i < 20; i++)
                {
                    if (dbAllLineMaxMinPercentAry[i] <= 0.05 && dicStockNoMapping.ContainsKey(strno))
                    {
                        iCount0_3++;


                    }
                }

                double dbopenmax = dbOpenArray.Max();
                double dbopenmin = dbOpenArray.Min();
                double dbclosemax = dbOpenArray.Max();
                double dbclosemin = dbOpenArray.Min();
                double[] dbrealKhilow = new double[] { dbopenmax, dbopenmin, dbclosemax, dbclosemin };

                double db60openmax = dbOpenArray.Take(60).Max();
                double db60openmin = dbOpenArray.Take(60).Min();
                double db60closemax = dbOpenArray.Take(60).Max();
                double db60closemin = dbOpenArray.Take(60).Min();
                double[] dbrealK60hilow = new double[] { db60openmax, db60openmin, db60closemax, db60closemin };

                double db20openmax = dbOpenArray.Take(20).Max();
                double db20openmin = dbOpenArray.Take(20).Min();
                double db20closemax = dbOpenArray.Take(20).Max();
                double db20closemin = dbOpenArray.Take(20).Min();
                double[] dbrealK20hilow = new double[] { db20openmax, db20openmin, db20closemax, db20closemin };

                double[] db4avghilow = new double[] { db5ary[0], db10ary[0], db20ary[0], db60ary[0], db120ary[0], db240ary[0] };

                double dballhi = dbHighArray.Max();
                double dballlow = dbLowArray.Min();

                double dball60hi = dbHighArray.Take(60).Max();
                double dball60low = dbLowArray.Take(60).Min();

                double dball20hi = dbHighArray.Take(20).Max();
                double dball20low = dbLowArray.Take(20).Min();


                double dbperc = (dball60hi - dball60low) / (dballhi - dballlow);
                double dbperc20 = (dball20hi - dball20low) / (dball60hi - dball60low);

                double dbpercrealK = (dbrealK60hilow.Max() - dbrealK60hilow.Min()) / (dbrealKhilow.Max() - dbrealKhilow.Min());
                double dbperc20realK = (dbrealK20hilow.Max() - dbrealK20hilow.Min()) / (dbrealK60hilow.Max() - dbrealK60hilow.Min());

                double dbavghilowpercent = (db4avghilow.Max() - db4avghilow.Min()) / db4avghilow.Min();

                double dbreal0 = Math.Abs(dbOpenArray[0] - dbCloseArray[0]);
                double dbreal1 = Math.Abs(dbOpenArray[1] - dbCloseArray[1]);
                double dbreal2 = Math.Abs(dbOpenArray[2] - dbCloseArray[2]);

                double dbol0 = Math.Abs(dbOpenArray[0] - dbLowArray[0]);
                double dbol1 = Math.Abs(dbOpenArray[1] - dbLowArray[1]);
                double dbol2 = Math.Abs(dbOpenArray[2] - dbLowArray[2]);

                if (dicStockNoMapping.ContainsKey(strno) && dbCloseArray[0] > db60ary[0] && dbol0 > dbreal0 * 2 && dbol1 > dbreal1 * 2 &&
                    dbLowArray[0] > dbLowArray[1] && dbLowArray[1] > dbLowArray[2]
                    )
                {
                    string trmsg = string.Format("{0},{1},{2:F2},{3:F5},{4:F5}", strno, dicStockNoMapping[strno], dbCloseArray[0], dbperc20realK, dbavghilowpercent);
                    Console.WriteLine(trmsg);
                }

            }
        }

        private void button59_Click(object sender, EventArgs e)
        {

            // 投信120沒買  最近5天買
            DateTime dtStart = monthCalendar1.SelectionStart;
            DateTime dtLoopEnd = dtStart.AddMonths(-6);


            foreach(string strno in dicStock3Party.Keys)
            {
                long trusttotal5days=0;
                long fortotal5days = 0;
                DateTime dtLoopStart = dtStart;
                int icount5=0;
                if (dicStock3Party[strno].Count > 140)
                {
                    for (int iCount = 0; iCount < 20; )
                    {
                        if (dicStock3Party[strno].ContainsKey(dtLoopStart))
                        {
                            long trusttotal = long.Parse(dicStock3Party[strno][dtLoopStart].TrustTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                            trusttotal5days += trusttotal;
                            fortotal5days += long.Parse(dicStock3Party[strno][dtLoopStart].ForeigenTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                            iCount++;
                            if(trusttotal > 0)
                            {
                                icount5++;
                            }
                        }
                        dtLoopStart = dtLoopStart.AddDays(-1);
                    }

                    long trusttotal20days = 0;
                    long fortotal20days = 0;
                    for (int iCount = 0; iCount < 20; )
                    {
                        if (dicStock3Party[strno].ContainsKey(dtLoopStart))
                        {

                            trusttotal20days += long.Parse(dicStock3Party[strno][dtLoopStart].TrustTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                            fortotal20days += long.Parse(dicStock3Party[strno][dtLoopStart].ForeigenTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                            iCount++;
                        }
                        dtLoopStart = dtLoopStart.AddDays(-1);
                    }

                    long trusttotal100days = 0;
                    long fortotal100days = 0;
                    for (int iCount = 0; iCount < 100; )
                    {
                        if (dicStock3Party[strno].ContainsKey(dtLoopStart))
                        {

                            trusttotal100days += long.Parse(dicStock3Party[strno][dtLoopStart].TrustTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                            fortotal100days += long.Parse(dicStock3Party[strno][dtLoopStart].ForeigenTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                            iCount++;
                        }
                        dtLoopStart = dtLoopStart.AddDays(-1);
                    }




                    if (trusttotal100days >= 0 && trusttotal5days >= 0 && trusttotal20days >= 0 && trusttotal5days < trusttotal20days / 10)
                    {
                        string trmsg = string.Format("{0},{1}", strno, dicStockNoMapping[strno]);
                        Console.WriteLine(trmsg);
                    }
     
                }

       
            }

            int mmm = 0;
        }

        private void button60_Click(object sender, EventArgs e)
        {
            DateTime dtStart = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);
            DateTime dtLoopEnd = dtStart.AddMonths(-12);
            DateTime dtLoopStart = dtStart;
            int iMaxContainDays = 300;


            string strBasePath = "D:\\Chips\\stock\\TDCC_OD_1-5_";
            int iCount = 8;
            Dictionary<string, StockLevelCount>[] weekLevelCount = new Dictionary<string, StockLevelCount>[iCount];
            int iCountWeek = 0;
            while (dtLoopStart > dtLoopEnd && iCountWeek < iCount)
            {
                string strcsv = strBasePath + dtLoopStart.ToString("yyyyMMdd") + ".csv";
                if (File.Exists(strcsv))
                {
                    weekLevelCount[iCountWeek] = new Dictionary<string, StockLevelCount>();
                    int iReadCount = 0;
                    IEnumerable<string> lines = File.ReadLines(strcsv);
                    foreach (string line in lines)
                    {
                        if (iReadCount > 0)
                        {
                            string[] strSplitLine = line.Split(',');
                            if (weekLevelCount[iCountWeek].ContainsKey(strSplitLine[1]))
                            {
                                int iLevel = int.Parse(strSplitLine[2]) - 1;
                                weekLevelCount[iCountWeek][strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                                weekLevelCount[iCountWeek][strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                                weekLevelCount[iCountWeek][strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                            }
                            else
                            {
                                weekLevelCount[iCountWeek].Add(strSplitLine[1], new StockLevelCount());
                                int iLevel = int.Parse(strSplitLine[2]) - 1;
                                weekLevelCount[iCountWeek][strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                                weekLevelCount[iCountWeek][strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                                weekLevelCount[iCountWeek][strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                            }
                        }
                        iReadCount++;
                    }
                    iCountWeek++;
                }
                dtLoopStart = dtLoopStart.AddDays(-1);
            }



            DateTime dtfail = DateTime.Now;
            foreach (string strno in dbEvenydayOLHCDic.Keys)
            {
                Dictionary<int, DateTime> backDaysmapDatetime = new Dictionary<int, DateTime>();
                double[] dbOpenArray = new double[iMaxContainDays];
                double[] dbLowArray = new double[iMaxContainDays];
                double[] dbHighArray = new double[iMaxContainDays];
                double[] dbCloseArray = new double[iMaxContainDays];
                long[] lvol = new long[iMaxContainDays];
                for (int i = 0; i < iMaxContainDays; i++)
                {
                    dbOpenArray[i] = 0;
                    dbLowArray[i] = 0;
                    dbHighArray[i] = 0;
                    dbCloseArray[i] = 0;
                    lvol[i] = 0;
                }
                int iloop = 0;
                DateTime dtchk = dtfail;
                foreach (DateTime dtloop in dbEvenydayOLHCDic[strno].Keys)
                {
                    if (dtchk != dtfail)
                    {
                        // 確認時間是照順序排
                        if (dtchk <= dtloop)
                        {
                            MessageBox.Show("False");
                        }
                    }
                    dtchk = dtloop;
                    if (dtchk <= dtStart)
                    {
                        backDaysmapDatetime.Add(iloop, dtloop);
                        lvol[iloop] = dbStockVolDic[strno][dtchk];
                        dbOpenArray[iloop] = dbEvenydayOLHCDic[strno][dtchk][0];
                        dbLowArray[iloop] = dbEvenydayOLHCDic[strno][dtchk][1];
                        dbHighArray[iloop] = dbEvenydayOLHCDic[strno][dtchk][2];
                        dbCloseArray[iloop] = dbEvenydayOLHCDic[strno][dtchk][3];
                        iloop++;
                        if (iloop == iMaxContainDays)
                            break;
                    }
                }
                double[] db5ary = new double[50];
                double[] db10ary = new double[50];
                double[] db20ary = new double[50];
                double[] db60ary = new double[50];
                double[] db120ary = new double[50];
                double[] db200ary = new double[50];
                double[] db240ary = new double[50];

                for (int icrossdays = 0; icrossdays < 50; icrossdays++)
                {
                    db5ary[icrossdays] = dbCloseArray.Skip(icrossdays).Take(5).Average();
                    db10ary[icrossdays] = dbCloseArray.Skip(icrossdays).Take(10).Average();
                    db20ary[icrossdays] = dbCloseArray.Skip(icrossdays).Take(20).Average();
                    db60ary[icrossdays] = dbCloseArray.Skip(icrossdays).Take(60).Average();
                    db120ary[icrossdays] = dbCloseArray.Skip(icrossdays).Take(120).Average();
                    db200ary[icrossdays] = dbCloseArray.Skip(icrossdays).Take(200).Average();
                    db240ary[icrossdays] = dbCloseArray.Skip(icrossdays).Take(240).Average();
                }
                int iChangeDirCount = 0;
                bool bull = false;
                double[] dbAllLineMaxMinPercentAry = new double[50];
                for (int icrossdays = 0; icrossdays < 50; icrossdays++)
                {
                    if(icrossdays > 0)
                    {
                        bool bb = false;
                        if(db60ary[icrossdays-1] >=db60ary[icrossdays])
                        {
                            bb = true;
                        }
                        else{
                            bb = false;
                        }
                        if(bull != bb)
                        {
                            bull = bb;
                            iChangeDirCount++;
                        }
                    }
                    

                    double[] dbAllLineAry = new double[7];
                    dbAllLineAry[0] = dbHighArray[icrossdays];
                    dbAllLineAry[1] = dbLowArray[icrossdays];
                    dbAllLineAry[2] = db5ary[icrossdays];
                    dbAllLineAry[3] = db10ary[icrossdays];
                    dbAllLineAry[4] = db20ary[icrossdays];
                    dbAllLineAry[5] = db60ary[icrossdays];
                    dbAllLineAry[6] = db120ary[icrossdays];
                    //dbAllLineAry[7] = db200ary[icrossdays];
                    dbAllLineMaxMinPercentAry[icrossdays] = (dbAllLineAry.Max() - dbAllLineAry.Min()) / dbAllLineAry.Min();

                }
                int iCount0_3 = 0;
                for (int i = 0; i < 20; i++)
                {
                    if (dbAllLineMaxMinPercentAry[i] <= 0.05 && dicStockNoMapping.ContainsKey(strno))
                    {
                        iCount0_3++;


                    }
                }

                int iRange = 20;
                long l3PTotal = 0;
                long lHedgeTotal = 0;
                long lTrust = 0;
                long lSelf = 0;
                long lTotal = 0;
                long lFor = 0;

                long ladd = 0;
                int iLoopDay = 0;
                dtLoopStart = dtStart;
                while (iLoopDay < iRange)  // 計算最近幾天
                {
                    if (workingdays.Contains(dtLoopStart))
                    {
                        if (dicStock3Party.ContainsKey(strno))
                        {
                            if (dicStock3Party[strno].ContainsKey(dtLoopStart))
                            {
                                int iTotal = int.Parse(dicStock3Party[strno][dtLoopStart].Party3Total, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                int iHedge = int.Parse(dicStock3Party[strno][dtLoopStart].SelfHedgeTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                int iHedgeBuy = int.Parse(dicStock3Party[strno][dtLoopStart].SelfHedgeBuy, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                int iHedgeSell = int.Parse(dicStock3Party[strno][dtLoopStart].SelfHedgeSell, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                int iForeigen = int.Parse(dicStock3Party[strno][dtLoopStart].ForeigenTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                int iTrust = int.Parse(dicStock3Party[strno][dtLoopStart].TrustTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                int iSelf = int.Parse(dicStock3Party[strno][dtLoopStart].SelfSelfTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                ladd += (iTotal - iHedge); // 計算不含避險3大買量
                                //ladd += iTotal; // 全算
                                //ladd += (iTrust + iForeigen); //外 信
                                //ladd += iHedge;

                                lTotal += (iTotal - iHedge);
                                l3PTotal += iTotal;
                                lHedgeTotal += iHedge;
                                lTrust += iTrust;
                                lSelf += iSelf;
                                lFor += iForeigen;
                            }
                        }
                        iLoopDay++;
                    }

                    dtLoopStart = dtLoopStart.AddDays(-1);
                }


                double dbopenmax = dbOpenArray.Max();
                double dbopenmin = dbOpenArray.Min();
                double dbclosemax = dbOpenArray.Max();
                double dbclosemin = dbOpenArray.Min();
                double[] dbrealKhilow = new double[] { dbopenmax, dbopenmin, dbclosemax, dbclosemin };

                double db60openmax = dbOpenArray.Take(60).Max();
                double db60openmin = dbOpenArray.Take(60).Min();
                double db60closemax = dbOpenArray.Take(60).Max();
                double db60closemin = dbOpenArray.Take(60).Min();
                double[] dbrealK60hilow = new double[] { db60openmax, db60openmin, db60closemax, db60closemin };

                double db20openmax = dbOpenArray.Take(20).Max();
                double db20openmin = dbOpenArray.Take(20).Min();
                double db20closemax = dbOpenArray.Take(20).Max();
                double db20closemin = dbOpenArray.Take(20).Min();
                double[] dbrealK20hilow = new double[] { db20openmax, db20openmin, db20closemax, db20closemin };

                double[] db4avghilow = new double[] { db5ary[0], db10ary[0], db20ary[0], db60ary[0], db120ary[0], db240ary[0] };

                double dballhi = dbHighArray.Max();
                double dballlow = dbLowArray.Min();

                double dball60hi = dbHighArray.Take(60).Max();
                double dball60low = dbLowArray.Take(60).Min();

                double dball20hi = dbHighArray.Take(20).Max();
                double dball20low = dbLowArray.Take(20).Min();


                double dbperc = (dball60hi - dball60low) / (dballhi - dballlow);
                double dbperc20 = (dball20hi - dball20low) / (dball60hi - dball60low);

                double dbpercrealK = (dbrealK60hilow.Max() - dbrealK60hilow.Min()) / (dbrealKhilow.Max() - dbrealKhilow.Min());
                double dbperc20realK = (dbrealK20hilow.Max() - dbrealK20hilow.Min()) / (dbrealK60hilow.Max() - dbrealK60hilow.Min());

                double dbavghilowpercent = (db4avghilow.Max() - db4avghilow.Min()) / db4avghilow.Min();

                if (dicStockNoMapping.ContainsKey(strno) && lvol.Average() > 500000 && weekLevelCount[0].ContainsKey(strno) && weekLevelCount[3].ContainsKey(strno) &&
                    weekLevelCount[0][strno].LevelPeople[1] <  weekLevelCount[3][strno].LevelPeople[1] &&
                    db60ary[0] > db60ary[1] && dbCloseArray[0] > db60ary[0] && lTrust > 0 && lFor>0
                    )
                {
                    string trmsg = string.Format("{0},{1},{2}", strno, dicStockNoMapping[strno], iChangeDirCount);
                    Console.WriteLine(trmsg);
                }

            }
        }

        private void button61_Click(object sender, EventArgs e)
        {
            //https://www.twse.com.tw//exchangeReport/MI_5MINS?response=json&date=20191227
            string strSavePath = "D:\\Chips\\5MINS\\";
            DateTime dt = DateTime.Now;


            for (int i = 0; i < 20; i++)
            {
                if (!checkBoxToday.Checked)
                    dt = dt.AddDays(-1);
                string strStockMon = string.Format("{0:0000}{1:00}{2:00}_5MINS.json", dt.Year, dt.Month, dt.Day);
                string strLocalFile = strSavePath + strStockMon;
                if (!File.Exists(strLocalFile))
                {

                    string url = string.Format("https://www.twse.com.tw//exchangeReport/MI_5MINS?response=json&date={0:0000}{1:00}{2:00}", dt.Year, dt.Month, dt.Day);

                    HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
                    httpWebRequest.ContentType = "application/json";
                    httpWebRequest.Method = WebRequestMethods.Http.Get;
                    httpWebRequest.Accept = "application/json";
                    string text;
                    var response = (HttpWebResponse)httpWebRequest.GetResponse();

                    using (var sr = new StreamReader(response.GetResponseStream()))
                    {
                        text = sr.ReadToEnd();
                    }

                    File.WriteAllText(strLocalFile, text);

                    response.Close();

                    httpWebRequest.Abort();

                    Thread.Sleep(5000);
                }


            }
        }

        private void button62_Click(object sender, EventArgs e)
        {

            DateTime dtStart = new DateTime(DateTime.Now.Year,DateTime.Now.Month,DateTime.Now.Day-1);
            DateTime dtLoopEnd = dtStart.AddMonths(-12);
            DateTime dtLoopStart = dtStart;


            int icountweek = 8;
            //Queue<string> qcsvfile = new Queue<string>();
            string strBasePath = "D:\\Chips\\stock\\TDCC_OD_1-5_";
            int iCount = 32;
            Dictionary<string, StockLevelCount>[] weekLevelCount = new Dictionary<string, StockLevelCount>[iCount];
            int iCountWeek = 0;
            while (dtLoopStart > dtLoopEnd && iCountWeek<icountweek)
            {
                string strcsv = strBasePath + dtLoopStart.ToString("yyyyMMdd") + ".csv";
                if (File.Exists(strcsv))
                {
                    weekLevelCount[iCountWeek] = new Dictionary<string, StockLevelCount>();
                    int iReadCount = 0;
                    IEnumerable<string> lines = File.ReadLines(strcsv);
                    foreach (string line in lines)
                    {
                        if (iReadCount > 0)
                        {
                            string[] strSplitLine = line.Split(',');
                            if (weekLevelCount[iCountWeek].ContainsKey(strSplitLine[1]))
                            {
                                int iLevel = int.Parse(strSplitLine[2]) - 1;
                                weekLevelCount[iCountWeek][strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                                weekLevelCount[iCountWeek][strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                                weekLevelCount[iCountWeek][strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                            }
                            else
                            {
                                weekLevelCount[iCountWeek].Add(strSplitLine[1], new StockLevelCount());
                                int iLevel = int.Parse(strSplitLine[2]) - 1;
                                weekLevelCount[iCountWeek][strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                                weekLevelCount[iCountWeek][strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                                weekLevelCount[iCountWeek][strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                            }
                        }
                        iReadCount++;
                    }
                    iCountWeek++;
                }
                dtLoopStart = dtLoopStart.AddDays(-1);
            }

            // 人數持續減2個月的股價變化
            foreach (KeyValuePair<string, StockLevelCount> item in weekLevelCount[0])
            {
                if (dbEvenydayOLHCDic.ContainsKey(item.Key) && dicStockNoMapping.ContainsKey(item.Key) && weekLevelCount[1].ContainsKey(item.Key) && weekLevelCount[2].ContainsKey(item.Key) && weekLevelCount[3].ContainsKey(item.Key) && weekLevelCount[4].ContainsKey(item.Key) && weekLevelCount[5].ContainsKey(item.Key) && weekLevelCount[6].ContainsKey(item.Key) && weekLevelCount[7].ContainsKey(item.Key))
                {
                    double dbpeople0 = weekLevelCount[0][item.Key].LevelPeople[16] - weekLevelCount[0][item.Key].LevelPeople[0];
                    double dbpeople1 = weekLevelCount[1][item.Key].LevelPeople[16] - weekLevelCount[1][item.Key].LevelPeople[0];
                    double dbpeople2 = weekLevelCount[2][item.Key].LevelPeople[16] - weekLevelCount[2][item.Key].LevelPeople[0];
                    double dbpeople3 = weekLevelCount[3][item.Key].LevelPeople[16] - weekLevelCount[3][item.Key].LevelPeople[0];
                    double dbpeople4 = weekLevelCount[4][item.Key].LevelPeople[16] - weekLevelCount[4][item.Key].LevelPeople[0];
                    double dbpeople5 = weekLevelCount[5][item.Key].LevelPeople[16] - weekLevelCount[5][item.Key].LevelPeople[0];
                    double dbpeople6 = weekLevelCount[6][item.Key].LevelPeople[16] - weekLevelCount[6][item.Key].LevelPeople[0];
                    double dbpeople7 = weekLevelCount[7][item.Key].LevelPeople[16] - weekLevelCount[7][item.Key].LevelPeople[0];


                    if (dbpeople0 < dbpeople1 && dbpeople1 < dbpeople2 && dbpeople2 < dbpeople3 && dbpeople3 < dbpeople4 && dbpeople4 < dbpeople5 && dbpeople5 < dbpeople6 && dbpeople6 < dbpeople7)
                    {
                        DateTime dtLatest = dtStart;
                        DateTime dtAge = dtLoopStart;

                        if (dbEvenydayOLHCDic[item.Key].ContainsKey(dtLatest) && dbEvenydayOLHCDic[item.Key].ContainsKey(dtAge))
                        {
                            string strhigh = string.Format("{0},{1},{2},{3},{4}", item.Key, dicStockNoMapping[item.Key], dbEvenydayOLHCDic[item.Key][dtLatest][3], dbEvenydayOLHCDic[item.Key][dtAge][3], dbEvenydayOLHCDic[item.Key][dtLatest][3]-dbEvenydayOLHCDic[item.Key][dtAge][3]);
  
                            Console.WriteLine(strhigh);
                        }
                        else
                        {
                            
                            int mmm = 0;
                        }
                    }
                }
            }
        }

        private void button63_Click(object sender, EventArgs e)
        {
            int iRange = 4;  //iRange天內
            double dbPercent = 0.015; // 成長數大於總股數
            double dbVolPercent = 0.25; // 成交量占比
            DateTime dtStart = monthCalendar1.SelectionStart;
            DateTime dtLoopStart = monthCalendar1.SelectionStart;

            string strMsg = string.Format("找{0}前{1}天內三大買超佔成交比重{2:F4}", dtStart.ToString("MM/dd"), iRange, dbVolPercent);
            Console.WriteLine(strMsg);

            CandidateParty3Stock.Clear();

            //CandidateWeekChipSet.Clear();            
            string strBasePath = "D:\\Chips\\stock\\TDCC_OD_1-5_";
            string strReadFile = "";
            int icountchip = 0;
            while (true)
            {
                string strcsv = strBasePath + dtLoopStart.ToString("yyyyMMdd") + ".csv";
                if (File.Exists(strcsv))
                {
                    strReadFile = strcsv;
                    break;
                }
                dtLoopStart = dtLoopStart.AddDays(-1);
            }
            Dictionary<string, StockLevelCount> weekLevelCount = new Dictionary<string, StockLevelCount>();
            int iReadCount = 0;
            IEnumerable<string> lines = File.ReadLines(strReadFile);
            foreach (string line in lines)  // 取流通在外數
            {
                if (iReadCount > 0)
                {
                    string[] strSplitLine = line.Split(',');
                    if (weekLevelCount.ContainsKey(strSplitLine[1]))
                    {
                        int iLevel = int.Parse(strSplitLine[2]) - 1;
                        weekLevelCount[strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                        weekLevelCount[strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                        weekLevelCount[strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                    }
                    else
                    {
                        weekLevelCount.Add(strSplitLine[1], new StockLevelCount());
                        int iLevel = int.Parse(strSplitLine[2]) - 1;
                        weekLevelCount[strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                        weekLevelCount[strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                        weekLevelCount[strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                    }
                }
                iReadCount++;
            }

            Dictionary<string, long> stock3PartyVol = new Dictionary<string, long>();
            foreach (string strStockNo in weekLevelCount.Keys)
            {
                long l3PTotal = 0;
                long lHedgeTotal = 0;
                long lTrust = 0;
                long lSelf = 0;
                long lTotal = 0;
                long lFor = 0;
                long lVol = 0;
                long ladd = 0;
                int iLoopDay = 0;
                dtLoopStart = dtStart;
                while (iLoopDay < iRange)  // 計算最近幾天
                {
                    if (workingdays.Contains(dtLoopStart))
                    {
                        if (dicStock3Party.ContainsKey(strStockNo))
                        {
                            if (dicStock3Party[strStockNo].ContainsKey(dtLoopStart))
                            {
                                int iTotal = int.Parse(dicStock3Party[strStockNo][dtLoopStart].Party3Total, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                int iHedge = int.Parse(dicStock3Party[strStockNo][dtLoopStart].SelfHedgeTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                int iHedgeBuy = int.Parse(dicStock3Party[strStockNo][dtLoopStart].SelfHedgeBuy, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                int iHedgeSell = int.Parse(dicStock3Party[strStockNo][dtLoopStart].SelfHedgeSell, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                int iForeigen = int.Parse(dicStock3Party[strStockNo][dtLoopStart].ForeigenTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                int iTrust = int.Parse(dicStock3Party[strStockNo][dtLoopStart].TrustTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                int iSelf = int.Parse(dicStock3Party[strStockNo][dtLoopStart].SelfSelfTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                ladd += (iTotal - iHedge); // 計算不含避險3大買量
                                //ladd += iTotal; // 全算
                                //ladd += (iTrust + iForeigen); //外 信
                                //ladd += iHedge;

                                lTotal += (iTotal - iHedge);
                                l3PTotal += iTotal;
                                lHedgeTotal += iHedge;
                                lTrust += iTrust;
                                lSelf += iSelf;
                                lFor += iForeigen;

                                lVol += dbStockVolDic[strStockNo][dtLoopStart];
                            }
                        }
                        iLoopDay++;
                    }

                    dtLoopStart = dtLoopStart.AddDays(-1);
                }

                stock3PartyVol.Add(strStockNo, ladd);

                if (weekLevelCount.ContainsKey(strStockNo))
                {
                    double dbper = (double)stock3PartyVol[strStockNo] / (double)weekLevelCount[strStockNo].LevelVol[16];
                    double dbpervol = (double)stock3PartyVol[strStockNo] / (double)lVol;
                    if (dbpervol > dbVolPercent) // 成長數
                    {
                        CandidateParty3Stock.Add(strStockNo);
                        string strmsg = string.Format("{0},{1:F4},{2:F4},{3},{4},外:{5},信:{6}", strStockNo, dbper, dbpervol, dicStockNoMapping[strStockNo], dbEvenydayOLHCDic[strStockNo][dtStart][3], lFor, lTrust);
                        Console.WriteLine(strmsg);
                    }
                }
            }    
        }

        private void button64_Click(object sender, EventArgs e)
        {
            DateTime dtLat = new DateTime(2020, 7, 3);
            DateTime dtBeg = new DateTime(2020, 7, 6);
            DateTime dtEnd = new DateTime(2020, 7, 10);
            string strReadFile = "D:\\Chips\\stock\\TDCC_OD_1-5_";
            strReadFile += dtLat.ToString("yyyyMMdd") + ".csv";
            Dictionary<string, StockLevelCount>[] weekLevelCount = new Dictionary<string, StockLevelCount>[2];
            int iReadCount = 0;
            IEnumerable<string> lines = File.ReadLines(strReadFile);
            weekLevelCount[0] = new  Dictionary<string, StockLevelCount>();
            foreach (string line in lines)  // 取流通在外數
            {
                if (iReadCount > 0)
                {
                    string[] strSplitLine = line.Split(',');
                    if (weekLevelCount[0].ContainsKey(strSplitLine[1]))
                    {
                        int iLevel = int.Parse(strSplitLine[2]) - 1;
                        weekLevelCount[0][strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                        weekLevelCount[0][strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                        weekLevelCount[0][strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                    }
                    else
                    {
                        weekLevelCount[0].Add(strSplitLine[1], new StockLevelCount());
                        int iLevel = int.Parse(strSplitLine[2]) - 1;
                        weekLevelCount[0][strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                        weekLevelCount[0][strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                        weekLevelCount[0][strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                    }
                }
                iReadCount++;
            }


            string strReadFile1 = "D:\\Chips\\stock\\TDCC_OD_1-5_";
            strReadFile1 += dtEnd.ToString("yyyyMMdd") + ".csv"; ;
            iReadCount = 0;
            IEnumerable<string> lines1 = File.ReadLines(strReadFile1);
             weekLevelCount[1] = new  Dictionary<string, StockLevelCount>();
            foreach (string line in lines1)  // 取流通在外數
            {
                if (iReadCount > 0)
                {
                    string[] strSplitLine = line.Split(',');
                    if (weekLevelCount[1].ContainsKey(strSplitLine[1]))
                    {
                        int iLevel = int.Parse(strSplitLine[2]) - 1;
                        weekLevelCount[1][strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                        weekLevelCount[1][strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                        weekLevelCount[1][strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                    }
                    else
                    {
                        weekLevelCount[1].Add(strSplitLine[1], new StockLevelCount());
                        int iLevel = int.Parse(strSplitLine[2]) - 1;
                        weekLevelCount[1][strSplitLine[1]].LevelPeople[iLevel] = long.Parse(strSplitLine[3]);
                        weekLevelCount[1][strSplitLine[1]].LevelVol[iLevel] = long.Parse(strSplitLine[4]);
                        weekLevelCount[1][strSplitLine[1]].LevelRate[iLevel] = double.Parse(strSplitLine[5]);
                    }
                }
                iReadCount++;
            }

            
            // 找出上週高點與收盤跌 10% , 但大持股+ 小股東- , 投信+  =  甩較  ， 結論: 本周開盤找低點買
            foreach(string strno in dbEvenydayOLHCDic.Keys)
            {
                double dbhigh = 0;

                double close = 0;
                if(dbEvenydayOLHCDic[strno].ContainsKey(dtBeg) && dbEvenydayOLHCDic[strno].ContainsKey(dtEnd))
                {
                    long lTrust = 0;
                    long lFor = 0;
                    long lSelf = 0;
                    foreach (DateTime dtDay in dbEvenydayOLHCDic[strno].Keys)
                    {
                        if (dtDay >= dtBeg && dtDay <= dtEnd)
                        {
                            if (dbEvenydayOLHCDic[strno][dtDay][2] > dbhigh)
                                dbhigh = dbEvenydayOLHCDic[strno][dtDay][2];

                            if (dicStock3Party.ContainsKey(strno) && dicStock3Party[strno].ContainsKey(dtDay))
                            {
                                int itruet = int.Parse(dicStock3Party[strno][dtDay].TrustTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                int iself = int.Parse(dicStock3Party[strno][dtDay].SelfSelfTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                                int ifor = int.Parse(dicStock3Party[strno][dtDay].ForeigenTotal, NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);

                                lTrust += itruet;
                                lSelf += iself;
                                lFor += ifor;
                            }

                        }
                    }
                    close = dbEvenydayOLHCDic[strno][dtEnd][3];

                    if(weekLevelCount[0].ContainsKey(strno) && weekLevelCount[1].ContainsKey(strno)  )
                    {
                        //大持股+  
                        double dbrate0 = weekLevelCount[0][strno].LevelRate[11] + weekLevelCount[0][strno].LevelRate[12] + weekLevelCount[0][strno].LevelRate[13] + weekLevelCount[0][strno].LevelRate[14];
                        double dbrate1 = weekLevelCount[0][strno].LevelRate[11] + weekLevelCount[1][strno].LevelRate[12] + weekLevelCount[1][strno].LevelRate[13] + weekLevelCount[1][strno].LevelRate[14];
                        //小股東-
                        double dbpeople0 = weekLevelCount[0][strno].LevelPeople[1] + weekLevelCount[0][strno].LevelPeople[2] + weekLevelCount[0][strno].LevelPeople[3];
                        double dbpeople1 = weekLevelCount[1][strno].LevelPeople[1] + weekLevelCount[1][strno].LevelPeople[2] + weekLevelCount[1][strno].LevelPeople[3];

                        if (dbhigh > close * 1.05 && dicStockNoMapping.ContainsKey(strno) && dbrate1 > dbrate0 && dbpeople0 > dbpeople1 && lTrust > 0)
                        {
                            string strmsg = string.Format("{0},{1},{2},{3},{4}", strno, dicStockNoMapping[strno], lFor, lTrust, lSelf);
                            Console.WriteLine(strmsg);
                        }
                    }

                }


            }
        }
        

        // Root myDeserializedClass = JsonConvert.DeserializeObject<Root>(myJsonResponse); 
        public class JSON_Root
        {
            public string stat { get; set; }
            public string title { get; set; }
            public string date { get; set; }
            public List<string> fields { get; set; }
            public List<List<string>> data { get; set; }
        }


        public class JSON_Root_5MINS
        {
            /// <summary>
            /// 
            /// </summary>
            public string stat { get; set; }
            /// <summary>
            /// 
            /// </summary>
            public string date { get; set; }
            /// <summary>
            /// 105年07月11日每5秒委託成交統計
            /// </summary>
            public string title { get; set; }
            /// <summary>
            /// 
            /// </summary>
            public List<string> fields { get; set; }
            /// <summary>
            /// 
            /// </summary>
            public List<List<string>> data { get; set; }
            /// <summary>
            /// 
            /// </summary>
            public List<string> notes { get; set; }
        }



        public class Group_MARGN
        {
            public int start { get; set; }
            public int span { get; set; }
            public string title { get; set; }
        }

        public class MARGN_DAY
        {
            public string stat { get; set; }
            public string creditTitle { get; set; }
            public List<string> creditFields { get; set; }
            public List<List<string>> creditList { get; set; }
            public List<string> creditNotes { get; set; }
            public string title { get; set; }
            public string fields { get; set; }
            public List<Group_MARGN> groups { get; set; }
            public List<object> notes { get; set; }
            public string data { get; set; }
            public string date { get; set; }
            public string selectType { get; set; }

       
        }
        public class ParamsOf3pty
        {
            public string response { get; set; }
            public string dayDate { get; set; }
            public string type { get; set; }
            public string controller { get; set; }
            public object format { get; set; }
            public string action { get; set; }
            public string lang { get; set; }
            public string monthDate { get; set; }
            public string weekDate { get; set; }
        }

  
        public class ObjectOf3pty
        {
            public string stat { get; set; }
            public string title { get; set; }
            public List<string> fields { get; set; }
            public string date { get; set; }
            public List<List<string>> data { get; set; }
            [JsonProperty(PropertyName = "params")]
            public ParamsOf3pty parameters { get; set; }
            public List<string> notes { get; set; }
        }

        private void button65_Click(object sender, EventArgs e)
        {
            Dictionary<DateTime, double> dtHi = new Dictionary<DateTime,double>();
            Dictionary<DateTime, double> dtLo = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtClose = new Dictionary<DateTime, double>();

            Dictionary<DateTime, double> dtDiffHi = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtDiffLo = new Dictionary<DateTime, double>();

            Dictionary<DateTime, double> dtDiffHi2Days = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtDiffLo2Days = new Dictionary<DateTime, double>();

            Dictionary<DateTime, double> dtDiffHi1Day = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtDiffLo1Day = new Dictionary<DateTime, double>();

            Dictionary<DateTime, double> dtBuyNOTBeg = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtBuyQutBeg = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtSelNOTBeg = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtSelQutBeg = new Dictionary<DateTime, double>();

            Dictionary<DateTime, double> dtBuyNOTEnd = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtBuyQutEnd = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtSelNOTEnd = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtSelQutEnd = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtCumTEnd   = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtCumQEnd   = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtCumPEnd   = new Dictionary<DateTime, double>();

            Dictionary<DateTime, double> dtEXRateUSDtoNTD = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtEXRateRMBtoNTD = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtEXRateEURtoUSD = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtEXRateEURtoJPD = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtEXRateGBPtoUSD = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtEXRateAUDtoUSD = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtEXRateUSDtoHKD = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtEXRateUSDtoRMB = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtEXRateUSDtoSAD = new Dictionary<DateTime, double>();

            Dictionary<DateTime, double> dtFinBuyQut = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtFinSelQut = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtFinCashReturn = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtFinTodayQut = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtFinYesterQut = new Dictionary<DateTime, double>();

            Dictionary<DateTime, double> dtMarBuyQut = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtMarSelQut = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtMarCashReturn = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtMarTodayQut = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtMarYesterQut = new Dictionary<DateTime, double>();

            Dictionary<DateTime, double> dtFinBuyPri = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtFinSelPri = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtFinCashReturnPri = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtFinTodayPri = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtFinYesterPri = new Dictionary<DateTime, double>();


            for (int iy = 2008; iy <= 2020; iy++)
            {
                string strfilename = string.Format("D:\\Chips\\EXRate\\{0}.csv", iy);

                IEnumerable<string> lines = File.ReadLines(strfilename);
                foreach (string line in lines)
                {
                    string[] strAtLine = line.Split(',');
                    if (strAtLine.Length == 11)
                    {
                        string strdate = strAtLine[0];
                        DateTime parsed;
                        if (DateTime.TryParse(strdate, out parsed))
                        {
                            double EXRateUSDtoNTD = double.Parse(strAtLine[1], NumberStyles.Float);
                            double EXRateRMBtoNTD = 5;
                            if(strAtLine[2] != "-")
                                EXRateRMBtoNTD = double.Parse(strAtLine[2], NumberStyles.Float);
                            double EXRateEURtoUSD = double.Parse(strAtLine[3], NumberStyles.Float);
                            double EXRateEURtoJPD = double.Parse(strAtLine[4], NumberStyles.Float);
                            double EXRateGBPtoUSD = double.Parse(strAtLine[5], NumberStyles.Float);
                            double EXRateAUDtoUSD = double.Parse(strAtLine[6], NumberStyles.Float);
                            double EXRateUSDtoHKD = double.Parse(strAtLine[7], NumberStyles.Float);
                            double EXRateUSDtoRMB = 6.22805;
                            if(strAtLine[8] != "-")
                                EXRateUSDtoRMB = double.Parse(strAtLine[8], NumberStyles.Float);
                            double EXRateUSDtoSAD = 11.32215;
                            if(strAtLine[9] != "-")
                                 EXRateUSDtoSAD = double.Parse(strAtLine[9], NumberStyles.Float);

                            dtEXRateUSDtoNTD.Add(parsed, EXRateUSDtoNTD);
                            dtEXRateRMBtoNTD.Add(parsed, EXRateRMBtoNTD);
                            dtEXRateEURtoUSD.Add(parsed, EXRateEURtoUSD);
                            dtEXRateEURtoJPD.Add(parsed, EXRateEURtoJPD);
                            dtEXRateGBPtoUSD.Add(parsed, EXRateGBPtoUSD);
                            dtEXRateAUDtoUSD.Add(parsed, EXRateAUDtoUSD);
                            dtEXRateUSDtoHKD.Add(parsed, EXRateUSDtoHKD);
                            dtEXRateUSDtoRMB.Add(parsed, EXRateUSDtoRMB);
                            dtEXRateUSDtoSAD.Add(parsed, EXRateUSDtoSAD);
                            // Use reformatted
                        }
                        else
                        {
                            int nn = 0;
                            // Log error, perhaps?
                        }
                    }
                }
            }
            string strbase = "D:\\Chips\\TWSE\\";
            List<string> listfiles = new List<string>();
            DateTime dtLoop = new DateTime(2008, 1, 1);
            for (; dtLoop < new DateTime(2020, 8, 31); dtLoop = dtLoop.AddMonths(1))
            {
                string strfilename = string.Format("{0}MI_5MINS_HIST_{1}.json", strbase, dtLoop.ToString("yyyyMMdd"));
                listfiles.Add(strfilename);

                string json = File.ReadAllText(strfilename);
                JSON_Root json_Dictionary = JsonConvert.DeserializeObject<JSON_Root>(json);
                List<List<string>> lstlist = json_Dictionary.data;
                foreach(List<string> contextinlst in lstlist)
                {
                    string[] YYYMMDD = contextinlst[0].Split('/');
                    int iyyy = int.Parse(YYYMMDD[0]);
                    int imm = int.Parse(YYYMMDD[1]);
                    int idd = int.Parse(YYYMMDD[2]);
                    //DateTime dt = DateTime.ParseExact(contextinlst[0], "yyy/MM/dd", CultureInfo.InvariantCulture).AddYears(1911);
                    DateTime dt = new DateTime(1911 + iyyy, imm, idd);
                    double dbOpen = double.Parse(contextinlst[1], NumberStyles.Float | NumberStyles.AllowThousands);
                    double dbHi= double.Parse(contextinlst[2], NumberStyles.Float | NumberStyles.AllowThousands);
                    double dbLo = double.Parse(contextinlst[3], NumberStyles.Float | NumberStyles.AllowThousands);
                    double dbClose= double.Parse(contextinlst[4], NumberStyles.Float | NumberStyles.AllowThousands);
                    dtHi.Add(dt, dbHi);
                    dtLo.Add(dt, dbLo);
                    dtClose.Add(dt, dbClose);
  
                }
            }
            DateTime dtEnd = new DateTime(2020, 8, 27);

            foreach(DateTime dtin in dtClose.Keys)
            {
                if (dtin < dtEnd)
                {
                    DateTime dtinstart = dtin;
                    int iNext2Count = 0;
                    while (iNext2Count < 3)
                    {
                        DateTime dtnext = dtinstart.AddDays(1);
                        if (dtClose.ContainsKey(dtnext))
                        {
                            iNext2Count++;
                            double dfhi = dtHi[dtnext] - dtClose[dtin];
                            double dflo = dtClose[dtin] - dtLo[dtnext];
                            if (iNext2Count==1)
                            {
                                dtDiffHi1Day.Add(dtin, dfhi);
                                dtDiffLo1Day.Add(dtin, dflo);
                            }
                            if (iNext2Count == 2)
                            {
                                if (dtDiffHi2Days.ContainsKey(dtin))
                                {
                                    if (dfhi > dtDiffHi2Days[dtin])
                                    {
                                        dtDiffHi2Days[dtin] = dfhi;
                                    }
                                }
                                else
                                {
                                    dtDiffHi2Days.Add(dtin, dfhi);
                                }

                                if (dtDiffLo2Days.ContainsKey(dtin))
                                {
                                    if (dflo > dtDiffLo2Days[dtin])
                                    {
                                        dtDiffLo2Days[dtin] = dflo;
                                    }
                                }
                                else
                                {
                                    dtDiffLo2Days.Add(dtin, dflo);
                                }
                            }
                            if (dtDiffHi.ContainsKey(dtin))
                            {
                                if (dfhi > dtDiffHi[dtin])
                                {
                                    dtDiffHi[dtin] = dfhi;
                                }
                            }
                            else
                            {
                                dtDiffHi.Add(dtin, dfhi);
                            }

                            if (dtDiffLo.ContainsKey(dtin))
                            {
                                if (dflo > dtDiffLo[dtin])
                                {
                                    dtDiffLo[dtin] = dflo;
                                }
                            }
                            else
                            {
                                dtDiffLo.Add(dtin, dflo);
                            }
                        }
                        dtinstart = dtnext;
                    }


                }

                {
                    string strfilename = string.Format("D:\\Chips\\TWSE5MINS\\MI_5MINS_{0}.json", dtin.ToString("yyyyMMdd"));
                    listfiles.Add(strfilename);
                    string json = File.ReadAllText(strfilename);
                    JSON_Root_5MINS json_Dictionary = JsonConvert.DeserializeObject<JSON_Root_5MINS>(json);
                    double BuyNOT = double.Parse(json_Dictionary.data[0][1], NumberStyles.Float | NumberStyles.AllowThousands);
                    double BuyQut = double.Parse(json_Dictionary.data[0][2], NumberStyles.Float | NumberStyles.AllowThousands);
                    double SelNOT = double.Parse(json_Dictionary.data[0][3], NumberStyles.Float | NumberStyles.AllowThousands);
                    double SelQut = double.Parse(json_Dictionary.data[0][4], NumberStyles.Float | NumberStyles.AllowThousands);

                    double BuyNOTEnd = double.Parse(json_Dictionary.data[json_Dictionary.data.Count - 1][1], NumberStyles.Float | NumberStyles.AllowThousands);
                    double BuyQutEnd = double.Parse(json_Dictionary.data[json_Dictionary.data.Count - 1][2], NumberStyles.Float | NumberStyles.AllowThousands);
                    double SelNOTEnd = double.Parse(json_Dictionary.data[json_Dictionary.data.Count - 1][3], NumberStyles.Float | NumberStyles.AllowThousands);
                    double SelQutEnd = double.Parse(json_Dictionary.data[json_Dictionary.data.Count - 1][4], NumberStyles.Float | NumberStyles.AllowThousands);
                    double CumTEnd = double.Parse(json_Dictionary.data[json_Dictionary.data.Count - 1][5], NumberStyles.Float | NumberStyles.AllowThousands);
                    double CumQEnd = double.Parse(json_Dictionary.data[json_Dictionary.data.Count - 1][6], NumberStyles.Float | NumberStyles.AllowThousands);
                    double CumPEnd = double.Parse(json_Dictionary.data[json_Dictionary.data.Count - 1][7], NumberStyles.Float | NumberStyles.AllowThousands);

                    dtBuyNOTBeg.Add(dtin, BuyNOT);
                    dtBuyQutBeg.Add(dtin, BuyQut);
                    dtSelNOTBeg.Add(dtin, SelNOT);
                    dtSelQutBeg.Add(dtin, SelQut);
                    dtBuyNOTEnd.Add(dtin, BuyNOTEnd);
                    dtBuyQutEnd.Add(dtin, BuyQutEnd);
                    dtSelNOTEnd.Add(dtin, SelNOTEnd);
                    dtSelQutEnd.Add(dtin, SelQutEnd);
                    dtCumTEnd.Add(dtin, CumTEnd);
                    dtCumQEnd.Add(dtin, CumQEnd);
                    dtCumPEnd.Add(dtin, CumPEnd);
                }

                {
                    string strfilename = string.Format("D:\\Chips\\MARGNMS\\MI_MARGN_MS_{0}.json", dtin.ToString("yyyyMMdd"));

                    string json = File.ReadAllText(strfilename);
                    MARGN_DAY json_Dictionary = JsonConvert.DeserializeObject<MARGN_DAY>(json);
                    string strFinBuyQut = json_Dictionary.creditList[0][1];
                    string strFinSelQut = json_Dictionary.creditList[0][2];
                    string strFinCashReturn = json_Dictionary.creditList[0][3];
                    string strFinTodayQut = json_Dictionary.creditList[0][4];
                    string strFinYesterQut = json_Dictionary.creditList[0][5];

                    string strMarBuyQut = json_Dictionary.creditList[1][1];
                    string strMarSelQut = json_Dictionary.creditList[1][2];
                    string strMarCashReturn = json_Dictionary.creditList[1][3];
                    string strMarTodayQut = json_Dictionary.creditList[1][4];
                    string strMarYesterQut = json_Dictionary.creditList[1][5];

                    string strFinBuyPri = json_Dictionary.creditList[2][1];
                    string strFinSelPri = json_Dictionary.creditList[2][2];
                    string strFinCashReturnPri = json_Dictionary.creditList[2][3];
                    string strFinTodayPri = json_Dictionary.creditList[2][4];
                    string strFinYesterPri = json_Dictionary.creditList[2][5];

                    dtFinBuyQut.Add(dtin, double.Parse(strFinBuyQut, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign));
                    dtFinSelQut.Add(dtin, double.Parse(strFinSelQut, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign));
                    dtFinCashReturn.Add(dtin, double.Parse(strFinCashReturn, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign));
                    dtFinTodayQut.Add(dtin, double.Parse(strFinTodayQut, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign)); 
                    dtFinYesterQut.Add(dtin, double.Parse(strFinYesterQut, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign)); 
                    
                    dtMarBuyQut.Add(dtin, double.Parse(strMarBuyQut, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign));  
                    dtMarSelQut.Add(dtin, double.Parse(strMarSelQut, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign));  
                    dtMarCashReturn.Add(dtin, double.Parse(strMarCashReturn, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign));  
                    dtMarTodayQut.Add(dtin, double.Parse(strMarTodayQut, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign)); 
                    dtMarYesterQut.Add(dtin, double.Parse(strMarYesterQut, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign));

                    dtFinBuyPri.Add(dtin, double.Parse(strFinBuyPri, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign));
                    dtFinSelPri.Add(dtin, double.Parse(strFinSelPri, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign));
                    dtFinCashReturnPri.Add(dtin, double.Parse(strFinCashReturnPri, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign));
                    dtFinTodayPri.Add(dtin, double.Parse(strFinTodayPri, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign));
                    dtFinYesterPri.Add(dtin, double.Parse(strFinYesterPri, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign));  
                }
            }

            string strwritefilename = "result.csv";
            File.Delete(strwritefilename);

            string stradd = string.Format("Date,HiDiffidIn1Day,LoDiffIn1Day,HiDiffidIn2Days,LoDiffIn2Days,HiDiffidIn3Days,LoDiffIn3Days,Close,High,Low,Close-1,High-1,Low-1,BuyNOT,BuyQut,SelNOT,SelQut,BuyNOTEnd,BuyQutEnd,SelNOTEnd,SelQutEnd,CumTEnd,CumQEnd,CumPEnd,USDtoNTD,RMBtoNTD,EURtoUSD,EURtoJPD,GBPtoUSD,AUDtoUSD,USDtoHKD,USDtoRMB,USDtoSAD,FinBuyQut,FinSelQut,FinCashReturn,FinTodayQut,FinYesterQut,MarBuyQut,MarSelQut,MarCashReturn,MarTodayQut,MarYesterQut,FinBuyPri,FinSelPri,FinCashReturnPri,FinTodayPri,FinYesterPri");
            stradd += "\r\n";
            File.AppendAllText(strwritefilename, stradd);

            int iTotaldays = 0;
            int iHiVIXdays = 0;
            int iBuffdays = 0;
            int iBeardays = 0;

            DateTime dtin_1 = dtClose.Keys.ElementAt(0);
            foreach (DateTime dtin in dtDiffHi.Keys)
            {
                stradd = string.Format("{0},", dtin.ToString("yyyy/MM/dd"));
                iTotaldays++;
                stradd += string.Format("{0},", dtDiffHi1Day[dtin]);
                stradd += string.Format("{0},", dtDiffLo1Day[dtin]);
                stradd += string.Format("{0},", dtDiffHi2Days[dtin]);
                stradd += string.Format("{0},", dtDiffLo2Days[dtin]);
                stradd += string.Format("{0},", dtDiffHi[dtin]);
                stradd += string.Format("{0},", dtDiffLo[dtin]);
                stradd += string.Format("{0},", dtClose[dtin]);
                stradd += string.Format("{0},", dtHi[dtin]);
                stradd += string.Format("{0},", dtLo[dtin]);
                stradd += string.Format("{0},", dtClose[dtin_1]);
                stradd += string.Format("{0},", dtHi[dtin_1]);
                stradd += string.Format("{0},", dtLo[dtin_1]);

                stradd += string.Format("{0},", dtBuyNOTBeg[dtin]);
                stradd += string.Format("{0},", dtBuyQutBeg[dtin]);
                stradd += string.Format("{0},", dtSelNOTBeg[dtin]);
                stradd += string.Format("{0},", dtSelQutBeg[dtin]);
                stradd += string.Format("{0},", dtBuyNOTEnd[dtin]);
                stradd += string.Format("{0},", dtBuyQutEnd[dtin]);
                stradd += string.Format("{0},", dtSelNOTEnd[dtin]);
                stradd += string.Format("{0},", dtSelQutEnd[dtin]);
                stradd += string.Format("{0},", dtCumTEnd[dtin]);
                stradd += string.Format("{0},", dtCumQEnd[dtin]);
                stradd += string.Format("{0},", dtCumPEnd[dtin]);

                stradd += string.Format("{0},", dtEXRateUSDtoNTD[dtin]);
                stradd += string.Format("{0},", dtEXRateRMBtoNTD[dtin]);
                stradd += string.Format("{0},", dtEXRateEURtoUSD[dtin]);
                stradd += string.Format("{0},", dtEXRateEURtoJPD[dtin]);
                stradd += string.Format("{0},", dtEXRateGBPtoUSD[dtin]);
                stradd += string.Format("{0},", dtEXRateAUDtoUSD[dtin]);
                stradd += string.Format("{0},", dtEXRateUSDtoHKD[dtin]);
                stradd += string.Format("{0},", dtEXRateUSDtoRMB[dtin]);
                stradd += string.Format("{0},", dtEXRateUSDtoSAD[dtin]);

                stradd += string.Format("{0},", dtFinBuyQut[dtin]);
                stradd += string.Format("{0},", dtFinSelQut[dtin]);
                stradd += string.Format("{0},", dtFinCashReturn[dtin]);
                stradd += string.Format("{0},", dtFinTodayQut[dtin]);
                stradd += string.Format("{0},", dtFinYesterQut[dtin]);
                stradd += string.Format("{0},", dtMarBuyQut[dtin]);
                stradd += string.Format("{0},", dtMarSelQut[dtin]);
                stradd += string.Format("{0},", dtMarCashReturn[dtin]);
                stradd += string.Format("{0},", dtMarTodayQut[dtin]);
                stradd += string.Format("{0},", dtMarYesterQut[dtin]);
                stradd += string.Format("{0},", dtFinBuyPri[dtin]);
                stradd += string.Format("{0},", dtFinSelPri[dtin]);
                stradd += string.Format("{0},", dtFinCashReturnPri[dtin]);
                stradd += string.Format("{0},", dtFinTodayPri[dtin]);
                stradd += string.Format("{0}", dtFinYesterPri[dtin]);

                stradd += "\r\n";
                File.AppendAllText(strwritefilename, stradd);

                dtin_1 = dtin;
            }
            int nnn=0;
        }

        private void button66_Click(object sender, EventArgs e)
        {
            string strTitle = ",SelfSelfBuy,SelfSelfSell,SelfSelfSum,SelfHedgeBuy,SelfHedgeSell,SelfHedgeSum,TrustBuy,TrustSell,TrustSum,ForBuy,ForSell,ForSum,AllBuy,AllSell,AllSum";
            string strfilename = string.Format("D:\\Chips\\Trade\\ExamChipTrade\\ExamChipTrade\\bin\\Debug\\result.csv");
            string strmake3tpyfilename = string.Format("D:\\Chips\\Trade\\ExamChipTrade\\ExamChipTrade\\bin\\Debug\\TWSEresult3pty.csv");
            IEnumerable<string> lines = File.ReadLines(strfilename);
            int iLine = 0;
            foreach (string line in lines)
            {
                if(iLine == 0)
                {
                    File.AppendAllText(strmake3tpyfilename, line + strTitle + "\r\n");
                }
                string[] strAtLine = line.Split(',');
                try
                {
                    DateTime dtin = DateTime.ParseExact(strAtLine[0], "yyyy/MM/dd", CultureInfo.InvariantCulture);
                    string strfilenameBFI82U = string.Format("D:\\Chips\\BFI82U\\BFI82U_day_{0}.json", dtin.ToString("yyyyMMdd"));
                    if(dtin.Year == 2016 && dtin.Month == 11 && dtin.Day == 11)
                    {
                        int mmm = 0;
                    }
               
                    string json = File.ReadAllText(strfilenameBFI82U);
                    ObjectOf3pty json_Dictionary = JsonConvert.DeserializeObject<ObjectOf3pty>(json);

                    string strSelfSelfBuy="0";
                    string strSelfSelfSell = "0";
                    string strSelfSelfSum = "0";
                    string strSelfHedgBuy = "0";
                    string strSelfHedgSell = "0";
                    string strSelfHedgSum = "0";
                    string strTrustBuy = "0";
                    string strTrustSell = "0";
                    string strTrustSum = "0";
                    string strForBuy = "0";
                    string strForSell = "0";
                    string strForSum = "0";
                    string strAllBuy = "0";
                    string strAllSell = "0";
                    string strAllSum = "0";
                    if(json_Dictionary.data.Count == 4)
                    {
                        strSelfSelfBuy = json_Dictionary.data[0][1];
                        strSelfSelfSell = json_Dictionary.data[0][2];
                        strSelfSelfSum = json_Dictionary.data[0][3];
                        strTrustBuy = json_Dictionary.data[1][1];
                        strTrustSell = json_Dictionary.data[1][2];
                        strTrustSum = json_Dictionary.data[1][3];
                        strForBuy = json_Dictionary.data[2][1];
                        strForSell = json_Dictionary.data[2][2];
                        strForSum = json_Dictionary.data[2][3];
                        strAllBuy = json_Dictionary.data[3][1];
                        strAllSell = json_Dictionary.data[3][2];
                        strAllSum = json_Dictionary.data[3][3];

                    }
                    else
                    {
                        strSelfSelfBuy = json_Dictionary.data[0][1];
                        strSelfSelfSell = json_Dictionary.data[0][2];
                        strSelfSelfSum = json_Dictionary.data[0][3];
                        strSelfHedgBuy = json_Dictionary.data[1][1];
                        strSelfHedgSell = json_Dictionary.data[1][2];
                        strSelfHedgSum = json_Dictionary.data[1][3];
                        strTrustBuy = json_Dictionary.data[2][1];
                        strTrustSell = json_Dictionary.data[2][2];
                        strTrustSum = json_Dictionary.data[2][3];
                        strForBuy = json_Dictionary.data[3][1];
                        strForSell = json_Dictionary.data[3][2];
                        strForSum = json_Dictionary.data[3][3];
                        strAllBuy = json_Dictionary.data[4][1];
                        strAllSell = json_Dictionary.data[4][2];
                        strAllSum = json_Dictionary.data[4][3];
                    }


                    double dbSelfSelfBuy  =  double.Parse(strSelfSelfBuy            , NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);    
                    double dbSelfSelfSell =  double.Parse(strSelfSelfSell            , NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);      
                    double dbSelfSelfSum  =  double.Parse(strSelfSelfSum              , NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);     
                    double dbSelfHedgBuy  =  double.Parse(strSelfHedgBuy              , NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);     
                    double dbSelfHedgSell =  double.Parse(strSelfHedgSell             , NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);        
                    double dbSelfHedgSum  =  double.Parse(strSelfHedgSum              , NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);      
                    double dbTrustBuy     =  double.Parse(strTrustBuy                 , NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                    double dbTrustSell    =  double.Parse(strTrustSell                , NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);     
                    double dbTrustSum     =  double.Parse(strTrustSum                 , NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);      
                    double dbForBuy       =  double.Parse(strForBuy                   , NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);        
                    double dbForSell      =  double.Parse(strForSell                  , NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);           
                    double dbForSum       =  double.Parse(strForSum                   , NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                    double dbAllBuy = double.Parse(strAllBuy, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                    double dbAllSell = double.Parse(strAllSell, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                    double dbAllSum = double.Parse(strAllSum, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);


                    if (dtin >= new DateTime(2017, 12, 18))
                    {
                        dbForBuy += dbAllBuy;
                        dbForSell += dbAllSell;
                        dbForSum += dbAllSum;

                        strAllBuy = json_Dictionary.data[5][1];
                        strAllSell = json_Dictionary.data[5][2];
                        strAllSum = json_Dictionary.data[5][3];

                        dbAllBuy = double.Parse(strAllBuy, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                        dbAllSell = double.Parse(strAllSell, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                        dbAllSum = double.Parse(strAllSum, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                    }
                    string strAddmore = string.Format(",{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14}", dbSelfSelfBuy, dbSelfSelfSell, dbSelfSelfSum, dbSelfHedgBuy, dbSelfHedgSell, dbSelfHedgSum,
                        dbTrustBuy, dbTrustSell, dbTrustSum, dbForBuy, dbForSell, dbForSum,dbAllBuy,dbAllSell,dbAllSum);
                    File.AppendAllText(strmake3tpyfilename, line + strAddmore + "\r\n");
                }
                catch(Exception exp)
                {
                    
                }
                iLine++;

            }
        }

        private void button67_Click(object sender, EventArgs e)
        {
            DateTime dtToday = new DateTime(2020, 9, 2);

            //https://www.twse.com.tw/fund/BFI82U?response=json&dayDate=20200831&type=day
            {
                string strSavePath = "D:\\Chips\\BFI82U\\";

                string strLocalFile = string.Format("{0}BFI82U_day_{1}.json", strSavePath, dtToday.ToString("yyyyMMdd"));
                if (!File.Exists(strLocalFile))
                {
                    string url = string.Format("https://www.twse.com.tw/fund/BFI82U?response=json&dayDate={0}&type=day", dtToday.ToString("yyyyMMdd"));

                    HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
                    httpWebRequest.ContentType = "application/json";
                    httpWebRequest.Method = WebRequestMethods.Http.Get;
                    httpWebRequest.Accept = "application/json";


                    string text;
                    var response = (HttpWebResponse)httpWebRequest.GetResponse();

                    using (var sr = new StreamReader(response.GetResponseStream()))
                    {
                        text = sr.ReadToEnd();
                    }
                    File.WriteAllText(strLocalFile, text);
                    response.Close();
                    httpWebRequest.Abort();
                }
            }


            // 3大
            //https://www.twse.com.tw/fund/BFI82U?response=json&dayDate=20200506&type=day
            {
                string strSavePath = "D:\\Chips\\BFI82U\\";
                DateTime dt = DateTime.Now;

                string strLocalFile = string.Format("{0}BFI82U_day_{1}.json", strSavePath, dtToday.ToString("yyyyMMdd"));
                if (!File.Exists(strLocalFile))
                {
                    string url = string.Format("https://www.twse.com.tw/fund/BFI82U?response=json&dayDate={0}&type=day", dtToday.ToString("yyyyMMdd"));

                    HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
                    httpWebRequest.ContentType = "application/json";
                    httpWebRequest.Method = WebRequestMethods.Http.Get;
                    httpWebRequest.Accept = "application/json";


                    string text;
                    var response = (HttpWebResponse)httpWebRequest.GetResponse();

                    using (var sr = new StreamReader(response.GetResponseStream()))
                    {
                        text = sr.ReadToEnd();
                    }

                    File.WriteAllText(strLocalFile, text);

                    response.Close();
                    httpWebRequest.Abort();

                }

            }
            
            {
                //信用交易統計
                //https://www.twse.com.tw/exchangeReport/MI_MARGN?response=json&date=20200803&selectType=MS
                string strSavePath = "D:\\Chips\\MARGNMS\\";
                DateTime dt = DateTime.Now;

                string strLocalFile = string.Format("{0}MI_MARGN_MS_{1}.json", strSavePath, dtToday.ToString("yyyyMMdd"));
                if (!File.Exists(strLocalFile))
                {
                    string url = string.Format("https://www.twse.com.tw/exchangeReport/MI_MARGN?response=json&date={0}&selectType=MS", dtToday.ToString("yyyyMMdd"));

                    HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
                    httpWebRequest.ContentType = "application/json";
                    httpWebRequest.Method = WebRequestMethods.Http.Get;
                    httpWebRequest.Accept = "application/json";


                    string text;
                    var response = (HttpWebResponse)httpWebRequest.GetResponse();

                    using (var sr = new StreamReader(response.GetResponseStream()))
                    {
                        text = sr.ReadToEnd();
                    }

                    File.WriteAllText(strLocalFile, text);

                    response.Close();
                    httpWebRequest.Abort();

                }
            }
            
            // 每日5秒紀錄
            //https://www.twse.com.tw/exchangeReport/MI_5MINS?response=csv&date=20200731
            {
                string strSavePath = "D:\\Chips\\TWSE5MINS\\";
                string strLocalFile = string.Format("{0}MI_5MINS_{1}.json", strSavePath, dtToday.ToString("yyyyMMdd"));
                if (!File.Exists(strLocalFile))
                {
                    string url = string.Format("https://www.twse.com.tw/exchangeReport/MI_5MINS?response=json&date={0}", dtToday.ToString("yyyyMMdd"));

                    HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
                    httpWebRequest.ContentType = "application/json";
                    httpWebRequest.Method = WebRequestMethods.Http.Get;
                    httpWebRequest.Accept = "application/json";


                    string text;
                    var response = (HttpWebResponse)httpWebRequest.GetResponse();

                    using (var sr = new StreamReader(response.GetResponseStream()))
                    {
                        text = sr.ReadToEnd();
                    }

                    File.WriteAllText(strLocalFile, text);

                    response.Close();
                    httpWebRequest.Abort();

                }


            }
            // 這是抓權指__*
            //https://www.twse.com.tw/indicesReport/MI_5MINS_HIST?response=csv&date=20160501
  
            {
                string strSavePath = "D:\\Chips\\TWSE\\";


                string strLocalFile = string.Format("{0}MI_5MINS_HIST_{1}01.json", strSavePath, dtToday.ToString("yyyyMM"));
                if (!File.Exists(strLocalFile))
                {
                    string url = string.Format("https://www.twse.com.tw/indicesReport/MI_5MINS_HIST?response=json&date={0}", dtToday.ToString("yyyyMMdd"));

                    HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
                    httpWebRequest.ContentType = "application/json";
                    httpWebRequest.Method = WebRequestMethods.Http.Get;
                    httpWebRequest.Accept = "application/json";


                    string text;
                    var response = (HttpWebResponse)httpWebRequest.GetResponse();

                    using (var sr = new StreamReader(response.GetResponseStream()))
                    {
                        text = sr.ReadToEnd();
                    }

                    File.WriteAllText(strLocalFile, text);

                    response.Close();
                    httpWebRequest.Abort();


                }
        
            }


            //  https://www.taifex.com.tw/cht/3/dailyFXRate    手動下載


            Dictionary<DateTime, double> dtHi = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtLo = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtClose = new Dictionary<DateTime, double>();

            Dictionary<DateTime, double> dtBuyNOTBeg = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtBuyQutBeg = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtSelNOTBeg = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtSelQutBeg = new Dictionary<DateTime, double>();

            Dictionary<DateTime, double> dtBuyNOTEnd = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtBuyQutEnd = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtSelNOTEnd = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtSelQutEnd = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtCumTEnd = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtCumQEnd = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtCumPEnd = new Dictionary<DateTime, double>();


            Dictionary<DateTime, double> dtEXRateUSDtoNTD = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtEXRateRMBtoNTD = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtEXRateEURtoUSD = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtEXRateEURtoJPD = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtEXRateGBPtoUSD = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtEXRateAUDtoUSD = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtEXRateUSDtoHKD = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtEXRateUSDtoRMB = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtEXRateUSDtoSAD = new Dictionary<DateTime, double>();

            Dictionary<DateTime, double> dtFinBuyQut = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtFinSelQut = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtFinCashReturn = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtFinTodayQut = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtFinYesterQut = new Dictionary<DateTime, double>();

            Dictionary<DateTime, double> dtMarBuyQut = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtMarSelQut = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtMarCashReturn = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtMarTodayQut = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtMarYesterQut = new Dictionary<DateTime, double>();

            Dictionary<DateTime, double> dtFinBuyPri = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtFinSelPri = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtFinCashReturnPri = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtFinTodayPri = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtFinYesterPri = new Dictionary<DateTime, double>();


            double dbSelfSelfBuy = 0;
            double dbSelfSelfSell = 0;
            double dbSelfSelfSum = 0;
            double dbSelfHedgBuy = 0;
            double dbSelfHedgSell = 0;
            double dbSelfHedgSum = 0;
            double dbTrustBuy = 0;
            double dbTrustSell = 0;
            double dbTrustSum = 0;
            double dbForBuy = 0;
            double dbForSell = 0;
            double dbForSum = 0;
            double dbAllBuy = 0;
            double dbAllSell = 0;
            double dbAllSum = 0;


            {
                string strbase = "D:\\Chips\\TWSE\\";
                DateTime dtlast = dtToday.AddMonths(-1);
                string strfilename = string.Format("{0}MI_5MINS_HIST_{1}01.json", strbase, dtToday.ToString("yyyyMM"));
                string strfilenamelastmonth = string.Format("{0}MI_5MINS_HIST_{1}01.json", strbase, dtlast.ToString("yyyyMM"));

                {
                    string json = File.ReadAllText(strfilename);
                    JSON_Root json_Dictionary = JsonConvert.DeserializeObject<JSON_Root>(json);
                    List<List<string>> lstlist = json_Dictionary.data;
                    foreach (List<string> contextinlst in lstlist)
                    {
                        string[] YYYMMDD = contextinlst[0].Split('/');
                        int iyyy = int.Parse(YYYMMDD[0]);
                        int imm = int.Parse(YYYMMDD[1]);
                        int idd = int.Parse(YYYMMDD[2]);
                        //DateTime dt = DateTime.ParseExact(contextinlst[0], "yyy/MM/dd", CultureInfo.InvariantCulture).AddYears(1911);
                        DateTime dt = new DateTime(1911 + iyyy, imm, idd);
                        double dbOpen = double.Parse(contextinlst[1], NumberStyles.Float | NumberStyles.AllowThousands);
                        double dbHi = double.Parse(contextinlst[2], NumberStyles.Float | NumberStyles.AllowThousands);
                        double dbLo = double.Parse(contextinlst[3], NumberStyles.Float | NumberStyles.AllowThousands);
                        double dbClose = double.Parse(contextinlst[4], NumberStyles.Float | NumberStyles.AllowThousands);
                        dtHi.Add(dt, dbHi);
                        dtLo.Add(dt, dbLo);
                        dtClose.Add(dt, dbClose);

                    }
                }

                {
                    string json = File.ReadAllText(strfilenamelastmonth);
                    JSON_Root json_Dictionary = JsonConvert.DeserializeObject<JSON_Root>(json);
                    List<List<string>> lstlist = json_Dictionary.data;
                    foreach (List<string> contextinlst in lstlist)
                    {
                        string[] YYYMMDD = contextinlst[0].Split('/');
                        int iyyy = int.Parse(YYYMMDD[0]);
                        int imm = int.Parse(YYYMMDD[1]);
                        int idd = int.Parse(YYYMMDD[2]);
                        //DateTime dt = DateTime.ParseExact(contextinlst[0], "yyy/MM/dd", CultureInfo.InvariantCulture).AddYears(1911);
                        DateTime dt = new DateTime(1911 + iyyy, imm, idd);
                        double dbOpen = double.Parse(contextinlst[1], NumberStyles.Float | NumberStyles.AllowThousands);
                        double dbHi = double.Parse(contextinlst[2], NumberStyles.Float | NumberStyles.AllowThousands);
                        double dbLo = double.Parse(contextinlst[3], NumberStyles.Float | NumberStyles.AllowThousands);
                        double dbClose = double.Parse(contextinlst[4], NumberStyles.Float | NumberStyles.AllowThousands);
                        dtHi.Add(dt, dbHi);
                        dtLo.Add(dt, dbLo);
                        dtClose.Add(dt, dbClose);

                    }
                }

            }
            {
                string strfilename = string.Format("D:\\Chips\\TWSE5MINS\\MI_5MINS_{0}.json", dtToday.ToString("yyyyMMdd"));
  
                string json = File.ReadAllText(strfilename);
                JSON_Root_5MINS json_Dictionary = JsonConvert.DeserializeObject<JSON_Root_5MINS>(json);
                double BuyNOT = double.Parse(json_Dictionary.data[0][1], NumberStyles.Float | NumberStyles.AllowThousands);
                double BuyQut = double.Parse(json_Dictionary.data[0][2], NumberStyles.Float | NumberStyles.AllowThousands);
                double SelNOT = double.Parse(json_Dictionary.data[0][3], NumberStyles.Float | NumberStyles.AllowThousands);
                double SelQut = double.Parse(json_Dictionary.data[0][4], NumberStyles.Float | NumberStyles.AllowThousands);

                double BuyNOTEnd = double.Parse(json_Dictionary.data[json_Dictionary.data.Count - 1][1], NumberStyles.Float | NumberStyles.AllowThousands);
                double BuyQutEnd = double.Parse(json_Dictionary.data[json_Dictionary.data.Count - 1][2], NumberStyles.Float | NumberStyles.AllowThousands);
                double SelNOTEnd = double.Parse(json_Dictionary.data[json_Dictionary.data.Count - 1][3], NumberStyles.Float | NumberStyles.AllowThousands);
                double SelQutEnd = double.Parse(json_Dictionary.data[json_Dictionary.data.Count - 1][4], NumberStyles.Float | NumberStyles.AllowThousands);
                double CumTEnd = double.Parse(json_Dictionary.data[json_Dictionary.data.Count - 1][5], NumberStyles.Float | NumberStyles.AllowThousands);
                double CumQEnd = double.Parse(json_Dictionary.data[json_Dictionary.data.Count - 1][6], NumberStyles.Float | NumberStyles.AllowThousands);
                double CumPEnd = double.Parse(json_Dictionary.data[json_Dictionary.data.Count - 1][7], NumberStyles.Float | NumberStyles.AllowThousands);

                dtBuyNOTBeg.Add(dtToday, BuyNOT);
                dtBuyQutBeg.Add(dtToday, BuyQut);
                dtSelNOTBeg.Add(dtToday, SelNOT);
                dtSelQutBeg.Add(dtToday, SelQut);
                dtBuyNOTEnd.Add(dtToday, BuyNOTEnd);
                dtBuyQutEnd.Add(dtToday, BuyQutEnd);
                dtSelNOTEnd.Add(dtToday, SelNOTEnd);
                dtSelQutEnd.Add(dtToday, SelQutEnd);
                dtCumTEnd.Add(dtToday, CumTEnd);
                dtCumQEnd.Add(dtToday, CumQEnd);
                dtCumPEnd.Add(dtToday, CumPEnd);
            }


            {
                string strfilename = string.Format("D:\\Chips\\EXRate\\{0}.csv", dtToday.ToString("yyyy"));

                IEnumerable<string> lines = File.ReadLines(strfilename);
                foreach (string line in lines)
                {
                    string[] strAtLine = line.Split(',');
                    if (strAtLine.Length == 11)
                    {
                        string strdate = strAtLine[0];
                        DateTime parsed;
                        if (DateTime.TryParse(strdate, out parsed))
                        {
                            double EXRateUSDtoNTD = double.Parse(strAtLine[1], NumberStyles.Float);
                            double EXRateRMBtoNTD = 5;
                            if (strAtLine[2] != "-")
                                EXRateRMBtoNTD = double.Parse(strAtLine[2], NumberStyles.Float);
                            double EXRateEURtoUSD = double.Parse(strAtLine[3], NumberStyles.Float);
                            double EXRateEURtoJPD = double.Parse(strAtLine[4], NumberStyles.Float);
                            double EXRateGBPtoUSD = double.Parse(strAtLine[5], NumberStyles.Float);
                            double EXRateAUDtoUSD = double.Parse(strAtLine[6], NumberStyles.Float);
                            double EXRateUSDtoHKD = double.Parse(strAtLine[7], NumberStyles.Float);
                            double EXRateUSDtoRMB = 6.22805;
                            if (strAtLine[8] != "-")
                                EXRateUSDtoRMB = double.Parse(strAtLine[8], NumberStyles.Float);
                            double EXRateUSDtoSAD = 11.32215;
                            if (strAtLine[9] != "-")
                                EXRateUSDtoSAD = double.Parse(strAtLine[9], NumberStyles.Float);

                            dtEXRateUSDtoNTD.Add(parsed, EXRateUSDtoNTD);
                            dtEXRateRMBtoNTD.Add(parsed, EXRateRMBtoNTD);
                            dtEXRateEURtoUSD.Add(parsed, EXRateEURtoUSD);
                            dtEXRateEURtoJPD.Add(parsed, EXRateEURtoJPD);
                            dtEXRateGBPtoUSD.Add(parsed, EXRateGBPtoUSD);
                            dtEXRateAUDtoUSD.Add(parsed, EXRateAUDtoUSD);
                            dtEXRateUSDtoHKD.Add(parsed, EXRateUSDtoHKD);
                            dtEXRateUSDtoRMB.Add(parsed, EXRateUSDtoRMB);
                            dtEXRateUSDtoSAD.Add(parsed, EXRateUSDtoSAD);
                            // Use reformatted
                        }
                        else
                        {
                            int nn = 0;
                            // Log error, perhaps?
                        }
                    }
                }
            }
            {
                string strfilename = string.Format("D:\\Chips\\MARGNMS\\MI_MARGN_MS_{0}.json", dtToday.ToString("yyyyMMdd"));

                string json = File.ReadAllText(strfilename);
                MARGN_DAY json_Dictionary = JsonConvert.DeserializeObject<MARGN_DAY>(json);
                string strFinBuyQut = json_Dictionary.creditList[0][1];
                string strFinSelQut = json_Dictionary.creditList[0][2];
                string strFinCashReturn = json_Dictionary.creditList[0][3];
                string strFinTodayQut = json_Dictionary.creditList[0][4];
                string strFinYesterQut = json_Dictionary.creditList[0][5];

                string strMarBuyQut = json_Dictionary.creditList[1][1];
                string strMarSelQut = json_Dictionary.creditList[1][2];
                string strMarCashReturn = json_Dictionary.creditList[1][3];
                string strMarTodayQut = json_Dictionary.creditList[1][4];
                string strMarYesterQut = json_Dictionary.creditList[1][5];

                string strFinBuyPri = json_Dictionary.creditList[2][1];
                string strFinSelPri = json_Dictionary.creditList[2][2];
                string strFinCashReturnPri = json_Dictionary.creditList[2][3];
                string strFinTodayPri = json_Dictionary.creditList[2][4];
                string strFinYesterPri = json_Dictionary.creditList[2][5];

                dtFinBuyQut.Add(dtToday, double.Parse(strFinBuyQut, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign));
                dtFinSelQut.Add(dtToday, double.Parse(strFinSelQut, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign));
                dtFinCashReturn.Add(dtToday, double.Parse(strFinCashReturn, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign));
                dtFinTodayQut.Add(dtToday, double.Parse(strFinTodayQut, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign));
                dtFinYesterQut.Add(dtToday, double.Parse(strFinYesterQut, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign));

                dtMarBuyQut.Add(dtToday, double.Parse(strMarBuyQut, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign));
                dtMarSelQut.Add(dtToday, double.Parse(strMarSelQut, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign));
                dtMarCashReturn.Add(dtToday, double.Parse(strMarCashReturn, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign));
                dtMarTodayQut.Add(dtToday, double.Parse(strMarTodayQut, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign));
                dtMarYesterQut.Add(dtToday, double.Parse(strMarYesterQut, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign));

                dtFinBuyPri.Add(dtToday, double.Parse(strFinBuyPri, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign));
                dtFinSelPri.Add(dtToday, double.Parse(strFinSelPri, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign));
                dtFinCashReturnPri.Add(dtToday, double.Parse(strFinCashReturnPri, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign));
                dtFinTodayPri.Add(dtToday, double.Parse(strFinTodayPri, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign));
                dtFinYesterPri.Add(dtToday, double.Parse(strFinYesterPri, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign));
            }
            {
                try
                {
                    string strfilenameBFI82U = string.Format("D:\\Chips\\BFI82U\\BFI82U_day_{0}.json", dtToday.ToString("yyyyMMdd"));


                    string json = File.ReadAllText(strfilenameBFI82U);
                    ObjectOf3pty json_Dictionary = JsonConvert.DeserializeObject<ObjectOf3pty>(json);

                    string strSelfSelfBuy = "0";
                    string strSelfSelfSell = "0";
                    string strSelfSelfSum = "0";
                    string strSelfHedgBuy = "0";
                    string strSelfHedgSell = "0";
                    string strSelfHedgSum = "0";
                    string strTrustBuy = "0";
                    string strTrustSell = "0";
                    string strTrustSum = "0";
                    string strForBuy = "0";
                    string strForSell = "0";
                    string strForSum = "0";
                    string strAllBuy = "0";
                    string strAllSell = "0";
                    string strAllSum = "0";
                    if (json_Dictionary.data.Count == 4)
                    {
                        strSelfSelfBuy = json_Dictionary.data[0][1];
                        strSelfSelfSell = json_Dictionary.data[0][2];
                        strSelfSelfSum = json_Dictionary.data[0][3];
                        strTrustBuy = json_Dictionary.data[1][1];
                        strTrustSell = json_Dictionary.data[1][2];
                        strTrustSum = json_Dictionary.data[1][3];
                        strForBuy = json_Dictionary.data[2][1];
                        strForSell = json_Dictionary.data[2][2];
                        strForSum = json_Dictionary.data[2][3];
                        strAllBuy = json_Dictionary.data[3][1];
                        strAllSell = json_Dictionary.data[3][2];
                        strAllSum = json_Dictionary.data[3][3];

                    }
                    else
                    {
                        strSelfSelfBuy = json_Dictionary.data[0][1];
                        strSelfSelfSell = json_Dictionary.data[0][2];
                        strSelfSelfSum = json_Dictionary.data[0][3];
                        strSelfHedgBuy = json_Dictionary.data[1][1];
                        strSelfHedgSell = json_Dictionary.data[1][2];
                        strSelfHedgSum = json_Dictionary.data[1][3];
                        strTrustBuy = json_Dictionary.data[2][1];
                        strTrustSell = json_Dictionary.data[2][2];
                        strTrustSum = json_Dictionary.data[2][3];
                        strForBuy = json_Dictionary.data[3][1];
                        strForSell = json_Dictionary.data[3][2];
                        strForSum = json_Dictionary.data[3][3];
                        strAllBuy = json_Dictionary.data[4][1];
                        strAllSell = json_Dictionary.data[4][2];
                        strAllSum = json_Dictionary.data[4][3];
                    }


                    dbSelfSelfBuy = double.Parse(strSelfSelfBuy, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                    dbSelfSelfSell = double.Parse(strSelfSelfSell, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                    dbSelfSelfSum = double.Parse(strSelfSelfSum, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                    dbSelfHedgBuy = double.Parse(strSelfHedgBuy, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                    dbSelfHedgSell = double.Parse(strSelfHedgSell, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                    dbSelfHedgSum = double.Parse(strSelfHedgSum, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                    dbTrustBuy = double.Parse(strTrustBuy, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                    dbTrustSell = double.Parse(strTrustSell, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                    dbTrustSum = double.Parse(strTrustSum, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                    dbForBuy = double.Parse(strForBuy, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                    dbForSell = double.Parse(strForSell, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                    dbForSum = double.Parse(strForSum, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                    dbAllBuy = double.Parse(strAllBuy, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                    dbAllSell = double.Parse(strAllSell, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                    dbAllSum = double.Parse(strAllSum, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);



                    dbForBuy += dbAllBuy;
                    dbForSell += dbAllSell;
                    dbForSum += dbAllSum;

                    strAllBuy = json_Dictionary.data[5][1];
                    strAllSell = json_Dictionary.data[5][2];
                    strAllSum = json_Dictionary.data[5][3];

                    dbAllBuy = double.Parse(strAllBuy, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                    dbAllSell = double.Parse(strAllSell, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                    dbAllSum = double.Parse(strAllSum, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                  
                    string strAddmore = string.Format(",{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14}", dbSelfSelfBuy, dbSelfSelfSell, dbSelfSelfSum, dbSelfHedgBuy, dbSelfHedgSell, dbSelfHedgSum,
                        dbTrustBuy, dbTrustSell, dbTrustSum, dbForBuy, dbForSell, dbForSum, dbAllBuy, dbAllSell, dbAllSum);
                   // File.AppendAllText(strmake3tpyfilename, line + strAddmore + "\r\n");
                }
                catch (Exception exp)
                {

                }
            }

            DateTime dtLastTime = dtToday.AddDays(-1);
            while(!dtClose.ContainsKey(dtLastTime))
            {
                dtLastTime = dtLastTime.AddDays(-1);
            }
            string strwritefilename = "today.csv";
            File.Delete(strwritefilename);

            string stradd = string.Format("Date,HiDiffidIn1Day,LoDiffIn1Day,HiDiffidIn2Days,LoDiffIn2Days,HiDiffidIn3Days,LoDiffIn3Days,Close,High,Low,Close-1,High-1,Low-1,BuyNOT,BuyQut,SelNOT,SelQut,BuyNOTEnd,BuyQutEnd,SelNOTEnd,SelQutEnd,CumTEnd,CumQEnd,CumPEnd,USDtoNTD,RMBtoNTD,EURtoUSD,EURtoJPD,GBPtoUSD,AUDtoUSD,USDtoHKD,USDtoRMB,USDtoSAD,FinBuyQut,FinSelQut,FinCashReturn,FinTodayQut,FinYesterQut,MarBuyQut,MarSelQut,MarCashReturn,MarTodayQut,MarYesterQut,FinBuyPri,FinSelPri,FinCashReturnPri,FinTodayPri,FinYesterPri,SelfSelfBuy,SelfSelfSell,SelfSelfSum,SelfHedgeBuy,SelfHedgeSell,SelfHedgeSum,TrustBuy,TrustSell,TrustSum,ForBuy,ForSell,ForSum,AllBuy,AllSell,AllSum");
            stradd += "\r\n";
            File.AppendAllText(strwritefilename, stradd, Encoding.UTF8);

            stradd = "";
            string straddvalue = string.Format("{0},NA,NA,NA,NA,NA,NA,{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},", dtToday.ToString("yyyy/MM/dd")
                                                                       , dtClose[dtToday]
                                                                       , dtHi[dtToday]
                                                                       , dtLo[dtToday]
                                                                       , dtClose[dtLastTime]
                                                                       , dtHi[dtLastTime]
                                                                       , dtLo[dtLastTime]
                                                                       , dtBuyNOTBeg[dtToday]
                                                                       , dtBuyQutBeg[dtToday]
                                                                       , dtSelNOTBeg[dtToday]
                                                                       , dtSelQutBeg[dtToday]
                                                                       , dtBuyNOTEnd[dtToday]
                                                                       , dtBuyQutEnd[dtToday]
                                                                       , dtSelNOTEnd[dtToday]
                                                                       , dtSelQutEnd[dtToday]
                                                                       , dtCumTEnd[dtToday]
                                                                       , dtCumQEnd[dtToday]
                                                                       , dtCumPEnd[dtToday]);

            // https://www.taifex.com.tw/cht/3/dailyFXRate
            stradd += string.Format("{0},", dtEXRateUSDtoNTD[dtToday]);
            stradd += string.Format("{0},", dtEXRateRMBtoNTD[dtToday]);
            stradd += string.Format("{0},", dtEXRateEURtoUSD[dtToday]);
            stradd += string.Format("{0},", dtEXRateEURtoJPD[dtToday]);
            stradd += string.Format("{0},", dtEXRateGBPtoUSD[dtToday]);
            stradd += string.Format("{0},", dtEXRateAUDtoUSD[dtToday]);
            stradd += string.Format("{0},", dtEXRateUSDtoHKD[dtToday]);
            stradd += string.Format("{0},", dtEXRateUSDtoRMB[dtToday]);
            stradd += string.Format("{0},", dtEXRateUSDtoSAD[dtToday]);


            stradd += string.Format("{0},", dtFinBuyQut[dtToday]);
            stradd += string.Format("{0},", dtFinSelQut[dtToday]);
            stradd += string.Format("{0},", dtFinCashReturn[dtToday]);
            stradd += string.Format("{0},", dtFinTodayQut[dtToday]);
            stradd += string.Format("{0},", dtFinYesterQut[dtToday]);
            stradd += string.Format("{0},", dtMarBuyQut[dtToday]);
            stradd += string.Format("{0},", dtMarSelQut[dtToday]);
            stradd += string.Format("{0},", dtMarCashReturn[dtToday]);
            stradd += string.Format("{0},", dtMarTodayQut[dtToday]);
            stradd += string.Format("{0},", dtMarYesterQut[dtToday]);
            stradd += string.Format("{0},", dtFinBuyPri[dtToday]);
            stradd += string.Format("{0},", dtFinSelPri[dtToday]);
            stradd += string.Format("{0},", dtFinCashReturnPri[dtToday]);
            stradd += string.Format("{0},", dtFinTodayPri[dtToday]);
            stradd += string.Format("{0},", dtFinYesterPri[dtToday]);
            straddvalue += stradd;


            string strAdd3pty = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14}", dbSelfSelfBuy, dbSelfSelfSell, dbSelfSelfSum, dbSelfHedgBuy, dbSelfHedgSell, dbSelfHedgSum,
    dbTrustBuy, dbTrustSell, dbTrustSum, dbForBuy, dbForSell, dbForSum, dbAllBuy, dbAllSell, dbAllSum);
            straddvalue += strAdd3pty;

            straddvalue += "\r\n";
            File.AppendAllText(strwritefilename, straddvalue, Encoding.UTF8);




        }

        private void button68_Click(object sender, EventArgs e)
        {
            DateTime dtToday = new DateTime(2020, 9, 1);

            //https://www.twse.com.tw/fund/BFI82U?response=json&dayDate=20200831&type=day
            {
                string strSavePath = "D:\\Chips\\BFI82U\\";

                string strLocalFile = string.Format("{0}BFI82U_day_{1}.json", strSavePath, dtToday.ToString("yyyyMMdd"));
                if (!File.Exists(strLocalFile))
                {
                    string url = string.Format("https://www.twse.com.tw/fund/BFI82U?response=json&dayDate={0}&type=day", dtToday.ToString("yyyyMMdd"));

                    HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
                    httpWebRequest.ContentType = "application/json";
                    httpWebRequest.Method = WebRequestMethods.Http.Get;
                    httpWebRequest.Accept = "application/json";


                    string text;
                    var response = (HttpWebResponse)httpWebRequest.GetResponse();

                    using (var sr = new StreamReader(response.GetResponseStream()))
                    {
                        text = sr.ReadToEnd();
                    }
                    File.WriteAllText(strLocalFile, text);
                    response.Close();
                    httpWebRequest.Abort();
                }
            }


            // 3大
            //https://www.twse.com.tw/fund/BFI82U?response=json&dayDate=20200506&type=day
            {
                string strSavePath = "D:\\Chips\\BFI82U\\";
                DateTime dt = DateTime.Now;

                string strLocalFile = string.Format("{0}BFI82U_day_{1}.json", strSavePath, dtToday.ToString("yyyyMMdd"));
                if (!File.Exists(strLocalFile))
                {
                    string url = string.Format("https://www.twse.com.tw/fund/BFI82U?response=json&dayDate={0}&type=day", dtToday.ToString("yyyyMMdd"));

                    HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
                    httpWebRequest.ContentType = "application/json";
                    httpWebRequest.Method = WebRequestMethods.Http.Get;
                    httpWebRequest.Accept = "application/json";


                    string text;
                    var response = (HttpWebResponse)httpWebRequest.GetResponse();

                    using (var sr = new StreamReader(response.GetResponseStream()))
                    {
                        text = sr.ReadToEnd();
                    }

                    File.WriteAllText(strLocalFile, text);

                    response.Close();
                    httpWebRequest.Abort();

                }

            }

            /*{
                //信用交易統計
                //https://www.twse.com.tw/exchangeReport/MI_MARGN?response=json&date=20200803&selectType=MS
                string strSavePath = "D:\\Chips\\MARGNMS\\";
                DateTime dt = DateTime.Now;

                string strLocalFile = string.Format("{0}MI_MARGN_MS_{1}.json", strSavePath, dtToday.ToString("yyyyMMdd"));
                if (!File.Exists(strLocalFile))
                {
                    string url = string.Format("https://www.twse.com.tw/exchangeReport/MI_MARGN?response=json&date={0}&selectType=MS", dtToday.ToString("yyyyMMdd"));

                    HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
                    httpWebRequest.ContentType = "application/json";
                    httpWebRequest.Method = WebRequestMethods.Http.Get;
                    httpWebRequest.Accept = "application/json";


                    string text;
                    var response = (HttpWebResponse)httpWebRequest.GetResponse();

                    using (var sr = new StreamReader(response.GetResponseStream()))
                    {
                        text = sr.ReadToEnd();
                    }

                    File.WriteAllText(strLocalFile, text);

                    response.Close();
                    httpWebRequest.Abort();

                }
            }*/

            // 每日5秒紀錄
            //https://www.twse.com.tw/exchangeReport/MI_5MINS?response=csv&date=20200731
            {
                string strSavePath = "D:\\Chips\\TWSE5MINS\\";
                string strLocalFile = string.Format("{0}MI_5MINS_{1}.json", strSavePath, dtToday.ToString("yyyyMMdd"));
                if (!File.Exists(strLocalFile))
                {
                    string url = string.Format("https://www.twse.com.tw/exchangeReport/MI_5MINS?response=json&date={0}", dtToday.ToString("yyyyMMdd"));

                    HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
                    httpWebRequest.ContentType = "application/json";
                    httpWebRequest.Method = WebRequestMethods.Http.Get;
                    httpWebRequest.Accept = "application/json";


                    string text;
                    var response = (HttpWebResponse)httpWebRequest.GetResponse();

                    using (var sr = new StreamReader(response.GetResponseStream()))
                    {
                        text = sr.ReadToEnd();
                    }

                    File.WriteAllText(strLocalFile, text);

                    response.Close();
                    httpWebRequest.Abort();

                }


            }
            // 這是抓權指__*
            //https://www.twse.com.tw/indicesReport/MI_5MINS_HIST?response=csv&date=20160501

            {
                string strSavePath = "D:\\Chips\\TWSE\\";


                string strLocalFile = string.Format("{0}MI_5MINS_HIST_{1}01.json", strSavePath, dtToday.ToString("yyyyMM"));
                if (!File.Exists(strLocalFile))
                {
                    string url = string.Format("https://www.twse.com.tw/indicesReport/MI_5MINS_HIST?response=json&date={0}", dtToday.ToString("yyyyMMdd"));

                    HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
                    httpWebRequest.ContentType = "application/json";
                    httpWebRequest.Method = WebRequestMethods.Http.Get;
                    httpWebRequest.Accept = "application/json";


                    string text;
                    var response = (HttpWebResponse)httpWebRequest.GetResponse();

                    using (var sr = new StreamReader(response.GetResponseStream()))
                    {
                        text = sr.ReadToEnd();
                    }

                    File.WriteAllText(strLocalFile, text);

                    response.Close();
                    httpWebRequest.Abort();


                }

            }


            //  https://www.taifex.com.tw/cht/3/dailyFXRate    手動下載


            Dictionary<DateTime, double> dtHi = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtLo = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtClose = new Dictionary<DateTime, double>();

            Dictionary<DateTime, double> dtBuyNOTBeg = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtBuyQutBeg = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtSelNOTBeg = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtSelQutBeg = new Dictionary<DateTime, double>();

            Dictionary<DateTime, double> dtBuyNOTEnd = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtBuyQutEnd = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtSelNOTEnd = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtSelQutEnd = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtCumTEnd = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtCumQEnd = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtCumPEnd = new Dictionary<DateTime, double>();


            Dictionary<DateTime, double> dtEXRateUSDtoNTD = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtEXRateRMBtoNTD = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtEXRateEURtoUSD = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtEXRateEURtoJPD = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtEXRateGBPtoUSD = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtEXRateAUDtoUSD = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtEXRateUSDtoHKD = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtEXRateUSDtoRMB = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtEXRateUSDtoSAD = new Dictionary<DateTime, double>();

            Dictionary<DateTime, double> dtFinBuyQut = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtFinSelQut = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtFinCashReturn = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtFinTodayQut = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtFinYesterQut = new Dictionary<DateTime, double>();

            Dictionary<DateTime, double> dtMarBuyQut = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtMarSelQut = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtMarCashReturn = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtMarTodayQut = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtMarYesterQut = new Dictionary<DateTime, double>();

            Dictionary<DateTime, double> dtFinBuyPri = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtFinSelPri = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtFinCashReturnPri = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtFinTodayPri = new Dictionary<DateTime, double>();
            Dictionary<DateTime, double> dtFinYesterPri = new Dictionary<DateTime, double>();


            double dbSelfSelfBuy = 0;
            double dbSelfSelfSell = 0;
            double dbSelfSelfSum = 0;
            double dbSelfHedgBuy = 0;
            double dbSelfHedgSell = 0;
            double dbSelfHedgSum = 0;
            double dbTrustBuy = 0;
            double dbTrustSell = 0;
            double dbTrustSum = 0;
            double dbForBuy = 0;
            double dbForSell = 0;
            double dbForSum = 0;
            double dbAllBuy = 0;
            double dbAllSell = 0;
            double dbAllSum = 0;


            {
                string strbase = "D:\\Chips\\TWSE\\";
                DateTime dtlast = dtToday.AddMonths(-1);
                string strfilename = string.Format("{0}MI_5MINS_HIST_{1}01.json", strbase, dtToday.ToString("yyyyMM"));
                string strfilenamelastmonth = string.Format("{0}MI_5MINS_HIST_{1}01.json", strbase, dtlast.ToString("yyyyMM"));

                {
                    string json = File.ReadAllText(strfilename);
                    JSON_Root json_Dictionary = JsonConvert.DeserializeObject<JSON_Root>(json);
                    List<List<string>> lstlist = json_Dictionary.data;
                    foreach (List<string> contextinlst in lstlist)
                    {
                        string[] YYYMMDD = contextinlst[0].Split('/');
                        int iyyy = int.Parse(YYYMMDD[0]);
                        int imm = int.Parse(YYYMMDD[1]);
                        int idd = int.Parse(YYYMMDD[2]);
                        //DateTime dt = DateTime.ParseExact(contextinlst[0], "yyy/MM/dd", CultureInfo.InvariantCulture).AddYears(1911);
                        DateTime dt = new DateTime(1911 + iyyy, imm, idd);
                        double dbOpen = double.Parse(contextinlst[1], NumberStyles.Float | NumberStyles.AllowThousands);
                        double dbHi = double.Parse(contextinlst[2], NumberStyles.Float | NumberStyles.AllowThousands);
                        double dbLo = double.Parse(contextinlst[3], NumberStyles.Float | NumberStyles.AllowThousands);
                        double dbClose = double.Parse(contextinlst[4], NumberStyles.Float | NumberStyles.AllowThousands);
                        dtHi.Add(dt, dbHi);
                        dtLo.Add(dt, dbLo);
                        dtClose.Add(dt, dbClose);

                    }
                }

                {
                    string json = File.ReadAllText(strfilenamelastmonth);
                    JSON_Root json_Dictionary = JsonConvert.DeserializeObject<JSON_Root>(json);
                    List<List<string>> lstlist = json_Dictionary.data;
                    foreach (List<string> contextinlst in lstlist)
                    {
                        string[] YYYMMDD = contextinlst[0].Split('/');
                        int iyyy = int.Parse(YYYMMDD[0]);
                        int imm = int.Parse(YYYMMDD[1]);
                        int idd = int.Parse(YYYMMDD[2]);
                        //DateTime dt = DateTime.ParseExact(contextinlst[0], "yyy/MM/dd", CultureInfo.InvariantCulture).AddYears(1911);
                        DateTime dt = new DateTime(1911 + iyyy, imm, idd);
                        double dbOpen = double.Parse(contextinlst[1], NumberStyles.Float | NumberStyles.AllowThousands);
                        double dbHi = double.Parse(contextinlst[2], NumberStyles.Float | NumberStyles.AllowThousands);
                        double dbLo = double.Parse(contextinlst[3], NumberStyles.Float | NumberStyles.AllowThousands);
                        double dbClose = double.Parse(contextinlst[4], NumberStyles.Float | NumberStyles.AllowThousands);
                        dtHi.Add(dt, dbHi);
                        dtLo.Add(dt, dbLo);
                        dtClose.Add(dt, dbClose);

                    }
                }

            }
            {
                string strfilename = string.Format("D:\\Chips\\TWSE5MINS\\MI_5MINS_{0}.json", dtToday.ToString("yyyyMMdd"));

                string json = File.ReadAllText(strfilename);
                JSON_Root_5MINS json_Dictionary = JsonConvert.DeserializeObject<JSON_Root_5MINS>(json);
                double BuyNOT = double.Parse(json_Dictionary.data[0][1], NumberStyles.Float | NumberStyles.AllowThousands);
                double BuyQut = double.Parse(json_Dictionary.data[0][2], NumberStyles.Float | NumberStyles.AllowThousands);
                double SelNOT = double.Parse(json_Dictionary.data[0][3], NumberStyles.Float | NumberStyles.AllowThousands);
                double SelQut = double.Parse(json_Dictionary.data[0][4], NumberStyles.Float | NumberStyles.AllowThousands);

                double BuyNOTEnd = double.Parse(json_Dictionary.data[json_Dictionary.data.Count - 1][1], NumberStyles.Float | NumberStyles.AllowThousands);
                double BuyQutEnd = double.Parse(json_Dictionary.data[json_Dictionary.data.Count - 1][2], NumberStyles.Float | NumberStyles.AllowThousands);
                double SelNOTEnd = double.Parse(json_Dictionary.data[json_Dictionary.data.Count - 1][3], NumberStyles.Float | NumberStyles.AllowThousands);
                double SelQutEnd = double.Parse(json_Dictionary.data[json_Dictionary.data.Count - 1][4], NumberStyles.Float | NumberStyles.AllowThousands);
                double CumTEnd = double.Parse(json_Dictionary.data[json_Dictionary.data.Count - 1][5], NumberStyles.Float | NumberStyles.AllowThousands);
                double CumQEnd = double.Parse(json_Dictionary.data[json_Dictionary.data.Count - 1][6], NumberStyles.Float | NumberStyles.AllowThousands);
                double CumPEnd = double.Parse(json_Dictionary.data[json_Dictionary.data.Count - 1][7], NumberStyles.Float | NumberStyles.AllowThousands);

                dtBuyNOTBeg.Add(dtToday, BuyNOT);
                dtBuyQutBeg.Add(dtToday, BuyQut);
                dtSelNOTBeg.Add(dtToday, SelNOT);
                dtSelQutBeg.Add(dtToday, SelQut);
                dtBuyNOTEnd.Add(dtToday, BuyNOTEnd);
                dtBuyQutEnd.Add(dtToday, BuyQutEnd);
                dtSelNOTEnd.Add(dtToday, SelNOTEnd);
                dtSelQutEnd.Add(dtToday, SelQutEnd);
                dtCumTEnd.Add(dtToday, CumTEnd);
                dtCumQEnd.Add(dtToday, CumQEnd);
                dtCumPEnd.Add(dtToday, CumPEnd);
            }


            {
                string strfilename = string.Format("D:\\Chips\\EXRate\\{0}.csv", dtToday.ToString("yyyy"));

                IEnumerable<string> lines = File.ReadLines(strfilename);
                foreach (string line in lines)
                {
                    string[] strAtLine = line.Split(',');
                    if (strAtLine.Length == 11)
                    {
                        string strdate = strAtLine[0];
                        DateTime parsed;
                        if (DateTime.TryParse(strdate, out parsed))
                        {
                            double EXRateUSDtoNTD = double.Parse(strAtLine[1], NumberStyles.Float);
                            double EXRateRMBtoNTD = 5;
                            if (strAtLine[2] != "-")
                                EXRateRMBtoNTD = double.Parse(strAtLine[2], NumberStyles.Float);
                            double EXRateEURtoUSD = double.Parse(strAtLine[3], NumberStyles.Float);
                            double EXRateEURtoJPD = double.Parse(strAtLine[4], NumberStyles.Float);
                            double EXRateGBPtoUSD = double.Parse(strAtLine[5], NumberStyles.Float);
                            double EXRateAUDtoUSD = double.Parse(strAtLine[6], NumberStyles.Float);
                            double EXRateUSDtoHKD = double.Parse(strAtLine[7], NumberStyles.Float);
                            double EXRateUSDtoRMB = 6.22805;
                            if (strAtLine[8] != "-")
                                EXRateUSDtoRMB = double.Parse(strAtLine[8], NumberStyles.Float);
                            double EXRateUSDtoSAD = 11.32215;
                            if (strAtLine[9] != "-")
                                EXRateUSDtoSAD = double.Parse(strAtLine[9], NumberStyles.Float);

                            dtEXRateUSDtoNTD.Add(parsed, EXRateUSDtoNTD);
                            dtEXRateRMBtoNTD.Add(parsed, EXRateRMBtoNTD);
                            dtEXRateEURtoUSD.Add(parsed, EXRateEURtoUSD);
                            dtEXRateEURtoJPD.Add(parsed, EXRateEURtoJPD);
                            dtEXRateGBPtoUSD.Add(parsed, EXRateGBPtoUSD);
                            dtEXRateAUDtoUSD.Add(parsed, EXRateAUDtoUSD);
                            dtEXRateUSDtoHKD.Add(parsed, EXRateUSDtoHKD);
                            dtEXRateUSDtoRMB.Add(parsed, EXRateUSDtoRMB);
                            dtEXRateUSDtoSAD.Add(parsed, EXRateUSDtoSAD);
                            // Use reformatted
                        }
                        else
                        {
                            int nn = 0;
                            // Log error, perhaps?
                        }
                    }
                }
            }
           
            {
                try
                {
                    string strfilenameBFI82U = string.Format("D:\\Chips\\BFI82U\\BFI82U_day_{0}.json", dtToday.ToString("yyyyMMdd"));


                    string json = File.ReadAllText(strfilenameBFI82U);
                    ObjectOf3pty json_Dictionary = JsonConvert.DeserializeObject<ObjectOf3pty>(json);

                    string strSelfSelfBuy = "0";
                    string strSelfSelfSell = "0";
                    string strSelfSelfSum = "0";
                    string strSelfHedgBuy = "0";
                    string strSelfHedgSell = "0";
                    string strSelfHedgSum = "0";
                    string strTrustBuy = "0";
                    string strTrustSell = "0";
                    string strTrustSum = "0";
                    string strForBuy = "0";
                    string strForSell = "0";
                    string strForSum = "0";
                    string strAllBuy = "0";
                    string strAllSell = "0";
                    string strAllSum = "0";
                    if (json_Dictionary.data.Count == 4)
                    {
                        strSelfSelfBuy = json_Dictionary.data[0][1];
                        strSelfSelfSell = json_Dictionary.data[0][2];
                        strSelfSelfSum = json_Dictionary.data[0][3];
                        strTrustBuy = json_Dictionary.data[1][1];
                        strTrustSell = json_Dictionary.data[1][2];
                        strTrustSum = json_Dictionary.data[1][3];
                        strForBuy = json_Dictionary.data[2][1];
                        strForSell = json_Dictionary.data[2][2];
                        strForSum = json_Dictionary.data[2][3];
                        strAllBuy = json_Dictionary.data[3][1];
                        strAllSell = json_Dictionary.data[3][2];
                        strAllSum = json_Dictionary.data[3][3];

                    }
                    else
                    {
                        strSelfSelfBuy = json_Dictionary.data[0][1];
                        strSelfSelfSell = json_Dictionary.data[0][2];
                        strSelfSelfSum = json_Dictionary.data[0][3];
                        strSelfHedgBuy = json_Dictionary.data[1][1];
                        strSelfHedgSell = json_Dictionary.data[1][2];
                        strSelfHedgSum = json_Dictionary.data[1][3];
                        strTrustBuy = json_Dictionary.data[2][1];
                        strTrustSell = json_Dictionary.data[2][2];
                        strTrustSum = json_Dictionary.data[2][3];
                        strForBuy = json_Dictionary.data[3][1];
                        strForSell = json_Dictionary.data[3][2];
                        strForSum = json_Dictionary.data[3][3];
                        strAllBuy = json_Dictionary.data[4][1];
                        strAllSell = json_Dictionary.data[4][2];
                        strAllSum = json_Dictionary.data[4][3];
                    }


                    dbSelfSelfBuy = double.Parse(strSelfSelfBuy, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                    dbSelfSelfSell = double.Parse(strSelfSelfSell, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                    dbSelfSelfSum = double.Parse(strSelfSelfSum, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                    dbSelfHedgBuy = double.Parse(strSelfHedgBuy, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                    dbSelfHedgSell = double.Parse(strSelfHedgSell, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                    dbSelfHedgSum = double.Parse(strSelfHedgSum, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                    dbTrustBuy = double.Parse(strTrustBuy, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                    dbTrustSell = double.Parse(strTrustSell, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                    dbTrustSum = double.Parse(strTrustSum, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                    dbForBuy = double.Parse(strForBuy, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                    dbForSell = double.Parse(strForSell, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                    dbForSum = double.Parse(strForSum, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                    dbAllBuy = double.Parse(strAllBuy, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                    dbAllSell = double.Parse(strAllSell, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                    dbAllSum = double.Parse(strAllSum, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);



                    dbForBuy += dbAllBuy;
                    dbForSell += dbAllSell;
                    dbForSum += dbAllSum;

                    strAllBuy = json_Dictionary.data[5][1];
                    strAllSell = json_Dictionary.data[5][2];
                    strAllSum = json_Dictionary.data[5][3];

                    dbAllBuy = double.Parse(strAllBuy, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                    dbAllSell = double.Parse(strAllSell, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                    dbAllSum = double.Parse(strAllSum, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);

                    string strAddmore = string.Format(",{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14}", dbSelfSelfBuy, dbSelfSelfSell, dbSelfSelfSum, dbSelfHedgBuy, dbSelfHedgSell, dbSelfHedgSum,
                        dbTrustBuy, dbTrustSell, dbTrustSum, dbForBuy, dbForSell, dbForSum, dbAllBuy, dbAllSell, dbAllSum);
                    // File.AppendAllText(strmake3tpyfilename, line + strAddmore + "\r\n");
                }
                catch (Exception exp)
                {

                }
            }

            DateTime dtLastTime = dtToday.AddDays(-1);
            while (!dtClose.ContainsKey(dtLastTime))
            {
                dtLastTime = dtLastTime.AddDays(-1);
            }
            string strwritefilename = "today.csv";
            File.Delete(strwritefilename);

            string stradd = string.Format("Date,HiDiffidIn1Day,LoDiffIn1Day,HiDiffidIn2Days,LoDiffIn2Days,HiDiffidIn3Days,LoDiffIn3Days,Close,High,Low,Close-1,High-1,Low-1,BuyNOT,BuyQut,SelNOT,SelQut,BuyNOTEnd,BuyQutEnd,SelNOTEnd,SelQutEnd,CumTEnd,CumQEnd,CumPEnd,USDtoNTD,RMBtoNTD,EURtoUSD,EURtoJPD,GBPtoUSD,AUDtoUSD,USDtoHKD,USDtoRMB,USDtoSAD,FinBuyQut,FinSelQut,FinCashReturn,FinTodayQut,FinYesterQut,MarBuyQut,MarSelQut,MarCashReturn,MarTodayQut,MarYesterQut,FinBuyPri,FinSelPri,FinCashReturnPri,FinTodayPri,FinYesterPri,SelfSelfBuy,SelfSelfSell,SelfSelfSum,SelfHedgeBuy,SelfHedgeSell,SelfHedgeSum,TrustBuy,TrustSell,TrustSum,ForBuy,ForSell,ForSum,AllBuy,AllSell,AllSum");
            stradd += "\r\n";
            File.AppendAllText(strwritefilename, stradd, Encoding.UTF8);

            stradd = "";
            string straddvalue = string.Format("{0},NA,NA,NA,NA,NA,NA,{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},", dtToday.ToString("yyyy/MM/dd")
                                                                       , dtClose[dtToday]
                                                                       , dtHi[dtToday]
                                                                       , dtLo[dtToday]
                                                                       , dtClose[dtLastTime]
                                                                       , dtHi[dtLastTime]
                                                                       , dtLo[dtLastTime]
                                                                       , dtBuyNOTBeg[dtToday]
                                                                       , dtBuyQutBeg[dtToday]
                                                                       , dtSelNOTBeg[dtToday]
                                                                       , dtSelQutBeg[dtToday]
                                                                       , dtBuyNOTEnd[dtToday]
                                                                       , dtBuyQutEnd[dtToday]
                                                                       , dtSelNOTEnd[dtToday]
                                                                       , dtSelQutEnd[dtToday]
                                                                       , dtCumTEnd[dtToday]
                                                                       , dtCumQEnd[dtToday]
                                                                       , dtCumPEnd[dtToday]);


            stradd += string.Format("{0},", dtEXRateUSDtoNTD[dtToday]);
            stradd += string.Format("{0},", dtEXRateRMBtoNTD[dtToday]);
            stradd += string.Format("{0},", dtEXRateEURtoUSD[dtToday]);
            stradd += string.Format("{0},", dtEXRateEURtoJPD[dtToday]);
            stradd += string.Format("{0},", dtEXRateGBPtoUSD[dtToday]);
            stradd += string.Format("{0},", dtEXRateAUDtoUSD[dtToday]);
            stradd += string.Format("{0},", dtEXRateUSDtoHKD[dtToday]);
            stradd += string.Format("{0},", dtEXRateUSDtoRMB[dtToday]);
            stradd += string.Format("{0},", dtEXRateUSDtoSAD[dtToday]);


            stradd += string.Format("NA,");
            stradd += string.Format("NA,");
            stradd += string.Format("NA,");
            stradd += string.Format("NA,");
            stradd += string.Format("NA,");
            stradd += string.Format("NA,");
            stradd += string.Format("NA,");
            stradd += string.Format("NA,");
            stradd += string.Format("NA,");
            stradd += string.Format("NA,");
            stradd += string.Format("NA,");
            stradd += string.Format("NA,");
            stradd += string.Format("NA,");
            stradd += string.Format("NA,");
            stradd += string.Format("NA,");
            straddvalue += stradd;


            string strAdd3pty = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14}", dbSelfSelfBuy, dbSelfSelfSell, dbSelfSelfSum, dbSelfHedgBuy, dbSelfHedgSell, dbSelfHedgSum,
    dbTrustBuy, dbTrustSell, dbTrustSum, dbForBuy, dbForSell, dbForSum, dbAllBuy, dbAllSell, dbAllSum);
            straddvalue += strAdd3pty;

            straddvalue += "\r\n";
            File.AppendAllText(strwritefilename, straddvalue, Encoding.UTF8);



        }

    }
}

