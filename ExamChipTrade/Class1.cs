using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExamChipTrade
{
    public partial class Form1 : Form
    {
        public class StockLevelCount
        {
            public long[] LevelPeople = new long[17];
            public long[] LevelVol = new long[17];
            public double[] LevelRate = new double[17];
        }

        class PartyParam
        {
            public string number { get; set; }
            public string name { get; set; }

            public string ForeigenBuy { get; set; }
            public string ForeigenSell { get; set; }
            public string ForeigenTotal { get; set; }
            public string ForeigenSelfBuy { get; set; }
            public string ForeigenSelfSell { get; set; }
            public string ForeigenSelfTotal { get; set; }
            public string ForeigenMainlandBuy { get; set; }
            public string ForeigenMainlandSell { get; set; }
            public string ForeigenMainlandTotal { get; set; }

            public string TrustBuy { get; set; }
            public string TrustSell { get; set; }
            public string TrustTotal { get; set; }

            public string SelfBuy { get; set; }
            public string SelfSell { get; set; }
            public string SelfTotal { get; set; }
            public string SelfSelfBuy { get; set; }
            public string SelfSelfSell { get; set; }
            public string SelfSelfTotal { get; set; }
            public string SelfHedgeBuy { get; set; }
            public string SelfHedgeSell { get; set; }
            public string SelfHedgeTotal { get; set; }

            public string Party3Total { get; set; }

            public void Assign(List<string> LS)
            {
                if(LS.Count == 19)
                {
                    number = LS.ElementAt(0);
                    name = LS.ElementAt(1);
                    ForeigenBuy = LS.ElementAt(2);
                    ForeigenSell = LS.ElementAt(3);
                    ForeigenTotal = LS.ElementAt(4);
                    ForeigenSelfBuy = LS.ElementAt(5);
                    ForeigenSelfSell = LS.ElementAt(6);
                    ForeigenSelfTotal = LS.ElementAt(7);
                    TrustBuy = LS.ElementAt(8);
                    TrustSell = LS.ElementAt(9);
                    TrustTotal = LS.ElementAt(10);
                    SelfTotal = LS.ElementAt(11);
                    SelfSelfBuy = LS.ElementAt(12);
                    SelfSelfSell = LS.ElementAt(13);
                    SelfSelfTotal = LS.ElementAt(14);
                    SelfHedgeBuy = LS.ElementAt(15);
                    SelfHedgeSell = LS.ElementAt(16);
                    SelfHedgeTotal = LS.ElementAt(17);
                    Party3Total = LS.ElementAt(18);
                }
                else if (LS.Count == 16)
                {
                    //1701 2020/6/3
                }
            }

            public void AssignTPEX(List<string> LS)
            {
                number = LS.ElementAt(0);
                name = LS.ElementAt(1);
                ForeigenBuy = LS.ElementAt(2);
                ForeigenSell = LS.ElementAt(3);
                ForeigenTotal = LS.ElementAt(4);
                ForeigenSelfBuy = LS.ElementAt(5);
                ForeigenSelfSell = LS.ElementAt(6);
                ForeigenSelfTotal = LS.ElementAt(7);
                ForeigenMainlandBuy = LS.ElementAt(8);
                ForeigenMainlandSell = LS.ElementAt(9);
                ForeigenMainlandTotal = LS.ElementAt(10);
                TrustBuy = LS.ElementAt(11);
                TrustSell = LS.ElementAt(12);
                TrustTotal = LS.ElementAt(13);

                SelfSelfBuy = LS.ElementAt(14);
                SelfSelfSell = LS.ElementAt(15);
                SelfSelfTotal = LS.ElementAt(16);

                SelfHedgeBuy = LS.ElementAt(17);
                SelfHedgeSell = LS.ElementAt(18);
                SelfHedgeTotal = LS.ElementAt(19);

                SelfBuy = LS.ElementAt(20);
                SelfSell = LS.ElementAt(21);
                SelfTotal = LS.ElementAt(22);

                Party3Total = LS.ElementAt(23);
            }
        }


        class PartyParamLong
        {
            public string number { get; set; }
            public string name { get; set; }

            public long ForeigenBuy { get; set; }
            public long ForeigenSell { get; set; }
            public long ForeigenTotal { get; set; }
            public long ForeigenSelfBuy { get; set; }
            public long ForeigenSelfSell { get; set; }
            public long ForeigenSelfTotal { get; set; }
            public long ForeigenMainlandBuy { get; set; }
            public long ForeigenMainlandSell { get; set; }
            public long ForeigenMainlandTotal { get; set; }

            public long TrustBuy { get; set; }
            public long TrustSell { get; set; }
            public long TrustTotal { get; set; }

            public long SelfBuy { get; set; }
            public long SelfSell { get; set; }
            public long SelfTotal { get; set; }
            public long SelfSelfBuy { get; set; }
            public long SelfSelfSell { get; set; }
            public long SelfSelfTotal { get; set; }
            public long SelfHedgeBuy { get; set; }
            public long SelfHedgeSell { get; set; }
            public long SelfHedgeTotal { get; set; }

            public long Party3Total { get; set; }

            public void Assign(List<string> LS)
            {
                number = LS.ElementAt(0);
                name = LS.ElementAt(1);
                ForeigenBuy = long.Parse(LS.ElementAt(2), NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                ForeigenSell = long.Parse(LS.ElementAt(3), NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                ForeigenTotal = long.Parse(LS.ElementAt(4), NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                ForeigenSelfBuy = long.Parse(LS.ElementAt(5), NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                ForeigenSelfSell = long.Parse(LS.ElementAt(6), NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                ForeigenSelfTotal = long.Parse(LS.ElementAt(7), NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                TrustBuy = long.Parse(LS.ElementAt(8), NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                TrustSell = long.Parse(LS.ElementAt(9), NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                TrustTotal = long.Parse(LS.ElementAt(10), NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);  
                SelfTotal = long.Parse(LS.ElementAt(11), NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);   
                SelfSelfBuy = long.Parse(LS.ElementAt(12), NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign); 
                SelfSelfSell = long.Parse(LS.ElementAt(13), NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                SelfSelfTotal = long.Parse(LS.ElementAt(14), NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                SelfHedgeBuy = long.Parse(LS.ElementAt(15), NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign); 
                SelfHedgeSell = long.Parse(LS.ElementAt(16), NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                SelfHedgeTotal = long.Parse(LS.ElementAt(17), NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                Party3Total = long.Parse(LS.ElementAt(18), NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
            }

            public void AssignTPEX(List<string> LS)
            {
                number = LS.ElementAt(0);
                name = LS.ElementAt(1);
                ForeigenBuy = long.Parse(LS.ElementAt(2), NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                ForeigenSell = long.Parse(LS.ElementAt(3), NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign); 
                ForeigenTotal = long.Parse(LS.ElementAt(4), NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                ForeigenSelfBuy = long.Parse(LS.ElementAt(5), NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign); 
                ForeigenSelfSell = long.Parse(LS.ElementAt(6), NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                ForeigenSelfTotal = long.Parse(LS.ElementAt(7), NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign); 
                ForeigenMainlandBuy = long.Parse(LS.ElementAt(8), NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                ForeigenMainlandSell = long.Parse(LS.ElementAt(9), NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                ForeigenMainlandTotal = long.Parse(LS.ElementAt(10), NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                TrustBuy = long.Parse(LS.ElementAt(11), NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign); 
                TrustSell = long.Parse(LS.ElementAt(12), NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                TrustTotal = long.Parse(LS.ElementAt(13), NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);

                SelfSelfBuy = long.Parse(LS.ElementAt(14), NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign); 
                SelfSelfSell = long.Parse(LS.ElementAt(15), NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                SelfSelfTotal = long.Parse(LS.ElementAt(16), NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);

                SelfHedgeBuy = long.Parse(LS.ElementAt(17), NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign); 
                SelfHedgeSell = long.Parse(LS.ElementAt(18), NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                SelfHedgeTotal = long.Parse(LS.ElementAt(19), NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);

                SelfBuy = long.Parse(LS.ElementAt(20), NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign); 
                SelfSell = long.Parse(LS.ElementAt(21), NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                SelfTotal = long.Parse(LS.ElementAt(22), NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);

                Party3Total = long.Parse(LS.ElementAt(23), NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign); 
            }
        }


        public List<string> GetStocks(DateTime dtDay,int iRange,double dbPercent) //當天，前幾天，3PARTY買的百分比
        {
            List<string> ret = new List<string>();
            DateTime dtLoopStart = dtDay;

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

            return ret;
        }

        public double StandardDeviation(IEnumerable<double> values, out double avgout)
        {
            double avg = values.Average();
            avgout = avg;
            return Math.Sqrt(values.Average(v => Math.Pow(v - avg, 2)));
        }


        public double StandardDeviation(Dictionary<DateTime, double> dicPrice, int CountDays, out double avgout)
        {
            if (dicPrice.Count < CountDays)
            {
                avgout = 0;
                return -1;
            }

            List<double> dbList = new List<double>();
            DateTime dtStart = dicPrice.Keys.Max<DateTime>();

            while (dbList.Count < CountDays)
            {
                if (dicPrice.ContainsKey(dtStart))
                {
                    dbList.Add(dicPrice[dtStart]);
                }
                dtStart = dtStart.AddDays(-1);
            }
            double dbavg = dicPrice.Values.Average();
            avgout = dbavg;

            return Math.Sqrt(dicPrice.Values.Average(v => Math.Pow(v - dbavg, 2)));
        }

        public double StandardDeviation(Dictionary<DateTime, int> dicPrice, int CountDays,out double avgout)
        {
            if (dicPrice.Count < CountDays)
            {
                avgout = 0;
                return -1;
            }
                

            List<double> dbList = new List<double>();
            DateTime dtStart = dicPrice.Keys.Max<DateTime>();

            while (dbList.Count < CountDays)
            {
                if(dicPrice.ContainsKey(dtStart))
                {
                    dbList.Add(dicPrice[dtStart]);
                }
                dtStart = dtStart.AddDays(-1);
            }
            double dbavg = dicPrice.Values.Average();
            avgout = dbavg;

            return Math.Sqrt((double)dicPrice.Values.Average(v => Math.Pow(v - dbavg, 2)));
            
        }

        public double SMA(Dictionary<DateTime, double[]> dicOLHCPrice,DateTime dtDay, int Days,int iOLHC=3) //3:close 2:High 1:Low 0:Open
        {
            if (dicOLHCPrice.Count < Days)
            {
                return -1;
            }
            DateTime dtMax = dicOLHCPrice.Keys.Max<DateTime>();
            DateTime dtMin = dicOLHCPrice.Keys.Min<DateTime>();
            if(dtDay>dtMax)
            {
                return -1;
            }
            List<double> dbList = new List<double>();
            DateTime dtStart = dtDay;
            while (dbList.Count < Days)
            {
                if (dicOLHCPrice.ContainsKey(dtStart))
                {
                    dbList.Add(dicOLHCPrice[dtStart][iOLHC]);
                }
                dtStart = dtStart.AddDays(-1);
                if(dtStart<dtMin)
                {
                    return -1;
                }
            }
            return dbList.Average();
        }
        public double STDev(Dictionary<DateTime, double[]> dicOLHCPrice, DateTime dtDay, int Days, out double avgout) 
        {
            if (dicOLHCPrice.Count < Days)
            {
                avgout = -1;
                return -1;
            }
            DateTime dtMax = dicOLHCPrice.Keys.Max<DateTime>();
            DateTime dtMin = dicOLHCPrice.Keys.Min<DateTime>();
            if (dtDay > dtMax)
            {
                avgout = -1;
                return -1;
            }
            List<double> dbList = new List<double>();
            DateTime dtStart = dtDay;
            while (dbList.Count < Days)
            {
                if (dicOLHCPrice.ContainsKey(dtStart))
                {
                    dbList.Add(dicOLHCPrice[dtStart][3]); //3:close 2:High 1:Low 0:Open
                }
                dtStart = dtStart.AddDays(-1);
                if (dtStart < dtMin)
                {
                    avgout = -1;
                    return -1;
                }
            }
            double dbavg = dbList.Average();
            avgout = dbavg;
            return Math.Sqrt((double)dbList.Average(v => Math.Pow(v - dbavg, 2)));
        }


        public bool GetKeyBuy(string strsn,int idays,out int ibuy,out int isell)
        {
            // 主力1 5 10 20 60 120
            //https://fubon-ebrokerdj.fbs.com.tw/z/zc/zco/zco_8155_6.djhtm


            string url = string.Format("https://fubon-ebrokerdj.fbs.com.tw/z/zc/zco/zco_{0}_{1}.djhtm", strsn, idays);

            HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
            httpWebRequest.ContentType = "application/x-www-form-urlencoded;charset=big5";
            httpWebRequest.Method = WebRequestMethods.Http.Get;

            string text = "";
            var response = (HttpWebResponse)httpWebRequest.GetResponse();
            Stream stream = response.GetResponseStream();


            StreamReader read = new StreamReader(stream, Encoding.UTF8);
            text = read.ReadToEnd();
            //File.WriteAllText("D:\\temp.days", text);


            response.Close();
            httpWebRequest.Abort();

            HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
            htmlDoc.LoadHtml(text);
            // find <TR id="oScrollFoot">
            Queue<string> strQtr = new Queue<string>();
            HtmlNode elementScoll = htmlDoc.GetElementbyId("oScrollFoot");
            foreach (HtmlNode table in htmlDoc.DocumentNode.SelectNodes("//table"))
            {
                if (table.Id == "oMainTable")
                {
                    foreach (HtmlNode trt in table.SelectNodes("//tr"))
                    {
                        //id = oMainTable
                        if (trt.Id == "oScrollFoot")
                        {

                            foreach (HtmlNode tdt in trt.SelectNodes("td"))
                            {
                                string strtd = tdt.InnerText.Trim();
                                strQtr.Enqueue(strtd);
                            }
                            break;
                        }
                    }
                }
            }

            if(strQtr.Count!=4)
            {
                ibuy = -1;
                isell = -1;
                return false;

            }


            ibuy = int.Parse(strQtr.ElementAt(1), NumberStyles.AllowThousands);
            isell = int.Parse(strQtr.ElementAt(3), NumberStyles.AllowThousands);
            return true;
        }
    }
}
