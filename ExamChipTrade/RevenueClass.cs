using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExamChipTrade
{
    class RevenueClass
    {
        public string number { get; set; }
        public string name { get; set; }
        public string catagory { get; set; }
        public long incomethismonth { get; set; }
        public long incomelastmonth { get; set; }
        public long incomelastyearthismonth { get; set; }
        public double comparelastmonth { get; set; }
        public double comparelastyearthismonth { get; set; }
        public long accincome { get; set; }
        public long accincomelastyear { get; set; }
        public double acccomparelastyear { get; set; }

        public void AssignTWSE(string strLine)
        {   //  0        1        2        3        4           5                   6               7                     8                         9                           10                      11                        12
            //出表日期,資料年月,公司代號,公司名稱,產業別,營業收入-當月營收,營業收入-上月營收,營業收入-去年當月營收,營業收入-上月比較增減(%),營業收入-去年同月增減(%),累計營業收入-當月累計營收,累計營業收入-去年累計營收,累計營業收入-前期比較增減(%),備註
            string[] strAry = strLine.Split(',');
            if(strAry.Length == 14)
            {
                string[] strNoAry = strAry[2].Split('\"');
                if(strNoAry.Length == 3)
                {
                    number = strNoAry[1];
                }
                string[] strNameAry = strAry[3].Split('\"');
                if (strNameAry.Length == 3)
                {
                    name = strNameAry[1];
                }
                string[] strcatAry = strAry[4].Split('\"');
                if (strcatAry.Length == 3)
                {
                    catagory = strcatAry[1];
                }
                string[] strincAry = strAry[5].Split('\"');
                if (strincAry.Length == 3)
                {
                    string strinc = strincAry[1];
                    incomethismonth = long.Parse(strinc);
                }
                string[] strlastincAry = strAry[6].Split('\"');
                if (strlastincAry.Length == 3)
                {
                    string strinc = strlastincAry[1];
                    incomelastmonth = long.Parse(strinc);
                }
                string[] strlastyearincAry = strAry[7].Split('\"');
                if (strlastyearincAry.Length == 3)
                {
                    string strinc = strlastyearincAry[1];
                    incomelastyearthismonth = long.Parse(strinc);
                }
                string[] strcomparelastmonthAry = strAry[8].Split('\"');
                if (strcomparelastmonthAry.Length == 3)
                {
                    string strinc = strcomparelastmonthAry[1];
                    if (strinc != "")
                    comparelastmonth = double.Parse(strinc, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                }
                string[] strcomparelastyearAry = strAry[9].Split('\"');
                if (strcomparelastyearAry.Length == 3)
                {
                    string strinc = strcomparelastyearAry[1];
                    if (strinc != "")
                        comparelastyearthismonth = double.Parse(strinc, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                }

                string[] straccincomeAry = strAry[10].Split('\"');
                if (straccincomeAry.Length == 3)
                {
                    string strinc = straccincomeAry[1];
                    accincome = long.Parse(strinc);
                }
                string[] straccincomelastyearAry = strAry[11].Split('\"');
                if (straccincomelastyearAry.Length == 3)
                {
                    string strinc = straccincomelastyearAry[1];
                    accincomelastyear = long.Parse(strinc);
                }

                string[] stracccomparelastyearAry = strAry[12].Split('\"');
                if (stracccomparelastyearAry.Length == 3)
                {
                    string strinc = stracccomparelastyearAry[1];
                    if (strinc != "")
                    acccomparelastyear = double.Parse(strinc, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);
                }
            }


        }

        public void AssignTPEXcsv(string strLine)
        {
            //     0        2         3       4        5         6        7        8           9               10
            //  公司名稱,本年上月,本年本月,本年累計,上年本月,上年累計,累計增減,累計增減%,本年背書保證金,上年背書保證金%
            string[] strAry = strLine.Split(',');
            if(strAry.Length == 11)
            {
                string[] strNoNameAry = strAry[0].Split(' ');
                if(strNoNameAry.Length == 2)
                {
                    number = strNoNameAry[0];
                    name = strNoNameAry[1];

                }

                string strlastmonth = strAry[2];
                string strthismonth = strAry[3];
                string straccincome = strAry[4];

                string strincomelastyearthismonth = strAry[5];
                string straccincomelastyear = strAry[6];
                string strSum = strAry[7];
                string strSumpercent = strAry[8];

                if (name != "Average")
                {
                    if (strlastmonth != "")
                        incomelastmonth = long.Parse(strlastmonth);
                    if (strthismonth != "")
                        incomethismonth = long.Parse(strthismonth);
                    if (straccincome != "")
                        accincome = long.Parse(straccincome);
                    if (strincomelastyearthismonth != "")
                        incomelastyearthismonth = long.Parse(strincomelastyearthismonth);
                    if (straccincomelastyear != "")
                        accincomelastyear = long.Parse(straccincomelastyear);

                    if (strSumpercent != "" && strSumpercent != "NA")
                        comparelastyearthismonth = double.Parse(strSumpercent, NumberStyles.Float | NumberStyles.AllowThousands | NumberStyles.AllowLeadingSign);

                }
                
                int m = 0;
            }

        }
    }
}
