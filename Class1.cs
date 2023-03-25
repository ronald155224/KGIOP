using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Net;
using System.IO;
using System.Text;
using System.Web;
using System.Collections.Generic;
using System.Collections.Specialized;
using HtmlAgilityPack;

namespace OP.AutoTrading
{
    public class TQuoteStruct
    {
        public string StrickPrice { get; set; }
        public string Stock { get; set; }
        public string OpenPrice { get; set; }
        public string UpDown { get; set; }
        public string Vol { get; set; }
        public string OI { get; set; }
        public string TotalVol { get; set; }
        public string HundfifDif { get; set; }
        public string HundDif { get; set; }
        public string DealPrice { get; set; }
        public string Symbol { get; set; }
        public string StrickPriceCenter { get; set; }
        public string Symbol2 { get; set; }
        public string DealPrice2 { get; set; }
        public string HundDif2 { get; set; }
        public string HundfifDif2 { get; set; }
        public string TotalVol2 { get; set; }
        public string OI2 { get; set; }
        public string Vol2 { get; set; }
        public string UpDown2 { get; set; }
        public string OpenPrice2 { get; set; }
        public string Stock2 { get; set; }

        public TQuoteStruct(List<string> tQuoteList)
        {
            StrickPrice = tQuoteList[0];
            Stock = tQuoteList[1];
            OpenPrice = tQuoteList[2];
            UpDown = tQuoteList[3];
            Vol = tQuoteList[4];
            OI = tQuoteList[5];
            TotalVol = tQuoteList[6];
            HundfifDif = tQuoteList[7];
            HundDif = tQuoteList[8];
            DealPrice = tQuoteList[9];
            Symbol = tQuoteList[10];
            StrickPriceCenter = tQuoteList[11];
            Symbol2 = tQuoteList[12];
            DealPrice2 = tQuoteList[13];
            HundDif2 = tQuoteList[14];
            HundfifDif2 = tQuoteList[15];
            TotalVol2 = tQuoteList[16];
            OI2 = tQuoteList[17];
            Vol2 = tQuoteList[18];
            UpDown2 = tQuoteList[19];
            OpenPrice2 = tQuoteList[20];
            Stock2 = tQuoteList[21];
        }
    }

    public class taifexCrawler
    {
        public List<List<string>> TaiFexList = new List<List<string>>();
        private DateTime dateTime = DateTime.Now;

        public void doCrawler()
        {
            //HtmlNodeCollection nodes;
            dateTime = DateTime.Now;

            opLog.txtLog("DoCrawler DateTime: " + dateTime.ToString("yyyy/MM/dd"));

            //14:00前只能拿昨天的OI，14:00過後才能拿到今日的OI
            switch (Convert.ToInt32(dateTime.DayOfWeek.ToString("d")))
            {
                case 6:
                    dateTime = dateTime.AddDays(-1);
                    break;
                case 7:
                    dateTime = dateTime.AddDays(-2);
                    break;
                default:
                    if (DateCalculateExtensions.IsLegalTime(dateTime, "000000-050000")) dateTime = dateTime.AddDays(-1);
                    else if (DateCalculateExtensions.IsLegalTime(dateTime, "050001-143000")) dateTime = dateTime.AddDays(-1);
                    break;
            }

            if (!read())
            {
                string url = "https://www.taifex.com.tw/cht/3/optDailyMarketReport";
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);


                //var wb = new WebClient();
                //var data = new NameValueCollection();

                var postData = "queryDate=" + Uri.EscapeDataString(dateTime.ToString("yyyy/MM/dd"));
                postData += "&queryType=" + Uri.EscapeDataString("2");
                postData += "&marketCode=" + Uri.EscapeDataString("0");
                postData += "&MarketCode=" + Uri.EscapeDataString("0");
                //postData += "dateaddcnt" + Uri.EscapeDataString("1");
                postData += "&commodity_id=" + Uri.EscapeDataString("TXO");
                postData += "&commodity_idt=" + Uri.EscapeDataString("TXO");
                var data = Encoding.ASCII.GetBytes(postData);

                request.UserAgent = "Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:47.0) Gecko/20100101 Firefox/47.0";
                request.Method = "POST";
                request.ContentType = "application/x-www-form-urlencoded";
                request.ContentLength = data.Length;

                using (var stream = request.GetRequestStream())
                {
                    stream.Write(data, 0, data.Length);
                }

                string responseStr = "";
                var doc = new HtmlAgilityPack.HtmlDocument();

                using (WebResponse response = (HttpWebResponse)request.GetResponse())
                {
                    using (StreamReader sr = new StreamReader(response.GetResponseStream(), Encoding.UTF8))
                    {
                        responseStr = sr.ReadToEnd();
                    }
                }

                //taifex網站爬蟲後會出現一個只有結尾的</td>Tag，所以需要把這多餘的Tag清除
                if(Regex.IsMatch(responseStr, @"&nbsp;查無資料"))
                    opLog.txtLog("查無期交所OI資料!!!!");
                else
                {
                    responseStr = Regex.Replace(responseStr, @"^\t+<\/td>[\r\n]+\t+[\r\n]+", "", RegexOptions.Multiline);
                    //Console.WriteLine(responseStr);
                    doc.LoadHtml(responseStr);                
                    TaiFexList = doc.DocumentNode.SelectSingleNode("//table[@class='table_f']")
                                        .Descendants("tr")
                                        .Skip(1)
                                        .Where(tr => tr.Elements("td").Count() > 1)
                                        .Select(tr => tr.Elements("td").Select(td => td.InnerText.Trim()).ToList())
                                        .ToList();
                }

                for (int i = 0; i < TaiFexList.Count(); i++)
                {
                    for (int j = 0; j < TaiFexList[i].Count(); j++)
                    {
                        opLog.txtLog("TaiFexList[" + i + "][" + j + "]: " + TaiFexList[i][j]);
                    }
                    opLog.txtLog("");
                }
                if (TaiFexList.Count == 0)
                {
                    opLog.txtLog("Can't get the response of httpPost!");
                }
                else
                {
                    //for (int i = 0; i < TaiFexList.Count(); i++)
                    //{
                    //    for (int j = 0; j < TaiFexList[i].Count(); j++)
                    //    {
                    //        opLog.txtLog("TaiFexList[" + i + "][" + j + "]: " + TaiFexList[i][j]);
                    //    }
                    //    opLog.txtLog("");
                    //}
                    save();
                }
            }

        }
        private void save()
        {
            string dir = Application.StartupPath + @"\Crawler\";
            opLog.txtLog("CrawlerToTxt " + dateTime.ToString("yyyyMMdd"));
            string fileName = dir + dateTime.ToString("yyyyMMdd") + ".txt";
            try
            {
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                if (!File.Exists(fileName))
                {
                    string row;
                    if (TaiFexList.Count() > 0)
                    {
                        FileStream fs = new FileStream(fileName, FileMode.OpenOrCreate, FileAccess.Write);
                        StreamWriter sw = new StreamWriter(fs);
                        sw.Flush();
                        sw.BaseStream.Seek(0, SeekOrigin.Begin);

                        for (int i = 0; i < TaiFexList.Count(); i++)
                        {
                            row = TaiFexList[i][0];
                            for (int j = 1; j < TaiFexList[i].Count(); j++)
                            {
                                row += "," + TaiFexList[i][j];
                            }
                            sw.WriteLine(row);
                        }

                        sw.Flush();
                        sw.Close();
                        fs.Close();
                    }
                }
                else opLog.txtLog("Had done crawler!!");
            }
            catch (IOException ex)
            {
                opLog.txtLog(
                    ex.GetType().Name + ": The write operation could not " +
                    "be performed because the specified " +
                    "part of the file is locked.");
            }
        }
        private bool read()
        {
            string dir = Application.StartupPath + @"\Crawler\";
            opLog.txtLog("ReadCrawlerTxt " + dateTime.ToString("yyyyMMdd"));
            string fileName = dir + dateTime.ToString("yyyyMMdd") + ".txt";
            string row = "";
            string[] rowSplit = new string[19];

            if (File.Exists(fileName))
            {
                TaiFexList.Clear();
                FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
                StreamReader sr = new StreamReader(fs);

                sr.BaseStream.Seek(0, SeekOrigin.Begin);
                row = sr.ReadLine();
                //Console.WriteLine("Row: "+ row);
                List<string> list = new List<string>();
                while ( row != null)
                {
                    row = row.Trim();
                    rowSplit = row.Split(',');
                    list = new List<string>();
                    for (int i = 0; i < rowSplit.Count(); i++)
                    {
                        list.Add(rowSplit[i]);
                    }
                    TaiFexList.Add(list);
                    row = sr.ReadLine();
                    //Console.WriteLine("Row: " + row);
                }
                //Console.WriteLine("before Close!");
                sr.Close();
                fs.Close();
                //Console.WriteLine("AFTER Close!");
                return true;
            }
            return false;
        }
    }
    public class dgvTQuote
    {
        public decimal StrikeP { get; set; }
        public decimal CallUpDown { get; set; }
        public int CallVol { get; set; }
        public int CallOI { get; set; }
        public decimal CallHundFifDif { get; set; }
        public decimal CallHundDif { get; set; }
        public decimal CallPrice { get; set; }
        public string CallSymbol { get; set; }
        public decimal StrikePCenter { get; set; }
        public string PutSymbol { get; set; }
        public decimal PutHundDif { get; set; }
        public decimal PutHundFifDif { get; set; }
        public int PutOI { get; set; }
        public int PutVol { get; set; }
        public decimal PutUpDown { get; set; }
        public decimal PutStrikeP { get; set; }
    }

    class OrderSettings
    {
        public string StrSettingIni;
        public List<List<decimal>> DailyPoints { get; set; }//DailyPoints[]: WED日,WED夜,THR日... DailyPoints[][0]:callpoints [][1]:putpoints [][2]:停利
        public decimal CallDownLim { get; set; }
        public decimal CallUpLim { get; set; }
        public decimal PutDownLim { get; set; }
        public decimal PutUpLim { get; set; }
        public decimal UpSafeDis { get; set; }
        public decimal DownSafeDis { get; set; }
        public decimal TakeProfitPointsSum { get; set; }
        public decimal TakeProfitAcc { get; set; }
        public ushort TakeProfitLots { get; set; }
        public decimal TakeProfitGap { get; set; }
        public decimal StopLossMagn { get; set; }
        public ushort StopLossLots { get; set; }
        public decimal StopLossGap { get; set; }
        public ushort EachOrderLots { get; set; }
        public decimal EachOrderGap { get; set; }
        public decimal OrderPriceShift { get; set; }
        public decimal ProfitStopCountMin { get; set; }
        public decimal ProfitStopSuccessRound { get; set; }
        public decimal ProfitStopFailRound { get; set; }
        public decimal EmergencyCountMin { get; set; }
        public decimal EmergencySuccessRound { get; set; }
        public decimal EmergencyFailRound { get; set; }
        public decimal AtleastMargin { get; set; }

        public OrderSettings(string _strSettingIni)
        {
            StrSettingIni = _strSettingIni;
            LoadOrderSettins();
        }
        public void LoadOrderSettins()
        {
            string allSetting = "";
            string[] strSettingSplit;
            List<List<decimal>> tmpDailyPoints = new List<List<decimal>>();
            List<decimal> sublist = new List<decimal>();
            try
            {
                if (!File.Exists(StrSettingIni))
                {
                    opLog.txtLog("The setting.ini file doesn't exist."
                        + Environment.NewLine
                        + "Create this file automatically!");

                    string dir = Application.StartupPath;
                    string fileName = dir + @"\" + StrSettingIni;

                    File.Create(fileName).Close();
                }
                else
                {
                    allSetting = System.IO.File.ReadAllText(StrSettingIni);

                    if (string.IsNullOrEmpty(allSetting))
                    {
                        opLog.txtLog("The setting is null."
                            + Environment.NewLine
                            + "Please fill this file with setting info.");
                    }
                    else
                    {
                        string[] tempS = new string[2];
                        allSetting = allSetting.Trim();
                        strSettingSplit = allSetting.Split('\n');
                        string[] members;

                        foreach (var sub in strSettingSplit)
                        {
                            sublist = new List<decimal>();
                            tempS = sub.Split('=');
                            if (String.Compare(tempS[0], "WEDNG") == 0)
                            {
                                members = tempS[1].Split(',');
                                for (int i = 0; i < members.Length; i++)
                                    sublist.Add(decimal.Parse(members[i]));
                                tmpDailyPoints.Add(sublist);
                            }
                            else if (String.Compare(tempS[0], "THUDT") == 0)
                            {
                                members = tempS[1].Split(',');
                                for (int i = 0; i < members.Length; i++)
                                    sublist.Add(decimal.Parse(members[i]));
                                tmpDailyPoints.Add(sublist);
                            }
                            else if (String.Compare(tempS[0], "THUNG") == 0)
                            {
                                members = tempS[1].Split(',');
                                for (int i = 0; i < members.Length; i++)
                                    sublist.Add(decimal.Parse(members[i]));
                                tmpDailyPoints.Add(sublist);
                            }
                            else if (String.Compare(tempS[0], "FRIDT") == 0)
                            {
                                members = tempS[1].Split(',');
                                for (int i = 0; i < members.Length; i++)
                                    sublist.Add(decimal.Parse(members[i]));
                                tmpDailyPoints.Add(sublist);
                            }
                            else if (String.Compare(tempS[0], "FRING") == 0)
                            {
                                members = tempS[1].Split(',');
                                for (int i = 0; i < members.Length; i++)
                                    sublist.Add(decimal.Parse(members[i]));
                                tmpDailyPoints.Add(sublist);
                            }
                            else if (String.Compare(tempS[0], "MONDT") == 0)
                            {
                                members = tempS[1].Split(',');
                                for (int i = 0; i < members.Length; i++)
                                    sublist.Add(decimal.Parse(members[i]));
                                tmpDailyPoints.Add(sublist);
                            }
                            else if (String.Compare(tempS[0], "MONNG") == 0)
                            {
                                members = tempS[1].Split(',');
                                for (int i = 0; i < members.Length; i++)
                                    sublist.Add(decimal.Parse(members[i]));
                                tmpDailyPoints.Add(sublist);
                            }
                            else if (String.Compare(tempS[0], "TUEDT") == 0)
                            {
                                members = tempS[1].Split(',');
                                for (int i = 0; i < members.Length; i++)
                                    sublist.Add(decimal.Parse(members[i]));
                                tmpDailyPoints.Add(sublist);
                            }
                            else if (String.Compare(tempS[0], "TUENG") == 0)
                            {
                                members = tempS[1].Split(',');
                                for (int i = 0; i < members.Length; i++)
                                    sublist.Add(decimal.Parse(members[i]));
                                tmpDailyPoints.Add(sublist);
                            }
                            else if (String.Compare(tempS[0], "WEDDT") == 0)
                            {
                                members = tempS[1].Split(',');
                                for (int i = 0; i < members.Length; i++)
                                    sublist.Add(decimal.Parse(members[i]));
                                tmpDailyPoints.Add(sublist);
                            }
                            else if (String.Compare(tempS[0], "CallDownLim") == 0)
                            {
                                CallDownLim = decimal.Parse(tempS[1]);
                            }
                            else if (String.Compare(tempS[0], "CallUpLim") == 0)
                            {
                                CallUpLim = decimal.Parse(tempS[1]);
                            }
                            else if (String.Compare(tempS[0], "PutDownLim") == 0)
                            {
                                PutDownLim = decimal.Parse(tempS[1]);
                            }
                            else if (String.Compare(tempS[0], "PutUpLim") == 0)
                            {
                                PutUpLim = decimal.Parse(tempS[1]);
                            }
                            else if (String.Compare(tempS[0], "UpSafeDis") == 0)
                            {
                                UpSafeDis = decimal.Parse(tempS[1]);
                            }
                            else if (String.Compare(tempS[0], "DownSafeDis") == 0)
                            {
                                DownSafeDis = decimal.Parse(tempS[1]);
                            }
                            else if (String.Compare(tempS[0], "TakeProfitPointsSum") == 0)
                            {
                                TakeProfitPointsSum = decimal.Parse(tempS[1]);
                            }
                            else if (String.Compare(tempS[0], "TakeProfitAcc") == 0)
                            {
                                TakeProfitAcc = decimal.Parse(tempS[1]);
                            }
                            else if (String.Compare(tempS[0], "TakeProfitLots") == 0)
                            {
                                TakeProfitLots = ushort.Parse(tempS[1]);
                            }
                            else if (String.Compare(tempS[0], "TakeProfitGap") == 0)
                            {
                                TakeProfitGap = decimal.Parse(tempS[1]);
                            }
                            else if (String.Compare(tempS[0], "StopLossMagn") == 0)
                            {
                                StopLossMagn = decimal.Parse(tempS[1]);
                            }
                            else if (String.Compare(tempS[0], "StopLossLots") == 0)
                            {
                                StopLossLots = ushort.Parse(tempS[1]);
                            }
                            else if (String.Compare(tempS[0], "StopLossGap") == 0)
                            {
                                StopLossGap = decimal.Parse(tempS[1]);
                            }
                            else if (String.Compare(tempS[0], "EachOrderLots") == 0)
                            {
                                EachOrderLots = ushort.Parse(tempS[1]);
                            }
                            else if (String.Compare(tempS[0], "EachOrderGap") == 0)
                            {
                                EachOrderGap = decimal.Parse(tempS[1]);
                            }
                            else if (String.Compare(tempS[0], "OrderPriceShift") == 0)
                            {
                                OrderPriceShift = decimal.Parse(tempS[1]);
                            }
                            else if (String.Compare(tempS[0], "ProfitStopCountMin") == 0)
                            {
                                ProfitStopCountMin = decimal.Parse(tempS[1]);
                            }
                            else if (String.Compare(tempS[0], "ProfitStopSuccessRound") == 0)
                            {
                                ProfitStopSuccessRound = decimal.Parse(tempS[1]);
                            }
                            else if (String.Compare(tempS[0], "ProfitStopFailRound") == 0)
                            {
                                ProfitStopFailRound = decimal.Parse(tempS[1]);
                            }
                            else if (String.Compare(tempS[0], "EmergencyCountMin") == 0)
                            {
                                EmergencyCountMin = decimal.Parse(tempS[1]);
                            }
                            else if (String.Compare(tempS[0], "EmergencySuccessRound") == 0)
                            {
                                EmergencySuccessRound = decimal.Parse(tempS[1]);
                            }
                            else if (String.Compare(tempS[0], "EmergencyFailRound") == 0)
                            {
                                EmergencyFailRound = decimal.Parse(tempS[1]);
                            }
                            else if (String.Compare(tempS[0], "AtleastMargin") == 0)
                            {
                                AtleastMargin = decimal.Parse(tempS[1]);
                            }
                        }
                        DailyPoints = tmpDailyPoints;
                    }
                }
            }
            catch (IOException ex)
            {
                opLog.txtLog(
                    ex.GetType().Name+": The write operation could not " +
                    "be performed because the specified " +
                    "part of the file is locked.");
            }

        }

        //計算依設定每日的累計配額
        public decimal[] cacDayRentSumAndDaily()
        {
            DateTime NowDate = DateTime.Now;
            int periodOfWeek = 0;
            decimal[] dayRent = new decimal[4];//dayRent[0]: call每日累積配額 dayRent[1]: put每日累積配額 dayRent[2]: call今日配額 dayRent[3]: put今日配額

            switch (Convert.ToInt32(NowDate.DayOfWeek.ToString("d")))
            {
                case 1:
                    if (DateCalculateExtensions.IsLegalTime(NowDate, "000000-050000")) periodOfWeek = (int)PERIOD_LIST.MONDT;
                    else if (DateCalculateExtensions.IsLegalTime(NowDate, "050001-143000")) periodOfWeek = (int)PERIOD_LIST.MONDT;
                    else if (DateCalculateExtensions.IsLegalTime(NowDate, "143001-235959")) periodOfWeek = (int)PERIOD_LIST.MONNG;
                    break;
                case 2:
                    if (DateCalculateExtensions.IsLegalTime(NowDate, "000000-050000")) periodOfWeek = (int)PERIOD_LIST.MONNG;
                    else if (DateCalculateExtensions.IsLegalTime(NowDate, "050001-143000")) periodOfWeek = (int)PERIOD_LIST.TUEDT;
                    else if (DateCalculateExtensions.IsLegalTime(NowDate, "143001-235959")) periodOfWeek = (int)PERIOD_LIST.TUENG;
                    break;
                case 3:
                    if (DateCalculateExtensions.IsLegalTime(NowDate, "000000-050000")) periodOfWeek = (int)PERIOD_LIST.TUENG;
                    else if (DateCalculateExtensions.IsLegalTime(NowDate, "050001-143000")) periodOfWeek = (int)PERIOD_LIST.WEDDT;
                    else if (DateCalculateExtensions.IsLegalTime(NowDate, "143001-235959")) periodOfWeek = (int)PERIOD_LIST.WEDNG;
                    break;
                case 4:
                    if (DateCalculateExtensions.IsLegalTime(NowDate, "000000-050000")) periodOfWeek = (int)PERIOD_LIST.WEDNG;
                    else if (DateCalculateExtensions.IsLegalTime(NowDate, "050001-143000")) periodOfWeek = (int)PERIOD_LIST.THUDT;
                    else if (DateCalculateExtensions.IsLegalTime(NowDate, "143001-235959")) periodOfWeek = (int)PERIOD_LIST.THUNG;
                    break;
                case 5:
                    if (DateCalculateExtensions.IsLegalTime(NowDate, "000000-050000")) periodOfWeek = (int)PERIOD_LIST.THUNG;
                    else if (DateCalculateExtensions.IsLegalTime(NowDate, "050001-143000")) periodOfWeek = (int)PERIOD_LIST.FRIDT;
                    else if (DateCalculateExtensions.IsLegalTime(NowDate, "143001-235959")) periodOfWeek = (int)PERIOD_LIST.FRING;
                    break;
                case 6:
                    if (DateCalculateExtensions.IsLegalTime(NowDate, "000000-050000")) periodOfWeek = (int)PERIOD_LIST.FRING;
                    break;
                default:
                    break;
            }

            for (int i = 0; i <= periodOfWeek; i++)
            {
                dayRent[0] += DailyPoints[i][0];
                dayRent[1] += DailyPoints[i][1];
            }
            dayRent[2] = DailyPoints[periodOfWeek][0];
            dayRent[3] = DailyPoints[periodOfWeek][1];

            //opLog.txtLog("dayRent[0]: " + dayRent[0]);
            //opLog.txtLog("dayRent[1]: " + dayRent[1]);
            //opLog.txtLog("dayRent[2]: " + dayRent[2]);
            //opLog.txtLog("dayRent[3]: " + dayRent[3]);

            //MyLog("dayRent[0]: " + dayRent[0]);
            //MyLog("dayRent[1]: " + dayRent[1]);
            //MyLog("dayRent[2]: " + dayRent[2]);
            //MyLog("dayRent[3]: " + dayRent[3]);

            return dayRent;
        }
    }

    class TodayOrderStatus
    {
        public decimal todayRentCall;
        public decimal todayRentPut;
        public bool OrderCallDone { get; set; }
        public bool OrderPutDone { get; set; }
        public TodayOrderStatus()
        {
            Read();
        }
        public decimal TodayRentPut
        {
            get
            {
                return todayRentPut;
            }
            set
            {
                if (value != todayRentPut)
                {
                    todayRentPut = value;
                    opLog.txtLog("todayRentPut changed!");
                    Save();
                }
            }
        }
        public decimal TodayRentCall
        {
            get
            {
                return todayRentCall;
            }
            set
            {
                if (value != todayRentCall)
                {
                    todayRentCall = value;
                    opLog.txtLog("todayRentCall changed!");
                    Save();                
                }
            }
        }
        public void cacTodayRent(decimal callSum, decimal putSum, decimal[] dayRent, decimal[] autoOrderPoints)
        {
            decimal[] dayRentSum = new decimal[] { dayRent[0], dayRent[1] };//至今日配額[0]:call [1]:put
            decimal[] todayRent = new decimal[] { dayRent[2], dayRent[3] };//今日配額[2]:call [3]:put

            if (callSum < dayRentSum[0])                     //庫存的call總點數<至今日的累積設定點數 && call今日配額未完成 && 今日配額尚未計算過
            {
                if (autoOrderPoints[0] >= todayRent[0])//如果今日由程式下單的點數已經超過今日設定點數，則不再下單
                {
                    todayRent[0] = 0;
                }
                else if (dayRentSum[0] - callSum >= todayRent[0]) todayRent[0] = todayRent[0];                                               //至今日的累積設定點數 - 庫存的call總點數 >= 今日設定點數 => 今日配額 = 今日設定點數
                else if (dayRentSum[0] - callSum < todayRent[0] && dayRentSum[0] - callSum > 0) todayRent[0] = dayRentSum[0] - callSum; //至今日的累積設定點數 - 庫存的call總點數 < 今日設定點數 && 至今日的累積設定點數 - 庫存的call總點數 > 0 => 今日配額 = 至今日的累積設定點數 - 庫存的call總點數
                else                                                                                                                    //今日配額 = 0; 並且今日Call的配額改完已完成
                {
                    todayRent[0] = 0;
                }
                todayRentCall = decimal.Round(todayRent[0], 2);
            }
            else todayRentCall = 0;
            if (putSum < dayRentSum[1]) //同上
            {
                if (autoOrderPoints[1] >= todayRent[1])//如果今日由程式下單的點數已經超過今日設定點數，則不再下單
                {
                    todayRent[1] = 0;
                }
                else if (dayRentSum[1] - putSum >= todayRent[1]) todayRent[1] = todayRent[1];
                else if (dayRentSum[1] - putSum < todayRent[1] && dayRentSum[1] - putSum > 0) todayRent[1] = dayRentSum[1] - putSum;
                else todayRent[1] = 0;
                todayRentPut = decimal.Round(todayRent[1], 2);
                //opLog.txtLog("dayRent[0]: " + dayRent[0]);
                //opLog.txtLog("dayRent[1]: " + dayRent[1]);
                //opLog.txtLog("dayRent[2]: " + dayRent[2]);
                //opLog.txtLog("dayRent[3]: " + dayRent[3]);
                //opLog.txtLog("todayOrderRecord.cacTodayAutoOrderPoints()[0]: " + autoOrderPoints[0]);
                //opLog.txtLog("todayOrderRecord.cacTodayAutoOrderPoints()[1]: " + autoOrderPoints[1]);
                //opLog.txtLog("todayOrderStatus.TodayRentCall: " + TodayRentCall);
                //opLog.txtLog("todayOrderStatus.TodayRentPut: " + TodayRentPut);
            }
            else todayRentPut = 0;
        }
        public void Read()
        {
            string dir = Application.StartupPath + @"\Order\";
            DateTime NowDate = DateTime.Now;
            string dayTimeOrNight = "0"; // 0: daytime 1:night

            switch (Convert.ToInt32(NowDate.DayOfWeek.ToString("d")))
            {
                case 1:
                    if (DateCalculateExtensions.IsLegalTime(NowDate, "000000-050000")) dayTimeOrNight = "0";
                    else if (DateCalculateExtensions.IsLegalTime(NowDate, "050001-143000")) dayTimeOrNight = "0";
                    else if (DateCalculateExtensions.IsLegalTime(NowDate, "143001-235959")) dayTimeOrNight = "1";
                    break;
                case 2:
                    if (DateCalculateExtensions.IsLegalTime(NowDate, "000000-050000"))
                    {
                        NowDate = NowDate.AddDays(-1); //夜盤超過12點還是要算前一天的夜盤
                        dayTimeOrNight = "1";
                    }
                    else if (DateCalculateExtensions.IsLegalTime(NowDate, "050001-143000")) dayTimeOrNight = "0";
                    else if (DateCalculateExtensions.IsLegalTime(NowDate, "143001-235959")) dayTimeOrNight = "1";
                    break;
                case 3:
                    if (DateCalculateExtensions.IsLegalTime(NowDate, "000000-050000"))
                    {
                        NowDate = NowDate.AddDays(-1); //夜盤超過12點還是要算前一天的夜盤
                        dayTimeOrNight = "1";
                    }
                    else if (DateCalculateExtensions.IsLegalTime(NowDate, "050001-143000")) dayTimeOrNight = "0";
                    else if (DateCalculateExtensions.IsLegalTime(NowDate, "143001-235959")) dayTimeOrNight = "1";
                    break;
                case 4:
                    if (DateCalculateExtensions.IsLegalTime(NowDate, "000000-050000"))
                    {
                        NowDate = NowDate.AddDays(-1); //夜盤超過12點還是要算前一天的夜盤
                        dayTimeOrNight = "1";
                    }
                    else if (DateCalculateExtensions.IsLegalTime(NowDate, "050001-143000")) dayTimeOrNight = "0";
                    else if (DateCalculateExtensions.IsLegalTime(NowDate, "143001-235959")) dayTimeOrNight = "1";
                    break;
                case 5:
                    if (DateCalculateExtensions.IsLegalTime(NowDate, "000000-050000"))
                    {
                        NowDate = NowDate.AddDays(-1); //夜盤超過12點還是要算前一天的夜盤
                        dayTimeOrNight = "1";
                    }
                    else if (DateCalculateExtensions.IsLegalTime(NowDate, "050001-143000")) dayTimeOrNight = "0";
                    else if (DateCalculateExtensions.IsLegalTime(NowDate, "143001-235959")) dayTimeOrNight = "1";
                    break;
                case 6:
                    if (DateCalculateExtensions.IsLegalTime(NowDate, "000000-050000"))
                    {
                        NowDate = NowDate.AddDays(-1); //夜盤超過12點還是要算前一天的夜盤
                        dayTimeOrNight = "1";
                    }
                    break;
                default:
                    break;
            }

            string fileName = dir + NowDate.ToString("yyyyMMdd") + dayTimeOrNight + ".txt";
            string allStatus = "";
            string[] strStatusSplit;
            decimal file_TodayRentCall = -1;
            decimal file_TodayRentPut = -1;
            bool file_OrderCallDone = false;
            bool file_OrderPutDone = false;
            try
            {
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                if (!File.Exists(fileName))
                {
                    File.Create(fileName).Close();
                    string temp = "";
                    temp += "TodayRentCall=" + file_TodayRentCall.ToString();
                    temp += "\n";
                    temp += "TodayRentPut=" + file_TodayRentPut.ToString();
                    temp += "\n";
                    temp += "OrderCallDone=" + file_OrderCallDone.ToString();
                    temp += "\n";
                    temp += "OrderPutDone=" + file_OrderPutDone.ToString();
                    temp += "\n";
                    File.WriteAllText(fileName, temp);
                }
                else
                {
                    allStatus = System.IO.File.ReadAllText(fileName);
                    if (string.IsNullOrEmpty(allStatus))
                    {
                        opLog.txtLog("The allStatus is null."
                            + Environment.NewLine
                            + "Please fill this file with allStatus info.");
                    }
                    else
                    {
                        string[] tempS = new string[2];
                        allStatus = allStatus.Trim();
                        strStatusSplit = allStatus.Split('\n');
                        if (strStatusSplit != null)
                        {
                            foreach (var sub in strStatusSplit)
                            {
                                tempS = sub.Split('=');
                                if (String.Compare(tempS[0], "TodayRentCall") == 0)
                                    file_TodayRentCall = Convert.ToDecimal(tempS[1]);
                                if (String.Compare(tempS[0], "TodayRentPut") == 0)
                                    file_TodayRentPut = Convert.ToDecimal(tempS[1]);
                                if (String.Compare(tempS[0], "OrderCallDone") == 0)
                                    file_OrderCallDone = Convert.ToBoolean(tempS[1]);
                                if (String.Compare(tempS[0], "OrderPutDone") == 0)
                                    file_OrderPutDone = Convert.ToBoolean(tempS[1]);
                            }
                            todayRentCall = file_TodayRentCall;
                            todayRentPut = file_TodayRentPut;
                            OrderCallDone = file_OrderCallDone;
                            OrderPutDone = file_OrderPutDone;
                        }
                        //string temp = "";
                        //temp += "TodayRentCall=" + TodayRentCall.ToString();
                        //temp += "\n";
                        //temp += "TodayRentPut=" + TodayRentPut.ToString();
                        //temp += "\n";
                        //temp += "OrderCallDone=" + OrderCallDone.ToString();
                        //temp += "\n";
                        //temp += "OrderPutDone=" + OrderPutDone.ToString();
                        //temp += "\n";
                        //File.WriteAllText(fileName, temp);
                    }
                }
            }
            catch (IOException ex)
            {
                opLog.txtLog(
                    ex.GetType().Name+": The write operation could not " +
                    "be performed because the specified " +
                    "part of the file is locked.");
            }
        }
        public void Save()
        {
            string dir = Application.StartupPath + @"\Order\";
            DateTime NowDate = DateTime.Now;
            string dayTimeOrNight = "0"; // 0: daytime 1:night

            switch (Convert.ToInt32(NowDate.DayOfWeek.ToString("d")))
            {
                case 1:
                    if (DateCalculateExtensions.IsLegalTime(NowDate, "000000-050000")) dayTimeOrNight = "0";
                    else if (DateCalculateExtensions.IsLegalTime(NowDate, "050001-143000")) dayTimeOrNight = "0";
                    else if (DateCalculateExtensions.IsLegalTime(NowDate, "143001-235959")) dayTimeOrNight = "1";
                    break;
                case 2:
                    if (DateCalculateExtensions.IsLegalTime(NowDate, "000000-050000"))
                    {
                        NowDate = NowDate.AddDays(-1); //夜盤超過12點還是要算前一天的夜盤
                        dayTimeOrNight = "1";
                    }
                    else if (DateCalculateExtensions.IsLegalTime(NowDate, "050001-143000")) dayTimeOrNight = "0";
                    else if (DateCalculateExtensions.IsLegalTime(NowDate, "143001-235959")) dayTimeOrNight = "1";
                    break;
                case 3:
                    if (DateCalculateExtensions.IsLegalTime(NowDate, "000000-050000"))
                    {
                        NowDate = NowDate.AddDays(-1); //夜盤超過12點還是要算前一天的夜盤
                        dayTimeOrNight = "1";
                    }
                    else if (DateCalculateExtensions.IsLegalTime(NowDate, "050001-143000")) dayTimeOrNight = "0";
                    else if (DateCalculateExtensions.IsLegalTime(NowDate, "143001-235959")) dayTimeOrNight = "1";
                    break;
                case 4:
                    if (DateCalculateExtensions.IsLegalTime(NowDate, "000000-050000"))
                    {
                        NowDate = NowDate.AddDays(-1); //夜盤超過12點還是要算前一天的夜盤
                        dayTimeOrNight = "1";
                    }
                    else if (DateCalculateExtensions.IsLegalTime(NowDate, "050001-143000")) dayTimeOrNight = "0";
                    else if (DateCalculateExtensions.IsLegalTime(NowDate, "143001-235959")) dayTimeOrNight = "1";
                    break;
                case 5:
                    if (DateCalculateExtensions.IsLegalTime(NowDate, "000000-050000"))
                    {
                        NowDate = NowDate.AddDays(-1); //夜盤超過12點還是要算前一天的夜盤
                        dayTimeOrNight = "1";
                    }
                    else if (DateCalculateExtensions.IsLegalTime(NowDate, "050001-143000")) dayTimeOrNight = "0";
                    else if (DateCalculateExtensions.IsLegalTime(NowDate, "143001-235959")) dayTimeOrNight = "1";
                    break;
                case 6:
                    if (DateCalculateExtensions.IsLegalTime(NowDate, "000000-050000"))
                    {
                        NowDate = NowDate.AddDays(-1); //夜盤超過12點還是要算前一天的夜盤
                        dayTimeOrNight = "1";
                    }
                    break;
                default:
                    break;
            }

            string fileName = dir + NowDate.ToString("yyyyMMdd") + dayTimeOrNight + ".txt";
            decimal file_TodayRentCall = -1;
            decimal file_TodayRentPut = -1;
            bool file_OrderCallDone = false;
            bool file_OrderPutDone = false;


            try
            {
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                if (!File.Exists(fileName))
                {
                    File.Create(fileName).Close();
                    string temp = "";
                    temp += "TodayRentCall=" + file_TodayRentCall.ToString();
                    temp += "\n";
                    temp += "TodayRentPut=" + file_TodayRentPut.ToString();
                    temp += "\n";
                    temp += "OrderCallDone=" + file_OrderCallDone.ToString();
                    temp += "\n";
                    temp += "OrderPutDone=" + file_OrderPutDone.ToString();
                    temp += "\n";
                    File.WriteAllText(fileName, temp);
                }
                else
                {
                    string temp = "";
                    temp += "TodayRentCall=" + TodayRentCall.ToString();
                    temp += "\n";
                    temp += "TodayRentPut=" + TodayRentPut.ToString();
                    temp += "\n";
                    temp += "OrderCallDone=" + OrderCallDone.ToString();
                    temp += "\n";
                    temp += "OrderPutDone=" + OrderPutDone.ToString();
                    temp += "\n";
                    File.WriteAllText(fileName, temp);
                }
            }
            catch (IOException ex)
            {
                opLog.txtLog(
                    ex.GetType().Name + ": The write operation could not " +
                    "be performed because the specified " +
                    "part of the file is locked.");
            }

        }
    }

    class DataRecord
    {
        public decimal LastMXF { get; set; } = 0;
        public void saveMXF(decimal lastMXF)
        {
            string dir = Application.StartupPath + @"\DataRecord\";
            string fileName = dir + "datarecord" + ".txt";
            string allElement = "";

            try
            {
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                if (!File.Exists(fileName))
                {
                    File.Create(fileName).Close();
                    string temp = "";
                    temp += "LastMXF=" + lastMXF.ToString();
                    temp += "\n";
                    File.WriteAllText(fileName, temp);
                }
                else
                {
                    allElement = System.IO.File.ReadAllText(fileName);
                    if (string.IsNullOrEmpty(allElement))
                    {
                        opLog.txtLog("The allStatus is null."
                            + Environment.NewLine
                            + "Please fill this file with allStatus info.");
                    }
                    else
                    {
                        string temp = "";
                        temp += "LastMXF=" + lastMXF.ToString();
                        temp += "\n";
                        File.WriteAllText(fileName, temp);
                    }
                }
            }
            catch (IOException ex)
            {
                opLog.txtLog(
                    ex.GetType().Name + ": The write operation could not " +
                    "be performed because the specified " +
                    "part of the file is locked.");
            }
        }
        public decimal readMXF()
        {
            string dir = Application.StartupPath + @"\DataRecord\";
            string fileName = dir + "datarecord" + ".txt";
            string allElement = "";
            string[] allElementSplit;
            string[] tempS = new string[2];

            try
            {
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                if (!File.Exists(fileName))
                {
                    File.Create(fileName).Close();
                    string temp = "";
                    temp += "LastMXF=" + "18000";
                    temp += "\n";
                    File.WriteAllText(fileName, temp);
                    allElement = System.IO.File.ReadAllText(fileName);
                }
                else
                {
                    allElement = System.IO.File.ReadAllText(fileName);
                    if (string.IsNullOrEmpty(allElement))
                    {
                        opLog.txtLog("The allStatus is null."
                            + Environment.NewLine
                            + "Please fill this file with allStatus info.");
                    }
                }
            }
            catch (IOException ex)
            {
                opLog.txtLog(
                    ex.GetType().Name + ": The write operation could not " +
                    "be performed because the specified " +
                    "part of the file is locked.");
            }
            allElement = allElement.Trim();
            allElementSplit = allElement.Split('\n');

            if (allElementSplit != null)
            {
                foreach (var sub in allElementSplit)
                {
                    tempS = sub.Split('=');
                    if (String.Compare(tempS[0], "LastMXF") == 0)
                        LastMXF = Convert.ToDecimal(tempS[1]);
                }
            }

            return LastMXF;
        }
    }

    class TodayOrderRecord
    {
        public Dictionary<string, List<string>> recordOrders = new Dictionary<string, List<string>>();
        public decimal[] autoOrderPoints = new decimal[2]; // autoOrderPoints[0]: call autoOrderPoints[1]: put
                                                           //public string ORDERSTATUS { get; set; }
                                                           //public string CNT { get; set; }
                                                           //public string ORDERNO { get; set; }
                                                           //public string TRADEDATE { get; set; }
                                                           //public string REPORTTIME { get; set; }
                                                           //public string SYMBOL { get; set; }
                                                           //public string SIDE { get; set; }
                                                           //public string DEALPRICE { get; set; }
                                                           //public string CUMQTY { get; set; }
                                                           //public string SYMBOL1 { get; set; }
                                                           //public string DEALPRICE1 { get; set; }
                                                           //public string QTY1 { get; set; }
                                                           //public string BS1 { get; set; }
                                                           //public string SYMBOL2 { get; set; }
                                                           //public string DEALPRICE2 { get; set; }
                                                           //public string QTY2 { get; set; }
                                                           //public string BS2 { get; set; }
        public TodayOrderRecord()
        {
            loadTxtOrder();
        }
        private void loadTxtOrder()
        {
            string dir = Application.StartupPath + @"\OrderRecord\";

            DateTime NowDate = DateTime.Now;
            string dayTimeOrNight = "0"; // 0: daytime 1:night

            switch (Convert.ToInt32(NowDate.DayOfWeek.ToString("d")))
            {
                case 1:
                    if (DateCalculateExtensions.IsLegalTime(NowDate, "000000-050000")) dayTimeOrNight = "0";
                    else if (DateCalculateExtensions.IsLegalTime(NowDate, "050001-143000")) dayTimeOrNight = "0";
                    else if (DateCalculateExtensions.IsLegalTime(NowDate, "143001-235959")) dayTimeOrNight = "1";
                    break;
                case 2:
                    if (DateCalculateExtensions.IsLegalTime(NowDate, "000000-050000"))
                    {
                        NowDate = NowDate.AddDays(-1); //夜盤超過12點還是要算前一天的夜盤
                        dayTimeOrNight = "1";
                    }
                    else if (DateCalculateExtensions.IsLegalTime(NowDate, "050001-143000")) dayTimeOrNight = "0";
                    else if (DateCalculateExtensions.IsLegalTime(NowDate, "143001-235959")) dayTimeOrNight = "1";
                    break;
                case 3:
                    if (DateCalculateExtensions.IsLegalTime(NowDate, "000000-050000"))
                    {
                        NowDate = NowDate.AddDays(-1); //夜盤超過12點還是要算前一天的夜盤
                        dayTimeOrNight = "1";
                    }
                    else if (DateCalculateExtensions.IsLegalTime(NowDate, "050001-143000")) dayTimeOrNight = "0";
                    else if (DateCalculateExtensions.IsLegalTime(NowDate, "143001-235959")) dayTimeOrNight = "1";
                    break;
                case 4:
                    if (DateCalculateExtensions.IsLegalTime(NowDate, "000000-050000"))
                    {
                        NowDate = NowDate.AddDays(-1); //夜盤超過12點還是要算前一天的夜盤
                        dayTimeOrNight = "1";
                    }
                    else if (DateCalculateExtensions.IsLegalTime(NowDate, "050001-143000")) dayTimeOrNight = "0";
                    else if (DateCalculateExtensions.IsLegalTime(NowDate, "143001-235959")) dayTimeOrNight = "1";
                    break;
                case 5:
                    if (DateCalculateExtensions.IsLegalTime(NowDate, "000000-050000"))
                    {
                        NowDate = NowDate.AddDays(-1); //夜盤超過12點還是要算前一天的夜盤
                        dayTimeOrNight = "1";
                    }
                    else if (DateCalculateExtensions.IsLegalTime(NowDate, "050001-143000")) dayTimeOrNight = "0";
                    else if (DateCalculateExtensions.IsLegalTime(NowDate, "143001-235959")) dayTimeOrNight = "1";
                    break;
                case 6:
                    if (DateCalculateExtensions.IsLegalTime(NowDate, "000000-050000"))
                    {
                        NowDate = NowDate.AddDays(-1); //夜盤超過12點還是要算前一天的夜盤
                        dayTimeOrNight = "1";
                    }
                    break;
                default:
                    break;
            }
            opLog.txtLog("loadtxt " + NowDate.ToString("yyyyMMdd"));
            string fileName = dir + NowDate.ToString("yyyyMMdd") + dayTimeOrNight + ".txt";
            string allOrderRecords = "";
            string[] allOrderRecordsSplit;
            try
            {
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                if (!File.Exists(fileName)) File.Create(fileName).Close();
                else
                {
                    string[] tempS = new string[2];
                    string[] orderContent;

                    allOrderRecords = System.IO.File.ReadAllText(fileName);
                    if (string.IsNullOrEmpty(allOrderRecords))
                    {
                        opLog.txtLog("The allStatus is null."
                            + Environment.NewLine
                            + "Please fill this file with allStatus info.");
                    }
                    else
                    {
                        allOrderRecords = allOrderRecords.Trim();
                        allOrderRecordsSplit = allOrderRecords.Split('\n');
                        if (allOrderRecordsSplit != null)
                        {
                            foreach (var sub in allOrderRecordsSplit)
                            {
                                tempS = sub.Split('=');
                                orderContent = tempS[1].Split(',');
                                if (orderContent != null)
                                    if (!recordOrders.ContainsKey(tempS[0]))
                                        recordOrders.Add(tempS[0], orderContent.ToList());
                            }
                        }
                    }
                }
            }
            catch (IOException ex)
            {
                opLog.txtLog(
                    ex.GetType().Name+": The write operation could not " +
                    "be performed because the specified " +
                    "part of the file is locked.");
            }
        }
        public void AddNewRecord(List<string> newOrder)
        {
            string dir = Application.StartupPath + @"\OrderRecord\";

            DateTime NowDate = DateTime.Now;
            string dayTimeOrNight = "0"; // 0: daytime 1:night

            switch (Convert.ToInt32(NowDate.DayOfWeek.ToString("d")))
            {
                case 1:
                    if (DateCalculateExtensions.IsLegalTime(NowDate, "000000-050000")) dayTimeOrNight = "0";
                    else if (DateCalculateExtensions.IsLegalTime(NowDate, "050001-143000")) dayTimeOrNight = "0";
                    else if (DateCalculateExtensions.IsLegalTime(NowDate, "143001-235959")) dayTimeOrNight = "1";
                    break;
                case 2:
                    if (DateCalculateExtensions.IsLegalTime(NowDate, "000000-050000"))
                    {
                        NowDate = NowDate.AddDays(-1); //夜盤超過12點還是要算前一天的夜盤
                        dayTimeOrNight = "1";
                    }
                    else if (DateCalculateExtensions.IsLegalTime(NowDate, "050001-143000")) dayTimeOrNight = "0";
                    else if (DateCalculateExtensions.IsLegalTime(NowDate, "143001-235959")) dayTimeOrNight = "1";
                    break;
                case 3:
                    if (DateCalculateExtensions.IsLegalTime(NowDate, "000000-050000"))
                    {
                        NowDate = NowDate.AddDays(-1); //夜盤超過12點還是要算前一天的夜盤
                        dayTimeOrNight = "1";
                    }
                    else if (DateCalculateExtensions.IsLegalTime(NowDate, "050001-143000")) dayTimeOrNight = "0";
                    else if (DateCalculateExtensions.IsLegalTime(NowDate, "143001-235959")) dayTimeOrNight = "1";
                    break;
                case 4:
                    if (DateCalculateExtensions.IsLegalTime(NowDate, "000000-050000"))
                    {
                        NowDate = NowDate.AddDays(-1); //夜盤超過12點還是要算前一天的夜盤
                        dayTimeOrNight = "1";
                    }
                    else if (DateCalculateExtensions.IsLegalTime(NowDate, "050001-143000")) dayTimeOrNight = "0";
                    else if (DateCalculateExtensions.IsLegalTime(NowDate, "143001-235959")) dayTimeOrNight = "1";
                    break;
                case 5:
                    if (DateCalculateExtensions.IsLegalTime(NowDate, "000000-050000"))
                    {
                        NowDate = NowDate.AddDays(-1); //夜盤超過12點還是要算前一天的夜盤
                        dayTimeOrNight = "1";
                    }
                    else if (DateCalculateExtensions.IsLegalTime(NowDate, "050001-143000")) dayTimeOrNight = "0";
                    else if (DateCalculateExtensions.IsLegalTime(NowDate, "143001-235959")) dayTimeOrNight = "1";
                    break;
                case 6:
                    if (DateCalculateExtensions.IsLegalTime(NowDate, "000000-050000"))
                    {
                        NowDate = NowDate.AddDays(-1); //夜盤超過12點還是要算前一天的夜盤
                        dayTimeOrNight = "1";
                    }
                    break;
                default:
                    break;
            }
            opLog.txtLog("addnewtxt " + NowDate.ToString("yyyyMMdd"));
            string fileName = dir + NowDate.ToString("yyyyMMdd") + dayTimeOrNight + ".txt";
            try
            {
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                if (!File.Exists(fileName))
                {
                    opLog.txtLog("File not found then Creat!!");
                    if (!recordOrders.ContainsKey(newOrder[(int)DEALRPT_LIST.CNT]))
                    {
                        string temp = "";
                        temp += newOrder[(int)DEALRPT_LIST.CNT] + "=";
                        temp += newOrder[(int)DEALRPT_LIST.ORDERSTATUS];
                        for (int i = 1; i < newOrder.Count; i++)
                        {
                            temp += "," + newOrder[i];
                        }
                        temp += "\n";
                        using (StreamWriter sw = File.AppendText(fileName))
                        {
                            sw.Write(temp);
                        }
                    }
                }
                else
                {
                    if (!recordOrders.ContainsKey(newOrder[(int)DEALRPT_LIST.CNT]))
                    {
                        string temp = "";
                        temp += newOrder[(int)DEALRPT_LIST.CNT] + "=";
                        temp += newOrder[(int)DEALRPT_LIST.ORDERSTATUS];
                        for (int i = 1; i < newOrder.Count; i++)
                        {
                            temp += "," + newOrder[i];
                        }
                        temp += "\n";
                        using (StreamWriter sw = File.AppendText(fileName))
                        {
                            sw.Write(temp);
                        }
                    }
                }
            }
            catch (IOException ex)
            {
                opLog.txtLog(
                    ex.GetType().Name+": The write operation could not " +
                    "be performed because the specified " +
                    "part of the file is locked.");
            }
            loadTxtOrder();
        }
        public decimal[] cacTodayAutoOrderPoints()
        {
            autoOrderPoints[0] = 0;
            autoOrderPoints[1] = 0;
            foreach (string key in recordOrders.Keys)
            {
                string market = recordOrders[key][(int)DEALRPT_LIST.MARKET];
                string symbol = recordOrders[key][(int)DEALRPT_LIST.SYMBOL];
                string symbol1 = recordOrders[key][(int)DEALRPT_LIST.SYMBOL1];
                string symbol2 = recordOrders[key][(int)DEALRPT_LIST.SYMBOL2];
                string side = recordOrders[key][(int)DEALRPT_LIST.SIDE];
                string bs1 = recordOrders[key][(int)DEALRPT_LIST.BS1];
                string bs2 = recordOrders[key][(int)DEALRPT_LIST.BS2];
                string dealprice = recordOrders[key][(int)DEALRPT_LIST.DEALPRICE];
                string dealprice1 = recordOrders[key][(int)DEALRPT_LIST.DEALPRICE1];
                string dealprice2 = recordOrders[key][(int)DEALRPT_LIST.DEALPRICE2];
                decimal qty = Convert.ToDecimal(recordOrders[key][(int)DEALRPT_LIST.DEALQTY]);
                if (String.Compare(market, "2") == 0)//選擇權單式
                {
                    if (symbol[symbol.Length - 2] >= 'A' && symbol[symbol.Length - 2] <= 'L') //Call
                    {
                        if (side[0] == 'S') autoOrderPoints[0] += decimal.Parse(dealprice)* qty;
                        if (side[0] == 'B') autoOrderPoints[0] -= decimal.Parse(dealprice) * qty;
                    }
                    else if (symbol[symbol.Length - 2] >= 'M' && symbol[symbol.Length - 2] <= 'X') //Put
                    {
                        if (side[0] == 'S') autoOrderPoints[1] += decimal.Parse(dealprice) * qty;
                        if (side[0] == 'B') autoOrderPoints[1] -= decimal.Parse(dealprice) * qty;
                    }
                }
                else if (String.Compare(market, "3") == 0)//選擇權複式
                {
                    if (symbol1[symbol1.Length - 2] >= 'A' && symbol1[symbol1.Length - 2] <= 'L') //Call
                    {
                        if (bs1[0] == 'S') autoOrderPoints[0] += decimal.Parse(dealprice1) * qty;
                        if (bs1[0] == 'B') autoOrderPoints[0] -= decimal.Parse(dealprice1) * qty;
                    }
                    else if (symbol1[symbol1.Length - 2] >= 'M' && symbol1[symbol1.Length - 2] <= 'X') //Put
                    {
                        if (bs1[0] == 'S') autoOrderPoints[1] += decimal.Parse(dealprice1) * qty;
                        if (bs1[0] == 'B') autoOrderPoints[1] -= decimal.Parse(dealprice1) * qty;
                    }
                    if (symbol2[symbol2.Length - 2] >= 'A' && symbol2[symbol2.Length - 2] <= 'L') //Call
                    {
                        if (bs2[0] == 'S') autoOrderPoints[0] += decimal.Parse(dealprice2) * qty;
                        if (bs2[0] == 'B') autoOrderPoints[0] -= decimal.Parse(dealprice2) * qty;
                    }
                    else if (symbol2[symbol2.Length - 2] >= 'M' && symbol2[symbol2.Length - 2] <= 'X') //Put
                    {
                        if (bs2[0] == 'S') autoOrderPoints[1] += decimal.Parse(dealprice2) * qty;
                        if (bs2[0] == 'B') autoOrderPoints[1] -= decimal.Parse(dealprice2) * qty;
                    }
                }
            }
            return autoOrderPoints;
        }
    }

    public static class DateCalculateExtensions
    {
        /// <summary>
        /// 判斷一個時間是否位於指定的時間段內
        /// </summary>
        /// <param name="time_interval">時間區間string</param>
        /// <returns></returns>
        public static bool IsLegalTime(DateTime dt, string time_intervals)
        {
            //當前時間
            int time_now = dt.Hour * 10000 + dt.Minute * 100 + dt.Second;
            //查看各個時間區間
            string[] time_interval = time_intervals.Split(';');
            foreach (string time in time_interval)
            {
                //Null直接跳過
                if (string.IsNullOrWhiteSpace(time))
                {
                    continue;
                }
                //時間格式：六個數字-六個數字
                if (!Regex.IsMatch(time, "^[0-9]{6}-[0-9]{6}$"))
                {
                    opLog.txtLog(time+"： 錯誤的時間數據");
                }
                string timea = time.Substring(0, 6);
                string timeb = time.Substring(7, 6);
                int time_a, time_b;
                //嘗試轉化為整數
                if (!int.TryParse(timea, out time_a))
                {
                    opLog.txtLog(timea+"： 轉化為整數失敗");
                }
                if (!int.TryParse(timeb, out time_b))
                {
                    opLog.txtLog(timeb+"： 轉化為整數失敗");
                }
                //如果當前時間大於等於開始時間，小於結束時間，return true
                if (time_a <= time_now && time_now <= time_b)
                {
                    return true;
                }
            }
            //不在任一個區間範圍內，return false
            return false;
        }
    }

    //用來計算這個月合約的Symbol與計算指定日期是今年的第幾週(以週三為第一天)
    public class CalculateComID
    {
        public string[] PartialComIDMX = new string[4];   //[0]:MX2   [1]:A2   [2]:202201   [3]:202201W4 期交所使用
        public string[] PartialComIDTXF = new string[3];
        public string[] PartialComIDC = new string[4];
        public string[] PartialComIDP = new string[4];
        public char[] monthSymbol = new char[2];
        public char[] weekSymbol = new char[2];
        public int MonthProduct = 0;//1~12
        public string WeekProduct;//1.2.O.3.4.5
        public int WeekProductMonthShift = 0;//0 or 1
        public int WeekProductYearShift = 0;//0 or 1
        public int MonthOrderYear = 0;

        public CalculateComID(DateTime NowDate)
        {
            //string testDate = "02/09/2022"; 
            //DateTime DateObject = DateTime.Parse(testDate);
            //DateTime NowDate = DateObject;
            //DateTime NowDate = DateTime.Now;
            if (NowDate.DayOfWeek == DayOfWeek.Wednesday && DateCalculateExtensions.IsLegalTime(DateTime.Now, "000000-050000") )
                NowDate = NowDate.AddDays(-1);

            GregorianCalendar gc = new GregorianCalendar();

            DateTime startMonth = NowDate.AddDays(1 - NowDate.Day);
            DateTime startNextMonth = startMonth.AddMonths(1);

            int nowWeekOfMonth = GetWeekOfMonth(NowDate);
            int startNextMonthWeekOfYear = 0;
            int nowWeekOfYear = 0;

            nowWeekOfYear = gc.GetWeekOfYear(NowDate, CalendarWeekRule.FirstFullWeek, DayOfWeek.Wednesday);
            startNextMonthWeekOfYear = gc.GetWeekOfYear(startNextMonth, CalendarWeekRule.FirstFullWeek, DayOfWeek.Wednesday);

            int nextStartMonthWeekOfYear = GetWeekOfMonth(NowDate);

            if (NowDate.Month == 12 && nowWeekOfMonth >= 3)
            {
                MonthProduct = 1;//1月合約
                if (nowWeekOfYear == startNextMonthWeekOfYear)
                {
                    WeekProduct = "1";
                    WeekProductMonthShift = -11;
                    WeekProductYearShift = 1;
                }
                else if (nowWeekOfMonth == 3) WeekProduct = "O";
                else WeekProduct = nowWeekOfMonth.ToString();
                MonthOrderYear = NowDate.Year + 1;
                //opLog.txtLog("nowWeekOfMonth: " + nowWeekOfMonth + Environment.NewLine);
                //opLog.txtLog("Month: " + NowDate.Month + Environment.NewLine);
                //opLog.txtLog("DayOfWeek: " + NowDate.DayOfWeek + Environment.NewLine);
            }
            else if (nowWeekOfMonth >= 3)
            {
                MonthProduct = NowDate.Month + 1;
                if (nowWeekOfYear == startNextMonthWeekOfYear)
                {
                    WeekProduct = "1";
                    WeekProductMonthShift = 1;
                }
                else if (nowWeekOfMonth == 3) WeekProduct = "O";
                else WeekProduct = nowWeekOfMonth.ToString();
                MonthOrderYear = NowDate.Year;
                //opLog.txtLog("nowWeekOfMonth: " + nowWeekOfMonth + Environment.NewLine);
                //opLog.txtLog("Month: " + NowDate.Month + Environment.NewLine);
                //opLog.txtLog("DayOfWeek: " + NowDate.DayOfWeek + Environment.NewLine);
            }
            else
            {
                MonthProduct = NowDate.Month;
                if (nowWeekOfYear == startNextMonthWeekOfYear)
                {
                    WeekProduct = "1";
                    WeekProductMonthShift = 1;
                }
                else WeekProduct = nowWeekOfMonth.ToString();
                MonthOrderYear = NowDate.Year;
                //opLog.txtLog("nowWeekOfMonth: " + nowWeekOfMonth + Environment.NewLine);
                //opLog.txtLog("Month: " + NowDate.Month + Environment.NewLine);
                //opLog.txtLog("DayOfWeek: " + NowDate.DayOfWeek + Environment.NewLine);
            }
            //opLog.txtLog("nowWeekOfMonth: " + nowWeekOfMonth.ToString());
            //opLog.txtLog("WeekProduct: "+ WeekProduct);
            FindSymbol(MonthProduct, monthSymbol);
            FindSymbol(NowDate.Month + WeekProductMonthShift, weekSymbol);

            string taifexWeekTag = "";
            if (string.Compare(WeekProduct, "O") == 0) taifexWeekTag = "";
            else taifexWeekTag = "W" + WeekProduct;

            if (string.Compare(WeekProduct, "O") == 0) PartialComIDMX[0] = "MXF";
            else PartialComIDMX[0] = "MX" + WeekProduct;
            PartialComIDMX[1] = weekSymbol[0].ToString() + ((NowDate.Year + WeekProductYearShift) % 10).ToString();
            PartialComIDMX[2] = (NowDate.Year + WeekProductYearShift).ToString() + String.Format("{0:00}", NowDate.Month + WeekProductMonthShift);
            PartialComIDMX[3] = (NowDate.Year + WeekProductYearShift).ToString() + String.Format("{0:00}", NowDate.Month + WeekProductMonthShift) + taifexWeekTag;
            PartialComIDTXF[0] = "TXF";
            PartialComIDTXF[1] = monthSymbol[0].ToString() + (MonthOrderYear % 10).ToString();
            PartialComIDTXF[2] = MonthOrderYear.ToString() + MonthProduct.ToString();
            PartialComIDC[0] = "TX" + WeekProduct;
            PartialComIDC[1] = weekSymbol[0].ToString() + ((NowDate.Year + WeekProductYearShift) % 10).ToString();
            PartialComIDC[2] = (NowDate.Year + WeekProductYearShift).ToString() + String.Format("{0:00}", NowDate.Month + WeekProductMonthShift);  //{0:00}:只有一位時前面補零
            PartialComIDC[3] = (NowDate.Year + WeekProductYearShift).ToString() + String.Format("{0:00}", NowDate.Month + WeekProductMonthShift) + taifexWeekTag;
            PartialComIDP[0] = "TX" + WeekProduct;
            PartialComIDP[1] = weekSymbol[1].ToString() + ((NowDate.Year + WeekProductYearShift) % 10).ToString();
            PartialComIDP[2] = (NowDate.Year + WeekProductYearShift).ToString() + String.Format("{0:00}", NowDate.Month + WeekProductMonthShift);
            PartialComIDP[3] = (NowDate.Year + WeekProductYearShift).ToString() + String.Format("{0:00}", NowDate.Month + WeekProductMonthShift) + taifexWeekTag;
            //opLog.txtLog("PartialComIDP[0]: " + PartialComIDP[0].Length +" "+ PartialComIDP[0] + Environment.NewLine);
            //opLog.txtLog("PartialComIDP[1]: " + PartialComIDP[1].Length +" "+ PartialComIDMX[1] + Environment.NewLine);
            //opLog.txtLog("partialComIDP: " + PartialComIDP + Environment.NewLine);
        }

        public static void FindSymbol(int month, char[] monthSymbol)
        {
            //opLog.txtLog("cbContractMonth.SelectedIndex: " + CbContractMonth_SelectedIndex + Environment.NewLine);
            switch (month)
            {
                case (int)TWELVEMONTH_SEQ.JAN:
                    monthSymbol[0] = 'A';
                    monthSymbol[1] = 'M';
                    break;
                case (int)TWELVEMONTH_SEQ.FEB:
                    monthSymbol[0] = 'B';
                    monthSymbol[1] = 'N';
                    break;
                case (int)TWELVEMONTH_SEQ.MAR:
                    monthSymbol[0] = 'C';
                    monthSymbol[1] = 'O';
                    break;
                case (int)TWELVEMONTH_SEQ.APR:
                    monthSymbol[0] = 'D';
                    monthSymbol[1] = 'P';
                    break;
                case (int)TWELVEMONTH_SEQ.MAY:
                    monthSymbol[0] = 'E';
                    monthSymbol[1] = 'Q';
                    break;
                case (int)TWELVEMONTH_SEQ.JUN:
                    monthSymbol[0] = 'F';
                    monthSymbol[1] = 'R';
                    break;
                case (int)TWELVEMONTH_SEQ.JUL:
                    monthSymbol[0] = 'G';
                    monthSymbol[1] = 'S';
                    break;
                case (int)TWELVEMONTH_SEQ.AUG:
                    monthSymbol[0] = 'H';
                    monthSymbol[1] = 'T';
                    break;
                case (int)TWELVEMONTH_SEQ.SEP:
                    monthSymbol[0] = 'I';
                    monthSymbol[1] = 'U';
                    break;
                case (int)TWELVEMONTH_SEQ.OCT:
                    monthSymbol[0] = 'J';
                    monthSymbol[1] = 'V';
                    break;
                case (int)TWELVEMONTH_SEQ.NOV:
                    monthSymbol[0] = 'K';
                    monthSymbol[1] = 'W';
                    break;
                case (int)TWELVEMONTH_SEQ.DEC:
                    monthSymbol[0] = 'L';
                    monthSymbol[1] = 'X';
                    break;
            }
            //opLog.txtLog("opMonthSymbol[0]: " + OpMonthSymbol[0] + Environment.NewLine);
            //opLog.txtLog("opMonthSymbol[1]: " + OpMonthSymbol[1] + Environment.NewLine);
        }

        public int GetWeekOfMonth(DateTime dt)
        {
            int endLastMonthOfWeeks = 0, dtOfWeeks = 0;
            DateTime endLastMonth = dt.AddDays((-1) * dt.Day);//上月月末
                                                              //opLog.txtLog("endLastMonth: " + endLastMonth + Environment.NewLine);

            GregorianCalendar gc = new GregorianCalendar();
            endLastMonthOfWeeks = gc.GetWeekOfYear(endLastMonth, CalendarWeekRule.FirstFullWeek, DayOfWeek.Wednesday);
            //opLog.txtLog("endLastMonthOfWeeks: " + endLastMonthOfWeeks + Environment.NewLine);

            dtOfWeeks = gc.GetWeekOfYear(dt, CalendarWeekRule.FirstFullWeek, DayOfWeek.Wednesday);
            //opLog.txtLog("dtOfWeeks: " + dtOfWeeks + Environment.NewLine);

            //opLog.txtLog("endLastMonthOfWeeks: "+ endLastMonthOfWeeks);
            //opLog.txtLog("dtOfWeeks: "+ dtOfWeeks);

            //1月的時候無法跟去年12月做差值來計算第幾周
            if (dt.Month == 1 && gc.GetWeekOfYear(dt, CalendarWeekRule.FirstFullWeek, DayOfWeek.Wednesday) >= 1 && gc.GetWeekOfYear(dt, CalendarWeekRule.FirstFullWeek, DayOfWeek.Wednesday) <= 5)
            {
                //opLog.txtLog("dt.DayOfWeek: "+ dt.DayOfWeek+ " dt.Hour: "+ dt.Hour+ " dtOfWeeks: "+ dtOfWeeks);
                //if (dt.DayOfWeek == DayOfWeek.Wednesday && dt.Hour >= 0 && dt.Hour <= 8) return dtOfWeeks;
                return dtOfWeeks + 1;
            }
            else
            {
                //opLog.txtLog("dt.DayOfWeek: " + dt.DayOfWeek + " dt.Hour: " + dt.Hour + " dtOfWeeks: " + dtOfWeeks);
                //if (dt.DayOfWeek == DayOfWeek.Wednesday && dt.Hour >= 0 && dt.Hour <= 8)
                //{
                //    return dtOfWeeks - endLastMonthOfWeeks;
                //}
                return dtOfWeeks - endLastMonthOfWeeks + 1;
            }
        }

    }
    class AccountInfo
    {
        ///// Account parameters
        string strFile = "";
        private string[] strAccountSplit;
        private string strAccountIni = "";

        public string BrokerID { get; set; } = "";
        public string TbTAccount { get; set; } = "";
        public string StrID { get; set; }
        public string StrPwd { get; set; }
        public string StrCenter { get; set; }
        public AccountInfo(string accSelectedIndex)
        {
            strAccountIni = "account" + accSelectedIndex + ".ini";

            opLog.txtLog("strAccountIni: " + strAccountIni);
            try
            {
                if (!File.Exists(strAccountIni))
                {
                    opLog.txtLog("The account.ini file doesn't exist."
                        + Environment.NewLine
                        + "Please create this file with account and password info.");
                }
                else
                {
                    strFile = System.IO.File.ReadAllText(strAccountIni);

                    if (string.IsNullOrEmpty(strFile))
                    {
                        opLog.txtLog("The id or pwd is null."
                            + Environment.NewLine
                            + "Please fill this file with account and password info.");
                    }
                    else
                    {
                        strFile = strFile.Trim();
                        strAccountSplit = strFile.Split(',');

                        StrID = strAccountSplit[0];
                        StrCenter = strAccountSplit[1];
                        StrPwd = strAccountSplit[2];
                    }
                }
            }
            catch (IOException ex)
            {
                opLog.txtLog(
                    ex.GetType().Name+": The write operation could not " +
                    "be performed because the specified " +
                    "part of the file is locked.");
            }
        }
    }

    //比較不會變動的設定 更改設定需重開程式
    class Settings
    {
        //private string brokerID = "";
        //private string tbTAccount = "";
        //private string token = "";
        //private string contract = "";

        public string Account { get; set; } = "1";
        public string SubAccount { get; set; } = "1";
        public string TestMode { get; set; } = "0";
        public string ReconnectGap { get; set; } = "5";
        public string PrintGap { get; set; } = "5";
        public string QuoteServerHost { get; set; }
        public string QuoteServerPort { get; set; }
        public string QuoteServerHostTest { get; set; }
        public string OrderServerHost { get; set; }
        public string OrderServerPort { get; set; }
        public string OrderServerHostTest { get; set; }
        public string ChatID { get; set; }
        public string Token { get; set; }
        public Settings(string strSettingIni)
        {
            string allSetting = "";
            string[] strSettingSplit;
            try
            {
                if (!File.Exists(strSettingIni))
                {
                    opLog.txtLog("The setting.ini file doesn't exist."
                        + Environment.NewLine
                        + "Please create this file with setting info.");
                }
                else
                {
                    allSetting = System.IO.File.ReadAllText(strSettingIni);

                    if (string.IsNullOrEmpty(allSetting))
                    {
                        opLog.txtLog("The setting is null."
                            + Environment.NewLine
                            + "Please fill this file with setting info.");
                    }
                    else
                    {
                        string[] tempS = new string[2];
                        allSetting = allSetting.Trim();
                        strSettingSplit = allSetting.Split('\n');
                        foreach (var sub in strSettingSplit)
                        {
                            tempS = sub.Split('=');
                            if (String.Compare(tempS[0], "Account") == 0)
                                Account = tempS[1].Trim();
                            else if (String.Compare(tempS[0], "SubAccount") == 0)
                                SubAccount = tempS[1].Trim();
                            else if (String.Compare(tempS[0], "TestMode") == 0)
                                TestMode = tempS[1].Trim();
                            else if (String.Compare(tempS[0], "ReconnectGap") == 0)
                                ReconnectGap = tempS[1].Trim();
                            else if (String.Compare(tempS[0], "PrintGap") == 0)
                                PrintGap = tempS[1].Trim();
                            else if (String.Compare(tempS[0], "ChatID") == 0)
                                ChatID = tempS[1].Trim();
                            else if (String.Compare(tempS[0], "Token") == 0)
                                Token = tempS[1].Trim();
                            else if (String.Compare(tempS[0], "QuoteServerHost") == 0)
                                QuoteServerHost = tempS[1].Trim();
                            else if (String.Compare(tempS[0], "QuoteServerHostTest") == 0)
                                QuoteServerHostTest = tempS[1].Trim();
                            else if (String.Compare(tempS[0], "QuoteServerPort") == 0)
                                QuoteServerPort = tempS[1].Trim();
                            else if (String.Compare(tempS[0], "OrderServerHost") == 0)
                                OrderServerHost = tempS[1].Trim();
                            else if (String.Compare(tempS[0], "OrderServerHostTest") == 0)
                                OrderServerHostTest = tempS[1].Trim();
                            else if (String.Compare(tempS[0], "OrderServerPort") == 0)
                                OrderServerPort = tempS[1].Trim();
                        }
                    }
                }
            }
            catch (IOException ex)
            {
                opLog.txtLog(
                    ex.GetType().Name+": The write operation could not " +
                    "be performed because the specified " +
                    "part of the file is locked.");
            }
        }

    }
    class opLog
    {
        public static void txtLog(string msg)
        {
            string dir = Application.StartupPath + @"\Log\";
            string fileName = dir + DateTime.Now.ToString("yyyyMMdd") + ".txt";
            try
            {
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                if (!File.Exists(fileName)) File.Create(fileName).Close();
                using (StreamWriter sw = File.AppendText(fileName))
                {
                    sw.Write("[{0}] ", DateTime.Now.ToString("HH:mm:ss"));
                    sw.Write(msg);
                    sw.WriteLine("");
                    sw.Close();
                }
            }
            catch (IOException ex)
            {
                System.Diagnostics.Debug.WriteLine("CAUGHT EXCEPTION:");
                System.Diagnostics.Debug.WriteLine(ex);
                opLog.txtLog(ex.Message);
            }
        }
    }
}
