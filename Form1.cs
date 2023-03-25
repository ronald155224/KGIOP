using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Intelligence;
using Package;
using Smart;
using Notifyspace;
using System.Data.SqlClient;
using System.Threading;
using System.Text.RegularExpressions;
using System.Diagnostics;
using Microsoft.VisualBasic;

namespace OP.AutoTrading
{
    public enum ORDERRPT_LIST //下單回報陣列
    {
        ORDERSTATUS = 0,
        CNT = 1,
        ORDERNO = 2,
        TRADEDATE = 3,
        REPORTTIME = 4,
        SYMBOL = 5,
        SIDE = 6,
        PRICE = 7,
        AFTERQTY = 8,
        POSITIONEFFECT = 9,
        CODE = 10,
        ERRMSG = 11,
    }
    public enum DEALRPT_LIST //成交回報陣列
    {
        ORDERSTATUS = 0,
        CNT = 1,
        ORDERNO = 2,
        TRADEDATE = 3,
        REPORTTIME = 4,
        SYMBOL = 5,
        SIDE = 6,
        MARKET = 7,
        DEALPRICE = 8,
        DEALQTY = 9,
        CUMQTY = 10,
        LEAVEQTY = 11,
        SYMBOL1 = 12,
        DEALPRICE1 = 13,
        QTY1 = 14,
        BS1 = 15,
        SYMBOL2 = 16,
        DEALPRICE2 = 17,
        QTY2 = 18,
        BS2 = 19,
    }
    public enum PERIOD_LIST //每日配額陣列
    {
        WEDNG = 0,
        THUDT = 1,
        THUNG = 2,
        FRIDT = 3,
        FRING = 4,
        MONDT = 5,
        MONNG = 6,
        TUEDT = 7,
        TUENG = 8,
        WEDDT = 9,
    }
    public enum FUTURE_LIST //期貨陣列
    {
        SYMBOL = 0,
        MONTH = 1,
        PRICE = 2,
        OPENPRICE = 3,
        UPDOWN = 4,
        VOL = 5,
    }

    public enum OPTION_LIST //選擇權陣列
    {
        SYMBOL = 0,
        MONTH = 1,  
        STOCK = 2,
        OPENPRICE = 3,
        UPDOWN = 4,
        VOL = 5,
        OI = 6,
        TOTALVOL = 7,
        HUNDFIFDIF = 8,
        HUNDDIF = 9,
        PRICE = 10,
    }

    public enum TWELVEMONTH_SEQ
    {
        JAN = 1,//A.M:一月
        FEB = 2,//B.N:二月
        MAR = 3,//C.O:三月
        APR = 4,//D.P:四月
        MAY = 5,//E.Q:五月
        JUN = 6,//F.R:六月
        JUL = 7,//G.S:七月
        AUG = 8,//H.T:八月
        SEP = 9,//I.U:九月
        OCT = 10,//J.V:十月
        NOV = 11,//K.W:十一月
        DEC = 12,//L.X:十二月
    }

    delegate void SetAddInfoCallBack(string text);

    public partial class Form1 : Form
    {
        private bool formClose = false;
        private bool formCloseFinish = false;
        private bool loginBtnFlag = false;
        private bool FlagInvokeLoopPrint = false;
        private bool FlagInvokeLoopOrder = false;
        private bool FlagInvoke = false;

        //setting物件
        OrderSettings orderSettings = new OrderSettings(@"orderSetting.txt");
        TodayOrderStatus todayOrderStatus = new TodayOrderStatus();
        TodayOrderRecord todayOrderRecord = new TodayOrderRecord();
        Settings settings = new Settings(@"setting.ini");
        private string QuoteServerHost = "";
        private string QuoteServerPort = "";
        private string OrderServerHost = "";
        private string OrderServerPort = "";

        AccountInfo accInfo;
        DataRecord dataRecord = new DataRecord();


        //訂閱報價用list
        StringBuilder subQuoteSB = new StringBuilder();
        List<string> subQuoteList = new List<string>();
        TaiFexCom tfcom;

        //每1小時重新訂閱報價
        System.Timers.Timer tmrUnsub = new System.Timers.Timer(1000 * 60 * 60);
        //每1分鐘重新要庫存、權益數
        System.Timers.Timer tmrPosition = new System.Timers.Timer(1000 * 60 * 1);
        //每日固定時間更新OIU
        System.Timers.Timer tmrUpdateOI = new System.Timers.Timer(1000 * 60 * 30);

        Intelligence.QuoteCom quoteCom;
        private UTF8Encoding encoding = new System.Text.UTF8Encoding();

        //Threads
        //ListView Print Thread
        private Thread thrListV;
        private Thread thrOrder;
        bool loopOrderTempSwitch;

        bool btSaving = false;

        //紀錄報價元件、下單元件的連線狀態 用來做斷線重連的判斷
        private bool quoteConnect = false;
        private bool tradeConnect = false;

        Dictionary<string, int> RecoverMap = new Dictionary<string, int>();

        Telegram tl = new Telegram();
        taifexCrawler taifexCrawler;

        //存期貨商品的Dictionary
        Dictionary<string, List<string>> symbolFPriceMap = new Dictionary<string, List<string>>();
        //存選擇權商品的Dictionary
        Dictionary<string, List<string>> symbolOPriceMap = new Dictionary<string, List<string>>();
        //存委託第一回覆的Dictionary Key:CNT
        Dictionary<long, string[]> orderFirstRPT = new Dictionary<long, string[]>();
        //存委託第二回覆與委託回報之後的Dictionary Key:CNT
        Dictionary<string, List<string>> orderRPT = new Dictionary<string, List<string>>();
        //存成交回報的Dictionary Key:CNT
        Dictionary<string, List<string>> dealRPT = new Dictionary<string, List<string>>();
        //存原始庫存
        Dictionary<string, List<string>> stockRPT = new Dictionary<string, List<string>>();
        //存整合後的庫存的
        List<List<string>> stockAggre = new List<List<string>>();
        List<TQuoteStruct> tquoteStructList;
        BindingSource tquoteBinding = new BindingSource();

        //call put 的總點數
        decimal callSum = 0;
        decimal putSum = 0;
        decimal callCost = 0;
        decimal putCost = 0;
        decimal TbarShift = 0;

        //權益數
        decimal equity = 0;

        public Form1()
        {
            InitializeComponent();
            btTbarShift.Text = TbarShift.ToString();
            //char cac = '⯅';
            //char cac = Convert.ToChar(9651);//9650 9651   //11205  11206
            //char cac = Convert.ToChar(11206);
            //int a = Convert.ToInt32(cac);
            //textBox1.Text = cac.ToString();

            LoadOrderSettingToForm(); 

            AccountInfo accInfo;
            if(string.Compare(settings.TestMode,"1")==0)
            {
                QuoteServerHost = settings.QuoteServerHostTest;
                QuoteServerPort = settings.QuoteServerPort;
                OrderServerHost = settings.OrderServerHostTest;
                OrderServerPort = settings.OrderServerPort;
            }
            else
            {
                QuoteServerHost = settings.QuoteServerHost;
                QuoteServerPort = settings.QuoteServerPort;
                OrderServerHost = settings.OrderServerHost;
                OrderServerPort = settings.OrderServerPort;
            }

            //Telegram ID,Token 設定
            tl.ChatID = settings.ChatID;
            tl.Token = settings.Token;

            //PrintThread
            thrListV = new Thread(LoopPrint);
            thrListV.Start();

            //OrderThread
            thrOrder = new Thread(LoopOrder);
            thrOrder.IsBackground = true;

            //FormObjectInit
            FormObjectInit();
            TradeConnectInit();
            QuoteConnectInit();

            //每1小時Timer更新週小台與所有訂閱
            tmrUnsub.Elapsed += delegate
            {
                TComid();
            };
            tmrUnsub.AutoReset = true;
            tmrUnsub.Enabled = false;
            GC.KeepAlive(tmrUnsub);

            //每1分鐘Timer更新庫存、權益
            tmrPosition.Elapsed += delegate
            {
                checkPositionSum();
            };
            tmrPosition.AutoReset = true;
            tmrPosition.Enabled = false;
            GC.KeepAlive(tmrPosition);

            //每1分鐘Timer檢查是否到停盤時間，若為停盤時間則更新最後價格
            tmrUpdateOI.Elapsed += delegate
            {
                updateOI();
            };
            tmrUpdateOI.AutoReset = true;
            tmrUpdateOI.Enabled = false;
            GC.KeepAlive(tmrUpdateOI);

            //爬蟲
            taifexCrawler = new taifexCrawler();

            opLog.txtLog("InitComplete!");
        }

        private void DataGridView1Init()
        {
            this.Invoke((Action)(() =>
            {
                if (formClose)
                {
                    FlagInvoke = false;
                    return;
                }
                FlagInvoke = true;

                try
                {
                    opLog.txtLog("before DataGridView1Init!");
                    dataGridView1.DataSource = tquoteStructList;
                    /*if (tquoteBinding.DataSource == null)
                    {
                        tquoteBinding.DataSource = tquoteStructList;
                        dataGridView1.DataSource = tquoteBinding;
                    }
                    else tquoteBinding.ResetBindings(false);*/

                    dataGridView1.Columns["StrickPrice"].HeaderText = "履約價";
                    dataGridView1.Columns["StrickPrice"].Visible = false;
                    dataGridView1.Columns["Stock"].HeaderText = "庫存";
                    dataGridView1.Columns["OpenPrice"].HeaderText = "開盤價";
                    dataGridView1.Columns["OpenPrice"].Visible = false;
                    dataGridView1.Columns["UpDown"].HeaderText = "漲跌";
                    dataGridView1.Columns["Vol"].HeaderText = "成交量";
                    dataGridView1.Columns["OI"].HeaderText = "未平倉";
                    dataGridView1.Columns["TotalVol"].HeaderText = "合計量";
                    dataGridView1.Columns["TotalVol"].Visible = false;
                    dataGridView1.Columns["HundfifDif"].HeaderText = "150差";
                    dataGridView1.Columns["HundDif"].HeaderText = "100差";
                    dataGridView1.Columns["DealPrice"].HeaderText = "成交價";
                    dataGridView1.Columns["Symbol"].HeaderText = "商品碼";
                    dataGridView1.Columns["Symbol"].Visible = false;
                    dataGridView1.Columns["StrickPriceCenter"].HeaderText = "履約價";
                    dataGridView1.Columns["Symbol2"].HeaderText = "商品碼";
                    dataGridView1.Columns["Symbol2"].Visible = false;
                    dataGridView1.Columns["DealPrice2"].HeaderText = "成交價";
                    dataGridView1.Columns["HundDif2"].HeaderText = "100差";
                    dataGridView1.Columns["HundfifDif2"].HeaderText = "150差";
                    dataGridView1.Columns["TotalVol2"].HeaderText = "合計量";
                    dataGridView1.Columns["TotalVol2"].Visible = false;
                    dataGridView1.Columns["OI2"].HeaderText = "未平倉";
                    dataGridView1.Columns["Vol2"].HeaderText = "成交量";
                    dataGridView1.Columns["UpDown2"].HeaderText = "漲跌";
                    dataGridView1.Columns["OpenPrice2"].HeaderText = "開盤價";
                    dataGridView1.Columns["OpenPrice2"].Visible = false;
                    dataGridView1.Columns["Stock2"].HeaderText = "庫存";
                    opLog.txtLog("after DataGridView1Init!");
                }
                catch (Exception ex)
                {
                    opLog.txtLog("DataGridView1Init ex: " + ex.ToString());
                }

                FlagInvoke = false;
                if (formClose)
                {
                    return;
                }
            }));
        }

        private void FormObjectInit()
        {
            //comboboxInit
            tcResult.SelectedIndex = 0;
            if (string.Compare(settings.TestMode, "1") == 0) btTestMode.Visible = true;
        }
        private void TQStructListInit(List<string> subList)
        {
            try
            {
                List<string> tqStructMember = new List<string>();
                TQuoteStruct tmpTQStruct;
                bool foundStrikePrice = false;
                tquoteStructList = new List<TQuoteStruct>();


                foreach (string tempS in subList)
                {
                    if (tempS.Length >= 8)
                    {
                        //建立只有訂閱的履約價其餘項目為"-"的tmpTQStruct
                        tqStructMember = new List<string>();
                        tqStructMember.Add(tempS.Substring(3, 5));
                        for (int i = 1; i <= 21; i++)
                            tqStructMember.Add("-");
                        tmpTQStruct = new TQuoteStruct(tqStructMember);

                        //確認tquoteStructList裡無此履約價才add進去
                        foundStrikePrice = false;
                        if (tquoteStructList != null)
                        {
                            foreach (TQuoteStruct tquoteS in tquoteStructList)
                            {
                                if (tquoteS != null && string.Compare(tquoteS.StrickPrice, tempS.Substring(3, 5)) == 0)
                                {
                                    foundStrikePrice = true;
                                }
                            }
                        }
                        if (foundStrikePrice == false)
                        {
                            tquoteStructList.Add(tmpTQStruct);
                        }
                    }
                }

                //for (int i = 0; i < tquoteStructList.Count; i++)
                //{
                //    Console.WriteLine("[" + i + "]" + "StrickPrice: " + tquoteStructList[i].StrickPrice);
                //    Console.WriteLine("[" + i + "]" + "Stock: " + tquoteStructList[i].Stock);
                //    Console.WriteLine("[" + i + "]" + "OpenPrice: " + tquoteStructList[i].OpenPrice);
                //    Console.WriteLine("[" + i + "]" + "UpDown: " + tquoteStructList[i].UpDown);
                //    Console.WriteLine("[" + i + "]" + "Vol: " + tquoteStructList[i].Vol);
                //    Console.WriteLine("[" + i + "]" + "OI: " + tquoteStructList[i].OI);
                //    Console.WriteLine("[" + i + "]" + "TotalVol: " + tquoteStructList[i].TotalVol);
                //    Console.WriteLine("[" + i + "]" + "HundfifDif: " + tquoteStructList[i].HundfifDif);
                //    Console.WriteLine("[" + i + "]" + "HundDif: " + tquoteStructList[i].HundDif);
                //    Console.WriteLine("[" + i + "]" + "DealPrice: " + tquoteStructList[i].DealPrice);
                //    Console.WriteLine("[" + i + "]" + "Symbol: " + tquoteStructList[i].Symbol);
                //    Console.WriteLine("[" + i + "]" + "StrickPriceCenter: " + tquoteStructList[i].StrickPriceCenter);
                //    Console.WriteLine("[" + i + "]" + "Symbol2: " + tquoteStructList[i].Symbol2);
                //    Console.WriteLine("[" + i + "]" + "DealPrice2: " + tquoteStructList[i].DealPrice2);
                //    Console.WriteLine("[" + i + "]" + "HundDif2: " + tquoteStructList[i].HundDif2);
                //    Console.WriteLine("[" + i + "]" + "HundfifDif2: " + tquoteStructList[i].HundfifDif2);
                //    Console.WriteLine("[" + i + "]" + "TotalVol2: " + tquoteStructList[i].TotalVol2);
                //    Console.WriteLine("[" + i + "]" + "OI2: " + tquoteStructList[i].OI2);
                //    Console.WriteLine("[" + i + "]" + "Vol2: " + tquoteStructList[i].Vol2);
                //    Console.WriteLine("[" + i + "]" + "UpDown2: " + tquoteStructList[i].UpDown2);
                //    Console.WriteLine("[" + i + "]" + "OpenPrice2: " + tquoteStructList[i].OpenPrice2);
                //    Console.WriteLine("[" + i + "]" + "Stock2: " + tquoteStructList[i].Stock2);
                //}
            }
            catch (Exception ex)
            {
                opLog.txtLog("TQStructListInit ex: " + ex.ToString());
            }
        }

        private void QuoteConnectInit()
        {
            quoteCom = new Intelligence.QuoteCom(QuoteServerHost, ushort.Parse(QuoteServerPort), "API", "b6eb"); // Host changed on Oct/01/2014
            quoteCom.OnRcvMessage += OnQuoteRcvMessage;
            quoteCom.OnGetStatus += OnQuoteGetStatus;
            quoteCom.OnRecoverStatus += OnRecoverStatusQuote;
            quoteCom.SourceId = "API";

            DateTime NowDate = DateTime.Now;

            quoteCom.Host = QuoteServerHost;
            quoteCom.Port = ushort.Parse(QuoteServerPort);
            //Console.WriteLine("strQuoteServerHost: " + strQuoteServerHost);
        }

        private void TradeConnectInit()
        {
            //先Create , 尚未正式連線,  IP 在正式連線時再輸入即可 ; 
            tfcom = new TaiFexCom(OrderServerHost, ushort.Parse(OrderServerPort), "API");
            /*tfcom = new TaiFexCom("10.4.99.71", 8000, "API");
            this.Text = "TradeCom 範例程式 [ Version : " + tfcom.version + " ]";*/
            tfcom.OnRcvMessage += OnRcvMessage;          //資料接收事件
            tfcom.OnGetStatus += OnGetStatus;               //狀態通知事件
            tfcom.OnRcvServerTime += OnRcvServerTime;   //接收主機時間
            tfcom.OnRecoverStatus += OnRecoverStatus;   //回補狀態通知*/

            tfcom.ServerHost = OrderServerHost;
            ushort.TryParse(OrderServerPort, out tfcom.ServerPort);
            //*********Auto Reconnect Push2 *************
            /*if (cbConn2Push.Checked){
                tfcom.AutoRedirectPushLogin = true;
                tfcom.ServerHost2 = cbHost2.Text;
                ushort.TryParse(cbPort2.Text, out tfcom.ServerPort2);
                tfcom.LoginTimeout = 8000;
            }*/
            //************************
            tfcom.AutoRetriveProductInfo = true;  //是否載入商品檔
                                                  //tfcom.AutoRetriveForeignProductInfo = cbForeignProduct.Checked; 
            tfcom.AutoSubReport = true;   //是否回補回報&註冊即時回報(國內)
            tfcom.AutoSubReportForeign = true;   //是否回補回報&註冊即時回報(國外)
                                                 //2015.7.9 將即時 & 回補 回報分離; 未設定時則 AutoRecoverReport=AutoSubRepor ; 
            tfcom.AutoRecoverReport = true;
            tfcom.AutoRecoverReportForeign = true;
            tfcom.ConnectTimeout = 5000;              //連線time out 時間 1/1000 秒
                                                      //tfcom.ReconnectInterval = 2000;           //自動重新連線時間間隔 1/1000 秒
                                                      //tfcom.ReconnectMaxCount = 2;             //自動重新連線.連線次數, 0,代表不自動連線
        }
        private async void Login()
        {
            CalculateComID calcuComID;
            //星期三早盤先不要換下周的新合約等待收盤過後再換
            if (string.Compare( DateTime.Now.DayOfWeek.ToString("d") , "3" )==0 && DateCalculateExtensions.IsLegalTime(DateTime.Now, "050001-143000") )
                calcuComID = new CalculateComID(DateTime.Now.AddDays(-1));
            else calcuComID = new CalculateComID(DateTime.Now);

            //quoteCom.Connect(quoteCom.Host, quoteCom.Port);
            //AccountInfo accInfo = new AccountInfo(cbID.SelectedIndex);
            ///// Quote
            //quoteCom.Login(accInfo.StrID, accInfo.StrPwd, ' ');

            //QuoteCom 伺服器連線+登入
            string weekMXSymbol = calcuComID.PartialComIDMX[0] + calcuComID.PartialComIDMX[1];
            Console.WriteLine("先拿週小台: " + weekMXSymbol);
            opLog.txtLog("先拿週小台: " + weekMXSymbol);
            quoteCom.Connect2Quote(quoteCom.Host, quoteCom.Port, accInfo.StrID, accInfo.StrPwd, ' ', "");
            await Task.Delay(5000);
            opLog.txtLog("Connect2Quote!");
            //先拿到週小台的最後報價
            short istatus = quoteCom.RetriveLastPrice(weekMXSymbol);
            if (istatus < 0)   //
                opLog.txtLog(quoteCom.GetSubQuoteMsg(istatus));
            quoteCom.LoadTaifexProductXML();
            opLog.txtLog("quoteCom.LoadTaifexProductXML!");

            //下單元件登入
            tfcom.LoginDirect(tfcom.ServerHost, tfcom.ServerPort, accInfo.StrID + ",," + accInfo.StrPwd);

            AddInfoTG("tfcom Login!");
            opLog.txtLog("tfcom Login!");
        }

        //帳務查詢
        private void checkPositionSum()
        {
            try
            {
                if (tradeConnect == true)
                {
                    string mk = "I";//I=國內   O=國外
                    long rtn = tfcom.RetrivePositionSum(mk, accInfo.BrokerID, accInfo.TbTAccount, "", "");
                    //long rtn = tfcom.RetrivePositionSum(mk, "F004000", "", "", "");
                    //dgv1616.Rows.Clear();
                    if (rtn == -1)
                    {
                        MessageBox.Show("無查詢權限1616!!!");
                        opLog.txtLog("無查詢權限1616!!!");
                        AddInfoTG("無查詢權限1616!!!");
                    }

                    rtn = tfcom.RetrivePositionDetail(mk, accInfo.BrokerID, accInfo.TbTAccount, "", "");
                    //rtn = tfcom.RetrivePositionDetail(mk, "F004000", "", "", "");
                    if (rtn == -1)
                    {
                        MessageBox.Show("無查詢權限1618!!!");
                        opLog.txtLog("無查詢權限1618!!!");
                        AddInfoTG("無查詢權限1618!!!");
                    }
                    rtn = tfcom.RetriveFMargin(mk, accInfo.BrokerID, accInfo.TbTAccount, "", "");
                    //rtn = tfcom.RetriveFMargin(mk, "F004000", "", "", "");
                    if (rtn == -1)
                    {
                        MessageBox.Show("無查詢權限1626!!!");
                        opLog.txtLog("無查詢權限1626!!!");
                        AddInfoTG("無查詢權限1626!!!");
                    }
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine("checkPositionSum ex:" + ex);
            }

        }

        private void TComid()
        {
            CalculateComID calcuComID;
            StringBuilder sb = new StringBuilder();

            //若是星期三早盤先不換下星期的新合約
            if (string.Compare(DateTime.Now.DayOfWeek.ToString("d"), "3") == 0 && DateCalculateExtensions.IsLegalTime(DateTime.Now, "050001-143000"))
                calcuComID = new CalculateComID(DateTime.Now.AddDays(-1));
            else calcuComID = new CalculateComID(DateTime.Now);

            //Console.WriteLine("TComid calcuComID.PartialComIDC: " + calcuComID.PartialComIDC[0]+ calcuComID.PartialComIDC[1]);
            opLog.txtLog("TComid calcuComID.PartialComIDC: " + calcuComID.PartialComIDC[0]+ calcuComID.PartialComIDC[1]);
            List<PT01802> rtn = quoteCom.GetTaifexProductDetailList(calcuComID.PartialComIDC[0]);
            opLog.txtLog("GetTaifexProductDetailList done!");
            short istatus;
            string weekMX = calcuComID.PartialComIDMX[0] + calcuComID.PartialComIDMX[1];

            //如果subQuoteSB有訂閱則清空
            if (subQuoteSB.Length > 0)
            {
                //Console.WriteLine("subQuoteSB.Length: " + subQuoteSB.Length);
                quoteCom.UnsubQuote(subQuoteSB.ToString());
                subQuoteList.Clear();
                subQuoteSB.Clear(); 
                AddInfo("UnsubQuote All!");
                AddInfoTG("UnsubQuote All!");
                opLog.txtLog("UnsubQuote All!");
            }
            try { if (string.IsNullOrEmpty(weekMX)) throw new ArgumentNullException("無法自動找到周小台的Symbol，請檢查!"); }
            catch(ArgumentNullException ex)
            {
                System.Diagnostics.Debug.WriteLine("CAUGHT EXCEPTION:");
                System.Diagnostics.Debug.WriteLine(ex);
                Console.WriteLine(ex.Message);
                string input = Interaction.InputBox("請輸入週小台商品代碼", "自訂週小台代碼", "", -1, -1);
                weekMX = input;
            }

            try
            {
                //取得小台的最後報價(此處是為了60分鐘一次更新使用)
                opLog.txtLog("weekMX: "+ weekMX+ " RetriveLastPrice");
                istatus = quoteCom.RetriveLastPrice(weekMX);
                if (istatus < 0)   //
                {
                    AddInfo(quoteCom.GetSubQuoteMsg(istatus));
                    AddInfoTG(quoteCom.GetSubQuoteMsg(istatus));
                    opLog.txtLog(quoteCom.GetSubQuoteMsg(istatus));
                    MessageBox.Show("Get不到週小台的最後報價，請檢查!!!!");
                }
                //將這個月台指加入訂閱List
                subQuoteSB.Append(weekMX);
                subQuoteList.Add(weekMX);
            }
            catch (ArgumentOutOfRangeException ex)
            {
                System.Diagnostics.Debug.WriteLine("CAUGHT EXCEPTION:");
                System.Diagnostics.Debug.WriteLine(ex);
                opLog.txtLog(ex.Message);
            }
            catch (System.Collections.Generic.KeyNotFoundException ex)
            {
                System.Diagnostics.Debug.WriteLine("CAUGHT EXCEPTION:");
                System.Diagnostics.Debug.WriteLine(ex);
                opLog.txtLog(ex.Message);
            }
            catch (Exception ex)
            {
                opLog.txtLog("取得小台的最後報價 ex:" + ex);
            }

            //while (!symbolFPriceMap.ContainsKey(weekMX))
            //{
            //    Console.WriteLine("Not yet got the MXprice!");
            //}
            try
            {
                //把小台以50點為單位無條件進位
                decimal weekMXPrice = 0;
                if (symbolFPriceMap.ContainsKey(weekMX) && decimal.Parse(symbolFPriceMap[weekMX][(int)FUTURE_LIST.PRICE]) > 0)
                {
                    weekMXPrice = Math.Round(decimal.Parse(symbolFPriceMap[weekMX][(int)FUTURE_LIST.PRICE]) / 50, 0, MidpointRounding.AwayFromZero) * 50 + TbarShift*50;
                    dataRecord.saveMXF(weekMXPrice);
                }
                else
                {
                    AddInfo("未收到週小台的價格回報，請檢查!!!!");
                    AddInfoTG("未收到週小台的價格回報，請檢查!!!!");
                    opLog.txtLog("未收到週小台的價格回報，請檢查!!!!");
                    MessageBox.Show("未收到週小台的價格回報，請檢查!!!!");
                    weekMXPrice = dataRecord.readMXF();
                }

                int countPLower = 0, countCUppon = 0;
                int tmpUpPrice = 0, tmpDownPrice = 0;

                tmpUpPrice = (int)weekMXPrice - 50 * 2;
                tmpDownPrice = (int)weekMXPrice + 50 * 2;
                for (int j = 0; j < 24; j++)
                {   //dump putoption in this month
                    for (int i = 0; i < rtn.Count; i++)
                    {
                        if (rtn[i].ComId.Contains(calcuComID.PartialComIDP[1])
                        //Put抓比小台價格小的履約價12檔
                        && tmpDownPrice == int.Parse(rtn[i].ComId.Substring(3, 5))
                        && countPLower < 12
                        )
                        {
                            sb.Append("ComType=").Append(rtn[i].ComType).Append("|");
                            sb.Append("ComId=").Append(rtn[i].ComId).Append("|");
                            sb.Append("RisePrice=").Append(rtn[i].RisePrice).Append("|");
                            sb.Append("FallPrice=").Append(rtn[i].FallPrice).Append("|");
                            sb.Append("PriceDecimal=").Append(rtn[i].PriceDecimal).Append("|");
                            sb.Append("StkPriceDecimal=").Append(rtn[i].StkPriceDecimal).Append("|");
                            sb.Append("Hot=").Append(rtn[i].Hot).Append("|").Append(Environment.NewLine);
                            subQuoteSB.Append("|").Append(rtn[i].ComId);
                            subQuoteList.Add(rtn[i].ComId);
                            //opLog.txtLog("rtn["+i+"].ComId: " + rtn[i].ComId);
                            //先得到最後報價最後再統一訂閱
                            istatus = quoteCom.RetriveLastPrice(rtn[i].ComId);
                            if (istatus < 0)   //
                            {
                                AddInfo(quoteCom.GetSubQuoteMsg(istatus));
                                AddInfoTG(quoteCom.GetSubQuoteMsg(istatus));
                                opLog.txtLog(quoteCom.GetSubQuoteMsg(istatus));
                            }
                            countPLower++;
                        }
                        //dump calloption in this month
                        else if (rtn[i].ComId.Contains(calcuComID.PartialComIDC[1])
                        //Call抓比小台價格大的履約價12檔
                        && tmpUpPrice == int.Parse(rtn[i].ComId.Substring(3, 5))
                        && countCUppon < 12
                        )
                        {
                            subQuoteSB.Append("|").Append(rtn[i].ComId);
                            subQuoteList.Add(rtn[i].ComId);
                            //先得到最後報價最後再統一訂閱
                            istatus = quoteCom.RetriveLastPrice(rtn[i].ComId);
                            if (istatus < 0)
                            {
                                AddInfo(quoteCom.GetSubQuoteMsg(istatus));
                                AddInfoTG(quoteCom.GetSubQuoteMsg(istatus));
                                opLog.txtLog(quoteCom.GetSubQuoteMsg(istatus));
                            }
                            countCUppon++;
                        }
                    }
                    tmpUpPrice += 50;
                    tmpDownPrice -= 50;
                    //AddInfo(sb.ToString());
                }

                //訂閱期貨、選擇權報價
                istatus = quoteCom.SubQuotes(subQuoteSB.ToString());
                if (istatus < 0)
                {
                    AddInfo(quoteCom.GetSubQuoteMsg(istatus));
                    AddInfoTG(quoteCom.GetSubQuoteMsg(istatus));
                    opLog.txtLog(quoteCom.GetSubQuoteMsg(istatus));
                }
                //foreach (string temp in subQuoteList)
                //{
                //    opLog.txtLog("Qlist: " + temp);
                //}
                subQuoteList.Sort();
                TQStructListInit(subQuoteList);
                DataGridView1Init();
            }
            catch (Exception ex)
            {
                Console.WriteLine("checkPositionSum ex:" + ex);
            }
        }

        //Telegram Func + Time
        private void AddInfoTG(string msg)
        {
            string fMsg = String.Format("[{0}] {1} {2}", DateTime.Now.ToString("HH:mm:ss:ffff"), msg, Environment.NewLine);
            try
            {
                if (fMsg.Length <= 100) tl.SendNotify(fMsg);
            }
            catch(Exception ex)
            {
                Console.WriteLine("AddInfoTG: "+ex.Message);
            }
        }

        #region 凱基元件的回報與Log
        //報價元件的回報
        private void OnQuoteRcvMessage(object sender, PackageBase package)
        {
            if (package.TOPIC != null)
                if (RecoverMap.ContainsKey(package.TOPIC)) RecoverMap[package.TOPIC]++;
            if (this.InvokeRequired)
            {
                Intelligence.OnRcvMessage_EventHandler d = new Intelligence.OnRcvMessage_EventHandler(OnQuoteRcvMessage);
                this.Invoke(d, new object[] { sender, package });
                return;
            }
            else
            {
                //Console.WriteLine("don't require invok!");
            }

            StringBuilder sb;
            List<string> tempList;
            try 
            {
                switch (package.DT)
                {
                    case (ushort)DT.LOGIN:
                        P001503 _p001503 = (P001503)package;
                        if (_p001503.Code == 0) AddInfo("可註冊檔數：" + _p001503.Qnum); //目前顯示50檔
                        break;
                    #region 公告 2014.12.12 ADD
                    case (ushort)DT.NOTICE_RPT:
                        P001701 p1701 = (P001701)package;
                        AddInfo(p1701.ToLog());
                        break;    //公告(被動查詢) 
                    case (ushort)DT.NOTICE:   //公告(主動)
                        P001702 p1702 = (P001702)package;
                        AddInfo(p1702.ToLog());
                        break;
                    #endregion
                    case (ushort)DT.QUOTE_I020:
                    case (ushort)DT.QUOTE_I022:  //2014.4.2. ADD 20022 盤前揭示  //4~5秒觸發一次
                        PI20020 i20020 = (PI20020)package;

                        sb = new StringBuilder(Environment.NewLine);
                        sb.Append("DT:[" + i20020.DT + "]");
                        sb.Append("資料時間:").Append(String.Format("{0:00:00:00\\.00}", i20020.MatchTime));
                        sb.Append("    Tick序號:").Append(i20020.InfoSeq).Append(Environment.NewLine);
                        sb.Append("最後封包:").Append(i20020.LastItem).Append(Environment.NewLine);
                        sb.Append("期貨/選擇權:").Append(i20020.Market).Append("  [").Append(i20020.Symbol).Append("]").Append(Environment.NewLine);
                        //sb.Append("價格正負號:").Append(i20020.MATCH_SIGN).Append(Environment.NewLine);
                        sb.Append("成交 [價:").Append(i20020.Price).Append("   量:").Append(i20020.MatchQuantity).Append("]").Append(Environment.NewLine);
                        sb.Append("累計成交  [量:").Append(i20020.MatchTotalQty).Append("  買筆數:")
                            .Append(i20020.MatchBuyCnt).Append("  賣筆數:")
                            .Append(i20020.MatchSellCnt).Append("]").Append(Environment.NewLine);
                        sb.Append("=============================");
                        //if (i20020.DT == 20020) AddInfo(sb.ToString());
                        //else AddInfoCover(i20020.ToLog());//AddInfoCover(sb.ToString()); 


                        if (i20020.Market == 'F')
                        {
                            //期貨 list順序"商品全碼", "月", "成交價", "開盤價", "漲跌", "成交量"
                            tempList = new List<string>();
                            tempList.Add(i20020.Symbol);
                            tempList.Add(i20020.Symbol[i20020.Symbol.Length - 2].ToString());
                            tempList.Add((i20020.Price).ToString());
                            tempList.Add("");
                            tempList.Add("");
                            tempList.Add((i20020.MatchTotalQty).ToString());

                            if (symbolFPriceMap.ContainsKey(i20020.Symbol))
                            {
                                symbolFPriceMap[i20020.Symbol][(int)FUTURE_LIST.VOL] = tempList[(int)FUTURE_LIST.VOL];
                            }
                            else
                            {
                                symbolFPriceMap.Add(i20020.Symbol, tempList);
                            }
                        }

                        if (i20020.Market == 'O')
                        {
                            //選擇權 list順序"商品全碼", "月", "成交價", "開盤價", "漲跌", "成交量", "50差", "100差"
                            tempList = new List<string>();
                            tempList.Add(i20020.Symbol);
                            tempList.Add(i20020.Symbol[i20020.Symbol.Length - 2].ToString());
                            tempList.Add("");
                            tempList.Add("");
                            tempList.Add("");
                            tempList.Add((i20020.MatchTotalQty).ToString());
                            tempList.Add("0");
                            tempList.Add("0");
                            tempList.Add("0");
                            tempList.Add("0");
                            tempList.Add((i20020.Price).ToString());
                            //此商品已存Dictionary裡在就只更新成交量
                            if (symbolOPriceMap.ContainsKey(i20020.Symbol))
                            {
                                symbolOPriceMap[i20020.Symbol][(int)OPTION_LIST.VOL] = tempList[(int)OPTION_LIST.VOL];
                            }
                            //此商品不存在Dictionary裡就add
                            else
                            {
                                symbolOPriceMap.Add(i20020.Symbol, tempList);
                            }
                        }
                        break;
                    case (ushort)DT.QUOTE_I020_RECOVER:   //I20 行情回補
                        PI20020 i21020 = (PI20020)package;
                        sb = new StringBuilder(Environment.NewLine);
                        sb.Append("===I020 資料回補 =====").Append(Environment.NewLine);
                        sb.Append("資料時間:").Append(String.Format("{0:00:00:00\\.00}", i21020.MatchTime));
                        sb.Append("    Tick序號:").Append(i21020.InfoSeq).Append(Environment.NewLine);
                        sb.Append("最後封包:").Append(i21020.LastItem).Append(Environment.NewLine);
                        sb.Append("期貨/選擇權:").Append(i21020.Market).Append("  [").Append(i21020.Symbol).Append("]").Append(Environment.NewLine);
                        sb.Append("成交 [價:").Append(i21020.Price).Append("   量:").Append(i21020.MatchQuantity).Append("]").Append(Environment.NewLine);
                        sb.Append("累計成交  [量:").Append(i21020.MatchTotalQty).Append("  買筆數:")
                            .Append(i21020.MatchBuyCnt).Append("  賣筆數:")
                            .Append(i21020.MatchSellCnt).Append("]").Append(Environment.NewLine);
                        AddInfoCover(sb.ToString());

                        break;
                    case (ushort)DT.QUOTE_I060:   //2015.9.8 Lynn Add: 現貨標的 (人民幣 RTF$)
                                                  //PI20060 i20060 = (PI20060)package;
                                                  //sb = new StringBuilder();
                                                  //sb.Append("******DataType:").Append(i20060.DT).Append("********").Append(Environment.NewLine)
                                                  //    .Append("商品代碼").Append(i20060.Symbol).Append("資料時間:")
                                                  //    .Append(String.Format("{0:00:00:00\\.00}", i20060.MatchTime))
                                                  //    .Append("成交價:").Append(i20060.MatchPrice)
                                                  //    .Append("買價:").Append(i20060.BidPrice)
                                                  //    .Append("賣價:").Append(i20060.AskPrice)
                                                  //    .Append("定盤價:").Append(i20060.FixPrice)
                                                  //    .Append("參考價:").Append(i20060.RefPrice)
                                                  //    .Append(Environment.NewLine);
                                                  //AddInfo(sb.ToString());
                        break;
                    case (ushort)DT.QUOTE_I021:
                        //PI20021 i20021 = (PI20021)package;
                        //sb = new StringBuilder(Environment.NewLine);
                        //sb.Append("******DataType:").Append(i20021.DT).Append("********");
                        //sb.Append("資料時間:").Append(String.Format("{0:00:00:00\\.00}", i20021.MatchTime));
                        //sb.Append("當日最高成交價格:").Append(i20021.DayHighPrice).Append("當日最低成交價格:").Append(i20021.DayLowPrice).Append(Environment.NewLine);
                        //AddInfo(sb.ToString());
                        break;
                    case (ushort)DT.QUOTE_I023:
                        //PI20023 i20023 = (PI20023)package;
                        //sb = new StringBuilder(Environment.NewLine);
                        //sb.Append("**********DataType:").Append(i20023.DT).Append("**********");
                        //AddInfo(sb.ToString());
                        break;
                    case (ushort)DT.QUOTE_I030:
                        PI20030 i20030 = (PI20030)package;
                        //AddInfo(i20030.ToLog());
                        //2018.6.28 Lynn TestAddInfoCover(i20030.ToLog());
                        break;
                    case (ushort)DT.QUOTE_BASE_P08:   //商品檔 :  RetriveQuoteList() 
                        /*PI20008 pi20008 = (PI20008)package;
                        if (pi20008.Market == 'F') break; 
                        sb = new StringBuilder();
                        sb.Append("商品:").Append(pi20008.Symbol).Append(" 價格小數位:").Append(pi20008.PriceDecimal).Append(" 履約價格小數位:").Append(pi20008.StrikePriceDecimal);
                        AddInfo(sb.ToString());

                        sb = new StringBuilder(Environment.NewLine);
                        sb.Append("期貨/選擇權:").Append(pi20008.Market).Append(Environment.NewLine);
                        sb.Append("商品代號:").Append(pi20008.Symbol).Append(Environment.NewLine);
                        sb.Append("商品IDX:").Append(pi20008.SymbolIdx).Append(Environment.NewLine);
                        sb.Append("第一漲停價:").Append(pi20008._RISE_LIMIT_PRICE1).Append(Environment.NewLine);
                        sb.Append("參考價:").Append(pi20008._REFERENCE_PRICE).Append(Environment.NewLine);
                        sb.Append("第一漲停價:").Append(pi20008._RISE_LIMIT_PRICE1).Append(Environment.NewLine);
                        sb.Append("第二漲停價:").Append(pi20008._RISE_LIMIT_PRICE2).Append(Environment.NewLine);
                        sb.Append("第二跌停價:").Append(pi20008._FALL_LIMIT_PRICE2).Append(Environment.NewLine);
                        sb.Append("第三漲停價:").Append(pi20008._RISE_LIMIT_PRICE3).Append(Environment.NewLine);
                        sb.Append("第三跌停價:").Append(pi20008._FALL_LIMIT_PRICE3).Append(Environment.NewLine);
                        sb.Append("契約種類:").Append(pi20008._PROD_KIND).Append(Environment.NewLine);
                        sb.Append("價格欄位小數位數:").Append(pi20008.PriceDecimal).Append(Environment.NewLine);
                        sb.Append("商品名稱:").Append(pi20008._PROD_NAME).Append(Environment.NewLine);
                         sb.Append("下市日期:").Append(pi20008.END_DATE).Append(Environment.NewLine);
                        sb.Append("=================================");
                        AddInfo(sb.ToString());
                        */
                        break;
                    case (ushort)DT.QUOTE_CLOSE_I070:  //收盤:  RetriveClosePrice()   //這會直接收所有商品，還在想要怎麼只收單個商品
                        PI20070 pb = (PI20070)package;
                        sb = new StringBuilder();
                        sb.Append("期貨/選擇權:").Append(pb.MESSAGE_KIND).Append(Environment.NewLine);
                        sb.Append("商品代號:").Append(pb.PROD_ID).Append(Environment.NewLine);
                        sb.Append("該期最高價:").Append(pb.TERM_HIGH_PRICE).Append(Environment.NewLine);
                        sb.Append("該期最低價:").Append(pb.TERM_LOW_PRICE).Append(Environment.NewLine);
                        sb.Append("該日最高價:").Append(pb.DAY_HIGH_PRICE).Append(Environment.NewLine);
                        sb.Append("該日最低價:").Append(pb.DAY_LOW_PRICE).Append(Environment.NewLine);
                        sb.Append("開盤價:").Append(pb.OPEN_PRICE).Append(Environment.NewLine);
                        sb.Append("最後買價:").Append(pb.BUY_PRICE).Append(Environment.NewLine);
                        sb.Append("最後賣價:").Append(pb.SELL_PRICE).Append(Environment.NewLine);
                        sb.Append("收盤價:").Append(pb.CLOSE_PRICE).Append(Environment.NewLine);
                        sb.Append("委託買進總筆數:").Append(pb.BO_COUNT_TAL).Append(Environment.NewLine);
                        sb.Append("委託買進總口數:").Append(pb.BO_QNTY_TAL).Append(Environment.NewLine);
                        sb.Append("委託賣出總筆數:").Append(pb.SO_COUNT_TAL).Append(Environment.NewLine);
                        sb.Append("委託賣出總口數:").Append(pb.SO_QNTY_TAL).Append(Environment.NewLine);
                        sb.Append("總成交筆數:").Append(pb.TOTAL_COUNT).Append(Environment.NewLine);
                        sb.Append("總成交量:").Append(pb.TOTAL_QNTY).Append(Environment.NewLine);
                        sb.Append("合併委託買進總筆數:").Append(pb.COMBINE_BO_COUNT_TAL).Append(Environment.NewLine);
                        sb.Append("合併委託買進總口數:").Append(pb.COMBINE_BO_QNTY_TAL).Append(Environment.NewLine);
                        sb.Append("合併委託賣出總筆數:").Append(pb.COMBINE_SO_COUNT_TAL).Append(Environment.NewLine);
                        sb.Append("合併委託賣出總口數:").Append(pb.COMBINE_SO_QNTY_TAL).Append(Environment.NewLine);
                        sb.Append("合併總成交量:").Append(pb.COMBINE_TOTAL_QNTY).Append(Environment.NewLine);
                        sb.Append("價格欄位小數點:").Append(pb.DECIMAL_LOCATOR).Append(Environment.NewLine);
                        sb.Append("=================");
                        //double decPoint = double.Parse(pb.DECIMAL_LOCATOR.ToString());
                        //foreach (string quoteSymbol in subQuoteList)
                        //{
                        //    if (string.Compare(quoteSymbol, pb.PROD_ID) == 0)
                        //    {
                        //        if (pb.MESSAGE_KIND == 'F')
                        //        {
                        //            //期貨 list順序"商品全碼", "月", "成交價", "開盤價", "漲跌", "成交量"
                        //            tempList = new List<string>();
                        //            tempList.Add(pb.PROD_ID);
                        //            tempList.Add(pb.PROD_ID[pb.PROD_ID.Length - 2].ToString());
                        //            tempList.Add( (pb.CLOSE_PRICE/ Math.Pow(10, decPoint) ).ToString() );
                        //            tempList.Add( (pb.OPEN_PRICE / Math.Pow(10, decPoint) ).ToString() );
                        //            tempList.Add("");
                        //            tempList.Add(pb.TOTAL_QNTY.ToString());

                        //            if (symbolFPriceMap.ContainsKey(pb.PROD_ID))
                        //            {
                        //                symbolFPriceMap[pb.PROD_ID] = tempList;
                        //            }
                        //            else
                        //            {
                        //                symbolFPriceMap.Add(pb.PROD_ID, tempList);
                        //            }
                        //            if (tempList != null)
                        //            foreach (string temp in tempList)
                        //            {
                        //                Console.WriteLine(temp);
                        //            }
                        //        }
                        //        else if (pb.MESSAGE_KIND == 'O')
                        //        {
                        //            tempList = new List<string>();
                        //            tempList.Add(pb.PROD_ID);
                        //            tempList.Add(pb.PROD_ID[pb.PROD_ID.Length - 2].ToString());
                        //            tempList.Add( ( pb.OPEN_PRICE/ Math.Pow(10, decPoint) ).ToString() );
                        //            tempList.Add("");
                        //            tempList.Add(pb.TOTAL_QNTY.ToString());
                        //            tempList.Add(pb.COMBINE_TOTAL_QNTY.ToString());
                        //            tempList.Add("0");
                        //            tempList.Add("0");
                        //            tempList.Add( ( pb.CLOSE_PRICE/ Math.Pow(10, decPoint) ).ToString() );

                        //            //如果商品已存在在Dictionary
                        //            if (symbolOPriceMap.ContainsKey(pb.PROD_ID))
                        //            {
                        //                symbolOPriceMap[pb.PROD_ID] = tempList;
                        //            }
                        //            else
                        //            {
                        //                symbolOPriceMap.Add(pb.PROD_ID, tempList);
                        //            }
                        //            if (tempList != null)
                        //            foreach (string temp in tempList)
                        //            {
                        //                Console.WriteLine(temp);
                        //            }
                        //        }
                        //        AddInfo(sb.ToString());
                        //    }
                        //}
                        //AddInfo(sb.ToString());
                        break;
                    case (ushort)DT.QUOTE_I080:
                    case (ushort)DT.QUOTE_I082: //case (ushort)DT.QUOTE_I082:   //2014.4.2 ADD 盤前揭示  //觸發最頻繁
                        PI20080 i20080 = (PI20080)package;
                        sb = new StringBuilder(Environment.NewLine);
                        sb.Append("DT:[" + i20080.DT + "]");
                        sb.Append("商品代號:").Append(i20080.Symbol).Append(Environment.NewLine);
                        for (int i = 0; i < 5; i++)
                            sb.Append(String.Format("五檔[{0}] 買[價:{1:N} 量:{2:N}]    賣[價:{3:N} 量:{4:N}]", i + 1, i20080.BUY_DEPTH[i].PRICE, i20080.BUY_DEPTH[i].QUANTITY, i20080.SELL_DEPTH[i].PRICE, i20080.SELL_DEPTH[i].QUANTITY)).Append(Environment.NewLine);
                        sb.AppendLine("衍生委託第一檔買進價格:" + i20080.FIRST_DERIVED_BUY_PRICE);
                        sb.AppendLine("衍生委託第一檔買進數量:" + i20080.FIRST_DERIVED_BUY_QTY);
                        sb.AppendLine("衍生委託第一檔賣出價格:" + i20080.FIRST_DERIVED_SELL_PRICE);
                        sb.AppendLine("衍生委託第一檔賣出數量" + i20080.FIRST_DERIVED_SELL_QTY);
                        sb.AppendLine("資料時間" + i20080.DATA_TIME);
                        sb.AppendLine("小數位位:" + i20080.PriceDecimal);
                        sb.Append("==============================");
                        //AddInfo(sb.ToString());
                        //if (i20080.DT == 20080) AddInfo(sb.ToString());
                        //else AddInfoCover(i20080.ToLog()); //AddInfoCover(sb.ToString()); 

                        //照理說這個是最頻繁的回報，但先前會先拿到其他資訊，所以這裡只更新買價
                        //if (symbolFPriceMap.ContainsKey(i20080.Symbol)) {
                        //    symbolFPriceMap[i20080.Symbol][(int)FUTURE_LIST.PRICE] = i20080.SELL_DEPTH[0].PRICE.ToString();
                        //}
                        //else if (symbolOPriceMap.ContainsKey(i20080.Symbol)) {
                        //    symbolOPriceMap[i20080.Symbol][(int)OPTION_LIST.PRICE] = i20080.BUY_DEPTH[0].PRICE.ToString();
                        //}
                        //else {
                        //    AddInfo("The PriceMap do not contain the key!!!");
                        //    AddInfoTG("The PriceMap do not contain the key!!!");
                        //    opLog.txtLog("The PriceMap do not contain the key!!!");
                        //}

                        break;
                    case (ushort)DT.QUOTE_I080_RECOVER:  //I080回補

                        PI20080 i21080 = (PI20080)package;
                        sb = new StringBuilder(Environment.NewLine);
                        sb.Append("==I080 回補資料 ===");
                        sb.Append("商品代號:").Append(i21080.Symbol).Append(Environment.NewLine);
                        for (int i = 0; i < 5; i++)
                            sb.Append(String.Format("五檔[{0}] 買[價:{1:N} 量:{2:N}]    賣[價:{3:N} 量:{4:N}]", i + 1, i21080.BUY_DEPTH[i].PRICE, i21080.BUY_DEPTH[i].QUANTITY, i21080.SELL_DEPTH[i].PRICE, i21080.SELL_DEPTH[i].QUANTITY)).Append(Environment.NewLine);
                        sb.AppendLine("衍生委託第一檔買進價格:" + i21080.FIRST_DERIVED_BUY_PRICE);
                        sb.AppendLine("衍生委託第一檔買進數量:" + i21080.FIRST_DERIVED_BUY_QTY);
                        sb.AppendLine("衍生委託第一檔賣出價格:" + i21080.FIRST_DERIVED_SELL_PRICE);
                        sb.AppendLine("衍生委託第一檔賣出數量" + i21080.FIRST_DERIVED_SELL_QTY);
                        sb.AppendLine("資料時間" + i21080.DATA_TIME);
                        AddInfoCover(sb.ToString());

                        break;
                    case (ushort)DT.QUOTE_LAST_PRICE:   // 最後價格: RetriveLastPrice()   //callAPI後只會觸發一次
                        PI20026 pi20026 = (PI20026)package;
                        sb = new StringBuilder(Environment.NewLine);
                        sb.Append("商品代號:").Append(pi20026.Symbol).Append(" 最後價格:").Append(pi20026.MatchPrice).Append(Environment.NewLine);
                        sb.Append("當日最高成交價格:").Append(pi20026.DayHighPrice).Append(" 當日最低成交價格:").Append(pi20026.DayLowPrice);
                        sb.Append("開盤價:").Append(pi20026.FirstMatchPrice).Append(" 開盤量:").Append(pi20026.FirstMatchQty).Append(Environment.NewLine);
                        sb.Append("參考價:").Append(pi20026.ReferencePrice).Append("累計成交量:").Append(pi20026.MatchTotalQty)
                            .Append("盤別:").Append(pi20026.Session)
                            .Append(" 交易暫停否:").Append(pi20026.Break_Mark).Append(Environment.NewLine);
                        //2016.5  20026 封包 新增4欄位(衍生第一檔)
                        sb.Append("衍生委託第一檔買進價格:").Append(pi20026.FirstDerivedBuyPrice).Append(Environment.NewLine)
                           .Append("衍生委託第一檔買進數量:").Append(pi20026.FirstDerivedBuyQty).Append(Environment.NewLine)
                           .Append("衍生委託第一檔賣出價格:").Append(pi20026.FirstDerivedSellPrice).Append(Environment.NewLine)
                           .Append("衍生委託第一檔賣出數量:").Append(pi20026.FirstDerivedSellQty).Append(Environment.NewLine);

                        for (int i = 0; i < 5; i++)
                            sb.Append(String.Format("五檔[{0}] 買[價:{1:N} 量:{2:N}]    賣[價:{3:N} 量:{4:N}]", i + 1, pi20026.BUY_DEPTH[i].PRICE, pi20026.BUY_DEPTH[i].QUANTITY, pi20026.SELL_DEPTH[i].PRICE, pi20026.SELL_DEPTH[i].QUANTITY)).Append(Environment.NewLine);
                        //2020.10.26 Lynn Add:盤中漲跌停價格
                        sb.Append("漲停價格階段:").Append(pi20026.Rise_Level).Append("漲停價格:").Append(pi20026.RiseStopPrice).Append(Environment.NewLine);
                        sb.Append("跌停價格階段:").Append(pi20026.Fall_Level).Append("跌停價格:").Append(pi20026.FallStopPrice).Append(Environment.NewLine);
                        sb.Append("==========================");
                        AddInfo(sb.ToString());

                        //小台大台 list順序"商品全碼", "月", "成交價", "開盤價", "漲跌", "成交量"
                        if (pi20026.Symbol.Contains("MX") || pi20026.Symbol.Contains("TXF"))
                        {
                            tempList = new List<string>();
                            tempList.Add(pi20026.Symbol);
                            tempList.Add(pi20026.Symbol[pi20026.Symbol.Length - 2].ToString());
                            tempList.Add(pi20026.MatchPrice.ToString());
                            tempList.Add(pi20026.FirstMatchPrice.ToString());
                            tempList.Add("");
                            tempList.Add(pi20026.MatchTotalQty.ToString());

                            //如果商品已存在在Dictionary裡則只更新 成交價、開盤價、成交量
                            if (symbolFPriceMap.ContainsKey(pi20026.Symbol))
                            {
                                if (int.Parse(tempList[(int)FUTURE_LIST.PRICE]) >= 1) symbolFPriceMap[pi20026.Symbol][(int)FUTURE_LIST.PRICE] = tempList[(int)FUTURE_LIST.PRICE];
                                if (int.Parse(tempList[(int)FUTURE_LIST.OPENPRICE]) >= 1) symbolFPriceMap[pi20026.Symbol][(int)FUTURE_LIST.OPENPRICE] = tempList[(int)FUTURE_LIST.OPENPRICE];
                                if (int.Parse(tempList[(int)FUTURE_LIST.VOL]) >= 1) symbolFPriceMap[pi20026.Symbol][(int)FUTURE_LIST.VOL] = tempList[(int)FUTURE_LIST.VOL];
                            }
                            else
                            {
                                symbolFPriceMap.Add(pi20026.Symbol, tempList);
                            }
                        }
                        //選擇權 list順序"商品全碼", "月", "成交價", "開盤價", "漲跌", "成交量", "50差", "100差" 這邊算是initial訂閱前會手動觸發一次
                        else if (pi20026.Symbol.Length > 8)
                        {
                            tempList = new List<string>();
                            tempList.Add(pi20026.Symbol);
                            tempList.Add(pi20026.Symbol[pi20026.Symbol.Length - 2].ToString());
                            tempList.Add("");
                            tempList.Add(pi20026.FirstMatchPrice.ToString());
                            tempList.Add("");
                            tempList.Add((pi20026.MatchTotalQty).ToString());
                            tempList.Add("0");
                            tempList.Add("0");
                            tempList.Add("0");
                            tempList.Add("0");
                            tempList.Add(pi20026.MatchPrice.ToString());

                            //如果商品已存在在Dictionary裡則只更新 成交價、開盤價、成交量
                            if (symbolOPriceMap.ContainsKey(pi20026.Symbol))
                            {
                                if (double.Parse(tempList[(int)OPTION_LIST.PRICE]) >= 0.1) symbolOPriceMap[pi20026.Symbol][(int)OPTION_LIST.PRICE] = tempList[(int)OPTION_LIST.PRICE];
                                if (double.Parse(tempList[(int)OPTION_LIST.OPENPRICE]) >= 0.1) symbolOPriceMap[pi20026.Symbol][(int)OPTION_LIST.OPENPRICE] = tempList[(int)OPTION_LIST.OPENPRICE];
                                if (int.Parse(tempList[(int)OPTION_LIST.VOL]) >= 1) symbolOPriceMap[pi20026.Symbol][(int)OPTION_LIST.VOL] = tempList[(int)OPTION_LIST.VOL];
                            }
                            else
                            {
                                symbolOPriceMap.Add(pi20026.Symbol, tempList);
                            }
                        }

                        break;
                    case (ushort)DT.QUOTE_I101: //2020.10.26 Lynn Add:盤中漲跌停異動通知
                        PI20101 pi20101 = (PI20101)package;
                        AddInfo(pi20101.ToLog());
                        sb = new StringBuilder();
                        sb.Append("漲跌停異動通知: ");
                        sb.Append(pi20101.ToLog());
                        sb.Append("==========================");
                        AddInfo(sb.ToString());
                        break;
                    case (ushort)DT.QUOTE_INFO:
                        PI20140 pi20140 = (PI20140)package;
                        sb = new StringBuilder(Environment.NewLine);
                        sb.Append("F/O :").Append(pi20140.Market).Append(" kind:").Append(pi20140.Kind).Append(" Reason:").Append(pi20140.Reason);
                        sb.Append(" Status:").Append(pi20140.Status).Append(" Count:").Append(pi20140.Count);
                        for (int i = 0; i < pi20140.Count; i++)
                            sb.Append(pi20140.Symbols[i]).Append(",");
                        AddInfo(sb.ToString());
                        break;
                    case (ushort)DT.QUOTE_SESSION: //2017.4.10 Lynn Add: 盤別 FOR 盤後新制
                        try
                        {
                            //PI05005 pi5005 = (PI05005)package;
                            //sb = new StringBuilder(Environment.NewLine);
                            //sb.Append("F/O :").Append(pi5005.Market).Append(" Symbol:").Append(pi5005.Symbol).Append(" FallLimitPrice:").Append(pi5005.FallLimitPrice);
                            //sb.Append(" RiseLimitPrice:").Append(pi5005.RiseLimitPrice).Append(" RefPrice:").Append(pi5005.RefPrice)
                            //    .Append(" Session:").Append(pi5005.Session).Append(" Status:").Append(pi5005.Status);
                            //AddInfo(sb.ToString());
                            //dataGridView1.Rows.Add(
                            //    pi5005.Market.ToString().Trim(),
                            //    pi5005.Symbol.ToString().Trim(),
                            //    pi5005.FallLimitPrice.ToString().Trim(),
                            //    pi5005.RiseLimitPrice.ToString().Trim(),
                            //    pi5005.RefPrice.ToString().Trim(),
                            //    pi5005.Session.ToString().Trim(),
                            //    pi5005.Status.ToString().Trim());
                        }
                        catch { };
                        break;
                    case (ushort)DT.QUOTE_FOREIGN_MATCH:
                        /*if (cb40020.Checked){
                            PI40020 pi40020 = (PI40020)package;
                            sb = new StringBuilder(Environment.NewLine);
                            sb.Append("[40020]交易所:").Append(pi40020.ExchangeNm);
                            sb.Append(" 商品代號:").Append(pi40020.Symbol).Append(" 類別:").Append(pi40020.ComType);
                            sb.Append(" 盤別:").Append(pi40020.Session).Append(Environment.NewLine);
                            sb.Append(" 成交時間:").Append(pi40020.TradeTime).Append(" 價格:").Append(pi40020.Price).Append(" 口數:").Append(pi40020.Quantity).Append(Environment.NewLine);
                            sb.Append(" 總量:").Append(pi40020.TotalQty);
                            AddInfo(sb.ToString());
                        }*/
                        break;
                    case (ushort)DT.QUOTE_FOREIGN_TOTALQTY:
                        /*if (cb40020.Checked){
                            PI40023 pi40023 = (PI40023)package;
                            sb = new StringBuilder(Environment.NewLine);
                            sb.Append("[40020]交易所:").Append(pi40023.ExchangeNm);
                            sb.Append(" 商品代號:").Append(pi40023.Symbol).Append(" 類別:").Append(pi40023.ComType);
                            sb.Append(" 盤別:").Append(pi40023.Session).Append(Environment.NewLine);
                            sb.Append(" 總量:").Append(pi40023.TotalQty);
                            AddInfo(sb.ToString());
                        }*/
                        break;
                    case (ushort)DT.QUOTE_FOREIGN_CLOSEINFO:
                        //PI40021 pi40021 = (PI40021)package;
                        //sb = new StringBuilder(Environment.NewLine);
                        //sb.Append("[40021]交易所:").Append(pi40021.ExchangeNm);
                        //sb.Append(" 商品代號:").Append(pi40021.Symbol).Append(" 類別:").Append(pi40021.ComType);
                        //sb.Append(" 盤別:").Append(pi40021.Session).Append(Environment.NewLine);
                        //sb.Append(" 收盤價:").Append(pi40021.ClosePrice).Append(" 收盤量:").Append(pi40021.CloseQty).Append(Environment.NewLine);
                        //AddInfo(sb.ToString());
                        //AddInfoCover(sb.ToString());
                        break;
                    case (ushort)DT.QUOTE_FOREIGN_SETTLEMENTPRICE:
                        //PI40022 pi40022 = (PI40022)package;
                        //sb = new StringBuilder(Environment.NewLine);
                        //sb.Append("[40022]交易所:").Append(pi40022.ExchangeNm);
                        //sb.Append(" 商品代號:").Append(pi40022.Symbol).Append(" 類別:").Append(pi40022.ComType);
                        //sb.Append(" 盤別:").Append(pi40022.Session);
                        //sb.Append(" 結算價:").Append(pi40022.SettlementPrice).Append(Environment.NewLine);
                        //AddInfo(sb.ToString());
                        //AddInfoCover(sb.ToString());
                        break;
                    case (ushort)DT.QUOTE_FOREIGN_DEPTH:
                        /*PI40080 pi40080 = (PI40080)package;
                        if (cb40080.Checked){
                            sb = new StringBuilder(Environment.NewLine);
                            sb.Append("[40080]交易所:").Append(pi40080.ExchangeNm).Append(" 來源:").Append(pi40080.Source);
                            sb.Append("商品代號:").Append(pi40080.Symbol).Append(" 類別:").Append(pi40080.ComType);
                            sb.Append(" 盤別:").Append(pi40080.Session).Append(" 買賣別:").Append(pi40080.Side).Append(Environment.NewLine);
                            sb.Append(" 頂端價:").Append(pi40080.TopMark).Append(" 價格:").Append(pi40080.Price);
                            sb.Append(" 口數:").Append(pi40080.Quantity).Append("檔位:").Append(pi40080.Position).Append(Environment.NewLine);
                            AddInfo(sb.ToString());
                        }*/
                        break;
                    case (ushort)DT.QUOTE_FOREIGN_LASTPRICE:
                        //PI40026 pi40026 = (PI40026)package;
                        //sb = new StringBuilder(Environment.NewLine);
                        //sb.Append("[40026]交易所:").Append(pi40026.ExchangeNm).Append(" 類別:").Append(pi40026.ComType).Append(" 盤別:").Append(pi40026.Session).Append(" 盤別狀態:").Append(pi40026.Status);
                        //sb.Append("商品代號:").Append(pi40026.Symbol).Append(Environment.NewLine);
                        //sb.Append(" 參考價:").Append(pi40026.ReferencePrice).Append(" 昨日成交量:").Append(pi40026.YesterdayQty).Append(Environment.NewLine);
                        //sb.Append(" 當日最高價:").Append(pi40026.DayHighPrice).Append("當日最低價:").Append(pi40026.DayLowPrice).Append(Environment.NewLine);
                        //sb.Append(" 開盤價:").Append(pi40026.OpenPrice).Append(" 開盤量:").Append(pi40026.OpenQty).Append(Environment.NewLine);
                        //sb.Append(" 最新成交價:").Append(pi40026.LastPrice).Append(" 最新成交量:").Append(pi40026.LastQty).Append(" 當日成交量:").Append(pi40026.DayTradeQty).Append(Environment.NewLine);
                        //sb.Append("盤別狀態:").Append(pi40026.Status).Append(Environment.NewLine);
                        //sb.Append("昨日收盤價:").Append(pi40026.ClosePrice).Append(Environment.NewLine);   //2020.8.10 Add
                        //sb.Append("買方檔數:" + pi40026.BuyDepthCount).Append(Environment.NewLine);
                        //for (int i = 0; i < pi40026.BuyDepthCount; i++)
                        //    sb.Append("價格:").Append(pi40026.BUY_DEPTH[i].PRICE).Append("數量:").Append(pi40026.BUY_DEPTH[i].QUANTITY).Append(Environment.NewLine);
                        //sb.Append("賣方檔數:" + pi40026.SellDepthCount).Append(Environment.NewLine);
                        //for (int i = 0; i < pi40026.SellDepthCount; i++)
                        //    sb.Append("價格:").Append(pi40026.SELL_DEPTH[i].PRICE).Append("數量:").Append(pi40026.SELL_DEPTH[i].QUANTITY).Append(Environment.NewLine);
                        //sb.Append("==========================");
                        //AddInfo(sb.ToString());
                        break;
                    //2015.12 Add: 現貨指數
                    case (ushort)DT.QUOTE_FOREIGN_INDEX:
                        //PI40060 pi40060 = (PI40060)package;
                        //sb = new StringBuilder(Environment.NewLine);
                        //sb.Append("[40060]交易所:").Append(pi40060.ExchangeNm);
                        //sb.Append(" 商品代號:").Append(pi40060.Symbol).Append(" 類別:").Append(pi40060.ComType);
                        //sb.Append(" 盤別:").Append(pi40060.Session).Append(Environment.NewLine);
                        //sb.Append(" 成交時間:").Append(pi40060.TradeTime).Append(" 價格:").Append(pi40060.LastPrice).Append(Environment.NewLine);
                        //AddInfo(sb.ToString());
                        break;
                    //2018.7.25 : Lynn (I90)  台灣期貨交易所編制指數資訊揭示訊息
                    case (ushort)DT.QUOTE_I090:
                        PI20090 pi20090 = (PI20090)package;
                        sb = new StringBuilder(Environment.NewLine);
                        sb.Append("[20090]").Append(pi20090.Market).Append("IndexID=").Append(pi20090.IndexID);
                        sb.Append(" IndexTime=").Append(pi20090.IndexTime).Append(" IndexPrice=").Append(pi20090.IndexPrice);

                        AddInfo(sb.ToString());
                        break;
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine("報價元件exception: "+ex);
            }
        }

        //報價元件的狀態回報
        private void OnQuoteGetStatus(object sender, COM_STATUS staus, byte[] msg)
        {
            QuoteCom com = (QuoteCom)sender;
            string smsg = "";
            this.Invoke((Action)(() =>
            {
                if (formClose)
                {
                    FlagInvoke = false;
                    return;
                }
                FlagInvoke = true;

                switch (staus)
                {
                    case COM_STATUS.LOGIN_READY:
                        quoteConnect = true;
                        AddInfo(String.Format("Quote LOGIN_READY:[{0}]", encoding.GetString(msg)));
                        AddInfoTG(String.Format("Quote LOGIN_READY:[{0}]", encoding.GetString(msg)));
                        opLog.txtLog(String.Format("Quote LOGIN_READY:[{0}]", encoding.GetString(msg)));
                        break;
                    case COM_STATUS.LOGIN_FAIL:
                        AddInfo(String.Format("Quote LOGIN FAIL:[{0}]", encoding.GetString(msg)));
                        AddInfoTG(String.Format("Quote LOGIN FAIL:[{0}]", encoding.GetString(msg)));
                        opLog.txtLog(String.Format("Quote LOGIN FAIL:[{0}]", encoding.GetString(msg)));
                        break;
                    case COM_STATUS.LOGIN_UNKNOW:
                        AddInfo(String.Format("Quote LOGIN UNKNOW:[{0}]", encoding.GetString(msg)));
                        AddInfoTG(String.Format("Quote LOGIN UNKNOW:[{0}]", encoding.GetString(msg)));
                        opLog.txtLog(String.Format("Quote LOGIN UNKNOW:[{0}]", encoding.GetString(msg)));
                        break;
                    case COM_STATUS.CONNECT_READY:
                        smsg = "QuoteCom: [" + encoding.GetString(msg) + "] MyIP=" + quoteCom.MyIP;
                        break;
                    case COM_STATUS.CONNECT_FAIL:
                        smsg = encoding.GetString(msg);
                        AddInfo("Quote CONNECT_FAIL:" + smsg);
                        AddInfoTG("Quote CONNECT_FAIL:" + smsg);
                        opLog.txtLog("Quote CONNECT_FAIL:" + smsg);
                        break;
                    case COM_STATUS.DISCONNECTED:
                        smsg = encoding.GetString(msg);
                        //如果斷線則開此TryContinueConnect Thread重連 尚未測試
                        quoteConnect = false;
                        TryContinueConnect();
                        Console.WriteLine("quoteConnect: " + quoteConnect);
                        AddInfo("Quote DISCONNECTED:" + smsg);
                        AddInfo("quoteCom.ComStatus = " + quoteCom.ComStatus);
                        AddInfoTG("Quote DISCONNECTED:" + smsg + "  quoteCom.ComStatus = " + quoteCom.ComStatus);
                        opLog.txtLog("Quote DISCONNECTED:" + smsg + "  quoteCom.ComStatus = " + quoteCom.ComStatus);
                        break;
                    case COM_STATUS.SUBSCRIBE:
                        smsg = encoding.GetString(msg, 0, msg.Length - 1);
                        AddInfo(String.Format("Quote SUBSCRIBE:[{0}]", smsg));
                        //AddInfoTG(String.Format("Quote SUBSCRIBE:[{0}]", smsg));
                        opLog.txtLog(String.Format("Quote SUBSCRIBE:[{0}]", smsg));
                        break;
                    case COM_STATUS.UNSUBSCRIBE:
                        smsg = encoding.GetString(msg, 0, msg.Length - 1);
                        AddInfo(String.Format("Quote UNSUBSCRIBE:[{0}]", smsg));
                        //AddInfoTG(String.Format("Quote UNSUBSCRIBE:[{0}]", smsg));
                        opLog.txtLog(String.Format("Quote UNSUBSCRIBE:[{0}]", smsg));
                        break;
                    case COM_STATUS.ACK_REQUESTID:
                        long RequestId = BitConverter.ToInt64(msg, 0);
                        byte status = msg[8];
                        AddInfo("Quote Request Id BACK: " + RequestId + " Status=" + status);
                        AddInfoTG("Quote Request Id BACK: " + RequestId + " Status=" + status);
                        opLog.txtLog("Quote Request Id BACK: " + RequestId + " Status=" + status);
                        break;
                    case COM_STATUS.RECOVER_DATA:
                        smsg = encoding.GetString(msg, 1, msg.Length - 1);
                        if (!RecoverMap.ContainsKey(smsg))
                            RecoverMap.Add(smsg, 0);

                        if (msg[0] == 0)
                        {
                            RecoverMap[smsg] = 0;
                            AddInfo(String.Format("Quote開始回補 Topic:[{0}]", smsg));
                        }

                        if (msg[0] == 1)
                        {
                            AddInfo(String.Format("Quote結束回補 Topic:[{0} 筆數:{1}]", smsg, RecoverMap[smsg]));
                        }
                        break;
                }

                FlagInvoke = false;
                if (formClose)
                {
                    return;
                }
            }));
            com.Processed();
        }

        private void OnRecoverStatusQuote(object sender, string Topic, RECOVER_STATUS status, uint RecoverCount)
        {
            if (this.InvokeRequired) {
                Intelligence.OnRecover_EvenHandler d = new Intelligence.OnRecover_EvenHandler(OnRecoverStatusQuote);
                this.Invoke(d, new object[] { sender, Topic, status, RecoverCount });
                return;
            }
            QuoteCom com = (QuoteCom)sender;
            switch (status)
            {
                case RECOVER_STATUS.RS_DONE:        //回補資料結束
                    AddInfo(String.Format("Quote結束回補 Topic:[{0}]{1}", Topic, RecoverCount));
                    break;
                case RECOVER_STATUS.RS_BEGIN:       //開始回補資料
                    AddInfo(String.Format("Quote開始回補 Topic:[{0}]", Topic));
                    break;
            }
        }

        private void AddInfo(string msg)
        {
            if (this.msgBox.InvokeRequired) {
                SetAddInfoCallBack d = new SetAddInfoCallBack(AddInfo);
                this.Invoke(d, new object[] { msg });
            }
            else
            {
                string fMsg = String.Format("[{0}] {1} {2}", DateTime.Now.ToString("HH:mm:ss:ffff"), msg, Environment.NewLine);
                try
                {
                    if (msgBox.TextLength > 50000) msgBox.ResetText();
                    msgBox.AppendText(fMsg);
                    //if(fMsg.Length <= 100)  tl.SendNotify(fMsg);
                }
                catch { };
            }
        }

        private void AddInfoCover(string msg)
        {
            if (this.txtRecover.InvokeRequired)
            {
                SetAddInfoCallBack d = new SetAddInfoCallBack(AddInfoCover);
                this.Invoke(d, new object[] { msg });
            }
            else
            {
                string fMsg = String.Format("{0} {1}", msg, Environment.NewLine);
                try
                {
                    if (txtRecover.TextLength > 5000) txtRecover.ResetText();
                    txtRecover.AppendText(fMsg);
                }
                catch { };
            }
        }

        private void OnRecoverStatus(object sender, string Topic, RECOVER_STATUS status, uint RecoverCount)
        {
            if (this.InvokeRequired)
            {
                Smart.OnRecover_EvenHandler d = new Smart.OnRecover_EvenHandler(OnRecoverStatus);
                this.Invoke(d, new object[] { sender, Topic, status, RecoverCount });
                return;
            }

            TaiFexCom com = (TaiFexCom)sender;
            switch (status)
            {
                case RECOVER_STATUS.RS_DONE:        //回補資料結束
                    if (RecoverCount == 0)
                        AddInfo(String.Format("Trade結束回補 Topic:[{0}]", Topic));
                    else AddInfo(String.Format("Trade結束回補 Topic:[{0} 筆數:{1}]", Topic, RecoverCount));
                    break;
                case RECOVER_STATUS.RS_BEGIN:       //開始回補資料
                    AddInfo(String.Format("Trade開始回補 Topic:[{0}]", Topic));
                    break;
            }
        }

        private void OnRcvServerTime(Object sender, DateTime serverTime, int ConnQuality)
        {
            if (this.InvokeRequired)
            {
                Smart.OnRcvServerTime_EventHandler d = new Smart.OnRcvServerTime_EventHandler(OnRcvServerTime);
                this.Invoke(d, new object[] { sender, serverTime, ConnQuality });
                return;
            }
            //ConnQuality : 本次與上次 HeatBeat 之時間差(milliseconds) 
            /*if (ConnQuality > 100)
                pbConnQuality.Value = 0;
            else
                pbConnQuality.Value = 100 - ConnQuality;
            labelServerTime.Text = String.Format("{0:hh:mm:ss.fff}", serverTime);
            labelConnQuality.Text = "[" + ConnQuality + "]";*/
        }

        private void OnGetStatus(object sender, COM_STATUS staus, byte[] msg)
        {
            TaiFexCom com = (TaiFexCom)sender;
            if (this.InvokeRequired)
            {
                Smart.OnGetStatus_EventHandler d = new Smart.OnGetStatus_EventHandler(OnGetStatus);
                this.Invoke(d, new object[] { sender, staus, msg });
                return;
            }
            OnGetStatusUpdateUI(sender, staus, msg);
        }

        //下單元件狀態回報
        private void OnGetStatusUpdateUI(object sender, COM_STATUS staus, byte[] msg)
        {
            TaiFexCom com = (TaiFexCom)sender;
            string smsg = null;
            switch (staus)
            {
                case COM_STATUS.LOGIN_READY:          //登入成功
                    string[] tempSplit;

                    AddInfo("Trade登入成功:" + com.Accounts);
                    tradeConnect = true;
                    Console.WriteLine("tradeConnect: " + tradeConnect);
                    AddInfoTG("Trade登入成功:" + com.Accounts);
                    opLog.txtLog("Trade登入成功:" + com.Accounts);

                    tempSplit = com.Accounts.Split(',');

                    if( tempSplit != null && int.Parse(settings.SubAccount) == 1)
                    {
                        accInfo.BrokerID = tempSplit[0];
                        accInfo.TbTAccount = tempSplit[1];
                        tbCompany.Text = accInfo.BrokerID;
                        tbAccount.Text = accInfo.TbTAccount;
                    }
                    else if (tempSplit != null && int.Parse(settings.SubAccount) == 2)
                    {
                        accInfo.BrokerID = tempSplit[0];
                        accInfo.TbTAccount = tempSplit[10];
                        tbCompany.Text = accInfo.BrokerID;
                        tbAccount.Text = accInfo.TbTAccount;
                    }
                    break;
                case COM_STATUS.LOGIN_FAIL:             //登入失敗
                    AddInfo(String.Format("Trade登入失敗:[{0}]", encoding.GetString(msg)));
                    AddInfoTG(String.Format("Trade登入失敗:[{0}]", encoding.GetString(msg)));
                    opLog.txtLog(String.Format("Trade登入失敗:[{0}]", encoding.GetString(msg)));
                    break;
                case COM_STATUS.LOGIN_UNKNOW:       //登入狀態不明
                    AddInfo(String.Format("Trade登入狀態不明:[{0}]", encoding.GetString(msg)));
                    AddInfoTG(String.Format("Trade登入狀態不明:[{0}]", encoding.GetString(msg)));
                    opLog.txtLog(String.Format("Trade登入狀態不明:[{0}]", encoding.GetString(msg)));
                    break;
                case COM_STATUS.CONNECT_READY:      //連線成功
                    smsg = "Trade伺服器" + tfcom.ServerHost + ":" + tfcom.ServerPort + Environment.NewLine +
                              "Trade伺服器回應: [" + encoding.GetString(msg) + "]" + Environment.NewLine +
                              "本身為" + ((tfcom.isInternal) ? "內" : "外") + "部 IP:" + tfcom.MyIP;
                    AddInfo(smsg);
                    AddInfoTG(smsg);
                    opLog.txtLog(smsg);
                    break;
                case COM_STATUS.CONNECT_FAIL:       //連線失敗
                    smsg = encoding.GetString(msg);
                    AddInfo("Trade連線失敗:" + smsg + " " + tfcom.ServerHost + ":" + tfcom.ServerPort);
                    AddInfoTG("Trade連線失敗:" + smsg + " " + tfcom.ServerHost + ":" + tfcom.ServerPort);
                    opLog.txtLog("Trade連線失敗:" + smsg + " " + tfcom.ServerHost + ":" + tfcom.ServerPort);
                    break;
                case COM_STATUS.DISCONNECTED:       //斷線
                    smsg = encoding.GetString(msg);
                    tradeConnect = false;
                    TryContinueConnect();
                    Console.WriteLine("tradeConnect: " + tradeConnect);
                    AddInfo("Trade斷線:" + smsg);
                    AddInfoTG("Trade斷線:" + smsg);
                    opLog.txtLog("Trade斷線:" + smsg);
                    break;
                case COM_STATUS.SUBSCRIBE:
                    com.WriterLog("msg.Length=" + msg.Length);
                    smsg = encoding.GetString(msg);
                    com.WriterLog(String.Format("Trade註冊:[{0}]", smsg));
                    AddInfo(String.Format("Trade註冊:[{0}]", smsg));
                    AddInfoTG(String.Format("Trade註冊:[{0}]", smsg));
                    opLog.txtLog(String.Format("Trade註冊:[{0}]", smsg));
                    break;
                case COM_STATUS.UNSUBSCRIBE:
                    smsg = encoding.GetString(msg);
                    AddInfo(String.Format("Trade取消註冊:[{0}]", smsg));
                    AddInfoTG(String.Format("Trade取消註冊:[{0}]", smsg));
                    opLog.txtLog(String.Format("Trade取消註冊:[{0}]", smsg));
                    break;
                case COM_STATUS.ACK_REQUESTID:          //下單或改單第一次回覆
                    long RequestId = BitConverter.ToInt64(msg, 0);
                    byte status = msg[8];
                    string[] tmp = new string[2];
                    tmp[0] = (status == 1) ? "收單" : "失敗";
                    tmp[1] = "";
                    if (orderFirstRPT.ContainsKey(RequestId))
                    {
                        opLog.txtLog("改單 || 刪單 第一次回覆");
                    }
                    else
                    {
                        orderFirstRPT.Add(RequestId, tmp);
                    }
                    AddInfo("序號回覆: " + RequestId + " 狀態=" + ((status == 1) ? "收單" : "失敗"));
                    AddInfoTG("序號回覆: " + RequestId + " 狀態=" + ((status == 1) ? "收單" : "失敗"));
                    opLog.txtLog("序號回覆: " + RequestId + " 狀態=" + ((status == 1) ? "收單" : "失敗"));
                    break;
                case COM_STATUS.QUEUE_WARNING:
                    smsg = encoding.GetString(msg);
                    AddInfo(smsg);
                    //lbQueueWarn.Text = smsg;
                    AddInfoTG(smsg);
                    opLog.txtLog(smsg);
                    break;
            }
        }

        private void OnRcvMessage(object sender, PackageBase package)
        {
            List<string> tempList = new List<string>();

            if (this.InvokeRequired)
            {
                Smart.OnRcvMessage_EventHandler d = new Smart.OnRcvMessage_EventHandler(OnRcvMessage);
                this.Invoke(d, new object[] { sender, package });
                return;
            }

            switch ((DT)package.DT)
            {
                #region Login 
                case DT.LOGIN:
                    P001503 p1503 = (P001503)package;
                    if (p1503.Code != 0)
                        AddInfo("登入失敗 CODE = " + p1503.Code + " " + tfcom.GetMessageMap(p1503.Code));
                    else
                    {
                        AddInfo("登入成功 ");
                        AddInfo(p1503.ToLog());
                    }
                    break;
                #endregion 

                //2020.10.13 Test....
                case DT.PreviousLogin:
                    P001507 p1507 = (P001507)package;
                    AddInfo("前次登入:" + p1507.ToLog());
                    break;
                #region 公告 2014.12.12 ADD
                case DT.NOTICE_RPT:
                    P001701 p1701 = (P001701)package;
                    AddInfo(p1701.ToLog());
                    break;    //公告(被動查詢) 
                case DT.NOTICE:   //公告(主動)
                    P001702 p1702 = (P001702)package;
                    AddInfo(p1702.ToLog());
                    break;
                #endregion

                #region 國內下單回報
                case DT.FUT_ORDER_ACK:   //下單第二回覆
                    PT02002 p2002 = (PT02002)package;
                    string[] tmp = new string[2];
                    if (orderFirstRPT.ContainsKey(p2002.RequestId))
                    {
                        opLog.txtLog("有經過第一回覆");
                    }
                    else
                    {
                        opLog.txtLog("跳過第一回覆");
                        tmp[0] = "收單";
                        tmp[1] = "";
                        orderFirstRPT.Add(p2002.RequestId, tmp);
                    }
                    orderFirstRPT[p2002.RequestId][1] = p2002.CNT;
                    AddInfo("Order Second report!");
                    opLog.txtLog("Order Second report!");
                        tempList = new List<string>();
                        tempList.Add("0");//狀態紀錄  -1:舊的單子 0:尚未收到委託回報 1:收到委託回報 2:成交
                        tempList.Add(p2002.CNT);  //CNT
                        tempList.Add(p2002.OrderNo);  //單號
                        tempList.Add("");//交易日期
                        tempList.Add("");//回報時間
                        tempList.Add("");//Symbol
                        tempList.Add("");//買賣別
                        tempList.Add("");//委託價
                        tempList.Add("");//改後數量
                        tempList.Add("");//新/平倉
                        tempList.Add(p2002.ErrorCode.ToString());
                        tempList.Add("");//error message
                        orderRPT.Add(p2002.CNT, tempList);
                    AddInfo(p2002.ToLog());//+ "訊息:" + tfcom.GetMessageMap(p2002.ErrorCode));
                    AddInfoTG(p2002.ToLog());
                    opLog.txtLog(p2002.ToLog());
                    break;
                case DT.FUT_ORDER_RPT: //委託回報
                    PT02010 p2010 = (PT02010)package;
                    //AddInfo("RCV 2010 [" + p2010.CNT + "," + p2010.OrderNo + "]");
                    string[] row2010 = { p2010.OrderFunc.ToString(), p2010.FrontOffice.ToString(), p2010.BrokerId, p2010.Account, p2010.OrderNo, p2010.TradeDate, p2010.ReportTime, p2010.ClientOrderTime, p2010.WebID, p2010.CNT, p2010.TimeInForce.ToString(), p2010.Symbol, p2010.Side.ToString(), p2010.PriceMark.ToString(), p2010.Price, p2010.PositionEffect.ToString(), p2010.BeforeQty, p2010.AfterQty, p2010.Code, p2010.ErrMsg, p2010.Trader };
                    //dgv2010.Rows.Add(row2010);

                    tempList = new List<string>();
                    tempList.Add(p2010.CNT);  //CNT
                    tempList.Add(p2010.OrderNo);  //單號
                    tempList.Add(p2010.TradeDate);//交易日期
                    tempList.Add(p2010.ReportTime);//回報時間
                    tempList.Add(p2010.Symbol);//Symbol
                    tempList.Add(p2010.Side.ToString());//買賣別
                    tempList.Add(p2010.Price);//委託價
                    tempList.Add(p2010.AfterQty);//改後數量
                    tempList.Add(p2010.PositionEffect.ToString());//新/平倉
                    tempList.Add(p2010.Code);
                    tempList.Add(p2010.ErrMsg);

                    if (orderRPT.ContainsKey(p2010.CNT))
                    {
                        AddInfo("Order last report!");
                        opLog.txtLog("Order last report!");
                        if ( String.IsNullOrEmpty(orderRPT[p2010.CNT][(int)ORDERRPT_LIST.TRADEDATE]) )
                            tempList.Insert((int)ORDERRPT_LIST.ORDERSTATUS, "1");//狀態紀錄  -1:之前的委託 0:尚未收到委託回報 1:收到委託回報 2:成交
                        else
                            tempList.Insert((int)ORDERRPT_LIST.ORDERSTATUS, orderRPT[p2010.CNT][(int)ORDERRPT_LIST.ORDERSTATUS]);
                        orderRPT[p2010.CNT] = tempList;
                    }
                    else
                    {   //舊的委託回報，或者沒經過第二回覆
                        tempList.Insert((int)ORDERRPT_LIST.ORDERSTATUS, "-1");//狀態紀錄  -1:之前的委託 0:尚未收到委託回報 1:收到委託回報 2:成交
                        orderRPT.Add(p2010.CNT, tempList);
                    }
                    AddInfo(p2010.ToLog());
                    AddInfoTG(p2010.ToLog());
                    opLog.txtLog(p2010.ToLog());
                    break;
                case DT.FUT_DEAL_RPT:   //成交回報
                    PT02011 p2011 = (PT02011)package;

                    tempList = new List<string>();
                    tempList.Add(p2011.CNT);  //CNT
                    tempList.Add(p2011.OrderNo);  //單號
                    tempList.Add(p2011.TradeDate);//交易日期
                    tempList.Add(p2011.ReportTime);//回報時間
                    tempList.Add(p2011.Symbol);
                    tempList.Add(p2011.Side.ToString());//買賣別
                    tempList.Add(p2011.Market.ToString());//1:期貨單式 2:選擇權單式 3:選擇權複式 4:期貨複式
                    tempList.Add(p2011.DealPrice); //成交價格
                    tempList.Add(p2011.DealQty); //原委託數量
                    tempList.Add(p2011.CumQty); //已成交總量
                    tempList.Add(p2011.LeaveQty); //剩餘未成交數量
                    tempList.Add(p2011.Symbol1);
                    tempList.Add(p2011.DealPrice1);
                    tempList.Add(p2011.Qty1);
                    tempList.Add(p2011.BS1.ToString());
                    tempList.Add(p2011.Symbol2);
                    tempList.Add(p2011.DealPrice2);
                    tempList.Add(p2011.Qty2);
                    tempList.Add(p2011.BS2.ToString());

                    if (orderRPT.ContainsKey(p2011.CNT))
                    {
                        if (String.Compare(orderRPT[p2011.CNT][(int)ORDERRPT_LIST.ORDERSTATUS], "1") == 0) //從委託回報再送下來的，為新成交
                        {
                            orderRPT[p2011.CNT][(int)ORDERRPT_LIST.ORDERSTATUS] = "2";//更新orderRPT的狀態紀錄變為成交  -1:之前的委託 0:尚未收到委託回報 1:收到委託回報 2:成交
                            if (dealRPT.ContainsKey(p2011.CNT))
                            {   //新成交時，照理說dealRPT需要新建本來會是空的，所以不會跑至這裡
                                AddInfo("Deal report1!");
                                tempList.Insert((int)DEALRPT_LIST.ORDERSTATUS, "1");//當作委託、成交狀態紀錄  0:尚未成交 1:新成交 -1:之前的成交單
                                dealRPT[p2011.CNT] = tempList;
                            }
                            else
                            {
                                AddInfo("Deal report1!");
                                tempList.Insert((int)DEALRPT_LIST.ORDERSTATUS, "1");//當作委託、成交狀態紀錄  0:尚未成交 1:新成交 -1:之前的成交單
                                dealRPT.Add(p2011.CNT, tempList);
                            }
                        }
                        else if (String.Compare(orderRPT[p2011.CNT][(int)ORDERRPT_LIST.ORDERSTATUS], "-1") == 0)//舊委託回報
                        {
                            if (dealRPT.ContainsKey(p2011.CNT))
                            {   //通常拿到舊的委託回報，同一個CNT的成交回報也會是舊的，所以不會跑至這裡
                                AddInfo("Deal report2!");
                                tempList.Insert((int)DEALRPT_LIST.ORDERSTATUS, "-1");//當作委託、成交狀態紀錄  0:尚未成交 1:新成交 -1:之前的成交單
                                dealRPT[p2011.CNT] = tempList;
                            }
                            else
                            {
                                AddInfo("Deal report2!");
                                tempList.Insert((int)DEALRPT_LIST.ORDERSTATUS, "-1");//當作委託、成交狀態紀錄  0:尚未成交 1:新成交 -1:之前的成交單
                                dealRPT.Add(p2011.CNT, tempList);
                            }
                        }
                        else if (String.Compare(orderRPT[p2011.CNT][(int)ORDERRPT_LIST.ORDERSTATUS], "2") == 0)
                        {
                            if (dealRPT.ContainsKey(p2011.CNT))
                            {   //跑至這裡代表重複收到兩次成交回報
                                AddInfo("Deal report3!");
                                tempList.Insert((int)DEALRPT_LIST.ORDERSTATUS, "-1");//當作委託、成交狀態紀錄  0:尚未成交 1:新成交 -1:之前的成交單
                                dealRPT[p2011.CNT] = tempList;
                            }
                            else
                            {   //跑至這裡代表重複收到兩次成交回報
                                AddInfo("Deal report3!");
                                tempList.Insert((int)DEALRPT_LIST.ORDERSTATUS, "-1");//當作委託、成交狀態紀錄  0:尚未成交 1:新成交 -1:之前的成交單
                                dealRPT.Add(p2011.CNT, tempList);
                            }
                        }
                        else
                        {
                            AddInfo("跳過委託回報!");
                            if (dealRPT.ContainsKey(p2011.CNT))
                            {   //跑至這裡代表跳過委託回報
                                AddInfo("Deal report4!");
                                tempList.Insert((int)DEALRPT_LIST.ORDERSTATUS, "-1");//當作委託、成交狀態紀錄  0:尚未成交 1:新成交 -1:之前的成交單
                                dealRPT[p2011.CNT] = tempList;
                            }
                            else
                            {   //跑至這裡代表跳過委託回報
                                AddInfo("Deal report4!");
                                tempList.Insert((int)DEALRPT_LIST.ORDERSTATUS, "-1");//當作委託、成交狀態紀錄  0:尚未成交 1:新成交 -1:之前的成交單
                                dealRPT.Add(p2011.CNT, tempList);
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine("不知道哪來的單子，沒經過委託回報!");
                        if (dealRPT.ContainsKey(p2011.CNT))
                        {   //跑至這裡代表跳過委託回報
                            AddInfo("Deal report5!");
                            tempList.Insert((int)DEALRPT_LIST.ORDERSTATUS, "X");//當作委託、成交狀態紀錄  0:尚未成交 1:新成交 -1:之前的成交單
                            dealRPT[p2011.CNT] = tempList;
                        }
                        else
                        {   //跑至這裡代表跳過委託回報
                            AddInfo("Deal report5!");
                            tempList.Insert((int)DEALRPT_LIST.ORDERSTATUS, "X");//當作委託、成交狀態紀錄  0:尚未成交 1:新成交 -1:之前的成交單
                            dealRPT.Add(p2011.CNT, tempList);
                        }
                    }
                    AddInfo("RCV 2011 [" + p2011.OrderNo + "]" + " CNT [" + p2011.CNT + "]");
                    //AddInfoTG("RCV 2011 [" + p2011.OrderNo + "]" + " CNT [" + p2011.CNT + "]");
                    opLog.txtLog("RCV 2011 [" + p2011.OrderNo + "]" + " CNT [" + p2011.CNT + "]");
                    string[] row2011 = { p2011.OrderFunc.ToString(), p2011.FrontOffice.ToString(), p2011.BrokerId, p2011.Account, p2011.OrderNo, p2011.TradeDate, p2011.ReportTime, p2011.WebID, p2011.CNT, p2011.Symbol, p2011.Side.ToString(), p2011.Market.ToString(), p2011.DealPrice, p2011.DealQty, p2011.CumQty, p2011.LeaveQty, p2011.MarketNo, p2011.Symbol1, p2011.DealPrice1, p2011.Qty1, p2011.BS1.ToString(), p2011.Symbol2, p2011.DealPrice2, p2011.Qty2, p2011.BS2.ToString(), p2011.Trader };
                    //dgv2011.Rows.Add(row2011);
                    break;
                #endregion 

                #region 國外下單回報
                case DT.FFUT_ORDER_ACK:   //下單第二回覆
                    PT03302 p3302 = (PT03302)package;
                    AddInfo(p3302.ToLog());

                    break;
                case DT.FFUT_ORDER_RPT: //委託回報
                    PT03310 p3310 = (PT03310)package;
                    AddInfo("RCV 3310 [" + p3310.CNT + "," + p3310.ASorderNo + "]");
                    string[] row3310 = { p3310.OrderFunc.ToString(), p3310.Exchange, p3310.FCM, p3310.FFUT_ACCOUNT, p3310.ORDNO, p3310.ASorderNo, p3310.BrokerId, p3310.Account, p3310.AE, p3310.TradeDate, p3310.ReportTime, p3310.WEBID, p3310.SOURCE, p3310.OrgCnt, p3310.CNT, p3310.Symbol, p3310.ComYM, p3310.StrikePrice, p3310.CP.ToString(), p3310.BS.ToString(), p3310.TimeInForce.ToString(), p3310.PriceFlag.ToString(), p3310.PositionEffect.ToString(), p3310.DayTrade.ToString(), p3310.Price, p3310.StopPrice, p3310.BeforeQty, p3310.AfterQty, p3310.KeyIn, p3310.ErrCode, p3310.ErrMsg };
                    //dgv3310.Rows.Add(row3310);
                    break;
                case DT.FFUT_DEAL_RPT:   //成交回報
                    PT03311 p3311 = (PT03311)package;
                    AddInfo("RCV 3311 [" + p3311.CNT + "]");
                    string[] row3311 = { p3311.OrderFunc.ToString(), p3311.Exchange, p3311.FCM, p3311.FFUT_ACCOUNT, p3311.ORDNO, p3311.ASorderNo, p3311.BrokerId, p3311.Account, p3311.AE, p3311.TradeDate, p3311.ReportTime, p3311.WEBID, p3311.SOURCE, p3311.CNT, p3311.Symbol, p3311.ComYM, p3311.StrikePrice, p3311.CP.ToString(), p3311.BS.ToString(), p3311.TimeInForce.ToString(), p3311.PriceFlag.ToString(), p3311.PositionEffect.ToString(), p3311.DayTrade.ToString(), p3311.DealPrice, p3311.AvgPrice, p3311.DealQty, p3311.PATSNo, p3311.LeavesQty, p3311.CumQty, p3311.KeyIn };
                    //dgv3311.Rows.Add(row3311);
                    break;

                case DT.FFUT_ORDER_RPT2: //委託回報-複式單
                    PT03314 p3314 = (PT03314)package;
                    AddInfo("RCV 3314 [" + p3314.CNT + "," + p3314.ASorderNo + "]");
                    string[] row3314 = { p3314.OrderFunc.ToString(), p3314.Exchange, p3314.FCM, p3314.FFUT_ACCOUNT, p3314.ORDNO, p3314.ASorderNo, p3314.BrokerId, p3314.Account, p3314.AE, p3314.TradeDate, p3314.ReportTime, p3314.WEBID, p3314.SOURCE, p3314.OrgCnt, p3314.CNT, p3314.Symbol, p3314.ComYM, p3314.StrikePrice, p3314.CP.ToString(), p3314.BS.ToString(), p3314.Symbol2, p3314.ComYM2, p3314.StrikePrice2, p3314.CP2.ToString(), p3314.BS2.ToString(), p3314.TimeInForce.ToString(), p3314.PriceFlag.ToString(), p3314.PositionEffect.ToString(), p3314.DayTrade.ToString(), p3314.Price, p3314.StopPrice, p3314.BeforeQty, p3314.AfterQty, p3314.OrgKeyin, p3314.KeyIn, p3314.ErrCode, p3314.ErrMsg };
                    //dgv3314.Rows.Add(row3314);
                    break;
                case DT.FFUT_DEAL_RPT2:   //成交回報-複式單
                    PT03315 p3315 = (PT03315)package;
                    AddInfo("RCV 3315 [" + p3315.CNT + "]");
                    string[] row3315 = { p3315.OrderFunc.ToString(), p3315.Exchange, p3315.FCM, p3315.FFUT_ACCOUNT, p3315.ORDNO, p3315.ASorderNo, p3315.BrokerId, p3315.Account, p3315.AE, p3315.TradeDate, p3315.ReportTime, p3315.WEBID, p3315.SOURCE, p3315.CNT, p3315.Symbol, p3315.ComYM, p3315.StrikePrice, p3315.CP.ToString(), p3315.BS.ToString(), p3315.TimeInForce.ToString(), p3315.PriceFlag.ToString(), p3315.PositionEffect.ToString(), p3315.DayTrade.ToString(), p3315.DealPrice, p3315.AvgPrice, p3315.DealQty, p3315.PATSNo, p3315.LeavesQty, p3315.CumQty, p3315.KeyIn };
                    //dgv3315.Rows.Add(row3315);
                    break;

                #endregion

                #region 帳務查詢
                case DT.FINANCIAL_COVER_TRADER: //1614 
                    P001614 p1614 = (P001614)package;
                    AddInfo("RCV 1614 [" + p1614.Rows + "]");
                    //dgv1614.Rows.Clear();
                    if (p1614.Rows > 0)
                    {
                        foreach (P001614_2 p1614b in p1614.p001614_2)
                        {
                            string[] row1614 = { p1614b.BrokerId, p1614b.Account, p1614b.Group, p1614b.Trader, p1614b.Exchange, p1614b.ComID, p1614b.ComYM, p1614b.StrikePrice, p1614b.CP, p1614b.CURRENCY, p1614b.PRTLOS, p1614b.CTAXAMT, p1614b.ORIGNFEE, p1614b.OSPRTLOS, p1614b.Qty, p1614b.OSPRTLOS_TWD };
                            //dgv1614.Rows.Add(row1614);
                        }
                    }
                    break;
                case DT.INVENTORY_TRADER: //1616
                    P001616 p1616 = (P001616)package;
                    //dgv1616.Rows.Clear();
                    AddInfo("RCV 1616 [" + p1616.Rows + "]");
                    if (p1616.Rows > 0)
                    {
                        callSum = 0;
                        putSum = 0;
                        callCost=0;
                        putCost=0;
                        foreach (P001616_2 p1616b in p1616.p001616_2)
                        {
                            string[] row1616 = { p1616b.BrokerId, p1616b.Account, p1616b.Group, p1616b.Trader, p1616b.Exchange, p1616b.ComType, p1616b.ComID, p1616b.ComYM, p1616b.StrikePrice, p1616b.CP, p1616b.BS, p1616b.DeliveryDate, p1616b.Currency, p1616b.OTQty, p1616b.TrdPrice, p1616b.MPrice, p1616b.PRTLOS, p1616b.DealPrice };
                            //dgv1616.Rows.Add(row1616);
                            if (String.Compare(p1616b.ComType, "O") == 0 && String.Compare(p1616b.CP, "C") == 0 && String.Compare(p1616b.BS, "S") == 0)
                            {
                                callSum += decimal.Parse(p1616b.MPrice)* Convert.ToDecimal(p1616b.OTQty);
                                callCost += decimal.Parse(p1616b.DealPrice) * Convert.ToDecimal(p1616b.OTQty);
                            }
                            else if (String.Compare(p1616b.ComType, "O") == 0 && String.Compare(p1616b.CP, "C") == 0 && String.Compare(p1616b.BS, "B") == 0)
                            {
                                callSum -= decimal.Parse(p1616b.MPrice) * Convert.ToDecimal(p1616b.OTQty);
                                callCost -= decimal.Parse(p1616b.DealPrice) * Convert.ToDecimal(p1616b.OTQty);
                            }
                            else if (String.Compare(p1616b.ComType, "O") == 0 && String.Compare(p1616b.CP, "P") == 0 && String.Compare(p1616b.BS, "S") == 0)
                            {
                                putSum += decimal.Parse(p1616b.MPrice) * Convert.ToDecimal(p1616b.OTQty);
                                putCost += decimal.Parse(p1616b.DealPrice) * Convert.ToDecimal(p1616b.OTQty);
                            }
                            else if (String.Compare(p1616b.ComType, "O") == 0 && String.Compare(p1616b.CP, "P") == 0 && String.Compare(p1616b.BS, "B") == 0)
                            {
                                putSum -= decimal.Parse(p1616b.MPrice) * Convert.ToDecimal(p1616b.OTQty);
                                putCost -= decimal.Parse(p1616b.DealPrice) * Convert.ToDecimal(p1616b.OTQty);
                            }
                        }
                    }
                    break;
                case DT.INVENTORY_DETAIL_TRADER:  //1618
                    P001618 p1618 = (P001618)package;
                    dgv1618.Rows.Clear();
                    AddInfo("RCV 1618 [" + p1618.Rows + "]");
                    Dictionary<string, List<string>> tmpstockRPT = new Dictionary<string, List<string>>();
                    List<List<string>> tmpstockAggre = new List<List<string>>();
                    //stockRPT.Clear();
                    //stockAggre.Clear();
                    foreach (string key in symbolOPriceMap.Keys)
                    {
                        symbolOPriceMap[key][(int)OPTION_LIST.STOCK] = "";
                    }
                    if (p1618.Rows > 0)
                    {
                        int i = 0;
                        foreach (P001618_2 p1618b in p1618.p001618_2)
                        {
                            string[] row1618 = { p1618b.BrokerId, p1618b.Account, p1618b.Group, p1618b.Trader, p1618b.Exchange, p1618b.SeqNo, p1618b.FcmActNo, p1618b.TradeType, p1618b.FCM, p1618b.DeliveryDate, p1618b.CloseDate, p1618b.WEB, p1618b.Cnt, p1618b.OrdNo, p1618b.MarketNo, p1618b.sNo, p1618b.TradeDate, p1618b.ComID, p1618b.BS.ToString(), p1618b.ComType.ToString(), p1618b.CP.ToString(), p1618b.StrikePrice, p1618b.ComYM, p1618b.Qty, p1618b.TrdPrice, p1618b.MPrice, p1618b.PRTLOS, p1618b.InitialMargin, p1618b.MTMargin, p1618b.Currency, p1618b.DealPrice, p1618b.MixQty, p1618b.DayTrade.ToString(), p1618b.SPREAD.ToString(), p1618b.spKey.ToString(), p1618b.OrdNo2, p1618b.MarketNo2, p1618b.sNo2, p1618b.TradeDate2, p1618b.ComID2, p1618b.BS2.ToString(), p1618b.ComType2.ToString(), p1618b.CP2.ToString(), p1618b.StrikePrice2, p1618b.ComYM2, p1618b.Qty2, p1618b.TrdPrice2, p1618b.MPrice2, p1618b.PRTLOS2, p1618b.InitialMargin2, p1618b.MTMargin2, p1618b.Currency2, p1618b.DealPrice2, p1618b.MixQty2, p1618b.DayTrade2.ToString() };
                            dgv1618.Rows.Add(row1618);
                            //Console.WriteLine("p1618b.SeqNo " + p1618b.SeqNo);
                            if (!tmpstockRPT.ContainsKey(p1618b.SeqNo))
                            {
                                //AddInfoTG("庫存陣列stockRPT可能沒有清除，請檢查!");
                                //opLog.txtLog("庫存陣列stockRPT可能沒有清除，請檢查!");
                                tmpstockRPT.Add(p1618b.SeqNo, row1618.ToList());
                            }
                            else tmpstockRPT[p1618b.SeqNo] = row1618.ToList();

                            List<string> oneStock = new List<string>();
                            if (!string.IsNullOrEmpty(p1618b.Qty) && !string.IsNullOrEmpty(p1618b.Qty2))
                            {
                                oneStock.Add(Math.Round(decimal.Parse(p1618b.StrikePrice2),0).ToString() + p1618b.CP2.ToString());//履約價 ex:18550C
                                oneStock.Add(Math.Round(decimal.Parse(p1618b.MPrice2) - decimal.Parse(p1618b.MPrice), 1).ToString());//市價和 
                                oneStock.Add(Math.Round(decimal.Parse(p1618b.DealPrice2) - decimal.Parse(p1618b.DealPrice),1).ToString());//成本和
                                oneStock.Add(Math.Round(decimal.Parse(oneStock[2]) - decimal.Parse(oneStock[1]),1).ToString());//損益和
                                oneStock.Add(Math.Round(decimal.Parse(p1618b.MPrice2),1).ToString() + "/" + Math.Round(decimal.Parse(p1618b.MPrice),1).ToString());//市價  第二隻腳/第一隻腳
                                oneStock.Add(Math.Round(decimal.Parse(p1618b.DealPrice2), 1).ToString() + "/" + Math.Round(decimal.Parse(p1618b.DealPrice), 1).ToString());//成本 第二隻腳/第一隻腳
                                oneStock.Add(Math.Round(decimal.Parse(p1618b.Qty2),1).ToString());//數量
                                tmpstockAggre.Add(oneStock);
                                //Console.Write("tmpstockAggre[i]: ");
                                //foreach (string a in tmpstockAggre[i])
                                //{
                                //    Console.Write(a+" ");
                                //}
                                //Console.WriteLine("");
                                i++;
                            }

                            //在option的dictionary裡加入庫存
                            char[] monthSymbol = new char[2];
                            string combineSymbol1 = "";
                            string combineSymbol2 = "";
                            CalculateComID.FindSymbol(int.Parse(p1618b.ComYM.Substring(4, 2)), monthSymbol);
                            if (!string.IsNullOrEmpty(p1618b.Qty))
                            {
                                if (p1618b.CP == 'C')
                                {
                                    combineSymbol1 = p1618b.ComID;
                                    combineSymbol1 += Math.Round(decimal.Parse(p1618b.StrikePrice),0).ToString();
                                    combineSymbol1 += monthSymbol[0].ToString() + p1618b.ComYM[3].ToString();
                                }
                                else
                                {
                                    combineSymbol1 = p1618b.ComID;
                                    combineSymbol1 += Math.Round(decimal.Parse(p1618b.StrikePrice), 0).ToString();
                                    combineSymbol1 += monthSymbol[1].ToString() + p1618b.ComYM[3].ToString();
                                }
                                if(!string.IsNullOrEmpty(combineSymbol1) && symbolOPriceMap.ContainsKey(combineSymbol1))
                                {
                                    if (string.IsNullOrEmpty(symbolOPriceMap[combineSymbol1][(int)OPTION_LIST.STOCK])) symbolOPriceMap[combineSymbol1][(int)OPTION_LIST.STOCK] = "0";
                                    if (p1618b.BS == 'S') symbolOPriceMap[combineSymbol1][(int)OPTION_LIST.STOCK] = (decimal.Parse(symbolOPriceMap[combineSymbol1][(int)OPTION_LIST.STOCK]) + (-1)*Math.Round(decimal.Parse(p1618b.Qty), 0)).ToString();
                                    else symbolOPriceMap[combineSymbol1][(int)OPTION_LIST.STOCK] = (decimal.Parse(symbolOPriceMap[combineSymbol1][(int)OPTION_LIST.STOCK]) + Math.Round(decimal.Parse(p1618b.Qty),0)).ToString();  //在option的dictionary裡加入庫存
                                }
                            }
                            if (!string.IsNullOrEmpty(p1618b.Qty2))
                            {
                                if (p1618b.CP2 == 'C')
                                {
                                    combineSymbol2 = p1618b.ComID2;
                                    combineSymbol2 += Math.Round(decimal.Parse(p1618b.StrikePrice2), 0).ToString();
                                    combineSymbol2 += monthSymbol[0].ToString() + p1618b.ComYM2[3].ToString();
                                }
                                else
                                {
                                    combineSymbol2 = p1618b.ComID2;
                                    combineSymbol2 += Math.Round(decimal.Parse(p1618b.StrikePrice2), 0).ToString();
                                    combineSymbol2 += monthSymbol[1].ToString() + p1618b.ComYM2[3].ToString();
                                }
                                if (!string.IsNullOrEmpty(combineSymbol2) && symbolOPriceMap.ContainsKey(combineSymbol2))
                                {
                                    if (string.IsNullOrEmpty(symbolOPriceMap[combineSymbol2][(int)OPTION_LIST.STOCK])) symbolOPriceMap[combineSymbol2][(int)OPTION_LIST.STOCK] = "0";
                                    if (p1618b.BS2 == 'S') symbolOPriceMap[combineSymbol2][(int)OPTION_LIST.STOCK] = (decimal.Parse(symbolOPriceMap[combineSymbol2][(int)OPTION_LIST.STOCK]) + (-1)*Math.Round(decimal.Parse(p1618b.Qty2), 0)).ToString();
                                    else symbolOPriceMap[combineSymbol2][(int)OPTION_LIST.STOCK] = (decimal.Parse(symbolOPriceMap[combineSymbol2][(int)OPTION_LIST.STOCK]) + Math.Round(decimal.Parse(p1618b.Qty2),0)).ToString();  //在option的dictionary裡加入庫存
                                }
                            }

                        }
                    }
                    stockRPT = tmpstockRPT;
                    stockAggre = tmpstockAggre;
                    break;
                case DT.FINANCIAL_COVER_TRADER_Detail:  //1624 
                    P001624 p1624 = (P001624)package;
                    //dgv1624.Rows.Clear();
                    AddInfo("RCV 1624 [" + p1624.Rows + "]");
                    if (p1624.Rows > 0)
                    {
                        foreach (P001624_2 p1624b in p1624.p001624_2)
                        {
                            string[] row1624 = { p1624b.BrokerId, p1624b.Account, p1624b.Group, p1624b.Trader, p1624b.Exchange, p1624b.OccDT, p1624b.TrdDT1, p1624b.OrdNo1, p1624b.FirmOrd1, p1624b.OffsetSpliteSeqNo, p1624b.TrdDT2, p1624b.OrdNo2, p1624b.FirmOrd2, p1624b.OffsetSpliteSeqNo2, p1624b.OffsetCode.ToString(), p1624b.offset.ToString(), p1624b.BS.ToString(), p1624b.ComID, p1624b.ComYM, p1624b.StrikePrice, p1624b.CP.ToString(), p1624b.ComID2, p1624b.Qty1, p1624b.Qty2, p1624b.TrdPrice1, p1624b.TrdPrice2, p1624b.PRTLOS, p1624b.AENO, p1624b.Currency, p1624b.CTAXAMT, p1624b.ORIGNFEE, p1624b.Premium1, p1624b.Premium2, p1624b.InNo1, p1624b.InNo2, p1624b.Cnt1, p1624b.Cnt2, p1624b.OSRTLOS };
                            //dgv1624.Rows.Add(row1624);
                        }
                    }
                    break;

                case DT.FINANCIAL_TRADERN: //1626 
                    P001626 p1626 = (P001626)package;
                    AddInfo("RCV 1626 [" + p1626.Count + "]");
                    dgv1626.Rows.Clear();
                    int j = 0;
                    if (p1626.Count > 0)
                    {
                        equity = 0;
                        foreach (P001626_2 p1626b in p1626.p001626_2)
                        {
                            //,p1626b.WithdrawMnt 
                            string[] row1626 = { p1626b.BrokerId, p1626b.Account, p1626b.Group, p1626b.Trader, p1626b.Currency, p1626b.LCTDAB, p1626b.ORIGNFEE, p1626b.TAXAMT, p1626b.CTAXAMT, p1626b.DWAMT, p1626b.OSPRTLOS, p1626b.PRTLOS, p1626b.BMKTVAL, p1626b.SMKTVAL, p1626b.OPREMIUM, p1626b.TPREMIUM, p1626b.EQUITY, p1626b.IAMT, p1626b.MAMT, p1626b.EXCESS, p1626b.ORDEXCESS, p1626b.ORDAMT, p1626b.ExProfit, p1626b.Premium, p1626b.PTime, p1626b.FloatProfit, p1626b.LASSPRTLOS, p1626b.CLOSEAMT, p1626b.ORDIAMT, p1626b.ORDMAMT, p1626b.DayTradeAMT, p1626b.ReductionAMT, p1626b.CreditAMT, p1626b.balance, p1626b.IPremium, p1626b.OPremium, p1626b.Securities, p1626b.SecuritiesOffset, p1626b.OffsetAMT, p1626b.Offset, p1626b.FULLMTRISK, p1626b.FULLRISK, p1626b.MarginCall, p1626b.SellVerticalSpread, p1626b.StrikePrice, p1626b.ActMarketValue, p1626b.TPRTLOS, p1626b.MarginCall1, p1626b.AddMargin, p1626b.ORDAMTNOCN, p1626b.WithdrawMnt };
                            dgv1626.Rows.Add(row1626);
                            if (string.Compare(p1626b.Currency, "TWD") == 0) equity = decimal.Round(decimal.Parse(p1626b.EQUITY), 2);
                            if (j == 0)
                            {
                                tbEQ1.Text = decimal.Round(decimal.Parse(p1626b.LCTDAB),0).ToString();
                                tbEQ2.Text = decimal.Round(decimal.Parse(p1626b.DWAMT), 0).ToString();
                                tbEQ3.Text = decimal.Round(decimal.Parse(p1626b.DWAMT), 0).ToString();
                                tbEQ4.Text = decimal.Round(decimal.Parse(p1626b.TPREMIUM), 0).ToString();
                                tbEQ5.Text = decimal.Round(decimal.Parse(p1626b.TPRTLOS), 0).ToString();
                                tbEQ6.Text = decimal.Round(decimal.Parse(p1626b.ORIGNFEE), 0).ToString();
                                tbEQ7.Text = decimal.Round(decimal.Parse(p1626b.CTAXAMT), 0).ToString();
                                tbEQ8.Text = decimal.Round(decimal.Parse(p1626b.FloatProfit), 0).ToString();
                                tbEQ9.Text = p1626b.FULLMTRISK;
                                tbEQ10.Text = decimal.Round(decimal.Parse(p1626b.EQUITY), 0).ToString();
                                tbEQ11.Text = decimal.Round(decimal.Parse(p1626b.ActMarketValue), 0).ToString();
                                tbEQ12.Text = decimal.Round(decimal.Parse(p1626b.IAMT), 0).ToString();
                                tbEQ13.Text = decimal.Round(decimal.Parse(p1626b.MAMT), 0).ToString();
                                tbEQ14.Text = decimal.Round(decimal.Parse(p1626b.ORDAMT), 0).ToString();
                                tbEQ15.Text = decimal.Round(decimal.Parse(p1626b.OPREMIUM), 0).ToString();
                                tbEQ16.Text = decimal.Round(decimal.Parse(p1626b.ORDEXCESS), 0).ToString();
                            }
                            j++;
                        }
                    }
                    break;

                case DT.FINANCIAL_CURRENCY: //1628 
                    P001628 p1628 = (P001628)package;
                    AddInfo("RCV 1628 CODE= " + p1628.Code + " MSG=" + p1628.ErrorMsg);
                    break;
                case DT.FINANCIAL_STRIKE:   //2017.2.4 1643  無效履約查詢
                                            //        string[] header1643 = { "分公司", "帳號", "組別", "交易員","履約日期","成交日期","委託單號","成交序號","拆單序號", "交易所", "商品代碼", "商品年月", "履約價", "CP", "BS", "口數", "交易稅幣別","交易稅","手續費幣別","手續費","交易幣別","權利金收支/履約盈虧","結算價"};

                    P001643 p1643 = (P001643)package;
                    AddInfo("RCV 1643 Code =  " + p1643.Code + "   Rows[" + p1643.Rows + "]");
                    //dgv1643.Rows.Clear();
                    if (p1643.Rows > 0)
                    {
                        foreach (P001643_2 p1643b in p1643.Detail)
                        {
                            string[] row1643 = { p1643b.BrokerId, p1643b.Account, p1643b.Group, p1643b.Trader, p1643b.DueDate, p1643b.TrdDate, p1643b.OrdNo, p1643b.FirmOrd, p1643b.SeqNo, p1643b.Exchange, p1643b.ComID, p1643b.ComYM, p1643b.StrikePrice, p1643b.CP.ToString(), p1643b.BS.ToString(), p1643b.Qty, p1643b.TaxCurr, p1643b.TaxAmt, p1643b.FeeCurr, p1643b.FeeAmt, p1643b.TrdCurr, p1643b.Premium, p1643b.TrdPre };
                            //dgv1643.Rows.Add(row1643);
                        }
                    }

                    break;
                case DT.FINANCIAL_COVERDH:     //2017.2.14  1645 ; 平倉明細歷史查詢

                    P001645 p1645 = (P001645)package;
                    AddInfo("RCV 1645 Code =  " + p1645.Code + "   Rows[" + p1645.Rows + "]");
                    //dgv1645.Rows.Clear();
                    if (p1645.Rows > 0)
                    {
                        foreach (P001645_2 p1645b in p1645.Detail)
                        {
                            //string[] header1645 = { "市場別", "分公司", "帳號"          , "組別",                       "交易員",                                  "交易所"              , "平倉日期",     "平倉成交日期", "平倉委託編號", "平倉成交序號",         "平倉拆單序號", "被平成交日期",     "被平委託編號",   "被平成交序號",     "被平拆單序號",          "指定平倉碼",                               "互抵",                                       "BS",                        "商品代號", "商品年月", "履約價 ", "CP", "被平商品代號", "平倉口數", "被平口數", "平倉成交價", "被平成交價", "平倉損益", "業務員代號", "幣別", "交易稅", "手續費", "平倉權利金", "被平權利金", "平倉場內編號", "被平場內編號", "平倉電子單號", "被平電子單號" };
                            string[] row1645 = { p1645b.Market, p1645b.BrokerId, p1645b.Account, p1645b.Group, p1645b.Trader, p1645b.Exchange, p1645b.OccDT, p1645b.TrdDT1, p1645b.OrdNo1, p1645b.FirmOrd1, p1645b.OffsetSpliteSeqNo, p1645b.TrdDT2, p1645b.OrdNo2, p1645b.FirmOrd2, p1645b.OffsetSpliteSeqNo2, p1645b.OffsetCode.ToString(), p1645b.offset.ToString(), p1645b.BS.ToString(), p1645b.ComID, p1645b.ComYM, p1645b.StrikePrice, p1645b.CP.ToString(), p1645b.ComID2, p1645b.Qty1, p1645b.Qty2, p1645b.TrdPrice1, p1645b.TrdPrice2, p1645b.PRTLOS, p1645b.AENO, p1645b.Currency, p1645b.CTAXAMT, p1645b.ORIGNFEE, p1645b.Premium1, p1645b.Premium2, p1645b.InNo1, p1645b.InNo2, p1645b.Cnt1, p1645b.Cnt2 };
                            //dgv1645.Rows.Add(row1645);
                        }
                    }
                    break;
                case DT.FINANCIAL_RECIPROCATE:  //2018.12 : 大小台互抵
                    P001647 p1647 = (P001647)package;
                    string msg = p1647.BrokerId + p1647.Account;
                    if (p1647.Status == "0")
                        AddInfo(msg + "大小台互抵成功: 大台 " + p1647.Qty1 + " ,小台 " + p1647.Qty2);
                    else
                        AddInfo(msg + " 大小台互抵失敗");
                    break;

                case DT.FINANCIAL_SERVICECHARGE: //1631  手續費; 
                    P001631 p1631 = (P001631)package;
                    AddInfo("RCV 1631 [" + p1631.Count + "]");
                    //dgv1631.Rows.Clear();
                    if (p1631.Count > 0)
                    {
                        foreach (P001631_2 p1631b in p1631.p001631_2)
                        {
                            string[] row1631 = { p1631b.BrokerId, p1631b.Account, p1631b.Group, p1631b.Trader, p1631b.Kind, p1631b.Class, p1631b.Qty, p1631b.Method.ToString(), p1631b.Rule.ToString(), p1631b.Amt, p1631b.Rate, p1631b.Currency };
                            //dgv1631.Rows.Add(row1631);
                        }
                    }
                    break;
                case DT.FINANCIAL_APPLYWITHDRAW: //1633 出金申請; 
                    P001633 p1633 = (P001633)package;
                    AddInfo("RCV 1633 ");
                    string[] row1633 = { p1633.Webid, p1633.Seqno, p1633.ApplyNo, p1633.UpdateTime, p1633.ApplyDate, p1633.ApplyTime, p1633.AMT, p1633.Code, p1633.ErrorMsg };
                    //dgv1633.Rows.Add(row1633);

                    break;
                case DT.FINANCIAL_WITHDRAW: //1635 ; 
                    P001635 p1635 = (P001635)package;
                    AddInfo("RCV 1635 [" + p1635.Count + "]");
                    //dgv1635.Rows.Clear();
                    if (p1635.Count > 0)
                    {
                        foreach (P001635_2 p1635b in p1635.p001635_2)
                        {
                            string[] row1635 = { p1635b.Webid, p1635b.Seqno, p1635b.Market, p1635b.Type, p1635b.UpdateDate, p1635b.WithdrawNo, p1635b.BankNo, p1635b.BankAccount, p1635b.Currency, p1635b.Amt, p1635b.NTAmt, p1635b.Status };
                            //dgv1635.Rows.Add(row1635);
                        }
                    }
                    break;
                    #endregion
            }
        }
        #endregion

        //ListView Print
        public void LoopPrint()
        {
            while (!this.IsHandleCreated)
            {; }
            while (true)
            {
                if (formClose) break;
                

                decimal dif = 0;
                decimal price = 0;
                decimal openPrice = 0;
                char upDownAscii;
                
                try
                {
                    List<string> keyList = new List<string>(symbolOPriceMap.Keys);
                    for (int i = 0; i < keyList.Count; i++)
                    {
                        //計算漲跌
                        if (keyList[i].Length > 3 && symbolOPriceMap[keyList[i]][(int)OPTION_LIST.PRICE] != null && symbolOPriceMap[keyList[i]][(int)OPTION_LIST.OPENPRICE] != null)
                        {
                            //Console.WriteLine("before updown!");
                            price = decimal.Parse(symbolOPriceMap[keyList[i]][(int)OPTION_LIST.PRICE]);
                            openPrice = decimal.Parse(symbolOPriceMap[keyList[i]][(int)OPTION_LIST.OPENPRICE]);
                            if (price - openPrice > 0)
                            {
                                upDownAscii = Convert.ToChar(11205);
                                symbolOPriceMap[keyList[i]][(int)OPTION_LIST.UPDOWN] = upDownAscii.ToString() + (price - openPrice).ToString();
                            }
                            else if (price - openPrice < 0)
                            {
                                upDownAscii = Convert.ToChar(11206);
                                symbolOPriceMap[keyList[i]][(int)OPTION_LIST.UPDOWN] = (price - openPrice).ToString().Replace('-', upDownAscii);
                            }
                            else symbolOPriceMap[keyList[i]][(int)OPTION_LIST.UPDOWN] = (price - openPrice).ToString();
                            //Console.WriteLine("af updown!");
                        }
                        //計算合計量
                        if (keyList[i].Length > 3 && symbolOPriceMap[keyList[i]][(int)OPTION_LIST.OI] != null && symbolOPriceMap[keyList[i]][(int)OPTION_LIST.VOL] != null
                              && string.Compare(symbolOPriceMap[keyList[i]][(int)OPTION_LIST.OI], "-") != 0 && string.Compare(symbolOPriceMap[keyList[i]][(int)OPTION_LIST.VOL], "-") != 0)
                        {
                            symbolOPriceMap[keyList[i]][(int)OPTION_LIST.TOTALVOL] = (decimal.Parse(symbolOPriceMap[keyList[i]][(int)OPTION_LIST.OI]) + decimal.Parse(symbolOPriceMap[keyList[i]][(int)OPTION_LIST.VOL])).ToString();
                        }
                        for (int j = 0; j < keyList.Count; j++)
                        {
                            //Console.WriteLine("before 計算合計量!");
                            if (keyList[i].Length > 8 && keyList[j].Length > 8)
                            {
                                //計算150差100差
                                if (keyList[i][keyList[i].Length - 2] >= 'M' && keyList[i][keyList[i].Length - 2] <= 'X' && keyList[i][keyList[i].Length - 2] == keyList[j][keyList[j].Length - 2])
                                {
                                    if (int.Parse(keyList[i].Substring(3, 5)) - int.Parse(keyList[j].Substring(3, 5)) == 150)
                                    {
                                        dif = decimal.Parse(symbolOPriceMap[keyList[i]][(int)OPTION_LIST.PRICE]) - decimal.Parse(symbolOPriceMap[keyList[j]][(int)OPTION_LIST.PRICE]);
                                        if (dif > 0)
                                            symbolOPriceMap[keyList[i]][(int)OPTION_LIST.HUNDFIFDIF] = dif.ToString();
                                    }
                                    if (int.Parse(keyList[i].Substring(3, 5)) - int.Parse(keyList[j].Substring(3, 5)) == 100)
                                    {
                                        dif = decimal.Parse(symbolOPriceMap[keyList[i]][(int)OPTION_LIST.PRICE]) - decimal.Parse(symbolOPriceMap[keyList[j]][(int)OPTION_LIST.PRICE]);
                                        if (dif > 0)
                                            symbolOPriceMap[keyList[i]][(int)OPTION_LIST.HUNDDIF] = dif.ToString();
                                    }
                                }
                                else if (keyList[i][keyList[i].Length - 2] >= 'A' && keyList[i][keyList[i].Length - 2] <= 'L' && keyList[i][keyList[i].Length - 2] == keyList[j][keyList[j].Length - 2])
                                {
                                    if (int.Parse(keyList[j].Substring(3, 5)) - int.Parse(keyList[i].Substring(3, 5)) == 150)
                                    {
                                        dif = decimal.Parse(symbolOPriceMap[keyList[i]][(int)OPTION_LIST.PRICE]) - decimal.Parse(symbolOPriceMap[keyList[j]][(int)OPTION_LIST.PRICE]);
                                        if (dif > 0)
                                            symbolOPriceMap[keyList[i]][(int)OPTION_LIST.HUNDFIFDIF] = dif.ToString();
                                    }
                                    if (int.Parse(keyList[j].Substring(3, 5)) - int.Parse(keyList[i].Substring(3, 5)) == 100)
                                    {
                                        dif = decimal.Parse(symbolOPriceMap[keyList[i]][(int)OPTION_LIST.PRICE]) - decimal.Parse(symbolOPriceMap[keyList[j]][(int)OPTION_LIST.PRICE]);
                                        if (dif > 0)
                                            symbolOPriceMap[keyList[i]][(int)OPTION_LIST.HUNDDIF] = dif.ToString();
                                    }
                                }
                            }
                            //Console.WriteLine("af 計算合計量!");
                        }
                    }
                }
                catch (Exception ex)
                {
                    opLog.txtLog("print計算 exception: " + ex.ToString());
                }

                //Call的部分
                //dataGridView1.Rows.Clear();
                List<string> row = new List<string>();

                bool findCall = false;
                bool findPut = false;

                try
                {
                    foreach (string key in symbolOPriceMap.Keys)
                    {
                        if (symbolOPriceMap[key][(int)OPTION_LIST.SYMBOL].Substring(3, 5) != null)//報價Dictionary的Key與第一個Member皆是商品代碼 ex:TX218000A2  substring(3,5)即是18000
                        {
                            if (key[key.Length - 2] >= 'A' && key[key.Length - 2] <= 'L')//找Call
                            {
                                for (int i = 0; i < tquoteStructList.Count; i++)
                                {
                                    if (string.Compare(tquoteStructList[i].StrickPrice, symbolOPriceMap[key][(int)OPTION_LIST.SYMBOL].Substring(3, 5)) == 0)
                                    {
                                        tquoteStructList[i].StrickPrice = symbolOPriceMap[key][(int)OPTION_LIST.SYMBOL].Substring(3, 5);//位置[0]填履約價方便排序，寬度再設零
                                        tquoteStructList[i].Stock = symbolOPriceMap[key][(int)OPTION_LIST.STOCK];
                                        tquoteStructList[i].OpenPrice = symbolOPriceMap[key][(int)OPTION_LIST.OPENPRICE];
                                        tquoteStructList[i].UpDown = symbolOPriceMap[key][(int)OPTION_LIST.UPDOWN];
                                        tquoteStructList[i].Vol = symbolOPriceMap[key][(int)OPTION_LIST.VOL];
                                        tquoteStructList[i].OI = symbolOPriceMap[key][(int)OPTION_LIST.OI];
                                        tquoteStructList[i].TotalVol = symbolOPriceMap[key][(int)OPTION_LIST.TOTALVOL];
                                        tquoteStructList[i].HundfifDif = symbolOPriceMap[key][(int)OPTION_LIST.HUNDFIFDIF];
                                        tquoteStructList[i].HundDif = symbolOPriceMap[key][(int)OPTION_LIST.HUNDDIF];
                                        tquoteStructList[i].DealPrice = symbolOPriceMap[key][(int)OPTION_LIST.PRICE];
                                        tquoteStructList[i].Symbol = symbolOPriceMap[key][(int)OPTION_LIST.SYMBOL];//T字報價履約價左邊，填Call的Symbol
                                        tquoteStructList[i].StrickPriceCenter = symbolOPriceMap[key][(int)OPTION_LIST.SYMBOL].Substring(3, 5);//中間填一次履約價當選擇權T字報價中間的值
                                        findPut = false;
                                        foreach (string keyPut in symbolOPriceMap.Keys)//看看Put裡面是否有跟剛剛的Call訂閱同樣的履約價
                                        {
                                            if (keyPut[keyPut.Length - 2] >= 'M' && keyPut[keyPut.Length - 2] <= 'X' && string.Compare(symbolOPriceMap[keyPut][(int)OPTION_LIST.SYMBOL].Substring(3, 5), symbolOPriceMap[key][(int)OPTION_LIST.SYMBOL].Substring(3, 5)) == 0)
                                            {
                                                findPut = true;
                                                tquoteStructList[i].Symbol2 = symbolOPriceMap[keyPut][(int)OPTION_LIST.SYMBOL];//T字報價履約價右邊填Put的Symbol方便點選下單使用
                                                tquoteStructList[i].DealPrice2 = symbolOPriceMap[keyPut][(int)OPTION_LIST.PRICE];
                                                tquoteStructList[i].HundDif2 = symbolOPriceMap[keyPut][(int)OPTION_LIST.HUNDDIF];
                                                tquoteStructList[i].HundfifDif2 = symbolOPriceMap[keyPut][(int)OPTION_LIST.HUNDFIFDIF];
                                                tquoteStructList[i].TotalVol2 = symbolOPriceMap[keyPut][(int)OPTION_LIST.TOTALVOL];
                                                tquoteStructList[i].OI2 = symbolOPriceMap[keyPut][(int)OPTION_LIST.OI];
                                                tquoteStructList[i].Vol2 = symbolOPriceMap[keyPut][(int)OPTION_LIST.VOL];
                                                tquoteStructList[i].UpDown2 = symbolOPriceMap[keyPut][(int)OPTION_LIST.UPDOWN];
                                                tquoteStructList[i].OpenPrice2 = symbolOPriceMap[keyPut][(int)OPTION_LIST.OPENPRICE];
                                                tquoteStructList[i].Stock2 = symbolOPriceMap[keyPut][(int)OPTION_LIST.STOCK];
                                            }
                                        }
                                        //如果此履約價沒訂閱Put則填"-"
                                        if (findPut == false)
                                        {
                                            tquoteStructList[i].Symbol2 = "-";
                                            tquoteStructList[i].DealPrice2 = "-";
                                            tquoteStructList[i].HundDif2 = "-";
                                            tquoteStructList[i].HundfifDif2 = "-";
                                            tquoteStructList[i].TotalVol2 = "-";
                                            tquoteStructList[i].OI2 = "-";
                                            tquoteStructList[i].Vol2 = "-";
                                            tquoteStructList[i].UpDown2 = "-";
                                            tquoteStructList[i].OpenPrice2 = "-";
                                            tquoteStructList[i].Stock2 = "-";
                                            break;
                                        }
                                        else break;
                                    }

                                }
                                //dataGridView1.Rows.Add(row.ToArray());
                            }
                            else if (key[key.Length - 2] >= 'M' && key[key.Length - 2] <= 'X')//找Put
                            {
                                findCall = false;
                                foreach (string keyCall in symbolOPriceMap.Keys)//看看Call裡面是否有跟剛剛的Put訂閱同樣的履約價，有的話則跳過，因為前一個if找Call的時候已經填過了
                                {
                                    if (keyCall[keyCall.Length - 2] >= 'A' && keyCall[keyCall.Length - 2] <= 'L' && string.Compare(symbolOPriceMap[keyCall][(int)OPTION_LIST.SYMBOL].Substring(3, 5), symbolOPriceMap[key][(int)OPTION_LIST.SYMBOL].Substring(3, 5)) == 0)
                                    {
                                        findCall = true;
                                    }
                                }
                                if (findCall == false)//剩下來的履約價都只有訂閱Put
                                {
                                    for (int i = 0; i < tquoteStructList.Count; i++)
                                    {
                                        if (string.Compare(tquoteStructList[i].StrickPrice, symbolOPriceMap[key][(int)OPTION_LIST.SYMBOL].Substring(3, 5)) == 0)
                                        {
                                            tquoteStructList[i].StrickPrice = symbolOPriceMap[key][(int)OPTION_LIST.SYMBOL].Substring(3, 5);//位置[0]填履約價方便排序，寬度再設零
                                            tquoteStructList[i].Stock = "-";
                                            tquoteStructList[i].OpenPrice = "-";
                                            tquoteStructList[i].UpDown = "-";
                                            tquoteStructList[i].Vol = "-";
                                            tquoteStructList[i].OI = "-";
                                            tquoteStructList[i].TotalVol = "-";
                                            tquoteStructList[i].HundfifDif = "-";
                                            tquoteStructList[i].HundDif = "-";
                                            tquoteStructList[i].DealPrice = "-";
                                            tquoteStructList[i].Symbol = "-";//T字報價履約價左邊，本來要填Call的Symbol，但此履約價無訂閱Call則填"-"
                                            tquoteStructList[i].StrickPriceCenter = symbolOPriceMap[key][(int)OPTION_LIST.SYMBOL].Substring(3, 5);//中間填一次履約價當選擇權T字報價中間的值
                                            tquoteStructList[i].Symbol2 = symbolOPriceMap[key][(int)OPTION_LIST.SYMBOL];//T字報價履約價右邊填Put的Symbol方便點選下單使用
                                            tquoteStructList[i].DealPrice2 = symbolOPriceMap[key][(int)OPTION_LIST.PRICE];
                                            tquoteStructList[i].HundDif2 = symbolOPriceMap[key][(int)OPTION_LIST.HUNDDIF];
                                            tquoteStructList[i].HundfifDif2 = symbolOPriceMap[key][(int)OPTION_LIST.HUNDFIFDIF];
                                            tquoteStructList[i].TotalVol2 = symbolOPriceMap[key][(int)OPTION_LIST.TOTALVOL];
                                            tquoteStructList[i].OI2 = symbolOPriceMap[key][(int)OPTION_LIST.OI];
                                            tquoteStructList[i].Vol2 = symbolOPriceMap[key][(int)OPTION_LIST.VOL];
                                            tquoteStructList[i].UpDown2 = symbolOPriceMap[key][(int)OPTION_LIST.UPDOWN];
                                            tquoteStructList[i].OpenPrice2 = symbolOPriceMap[key][(int)OPTION_LIST.OPENPRICE];
                                            tquoteStructList[i].Stock2 = symbolOPriceMap[key][(int)OPTION_LIST.STOCK];
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    opLog.txtLog("tquoteStructListassign exception: " + ex.ToString());
                }

                this.Invoke((Action)(() =>
                {
                    if(formClose)
                    {
                        FlagInvokeLoopPrint = false;
                        return;
                    }
                    FlagInvokeLoopPrint = true;

                    try 
                    {
                        tbScallSum.Text = decimal.Round(callSum, 2).ToString();
                        tbSputSum.Text = decimal.Round(putSum, 2).ToString();

                        decimal[] _dayRent = new decimal[4];
                        _dayRent = orderSettings.cacDayRentSumAndDaily();

                        tbTodayAccPointCall.Text = _dayRent[0].ToString();
                        tbTodayAccPointPut.Text = _dayRent[1].ToString();
                        tbTodaySetPointCall.Text = _dayRent[2].ToString();
                        tbTodaySetPointPut.Text = _dayRent[3].ToString();
                        tbDoneAndTargetPointCall.Text = todayOrderRecord.cacTodayAutoOrderPoints()[0].ToString() + " / " + todayOrderStatus.TodayRentCall.ToString();
                        tbDoneAndTargetPointPut.Text = todayOrderRecord.cacTodayAutoOrderPoints()[1].ToString() + " / " + todayOrderStatus.TodayRentPut.ToString();
                        //tquoteBinding.ResetBindings(false);
                        dataGridView1.Update();
                        dataGridView1.Refresh();
                        //if(dataGridView1.ColumnCount > 0 && dataGridView1.Columns["StrickPrice"]!=null && dataGridView1.DataSource !)
                        //    this.dataGridView1.Sort(this.dataGridView1.Columns["StrickPrice"], ListSortDirection.Ascending);

                        int[] largestCallTotoalVol = new int[3] { 0, 0, 0 };
                        int[] largestPutTotoalVol = new int[3] { 0, 0, 0 };
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            for (int j = 0; j < dataGridView1.ColumnCount; j++)
                            {
                                if (dataGridView1.Rows[i].Cells[j].Value != null && string.Compare(dataGridView1.Rows[i].Cells[j].Value.ToString(), "-") != 0)
                                {
                                    if (j == 6)
                                    {
                                        if (decimal.Parse(dataGridView1.Rows[i].Cells[j].Value.ToString()) > largestCallTotoalVol[2])
                                        {
                                            largestCallTotoalVol[0] = i;
                                            largestCallTotoalVol[1] = j;
                                            largestCallTotoalVol[2] = (int)decimal.Parse(dataGridView1.Rows[i].Cells[j].Value.ToString());
                                        }
                                    }
                                    else if (j == 16)
                                    {
                                        if (decimal.Parse(dataGridView1.Rows[i].Cells[j].Value.ToString()) > largestPutTotoalVol[2])
                                        {
                                            largestPutTotoalVol[0] = i;
                                            largestPutTotoalVol[1] = j;
                                            largestPutTotoalVol[2] = (int)decimal.Parse(dataGridView1.Rows[i].Cells[j].Value.ToString());
                                        }
                                    }
                                    if (dataGridView1.Rows[i].Cells[j].Style.BackColor == Color.Orange)
                                        dataGridView1.Rows[i].Cells[j].Style.BackColor = Color.Black;
                                    //Console.WriteLine("[" + i + "][" + j + "]: " + dataGridView1.Rows[i].Cells[j].Value.ToString() + "  ");
                                }
                            }
                            //Console.WriteLine("");
                        }
                        //找出Call、Put各自的最大合計量，並上色
                        if (largestCallTotoalVol[0] != 0 && largestCallTotoalVol[1] != 0 && largestCallTotoalVol[2] != 0)
                        {
                            dataGridView1.Rows[largestCallTotoalVol[0]].Cells[largestCallTotoalVol[1]].Style.BackColor = Color.Orange;
                            dataGridView1.Rows[largestCallTotoalVol[0]].Cells[largestCallTotoalVol[1] - 1].Style.BackColor = Color.Orange;
                            dataGridView1.Rows[largestCallTotoalVol[0]].Cells[largestCallTotoalVol[1] - 2].Style.BackColor = Color.Orange;
                        }
                        if (largestPutTotoalVol[0] != 0 && largestPutTotoalVol[1] != 0 && largestPutTotoalVol[2] != 0)
                        {
                            dataGridView1.Rows[largestPutTotoalVol[0]].Cells[largestPutTotoalVol[1]].Style.BackColor = Color.Orange;
                            dataGridView1.Rows[largestPutTotoalVol[0]].Cells[largestPutTotoalVol[1] + 1].Style.BackColor = Color.Orange;
                            dataGridView1.Rows[largestPutTotoalVol[0]].Cells[largestPutTotoalVol[1] + 2].Style.BackColor = Color.Orange;
                        }

                        CalculateComID calcuComID;
                        if (string.Compare(DateTime.Now.DayOfWeek.ToString("d"), "3") == 0 && DateCalculateExtensions.IsLegalTime(DateTime.Now, "050001-143000"))
                            calcuComID = new CalculateComID(DateTime.Now.AddDays(-1));
                        else calcuComID = new CalculateComID(DateTime.Now);
                        string weekMX = calcuComID.PartialComIDMX[0] + calcuComID.PartialComIDMX[1];
                        //Console.WriteLine("loopPrintweekMX: " + weekMX);
                        if (symbolFPriceMap.ContainsKey(weekMX) && dataGridView1.ColumnCount > 0)
                        {
                            decimal weekMXPrice = Math.Round(decimal.Parse(symbolFPriceMap[weekMX][(int)FUTURE_LIST.PRICE]) / 50, 0, MidpointRounding.AwayFromZero) * 50;
                            int centerIndex = dataGridView1.Columns["StrickPriceCenter"].Index;
                            for (int i = 0; i < dataGridView1.Rows.Count; i++)
                            {
                                if (dataGridView1.Rows[i].Cells[centerIndex].Value != null && string.Compare(dataGridView1.Rows[i].Cells[centerIndex].Value.ToString(), weekMXPrice.ToString()) == 0)
                                {
                                    dataGridView1.Rows[i].Cells[centerIndex].Style.ForeColor = Color.Yellow;
                                }
                            }
                        }
                        //Console.WriteLine("after 合計量上色");
                        //找出Call、Put各自是上漲還下跌，並上色
                        if (dataGridView1.ColumnCount > 0)
                        {
                            int centerIndex = dataGridView1.Columns["StrickPriceCenter"].Index;
                            int callUpDownIndex = dataGridView1.Columns["UpDown"].Index;
                            int putUpDownIndex = dataGridView1.Columns["UpDown2"].Index;
                            for (int i = 0; i < dataGridView1.Rows.Count; i++)
                            {
                                if (dataGridView1.Rows[i].Cells[callUpDownIndex].Value != null && string.Compare(dataGridView1.Rows[i].Cells[callUpDownIndex].Value.ToString(), "-") != 0)
                                {
                                    for (int j = 1; j < centerIndex; j++)
                                    {
                                        if (dataGridView1.Rows[i].Cells[callUpDownIndex].Value.ToString()[0] == Convert.ToChar(11205)) //上三角形
                                            dataGridView1.Rows[i].Cells[j].Style.ForeColor = Color.Red;
                                        else if (dataGridView1.Rows[i].Cells[callUpDownIndex].Value.ToString()[0] == Convert.ToChar(11206)) //下三角形
                                            dataGridView1.Rows[i].Cells[j].Style.ForeColor = Color.Green;
                                    }
                                }
                                if (dataGridView1.Rows[i].Cells[putUpDownIndex].Value != null && string.Compare(dataGridView1.Rows[i].Cells[putUpDownIndex].Value.ToString(), "-") != 0)
                                {
                                    for (int j = centerIndex + 1; j < dataGridView1.ColumnCount; j++)
                                    {
                                        if (dataGridView1.Rows[i].Cells[putUpDownIndex].Value.ToString()[0] == Convert.ToChar(11205)) //上三角形
                                            dataGridView1.Rows[i].Cells[j].Style.ForeColor = Color.Red;
                                        else if (dataGridView1.Rows[i].Cells[putUpDownIndex].Value.ToString()[0] == Convert.ToChar(11206)) //下三角形
                                            dataGridView1.Rows[i].Cells[j].Style.ForeColor = Color.Green;
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        opLog.txtLog("largestTOTVOL exception: " + ex.ToString());
                    }

                    var item = new ListViewItem();
                    listView2.Items.Clear();
                    for (int i=0; i<stockAggre.Count ;i++)
                    {
                        item = new ListViewItem(stockAggre[i][0]);
                        for (int j = 1; j < stockAggre[i].Count; j++)
                        {
                            item.SubItems.Add(stockAggre[i][j]);
                        }
                        listView2.Items.Add(item);
                    }
                    //for (int i = 0; i < listView2.Columns.Count; i++)
                    //    listView2.Columns[i].Width = -2;


                    listView3.Items.Clear();
                    foreach (string key in orderRPT.Keys)
                    {
                        item = new ListViewItem(orderRPT[key][0]);
                        for (int i = 1; i < orderRPT[key].Count; i++)
                            item.SubItems.Add(orderRPT[key][i]);
                        listView3.Items.Add(item);
                    }
                    for (int i = 0; i < listView3.Columns.Count; i++)
                        listView3.Columns[i].Width = -2;

                    listView4.Items.Clear();
                    foreach (string key in dealRPT.Keys)
                    {
                        item = new ListViewItem(dealRPT[key][0]);
                        for (int i = 1; i < dealRPT[key].Count; i++)
                            item.SubItems.Add(dealRPT[key][i]);
                        listView4.Items.Add(item);
                    }
                    for (int i = 0; i < listView4.Columns.Count; i++)
                        listView4.Columns[i].Width = -2;

                    FlagInvokeLoopPrint = false;
                    if (formClose)
                    {
                        return;
                    }
                }));
                if(!formClose) Thread.Sleep(int.Parse(settings.PrintGap) * 1000);
            }
            formCloseFinish = true;
        }

        public void LoopOrder()
        {
            //int count = 0;
            bool emergencyBreak = false;
            bool emergencyBreakStopLoss = false;
            while (!this.IsHandleCreated)
            {; }
            while (true)
            {
                //count++;
                if (formClose) break;
                if (emergencyBreak) AddInfoTG("EmergencyBreak, Please check!!!!");
                if (emergencyBreakStopLoss) AddInfoTG("emergencyBreakStopLoss, Please check!!!!");

                decimal[] _dayRent = new decimal[4];
                _dayRent = orderSettings.cacDayRentSumAndDaily();

                decimal[] dayRentSum = new decimal[] { _dayRent[0], _dayRent[1] };//至今日累計設定配額 [0]:call [1]:put
                todayOrderStatus.cacTodayRent(callSum, putSum, _dayRent, todayOrderRecord.cacTodayAutoOrderPoints());

                #region Call下單
                //Call
                ushort partialDealRemains = 0;
                ushort lots;
                decimal tmpDec;
                int failCount = 0;
                int successCount = 0;
                Stopwatch stopWatch = new Stopwatch();
                //至今日Call累積設定點數 > 庫存Call總點數 && 今日由程式下單的Call總點數 < 今日配額 && 今日配額 > 0 && 非緊急停止狀態 && 自動下單開關 && form1關閉時跳出
                while (!formClose && !emergencyBreak && loopOrderTempSwitch && dayRentSum[0] > callSum && todayOrderStatus.TodayRentCall > 0 && todayOrderRecord.cacTodayAutoOrderPoints()[0]+ orderSettings.CallDownLim < todayOrderStatus.TodayRentCall )
                {
                    lots = 0;
                    tmpDec = 0;
                    //計算今日剩餘配額除以單口價格=今日需要下幾口單 => 再去計算是否能下設定單次最大口數，如果不行則以上面餘數來下單
                    if(orderSettings.CallDownLim>0) tmpDec = Math.Floor(todayOrderStatus.TodayRentCall - todayOrderRecord.cacTodayAutoOrderPoints()[0]) / orderSettings.CallDownLim;
                    if (partialDealRemains > 0) lots = partialDealRemains;
                    else if (tmpDec > 0 && (ushort)tmpDec < orderSettings.EachOrderLots) lots = (ushort)tmpDec;
                    else if (tmpDec > 0 && (ushort)tmpDec >= orderSettings.EachOrderLots) lots = (ushort)orderSettings.EachOrderLots;
                    else lots = 0;

                    StringBuilder sb = new StringBuilder();
                    long rtn = 0;
                    string[] combineSymbol = new string[4];
                    combineSymbol = findHundPointInHundred(); //找出100差裡在設定點數範圍的價差合約
                    if (string.IsNullOrEmpty(combineSymbol[0])) break;

                    //emergencyBreak判斷用碼錶
                    if (successCount == 0) stopWatch.Restart();  

                    //取得送單序號
                    long RequestId = tfcom.GetRequestId();
                    bool targetOrderFound = false;

                    ORDER_TYPE ordertype = ORDER_TYPE.OT_NEW;
                    SIDE_FLAG SideFlag = SIDE_FLAG.SF_SELL;//(tbTSide.Text[0] == 'B') ? SIDE_FLAG.SF_BUY : SIDE_FLAG.SF_SELL;

                    sb.Append("Call RequestId: " + RequestId).Append(Environment.NewLine);
                    sb.Append("SideFlag: " + SideFlag).Append(Environment.NewLine);
                    sb.Append("Symbol: " + combineSymbol[0]).Append(Environment.NewLine);
                    sb.Append("Lots: " + lots).Append(Environment.NewLine);
                    sb.Append("Price: " + orderSettings.CallDownLim.ToString() ).Append(Environment.NewLine);
                    //sb.Append("DownSafeDis: " + orderSettings.DownSafeDis).Append(Environment.NewLine);
                    //sb.Append("UpSafeDis: " + orderSettings.UpSafeDis).Append(Environment.NewLine);

                    AddInfo(sb.ToString());
                    opLog.txtLog(sb.ToString());
                    this.Invoke((Action)(() =>
                    {
                        if (formClose)
                        {
                            FlagInvokeLoopOrder = false;
                            return;
                        }
                        FlagInvokeLoopOrder = true;

                        //顯示此輪是否進行部分成交的再送單
                        tbPartialDeal.Text = partialDealRemains.ToString();
                        if (!String.IsNullOrEmpty(combineSymbol[0]) && decimal.Parse(combineSymbol[2]) > 0)
                        {
                            if (lots > 0 && tradeConnect && quoteConnect)
                            {
                                if(cbOrderNotReview.Checked)
                                {
                                    targetOrderFound = true;
                                    rtn = tfcom.Order(ordertype, MARKET_FLAG.MF_OPT, RequestId, accInfo.BrokerID, accInfo.TbTAccount, "",
                                                                combineSymbol[0] /*tb20Code.Text*/ , SideFlag, PRICE_FLAG.PF_SPECIFIED, orderSettings.CallDownLim, TIME_IN_FORCE.TIF_IOC, lots, POSITION_EFFECT.PE_OPEN, OFFICE_FLAG.OF_SPEEDY);
                                }
                                else
                                {
                                    if (MessageBox.Show(sb.ToString() + " 是否下單?", "問題", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                                    {
                                        targetOrderFound = true;
                                        rtn = tfcom.Order(ordertype, MARKET_FLAG.MF_OPT, RequestId, accInfo.BrokerID, accInfo.TbTAccount, "",
                                                                    combineSymbol[0] /*tb20Code.Text*/ , SideFlag, PRICE_FLAG.PF_SPECIFIED, orderSettings.CallDownLim, TIME_IN_FORCE.TIF_IOC, lots, POSITION_EFFECT.PE_OPEN, OFFICE_FLAG.OF_SPEEDY);
                                    }
                                    else
                                    {
                                        MessageBox.Show("Rcv No!");
                                    }
                                }
                            }
                            else if(lots == 0)
                            {
                                AddInfo("計算後的下單口數為零，點數已接近目標點數，再下可能會超出點數，應該早跳出while loop!請檢查!!!!");
                                AddInfoTG("計算後的下單口數為零，點數已接近目標點數，再下可能會超出點數，應該早跳出while loop!請檢查!!!!");
                                opLog.txtLog("計算後的下單口數為零，點數已接近目標點數，再下可能會超出點數，應該早跳出while loop!請檢查!!!!");
                            }
                            else opLog.txtLog("Call合適的價差單在安全距離內，所以不下單");
                        }
                        else opLog.txtLog("Call找無合適的價差單，所以不下單");

                        FlagInvokeLoopOrder = false;
                        if (formClose)
                        {
                            return;
                        }
                    }));
                        if (rtn != 0)
                        {
                            //MessageBox.Show(tfcom.GetOrderErrMsg(rtn));
                            AddInfo(tfcom.GetOrderErrMsg(rtn));
                            AddInfoTG(tfcom.GetOrderErrMsg(rtn));
                            opLog.txtLog(tfcom.GetOrderErrMsg(rtn));
                            failCount++;
                            targetOrderFound = false;
                        }

                    int timeoutCounter = 0;
                    while(targetOrderFound == true && timeoutCounter < 10 )
                    {
                            
                        if(orderFirstRPT.ContainsKey(RequestId))
                        {
                            break;
                        }
                        Thread.Sleep(500);
                        timeoutCounter++;
                    }
                    if (targetOrderFound==true && orderFirstRPT.ContainsKey(RequestId) && string.Compare(orderFirstRPT[RequestId][0], "收單") == 0)
                    {
                        Console.WriteLine("Call " + RequestId +": 收單成功");
                        AddInfo("Call " + RequestId + "收單成功");
                        opLog.txtLog("Call "+RequestId + "收單成功");
                        //檢查 orderFirstRPT[RequestId][1] (CNT) 是否收到
                        if (orderFirstRPT[RequestId][1] != null)
                        {
                            string cnt = orderFirstRPT[RequestId][1];
                            AddInfo("Call Rcv orderSecondRPT CNT: " + cnt);
                            opLog.txtLog("Call Rcv orderSecondRPT CNT: " + cnt);

                            AddInfo("Call 第二次下單回覆: " + orderRPT[cnt][(int)ORDERRPT_LIST.ORDERSTATUS]);
                            opLog.txtLog("Call 第二次下單回覆: " + orderRPT[cnt][(int)ORDERRPT_LIST.ORDERSTATUS]);

                            AddInfo("Call Rcv orderSecondRPT errorcode: " + orderRPT[cnt][(int)ORDERRPT_LIST.CODE]);
                            AddInfoTG("Call Rcv orderSecondRPT errorcode: " + orderRPT[cnt][(int)ORDERRPT_LIST.CODE]);
                            opLog.txtLog("Call Rcv orderSecondRPT errorcode: " + orderRPT[cnt][(int)ORDERRPT_LIST.CODE]);

                            if (dealRPT != null && dealRPT.ContainsKey(cnt))
                            {
                                opLog.txtLog("Call dealRPT[cnt][ORDERSTATUS]: " + dealRPT[cnt][(int)DEALRPT_LIST.ORDERSTATUS]);
                                if (string.Compare(dealRPT[cnt][(int)DEALRPT_LIST.ORDERSTATUS], "1") == 0)
                                {
                                    successCount++;
                                    todayOrderRecord.AddNewRecord(dealRPT[cnt]);
                                    AddInfo(cnt + " 收到成交回報 Call自動下單流程完成! 並將此單加入今日自動下單列表");
                                    AddInfoTG(cnt + " 收到成交回報 Call自動下單流程完成! 並將此單加入今日自動下單列表");
                                    opLog.txtLog(cnt + " 收到成交回報 Call自動下單流程完成! 並將此單加入今日自動下單列表");
                                    if (todayOrderRecord.cacTodayAutoOrderPoints()[0] >= todayOrderStatus.TodayRentCall)//如果今日由程式下單之Call總點數 >= 今日配額
                                    {
                                        AddInfo("今日由程式完成Call下單點數大於等於今日配額! 停止下單!");
                                        AddInfoTG("今日由程式完成Call下單點數大於等於今日配額! 停止下單!");
                                        opLog.txtLog("今日由程式完成Call下單點數大於等於今日配額! 停止下單!");
                                        partialDealRemains = 0;
                                    }
                                    else if (ushort.Parse(dealRPT[cnt][(int)DEALRPT_LIST.LEAVEQTY]) > 0) //如果是部分成交
                                    {
                                        AddInfo("Call部分成交，立即再送出未成交之數量!");
                                        AddInfoTG("Call部分成交，立即再送出未成交之數量!");
                                        opLog.txtLog("Call部分成交，立即再送出未成交之數量!");
                                        partialDealRemains = ushort.Parse(dealRPT[cnt][(int)DEALRPT_LIST.LEAVEQTY]);
                                    }
                                    else
                                    {
                                        AddInfo("Call完全成交!");
                                        AddInfoTG("Call完全成交!");
                                        opLog.txtLog("Call完全成交!");
                                        partialDealRemains = 0;
                                    }
                                    rtn = tfcom.RetrivePositionDetail("I", accInfo.BrokerID, accInfo.TbTAccount, "", "");
                                    if (rtn == -1)
                                    {
                                        MessageBox.Show("無查詢權限1618!!!");
                                        opLog.txtLog("無查詢權限1618!!!");
                                        AddInfoTG("無查詢權限1618!!!");
                                    }
                                    rtn = tfcom.RetrivePositionSum("I", accInfo.BrokerID, accInfo.TbTAccount, "", "");
                                    if (rtn == -1)
                                    {
                                        MessageBox.Show("無查詢權限1616!!!");
                                        opLog.txtLog("無查詢權限1616!!!");
                                        AddInfoTG("無查詢權限1616!!!");
                                    }
                                }
                                else
                                {
                                    failCount++;
                                    AddInfo("可能收到舊的成交回報 Call自動下單流程失敗! 繼續下一輪!!!!");
                                    AddInfoTG("可能收到舊的成交回報 Call自動下單流程失敗! 繼續下一輪!!!!");
                                    opLog.txtLog("可能收到舊的成交回報 Call自動下單流程失敗! 繼續下一輪!!!!");
                                    partialDealRemains = 0;
                                }
                            }
                            else
                            {
                                failCount++;
                                AddInfo("未收到成交回報 Call自動下單流程失敗! 請檢查!!!!");
                                AddInfoTG("未收到成交回報 Call自動下單流程失敗! 請檢查!!!!");
                                opLog.txtLog("未收到成交回報 Call自動下單流程失敗! 請檢查!!!!");
                                partialDealRemains = 0;
                            }
                        }
                        else
                        {
                            failCount++;
                            AddInfo("未收到第一回覆 Call自動下單流程失敗! 請檢查!!!!");
                            AddInfoTG("未收到第一回覆 Call自動下單流程失敗! 請檢查!!!!");
                            opLog.txtLog("未收到第一回覆 Call自動下單流程失敗! 請檢查!!!!");
                            partialDealRemains = 0;
                        }
                    }
                    else if(targetOrderFound == false)
                    {
                        AddInfo(RequestId + " Call未找到符合條件的合約!");
                        AddInfoTG(RequestId + " Call未找到符合條件的合約!");
                        opLog.txtLog(RequestId + " Call未找到符合條件的合約!");
                        partialDealRemains = 0;
                    }
                    else
                    {
                        failCount++;
                        AddInfo(RequestId + " Call收單失敗! 請檢查!!!!");
                        AddInfoTG(RequestId + " Call收單失敗! 請檢查!!!!");
                        opLog.txtLog(RequestId + " Call收單失敗! 請檢查!!!!");
                        partialDealRemains = 0;
                    }
                    //多少時間內 成交多少單或失敗多少單 則進保險
                    //保證金保險 保證金低於多少不下單
                    stopWatch.Stop();
                    double ts = stopWatch.Elapsed.TotalMinutes;

                    //opLog.txtLog("ts= "+ ts);
                    //opLog.txtLog("orderSettings.EmergencyCountMin= " + orderSettings.EmergencyCountMin);
                    //opLog.txtLog("orderSettings.EmergencySuccessRound= " + orderSettings.EmergencySuccessRound);
                    //opLog.txtLog("orderSettings.EmergencyFailRound= " + orderSettings.EmergencyFailRound);
                    //opLog.txtLog("successCount= " + successCount);
                    //opLog.txtLog("failCount= " + failCount);

                    if (equity < orderSettings.AtleastMargin)
                    {
                        AddInfo("保證金不足! 停止下單! 請檢查!!!!");
                        AddInfoTG("保證金不足! 停止下單! 請檢查!!!!");
                        opLog.txtLog("保證金不足! 停止下單! 請檢查!!!!");
                        emergencyBreak = true;
                    }
                    else if (ts <= (double)orderSettings.EmergencyCountMin && (successCount >= orderSettings.EmergencySuccessRound || failCount >= orderSettings.EmergencyFailRound))
                    {
                        AddInfo("Call在" + orderSettings.EmergencyCountMin + "分鐘內count次數超過，" + "successCount=" + successCount + " failCount=" + failCount + "  emergencyBreak!!!!");
                        AddInfoTG("Call在" + orderSettings.EmergencyCountMin + "分鐘內count次數超過，" + "successCount=" + successCount + " failCount=" + failCount + "  emergencyBreak!!!!");
                        opLog.txtLog("Call在" + orderSettings.EmergencyCountMin + "分鐘內count次數超過，" + "successCount=" + successCount + " failCount=" + failCount + "  emergencyBreak!!!!");
                        emergencyBreak = true;
                        MessageBox.Show("Call在" + orderSettings.EmergencyCountMin + "分鐘內count次數超過，" + "successCount=" + successCount + " failCount=" + failCount + "  emergencyBreak!!!!");
                    }
                    else if (ts > (double)orderSettings.EmergencyCountMin)
                    {
                        opLog.txtLog("Call " + orderSettings.EmergencyCountMin + "分鐘已到 stopWatch&successCount&failCount 歸零!");
                        stopWatch.Reset();
                        successCount = 0;
                        failCount = 0;
                    }
                    else stopWatch.Start();
                    todayOrderStatus.Read();//每次流程完需要更新一次，確認是否換日
                    if (partialDealRemains == 0 && !formClose)//如果沒有部份成交需要馬上再下單的才進Sleep
                        Thread.Sleep((int)orderSettings.EachOrderGap * 1000);
                }
                #endregion
                #region Put下單
                //Put
                partialDealRemains = 0;
                failCount = 0;
                successCount = 0;
                stopWatch.Reset();
                //至今日Put累積設定點數 > 庫存Put總點數 && 今日由程式下單的Put總點數 < 今日配額 && 今日配額 > 0 && 非緊急停止狀態 && 自動下單開關 && form1關閉時跳出
                while (!formClose && !emergencyBreak && loopOrderTempSwitch && dayRentSum[1] > putSum && todayOrderStatus.TodayRentPut > 0 && todayOrderRecord.cacTodayAutoOrderPoints()[1] + orderSettings.PutDownLim < todayOrderStatus.TodayRentPut)
                {
                    lots = 0;
                    tmpDec = 0;
                    if (orderSettings.PutDownLim > 0) tmpDec = Math.Floor(todayOrderStatus.TodayRentPut - todayOrderRecord.cacTodayAutoOrderPoints()[1]) / orderSettings.PutDownLim;
                    if (partialDealRemains > 0) lots = partialDealRemains;
                    else if (tmpDec > 0 && (ushort)tmpDec < orderSettings.EachOrderLots) lots = (ushort)tmpDec;
                    else if (tmpDec > 0 && (ushort)tmpDec >= orderSettings.EachOrderLots) lots = (ushort)orderSettings.EachOrderLots;
                    else lots = 0;

                    StringBuilder sb = new StringBuilder();
                    long rtn = 0;
                    string[] combineSymbol = new string[4];
                    combineSymbol = findHundPointInHundred(); //找出100差裡在設定點數範圍的價差合約
                    if (string.IsNullOrEmpty(combineSymbol[1])) break;

                    //emergencyBreak判斷用碼錶
                    if (successCount == 0) stopWatch.Restart();

                    //取得送單序號
                    long RequestId = tfcom.GetRequestId();
                    bool targetOrderFound = false;

                    ORDER_TYPE ordertype = ORDER_TYPE.OT_NEW;
                    SIDE_FLAG SideFlag = SIDE_FLAG.SF_SELL;//(tbTSide.Text[0] == 'B') ? SIDE_FLAG.SF_BUY : SIDE_FLAG.SF_SELL;

                    sb.Append("Put RequestId: " + RequestId).Append(Environment.NewLine);
                    sb.Append("SideFlag: " + SideFlag).Append(Environment.NewLine);
                    sb.Append("Symbol: " + combineSymbol[1]).Append(Environment.NewLine);
                    sb.Append("Lots: " + lots).Append(Environment.NewLine);
                    sb.Append("Price: " + orderSettings.PutDownLim.ToString()).Append(Environment.NewLine);
                    //sb.Append("DownSafeDis: " + orderSettings.DownSafeDis).Append(Environment.NewLine);
                    //sb.Append("UpSafeDis: " + orderSettings.UpSafeDis).Append(Environment.NewLine);

                    AddInfo(sb.ToString());
                    opLog.txtLog(sb.ToString());
                    this.Invoke((Action)(() =>
                    {
                        if (formClose)
                        {
                            FlagInvokeLoopOrder = false;
                            return;
                        }
                        FlagInvokeLoopOrder = true;

                        tbPartialDeal.Text = partialDealRemains.ToString();
                        if (!String.IsNullOrEmpty(combineSymbol[1]) && decimal.Parse(combineSymbol[3]) > 0)
                        {
                            if (lots > 0 && tradeConnect && quoteConnect)
                            {
                                if (cbOrderNotReview.Checked)
                                {
                                    targetOrderFound = true;
                                    rtn = tfcom.Order(ordertype, MARKET_FLAG.MF_OPT, RequestId, accInfo.BrokerID, accInfo.TbTAccount, "",
                                                                combineSymbol[1] /*tb20Code.Text*/ , SideFlag, PRICE_FLAG.PF_SPECIFIED, orderSettings.PutDownLim, TIME_IN_FORCE.TIF_IOC, lots, POSITION_EFFECT.PE_OPEN, OFFICE_FLAG.OF_SPEEDY);
                                }
                                else
                                {
                                    if (MessageBox.Show(sb.ToString() + "是否下單?", "問題", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                                    {
                                        targetOrderFound = true;
                                        rtn = tfcom.Order(ordertype, MARKET_FLAG.MF_OPT, RequestId, accInfo.BrokerID, accInfo.TbTAccount, "",
                                                                    combineSymbol[1] /*tb20Code.Text*/ , SideFlag, PRICE_FLAG.PF_SPECIFIED, orderSettings.PutDownLim, TIME_IN_FORCE.TIF_IOC, lots, POSITION_EFFECT.PE_OPEN, OFFICE_FLAG.OF_SPEEDY);
                                    }
                                    else
                                    {
                                        MessageBox.Show("Rcv No!");
                                    }
                                }
                            }
                            else if (lots == 0)
                            {
                                AddInfo("計算後的下單口數為零，點數已接近目標點數，再下可能會超出點數，應該早跳出while loop!請檢查!!!!");
                                AddInfoTG("計算後的下單口數為零，點數已接近目標點數，再下可能會超出點數，應該早跳出while loop!請檢查!!!!");
                                opLog.txtLog("計算後的下單口數為零，點數已接近目標點數，再下可能會超出點數，應該早跳出while loop!請檢查!!!!");
                            }
                            else opLog.txtLog("Put合適的價差單在安全距離內，所以不下單");
                        }
                        else opLog.txtLog("Put找無合適的價差單，所以不下單");

                        FlagInvokeLoopOrder = false;
                        if (formClose)
                        {
                            return;
                        }
                    }));
                    if (rtn != 0)
                    {
                        //MessageBox.Show(tfcom.GetOrderErrMsg(rtn));
                        AddInfo(tfcom.GetOrderErrMsg(rtn));
                        AddInfoTG(tfcom.GetOrderErrMsg(rtn));
                        opLog.txtLog(tfcom.GetOrderErrMsg(rtn));
                        failCount++;
                        targetOrderFound = false;
                    }

                    int timeoutCounter = 0;
                    while (targetOrderFound == true && timeoutCounter < 10)
                    {

                        if (orderFirstRPT.ContainsKey(RequestId))
                        {
                            break;
                        }
                        Thread.Sleep(500);
                        timeoutCounter++;
                    }
                    if (targetOrderFound == true && orderFirstRPT.ContainsKey(RequestId) && string.Compare(orderFirstRPT[RequestId][0], "收單") == 0)
                    {
                        Console.WriteLine("Put " + RequestId + ": 收單成功");
                        AddInfo("Put " + RequestId + "收單成功");
                        opLog.txtLog("Put " + RequestId + "收單成功");
                        //檢查 orderFirstRPT[RequestId][1] (CNT) 是否收到
                        if (orderFirstRPT[RequestId][1] != null)
                        {
                            string cnt = orderFirstRPT[RequestId][1];
                            AddInfo("Put Rcv orderSecondRPT CNT: " + cnt);
                            opLog.txtLog("Put Rcv orderSecondRPT CNT: " + cnt);

                            AddInfo("Put 第二次下單回覆: " + orderRPT[cnt][(int)ORDERRPT_LIST.ORDERSTATUS]);
                            opLog.txtLog("Put 第二次下單回覆: " + orderRPT[cnt][(int)ORDERRPT_LIST.ORDERSTATUS]);

                            AddInfo("Put Rcv orderSecondRPT errorcode: " + orderRPT[cnt][(int)ORDERRPT_LIST.CODE]);
                            AddInfoTG("Put Rcv orderSecondRPT errorcode: " + orderRPT[cnt][(int)ORDERRPT_LIST.CODE]);
                            opLog.txtLog("Put Rcv orderSecondRPT errorcode: " + orderRPT[cnt][(int)ORDERRPT_LIST.CODE]);

                            if (dealRPT != null && dealRPT.ContainsKey(cnt))
                            {
                                opLog.txtLog("Put dealRPT[cnt][ORDERSTATUS]: " + dealRPT[cnt][(int)DEALRPT_LIST.ORDERSTATUS]);
                                if (string.Compare(dealRPT[cnt][(int)DEALRPT_LIST.ORDERSTATUS], "1") == 0)
                                {
                                    successCount++;
                                    todayOrderRecord.AddNewRecord(dealRPT[cnt]);
                                    AddInfo(cnt + " 收到成交回報 Put自動下單流程完成! 並將此單加入今日自動下單列表");
                                    AddInfoTG(cnt + " 收到成交回報 Put自動下單流程完成! 並將此單加入今日自動下單列表");
                                    opLog.txtLog(cnt + " 收到成交回報 Put自動下單流程完成! 並將此單加入今日自動下單列表");
                                    if (todayOrderRecord.cacTodayAutoOrderPoints()[1] >= todayOrderStatus.TodayRentPut)//如果今日由程式下單之Put總點數 >= 今日配額
                                    {
                                        AddInfo("今日由程式完成Put下單點數大於等於今日配額! 停止下單!");
                                        AddInfoTG("今日由程式完成Put下單點數大於等於今日配額! 停止下單!");
                                        opLog.txtLog("今日由程式完成Put下單點數大於等於今日配額! 停止下單!");
                                        partialDealRemains = 0;
                                    }
                                    else if (ushort.Parse(dealRPT[cnt][(int)DEALRPT_LIST.LEAVEQTY]) > 0) //如果是部分成交
                                    {
                                        AddInfo("Put部分成交，立即再送出未成交之數量!");
                                        AddInfoTG("Put部分成交，立即再送出未成交之數量!");
                                        opLog.txtLog("Put部分成交，立即再送出未成交之數量!");
                                        partialDealRemains = ushort.Parse(dealRPT[cnt][(int)DEALRPT_LIST.LEAVEQTY]);
                                    }
                                    else
                                    {
                                        AddInfo("Put完全成交!");
                                        AddInfoTG("Put完全成交!");
                                        opLog.txtLog("Put完全成交!");
                                        partialDealRemains = 0;
                                    }
                                    rtn = tfcom.RetrivePositionDetail("I", accInfo.BrokerID, accInfo.TbTAccount, "", "");
                                    if (rtn == -1)
                                    {
                                        MessageBox.Show("無查詢權限1618!!!");
                                        opLog.txtLog("無查詢權限1618!!!");
                                        AddInfoTG("無查詢權限1618!!!");
                                    }
                                    rtn = tfcom.RetrivePositionSum("I", accInfo.BrokerID, accInfo.TbTAccount, "", "");
                                    if (rtn == -1)
                                    {
                                        MessageBox.Show("無查詢權限1616!!!");
                                        opLog.txtLog("無查詢權限1616!!!");
                                        AddInfoTG("無查詢權限1616!!!");
                                    }
                                }
                                else
                                {
                                    failCount++;
                                    AddInfo("可能收到舊的成交回報 Put自動下單流程失敗! 繼續下一輪!!!!");
                                    AddInfoTG("可能收到舊的成交回報 Put自動下單流程失敗! 繼續下一輪!!!!");
                                    opLog.txtLog("可能收到舊的成交回報 Put自動下單流程失敗! 繼續下一輪!!!!");
                                    partialDealRemains = 0;
                                }
                            }
                            else
                            {
                                failCount++;
                                AddInfo("未收到成交回報 Put自動下單流程失敗! 請檢查!!!!");
                                AddInfoTG("未收到成交回報 Put自動下單流程失敗! 請檢查!!!!");
                                opLog.txtLog("未收到成交回報 Put自動下單流程失敗! 請檢查!!!!");
                                partialDealRemains = 0;
                            }
                        }
                        else
                        {
                            failCount++;
                            AddInfo("未收到第一回覆 Put自動下單流程失敗! 請檢查!!!!");
                            AddInfoTG("未收到第一回覆 Put自動下單流程失敗! 請檢查!!!!");
                            opLog.txtLog("未收到第一回覆 Put自動下單流程失敗! 請檢查!!!!");
                            partialDealRemains = 0;
                        }
                    }
                    else if (targetOrderFound == false)
                    {
                        AddInfo(RequestId + " Put未找到符合條件的合約!");
                        AddInfoTG(RequestId + " Put未找到符合條件的合約!");
                        opLog.txtLog(RequestId + " Put未找到符合條件的合約!");
                        partialDealRemains = 0;
                    }
                    else
                    {
                        failCount++;
                        AddInfo(RequestId + " Put收單失敗! 請檢查!!!!");
                        AddInfoTG(RequestId + " Put收單失敗! 請檢查!!!!");
                        opLog.txtLog(RequestId + " Put收單失敗! 請檢查!!!!");
                        partialDealRemains = 0;
                    }
                    //多少時間內 成交多少單或失敗多少單 則進保險
                    //保證金保險 保證金低於多少不下單
                    stopWatch.Stop();
                    double ts = stopWatch.Elapsed.TotalMinutes;
                    if (equity < orderSettings.AtleastMargin)
                    {
                        AddInfo("保證金不足! 停止下單! 請檢查!!!!");
                        AddInfoTG("保證金不足! 停止下單! 請檢查!!!!");
                        opLog.txtLog("保證金不足! 停止下單! 請檢查!!!!");
                        emergencyBreak = true;
                    }
                    else if (ts <= (double)orderSettings.EmergencyCountMin && (successCount >= orderSettings.EmergencySuccessRound || failCount >= orderSettings.EmergencyFailRound))
                    {
                        AddInfo("Put在" + orderSettings.EmergencyCountMin + "分鐘內count次數超過，" + "successCount=" + successCount + " failCount=" + failCount + "  emergencyBreak!!!!");
                        AddInfoTG("Put在" + orderSettings.EmergencyCountMin + "分鐘內count次數超過，" + "successCount=" + successCount + " failCount=" + failCount + "  emergencyBreak!!!!");
                        opLog.txtLog("Put在" + orderSettings.EmergencyCountMin + "分鐘內count次數超過，" + "successCount=" + successCount + " failCount=" + failCount + "  emergencyBreak!!!!");
                        emergencyBreak = true;
                        MessageBox.Show("Put在" + orderSettings.EmergencyCountMin + "分鐘內count次數超過，" + "successCount=" + successCount + " failCount=" + failCount + "  emergencyBreak!!!!");
                    }
                    else if (ts > (double)orderSettings.EmergencyCountMin)
                    {
                        opLog.txtLog("Put " + orderSettings.EmergencyCountMin + "分鐘已到 stopWatch&successCount&failCount 歸零!");
                        stopWatch.Reset();
                        successCount = 0;
                        failCount = 0;
                    }
                    else stopWatch.Start();
                    todayOrderStatus.Read();//每次流程完需要更新一次，確認是否換日
                    if (partialDealRemains == 0 && !formClose)//如果沒有部份成交需要馬上再下單的才進Sleep
                        Thread.Sleep((int)orderSettings.EachOrderGap * 1000);
                }
                #endregion
                #region Call停損
                int profitStopFailCount = 0;
                int profitStopSuccessCount = 0;
                string[] stoplossCombineSymbol = new string[4];
                stopWatch.Reset();
                while (!formClose && !emergencyBreakStopLoss && loopOrderTempSwitch && callSum > dayRentSum[0] && callSum > callCost*orderSettings.StopLossMagn )//庫存總點數 > 累計設定點數 && 庫存總點數 > 庫存成本總點數*停損倍率
                {
                    lots = 0;

                    lots = orderSettings.StopLossLots;
                    stoplossCombineSymbol = new string[4];
                    stoplossCombineSymbol = findNearestNowPriceInStock(); //找出庫存裡面最接近價平的價差合約
                    if (ushort.Parse(symbolOPriceMap[stoplossCombineSymbol[0]][(int)OPTION_LIST.STOCK]) < lots && !string.IsNullOrEmpty(symbolOPriceMap[stoplossCombineSymbol[0]][(int)OPTION_LIST.STOCK]))
                        lots = ushort.Parse(symbolOPriceMap[stoplossCombineSymbol[0]][(int)OPTION_LIST.STOCK]);

                    StringBuilder sb = new StringBuilder();
                    long rtn = 0;
                    //emergencyBreakStopLoss判斷用碼錶
                    if (profitStopSuccessCount == 0) stopWatch.Restart();

                    //取得送單序號
                    long RequestId = tfcom.GetRequestId();
                    bool targetOrderFound = false;

                    ORDER_TYPE ordertype = ORDER_TYPE.OT_NEW;
                    SIDE_FLAG SideFlag = SIDE_FLAG.SF_BUY;//(tbTSide.Text[0] == 'B') ? SIDE_FLAG.SF_BUY : SIDE_FLAG.SF_SELL; //我們平常價差單停損目前都是要用Buy

                    sb.Append("Call停損 RequestId: " + RequestId).Append(Environment.NewLine);
                    sb.Append("SideFlag: " + SideFlag).Append(Environment.NewLine);
                    sb.Append("Symbol: " + stoplossCombineSymbol[0]).Append(Environment.NewLine);
                    sb.Append("Lots: " + lots).Append(Environment.NewLine);
                    sb.Append("Price: " + (decimal.Parse(stoplossCombineSymbol[2])).ToString()).Append(Environment.NewLine);
                    //sb.Append("DownSafeDis: " + orderSettings.DownSafeDis).Append(Environment.NewLine);
                    //sb.Append("UpSafeDis: " + orderSettings.UpSafeDis).Append(Environment.NewLine);

                    AddInfo(sb.ToString());
                    opLog.txtLog(sb.ToString());
                    this.Invoke((Action)(() =>
                    {
                        if (formClose)
                        {
                            FlagInvokeLoopOrder = false;
                            return;
                        }
                        FlagInvokeLoopOrder = true;

                        if (!String.IsNullOrEmpty(stoplossCombineSymbol[0]) && decimal.Parse(stoplossCombineSymbol[2]) > 0 && tradeConnect && quoteConnect )
                        {
                            if (cbOrderNotReview.Checked)
                            {
                                targetOrderFound = true;
                                rtn = tfcom.Order(ordertype, MARKET_FLAG.MF_OPT, RequestId, accInfo.BrokerID, accInfo.TbTAccount, "",
                                                            stoplossCombineSymbol[0] /*tb20Code.Text*/ , SideFlag, PRICE_FLAG.PF_SPECIFIED, decimal.Parse(stoplossCombineSymbol[2]) * orderSettings.StopLossMagn, TIME_IN_FORCE.TIF_IOC, lots, POSITION_EFFECT.PE_CLOSE, OFFICE_FLAG.OF_SPEEDY);
                            }
                            else
                            {
                                if (MessageBox.Show(sb.ToString() + "是否下停損單?", "問題", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                                {
                                    targetOrderFound = true;
                                    rtn = tfcom.Order(ordertype, MARKET_FLAG.MF_OPT, RequestId, accInfo.BrokerID, accInfo.TbTAccount, "",
                                                                stoplossCombineSymbol[0] /*tb20Code.Text*/ , SideFlag, PRICE_FLAG.PF_SPECIFIED, decimal.Parse(stoplossCombineSymbol[2]) * orderSettings.StopLossMagn, TIME_IN_FORCE.TIF_IOC, lots, POSITION_EFFECT.PE_CLOSE, OFFICE_FLAG.OF_SPEEDY);
                                }
                                else
                                {
                                    MessageBox.Show("Rcv No!");
                                }
                            }

                        }
                        else
                        {
                            opLog.txtLog("Call停損找無合適的合約，請檢查!!!!");
                            AddInfoTG("Call停損找無合適的合約，請檢查!!!!");
                        }

                        FlagInvokeLoopOrder = false;
                        if (formClose)
                        {
                            return;
                        }
                    }));
                    if (rtn != 0)
                    {
                        //MessageBox.Show(tfcom.GetOrderErrMsg(rtn));
                        AddInfo(tfcom.GetOrderErrMsg(rtn));
                        AddInfoTG(tfcom.GetOrderErrMsg(rtn));
                        opLog.txtLog(tfcom.GetOrderErrMsg(rtn));
                        profitStopFailCount++;
                        targetOrderFound = false;
                    }

                    int timeoutCounter = 0;
                    while (targetOrderFound == true && timeoutCounter < 10)
                    {

                        if (orderFirstRPT.ContainsKey(RequestId))
                        {
                            break;
                        }
                        Thread.Sleep(500);
                        timeoutCounter++;
                    }
                    if (targetOrderFound == true && orderFirstRPT.ContainsKey(RequestId) && string.Compare(orderFirstRPT[RequestId][0], "收單") == 0)
                    {
                        Console.WriteLine("Call停損 " + RequestId + ": 收單成功");
                        AddInfo("Call停損 " + RequestId + "收單成功");
                        opLog.txtLog("Call停損 " + RequestId + "收單成功");
                        //檢查 orderFirstRPT[RequestId][1] (CNT) 是否收到
                        if (orderFirstRPT[RequestId][1] != null)
                        {
                            string cnt = orderFirstRPT[RequestId][1];
                            AddInfo("Call停損 Rcv orderSecondRPT CNT: " + cnt);
                            opLog.txtLog("Call停損 Rcv orderSecondRPT CNT: " + cnt);

                            AddInfo("Call停損 第二次下單回覆: " + orderRPT[cnt][(int)ORDERRPT_LIST.ORDERSTATUS]);
                            opLog.txtLog("Call停損 第二次下單回覆: " + orderRPT[cnt][(int)ORDERRPT_LIST.ORDERSTATUS]);

                            AddInfo("Call停損 Rcv orderSecondRPT errorcode: " + orderRPT[cnt][(int)ORDERRPT_LIST.CODE]);
                            AddInfoTG("Call停損 Rcv orderSecondRPT errorcode: " + orderRPT[cnt][(int)ORDERRPT_LIST.CODE]);
                            opLog.txtLog("Call停損 Rcv orderSecondRPT errorcode: " + orderRPT[cnt][(int)ORDERRPT_LIST.CODE]);

                            if (dealRPT != null && dealRPT.ContainsKey(cnt))
                            {
                                opLog.txtLog("Call停損 dealRPT[cnt][ORDERSTATUS]: " + dealRPT[cnt][(int)DEALRPT_LIST.ORDERSTATUS]);
                                if (string.Compare(dealRPT[cnt][(int)DEALRPT_LIST.ORDERSTATUS], "1") == 0)
                                {
                                    profitStopSuccessCount++;
                                    AddInfo(cnt + " 收到成交回報 Call停損自動下單流程完成!");
                                    AddInfoTG(cnt + " 收到成交回報 Call停損自動下單流程完成!");
                                    opLog.txtLog(cnt + " 收到成交回報 Call停損自動下單流程完成!");
                                    if (ushort.Parse(dealRPT[cnt][(int)DEALRPT_LIST.LEAVEQTY]) > 0) //如果是部分成交
                                    {
                                        AddInfo("Call停損部分成交!");
                                        AddInfoTG("Call停損部分成交!");
                                        opLog.txtLog("Call停損部分成交!");
                                    }
                                    else
                                    {
                                        AddInfo("Call停損完全成交!");
                                        AddInfoTG("Call停損完全成交!");
                                        opLog.txtLog("Call停損完全成交!");
                                    }
                                    rtn = tfcom.RetrivePositionDetail("I", accInfo.BrokerID, accInfo.TbTAccount, "", "");
                                    if (rtn == -1)
                                    {
                                        MessageBox.Show("無查詢權限1618!!!");
                                        opLog.txtLog("無查詢權限1618!!!");
                                        AddInfoTG("無查詢權限1618!!!");
                                    }
                                    rtn = tfcom.RetrivePositionSum("I", accInfo.BrokerID, accInfo.TbTAccount, "", "");
                                    if (rtn == -1)
                                    {
                                        MessageBox.Show("無查詢權限1616!!!");
                                        opLog.txtLog("無查詢權限1616!!!");
                                        AddInfoTG("無查詢權限1616!!!");
                                    }
                                }
                                else
                                {
                                    profitStopFailCount++;
                                    AddInfo("可能收到舊的成交回報 Call停損自動下單流程失敗! 請檢查!!!!");
                                    AddInfoTG("可能收到舊的成交回報 Call停損自動下單流程失敗! 請檢查!!!!");
                                    opLog.txtLog("可能收到舊的成交回報 Call停損自動下單流程失敗! 請檢查!!!!");
                                }
                            }
                            else
                            {
                                profitStopFailCount++;
                                AddInfo("未收到成交回報 Call停損自動下單流程失敗! 重新送單!!!!");
                                AddInfoTG("未收到成交回報 Call停損自動下單流程失敗! 重新送單!!!!");
                                opLog.txtLog("未收到成交回報 Call停損自動下單流程失敗! 重新送單!!!!");
                            }
                        }
                    }
                    else if (targetOrderFound == false)
                    {
                        profitStopFailCount++;
                        AddInfo(RequestId + " Call停損未找到符合條件的合約，請檢查!!!!");
                        AddInfoTG(RequestId + " Call停損未找到符合條件的合約，請檢查!!!!");
                        opLog.txtLog(RequestId + " Call停損未找到符合條件的合約，請檢查!!!!");
                    }
                    else
                    {
                        profitStopFailCount++;
                        AddInfo(RequestId + " Call停損收單失敗，請檢查!!!!");
                        AddInfoTG(RequestId + " Call停損收單失敗，請檢查!!!!");
                        opLog.txtLog(RequestId + " Call停損收單失敗，請檢查!!!!");
                    }
                    
                    //多少時間內 成交多少單或失敗多少單 則進保險
                    stopWatch.Stop();
                    double ts = stopWatch.Elapsed.TotalMinutes;
                    if (ts <= (double)orderSettings.ProfitStopCountMin && (profitStopSuccessCount >= orderSettings.ProfitStopSuccessRound || profitStopFailCount >= orderSettings.ProfitStopFailRound))
                    {
                        AddInfo("Call停損在" + orderSettings.ProfitStopCountMin + "分鐘內count次數超過，" + "profitStopSuccessCount=" + profitStopSuccessCount + " profitStopFailCount=" + profitStopFailCount + "  emergencyBreakStopLoss!!!!");
                        AddInfoTG("Call停損在" + orderSettings.ProfitStopCountMin + "分鐘內count次數超過，" + "profitStopSuccessCount=" + profitStopSuccessCount + " profitStopFailCount=" + profitStopFailCount + "  emergencyBreakStopLoss!!!!");
                        opLog.txtLog("Call停損在" + orderSettings.ProfitStopCountMin + "分鐘內count次數超過，" + "profitStopSuccessCount=" + profitStopSuccessCount + " profitStopFailCount=" + profitStopFailCount + "  emergencyBreakStopLoss!!!!");
                        emergencyBreakStopLoss = true;
                        MessageBox.Show("Call停損在" + orderSettings.ProfitStopCountMin + "分鐘內count次數超過，" + "profitStopSuccessCount=" + profitStopSuccessCount + " profitStopFailCount=" + profitStopFailCount + "  emergencyBreakStopLoss!!!!");
                    }
                    else if (ts > (double)orderSettings.ProfitStopCountMin)
                    {
                        opLog.txtLog("Call停損 " + orderSettings.ProfitStopCountMin + "分鐘已到 stopWatch&profitStopSuccessCount&profitStopFailCount 歸零!");
                        stopWatch.Reset();
                        profitStopSuccessCount = 0;
                        profitStopFailCount = 0;
                    }
                    else stopWatch.Start();
                    if(!formClose) Thread.Sleep(Convert.ToInt32(orderSettings.StopLossGap * 1000));
                }
                #endregion
                #region Put停損
                profitStopFailCount = 0;
                profitStopSuccessCount = 0;
                stoplossCombineSymbol = new string[4];
                stopWatch.Reset();
                while (!formClose && !emergencyBreakStopLoss && loopOrderTempSwitch && putSum > dayRentSum[1] && putSum > putCost * orderSettings.StopLossMagn)//庫存總點數 > 累計設定點數 && 庫存總點數 > 庫存成本總點數*停損倍率
                {
                    lots = 0;

                    lots = orderSettings.StopLossLots;
                    stoplossCombineSymbol = new string[4];
                    stoplossCombineSymbol = findNearestNowPriceInStock(); //找出庫存裡面最接近價平的價差合約
                    if (ushort.Parse(symbolOPriceMap[stoplossCombineSymbol[1]][(int)OPTION_LIST.STOCK]) < lots && !string.IsNullOrEmpty(symbolOPriceMap[stoplossCombineSymbol[1]][(int)OPTION_LIST.STOCK]))
                        lots = ushort.Parse(symbolOPriceMap[stoplossCombineSymbol[1]][(int)OPTION_LIST.STOCK]);

                    StringBuilder sb = new StringBuilder();
                    long rtn = 0;
                    //emergencyBreakStopLoss判斷用碼錶
                    if (profitStopSuccessCount == 0) stopWatch.Restart();

                    //取得送單序號
                    long RequestId = tfcom.GetRequestId();
                    bool targetOrderFound = false;

                    ORDER_TYPE ordertype = ORDER_TYPE.OT_NEW;
                    SIDE_FLAG SideFlag = SIDE_FLAG.SF_BUY;//(tbTSide.Text[0] == 'B') ? SIDE_FLAG.SF_BUY : SIDE_FLAG.SF_SELL; //我們平常價差單停損目前都是要用Buy

                    sb.Append("Put停損 RequestId: " + RequestId).Append(Environment.NewLine);
                    sb.Append("SideFlag: " + SideFlag).Append(Environment.NewLine);
                    sb.Append("Symbol: " + stoplossCombineSymbol[1]).Append(Environment.NewLine);
                    sb.Append("Lots: " + lots).Append(Environment.NewLine);
                    sb.Append("Price: " + (decimal.Parse(stoplossCombineSymbol[3])).ToString()).Append(Environment.NewLine);
                    //sb.Append("DownSafeDis: " + orderSettings.DownSafeDis).Append(Environment.NewLine);
                    //sb.Append("UpSafeDis: " + orderSettings.UpSafeDis).Append(Environment.NewLine);

                    AddInfo(sb.ToString());
                    opLog.txtLog(sb.ToString());
                    this.Invoke((Action)(() =>
                    {
                        if (formClose)
                        {
                            FlagInvokeLoopOrder = false;
                            return;
                        }
                        FlagInvokeLoopOrder = true;

                        if (!String.IsNullOrEmpty(stoplossCombineSymbol[1]) && decimal.Parse(stoplossCombineSymbol[3]) > 0 && tradeConnect && quoteConnect)
                        {
                            if (cbOrderNotReview.Checked)
                            {
                                targetOrderFound = true;
                                rtn = tfcom.Order(ordertype, MARKET_FLAG.MF_OPT, RequestId, accInfo.BrokerID, accInfo.TbTAccount, "",
                                                            stoplossCombineSymbol[1] /*tb20Code.Text*/ , SideFlag, PRICE_FLAG.PF_SPECIFIED, decimal.Parse(stoplossCombineSymbol[3]) * orderSettings.StopLossMagn, TIME_IN_FORCE.TIF_IOC, lots, POSITION_EFFECT.PE_CLOSE, OFFICE_FLAG.OF_SPEEDY);
                            }
                            else
                            {
                                if (MessageBox.Show(sb.ToString() + "是否下停損單?", "問題", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                                {
                                    targetOrderFound = true;
                                    rtn = tfcom.Order(ordertype, MARKET_FLAG.MF_OPT, RequestId, accInfo.BrokerID, accInfo.TbTAccount, "",
                                                                stoplossCombineSymbol[1] /*tb20Code.Text*/ , SideFlag, PRICE_FLAG.PF_SPECIFIED, decimal.Parse(stoplossCombineSymbol[3]) * orderSettings.StopLossMagn, TIME_IN_FORCE.TIF_IOC, lots, POSITION_EFFECT.PE_CLOSE, OFFICE_FLAG.OF_SPEEDY);
                                }
                                else
                                {
                                    MessageBox.Show("Rcv No!");
                                }
                            }
                        }
                        else
                        {
                            opLog.txtLog("Put停損找無合適的合約，請檢查!!!!");
                            AddInfoTG("Put停損找無合適的合約，請檢查!!!!");
                        }

                        FlagInvokeLoopOrder = false;
                        if (formClose)
                        {
                            return;
                        }
                    }));
                    if (rtn != 0)
                    {
                        //MessageBox.Show(tfcom.GetOrderErrMsg(rtn));
                        AddInfo(tfcom.GetOrderErrMsg(rtn));
                        AddInfoTG(tfcom.GetOrderErrMsg(rtn));
                        opLog.txtLog(tfcom.GetOrderErrMsg(rtn));
                        profitStopFailCount++;
                        targetOrderFound = false;
                    }

                    int timeoutCounter = 0;
                    while (targetOrderFound == true && timeoutCounter < 10)
                    {

                        if (orderFirstRPT.ContainsKey(RequestId))
                        {
                            break;
                        }
                        Thread.Sleep(500);
                        timeoutCounter++;
                    }
                    if (targetOrderFound == true && orderFirstRPT.ContainsKey(RequestId) && string.Compare(orderFirstRPT[RequestId][0], "收單") == 0)
                    {
                        Console.WriteLine("Put停損 " + RequestId + ": 收單成功");
                        AddInfo("Put停損 " + RequestId + "收單成功");
                        opLog.txtLog("Put停損 " + RequestId + "收單成功");
                        //檢查 orderFirstRPT[RequestId][1] (CNT) 是否收到
                        if (orderFirstRPT[RequestId][1] != null)
                        {
                            string cnt = orderFirstRPT[RequestId][1];
                            AddInfo("Put停損 Rcv orderSecondRPT CNT: " + cnt);
                            opLog.txtLog("Put停損 Rcv orderSecondRPT CNT: " + cnt);

                            AddInfo("Put停損 第二次下單回覆: " + orderRPT[cnt][(int)ORDERRPT_LIST.ORDERSTATUS]);
                            opLog.txtLog("Put停損 第二次下單回覆: " + orderRPT[cnt][(int)ORDERRPT_LIST.ORDERSTATUS]);

                            AddInfo("Put停損 Rcv orderSecondRPT errorcode: " + orderRPT[cnt][(int)ORDERRPT_LIST.CODE]);
                            AddInfoTG("Put停損 Rcv orderSecondRPT errorcode: " + orderRPT[cnt][(int)ORDERRPT_LIST.CODE]);
                            opLog.txtLog("Put停損 Rcv orderSecondRPT errorcode: " + orderRPT[cnt][(int)ORDERRPT_LIST.CODE]);

                            if (dealRPT != null && dealRPT.ContainsKey(cnt))
                            {
                                opLog.txtLog("Put停損 dealRPT[cnt][ORDERSTATUS]: " + dealRPT[cnt][(int)DEALRPT_LIST.ORDERSTATUS]);
                                if (string.Compare(dealRPT[cnt][(int)DEALRPT_LIST.ORDERSTATUS], "1") == 0)
                                {
                                    profitStopSuccessCount++;
                                    AddInfo(cnt + " 收到成交回報 Put停損自動下單流程完成!");
                                    AddInfoTG(cnt + " 收到成交回報 Put停損自動下單流程完成!");
                                    opLog.txtLog(cnt + " 收到成交回報 Put停損自動下單流程完成!");
                                    if (ushort.Parse(dealRPT[cnt][(int)DEALRPT_LIST.LEAVEQTY]) > 0) //如果是部分成交
                                    {
                                        AddInfo("Put停損部分成交!");
                                        AddInfoTG("Put停損部分成交!");
                                        opLog.txtLog("Put停損部分成交!");
                                    }
                                    else
                                    {
                                        AddInfo("Put停損完全成交!");
                                        AddInfoTG("Put停損完全成交!");
                                        opLog.txtLog("Put停損完全成交!");
                                    }
                                    rtn = tfcom.RetrivePositionDetail("I", accInfo.BrokerID, accInfo.TbTAccount, "", "");
                                    if (rtn == -1)
                                    {
                                        MessageBox.Show("無查詢權限1618!!!");
                                        opLog.txtLog("無查詢權限1618!!!");
                                        AddInfoTG("無查詢權限1618!!!");
                                    }
                                    rtn = tfcom.RetrivePositionSum("I", accInfo.BrokerID, accInfo.TbTAccount, "", "");
                                    if (rtn == -1)
                                    {
                                        MessageBox.Show("無查詢權限1616!!!");
                                        opLog.txtLog("無查詢權限1616!!!");
                                        AddInfoTG("無查詢權限1616!!!");
                                    }
                                }
                                else
                                {
                                    profitStopFailCount++;
                                    AddInfo("可能收到舊的成交回報 Put停損自動下單流程失敗! 請檢查!!!!");
                                    AddInfoTG("可能收到舊的成交回報 Put停損自動下單流程失敗! 請檢查!!!!");
                                    opLog.txtLog("可能收到舊的成交回報 Put停損自動下單流程失敗! 請檢查!!!!");
                                }
                            }
                            else
                            {
                                profitStopFailCount++;
                                AddInfo("未收到成交回報 Put停損自動下單流程失敗! 重新送單!!!!");
                                AddInfoTG("未收到成交回報 Put停損自動下單流程失敗! 重新送單!!!!");
                                opLog.txtLog("未收到成交回報 Put停損自動下單流程失敗! 重新送單!!!!");
                            }
                        }
                    }
                    else if (targetOrderFound == false)
                    {
                        profitStopFailCount++;
                        AddInfo(RequestId + " Put停損未找到符合條件的合約，請檢查!!!!");
                        AddInfoTG(RequestId + " Put停損未找到符合條件的合約，請檢查!!!!");
                        opLog.txtLog(RequestId + " Put停損未找到符合條件的合約，請檢查!!!!");
                    }
                    else
                    {
                        profitStopFailCount++;
                        AddInfo(RequestId + " Put停損收單失敗，請檢查!!!!");
                        AddInfoTG(RequestId + " Put停損收單失敗，請檢查!!!!");
                        opLog.txtLog(RequestId + " Put停損收單失敗，請檢查!!!!");
                    }

                    //多少時間內 成交多少單或失敗多少單 則進保險
                    stopWatch.Stop();
                    double ts = stopWatch.Elapsed.TotalMinutes;
                    if (ts <= (double)orderSettings.ProfitStopCountMin && (profitStopSuccessCount >= orderSettings.ProfitStopSuccessRound || profitStopFailCount >= orderSettings.ProfitStopFailRound))
                    {
                        AddInfo("Put停損在" + orderSettings.ProfitStopCountMin + "分鐘內count次數超過，" + "profitStopSuccessCount=" + profitStopSuccessCount + " profitStopFailCount=" + profitStopFailCount + "  emergencyBreakStopLoss!!!!");
                        AddInfoTG("Put停損在" + orderSettings.ProfitStopCountMin + "分鐘內count次數超過，" + "profitStopSuccessCount=" + profitStopSuccessCount + " profitStopFailCount=" + profitStopFailCount + "  emergencyBreakStopLoss!!!!");
                        opLog.txtLog("Put停損在" + orderSettings.ProfitStopCountMin + "分鐘內count次數超過，" + "profitStopSuccessCount=" + profitStopSuccessCount + " profitStopFailCount=" + profitStopFailCount + "  emergencyBreakStopLoss!!!!");
                        emergencyBreakStopLoss = true;
                        MessageBox.Show("Put停損在" + orderSettings.ProfitStopCountMin + "分鐘內count次數超過，" + "profitStopSuccessCount=" + profitStopSuccessCount + " profitStopFailCount=" + profitStopFailCount + "  emergencyBreakStopLoss!!!!");
                    }
                    else if (ts > (double)orderSettings.ProfitStopCountMin)
                    {
                        opLog.txtLog("Put停損 " + orderSettings.ProfitStopCountMin + "分鐘已到 stopWatch&profitStopSuccessCount&profitStopFailCount 歸零!");
                        stopWatch.Reset();
                        profitStopSuccessCount = 0;
                        profitStopFailCount = 0;
                    }
                    else stopWatch.Start();
                    if(!formClose) Thread.Sleep(Convert.ToInt32(orderSettings.StopLossGap * 1000));
                }
                #endregion
                #region Call停利
                profitStopFailCount = 0;
                profitStopSuccessCount = 0;
                string[] stopprofitCombineSymbol = new string[4];
                stopWatch.Reset();
                while (!formClose && !emergencyBreakStopLoss && loopOrderTempSwitch && !string.IsNullOrEmpty(findStopProfitInStock()[0]))
                {
                    lots = 0;
                    lots = orderSettings.TakeProfitLots;
                    stopprofitCombineSymbol = new string[7];
                    stopprofitCombineSymbol = findStopProfitInStock(); //找出庫存裡面需要停利的合約
                    if (ushort.Parse(stopprofitCombineSymbol[4]) < lots && !string.IsNullOrEmpty(stopprofitCombineSymbol[4]))
                        lots = ushort.Parse(stopprofitCombineSymbol[4]);

                    StringBuilder sb = new StringBuilder();
                    long rtn = 0;
                    //emergencyBreakStopLoss判斷用碼錶
                    if (profitStopSuccessCount == 0) stopWatch.Restart();

                    //取得送單序號
                    long RequestId = tfcom.GetRequestId();
                    bool targetOrderFound = false;

                    ORDER_TYPE ordertype = ORDER_TYPE.OT_NEW;
                    SIDE_FLAG SideFlag = SIDE_FLAG.SF_BUY;//(tbTSide.Text[0] == 'B') ? SIDE_FLAG.SF_BUY : SIDE_FLAG.SF_SELL; //我們平常價差單停利目前都是要用Buy

                    sb.Append("Call停利 RequestId: " + RequestId).Append(Environment.NewLine);
                    sb.Append("SideFlag: " + SideFlag).Append(Environment.NewLine);
                    sb.Append("Symbol: " + stopprofitCombineSymbol[0]).Append(Environment.NewLine);
                    sb.Append("Lots: " + lots).Append(Environment.NewLine);
                    sb.Append("Price: " + (decimal.Parse(stopprofitCombineSymbol[6])).ToString()).Append(Environment.NewLine);

                    AddInfo(sb.ToString());
                    opLog.txtLog(sb.ToString());
                    this.Invoke((Action)(() =>
                    {
                        if (formClose)
                        {
                            FlagInvokeLoopOrder = false;
                            return;
                        }
                        FlagInvokeLoopOrder = true;

                        if (!String.IsNullOrEmpty(stopprofitCombineSymbol[0]) && decimal.Parse(stopprofitCombineSymbol[2]) > 0 && tradeConnect && quoteConnect)
                        {
                            if (cbOrderNotReview.Checked)
                            {
                                targetOrderFound = true;
                                rtn = tfcom.Order(ordertype, MARKET_FLAG.MF_OPT, RequestId, accInfo.BrokerID, accInfo.TbTAccount, "",
                                                            stopprofitCombineSymbol[0] /*tb20Code.Text*/ , SideFlag, PRICE_FLAG.PF_SPECIFIED, decimal.Parse(stopprofitCombineSymbol[6]), TIME_IN_FORCE.TIF_IOC, lots, POSITION_EFFECT.PE_CLOSE, OFFICE_FLAG.OF_SPEEDY);
                            }
                            else
                            {
                                if (MessageBox.Show(sb.ToString() + "是否下停利單?", "問題", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                                {
                                    targetOrderFound = true;
                                    rtn = tfcom.Order(ordertype, MARKET_FLAG.MF_OPT, RequestId, accInfo.BrokerID, accInfo.TbTAccount, "",
                                                                stopprofitCombineSymbol[0] /*tb20Code.Text*/ , SideFlag, PRICE_FLAG.PF_SPECIFIED, decimal.Parse(stopprofitCombineSymbol[6]), TIME_IN_FORCE.TIF_IOC, lots, POSITION_EFFECT.PE_CLOSE, OFFICE_FLAG.OF_SPEEDY);
                                }
                                else
                                {
                                    MessageBox.Show("Rcv No!");
                                }
                            }

                        }
                        else
                        {
                            opLog.txtLog("Call停利找無合適的合約，請檢查!!!!");
                            AddInfoTG("Call停利找無合適的合約，請檢查!!!!");
                        }

                        FlagInvokeLoopOrder = false;
                        if (formClose)
                        {
                            return;
                        }
                    }));
                    if (rtn != 0)
                    {
                        //MessageBox.Show(tfcom.GetOrderErrMsg(rtn));
                        AddInfo(tfcom.GetOrderErrMsg(rtn));
                        AddInfoTG(tfcom.GetOrderErrMsg(rtn));
                        opLog.txtLog(tfcom.GetOrderErrMsg(rtn));
                        profitStopFailCount++;
                        targetOrderFound = false;
                    }

                    int timeoutCounter = 0;
                    while (targetOrderFound == true && timeoutCounter < 10)
                    {

                        if (orderFirstRPT.ContainsKey(RequestId))
                        {
                            break;
                        }
                        Thread.Sleep(500);
                        timeoutCounter++;
                    }
                    if (targetOrderFound == true && orderFirstRPT.ContainsKey(RequestId) && string.Compare(orderFirstRPT[RequestId][0], "收單") == 0)
                    {
                        Console.WriteLine("Call停利 " + RequestId + ": 收單成功");
                        AddInfo("Call停利 " + RequestId + "收單成功");
                        opLog.txtLog("Call停利 " + RequestId + "收單成功");
                        //檢查 orderFirstRPT[RequestId][1] (CNT) 是否收到
                        if (orderFirstRPT[RequestId][1] != null)
                        {
                            string cnt = orderFirstRPT[RequestId][1];
                            AddInfo("Call停利 Rcv orderSecondRPT CNT: " + cnt);
                            opLog.txtLog("Call停利 Rcv orderSecondRPT CNT: " + cnt);

                            AddInfo("Call停利 第二次下單回覆: " + orderRPT[cnt][(int)ORDERRPT_LIST.ORDERSTATUS]);
                            opLog.txtLog("Call停利 第二次下單回覆: " + orderRPT[cnt][(int)ORDERRPT_LIST.ORDERSTATUS]);

                            AddInfo("Call停利 Rcv orderSecondRPT errorcode: " + orderRPT[cnt][(int)ORDERRPT_LIST.CODE]);
                            AddInfoTG("Call停利 Rcv orderSecondRPT errorcode: " + orderRPT[cnt][(int)ORDERRPT_LIST.CODE]);
                            opLog.txtLog("Call停利 Rcv orderSecondRPT errorcode: " + orderRPT[cnt][(int)ORDERRPT_LIST.CODE]);

                            if (dealRPT != null && dealRPT.ContainsKey(cnt))
                            {
                                opLog.txtLog("Call停利 dealRPT[cnt][ORDERSTATUS]: " + dealRPT[cnt][(int)DEALRPT_LIST.ORDERSTATUS]);
                                if (string.Compare(dealRPT[cnt][(int)DEALRPT_LIST.ORDERSTATUS], "1") == 0)
                                {
                                    profitStopSuccessCount++;
                                    AddInfo(cnt + " 收到成交回報 Call停利自動下單流程完成!");
                                    AddInfoTG(cnt + " 收到成交回報 Call停利自動下單流程完成!");
                                    opLog.txtLog(cnt + " 收到成交回報 Call停利自動下單流程完成!");
                                    if (ushort.Parse(dealRPT[cnt][(int)DEALRPT_LIST.LEAVEQTY]) > 0) //如果是部分成交
                                    {
                                        AddInfo("Call停利部分成交!");
                                        AddInfoTG("Call停利部分成交!");
                                        opLog.txtLog("Call停利部分成交!");
                                    }
                                    else
                                    {
                                        AddInfo("Call停利完全成交!");
                                        AddInfoTG("Call停利完全成交!");
                                        opLog.txtLog("Call停利完全成交!");
                                    }
                                    rtn = tfcom.RetrivePositionDetail("I", accInfo.BrokerID, accInfo.TbTAccount, "", "");
                                    if (rtn == -1)
                                    {
                                        MessageBox.Show("無查詢權限1618!!!");
                                        opLog.txtLog("無查詢權限1618!!!");
                                        AddInfoTG("無查詢權限1618!!!");
                                    }
                                    rtn = tfcom.RetrivePositionSum("I", accInfo.BrokerID, accInfo.TbTAccount, "", "");
                                    if (rtn == -1)
                                    {
                                        MessageBox.Show("無查詢權限1616!!!");
                                        opLog.txtLog("無查詢權限1616!!!");
                                        AddInfoTG("無查詢權限1616!!!");
                                    }
                                }
                                else
                                {
                                    profitStopFailCount++;
                                    AddInfo("可能收到舊的成交回報 Call停利自動下單流程失敗! 請檢查!!!!");
                                    AddInfoTG("可能收到舊的成交回報 Call停利自動下單流程失敗! 請檢查!!!!");
                                    opLog.txtLog("可能收到舊的成交回報 Call停利自動下單流程失敗! 請檢查!!!!");
                                }
                            }
                            else
                            {
                                profitStopFailCount++;
                                AddInfo("未收到成交回報 Call停利自動下單流程失敗! 重新送單!!!!");
                                AddInfoTG("未收到成交回報 Call停利自動下單流程失敗! 重新送單!!!!");
                                opLog.txtLog("未收到成交回報 Call停利自動下單流程失敗! 重新送單!!!!");
                            }
                        }
                    }
                    else if (targetOrderFound == false)
                    {
                        profitStopFailCount++;
                        AddInfo(RequestId + " Call停利未找到符合條件的合約，請檢查!!!!");
                        AddInfoTG(RequestId + " Call停利未找到符合條件的合約，請檢查!!!!");
                        opLog.txtLog(RequestId + " Call停利未找到符合條件的合約，請檢查!!!!");
                    }
                    else
                    {
                        profitStopFailCount++;
                        AddInfo(RequestId + " Call停利收單失敗，請檢查!!!!");
                        AddInfoTG(RequestId + " Call停利收單失敗，請檢查!!!!");
                        opLog.txtLog(RequestId + " Call停利收單失敗，請檢查!!!!");
                    }

                    //多少時間內 成交多少單或失敗多少單 則進保險
                    stopWatch.Stop();
                    double ts = stopWatch.Elapsed.TotalMinutes;
                    if (ts <= (double)orderSettings.ProfitStopCountMin && (profitStopSuccessCount >= orderSettings.ProfitStopSuccessRound || profitStopFailCount >= orderSettings.ProfitStopFailRound))
                    {
                        AddInfo("Call停利在" + orderSettings.ProfitStopCountMin + "分鐘內count次數超過，" + "profitStopSuccessCount=" + profitStopSuccessCount + " profitStopFailCount=" + profitStopFailCount + "  emergencyBreakStopLoss!!!!");
                        AddInfoTG("Call停利在" + orderSettings.ProfitStopCountMin + "分鐘內count次數超過，" + "profitStopSuccessCount=" + profitStopSuccessCount + " profitStopFailCount=" + profitStopFailCount + "  emergencyBreakStopLoss!!!!");
                        opLog.txtLog("Call停利在" + orderSettings.ProfitStopCountMin + "分鐘內count次數超過，" + "profitStopSuccessCount=" + profitStopSuccessCount + " profitStopFailCount=" + profitStopFailCount + "  emergencyBreakStopLoss!!!!");
                        emergencyBreakStopLoss = true;
                        MessageBox.Show("Call停利在" + orderSettings.ProfitStopCountMin + "分鐘內count次數超過，" + "profitStopSuccessCount=" + profitStopSuccessCount + " profitStopFailCount=" + profitStopFailCount + "  emergencyBreakStopLoss!!!!");
                    }
                    else if (ts > (double)orderSettings.ProfitStopCountMin)
                    {
                        opLog.txtLog("Call停利 " + orderSettings.ProfitStopCountMin + "分鐘已到 stopWatch&profitStopSuccessCount&profitStopFailCount 歸零!");
                        stopWatch.Reset();
                        profitStopSuccessCount = 0;
                        profitStopFailCount = 0;
                    }
                    else stopWatch.Start();
                    if(!formClose) Thread.Sleep(Convert.ToInt32(orderSettings.TakeProfitGap * 1000));
                }
                #endregion
                #region Put停利
                profitStopFailCount = 0;
                profitStopSuccessCount = 0;
                stopprofitCombineSymbol = new string[4];
                stopWatch.Reset();
                while (!formClose && !emergencyBreakStopLoss && loopOrderTempSwitch && !string.IsNullOrEmpty(findStopProfitInStock()[1]) )
                {
                    lots = 0;

                    lots = orderSettings.TakeProfitLots;
                    stopprofitCombineSymbol = new string[7];
                    stopprofitCombineSymbol = findStopProfitInStock(); //找出庫存裡面需要停利的合約
                    if (ushort.Parse(stopprofitCombineSymbol[5]) < lots && !string.IsNullOrEmpty(stopprofitCombineSymbol[5]))
                        lots = ushort.Parse(stopprofitCombineSymbol[5]);

                    StringBuilder sb = new StringBuilder();
                    long rtn = 0;
                    //emergencyBreakStopLoss判斷用碼錶
                    if (profitStopSuccessCount == 0) stopWatch.Restart();

                    //取得送單序號
                    long RequestId = tfcom.GetRequestId();
                    bool targetOrderFound = false;

                    ORDER_TYPE ordertype = ORDER_TYPE.OT_NEW;
                    SIDE_FLAG SideFlag = SIDE_FLAG.SF_BUY;//(tbTSide.Text[0] == 'B') ? SIDE_FLAG.SF_BUY : SIDE_FLAG.SF_SELL; //我們平常價差單停利目前都是要用Buy

                    sb.Append("Put停利 RequestId: " + RequestId).Append(Environment.NewLine);
                    sb.Append("SideFlag: " + SideFlag).Append(Environment.NewLine);
                    sb.Append("Symbol: " + stopprofitCombineSymbol[1]).Append(Environment.NewLine);
                    sb.Append("Lots: " + lots).Append(Environment.NewLine);
                    sb.Append("Price: " + (decimal.Parse(stopprofitCombineSymbol[6])).ToString()).Append(Environment.NewLine);

                    AddInfo(sb.ToString());
                    opLog.txtLog(sb.ToString());
                    this.Invoke((Action)(() =>
                    {
                        if (formClose)
                        {
                            FlagInvokeLoopOrder = false;
                            return;
                        }
                        FlagInvokeLoopOrder = true;

                        if (!String.IsNullOrEmpty(stopprofitCombineSymbol[1]) && decimal.Parse(stopprofitCombineSymbol[3]) > 0 && tradeConnect && quoteConnect)
                        {
                            if (cbOrderNotReview.Checked)
                            {
                                targetOrderFound = true;
                                rtn = tfcom.Order(ordertype, MARKET_FLAG.MF_OPT, RequestId, accInfo.BrokerID, accInfo.TbTAccount, "",
                                                            stopprofitCombineSymbol[1] /*tb20Code.Text*/ , SideFlag, PRICE_FLAG.PF_SPECIFIED, decimal.Parse(stopprofitCombineSymbol[6]), TIME_IN_FORCE.TIF_IOC, lots, POSITION_EFFECT.PE_CLOSE, OFFICE_FLAG.OF_SPEEDY);
                            }
                            else
                            {
                                if (MessageBox.Show(sb.ToString() + "是否下停利單?", "問題", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                                {
                                    targetOrderFound = true;
                                    rtn = tfcom.Order(ordertype, MARKET_FLAG.MF_OPT, RequestId, accInfo.BrokerID, accInfo.TbTAccount, "",
                                                                stopprofitCombineSymbol[1] /*tb20Code.Text*/ , SideFlag, PRICE_FLAG.PF_SPECIFIED, decimal.Parse(stopprofitCombineSymbol[6]), TIME_IN_FORCE.TIF_IOC, lots, POSITION_EFFECT.PE_CLOSE, OFFICE_FLAG.OF_SPEEDY);
                                }
                                else
                                {
                                    MessageBox.Show("Rcv No!");
                                }
                            }
                        }
                        else
                        {
                            opLog.txtLog("Put停利找無合適的合約，請檢查!!!!");
                            AddInfoTG("Put停利找無合適的合約，請檢查!!!!");
                        }

                        FlagInvokeLoopOrder = false;
                        if (formClose)
                        {
                            return;
                        }
                    }));
                    if (rtn != 0)
                    {
                        //MessageBox.Show(tfcom.GetOrderErrMsg(rtn));
                        AddInfo(tfcom.GetOrderErrMsg(rtn));
                        AddInfoTG(tfcom.GetOrderErrMsg(rtn));
                        opLog.txtLog(tfcom.GetOrderErrMsg(rtn));
                        profitStopFailCount++;
                        targetOrderFound = false;
                    }

                    int timeoutCounter = 0;
                    while (targetOrderFound == true && timeoutCounter < 10)
                    {

                        if (orderFirstRPT.ContainsKey(RequestId))
                        {
                            break;
                        }
                        Thread.Sleep(500);
                        timeoutCounter++;
                    }
                    if (targetOrderFound == true && orderFirstRPT.ContainsKey(RequestId) && string.Compare(orderFirstRPT[RequestId][0], "收單") == 0)
                    {
                        Console.WriteLine("Put停利 " + RequestId + ": 收單成功");
                        AddInfo("Put停利 " + RequestId + "收單成功");
                        opLog.txtLog("Put停利 " + RequestId + "收單成功");
                        //檢查 orderFirstRPT[RequestId][1] (CNT) 是否收到
                        if (orderFirstRPT[RequestId][1] != null)
                        {
                            string cnt = orderFirstRPT[RequestId][1];
                            AddInfo("Put停利 Rcv orderSecondRPT CNT: " + cnt);
                            opLog.txtLog("Put停利 Rcv orderSecondRPT CNT: " + cnt);

                            AddInfo("Put停利 第二次下單回覆: " + orderRPT[cnt][(int)ORDERRPT_LIST.ORDERSTATUS]);
                            opLog.txtLog("Put停利 第二次下單回覆: " + orderRPT[cnt][(int)ORDERRPT_LIST.ORDERSTATUS]);

                            AddInfo("Put停利 Rcv orderSecondRPT errorcode: " + orderRPT[cnt][(int)ORDERRPT_LIST.CODE]);
                            AddInfoTG("Put停利 Rcv orderSecondRPT errorcode: " + orderRPT[cnt][(int)ORDERRPT_LIST.CODE]);
                            opLog.txtLog("Put停利 Rcv orderSecondRPT errorcode: " + orderRPT[cnt][(int)ORDERRPT_LIST.CODE]);

                            if (dealRPT != null && dealRPT.ContainsKey(cnt))
                            {
                                opLog.txtLog("Put停利 dealRPT[cnt][ORDERSTATUS]: " + dealRPT[cnt][(int)DEALRPT_LIST.ORDERSTATUS]);
                                if (string.Compare(dealRPT[cnt][(int)DEALRPT_LIST.ORDERSTATUS], "1") == 0)
                                {
                                    profitStopSuccessCount++;
                                    AddInfo(cnt + " 收到成交回報 Put停利自動下單流程完成!");
                                    AddInfoTG(cnt + " 收到成交回報 Put停利自動下單流程完成!");
                                    opLog.txtLog(cnt + " 收到成交回報 Put停利自動下單流程完成!");
                                    if (ushort.Parse(dealRPT[cnt][(int)DEALRPT_LIST.LEAVEQTY]) > 0) //如果是部分成交
                                    {
                                        AddInfo("Put停利部分成交!");
                                        AddInfoTG("Put停利部分成交!");
                                        opLog.txtLog("Put停利部分成交!");
                                    }
                                    else
                                    {
                                        AddInfo("Put停利完全成交!");
                                        AddInfoTG("Put停利完全成交!");
                                        opLog.txtLog("Put停利完全成交!");
                                    }
                                    rtn = tfcom.RetrivePositionDetail("I", accInfo.BrokerID, accInfo.TbTAccount, "", "");
                                    if (rtn == -1)
                                    {
                                        MessageBox.Show("無查詢權限1618!!!");
                                        opLog.txtLog("無查詢權限1618!!!");
                                        AddInfoTG("無查詢權限1618!!!");
                                    }
                                    rtn = tfcom.RetrivePositionSum("I", accInfo.BrokerID, accInfo.TbTAccount, "", "");
                                    if (rtn == -1)
                                    {
                                        MessageBox.Show("無查詢權限1616!!!");
                                        opLog.txtLog("無查詢權限1616!!!");
                                        AddInfoTG("無查詢權限1616!!!");
                                    }
                                }
                                else
                                {
                                    profitStopFailCount++;
                                    AddInfo("可能收到舊的成交回報 Put停利自動下單流程失敗! 請檢查!!!!");
                                    AddInfoTG("可能收到舊的成交回報 Put停利自動下單流程失敗! 請檢查!!!!");
                                    opLog.txtLog("可能收到舊的成交回報 Put停利自動下單流程失敗! 請檢查!!!!");
                                }
                            }
                            else
                            {
                                profitStopFailCount++;
                                AddInfo("未收到成交回報 Put停利自動下單流程失敗! 重新送單!!!!");
                                AddInfoTG("未收到成交回報 Put停利自動下單流程失敗! 重新送單!!!!");
                                opLog.txtLog("未收到成交回報 Put停利自動下單流程失敗! 重新送單!!!!");
                            }
                        }
                    }
                    else if (targetOrderFound == false)
                    {
                        profitStopFailCount++;
                        AddInfo(RequestId + " Put停利未找到符合條件的合約，請檢查!!!!");
                        AddInfoTG(RequestId + " Put停利未找到符合條件的合約，請檢查!!!!");
                        opLog.txtLog(RequestId + " Put停利未找到符合條件的合約，請檢查!!!!");
                    }
                    else
                    {
                        profitStopFailCount++;
                        AddInfo(RequestId + " Put停利收單失敗，請檢查!!!!");
                        AddInfoTG(RequestId + " Put停利收單失敗，請檢查!!!!");
                        opLog.txtLog(RequestId + " Put停利收單失敗，請檢查!!!!");
                    }

                    //多少時間內 成交多少單或失敗多少單 則進保險
                    stopWatch.Stop();
                    double ts = stopWatch.Elapsed.TotalMinutes;
                    if (ts <= (double)orderSettings.ProfitStopCountMin && (profitStopSuccessCount >= orderSettings.ProfitStopSuccessRound || profitStopFailCount >= orderSettings.ProfitStopFailRound))
                    {
                        AddInfo("Put停利在" + orderSettings.ProfitStopCountMin + "分鐘內count次數超過，" + "profitStopSuccessCount=" + profitStopSuccessCount + " profitStopFailCount=" + profitStopFailCount + "  emergencyBreakStopLoss!!!!");
                        AddInfoTG("Put停利在" + orderSettings.ProfitStopCountMin + "分鐘內count次數超過，" + "profitStopSuccessCount=" + profitStopSuccessCount + " profitStopFailCount=" + profitStopFailCount + "  emergencyBreakStopLoss!!!!");
                        opLog.txtLog("Put停利在" + orderSettings.ProfitStopCountMin + "分鐘內count次數超過，" + "profitStopSuccessCount=" + profitStopSuccessCount + " profitStopFailCount=" + profitStopFailCount + "  emergencyBreakStopLoss!!!!");
                        emergencyBreakStopLoss = true;
                        MessageBox.Show("Put停利在" + orderSettings.ProfitStopCountMin + "分鐘內count次數超過，" + "profitStopSuccessCount=" + profitStopSuccessCount + " profitStopFailCount=" + profitStopFailCount + "  emergencyBreakStopLoss!!!!");
                    }
                    else if (ts > (double)orderSettings.ProfitStopCountMin)
                    {
                        opLog.txtLog("Put停利 " + orderSettings.ProfitStopCountMin + "分鐘已到 stopWatch&profitStopSuccessCount&profitStopFailCount 歸零!");
                        stopWatch.Reset();
                        profitStopSuccessCount = 0;
                        profitStopFailCount = 0;
                    }
                    else stopWatch.Start();
                    if(!formClose) Thread.Sleep(Convert.ToInt32(orderSettings.TakeProfitGap * 1000));
                }
                #endregion

            }
            formCloseFinish = true;
        }

        //找出100差在目標範圍內的Call Put之Symbol
        private string[] findHundPointInHundred()//[0]: combineSymbolCall  [1]: combineSymbolPut [2]: HundredDiff of thisSymbolCall [3]: HundredDiff of thisSymbolPut
        {
            decimal targetCallUpLim = orderSettings.CallUpLim, targetCallDownLim = orderSettings.CallDownLim;
            //opLog.txtLog("orderSettings.CallUpLim: "+ orderSettings.CallUpLim);
            //opLog.txtLog("orderSettings.CallDownLim: " + orderSettings.CallDownLim);
            Dictionary<string, List<string>> _symbolOPriceMap = new Dictionary<string, List<string>>();

            _symbolOPriceMap = symbolOPriceMap;
            var resultCallFifty = _symbolOPriceMap
                .Select(x => x.Value).Cast<List<string>>()
                .Where(y => y[(int)OPTION_LIST.MONTH][0] >= 'A' && y[(int)OPTION_LIST.MONTH][0] <= 'L')
                .Where(y => decimal.Parse(y[(int)OPTION_LIST.SYMBOL].Substring(3,5)) >= orderSettings.UpSafeDis)
                .Where(y => decimal.Parse(y[(int)OPTION_LIST.HUNDDIF]) <= targetCallUpLim)
                .Where(y => decimal.Parse(y[(int)OPTION_LIST.HUNDDIF]) >= targetCallDownLim - orderSettings.OrderPriceShift)
                .Select(y => new { symbol = y[(int)OPTION_LIST.SYMBOL], diff = Math.Abs(targetCallDownLim - decimal.Parse(y[(int)OPTION_LIST.HUNDDIF])), hunddiff= y[(int)OPTION_LIST.HUNDDIF] })
                .OrderBy(y => y.diff).FirstOrDefault();

            //if (resultCallFifty != null)
            //{
            //    opLog.txtLog("result.symbol: " + resultCallFifty.symbol);
            //}

            decimal targetPutUpLim = orderSettings.PutUpLim, targetPutDownLim = orderSettings.PutDownLim;
            //opLog.txtLog("orderSettings.PutUpLim: " + orderSettings.PutUpLim);
            //opLog.txtLog("orderSettings.PutDownLim: " + orderSettings.PutDownLim);
            var resultPutFifty = _symbolOPriceMap
                .Select(x => x.Value).Cast<List<string>>()
                .Where(y => y[(int)OPTION_LIST.MONTH][0] >= 'M' && y[(int)OPTION_LIST.MONTH][0] <= 'X')
                .Where(y => decimal.Parse(y[(int)OPTION_LIST.SYMBOL].Substring(3, 5)) <= orderSettings.DownSafeDis)
                .Where(y => decimal.Parse(y[(int)OPTION_LIST.HUNDDIF]) <= targetPutUpLim)
                .Where(y => decimal.Parse(y[(int)OPTION_LIST.HUNDDIF]) >= targetPutDownLim - orderSettings.OrderPriceShift)
                .Select(y => new { symbol = y[(int)OPTION_LIST.SYMBOL], diff = Math.Abs(targetPutDownLim - decimal.Parse(y[(int)OPTION_LIST.HUNDDIF])), hunddiff = y[(int)OPTION_LIST.HUNDDIF] })
                .OrderBy(y => y.diff).FirstOrDefault();


            //if (resultPutFifty != null)
            //{
            //    opLog.txtLog("result.symbol: " + resultPutFifty.symbol);
            //}

            string combineSymbolCall = "", combineSymbolPut = "", hunddiffCall = "0", hunddiffPut = "0";
            

            if (resultCallFifty != null)
            {
                if(resultCallFifty.symbol != null)
                {
                    combineSymbolCall = resultCallFifty.symbol;
                    combineSymbolCall = combineSymbolCall.Insert(3, (decimal.Parse(combineSymbolCall.Substring(3, 5)) + 100).ToString() + @"/");
                    //tb20Code.Text = combineSymbolCall;
                    //tbTStrikePrice2.Text = resultCallFifty.symbol.Substring(3, 5);
                    //tbTSide2.Text = "S";
                    //tbTComym2.Text = CalculateComID.PartialComIDC[2];
                    //tbTStrikePrice.Text = (decimal.Parse(resultCallFifty.symbol.Substring(3, 5)) + 100).ToString();
                    //tbTSide.Text = "B";
                    //tbTComym.Text = CalculateComID.PartialComIDC[2];
                    //rbMulti.Checked = true;
                }
                if (resultCallFifty.hunddiff != null) hunddiffCall = resultCallFifty.hunddiff;
            }
            

            if (resultPutFifty != null)
            {
                if (resultPutFifty.symbol != null)
                {
                    combineSymbolPut = resultPutFifty.symbol;
                    combineSymbolPut = combineSymbolPut.Insert(3, (decimal.Parse(combineSymbolPut.Substring(3, 5)) - 100).ToString() + @"/");
                }
                if (resultPutFifty.hunddiff != null) hunddiffPut = resultPutFifty.hunddiff;
            }

            string[] combineSymbol = new string[] { combineSymbolCall, combineSymbolPut, hunddiffCall, hunddiffPut };
            return combineSymbol;
        }

        //找出離價平最近的庫存
        private string[] findNearestNowPriceInStock()
        {
            CalculateComID calcuComID;
            if (string.Compare(DateTime.Now.DayOfWeek.ToString("d"), "3") == 0 && DateCalculateExtensions.IsLegalTime(DateTime.Now, "050001-143000"))
                calcuComID = new CalculateComID(DateTime.Now.AddDays(-1));
            else calcuComID = new CalculateComID(DateTime.Now);

            string weekMX = calcuComID.PartialComIDMX[0] + calcuComID.PartialComIDMX[1];
            decimal targetFPrice = 0;
            Dictionary<string, List<string>> _stockRPT = new Dictionary<string, List<string>>();

            if (symbolFPriceMap.ContainsKey(weekMX))
                targetFPrice = decimal.Parse(symbolFPriceMap[weekMX][(int)FUTURE_LIST.PRICE]);

            opLog.txtLog("targetFPrice: " + targetFPrice);
            _stockRPT = stockRPT;
            var resultNearestCall = _stockRPT
                .Select(x => x.Value).Cast<List<string>>()
                .Where(y => y[19][0] == '1' && y[20][0] == 'C' && y[34][0] == '1' && y[18][0] == 'B' && y[40][0] == 'S')  //18: BS1  19: F:期 O:權  20:C/P 34:複式組合單號 "1" 40: BS2  
                .Select(y => new { symbol = y[17], strikePrice1 = y[21], strikePrice2 = y[43], comYM2 = y[44], diff = Math.Abs(targetFPrice - decimal.Parse(y[43])), marketPrice = decimal.Parse(y[47]) - decimal.Parse(y[25]), dealPrice = decimal.Parse(y[52]) - decimal.Parse(y[30]) })  //21:履約價1 43:履約價2  17:商品代碼 "TX2"  44: 商品年月2  202201  47:MarketPrice2  25:MarketPrice1 52:DealPrice2 30:DealPrice
                .OrderBy(y => y.diff).FirstOrDefault();

            if (resultNearestCall != null)
            {
                opLog.txtLog("resultCall.symbol: " + resultNearestCall.symbol);
                opLog.txtLog("resultCall.strikePrice1: " + resultNearestCall.strikePrice1);
                opLog.txtLog("resultCall.strikePrice2: " + resultNearestCall.strikePrice2);
                opLog.txtLog("resultCall.comYM2: " + resultNearestCall.comYM2);
                opLog.txtLog("resultCall.diff: " + resultNearestCall.diff);
            }

            var resultNearestPut = _stockRPT
                .Select(x => x.Value).Cast<List<string>>()
                .Where(y => y[19][0] == '1' && y[20][0] == 'P' && y[34][0] == '1' && y[18][0] == 'B' && y[40][0] == 'S')  //18: BS1  19: F:期 O:權  20:C/P 34:複式組合單號 "1" 40: BS2  
                .Select(y => new { symbol = y[17], strikePrice1 = y[21], strikePrice2 = y[43], comYM2 = y[44], diff = Math.Abs(targetFPrice - decimal.Parse(y[43])), marketPrice = decimal.Parse(y[47]) - decimal.Parse(y[25]), dealPrice = decimal.Parse(y[52]) - decimal.Parse(y[30]) })  //21:履約價1 43:履約價2  17:商品代碼 "TX2"  44: 商品年月2  202201   47:MarketPrice2  25:MarketPrice1 52:DealPrice2 30:DealPrice
                .OrderBy(y => y.diff).FirstOrDefault();


            if (resultNearestPut != null)
            {
                opLog.txtLog("resultNearestPut.symbol: " + resultNearestPut.symbol);
            }


            string combineSymbolCall = "", combineSymbolPut = "", dealPriceCall = "0", dealPricePut = "0";
            char[] monthSymbol = new char[2];
            if (resultNearestCall != null)
                CalculateComID.FindSymbol(int.Parse(resultNearestCall.comYM2.Substring(4, 2)), monthSymbol);

            if (resultNearestCall != null)
            {
                if (!string.IsNullOrEmpty(resultNearestCall.symbol))
                {
                    combineSymbolCall = resultNearestCall.symbol.ToString();
                    combineSymbolCall = combineSymbolCall + ((int)double.Parse(resultNearestCall.strikePrice1)).ToString() + @"/";
                    combineSymbolCall = combineSymbolCall + ((int)double.Parse(resultNearestCall.strikePrice2)).ToString();
                    combineSymbolCall = combineSymbolCall + monthSymbol[0].ToString() + resultNearestCall.comYM2[3].ToString();
                }
                if (resultNearestCall.dealPrice > 0) dealPriceCall = resultNearestCall.dealPrice.ToString();
            }
            else combineSymbolCall = "";

            if (resultNearestPut != null)
            {
                if (resultNearestPut.symbol != null)
                {
                    combineSymbolPut = resultNearestPut.symbol.ToString();
                    combineSymbolPut = combineSymbolPut + ((int)double.Parse(resultNearestPut.strikePrice1)).ToString() + @"/";
                    combineSymbolPut = combineSymbolPut + ((int)double.Parse(resultNearestPut.strikePrice2)).ToString();
                    combineSymbolPut = combineSymbolPut + monthSymbol[0].ToString() + resultNearestPut.comYM2[3].ToString();
                }
                if (resultNearestPut.dealPrice > 0) dealPricePut = resultNearestPut.dealPrice.ToString();
            }
            else combineSymbolPut = "";

            string[] combineSymbol = new string[] { combineSymbolCall, combineSymbolPut, dealPriceCall, dealPricePut };
            opLog.txtLog("NearestCall: " + combineSymbol[0]);
            opLog.txtLog("NearestPut: " + combineSymbol[1]);
            opLog.txtLog("dealPriceCall: " + combineSymbol[2]);
            opLog.txtLog("dealPricePut: " + combineSymbol[3]);

            return combineSymbol;
        }

        //找出庫存中需要停利的商品
        private string[] findStopProfitInStock()
        {
            DateTime NowDate = DateTime.Now;
            int periodOfWeek = 0;
            Dictionary<string, List<string>> _stockRPT = new Dictionary<string, List<string>>();

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

            _stockRPT = stockRPT;
            var resultStopProfitCall = _stockRPT
                .Select(x => x.Value).Cast<List<string>>()
                .Where(y => y[19][0] == '1' && y[20][0] == 'C' && y[34][0] == '1' && y[18][0] == 'B' && y[40][0] == 'S' && (decimal.Parse(y[47]) - decimal.Parse(y[25])) <= orderSettings.DailyPoints[periodOfWeek][2] )  //18: BS1  19: F:期 O:權  20:C/P 34:複式組合單號 "1" 40: BS2  
                .Select(y => new { symbol = y[17], strikePrice1 = y[21], strikePrice2 = y[43], comYM2 = y[44], marketPrice = decimal.Parse(y[47]) - decimal.Parse(y[25]), dealPrice = decimal.Parse(y[52]) - decimal.Parse(y[30]), qty=ushort.Parse(y[23]) })  //21:履約價1 43:履約價2  17:商品代碼 "TX2"  44: 商品年月2  202201  47:MarketPrice2  25:MarketPrice1 52:DealPrice2 30:DealPrice
                .OrderBy(y => y.marketPrice).FirstOrDefault();

            if (resultStopProfitCall != null)
            {
                opLog.txtLog("resultStopProfitCall.symbol: " + resultStopProfitCall.symbol);
                opLog.txtLog("resultStopProfitCall.strikePrice1: " + resultStopProfitCall.strikePrice1);
                opLog.txtLog("resultStopProfitCall.strikePrice2: " + resultStopProfitCall.strikePrice2);
                opLog.txtLog("resultStopProfitCall.comYM2: " + resultStopProfitCall.comYM2);
                opLog.txtLog("resultStopProfitCall.marketPrice: " + resultStopProfitCall.marketPrice);
                opLog.txtLog("resultStopProfitCall.qty: " + resultStopProfitCall.qty);
            }

            var resultStopProfitPut = _stockRPT
                .Select(x => x.Value).Cast<List<string>>()
                .Where(y => y[19][0] == '1' && y[20][0] == 'P' && y[34][0] == '1' && y[18][0] == 'B' && y[40][0] == 'S' && (decimal.Parse(y[47]) - decimal.Parse(y[25])) <= orderSettings.DailyPoints[periodOfWeek][2] )  //18: BS1  19: F:期 O:權  20:C/P 34:複式組合單號 "1" 40: BS2  
                .Select(y => new { symbol = y[17], strikePrice1 = y[21], strikePrice2 = y[43], comYM2 = y[44], marketPrice = decimal.Parse(y[47]) - decimal.Parse(y[25]), dealPrice = decimal.Parse(y[52]) - decimal.Parse(y[30]), qty = ushort.Parse(y[23]) })  //21:履約價1 43:履約價2  17:商品代碼 "TX2"  44: 商品年月2  202201   47:MarketPrice2  25:MarketPrice1 52:DealPrice2 30:DealPrice
                .OrderBy(y => y.marketPrice).FirstOrDefault();

            if (resultStopProfitPut != null)
            {
                opLog.txtLog("resultStopProfitPut.symbol: " + resultStopProfitPut.symbol);
                opLog.txtLog("resultStopProfitPut.strikePrice1: " + resultStopProfitPut.strikePrice1);
                opLog.txtLog("resultStopProfitPut.strikePrice2: " + resultStopProfitPut.strikePrice2);
                opLog.txtLog("resultStopProfitPut.comYM2: " + resultStopProfitPut.comYM2);
                opLog.txtLog("resultStopProfitPut.marketPrice: " + resultStopProfitPut.marketPrice);
                opLog.txtLog("resultStopProfitPut.qty: " + resultStopProfitPut.qty);
            }

            string combineSymbolCall = "", combineSymbolPut = "", marketPriceCall = "0", marketPricePut = "0", qtyCall = "", qtyPut = "";
            char[] monthSymbol = new char[2];
            if (resultStopProfitCall != null)
                CalculateComID.FindSymbol(int.Parse(resultStopProfitCall.comYM2.Substring(4, 2)), monthSymbol);

            if (resultStopProfitCall != null)
            {
                if (!string.IsNullOrEmpty(resultStopProfitCall.symbol))
                {
                    combineSymbolCall = resultStopProfitCall.symbol.ToString();
                    combineSymbolCall = combineSymbolCall + ((int)double.Parse(resultStopProfitCall.strikePrice1)).ToString() + @"/";
                    combineSymbolCall = combineSymbolCall + ((int)double.Parse(resultStopProfitCall.strikePrice2)).ToString();
                    combineSymbolCall = combineSymbolCall + monthSymbol[0].ToString() + resultStopProfitCall.comYM2[3].ToString();
                }
                if (resultStopProfitCall.marketPrice > 0) marketPriceCall = resultStopProfitCall.marketPrice.ToString();
                qtyCall = resultStopProfitCall.qty.ToString();
            }
            else combineSymbolCall = "";

            if (resultStopProfitPut != null)
            {
                if (resultStopProfitPut.symbol != null)
                {
                    combineSymbolPut = resultStopProfitPut.symbol.ToString();
                    combineSymbolPut = combineSymbolPut + ((int)double.Parse(resultStopProfitPut.strikePrice1)).ToString() + @"/";
                    combineSymbolPut = combineSymbolPut + ((int)double.Parse(resultStopProfitPut.strikePrice2)).ToString();
                    combineSymbolPut = combineSymbolPut + monthSymbol[0].ToString() + resultStopProfitPut.comYM2[3].ToString();
                }
                if (resultStopProfitPut.marketPrice > 0) marketPricePut = resultStopProfitPut.marketPrice.ToString();
                qtyPut = resultStopProfitPut.qty.ToString();
            }
            else combineSymbolPut = "";

            string[] combineSymbol = new string[] { combineSymbolCall, combineSymbolPut, marketPriceCall, marketPricePut, qtyCall, qtyPut , orderSettings.DailyPoints[periodOfWeek][2].ToString() };
            //opLog.txtLog("StopProfitCall: " + combineSymbol[0]);
            //opLog.txtLog("StopProfitPut: " + combineSymbol[1]);
            //opLog.txtLog("marketPriceCall: " + combineSymbol[2]);
            //opLog.txtLog("marketPricePut: " + combineSymbol[3]);
            
            return combineSymbol;
        }

        //斷線重連Func
        private async void TryContinueConnect()
        {
            int count = 0;
            if (quoteConnect && tradeConnect) return;
            if (formClose) return;

            CalculateComID calcuComID;
            if (string.Compare(DateTime.Now.DayOfWeek.ToString("d"), "3") == 0 && DateCalculateExtensions.IsLegalTime(DateTime.Now, "050001-143000"))
                calcuComID = new CalculateComID(DateTime.Now.AddDays(-1));
            else calcuComID = new CalculateComID(DateTime.Now);

            while ((!quoteConnect || !tradeConnect)&& !formClose)
            {
                try
                {
                    count++;
                    opLog.txtLog("Disconnect try relogin!");
                    AddInfoTG("Disconnect try relogin!");
                    if (!quoteConnect && !tradeConnect)
                    {
                        //QuoteConnectInit();
                        quoteCom.Connect(quoteCom.Host, quoteCom.Port);
                        await Task.Delay(5000);
                        quoteCom.Login(accInfo.StrID, accInfo.StrPwd, ' ');
                        quoteCom.LoadTaifexProductXML();
                        short istatus = quoteCom.RetriveLastPrice(calcuComID.PartialComIDMX[0] + calcuComID.PartialComIDMX[1]);
                        if (istatus < 0)   //
                            opLog.txtLog(quoteCom.GetSubQuoteMsg(istatus));
                        await Task.Delay(3000);
                        tfcom.LoginDirect(tfcom.ServerHost, tfcom.ServerPort, accInfo.StrID + ",," + accInfo.StrPwd);
                        await Task.Delay(1000);
                    }
                    else if(!quoteConnect && tradeConnect)
                    {
                        quoteCom.Connect(quoteCom.Host, quoteCom.Port);
                        await Task.Delay(5000);
                        quoteCom.Login(accInfo.StrID, accInfo.StrPwd, ' ');
                        quoteCom.LoadTaifexProductXML();
                        short istatus = quoteCom.RetriveLastPrice(calcuComID.PartialComIDMX[0] + calcuComID.PartialComIDMX[1]);
                        if (istatus < 0)   //
                            opLog.txtLog(quoteCom.GetSubQuoteMsg(istatus));
                        await Task.Delay(int.Parse(settings.ReconnectGap) * 1000);
                    }
                    else if(quoteConnect && !tradeConnect)
                    {
                        tfcom.LoginDirect(tfcom.ServerHost, tfcom.ServerPort, accInfo.StrID + ",," + accInfo.StrPwd);
                        await Task.Delay(1000);
                    }
                    else if(quoteConnect && tradeConnect)
                    {
                        opLog.txtLog("Already Login!");
                        AddInfoTG("Already Login!");
                    }
                    else
                    {
                        opLog.txtLog("Fail! do relogin again!");
                        AddInfoTG("Fail! do relogin again!");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("reconnect exception:" + ex.Message);
                    opLog.txtLog("reconnect exception:" + ex.Message);
                    AddInfoTG("reconnect exception:" + ex.Message);
                }

                // 如果還沒連線不符合結束條件則睡設定秒數
                if (!quoteConnect || !tradeConnect)
                {
                    Console.WriteLine($"retry {count} times! quoteConnect={quoteConnect} tradeConnect={tradeConnect}");
                    opLog.txtLog($"retry {count} times! quoteConnect={quoteConnect} tradeConnect={tradeConnect}");
                    AddInfoTG($"retry {count} times! quoteConnect={quoteConnect} tradeConnect={tradeConnect}");
                    await Task.Delay(int.Parse(settings.ReconnectGap)*1000);
                }
                else
                {
                    Console.WriteLine($"retry {count} times! Reconnect Success! quoteConnect={quoteConnect} tradeConnect={tradeConnect}");
                    opLog.txtLog($"retry {count} times! Reconnect Success! quoteConnect={quoteConnect} tradeConnect={tradeConnect}");
                    AddInfoTG($"retry {count} times! Reconnect Success! quoteConnect={quoteConnect} tradeConnect={tradeConnect}");
                    TComid();
                    await Task.Delay(5000);
                }
            }
        }

        private void addTestRecord(string tempSymbol, string tempPrice, string tempVol)
        {
            List<string> tempList;
            string tempS = "";
            string[] tempSplit;

            tempS += tempSymbol;
            tempS += ",X,";
            tempS += tempPrice + ",0.1," + tempVol + ",NoN,NoN," + tempPrice;

            tempSplit = tempS.Split(',');

            tempList = new List<string>();
            foreach (string i in tempSplit)
            {
                tempList.Add(i);
                Console.WriteLine("addTestRecord:" + i);
            }

            if (symbolOPriceMap.ContainsKey(tempSymbol))
                symbolOPriceMap[tempSymbol] = tempList;
            else
                symbolOPriceMap.Add(tempSymbol, tempList);
        }

        private void btnOrder_Click(object sender, EventArgs e)
        {
            //請依您的需求自行增加下單欄位的檢核(ex 欄位長度/欄位值合法性 等等) ;  
            long rtn = 0;
            #region 
            MARKET_FLAG MF = (cbMarket.SelectedIndex == 1) ? MARKET_FLAG.MF_FUT : MARKET_FLAG.MF_OPT;

            SIDE_FLAG SideFlag = (tbTSide.Text[0] == 'B') ? SIDE_FLAG.SF_BUY : SIDE_FLAG.SF_SELL;

            if (rbMulti.Checked)  //2019.1.2 Add: 國內複式單以 第二支腳為主 
                SideFlag = (tbTSide2.Text[0] == 'B') ? SIDE_FLAG.SF_BUY : SIDE_FLAG.SF_SELL;

            //PRICE_FLAG priceFlag = (cbMarketPrice.Checked) ? PRICE_FLAG.PF_MARKET : PRICE_FLAG.PF_SPECIFIED;
            PRICE_FLAG priceFlag = PRICE_FLAG.PF_SPECIFIED;

            //測試期間不使用市價
            //if (rbMP.Checked) priceFlag = PRICE_FLAG.PF_MARKET;
            //else if (rbPP.Checked) priceFlag = PRICE_FLAG.PF_MARKET_RANGE; //2014.4.2 ADD 一定範圍市價單 

            decimal Price;
            if (priceFlag == PRICE_FLAG.PF_SPECIFIED)
                Price = decimal.Parse(tbTPrice.Text);
            else Price = 0;

            //預設IOC
            TIME_IN_FORCE TIF = TIME_IN_FORCE.TIF_IOC;
            if (cbROD.Text == "FOK") TIF = TIME_IN_FORCE.TIF_FOK;
            switch (cbROD.Text)
            {
                case "ROD": TIF = TIME_IN_FORCE.TIF_ROD; break;
                case "FOK": TIF = TIME_IN_FORCE.TIF_FOK; break;
                case "IOC": TIF = TIME_IN_FORCE.TIF_IOC; break;
                default:
                    MessageBox.Show("ROD/FOK/IOC 項目錯誤");
                    return;
            }
            ushort Qty = 0;
            try
            {
                Qty = ushort.Parse(tbTQty.Text);
            }
            catch
            {
                MessageBox.Show("口數輸入錯誤");
                return;
            }
            OFFICE_FLAG officeFlag = OFFICE_FLAG.OF_SPEEDY;
            POSITION_EFFECT pe;
            switch (cbOpen.SelectedIndex)
            {
                case 0: pe = POSITION_EFFECT.PE_OPEN; break;
                case 1: pe = POSITION_EFFECT.PE_CLOSE; break;
                case 2: pe = POSITION_EFFECT.PE_DAY_TRADE; break;
                case 3: pe = POSITION_EFFECT.PE_AUTO; break;
                default:
                    MessageBox.Show("倉別錯誤");
                    return;
            }
            ORDER_TYPE ordertype = ORDER_TYPE.OT_NEW;
            switch (cbFunc.SelectedIndex)
            {
                case 0: ordertype = ORDER_TYPE.OT_NEW; break;
                case 1: ordertype = ORDER_TYPE.OT_CANCEL; break;
                case 2: ordertype = ORDER_TYPE.OT_MODIFY_QTY; break;
                case 3: ordertype = ORDER_TYPE.OT_MODIFY_PRICE; break;
            }

            tb20Code.Text = GenFullSymbol();
            #endregion
            //****注意: RequestID必須經由GetRequestID取得, 且不可重覆********
            long RequestId;

            //2020.11.19 Lynn Test 自key 商品全碼
            string symbol = tb20Code.Text;
            RequestId = tfcom.GetRequestId();      //取得送單序號
            if (ordertype == ORDER_TYPE.OT_NEW)    //新單
                rtn = tfcom.Order(ORDER_TYPE.OT_NEW, MF, RequestId, accInfo.BrokerID, accInfo.TbTAccount, "",
                                            symbol /*tb20Code.Text*/ , SideFlag, priceFlag, Price, TIF, Qty, pe, officeFlag);
            else
            {
                //刪/改單
                //rtn = tfcom.Order(ordertype, MF, RequestId, accInfo.BrokerID, accInfo.TbTAccount, "",
                //                symbol /*tb20Code.Text*/, SideFlag, priceFlag, Price, TIF, Qty, pe, officeFlag, tbMWebID.Text, tbMCnt.Text, tbMOrderNo.Text);

            }
            if (rtn != 0)
                MessageBox.Show(tfcom.GetOrderErrMsg(rtn));
        }
        private string GenFullSymbol()
        {
            string fullsymbol = "";
            if (cbMarket.SelectedIndex == 1)
            { //期貨
                if (rbSingle.Checked)  //單式單
                    fullsymbol = tfcom.GenFutSymbol(tbTSymbol.Text, tbTComym.Text, "");
                else   //複式單
                    fullsymbol = tfcom.GenFutSymbol(tbTSymbol.Text, tbTComym.Text, tbTComym2.Text);
            }
            else
            {   //選擇權
                if (rbSingle.Checked)  //單式單
                    fullsymbol = tfcom.GenOptSymbol(tbTSymbol.Text, tbTComym.Text, tbTStrikePrice.Text, tbTCallPut.Text);
                else fullsymbol = tfcom.GenOptDoubleSymbol(tbTSymbol.Text, tbTComym.Text, tbTStrikePrice.Text, tbTCallPut.Text, tbTSide.Text, tbTComym2.Text, tbTStrikePrice2.Text, tbTCallPut2.Text, tbTSide2.Text);
            }
            return fullsymbol;

        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            //MessageBox.Show("FormClosed");
            //tl.CloseHttpClient();
            //tfcom.Logout();
            //tfcom.Dispose();
            //quoteCom.Logout();
            //quoteCom.Dispose();
            //thrListV.Abort();
            //thrOrder.Abort();
            //thrListV.Interrupt();
            //thrOrder.Interrupt();
            //tmrUnsub.Stop();
            //tmrUnsub.Close();
            //tmrPosition.Stop();
            //tmrPosition.Close();
            //Application.Exit();

            //Environment.Exit(Environment.ExitCode);
            Application.Exit();
            
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            //MessageBox.Show("Closing event\n");
            DialogResult dr = MessageBox.Show("確定要關閉程式嗎?",
                "Closing event!", MessageBoxButtons.YesNo);
            if (dr == DialogResult.No)
                e.Cancel = true;//取消離開
            else
                e.Cancel = false;
            formClose = true;
            tl.CloseHttpClient();
            tmrUnsub.Stop();
            tmrUnsub.Close();
            tmrPosition.Stop();
            tmrPosition.Close();
            //tfcom.Logout();
            //tfcom.Dispose();
            //quoteCom.Logout();
            //quoteCom.Dispose();
            while ( FlagInvoke || FlagInvokeLoopPrint || FlagInvokeLoopOrder)
            {
                Console.WriteLine("Form1_FormClosing!");
                Application.DoEvents();
                //if (FlagInvoke == false) this.Close();
                //Environment.Exit(Environment.ExitCode);
            }
            //try
            //{
            //    string[] zero = new string[1];
            //    Console.WriteLine(zero[2]);
            //}
            //catch { Console.WriteLine("Program Closed!"); }
            //MessageBox.Show("Closing event End\n");
        }

        private void Form2_FormClosed(object sender, FormClosedEventArgs e)
        {
            Form2 sub = (Form2)sender;
            //parm1 = sub.parm2;
            //this.textBox1.Text = parm1.FuncNum;
        }

        private void btnGenSymbol_Click(object sender, EventArgs e)
        {
            //請依您的需求增加各欄位的檢核(ex 欄位長度,合理性等等) ;  在此不做欄位檢核
            tb20Code.Text = GenFullSymbol();
        }

        private void btAutoContract_Click(object sender, EventArgs e)
        {
            CalculateComID calcuComID;
            if (string.Compare(DateTime.Now.DayOfWeek.ToString("d"), "3") == 0 && DateCalculateExtensions.IsLegalTime(DateTime.Now, "050001-143000"))
                calcuComID = new CalculateComID(DateTime.Now.AddDays(-1));
            else calcuComID = new CalculateComID(DateTime.Now);

            if (String.Compare(cbMarket.Text, "O.選擇權") == 0)
            {
                if (String.Compare(tbTCallPut.Text, "CALL") == 0)
                {
                    tbTSymbol.Text = calcuComID.PartialComIDC[0];
                    tbTComym.Text = calcuComID.PartialComIDC[2];
                    tbTComym2.Text = calcuComID.PartialComIDC[2];
                }
                else if (String.Compare(tbTCallPut.Text, "PUT") == 0)
                {
                    tbTSymbol.Text = calcuComID.PartialComIDP[0];
                    tbTComym.Text = calcuComID.PartialComIDP[2];
                    tbTComym2.Text = calcuComID.PartialComIDP[2];
                }
            }
            else
            {
                tbTSymbol.Text = calcuComID.PartialComIDMX[0];
                tbTComym.Text = calcuComID.PartialComIDMX[2];
            }
        }

        private void btSave_Click(object sender, EventArgs e)
        {
            try
            {
                btSave.Enabled = false;

                string dir = Application.StartupPath;
                string fileName = dir + @"\" + orderSettings.StrSettingIni;
                string temp = "";

                temp += "WEDNG=" + tbWEDNGcall.Text + "," + tbWEDNGput.Text + "," + tbWEDNGtkprofit.Text;
                temp += "\n";

                temp += "THUDT=" + tbTHUDTcall.Text + "," + tbTHUDTput.Text + "," + tbTHUDTtkprofit.Text;
                temp += "\n";

                temp += "THUNG=" + tbTHUNGcall.Text + "," + tbTHUNGput.Text + "," + tbTHUNGtkprofit.Text;
                temp += "\n";

                temp += "FRIDT=" + tbFRIDTcall.Text + "," + tbFRIDTput.Text + "," + tbFRIDTtkprofit.Text;
                temp += "\n";

                temp += "FRING=" + tbFRINGcall.Text + "," + tbFRINGput.Text + "," + tbFRINGtkprofit.Text;
                temp += "\n";

                temp += "MONDT=" + tbMONDTcall.Text + "," + tbMONDTput.Text + "," + tbMONDTtkprofit.Text;
                temp += "\n";

                temp += "MONNG=" + tbMONNGcall.Text + "," + tbMONNGput.Text + "," + tbMONNGtkprofit.Text;
                temp += "\n";

                temp += "TUEDT=" + tbTUEDTcall.Text + "," + tbTUEDTput.Text + "," + tbTUEDTtkprofit.Text;
                temp += "\n";

                temp += "TUENG=" + tbTUENGcall.Text + "," + tbTUENGput.Text + "," + tbTUENGtkprofit.Text;
                temp += "\n";

                temp += "WEDDT=" + tbWEDDTcall.Text + "," + tbWEDDTput.Text + "," + tbWEDDTtkprofit.Text;
                temp += "\n";

                temp += "CallDownLim=" + tbCallDownLim.Text;
                temp += "\n";

                temp += "CallUpLim=" + tbCallUpLim.Text;
                temp += "\n";

                temp += "PutDownLim=" + tbPutDownLim.Text;
                temp += "\n";

                temp += "PutUpLim=" + tbPutUpLim.Text;
                temp += "\n";

                temp += "UpSafeDis=" + tbUpSafeDis.Text;
                temp += "\n";

                temp += "DownSafeDis=" + tbDownSafeDis.Text;
                temp += "\n";

                temp += "TakeProfitPointsSum=" + tbTakeProfitPointsSum.Text;
                temp += "\n";

                temp += "TakeProfitAcc=" + tbTakeProfitAcc.Text;
                temp += "\n";

                temp += "TakeProfitLots=" + tbTakeProfitLots.Text;
                temp += "\n";

                temp += "TakeProfitGap=" + tbTakeProfitGap.Text;
                temp += "\n";

                temp += "StopLossMagn=" + tbStopLossMagn.Text;
                temp += "\n";

                temp += "StopLossLots=" + tbStopLossLots.Text;
                temp += "\n";

                temp += "StopLossGap=" + tbStopLossGap.Text;
                temp += "\n";

                temp += "EachOrderLots=" + tbEachOrderLots.Text;
                temp += "\n";

                temp += "EachOrderGap=" + tbEachOrderGap.Text;
                temp += "\n";

                temp += "OrderPriceShift=" + tbOrderPriceShift.Text;
                temp += "\n";

                temp += "ProfitStopCountMin=" + tbProfitStopCountMin.Text;
                temp += "\n";

                temp += "ProfitStopSuccessRound=" + tbProfitStopSuccessRound.Text;
                temp += "\n";

                temp += "ProfitStopFailRound=" + tbProfitStopFailRound.Text;
                temp += "\n";

                temp += "EmergencyCountMin=" + tbEmergencyCountMin.Text;
                temp += "\n";

                temp += "EmergencySuccessRound=" + tbEmergencySuccessRound.Text;
                temp += "\n";

                temp += "EmergencyFailRound=" + tbEmergencyFailRound.Text;
                temp += "\n";

                temp += "AtleastMargin=" + tbAtleastMargin.Text;
                temp += "\n";

                try
                {
                    if (!File.Exists(fileName))
                    {
                        // Create a file to write to.
                        File.Create(fileName).Close();
                        File.WriteAllText(fileName, temp);
                    }
                    else File.WriteAllText(fileName, temp);
                }
                catch (IOException ex)
                {
                    opLog.txtLog(
                        ex.GetType().Name + ": The write operation could not " +
                        "be performed because the specified " +
                        "part of the file is locked."
                        );
                }

                orderSettings.LoadOrderSettins();
                todayOrderStatus.TodayRentCall = -1;
                todayOrderStatus.TodayRentPut = -1;

                btSave.Enabled = true;
            }
            catch (Exception ex)
            {
                opLog.txtLog("btSave_Click ex: " + ex);
            }
        }

        private void LoadOrderSettingToForm()
        {
            try
            {
                if (orderSettings.DailyPoints[0] != null)
                {
                    //每日日盤、夜盤分配點數
                    tbWEDNGcall.Text = orderSettings.DailyPoints[0][0].ToString();
                    tbWEDNGput.Text = orderSettings.DailyPoints[0][1].ToString();
                    tbWEDNGtkprofit.Text = orderSettings.DailyPoints[0][2].ToString();

                    tbTHUDTcall.Text = orderSettings.DailyPoints[1][0].ToString();
                    tbTHUDTput.Text = orderSettings.DailyPoints[1][1].ToString();
                    tbTHUDTtkprofit.Text = orderSettings.DailyPoints[1][2].ToString();

                    tbTHUNGcall.Text = orderSettings.DailyPoints[2][0].ToString();
                    tbTHUNGput.Text = orderSettings.DailyPoints[2][1].ToString();
                    tbTHUNGtkprofit.Text = orderSettings.DailyPoints[2][2].ToString();

                    tbFRIDTcall.Text = orderSettings.DailyPoints[3][0].ToString();
                    tbFRIDTput.Text = orderSettings.DailyPoints[3][1].ToString();
                    tbFRIDTtkprofit.Text = orderSettings.DailyPoints[3][2].ToString();

                    tbFRINGcall.Text = orderSettings.DailyPoints[4][0].ToString();
                    tbFRINGput.Text = orderSettings.DailyPoints[4][1].ToString();
                    tbFRINGtkprofit.Text = orderSettings.DailyPoints[4][2].ToString();

                    tbMONDTcall.Text = orderSettings.DailyPoints[5][0].ToString();
                    tbMONDTput.Text = orderSettings.DailyPoints[5][1].ToString();
                    tbMONDTtkprofit.Text = orderSettings.DailyPoints[5][2].ToString();

                    tbMONNGcall.Text = orderSettings.DailyPoints[6][0].ToString();
                    tbMONNGput.Text = orderSettings.DailyPoints[6][1].ToString();
                    tbMONNGtkprofit.Text = orderSettings.DailyPoints[6][2].ToString();

                    tbTUEDTcall.Text = orderSettings.DailyPoints[7][0].ToString();
                    tbTUEDTput.Text = orderSettings.DailyPoints[7][1].ToString();
                    tbTUEDTtkprofit.Text = orderSettings.DailyPoints[7][2].ToString();

                    tbTUENGcall.Text = orderSettings.DailyPoints[8][0].ToString();
                    tbTUENGput.Text = orderSettings.DailyPoints[8][1].ToString();
                    tbTUENGtkprofit.Text = orderSettings.DailyPoints[8][2].ToString();

                    tbWEDDTcall.Text = orderSettings.DailyPoints[9][0].ToString();
                    tbWEDDTput.Text = orderSettings.DailyPoints[9][1].ToString();
                    tbWEDDTtkprofit.Text = orderSettings.DailyPoints[9][2].ToString();

                    //點數大於、點數小於
                    tbCallDownLim.Text = orderSettings.CallDownLim.ToString();
                    tbCallUpLim.Text = orderSettings.CallUpLim.ToString();
                    tbPutDownLim.Text = orderSettings.PutDownLim.ToString();
                    tbPutUpLim.Text = orderSettings.PutUpLim.ToString();

                    //安全距離
                    tbUpSafeDis.Text = orderSettings.UpSafeDis.ToString();
                    tbDownSafeDis.Text = orderSettings.DownSafeDis.ToString();

                    //停利設定
                    tbTakeProfitPointsSum.Text = orderSettings.TakeProfitPointsSum.ToString();
                    tbTakeProfitAcc.Text = orderSettings.TakeProfitAcc.ToString();
                    tbTakeProfitLots.Text = orderSettings.TakeProfitLots.ToString();
                    tbTakeProfitGap.Text = orderSettings.TakeProfitGap.ToString();

                    //停損倍率、停損單間隔秒數
                    tbStopLossMagn.Text = orderSettings.StopLossMagn.ToString();
                    tbStopLossLots.Text = orderSettings.StopLossLots.ToString();
                    tbStopLossGap.Text = orderSettings.StopLossGap.ToString();

                    //下單上限   
                    tbEachOrderLots.Text = orderSettings.EachOrderLots.ToString();
                    tbEachOrderGap.Text = orderSettings.EachOrderGap.ToString();
                    tbOrderPriceShift.Text = orderSettings.OrderPriceShift.ToString();

                    //EmergencyBreak設定
                    tbProfitStopCountMin.Text = orderSettings.ProfitStopCountMin.ToString();
                    tbProfitStopSuccessRound.Text = orderSettings.ProfitStopSuccessRound.ToString();
                    tbProfitStopFailRound.Text = orderSettings.ProfitStopFailRound.ToString();
                    tbEmergencyCountMin.Text = orderSettings.EmergencyCountMin.ToString();
                    tbEmergencySuccessRound.Text = orderSettings.EmergencySuccessRound.ToString();
                    tbEmergencyFailRound.Text = orderSettings.EmergencyFailRound.ToString();
                    tbAtleastMargin.Text = orderSettings.AtleastMargin.ToString();
                }
            }
            catch (IndexOutOfRangeException ex)
            {
                opLog.txtLog("DailyPoints array exception: " + ex);
            }
        }

        private void btFind100_Click_1(object sender, EventArgs e)
        {
            CalculateComID calcuComID;
            if (string.Compare(DateTime.Now.DayOfWeek.ToString("d"), "3") == 0 && DateCalculateExtensions.IsLegalTime(DateTime.Now, "050001-143000"))
                calcuComID = new CalculateComID(DateTime.Now.AddDays(-1));
            else calcuComID = new CalculateComID(DateTime.Now);

            string[] symbol100 = findHundPointInHundred();
            opLog.txtLog("symbol100[0]: " + symbol100[0]);
            opLog.txtLog("symbol100[1]: " + symbol100[1]);
            opLog.txtLog("symbol100[2]: " + symbol100[2]);
            opLog.txtLog("symbol100[3]: " + symbol100[3]);

            tb20Code.Text = symbol100[0];
            tbTStrikePrice2.Text = symbol100[0].Substring(9, 5);
            tbTSide2.Text = "S";
            tbTComym2.Text = calcuComID.PartialComIDC[2];
            tbTStrikePrice.Text = decimal.Parse(symbol100[0].Substring(3, 5)).ToString();
            tbTSide.Text = "B";
            tbTComym.Text = calcuComID.PartialComIDC[2];
            rbMulti.Checked = true;

        }

        private void button3_Click(object sender, EventArgs e)
        {
            findNearestNowPriceInStock();
        }

        private void btSettingReset_Click(object sender, EventArgs e)
        {
            tbWEDNGcall.Text = "300";
            tbWEDNGput.Text = "300";
            tbWEDNGtkprofit.Text = "4";

            tbTHUDTcall.Text = "300";
            tbTHUDTput.Text = "300";
            tbTHUDTtkprofit.Text = "4";

            tbTHUNGcall.Text = "300";
            tbTHUNGput.Text = "300";
            tbTHUNGtkprofit.Text = "3.5";

            tbFRIDTcall.Text = "300";
            tbFRIDTput.Text = "300";
            tbFRIDTtkprofit.Text = "3";

            tbFRINGcall.Text = "300";
            tbFRINGput.Text = "300";
            tbFRINGtkprofit.Text = "2.5";

            tbMONDTcall.Text = "300";
            tbMONDTput.Text = "300";
            tbMONDTtkprofit.Text = "2";

            tbMONNGcall.Text = "300";
            tbMONNGput.Text = "300";
            tbMONNGtkprofit.Text = "1.5";

            tbTUEDTcall.Text = "300";
            tbTUEDTput.Text = "300";
            tbTUEDTtkprofit.Text = "1";

            tbTUENGcall.Text = "300";
            tbTUENGput.Text = "300";
            tbTUENGtkprofit.Text = "1";

            tbWEDDTcall.Text = "300";
            tbWEDDTput.Text = "300";
            tbWEDDTtkprofit.Text = "1";

            //點數大於、點數小於
            tbCallDownLim.Text = "15.00";
            tbCallUpLim.Text = "16.00";
            tbPutDownLim.Text = "13.50";
            tbPutUpLim.Text = "15.00";

            //安全距離
            tbUpSafeDis.Text = "18250";
            tbDownSafeDis.Text = "17900";

            //停利設定
            tbTakeProfitPointsSum.Text = "2100";
            tbTakeProfitAcc.Text = "20.70";
            tbTakeProfitLots.Text = "10";
            tbTakeProfitGap.Text = "3";

            //停損倍率、停損單間隔秒數
            tbStopLossMagn.Text = "2.5";
            tbStopLossLots.Text = "1";
            tbStopLossGap.Text = "5";

            //下單上限
            tbEachOrderLots.Text = "10";
            tbEachOrderGap.Text = "15";
            tbOrderPriceShift.Text = "0.6";

            //EmergencyBreak設定
            tbProfitStopCountMin.Text = "1";
            tbProfitStopSuccessRound.Text = "10";
            tbProfitStopFailRound.Text = "70";
            tbEmergencyCountMin.Text = "1";
            tbEmergencySuccessRound.Text = "10";
            tbEmergencyFailRound.Text = "10";
            tbAtleastMargin.Text = "20000";
        }

        private void btLoadSetting_Click(object sender, EventArgs e)
        {
            LoadOrderSettingToForm();
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            List<string> sb = new List<string>();
            string temp = "";
            if (listView1.SelectedItems.Count > 0)
            {
                ListView.SelectedListViewItemCollection selected = listView1.SelectedItems;
                //foreach (ListViewItem item in selected)
                //        sb.Append(item.ToString()).Append(Environment.NewLine);
                //tbTest.Text = sb.ToString();

                //for (int i=0; i<selected[0].SubItems.Count; i++)
                //{
                //    sb.Add(selected[0].SubItems[i].ToString());
                //}
                //for (int i = 0; i < sb.Count; i++)
                //{{
                //    temp += sb[i];
                //}
                if (string.Compare(tbTCallPut2.Text, "CALL") == 0 )
                {
                    temp = selected[0].SubItems[8].ToString();
                    if ( temp.Length-1 - temp.IndexOf("{", 0, temp.Length - 1) > 2)
                    {
                        if (temp.IndexOf("{", 0, temp.Length) > 1)
                            temp = temp.Substring(temp.IndexOf("{", 0, temp.Length - 1) + 1, 10);
                        temp = temp.Insert(3, (decimal.Parse(temp.Substring(3, 5)) + 100).ToString() + "/");
                        //tbTest.Text = temp;
                        tb20Code.Text = temp;
                    }
                }
                else if (string.Compare(tbTCallPut2.Text, "PUT") == 0 )
                {
                    temp = selected[0].SubItems[10].ToString();
                    if (temp.Length-1 - temp.IndexOf("{", 0, temp.Length - 1) > 2)
                    {
                        if (temp.IndexOf("{", 0, temp.Length) > 1)
                            temp = temp.Substring(temp.IndexOf("{", 0, temp.Length - 1) + 1, 10);
                        temp = temp.Insert(3, (decimal.Parse(temp.Substring(3, 5)) - 100).ToString()+"/");
                        //tbTest.Text = temp;
                        tb20Code.Text = temp;
                    }
                }
                //temp = selected[0].SubItems[8].ToString();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //updateOI();
            try
            {
                TComid();
            }
            catch (Exception ex)
            {
                opLog.txtLog("TComid ex: "+ex);
            }

        }
        private void updateOI()
        {
            CalculateComID calcuComID;
            if (string.Compare(DateTime.Now.DayOfWeek.ToString("d"), "3") == 0 && DateCalculateExtensions.IsLegalTime(DateTime.Now, "050001-143000"))
                calcuComID = new CalculateComID(DateTime.Now.AddDays(-1));
            else calcuComID = new CalculateComID(DateTime.Now);

            string combineSymbol = "";
            char[] temp= new char[2];
            taifexCrawler.doCrawler();
            for(int i=0; i< taifexCrawler.TaiFexList.Count(); i++)
            {
                if (taifexCrawler.TaiFexList[i] != null)
                {

                    //calcuComID.FindSymbol(int.Parse(taifexCrawler.TaiFexList[i][1].Substring(4, 2)), temp);
                    //Console.WriteLine("calcuComID.PartialComIDC[3]: "+ calcuComID.PartialComIDC[3]);
                    if (string.Compare(calcuComID.PartialComIDC[3], taifexCrawler.TaiFexList[i][1]) ==0)
                    {
                        if (string.Compare(taifexCrawler.TaiFexList[i][3], "Call") == 0)
                        {
                            combineSymbol = calcuComID.PartialComIDC[0] + taifexCrawler.TaiFexList[i][2] + calcuComID.PartialComIDC[1];
                            //Console.WriteLine("crawlercombineSymbol: " + combineSymbol);
                            if (!string.IsNullOrEmpty(taifexCrawler.TaiFexList[i][14]) && string.Compare(taifexCrawler.TaiFexList[i][14], "-") != 0)
                            {
                                //Console.WriteLine("crawlercombineSymbol2: " + combineSymbol);
                                if (symbolOPriceMap.ContainsKey(combineSymbol))
                                {
                                    //Console.WriteLine("OI: " + taifexCrawler.TaiFexList[i][14]);
                                    //Console.WriteLine("TOTALVOL: " + (decimal.Parse(taifexCrawler.TaiFexList[i][14]) + decimal.Parse(symbolOPriceMap[combineSymbol][(int)OPTION_LIST.VOL])).ToString());
                                    symbolOPriceMap[combineSymbol][(int)OPTION_LIST.OI] = taifexCrawler.TaiFexList[i][14];
                                    symbolOPriceMap[combineSymbol][(int)OPTION_LIST.TOTALVOL] = ( decimal.Parse(taifexCrawler.TaiFexList[i][14])+decimal.Parse(symbolOPriceMap[combineSymbol][(int)OPTION_LIST.VOL]) ).ToString() ;
                                }
                            }
                        }
                        else if (string.Compare(taifexCrawler.TaiFexList[i][3], "Put") == 0)
                        {
                            combineSymbol = calcuComID.PartialComIDP[0] + taifexCrawler.TaiFexList[i][2] + calcuComID.PartialComIDP[1];
                            if (!string.IsNullOrEmpty(taifexCrawler.TaiFexList[i][14]) && string.Compare(taifexCrawler.TaiFexList[i][14], "-") != 0)
                            {
                                if (symbolOPriceMap.ContainsKey(combineSymbol))
                                {
                                    symbolOPriceMap[combineSymbol][(int)OPTION_LIST.OI] = taifexCrawler.TaiFexList[i][14];
                                    symbolOPriceMap[combineSymbol][(int)OPTION_LIST.TOTALVOL] = (decimal.Parse(taifexCrawler.TaiFexList[i][14]) + decimal.Parse(symbolOPriceMap[combineSymbol][(int)OPTION_LIST.VOL])).ToString();
                                }
                            }
                        }
                    }
                }

            }
            opLog.txtLog("UpdateOI!");
            AddInfoTG("UpdateOI!");
            AddInfo("UpdateOI!");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //taifexCrawler.doCrawler();
            CalculateComID calcuComID;
            if (string.Compare(DateTime.Now.DayOfWeek.ToString("d"), "3") == 0 && DateCalculateExtensions.IsLegalTime(DateTime.Now, "050001-143000"))
                calcuComID = new CalculateComID(DateTime.Now.AddDays(-1));
            else calcuComID = new CalculateComID(DateTime.Now);

            string combineSymbol = "";
            char[] temp = new char[2];
            taifexCrawler.doCrawler();
            for (int i = 0; i < taifexCrawler.TaiFexList.Count(); i++)
            {
                if (taifexCrawler.TaiFexList[i] != null)
                {
                    //calcuComID.FindSymbol(int.Parse(taifexCrawler.TaiFexList[i][1].Substring(4, 2)), temp);
                    if (string.Compare(calcuComID.PartialComIDC[3], taifexCrawler.TaiFexList[i][1]) == 0)
                    {
                        if (string.Compare(taifexCrawler.TaiFexList[i][3], "Call") == 0)
                        {
                            combineSymbol = calcuComID.PartialComIDC[0] + taifexCrawler.TaiFexList[i][2] + calcuComID.PartialComIDC[1];
                            if (!string.IsNullOrEmpty(taifexCrawler.TaiFexList[i][14]) && string.Compare(taifexCrawler.TaiFexList[i][14], "-") != 0)
                            {
                                if (symbolOPriceMap.ContainsKey(combineSymbol))
                                {
                                    symbolOPriceMap[combineSymbol][(int)OPTION_LIST.OI] = taifexCrawler.TaiFexList[i][14];
                                    symbolOPriceMap[combineSymbol][(int)OPTION_LIST.TOTALVOL] = (decimal.Parse(taifexCrawler.TaiFexList[i][14]) + decimal.Parse(symbolOPriceMap[combineSymbol][(int)OPTION_LIST.VOL])).ToString();
                                }
                            }
                        }
                        else if (string.Compare(taifexCrawler.TaiFexList[i][3], "Put") == 0)
                        {
                            combineSymbol = calcuComID.PartialComIDP[0] + taifexCrawler.TaiFexList[i][2] + calcuComID.PartialComIDP[1];
                            if (!string.IsNullOrEmpty(taifexCrawler.TaiFexList[i][14]) && string.Compare(taifexCrawler.TaiFexList[i][14], "-") != 0)
                            {
                                if (symbolOPriceMap.ContainsKey(combineSymbol))
                                {
                                    symbolOPriceMap[combineSymbol][(int)OPTION_LIST.OI] = taifexCrawler.TaiFexList[i][14];
                                    symbolOPriceMap[combineSymbol][(int)OPTION_LIST.TOTALVOL] = (decimal.Parse(taifexCrawler.TaiFexList[i][14]) + decimal.Parse(symbolOPriceMap[combineSymbol][(int)OPTION_LIST.VOL])).ToString();
                                }
                            }
                        }
                    }
                }

            }
            opLog.txtLog("UpdateOI!");
            AddInfoTG("UpdateOI!");
            AddInfo("UpdateOI!");
        }

        private async void btLogin_Click(object sender, EventArgs e)
        {
            accInfo = new AccountInfo(settings.Account);
            if (string.Compare(settings.TestMode, "1") == 0)
                accInfo.StrPwd = "0000";

            Login();
            //訂閱、開啟Timer
            await Task.Delay(12000);
            TComid();
            checkPositionSum();
            tmrUnsub.Start();
            tmrPosition.Start();
            tmrUpdateOI.Start();
            await Task.Delay(5000);
            updateOI();
            loginBtnFlag = true;
        }

        private void btStart_Click(object sender, EventArgs e)
        {
            if (loginBtnFlag == true && loopOrderTempSwitch == false)
            {
                loopOrderTempSwitch = true;
                thrOrder.Start();
                //Console.WriteLine("thrOrder started!");
                AddInfo("自動下單啟動!");
                AddInfoTG("自動下單啟動!");
                opLog.txtLog("自動下單啟動!");
            }
            else if(loopOrderTempSwitch == true) MessageBox.Show("您已經啟動自動下單");
            else MessageBox.Show("您尚未登入完成，請先登入完成再啟動自動下單!");
        }

        private async void btLogout_Click(object sender, EventArgs e)
        {
            quoteCom.Logout();
            await Task.Delay(5000);
            tfcom.Logout();
            await Task.Delay(5000);
        }

        private void btPause_Click(object sender, EventArgs e)
        {
            if (loopOrderTempSwitch == true)
            {
                loopOrderTempSwitch = false;
                AddInfo("自動下單暫停!");
                AddInfoTG("自動下單暫停!");
                opLog.txtLog("自動下單暫停!");
            }
            else
            {
                loopOrderTempSwitch = true;
                AddInfo("自動下單重啟!");
                AddInfoTG("自動下單重啟!");
                opLog.txtLog("自動下單重啟!");
            }
        }

        private void btTestData_Click(object sender, EventArgs e)
        {
            addTestRecord("TX417350X1", "12.5", "1278");
            addTestRecord("TX417400X1", "16.5", "2798");
            addTestRecord("TX417450X1", "22", "3009");
            addTestRecord("TX417500X1", "28", "7175");
            addTestRecord("TX417550X1", "33", "3610");
            addTestRecord("TX417600X1", "42", "7009");
            addTestRecord("TX417650X1", "50", "5355");
            addTestRecord("TX417700X1", "67", "9554");
            addTestRecord("TX417750X1", "78", "7650");
            addTestRecord("TX417800X1", "104", "8039");
            addTestRecord("TX417850X1", "133", "3280");

            addTestRecord("TX417600L1", "221", "1278");
            addTestRecord("TX417650L1", "186", "2798");
            addTestRecord("TX417700L1", "150", "3009");
            addTestRecord("TX417750L1", "118", "7175");
            addTestRecord("TX417800L1", "91", "3610");
            addTestRecord("TX417850L1", "66", "7009");
            addTestRecord("TX417900L1", "45", "5355");
            addTestRecord("TX417950L1", "33.5", "9554");
            addTestRecord("TX418000L1", "20.5", "7650");
            addTestRecord("TX418050L1", "12.5", "8039");
            addTestRecord("TX418100L1", "7.1", "3280");

            /*foreach(string key in symbolFPriceMap.Keys)
                foreach(var sList in symbolFPriceMap[key])
                        listView1.Items.Add(sList);*/
            foreach (string key in symbolOPriceMap.Keys)
                foreach (var sList in symbolOPriceMap[key])
                    Console.WriteLine(sList);
        }

        private void dataGridView1_CellMouseUp(object sender, DataGridViewCellMouseEventArgs e)
        {
            //在dataGridView1按右鍵時也一併選取那個Cell
            if (e.Button == System.Windows.Forms.MouseButtons.Right && e.RowIndex >= 0)
            {
                dataGridView1.CurrentCell = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex];
                AddInfo("e.RowIndex: " + e.RowIndex + "e.ColumnIndex" + e.ColumnIndex);
                if (dataGridView1.CurrentCell.ColumnIndex > 0 && dataGridView1.CurrentCell.ColumnIndex < dataGridView1.Columns["StrickPriceCenter"].Index)
                {
                    if (dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0] != null && string.Compare(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString(), "-") != 0)
                        tbUpSafeDis.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString();
                }
                else if (dataGridView1.CurrentCell.ColumnIndex > dataGridView1.Columns["StrickPriceCenter"].Index && dataGridView1.CurrentCell.ColumnIndex < dataGridView1.ColumnCount)
                {
                    if (dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0] != null && string.Compare(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString(), "-") != 0)
                        tbDownSafeDis.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString();
                }
            }
        }

        private void 填入安全邊界ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if( dataGridView1.CurrentCell.ColumnIndex > 0 && dataGridView1.CurrentCell.ColumnIndex < dataGridView1.Columns["StrickPriceCenter"].Index)
            {
                if(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0]!=null && string.Compare( dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString(), "-")!=0  )
                    tbUpSafeDis.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString();
            }
            else if (dataGridView1.CurrentCell.ColumnIndex > dataGridView1.Columns["StrickPriceCenter"].Index && dataGridView1.CurrentCell.ColumnIndex < dataGridView1.ColumnCount)
            {
                if (dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0] != null && string.Compare(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString(), "-") != 0)
                    tbDownSafeDis.Text = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString();
            }
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            decimal whichDif = 0;
            if (string.Compare(cbWhichDif.Text, "100差") == 0) whichDif = 100;
            else if(string.Compare(cbWhichDif.Text, "150差") == 0) whichDif = 150;

            CalculateComID calcuComID;
            if (string.Compare(DateTime.Now.DayOfWeek.ToString("d"), "3") == 0 && DateCalculateExtensions.IsLegalTime(DateTime.Now, "050001-143000"))
                calcuComID = new CalculateComID(DateTime.Now.AddDays(-1));
            else calcuComID = new CalculateComID(DateTime.Now);

            string strikePrice="";
            string strikePrice2="";

            if (dataGridView1.CurrentCell.ColumnIndex > 0 && dataGridView1.CurrentCell.ColumnIndex < dataGridView1.Columns["StrickPriceCenter"].Index)
            {
                if (dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0] != null && string.Compare(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString(), "-") != 0)
                {
                    strikePrice = (decimal.Parse(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString()) + whichDif).ToString();
                    strikePrice2 = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString();
                    tbTStrikePrice.Text = strikePrice;
                    tbTStrikePrice2.Text = strikePrice2;
                    tbTCallPut.Text = "CALL";
                    tbTCallPut2.Text = "CALL";
                    tbTComym.Text = calcuComID.PartialComIDC[2];
                    tbTComym2.Text = calcuComID.PartialComIDC[2];
                    tbTSymbol.Text = calcuComID.PartialComIDC[0];
                    tb20Code.Text = calcuComID.PartialComIDC[0] + strikePrice + "/" + strikePrice2 + calcuComID.PartialComIDC[1];
                    tbTPrice.Text = orderSettings.CallDownLim.ToString();
                    tbTSide.Text = "B";
                    tbTSide2.Text = "S";
                    cbFunc.Text = "O.下單";
                    cbMarket.Text = "O.選擇權";
                    cbROD.Text = "IOC";
                    cbOpen.Text = "O 新倉";
                    rbMulti.Checked = true;
                }
            }
            else if (dataGridView1.CurrentCell.ColumnIndex > dataGridView1.Columns["StrickPriceCenter"].Index && dataGridView1.CurrentCell.ColumnIndex < dataGridView1.ColumnCount)
            {
                if (dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0] != null && string.Compare(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString(), "-") != 0)
                {
                    strikePrice = (decimal.Parse(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString()) - whichDif).ToString();
                    strikePrice2 = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString();
                    tbTStrikePrice.Text = strikePrice;
                    tbTStrikePrice2.Text = strikePrice2;
                    tbTCallPut.Text = "PUT";
                    tbTCallPut2.Text = "PUT";
                    tbTComym.Text = calcuComID.PartialComIDP[2];
                    tbTComym2.Text = calcuComID.PartialComIDP[2];
                    tbTSymbol.Text = calcuComID.PartialComIDP[0];
                    tb20Code.Text = calcuComID.PartialComIDP[0] + strikePrice + "/" + strikePrice2 + calcuComID.PartialComIDP[1];
                    tbTPrice.Text = orderSettings.PutDownLim.ToString();
                    tbTSide.Text = "B";
                    tbTSide2.Text = "S";
                    cbFunc.Text = "O.下單";
                    cbMarket.Text = "O.選擇權";
                    cbROD.Text = "IOC";
                    cbOpen.Text = "O 新倉";
                    rbMulti.Checked = true;
                }
            }
        }

        private void btTbarUp_Click(object sender, EventArgs e)
        {
            btTbarUp.Enabled = false;
            TbarShift += 1;
            btTbarShift.Text = TbarShift.ToString();
            btTbarUp.Enabled = true;
        }

        private void btTbarDown_Click(object sender, EventArgs e)
        {
            btTbarDown.Enabled = false;
            TbarShift -= 1;
            btTbarShift.Text = TbarShift.ToString();
            btTbarDown.Enabled = true;
        }

        private async void btTbarShift_Click(object sender, EventArgs e)
        {
            btTbarShift.Enabled = false;
            await Task.Delay(2000);
            TComid();
            await Task.Delay(5000);
            btTbarShift.Enabled = true;
        }
    }
}
