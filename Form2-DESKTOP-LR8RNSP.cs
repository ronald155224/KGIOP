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

namespace KGIOP
{
    public partial class Form2 : Form
    {
        OrderSettings orderSetting = new OrderSettings(@"orderSetting.txt");
        public Form2()
        {
            InitializeComponent();
            LoadOrderSettingToForm();
        }

        private void LoadOrderSettingToForm()
        {
            try
            {
                if (orderSetting.DailyPoints[0] != null)
                {
                    tbWEDNGcall.Text = orderSetting.DailyPoints[0][0].ToString();
                    tbWEDNGput.Text = orderSetting.DailyPoints[0][1].ToString();
                    tbWEDNGdis.Text = orderSetting.DailyPoints[0][2].ToString();

                    tbTHUDTcall.Text = orderSetting.DailyPoints[1][0].ToString();
                    tbTHUDTput.Text = orderSetting.DailyPoints[1][1].ToString();
                    tbTHUDTdis.Text = orderSetting.DailyPoints[1][2].ToString();

                    tbTHUNGcall.Text = orderSetting.DailyPoints[2][0].ToString();
                    tbTHUNGput.Text = orderSetting.DailyPoints[2][1].ToString();
                    tbTHUNGdis.Text = orderSetting.DailyPoints[2][2].ToString();

                    tbFRIDTcall.Text = orderSetting.DailyPoints[3][0].ToString();
                    tbFRIDTput.Text = orderSetting.DailyPoints[3][1].ToString();
                    tbFRIDTdis.Text = orderSetting.DailyPoints[3][2].ToString();

                    tbFRINGcall.Text = orderSetting.DailyPoints[4][0].ToString();
                    tbFRINGput.Text = orderSetting.DailyPoints[4][1].ToString();
                    tbFRINGdis.Text = orderSetting.DailyPoints[4][2].ToString();

                    tbMONDTcall.Text = orderSetting.DailyPoints[5][0].ToString();
                    tbMONDTput.Text = orderSetting.DailyPoints[5][1].ToString();
                    tbMONDTdis.Text = orderSetting.DailyPoints[5][2].ToString();

                    tbMONNGcall.Text = orderSetting.DailyPoints[6][0].ToString();
                    tbMONNGput.Text = orderSetting.DailyPoints[6][1].ToString();
                    tbMONNGdis.Text = orderSetting.DailyPoints[6][2].ToString();

                    tbTUEDTcall.Text = orderSetting.DailyPoints[7][0].ToString();
                    tbTUEDTput.Text = orderSetting.DailyPoints[7][1].ToString();
                    tbTUEDTdis.Text = orderSetting.DailyPoints[7][2].ToString();

                    tbTUENGcall.Text = orderSetting.DailyPoints[8][0].ToString();
                    tbTUENGput.Text = orderSetting.DailyPoints[8][1].ToString();
                    tbTUENGdis.Text = orderSetting.DailyPoints[8][2].ToString();

                    tbWEDDTcall.Text = orderSetting.DailyPoints[9][0].ToString();
                    tbWEDDTput.Text = orderSetting.DailyPoints[9][1].ToString();
                    tbWEDDTdis.Text = orderSetting.DailyPoints[9][2].ToString();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("DailyPoints array exception: " + ex);
            }

        }
        class OrderSettings
        {
            public string StrSettingIni;
            public List<int>[] DailyPoints { get; set; } = {};//DailyPoints[]: WED日,WED夜,THR日... DailyPoints[][0]:callpoints [][1]:putpoints [][2]:安全距離
            public int SumCallPoints { get; set; }
            public int SumPutPoints { get; set; }
            public int StopLoss { get; set; }
            //public string DailyOrderStopLossLim { get; set; }

            public OrderSettings(string strSettingIni)
            {
                StrSettingIni = strSettingIni;
                string allSetting = "";
                string[] strSettingSplit;
                DailyPoints = new List<int>[10];

                if (!File.Exists(strSettingIni))
                {
                    Console.WriteLine("The setting.ini file doesn't exist."
                        + Environment.NewLine
                        + "Create this file automatically!");

                    string dir = Application.StartupPath;
                    string fileName = dir + @"\"+ strSettingIni;
                    
                    File.Create(fileName).Close();
                }
                else
                {
                    allSetting = System.IO.File.ReadAllText(strSettingIni);

                    if (string.IsNullOrEmpty(allSetting))
                    {
                        Console.WriteLine("The setting is null."
                            + Environment.NewLine
                            + "Please fill this file with setting info.");
                    }
                    else
                    {
                        string[] tempS = new string[2];
                        allSetting = allSetting.Trim();
                        strSettingSplit = allSetting.Split('\n');
                        string[] members = new string[3];

                        foreach (var sub in strSettingSplit)
                        {
                            tempS = sub.Split('=');
                            if (String.Compare(tempS[0], "WEDNG") == 0)
                            {
                                members = tempS[1].Split(',');
                                for(int i=0; i < members.Length; i++)
                                    DailyPoints[0].Add(int.Parse(members[i]));
                            }
                            else if (String.Compare(tempS[0], "THUDT") == 0)
                            {
                                members = tempS[1].Split(',');
                                for (int i = 0; i < members.Length; i++)
                                    DailyPoints[1].Add(int.Parse(members[i]));
                            }
                            else if (String.Compare(tempS[0], "THUNG") == 0)
                            {
                                members = tempS[1].Split(',');
                                for (int i = 0; i < members.Length; i++)
                                    DailyPoints[2].Add(int.Parse(members[i]));
                            }
                            else if (String.Compare(tempS[0], "FRIDT") == 0)
                            {
                                members = tempS[1].Split(',');
                                for (int i = 0; i < members.Length; i++)
                                    DailyPoints[3].Add(int.Parse(members[i]));
                            }
                            else if (String.Compare(tempS[0], "FRING") == 0)
                            {
                                members = tempS[1].Split(',');
                                for (int i = 0; i < members.Length; i++)
                                    DailyPoints[4].Add(int.Parse(members[i]));
                            }
                            else if (String.Compare(tempS[0], "MONDT") == 0)
                            {
                                members = tempS[1].Split(',');
                                for (int i = 0; i < members.Length; i++)
                                    DailyPoints[5].Add(int.Parse(members[i]));
                            }
                            else if (String.Compare(tempS[0], "MONNG") == 0)
                            {
                                members = tempS[1].Split(',');
                                for (int i = 0; i < members.Length; i++)
                                    DailyPoints[6].Add(int.Parse(members[i]));
                            }
                            else if (String.Compare(tempS[0], "TUEDT") == 0)
                            {
                                members = tempS[1].Split(',');
                                for (int i = 0; i < members.Length; i++)
                                    DailyPoints[7].Add(int.Parse(members[i]));
                            }
                            else if (String.Compare(tempS[0], "TUENG") == 0)
                            {
                                members = tempS[1].Split(',');
                                for (int i = 0; i < members.Length; i++)
                                    DailyPoints[8].Add(int.Parse(members[i]));
                            }
                            else if (String.Compare(tempS[0], "WEDDT") == 0)
                            {
                                members = tempS[1].Split(',');
                                for (int i = 0; i < members.Length; i++)
                                    DailyPoints[9].Add(int.Parse(members[i]));
                            }
                            else if (String.Compare(tempS[0], "SumCallPoints") == 0)
                            {
                                SumCallPoints = int.Parse(tempS[1]);
                            }
                            else if (String.Compare(tempS[0], "SumPutPoints") == 0)
                            {
                                SumPutPoints = int.Parse(tempS[1]);
                            }
                            else if (String.Compare(tempS[0], "StopLoss") == 0)
                            {
                                StopLoss = int.Parse(tempS[1]);
                            }
                        }
                    }
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string dir = Application.StartupPath;
            string fileName = dir + @"\" + orderSetting.StrSettingIni;
            string temp = "";

            temp += "WEDNG=" + tbWEDNGcall.Text + ","+ tbWEDNGput.Text + "," + tbWEDNGdis.Text;
            temp += "\n";

            temp += ("THUDT=" + tbTHUDTcall.Text + "," + tbTHUDTput.Text + "," + tbTHUDTdis.Text);
            temp += "\n";

            temp += ("THUNG=" + tbTHUNGcall.Text + "," + tbTHUNGput.Text + "," + tbTHUNGdis.Text);
            temp += "\n";

            temp += ("FRIDT="+ tbFRIDTcall.Text + "," + tbFRIDTput.Text + "," + tbFRIDTdis.Text);
            temp += "\n";

            temp += ("FRING="+ tbFRINGcall.Text + "," + tbFRINGput.Text + "," + tbFRINGdis.Text);
            temp += "\n";

            temp += ("MONDT="+ tbMONDTcall.Text + "," + tbMONDTput.Text + "," + tbMONDTdis.Text);
            temp += "\n";

            temp += ("MONNG="+ tbMONNGcall.Text + "," + tbMONNGput.Text + "," + tbMONNGdis.Text);
            temp += "\n";

            temp += ("TUEDT="+ tbTUEDTcall.Text + "," + tbTUEDTput.Text + "," + tbTUEDTdis.Text);
            temp += "\n";

            temp += ("TUENG="+ tbTUENGcall.Text + "," + tbTUENGput.Text + "," + tbTUENGdis.Text);
            temp += "\n";

            temp += ("WEDDT="+ tbWEDDTcall.Text + "," + tbWEDDTput.Text + "," + tbWEDDTdis.Text);
            temp += "\n";

            if (!File.Exists(fileName))
            {
                // Create a file to write to.
                File.WriteAllText(fileName, temp);
            }
            else File.WriteAllText(fileName, temp);
        }
    }
}
