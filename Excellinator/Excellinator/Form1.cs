using System;
using System.Web;
using System.ComponentModel;
using System.Data;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Windows.Forms;

namespace Excellinator
{
    public partial class Form1 : Form
    {
        string fileName = "";
        Form1x App = new Form1x();
        public Form1()
        {
            InitializeComponent();
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            pbProgress.Value = 0;

            _Application myApp = new Microsoft.Office.Interop.Excel.Application();
            _Workbook myWorkbook = myApp.Workbooks.Add(Type.Missing);
            _Worksheet myWorksheet = null;

            myWorksheet = myWorkbook.Sheets["Sheet1"];
            myWorksheet = myWorkbook.ActiveSheet;
            myWorksheet.Name = "Activity log";

            for (int i = 1; i < dgActivityLogs.Columns.Count + 1; i++)
            {
                myWorksheet.Cells[1, i] = dgActivityLogs.Columns[i - 1].HeaderText;
            }

            for (int r = 0; r < dgActivityLogs.Rows.Count; r++)
            {
                for (int c = 0; c < dgActivityLogs.Columns.Count; c++)
                {
                    try
                    {
                        myWorksheet.Cells[r + 2, c + 1] = dgActivityLogs.Rows[r].Cells[c].Value.ToString();
                    }
                    catch
                    {
                        break;
                    }
                    pbProgress.Increment(1);
                }
               
            }            

            var saveFileDialoge = new SaveFileDialog();
            saveFileDialoge.FileName = ofdImport.SafeFileName.TrimEnd(".txt".ToCharArray());
            saveFileDialoge.DefaultExt = ".xlsx";

            pbProgress.Value = 4000;

            if (saveFileDialoge.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    myWorkbook.SaveAs(saveFileDialoge.FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                }
                catch
                {

                }
            }

            myApp.Quit();

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            AcceptButton = btnImport;
            App.Start();
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            pbProgress.Value = 0;

            dgActivityLogs.Controls.Clear();

            try
            {
                DialogResult result = ofdImport.ShowDialog();
                if (result == DialogResult.OK) // Test result.
                {
                    fileName = ofdImport.FileName;
                }
            }
            catch
            {
                MessageBox.Show("Please select a file to import.", "Import file");
            }

            string[] myArray;
            string Duration = "";
            string EventDesc = "";
            string start_time = "";
            System.Data.DataTable myLogs;
            string[] myDate = new string[3];
            TimeSpan myDuration = new TimeSpan();

            using (myLogs = new System.Data.DataTable())
            {
                myLogs.Columns.Add(new DataColumn("Channel Name", typeof(string)));
                myLogs.Columns.Add(new DataColumn("Events Date", typeof(string)));
                myLogs.Columns.Add(new DataColumn("Start Time", typeof(string)));
                myLogs.Columns.Add(new DataColumn("End Time", typeof(string)));
                myLogs.Columns.Add(new DataColumn("Duration", typeof(string)));
                myLogs.Columns.Add(new DataColumn("Events type", typeof(string)));
                myLogs.Columns.Add(new DataColumn("Description of Events", typeof(string)));
                myLogs.Columns.Add(new DataColumn("Flighting Code", typeof(string)));
                myLogs.Columns.Add(new DataColumn("Tape number/House number", typeof(string)));
                myLogs.Columns.Add(new DataColumn("Advertiser name", typeof(string)));
                myLogs.Columns.Add(new DataColumn("Category", typeof(string)));

                StreamReader myReader = new StreamReader(fileName);


                while (myReader.EndOfStream == false)
                {
                    myArray = myReader.ReadLine().Split(',');
                    EventDesc = "";
                    try
                    {
                        TimeSpan.Parse(myArray[10].Trim('"')).ToString();
                    }
                    catch
                    {
                        try
                        {
                            myArray[10] = TimeSpan.Parse("00:00:00").ToString();
                        }
                        catch
                        {
                            continue;
                        }
                    }

                    try
                    {
                        myArray[1] = myArray[1].Trim('"').Substring(0, 8);
                        myArray[10] = myArray[10].Trim('"').Remove(8, 3);
                    }
                    catch
                    {

                    }
                    try
                    {
                        myDuration = TimeSpan.FromSeconds(double.Parse(myArray[11].Trim('"')));
                    }
                    catch
                    {
                        try
                        {
                            myArray[11] = myArray[11].Replace('.', ',');
                            myDuration = TimeSpan.FromSeconds(double.Parse(myArray[11].Trim('"')));
                        }
                        catch
                        {
                            myDuration = TimeSpan.Parse("00:00:00");
                        }
                    }

                    if (myArray[11] == "")
                    {
                        Duration = myArray[10].Trim('"');
                    }
                    else
                    {
                        Duration = myDuration.Hours.ToString().PadLeft(2, '0') + ":" + myDuration.Minutes.ToString().PadLeft(2, '0') + ":" + myDuration.Seconds.ToString().PadLeft(2, '0');
                    }

                    start_time = (TimeSpan.Parse(myArray[1].Trim('"')).Subtract(TimeSpan.Parse(Duration))).ToString();

                    if (start_time.Contains("-"))
                    {
                        try
                        {
                            start_time = (TimeSpan.Parse("23:59:59") + TimeSpan.Parse(start_time)).ToString();
                        }
                        catch
                        {

                        }
                    }

                    if (myArray[3].Trim('"').ToUpper() == "STOP" && myArray[6].Trim('"') != "Show   as a logo " && !myArray[5].Trim('"').Contains("Cinegy Type Layer "))
                    {
                        try
                        {
                            if (IsNumeric(myArray[5].Substring(1, 4)))
                            {
                                continue;
                            }
                        }
                        catch
                        {
                            
                        }

                        if (myArray[6] != "Filler")
                        {
                            if ((myArray[5].ToUpper()).Contains("EZM"))
                            {
                                myArray[6] = "Program";
                                EventDesc = "Ezabantwana";
                            }
                            else if ((myArray[5]).Contains("DSTV") || (myArray[5].ToUpper()).Contains("MULC") || (myArray[5].ToUpper()).Contains("BACR") || (myArray[5].ToUpper()).Contains("BACR") || (myArray[5].ToUpper()).Contains("ERDA") || (myArray[5].ToUpper()).Contains("BACW") || (myArray[5].ToUpper()).Contains("IKGB"))
                            {
                                myArray[6] = "Spot";
                                EventDesc = "DSTV promo";
                            }
                            else if (((myArray[5].ToUpper()).Contains("PROMO") || (myArray[5].Trim('"')).Contains("BRANDER")) && !(myArray[5].Trim('"').ToUpper()).Contains("STARSAT"))
                            {
                                myArray[6] = "Spot";
                                EventDesc = "Channel promo";
                            }
                            else if ((myArray[5]).Contains("KZNNT"))
                            {
                                myArray[6] = "Program";
                                EventDesc = "KZN News talk";
                            }
                            else if ((myArray[5]).Contains("SIT"))
                            {
                                myArray[6] = "Program";
                                EventDesc = "1 Sithombe";
                            }
                            else if ((myArray[5]).Contains("ENER"))
                            {
                                myArray[6] = "Program";
                                EventDesc = "Energade";
                            }
                            else if ((myArray[5]).Contains("GMM") || (myArray[5]).Contains("GMW"))
                            {
                                myArray[6] = "Program";
                                EventDesc = "Gospel mix";
                            }
                            else if ((myArray[5]).Contains("KZNFILM"))
                            {
                                myArray[6] = "Spot";
                                EventDesc = "KZN Film commission";
                            }
                            else if ((myArray[5]).Contains("JBM") || (myArray[5].Trim('"')).Contains("JBT") || (myArray[5].Trim('"')).Contains("JBW"))
                            {
                                myArray[6] = "Program";
                                EventDesc = "my juke box";
                            }
                            else if ((myArray[5]).Contains("NEWS"))
                            {
                                myArray[6] = "Program";
                                EventDesc = "News";
                            }
                            else if ((myArray[5]).Contains("MAKT"))
                            {
                                myArray[6] = "Program";
                                EventDesc = "Makamu";
                            }
                            else if ((myArray[5]).Contains("ULM") || (myArray[5]).Contains("ULT"))
                            {
                                myArray[6] = "Program";
                                EventDesc = "Ugubhu Lwami";
                            }
                            else if ((myArray[5]).Contains("ZZ"))
                            {
                                myArray[6] = "Program";
                                EventDesc = "Ziphuma zishisa";
                            }
                            else if ((myArray[5]).Contains("TNB"))
                            {
                                myArray[6] = "Program";
                                EventDesc = "The next billionaire";
                            }
                            else if ((myArray[5]).Contains("MAQ"))
                            {
                                myArray[6] = "Spot";
                                EventDesc = "MAQ";
                            }
                            else if ((myArray[5]).Contains("WOZA"))
                            {
                                myArray[6] = "Program";
                                EventDesc = "Wozodlala";
                            }
                            else if ((myArray[5]).Contains("REV"))
                            {
                                myArray[6] = "Program";
                                EventDesc = "Revival Imvuselelo";
                            }
                            else if ((myArray[5]).Contains("ABK"))
                            {
                                myArray[6] = "Program";
                                EventDesc = "Abakhulume";
                            }
                            else if ((myArray[5]).Contains("ABV"))
                            {
                                myArray[6] = "Program";
                                EventDesc = "Abavumi";
                            }
                            else if ((myArray[5]).Contains("T1T"))
                            {
                                myArray[6] = "Program";
                                EventDesc = "Twenty 1";
                            }
                            else if ((myArray[5]).Contains("EBK"))
                            {
                                myArray[6] = "Program";
                                EventDesc = "Ebokhusini";
                            }
                            else if ((myArray[5]).Contains("KUNTH"))
                            {
                                myArray[6] = "Program";
                                EventDesc = "Kungengeka";
                            }
                            else if ((myArray[5]).Contains("STARSAT"))
                            {
                                myArray[6] = "Spot";
                                EventDesc = "Starsat promo";
                            }
                            else if ((myArray[5]).Contains("NCA"))
                            {
                                myArray[6] = "Spot";
                                EventDesc = "Church promo";
                            }
                            else if ((myArray[5]).Contains("OPGX"))
                            {
                                myArray[6] = "Spot";
                                EventDesc = "Opel";
                            }
                            else if ((myArray[5]).Contains("CFML") || (myArray[5].Trim('"')).Contains("CRLL"))
                            {
                                myArray[6] = "Spot";
                                EventDesc = "Clere";
                            }
                            else if ((myArray[5]).Contains("SPKK") || (myArray[5].Trim('"')).Contains("SPEK"))
                            {
                                myArray[6] = "Spot";
                                EventDesc = "Spekko";
                            }
                            else if ((myArray[5]).Contains("ZIK"))
                            {
                                myArray[6] = "Program";
                                EventDesc = "Zikhipani";
                            }
                            else if ((myArray[5]).Contains("DUMw"))
                            {
                                myArray[6] = "Program";
                                EventDesc = "Dumisa";
                            }
                            else if ((myArray[5]).Contains("MFC"))
                            {
                                myArray[6] = "Program";
                                EventDesc = "MFC fight zone";
                            }
                            else if ((myArray[5]).Contains("PBLY"))
                            {
                                myArray[6] = "Spot";
                                EventDesc = "Playboy";
                            }
                            else if ((myArray[5]).Contains("PGLY"))
                            {
                                myArray[6] = "Spot";
                                EventDesc = "Playgirl";
                            }
                            else if ((myArray[5]).Contains("ZR"))
                            {
                                myArray[6] = "Program";
                                EventDesc = "Zion reloaded";
                            }
                            else if ((myArray[5]).Contains("KHUL"))
                            {
                                myArray[6] = "Program";
                                EventDesc = "Khuluma sizwe";
                            }
                            else if ((myArray[5]).Contains("PEDI"))
                            {
                                myArray[6] = "Program";
                                EventDesc = "Pedigree";
                            }
                            else if ((myArray[5]).Contains("DJ MIX"))
                            {
                                myArray[6] = "Filler";
                                EventDesc = "DJ MIX";
                            }
                            else if ((myArray[5]).Contains("NWC"))
                            {
                                myArray[6] = "Spot";
                                EventDesc = "Ncwane Afrigospel";
                            }
                            else if ((myArray[5]).Contains("VOA"))
                            {
                                myArray[6] = "Program";
                                EventDesc = "Africa 54";
                            }
                            else if ((myArray[5]).Contains("SIY"))
                            {
                                myArray[6] = "Program";
                                EventDesc = "Siya pheka";
                            }
                            else if ((myArray[5]).Contains("MAS"))
                            {
                                myArray[6] = "Program";
                                EventDesc = "Masimbonge clap & tap";
                            }
                            else if ((myArray[5]).Contains("STA"))
                            {
                                myArray[6] = "Program";
                                EventDesc = "Straight talk Africa";
                            }
                        }

                        if (myArray[5].Contains("_"))
                        {
                            myArray[5] = myArray[5].Replace('_', '/');
                        }
                        myDate = myArray[0].Trim('"').Split('/');
                        myLogs.Rows.Add("1KZN", myDate[1] + '/' + myDate[0] + '/' + myDate[2], start_time, myArray[1].Trim('"'), Duration, myArray[6].Trim('"'), EventDesc, myArray[5].Trim('"'), "", "", "");

                        pbProgress.Increment(10);
                    }


                }
                dgActivityLogs.DataSource = myLogs;
            };
            pbProgress.Value = 4000;
            AcceptButton = btnExport;
        }
        public static bool IsNumeric(object Expression)
        {
            double retNum;

            bool isNum = Double.TryParse(Convert.ToString(Expression), System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out retNum);
            return isNum;
        }
    }
}
