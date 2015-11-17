using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace Bank_Activity
{
    class MainProgram
    {
        public static void Start(string filePath, string savePath, string bank)
        {
            Excel.Application oXL;
            Excel.Workbook oWB;
            Excel.Workbook oWB2;
            Excel.Worksheet oSht;
            Excel.Worksheet oSht2;

            oXL = new Excel.Application();
            oXL.Visible = false;

            oWB = (Excel.Workbook)(oXL.Workbooks.Open(filePath));
            oSht = (Excel.Worksheet)oWB.Sheets[1];

            oWB2 = (Excel.Workbook)(oXL.Workbooks.Add());
            oSht2 = (Excel.Worksheet)oWB2.Sheets[1];


            try
            {
                switch (bank)
                {
                    case "TD Bank":
                        TDBank.LtcFormat(oSht, oSht2);
                        break;

                    case "Citi Bank":
                        CitiBank.LtcFormat(oSht, oSht2);
                        break;

                    case "Wells Fargo Bank":
                        WellsFargo.LtcFormat(oSht, oSht2);
                        break;

                    case "Private Bank":
                        PrivateBank.LtcFormat(oSht, oSht2);
                        break;

                    case "PNC Bank":
                        PNCBank.LtcFormat(oSht, oSht2);
                        break;

                    case "Pacific Western Bank":
                        PWB.LtcFormat(oSht, oSht2);
                        break;

                    case "Community & Southern":
                        CommunityandSouthern.LtcFormat(oSht, oSht2);
                        break;

                    case "Citizens Bank":
                        CitizensBank.LtcFormat(oSht, oSht2);
                        break;
                }


                //Now that the format is considered ltc grade, let's turn into an intacct grade
                FormatIntacct(oSht2);

                //Let's now close the original without saving, and saveas the new workbook as a xlsx file.
                oWB2.SaveAs(savePath);
                oWB.Close(false);
                oWB2.Close();
                oXL.Quit();

            }
            catch
            {
                MessageBox.Show("There was an error running the program for " + bank + " please make sure the bank has not changed there template.");
                oWB.Close(false);
                oWB2.Close(false);
                oXL.Quit();
            }
        }

        static void FormatIntacct(Excel.Worksheet oSht2)
        {
            Excel.Range line = oSht2.Columns[1];
            line.NumberFormat = "m/d/yyyy";

            for (int i = 2; i <= 6; i++)
            {
                line = oSht2.Columns[i];
                line.NumberFormat = "General";
            }

        }
    }
}
