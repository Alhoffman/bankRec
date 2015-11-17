using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace Bank_Activity
{
    public class CitiBank
    {
        public static void LtcFormat(Excel.Worksheet oSht, Excel.Worksheet oSht2)
        {

            Excel.Range last = oSht.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);

            int j = 1;

            for (int i = 1; i <= last.Row; i++)
            {

                oSht2.Cells[j, 1] = oSht.Cells[i, 1];

                //int word = oSht.Cells[i, 9].Value);
                if (Convert.ToDouble(oSht.Cells[i, 3].Value) < 0.0)
                {
                    oSht2.Cells[j, 2] = "Withdrawal";
                }
                else
                {
                    oSht2.Cells[j, 2] = "Deposit";
                }


                // Check to see if this transaction is a check or another method
                string dType = Convert.ToString(oSht.Cells[i, 2].Value);

                if (dType.ToUpper().IndexOf("CHECK") > -1)
                {
                    oSht2.Cells[j, 3] = "Check";
                    oSht2.Cells[j, 4] = dType.Substring(6);
                }
                else
                {
                    oSht2.Cells[j, 3] = "ACH";
                    oSht2.Cells[j, 4] = "ACH";
                }

                //Payee will be empty 
                oSht2.Cells[j, 5] = "xxx";

                //Amount does not need to be changed just moved
                oSht2.Cells[j, 6] = oSht.Cells[i, 3];

                j++;
            }
        }
    }

    class CitizensBank
    {
        public static void LtcFormat(Excel.Worksheet oSht, Excel.Worksheet oSht2)
        {

            Excel.Range last = oSht.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);

            int j = 1;
            for (int i = 12; i <= last.Row; i++)
            {

                Excel.Range cell = oSht.Cells[i, 1];
                string numberFormat = cell.NumberFormat;

                if (numberFormat == "m/d/yyyy")
                {
                    oSht2.Cells[j, 1] = oSht.Cells[i, 1];

                    //int word = oSht.Cells[i, 9].Value);
                    if (Convert.ToDouble(oSht.Cells[i, 3].Value) > 0)
                    {
                        oSht2.Cells[j, 2] = "Withdrawal";
                    }
                    else
                    {
                        oSht2.Cells[j, 2] = "Deposit";
                    }


                    // Check to see if this transaction is a check or another method
                    string dType = Convert.ToString(oSht.Cells[i, 2].Value);

                    if (dType.ToUpper().IndexOf("CHECK") > -1)
                    {
                        oSht2.Cells[j, 3] = "Check";
                        oSht2.Cells[j, 4] = oSht.Cells[i, 5];
                    }
                    else
                    {
                        oSht2.Cells[i, 3] = "ACH";
                        oSht2.Cells[i, 4] = "ACH";
                    }


                    //Payee will be empty 
                    oSht2.Cells[j, 5] = "xxx";

                    //Amount does not need to be changed just moved
                    oSht2.Cells[j, 6] = oSht.Cells[i, 3];

                    //Increase j
                    j++;

                }
            }
        }
    }

    class CommunityandSouthern
    {
        public static void LtcFormat(Excel.Worksheet oSht, Excel.Worksheet oSht2)
        {

            Excel.Range last = oSht.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);

            int j = 1;

            for (int i = 2; i <= last.Row; i++)
            {

                oSht2.Cells[i - 1, 1] = oSht.Cells[i, 2];

                //This next part decides if the type is either Withdrawal or deposit
                //It also moves the amount over because they are the same logic.
                if (Convert.ToDouble(oSht.Cells[i, 5].Value) != 0)
                {
                    oSht2.Cells[j, 2] = "Withdrawal";
                    oSht2.Cells[j, 6] = oSht.Cells[i, 5];
                }
                else
                {
                    oSht2.Cells[j, 2] = "Deposit";
                    oSht2.Cells[j, 6] = oSht.Cells[i, 6];
                }


                // Check to see if this transaction is a check or another method
                string dType = Convert.ToString(oSht.Cells[i, 4].Value);

                if (dType.ToUpper().IndexOf("CHECK") > -1)
                {
                    oSht2.Cells[j, 3] = "Check";
                    oSht2.Cells[j, 4] = oSht.Cells[i, 8];
                }
                else
                {
                    oSht2.Cells[j, 3] = "ACH";
                    oSht2.Cells[j, 4] = "ACH";
                }

                //Payee will be empty 
                oSht2.Cells[j, 5] = "xxx";

                j++;
            }
        }
    }

    class PNCBank
    {
        public static void LtcFormat(Excel.Worksheet oSht, Excel.Worksheet oSht2)
        {

            Excel.Range last = oSht.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);

            int j = 1;

            for (int i = 2; i <= last.Row; i++)
            {

                oSht2.Cells[j, 1] = oSht.Cells[i, 1];

                //int word = oSht.Cells[i, 9].Value);
                if (Convert.ToString(oSht.Cells[i, 9].Value) == "DB")
                {
                    oSht2.Cells[j, 2] = "Withdrawal";
                }
                else
                {
                    oSht2.Cells[j, 2] = "Deposit";
                }


                // Check to see if this transaction is a check or another method
                string dType = Convert.ToString(oSht.Cells[i, 7].Value);
                string refNum = Convert.ToString(oSht.Cells[i, 10].Value);

                if (dType.ToUpper().IndexOf("CHECK") > -1)
                {
                    oSht2.Cells[j, 3] = "Check";
                    oSht2.Cells[j, 4] = refNum.Substring(1);
                }
                else
                {
                    oSht2.Cells[j, 3] = "ACH";
                    oSht2.Cells[j, 4] = "ACH";
                }

                //Payee will be empty 
                oSht2.Cells[j, 5] = "xxx";

                //Amount does not need to be changed just moved
                oSht2.Cells[j, 6] = oSht.Cells[i, 8];

                j++;
            }
        }
    }

    class PrivateBank
    {
        public static void LtcFormat(Excel.Worksheet oSht, Excel.Worksheet oSht2)
        {

            Excel.Range last = oSht.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);

            int j = 1;

            for (int i = 2; i <= last.Row; i++)
            {

                oSht2.Cells[j, 1] = oSht.Cells[i, 6];

                //int word = oSht.Cells[i, 9].Value);
                if (Convert.ToDouble(oSht.Cells[i, 8].Value) < 0.0)
                {
                    oSht2.Cells[j, 2] = "Withdrawal";
                }
                else
                {
                    oSht2.Cells[j, 2] = "Deposit";
                }


                // Check to see if this transaction is a check or another method
                string dType = Convert.ToString(oSht.Cells[i, 11].Value);

                if (dType.ToUpper().IndexOf("CHECK") > -1)
                {
                    oSht2.Cells[j, 3] = "Check";
                    oSht2.Cells[j, 4] = oSht.Cells[i, 4];
                }
                else
                {
                    oSht2.Cells[j, 3] = "ACH";
                    oSht2.Cells[j, 4] = "ACH";
                }

                //Payee will be empty 
                oSht2.Cells[j, 5] = "xxx";

                //Amount does not need to be changed just moved
                oSht2.Cells[j, 6] = oSht.Cells[i, 8];

                j++;
            }
        }
    }

    class PWB
    {
        public static void LtcFormat(Excel.Worksheet oSht, Excel.Worksheet oSht2)
        {

            Excel.Range last = oSht.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);

            int j = 1;

            for (int i = 2; i <= last.Row; i++)
            {

                oSht2.Cells[j, 1] = oSht.Cells[i, 1];

                //This next part decides if the type is either Withdrawal or deposit
                //It also moves the amount over because they are the same logic.
                if (Convert.ToDouble(oSht.Cells[i, 4].Value) > 0)
                {
                    oSht2.Cells[j, 2] = "Withdrawal";
                    oSht2.Cells[j, 6] = oSht.Cells[i, 4];
                }
                else
                {
                    oSht2.Cells[j, 2] = "Deposit";
                    oSht2.Cells[j, 6] = oSht.Cells[i, 5];
                }


                // Check to see if this transaction is a check or another method
                string dType = Convert.ToString(oSht.Cells[i, 3].Value);

                if (dType.ToUpper().IndexOf("CHECK") > -1)
                {
                    oSht2.Cells[j, 3] = "Check";
                    oSht2.Cells[j, 4] = oSht.Cells[i, 2];
                }
                else
                {
                    oSht2.Cells[j, 3] = "ACH";
                    oSht2.Cells[j, 4] = "ACH";
                }

                //Payee will be empty 
                oSht2.Cells[j, 5] = "xxx";

                j++;
            }
        }
    }

    public class TDBank
    {
        public static void LtcFormat(Excel.Worksheet oSht, Excel.Worksheet oSht2)
        {

            Excel.Range last = oSht.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);

            int j = 1;

            for (int i = 1; i <= last.Row; i++)
            {
                //Get the date and put in the new workbook column A
                oSht2.Cells[j, 1] = oSht.Cells[i, 1];

                //int word = oSht.Cells[i, 9].Value);
                if (Convert.ToDouble(oSht.Cells[i, 6].Value) != 0.0)
                {
                    oSht2.Cells[j, 2] = "Withdrawal";
                    oSht2.Cells[j, 6] = oSht.Cells[i, 6];
                }
                else
                {
                    oSht2.Cells[j, 2] = "Deposit";
                    oSht2.Cells[j, 6] = oSht.Cells[i, 7];
                }


                // Check to see if this transaction is a check or another method
                string dType = Convert.ToString(oSht.Cells[i, 5].Value);
                //int index = dType.IndexOfAny

                if (dType.IndexOf("CHECK") > -1)
                {
                    oSht2.Cells[j, 3] = "Check";
                    oSht2.Cells[j, 4] = oSht.Cells[i, 8];
                }
                else
                {
                    oSht2.Cells[j, 3] = "ACH";
                    oSht2.Cells[j, 4] = "ACH";
                }

                //Payee will be empty 
                oSht2.Cells[j, 5] = "xxx";

                j++;

            }
        }
    }

    class WellsFargo
    {

        public static void LtcFormat(Excel.Worksheet oSht, Excel.Worksheet oSht2)
        {

            Excel.Range last = oSht.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);

            int j = 1;

            for (int i = 1; i <= last.Row; i++)
            {

                oSht2.Cells[j, 1] = oSht.Cells[i, 1];

                //int word = oSht.Cells[i, 9].Value);
                if (Convert.ToDouble(oSht.Cells[i, 2].Value) < 0.0)
                {
                    oSht2.Cells[j, 2] = "Withdrawal";
                }
                else
                {
                    oSht2.Cells[j, 2] = "Deposit";
                }


                // Check to see if this transaction is a check or another method
                string dType = Convert.ToString(oSht.Cells[i, 5].Value);

                if (dType.ToUpper().IndexOf("CHECK") > -1)
                {
                    oSht2.Cells[j, 3] = "Check";
                    oSht2.Cells[j, 4] = oSht.Cells[i, 4];
                }
                else
                {
                    oSht2.Cells[j, 3] = "ACH";
                    oSht2.Cells[j, 4] = "ACH";
                }

                //Payee will be empty 
                oSht2.Cells[j, 5] = "xxx";

                //Amount does not need to be changed just moved
                oSht2.Cells[j, 6] = oSht.Cells[i, 2];

                j++;

            }
        }
    }
}
