using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        public void Texterize()
        {
            var ExcelApp = this.Application;
            Excel.Range selectedRng = ExcelApp.Selection;

            if (selectedRng.Areas.Count > 1)
            {
                System.Windows.Forms.MessageBox.Show("Not allow multiple selection, Please select 1 area at once", "Texterize");
                return;
            }

            int lastRow = selectedRng.Find("*", LookIn: XlFindLookIn.xlFormulas, SearchOrder: XlSearchOrder.xlByRows, SearchDirection: XlSearchDirection.xlPrevious).Row;
            int lastCol = selectedRng.Find("*", LookIn: XlFindLookIn.xlFormulas, SearchOrder: XlSearchOrder.xlByColumns, SearchDirection: XlSearchDirection.xlPrevious).Column;

            int tRow = selectedRng.Rows.Count;
            int tCol = selectedRng.Columns.Count;

            if (tRow == 1 && tCol == 1)
            {
                selectedRng.Value = "'" + selectedRng.Value2;
            }
            else
            {
                if (lastRow < tRow) tRow = lastRow;
                if (lastCol < tCol) tCol = lastCol;
                object[,] arr = new object[tRow, tCol];
                arr = selectedRng.Value2;


                for (int nRow = 1; nRow <= tRow; nRow++)
                {
                    for (int nCol = 1; nCol <= tCol; nCol++)
                    {
                        if (arr[nRow, nCol] != null)
                            arr[nRow, nCol] = "'" + arr[nRow, nCol];
                    }
                }

                selectedRng.Value = arr;
            }

            System.Windows.Forms.MessageBox.Show("Done", "Texterize");
        }

        public void UnTexterize()
        {
            double number;

            var ExcelApp = this.Application;
            Excel.Range selectedRng = ExcelApp.Selection;

            if (selectedRng.Areas.Count > 1)
            {
                System.Windows.Forms.MessageBox.Show("Not allow multiple selection, Please select 1 area at once", "UnTexterize");
                return;
            }

            int lastRow = selectedRng.Find("*", LookIn: XlFindLookIn.xlFormulas, SearchOrder: XlSearchOrder.xlByRows, SearchDirection: XlSearchDirection.xlPrevious).Row;
            int lastCol = selectedRng.Find("*", LookIn: XlFindLookIn.xlFormulas, SearchOrder: XlSearchOrder.xlByColumns, SearchDirection: XlSearchDirection.xlPrevious).Column;

            int tRow = selectedRng.Rows.Count;
            int tCol = selectedRng.Columns.Count;

            if (tRow == 1 && tCol == 1)
            {
                if (Double.TryParse(selectedRng.Value2.ToString(), out number))
                {
                    selectedRng.Value = number;
                }
                else
                {
                    selectedRng.Value = selectedRng.Value2;
                }
            }
            else
            {
                if (lastRow < tRow) tRow = lastRow;
                if (lastCol < tCol) tCol = lastCol;

                object[,] arr = new object[tRow, tCol];
                arr = selectedRng.Value;

                for (int nRow = 1; nRow <= tRow; nRow++)
                {
                    for (int nCol = 1; nCol <= tCol; nCol++)
                    {
                        if (arr[nRow, nCol] != null)
                        {
                            if (Double.TryParse(arr[nRow, nCol].ToString(), out number))
                            {
                                arr[nRow, nCol] = number;
                            }
                        }

                    }
                }

                selectedRng.Value = arr;
            }

            System.Windows.Forms.MessageBox.Show("Done", "Texterize");
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
