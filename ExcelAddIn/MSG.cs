using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace ExcelAddIn
{
    public partial class MSG
    {
        private void MSG_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnTexterize_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Texterize();
        }

        private void btnUnTexterize_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.UnTexterize();
        }
    }
}
