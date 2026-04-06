using ExcelDna.Integration;
using ExcelDna.IntelliSense;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TESTEXDNA
{
    public class IntelliSenseAddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            IntelliSenseServer.Install();
        }
        public void AutoClose()
        {
            IntelliSenseServer.Uninstall();
        }
    }
}
