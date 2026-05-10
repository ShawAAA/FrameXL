using ExcelDna.Integration;
using ExcelDna.IntelliSense;

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
