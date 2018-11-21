using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

using ExcelDna.ComInterop;
using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;

namespace RtdMinimal
{
    public class ThisAddin : IExcelAddIn
    {
        public void AutoOpen()
        {
            ComServer.DllRegisterServer();

            var application = ExcelUtil.Application;
            application.WorkbookOpen += ApplicationOnWorkbookOpen;
        }

        private void ApplicationOnWorkbookOpen(Excel.Workbook wb)
        {
            ReportRam();
        }

        public void AutoClose()
        {
        }

        public static void ReportRam()
        {
            var process = Process.GetCurrentProcess();
            var ramUsedTotal = process.PrivateMemorySize64;
            var ramUsedManaged = GC.GetTotalMemory(true);

            Trace.WriteLine($"\t\t>>>\t{ramUsedManaged:n0} {ramUsedTotal:n0}");
        }
    }
}
