using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;

namespace RtdMinimal
{
    public static class ExcelUtil
    {
        public static Excel.Application Application
        {
            get
            {
                var res = (Excel.Application)ExcelDnaUtil.Application;
                return res;
            }
        }

        public static ExcelReference GetCallingRange()
        {
            var res = (ExcelReference)XlCall.Excel(XlCall.xlfCaller);
            return res;
        }

        public static string GetAddressText(this ExcelReference range)
        {
            return range.GetSheetName() + "!" + range.GetLocalAddressText();
        }

        public static string GetLocalAddressText(this ExcelReference range)
        {
            var firstCell = (string)XlCall.Excel(XlCall.xlfAddress, 1 + range.RowFirst, 1 + range.ColumnFirst);
            var lastCell = (string)XlCall.Excel(XlCall.xlfAddress, 1 + range.RowLast, 1 + range.ColumnLast);

            var res = firstCell;
            if (firstCell != lastCell)
                res += ":" + lastCell;
            return res;
        }

        public static string GetSheetName(this ExcelReference range)
        {
            var res = (string)XlCall.Excel(XlCall.xlSheetNm, range);
            return res;
        }
    }
}
