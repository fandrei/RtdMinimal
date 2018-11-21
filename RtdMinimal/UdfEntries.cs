using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;

using ExcelDna.Integration;

namespace RtdMinimal
{
    public static class UdfEntries
    {
        [ExcelFunction(Category = "RtdMinimalAddin", IsVolatile = false, IsThreadSafe = false, IsMacroType = true, IsExceptionSafe = true)]
        public static object RtdMinimal1()
        {
            Trace.WriteLine("\t--- =RtdMinimal1()");

            var location = ExcelUtil.GetCallingRange();
            var res = XlCall.Excel(XlCall.xlfRtd, Const.RtdServerProgId, Const.RtdServerGuid, "TEST", location.GetAddressText());
            return res;
        }
    }
}
