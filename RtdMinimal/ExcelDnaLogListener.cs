using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace RtdMinimal
{
    public class ExcelDnaLogListener : TraceListener
    {
        public ExcelDnaLogListener()
        {
            Trace.WriteLine($"--------------- {nameof(ExcelDnaLogListener)}");
        }

        public ExcelDnaLogListener(string name)
            : base(name)
        {
            Trace.WriteLine($"--------------- {nameof(ExcelDnaLogListener)}");
        }

        public override void Write(string message)
        {
            Trace.WriteLine($"{Name} {message}");
        }

        public override void WriteLine(string message)
        {
            Trace.WriteLine($"{Name} {message}");
        }
    }
}
