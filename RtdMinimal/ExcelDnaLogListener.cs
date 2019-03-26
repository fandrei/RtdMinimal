using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace RtdMinimal
{
    public class ExcelDnaLogListener : TraceListener
    {
        public ExcelDnaLogListener()
        {
            Trace.WriteLine($"--------------- {nameof(ExcelDnaLogListener)}");
            InitFile();
        }

        public ExcelDnaLogListener(string name)
            : base(name)
        {
            Trace.WriteLine($"--------------- {nameof(ExcelDnaLogListener)}");
            InitFile();
        }

        void InitFile()
        {
            lock (Sync)
            {
                if (_writer == null)
                    _writer = new StreamWriter($"Log_{DateTime.Now:yyyy-MM-dd HH_mm_ss.fffffff}.txt");
            }
        }

        public override void Write(string message)
        {
            Trace.Write($"{Name} {message}");
            lock (Sync)
            {
                _writer.Write($"{Name} {message}");
                _writer.Flush();
            }
        }

        public override void WriteLine(string message)
        {
            Trace.WriteLine($"{Name} {message}");
            lock (Sync)
            {
                _writer.WriteLine($"{Name} {message}");
                _writer.Flush();
            }
        }

        private static StreamWriter _writer;
        static readonly object Sync = new object();
    }
}
