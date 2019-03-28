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
                {
                    _writer = AppDomain.CurrentDomain.GetData(LogDataName) as StreamWriter;
                    if (_writer != null)
                        return;

                    _writer = new StreamWriter($"Log_{DateTime.Now:yyyy-MM-dd HH_mm_ss.fffffff}.txt");
                    AppDomain.CurrentDomain.SetData(LogDataName, _writer);

                    WriteFile("Init log");
                }
            }
        }

        public override void Write(string message)
        {
            Trace.WriteLine($"{Name} {message}");
            WriteFile(message);
        }

        public override void WriteLine(string message)
        {
            Trace.WriteLine($"{Name} {message}");
            WriteFile(message);
        }

        private static void WriteFile(string message)
        {
            lock (Sync)
            {
                _writer.Write(AppDomain.CurrentDomain.Id + " " + message + "\r\n");
                _writer.Flush();
            }
        }

        static readonly object Sync = new object();
        private static StreamWriter _writer;
        private const string LogDataName = "ExcelDnaLogListener";
    }
}
