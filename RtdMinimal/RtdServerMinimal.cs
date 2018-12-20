using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

using ExcelDna.Integration;
using ExcelDna.Integration.Rtd;

namespace RtdMinimal
{
    [ComVisible(true)]
    [ProgId(Const.RtdServerProgId), Guid(Const.RtdServerGuid)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class RtdServerMinimal : ExcelRtdServer
    {
        public RtdServerMinimal()
        {
            Trace.WriteLine(">>> Exception");
            throw new ApplicationException("TEST_EXCEPTION");
        }

        protected override bool ServerStart()
        {
            Trace.WriteLine("\t--- RtdServerMinimal.ServerStart()");

            return base.ServerStart();
        }

        protected override object ConnectData(Topic topic, IList<string> topicInfo, ref bool newValues)
        {
            Trace.WriteLine("\t--- RtdServerMinimal.ConnectData()");

            base.ConnectData(topic, topicInfo, ref newValues);

            topic.UpdateValue("TEST_" + topic.TopicId);

            if (newValues)
                return topic.Value;
            else
                return ExcelError.ExcelErrorNA;
        }
    }
}
