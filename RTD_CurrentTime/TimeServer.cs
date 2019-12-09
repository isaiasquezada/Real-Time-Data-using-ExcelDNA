using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using ExcelDna.Integration;
using ExcelDna.Integration.Rtd;
using ExcelDna.IntelliSense;
using System.Runtime.InteropServices; // Used to set GUID, ProgramID attributes

namespace RTD_CurrentTime
{
    [Guid("33eefd8a-98cb-410a-82ff-1180d8591c91")]
 
    public class TimeServer : ExcelRtdServer
    {
        private readonly List<Topic> topics = new List<Topic>();
        private Timer timer;

        public TimeServer()
        {
            timer = new Timer(Callback);
        }

        private void Start()
        {
            timer.Change(500, 500);
        }

        private void Stop()
        {
            timer.Change(-1, -1);
        }

        private void Callback(object o)
        {
            Stop();
            foreach (Topic topic in topics)
                topic.UpdateValue(GetTime());
            Start();
        }

        private static string GetTime()
        {
            return DateTime.Now.ToString("HH:mm:ss.fff");
        }

        protected override void ServerTerminate()
        {
            timer.Dispose();
            timer = null;
        }

        protected override object ConnectData(Topic topic, IList<string> topicInfo, ref bool newValues)
        {
            topics.Add(topic);
            Start();
            return GetTime();
        }

        protected override void DisconnectData(Topic topic)
        {
            topics.Remove(topic);
            if (topics.Count == 0)
                Stop();
        }
    }

    public static class Timefunctions
    {
        public static object GetCurrentTime()
        {
           return XlCall.RTD("RTD_CurrentTime.TimeServer", null, "HOY");
        }
    }
}
