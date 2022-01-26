using System;
using System.Reflection;
using System.Threading;
using log4net;

namespace coolOrange.AutoCADElectrical.Helpers
{
    public static class AcadDocHelper
    {
        static readonly ILog Log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);

        public static void SendCommandWait(dynamic acadDocument, string cmd)
        {
            Log.Debug($"Sending command '{cmd}' ... (Timeout {Properties.Settings.Default.AcadCmdTimeout}ms) ");
            var duration = 0;
            var sentCmd = false;
            while (duration < Properties.Settings.Default.AcadCmdTimeout)
            {
                try
                {
                    acadDocument.SendCommand(cmd);
                    Log.Debug("Successfully send command");
                    sentCmd = true;
                    break;
                }
                catch (Exception ex)
                {
                    Log.Warn($"SendCommandWait(): AutoCAD not ready yet: {ex.Message}. Waiting ...");
                }
                duration += 1000;
                Thread.Sleep(1000);
            }
            if (!sentCmd)
                throw new ApplicationException($"Failed to send command '{cmd}' (Timeout)!");
        }
    }
}
