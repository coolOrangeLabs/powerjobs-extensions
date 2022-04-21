using System;
using System.Reflection;
using log4net;

namespace coolOrange.AutoCADElectrical.Helpers
{
    public static class AcadDocHelper
    {
        static readonly ILog Log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);
        
        public static void SendCommandWait(dynamic acadDocument, string cmd)
        {
            Log.Debug($"Sending command '{cmd}'");
            var sentCmd = false;
            try
            {
                acadDocument.SendCommand(cmd);
                Log.Debug("Successfully send command");
                sentCmd = true;
            }
            catch (Exception ex)
            {
                Log.Warn($"SendCommandWait(): {ex.Message}.");
            }
            if (!sentCmd)
                throw new ApplicationException($"Failed to send command '{cmd}' (Timeout)!");
        }
    }
}
