using System;
using System.Reflection;
using System.Threading;
using Autodesk.AutoCAD.Interop;
using log4net;

namespace coolOrange.AutoCADElectrical.Helpers
{
    public static class AcadAppHelper
    {
        static readonly ILog Log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);

        public static void WaitUntilReady(this AcadApplication acadApplication, int additionalWaitTime = 0)
        {
            Log.Debug($"Wait for AutoCAD to be ready for {Properties.Settings.Default.AcadCmdTimeout}ms");
            var duration = 0;
            while (duration < Properties.Settings.Default.AcadCmdTimeout)
            {
                try
                {
                    var acadState = acadApplication.GetAcadState();
                    if (acadState.IsQuiescent) break;
                }
                catch (Exception ex)
                {
                    Log.Warn($"WaitUntilReady(): AutoCAD not ready yet: {ex.Message}. Waiting ...");
                }
                duration += 1000;
                Thread.Sleep(1000);
            }
            if (additionalWaitTime > 0)
                Thread.Sleep(additionalWaitTime); // wait for power... splash
            Log.Debug("AutoCAD ready!");
        }

        public static AcadDocument GetActiveDocument(this AcadApplication acadApplication)
        {
            Log.Debug($"Getting active document ... (Timeout {Properties.Settings.Default.AcadCmdTimeout}ms) ");
            var duration = 0;
            AcadDocument activeDocument = null;
            while (duration < Properties.Settings.Default.AcadCmdTimeout)
            {
                try
                {
                    activeDocument = acadApplication.ActiveDocument;
                    if (activeDocument != null) break;
                }
                catch (Exception ex)
                {
                    Log.Warn($"GetActiveDocument(): AutoCAD not ready yet: {ex.Message}. Waiting ...");
                }
                duration += 1000;
                Thread.Sleep(1000);
            }
            if (activeDocument == null)
                throw new ApplicationException($"Failed to get active document (Timeout)!");
            Log.Debug("Successfully got active document");
            return activeDocument;
        }

        public static string GetPrinterConfigPath(this AcadApplication acadApplication)
        {
            Log.Debug($"Getting printer config path from preferences ... (Timeout {Properties.Settings.Default.AcadCmdTimeout}ms) ");
            var duration = 0;
            string printerConfigPath = null;
            while (duration < Properties.Settings.Default.AcadCmdTimeout)
            {
                try
                {
                    printerConfigPath = acadApplication.Preferences.Files.PrinterConfigPath;
                    if (printerConfigPath != null) break;
                }
                catch (Exception ex)
                {
                    Log.Warn($"GetPrinterConfigPath(): AutoCAD not ready yet: {ex.Message}. Waiting ...");
                }
                duration += 1000;
                Thread.Sleep(1000);
            }
            if (printerConfigPath == null)
                throw new ApplicationException($"Failed to get printer config path (Timeout)!");
            Log.Debug("Successfully got printer config path");
            return printerConfigPath;
        }
    }
}
