using System;
using System.Reflection;
using log4net;

namespace coolOrange.AutoCADElectrical.Helpers
{
    public static class AcadAppHelper
    {
        static readonly ILog Log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);
        
        public static dynamic GetActiveDocument(dynamic acadApplication)
        {
            Log.Debug($"Getting active document) ");
            dynamic activeDocument = null;
            try
            {
                activeDocument = acadApplication.ActiveDocument;
            }
            catch (Exception ex)
            {
                Log.Warn($"GetActiveDocument(): {ex.Message}.");
            }
            if (activeDocument == null)
                throw new ApplicationException($"Failed to get active document!");
            Log.Debug("Successfully got active document");
            return activeDocument;
        }
        public static string GetPrinterConfigPath(dynamic acadApplication)
        {
            Log.Debug($"Getting printer config path from preferences ... ");
            string printerConfigPath = null;
            try
            {
                printerConfigPath = acadApplication.Preferences.Files.PrinterConfigPath;
            }
            catch (Exception ex)
            {
                Log.Warn($"GetPrinterConfigPath(): {ex.Message}.");
            }
            if (printerConfigPath == null)
                throw new ApplicationException($"Failed to get printer config path!");
            Log.Debug("Successfully got printer config path");
            return printerConfigPath;
        }
    }
}
