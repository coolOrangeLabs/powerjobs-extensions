using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Autodesk.AutoCAD.Interop;
using coolOrange.AutoCADElectrical.Helpers;
using log4net;
using Microsoft.Win32;
using powerJobs.Common.Applications;

namespace coolOrange.AutoCADElectrical
{
    public class Application : ApplicationBase
    {
        static readonly ILog Log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);

        public AcadApplication AcadApplication { get; set; }

        private string AcadProgId { get; set; }
        private string AcadExePath { get; set; }

        public Application()
        {
            var logConfigFile = new FileInfo(@"C:\Program Files\coolOrange\powerJobs Processor\powerJobs Processor.log4net");
            if (!logConfigFile.Exists)
                logConfigFile = new FileInfo(@"C:\Program Files\coolOrange\powerJobs\powerJobs.log4net");
            if (logConfigFile.Exists)
                log4net.Config.XmlConfigurator.Configure(logConfigFile);

            Exporter = new AcadElectricalExporter(this);
        }

        public override string Name => "AutoCAD Electrical";

        public override bool IsRunning
        {
            get
            {
                if (AcadApplication == null)
                    return false;

                try
                {
                    return !string.IsNullOrEmpty(AcadApplication.Caption);
                }
                catch (Exception ex)
                {
                    Log.Error($"AutoCAD Electrical Application doesn't exist anymore: {ex.Message}", ex);
                    AcadApplication = null;
                    return false;
                }
            }
        }

        public override HashSet<string> SupportedFileTypes => new HashSet<string>() { ".wdp" };

        public override void Start()
        {
            // start the application and prepare it
            try
            {
                Log.Info("Starting AutoCAD Electrical ...");
                var startParams = Properties.Settings.Default.AcadStartParameters;
                var psi = new ProcessStartInfo(AcadExePath, startParams)
                {
                    WorkingDirectory = @"C:\temp"
                };
                var acadProcess = Process.Start(psi);
                acadProcess.WaitForInputIdle();

                // TODO - use ROT to make sure using correct AutoCAD instance in case there are more than one (https://adndevblog.typepad.com/autocad/2013/12/accessing-com-applications-from-the-running-object-table.html)
                while (AcadApplication == null)
                {
                    try
                    {
                        AcadApplication = (AcadApplication)Marshal.GetActiveObject(AcadProgId);
                    }
                    catch
                    {
                        System.Windows.Forms.Application.DoEvents();
                    }
                }
                AcadApplication.WaitUntilReady(Properties.Settings.Default.StartWaitTime);
                Log.Info("Successfully started AutoCAD Electrical!");
            }
            catch (Exception ex)
            {
                Log.Error($"Failed to start AutoCAD Electrical: {ex.Message}", ex);
                throw;
            }
        }

        protected override bool IsInstalled_Internal()
        {
            try
            {
                if (!string.IsNullOrEmpty(AcadExePath))
                    return true;

                Log.Info($"Checking if AutoCAD Electrical is installed ...");
                var curVer = Registry.GetValue(@"HKEY_CURRENT_USER\SOFTWARE\Autodesk\AutoCAD", "CurVer", null);
                if (curVer != null)
                {
                    var curNum = curVer.ToString().Substring(1,2);
                    AcadProgId = $"AutoCAD.Application.{curNum}";
                    using (var acadProductsKey = Registry.LocalMachine.OpenSubKey($@"SOFTWARE\Autodesk\AutoCAD\{curVer}"))
                    {
                        foreach (var acedProdKeyName in acadProductsKey.GetSubKeyNames())
                        {
                            using (var acedProdKey = acadProductsKey.OpenSubKey(acedProdKeyName))
                            {
                                var productName = acedProdKey.GetValue("ProductNameGlob");
                                if (productName != null && productName.ToString().StartsWith("AutoCAD Electrical"))
                                {
                                    Log.Info($"AutoCAD Electrical found: {acedProdKey.GetValue("ProductNameGlob")}");
                                    AcadExePath = $"{acedProdKey.GetValue("AcadLocation")}\\acad.exe";
                                }
                                acedProdKey.Close();
                            }
                        }
                        acadProductsKey.Close();
                    }
                }
                if (string.IsNullOrEmpty(AcadExePath))
                    Log.Error("No AutoCAD Electrical found on this machine!");
                return !string.IsNullOrEmpty(AcadExePath);
            }
            catch (Exception ex)
            {
                Log.Error($"Failed to check if AutoCAD Electrical is installed: {ex.Message}", ex);
                return false;
            }
        }

        protected override IDocument OpenDocument_Internal(OpenDocumentSettings openSettings)
        {
            //open in the application the file with the passed settings
            try
            {
                if (AcadApplication == null)
                    throw new ApplicationException("No AutoCAD Electrical application available");

                if (!IsSupportedFile(openSettings.File))
                    throw new ApplicationException($"Files with extension {openSettings.File.Extension} are not supported!");

                AcadApplication.WaitUntilReady(Properties.Settings.Default.OpenWaitTime);
                if (AcadApplication.Documents.Count == 0)
                {
                    Log.Debug("Adding empty document, to have at least on document to send commands ...");
                    AcadApplication.Documents.Add();
                }
                return new AcadElectricalDocument(this, openSettings);
            }
            catch (Exception ex)
            {
                Log.Error($"Failed to open file: {ex.Message}", ex);
                throw;
            }
        }

        protected override void Stop_Internal()
        {
            if (IsRunning)
            {
                try
                {
                    Log.Info("Closing AutoCAD Electrical Application ...");
                    if (AcadApplication.ActiveDocument != null)
                        AcadApplication.ActiveDocument.Close();
                    AcadApplication.Quit();
                    AcadApplication = null;
                    Log.Info("Successfully closed AutoCAD Electrical Application!");
                }
                catch (Exception ex)
                {
                    Log.Error($"Failed to close AutoCAD Electrical Application: {ex.Message}", ex);
                }
            }
        }
    }
}