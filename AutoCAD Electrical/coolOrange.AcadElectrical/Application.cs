using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using log4net;
using Microsoft.Win32;
using powerJobs.Common.Applications;

namespace coolOrange.AutoCADElectrical
{
    #region IMessageFilter

    // For more information on IMessageFilter:
    // http://msdn.microsoft.com/en-us/library/ms693740(VS.85).aspx

    [ComImport,
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown),
     Guid("00000016-0000-0000-C000-000000000046")]
    public interface IMessageFilter
    {
        [PreserveSig]
        int HandleInComingCall(
            int dwCallType, IntPtr hTaskCaller,
            int dwTickCount, IntPtr lpInterfaceInfo
        );
        [PreserveSig]
        int RetryRejectedCall(
            IntPtr hTaskCallee, int dwTickCount, int dwRejectType
        );
        [PreserveSig]
        int MessagePending(
            IntPtr hTaskCallee, int dwTickCount, int dwPendingType
        );
    }

    #endregion
    public class Application : ApplicationBase, IMessageFilter
    {
        static readonly ILog Log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);

        internal dynamic AcadApplication { get; set; }

        private string AcadProgId { get; set; }
        private string AcadExePath { get; set; }

        [DllImport("ole32.dll")]
        static extern int CoRegisterMessageFilter(IMessageFilter lpMessageFilter, out IMessageFilter lplpMessageFilter);

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

        public void RegisterMessageFilter()
        {
            IMessageFilter oldFilter = null;
            var result = CoRegisterMessageFilter(this, out oldFilter);
            if (result != 0)
                throw new Exception("Registration of MessageFilter failed, approve the registration is made in a STA Thread");
        }

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
            bool finished = false;
            Proxy.Instance.Collection.Add(() =>
            {
                // start the application and prepare it
                try
                {
                    //Log.Info("Starting AutoCAD Electrical ...");
                    //var startParams = Properties.Settings.Default.AcadStartParameters;
                    //var psi = new ProcessStartInfo(AcadExePath, startParams)
                    //{
                    //    WorkingDirectory = @"C:\temp"
                    //};
                    //var acadProcess = Process.Start(psi);
                    //acadProcess.WaitForInputIdle();

                    // TODO - use ROT to make sure using correct AutoCAD instance in case there are more than one (https://adndevblog.typepad.com/autocad/2013/12/accessing-com-applications-from-the-running-object-table.html)
                    try
                    {
                        AcadApplication = Marshal.GetActiveObject(AcadProgId);
                        RegisterMessageFilter();
                        Console.WriteLine("Get object of type \"" + AcadProgId + "\"");
                    }
                    catch
                    {

                        try
                        {
                            Type acType = Type.GetTypeFromProgID(AcadProgId);
                            AcadApplication = Activator.CreateInstance(acType, true);
                            RegisterMessageFilter();
                            Console.WriteLine("create object of type \"" + AcadProgId + "\"");
                        }
                        catch
                        {
                            Log.Error("Cannot create object of type \"" + AcadProgId + "\"");
                        }
                    }

                    Log.Info("Successfully started AutoCAD Electrical!");
                }
                catch (Exception ex)
                {
                    Log.Error($"Failed to start AutoCAD Electrical: {ex.Message}", ex);
                    throw;
                }
                finally
                {
                    finished = true;
                }
            });
            Proxy.Instance.Run();
            while (!finished)
            {
                Thread.Sleep(1000);
            }
        }

        protected override bool IsInstalled_Internal()
        {
            bool result;
            try
            {
                if (!string.IsNullOrEmpty(AcadExePath))
                   result = true;

                Log.Info($"Checking if AutoCAD Electrical is installed ...");
                var curVer = Registry.GetValue(@"HKEY_CURRENT_USER\SOFTWARE\Autodesk\AutoCAD", "CurVer", null);
                if (curVer != null)
                {
                    var curNum = curVer.ToString().Substring(1, 2);
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
                                    break;
                                }
                                acedProdKey.Close();
                            }
                        }
                        acadProductsKey.Close();
                    }
                }
                if (string.IsNullOrEmpty(AcadExePath))
                    Log.Error("No AutoCAD Electrical found on this machine!");
                result =  !string.IsNullOrEmpty(AcadExePath);
            }
            catch (Exception ex)
            {
                Log.Error($"Failed to check if AutoCAD Electrical is installed: {ex.Message}", ex);
                result = false;
            }
            return result;
        }
        protected override IDocument OpenDocument_Internal(OpenDocumentSettings openSettings)
        {

            if (!IsSupportedFile(openSettings.File))
                throw new ApplicationException($"Files with extension {openSettings.File.Extension} are not supported!");
            return new AcadElectricalDocument(this, openSettings);
        }
        
        protected override void Stop_Internal()
        {
            Proxy.Instance.Collection.Add(() =>
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
                    finally
                    {
                        IMessageFilter oldFilter = null;
                        CoRegisterMessageFilter(null, out oldFilter);
                    }
                }
            });
            Proxy.Instance.Run();
            while (Proxy.Instance.Collection.Count == 0)
            {
                Thread.Sleep(1000);
            }
        }

        #region IMessageFilter


        int IMessageFilter.HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo)
        {
            return 0; // SERVERCALL_ISHANDLED
        }

        int IMessageFilter.RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType)
        {
            return 99;
        }

        int IMessageFilter.MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType)
        {
            return 1; // PENDINGMSG_WAITNOPROCESS
        }

        #endregion
    }
}