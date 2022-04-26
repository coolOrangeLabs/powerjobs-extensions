using System;
using System.IO;
using System.Reflection;
using System.Threading;
using coolOrange.AutoCADElectrical.Helpers;
using log4net;
using powerJobs.Common.Applications;

namespace coolOrange.AutoCADElectrical
{
    public class AcadElectricalDocument : DocumentBase
    {
        static readonly ILog Log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);

        public string AcadEProjectFilename { get; set; }

        public AcadElectricalDocument(IApplication application, OpenDocumentSettings openSettings)
            : base(application, openSettings)
        {
            var finished = false;
            Proxy.Instance.Collection.Add(() =>
            {
                //open in the application the file with the passed settings
                try
                {
                    if (application == null)
                        throw new ApplicationException("No AutoCAD Electrical application available");

                    dynamic acad = ((Application) application).AcadApplication;
                    dynamic documents = acad.Documents;
                    if (documents.Count == 0)
                    {
                        Log.Debug("Adding empty document, to have at least on document to send commands ...");
                        documents.Add();
                    }

                    AcadEProjectFilename = openSettings.File.FullName;
                }
                catch (Exception ex)
                {
                    Log.Error($"Failed to open file: {ex.Message}", ex);
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


        
        protected override void Close_Internal(bool save = false)
        {
            var finished = false;
            Proxy.Instance.Collection.Add(() =>
            {
                try
                {
                    Log.Info("Deactivate current project by activating dummy project ...");
                    dynamic acad = ((Application) Application).AcadApplication;
                    dynamic acadActiveDocument = acad.ActiveDocument;

                    var indexOf = acad.Caption.IndexOf(" 2");
                    var version = acad.Caption.Substring(indexOf + 1, 4);
                    var wdpFilename =
                        string.Format(
                            Environment.ExpandEnvironmentVariables(Properties.Settings.Default
                                .DummyAcadElectricalProjectFile), version);
                    if (!File.Exists(wdpFilename))
                    {
                        Log.Error($"Dummy project '{wdpFilename}' doesn't exist!");
                        return;
                    }

                    wdpFilename = Environment.ExpandEnvironmentVariables(wdpFilename).Replace('\\', '/');
                    Log.Info($"Activating project {wdpFilename}");
                    AcadDocHelper.SendCommandWait(acadActiveDocument,
                        $"(c:wd_makeproj_current \"{wdpFilename}\"){System.Environment.NewLine}");

                    AcadEProjectFilename = null;
                    Log.Info("Successfully deactivated current project.");
                }
                catch (Exception ex)
                {
                    Log.Error($"Error deactivating current project: {ex.Message}", ex);
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
    }
}
