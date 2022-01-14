﻿using System;
using System.IO;
using System.Reflection;
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
            AcadEProjectFilename = openSettings.File.FullName;
        }

        protected override void Close_Internal(bool save = false)
        {
            try
            {
                Log.Info("Deactivate current project by activating dummy project ...");
                var acadActiveDocument = ((Application)Application).AcadApplication.GetActiveDocument();

                var wdpFilename = Environment.ExpandEnvironmentVariables(Properties.Settings.Default.DummyAcadElectricalProjectFile);
                if (!File.Exists(wdpFilename))
                {
                    Log.Error($"Dummy project '{wdpFilename}' doesn't exist!");
                    return;
                }

                wdpFilename = Environment.ExpandEnvironmentVariables(wdpFilename).Replace('\\', '/');
                Log.Info($"Activating project {wdpFilename}");
                acadActiveDocument.SendCommandWait($"(c:wd_makeproj_current \"{wdpFilename}\"){System.Environment.NewLine}");

                AcadEProjectFilename = null;
                ((Application)Application).AcadApplication.WaitUntilReady(Properties.Settings.Default.MdbCreationWaitTime);
                Log.Info("Successfully deactivated current project.");
            }
            catch (Exception ex)
            {
                Log.Error($"Error deactivating current project: {ex.Message}", ex);
            }
        }
	}
}