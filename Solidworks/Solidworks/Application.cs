using powerJobs.Common.Applications;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using log4net;
using log4net.Config;
using System.Reflection;
using SolidWorks.Interop.sldworks;


namespace Solidworks
{
    public class Application : ApplicationBase
    {
        static readonly ILog Log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);
        public SldWorks _solidworks;
        public Application()
        {
            Exporter = new Exporter(this);
        }

        public override string Name
        {
            get { return "SolidWorksVersion2022"; }
        }

        public override HashSet<string> SupportedFileTypes
        {
            get
            {
                return new HashSet<string>()
            {
                ".slddrw"
            };
            }//speficy which file formats are supported from this application, e.g. ".exe", ".ico", ".iso", ".png"
        }

        public override bool IsRunning
        {
            get { return _solidworks != null; } //return the information whether the application is running or not. powerJobs Processor detects if the application has to be restarted
        }

        public override void Start()
        {
            _solidworks = (SldWorks)Activator.CreateInstance(System.Type.GetTypeFromProgID("SldWorks.Application.30"));
            // initialize and start the application preferably in silent mode
        }
        public SldWorks SolidWorks { get => _solidworks; set => _solidworks = value; }
        protected override bool IsInstalled_Internal()
        {
            return true; //check if the application is installed, for instance within the registry
        }

        protected override IDocument OpenDocument_Internal(OpenDocumentSettings openSettings)
        {
            //open the file with the passed settings
            return new Document(this, openSettings);
        }

        protected override void Stop_Internal()
        {
            if (IsRunning)
            {
                //Close the application and dispose all its resources
                _solidworks.ExitApp();
                Marshal.FinalReleaseComObject(_solidworks);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                _solidworks = null;
            }
        }
    }
}