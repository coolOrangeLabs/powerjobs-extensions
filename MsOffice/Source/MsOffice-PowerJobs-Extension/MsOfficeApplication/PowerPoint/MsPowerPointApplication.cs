using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using powerJobs.Common.Applications;

namespace MsOfficeApplication.PowerPoint
{
    public class MsPowerPointApplication : ApplicationBase
    {
        private Microsoft.Office.Interop.PowerPoint.Application _powerpoint;

        public MsPowerPointApplication()
        {
            Exporter = new MsPowerPointApplicationExporter(this);
        }
        public override void Start()
        {
            // start the application and prepare it
            _powerpoint = (Microsoft.Office.Interop.PowerPoint.Application)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("91493441-5A91-11CF-8700-00AA0060263B")));
            //_powerpoint.Visible = true;
        }

        public override string Name { get { return "MsPowerPointApplication"; } }

        public override bool IsRunning
        {
            get
            {
                return _powerpoint != null;
            } //return if app is running. powerJobs Processor detects if application has to be restarted
        }

        public override HashSet<string> SupportedFileTypes
        {
            get
            {
                return new HashSet<string>() {
                    ".ppt",
                    ".pptx",
                    ".pptm"
                };
            } //add all supported files like ".exe", ".ico", "iso", "png"
        }

        public Microsoft.Office.Interop.PowerPoint.Application PowerPoint { get => _powerpoint; set => _powerpoint = value; }
        protected override bool IsInstalled_Internal()
        {
           return true;
        }

        protected override IDocument OpenDocument_Internal(OpenDocumentSettings openSettings)
        {
            //open in the application the file with the passed settings
            return new MsPowerPointApplicationDocument(this, openSettings);
        }

        protected override void Stop_Internal()
        {
            if (!IsRunning)
                return;
            else
            {
                //Close the application if it is running and dispose all resources
                Microsoft.Office.Interop.PowerPoint.Application powerpoint = _powerpoint;
                _powerpoint.Quit();
                Marshal.FinalReleaseComObject(_powerpoint);
                GC.Collect();
                GC.WaitForPendingFinalizers();

                _powerpoint = null;
            }
        }
    }
}
