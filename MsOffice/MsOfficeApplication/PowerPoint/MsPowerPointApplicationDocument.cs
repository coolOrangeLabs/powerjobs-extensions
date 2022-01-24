using System.Runtime.InteropServices;
using Microsoft.Office.Interop.PowerPoint;
using powerJobs.Common.Applications;

namespace coolOrange.MsOffice.PowerPoint
{
    public class MsPowerPointApplicationDocument : DocumentBase
    {
        private Presentation _presentation;
        public MsPowerPointApplicationDocument(IApplication application, OpenDocumentSettings openSettings) : base(application, openSettings)
        {
            //Do something with the OpenSettings like cache information and on Close_Internal(), close a connection
            Presentations presentations = (application as Application)?.PowerPoint.Presentations;
            var fullName = openSettings.File.FullName;
            _presentation = presentations.Open(fullName, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
        }

        public Presentation Presentation { get => _presentation; set => _presentation = value; }


        protected override void Close_Internal(bool save = false)
        {
            if (_presentation == null)
                return;
            //close the opened file and save it depending on the argument
            Presentation pp = _presentation;
            _presentation.Close();
            Marshal.ReleaseComObject(_presentation);
            _presentation = null;
        }
    }
}