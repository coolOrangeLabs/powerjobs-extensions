using powerJobs.Common.Applications;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Runtime.InteropServices;

namespace Solidworks
{
    public class Document : DocumentBase
    {
        private ModelDoc2 _document;
        private SldWorks _swApp;
        public Document(IApplication application, OpenDocumentSettings openSettings)
            : base(application, openSettings)
        {
            _swApp = (application as Application).SolidWorks;
            var fullName = openSettings.File.FullName;
            _document = _swApp.OpenDoc6(fullName, (int)swDocumentTypes_e.swDocDRAWING, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", 0, 0);
            //OpenSettings can be used
        }
        public ModelDoc2 ModelDoc2 { get => _document; set => _document = value; }

        public SldWorks SolidWorks { get => _swApp; set => _swApp = value; }
        protected override void Close_Internal(bool save = false)
        {
            //close the opened file and save it depending on the argument
            if (_document == null) return;
            //close the opened file and save it depending on the argument
            _swApp.CloseDoc(_document.GetTitle());
            //_document.Close();
            Marshal.ReleaseComObject(_document);
            _document = null;
        }
    }
}
