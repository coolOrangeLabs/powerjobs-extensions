using Microsoft.Office.Interop.Word;
using System;
using System.Runtime.InteropServices;
using powerJobs.Common.Applications;

namespace MsWordApplication
{
    public class MsWordApplicationDocument : DocumentBase
    {
        private Document _document;

        public MsWordApplicationDocument(IApplication application, OpenDocumentSettings openSettings)
            : base(application, openSettings)
        {
            //Do something with the OpenSettings like cache information and on Close_Internal(), close a connection
            Documents documents = (application as Application).Word.Documents;
            var fullName = openSettings.File.FullName;
            _document = documents.Open(fullName);
        }

        public Document Document { get => _document; set => _document = value; }

        protected override void Close_Internal(bool save = false)
        {
            if (_document == null)
                return;
            //close the opened file and save it depending on the argument
            Document doc = _document;
            object obj = (object)WdSaveOptions.wdDoNotSaveChanges;
            ref object local1 = ref obj;
            object missing1 = Type.Missing;
            ref object local2 = ref missing1;
            object missing2 = Type.Missing;
            ref object local3 = ref missing2;
            _document.Close(ref local1, ref local2, ref local3);
            Marshal.ReleaseComObject(_document);
            _document = null;
        }
    }
}
