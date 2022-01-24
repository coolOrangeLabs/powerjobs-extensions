using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using powerJobs.Common.Applications;

namespace coolOrange.MsOffice.Excel
{
    public class MsExcelApplicationDocument : DocumentBase
    {
        private Workbook _workbook;

        public MsExcelApplicationDocument(IApplication application, OpenDocumentSettings openSettings) : base(application, openSettings)
        {
            //Do something with the OpenSettings like cache information and on Close_Internal(), close a connection
            Workbooks workbooks = (application as Application)?.Excel.Workbooks;
            var fullName = openSettings.File.FullName;
            _workbook = workbooks.Open(fullName);
        }


        public Workbook Workbook { get => _workbook; set => _workbook = value; }

        protected override void Close_Internal(bool save = false)
        {
            if (_workbook == null)
                return;
            //close the opened file and save it depending on the argument
            Workbook wb = _workbook;
            _workbook.Close(false, "", false);
            Marshal.ReleaseComObject(_workbook);
            _workbook = null;
        }
    }
}
