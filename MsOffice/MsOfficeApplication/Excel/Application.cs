using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using powerJobs.Common.Applications;

namespace coolOrange.MsOffice.Excel
{
    public class Application : ApplicationBase
    {
        private Microsoft.Office.Interop.Excel.Application _excel;

        public Application()
        {
            Exporter = new MsExcelApplicationExporter(this);
        }

        public override string Name
        {
            get { return "MsExcelApplication"; }
        }

        public override bool IsRunning
        {
            get { return _excel != null; } //return if app is running. powerJobs Processor detects if application has to be restarted
        }

        public override HashSet<string> SupportedFileTypes
        {
            get
            {
                return new HashSet<string>() {
                    ".xls",
                    ".xlsx",
                    ".xlsm"
                };
            } //add all supported files like ".exe", ".ico", "iso", "png"
        }

        public Microsoft.Office.Interop.Excel.Application Excel { get => _excel; set => _excel = value; }

        public override void Start()
        {
            // start the application and prepare it
            _excel = (Microsoft.Office.Interop.Excel.Application)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("00024500-0000-0000-c000-000000000046")));
            //_excel.Visible = true;
        }

        protected override bool IsInstalled_Internal()
        {
            return true; //for instance check in the registry if application is installed
        }

        protected override IDocument OpenDocument_Internal(OpenDocumentSettings openSettings)
        {
            //open in the application the file with the passed settings
            return new MsExcelApplicationDocument(this, openSettings);
        }

        protected override void Stop_Internal()
        {
            if (!IsRunning)
                return;
            else
            {
                //Close the application if it is running and dispose all resources
                Microsoft.Office.Interop.Excel.Application excel = _excel;
                _excel.Quit();
                Marshal.FinalReleaseComObject(_excel);
                GC.Collect();
                GC.WaitForPendingFinalizers();

                _excel = null;
            }
        }
    }
}
