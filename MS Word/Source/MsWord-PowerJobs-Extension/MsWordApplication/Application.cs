using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using powerJobs.Common.Applications;

namespace MsWordApplication
{
    public class Application : ApplicationBase
    {
        private Microsoft.Office.Interop.Word.Application _word;

        public Application()
        {
            Exporter = new MsWordApplicationExporter(this);
        }

        public override string Name
        {
            get { return "MsWordApplication"; }
        }

        public override bool IsRunning
        {
            get { return _word != null; } //return if app is running. powerJobs Processor detects if application has to be restarted
        }

        public override HashSet<string> SupportedFileTypes
        {
            get
            {
                return new HashSet<string>() {
                    ".doc",
                    ".docx",
                    ".docm"
                };
            } //add all supported files like ".exe", ".ico", "iso", "png"
        }

        public Microsoft.Office.Interop.Word.Application Word { get => _word; set => _word = value; }

        public override void Start()
        {
            // start the application and prepare it
            _word = (Microsoft.Office.Interop.Word.Application)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("000209FF-0000-0000-C000-000000000046")));
            //_word.Visible = true;
        }

        protected override bool IsInstalled_Internal()
        {
            return true; //for instance check in the registry if application is installed
        }

        protected override IDocument OpenDocument_Internal(OpenDocumentSettings openSettings)
        {
            //open in the application the file with the passed settings
            return new MsWordApplicationDocument(this, openSettings);
        }

        protected override void Stop_Internal()
        {
            if (!IsRunning)
                return;
            else
            {
                //Close the application if it is running and dispose all resources
                Microsoft.Office.Interop.Word.Application word = _word;
                object missing1 = Type.Missing;
                ref object local1 = ref missing1;
                object missing2 = Type.Missing;
                ref object local2 = ref missing2;
                object missing3 = Type.Missing;
                ref object local3 = ref missing3;
                _word.Quit(ref local1, ref local2, ref local3);
                Marshal.FinalReleaseComObject(_word);
                GC.Collect();
                GC.WaitForPendingFinalizers();

                _word = null;
            }
        }
    }
}