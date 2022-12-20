using System.Collections.Generic;
using powerJobs.Common.Applications;
using SolidWorks.Interop.sldworks;

namespace coolOrange.Solidworks.Exports
{
    public class PdfExport : DocumentExportBase
    {
        public override string Name
        {
            get { return "Pdf"; } //Name which can be passed to Export-Document -Format 'Pdf'
        }

        public override HashSet<string> SupportedDocumentTypes
        {
            get { return ((Application)SourceDocument.Application).SupportedFileTypes; } //WATCH OUT: When using multiple Exports this could be dangerous, because then every Export supports ALL file formats from the Application and this could be a lot more then an individual export usually supports.
        }

        public PdfExport(Document sourceDocument, ExportSettings settings)
            : base(sourceDocument, settings)
        {
            //read the ExportSettings
        }

        public override void Execute()
        {
            //export the document to PDF and save the outcome to settings.DestinationFile.FullName
            ModelDoc2 doc = ((Document)SourceDocument).ModelDoc2;
            var fullName = DestinationFile.FullName;
            doc.SaveAs(fullName);
        }
    }
}