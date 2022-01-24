using System.Collections.Generic;
using powerJobs.Common.Applications;

namespace coolOrange.MsOffice.Word
{
    public class MsWordPdfExport : DocumentExportBase
    {
        public override HashSet<string> SupportedDocumentTypes
        {
            get { return ((Application)SourceDocument.Application).SupportedFileTypes; } //WATCH OUT: When using multiple Exports this could be dangerous, because then every Export supports ALL file formats from the Application and this could be a lot more then an individual export usually supports.
        }


        public MsWordPdfExport(MsWordApplicationDocument sourceDocument, ExportSettings settings)
            : base(sourceDocument, settings)
        {
            //read the ExportSettings and do something
        }

        public override void Execute()
        {
            //code which makes for instance the export and saves the file to settings.DestinationFile.FullName
            var doc = ((MsWordApplicationDocument)SourceDocument).Document;
            var fullName = DestinationFile.FullName;
            doc.ExportAsFixedFormat(fullName, Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF);
        }

        public override string Name
        {
            get { return "Pdf"; } //Name which can be passed to Export-Document -Format 'Pdf'
        }
    }
}