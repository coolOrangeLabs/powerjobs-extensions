using System.Collections.Generic;
using powerJobs.Common.Applications;

namespace MsOfficeApplication.PowerPoint
{
    public class MsPowerPointPdfExport : DocumentExportBase
    {
        public MsPowerPointPdfExport(MsPowerPointApplicationDocument sourceDocument, ExportSettings settings) : base(sourceDocument, settings)
        {
        }

        public override void Execute()
        {
            //code which makes for instance the export and saves the file to settings.DestinationFile.FullName
            var pp = ((MsPowerPointApplicationDocument)SourceDocument).Presentation;
            var fullName = DestinationFile.FullName;
            pp.ExportAsFixedFormat(fullName, Microsoft.Office.Interop.PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF);
        }

        public override string Name
        {
            get { return "Pdf"; } //Name which can be passed to Export-Document -Format 'Pdf'
        }
        public override HashSet<string> SupportedDocumentTypes
        {
            get { return ((MsPowerPointApplication)SourceDocument.Application).SupportedFileTypes; } //WATCH OUT: When using multiple Exports this could be dangerous, because then every Export supports ALL file formats from the Application and this could be a lot more then an individual export usually supports.
        }

    }
}