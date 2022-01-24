using System.Collections.Generic;
using powerJobs.Common.Applications;

namespace coolOrange.MsOffice.Excel
{
    public class MsExcelPdfExport : DocumentExportBase
    {
        public MsExcelPdfExport(MsExcelApplicationDocument sourceDocument, ExportSettings settings) : base(
            sourceDocument, settings)
        {
            //read the ExportSettings and do something}
        } 

        public override void Execute()
        {
            //code which makes for instance the export and saves the file to settings.DestinationFile.FullName
            var wb = ((MsExcelApplicationDocument)SourceDocument).Workbook;
            var fullName = DestinationFile.FullName;
            wb.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, fullName);
        }

        public override string Name
        {
            get { return "Pdf"; } //Name which can be passed to Export-Document -Format 'Pdf'
        }

        public override HashSet<string> SupportedDocumentTypes
        {
            get
            {
                return ((Application)SourceDocument.Application).SupportedFileTypes;
            } //WATCH OUT: When using multiple Exports this could be dangerous, because then every Export supports ALL file formats from the Application and this could be a lot more then an individual export usually supports.
        }
    }
}