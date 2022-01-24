using powerJobs.Common.Applications;

namespace coolOrange.MsOffice.Excel
{
    public class MsExcelApplicationExporter : DocumentExporterBase
    {
        public MsExcelApplicationExporter(Application msExcelApplication) : base(msExcelApplication)
        {
            ExportTypes.Add(typeof(MsExcelPdfExport)); //Add your export types which are inherited from IDocumentExport

        }
    }
}