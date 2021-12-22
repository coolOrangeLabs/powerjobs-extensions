using powerJobs.Common.Applications;

namespace MsOfficeApplication.Excel
{
    public class MsExcelApplicationExporter : DocumentExporterBase
    {
        public MsExcelApplicationExporter(MsExcelApplication msExcelApplication) : base(msExcelApplication)
        {
            ExportTypes.Add(typeof(MsExcelPdfExport)); //Add your export types which are inherited from IDocumentExport

        }
    }
}