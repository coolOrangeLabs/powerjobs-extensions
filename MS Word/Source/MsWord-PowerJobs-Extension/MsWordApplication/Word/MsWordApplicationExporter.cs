using powerJobs.Common.Applications;

namespace MsOfficeApplication.Word
{
    public class MsWordApplicationExporter : DocumentExporterBase
    {
        public MsWordApplicationExporter(MsWordApplication MsWordApplication) : base(MsWordApplication)
        {
            ExportTypes.Add(typeof(MsWordPdfExport)); //Add your export types which are inherited from IDocumentExport
        }
    }
}