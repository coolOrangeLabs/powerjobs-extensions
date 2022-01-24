using powerJobs.Common.Applications;

namespace coolOrange.MsOffice.Word
{
    public class MsWordApplicationExporter : DocumentExporterBase
    {
        public MsWordApplicationExporter(Application MsWordApplication) : base(MsWordApplication)
        {
            ExportTypes.Add(typeof(MsWordPdfExport)); //Add your export types which are inherited from IDocumentExport
        }
    }
}