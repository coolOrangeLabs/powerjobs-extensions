using MsWordApplication.Exports;
using powerJobs.Common.Applications;

namespace MsWordApplication
{
    public class MsWordApplicationExporter : DocumentExporterBase
    {
        public MsWordApplicationExporter(Application MsWordApplication) : base(MsWordApplication)
        {
            ExportTypes.Add(typeof(PdfExport)); //Add your export types which are inherited from IDocumentExport
        }
    }
}