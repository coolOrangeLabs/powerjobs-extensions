using powerJobs.Common.Applications;
using Solidworks.Exports;

namespace Solidworks
{
    public class Exporter : DocumentExporterBase
    {
        public Exporter(Application Solidworks) : base(Solidworks)
        {
            ExportTypes.Add(typeof(PdfExport)); //Register all your export types (those must implement IDocumentExport)
        }
    }
}