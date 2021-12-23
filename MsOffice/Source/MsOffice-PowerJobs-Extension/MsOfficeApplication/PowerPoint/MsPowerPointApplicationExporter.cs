using powerJobs.Common.Applications;

namespace MsOfficeApplication.PowerPoint
{
    public class MsPowerPointApplicationExporter : DocumentExporterBase
    {
        public MsPowerPointApplicationExporter(MsPowerPointApplication msPowerPointApplication) : base(msPowerPointApplication)
        {
            ExportTypes.Add(typeof(MsPowerPointPdfExport));
        }
        
    }
}