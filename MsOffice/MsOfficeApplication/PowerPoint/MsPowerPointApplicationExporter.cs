using powerJobs.Common.Applications;

namespace coolOrange.MsOffice.PowerPoint
{
    public class MsPowerPointApplicationExporter : DocumentExporterBase
    {
        public MsPowerPointApplicationExporter(Application msPowerPointApplication) : base(msPowerPointApplication)
        {
            ExportTypes.Add(typeof(MsPowerPointPdfExport));
        }
        
    }
}