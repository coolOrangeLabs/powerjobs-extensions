using coolOrange.AutoCADElectrical.Exports;
using powerJobs.Common.Applications;

namespace coolOrange.AutoCADElectrical
{
    public class AcadElectricalExporter : DocumentExporterBase
	{
		public AcadElectricalExporter(Application acadElectricalApplication) : base(acadElectricalApplication)
		{
            ExportTypes.Add(typeof(PdfExport));
            ExportTypes.Add(typeof(DwfExport));
            ExportTypes.Add(typeof(DwfxExport));
        }
    }
}