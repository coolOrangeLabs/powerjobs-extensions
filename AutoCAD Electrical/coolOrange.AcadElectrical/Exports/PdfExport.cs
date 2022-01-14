using System.Collections.Generic;
using powerJobs.Common.Applications;

namespace coolOrange.AutoCADElectrical.Exports
{
    public class PdfExport : AcadElectricalExportBase
    {
        public override HashSet<string> SupportedDocumentTypes => ((Application)SourceDocument.Application).SupportedFileTypes;

        public PdfExport(AcadElectricalDocument sourceDocument, ExportSettings settings)
            : base(sourceDocument, settings)
        {
        }

        public override string Name => "PDF";
    }
}