using System.Collections.Generic;
using powerJobs.Common.Applications;

namespace coolOrange.AutoCADElectrical.Exports
{
    public class DwfExport : AcadElectricalExportBase
    {
        public override HashSet<string> SupportedDocumentTypes => ((Application)SourceDocument.Application).SupportedFileTypes;

        public DwfExport(IDocument sourceDocument, ExportSettings settings)
            : base(sourceDocument, settings)
        {
        }

        public override string Name => "DWF";
    }
}
