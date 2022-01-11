using System.Collections.Generic;
using powerJobs.Common.Applications;

namespace coolOrange.AutoCADElectrical.Exports
{
    class DwfxExport : AcadElectricalExportBase
    {
        public override HashSet<string> SupportedDocumentTypes => ((Application)SourceDocument.Application).SupportedFileTypes;

        public DwfxExport(IDocument sourceDocument, ExportSettings settings)
            : base(sourceDocument, settings)
        {
        }

        public override string Name => "DWFx";
    }
}
