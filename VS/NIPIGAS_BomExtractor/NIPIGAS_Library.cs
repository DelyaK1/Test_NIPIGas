using System.Activities;
using System.Data;
using System.ComponentModel;
using NIPIGAS_BomExtractorLogic;
using System.Collections.Generic;

namespace NIPIGAS_BomExtractor
{
    public class BomExtractor : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> PdfPath { get; set; }

        [Category("Output")]
        [RequiredArgument]
        public OutArgument<DataTable> BomDataTable { get; set; }

        [Category("Output")]
        [RequiredArgument]
        public OutArgument<Dictionary<string, string>> DrawingElements { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            var BOME = new BOM_Extractor
            {
                FilePath = PdfPath.Get(context)
            };
            BOME.ExtractRawData();
            BOME.ParseBomDataTable();
            BomDataTable.Set(context, BOME.BomDataTable);
            DrawingElements.Set(context, BOME.DrawingElements);
        }
    }
}
