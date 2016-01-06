using System.Collections.Generic;

namespace LFEG.Infrastructure.Styling
{
    public class StyleProvider : IStyleProvider, IColumnStyler
    {
        public IList<NumberFormat> NumberFormats { get; } = new List<NumberFormat>();
        public IList<CellFormat> CellFormats { get; } = new List<CellFormat>();

        public StyleProvider()
        {
            CellFormats.Add(new CellFormat
            {
                FormatId = 0,
                Id = 0
            });

        }

        public int GetColumnStyle(string dataFormat)
        {
            var numberFormat = new NumberFormat
            {
                FormatCode = dataFormat,
                Id = 164 + NumberFormats.Count // custom number formats start from numFmtId=164
            };

            var cellFormat = new CellFormat
            {
                FormatId = numberFormat.Id,
                Id = CellFormats.Count
            };

            NumberFormats.Add(numberFormat);
            CellFormats.Add(cellFormat);

            return cellFormat.Id;
        }
    }
}