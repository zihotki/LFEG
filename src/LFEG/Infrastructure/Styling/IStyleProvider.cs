using System.Collections.Generic;

namespace LFEG.Infrastructure.Styling
{
    public interface IStyleProvider
    {
        IList<NumberFormat> NumberFormats { get; }
        IList<CellFormat> CellFormats { get; }
    }
}