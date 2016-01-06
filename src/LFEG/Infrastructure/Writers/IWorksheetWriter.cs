using System.Collections.Generic;
using Ionic.Zip;

namespace LFEG.Infrastructure.Writers
{
    public interface IWorksheetWriter
    {
        void Write<T>(ZipFile zip, IEnumerable<T> items, ExcelColumn<T>[] columns);
    }
}