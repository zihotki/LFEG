using System.Collections.Generic;
using System.Data;

namespace LFEG.Infrastructure
{
    public static class DataReaderEnumerator
    {
        public static IEnumerable<IDataReader> Enumerate(this IDataReader reader)
        {
            while (reader.Read())
            {
                yield return reader;
            }
        }
    }
}