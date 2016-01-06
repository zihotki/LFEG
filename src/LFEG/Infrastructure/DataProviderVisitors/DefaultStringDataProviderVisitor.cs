using System;

namespace LFEG.Infrastructure.DataProviderVisitors
{
    public class DefaultStringDataProviderVisitor : IDataProviderVisitor
    {
        public bool Visit<T>(ExcelColumn<T> column, Type dataType, Func<T, object> dataProvider)
        {
            column.ExcelDataType = ExcelDataTypes.InlineString;

            column.DataProvider = i =>
            {
                var val = dataProvider(i);
                if (val == null || val == DBNull.Value)
                    return string.Empty;

                return val.ToString();
            };

            return true;
        }
    }
}