using System;

namespace LFEG.Infrastructure.DataProviderVisitors
{
    public interface IDataProviderVisitor
    {
        bool Visit<T>(ExcelColumn<T> column, Type dataType, Func<T, object> dataProvider);
    }
}