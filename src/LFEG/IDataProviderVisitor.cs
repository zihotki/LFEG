using System;

namespace LFEG
{
    public interface IDataProviderVisitor
    {
        bool Visit<T>(ExcelColumn<T> column, Type dataType, Func<T, object> dataProvider);
    }
}