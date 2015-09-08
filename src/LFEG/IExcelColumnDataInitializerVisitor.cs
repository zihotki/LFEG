using System;

namespace LFEG
{
    public interface IExcelColumnDataInitializerVisitor
    {
        bool Visit(ExcelColumn column, Type dataType, Func<object, object> dataProvider);
    }
}