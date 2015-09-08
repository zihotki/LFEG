using System;

namespace LFEG.ExcelColumnDataInitializerVisitors
{
    public class DefaultStringDataInitializerVisitor : IExcelColumnDataInitializerVisitor
    {
        public bool Visit(ExcelColumn column, Type dataType, Func<object, object> dataProvider)
        {
            column.ExcelDataType = ExcelDataTypes.String;

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