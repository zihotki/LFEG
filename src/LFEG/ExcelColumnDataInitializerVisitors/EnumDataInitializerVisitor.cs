using System;

namespace LFEG.ExcelColumnDataInitializerVisitors
{
    public class EnumDataInitializerVisitor : IExcelColumnDataInitializerVisitor
    {
        public bool Visit(ExcelColumn column, Type dataType, Func<object, object> dataProvider)
        {
            if (!TypeHelper.IsEnum(dataType))
            {
                return false;
            }

            column.ExcelDataType = ExcelDataTypes.String;
                
            column.DataProvider = i =>
            {
                var value = dataProvider(i);

                if (value == null || value == DBNull.Value)
                    return string.Empty;

                return TypeHelper.GetValueCaption(dataType, value);
            };

            return true;
        }
    }
}