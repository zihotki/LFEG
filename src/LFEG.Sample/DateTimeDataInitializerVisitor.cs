using System;

namespace LFEG.Sample
{
    public class DateTimeDataInitializerVisitor : IExcelColumnDataInitializerVisitor
    {
        public bool Visit(ExcelColumn column, Type dataType, Func<object, object> dataProvider)
        {
            if (!IsDateType(dataType))
                return false;

            column.ExcelDataType = ExcelDataTypes.String;
            column.DataProvider = i =>
            {
                var val = dataProvider(i);
                if (val == null || val == DBNull.Value)
                    return string.Empty;

                var dateVal = (DateTime) val;
                if (dateVal == DateTime.MinValue)
                    return string.Empty;

                return ((DateTime)val).ToString("dd-MM-yyyy");
            };

            return true;
        }

        private static bool IsDateType(Type type)
        {
            if (type == null)
            {
                return false;
            }

            switch (Type.GetTypeCode(type))
            {
                case TypeCode.DateTime:
                    return true;
                case TypeCode.Object:
                    if (type.IsGenericType && type.GetGenericTypeDefinition() == typeof(Nullable<>))
                    {
                        return IsDateType(Nullable.GetUnderlyingType(type));
                    }
                    return false;
            }

            return false;
        }
    }
}