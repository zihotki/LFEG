using System;

namespace LFEG.Infrastructure.DataProviderVisitors
{
    public class DateTimeDataProviderVisitor : IDataProviderVisitor
    {
        private readonly int? _defaultStyle;

        public DateTimeDataProviderVisitor(int? defaultStyle)
        {
            _defaultStyle = defaultStyle;
        }

        public bool Visit<T>(ExcelColumn<T> column, Type dataType, Func<T, object> dataProvider)
        {
            if (!IsDateType(dataType))
                return false;

            column.ExcelDataType = ExcelDataTypes.Date;

            if (column.Style.HasValue == false)
            {
                column.Style = _defaultStyle;
            }

            column.DataProvider = i =>
            {
                var val = dataProvider(i);
                if (val == null || val == DBNull.Value)
                    return string.Empty;

                var dateVal = (DateTime) val;
                if (dateVal == DateTime.MinValue)
                    return string.Empty;

                return ((DateTime)val).ToOADate().ToString();
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