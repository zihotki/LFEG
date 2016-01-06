using System;

namespace LFEG.Infrastructure.DataProviderVisitors
{
    public class BooleanDataProviderVisitor : IDataProviderVisitor
    {
        private readonly int? _defaultStyle;

        public BooleanDataProviderVisitor(int? defaultStyle)
        {
            _defaultStyle = defaultStyle;
        }

        public bool Visit<T>(ExcelColumn<T> column, Type dataType, Func<T, object> dataProvider)
        {
            if (!IsBooleanType(dataType))
            {
                return false;
            }

            if (column.Style.HasValue == false)
            {
                column.Style = _defaultStyle;
            }

            column.ExcelDataType = column.Style.HasValue ? ExcelDataTypes.Number : ExcelDataTypes.Boolean;

            column.DataProvider = i =>
            {
                var value = dataProvider(i);
                if (value == null || value == DBNull.Value)
                    return string.Empty;

                if ((bool) value)
                    return "1";

                return "0";
            };

            return true;
        }

        private static bool IsBooleanType(Type type)
        {
            if (type == null)
            {
                return false;
            }

            switch (Type.GetTypeCode(type))
            {
                case TypeCode.Boolean:
                    return true;
                case TypeCode.Object:
                    if (type.IsGenericType && type.GetGenericTypeDefinition() == typeof(Nullable<>))
                    {
                        return IsBooleanType(Nullable.GetUnderlyingType(type));
                    }
                    return false;
            }

            return false;
        }
    }
}