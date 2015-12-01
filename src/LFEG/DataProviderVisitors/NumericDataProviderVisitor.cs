using System;

namespace LFEG.ExcelColumnDataInitializerVisitors
{
    public class NumericDataProviderVisitor : IDataProviderVisitor
    {
        public bool Visit<T>(ExcelColumn<T> column, Type dataType, Func<T, object> dataProvider)
        {
            if (!IsNumericType(dataType))
            {
                return false;
            }

            column.ExcelDataType = ExcelDataTypes.Number;

            column.DataProvider = i =>
            {
                var val = dataProvider(i);
                if (val == null || val == DBNull.Value)
                    return string.Empty;

                return val.ToString();
            };

            return true;
        }

        private static bool IsNumericType(Type type)
        {
            if (type == null)
            {
                return false;
            }

            switch (Type.GetTypeCode(type))
            {
                case TypeCode.Byte:
                case TypeCode.Decimal:
                case TypeCode.Double:
                case TypeCode.Int16:
                case TypeCode.Int32:
                case TypeCode.Int64:
                case TypeCode.SByte:
                case TypeCode.Single:
                case TypeCode.UInt16:
                case TypeCode.UInt32:
                case TypeCode.UInt64:
                    return true;
                case TypeCode.Object:
                    if (type.IsGenericType && type.GetGenericTypeDefinition() == typeof(Nullable<>))
                    {
                        return IsNumericType(Nullable.GetUnderlyingType(type));
                    }
                    return false;
            }
            return false;
        }
    }
}