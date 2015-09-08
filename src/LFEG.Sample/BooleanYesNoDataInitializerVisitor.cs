using System;

namespace LFEG.Sample
{
    public class BooleanYesNoDataInitializerVisitor : IExcelColumnDataInitializerVisitor
    {
        public bool Visit(ExcelColumn column, Type dataType, Func<object, object> dataProvider)
        {
            if (!IsBooleanType(dataType))
            {
                return false;
            }

            column.ExcelDataType = ExcelDataTypes.String;

            column.DataProvider = i =>
            {
                var value = dataProvider(i);
                if (value == null || value == DBNull.Value)
                    return string.Empty;

                if ((bool) value)
                    return "Yes";

                return "No";
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