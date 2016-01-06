using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text.RegularExpressions;

namespace LFEG.Infrastructure.DataProviderVisitors
{
    public class EnumDataProviderVisitor : IDataProviderVisitor
    {
        private readonly Dictionary<object, string> _enumValues = new Dictionary<object, string>();

        public bool Visit<T>(ExcelColumn<T> column, Type dataType, Func<T, object> dataProvider)
        {
            if (!TypeHelper.IsEnum(dataType))
            {
                return false;
            }

            column.ExcelDataType = ExcelDataTypes.String;

            var enumType = TypeHelper.GetUnderlyingType(dataType);
            var values = Enum.GetValues(enumType);

            foreach (var value in values)
            {
                _enumValues.Add(value, GetCaption(enumType, value));
            }

            column.DataProvider = i =>
            {
                var value = dataProvider(i);

                if (value == null || value == DBNull.Value || _enumValues.ContainsKey(value) == false)
                    return string.Empty;

                return _enumValues[value];
            };

            return true;
        }

        private static string GetCaption(Type enumType, object value)
        {
            var name = Enum.GetName(enumType, value);
            if (string.IsNullOrEmpty(name))
            {
                return string.Empty;
            }

            var fi = enumType.GetField(name);

            var descriptionAttributes = (DescriptionAttribute[])fi.GetCustomAttributes(typeof(DescriptionAttribute), false);
            if (descriptionAttributes.Length > 0)
            {
                return descriptionAttributes[0].Description;
            }

            name = Regex.Replace(name, "([a-z](?=[A-Z])|[A-Z](?=[A-Z][a-z]))", "$1 ").TrimEnd();
            return name;
        }
    }
}