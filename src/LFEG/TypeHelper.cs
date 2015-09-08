using System;
using System.ComponentModel;
using System.Data;
using System.Reflection;
using System.Text.RegularExpressions;

namespace LFEG
{
    public static class TypeHelper
    {
        public static bool IsEnum<T>()
        {
            return GetUnderlyingType<T>().IsEnum;
        }

        public static bool IsEnum(Type type)
        {
            return GetUnderlyingType(type).IsEnum;
        }

        public static bool IsEnum(Object obj)
        {
            return GetUnderlyingType(obj.GetType()).IsEnum;
        }

        public static Type GetUnderlyingType(Type type)
        {
            var underlyingType = Nullable.GetUnderlyingType(type);

            return underlyingType ?? type;
        }

        public static Type GetUnderlyingType<T>()
        {
            var realModelType = typeof(T);
            return GetUnderlyingType(realModelType);
        }

        public static string GetCaption(DataColumn column)
        {
            return string.IsNullOrEmpty(column.Caption)
                ? column.ColumnName.Replace("_", " ")
                : column.Caption.Replace("_", " ");
        }


        public static string GetCaption(PropertyInfo property)
        {
            /*var displayAttribute = property.GetCustomAttribute<DisplayAttribute>();
            if (displayAttribute != null)
            {
                return displayAttribute.Name;
            }*/

            var displayNameAttribute = property.GetCustomAttribute<DisplayNameAttribute>();
            if (displayNameAttribute != null)
            {
                return displayNameAttribute.DisplayName;
            }

            var descriptionAttribute = property.GetCustomAttribute<DescriptionAttribute>();
            if (descriptionAttribute != null)
            {
                return descriptionAttribute.Description;
            }

            return property.Name;
        }


        public static string GetValueCaption(Type enumType, object value)
        {
            if (value == null)
            {
                return string.Empty;
            }

            if (enumType == null)
            {
                return string.Empty;
            }

            enumType = Nullable.GetUnderlyingType(enumType) ?? enumType;

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