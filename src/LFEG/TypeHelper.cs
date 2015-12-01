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
            var displayNameAttribute = (DisplayNameAttribute)Attribute.GetCustomAttribute(property, typeof(DisplayNameAttribute));
            if (displayNameAttribute != null)
            {
                return displayNameAttribute.DisplayName;
            }

            var descriptionAttribute = (DescriptionAttribute)Attribute.GetCustomAttribute(property, typeof(DescriptionAttribute));
            if (descriptionAttribute != null)
            {
                return descriptionAttribute.Description;
            }

            return property.Name;
        }
    }
}