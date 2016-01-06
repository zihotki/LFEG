using System;
using System.ComponentModel;
using System.Data;
using System.Reflection;

namespace LFEG.Infrastructure
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

        public static bool IsEnum(object obj)
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
            var exportAttribute = property.GetCustomAttribute<ExcelExportAttribute>();
            if (!string.IsNullOrEmpty(exportAttribute?.Caption))
            {
                return exportAttribute.Caption;
            }

            var displayNameAttribute = (DisplayNameAttribute)Attribute.GetCustomAttribute(property, typeof(DisplayNameAttribute));
            if (!string.IsNullOrEmpty(displayNameAttribute?.DisplayName))
            {
                return displayNameAttribute.DisplayName;
            }

            var descriptionAttribute = (DescriptionAttribute)Attribute.GetCustomAttribute(property, typeof(DescriptionAttribute));
            if (!string.IsNullOrEmpty(descriptionAttribute?.Description))
            {
                return descriptionAttribute.Description;
            }

            return property.Name;
        }
    }
}