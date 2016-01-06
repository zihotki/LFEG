using System;
using System.Data;
using System.Linq.Expressions;
using System.Reflection;
using LFEG.Infrastructure.DataProviderVisitors;
using LFEG.Infrastructure.Styling;

namespace LFEG.Infrastructure
{
    public class ExcelColumnFactory
    {
        protected IDataProviderVisitor[] DataProviderVisitors;
        private readonly IColumnStyler _styler;

        // we need new visitor instances for each new 
        protected internal ExcelColumnFactory(IDataProviderVisitor[] visitors, IColumnStyler styler)
        {
            DataProviderVisitors = visitors;
            _styler = styler;
        }

        public virtual ExcelColumn<T> CreatePropertyColumn<T>(PropertyInfo propertyInfo)
        {
            var column = new ExcelColumn<T>
            {
                Caption = TypeHelper.GetCaption(propertyInfo)
            };

            var instance = Expression.Parameter(propertyInfo.DeclaringType, "i");
            var property = Expression.Property(instance, propertyInfo);
            var convert = Expression.TypeAs(property, typeof(object));

            var propertyAccessor = (Func<T, object>)Expression.Lambda(convert, instance).Compile();
            
            InitializeByDataType(column, propertyInfo.PropertyType, propertyAccessor);
            InitializeByAttribute(column, propertyInfo);

            return column;
        }

        
        public virtual ExcelColumn<IDataReader> CreateDataReaderColumn(DataColumn dataColumn, int index)
        {
            var column = new ExcelColumn<IDataReader>
            {
                Caption = TypeHelper.GetCaption(dataColumn),
            };

            InitializeByDataType(column, dataColumn.DataType, r => r[index]);

            return column;
        }

        public virtual ExcelColumn<DataRow> CreateDataTableColumn(DataColumn dataColumn, int index)
        {
            var column = new ExcelColumn<DataRow>
            {
                Caption = TypeHelper.GetCaption(dataColumn),
            };

            InitializeByDataType(column, dataColumn.DataType, r => r[index]);

            return column;
        }

        protected virtual void InitializeByDataType<T>(ExcelColumn<T> column, Type dataType, Func<T, object> dataProvider)
        {
            foreach (var visitor in DataProviderVisitors)
            {
                if (visitor.Visit(column, dataType, dataProvider))
                    return;
            }
        }

        protected virtual void InitializeByAttribute<T>(ExcelColumn<T> column, PropertyInfo propertyInfo)
        {
            var attr = propertyInfo.GetCustomAttribute<ExcelExportAttribute>();
            if (attr == null)
            {
                return;
            }

            column.Width = attr.Width;
            
            if (column.ExcelDataType == ExcelDataTypes.InlineString && attr.Intern)
            {
                column.ExcelDataType = ExcelDataTypes.String;
            }

            if (!string.IsNullOrEmpty(attr.DataFormat))
            {
                column.Style = _styler.GetColumnStyle(attr.DataFormat);

                if (column.ExcelDataType == ExcelDataTypes.Boolean)
                {
                    // number formatting can't be applied to boolean columns
                    column.ExcelDataType = ExcelDataTypes.Number;
                }
            }
        }
    }
}