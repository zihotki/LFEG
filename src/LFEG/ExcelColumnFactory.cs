using System;
using System.Data;
using System.Linq.Expressions;
using System.Reflection;

namespace LFEG
{
    public class ExcelColumnFactory
    {
        protected IDataProviderVisitor[] DataProviderVisitors;

        // we need new visitor instances for each new 
        protected internal ExcelColumnFactory(IDataProviderVisitor[] visitors)
        {
            DataProviderVisitors = visitors;
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
            
            InitializeForDataType(column, propertyInfo.PropertyType, propertyAccessor);

            return column;
        }

        public virtual ExcelColumn<IDataReader> CreateDataReaderColumn(DataColumn dataColumn, int index)
        {
            var column = new ExcelColumn<IDataReader>
            {
                Caption = TypeHelper.GetCaption(dataColumn),
            };

            InitializeForDataType(column, dataColumn.DataType, r => r[index]);

            return column;
        }

        public virtual ExcelColumn<DataRow> CreateDataTableColumn(DataColumn dataColumn, int index)
        {
            var column = new ExcelColumn<DataRow>
            {
                Caption = TypeHelper.GetCaption(dataColumn),
            };

            InitializeForDataType(column, dataColumn.DataType, r => r[index]);

            return column;
        }

        protected virtual void InitializeForDataType<T>(ExcelColumn<T> column, Type dataType, Func<T, object> dataProvider)
        {
            foreach (var visitor in DataProviderVisitors)
            {
                if (visitor.Visit(column, dataType, dataProvider))
                    return;
            }
        }
    }
}