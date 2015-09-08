using System;
using System.Data;
using System.Reflection;
using LFEG.ExcelColumnDataInitializerVisitors;

namespace LFEG
{
    public class ExcelColumnFactory
    {
        public static IExcelColumnDataInitializerVisitor[] DefaultDataInitializerVisitors =
        {
            new EnumDataInitializerVisitor(),
            new NumericDataInitializerVisitor(),
            new DefaultStringDataInitializerVisitor(),
        };
        
        protected IExcelColumnDataInitializerVisitor[] DataInitializerVisitors;

        public ExcelColumnFactory(IExcelColumnDataInitializerVisitor[] visitors = null)
        {
            DataInitializerVisitors = visitors ?? DefaultDataInitializerVisitors;
        }

        public virtual ExcelColumn CreatePropertyColumn(PropertyInfo property)
        {
            var column = new ExcelColumn
            {
                Caption = TypeHelper.GetCaption(property)
            };

            InitializeForDataType(column, property.PropertyType, property.GetValue);

            return column;
        }

        public virtual ExcelColumn CreateDataReaderColumn(DataColumn dataColumn, int index)
        {
            var column = new ExcelColumn
            {
                Caption = TypeHelper.GetCaption(dataColumn),
            };

            InitializeForDataType(column, dataColumn.DataType, r => ((IDataReader)r)[index]);

            return column;
        }

        public virtual ExcelColumn CreateDataTableColumn(DataColumn dataColumn, int index)
        {
            var column = new ExcelColumn
            {
                Caption = TypeHelper.GetCaption(dataColumn),
            };

            InitializeForDataType(column, dataColumn.DataType, r => ((DataRow)r)[index]);

            return column;
        }

        protected virtual void InitializeForDataType(ExcelColumn column, Type dataType, Func<object, object> dataProvider)
        {
            foreach (var visitor in DataInitializerVisitors)
            {
                if (visitor.Visit(column, dataType, dataProvider))
                    return;
            }
        }
    }
}