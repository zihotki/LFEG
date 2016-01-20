using System;

namespace LFEG.Infrastructure.DataProviderVisitors
{
    public class DefaultStringDataProviderVisitor : IDataProviderVisitor
    {
        private readonly bool _throwOnLimit;
        private const int ExcelCellLengthLimit = 32767;

        public DefaultStringDataProviderVisitor(bool throwOnLimit)
        {
            _throwOnLimit = throwOnLimit;
        }

        public bool Visit<T>(ExcelColumn<T> column, Type dataType, Func<T, object> dataProvider)
        {
            column.ExcelDataType = ExcelDataTypes.InlineString;

            column.DataProvider = i =>
            {
                var objValue = dataProvider(i);
                if (objValue == null || objValue == DBNull.Value)
                {
                    return string.Empty;
                }

                var value = objValue.ToString();

                if (value.Length <= ExcelCellLengthLimit)
                {
                    return value;
                }

                if (_throwOnLimit)
                {
                    throw new ValueExceedsExcelLimitException(
                        $"{column.Caption} column value is too long. Maximal length of a value in Excel is {ExcelCellLengthLimit}.");
                }

                return value.Substring(0, ExcelCellLengthLimit);
            };

            return true;
        }
    }
}