using System;
using System.Runtime.Serialization;

namespace LFEG.Infrastructure
{
    [Serializable]
    public class ValueExceedsExcelLimitException : Exception
    {
        public ValueExceedsExcelLimitException()
        {
        }

        public ValueExceedsExcelLimitException(string message) : base(message)
        {
        }

        public ValueExceedsExcelLimitException(string message, Exception inner) : base(message, inner)
        {
        }

        protected ValueExceedsExcelLimitException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }
    }
}