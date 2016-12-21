using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPLus.Extensions
{
    /// <summary>
    /// Class contains exception circumstances
    /// </summary>
    public class ExcelTableConvertExceptionArgs
    {
        public string propertyName { get; set; }
        public string columnName { get; set; }
        public Type expectedType { get; set; }
        public object cellValue { get; set; }
    }

    /// <summary>
    /// Class extends exception to hold casting exception circumstances
    /// </summary>
    public class ExcelTableConvertException : Exception
    {
        public ExcelTableConvertExceptionArgs args { get; private set; }

        public ExcelTableConvertException()
        {
        }

        public ExcelTableConvertException(string message)
        : base(message)
        {
        }

        public ExcelTableConvertException(string message, Exception inner)
        : base(message, inner)
        {
        }

        public ExcelTableConvertException(string message, Exception inner, ExcelTableConvertExceptionArgs args)
        : base(message, inner)
        {
            this.args = args;
        }
    }
}
