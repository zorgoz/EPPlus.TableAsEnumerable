/* 
 * Copyright (c) 2016 Zoltán Zörgő
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), 
 * to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, 
 * and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
 * 
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER 
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS 
 * IN THE SOFTWARE.
 */

using System;

namespace EPPlus.Extensions
{
    /// <summary>
    /// Class contains exception circumstances
    /// </summary>
    public class ExcelTableConvertExceptionArgs
    {
        /// <summary>
        /// Property that was tried to set
        /// </summary>
        public string propertyName { get; set; }
        
        /// <summary>
        /// Column that was mapped to this property
        /// </summary>
        public string columnName { get; set; }

        /// <summary>
        /// Type of the property
        /// </summary>
        public Type expectedType { get; set; }

        /// <summary>
        /// Cell value returned by EPPlus
        /// </summary>
        public object cellValue { get; set; }
    }

    /// <summary>
    /// Class extends exception to hold casting exception circumstances
    /// </summary>
    public class ExcelTableConvertException : Exception
    {
        public ExcelTableConvertExceptionArgs args { get; private set; }

        /// <summary>
        /// Default constructor
        /// </summary>
        public ExcelTableConvertException()
        {
        }

        /// <summary>
        /// Constructor with message
        /// </summary>
        /// <param name="message">Exception message</param>
        public ExcelTableConvertException(string message)
        : base(message)
        {
        }

        /// <summary>
        /// Constructor with message and inner exception
        /// </summary>
        /// <param name="message">Exception message</param>
        /// <param name="inner">Inner exception</param>
        public ExcelTableConvertException(string message, Exception inner)
        : base(message, inner)
        {
        }

        /// <summary>
        /// Custom constructor that takes message,, inner exception and conversion arguments
        /// </summary>
        /// <param name="message">Exception message</param>
        /// <param name="inner">Actual conversion exception catched</param>
        /// <param name="args">Information related to the circumstances of the exception</param>
        public ExcelTableConvertException(string message, Exception inner, ExcelTableConvertExceptionArgs args)
        : base(message, inner)
        {
            this.args = args;
        }
    }
}
