using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MailLib
{
    /// <summary>
    /// コンストラクタ
    /// </summary>
    public class PopException : Exception
    {
        /// <summary>
        /// コンストラクタです。
        /// </summary>
        public PopException()
        {
        }

        /// <summary>
        /// コンストラクタです。
        /// </summary>
        /// <param name="message"></param>
        public PopException(string message) : base(message)
        {
        }

        /// <summary>
        /// コンストラクタです。
        /// </summary>
        /// <param name="message"></param>
        /// <param name="innerException"></param>
        public PopException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}
