#region Header
/*
Copyright (C) 2013 =>

Creator:           Peter Gu, Australia
Developer:         Wolfgang Stamm, Germany

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and
associated documentation files (the "Software"), to deal in the Software without restriction,
including without limitation the rights to use, copy, modify, merge, publish, distribute,
sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or
substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING
BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

This program is free software; you can redistribute it and/or modify it under the terms of the
GNU General Public License as published by the Free Software Foundation; either version 2 of
the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY;
without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
See the GNU General Public License for more details.

You should have received a copy of the GNU General Public License along with this program;
if not, write to the Free Software Foundation, Inc., 51 Franklin Street, Fifth Floor,
Boston, MA 02110-1301 USA.
*/
#endregion Header

namespace ExcelMvc.Extensions
{
    using System;
    using System.Runtime.InteropServices;

    /// <summary>
    /// Encapsulates commonly used extensions for Action
    /// </summary>
    public static class ActionExtensions
    {
        #region Delegates

        /// <summary>
        /// Defines an Exception handler
        /// </summary>
        /// <param name="ex">Exception to be handled</param>
        public delegate void ExceptionHandler(Exception ex);

        #endregion Delegates

        #region Methods

        /// <summary>
        /// Executes an action, catches and/or handles any exceptions
        /// </summary>
        /// <param name="action">Action to be executed</param>
        /// <param name="handler">Exception hadler to be used</param>
        /// <returns>Exception caught</returns>
        public static Exception Try(Action action, ExceptionHandler handler = null)
        {
            Exception status = null;
            try
            {
                action();
            }
            catch (Exception ex)
            {
                status = ex;
                if (handler != null)
                    handler(ex);
            }

            return status;
        }

        /// <summary>
        /// Executes an action and wraps the exception with ErrorWrapper
        /// </summary>
        /// <param name="action">Action to be executed</param>
        /// <returns>null or an instance of ErrorWrapper</returns>
        public static ErrorWrapper Wrap(Action action)
        {
            ErrorWrapper status = null;
            try
            {
                action();
            }
            catch (Exception ex)
            {
                status = new ErrorWrapper(ex);
            }

            return status;
        }

        #endregion Methods
    }
}