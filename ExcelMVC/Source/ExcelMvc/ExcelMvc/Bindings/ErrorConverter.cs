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

namespace ExcelMvc.Bindings
{
    using System.Runtime.InteropServices;

    #region Enumerations

    /// <summary>
    /// Defines Excel errors
    /// </summary>
    internal enum ErrorCode
    {
        ErrDiv0  = -2146826281,
        ErrNA    = -2146826246,
        ErrName  = -2146826259,
        ErrNull  = -2146826288,
        ErrNum   = -2146826252,
        ErrRef   = -2146826265,
        ErrValue = -2146826273
    }

    #endregion Enumerations

    /// <summary>
    /// Wraps commonly used Error conversion functions
    /// </summary>
    internal static class ErrorConverter
    {
        #region Methods

        /// <summary>
        /// Converts an int value to ErrorCode
        /// </summary>
        /// <param name="value">Value to be convented</param>
        /// <returns>ErrorCode or null if the number is not an error code</returns>
        public static ErrorCode? IntToErrorCode(int value)
        {
            ErrorCode? result = null;
            if (value > -2146826246)
                return null;

            switch ((ErrorCode)value)
            {
                case ErrorCode.ErrDiv0:
                    result = ErrorCode.ErrDiv0;
                    break;
                case ErrorCode.ErrNA:
                    result = ErrorCode.ErrNA;
                    break;
                case ErrorCode.ErrName:
                    result = ErrorCode.ErrName;
                    break;
                case ErrorCode.ErrNull:
                    result = ErrorCode.ErrNull;
                    break;
                case ErrorCode.ErrNum:
                    result = ErrorCode.ErrNum;
                    break;
                case ErrorCode.ErrRef:
                    result = ErrorCode.ErrRef;
                    break;
                case ErrorCode.ErrValue:
                    result = ErrorCode.ErrValue;
                    break;
            }

            return result;
        }

        /// <summary>
        /// Converts an error code to Interop ErrorWrapper
        /// </summary>
        /// <param name="code">Error code </param>
        /// <returns>ErrorWrapper instance</returns>
        public static ErrorWrapper ToErrorWrapper(ErrorCode code)
        {
            return new ErrorWrapper((int)code);
        }

        #endregion Methods
    }
}