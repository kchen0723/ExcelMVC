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
    using System;

    using Microsoft.Office.Interop.Excel;

    /// <summary>
    /// Wraps commonly used range conversion functions
    /// </summary>
    internal static class RangeConversion
    {
        #region Fields

        private static readonly int[] OneBased = { 1, 1 };

        #endregion Fields

        #region Methods

        /// <summary>
        /// Merges the value of a range to the value of another range
        /// </summary>
        /// <param name="changed">Range to be merged</param>
        /// <param name="to">Range being merged</param>
        /// <param name="toValue">Value of the To range</param>
        /// <returns>Merge result</returns>
        public static MergeResult MergeChangedValue(Range changed, Range to, object toValue)
        {
            var count = 0;
            if (to.Count == 1)
            {
                object newValue = (toValue == null || changed.Value == null) ? changed.Value
                    : Convert.ChangeType(changed.Value, toValue.GetType());
                if (!Equals(toValue, newValue))
                {
                    toValue = newValue;
                    count++;
                }
            }
            else
            {
                var toArray = (object[,])toValue;
                if (changed.Count == 1)
                {
                    if (!Equals(toArray[changed.Row - to.Row + 1, changed.Column  - to.Column + 1], changed.Value))
                    {
                        toArray[changed.Row - to.Row + 1, changed.Column  - to.Column + 1] = changed.Value;
                        count++;
                    }
                }
                else
                {
                    var changeArray = (object[,])changed.Value;
                    for (var idx = changeArray.GetLowerBound(0); idx <= changeArray.GetUpperBound(0); idx++)
                    {
                        for (var jdx = changeArray.GetLowerBound(1); jdx <= changeArray.GetUpperBound(1); jdx++)
                        {
                            if (!Equals(changeArray[idx, jdx], toArray[changed.Row - to.Row + idx, changed.Column - to.Column + jdx]))
                            {
                                toArray[changed.Row - to.Row + idx, changed.Column - to.Column + jdx] =
                                    changeArray[idx, jdx];
                                count++;
                            }
                        }
                    }
                }
            }

            MergeResult result;
            result.Value = toValue;
            result.Changed = count > 0;
            return result;
        }

        /// <summary>
        ///  Converts a range to an instance of Matrix
        /// </summary>
        /// <param name="range">Range to be converted</param>
        /// <param name="isErrorChecked">Indicates if Excel errors are checked</param>
        /// <param name="isErrorFilled">Indicates if Excel errors are filed</param>
        /// <param name="errorFiller">Error filler</param>
        /// <returns>Matrix instance</returns>
        private static Matrix RangeToMatrix(Range range, bool isErrorChecked, bool isErrorFilled, object errorFiller)
        {
            Matrix result;
            if (range.Count == 1)
            {
                result.Value = (object[,])Array.CreateInstance(typeof(object), OneBased, OneBased);
                result.Value[1, 1] = range.Value;
            }
            else
            {
                result.Value = (object[,])range.Value;
            }

            result.Error = null;
            if (!isErrorChecked)
                return result;

            var value = result.Value;
            var funs = range.Worksheet.Application.WorksheetFunction;
            result.Error = (ErrorCode?[,])Array.CreateInstance(typeof(ErrorCode?), new[] { value.GetLength(0), value.GetLength(1) }, OneBased);
            for (var idx = value.GetLowerBound(0); idx <= value.GetUpperBound(0); idx++)
            {
                for (var jdx = value.GetLowerBound(1); jdx <= value.GetUpperBound(1); jdx++)
                {
                    ErrorCode? code;
                    if (value[idx, jdx] is int && (code = ErrorConverter.IntToErrorCode((int)value[idx, jdx])) != null)
                    {
                        result.Error[idx, jdx] = funs.IsError(range.Cells[idx, jdx]) ? code : null;
                        if (isErrorFilled)
                            value[idx, jdx] = errorFiller;
                    }
                }
            }

            return result;
        }

        #endregion Methods

        #region Nested Types

        /// <summary>
        /// Struct that captures range values and error codes
        /// </summary>
        private struct Matrix
        {
            #region Fields

            public ErrorCode?[,] Error;
            public object[,] Value;

            #endregion Fields
        }

        #endregion Nested Types
    }
}