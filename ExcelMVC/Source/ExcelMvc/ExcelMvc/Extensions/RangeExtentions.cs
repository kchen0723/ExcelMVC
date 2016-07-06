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
    using Microsoft.Office.Interop.Excel;
    using Views;

    /// <summary>
    /// Encapsulates commonly used extensions for Range
    /// </summary>
    internal static class RangeExtentions
    {
        #region Methods

        /// <summary>
        /// Executes an action on a protected host
        /// </summary>
        /// <param name="host">Hosting View</param>
        /// <param name="action">Action to be executed</param>
        public static void ExecuteProtected(this View host, System.Action action)
        {
            host.ExecuteBinding(() =>
            {
                var sheet = ((Sheet)host).Underlying;
                if (!sheet.ProtectContents)
                {
                    action();
                    return;
                }

                // stops screen flickering
                var updating = App.Instance.Underlying.ScreenUpdating;
                App.Instance.Underlying.ScreenUpdating = false;

                var args = new ViewEventArgs(host);
                if (args.State == null)
                    sheet.Unprotect();
                else 
                    sheet.Unprotect(args.State as string);
                try
                {
                    action();
                }
                finally
                {
                    if (args.State == null)
                        sheet.Protect();
                    else
                        sheet.Protect(args.State as string);
                    App.Instance.Underlying.ScreenUpdating = updating;
                }
            });
        }

        /// <summary>
        /// Makes a new range
        /// </summary>
        /// <param name="range">Base range</param>
        /// <param name="rowOffset">Start row offset</param>
        /// <param name="rows">Rows to extend from the binding Cell</param>
        /// <param name="columnOffset">Start column offset</param>
        /// <param name="columns">Columns to extend from the binding Cell</param>
        /// <returns>Column range</returns>
        public static Range MakeRange(this Range range, int rowOffset, int rows, int columnOffset, int columns)
        {
            var start = range.Worksheet.Cells[range.Row + rowOffset, range.Column + columnOffset];
            var end = range.Worksheet.Cells[range.Row + rowOffset + rows - 1, range.Column + +columnOffset + columns - 1];
            return range.Worksheet.Range[start, end];
        }

        #endregion Methods
    }
}