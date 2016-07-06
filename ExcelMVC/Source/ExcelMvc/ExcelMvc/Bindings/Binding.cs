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
    using System.Windows.Data;
    using Microsoft.Office.Interop.Excel;
    using Views;

    /// <summary>
    /// Represents either a form field binding or a table column binding between 
    /// the View (Excel) and its view model
    /// </summary>
    public class Binding
    {
        #region Constructors

        /// <summary>
        /// Initialises an instance of ExcelMvc.Binding
        /// </summary>
        public Binding()
        {
            Type = ViewType.None;
            Mode = ModeType.OneWay;
            Name = string.Empty;
            Path = string.Empty;
            Visible = true;
        }

        #endregion Constructors

        #region Properties

        /// <summary>
        /// Start cell 
        /// </summary>
        public Range StartCell
        {
            get; internal set;
        }

        /// <summary>
        /// End cell (null for no binding boundary limit)
        /// </summary>
        public Range EndCell
        {
            get;
            internal set;
        }

        /// <summary>
        /// Value converter
        /// </summary>
        public IValueConverter Converter
        {
            get; set;
        }

        /// <summary>
        /// Gets and sets the mode Type
        /// </summary>
        public ModeType Mode
        {
            get; set;
        }

        /// <summary>
        /// Property path
        /// </summary>
        public string Path
        {
            get; set;
        }

        /// <summary>
        /// Validation list address
        /// </summary>
        public string ValidationList
        {
            get; set;
        }

        /// <summary>
        /// Visible
        /// </summary>
        public bool Visible
        {
            get; set;
        }

        /// <summary>
        /// View name
        /// </summary>
        internal string Name
        {
            get; set;
        }

        /// <summary>
        /// Gets and sets the View Type
        /// </summary>
        internal ViewType Type
        {
            get;
            set;
        }

        #endregion Properties

        #region Methods

        /// <summary>
        /// Makes a range from the binding cell
        /// </summary>
        /// <param name="rowOffset">Start row offset</param>
        /// <param name="rows">Rows to extend from the binding Cell</param>
        /// <param name="columnOffset">Start column offset</param>
        /// <param name="cols">Columns to extend from the binding Cell</param>
        /// <returns>Column range</returns>
        public Range MakeRange(int rowOffset, int rows, int columnOffset, int cols)
        {
            var start = StartCell.Worksheet.Cells[StartCell.Row + rowOffset, StartCell.Column + columnOffset];
            var end = StartCell.Worksheet.Cells[StartCell.Row + rowOffset + rows - 1, StartCell.Column + +columnOffset + cols - 1];
            return StartCell.Worksheet.Range[start, end];
        }

        #endregion Methods
    }
}