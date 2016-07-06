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
namespace ExcelMvc.Runtime
{
    using System;
    using System.Runtime.InteropServices;
    using System.Threading;
    using Diagnostics;
    using Extensions;

    using Microsoft.Office.Interop.Excel;
    using Views;

    /// <summary>
    /// Encapsulates Range updating functions
    /// </summary>
    internal class RangeUpdator
    {
        #region Fields

        private static readonly Lazy<RangeUpdator> LazyInstance = new Lazy<RangeUpdator>(() => new RangeUpdator());

        #endregion Fields

        #region Constructors

        private RangeUpdator()
        {
        }

        #endregion Constructors

        #region Properties

        /// <summary>
        /// Singleton
        /// </summary>
        public static RangeUpdator Instance
        {
            get { return LazyInstance.Value; }
        }

        internal static string NameOfAsynUpdateThread
        {
            get { return "ExcelMvcAsynUpdateThread"; }
        }

        #endregion Properties

        #region Methods

        public void Update(Range range, int rowOffset, int rows, int columnOffset, int columns, object value)
        {
            if (IsAsyncUpdateThread())
                Enqueue(new Item { Range = range, RowOffset = rowOffset, Rows = rows, ColumnOffset = columnOffset, Columns = columns, Value = value });
            else
                AssignRangeValue(range.MakeRange(rowOffset, rows, columnOffset, columns), value);
        }

        public void Update(Range range, Range rowIdStart, int rowCount, string rowId, int rows, int columnOffset, int columns, object value)
        {
            if (IsAsyncUpdateThread())
                Enqueue(new Item
                {
                    Range = range,
                    RowIdStart = rowIdStart,
                    RowId = rowId,
                    RowCount = rowCount,
                    Rows = rows,
                    RowOffset = int.Parse(rowId),
                    ColumnOffset = columnOffset,
                    Columns = columns,
                    Value = value
                });
            else
                AssignRangeValue(range.MakeRange(RowOffsetFromRowId(rowIdStart, rowCount, rowId), rows, columnOffset, columns), value);
        }

        public void Update(Range range, int rowOffset, int rows, Range colIdStart, int colCount, string colId, int columns, object value)
        {
            if (IsAsyncUpdateThread())
                Enqueue(new Item
                {
                    Range = range,
                    ColIdStart = colIdStart,
                    ColId = colId,
                    ColCount = colCount,
                    Rows = rows,
                    RowOffset = rowOffset,
                    ColumnOffset = int.Parse(colId),
                    Columns = columns,
                    Value = value
                });
            else
                AssignRangeValue(range.MakeRange(rowOffset, rows, ColOffsetFromColId(colIdStart, colCount, colId), columns), value);
        }

        private static int ColOffsetFromColId(Range start, int count, string colId)
        {
            var offset = -1;
            if (start == null)
            {
                // columns are assumed not shuffled after binding
                offset = int.Parse(colId);
            }
            else
            {
                // this can be very slow (needs a better way)
                var row = start.MakeRange(0, 1, 0, count);
                for (var idx = 0; idx < count; idx++)
                {
                    if (((Range)row.Cells[1, idx + 1]).ID != colId)
                        continue;

                    offset = idx;
                    break;
                }
            }

            return offset;
        }

        private static bool IsAsyncUpdateThread()
        {
            var threadName = Thread.CurrentThread.Name;
            return !string.IsNullOrEmpty(threadName) && threadName.CompareOrdinalIgnoreCase(NameOfAsynUpdateThread) == 0;
        }

        private static int RowOffsetFromRowId(Range start, int count, string rowId)
        {
            var offset = -1;
            if (start == null)
            {
                // rows are assumed not sorted after binding
                offset = int.Parse(rowId);
            }
            else
            {
                // this can be very slow (needs a better way)
                var column = start.MakeRange(0, count, 0, 1);
                for (var idx = 0; idx < count; idx++)
                {
                    if (((Range)column.Cells[idx + 1, 1]).ID != rowId)
                        continue;

                    offset = idx;
                    break;
                }
            }

            return offset;
        }

        private static void Enqueue(Item item, int pumpMilliseconds = 0)
        {
            AsyncActions.Post(UpdateAsync, item, true, pumpMilliseconds);
        }

        private static void UpdateAsync(object state)
        {
            var item = (Item) state;
            try
            {
                var rowOffset = item.RowIdStart == null ? item.RowOffset
                    : RowOffsetFromRowId(item.RowIdStart, item.RowCount, item.RowId);
                var colOffset = item.ColIdStart == null ? item.ColumnOffset
                    : ColOffsetFromColId(item.ColIdStart, item.ColCount, item.ColId);
                AssignRangeValue(item.Range.MakeRange(rowOffset, item.Rows, colOffset, item.Columns), item.Value);
            }
            catch(Exception ex)
            {
                var comex = (ex as COMException) ?? ex.InnerException as COMException;
                if (IsRecoverable(comex))
                {
                    item.AgeMilliseconds += 100;
                    if ( item.AgeMilliseconds > 10000)
                        MessageWindow.AddErrorLine(ex);
                    else
                        Enqueue(item, 100);
                }
                else
                {
                    MessageWindow.AddErrorLine(ex);
                }
            }
        }

        private static void AssignRangeValue(Range range, object value)
        {
            var locked = Convert.ToBoolean(range.Locked);
            if (locked && range.Worksheet.ProtectContents)
            {
                var book = App.Instance.Find(ViewType.Book, (range.Worksheet.Parent as Workbook).Name);
                var sheet = book.Find(ViewType.Sheet, range.Worksheet.Name);
                sheet.ExecuteProtected(() => range.Value = value);
            }
            else
            {
                range.Value = value;
            }
        }

        static bool IsRecoverable(COMException ex)
        {
            const uint RPC_E_SERVERCALL_RETRYLATER = 0x8001010A;
            const uint RPC_E_CALL_REJECTED = 0x80010001;
            const uint VBA_E_IGNORE = 0x800AC472;
            const uint NAME_NOT_FOUND = 0x800A03EC;
            var errorCode = (uint)ex.ErrorCode;
            switch (errorCode)
            {
                case RPC_E_SERVERCALL_RETRYLATER:
                case VBA_E_IGNORE:
                case NAME_NOT_FOUND:
                case RPC_E_CALL_REJECTED:
                    return true;
                default:
                    return false;
            }
        }

        #endregion Methods

        #region Nested Types

        private class Item
        {
            #region Properties

            public int ColCount
            {
                get;
                set;
            }

            public string ColId
            {
                get;
                set;
            }

            public Range ColIdStart
            {
                get;
                set;
            }

            public int ColumnOffset
            {
                get;
                set;
            }

            public int Columns
            {
                get;
                set;
            }

            public Range Range
            {
                get;
                set;
            }

            public int RowCount
            {
                get;
                set;
            }

            public string RowId
            {
                get;
                set;
            }

            public Range RowIdStart
            {
                get;
                set;
            }

            public int RowOffset
            {
                get;
                set;
            }

            public int Rows
            {
                get;
                set;
            }

            public object Value
            {
                get;
                set;
            }

            public int AgeMilliseconds
            {
                get;
                set;
            }

            #endregion Properties
        }

        #endregion Nested Types
    }
}