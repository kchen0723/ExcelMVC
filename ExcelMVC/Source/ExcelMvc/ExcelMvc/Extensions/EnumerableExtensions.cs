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
    using System.Collections;
    using System.Collections.Generic;

    /// <summary>
    /// Encapsulates commonly used extensions for IEnumerable
    /// </summary>
    public static class EnumerableExtensions
    {
        #region Methods

        /// <summary>
        /// Gets the index of an item from an IEnumerable
        /// </summary>
        /// <param name="source">Source to be iterated</param>
        /// <param name="item">Item to be searched</param>
        /// <returns>Index of the item, or -1 if not found</returns>
        public static int GetIndex(this IEnumerable source, object item)
        {
            var items = new List<object>();
            var iterator = source.GetEnumerator();
            var idx = 0;
            while (iterator.MoveNext())
            {
                if (iterator.Current == item)
                    return idx;
                idx++;
            }

            return -1;
        }

        /// <summary>
        /// Gets items from an IEnumerable
        /// </summary>
        /// <param name="source">Source to be iterated</param>
        /// <param name="start">Start index</param>
        /// <param name="count">Number of items to get</param>
        /// <returns>Items fetched</returns>
        public static IEnumerable<object> GetItems(this IEnumerable source, int start, int count)
        {
            var items = new List<object>();
            var iterator = source.GetEnumerator();
            var idx = 0;
            while (iterator.MoveNext())
            {
                if (idx >= start)
                    items.Add(iterator.Current);
                if ((++idx) >= start + count)
                    break;
            }

            return items;
        }

        /// <summary>
        /// Converts this enumerable to List&lt;object&gt;
        /// </summary>
        /// <param name="source">Enumerable to be converted</param>
        /// <returns>List of objects</returns>
        public static IList<object> ToList(this IEnumerable source)
        {
            var items = new List<object>();
            var iterator = source.GetEnumerator();
            while (iterator.MoveNext())
                items.Add(iterator.Current);
            return items;
        }

        #endregion Methods
    }
}