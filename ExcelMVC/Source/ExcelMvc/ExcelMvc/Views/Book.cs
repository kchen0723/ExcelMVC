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

namespace ExcelMvc.Views
{
    using System.Collections.Generic;

    using Bindings;
    using Controls;
    using Extensions;
    using Microsoft.Office.Core;
    using Microsoft.Office.Interop.Excel;

    /// <summary>
    /// Represents a visual over an Excel workbook
    /// </summary>
    public class Book : View
    {
        #region Fields

        private readonly Dictionary<Worksheet, Sheet> sheets = new Dictionary<Worksheet, Sheet>();

        #endregion Fields

        #region Constructors

        internal Book(View parent, Workbook book)
        {
            Parent = parent;
            Underlying = book;
        }

        #endregion Constructors

        #region Properties

        /// <summary>
        /// Gets the child views
        /// </summary>
        public override IEnumerable<View> Children
        {
            get
            {
                foreach (var value in sheets.Values)
                    yield return value as View;
            }
        }

        /// <summary>
        /// Gets the full book name
        /// </summary>
        public string FullName
        {
            get { return Underlying.FullName; }
        }

        /// <summary>
        /// Gets the book id, as defined by the Custom Document Propety named "ExcelMvc"
        /// </summary>
        public override string Id
        {
            get
            {
                var value = string.Empty;
                ActionExtensions.Try(() => value = ((DocumentProperties)Underlying.CustomDocumentProperties)[App.ExcelMvc].Value.ToString());
                return value;
            }
        }

        /// <summary>
        /// Gets the book name
        /// </summary>
        public override string Name
        {
            get { return Underlying.Name; }
        }

        /// <summary>
        /// Gets the view type
        /// </summary>
        public override ViewType Type
        {
            get { return ViewType.Book; }
        }

        /// <summary>
        /// Underlying Excel Workbook
        /// </summary>
        internal Workbook Underlying
        {
            get; private set;
        }

        #endregion Properties

        #region Methods

        /// <summary>
        /// Disposes resources
        /// </summary>
        public override void Dispose()
        {
            ExecuteBinding(() =>
            {
                if (Underlying != null)
                {
                    Underlying.SheetActivate -= Underlying_SheetActivate;
                    Underlying.SheetDeactivate -= Underlying_SheetDeactivate;
                }

                foreach (var item in sheets.Values)
                    item.Dispose();
                sheets.Clear();
            });
        }

        /// <summary>
        /// Collects bindings and rebinds the view
        /// </summary>
        /// <param name="recursive"></param>
        public override void Rebind(bool recursive)
        {
            var bindings = new BindingCollector(Underlying).Process();
            foreach (var view in Children)
                view.Rebind(bindings, recursive);
        }

        /// <summary>
        /// Finds a specfic command
        /// </summary>
        /// <param name="sheet">Sheet hosts the command</param>
        /// <param name="name">Name of the command</param>
        /// <returns>Command found or nullf if not found</returns>
        internal Command FindCommand(Worksheet sheet, string name)
        {
            Command cmd = null;
            Sheet item;
            if (sheets.TryGetValue(sheet, out item))
                cmd = item.FindCommand(name);
            return cmd;
        }

        /// <summary>
        /// Initialises the view
        /// </summary>
        internal void Initialise()
        {
            Dispose();
            ExecuteBinding(() =>
            {
                var bindings = new BindingCollector(Underlying).Process();
                foreach (Worksheet item in Underlying.Worksheets)
                {
                    var view = new Sheet(this, item);
                    List<Binding> sheetBindings;
                    bindings.TryGetValue(item, out sheetBindings);
                    view.Initialise(sheetBindings);
                    sheets[item] = view;
                }

                Underlying.SheetActivate += Underlying_SheetActivate;
                Underlying.SheetDeactivate += Underlying_SheetDeactivate;
            });
        }

        private void Underlying_SheetActivate(object sh)
        {
            OnActivated(new ViewEventArgs(sheets[(Worksheet)sh]));
        }

        private void Underlying_SheetDeactivate(object sh)
        {
            OnDeactivated(new ViewEventArgs(sheets[(Worksheet)sh]));
        }

        #endregion Methods
    }
}