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
    using System;
    using System.Collections.Generic;
    using System.Linq;

    using Bindings;
    using Controls;
    using Extensions;

    using Microsoft.Office.Interop.Excel;

    /// <summary>
    /// Represents a visual over an Excel worksheet
    /// </summary>
    public class Sheet : View
    {
        #region Fields

        private readonly Dictionary<string, Command> commands = 
            new Dictionary<string, Command>(StringComparer.OrdinalIgnoreCase);

        private readonly Dictionary<string, Form> forms = 
            new Dictionary<string, Form>(StringComparer.OrdinalIgnoreCase);
        
        private readonly Dictionary<string, Table> tables = 
            new Dictionary<string, Table>(StringComparer.OrdinalIgnoreCase);

        #endregion Fields

        #region Constructors

        /// <summary>
        /// Initiaalises an instance of ExcelMvc.Views.Workspace
        /// </summary>
        /// <param name="parent"></param>
        /// <param name="sheet">The underlying Excel Worksheet</param>
        internal Sheet(View parent, Worksheet sheet)
        {
            Parent = parent;
            Underlying = sheet;
        }

        #endregion Constructors

        #region Events

        /// <summary>
        /// Occurs when a command is clicked
        /// </summary>
        public event ClickedHandler Clicked = delegate { };

        #endregion Events

        #region Properties

        /// <summary>
        /// Gets the child views
        /// </summary>
        public override IEnumerable<View> Children
        {
            get
            {
                var lforms = from form in forms.Values
                             select (View)form;
                var ltables = from table in tables.Values
                              select (View)table;
                return lforms.Concat(ltables);
            }
        }

        /// <summary>
        /// Gets the Commands on the sheet
        /// </summary>
        public override IEnumerable<Command> Commands
        {
            get { return commands.Values.ToList(); }
        }

        /// <summary>
        /// Gets the view id
        /// </summary>
        public override string Id
        {
            get { return Name; }
        }

        /// <summary>
        /// Gets the view name
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
            get { return ViewType.Sheet; }
        }

        /// <summary>
        /// The underlying Excel sheet
        /// </summary>
        internal Worksheet Underlying
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
            foreach (var cmd in commands.Values)
                cmd.Dispose();
            commands.Clear();

            foreach (var form in forms.Values)
                form.Dispose();
            forms.Clear();

            foreach (var table in tables.Values)
                table.Dispose();
            tables.Clear();
        }

        /// <summary>
        /// Rebinds the view with bindings supplied
        /// </summary>
        /// <param name="bindings"></param>
        /// <param name="recursive"></param>
        internal override void Rebind(Dictionary<Worksheet, List<Binding>> bindings, bool recursive)
        {
            List<Binding> sheetBindings;
            if (!bindings.TryGetValue(Underlying, out sheetBindings))
                return;

            if (!recursive)
                return;

            foreach (var view in Children)
                view.Rebind(bindings, true);
        }

        internal void Initialise(IEnumerable<Binding> bindings)
        {
            Dispose();

            if (bindings != null)
                CreateViews(bindings);

            CommandFactory.Create(Underlying, this, commands);
            foreach (var cmd in commands.Values)
                cmd.Clicked += Cmd_Clicked;
        }

        private void Cmd_Clicked(object sender, CommandEventArgs args)
        {
            Clicked(sender, args);
            if (args.Handled)
                return;

            var views = forms.Values.Select(x => x as BindingView).ToList();
            views.AddRange(tables.Values.Select(x => x as BindingView));
            foreach (var view in views)
            {
                view.FireClicked(sender, args);
                if (args.Handled)
                    return;
            }
        }

        private void CreateViews(IEnumerable<Binding> bindings)
        {
            CreateViews(bindings, ViewType.Form, (x, y) => new Form(x, y), forms);
            CreateViews(bindings, ViewType.Table, (x, y) => new Table(x, y), tables);
        }

        private void CreateViews<T>(
            IEnumerable<Binding> bindings,
            ViewType type, 
            Func<Sheet, IEnumerable<Binding>, T> create,
            Dictionary<string, T> views) where T : View
        {
            var names = bindings.Where(x => x.Type == type).Select(x => x.Name).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            foreach (var item in names)
            {
                var name = item;
                var fields = bindings.Where(x => x.Type == type && x.Name.CompareOrdinalIgnoreCase(name) == 0).ToList();
                var view = create(this, fields);
                var args = new ViewEventArgs(view);
                views[name] = view;
                OnOpened(args);
            }
        }

        #endregion Methods
    }
}