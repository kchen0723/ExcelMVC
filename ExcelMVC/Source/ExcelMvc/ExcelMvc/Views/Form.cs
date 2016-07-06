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
    using System.ComponentModel;
    using System.Linq;

    using Bindings;
    using Extensions;
    using Microsoft.Office.Interop.Excel;
    using Runtime;

    /// <summary>
    /// Represents a visual consists with scattered fields
    /// </summary>
    public class Form : BindingView
    {
        #region Fields

        private INotifyPropertyChanged notifyPropertyChanged;

        #endregion Fields

        #region Constructors

        /// <summary>
        /// Initialises an instances of ExcelMvc.Views.Form
        /// </summary>
        /// <param name="parent"></param>
        /// <param name="bindings">Bindings for the view</param>
        internal Form(View parent, IEnumerable<Binding> bindings)
            : base(parent, bindings)
        {
            SelectedBindings = new List<Binding>();
        }

        #endregion Constructors

        #region Properties

        /// <summary>
        /// 
        /// </summary>
        public override object Model
        {
            set
            {
                base.Model = value;
                UpdateView();
                OneWayToSource();
            }
        }

        /// <summary>
        /// Gets the selected bindings.
        /// </summary>
        public List<Binding> SelectedBindings
        {
            get; private set;
        }

        /// <summary>
        /// 
        /// </summary>
        public override ViewType Type
        {
            get { return ViewType.Form; }
        }

        #endregion Properties

        #region Methods

        /// <summary>
        /// Disposes resources
        /// </summary>
        public override void Dispose()
        {
            base.Model = null;
            UnhookModelEvents();
            UnhookViewEvents();
        }

        /// <summary>
        /// Rebinds the view with bindings supplied
        /// </summary>
        /// <param name="bindings"></param>
        /// <param name="recursive"></param>
        internal override void Rebind(Dictionary<Worksheet, List<Binding>> bindings, bool recursive)
        {
            List<Binding> sheetBindings;
            if (bindings.TryGetValue(((Sheet)Parent).Underlying, out sheetBindings))
            {
                // clear current view
                var current = Model;
                Model = null;
                
                // rebind
                Bindings = sheetBindings.Where(x => x.Type == Type && x.Name.CompareOrdinalIgnoreCase(Name) == 0).ToList();
                Model = current;
            }
        }

        private void HookModelEvents()
        {
            UnhookModelEvents();
            notifyPropertyChanged = Model as INotifyPropertyChanged;
            if (notifyPropertyChanged != null)
                notifyPropertyChanged.PropertyChanged += Notify_PropertyChanged;
        }

        private void HookViewEvents()
        {
            UnhookViewEvents();
            var sheet = (Sheet)Parent;
            sheet.Underlying.Change += Underlying_Change;
            sheet.Underlying.SelectionChange += Underlying_SelectionChange;
        }

        private void OneWayToSource()
        {
            var oneways = Bindings.Where(x => (x.Mode == ModeType.OneWayToSource));
            foreach (var oneway in oneways)
                UpdateObject(oneway, oneway.StartCell);
        }

        private void Notify_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            UpdateView(e.PropertyName ?? "*");
        }

        private void Underlying_Change(Range target)
        {
            UpdateObject(target);
        }

        private void Underlying_SelectionChange(Range target)
        {
            var count = SelectedBindings.Count;
            SelectedBindings.Clear();
            SelectedBindings.AddRange(Bindings.Where(binding => target.Application.Intersect(binding.StartCell, target) != null));
            if (count != 0 || SelectedBindings.Count != 0)
                OnSelectionChanged(new[] { Model }, SelectedBindings);
        }

        private void UnhookModelEvents()
        {
            if (notifyPropertyChanged != null)
                notifyPropertyChanged.PropertyChanged -= Notify_PropertyChanged;
        }

        private void UnhookViewEvents()
        {
            var sheet = (Sheet)Parent;
            sheet.Underlying.Change -= Underlying_Change;
            sheet.Underlying.SelectionChange -= Underlying_SelectionChange;
        }

        private void UpdateObject(Range target)
        {
            var toSource = Bindings.Where(x => (x.Mode == ModeType.TwoWay || x.Mode == ModeType.OneWayToSource));
            foreach (var binding in toSource)
                UpdateObject(binding, target);
        }

        private void UpdateObject(Binding binding, Range target)
        {
            ExecuteBinding(() =>
            {
                var range = binding.StartCell;
                var changed = target.Application.Intersect(range, target);
                if (changed != null)
                {
                    var value = RangeConversion.MergeChangedValue(changed, range, ObjectBinding.GetPropertyValue(Model, binding));
                    if (value.Changed)
                    {
                        ObjectBinding.SetPropertyValue(Model, binding, value.Value);
                        OnObjectChanged(new[] { Model }, new[] { binding.Path });
                    }
                }
            });
        }

        private void UpdateView()
        {
            ExecuteBinding(
                () =>
                {
                    UnhookViewEvents();
                    UnhookModelEvents();
                    UpdateView("*");
                    BindValidationLists(1);
                }, 
                () =>
                {
                    HookViewEvents();
                    HookModelEvents();
                });
        }

        private void UpdateView(string path)
        {
            ExecuteBinding(() =>
            {
                var match = string.IsNullOrEmpty(path) ? null : Bindings.FirstOrDefault(x => x.Path == path);
                if (match != null)
                {
                    UpdateView(match);
                }
                else if (path == "*")
                {
                    foreach (var binding in Bindings)
                        UpdateView(binding);
                }
            });
        }

        private void UpdateView(Binding binding)
        {
            if (binding.Mode == ModeType.OneWayToSource)
                return;

            ExecuteBinding(() =>
                {
                    var value = ObjectBinding.GetPropertyValue(Model, binding);
                    RangeUpdator.Instance.Update(binding.StartCell, 0, 1, 0, 1, value);
                });
        }

        #endregion Methods
    }
}