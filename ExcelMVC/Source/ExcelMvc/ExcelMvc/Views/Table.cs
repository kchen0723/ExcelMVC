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
    using System.Collections;
    using System.Collections.Generic;
    using System.Collections.Specialized;
    using System.ComponentModel;
    using System.Globalization;
    using System.Linq;
    using Bindings;
    using Extensions;
    using Microsoft.Office.Interop.Excel;
    using Runtime;

    /// <summary>
    /// Represents a rectangular visual with rows and columns
    /// </summary>
    public class Table : BindingView
    {
        #region Fields

        private readonly List<string> categoryIds = new List<string>();
        private bool orderingAllowed;

        private IEnumerable enumerable;
        private IList<object> itemsBound;
        private INotifyCollectionChanged notifyCollectionChanged;
        private INotifyPropertyChanged notifyPropertyChanged;

        #endregion Fields

        #region Constructors

        /// <summary>
        /// Initialises an instances of ExcelMvc.Views.Table
        /// </summary>
        /// <param name="parent"></param>
        /// <param name="bindings">Bindings for the view</param>
        internal Table(View parent, IEnumerable<Binding> bindings)
            : base(parent, bindings)
        {
            DeriveOrientation();
            SelectedItems = new List<object>();
            SelectedBindings = new List<Binding>();
            SetCategoryVisibility();
        }

        #endregion Constructors

        #region Properties

        /// <summary>
        /// Sets the underlying model
        /// </summary>
        public override object Model
        {
            set
            {
                AttachModel(value, true);
            }
        }

        /// <summary>
        /// Gets the selected bindings.
        /// </summary>
        public List<Binding> SelectedBindings
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the selected objects.
        /// </summary>
        public List<object> SelectedItems
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the view type
        /// </summary>
        public override ViewType Type
        {
            get { return ViewType.Table; }
        }

        /// <summary>
        /// Gets the maximum number of items to bind
        /// </summary>
        public int MaxItemsToBind
        {
            get
            {
                return Orientation == ViewOrientation.Portrait
                    ? Bindings.Max(x => x.EndCell == null ? int.MaxValue : (x.EndCell.Row - x.StartCell.Row + 1))
                    : Bindings.Max(x => x.EndCell == null ? int.MaxValue : (x.EndCell.Column - x.StartCell.Column + 1));
            }
        }

        /// <summary>
        /// Enables or disables handling row sorting or column re-ordering. Note enabling ordering can result in
        /// slow performance in binding. (Not recommended)
        /// </summary>
        public bool IsOrderingAllowed
        {
            get
            {
                return orderingAllowed;
            }

            set
            {
                if (value != orderingAllowed)
                {
                    orderingAllowed = value;
                    if (orderingAllowed && itemsBound != null)
                        AssignCategoryIds();
                }
            }
        }

        #endregion Properties

        #region Methods

        /// <summary>
        /// Disposes resources
        /// </summary>
        public override void Dispose()
        {
            HookViewEvents(false);
            HookItemsEvents(false);
            UnhookModelEvents();
            base.Model = null;
        }

        /// <summary>
        /// Sets the Model property to null, but does not clear the view's current content 
        /// </summary>
        public override void DetachModel()
        {
           AttachModel(null, false);
        }

        /// <summary>
        /// Sets visibiliy for all columns (portrait table) or for all rows (landscape table)
        /// </summary>
        public void SetCategoryVisibility()
        {
            foreach (var binding in Bindings)
                SetCategoryVisibility(binding);
        }

        /// <summary>
        /// Toggles the visibility of a column (portrait table) or row (landscape table)
        /// </summary>
        /// <param name="path">Binding path of the column or row</param>
        /// <returns>true if visible, false otherwise</returns>
        public bool ToggleCategoryVisibility(string path)
        {
            var binding = Bindings.FirstOrDefault(x => x.Path == path);
            if (binding != null)
            {
                binding.Visible = !binding.Visible;
                switch (Orientation)
                {
                    case ViewOrientation.Portrait:
                        Parent.ExecuteProtected(() => binding.StartCell.EntireColumn.Hidden = !binding.Visible);
                        break;
                    case ViewOrientation.Landscape:
                        Parent.ExecuteProtected(() => binding.StartCell.EntireRow.Hidden = !binding.Visible);
                        break;
                }
            }

            return binding != null && binding.Visible;
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
                // clear previous view
                var current = Model;
                Model = null;

                // rebind
                Bindings = sheetBindings.Where(x => x.Type == Type && x.Name.CompareOrdinalIgnoreCase(Name) == 0).ToList();
                DeriveOrientation();
                Model = current;
            }
        }

        private void AttachModel(object model, bool clearifnull) 
        {
            base.Model = model;
            UpdateView(clearifnull);
            OneWayToSource();
        }

        private void ClearRange(List<Binding> bindings, int numberItems)
        {
            var first = bindings[0];
            var numberCategories = bindings.Count;
            var numberRows = Orientation == ViewOrientation.Portrait ? numberItems : numberCategories;
            var numberCols = Orientation == ViewOrientation.Portrait ? numberCategories : numberItems;
            var cells = new object[numberRows, numberCols];
            RangeUpdator.Instance.Update(first.StartCell, 0, numberRows, 0, numberCols, cells);
        }

        private void ClearView(IEnumerable<List<Binding>> groups, int numberItems)
        {
            foreach (var group in groups)
                ClearRange(group, numberItems);
        }

        private List<List<Binding>> GroupBindings(IEnumerable<Binding> bindings)
        {
            var groups = new List<List<Binding>>();
            var isPortrait = Orientation == ViewOrientation.Portrait;
            var ordered = (from x in bindings orderby isPortrait ? x.StartCell.Column : x.StartCell.Row select x).ToList();
            while (ordered.Count > 0)
            {
                int idx;
                for (idx = 1; idx < ordered.Count; idx++)
                {
                    var currentBindingCell = isPortrait ? ordered[idx].StartCell.Column : ordered[idx].StartCell.Row;
                    var successorOfPreviousBindingCell = isPortrait ? ordered[idx - 1].StartCell.Column + 1 : ordered[idx - 1].StartCell.Row + 1;
                    if (currentBindingCell != successorOfPreviousBindingCell)
                        break;
                }

                groups.Add(ordered.Take(idx).ToList());
                ordered = ordered.Skip(idx).ToList();
            }

            return groups;
        }

        private bool UpdateObject(Binding binding, int itemOffset, object model, Range target)
        {
            var range = Orientation == ViewOrientation.Portrait ? binding.MakeRange(itemOffset, 1, 0, 1) : binding.MakeRange(0, 1, itemOffset, 1);
            var changed = range.Application.Intersect(range, target);
            var value = RangeConversion.MergeChangedValue(changed, range, ObjectBinding.GetPropertyValue(model, binding));
            if (value.Changed)
                ObjectBinding.SetPropertyValue(model, binding, value.Value);
            return value.Changed;
        }

        private void UpdateRange(List<Binding> bindings, IDictionary<Binding, List<object>> bindingValues, int numberItems)
        {
            var first = bindings[0];
            var numberCategories = bindings.Count;
            var numberRows = Orientation == ViewOrientation.Portrait ? numberItems : numberCategories;
            var numberCols = Orientation == ViewOrientation.Portrait ? numberCategories : numberItems;
            var cells = new object[numberRows, numberCols];
            for (var categoryIndex = 0; categoryIndex < bindings.Count; categoryIndex++)
            {
                var values = bindingValues[bindings[categoryIndex]];
                for (var itemIndex = 0; itemIndex < numberItems; itemIndex++)
                {
                    var idx = Orientation == ViewOrientation.Portrait ? itemIndex : categoryIndex;
                    var jdx = Orientation == ViewOrientation.Portrait ? categoryIndex : itemIndex;
                    cells[idx, jdx] = values[itemIndex];
                }
            }

            RangeUpdator.Instance.Update(first.StartCell, 0, numberRows, 0, numberCols, cells);
        }

        private void UpdateView(IEnumerable<List<Binding>> groups, IDictionary<Binding, List<object>> bindingValues, int numberItems)
        {
            foreach (var group in groups)
                UpdateRange(group, bindingValues, numberItems);
        }

        private void AssignCategoryIds()
        {
            var first = Bindings.First();
            var categoryRange = GetCategoryRange(first);
            Parent.ExecuteProtected(() =>
            {
                for (var itemIndex = 1; itemIndex <= itemsBound.Count; itemIndex++)
                {
                    var idx = Orientation == ViewOrientation.Portrait ? itemIndex : 1;
                    var jdx = Orientation == ViewOrientation.Portrait ? 1 : itemIndex;
                    ((Range)categoryRange.Cells[idx, jdx]).ID = (itemIndex - 1).ToString(CultureInfo.InvariantCulture);
                }
            });
        }

        private Range GetCategoryRange(Binding binding)
        {
            switch (Orientation)
            {
                case ViewOrientation.Portrait:
                    return binding.MakeRange(0, itemsBound.Count, 0, 1);
                case ViewOrientation.Landscape:
                    return binding.MakeRange(0, 1, 0, itemsBound.Count);
                default:
                    return null;
            }
        }

        private RangeObjects GetRangeObjects(Range target)
        {
            RangeObjects result;
            result.Items = null;
            result.Bindings = null;
            result.Intersection = null;
            if (itemsBound == null || itemsBound.Count == 0)
                return result;

            var first = Bindings.First();
            var whole = GetWholeRange(first);
            result.Intersection = target.Application.Intersect(whole, target);
            if (result.Intersection == null)
                return result;

            var items = new List<object>();
            var isPortrait = Orientation == ViewOrientation.Portrait;
            foreach (Range item in isPortrait ? result.Intersection.Rows : result.Intersection.Columns)
            {
                if (isPortrait)
                {
                    var idCell = (Range)item.Worksheet.Cells[item.Row, first.StartCell.Column];
                    items.Add(itemsBound[IsOrderingAllowed ? int.Parse(idCell.ID) : item.Row - first.StartCell.Row]);
                }
                else
                {
                    var idCell = (Range)item.Worksheet.Cells[first.StartCell.Row, item.Column];
                    items.Add(itemsBound[IsOrderingAllowed ? int.Parse(idCell.ID) : item.Column - first.StartCell.Column]);
                }
            }

            result.Items = items;
            var skipItems = isPortrait ? result.Intersection.Column - first.StartCell.Column : result.Intersection.Row - first.StartCell.Row;
            var takeItems = isPortrait ? result.Intersection.Columns.Count : result.Intersection.Rows.Count;
            result.Bindings = Bindings.Skip(skipItems).Take(takeItems).ToList();

            return result;
        }

        private Range GetWholeRange(Binding binding)
        {
            switch (Orientation)
            {
                case ViewOrientation.Portrait:
                    return binding.MakeRange(0, itemsBound.Count, 0, Bindings.Count());
                case ViewOrientation.Landscape:
                    return binding.MakeRange(0, Bindings.Count(), 0, itemsBound.Count);
            }

            return null;
        }

        private void HookItemsEvents(bool toHook)
        {
            if (itemsBound == null)
                return;

            foreach (var item in itemsBound)
            {
                var itemNotify = item as INotifyPropertyChanged;
                if (itemNotify != null)
                {
                    if (toHook)
                        itemNotify.PropertyChanged += ItemNotify_PropertyChanged;
                    else
                        itemNotify.PropertyChanged -= ItemNotify_PropertyChanged;
                }
            }
        }

        private void HookModelEvents()
        {
            enumerable = Model as IEnumerable;
            if (enumerable == null && Model != null)
                throw new Exception(string.Format(Resource.ErrorNoIEnumuerable, Model.GetType().FullName, Name));

            UnhookModelEvents();

            notifyPropertyChanged = Model as INotifyPropertyChanged;
            if (notifyPropertyChanged != null)
                notifyPropertyChanged.PropertyChanged += NotifyPropertyChanged_PropertyChanged;

            notifyCollectionChanged = Model as INotifyCollectionChanged;
            if (notifyCollectionChanged != null)
                notifyCollectionChanged.CollectionChanged += NotifyCollectionChanged_CollectionChanged;
        }

        private void HookViewEvents(bool isHook)
        {
            var sheet = (Sheet)Parent;
            if (isHook)
            {
                sheet.Underlying.SelectionChange += Underlying_SelectionChange;
                sheet.Underlying.Change += Underlying_Change;
            }
            else
            {
                sheet.Underlying.SelectionChange -= Underlying_SelectionChange;
                sheet.Underlying.Change -= Underlying_Change;
            }
        }

        private void OneWayToSource()
        {
            if (Model == null)
                return;

            var oneways = Bindings.Where(x => (x.Mode == ModeType.OneWayToSource));
            var isPortrait = Orientation == ViewOrientation.Portrait;
            foreach (var oneway in oneways)
            {
                var ranage = isPortrait ? oneway.MakeRange(0, itemsBound.Count, 0, 1) : oneway.MakeRange(0, 1, 0, itemsBound.Count);
                UpdateObjects(ranage);
            }
        }

        private void ItemNotify_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            UpdateCell(sender, e.PropertyName ?? "*");
        }

        private void RestoreCategoryIds(Range selection)
        {
            if (categoryIds.Count == 0)
                return;

            var first = Bindings.First();
            var categoryRange = GetCategoryRange(first);
            var section = selection.Application.Intersect(categoryRange, selection);
            if (section != null)
            {
                Parent.ExecuteProtected(() =>
                {
                    var isPortrait = Orientation == ViewOrientation.Portrait;
                    for (var itemIndex = 1; itemIndex <= (isPortrait ? section.Rows.Count : section.Columns.Count); itemIndex++)
                    {
                        var idx = isPortrait ? itemIndex : 1;
                        var jdx = isPortrait ? 1 : itemIndex;
                        ((Range)section.Cells[idx, jdx]).ID = categoryIds[itemIndex - 1];
                    }
                });
            }
        }

        private void SaveCategoryIds(Range selection)
        {
            categoryIds.Clear();
            var first = Bindings.First();
            var categoryRange = GetCategoryRange(first);
            var section = selection.Application.Intersect(categoryRange, selection);
            if (section != null)
            {
                var isPortrait = Orientation == ViewOrientation.Portrait;
                foreach (Range item in isPortrait ? section.Rows : section.Columns)
                    categoryIds.Add(((Range)item.Cells[1, 1]).ID);
            }

            RestoreCategoryIds(selection);
        }

        /// <summary>
        /// Sets visibiliy for a single column (portrait table) or for a single row (landscape table)
        /// </summary>
        private void SetCategoryVisibility(Binding binding)
        {
            switch (Orientation)
            {
                case ViewOrientation.Portrait:
                    {
                        // set if different, otherwise the column is going to flicker
                        if (Convert.ToBoolean(binding.StartCell.EntireColumn.Hidden) != !binding.Visible)
                            Parent.ExecuteProtected(() => binding.StartCell.EntireColumn.Hidden = !binding.Visible);
                        break;
                    }

                case ViewOrientation.Landscape:
                    {
                        // set if different, otherwise the row is going to flicker
                        if (Convert.ToBoolean(binding.StartCell.EntireRow.Hidden) != !binding.Visible)
                            Parent.ExecuteProtected(() => binding.StartCell.EntireRow.Hidden = !binding.Visible);
                        break;
                    }
            }
        }

        private void Underlying_Change(Range target)
        {
            RestoreCategoryIds(target);
            UpdateRangeObjects(target);
        }

        private void Underlying_SelectionChange(Range target)
        {
            if (IsOrderingAllowed) SaveCategoryIds(target);
            var rangeObjs = GetRangeObjects(target);
            SelectedItems.Clear();
            SelectedBindings.Clear();
            if (rangeObjs.Items != null)
            {
                SelectedItems.AddRange(rangeObjs.Items);
                SelectedBindings.AddRange(rangeObjs.Bindings);
                OnSelectionChanged(rangeObjs.Items, rangeObjs.Bindings);
            }
        }

        private void UnhookModelEvents()
        {
            if (notifyPropertyChanged != null)
                notifyPropertyChanged.PropertyChanged -= NotifyPropertyChanged_PropertyChanged;
            if (notifyCollectionChanged != null)
                notifyCollectionChanged.CollectionChanged -= NotifyCollectionChanged_CollectionChanged;
        }

        private void UpdateCell(object model, string propertyName)
        {
            var objectId = itemsBound.IndexOf(model);
            if (objectId < 0)
                return;

            var match = Bindings.FirstOrDefault(x => x.Path == propertyName);
            if (match != null)
            {
                UpdateCell(match, model, objectId);
            }
            else if (propertyName == "*")
            {
                foreach (var binding in Bindings)
                    UpdateCell(binding, model, objectId);
            }
        }

        private void UpdateCell(Binding binding, object model, int objectId)
        {
            ExecuteBinding(() =>
            {
                var value = ObjectBinding.GetPropertyValue(model, binding);
                switch (Orientation)
                {
                    case ViewOrientation.Portrait:
                        RangeUpdator.Instance.Update(
                            binding.StartCell,
                            IsOrderingAllowed ? Bindings.First().StartCell : null,
                            itemsBound.Count,
                            objectId.ToString(CultureInfo.InvariantCulture),
                            1,
                            0,
                            1,
                            value);
                        break;
                    case ViewOrientation.Landscape:
                        RangeUpdator.Instance.Update(
                            binding.StartCell,
                            0,
                            1,
                            IsOrderingAllowed ? Bindings.First().StartCell : null,
                            itemsBound.Count,
                            objectId.ToString(CultureInfo.InvariantCulture),
                            1,
                            value);
                        break;
                }
            });
        }

        private bool UpdateObjects(Range target)
        {
            var updated = false;
            ExecuteBinding(() =>
            {
                var rangeObjs = GetRangeObjects(target);
                if (rangeObjs.Items != null)
                {
                    updated = true;
                    if (UpdateObjects(rangeObjs) > 0)
                        OnObjectChanged(rangeObjs.Items, null);
                }
            });
            return updated;
        }

        private int UpdateObjects(RangeObjects rangeItems)
        {
            var first = Bindings.First();
            var target = rangeItems.Intersection;
            var itemOffset = Orientation == ViewOrientation.Portrait ? target.Row - first.StartCell.Row : target.Column - first.StartCell.Column;
            var skipCategories = Orientation == ViewOrientation.Portrait ? target.Column - first.StartCell.Column : target.Row - first.StartCell.Row;
            var takeCategories = Orientation == ViewOrientation.Portrait ? target.Columns.Count : target.Rows.Count;
            var toSource = Bindings.Skip(skipCategories).Take(takeCategories)
                .Where(x => (x.Mode == ModeType.TwoWay || x.Mode == ModeType.OneWayToSource)).ToList();

            var updated = 0;
            foreach (var model in rangeItems.Items)
            {
                if (itemOffset >= MaxItemsToBind)
                    break;
                updated += toSource.Count(binding => UpdateObject(binding, itemOffset, model, rangeItems.Intersection));
                itemOffset++;
            }

            return updated;
        }

        private void UpdateRangeObjects(Range target)
        {
            var from = target;
            while (target != null)
            {
                if (UpdateObjects(target))
                    break;

                Range dependents = null;
                var ltarget = target;
                ActionExtensions.Try(() => dependents = ltarget.DirectDependents);
                target = dependents;

                // break if circular referenced
                if (target != null && from.Application.Intersect(from, target) != null)
                    break;
            }
        }

        private void UpdateView(bool clearifnull)
        {
            ExecuteBinding(
                () =>
                {
                    HookViewEvents(false);
                    HookModelEvents();
                    HookItemsEvents(false);
                    UpdateViewEx(clearifnull);
                },
                () =>
                {
                    if (Model != null)
                    {
                        HookViewEvents(true);
                        HookModelEvents();
                        HookItemsEvents(true);
                    }
                });
        }

        private void UpdateViewEx(bool clearifnull)
        {
            Dictionary<Binding, List<object>> bindingValues = null;
            var numberItemsBound = itemsBound == null ? 0 : itemsBound.Count;
            var toView = from binding in Bindings
                         where binding.Mode != ModeType.OneWayToSource
                         select binding;

            if (enumerable != null)
            {
                bindingValues = new Dictionary<Binding, List<object>>();
                foreach (var binding in toView)
                    bindingValues[binding] = new List<object>();

                itemsBound = enumerable.ToList();
                if (itemsBound.Count > MaxItemsToBind)
                    itemsBound = itemsBound.Take(MaxItemsToBind).ToList();

                foreach (var item in itemsBound)
                    foreach (var binding in toView)
                        bindingValues[binding].Add(ObjectBinding.GetPropertyValue(item, binding));
            }
            else
            {
                itemsBound = null;
            }

            var newItems = itemsBound == null ? 0 : itemsBound.Count;
            var groupBindings = GroupBindings(toView);
            if (numberItemsBound != newItems && (itemsBound != null || clearifnull))
            {
                if (numberItemsBound > 0)
                {
                    ClearView(groupBindings, numberItemsBound);
                    UnbindValidationLists(numberItemsBound);
                }
            }

            if (newItems > 0)
            {
                if (IsOrderingAllowed) AssignCategoryIds();
                UpdateView(groupBindings, bindingValues, newItems);
                BindValidationLists(newItems);
            }
        }

        private void NotifyCollectionChanged_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            UpdateView(true);
        }

        private void NotifyPropertyChanged_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
        }

        private void DeriveOrientation()
        {
            var origin = Bindings.First().StartCell;
            if (Bindings.All(x => x.StartCell.Row == origin.Row))
                Orientation = ViewOrientation.Portrait;
            else if (Bindings.All(x => x.StartCell.Column == origin.Column))
                Orientation = ViewOrientation.Landscape;
            else
            {
                ExecuteBinding(() =>
                {
                    throw new InvalidOperationException(string.Format(Resource.ErrorInvalidTableOrientation, Name));
                });
            }
        }

        #endregion Methods

        #region Nested Types

        private struct RangeObjects
        {
            #region Fields

            public IEnumerable<Binding> Bindings;
            public Range Intersection;
            public IEnumerable<object> Items;

            #endregion Fields
        }

        #endregion Nested Types
    }
}