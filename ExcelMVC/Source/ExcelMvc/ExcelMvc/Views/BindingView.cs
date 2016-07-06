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
    using System.Linq;
    using System.Reflection;

    using Bindings;
    using Controls;
    using Extensions;

    using Microsoft.Office.Core;
    using Microsoft.Office.Interop.Excel;

    using Shape = Microsoft.Office.Interop.Excel.Shape;

    /// <summary>
    /// Defines an abstract interface for Views
    /// </summary>
    public abstract class BindingView : View
    {
        #region Fields

        private readonly string name;

        #endregion Fields
        #region Constructors

        /// <summary>
        /// Initialises an instances of ExcelMvc.Views.Panel
        /// </summary>
        /// <param name="parent"></param>
        /// <param name="bindings">Bindings for the view</param>
        protected BindingView(View parent, IEnumerable<Binding> bindings)
        {
            Bindings = (from o in bindings orderby o.StartCell.Row, o.StartCell.Column select o).ToList();
            name = Bindings.First().Name;
            Parent = parent;
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
        /// Gets the bindings on the View
        /// </summary>
        public IEnumerable<Binding> Bindings
        {
            get;
            protected set;
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
            get { return name; }
        }

        internal ViewOrientation Orientation
        {
            get;
            set;
        }

        #endregion Properties

        #region Methods

        /// <summary>
        /// Fires the Clicked event
        /// </summary>
        public void FireClicked(object sender, CommandEventArgs args)
        {
            Clicked(sender, args);
        }

        /// <summary>
        /// Unbinds validation lists
        /// </summary>
        /// <param name="numberItems">Number of rows to unbind</param>
        protected void BindValidationLists(int numberItems)
        {
            foreach (var binding in Bindings.Where(binding => !string.IsNullOrEmpty(binding.ValidationList)))
            {
                if (IsBoolValidationList(binding.ValidationList))
                    BindCheckBoxes(binding, numberItems);
                else
                    BindValidationLists(binding, numberItems);
            }
        }

        /// <summary>
        /// Unbinds validation lists
        /// </summary>
        /// <param name="numberItems">Number of rows to unbind</param>
        protected void UnbindValidationLists(int numberItems)
        {
            foreach (var binding in Bindings.Where(binding => !string.IsNullOrEmpty(binding.ValidationList)))
            {
                if (IsBoolValidationList(binding.ValidationList))
                    UnbindCheckBoxes(binding, numberItems);
                else
                    UnbindValidationLists(binding, numberItems);
            }
        }

        private static bool IsBoolValidationList(string list)
        {
            return list.CompareOrdinalIgnoreCase("True/False") == 0;
        }

        private void BindCheckBoxes(Binding binding, int numberItems)
        {
            var worksheet = ((Sheet)Parent).Underlying;
            var boxes = worksheet.Shapes;
            var lbinding = binding;
            Parent.ExecuteProtected(() =>
            {
                // for (var idx = 0; idx < rows; idx++)
                {
                    var cell = lbinding.MakeRange(0, 1, 0, 1);
                    Shape box = boxes.AddShape(MsoAutoShapeType.msoShapeRoundedRectangle, (int)cell.Left + 2, (int)cell.Top + 2, 12, 12);
                    box.Fill.Visible = MsoTriState.msoFalse;
                    box.TextFrame2.TextRange.ParagraphFormat.Alignment = MsoParagraphAlignment.msoAlignCenter;
                    box.TextFrame2.VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle;
                    box.TextFrame2.TextRange.Characters.Text = "X";
                    box.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = 192;
                    box.Line.Weight = 1;
                    cell.Select();

                    var range = Orientation == ViewOrientation.Portrait ? lbinding.MakeRange(0, 1000, 0, 1) : lbinding.MakeRange(0, 1, 0, 1000);
                    cell.AutoFill(range);

                    // range.FillDown();
                }
            });
        }

        private void BindValidationLists(Binding binding, int numberItems)
        {
            var lbinding = binding;
            Parent.ExecuteProtected(() =>
            {
                var rangeCategory = Orientation == ViewOrientation.Portrait ? lbinding.MakeRange(0, numberItems, 0, 1) : lbinding.MakeRange(0, 1, 0, numberItems);
                rangeCategory.Validation.Delete();

                rangeCategory.Validation.Add(
                    XlDVType.xlValidateList,
                    XlDVAlertStyle.xlValidAlertStop,
                    XlFormatConditionOperator.xlBetween,
                    MarkValidationListFormula(lbinding.ValidationList));

                rangeCategory.Validation.IgnoreBlank = true;
                rangeCategory.Validation.InCellDropdown = true;
                rangeCategory.Validation.InputTitle = string.Empty;
                rangeCategory.Validation.ErrorTitle = string.Empty;
                rangeCategory.Validation.InputMessage = string.Empty;
                rangeCategory.Validation.ErrorMessage = string.Empty;
                rangeCategory.Validation.ShowInput = true;
                rangeCategory.Validation.ShowError = true;
            });
        }

        private string MarkValidationListFormula(string list)
        {
            Range range;
            if (list.Contains("["))
            {
                // "[book]sheet!start:[book]sheet!end"
                range = ((Sheet)Parent).Underlying.Application.Range[list];
            }
            else if (list.Contains("!"))
            {
                // "sheet!start:sheet!end", make it a book address
                var book = Parent.Parent.Name;
                list = string.Join(":", list.Split(':').Select(x => string.Format("[{0}]{1}", book, x)).ToArray());
                range = ((Sheet)Parent).Underlying.Application.Range[list];
            }
            else
            {
                // start:end
                range = ((Sheet)Parent).Underlying.Range[list];
            }

            // exclude trailing blank rows or trailing blank columns
            var value = (object[,])range.Value;
            for (var idx = value.GetUpperBound(0); idx >= value.GetLowerBound(0); idx--)
            {
                if (value[idx, 1] == null)
                    continue;
                var rows = idx - value.GetLowerBound(0) + 1;
                range = range.Worksheet.Range[range.Cells[1, 1], range.Cells[rows, 1]];
                break;
            }

            var address = range.Address[Missing.Value, Missing.Value, XlReferenceStyle.xlA1, true, Missing.Value];
            return string.Format("={0}", address);
        }

        private void UnbindCheckBoxes(Binding binding, int rows)
        {
            /*
            var worksheet = ((Sheet)Parent).Underlying;
            CheckBoxes boxes = worksheet.CheckBoxes();
            var lbinding = binding;
            Parent.ExecuteProtected(() =>
            {
                for (var idx = 0; idx < rows; idx++)
                {
                    var cell = lbinding.MakeRange(idx, 1, 0, 1);
                    var name = "_ExcelMvc_" + cell.Address;
                    CheckBox box = null;
                    ActionExtensions.Try(() => box = boxes.Item(name));
                    if (box != null)
                        box.Delete();
                }
            });*/
        }

        private void UnbindValidationLists(Binding binding, int numberItems)
        {
            var lbinding = binding;
            Parent.ExecuteProtected(() =>
            {
                var rangeCategory = Orientation == ViewOrientation.Portrait ? lbinding.MakeRange(0, numberItems, 0, 1) : lbinding.MakeRange(0, 1, 0, numberItems);
                rangeCategory.Validation.Delete();
            });
        }

        #endregion Methods
    }
}