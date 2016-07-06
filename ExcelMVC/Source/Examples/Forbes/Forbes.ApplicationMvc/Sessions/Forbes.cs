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

namespace Forbes.Application.Sessions
{
    using System;
    using System.Windows.Forms;
    using ExcelMvc.Bindings;
    using ExcelMvc.Views;
    using Models;
    using View = ExcelMvc.Views.View;

    internal class Forbes
    {
        #region Constructors

        public Forbes(View view)
        {
            view.HookBindingFailed(View_BindingFailed, true);

            Tests = new CommandTests((Sheet)view.Find(ViewType.Sheet, "Tests"));

            var settingsForms = (ExcelMvc.Views.Form)view.Find(ViewType.Form, "Settings");
            var settingsModel = new Settings();
            settingsForms.Model = settingsModel;

            view.Find("Table.AppSettings").Model = new AppConfigSettings();

            // portrait
            var parent = view.Find(ViewType.Sheet, "Forbes");
            ForbesTest = new Forbes2000(view, parent, settingsModel, "Company", "Company");

            // landscape/transposed
            parent = view.Find(ViewType.Sheet, "Forbes_transposed");
            ForbesTestTransposed = new Forbes2000(view, parent, settingsModel, "CompanyTransposed", "CompanyTransposed");
        }

        #endregion Constructors

        #region Properties

        private Forbes2000 ForbesTest
        {
            get; set;
        }

        private Forbes2000 ForbesTestTransposed
        {
            get; set;
        }

        private CommandTests Tests
        {
            get; set;
        }

        #endregion Properties

        #region Methods

        private void View_BindingFailed(object sender, BindingFailedEventArgs args)
        {
            DisplayException(args.Exception, args.View.Name);
        }

        private void DisplayException(Exception ex, string title)
        {
            MessageBox.Show(ex.Message + ex.StackTrace, title, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        #endregion Methods
    }
}