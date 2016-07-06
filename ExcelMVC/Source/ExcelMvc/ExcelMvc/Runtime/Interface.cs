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

namespace ExcelMvc.Runtime
{
    using System;
    using System.Windows.Forms;

    using Diagnostics;
    using Extensions;
    using Views;

    /// <summary>
    /// Defines interfacing methods for ExcelMvc
    /// </summary>
    public static class Interface
    {
        #region Methods

        /// <summary>
        /// Attaches the current Excel session to ExcelMvc
        /// </summary>
        /// <param name="app">Excel Application object</param>
        /// <returns>error string, null if success</returns>
        public static string Attach(object app)
        {
            var status = ActionExtensions.Try(() => App.Instance.Attach(app));
            return TestStauts(status);
        }

        /// <summary>
        /// Attaches the current Excel session to ExcelMvc
        /// </summary>
        /// <returns>error string, null if success</returns>
        public static string Attach()
        {
            var status = ActionExtensions.Try(() => App.Instance.Attach(null));
            return TestStauts(status);
        }

        /// <summary>
        /// Detaches the current Excel session from ExcelMvc 
        /// </summary>
        /// <returns>error string, null if success</returns>
        public static string Detach()
        {
            var status = ActionExtensions.Try(() => App.Instance.Detach());
            return TestStauts(status);
        }

        /// <summary>
        /// Fires clicked event for the caller
        /// </summary>
        /// <returns>error string, null if success</returns>
        public static string FireClicked()
        {
            var status = ActionExtensions.Try(() => App.Instance.FireClicked());
            return TestStauts(status);
        }

        /// <summary>
        /// Tests status
        /// </summary>
        /// <param name="status">Exception object</param>
        /// <returns>error string, null if success</returns>
        public static string TestStauts(Exception status)
        {
            string result = null;
            if (status != null)
            {
                result = status.Message + Environment.NewLine + status.StackTrace;
                MessageWindow.AddErrorLine(status);
                MessageBox.Show(result, typeof(Interface).Namespace, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return result;
        }

        /// <summary>
        /// Shows the ExcelMvc window
        /// </summary>
        /// <returns>error string, null if success</returns>
        public static string Show()
        {
            MessageWindow.ShowInstance();
            return null;
        }

        /// <summary>
        /// Hides the ExcelMvc window
        /// </summary>
        /// <returns>error string, null if success</returns>
        public static string Hide()
        {
            MessageWindow.HideInstance();
            return null;
        }

        /// <summary>
        /// Runs the next action in the Async queue
        /// </summary>
        /// <returns>error string, null if success</returns>
        public static string Run()
        {
            AsyncActions.Execute(true);
            return null;
        }
        #endregion Methods
    }
}