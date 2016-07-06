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
    using System.Windows.Forms;
    using ExcelMvc.Bindings;
    using ExcelMvc.Runtime;
    using ExcelMvc.Views;

    public class Session : ISession
    {
        #region Fields

        private const string ViewName = "Forbes2000";

        #endregion Fields

        #region Constructors

        public Session()
        {
            App.Instance.Opening += Book_Opening;
            App.Instance.Opened += Book_Opened;
            App.Instance.Closing += Book_Closing;
            App.Instance.Closed += Book_Closed;
        }

        #endregion Constructors

        #region Methods

        public void Dispose()
        {
        }

        private void Book_Closed(object sender, ViewEventArgs args)
        {
            if (args.View.Id == ViewName)
            {
                args.Accept();
            }
        }

        private void Book_Closing(object sender, ViewEventArgs args)
        {
            if (args.View.Id == ViewName)
                args.Accept();
        }

        private void Book_Opened(object sender, ViewEventArgs args)
        {
            if (args.View.Id == ViewName)
            {
                args.Accept();
                args.View.Model = new Forbes(args.View);
            }
        }

        private void Book_Opening(object sender, ViewEventArgs args)
        {
            // accept if the book being opened is "Forbes2000", whose view id is
            // defined by the Custom Document Propety named "ExcelMvc".
            if (args.View.Id == ViewName)
            {
                args.Accept();
                args.View.BindingFailed += View_BindingFailed;
            }
        }

        private void View_BindingFailed(object sender, BindingFailedEventArgs args)
        {
            MessageBox.Show(args.Exception.Message, args.View.Name, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        #endregion Methods
    }
}