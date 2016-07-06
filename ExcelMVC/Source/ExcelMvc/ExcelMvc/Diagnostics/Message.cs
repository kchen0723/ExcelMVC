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

namespace ExcelMvc.Diagnostics
{
    using System;
    using System.ComponentModel;

    internal class Message : INotifyPropertyChanged
    {
        public Message()
        {
            LineLimit = 2000;
        }

        public event PropertyChangedEventHandler PropertyChanged = delegate { }; 

        public string Error { get; private set; }

        public string Info { get; private set; }

        public int LineLimit { get; set; }

        public int ErrorLines { get; set; }

        public int InfoLines { get;  set; }

        public void Clear()
        {
            Error = null;
            Info = null;
            ErrorLines = 0;
            InfoLines = 0;
            RaiseErrorChanged();
            RaiseInfoChanged();
        }

        public void AddErrorLine(Exception ex)
        {
            AddErrorLine(string.Format("{0} [{1}]", ex.Message, ex.StackTrace));
        }

        public void AddErrorLine(string message)
        {
            if (ErrorLines > LineLimit)
            {
                ErrorLines = 0;
                Error = null;
            }

            Error = (Error ?? string.Empty) + message + Environment.NewLine;
            RaiseErrorChanged();
        }

        public void AddInfoLine(string message)
        {
            if (InfoLines > LineLimit)
            {
                InfoLines = 0;
                Info = null;
            }

            Info = (Info ?? string.Empty) + message + Environment.NewLine;
            RaiseInfoChanged();
        }

        private void RaiseErrorChanged()
        {
            PropertyChanged(this, new PropertyChangedEventArgs("Error"));
        }

        private void RaiseInfoChanged()
        {
            PropertyChanged(this, new PropertyChangedEventArgs("Info"));
        }
    }
}
