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

namespace ExcelMvc.Controls
{
    using System;
    using System.Windows.Input;

    using Views;

    /// <summary>
    /// Defines an abstract base class for a Command
    /// </summary>
    public abstract class Command : IDisposable
    {
        #region Fields

        private ICommand model;

        #endregion Fields

        #region Constructors

        /// <summary>
        /// Constructs an instance
        /// </summary>
        /// <param name="host">Command host</param>
        /// <param name="name">Command name</param>
        protected Command(View host, string name)
        {
            Host = host;
            Name = name;
        }

        #endregion Constructors

        #region Events

        /// <summary>
        /// Occurs when the command is clicked
        /// </summary>
        public event ClickedHandler Clicked = delegate { };

        #endregion Events

        #region Properties

        /// <summary>
        /// Caption of the command
        /// </summary>
        public abstract string Caption
        {
            get;
            set;
        }

        /// <summary>
        /// Caption to be swaped when the command is clicked
        /// </summary>
        public string ClickedCaption
        {
            get;
            set;
        }

        /// <summary>
        /// Gets the host view 
        /// </summary>
        public View Host
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets and sets the Enabled state
        /// </summary>
        public abstract bool IsEnabled
        {
            get;
            set;
        }

        /// <summary>
        /// Name of the command
        /// </summary>
        public string Name
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets and sets the command value
        /// </summary>
        public abstract object Value
        {
            get;
            set;
        }

        /// <summary>
        /// Gets and sets te command model
        /// </summary>
        public ICommand Model
        {
            get
            {
                return model;
            }

            set
            {
                if (model != null)
                    model.CanExecuteChanged -= Model_CanExecuteChanged;

                model = value;

                if (model != null)
                    model.CanExecuteChanged += Model_CanExecuteChanged;
            }
        }

        /// <summary>
        ///  Gets and sets an application specific object
        /// </summary>
        public object State
        {
            get;
            set;
        }

        #endregion Properties

        #region Methods

        /// <summary>
        /// Disposes resources
        /// </summary>
        public virtual void Dispose()
        {
            Host = null;
            if (Model != null)
                Model.CanExecuteChanged -= Model_CanExecuteChanged;
        }

        /// <summary>
        /// Fires the Clicked event
        /// </summary>
        public virtual void FireClicked()
        {
            Host.ExecuteBinding(() =>
            {
                var args = new CommandEventArgs { Source = this };
                Clicked(this, args);
                if (Model != null)
                    Model.Execute(args);

                if (ClickedCaption != null)
                {
                    var current = Caption;
                    Caption = ClickedCaption;
                    ClickedCaption = current;
                }
            });
        }

        private void Model_CanExecuteChanged(object sender, EventArgs e)
        {
            Host.ExecuteBinding(() =>
            {
                IsEnabled = Model.CanExecute(new CommandEventArgs { Source = this });
            });
        }

        #endregion Methods
    }
}