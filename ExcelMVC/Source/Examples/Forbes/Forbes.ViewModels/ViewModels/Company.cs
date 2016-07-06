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

namespace Forbes.Application.ViewModels
{
    using System.ComponentModel;
    using Models;

    public class Company : INotifyPropertyChanged
    {
        #region Constructors

        public Company()
        {
            Model = new CompanyModel();
        }

        #endregion Constructors

        #region Events

        public event PropertyChangedEventHandler PropertyChanged = delegate { };

        #endregion Events

        #region Properties

        public double Assets
        {
            get { return Model.Assets; }
            set { Model.Assets = value; }
        }

        public string Country
        {
            get { return Model.Country; }
            set { Model.Country = value; }
        }

        public string Industry
        {
            get { return Model.Industry; }
            set { Model.Industry = value; }
        }

        public bool Listed
        {
            get { return Model.Listed; }
            set { Model.Listed = value; }
        }

        public double MarketValue
        {
            get { return Model.MarketValue; }
            set { Model.MarketValue = value; }
        }

        public CompanyModel Model
        {
            get; set;
        }

        public string Name
        {
            get { return Model.Name; }
            set { Model.Name = value; }
        }

        public double Profits
        {
            get { return Model.Profits; }
            set { Model.Profits = value; }
        }

        public int Rank
        {
            get { return Model.Rank; }
            set { Model.Rank = value; }
        }

        public double Sales
        {
            get { return Model.Sales; }
            set { Model.Sales = value; }
        }

        #endregion Properties

        #region Methods

        public void RaiseChanged(string propertyName)
        {
            PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
        }

        public void RaiseChanged()
        {
            RaiseChanged("Name");
            RaiseChanged("Industry");
            RaiseChanged("Country");
            RaiseChanged("MarketValue");
            RaiseChanged("Profits");
            RaiseChanged("Sales");
            RaiseChanged("Assets");
            RaiseChanged("Rank");
            RaiseChanged("Listed");
        }

        #endregion Methods
    }
}