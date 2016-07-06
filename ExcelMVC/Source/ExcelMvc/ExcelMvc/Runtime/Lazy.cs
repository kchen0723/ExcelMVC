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
namespace ExcelMvc.Runtime
{
    using System;

    /// <summary>
    /// Provides support for lazy initialization.
    /// </summary>
    /// <typeparam name="T">Specifies the type of object that is being lazily initialized.</typeparam>
    public sealed class Lazy<T>
    {
        #region Fields

        private readonly Func<T> createValue;
        private readonly object padlock = new object();

        private bool isValueCreated;
        private T value;

        #endregion Fields

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the Lazy{T} class.
        /// </summary>
        /// <param name="createValue">The delegate that produces the value when it is needed.</param>
        public Lazy(Func<T> createValue)
        {
            if (createValue == null) throw new ArgumentNullException("createValue");

            this.createValue = createValue;
        }

        #endregion Constructors

        #region Properties

        /// <summary>
        /// Gets a value that indicates whether a value has been created for this Lazy{T} instance.
        /// </summary>
        public bool IsValueCreated
        {
            get
            {
                lock (padlock)
                {
                    return isValueCreated;
                }
            }
        }

        /// <summary>
        /// Gets the lazily initialized value of the current Lazy{T} instance.
        /// </summary>
        public T Value
        {
            get
            {
                if (!isValueCreated)
                {
                    lock (padlock)
                    {
                        if (!isValueCreated)
                        {
                            value = createValue();
                            isValueCreated = true;
                        }
                    }
                }

                return value;
            }
        }

        #endregion Properties

        #region Methods

        /// <summary>
        /// Creates and returns a string representation of the Lazy{T}.Value.
        /// </summary>
        /// <returns>The string representation of the Lazy{T}.Value property.</returns>
        public override string ToString()
        {
            return Value.ToString();
        }

        #endregion Methods
    }
}