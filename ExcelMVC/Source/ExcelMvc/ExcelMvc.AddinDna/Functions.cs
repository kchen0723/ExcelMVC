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

namespace ExcelMvc.AddinDna
{
    using ExcelDna.Integration;

    using Extensions;
    using Runtime;

    public static class Functions
    {
        #region Methods

        /// <summary>
        /// Attaches the current Excel session to ExcelMVC
        /// </summary>
        /// <returns></returns>
        [ExcelFunction(Description = "Attach Excel to ExcelMvc", Category = "ExcelMvc", IsVolatile = false, Name = "ExcelMvcAttach")]
        public static object ExcelMvcAttach()
        {
            return ActionExtensions.Wrap(() => Interface.Attach(null)) ?? (object)true;
        }

        /// <summary>
        /// Detaches the current Excel session from ExcelMVC
        /// </summary>
        /// <returns></returns>
        [ExcelFunction(Description = "Detach Excel from ExcelMvc", Category = "ExcelMvc", IsVolatile = false, Name = "ExcelMvcDetach")]
        public static object ExcelMvcDetach()
        {
            return ActionExtensions.Wrap(() => Interface.Detach()) ?? (object)true;
        }

        /// <summary>
        /// Shows the ExcelMVC status window
        /// </summary>
        /// <returns></returns>
        [ExcelFunction(Description = "Shows the ExcelMvc window", Category = "ExcelMvc", IsVolatile = false, Name = "ExcelMvcShow")]
        public static object ExcelMvcShow()
        {
            return ActionExtensions.Wrap(() => Interface.Show()) ?? (object)true;
        }

        /// <summary>
        /// Hides the ExcelMVC status window
        /// </summary>
        /// <returns></returns>
        [ExcelFunction(Description = "Hides the ExcelMvc window", Category = "ExcelMvc", IsVolatile = false, Name = "ExcelMvcHide")]
        public static object ExcelMvcHide()
        {
            return ActionExtensions.Wrap(() => Interface.Hide()) ?? (object)true;
        }

        /// <summary>
        /// Fires the Clicked event on the current command
        /// </summary>
        /// <returns></returns>
        [ExcelFunction(Description = "Called by a command", Category = "ExcelMvc", IsVolatile = false, Name = "ExcelMvcRunCommandAction", IsHidden = true)]
        public static void ExcelMvcRunCommandAction()
        {
            Interface.FireClicked();
        }

        /// <summary>
        /// Detaches the current Excel session from ExcelMVC
        /// </summary>
        /// <returns></returns>
        [ExcelFunction(Description = "Runs the next action in the async queue", Category = "ExcelMvc", IsVolatile = false, Name = "ExcelMvcRun", IsHidden = true)]
        public static object ExcelMvcRun()
        {
            return ActionExtensions.Wrap(() => Interface.Run()) ?? (object)true;
        }

        #endregion Methods
    }
}