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
    using System.Collections.Generic;
    using System.Runtime.CompilerServices;
    using Extensions;
    using Views;

    /// <summary>
    /// Posts and handles asynchronous actions
    /// </summary>
    internal static class AsyncActions
    {
        #region Fields

        private class Item
        {
            public Action<object> Action { get; set; }
            public object State  { get; set; }
        }

        private static AsyncWindow Context { get; set; }
        private static Queue<Item> Actions { get; set; }
        private static Queue<Item> Macros { get; set; }
    
        #endregion Fields

        #region Constructors

        static AsyncActions()
        {
            Context = new AsyncWindow();
            Actions = new Queue<Item>();
            Macros = new Queue<Item>();
            Context.AsyncActionReceived += MainWindow_AsyncActionReceived;
            Context.AsyncMacroReceived += MainWindow_AsyncMacroReceived;
        }

        static void MainWindow_AsyncActionReceived(object sender, EventArgs args)
        {
            ActionExtensions.Try(() => Execute(false));
        }

        static void MainWindow_AsyncMacroReceived(object sender, EventArgs args)
        {
            ActionExtensions.Try(() => App.Instance.Underlying.Run("ExcelMvcRun"));
        }

        #endregion Constructors

        #region Properties

        /// <summary>
        /// Gets the number of outstanding actions
        /// </summary>
        /// <returns>Number of items</returns>
        [MethodImpl(MethodImplOptions.Synchronized)]
        public static int GetActionDepth()
        {
            return Actions.Count;
        }

        /// <summary>
        /// Gets the number of outstanding macros
        /// </summary>
        /// <returns>Number of items</returns>
        [MethodImpl(MethodImplOptions.Synchronized)]
        public static int GetMacroDepth()
        {
            return Actions.Count;
        }

        #endregion Properties

        #region Methods

        /// <summary>
        /// Initialise class static states
        /// </summary>
        [MethodImpl(MethodImplOptions.Synchronized)]
        public static void Initialise()
        {
            // nothing needs to be done, as the static contructor will be called
        }

        /// <summary>
        /// Posts an Async action
        /// </summary>
        /// <param name="action">Action to be executed</param>
        /// <param name="state">State object</param>
        /// <param name="exectueAsMacro">Execute as a macro</param>
        /// <param name="pumpMilliseconds">Pumping message</param>
        [MethodImpl(MethodImplOptions.Synchronized)]
        public static void Post(Action<object> action, object state, bool exectueAsMacro, int pumpMilliseconds = 0)
        {
            var item = new Item {Action = action, State = state};
            if (exectueAsMacro)
            {
                Macros.Enqueue(item);
                Context.PostAsyncMacroMessage(pumpMilliseconds);
            }
            else
            {
                Actions.Enqueue(item);
                Context.PostAsyncActionMessage();
            }
        }

        /// <summary>
        /// Executes the next action in the queue
        /// </summary>
        /// <param name="exectueMacro">Execute the next macro</param>
        [MethodImpl(MethodImplOptions.Synchronized)]
        public static void Execute(bool exectueMacro)
        {
            var item = exectueMacro
                ? (Macros.Count == 0 ? null : Macros.Dequeue())
                : Actions.Count == 0 ? null : Actions.Dequeue();
            if (item != null)
                item.Action(item.State);
        }

        #endregion Methods
    }
}