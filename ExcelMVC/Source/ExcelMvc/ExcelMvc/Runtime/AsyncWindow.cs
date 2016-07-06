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
    using System.Runtime.InteropServices;
    using System.Windows.Forms;
    using Views;

    /// <summary>
    /// Subclasses a window
    /// </summary>
    internal sealed class AsyncWindow : NativeWindow
    {
        #region Fields

        private static readonly uint AsyncActionMessage;
        private static readonly uint AsyncMacroMessage;
        private static readonly uint WindowsTimerMessage;
        private static int TimerId;
    
        #endregion Fields

        #region Constructors

        static AsyncWindow()
        {
            AsyncActionMessage = RegisterWindowMessage("__ExcelMvcAsyncAction__");
            AsyncMacroMessage = RegisterWindowMessage("__ExcelMvcAsyncMacro__");
            WindowsTimerMessage = 0x113;
            TimerId = 0;
        }

        /// <summary>
        /// Intialises an instance of Window
        /// </summary>
        public AsyncWindow()
        {
            var cp = new CreateParams {Parent = new IntPtr(App.Instance.Underlying.Application.Hwnd)};
            CreateHandle(cp);
        }

        #endregion Constructors

        #region Delegates

        /// <summary>
        /// Handler for a AsyncAction
        /// </summary>
        /// <param name="sender">Event sender</param>
        /// <param name="args">EventArgs</param>
        public delegate void AsyncActionReceivedHandler(object sender, EventArgs args);

        /// <summary>
        /// Handler for a AsyncAction
        /// </summary>
        /// <param name="sender">Event sender</param>
        /// <param name="args">EventArgs</param>
        public delegate void AsyncMacroReceivedHandler(object sender, EventArgs args);

        #endregion Delegates

        #region Events

        /// <summary>
        /// Occurs when an async action message is received
        /// </summary>
        public event AsyncActionReceivedHandler AsyncActionReceived = delegate { };

        /// <summary>
        /// Occurs when an async macro message is received
        /// </summary>
        public event AsyncMacroReceivedHandler AsyncMacroReceived = delegate { };

        #endregion Events

        #region Methods
        
        /// <summary>
        /// Posts an async action message
        /// </summary>
        public void PostAsyncActionMessage()
        {
            PostAsyncMessage(AsyncActionMessage, 0);
        }

        /// <summary>
        /// Posts an async macro message
        /// </summary>
        /// <param name="pumpMilliseconds">Pumping messages</param>
        public void PostAsyncMacroMessage(int pumpMilliseconds = 0)
        {
            PostAsyncMessage(AsyncMacroMessage, pumpMilliseconds);
        }

        public void PostAsyncMessage(uint message, int pumpMilliseconds)
        {
            if (pumpMilliseconds > 0)
            {
                TimerId = TimerId == int.MaxValue ? 0 : TimerId + 1;
                SetTimer(Handle.ToInt32(), TimerId, pumpMilliseconds, IntPtr.Zero);
            }
            else
            {
                PostMessage(Handle, (int)message, 0, 0);
            }
        }

        /// <summary>
        /// Windows proc
        /// </summary>
        /// <param name="m">Message instance</param>
        protected override void WndProc(ref Message m)
        {
            if (m.Msg == AsyncActionMessage)
            {
                AsyncActionReceived(this, EventArgs.Empty);
                return;
            }
            if (m.Msg == AsyncMacroMessage)
            {
                AsyncMacroReceived(this, EventArgs.Empty);
                return;
            }
            if (m.Msg == WindowsTimerMessage)
            {
                KillTimer(Handle, (int) m.WParam);
                PostAsyncMacroMessage();
                return;
            }
            base.WndProc(ref m);
        }

        [DllImport("user32.dll")]
        private static extern int PostMessage(IntPtr hwnd, int msg, int wParam, int lParam);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern uint RegisterWindowMessage(string lpProcName);

        [DllImport("user32")]
        public static extern int SetTimer(int hwnd, int nIDEvent, int uElapse, IntPtr lpTimerFunc);

        [DllImport("user32")]
        private static extern int KillTimer(IntPtr hwnd, int nIDEvent);
  
        #endregion Methods
    }
}