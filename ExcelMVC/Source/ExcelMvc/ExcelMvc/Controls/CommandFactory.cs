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
    using System.Collections;
    using System.Collections.Generic;
    using System.Linq;

    using Diagnostics;
    using Extensions;
    using Microsoft.Office.Interop.Excel;
    using Views;

    /// <summary>
    /// Creates command on a sheet
    /// </summary>
    internal static class CommandFactory
    {
        #region Fields

        /// <summary>
        /// command names may be prefixed 
        /// </summary>
        private const string CommandPrefix = "ExcelMvc.";
        private const string CommandFullPrefix = "ExcelMvc.Command.";

        #endregion Fields

        #region Methods

        /// <summary>
        /// Creates commands on a sheet
        /// </summary>
        /// <param name="sheet">Sheet where commands are declared</param>
        /// <param name="host">View to host the commands</param>
        /// <param name="commands">Commands created</param>
        public static void Create(Worksheet sheet, View host, Dictionary<string, Command> commands)
        {
            var names = (from Comment item in sheet.Comments select item.Shape.Name).ToList();
            names.Sort();

            Create(sheet, host, (GroupObjects)sheet.GroupObjects(), names, commands);
            Create(sheet, host, (Buttons)sheet.Buttons(), names, commands);
            Create(sheet, host, (CheckBoxes)sheet.CheckBoxes(), names, commands);
            Create(sheet, host, (OptionButtons)sheet.OptionButtons(), names, commands);
            Create(sheet, host, (ListBoxes)sheet.ListBoxes(), names, commands);
            Create(sheet, host, (DropDowns)sheet.DropDowns(), names, commands);
            Create(sheet, host, (Spinners)sheet.Spinners(), names, commands);
            Create(sheet, host, sheet.Shapes, names, commands);

            foreach (var cmd in commands.Values)
                MessageWindow.AddInfoLine(string.Format(Resource.InfoCmdCreated, cmd.Name, cmd.GetType().Name,  cmd.Host.Name));
        }

        /// <summary>
        /// Removes ExcelMVC prefix from a command name
        /// </summary>
        /// <param name="name">Command name</param>
        /// <returns>Command name without prefix</returns>
        public static string RemovePrefix(string name)
        {
            if (StartsWithFullPrefix(name))
                return name.Substring(CommandFullPrefix.Length);
            return StartsWithPrefix(name) ? name.Substring(CommandPrefix.Length) : name;
        }

        private static void Create(Worksheet sheet, View host, IEnumerable items, List<string> names, Dictionary<string, Command> commands)
        {
            foreach (var item in items)
            {
               var button = item as Button;
               if (Create(button, () => button.Name, () => new CommandButton(host, button, RemovePrefix(button.Name)), commands, names))
                   continue;

               var cbox = item as CheckBox;
               if (Create(cbox, () => cbox.Name, () => new CommandCheckBox(host, cbox, RemovePrefix(cbox.Name)), commands, names))
                   continue;

               var option = item as OptionButton;
               if (Create(option, () => option.Name, () => new CommandOptionButton(host, option, RemovePrefix(option.Name)), commands, names))
                   continue;

               var lbox = item as ListBox;
               if (Create(lbox, () => lbox.Name, () => new CommandListBox(host, lbox, RemovePrefix(lbox.Name)), commands, names))
                   continue;

               var dbox = item as DropDown;
               if (Create(dbox, () => dbox.Name, () => new CommandDropDown(host, dbox, RemovePrefix(dbox.Name)), commands, names))
                   continue;

               var spin = item as Spinner;
               if (Create(spin, () => spin.Name, () => new CommandSpinner(host, spin, RemovePrefix(spin.Name)), commands, names))
                   continue;

               if (Create(sheet, host, item as GroupObject, commands, names))
                   continue;

                if (Create(host, item as Shape, commands, names))
                {
                }
            }
        }

        private static bool Create(object item, Func<string> getName, Func<Command> createCmd, Dictionary<string, Command> commands, List<string> names)
        {
            if (item == null)
                return false;

            var name = getName();
             int idx;
            if ((idx = names.BinarySearch(name)) >= 0 || !IsCreateable(name))
                return true;

            ActionExtensions.Try(() =>
            {
                commands[name] = createCmd();
                names.Insert(~idx, name);
            });

            return true;
        }

        private static bool Create(Worksheet sheet, View host, GroupObject item, Dictionary<string, Command> commands, List<string> names)
        {
            if (item == null)
                return false;

            var name = item.Name;
            int idx;
            if ((idx = names.BinarySearch(name)) >= 0)
                return true;

            ActionExtensions.Try(() =>
            {
                names.Insert(~idx, name);
                var shapes = (from Shape x in item.ShapeRange from Shape y in x.GroupItems select y).ToArray();
                item.Ungroup();
                Create(sheet, host, shapes, names, commands);
                sheet.Shapes.Range[(from Shape x in shapes select x.Name).ToArray()].Regroup();
            });
            return true;
        }

        private static bool Create(View host, Shape item, Dictionary<string, Command> commands, List<string> names)
        {
            if (item == null)
                return false;

            var name = item.Name;
            int idx;
            if ((idx = names.BinarySearch(name)) >= 0 || !IsCreateable(name))
                return true;

            ActionExtensions.Try(() =>
            {
                names.Insert(~idx, name);
                GroupShapes unused = null;
                ActionExtensions.Try(() => unused = item.GroupItems);
                if (unused == null)
                    commands[name] = new CommandShape(host, item, RemovePrefix(name));
            });

            return true;
        }

        private static bool IsCreateable(string name)
        {
            return StartsWithPrefix(name) || StartsWithFullPrefix(name);
        }

        private static bool StartsWithPrefix(string name)
        {
            return name.StartsWith(CommandPrefix, StringComparison.OrdinalIgnoreCase);
        }

        private static bool StartsWithFullPrefix(string name)
        {
            return name.StartsWith(CommandFullPrefix, StringComparison.OrdinalIgnoreCase);
        }

        #endregion Methods
    }
}