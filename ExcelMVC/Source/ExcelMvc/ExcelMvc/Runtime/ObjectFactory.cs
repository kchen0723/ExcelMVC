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
    using System.IO;
    using System.Linq;
    using System.Reflection;
    using System.Reflection.Emit;
    using System.Runtime.CompilerServices;
    using System.Windows;
    using Extensions;

    /// <summary>
    /// Generic object factory
    /// </summary>
    /// <typeparam name="T">Type of object</typeparam>
    public static class ObjectFactory<T>
    {
        #region Properties

        private static List<T> Instances
        {
            get; set;
        }

        #endregion Properties

        #region Methods

        /// <summary>
        /// Create instances of type T in the current AppDomain
        /// </summary>
        [MethodImpl(MethodImplOptions.Synchronized)]
        public static void CreateAll()
        {
            Instances = new List<T>();
            Instances.Clear();
            var itype = typeof(T);

            var asms = AppDomain.CurrentDomain.GetAssemblies();
            foreach (var asm in AppDomain.CurrentDomain.GetAssemblies())
            {
                foreach (var type in asm.GetTypes())
                {
                    if (itype.IsAssignableFrom(type) && !type.IsInterface && !type.IsAbstract)
                    {
                        if (type.GetConstructor(Type.EmptyTypes) != null)
                            Instances.Add((T)Activator.CreateInstance(type));
                    }
                }
            }

            var location = typeof(ObjectFactory<T>).Assembly.Location;
            if (string.IsNullOrEmpty(location))
            {
                // from a sandbox execution(e.g.Excel-DNA)
                return;
            }

            var path = Path.GetDirectoryName(location);
            var files = Directory.GetFiles(path, "*.dll", SearchOption.AllDirectories);

            // .NET 4 Assembly.IsDynamic equvialent
            Func<Assembly, bool> isDynamic = asm =>
            {
                if (asm.ManifestModule is ModuleBuilder)
                    return true;

                // the above test does not really return true for a dynamic assembly, hence use the try ignore 
                // method
                var asmPath = "";
                ActionExtensions.Try(() => asmPath = asm.Location);
                return string.IsNullOrEmpty(asmPath);
            };
            var nonDynamicAsms = asms.Where(x=> !isDynamic(x));
    
            // exclude files already loaded
            files = files.Where(x => nonDynamicAsms.All(y => y.Location.CompareOrdinalIgnoreCase(x) != 0)).ToArray();
            foreach (var file in files)
                Discover(file);
        }

        /// <summary>
        /// Deletes instance created
        /// </summary>
        [MethodImpl(MethodImplOptions.Synchronized)]
        public static void DeleteAll(Action<T> disposer)
        {
            if (Instances != null)
            {
                if (disposer != null)
                    Instances.ForEach(disposer);
                Instances.Clear();
            }
        }

        /// <summary>
        /// Finds the instance matching the full type name specified
        /// </summary>
        /// <param name="fullTypeName"></param>
        /// <returns></returns>
        [MethodImpl(MethodImplOptions.Synchronized)]
        public static T Find(string fullTypeName)
        {
            var idx = Instances.FindIndex(x => x.GetType().FullName == fullTypeName);
            if (idx < 0)
                idx = Instances.FindIndex(x => x.GetType().AssemblyQualifiedName == fullTypeName);
            return idx < 0 ? default(T) : Instances[idx];
        }

        private static void Discover(string assemblyPath)
        {
            ActionExtensions.Try(() =>
            {
                var itype = typeof(T);
                var types = TypeDiscoveryDomains.Discover(assemblyPath, itype);
                foreach (var type in types)
                    Instances.Add((T)Activator.CreateInstance(Type.GetType(type)));
            });
        }

        #endregion Methods
    }
}