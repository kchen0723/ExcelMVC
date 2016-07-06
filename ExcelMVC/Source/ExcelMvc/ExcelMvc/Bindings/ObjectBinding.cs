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

namespace ExcelMvc.Bindings
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Reflection;
    using System.Text.RegularExpressions;

    #region Delegates

    /// <summary>
    /// 
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="args"></param>
    public delegate void BindingFailedHandler(object sender, BindingFailedEventArgs args);

    #endregion Delegates

    /// <summary>
    /// Encapsulates binding functions between View models and Views
    /// </summary>
    internal static class ObjectBinding
    {
        #region Fields

        private const string NoQuoteCommaPattern = "[^\",]*";
        private const string QuotedStringPattern = "\"([^\"]*|\"\")*\"";

        private static readonly Regex IndexerPattern = new Regex(string.Format("({0}|{1})(,({0}|{1}))*", QuotedStringPattern, NoQuoteCommaPattern));
        private static readonly Type MatrixType = typeof(object[,]);
        private static readonly char[] PathSeparators = { '.', '[' };

        #endregion Fields

        #region Methods

        /// <summary>
        /// Changes the lower bounds of an array
        /// </summary>
        /// <typeparam name="T">Element type</typeparam>
        /// <param name="value">Array to be changed</param>
        /// <param name="lowerBound">Lower bound</param>
        /// <returns>Changed array</returns>
        public static Array ChangeLBound<T>(Array value, int lowerBound)
        {
            var lbounds = new int[value.Rank];
            var lengths = new int[value.Rank];
            for (var idx = 0; idx < value.Rank; idx++)
            {
                lbounds[idx] = lowerBound;
                lengths[idx] = value.GetLength(idx);
            }

            var to = Array.CreateInstance(typeof(object), lengths, lbounds);
            Array.Copy(value, to, value.Length);
            return to;
        }

        /// <summary>
        /// Gets a list of Get property names
        /// </summary>
        /// <param name="value">Object to be Interrogated</param>
        /// <returns>a list of Get property names</returns>
        public static IEnumerable<string> GetGetPropertyNames(object value)
        {
            const BindingFlags Flags = BindingFlags.GetProperty | BindingFlags.Instance | BindingFlags.Public;
            return value.GetType().GetProperties(Flags).Select(x => x.Name);
        }

        /// <summary>
        /// Gets the property value by Binding
        /// </summary>
        /// <param name="source">Source object</param>
        /// <param name="binding">Binding object</param>
        /// <returns>Property value</returns>
        public static object GetPropertyValue(object source, Binding binding)
        {
            var value = GetPropertyValue(source, binding.Path);
            if (binding.Converter != null)
                value = binding.Converter.Convert(value, null, null, null);
            return value;
        }

        /// <summary>
        /// Gets the property value by full path
        /// </summary>
        /// <param name="source">Source object</param>
        /// <param name="path">Property path</param>
        /// <returns>Property value</returns>
        public static object GetPropertyValue(object source, string path)
        {
            ReducePath(ref source, ref path);
            if (path == string.Empty)
            {
                return source;
            }

            var parameters = ParseIndexerParams(path);
            var property = GetProperty(BindingFlags.GetProperty, source, path, parameters);
            if (property == null)
                throw new Exception(string.Format(Resource.ErrorNoPropertyGet, path, source.GetType().FullName));
            return property.GetValue(source, parameters.Count > 0 ? parameters.ToArray() : null);
        }

        /// <summary>
        /// Sets property value by full property path
        /// </summary>
        /// <param name="source">Source object</param>
        /// <param name="binding">Binding object</param>
        /// <param name="value">Property value</param>
        public static void SetPropertyValue(object source, Binding binding, object value)
        {
            if (binding.Converter != null)
                value = binding.Converter.ConvertBack(value, null, null, null);
            SetPropertyValue(source, binding.Path, value);
        }

        /// <summary>
        /// Sets property value by full property path
        /// </summary>
        /// <param name="source">Source object</param>
        /// <param name="path">Property path</param>
        /// <param name="value">Property value</param>
        public static void SetPropertyValue(object source, string path, object value)
        {
            ReducePath(ref source, ref path);
            var parameters = ParseIndexerParams(path);

            var property = GetProperty(BindingFlags.SetProperty, source, path, parameters);
            if (property == null)
                throw new Exception(string.Format(Resource.ErrorNoPropertySet, path, source.GetType().FullName));
            if (property.PropertyType == MatrixType)
            {
                if (value == null || value.GetType() != MatrixType)
                    value = new[,] { { value } };
            }
            else if (value != null && value.GetType() == MatrixType)
            {
                var matrix = (object[,])value;
                value = matrix[matrix.GetLowerBound(0), matrix.GetLowerBound(1)];
            }

            if (value != null)
                value = Convert.ChangeType(value, property.PropertyType);
            property.SetValue(source, value, parameters.Count > 0 ? parameters.ToArray() : null);
        }

        private static PropertyInfo GetProperty(BindingFlags forFlag, object source, string path, List<object> parameters)
        {
            if (parameters.Count > 0)
            {
                var types = parameters.Select(x => x.GetType()).ToArray();
                return source.GetType().GetProperty("Item", types);
            }

            var flags = forFlag | BindingFlags.Instance | BindingFlags.Public;
            return source.GetType().GetProperty(path, flags);
        }

        private static List<object> ParseIndexerParams(string path)
        {
            var parameters = new List<object>();
            if (!path.StartsWith("[") || !path.EndsWith("]"))
                return parameters;

            var matches = IndexerPattern.Matches(path.Substring(1, path.Length - 2));
            foreach (Match match in matches)
            {
                var value = match.Value.Trim();
                if (value == string.Empty)
                    continue;
                if (value.StartsWith(","))
                    value = value.Substring(1);

                if (value.StartsWith("\"") && value.EndsWith("\""))
                    value = value.Substring(1, value.Length).Replace("\"\"", "\"");
                int ivalue;
                if (int.TryParse(value, out ivalue))
                    parameters.Add(value);
                else
                    parameters.Add(value);
            }

            return parameters;
        }

        private static void ReducePath(ref object source, ref string path)
        {
            var pos = path.IndexOfAny(PathSeparators);
            while (pos >= 0)
            {
                if (pos > 0)
                {
                    source = GetPropertyValue(source, path.Substring(0, pos));
                    path = path.Substring(pos);
                }

                if (path[0] == '[')
                {
                    // find matching ']'
                    var end = path.IndexOf(']', pos + 1);
                    if (end < 0)
                        throw new Exception(string.Format(Resource.ErrorBindingPathInvalidIndexer, path));
                    if (end == path.Length - 1)
                        break;
                    source = GetPropertyValue(source, path.Substring(0, end + 1));
                    path = path.Substring(end + 1);
                }
                else
                {
                    // skip .
                    path = path.Substring(1);
                }

                pos = path.IndexOfAny(PathSeparators);
            }
        }

        #endregion Methods
    }
}