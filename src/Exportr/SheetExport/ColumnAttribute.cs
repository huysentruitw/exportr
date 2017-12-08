/*
 * Copyright 2017 Wouter Huysentruit
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *   http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

using System;

namespace Exportr.SheetExport
{
    /// <summary>
    /// Defines the column name and order of properties added to a custom <see cref="RowData"/> override.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class ColumnAttribute : Attribute
    {
        /// <summary>
        /// Constructs a new <see cref="ColumnAttribute"/> instance with the given name.
        /// </summary>
        /// <param name="name">The name of the column the property is mapped to.</param>
        public ColumnAttribute(string name)
        {
            Name = name;
        }

        /// <summary>
        /// Gets the name of the column the property is mapped to.
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// Gets or sets the zero-based order of the column the property is mapped to.
        /// </summary>
        public int Order { get; set; }
    }
}
