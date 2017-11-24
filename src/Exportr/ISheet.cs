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

namespace Exportr
{
    /// <summary>
    /// Interface that describes a sheet in a document.
    /// </summary>
    public interface ISheet : IDisposable
    {
        /// <summary>
        /// Adds a header row to the sheet.
        /// </summary>
        /// <param name="labels">The labels of the header row cells.</param>
        void AddHeaderRow(params string[] labels);

        /// <summary>
        /// Adds a new row to the sheet.
        /// </summary>
        /// <param name="values">The values of the cells for the new row.</param>
        void AddRow(params object[] values);
    }
}
