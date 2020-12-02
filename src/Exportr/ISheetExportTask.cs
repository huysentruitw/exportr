/*
 * Copyright 2017 - 2020 Wouter Huysentruit
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

using System.Collections.Generic;

namespace Exportr
{
    /// <summary>
    /// Interface that describes a sheet export.
    /// </summary>
    public interface ISheetExportTask
    {
        /// <summary>
        /// Gets the name of the sheet export task. This name can/will be used to label the sheet in the document.
        /// </summary>
        string Name { get; }

        /// <summary>
        /// Gets the column labels for the header row.
        /// </summary>
        /// <returns>The column labels.</returns>
        string[] GetColumnLabels();

        /// <summary>
        /// Enumerates the row data.
        /// </summary>
        /// <returns>Row data: an <see cref="IEnumerable{T}"/> per row that contains the cell values for that row.</returns>
        IAsyncEnumerable<IEnumerable<object>> EnumRowData();
    }
}
