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

using DocumentFormat.OpenXml.Spreadsheet;

namespace Exportr.OpenXml
{
    /// <summary>
    /// Interface that describes an excel value converter.
    /// </summary>
    public interface IExcelValueConverter
    {
        /// <summary>
        /// Convert a cell value into a <see cref="Cell"/> object.
        /// </summary>
        /// <param name="value">The cell value.</param>
        /// <returns>A <see cref="Cell"/> object.</returns>
        Cell Convert(object value);
    }
}
