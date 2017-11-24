/*
 * Copyright 2017 Wouter Huysentruit, Jeroen Verhaeghe
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
using DocumentFormat.OpenXml.Spreadsheet;

namespace Exportr.OpenXml
{
    /// <summary>
    /// A default value converter implementation.
    /// </summary>
    public class DefaultExcelValueConverter : IExcelValueConverter
    {
        private readonly IFormatProvider _formatProvider;
        private readonly string _dateTimeFormat;

        /// <summary>
        /// Constructs a new <see cref="DefaultExcelValueConverter"/> instance.
        /// </summary>
        /// <param name="formatProvider">The format provider for formatting numbers and datetime.</param>
        /// <param name="dateTimeFormat">The datetime format specifier.</param>
        public DefaultExcelValueConverter(IFormatProvider formatProvider, string dateTimeFormat)
        {
            _formatProvider = formatProvider ?? throw new ArgumentNullException(nameof(formatProvider));
            if (string.IsNullOrEmpty(dateTimeFormat)) throw new ArgumentNullException(nameof(dateTimeFormat));
            _dateTimeFormat = dateTimeFormat;
        }

        /// <summary>
        /// Convert a cell value into a <see cref="Cell"/> object.
        /// </summary>
        /// <param name="value">The cell value.</param>
        /// <returns>A <see cref="Cell"/> object.</returns>
        public Cell Convert(object value)
        {
            switch (value)
            {
                case string text: return ToInlineStringCell(text);
                case Guid guid: return ToInlineStringCell(guid.ToString());
                case DateTime dateTime: return ToInlineStringCell(dateTime.ToString(_dateTimeFormat, _formatProvider));
                case DateTimeOffset dateTimeOffset: return ToInlineStringCell(dateTimeOffset.UtcDateTime.ToString(_dateTimeFormat, _formatProvider));
                case bool @bool: return ToInlineStringCell(@bool.ToString());
                case int @int: return ToNumberCell(@int);
                case uint @uint: return ToNumberCell(@uint);
                case long @long: return ToNumberCell(@long);
                case ulong @ulong: return ToNumberCell(@ulong);
                case double @double: return ToNumberCell((decimal)@double);
                case decimal @decimal: return ToNumberCell(@decimal);
                default: return new Cell();
            }
        }

        // ReSharper disable once MemberCanBeMadeStatic.Local
        private Cell ToInlineStringCell(string value)
        {
            var cell = new Cell { DataType = CellValues.InlineString };
            // Excel has a limit of 32767 characters per cell.
            // See https://support.office.com/en-sg/article/Excel-specifications-and-limits-16c69c74-3d6a-4aaf-ba35-e6eb276e8eaa
            var inlineString = new InlineString();
            inlineString.AppendChild(new Text { Text = value.Substring(0, Math.Min(value.Length, 32764)) });
            cell.AppendChild(inlineString);
            return cell;
        }

        private Cell ToNumberCell(decimal value)
            => new Cell
            {
                DataType = CellValues.Number,
                CellValue = new CellValue(value.ToString(_formatProvider)),
            };
    }
}
