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

using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Exportr.OpenXml
{
    internal class ExcelSheet : ISheet
    {
        private readonly IExcelValueConverter _valueConverter;
        private readonly OpenXmlWriter _writer;

        public ExcelSheet(WorksheetPart part, IExcelValueConverter valueConverter)
        {
            if (part == null) throw new ArgumentNullException(nameof(part));
            _valueConverter = valueConverter ?? throw new ArgumentNullException(nameof(valueConverter));
            _writer = OpenXmlWriter.Create(part);
            _writer.WriteStartElement(new Worksheet());
            _writer.WriteStartElement(new SheetData());
        }

        public void Dispose()
        {
            _writer.WriteEndElement();
            _writer.WriteEndElement();
            _writer.Dispose();
        }

        public void AddHeaderRow(params string[] labels)
        {
            var row = new Row(labels.Select(x => new Cell
            {
                DataType = CellValues.String,
                CellValue = new CellValue(x)
            }));

            _writer.WriteElement(row);
        }

        public void AddRow(IEnumerable<object> values)
        {
            var row = new Row(values.Select(_valueConverter.Convert));
            _writer.WriteElement(row);
        }
    }
}
