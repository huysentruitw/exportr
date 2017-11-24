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
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Exportr.OpenXml
{
    internal class ExcelDocument : IDocument
    {
        private readonly SpreadsheetDocument _document;
        private readonly IExcelValueConverter _valueConverter;
        private uint _sheetCounter = 1;

        public ExcelDocument(Stream stream, IExcelValueConverter valueConverter)
        {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            _valueConverter = valueConverter ?? throw new ArgumentNullException(nameof(valueConverter));
            _document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
            _document.AddPart(_document.AddWorkbookPart());
            _document.WorkbookPart.Workbook = new Workbook { Sheets = new Sheets() };
        }

        public void Dispose()
        {
            _document.Save();
            _document.Dispose();
        }

        public ISheet CreateSheet(string name)
        {
            var worksheetPart = _document.WorkbookPart.AddNewPart<WorksheetPart>();

            _document.WorkbookPart.Workbook.Sheets.AppendChild(new Sheet
            {
                Id = _document.WorkbookPart.GetIdOfPart(worksheetPart),
                SheetId = _sheetCounter++,
                Name = RemoveInvalidSheetNameCharacters(name)
            });

            return new ExcelSheet(worksheetPart, _valueConverter);
        }

        private static string RemoveInvalidSheetNameCharacters(string name)
        {
            var invalidCharacters = new[] { ':', '[', ']', '\\', '/', '*', '?' };
            return new string(name.Where(x => !invalidCharacters.Contains(x)).ToArray());
        }
    }
}
