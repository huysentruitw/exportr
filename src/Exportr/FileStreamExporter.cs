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
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace Exportr
{
    /// <summary>
    /// Takes an export task, proposes a filename and writes it to a stream.
    /// </summary>
    public sealed class FileStreamExporter : IFileStreamExporter
    {
        private readonly IDocumentFactory _documentFactory;
        private readonly IExportTask _exportTask;
        private readonly Func<DateTime> _nowFactory;

        /// <summary>
        /// Constructs a new <see cref="FileStreamExporter"/> instance.
        /// </summary>
        /// <param name="documentFactory">The export document factory.</param>
        /// <param name="exportTask">The export task.</param>
        public FileStreamExporter(IDocumentFactory documentFactory, IExportTask exportTask)
            : this(documentFactory, exportTask, () => DateTime.Now)
        {
        }

        internal FileStreamExporter(IDocumentFactory documentFactory, IExportTask exportTask, Func<DateTime> nowFactory)
        {
            _documentFactory = documentFactory ?? throw new ArgumentNullException(nameof(documentFactory));
            _exportTask = exportTask ?? throw new ArgumentNullException(nameof(exportTask));
            _nowFactory = nowFactory ?? throw new ArgumentNullException(nameof(nowFactory));
        }

        /// <summary>
        /// Gets the proposed filename.
        /// </summary>
        /// <returns>The proposed filename.</returns>
        public string GetFileName()
        {
            var name = _exportTask.Name;
            if (string.IsNullOrEmpty(name)) throw new InvalidOperationException("Failed to get the name of the export task");
            var extension = _documentFactory.FileExtension ?? string.Empty;
            if (extension.Any() && !extension.StartsWith(".")) extension = $".{extension}";
            return $"{FileNameEncode(name)} {_nowFactory():yyyyMMdd}{extension}";
        }

        /// <summary>
        /// Exports the export data to the given stream.
        /// </summary>
        /// <param name="stream">The stream to write the export to.</param>
        public async Task ExportToStream(Stream stream)
        {
            using var document = _documentFactory.CreateDocument(stream);
            foreach (var sheetTask in _exportTask.EnumSheetExportTasks())
            {
                using var sheet = document.CreateSheet(sheetTask.Name);
                sheet.AddHeaderRow(sheetTask.GetColumnLabels());

                await foreach (IEnumerable<object> rowData in sheetTask.EnumRowData())
                {
                    sheet.AddRow(rowData);
                }
            }
        }

        private static string FileNameEncode(string value)
            => Path.GetInvalidFileNameChars().Aggregate(value, (seed, invalidChar) => seed.Replace(invalidChar, '_'));
    }
}
