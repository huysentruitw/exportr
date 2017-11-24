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
using System.Globalization;
using System.IO;

namespace Exportr.OpenXml
{
    /// <summary>
    /// Factory for creating excel documents.
    /// </summary>
    public class ExcelDocumentFactory : IDocumentFactory
    {
        private readonly IExcelValueConverter _valueConverter;

        /// <summary>
        /// Constructs a new <see cref="ExcelDocumentFactory"/> instance with the default value converter.
        /// </summary>
        public ExcelDocumentFactory()
        {
            _valueConverter = new DefaultExcelValueConverter(CultureInfo.InvariantCulture, "O");
        }

        /// <summary>
        /// Constructs a new <see cref="ExcelDocumentFactory"/> with a custom value converter.
        /// </summary>
        /// <param name="valueConverter">The custom value converter.</param>
        public ExcelDocumentFactory(IExcelValueConverter valueConverter)
        {
            _valueConverter = valueConverter ?? throw new ArgumentNullException(nameof(valueConverter));
        }

        /// <summary>
        /// Creates a document for writing to an Excel stream.
        /// </summary>
        /// <param name="stream">The stream to write to.</param>
        /// <returns>An Excel specific <see cref="IDocument"/> implementation.</returns>
        public IDocument CreateDocument(Stream stream)
            => new ExcelDocument(stream, _valueConverter);
    }
}
