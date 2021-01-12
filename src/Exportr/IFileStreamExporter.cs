﻿/*
 * Copyright 2017 - 2021 Wouter Huysentruit
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

using System.IO;
using System.Threading.Tasks;

namespace Exportr
{
    /// <summary>
    /// Interface that describes a file stream exporter.
    /// </summary>
    public interface IFileStreamExporter
    {
        /// <summary>
        /// Gets the proposed filename.
        /// </summary>
        /// <returns>The proposed filename.</returns>
        string GetFileName();

        /// <summary>
        /// Exports the export data to the given stream.
        /// </summary>
        /// <param name="stream">The stream to write the export to.</param>
        Task ExportToStream(Stream stream);
    }
}
