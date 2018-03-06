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
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace Exportr.SheetExport
{
    /// <summary>
    /// A generic implementation that can be used to create a <see cref="ISheetExportTask"/> inline.
    /// </summary>
    /// <typeparam name="TEntity">The entity model that contains all information needed for producing one or more rows.</typeparam>
    /// <typeparam name="TRowData">The row data model describes columnnames and row values.</typeparam>
    public class InlineSheetExportTask<TEntity, TRowData> : ISheetExportTask
        where TRowData : RowData
    {
        private readonly Func<IEnumerable<TEntity>> _fetchEntities;
        private readonly Func<TEntity, IEnumerable<TRowData>> _parseEntity;
        private readonly IEnumerable<string> _additionalColumnNames;
        private readonly ColumnInfo[] _orderedColumnInfos;

        /// <summary>
        /// Create a <see cref="InlineSheetExportTask{TEntity,TRowData}"/> instance that produces one row for each entity.
        /// </summary>
        /// <param name="name">The name of the sheet.</param>
        /// <param name="fetchEntities">A method that returns the entities to be used.</param>
        /// <param name="parseEntity">A method that parses an entity into a TRowData instance.</param>
        /// <param name="additionalColumnNames">The names of zero or more additional columns. The number of names should equal the number of AdditionalValues returned in the TRowData instance.</param>
        /// <returns>A <see cref="InlineSheetExportTask{TEntity,TRowData}"/> instance.</returns>
        public static InlineSheetExportTask<TEntity, TRowData> SingleParse(
            string name,
            Func<IEnumerable<TEntity>> fetchEntities,
            Func<TEntity, TRowData> parseEntity,
            IEnumerable<string> additionalColumnNames = null)
        {
            if (parseEntity == null) throw new ArgumentNullException(nameof(parseEntity));
            return new InlineSheetExportTask<TEntity, TRowData>(name, fetchEntities, x => new[] { parseEntity(x) }, additionalColumnNames);
        }

        /// <summary>
        /// Create a <see cref="InlineSheetExportTask{TEntity,TRowData}"/> instance that produces a variable number of rows for each entity.
        /// </summary>
        /// <param name="name">The name of the sheet.</param>
        /// <param name="fetchEntities">A method that returns the entities to be used.</param>
        /// <param name="parseEntity">A method that parses an entity into zero or more TRowData instances.</param>
        /// <param name="additionalColumnNames">The names of zero or more additional columns. The number of names should equal the number of AdditionalValues returned in the TRowData instance.</param>
        /// <returns>A <see cref="InlineSheetExportTask{TEntity,TRowData}"/> instance.</returns>
        public static InlineSheetExportTask<TEntity, TRowData> MultiParse(
            string name,
            Func<IEnumerable<TEntity>> fetchEntities,
            Func<TEntity, IEnumerable<TRowData>> parseEntity,
            IEnumerable<string> additionalColumnNames = null)
            => new InlineSheetExportTask<TEntity, TRowData>(name, fetchEntities, parseEntity, additionalColumnNames);

        private InlineSheetExportTask(
            string name,
            Func<IEnumerable<TEntity>> fetchEntities,
            Func<TEntity, IEnumerable<TRowData>> parseEntity,
            IEnumerable<string> additionalColumnNames = null)
        {
            if (string.IsNullOrEmpty(name)) throw new ArgumentNullException(nameof(name));
            Name = name;
            _fetchEntities = fetchEntities ?? throw new ArgumentNullException(nameof(fetchEntities));
            _parseEntity = parseEntity ?? throw new ArgumentNullException(nameof(parseEntity));
            _additionalColumnNames = additionalColumnNames ?? Enumerable.Empty<string>();
            _orderedColumnInfos = typeof(TRowData)
                .GetProperties(BindingFlags.Instance | BindingFlags.Public)
                .Select(propertyInfo =>
                {
                    var columnAttribute = propertyInfo.GetCustomAttribute<ColumnAttribute>();
                    return new ColumnInfo
                    {
                        Name = columnAttribute?.Name,
                        Order = columnAttribute?.Order ?? 0,
                        PropertyInfo = propertyInfo
                    };
                })
                .Where(columnInfo => !string.IsNullOrEmpty(columnInfo.Name))
                .OrderBy(columnInfo => columnInfo.Order)
                .ToArray();
        }

        /// <summary>
        /// Gets the name of the sheet.
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// Gets the column labels of the sheet.
        /// </summary>
        /// <returns>The column labels of the sheet.</returns>
        public string[] GetColumnLabels() => _orderedColumnInfos.Select(x => x.Name).Concat(_additionalColumnNames).ToArray();

        /// <summary>
        /// Gets the row data of the sheet.
        /// </summary>
        /// <returns>The row data of the sheet.</returns>
        public IEnumerable<object[]> EnumRowData()
            => _fetchEntities()
                .SelectMany(entity =>
                {
                    var rowDatas = _parseEntity(entity);
                    return rowDatas.Select(rowData =>
                    {
                        var additionalValues = rowData.AdditionalValues ?? Enumerable.Empty<object>();
                        return _orderedColumnInfos
                            .Select(columnInfo => columnInfo.PropertyInfo.GetValue(rowData))
                            .Concat(additionalValues)
                            .ToArray();
                    });
                });

        private class ColumnInfo
        {
            public string Name { get; set; }
            public int Order { get; set; }
            public PropertyInfo PropertyInfo { get; set; }
        }
    }
}
