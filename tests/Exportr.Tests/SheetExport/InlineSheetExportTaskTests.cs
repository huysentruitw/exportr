using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Exportr.SheetExport;
using NUnit.Framework;

namespace Exportr.Tests.SheetExport
{
    [TestFixture]
    public class InlineSheetExportTaskTests
    {
        [Test]
        public void SingleParse_InvalidArguments_ShouldThrowException()
        {
            ArgumentException ex = Assert.Throws<ArgumentNullException>(() => InlineSheetExportTask<SomeEntityModel, SomeRowData>.SingleParse(null, EmptyAsyncEnumerable<SomeEntityModel>, _ => new SomeRowData()));
            Assert.That(ex.ParamName, Is.EqualTo("name"));

            ex = Assert.Throws<ArgumentNullException>(() => InlineSheetExportTask<SomeEntityModel, SomeRowData>.SingleParse(string.Empty, EmptyAsyncEnumerable<SomeEntityModel>, _ => new SomeRowData()));
            Assert.That(ex.ParamName, Is.EqualTo("name"));

            ex = Assert.Throws<ArgumentNullException>(() => InlineSheetExportTask<SomeEntityModel, SomeRowData>.SingleParse("SomeName", null, _ => new SomeRowData()));
            Assert.That(ex.ParamName, Is.EqualTo("fetchEntities"));

            ex = Assert.Throws<ArgumentNullException>(() => InlineSheetExportTask<SomeEntityModel, SomeRowData>.SingleParse("SomeName", EmptyAsyncEnumerable<SomeEntityModel>, null));
            Assert.That(ex.ParamName, Is.EqualTo("parseEntity"));
        }

        [Test]
        public void SingleParse_PassName_ShouldSetNameOnProperty()
        {
            var name = Guid.NewGuid().ToString("N");
            var task = InlineSheetExportTask<SomeEntityModel, SomeRowData>.SingleParse(name, EmptyAsyncEnumerable<SomeEntityModel>, _ => new SomeRowData());
            Assert.That(task.Name, Is.EqualTo(name));
        }

        [Test]
        public void MultiParse_InvalidArguments_ShouldThrowException()
        {
            ArgumentException ex = Assert.Throws<ArgumentNullException>(() => InlineSheetExportTask<SomeEntityModel, SomeRowData>.MultiParse(null, EmptyAsyncEnumerable<SomeEntityModel>, _ => EmptyAsyncEnumerable<SomeRowData>()));
            Assert.That(ex.ParamName, Is.EqualTo("name"));

            ex = Assert.Throws<ArgumentNullException>(() => InlineSheetExportTask<SomeEntityModel, SomeRowData>.MultiParse(string.Empty, EmptyAsyncEnumerable<SomeEntityModel>, _ => EmptyAsyncEnumerable<SomeRowData>()));
            Assert.That(ex.ParamName, Is.EqualTo("name"));

            ex = Assert.Throws<ArgumentNullException>(() => InlineSheetExportTask<SomeEntityModel, SomeRowData>.MultiParse("SomeName", null, _ => EmptyAsyncEnumerable<SomeRowData>()));
            Assert.That(ex.ParamName, Is.EqualTo("fetchEntities"));

            ex = Assert.Throws<ArgumentNullException>(() => InlineSheetExportTask<SomeEntityModel, SomeRowData>.MultiParse("SomeName", EmptyAsyncEnumerable<SomeEntityModel>, null));
            Assert.That(ex.ParamName, Is.EqualTo("parseEntity"));
        }

        [Test]
        public void MultiParse_PassName_ShouldSetNameOnProperty()
        {
            var name = Guid.NewGuid().ToString("N");
            var task = InlineSheetExportTask<SomeEntityModel, SomeRowData>.MultiParse(name, EmptyAsyncEnumerable<SomeEntityModel>, _ => EmptyAsyncEnumerable<SomeRowData>());
            Assert.That(task.Name, Is.EqualTo(name));
        }

        [Test]
        public async Task MultiParse_NoAdditionalData_ShouldExportColumnsSpecifiedInRowDataModel()
        {
#pragma warning disable 1998
            static async IAsyncEnumerable<SomeEntityModel> Fetch()
#pragma warning restore 1998
            {
                yield return new SomeEntityModel { Title = "Title1", Description = "DescriptionA" };
                yield return new SomeEntityModel { Title = "Title2", Description = "DescriptionB" };
            }

#pragma warning disable 1998
            static async IAsyncEnumerable<SomeRowData> Parse(SomeEntityModel entity)
#pragma warning restore 1998
            {
                yield return new SomeRowData { Title = entity.Title, Description = entity.Description, OrderA = "aa", OrderB = "bb", OrderC = "cc" };
            }

            var task = InlineSheetExportTask<SomeEntityModel, SomeRowData>.MultiParse("SomeName", Fetch, Parse);

            var columns = task.GetColumnLabels();
            Assert.That(columns, Is.EquivalentTo(new object[] { "Title", "Description", "OrderC", "OrderA", "OrderB" }));

            var data = await ToArray(task.EnumRowData());
            Assert.That(data, Has.Length.EqualTo(2));
            Assert.That(data[0], Is.EquivalentTo(new object[] { "Title1", "DescriptionA", "cc", "aa", "bb" }));
            Assert.That(data[1], Is.EquivalentTo(new object[] { "Title2", "DescriptionB", "cc", "aa", "bb" }));
        }

        [Test]
        public async Task MultiParse_AdditionalData_ShouldExportColumnsSpecifiedInRowDataModelAndAdditionalData()
        {
#pragma warning disable 1998
            static async IAsyncEnumerable<SomeEntityModel> Fetch()
#pragma warning restore 1998
            {
                yield return new SomeEntityModel { Title = "Title1", Description = "DescriptionA" };
                yield return new SomeEntityModel { Title = "Title2", Description = "DescriptionB" };
            }

#pragma warning disable 1998
            static async IAsyncEnumerable<SomeRowData> Parse(SomeEntityModel entity)
#pragma warning restore 1998
            {
                yield return new SomeRowData
                {
                    Title = entity.Title,
                    Description = entity.Description,
                    OrderA = "aa",
                    OrderB = "bb",
                    OrderC = "cc",
                    AdditionalValues = new object[] { true, 1 },
                };
            }

            var task = InlineSheetExportTask<SomeEntityModel, SomeRowData>.MultiParse("SomeName", Fetch, Parse, new[] { "Additional1", "Additional2" });

            var columns = task.GetColumnLabels();
            Assert.That(columns, Is.EquivalentTo(new object[] { "Title", "Description", "OrderC", "OrderA", "OrderB", "Additional1", "Additional2" }));

            var data = await ToArray(task.EnumRowData());
            Assert.That(data, Has.Length.EqualTo(2));
            Assert.That(data[0], Is.EquivalentTo(new object[] { "Title1", "DescriptionA", "cc", "aa", "bb", true, 1 }));
            Assert.That(data[1], Is.EquivalentTo(new object[] { "Title2", "DescriptionB", "cc", "aa", "bb", true, 1 }));
        }

        [Test]
        public async Task MultiParse_AdditionalData_DuplicateColumnNames_ShouldExportDuplicateColumnNames()
        {
#pragma warning disable 1998
            static async IAsyncEnumerable<SomeEntityModel> Fetch()
#pragma warning restore 1998
            {
                yield return new SomeEntityModel { Title = "Title1", Description = "DescriptionA" };
                yield return new SomeEntityModel { Title = "Title2", Description = "DescriptionB" };
            }

#pragma warning disable 1998
            static async IAsyncEnumerable<SomeRowData> Parse(SomeEntityModel entity)
#pragma warning restore 1998
            {
                yield return new SomeRowData
                {
                    Title = entity.Title,
                    Description = entity.Description,
                    OrderA = "aa",
                    OrderB = "bb",
                    OrderC = "cc",
                    AdditionalValues = new object[] { true, 1 },
                };
            }

            var task = InlineSheetExportTask<SomeEntityModel, SomeRowData>.MultiParse("SomeName", Fetch, Parse, new[] { "OrderA", "OrderB" });

            var columns = task.GetColumnLabels();
            Assert.That(columns, Is.EquivalentTo(new object[] { "Title", "Description", "OrderC", "OrderA", "OrderB", "OrderA", "OrderB" }));

            var data = await ToArray(task.EnumRowData());
            Assert.That(data, Has.Length.EqualTo(2));
            Assert.That(data[0], Is.EquivalentTo(new object[] { "Title1", "DescriptionA", "cc", "aa", "bb", true, 1 }));
            Assert.That(data[1], Is.EquivalentTo(new object[] { "Title2", "DescriptionB", "cc", "aa", "bb", true, 1 }));
        }

        [Test]
        public async Task MultiParse_AdditionalData_DuplicateColumnValues_ShouldExportDuplicateValues()
        {
#pragma warning disable 1998
            static async IAsyncEnumerable<SomeEntityModel> Fetch()
#pragma warning restore 1998
            {
                yield return new SomeEntityModel { Title = "Title1", Description = "DescriptionA" };
                yield return new SomeEntityModel { Title = "Title2", Description = "DescriptionB" };
            }

#pragma warning disable 1998
            static async IAsyncEnumerable<SomeRowData> Parse(SomeEntityModel entity)
#pragma warning restore 1998
            {
                yield return new SomeRowData
                {
                    Title = entity.Title,
                    Description = entity.Description,
                    OrderA = "aa",
                    OrderB = "aa",
                    OrderC = "cc",
                    AdditionalValues = new object[] { "cc", "aa" }
                };
            }

            var task = InlineSheetExportTask<SomeEntityModel, SomeRowData>.MultiParse("SomeName", Fetch, Parse, new[] { "OrderA", "OrderB" });

            var columns = task.GetColumnLabels();
            Assert.That(columns, Is.EquivalentTo(new object[] { "Title", "Description", "OrderC", "OrderA", "OrderB", "OrderA", "OrderB" }));

            var data = await ToArray(task.EnumRowData());
            Assert.That(data, Has.Length.EqualTo(2));
            Assert.That(data[0], Is.EquivalentTo(new object[] { "Title1", "DescriptionA", "cc", "aa", "aa", "cc", "aa" }));
            Assert.That(data[1], Is.EquivalentTo(new object[] { "Title2", "DescriptionB", "cc", "aa", "aa", "cc", "aa" }));
        }

#pragma warning disable 1998
        private static async IAsyncEnumerable<T> EmptyAsyncEnumerable<T>()
#pragma warning restore 1998
        {
            yield break;
        }

        private static async Task<T[]> ToArray<T>(IAsyncEnumerable<T> items)
        {
            var list = new List<T>();
            await foreach (var item in items)
                list.Add(item);
            return list.ToArray();
        }

        private class SomeEntityModel
        {
            public string Title { get; set; }
            public string Description { get; set; }
        }

        private class SomeRowData : RowData
        {
            [Column("Title", Order = 0)]
            public string Title { get; set; }
            [Column("Description", Order = 1)]
            public string Description { get; set; }

            [Column("OrderA", Order = 3)]
            public string OrderA { get; set; }
            [Column("OrderB", Order = 4)]
            public string OrderB { get; set; }
            [Column("OrderC", Order = 2)]
            public string OrderC { get; set; }
        }
    }
}
