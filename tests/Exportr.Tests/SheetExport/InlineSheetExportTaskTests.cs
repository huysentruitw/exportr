using System;
using System.Linq;
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
            ArgumentException ex = Assert.Throws<ArgumentNullException>(() => InlineSheetExportTask<SomeEntityModel, SomeRowData>.SingleParse(null, Enumerable.Empty<SomeEntityModel>, _ => new SomeRowData()));
            Assert.That(ex.ParamName, Is.EqualTo("name"));

            ex = Assert.Throws<ArgumentNullException>(() => InlineSheetExportTask<SomeEntityModel, SomeRowData>.SingleParse(string.Empty, Enumerable.Empty<SomeEntityModel>, _ => new SomeRowData()));
            Assert.That(ex.ParamName, Is.EqualTo("name"));

            ex = Assert.Throws<ArgumentNullException>(() => InlineSheetExportTask<SomeEntityModel, SomeRowData>.SingleParse("SomeName", null, _ => new SomeRowData()));
            Assert.That(ex.ParamName, Is.EqualTo("fetchEntities"));

            ex = Assert.Throws<ArgumentNullException>(() => InlineSheetExportTask<SomeEntityModel, SomeRowData>.SingleParse("SomeName", Enumerable.Empty<SomeEntityModel>, null));
            Assert.That(ex.ParamName, Is.EqualTo("parseEntity"));
        }

        [Test]
        public void SingleParse_PassName_ShouldSetNameOnProperty()
        {
            var name = Guid.NewGuid().ToString("N");
            var task = InlineSheetExportTask<SomeEntityModel, SomeRowData>.SingleParse(name, Enumerable.Empty<SomeEntityModel>, _ => new SomeRowData());
            Assert.That(task.Name, Is.EqualTo(name));
        }

        [Test]
        public void MultiParse_InvalidArguments_ShouldThrowException()
        {
            ArgumentException ex = Assert.Throws<ArgumentNullException>(() => InlineSheetExportTask<SomeEntityModel, SomeRowData>.MultiParse(null, Enumerable.Empty<SomeEntityModel>, _ => Enumerable.Empty<SomeRowData>()));
            Assert.That(ex.ParamName, Is.EqualTo("name"));

            ex = Assert.Throws<ArgumentNullException>(() => InlineSheetExportTask<SomeEntityModel, SomeRowData>.MultiParse(string.Empty, Enumerable.Empty<SomeEntityModel>, _ => Enumerable.Empty<SomeRowData>()));
            Assert.That(ex.ParamName, Is.EqualTo("name"));

            ex = Assert.Throws<ArgumentNullException>(() => InlineSheetExportTask<SomeEntityModel, SomeRowData>.MultiParse("SomeName", null, _ => Enumerable.Empty<SomeRowData>()));
            Assert.That(ex.ParamName, Is.EqualTo("fetchEntities"));

            ex = Assert.Throws<ArgumentNullException>(() => InlineSheetExportTask<SomeEntityModel, SomeRowData>.MultiParse("SomeName", Enumerable.Empty<SomeEntityModel>, null));
            Assert.That(ex.ParamName, Is.EqualTo("parseEntity"));
        }

        [Test]
        public void MultiParse_PassName_ShouldSetNameOnProperty()
        {
            var name = Guid.NewGuid().ToString("N");
            var task = InlineSheetExportTask<SomeEntityModel, SomeRowData>.MultiParse(name, Enumerable.Empty<SomeEntityModel>, _ => Enumerable.Empty<SomeRowData>());
            Assert.That(task.Name, Is.EqualTo(name));
        }

        [Test]
        public void MultiParse_NoAdditionalData_ShouldExportColumnsSpecifiedInRowDataModel()
        {
            var task = InlineSheetExportTask<SomeEntityModel, SomeRowData>.MultiParse("SomeName", () => new[]
            {
                new SomeEntityModel { Title = "Title1", Description = "DescriptionA" },
                new SomeEntityModel { Title = "Title2", Description = "DescriptionB" },
            }, entity => new[]
            {
                new SomeRowData { Title = entity.Title, Description = entity.Description, OrderA = "aa", OrderB = "bb", OrderC = "cc" },
            });

            var columns = task.GetColumnLabels();
            Assert.That(columns, Is.EquivalentTo(new object[] { "Title", "Description", "OrderC", "OrderA", "OrderB" }));

            var data = task.EnumRowData().ToArray();
            Assert.That(data, Has.Length.EqualTo(2));
            Assert.That(data[0], Is.EquivalentTo(new object[] { "Title1", "DescriptionA", "cc", "aa", "bb" }));
            Assert.That(data[1], Is.EquivalentTo(new object[] { "Title2", "DescriptionB", "cc", "aa", "bb" }));
        }

        [Test]
        public void MultiParse_AdditionalData_ShouldExportColumnsSpecifiedInRowDataModelAndAdditionalData()
        {
            var task = InlineSheetExportTask<SomeEntityModel, SomeRowData>.MultiParse("SomeName", () => new[]
            {
                new SomeEntityModel { Title = "Title1", Description = "DescriptionA" },
                new SomeEntityModel { Title = "Title2", Description = "DescriptionB" },
            }, entity => new[]
            {
                new SomeRowData
                {
                    Title = entity.Title,
                    Description = entity.Description,
                    OrderA = "aa",
                    OrderB = "bb",
                    OrderC = "cc",
                    AdditionalValues = new object[] { true, 1 }
                },
            }, new[] { "Additional1", "Additional2" });

            var columns = task.GetColumnLabels();
            Assert.That(columns, Is.EquivalentTo(new object[] { "Title", "Description", "OrderC", "OrderA", "OrderB", "Additional1", "Additional2" }));

            var data = task.EnumRowData().ToArray();
            Assert.That(data, Has.Length.EqualTo(2));
            Assert.That(data[0], Is.EquivalentTo(new object[] { "Title1", "DescriptionA", "cc", "aa", "bb", true, 1 }));
            Assert.That(data[1], Is.EquivalentTo(new object[] { "Title2", "DescriptionB", "cc", "aa", "bb", true, 1 }));
        }

        [Test]
        public void MultiParse_AdditionalData_DuplicateColumnNames_ShouldExportDuplicateColumnNames()
        {
            var task = InlineSheetExportTask<SomeEntityModel, SomeRowData>.MultiParse("SomeName", () => new[]
            {
                new SomeEntityModel { Title = "Title1", Description = "DescriptionA" },
                new SomeEntityModel { Title = "Title2", Description = "DescriptionB" },
            }, entity => new[]
            {
                new SomeRowData
                {
                    Title = entity.Title,
                    Description = entity.Description,
                    OrderA = "aa",
                    OrderB = "bb",
                    OrderC = "cc",
                    AdditionalValues = new object[] { true, 1 }
                },
            }, new[] { "OrderA", "OrderB" });

            var columns = task.GetColumnLabels();
            Assert.That(columns, Is.EquivalentTo(new object[] { "Title", "Description", "OrderC", "OrderA", "OrderB", "OrderA", "OrderB" }));

            var data = task.EnumRowData().ToArray();
            Assert.That(data, Has.Length.EqualTo(2));
            Assert.That(data[0], Is.EquivalentTo(new object[] { "Title1", "DescriptionA", "cc", "aa", "bb", true, 1 }));
            Assert.That(data[1], Is.EquivalentTo(new object[] { "Title2", "DescriptionB", "cc", "aa", "bb", true, 1 }));
        }

        [Test]
        public void MultiParse_AdditionalData_DuplicateColumnValues_ShouldExportDuplicateValues()
        {
            var task = InlineSheetExportTask<SomeEntityModel, SomeRowData>.MultiParse("SomeName", () => new[]
{
                new SomeEntityModel { Title = "Title1", Description = "DescriptionA" },
                new SomeEntityModel { Title = "Title2", Description = "DescriptionB" },
            }, entity => new[]
            {
                new SomeRowData
                {
                    Title = entity.Title,
                    Description = entity.Description,
                    OrderA = "aa",
                    OrderB = "aa",
                    OrderC = "cc",
                    AdditionalValues = new object[] { "cc", "aa" }
                },
            }, new[] { "OrderA", "OrderB" });

            var columns = task.GetColumnLabels();
            Assert.That(columns, Is.EquivalentTo(new object[] { "Title", "Description", "OrderC", "OrderA", "OrderB", "OrderA", "OrderB" }));

            var data = task.EnumRowData().ToArray();
            Assert.That(data, Has.Length.EqualTo(2));
            Assert.That(data[0], Is.EquivalentTo(new object[] { "Title1", "DescriptionA", "cc", "aa", "aa", "cc", "aa" }));
            Assert.That(data[1], Is.EquivalentTo(new object[] { "Title2", "DescriptionB", "cc", "aa", "aa", "cc", "aa" }));
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
