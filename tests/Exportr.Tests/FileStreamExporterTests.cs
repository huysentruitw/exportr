using System;
using System.IO;
using Moq;
using NUnit.Framework;

namespace Exportr.Tests
{
    [TestFixture]
    public class FileStreamExporterTests
    {
        private readonly Mock<IDocumentFactory> _documentFactoryMock = new Mock<IDocumentFactory>();
        private readonly Mock<IExportTask> _exportTaskMock = new Mock<IExportTask>();

        [SetUp]
        public void SetUp()
        {
            _documentFactoryMock.Reset();
            _exportTaskMock.Reset();
            _exportTaskMock.SetupGet(x => x.Name).Returns("SomeExportName");
        }

        [Test]
        public void Constructor_InvalidArguments_ShouldThrowException()
        {
            var ex = Assert.Throws<ArgumentNullException>(() => new FileStreamExporter(null, new Mock<IExportTask>().Object));
            Assert.That(ex.ParamName, Is.EqualTo("documentFactory"));

            ex = Assert.Throws<ArgumentNullException>(() => new FileStreamExporter(new Mock<IDocumentFactory>().Object, null));
            Assert.That(ex.ParamName, Is.EqualTo("task"));
        }

        [Test]
        public void GetFileName_InvalidTaskName_ShouldThrowException()
        {
            var exporter = new FileStreamExporter(_documentFactoryMock.Object, _exportTaskMock.Object);

            _exportTaskMock.SetupGet(x => x.Name).Returns((string)null);
            var ex = Assert.Throws<InvalidOperationException>(() => exporter.GetFileName());
            Assert.That(ex.Message, Is.EqualTo("Failed to get the name of the export task"));

            _exportTaskMock.SetupGet(x => x.Name).Returns(string.Empty);
            ex = Assert.Throws<InvalidOperationException>(() => exporter.GetFileName());
            Assert.That(ex.Message, Is.EqualTo("Failed to get the name of the export task"));
        }

        [Test]
        public void GetFileName_DocumentFactoryReturningNullAsExtension_ShouldReturnFileNameWithoutExtension()
        {
            _documentFactoryMock.SetupGet(x => x.FileExtension).Returns((string)null);
            var exporter = new FileStreamExporter(_documentFactoryMock.Object, _exportTaskMock.Object, () => DateTime.Parse("2018/01/05 12:00"));
            Assert.That(exporter.GetFileName(), Is.EqualTo("SomeExportName 20180105"));
        }

        [Test]
        public void GetFileName_DocumentFactoryReturningEmptyStringAsExtension_ShouldReturnFileNameWithoutExtension()
        {
            _documentFactoryMock.SetupGet(x => x.FileExtension).Returns(string.Empty);
            var exporter = new FileStreamExporter(_documentFactoryMock.Object, _exportTaskMock.Object, () => DateTime.Parse("2018/01/05 12:00"));
            Assert.That(exporter.GetFileName(), Is.EqualTo("SomeExportName 20180105"));
        }

        [Test]
        public void GetFileName_DocumentFactoryReturningAnExtension_ShouldReturnFileNameWithExtension()
        {
            var exporter = new FileStreamExporter(_documentFactoryMock.Object, _exportTaskMock.Object, () => DateTime.Parse("2018/01/05 12:00"));

            _documentFactoryMock.SetupGet(x => x.FileExtension).Returns("xlsx");
            Assert.That(exporter.GetFileName(), Is.EqualTo("SomeExportName 20180105.xlsx"));

            _documentFactoryMock.SetupGet(x => x.FileExtension).Returns(".zip");
            Assert.That(exporter.GetFileName(), Is.EqualTo("SomeExportName 20180105.zip"));
        }

        [Test]
        public void ExportToStream_ShouldCallInterfaceMethodsInCorrectOrder()
        {
            var stream = new MemoryStream();
            var columnLabels = new[] { "col1", "col2", "col3" };
            var rowData = new[]
            {
                new object[] { true, 1, "r1" },
                new object[] { false, 2, "r2" }
            };

            var documentMock = new Mock<IDocument>();
            var sheetExportTaskMock = new Mock<ISheetExportTask>();
            var sheetMock = new Mock<ISheet>();
            _documentFactoryMock.SetupGet(x => x.FileExtension).Returns("xlsx");
            _documentFactoryMock.Setup(x => x.CreateDocument(stream)).Returns(documentMock.Object);
            _exportTaskMock.Setup(x => x.EnumSheetExportTasks()).Returns(new[] { sheetExportTaskMock.Object });
            documentMock.Setup(x => x.CreateSheet("MySheet A")).Returns(sheetMock.Object);
            sheetExportTaskMock.SetupGet(x => x.Name).Returns("MySheet A");
            sheetExportTaskMock.Setup(x => x.GetColumnLabels()).Returns(columnLabels);
            sheetExportTaskMock.Setup(x => x.EnumRowData()).Returns(rowData);

            var exporter = new FileStreamExporter(_documentFactoryMock.Object, _exportTaskMock.Object);
            exporter.ExportToStream(stream);

            _documentFactoryMock.Verify(x => x.CreateDocument(It.IsAny<Stream>()), Times.Once());
            _exportTaskMock.Verify(x => x.EnumSheetExportTasks(), Times.Once);
            documentMock.Verify(x => x.CreateSheet(It.IsAny<string>()), Times.Once);
            sheetExportTaskMock.Verify(x => x.GetColumnLabels(), Times.Once);
            sheetMock.Verify(x => x.AddHeaderRow(columnLabels), Times.Once);
            sheetMock.Verify(x => x.AddRow(It.IsAny<object[]>()), Times.Exactly(2));
            sheetMock.Verify(x => x.AddRow(rowData[0]), Times.Exactly(1));
            sheetMock.Verify(x => x.AddRow(rowData[1]), Times.Exactly(1));
        }
    }
}
