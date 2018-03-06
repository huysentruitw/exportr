using System;
using Exportr.SheetExport;
using NUnit.Framework;

namespace Exportr.Tests.SheetExport
{
    [TestFixture]
    public class ColumnAttributeTests
    {
        [Test]
        public void Constructor_PassName_ShouldSetNameProperty()
        {
            var name = Guid.NewGuid().ToString("N");
            var attribute = new ColumnAttribute(name);
            Assert.That(attribute.Name, Is.EqualTo(name));
        }
    }
}
