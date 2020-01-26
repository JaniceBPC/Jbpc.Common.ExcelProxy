using System;
using System.Collections.Generic;
using System.Text;
using Jbpc.Common.ExtensionMethods;
using NUnit.Framework;

namespace Jbpc.Common.UnitTests.ExtensionMethods
{
    [TestFixture]
    class ExtensionMethodsOthers
    {
        [Test]
        public void ExpandedTypeName_SimpleType()
        {
            var name = typeof(ValueSourceAttribute).GetExpandedTypeName();
            var nameOf = nameof(ValueSourceAttribute);

            Assert.AreEqual(name, nameOf);
        }
        [Test]
        public void ExpandedTypeName_Generic_TypeName()
        {
            var name = typeof(List<ValueSourceAttribute>).GetExpandedTypeName();

            Assert.AreEqual(name, "List<ValueSourceAttribute>");
        }
        [Test]
        public void ExpandedTypeName_Nester_Generic_TypeNameX()
        {
            var name = typeof(List<List<ValueSourceAttribute>>).GetExpandedTypeName();

            Assert.AreEqual(name, "List<List<ValueSourceAttribute>>");
        }
    }
}
