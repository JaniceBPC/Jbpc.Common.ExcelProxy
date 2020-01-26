using System.Collections.Generic;
using NUnit.Framework;
using Jbpc.Common;
namespace Jbpc.Common.UnitTests
{
    class TestSettings
    {
        public string Name { get; set; }
        public int Id { get; set; }
        public float Weight { get; set; }
        public string[] RideLegs { get; set; }
    }
    [TestFixture]
    public class SettingsTests
    {
        private const string SettingsFileName = "MyTestSettings";
        private TestSettings testSettings;
        [SetUp]
        public void Setup()
        {
            testSettings = new TestSettings()
            {
                Name = "Janice",
                Id = 123,
                Weight = 123,
                RideLegs = new string[] {"Gateway to Bear", "Bear to Gateway"}
            };
        }

        [Test]
        public void Test1()
        {
            var saveSettings = new SettingsJsonPersistence<TestSettings>(SettingsFileName);

            saveSettings.SaveSettings(testSettings);

            var savedSettings = saveSettings.LoadSettings();

            Assert.AreEqual(testSettings.Name,savedSettings.Name);
            Assert.AreEqual(testSettings.Id,savedSettings.Id);
            Assert.AreEqual(testSettings.Weight,savedSettings.Weight);
            Assert.AreEqual(testSettings.RideLegs, savedSettings.RideLegs);
        }
    }
}