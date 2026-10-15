using System.IO;
using Newtonsoft.Json.Linq;
using WordTemplater;
using Xunit;

namespace WordTemplater.Tests
{
    public class WordTemplateLifecycleTests
    {
        [Fact]
        public void MultipleExportsOnSameTemplateInstance_ShouldNotCorruptState()
        {
            var dir = new DirectoryInfo(AppContext.BaseDirectory);
            while (dir != null && !File.Exists(Path.Combine(dir.FullName, "src", "WordTemplater.sln")))
            {
                dir = dir.Parent;
            }
            var templatePath = Path.Combine(dir!.FullName, "src", "WordTemplater.Example", "Templates", "Template.docx");

            Assert.True(File.Exists(templatePath), $"Template file not found at: {templatePath}");

            var data1 = new JObject { ["Author"] = "Alice" };
            var data2 = new JObject { ["Author"] = "Bob" };

            using var templateStream = File.OpenRead(templatePath);
            using var wordTemplate = new WordTemplate(templateStream);

            using var stream1 = wordTemplate.Export(data1);
            Assert.NotNull(stream1);
            Assert.True(stream1.Length > 0);

            // Second export must not fail with ObjectDisposedException or corrupt render contexts
            using var stream2 = wordTemplate.Export(data2);
            Assert.NotNull(stream2);
            Assert.True(stream2.Length > 0);
        }
    }
}
