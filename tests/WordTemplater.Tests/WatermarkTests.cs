using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace WordTemplater.Tests
{
    public class WatermarkTests
    {
        [Fact]
        public void ParagraphInnerXml_ShouldParseValidXml()
        {
            var p = new Paragraph();
            p.InnerXml = @"<w:r xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""><w:t>Hello</w:t></w:r>";
            Assert.Contains("Hello", p.InnerText);
        }
    }
}
