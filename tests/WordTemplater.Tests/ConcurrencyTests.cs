using System;
using System.Collections.Concurrent;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using Newtonsoft.Json.Linq;
using WordTemplater;
using Xunit;

namespace WordTemplater.Tests
{
    public class ConcurrencyTests
    {
        private const int Exports = 48;
        private const int Rows = 200;

        [Fact]
        public void ConcurrentExports_OnSharedTemplate_ProduceSameOutputAsSequential()
        {
            using var templateStream = new MemoryStream(TestTemplates.TableTemplate());
            using var template = new WordTemplate(templateStream);
            var baseline = BodyText(template, TestTemplates.Data(Rows));

            RunConcurrently(template, baseline, _ => TestTemplates.Data(Rows));
        }

        [Fact]
        public void ConcurrentExports_SharingSameDataObject_ProduceSameOutput()
        {
            using var templateStream = new MemoryStream(TestTemplates.LoopIfTemplate());
            using var template = new WordTemplate(templateStream);
            // Baseline from a separate instance so the shared object is untouched when the parallel exports start.
            var baseline = BodyText(template, TestTemplates.Data(Rows));

            var shared = TestTemplates.Data(Rows);
            RunConcurrently(template, baseline, _ => shared);
        }

        private static void RunConcurrently(WordTemplate template, string baseline, Func<int, JObject> dataFor)
        {
            int ok = 0, mismatches = 0;
            var errors = new ConcurrentQueue<Exception>();

            Parallel.For(0, Exports, new ParallelOptions { MaxDegreeOfParallelism = 8 }, i =>
            {
                try
                {
                    if (BodyText(template, dataFor(i)) == baseline) Interlocked.Increment(ref ok);
                    else Interlocked.Increment(ref mismatches);
                }
                catch (Exception ex)
                {
                    errors.Enqueue(ex);
                }
            });

            var firstError = errors.TryPeek(out var ex0) ? $" First exception: {ex0.GetType().Name}: {ex0.Message.Split('\n')[0]}" : string.Empty;
            Assert.True(mismatches == 0 && errors.IsEmpty,
                $"{Exports} concurrent exports: {ok} matched the sequential baseline, {mismatches} produced different output, {errors.Count} threw.{firstError}");
        }

        private static string BodyText(WordTemplate template, JObject data)
        {
            using var output = template.Export(data);
            Assert.NotNull(output);
            using var doc = WordprocessingDocument.Open(output!, false);
            return doc.MainDocumentPart!.Document.Body!.InnerText;
        }
    }
}
