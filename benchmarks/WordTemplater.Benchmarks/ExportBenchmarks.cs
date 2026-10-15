using BenchmarkDotNet.Attributes;
using Newtonsoft.Json.Linq;
using WordTemplater.Tests;

namespace WordTemplater.Benchmarks
{
    /// <summary>
    /// Cached-template scenario: templates, WordTemplate instances and data are built once;
    /// each benchmark measures a single Export (output stream disposed).
    /// </summary>
    [MemoryDiagnoser]
    [SimpleJob(warmupCount: 2, iterationCount: 5)]
    public class ExportBenchmarks
    {
        private WordTemplate _table = null!;
        private WordTemplate _loopIf = null!;
        private WordTemplate _nested = null!;
        private WordTemplate _tableWithFooter = null!;
        private JObject _data1000 = null!;
        private JObject _data2000 = null!;
        private JObject _data4000 = null!;
        private JObject _data8000 = null!;
        private JObject _nestedData = null!;

        [GlobalSetup]
        public void Setup()
        {
            _table = Load(TestTemplates.TableTemplate());
            _loopIf = Load(TestTemplates.LoopIfTemplate());
            _nested = Load(TestTemplates.NestedLoopTemplate());
            _tableWithFooter = Load(TestTemplates.TableWithFooterTemplate());
            _data1000 = TestTemplates.Data(1000);
            _data2000 = TestTemplates.Data(2000);
            _data4000 = TestTemplates.Data(4000);
            _data8000 = TestTemplates.Data(8000);
            _nestedData = TestTemplates.NestedData(100, 20);
        }

        [GlobalCleanup]
        public void Cleanup()
        {
            _table.Dispose();
            _loopIf.Dispose();
            _nested.Dispose();
            _tableWithFooter.Dispose();
        }

        [Benchmark]
        public long Table_1000() => Export(_table, _data1000);

        [Benchmark]
        public long Table_8000() => Export(_table, _data8000);

        [Benchmark]
        public long LoopIf_1000() => Export(_loopIf, _data1000);

        [Benchmark]
        public long LoopIf_4000() => Export(_loopIf, _data4000);

        [Benchmark]
        public long Nested_100x20() => Export(_nested, _nestedData);

        [Benchmark]
        public long TableWithFooter_2000() => Export(_tableWithFooter, _data2000);

        private static WordTemplate Load(byte[] templateBytes) => new WordTemplate(new MemoryStream(templateBytes));

        private static long Export(WordTemplate template, JObject data)
        {
            using var output = template.Export(data) ?? throw new InvalidOperationException("Export returned null.");
            return output.Length;
        }
    }
}
