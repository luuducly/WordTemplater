using System.Collections.Generic;
using Newtonsoft.Json.Linq;
using WordTemplater;
using Xunit;

namespace WordTemplater.Tests
{
    public class DataScopeTests
    {
        [Fact]
        public void Resolve_LoopItem_MatchesLegacyWriteMetadataThenResolve()
        {
            var items = new[]
            {
                "{}",
                "{\"_INDEX\":\"X\"}",
                "{\"_index\":\"X\"}",
                "{\"_Index\":\"Y\"}",
                "{\"_index\":\"X\",\"_Index\":\"Y\"}",
                "{\"_Index\":\"Y\",\"_index\":\"X\"}",
                "{\"Name\":\"n\",\"_LAST\":\"L\"}",
            };
            var paths = new[] { "_index", "_INDEX", "_Index", "_first", "_FIRST", "_last", "_LAST", "_even", "_odd", "Name", "missing" };
            var positions = new[] { (Index: 0, Count: 3), (Index: 1, Count: 3), (Index: 2, Count: 3) };

            var mismatches = new List<string>();
            foreach (var json in items)
            {
                var item = JObject.Parse(json);
                foreach (var path in paths)
                {
                    foreach (var (index, count) in positions)
                    {
                        var expected = LegacyResolve(item, index, count, path)?.ToString();
                        var actual = new DataScope(item, index, count).Resolve(path)?.ToString();
                        if (expected != actual)
                            mismatches.Add($"item {json}, path {path}, ({index},{count}): expected {expected ?? "null"}, got {actual ?? "null"}");
                    }
                }
            }

            Assert.True(mismatches.Count == 0, string.Join("\n", mismatches));
        }

        // The behavior before DataScope: metadata written into the item with the indexer, then the field resolved against it.
        private static JToken? LegacyResolve(JObject item, int index, int count, string path)
        {
            var clone = (JObject)item.DeepClone();
            clone["_index"] = index + 1;
            clone["_first"] = index == 0;
            clone["_last"] = index == count - 1;
            clone["_even"] = (index + 1) % 2 == 0;
            clone["_odd"] = (index + 1) % 2 == 1;
            return Utils.ResolveToken(clone, path);
        }
    }
}
