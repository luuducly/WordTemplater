using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime;
using System.Runtime.CompilerServices;
using WordTemplater;
using Xunit;

namespace WordTemplater.Tests
{
    /// <summary>
    /// Tests in this collection take process-wide measurements (managed heap, live SkiaSharp handles),
    /// so they must not run in parallel with anything else.
    /// </summary>
    [CollectionDefinition("NonParallel", DisableParallelization = true)]
    public class NonParallelCollection
    {
    }

    [Collection("NonParallel")]
    public class ResourceLifetimeTests
    {
        private const double OneMb = 1024 * 1024;

        [Fact]
        public void Export_DoesNotRetainDataOrDocument_WhenTemplateIsKept()
        {
            using var templateStream = new MemoryStream(TestTemplates.TableTemplate());
            using var template = new WordTemplate(templateStream);
            using (template.Export(TestTemplates.Data(5))) { } // warm-up

            long before = HeapAfterFullGc();
            var data = ExportAndReturnWeakData(template, 2000);
            long after = HeapAfterFullGc();
            GC.KeepAlive(template);

            double growthMb = (after - before) / OneMb;
            Assert.False(data.IsAlive,
                $"Data JObject passed to Export is still reachable after Export returned and the output was disposed (template kept alive); managed heap grew {growthMb:F1} MB.");
            Assert.True(growthMb < 5,
                $"Managed heap grew {growthMb:F1} MB after one Export(2000 rows) with the output disposed (limit 5 MB): the template retains per-export state.");
        }

        [Theory]
        [InlineData("Code:barcode()")]
        [InlineData("Code:qrcode()")]
        public void Export_DisposesNativeSkiaObjects(string extraCellCode)
        {
            var instances = SkiaHandleInstances();
            using var templateStream = new MemoryStream(TestTemplates.TableTemplate(extraCellCode));
            using var template = new WordTemplate(templateStream);
            using (template.Export(TestTemplates.Data(5, true))) { } // warm-up: static Skia/QR/barcode caches

            HeapAfterFullGc(); // drain finalizers of anything left over from the warm-up
            var before = LiveByType(instances);

            // Measure right after the export without letting a GC (and the SKObject finalizers it triggers)
            // hide undisposed objects. If the no-GC region cannot be entered or is exceeded, the measurement
            // can only under-count leaks, never over-count them.
            const int rows = 50;
            Dictionary<string, int> after;
            bool noGcRegion = TryStartNoGcRegion();
            try
            {
                using (template.Export(TestTemplates.Data(rows, true))) { }
                after = LiveByType(instances);
            }
            finally
            {
                if (noGcRegion && GCSettings.LatencyMode == GCLatencyMode.NoGCRegion)
                {
                    GC.EndNoGCRegion();
                }
            }

            var allowance = ThirdPartyResidue(extraCellCode, rows);
            var delta = after.Keys.Union(before.Keys)
                .Select(type => (Type: type, Delta: after.GetValueOrDefault(type) - before.GetValueOrDefault(type)))
                .Where(x => x.Delta != 0)
                .OrderBy(x => x.Type, StringComparer.Ordinal)
                .ToList();
            var violations = delta.Where(x => x.Delta > allowance.GetValueOrDefault(x.Type)).ToList();

            Assert.True(violations.Count == 0,
                $"{extraCellCode}: SkiaSharp native objects left undisposed by one Export of {rows} rows: "
                + string.Join(", ", delta.Select(x => $"{x.Type} +{x.Delta}" + (allowance.ContainsKey(x.Type) ? $" (allowed {allowance[x.Type]})" : string.Empty)))
                + $" [live before={before.Values.Sum()}, after={after.Values.Sum()}, no-GC region={noGcRegion}].");
        }

        /// <summary>
        /// Objects that remain live even when WordTemplater disposes everything it creates.
        /// BarcodeLib 3.0.3 itself leaves one SKImage per Barcode.Encode call undisposed (even when the Barcode,
        /// the returned SKImage and the encoded SKData are all disposed), and its label font holds one SKTypeface
        /// wrapper (re-created after a GC, but it does not grow with the row count).
        /// Measured with a probe that disposes every object WordTemplater owns; QR code generation leaves nothing.
        /// Remove this allowance if BarcodeLib is upgraded/replaced with one that does not leak.
        /// </summary>
        private static Dictionary<string, int> ThirdPartyResidue(string extraCellCode, int rows) =>
            extraCellCode.Contains("barcode", StringComparison.OrdinalIgnoreCase)
                ? new Dictionary<string, int> { ["SKImage"] = rows, ["SKTypeface"] = 1 }
                : new Dictionary<string, int>();

        [Fact]
        public void Export_RepeatedOnCachedTemplate_DoesNotGrowManagedHeap()
        {
            using var templateStream = new MemoryStream(TestTemplates.TableTemplate());
            using var template = new WordTemplate(templateStream);

            for (int i = 0; i < 30; i++)
            {
                using (template.Export(TestTemplates.Data(20))) { }
            }
            long afterWarmUp = HeapAfterFullGc();

            for (int i = 30; i < 300; i++)
            {
                using (template.Export(TestTemplates.Data(20))) { }
            }
            long afterSoak = HeapAfterFullGc();
            GC.KeepAlive(template);

            double growthMb = (afterSoak - afterWarmUp) / OneMb;
            Assert.True(growthMb < 3,
                $"Managed heap grew {growthMb:F2} MB between export #30 and #300 on one cached template (limit 3 MB).");
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private static WeakReference ExportAndReturnWeakData(WordTemplate template, int rows)
        {
            var data = TestTemplates.Data(rows);
            using (template.Export(data)) { }
            return new WeakReference(data);
        }

        private static bool TryStartNoGcRegion()
        {
            // The allowed budget depends on the GC configuration; take the largest one that is accepted.
            foreach (var megabytes in new long[] { 240, 200, 128, 64 })
            {
                try
                {
                    return GC.TryStartNoGCRegion(megabytes * 1024 * 1024);
                }
                catch (ArgumentOutOfRangeException)
                {
                }
            }
            return false;
        }

        private static long HeapAfterFullGc()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            return GC.GetTotalMemory(true);
        }

        /// <summary>
        /// SkiaSharp (pinned to 2.88.6) registers every live SKObject handle in the static
        /// SkiaSharp.HandleDictionary.instances dictionary and removes it on Dispose/finalization.
        /// </summary>
        private static IDictionary SkiaHandleInstances()
        {
            var handleDictionary = typeof(SkiaSharp.SKObject).Assembly.GetType("SkiaSharp.HandleDictionary");
            Assert.True(handleDictionary != null, "SkiaSharp.HandleDictionary type not found (SkiaSharp version changed?).");
            var field = handleDictionary!.GetField("instances", BindingFlags.NonPublic | BindingFlags.Static);
            Assert.True(field != null, "SkiaSharp.HandleDictionary.instances field not found (SkiaSharp version changed?).");
            var instances = field!.GetValue(null) as IDictionary;
            Assert.True(instances != null, "SkiaSharp.HandleDictionary.instances is not an IDictionary.");
            return instances!;
        }

        private static Dictionary<string, int> LiveByType(IDictionary instances)
        {
            var result = new Dictionary<string, int>(StringComparer.Ordinal);
            lock (instances)
            {
                foreach (DictionaryEntry entry in instances)
                {
                    var target = entry.Value is WeakReference weak ? weak.Target : entry.Value;
                    var type = target?.GetType().Name ?? "(collected)";
                    result[type] = result.GetValueOrDefault(type) + 1;
                }
            }
            return result;
        }
    }
}
