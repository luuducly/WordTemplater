using System;
using System.Collections;
using System.IO;
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
            int before = LiveCount(instances);

            // Measure right after the export without letting a GC (and the SKObject finalizers it triggers)
            // hide undisposed objects. If the no-GC region cannot be entered or is exceeded, the measurement
            // can only under-count leaks, never over-count them.
            int after;
            bool noGcRegion = TryStartNoGcRegion();
            try
            {
                using (template.Export(TestTemplates.Data(50, true))) { }
                after = LiveCount(instances);
            }
            finally
            {
                if (noGcRegion && GCSettings.LatencyMode == GCLatencyMode.NoGCRegion)
                {
                    GC.EndNoGCRegion();
                }
            }

            Assert.True(after - before == 0,
                $"{extraCellCode}: {after - before} SkiaSharp native objects left undisposed by one Export of 50 rows (live before={before}, after={after}, no-GC region={noGcRegion}).");
        }

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

        private static int LiveCount(IDictionary instances)
        {
            lock (instances)
            {
                return instances.Count;
            }
        }
    }
}
