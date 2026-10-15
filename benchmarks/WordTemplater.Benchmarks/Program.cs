using BenchmarkDotNet.Running;

BenchmarkSwitcher.FromAssembly(typeof(WordTemplater.Benchmarks.ExportBenchmarks).Assembly).Run(args);
