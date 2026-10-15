using System;
using System.IO;

namespace WordTemplater.Tests
{
    /// <summary>
    /// Locates the repository root (the folder containing src/WordTemplater.sln) and well-known asset paths.
    /// </summary>
    internal static class TestPaths
    {
        private static readonly Lazy<string> _repoRoot = new Lazy<string>(FindRepoRoot);

        public static string RepoRoot => _repoRoot.Value;

        public static string TemplatesDir => Path.Combine(RepoRoot, "src", "WordTemplater.Example", "Templates");

        public static string DataSamplesDir => Path.Combine(RepoRoot, "src", "WordTemplater.Example", "DataSamples");

        /// <summary>
        /// Source folder for approved snapshots (not the bin output folder), so new/received files land in the repo.
        /// </summary>
        public static string SnapshotsDir => Path.Combine(RepoRoot, "tests", "WordTemplater.Tests", "Snapshots");

        public static string Template(string fileName) => Path.Combine(TemplatesDir, fileName);

        public static string DataSample(string fileName) => Path.Combine(DataSamplesDir, fileName);

        private static string FindRepoRoot()
        {
            var dir = new DirectoryInfo(AppContext.BaseDirectory);
            while (dir != null && !File.Exists(Path.Combine(dir.FullName, "src", "WordTemplater.sln")))
            {
                dir = dir.Parent;
            }
            if (dir == null)
            {
                throw new InvalidOperationException($"Could not locate repository root (src/WordTemplater.sln) above {AppContext.BaseDirectory}");
            }
            return dir.FullName;
        }
    }
}
