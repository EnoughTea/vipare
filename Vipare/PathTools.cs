using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security;

namespace Vipare {
    /// <summary> Helper methods for path operations. </summary>
    internal static class PathTools {
        /// <summary> Converts a console file args into actual files. </summary>
        /// <remarks> Sequences like "test.vsdx *.vsdx .\test.vsdx" will be correctly recognized,
        /// no file duplicates will be created. </remarks>
        /// <param name="passedFileArgs">File arguments in the console input.</param>
        /// <returns>Sequence of files recognized in the console input.</returns>
        public static IEnumerable<FileInfo> GatherFiles(IList<string> passedFileArgs) {
            // Expand passed console input with possible wildcards and duplications into real unique file names:
            var expandedPaths = new HashSet<string>();
            foreach (var inputItem in passedFileArgs) {
                if (string.IsNullOrWhiteSpace(inputItem)) { continue; }

                string[] files;
                if (Directory.Exists(inputItem)) {
                    files = GetFiles(inputItem);
                } else {
                    string path = Path.GetDirectoryName(inputItem);
                    string filename = Path.GetFileName(inputItem);
                    if (filename == "*") {
                        filename = "*.*";
                    }

                    files = GetFiles(path, filename);
                }

                foreach (var file in files) {
                    expandedPaths.Add(file);
                }
            }
            // Expanded file names can be safely turned into FileInfos, since GetFiles()
            // only return real files which can be accessed.
            return expandedPaths.Select(expandedPath => new FileInfo(expandedPath));
        }

        private const string DotSlash = @".\";

        private static string[] GetFiles(string path, string pattern = null) {
            string fixedPath = (string.IsNullOrWhiteSpace(path)) ? DotSlash : path;

            try {
                string[] paths = (pattern == null)
                    ? Directory.GetFiles(fixedPath)
                    : Directory.GetFiles(fixedPath, pattern);
                if (!path.StartsWith(DotSlash, StringComparison.Ordinal)) { RemoveDotSlash(paths); }

                return paths;
            } catch (SecurityException) {
                Console.WriteLine($"Access denied for path '{fixedPath}'.");
            } catch (UnauthorizedAccessException) {
                Console.WriteLine($"Access denied for path '{fixedPath}'.");
            }

            return new string[0];
        }

        private static void RemoveDotSlash(IList<string> paths) {
            for (int i = 0; i < paths.Count; i++) {
                string path = paths[i];
                if (!string.IsNullOrEmpty(path) && path.StartsWith(DotSlash, StringComparison.Ordinal)) {
                    paths[i] = path.Substring(DotSlash.Length);
                }
            }
        }
    }
}
