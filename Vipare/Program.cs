using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using CommandLine;
using Microsoft.Office.Interop.Visio;
using Path = System.IO.Path;
using System.Diagnostics.Contracts;

namespace Vipare {
    class Program {
        static void Main(string[] args) {
            bool debug = Debugger.IsAttached;
            if (!debug) { AppDomain.CurrentDomain.UnhandledException += OnUnhandledException; }

            var options = new CommandLineOptions();
            if (Parser.Default.ParseArgumentsStrict(args, options)) {
                ProcessDiagrams(options);

                if (debug) { Console.ReadKey(); }
            }
        }

        private static void ProcessDiagrams(CommandLineOptions options) {
            // Read and verify passed command-line options:
            string outputFolder = Directory.GetCurrentDirectory();
            if (!string.IsNullOrWhiteSpace(options.OutputFolder) &&
                options.OutputFolder != CommandLineOptions.CurrentDirectoryMark) {
                outputFolder = Directory.CreateDirectory(options.OutputFolder).FullName;
            }

            var diagramFiles = PathTools.GatherFiles(options.VisioFiles).ToArray();

            string format = options.Format.TrimStart('.');
            if (!SupportedFormats.Contains(format)) {
                string expected = string.Join(", ", SupportedFormats.ToArray());
                throw new ArgumentException($"Passed unknown format '{format}', expected one of {{ {expected} }}.");
            }

            // We have options in check, so start actual export now:
            foreach (var diagramFile in diagramFiles) {
                Console.WriteLine($"Exporting '{diagramFile.FullName}':");
                ExportPages(diagramFile.FullName, outputFolder, format);
                Console.WriteLine($"Finished '{diagramFile}'.");
                Console.WriteLine();
            }
        }

        /// <summary> Exports all appropriate pages from the selected Visio diagram. </summary>
        /// <param name="diagramFile">Visio diagram file.</param>
        /// <param name="outputFolder">Output folder.</param>
        /// <param name="format">Export format.</param>
        private static void ExportPages(string diagramFile, string outputFolder, string format) {
            Contract.Requires(diagramFile != null);
            Contract.Requires(outputFolder != null);
            Contract.Requires(format != null);

            InvisibleApp app = null;
            Documents docs = null;
            Document doc = null;
            Pages pages = null;
            try {
                app = new InvisibleApp { ShowChanges = false };
                docs = app.Documents;
                try {
                    doc = docs.Open(diagramFile);
                } catch (COMException e) {
                    Console.WriteLine($"Could not open '{diagramFile}': {e.Message}");
                    return;
                }

                pages = doc.Pages;
                // Iterators and COM are best kept separated.
                for (int i = 1; i <= pages.Count; i++) {
                    ExportPage(pages, i, outputFolder, format);
                }

                doc.Close();
                app.Quit();
            } finally {
                if (pages != null) Marshal.ReleaseComObject(pages);
                if (doc != null) Marshal.ReleaseComObject(doc);
                if (docs != null) Marshal.ReleaseComObject(docs);
                if (app != null) Marshal.ReleaseComObject(app);
            }
        }

        /// <summary> Exports page with given index from the specified Visio pages collection. </summary>
        private static void ExportPage(Pages pages, int pageIndex, string outputFolder, string format) {
            Page page = null;
            try {
                page = pages[pageIndex];
                string imageName = page.NameU;
                if (ShouldIgnorePage(page)) {
                    Console.WriteLine("{0,-2}", $"Ignoring '{imageName}'.");
                    return;
                }

                string imageFileName = Path.Combine(outputFolder, imageName);
                page.Export(imageFileName + "." + format);
                Console.WriteLine("{0,-2}", $"'{imageName}' done.");
            } finally {
                if (page != null) Marshal.ReleaseComObject(page);
            }
        }

        private static bool ShouldIgnorePage(Page page) {
            string imageName = page.NameU;
            return (imageName.StartsWith("~", StringComparison.Ordinal) ||
                imageName.StartsWith("`", StringComparison.Ordinal) ||
                imageName.StartsWith("!", StringComparison.Ordinal));
        }

        private static void OnUnhandledException(object sender, UnhandledExceptionEventArgs arg) {
            var e = arg.ExceptionObject as Exception;
            if (e != null) {
                var fg = Console.ForegroundColor;
                Console.ForegroundColor = ConsoleColor.Red;
                Console.Write("Program halt: ");
                Console.WriteLine(e.Message);
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine(e.StackTrace);
                Console.ForegroundColor = fg;
            }

            Environment.Exit(1);
        }

        internal static readonly HashSet<string> SupportedFormats = new HashSet<string> {
            "bmp", "dib", "dwg", "dxf", "emf", "emz", "gif", "htm", "jpg", "png", "svg", "svgz", "tif", "wmf"
        };
    }
}
