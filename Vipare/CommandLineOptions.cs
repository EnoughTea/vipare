using System;
using System.Collections.Generic;
using System.Linq;
using CommandLine;
using CommandLine.Text;

namespace Vipare {
    /// <summary> Holds all available command line options. </summary>
    /// <remarks> Note that filenames are divided by space, so they must be quoted.</remarks>
    internal sealed class CommandLineOptions {
        public const string CurrentDirectoryMark = "<current directory>";

        [Option('o', "output", Required = false, DefaultValue = CurrentDirectoryMark,
            HelpText = "Defines output directory for exported images.")]
        public string OutputFolder { get; set; }

        [Option('f', "format", Required = false, DefaultValue = "png",
            HelpText = "Format indicates which export filter to use. Supply one of file formats supported by Visio export (bmp, dib, dwg, dxf, emf, emz, gif, htm, jpg, png, svg, svgz, tif, or wmf). Default preference settings for the specified filter will be used.")]
        public string Format { get; set; }

        /// <summary> Gets the passed visio diagrams. </summary>
        [ValueList(typeof(List<string>))]
        public IList<string> VisioFiles { get; set; }

        [ParserState]
        public IParserState LastParserState { get; set; }

        [HelpOption]
        public string GetUsage() {
            var help = new HelpText {
                Heading = "vipare is a simple command-line utility used to export pages from Visio diagrams.",
                AdditionalNewLineAfterOption = true,
                AddDashesToOption = true
            };

            HandleParsingErrorsInHelp(help);
            help.AddPostOptionsLine("Since a diagram can contain utility pages which you don't want to export, this tool ignores pages with names starting with one of the following 3 symbols: '`' '~' '!'.");
            help.AddPostOptionsLine(string.Empty);
            help.AddPostOptionsLine(string.Empty);
            help.AddPostOptionsLine("Usage examples");
            help.AddPostOptionsLine(string.Empty);
            help.AddPostOptionsLine("Export pages from one Visio diagram to png files and store them in the current directory:");
            help.AddPostOptionsLine("vipare \"file 1.vsdx\"");
            help.AddPostOptionsLine(string.Empty);
            help.AddPostOptionsLine("Export pages from 3 Visio diagrams to bmp files and store them in the specified output folder:");
            help.AddPostOptionsLine("vipare -f bmp -o \"D:\\Resulting images\\\" \"file 1.vsdx\" \"..\\another file 2.vsdx\" \"C:\\some folder\\other file 3.vsdx\"");
            help.AddOptions(this);
            return help;
        }

        private void HandleParsingErrorsInHelp(HelpText help) {
            if (LastParserState != null && LastParserState.Errors.Any()) {
                var errors = help.RenderParsingErrorsText(this, 2);
                if (!string.IsNullOrEmpty(errors)) {
                    help.AddPreOptionsLine(Environment.NewLine + "Error(s):");
                    help.AddPreOptionsLine(errors);
                }
            }
        }
    }
}
