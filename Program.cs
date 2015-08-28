using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace TabulateSmarterTestResults
{
    class Program
    {

        static string sSyntax =
@"This tool tabulates XML test results files in SmarterApp Test Results
Transmission Format into CSV files that match the SmarterApp Data Warehouse
Student Assessment format and Data Warehouse Item Level format.

Command-line parameters:
 -i <input file>
Filenames may include wildcards. If the file is a .zip then all .xml
files in the .zip will be processed. The -i parameter may be repeated
to process multiple input sources.

 -os <output file>
Output file for student test results. This will be a .csv formatted file that
contains test results including student information and test scores.

 -oi <output file>
Output file for item results. This will be a .csv formatted file that
contains item scores.";

        static void Main(string[] args)
        {
            try
            {
                List<string> inputFilenames = new List<string>();
                string osFilename = null;
                string oiFilename = null;
                bool help = false;

                for (int i=0; i<args.Length; ++i)
                {
                    switch (args[i])
                    {
                        case "-h":
                            help = true;
                            break;

                        case "-i":
                            {
                                ++i;
                                if (i >= args.Length) throw new ArgumentException("Invalid command line. '-i' option not followed by filename.");
                                inputFilenames.Add(args[i]);
                            }
                            break;

                        case "-os":
                            {
                                ++i;
                                if (i >= args.Length) throw new ArgumentException("Invalid command line. '-o' option not followed by filename.");
                                if (osFilename != null) throw new ArgumentException("Only one output file may be specified.");
                                string filename = Path.GetFullPath(args[i]);
                                osFilename = filename;
                            }
                            break;

                        case "-oi":
                            {
                                ++i;
                                if (i >= args.Length) throw new ArgumentException("Invalid command line. '-o' option not followed by filename.");
                                if (oiFilename != null) throw new ArgumentException("Only one output file may be specified.");
                                string filename = Path.GetFullPath(args[i]);
                                oiFilename = filename;
                            }
                            break;

                        default:
                            throw new ArgumentException(string.Format("Unknown command line option '{0}'. Use '-h' for syntax help.", args[i]));
                    }
                }

                if (help || args.Length == 0)
                {
                    Console.WriteLine(sSyntax);
                }

                else
                {
                    if (inputFilenames.Count == 0 || (osFilename == null && oiFilename == null)) throw new ArgumentException("Invalid command line. Use '-h' for syntax help");

                    if (osFilename != null)
                        Console.WriteLine("Writing student assessment results to: " + osFilename);
                    if (oiFilename != null)
                        Console.WriteLine("Writing item level results to: " + oiFilename);
                    using (ITestResultProcessor processor = new ToCsvProcessor(osFilename, oiFilename))
                    {
                        foreach (string filename in inputFilenames)
                        {
                            ProcessInputFilename(filename, processor);
                        }
                    }
                }
            }
            catch (Exception err)
            {
                Console.WriteLine();
#if DEBUG
                Console.WriteLine(err.ToString());
#else
                Console.WriteLine(err.Message);
#endif
            }

            if (Win32Interop.ConsoleHelper.IsSoleConsoleOwner)
            {
                Console.Write("Press any key to exit.");
                Console.ReadKey(true);
            }
        }

        static void ProcessInputFilename(string filenamePattern, ITestResultProcessor processor)
        {
            int count = 0;
            string directory = Path.GetDirectoryName(filenamePattern);
            if (string.IsNullOrEmpty(directory)) directory = Environment.CurrentDirectory;
            string pattern = Path.GetFileName(filenamePattern);
            foreach (string filename in Directory.GetFiles(directory, pattern))
            {
                switch (Path.GetExtension(filename).ToLower())
                {
                    case ".xml":
                        ProcessInputTrtFile(filename, processor);
                        break;

                    case ".zip":
                        ProcessInputZipFile(filename, processor);
                        break;

                    default:
                        throw new ArgumentException(string.Format("Input file '{0}' is of unsupported time. Only .xml and .zip are supported.", filename));
                }
                ++count;
            }
            if (count == 0) throw new ArgumentException(string.Format("Input file '{0}' not found!", filenamePattern));
        }

        static void ProcessInputTrtFile(string filename, ITestResultProcessor processor)
        {
            Console.WriteLine("Processing: " + filename);
            using (FileStream stream = new FileStream(filename, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                processor.ProcessResult(stream);
            }
            Console.WriteLine();
        }

        static void ProcessInputZipFile(string filename, ITestResultProcessor processor)
        {
            Console.WriteLine("Processing: " + filename);
            Console.WriteLine();
        }

    }

    interface ITestResultProcessor : IDisposable
    {
        void ProcessResult(Stream input);
    }
}
