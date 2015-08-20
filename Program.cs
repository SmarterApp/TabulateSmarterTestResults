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
            @"Replace this with syntax description";

        static void Main(string[] args)
        {
            try
            {
                List<string> inputFilenames = new List<string>();
                string outputFilename = null;
                bool help = false;

                for (int i=0; i<args.Length; ++i)
                {
                    switch (args[i])
                    {
                        case "-i":
                            {
                                ++i;
                                if (i >= args.Length) throw new ArgumentException("Invalid command line. '-i' option not followed by filename.");
                                inputFilenames.Add(args[i]);
                            }
                            break;

                        case "-o":
                            {
                                ++i;
                                if (i >= args.Length) throw new ArgumentException("Invalid command line. '-o' option not followed by filename.");
                                if (outputFilename != null) throw new ArgumentException("Only one output file may be specified.");
                                string filename = Path.GetFullPath(args[i]);
                                outputFilename = filename;
                            }
                            break;

                        default:
                            throw new ArgumentException(string.Format("Unknown command line option '{0}'. Use '-h' for syntax help."));
                    }
                }

                if (help)
                {
                    Console.WriteLine(sSyntax);
                }

                if (inputFilenames.Count == 0 || outputFilename == null) throw new ArgumentException("Invalid command line. Use '-h' for syntax help");

                Console.WriteLine("Writing result to: " + outputFilename);
                using (ITestResultProcessor processor = new ToCsvProcessor(outputFilename))
                {
                    foreach (string filename in inputFilenames)
                    {
                        ProcessInputFilename(filename, processor);
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
