using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.IO.Compression;

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
contains test results including student information and test scores.  Include the
.csv file extension.

 -oi <output file>
Output file for item results. This will be a .csv formatted file that
contains item scores.  Include the .csv file extension.

 -hid <passphrase>
Hash the Student ID using the specified passphrase as a key. And place the
result in the AlternateSSID (ExternalSSID) field. This is usually combined
with de-identification See below for details of the hashing algorithm.

 -did <flags>
De-identify the data by setting certain fields to blank. The flags indicate
which fields should be set to blank. See below for details.

 -fmt <format>
Specifies the output format. Default is 'dw' which is the format accepted
by the Data Warehouse. The 'all' format includes all fields defined in the
Smarter Balanced ""Test Results Logical Data Model"". The 'allshort' format
includes all fields but only includes the student response if it is 10
characters or shorter in length.

 -notxl
When using the 'all' format (-fmt all) studentID and AlternateSSID fields are
padded with a trailing tab character to prevent Microsoft Excel from treating
them as numbers. This option suppresses the tab padding for use by other CSV
readers. The data warehouse format never pads and this option has no effect
in that mode.

Student ID Hash:
The Student ID hash is prepared as follows:
 1. The passphrase is encoded into UTF8 and hashed into a 160-bit key using
    the SHA1 algorithm. Case is preserved.
 2. The student id is encoded into UTF8 and hashed into a 160-bit digest
    using the key and the HMAC-SHA1 algorithm.
 3. The resulting hash from step 2 is converted to an upper-case hexadecimal
    string.
 4. The hexadecimal string is placed in the AlternateSSID (ExternalSSID)
    field replacing any existing contents.

The De-Identification Option (-did):
One or more of the following flags should follow the -did option.
A space should be between -did and the flags. No space should be between
the flags. For example, ""-did inb"" When values are removed, the fields
are set to blank. Since the sensitivity of student groups is unknown any
de-identification option will cause student groups to be removed.

 i  Replace the StudentID with the AlternateSSID. If the AlternateSSID
    is blank then StudentID will be blank in both the student and item
    files. If the -hid is specified then the the newly generated
    AlternateSSID will be used for both AlternateSSID and StudentID.
 n  Remove first, middle, and last names.
 b  Remove birthdate.
 d  Remove demographic information (gender, race, ethnicity)
 s  Remove school and district information. Also removes session location
    id and name.

Syntax example:
TabulateSmarterTestResults.exe -i testresults.zip -os studentresults.csv -oi itemresults.csv -hid smarter -did inbds -fmt all
";

        static int s_ErrorCount = 0;

        static void Main(string[] args)
        {
            try
            {
                List<string> inputFilenames = new List<string>();
                string osFilename = null;
                string oiFilename = null;
                bool notExcel = false;
                string hashPassPhrase = null;
                DIDFlags didFlags = DIDFlags.None;
                OutputFormat outputFormat = OutputFormat.Dw;
                int maxResponse = 0;
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

                        case "-notxl":
                            notExcel = true;
                            break;

                        case "-hid":
                            {
                                ++i;
                                if (i >= args.Length) throw new ArgumentException("Invalid command line. '-hid' option not followed by passphrase.");
                                hashPassPhrase = args[i];
                            }
                            break;

                        case "-did":
                            {
                                ++i;
                                if (i >= args.Length) throw new ArgumentException("Invalid command line. '-did' option not followed by flags.");
                                foreach (char c in args[i])
                                {
                                    switch (Char.ToLowerInvariant(c))
                                    {
                                        case 'i':
                                            didFlags |= DIDFlags.Id;
                                            break;
                                                
                                        case 'n':
                                            didFlags |= DIDFlags.Name;
                                            break;

                                        case 'b':
                                            didFlags |= DIDFlags.Birthdate;
                                            break;

                                        case 'd':
                                            didFlags |= DIDFlags.Demographics;
                                            break;

                                        case 's':
                                            didFlags |= DIDFlags.School;
                                            break;

                                        default:
                                            throw new ArgumentException(string.Format("Invalid command line. '.did' flag '{0}' is undefined.", c));
                                    }
                                }

                            }
                            break;

                        case "-fmt":
                            ++i;
                            if (i >= args.Length) throw new ArgumentException("Invalid command line. '-fmt' option not followed by format type.");
                            switch (args[i])
                            {
                                case "dw":
                                    outputFormat = OutputFormat.Dw;
                                    break;

                                case "all":
                                    outputFormat = OutputFormat.All;
                                    break;

                                case "allshort":
                                    outputFormat = OutputFormat.All;
                                    maxResponse = 10;
                                    break;

                                default:
                                    throw new ArgumentException(string.Format("Invalid command line. Output format '{0}' is unknown.", args[i]));
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
                    using (ToCsvProcessor processor = new ToCsvProcessor(osFilename, oiFilename, outputFormat))
                    {
                        processor.HashPassPhrase = hashPassPhrase;
                        processor.DIDFlags = didFlags;
                        processor.MaxResponse = maxResponse;
                        processor.NotExcel = notExcel;

                        foreach (string filename in inputFilenames)
                        {
                            ProcessInputFilename(filename, processor);
                        }
                    }
                }

                if (s_ErrorCount > 0)
                {
                    Console.Error.WriteLine("{0} total errors", s_ErrorCount);
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
                try
                {
                    processor.ProcessResult(stream);
                }
                catch (Exception err)
                {
                    Console.Error.WriteLine("Error processing input file '{0}\r\n{1}\r\n", filename, err.Message);
                    ++s_ErrorCount;
                }

            }
            Console.WriteLine();
        }

        static void ProcessInputZipFile(string filename, ITestResultProcessor processor)
        {
            Console.WriteLine("Processing: " + filename);
            using (ZipArchive zip = ZipFile.Open(filename, ZipArchiveMode.Read))
            {
                foreach(ZipArchiveEntry entry in zip.Entries)
                {
                    // Must not be folder (empty name) and must have .xml extension
                    if (!string.IsNullOrEmpty(entry.Name) && Path.GetExtension(entry.Name).Equals(".xml", StringComparison.OrdinalIgnoreCase))
                    {
                        Console.WriteLine("   Processing: " + entry.FullName);
                        using (Stream stream = entry.Open())
                        {
                            try
                            {
                                processor.ProcessResult(stream);
                            }
                            catch (Exception err)
                            {
                                Console.Error.WriteLine("Error processing input file '{0}/{1}\r\n{2}\r\n", filename, entry.FullName, err.Message);
                                ++s_ErrorCount;
                            }
                        }
                    }
                }
            }
            Console.WriteLine();
        }

    }

    [Flags]
    enum DIDFlags : int
    {
        None = 0,
        Id = 1,             // Student ID
        Name = 2,           // Student Name
        Birthdate = 4,
        Demographics = 8,   // Sex, Race, Ethnicity
        School = 16          // School and districtID or ExternalSSID is unaffected
    }

    enum OutputFormat : int
    {
        Dw = 0,             // Data Warehouse Format
        All = 1             // All fields format
    }

    interface ITestResultProcessor : IDisposable
    {
        void ProcessResult(Stream input);
    }
}
