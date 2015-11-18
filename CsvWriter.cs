using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Diagnostics;

namespace Parse
{
    class CsvWriter : IDisposable
    {
        static readonly UTF8Encoding sUTF8_NoBom = new UTF8Encoding(false);
        static readonly char[] sCsvSpecialChars = new char[] { ',', '"', '\r', '\n' };

        const int cBufSize = 8192;
        TextWriter mWriter;


        public CsvWriter(TextWriter Writer)
        {
            mWriter = Writer;
        }

        public CsvWriter(Stream stream, Encoding encoding = null, bool leaveOpen = false)
        {
            if (encoding == null) encoding = sUTF8_NoBom;
            mWriter = new StreamWriter(stream, encoding, cBufSize, leaveOpen);
        }

        public CsvWriter(String path, bool append, Encoding encoding = null)
        {
            if (encoding == null) encoding = sUTF8_NoBom;
            mWriter = new StreamWriter(path, append, encoding, cBufSize);
        }

        public void Write(string[] values)
        {
            for (int i = 0; i < values.Length; ++i)
            {
                string value = values[i];
                if (value.IndexOfAny(sCsvSpecialChars) >= 0)
                {
                    if (value.IndexOf('\r') >= 0) value = value.Replace("\r", "");  // For Excel - substitutes \n for \r\n and newlines are tolerated.
                    if (value.IndexOf('"') >= 0) value = value.Replace("\"", "\"\"");
                    mWriter.Write('"');
                    mWriter.Write(value);
                    mWriter.Write('"');
                }
                else
                {
                    mWriter.Write(value);
                }
                if (i < values.Length - 1)
                    mWriter.Write(',');
            }
            mWriter.WriteLine();
        }

        public void Dispose()
        {
            Dispose(true);
        }

        ~CsvWriter()
        {
            Dispose(false);
        }

        private void Dispose(bool disposing)
        {
            if (mWriter != null)
            {
#if DEBUG
                if (!disposing) Debug.Fail("Failed to dispose CsvWriter");
#endif
                mWriter.Dispose();
                mWriter = null;
            }
            if (disposing)
            {
                GC.SuppressFinalize(this);
            }
        }
    }

}
