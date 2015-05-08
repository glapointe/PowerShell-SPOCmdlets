using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Management.Automation;
using System.Net;
using System.Text;

namespace Lapointe.SharePointOnline.PowerShell.Utilities
{
    internal static class StreamHelper
    {
        private const int activityId = 0xa681412;
        internal const int ChunkSize = 0x2710;
        internal const int DefaultReadBuffer = 0x186a0;

        internal static Encoding GetEncoding(string characterSet)
        {
            string name = string.IsNullOrEmpty(characterSet) ? "ISO-8859-1" : characterSet;
            return Encoding.GetEncoding(name);
        }
        internal static Encoding GetDefaultEncoding()
        {
            return GetEncoding((string)null);
        }

        internal static string DecodeStream(MemoryStream stream, string characterSet)
        {
            Encoding encoding = GetEncoding(characterSet);
            return DecodeStream(stream, encoding);
        }

        internal static string DecodeStream(MemoryStream stream, Encoding encoding)
        {
            if (encoding == null)
            {
                encoding = GetDefaultEncoding();
            }
            byte[] bytes = stream.ToArray();
            return encoding.GetString(bytes);
        }

        internal static byte[] EncodeToBytes(string str)
        {
            return EncodeToBytes(str, null);
        }

        internal static byte[] EncodeToBytes(string str, Encoding encoding)
        {
            if (encoding == null)
            {
                encoding = GetDefaultEncoding();
            }
            return encoding.GetBytes(str);
        }

        internal static Stream GetResponseStream(WebResponse response)
        {
            Stream responseStream = response.GetResponseStream();
            string str = response.Headers["Content-Encoding"];
            if (str != null)
            {
                if (str.IndexOf("gzip", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return new GZipStream(responseStream, CompressionMode.Decompress);
                }
                if (str.IndexOf("deflate", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    responseStream = new DeflateStream(responseStream, CompressionMode.Decompress);
                }
            }
            return responseStream;
        }

        internal static MemoryStream ReadStream(Stream stream, long contentLength, PSCmdlet cmdlet)
        {
            MemoryStream stream3;
            if (stream == null)
            {
                throw new ArgumentNullException("stream");
            }
            if (!stream.CanRead)
            {
                throw new ArgumentOutOfRangeException("stream");
            }
            if (0L >= contentLength)
            {
                contentLength = 0x186a0L;
            }
            int capacity = (int)Math.Min(contentLength, 0x7fffffffL);
            MemoryStream stream2 = new MemoryStream(capacity);
            try
            {
                long o = 0L;
                byte[] buffer = new byte[0x2710];
                int count = 1;
                while (0 < count)
                {
                    if (cmdlet != null)
                    {
                        ProgressRecord progressRecord = new ProgressRecord(0xa681412, "Reading Web Response", string.Format("Reading response stream... (Number of bytes read: {0})", o));
                        cmdlet.WriteProgress(progressRecord);
                    }
                    count = stream.Read(buffer, 0, buffer.Length);
                    if (0 < count)
                    {
                        stream2.Write(buffer, 0, count);
                    }
                    o += count;
                }
                if (cmdlet != null)
                {
                    ProgressRecord record2 = new ProgressRecord(0xa681412, "Reading Web Response", string.Format("Reading web response completed. (Number of bytes read: {0})", o))
                    {
                        RecordType = ProgressRecordType.Completed
                    };
                    cmdlet.WriteProgress(record2);
                }
                stream2.SetLength(o);
                stream3 = stream2;
            }
            catch (Exception)
            {
                stream2.Close();
                throw;
            }
            return stream3;
        }



        internal static void SaveStreamToFile(Stream stream, string filePath, PSCmdlet cmdlet)
        {
            using (FileStream stream2 = File.Create(filePath))
            {
                long position = stream.Position;
                stream.Position = 0L;
                WriteToStream(stream, stream2, cmdlet);
                stream.Position = position;
            }
        }

        internal static void WriteToStream(byte[] input, Stream output)
        {
            output.Write(input, 0, input.Length);
            output.Flush();
        }

        internal static void WriteToStream(Stream input, Stream output, PSCmdlet cmdlet)
        {
            byte[] buffer = new byte[0x2710];
            long length = input.Length;
            int count = 1;
            while (0 < count)
            {
                if (cmdlet != null)
                {
                    ProgressRecord progressRecord = new ProgressRecord(0xa681412, "Writing web request.", string.Format("Writing request stream... (Number of bytes remaining: {0})", length));
                    cmdlet.WriteProgress(progressRecord);
                }
                int num3 = (int) Math.Min(length, 0x2710L);
                count = input.Read(buffer, 0, num3);
                if (0 < count)
                {
                    output.Write(buffer, 0, count);
                }
                length -= count;
            }
            if (cmdlet != null)
            {
                ProgressRecord record2 = new ProgressRecord(0xa681412, "Writing web request.", string.Format("Writing web request completed. (Number of bytes remaining: {0})", length))
                {
                    RecordType = ProgressRecordType.Completed
                };
                cmdlet.WriteProgress(record2);
            }
            output.Flush();
        }
    }
    
}
