using CsvHelper;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace VendorEDI
{
    class FileUtilities
    {
        private static readonly string HEADER_LEAD = AppSettings.Get<string>("header.lead");
        private static readonly string DATA_LEAD = AppSettings.Get<string>("data.lead");

        public static string CleanFile(string fileName)
        {
            string cleanFile = string.Format(@"{0}\Clean-{1}.csv", Path.GetDirectoryName(fileName), Path.GetFileNameWithoutExtension(fileName));

            using (var writer = new StreamWriter(cleanFile))
            using (var stream = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var reader = new StreamReader(stream))
            {
                var csv = new CsvWriter(writer);

                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    if (line.StartsWith(HEADER_LEAD, StringComparison.InvariantCultureIgnoreCase))
                    {
                        WriteCsvRecord(csv, line);
                    }
                    else if (line.StartsWith(DATA_LEAD, StringComparison.InvariantCultureIgnoreCase))
                    {
                        WriteCsvRecord(csv, line);
                    }
                }
            }

            return cleanFile;
        }

        private static void WriteCsvRecord(CsvWriter csv, string line)
        {
            string[] items = line.Split('\t');
            foreach (var item in items)
            {
                csv.WriteField(item);
            }
            csv.NextRecord();
        }
    }
}
