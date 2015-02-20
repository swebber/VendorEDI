using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Text;

namespace VendorEDI
{
    class DbUtilities
    {
        private static readonly NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();

        private const string DB_FILENAME = "VendorEDI.sqlite";
        private static readonly string DATA_LEAD = AppSettings.Get<string>("data.lead");

        private string dbPath;
        private string dbFile;
        private string connStr;

        public bool Initialize(string fileName)
        {
            try
            {
                dbPath = Path.GetDirectoryName(fileName);
                dbFile = string.Format("{0}\\{1}", dbPath, DB_FILENAME);
                connStr = string.Format(@"Data Source={0};Version=3;", dbFile);

                SQLiteConnection.CreateFile(dbFile);

                using (var conn = new SQLiteConnection(connStr))
                {
                    conn.Open();

                    string sql = "CREATE TABLE IF NOT EXISTS \"AccountsPayable\" (\"VendorName\" VARCHAR(64) NOT NULL , \"VendorNumber\" VARCHAR(32) NOT NULL , \"CheckNumber\" VARCHAR(16) NOT NULL , \"BatchNumber\" VARCHAR(16) NOT NULL , \"OrderNumber\" VARCHAR(32) NOT NULL , \"VendorOrderInvoiceNumber\" VARCHAR(32) NOT NULL , \"SknId\" VARCHAR(16) NOT NULL , \"VendorSkuCode\" VARCHAR(32) NOT NULL , \"UnitsShipped\" INTEGER NOT NULL DEFAULT 0, \"VendorItemCost\" INTEGER NOT NULL DEFAULT 0, \"VendorShippingCost\" INTEGER NOT NULL DEFAULT 0, \"VendorHandlingCost\" INTEGER NOT NULL DEFAULT 0, \"ItemCost\" INTEGER NOT NULL DEFAULT 0, \"TotalCost\" INTEGER NOT NULL DEFAULT 0, \"VendorTotalCost\" INTEGER NOT NULL DEFAULT 0)";
                    using (var cmd = new SQLiteCommand(sql, conn))
                    {
                        cmd.ExecuteNonQuery();
                    }

                    sql = "DELETE FROM \"main\".\"AccountsPayable\"";
                    using (var cmd = new SQLiteCommand(sql, conn))
                    {
                        cmd.ExecuteNonQuery();
                    }

                    using (var stream = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    using (var reader = new StreamReader(stream))
                    {
                        string line;
                        while ((line = reader.ReadLine()) != null)
                        {
                            if (line.StartsWith(DATA_LEAD, StringComparison.InvariantCultureIgnoreCase))
                            {
                                string[] payableItems = line.Split('\t');
                            }
                        }
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message, ex);
                return false;
            }
        }
    }
}
