using System;
using System.Collections.Generic;
using System.Data.Entity.Core.EntityClient;
using System.Data.SQLite;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace VendorEDI
{
    class DbUtilities
    {
        private static readonly NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();

        private const int INSERT_COUNT_THRESHOLD = 100;
        private const string DB_FILENAME = "VendorEDI.sqlite";
        private static readonly string DATA_LEAD = AppSettings.Get<string>("data.lead");

        private string dbPath;
        private string dbFile;
        private string connStr;
        private string entityConnStr;

        private string GetEntityConnectionString()
        {
            string originalConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings["VendorEdiEntities"].ConnectionString;
            var ecsBuilder = new EntityConnectionStringBuilder(originalConnectionString);
            var sqlCsBuilder = new SQLiteConnectionStringBuilder(ecsBuilder.ProviderConnectionString)
            {
                DataSource = dbFile
            };
            var providerConnectionString = sqlCsBuilder.ToString();
            ecsBuilder.ProviderConnectionString = providerConnectionString;

            return ecsBuilder.ToString();
        }

        private int ToInteger(string value)
        {
            int result = 0;
            int.TryParse(value, out result);
            return result;
        }

        private int ToMoney(string value)
        {
            int result = 0;

            decimal amount = 0m;
            if (decimal.TryParse(value, out amount))
            {
                amount *= 100m;
                result = (int)amount;
            }

            return result;
        }

        public bool Initialize(string fileName)
        {
            try
            {
                dbPath = Path.GetDirectoryName(fileName);
                dbFile = string.Format("{0}\\{1}", dbPath, DB_FILENAME);
                connStr = string.Format(@"Data Source={0};Version=3;", dbFile);
                entityConnStr = GetEntityConnectionString();

                SQLiteConnection.CreateFile(dbFile);

                using (var conn = new SQLiteConnection(connStr))
                {
                    conn.Open();

                    // string sql = "CREATE TABLE IF NOT EXISTS \"AccountsPayable\" (\"VendorName\" VARCHAR(64) NOT NULL , \"VendorNumber\" VARCHAR(32) NOT NULL , \"CheckNumber\" VARCHAR(16) NOT NULL , \"BatchNumber\" VARCHAR(16) NOT NULL , \"OrderNumber\" VARCHAR(32) NOT NULL , \"VendorOrderInvoiceNumber\" VARCHAR(32) NOT NULL , \"SknId\" VARCHAR(16) NOT NULL , \"VendorSkuCode\" VARCHAR(32) NOT NULL , \"UnitsShipped\" INTEGER NOT NULL DEFAULT 0, \"VendorItemCost\" INTEGER NOT NULL DEFAULT 0, \"VendorShippingCost\" INTEGER NOT NULL DEFAULT 0, \"VendorHandlingCost\" INTEGER NOT NULL DEFAULT 0, \"ItemCost\" INTEGER NOT NULL DEFAULT 0, \"TotalCost\" INTEGER NOT NULL DEFAULT 0, \"VendorTotalCost\" INTEGER NOT NULL DEFAULT 0)";
                    // string sql = "CREATE TABLE \"main\".\"AccountsPayable\" (\"Id\" INTEGER PRIMARY KEY NOT NULL, \"VendorName\" VARCHAR(64) NOT NULL , \"VendorNumber\" VARCHAR(32) NOT NULL , \"CheckNumber\" VARCHAR(16) NOT NULL , \"BatchNumber\" VARCHAR(16) NOT NULL , \"OrderNumber\" VARCHAR(32) NOT NULL , \"VendorOrderInvoiceNumber\" VARCHAR(32) NOT NULL , \"SknId\" VARCHAR(16) NOT NULL , \"VendorSkuCode\" VARCHAR(32) NOT NULL , \"UnitsShipped\" INTEGER NOT NULL DEFAULT 0, \"VendorItemCost\" INTEGER NOT NULL DEFAULT 0, \"VendorShippingCost\" INTEGER NOT NULL DEFAULT 0, \"VendorHandlingCost\" INTEGER NOT NULL DEFAULT 0, \"ItemCost\" INTEGER NOT NULL DEFAULT 0, \"TotalCost\" INTEGER NOT NULL DEFAULT 0, \"VendorTotalCost\" INTEGER NOT NULL DEFAULT 0)";
                    // string sql = "CREATE TABLE \"main\".\"AccountsPayable\" (\"Id\" INTEGER PRIMARY KEY AUTOINCREMENT, \"VendorName\" VARCHAR(64) NOT NULL , \"VendorNumber\" VARCHAR(32) NOT NULL , \"CheckNumber\" VARCHAR(16) NOT NULL , \"BatchNumber\" VARCHAR(16) NOT NULL , \"OrderNumber\" VARCHAR(32) NOT NULL , \"VendorOrderInvoiceNumber\" VARCHAR(32) NOT NULL , \"SknId\" VARCHAR(16) NOT NULL , \"VendorSkuCode\" VARCHAR(32) NOT NULL , \"UnitsShipped\" INTEGER NOT NULL DEFAULT 0, \"VendorItemCost\" INTEGER NOT NULL DEFAULT 0, \"VendorShippingCost\" INTEGER NOT NULL DEFAULT 0, \"VendorHandlingCost\" INTEGER NOT NULL DEFAULT 0, \"ItemCost\" INTEGER NOT NULL DEFAULT 0, \"TotalCost\" INTEGER NOT NULL DEFAULT 0, \"VendorTotalCost\" INTEGER NOT NULL DEFAULT 0)";
                    string sql = "CREATE TABLE \"main\".\"AccountsPayable\" (\"Id\" INTEGER PRIMARY KEY AUTOINCREMENT, \"VendorName\" VARCHAR(64) NOT NULL , \"VendorNumber\" VARCHAR(32) NOT NULL , \"CheckNumber\" VARCHAR(16) NOT NULL , \"BatchNumber\" VARCHAR(16) NOT NULL , \"OrderNumber\" VARCHAR(32) NOT NULL , \"VendorOrderInvoiceNumber\" VARCHAR(32) NOT NULL , \"SknId\" VARCHAR(16) NOT NULL , \"VendorSkuCode\" VARCHAR(32) NOT NULL , \"UnitsShipped\" INTEGER NOT NULL DEFAULT 0, \"VendorItemCost\" INTEGER NOT NULL DEFAULT 0, \"VendorShippingCost\" INTEGER NOT NULL DEFAULT 0, \"VendorHandlingCost\" INTEGER NOT NULL DEFAULT 0, \"ItemCost\" INTEGER NOT NULL DEFAULT 0, \"TotalCost\" INTEGER NOT NULL DEFAULT 0, \"VendorTotalCost\" INTEGER NOT NULL DEFAULT 0, \"IsProcessed\" BOOLEAN NOT NULL DEFAULT 0)";
                    using (var cmd = new SQLiteCommand(sql, conn))
                    {
                        cmd.ExecuteNonQuery();
                    }

                    sql = "DELETE FROM \"main\".\"AccountsPayable\"";
                    using (var cmd = new SQLiteCommand(sql, conn))
                    {
                        cmd.ExecuteNonQuery();
                    }

                    sql = "DELETE FROM \"main\".\"sqlite_sequence\" WHERE name=\"AccountsPayable\"";
                    using (var cmd = new SQLiteCommand(sql, conn))
                    {
                        cmd.ExecuteNonQuery();
                    }
                }

                using (var stream = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var reader = new StreamReader(stream))
                using (var db = new VendorEdiEntities(entityConnStr))
                {
                    db.Configuration.AutoDetectChangesEnabled = false;
                    db.Configuration.ValidateOnSaveEnabled = false;

                    logger.Debug("Start data load.");
                    var sw = new Stopwatch();
                    sw.Start();

                    string line;
                    int insertCount = 0;
                    while ((line = reader.ReadLine()) != null)
                    {
                        if (line.StartsWith(DATA_LEAD, StringComparison.InvariantCultureIgnoreCase))
                        {
                            string[] payableItem = line.Split('\t');

                            var ap = new AccountsPayable
                            {
                                VendorName = payableItem[0],
                                VendorNumber = payableItem[1],
                                CheckNumber = payableItem[2],
                                BatchNumber = payableItem[3],
                                OrderNumber = payableItem[4],
                                VendorOrderInvoiceNumber = payableItem[5],
                                SknId = payableItem[6],
                                VendorSkuCode = payableItem[7],
                                UnitsShipped = ToInteger(payableItem[8]),
                                VendorItemCost = ToMoney(payableItem[9]),
                                VendorShippingCost = ToMoney(payableItem[10]),
                                VendorHandlingCost = ToMoney(payableItem[11]),
                                ItemCost = ToMoney(payableItem[12]),
                                TotalCost = ToMoney(payableItem[13]),
                                VendorTotalCost = ToMoney(payableItem[14]),
                                IsProcessed = false
                            };

                            db.AccountsPayable.Add(ap);

                            if (++insertCount >= INSERT_COUNT_THRESHOLD)
                            {
                                insertCount = 0;
                                db.SaveChanges();
                            }
                        }
                    }

                    if (insertCount >= INSERT_COUNT_THRESHOLD)
                    {
                        db.SaveChanges();
                    }

                    sw.Stop();
                    logger.Debug("Elapsed Load Time: {0}", sw.Elapsed);
                }

                return true;
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message, ex);
                return false;
            }
        }

        public void ReportMissing()
        {

        }
    }
}
