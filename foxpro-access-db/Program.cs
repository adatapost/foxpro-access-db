using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;

namespace foxpro_access_db
{
    class Program
    {
        static void Main(string[] args)
        {
            ReadAccessAndUpdateFoxpro();
        }
        static void ReadAccessAndUpdateFoxpro()
        {
            using (var accessCn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=c:\csnet\foxpro-access-db\foxpro-access-db\accessdb\acc-db.accdb;Persist Security Info=True"))
            {

                accessCn.Open();
                using (var accessCmd = accessCn.CreateCommand())
                {
                    accessCmd.CommandText = "select * from ledger";
                    using (var aDr = accessCmd.ExecuteReader())
                    {
                        using (var foxCn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\csnet\foxpro-access-db\foxpro-access-db\foxdb;Extended Properties=dBASE IV;User ID=Admin;Password=;"))
                        {
                            foxCn.Open();
                            using (var foxCmd = foxCn.CreateCommand())
                            {
                                foxCmd.CommandText = "Update ledger set ac_op=@a1 where ac_no=@a2";
                                foxCmd.Parameters.Add("@a1", OleDbType.Integer, 4);
                                foxCmd.Parameters.Add("@a2", OleDbType.Integer, 4);

                                while (aDr.Read())
                                {
                                    Console.WriteLine(aDr[0] + "  " + aDr[1] + "  " + aDr[2]);

                                    // Update foxpro table
                                    foxCmd.Parameters["@a1"].Value = aDr[2];
                                    foxCmd.Parameters["@a2"].Value = aDr[0];
                                    foxCmd.ExecuteNonQuery();
                                }
                            }
                        }

                    }
                }

            }

        }

        static void ReadAccessTable()
        {
            using (var accessCn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=c:\csnet\foxpro-access-db\foxpro-access-db\accessdb\acc-db.accdb;Persist Security Info=True"))
            {
                using (var foxCn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\csnet\foxpro-access-db\foxpro-access-db\foxdb;Extended Properties=dBASE IV;User ID=Admin;Password=;"))
                {
                    accessCn.Open();
                    using (var accessCmd = accessCn.CreateCommand())
                    {
                        accessCmd.CommandText = "select * from ledger";
                        using (var aDr = accessCmd.ExecuteReader())
                        {
                            while (aDr.Read())
                            {
                                Console.WriteLine(aDr[0] + "  " + aDr[1] + "  " + aDr[2]);
                            }
                        }
                    }
                }
            }

        }

        static void TestReadFoxProTable()
        {
            using (var accessCn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=c:\csnet\foxpro-access-db\acc-db.accdb;Extended Properties=dBASE IV;User ID=Admin;"))
            {
                using (var foxCn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\csnet\foxpro-access-db\foxpro-access-db\foxdb;Extended Properties=dBASE IV;User ID=Admin;Password=;"))
                {
                    foxCn.Open();
                    using (var foxCmd = foxCn.CreateCommand())
                    {
                        foxCmd.CommandText = "select * from ledger";
                        using (var foxDr = foxCmd.ExecuteReader())
                        {
                            while (foxDr.Read())
                            {
                                Console.WriteLine(foxDr[0] + "  " + foxDr[1] + "  " + foxDr[2]);
                            }
                        }
                    }
                }
            }
        }
    }
}
