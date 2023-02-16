using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Dapper;
using System.Configuration;
using System.Data.SqlClient;
using GemBox.Spreadsheet;
using static System.Runtime.CompilerServices.RuntimeHelpers;

namespace CloseMerchants
{
    static class Program
    {
        public static string filepath;
        public static int ordinalMid;

        static void Main(string[] args)
        {
            SpreadsheetInfo.SetLicense(ConfigurationManager.AppSettings["GemBoxKey"]);

            Console.WriteLine("Run against 1 - Staging or 2 - Prod (1 or 2):");
            string environment = Console.ReadLine();

            Console.WriteLine("Close Date:");
            string CloseDate = Console.ReadLine();

            Console.WriteLine("Close Reason:");
            string CloseReason = Console.ReadLine();

            Console.WriteLine("Enter Input File (optional):");
            filepath = Console.ReadLine();

            if (!string.IsNullOrEmpty(filepath))
            {
                // get csv import file details
                GetInputFileDetails(
                    out int ordinalMid);
            }

            Console.WriteLine("Enter Sortcode (optional):");
            string Sortcode = Console.ReadLine();

            Console.WriteLine("Enter user ID:");
            string UserID = Console.ReadLine();

            string dbString = string.Empty;
            if (environment == "2")
            {
                dbString = ConfigurationManager.AppSettings["MESstring2"];
            }
            else
            {
                dbString = ConfigurationManager.AppSettings["MESstring1"];
            }

            using (SqlConnection con = new SqlConnection(dbString))
            {
                if (!string.IsNullOrEmpty(filepath))
                {
                    ExcelFile ef = ExcelFile.Load(filepath);
                    foreach (ExcelWorksheet sheet in ef.Worksheets)
                    {
                        // Iterate through all rows in an Excel worksheet.
                        foreach (ExcelRow row in sheet.Rows)
                        {
                            if (row.Cells[ordinalMid].Value.ToString() != "MID")
                            {
                                string MID = row.Cells[ordinalMid].Value.ToString();
                                Console.WriteLine(MID);
                                con.Execute("sp_CloseMerchant", new { MID, CloseDate, CloseReason, UserID }, commandType: System.Data.CommandType.StoredProcedure);

                                List<string> flexFees = new List<string>();
                                string sql = "SELECT TranCode FROM tblFlexFee WHERE strMID = '" + MID + "'";
                                flexFees = con.Query<string>(sql, commandType: System.Data.CommandType.Text).ToList();

                                foreach(string TranCode in flexFees)
                                {
                                    con.Execute("sp_DumpFlexFee", new { MID, TranCode, UserID }, commandType: System.Data.CommandType.StoredProcedure);
                                }

                                con.Execute("sp_ZeroOutFees", new {MID, UserID}, commandType: System.Data.CommandType.StoredProcedure);
                            }
                        }
                    }
                }
                else
                {
                    string sql = "SELECT strMID FROM tblMrch WHERE SortCode LIKE '" + Sortcode + "%'";
                    var MIDList = con.Query<string>(sql, commandType: System.Data.CommandType.Text).ToList();

                    foreach (var MID in MIDList)
                    {
                        Console.WriteLine(MID);
                        con.Execute("sp_CloseMerchant", new { MID, CloseDate, UserID }, commandType: System.Data.CommandType.StoredProcedure);
                    }
                }
            }

            Console.WriteLine("Processing complete");
            Console.ReadLine();
        }

        public static Dictionary<int, string> GetOrdinals()
        {
            if (string.IsNullOrEmpty(filepath))
                throw new InvalidOperationException("File path has not been set!");

            Dictionary<int, string> ordinals = new Dictionary<int, string>();
            ExcelFile excelFile = ExcelFile.Load(filepath);
            ExcelWorksheet worksheet = excelFile.Worksheets[0];

            int colCount = worksheet.CalculateMaxUsedColumns();

            for (int i = 0; i < colCount; i++)
            {
                string header = worksheet.Rows[0].Cells[i].Value?.ToString() ?? string.Empty;

                ordinals.Add(i, header);
            }

            return ordinals;
        }

        public static void GetInputFileDetails(out int ordinalMid)
        {
            Console.WriteLine("Ordinal headers: ");
            var ordinals = GetOrdinals();
            foreach (var ordinal in ordinals)
            {
                Console.WriteLine($"{ordinal.Key} - {ordinal.Value}");
            }

            ordinalMid = 999;
            while (!ordinals.ContainsKey(ordinalMid))
            {
                ordinalMid = GetNumericInput("MID Ordinal:");
            }
        }

        private static int GetNumericInput(string promptText)
        {
            string input;
            int number;
            Console.WriteLine(promptText);
            input = Console.ReadLine();

            if (!Int32.TryParse(input, out number))
            {
                Console.WriteLine("Invalid input, please ensure input is numeric!");
                number = GetNumericInput(promptText);
            }

            return number;
        }
    }
}
