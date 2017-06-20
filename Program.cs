using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using XbrlTable;

namespace xw
{
    class Program
    {
        static string rootFolder = Environment.GetEnvironmentVariable("XBRL_ROOT");

        static void Main(string[] args)
        {
            var modulePath = "/home/john/xbrl/eiopa.europa.eu/eu/xbrl/s2md/fws/solvency/solvency2/2016-07-15/mod/aes.xsd";
            var tables = Parsing.ParseTables(modulePath);
            var workbookPath = Path.ChangeExtension(Path.GetFileName(modulePath), "xlsx");
            File.Delete(workbookPath);

            FileInfo newFile = new FileInfo(workbookPath);

            using (ExcelPackage package = new ExcelPackage(newFile))
            {
                foreach (var table in tables)
                {
                    var worksheet = package.Workbook.Worksheets.Add(table.Code);
                    worksheet.AddTable(table, 1, 1);
                    worksheet.Pretty();
                }

                package.Save();
            }
            Console.WriteLine(newFile.FullName);
        }
    }
}
