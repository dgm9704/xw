namespace xw
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using OfficeOpenXml;
    using XbrlTable;

    class Program
    {
        static string rootFolder = Environment.GetEnvironmentVariable("XBRL_ROOT");

        static void Main(string[] args)
        {
            var modulePath = "/home/john/xbrl/eiopa.europa.eu/eu/xbrl/s2md/fws/solvency/solvency2/2016-07-15/mod/adh.xsd";
            var workbookPath = Path.ChangeExtension(Path.GetFileName(modulePath), "xlsx");
            CreateWorkbookForModule(modulePath, workbookPath);
        }

        private static void CreateWorkbookForModule(string modulePath, string workbookPath)
        {
            var tables = Parsing.ParseTables(modulePath);

            File.Delete(workbookPath);

            var file = new FileInfo(workbookPath);

            using (var package = new ExcelPackage(file))
            {
                foreach (var table in tables)
                {
                    var worksheet = package.Workbook.Worksheets.Add(table.Code);
                    var endCoordinate = table.WriteToWorksheet(worksheet, new ExcelCoordinate(1, 1));
                    worksheet.Pretty();
                }

                package.Save();
            }
            Console.WriteLine(file.FullName);
        }
    }
}
