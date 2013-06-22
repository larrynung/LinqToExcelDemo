using System.Linq;
using LinqToExcel;

namespace ConsoleApplication1
{
    class Program
    {
        static void Main(string[] args)
        {
            System.Console.WriteLine("Data.xls (Excel 2003)...");
            TestLinqToExcel("Data.xls");
            System.Console.WriteLine(new string('=', 50));
            System.Console.WriteLine("Data.xlsx (Excel 2007)...");
            TestLinqToExcel("Data.xlsx");
        }

        void ShowColumnNames(string excelFile, string sheetName)
        {
            var excel = new ExcelQueryFactory(excelFile);
            var columnNames = excel.GetColumnNames(sheetName);
            foreach (var item in columnNames)
            {
                System.Console.WriteLine(item.ToString());
            }
        }

        void ShowWorkSheetNames(string excelFile)
        {
            var excel = new ExcelQueryFactory(excelFile);
            var workSheetNames = excel.GetWorksheetNames();
            foreach (var item in workSheetNames)
            {
                System.Console.WriteLine(item.ToString());
            }
        }

        void ShowWorkSheetData(string excelFile, string sheetName)
        {
            var excel = new ExcelQueryFactory(excelFile);

            //自己可自行加要過濾的條件，這邊只是示範
            var linq = from item in excel.Worksheet<Blogger>(sheetName)
                       where item.Sex == SexType.Boy
                       select item;
            foreach (var item in linq)
            {
                System.Console.WriteLine(item.ToString());
            }
        }

        static void TestLinqToExcel(string excelFile)
        {
            const string FIRST_SHEET = "BlogData1";
            const string SECOND_SHEET = "BlogData2";
            const string THIRD_SHEET = "Sheet3";

            var excel = new ExcelQueryFactory(excelFile);
            System.Console.WriteLine("Excel File: {0}", excel.FileName);

            System.Console.WriteLine();
            System.Console.WriteLine("WorksheetNames...");
            var workSheetNames = excel.GetWorksheetNames();
            foreach (var item in workSheetNames)
            {
                System.Console.WriteLine(item.ToString());
            }

            System.Console.WriteLine();
            System.Console.WriteLine("BlogData1's Columns...");
            var columnNames = excel.GetColumnNames(FIRST_SHEET);
            foreach (var item in columnNames)
            {
                System.Console.WriteLine(item.ToString());
            }

            System.Console.WriteLine();
            System.Console.WriteLine("BlogData2's Columns...");
            columnNames = excel.GetColumnNames(SECOND_SHEET);
            foreach (var item in columnNames)
            {
                System.Console.WriteLine(item.ToString());
            }



            System.Console.WriteLine();
            System.Console.WriteLine("Sheet3's Columns...");
            columnNames = excel.GetColumnNames(THIRD_SHEET);
            foreach (var item in columnNames)
            {
                System.Console.WriteLine(item.ToString());
            }

            System.Console.WriteLine();
            System.Console.WriteLine("BlogData1 With ExcelQueryFactory.Worksheet...");
            excel.AddMapping<Blogger>(item => item.FirstName, "First Name");
            excel.AddMapping<Blogger>(item => item.LastName, "Last Name");
            excel.AddTransformation<Blogger>(item => item.Sex, item => (item == "Boy") ? SexType.Boy : SexType.Girl);

            foreach (var item in excel.Worksheet<Blogger>(FIRST_SHEET))
            {
                System.Console.WriteLine(item.ToString());
            }

            System.Console.WriteLine();
            System.Console.WriteLine("BlogData2 With ExcelQueryFactory.Worksheet...");

            foreach (var item in excel.Worksheet<Blogger>(SECOND_SHEET))
            {
                System.Console.WriteLine(item.ToString());
            }

            System.Console.WriteLine();
            System.Console.WriteLine("BlogData2 With ExcelQueryFactory.WorksheetRange...");

            foreach (var item in excel.WorksheetRange<Blogger>("B2", "G3", SECOND_SHEET))
            {
                System.Console.WriteLine(item.ToString());
            }
        }

    }
}
