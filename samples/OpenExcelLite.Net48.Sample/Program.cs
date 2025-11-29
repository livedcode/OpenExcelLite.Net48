using System;


namespace OpenExcelLite.Net48.Sample
{
    internal class Program
    {
        static void Main()
        {
            Console.WriteLine("Generating Excel example files...");

            // ---------------------------------------------------
            // Call all demos here (same grouping as OpenExcelLite)
            // ---------------------------------------------------
            GenerateInMemory();
            GenerateInMemoryWithEmptyRows();
            GenerateInMemoryWithAfterHeaderEmptyRows();
            GenerateInMemoryHyperlinks();
            GenerateInMemoryHyperlinksWithEmptyRows();
            GenerateInMemoryMultiSheet();
            GenerateInMemoryMultiSheetHyperlinks();
            GenerateInMemoryMultiSheetWithEmptyRows();
            GenerateInMemoryTenSheets();

            Console.WriteLine("Done.");
        }

        // ============================================================
        // 1) In-Memory Excel Demo
        // ============================================================
        static void GenerateInMemory()
        {
            var wb = new ExcelWorkbook();
            var s = wb.AddSheet("Employees");

            s.AddRow("Id", "Name", "JoinDate", "Salary", "Active");
            s.AddRow(1, "Alex", DateTime.Today, 5000.5m, true);
            s.AddRow(2, "Brian", DateTime.Today.AddDays(-3), 6500.75m, true);
            s.AddRow(3, "Cindy", DateTime.Today.AddDays(-10), 7200m, false);

            s.AutoFitColumns();

            wb.SaveToFile("InMemory.xlsx");
        }

        // ============================================================
        // 2) In-Memory Empty Rows (Before Header)
        // ============================================================
        static void GenerateInMemoryWithEmptyRows()
        {
            var wb = new ExcelWorkbook();
            var s = wb.AddSheet("Employees");

            s.AddEmptyRows(2);
            s.AddRow("Id", "Name", "JoinDate", "Salary", "Active");
            s.AddRow(1, "Alex", DateTime.Today, 5000.5m, true);
            s.AddRow(2, "Brian", DateTime.Today.AddDays(-3), 6500.75m, true);
            s.AddRow(3, "Cindy", DateTime.Today.AddDays(-10), 7200m, false);

            s.AutoFitColumns();

            wb.SaveToFile("InMemoryEmptyRows.xlsx");
        }

        // ============================================================
        // 3) In-Memory Empty Rows (After Header)
        // ============================================================
        static void GenerateInMemoryWithAfterHeaderEmptyRows()
        {
            var wb = new ExcelWorkbook();
            var s = wb.AddSheet("Employees");

            s.AddRow("Id", "Name", "JoinDate", "Salary", "Active");
           
            s.AddRow(1, "Alex", DateTime.Today, 5000.5m, true);
            s.AddEmptyRows(2);
            s.AddRow(2, "Brian", DateTime.Today.AddDays(-3), 6500.75m, true);
            s.AddEmptyRows(2);
            s.AddRow(3, "Cindy", DateTime.Today.AddDays(-10), 7200m, false);

            s.AutoFitColumns();

            wb.SaveToFile("InMemoryEmptyRowsAF.xlsx");
        }

        // ============================================================
        // 4) In-Memory Hyperlinks
        // ============================================================
        static void GenerateInMemoryHyperlinks()
        {
            var wb = new ExcelWorkbook();
            var s = wb.AddSheet("Links");

            s.AddRow("Name", "Website");
            s.AddRow("Google", new HyperlinkCell("Visit Google", "https://google.com"));
            s.AddRow("Repo", new HyperlinkCell("GitHub", "https://github.com/livedcode/OpenExcelLite"));

            wb.SaveToFile("InMemoryHyperlinks.xlsx");
        }

        // ============================================================
        // 5) In-Memory Hyperlinks + Empty Rows
        // ============================================================
        static void GenerateInMemoryHyperlinksWithEmptyRows()
        {
            var wb = new ExcelWorkbook();
            var s = wb.AddSheet("Links");

            s.AddEmptyRows(2);
            s.AddRow("Name", "Website");
            s.AddRow("Google", new HyperlinkCell("Visit Google", "https://google.com"));
            s.AddRow("Repo", new HyperlinkCell("GitHub", "https://github.com/livedcode/OpenExcelLite"));

            wb.SaveToFile("InMemoryHyperlinksEmptyRows.xlsx");
        }

        // ============================================================
        // 6) In-Memory Multi-Sheet
        // ============================================================
        static void GenerateInMemoryMultiSheet()
        {
            var wb = new ExcelWorkbook();

            var s1 = wb.AddSheet("Employees");
            s1.AddRow("Id", "Name");
            s1.AddRow(1, "Alex");
            s1.AddRow(2, "Brian");

            var s2 = wb.AddSheet("Departments");
            s2.AddRow("DeptId", "Department");
            s2.AddRow(10, "Finance");
            s2.AddRow(20, "IT");

            var s3 = wb.AddSheet("Summary");
            s3.AddRow("Generated", DateTime.Now);

            wb.SaveToFile("InMemoryMultiSheet.xlsx");
        }

        // ============================================================
        // 7) In-Memory Multi-Sheet Hyperlinks
        // ============================================================
        static void GenerateInMemoryMultiSheetHyperlinks()
        {
            var wb = new ExcelWorkbook();

            wb.AddSheet("Links1")
              .AddRow("Name", "Website")
              .AddRow("Google", new HyperlinkCell("Visit Google", "https://google.com"));

            wb.AddSheet("Links2")
              .AddRow("API", "URL")
              .AddRow("Users", new HyperlinkCell("User API", "https://yourapi.com/users"));

            wb.AddSheet("Links3")
              .AddRow("Doc", "URL")
              .AddRow("README", new HyperlinkCell("README", "https://github.com/livedcode/OpenExcelLite/blob/main/README.md"));

            wb.SaveToFile("InMemoryMultiSheetHyperlinks.xlsx");
        }

        // ============================================================
        // 8) In-Memory Multi-Sheet with Empty Rows
        // ============================================================
        static void GenerateInMemoryMultiSheetWithEmptyRows()
        {
            var wb = new ExcelWorkbook();

            wb.AddSheet("A")
              .AddEmptyRows(3)
              .AddRow("Id", "Value")
              .AddRow(1, "AAA");

            wb.AddSheet("B")
              .AddRow("Key", "Result")
              .AddEmptyRows(2)
              .AddRow("X", 111);

            wb.AddSheet("C")
              .AddEmptyRows(5)
              .AddRow("Title", "Data")
              .AddRow("Demo", 999);

            wb.SaveToFile("InMemoryMultiSheetEmptyRows.xlsx");
        }

        // ============================================================
        // 9) In-Memory 10 Sheets
        // ============================================================
        static void GenerateInMemoryTenSheets()
        {
            var wb = new ExcelWorkbook();

            for (int i = 1; i <= 10; i++)
            {
                var s = wb.AddSheet("Sheet_" + i);
                s.AddRow("Row", "Value");
                for (int r = 1; r <= 5; r++)
                    s.AddRow(r, "Data " + r + " in Sheet " + i);
            }

            wb.SaveToFile("InMemoryTenSheets.xlsx");
        }
    }
}
