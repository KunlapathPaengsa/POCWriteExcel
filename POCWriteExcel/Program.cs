using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using OfficeOpenXml;

namespace POCWriteExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            string path = @"C:\Users\Lenovo\Desktop\MyWorkbook.xlsx";
            var fileInfo = new FileInfo(path);
            if (fileInfo.Exists)
            {
                fileInfo.Delete();
            }
            var memberships = GetMemberships();
            var excelHelper = new ExcelHelper(path, new ExcelHelper.ExcelOptions
            {
                IsWriteHeader = true
                , CustomColumnNames = new List<string> { "", "", "" } 
                ,
                DateFormatProvider = new ExcelHelper.DateFormatProvider
                {
                    Format = "yyyy-MM-dd HH:mm:ss.ffff",
                    CultureInfo = null//new CultureInfo("en-US")
                }
            });
            excelHelper.WriteFile(nameof(Membership), memberships);
        }

        static void WriteExcelFile(FileInfo fileInfo)
        {
            using (var package = new ExcelPackage(fileInfo))
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");
                PassingMemberData(sheet);
                package.Save();
            }
        }

        static void PassingMemberData(ExcelWorksheet sheet)
        {
            int limit = 10;
            //int limit = 1000000;
            for (int i = 0; i < limit; i++)
            {
                var birthday = DateTime.Now.AddDays(-10);
                var row = i + 1;
                sheet.Cells[row, 1].Value = row;
                sheet.Cells[row, 2].Value = Guid.NewGuid();
                sheet.Cells[row, 3].Value = $"Member_{i}";

                sheet.Cells[row, 4].Value = birthday.ToString("yyyy-MM-dd HH:mm:ss.ffff", new CultureInfo("en-US"));

                sheet.Cells[row, 5].Value = DateTime.Now.Year - birthday.Year;
                sheet.Cells[row, 6].Value = i % 2 == 0;
                sheet.Cells[row, 7].Value = i * 10000m;
                sheet.Cells[row, 8].Value = DateTime.Now;

                //sheet.Cells["A1:Z1"].Value = "1";

                //if (row == 1)
                //{
                //    for (int col = 9; col <= 500; col++)
                //    {
                //        sheet.Cells[row, col].Value = DateTime.Now;
                //    }
                //}

                Console.WriteLine($"Write excel file row : {row}.");
            }
        }

        static List<Membership> GetMemberships()
        {
            var memberships = new List<Membership>();
            int limit = 100;
            for (int i = 0; i < limit; i++)
            {
                var birthday = DateTime.Now.AddYears(-5);
                memberships.Add(new Membership
                {
                    Id = Guid.NewGuid(),
                    Name = $"Member_{i}",
                    Age = DateTime.Now.Year - birthday.Year,
                    Birthday = birthday,
                    IsActive = i % 3 == 0,
                    Amount = i * 10000,
                    CreatedDate = DateTime.Now
                });
            }
            return memberships;
        }
    }

    public class Membership
    {
        public Guid Id { get; set; }
        public string Name { get; set; }
        public int Age { get; set; }
        public DateTime Birthday { get; set; }
        public bool IsActive { get; set; }
        public decimal Amount { get; set; }
        public DateTime CreatedDate { get; set; }
    }
}
