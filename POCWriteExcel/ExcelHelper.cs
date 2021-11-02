using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using OfficeOpenXml;

namespace POCWriteExcel
{
    public class ExcelHelper
    {
        public string PathFile { get; }
        public ExcelOptions Options { get; }

        public DateFormatProvider DefaultDateFormatProvider => new DateFormatProvider
        {
            Format = "yyyy-MM-dd",
            CultureInfo = new CultureInfo("en-US")
        };

        public ExcelHelper(string pathFile) : this(pathFile, new ExcelOptions())
        { }

        public ExcelHelper(string pathFile, ExcelOptions options)
        {
            PathFile = pathFile ?? throw new ArgumentNullException(nameof(pathFile));

            Options = options;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        public FileInfo WriteFile<T>(string sheetName, IEnumerable<T> data) where T : class
        {
            var fileInfo = new FileInfo(PathFile);
            using (var package = new ExcelPackage(fileInfo))
            {
                var sheet = package.Workbook.Worksheets.Add(sheetName ?? "sheet1");
                var dataTable = data.ToDataTable();

                var headerColumnNames = Options.IsWriteHeader && Options.CustomColumnNames.Any()
                                        ? Options.CustomColumnNames.ToArray()
                                        : dataTable.Columns.Cast<DataColumn>().Select(s => s.ColumnName).ToArray();

                //Passing header name into cell
                PassingHeaders(sheet, headerColumnNames);

                //Passing data into cell
                PassingData(sheet, data.ToDataTable());
                package.Save();
            }
            return fileInfo;
        }

        private void PassingData(ExcelWorksheet sheet, DataTable data)
        {
            int startRow = Options.IsWriteHeader ? 2 : 1;
            for (int row = 0; row < data.Rows.Count; row++)
            {
                for (int col = 0; col < data.Columns.Count; col++)
                {
                    var column = data.Columns[col];
                    var value = data.Rows[row][col];

                    if (column.DataType == typeof(DateTime) || column.DataType == typeof(DateTime?))
                    {
                        value = ConvertToDateTime(value);
                    }

                    sheet.Cells[startRow, col + 1].Value = value;
                }
                startRow++;
            }
        }

        private void PassingHeaders(ExcelWorksheet sheet, string[] columnNames)
        {
            if (!Options.IsWriteHeader)
            {
                return;
            }

            int col = 1;
            foreach (var name in columnNames)
            {
                sheet.Cells[1, col].Value = name;
                col++;
            }
        }

        private object ConvertToDateTime(object value)
        {
            if (value == null)
            {
                return value;
            }

            if (Options.DateFormatProvider.IsEmpty)//If empty use default.
            {
                value = Convert.ToDateTime(value, DefaultDateFormatProvider.CultureInfo)
                               .ToString(DefaultDateFormatProvider.Format);
            }
            else if (!Options.DateFormatProvider.IsEmpty)
            {
                value = Convert.ToDateTime(value, Options.DateFormatProvider.CultureInfo)
                               .ToString(Options.DateFormatProvider.Format);
            }
            else if (!string.IsNullOrWhiteSpace(Options.DateFormatProvider.Format))
            {
                value = Convert.ToDateTime(value).ToString(Options.DateFormatProvider.Format);
            }
            else if (Options.DateFormatProvider.CultureInfo != null)
            {
                value = Convert.ToDateTime(value, Options.DateFormatProvider.CultureInfo);
            }

            return value;
        }

        public class ExcelOptions
        {
            /// <summary>
            /// Write header status. If true is write header, false no write.
            /// </summary>
            public bool IsWriteHeader { get; set; }
            /// <summary>
            /// Column name for write into excel file for custom.
            /// But must define <see cref="IsWriteHeader"/> is true.
            /// </summary>
            public List<string> CustomColumnNames { get; set; }

            /// <summary>
            /// Provider for set format datetime data for write into excel file.
            /// </summary>
            public DateFormatProvider DateFormatProvider { get; set; }

            public ExcelOptions()
            {
                CustomColumnNames = new List<string>();
                DateFormatProvider = new DateFormatProvider();
            }
        }

        public class DateFormatProvider
        {
            public string Format { get; set; }
            public CultureInfo CultureInfo { get; set; }

            internal bool IsEmpty => string.IsNullOrWhiteSpace(Format) && CultureInfo == null;
        }
    }

    public static class HelperExtension
    {
        public static DataTable ToDataTable<T>(this IEnumerable<T> data, string tableName = "") where T : class
        {
            if (data == null)
            {
                throw new ArgumentNullException(nameof(data));
            }

            var properties = typeof(T).GetProperties();
            var results = new DataTable(tableName ?? typeof(T).Name);

            //Create columns.
            results.Columns.AddRange(properties.Select(s => new DataColumn(s.Name, s.PropertyType)).ToArray());

            foreach (var row in data)
            {
                var newRow = results.NewRow();
                foreach (var property in properties)
                {
                    newRow[property.Name] = property.GetValue(row);
                }
                results.Rows.Add(newRow);
            }

            return results;
        }
    }
}
