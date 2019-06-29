using System;
using ClosedXML;
using ClosedXML.Excel;

namespace supa_cataloger
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");

            var workbook = new XLWorkbook("data/TestData.xlsx");
            var sheet = workbook.Worksheets.Worksheet(1);
            sheet.Name = "Catalog";

            InitializeSheet(workbook, "Yincludes", "VERIFY", "HIERARCHY", "Date Added", "Track Item", "Retailer", "Retailer Item ID", "TLD", "UPC", 
                "Title", "Manufacturer", "Brand", "Client Product Group", "Category", "Subcategory", "Amazon Sub Category", "Vat Rate Id", 
                "CTN", "Similar ASIN Grouping");
            InitializeSheet(workbook, "Manufacturers and Brands", "Brands", "Manufacturers");
            InitializeSheet(workbook, "Categories", "Valid Hierarchy", "Client Product Group", "Category", "Subcategory");
            InitializeSheet(workbook, "Zexcludes", "MATCHES FOUND", "Date Added", "Track Item", "Retailer", "Retailer Item ID", "TLD", "UPC", 
                "Title", "Manufacturer", "Brand", "Client Product Group", "Category", "Subcategory", "Amazon Sub Category", "Vat Rate Id", 
                "CTN", "Similar ASIN Grouping");



            var rangeUsed = sheet.RangeUsed();
            var rowsUsed = rangeUsed.Rows();
            var numColumns = rangeUsed.ColumnCount();
            var headerRow = rowsUsed.Cells();
            //Console.WriteLine(headerRow.);


            /*foreach (var row in rowsUsed)
            {
                Console.WriteLine(row.);
            }*/

            //sheet.Range(headerRow);

            workbook.SaveAs("Updated.xlsx");

            Console.WriteLine(headerRow);
        }

        static void InitializeSheet(XLWorkbook workbook, string sheetName, params string[] headers)
        {
            workbook.AddWorksheet(sheetName);
            var sheet = workbook.Worksheet(sheetName);
            for (int i = 0; i < headers.Length; i++)
            {
                // Excel has 1 index not 0
                sheet.Cell(1, i + 1).Value = headers[i];
            }
        }

    }
}
