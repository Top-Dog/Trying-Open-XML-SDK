/* Test program to demonstrate the basic functionality of 
 * the Open XML SDK. This example program creates a new
 * Excel document called "Report.xlsx" with one sheet called 
 * "Test Sheet".
 * 
 * Based off the tutorial here:
 * http://www.dispatchertimer.com/tutorial/how-to-create-an-excel-file-in-net-using-openxml-part-2-export-a-collection-to-spreadsheet/
 * 
 * Author: Sean O'Connor
 * Date: 17 June 2016
 * 
 * Builds using MSBuild (Windows), or
 * xbuild (linux).
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

// Open XML SDK namespaces
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Testing_Open_XML_SDK
{
    public class Report
    {
        private uint NumberOfBooks = 0;

        private uint NumberOfSheets = 0;
        static List<SheetData> _sheets;
        
        private SpreadsheetDocument document;
        private WorkbookPart workbookPart;
        private WorksheetPart worksheetPart;
        private Sheets sheets;

        /* Run the program
         * @param filename: The name of the file to create
         */
        static void Main(string[] args)
        {
            // Create an instance of Report class, and call the CreateExcelDoc method
            Report report = new Report();
            report.CreateExcelDoc(@"H:\My Documents\Visual Studio 2013\Projects\Testing Open XML SDK\Report.xlsx", "Sean's Test sheet");
            report.AddEmployees();
            report.CloseDocument();

            Console.WriteLine("Excel file has been created!");

            Excel xl = new Excel();
            xl.Initialize();
            xl.AddWorkbook(@"H:\My Documents\Visual Studio 2013\Projects\Testing Open XML SDK\Report 2.xlsx");
            xl.AddWorksheet(@"H:\My Documents\Visual Studio 2013\Projects\Testing Open XML SDK\Report 2.xlsx", "Sheet QRTY");
            //xl.AddWorksheet(@"H:\My Documents\Visual Studio 2013\Projects\Testing Open XML SDK\Report 2.xlsx", "Sheet QRTY 2");
            //xl.AddWorksheet(@"H:\My Documents\Visual Studio 2013\Projects\Testing Open XML SDK\Report 2.xlsx", "Sheet QRTY 3");
            xl.CloseDocument(@"H:\My Documents\Visual Studio 2013\Projects\Testing Open XML SDK\Report 2.xlsx");

            Console.Read();
        }

        public void CreateExcelDoc(string fileName, string sheetName)
        {
            // Create a new spreadsheet document with a name spescified by the method fileName parameter
            document = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook);
            //using (SpreadsheetDocument document = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook)) {} // using is like python's with
            
            // Add a workbookPart to the document
            workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();
            sheets = workbookPart.Workbook.AppendChild(new Sheets());

            // Add a WorksheetPart to the workbookPart
            worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet();

            _sheets = new List<SheetData>();
            // Add the first sheet
            AddSheet(sheetName);

            // Save the new document
            workbookPart.Workbook.Save();
        }

        public void AddWorkBook(string workbookName)
        {
            NumberOfBooks += 1;
        }

        public void AddSheet(string sheetName)
        {
            NumberOfSheets += 1;
            // Create a sheet (associated with the WorksheetPart) and append it to the "sheets" abstract object.
            Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = NumberOfSheets, Name = sheetName };
            sheets.Append(sheet);
            // SheetData is a container for all the rows and cols in this newly created sheet. Add it to a list of sheets in this workbook.
            _sheets.Add(worksheetPart.Worksheet.AppendChild(new SheetData()));
        }

        public void CloseDocument()
        {
            document.Close();
        }

        /* Add employees to the newly created file
         */
        public void AddEmployees()
        {
            // Get the list of employees from the Employees Class
            List<Employee> employees = Employees.EmployeesList;

            // Append a sheet data class to the worksheet. This acts
            // as container for all the rows and columns.
            //SheetData sheetData = worksheetPart.Worksheet.AppendChild(new SheetData());

            // Constructing header
            Row row = new Row();

            row.Append(
                ConstructCell("Id", CellValues.String),
                ConstructCell("Name", CellValues.String),
                ConstructCell("Birth Date", CellValues.String),
                ConstructCell("Salary", CellValues.String));

            // Insert the header row to the Sheet Data
            _sheets[0].AppendChild(row);

            // Inserting each employee
            foreach (var employee in employees)
            {
                row = new Row();

                row.Append(
                    ConstructCell(employee.Id.ToString(), CellValues.Number),
                    ConstructCell(employee.Name, CellValues.String),
                    ConstructCell(employee.DOB.ToString("yyyy/MM/dd"), CellValues.String),
                    ConstructCell(employee.Salary.ToString(), CellValues.Number));
                // Insert the data into the Sheet Data
                _sheets[0].AppendChild(row);
            }
            worksheetPart.Worksheet.Save();
        }

        /* Create a cell object and populate it. Takes the value of the cell and type of data
         * being put in as paramters.
         * @param value: the value to be shown in Excel
         * @param dataType: the type of data for Excel to handle 
         * @return: A cell object
         */ 
        private Cell ConstructCell(string value, CellValues dataType)
        {
            return new Cell()
            {
                CellValue = new CellValue(value),
                DataType = new EnumValue<CellValues>(dataType)
            };
        }
    }
}
