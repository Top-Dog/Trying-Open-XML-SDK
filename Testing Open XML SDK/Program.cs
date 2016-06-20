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
        private uint NumberOfSheets = 0;
        private WorkbookPart workbookPart;
        private Sheets sheets;
        private WorksheetPart worksheetPart;
        private SheetData sheetData;
        private SpreadsheetDocument document;

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
        }

        public void CreateExcelDoc(string fileName, string sheetName)
        {
            // Create a new spreadsheet document with a name spescified by the method fileName parameter
            document = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook);
            //using (SpreadsheetDocument document = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook))
            //{
                // Add a workbookPart to the document
                workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();
                sheets = workbookPart.Workbook.AppendChild(new Sheets());

                // Add a WorksheetPart to the workbookPart
                worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet();
                sheetData = worksheetPart.Worksheet.AppendChild(new SheetData());

                // Add the first sheet
                AddSheet(sheetName);

                // Save the new document
                workbookPart.Workbook.Save();

                //AddEmployees(); // try and move this... 'using' using might be like python 'with', it closes the document afterwards
            //}
        }

        public void AddSheet(string sheetName)
        {
            // Create a sheet (associated with the WorksheetPart) and append it to the "sheets" abstract object.
            NumberOfSheets += 1;
            Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = NumberOfSheets, Name = sheetName };
            sheets.Append(sheet);
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
            sheetData.AppendChild(row);

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
                sheetData.AppendChild(row);
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
