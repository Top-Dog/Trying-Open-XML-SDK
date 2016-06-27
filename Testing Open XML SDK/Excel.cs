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
    public class Document
    {
        public string fileName { get; set; }
        public SpreadsheetDocument document { get; set; }
        public WorkbookPart workbookPart { get; set; }
        public WorksheetPart worksheetPart { get; set; }
        public Sheets sheets { get; set; }
        public List<Sheet> _sheet;
    }

    public class Excel
    {
        static List<Document> _documents;

        public void Initialize()
        {
            _documents = new List<Document>();
        }

        public void AddWorkbook(string fileName)
        {
            // Create a new spreadsheet document with a name spescified by the method fileName parameter
            SpreadsheetDocument document = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook);

            // Add a workbookPart to the document
            WorkbookPart workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            // Add a WorksheetPart to the workbookPart
            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

            // Add a container for all the sheets
            Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());

            // Save the new document
            workbookPart.Workbook.Save();

            // Add the new workbook to a list referenced by fileName
            _documents.Add(new Document() 
                { 
                    fileName = fileName,
                    document = document,
                    workbookPart = workbookPart,
                    worksheetPart = worksheetPart,
                    sheets = sheets,
                    _sheet = new List<Sheet>()
                });
        }

        public void AddWorksheet(string fileName, string sheetName)
        {
            WorkbookPart activewbPart = GetDocument(fileName).workbookPart; // might return null
            WorksheetPart activewsPart = GetDocument(fileName).worksheetPart;
            Sheets sheets = GetDocument(fileName).sheets;

            activewsPart.Worksheet = new Worksheet();

            // Create a sheet (associated with the WorksheetPart) and append it to the "sheets" abstract object.
            Sheet sheet = new Sheet() { Id = activewbPart.GetIdOfPart(activewsPart), SheetId = GetNumberWorksheets(GetAllWorksheets(activewbPart)) + 1, Name = sheetName };
            sheets.Append(sheet);
            // SheetData is a container for all the rows and cols in this newly created sheet. Add it to a list of sheets in this workbook.
            activewsPart.Worksheet.AppendChild(new SheetData());

            //GetDocument(fileName)._sheet.AppendChild(activewsPart.Worksheet)
        }

        public void GetWorkbook(string workbook)
        {

        }

        public void GetWorksheet(string workbook, string worksheet)
        {

        }

        public void CloseDocument(string fileName)
        {
            Document doc = GetDocument(fileName);
            if (doc != null)
            {
                doc.workbookPart.Workbook.Save();
                doc.worksheetPart.Worksheet.Save();
                doc.document.Close();
            }
        }

        // Retrieve a List of all the sheets in a workbook.
        // The Sheets class contains a collection of 
        // OpenXmlElement objects, each representing one of 
        // the sheets.
        public static Sheets GetAllWorksheets(WorkbookPart wbPart)
        {
            return wbPart.Workbook.Sheets;
        }

        public static uint GetNumberWorksheets(Sheets sheets)
        {
            uint count = 0;
            foreach (Sheet sheet in sheets) 
            {
                count += 1;
            }
            return count;
        }

        /* Gets the documents class oject from the list */
        public static Document GetDocument(string fileName)
        {
            foreach (Document doc in _documents)
            {
                if (doc.fileName == fileName)
                {
                    return doc;
                }
            }
            return null;
        }
    }
}
