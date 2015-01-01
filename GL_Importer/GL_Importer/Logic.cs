using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;

namespace GL_Importer
{
    class Logic
    {
        private static JournalEntry GetCell(Row r, int row, WorkbookPart workbookPart)
        {
            int i = 0;
            string[] column = new string[5]{("A"), ("B"), ("C"), ("D"), ("E")};
            Cell cell = new Cell();
            List<string> list = new List<string>();
            string text = "";
            while(i < 5)
            {
                cell = r.Elements<Cell>().Where (c => string.Compare (c.CellReference.Value, column[i] + row, true) == 0).First();
                text = cell.CellValue.Text;
                if(i == 4)
                {
                    text = JournalEntry.stringConverter(workbookPart, Convert.ToInt32(cell.CellValue.Text));
                }
                list.Add(text);
                i++;
            }
            JournalEntry je = new JournalEntry();
            je = JournalEntry.Build(list);
            return (je);
        }
        public static List<JournalEntry> Program()
        {
            string hardcoded;
            hardcoded = @"GL Import 2010.xlsx";
            //Workbook will break if more than 1 worksheet
            List<string> testList = new List<string>();
            List<JournalEntry> storage = new List<JournalEntry>();

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(hardcoded, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                JournalEntry je = new JournalEntry();
                int i = 0;
                decimal zeroBal = 0;
                int count = 0;
                foreach(Row r in sheetData.Elements<Row>())
                {
                    i = Convert.ToInt32((r.RowIndex.ToString()));
                    je = Logic.GetCell(r, i, workbookPart);
                    zeroBal = zeroBal + je.Amount;
                    storage.Add(je);
                    count++;

                }
                if (zeroBal != 0)
                {
                    //throw exception
                }
                return storage;
            }
        }
    }
}
