using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace GL_Importer
{
    public class JournalEntries
    {
        public List<JournalEntry> Entries { get; set; }

        public JournalEntries()
        {
            this.Entries = new List<JournalEntry>();
        }

        private JournalEntry GetValueFromRow(Row r, SharedStringTablePart sharedStrings)
        {
            foreach (var item in r.Elements<Cell>())
            {
                if (item.DataType != CellValues.SharedString)
                {
                    Console.WriteLine(item.CellValue.ToString());
                }
                else
                {
                    int sharedStringID = -1;
                    if (Int32.TryParse(item.CellValue.ToString(), out sharedStringID))
                    {
                        var ssi = sharedStrings.SharedStringTable.Elements<SharedStringItem>().ElementAt(sharedStringID);
                    }
                    else
                    {

                    }
                }
            }
            return new JournalEntry();
        }

        private JournalEntry GetValueAt(Row r, string colIndex, SharedStringTablePart sharedStrings)
        {
            return new JournalEntry();
        }

        private List<string> Validation()
        {
            List<string> errors = new List<string>();
            return errors;
        }

        public JournalEntries(string path)
        {
            this.Entries = new List<JournalEntry>();

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(path, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                foreach (Row r in sheetData.Elements<Row>())
                {
                    this.GetValueFromRow(r, workbookPart.SharedStringTablePart);
                }
            }
        }
    }
}
