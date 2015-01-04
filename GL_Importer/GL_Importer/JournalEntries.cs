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

        private object GetValueAt(Row r, int colIndex, SharedStringTablePart sharedStrings)
        {
            Cell cell = r.Descendants<Cell>().ElementAt(colIndex);
            return cell.CellValue.InnerText;
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
                int i = 0; 
                foreach (Row r in sheetData.Elements<Row>())
                {
                    try
                    {
                        Entries.Add(new JournalEntry
                        {
                            seg1 = Int32.Parse(GetValueAt(r, 0, workbookPart.SharedStringTablePart).ToString()),
                            seg2 = Int32.Parse(GetValueAt(r, 1, workbookPart.SharedStringTablePart).ToString()),
                            seg3 = Int32.Parse(GetValueAt(r, 2, workbookPart.SharedStringTablePart).ToString()),
                            Amount = decimal.Parse(GetValueAt(r, 3, workbookPart.SharedStringTablePart).ToString()),
                            lineItem = GetValueAt(r, 4, workbookPart.SharedStringTablePart).ToString()
                        });
                    }
                    catch (Exception ex)
                    {
                        //TODO: REMOVE THROW AND LOG USER FRIENDLY ERROR TO DISPLAY IN THE UI
                        throw ex;
                    }

                    //this.GetValueAt(r, i, workbookPart.SharedStringTablePart);
                }
            }
        }
    }
}
