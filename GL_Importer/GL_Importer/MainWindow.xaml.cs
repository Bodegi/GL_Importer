using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;

namespace GL_Importer
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        //Logic for picking file and checking if it is proper extension
        private void btnBrowse_Click(object sender, RoutedEventArgs e)
        {
            var open = new System.Windows.Forms.OpenFileDialog();
            var fileInteract = open.ShowDialog();
            switch(fileInteract)
            {
                case System.Windows.Forms.DialogResult.OK:
                    var file = open.FileName;
                    txtFile.Text = file;
                    txtFile.ToolTip = file;
                    string extension = "";
                    int count = 0;
                    count = file.IndexOf(".");
                    extension = file.Substring(count).ToUpper();
                    switch (extension)
                    {
                        case ".XML":
                            break;
                        case ".XLS":
                            break;
                        case ".XLSX":
                            break;
                        default:
                            MessageBox.Show("Please select a valid .XML or .XLS file", "Error");
                            txtFile.Clear();
                            break;
                    }
                    break;
                case System.Windows.Forms.DialogResult.Cancel:
                default:
                    txtFile.Text = null;
                    txtFile.ToolTip = null;
                    break;
            }

        }

        //Logic for validating the data inside the selected file
        private void btnTestData_Click(object sender, RoutedEventArgs e)
        {
            string hardcoded;
            hardcoded = @"C:\Users\Chuck\Desktop\GL Import 2010.xlsx";
            //Workbook will break if more than 1 worksheet
            List<string> testList = new List<string>();
            List<JournalEntry> storage = new List<JournalEntry>();
            using (SpreadsheetDocument test = SpreadsheetDocument.Open(hardcoded, false))
            {
                WorkbookPart workbookPart = test.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();

                OpenXmlReader reader = OpenXmlReader.Create(worksheetPart);
                string text;
                while (reader.Read())
                {
                    if (reader.ElementType == typeof(CellValue))
                    {
                        text = reader.GetText();
                        testList.Add(text);
                    }
                }
                
                //WorkbookPart workbookPart = test.WorkbookPart;
                //WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                //SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                //string text;
                //foreach (Row r in sheetData.Elements<Row>())
                //{
                //    testList.Clear();
                //    foreach(Cell c in r.Elements<Cell>())
                //    {
                //        if(c.DataType == "s")
                //        {
                            
                //        }
                //        text = c.CellValue.Text;
                //        testList.Add(text);
                //    }
                //    JournalEntry je = new JournalEntry();
                //    je = JournalEntry.Build(testList);
                //    storage.Add(je);
                //}
            }
        }
    }
}
