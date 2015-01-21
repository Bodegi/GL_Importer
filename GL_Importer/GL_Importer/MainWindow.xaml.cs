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
                        case ".XLSX":
                            break;
                        default:
                            MessageBox.Show("Please select a valid .XLSX file", "Error");
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
            string path = txtFile.Text;
            JournalEntries entry = new JournalEntries(path);
            var message = "";
            if (entry.Errors.Count != 0)
            {
                    message = string.Join(Environment.NewLine, entry.Errors);
                    MessageBox.Show(message, "Errors in " + path);
            }
            else
            {
                MessageBox.Show("No Errors detected", "Errors in " + path);
                JournalEntries.ExportToXml(entry.Entries);
                MessageBox.Show("Xml file created", "XML Creation Successfull");
            }
        }
    }
}