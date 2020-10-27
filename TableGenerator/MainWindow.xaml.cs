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
using System.IO;
using System.Collections.Specialized;

namespace TableGenerator
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml.
    /// Contains all methods for interaction with MainWindow.xaml.
    /// </summary>
    public partial class MainWindow : Window
    {
        /// <summary>
        /// Variable for selected directory.
        /// </summary>
        private string folderName;

        /// <summary>
        /// Variable for collection of files paths.
        /// </summary>
        private string[] filesPaths;

        /// <summary>
        /// Class constructor.
        /// </summary>
        public MainWindow()
        {
            InitializeComponent();

            ((INotifyCollectionChanged)FilesNamesListBox.Items).CollectionChanged += FilesNamesListBox_CollectionChanged;
        }

        /// <summary>
        /// Method for deleting all items in FilesNamesListBox.
        /// </summary>
        /// <param name="sender">Referance to object that raised event.</param>
        /// <param name="e">Event data.</param>
        private void DeleteAllButton_Click(object sender, RoutedEventArgs e)
        {
            FilesNamesListBox.Items.Clear();
        }

        /// <summary>
        /// Method for deleting selected item in FilesNamesListBox.
        /// </summary>
        /// <param name="sender">Referance to object that raised event.</param>
        /// <param name="e">Event data.</param>
        private void DeleteButton_Click(object sender, RoutedEventArgs e)
        {
            FilesNamesListBox.Items.Remove(FilesNamesListBox.SelectedItem);
        }

        /// <summary>
        /// Method for selecting directory with files and running method for getting files.
        /// </summary>
        /// <param name="sender">Referance to object that raised event.</param>
        /// <param name="e">Event data.</param>
        private void SelectButton_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog folderDlg = new System.Windows.Forms.FolderBrowserDialog
            {
                ShowNewFolderButton = true
            };

            System.Windows.Forms.DialogResult result = folderDlg.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                this.folderName = folderDlg.SelectedPath;
                GetFilesPathByFolderName();
            }
        }

        /// <summary>
        /// Method for getting files paths and setting names in FilesNamesListBox.
        /// </summary>
        private void GetFilesPathByFolderName()
        {
            this.filesPaths = Directory.GetFiles(this.folderName, "*.*", SearchOption.AllDirectories);
            List<string> filesNames = StringCollectionLibrary.getNamesOfFilesFromPaths(this.filesPaths);
            AppendToListBox(FilesNamesListBox, filesNames);
        }

        /// <summary>
        /// Method for appending text fields to list box.
        /// </summary>
        /// <param name="listBox">List box to append in.</param>
        /// <param name="list">List of appending items.</param>
        private void AppendToListBox(ListBox listBox, List<string> list)
        {
            listBox.Items.Clear();
            list.ForEach(delegate (string name)
            {
                TextBlock textBlock = new TextBlock();
                textBlock.Text = name;
                listBox.Items.Add(textBlock);
            });
        }

        /// <summary>
        /// Method for handling content changes in FilesNamesListBox.
        /// </summary>
        /// <param name="sender">Referance to object that raised event.</param>
        /// <param name="e">Event data.</param>
        private void FilesNamesListBox_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            int countOfItems = FilesNamesListBox.Items.Count;

            if(countOfItems > 0)
            {
                DeleteAllButton.IsEnabled = true;
                GenerateButton.IsEnabled = true;
                DeleteButton.IsEnabled = true;
            } else
            {
                DeleteAllButton.IsEnabled = false;
                GenerateButton.IsEnabled = false;
                DeleteButton.IsEnabled = false;
            }
        }

        /// <summary>
        /// Method for generating Microsoft Excel table from files paths.
        /// </summary>
        /// <param name="sender">Referance to object that raised event.</param>
        /// <param name="e">Event data.</param>
        private void GenerateButton_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            if (excelApp == null)
            {
                MessageBox.Show("Excel is not properly installed on your computer!");
            } else
            {
                string fileSavePath = getSaveFilePath();
                if (fileSavePath != "")
                {
                    Microsoft.Office.Interop.Excel.Workbook workbook = excelApp.Workbooks.Add();
                    Microsoft.Office.Interop.Excel.Worksheet worksheet = workbook.Worksheets[1];

                    worksheet.Name = "List of materials";
                    string[] header = new string[6] { "Type", "Dimensions", "Steel grade", "Heat Nr", "Certificate nr", "Link" };
                    for (int column = 1; column <= header.Length; column++)
                    {
                        worksheet.Cells[1, column].Value = header[column - 1];
                    }

                    List<string> filesNames = StringCollectionLibrary.getNamesOfFilesFromPaths(this.filesPaths);
                    List<string[]> dataCollection = StringCollectionLibrary.getCollectionOfFormattedNames(filesNames);
                    for (int row = 0; row < dataCollection.Count; row++)
                    {
                        for (int column = 1; column <= header.Length; column++)
                        {
                            if (column < 6)
                            {
                                worksheet.Cells[row + 2, column].Value = dataCollection[row][column - 1];
                            }
                            else
                            {
                                Microsoft.Office.Interop.Excel.Range excelCell = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[row + 2, column];
                                worksheet.Hyperlinks.Add(excelCell, this.filesPaths[row], Type.Missing, this.filesPaths[row], this.filesPaths[row]);
                                worksheet.Cells[row + 2, column].Value = this.filesPaths[row];
                            }
                        }

                    }

                    workbook.SaveAs(fileSavePath);
                    workbook.Close();
                    MessageBox.Show("File is saved sucessfully!");
                }
            }
        }

        /// <summary>
        /// Method for getting save path for excel file.
        /// </summary>
        /// <returns>Path for saving file.</returns>
        private string getSaveFilePath()
        {
            System.Windows.Forms.SaveFileDialog saveFileDialog = new System.Windows.Forms.SaveFileDialog();
            saveFileDialog.Filter = "Excel |*.xlsx";
            saveFileDialog.ShowDialog();
            string fileSavePath = saveFileDialog.FileName;

            return fileSavePath;
        }
    }
}