using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using Forms = System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelContactList
{
    /// <summary>
    /// Interaction logic for setup.xaml
    /// </summary>
    public partial class setup : Window
    {
        
        public setup()
        {
            InitializeComponent();
        }

        /* Disable file selection for using existing file, and enable file
         * selection for creating new file.
         */
        private void New_RadioButton_Checked(object sender, RoutedEventArgs e)
        {
            if (newButton != null)
                newButton.IsEnabled = true;
            if (openButton != null)
                openButton.IsEnabled = false;
        }

        /* Disable file selection for creating new file, and enable file selection
         * for using existing file.
         */
        private void Open_RadioButton_Checked(object sender, RoutedEventArgs e)
        {
            if (newButton != null)
                newButton.IsEnabled = false;
            if (openButton != null)
                openButton.IsEnabled = true;
        }

        /* File selection for new file. */
        private void newButton_Click(object sender, RoutedEventArgs e)
        {
            Forms.SaveFileDialog dlg = new Forms.SaveFileDialog();
            dlg.Filter = "Worksheets (*.xlsx;*.xlsm;*.xlsb;*.xls)|*.xlsx;*.xlsm;*.xlsb;*.xls|All files (*.*)|*.*";
            dlg.FileName = "contact-list";
            dlg.DefaultExt = "*.xlsx";
            if (dlg.ShowDialog() == Forms.DialogResult.OK)
            {
                newText.Text = dlg.FileName;
            }
        }

        /* File selection for existing file. */
        private void openButton_Click(object sender, RoutedEventArgs e)
        {
            Forms.OpenFileDialog dlg = new Forms.OpenFileDialog();
            dlg.Filter = "Worksheets (*.xlsx;*.xlsm;*.xlsb;*.xls)|*.xlsx;*.xlsm;*.xlsb;*.xls|All files (*.*)|*.*";
            if (dlg.ShowDialog() == Forms.DialogResult.OK)
            {
                openText.Text = dlg.FileName;
            }
        }

        /* Enable file selection for background image. */
        private void imgCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            imgButton.IsEnabled = true;
        }

        /* Disable file selection for background image. */
        private void imgCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            imgButton.IsEnabled = false;
        }

        /* File selection for background image */
        private void imgButton_Click(object sender, RoutedEventArgs e)
        {
            Forms.OpenFileDialog dlg = new Forms.OpenFileDialog();
            dlg.Filter = "Image (*.jpg;*.jpeg;*.png)|*.jpg;*.jpeg;*.png|All files (*.*)|*.*";
            dlg.Multiselect = true;
            if (dlg.ShowDialog() == Forms.DialogResult.OK)
            {
                
                string[] files = dlg.FileNames;
                if (files.Length == 1)
                {
                    imgText.Text = dlg.FileName;
                }
                else
                {
                    /* If more than one file is selected, file names are enclosed in quotes and
                     * separated by a space in the text box. */
                    string fileTxt = "\"" + files[0] + "\"";
                    int i = 1;
                    while (i < files.Length)
                    {
                        fileTxt += " \"" + files[i] + "\"";
                        i++;
                    }
                    imgText.Text = fileTxt;
                }
            }
        }

        /* If neccessary files have been selected, call InitializeFile, show main window, and
         * close this window. Otherwise, display error message.
         */
        private void okButton_Click(object sender, RoutedEventArgs e)
        {
            bool validated = true;
            bool newFile = false;
            string filePath = "";
            string imgFile = null;

            if (newRadio.IsChecked.HasValue && newRadio.IsChecked.Value)
            {
                if (newText.Text == "")
                {
                    validated = false;
                    Forms.MessageBox.Show("Enter a file name to create.");
                }
                else
                {
                    newFile = true;
                    filePath = newText.Text;
                }
            }
            else if (openRadio.IsChecked.HasValue && openRadio.IsChecked.Value)
            {
                if (openText.Text == "")
                {
                    validated = false;
                    Forms.MessageBox.Show("Select a file to use.");
                }
                else
                {
                    newFile = false;
                    filePath = openText.Text;
                }
            }
            else
            {
                validated = false;
                Forms.MessageBox.Show("Neither radio button selected.", "Error", Forms.MessageBoxButtons.OK, Forms.MessageBoxIcon.Error);
            }
            if (imgCheckBox.IsChecked.HasValue && imgCheckBox.IsChecked.Value)
            {
                if (imgText.Text == "")
                {
                    validated = false;
                    Forms.MessageBox.Show("Select an image file to use.");
                }
                else
                {
                    imgFile = imgText.Text;
                }
            }
            else
            {
                imgFile = null;
            }
            if (validated)
            {
                try
                {
                    if (newRadio.IsChecked.HasValue && newRadio.IsChecked.Value)
                        InitializeFile(filePath);
                    new MainWindow(filePath, imgFile).Show();
                    this.Close();
                }
                catch (Exception ex)
                {
                    Forms.MessageBox.Show(ex.Message, ex.GetType().ToString(), Forms.MessageBoxButtons.OK, Forms.MessageBoxIcon.Error);
                }
            }
        }

        /* Save column headers to new excel file. */
        private void InitializeFile(string fileName)
        {
            Excel.Application excelApp = null;
            try
            {
                excelApp = new Excel.Application();
                excelApp.Workbooks.Add();
                Excel.Worksheet sheet = excelApp.ActiveSheet;
                sheet.Cells[1, "A"] = "Title";
                sheet.Cells[1, "B"] = "First Name";
                sheet.Cells[1, "C"] = "Last Name";
                sheet.Cells[1, "D"] = "Address Line 1";
                sheet.Cells[1, "E"] = "Address Line 2";
                sheet.Cells[1, "F"] = "City";
                sheet.Cells[1, "G"] = "State";
                sheet.Cells[1, "H"] = "Zip Code";
                sheet.Cells[1, "I"] = "E-mail";
                sheet.Cells[1, "J"] = "Phone";
                sheet.Cells[1, "K"] = "Date Added";
                for (int i = 1; i <= 11; i++)
                    sheet.Columns[i].ColumnWidth = 20;
                excelApp.DisplayAlerts = false;  // diables overwrite file message when calling SaveAs
                try
                {
                    sheet.SaveAs(fileName);
                }
                catch (System.Runtime.InteropServices.COMException err)
                {
                    string msg = "Error saving new contact list file. If new file should overwrite another," +
                        " the overwritten file must not be open in another application.";
                    throw new Exception(msg);

                }
                excelApp.DisplayAlerts = true;
            }
            finally
            {
                if (excelApp != null)
                {
                    excelApp.Quit();
                }
            }
        }

        private void exitButton_Click(object sender, RoutedEventArgs e)
        {
            App.Current.Shutdown();
        }

    }
}
