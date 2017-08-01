using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using Forms = System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelContactList
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        string fileName;            // file name of contact list excel file
        string[] bgImagePaths;      // file names of background images
        bool shuttingDown = false;  // checked by background image thread to know when to stop

        enum MainField { _title, _fName, _lName, _addrLine1, _addrLine2, _city, _zip, _email, _phone}   // used for indexing into TEXTINPUT and DEFAULTTXT
        bool[] textInput = new bool[] { false, false, false, false, false, false, false, false, false}; // true if user has input text into field, false otherwise
        string[] defaultTxt = new String[] { "Title", "First Name", "Last Name", "Address Line 1", "Address Line 2", "City", "Zip Code", "", "XXX-XXX-XXXX" };
        SolidColorBrush defaultColor = new SolidColorBrush(Colors.Gray);
        SolidColorBrush inputColor = new SolidColorBrush(Colors.Black);

        public MainWindow(string filePath, string imgPath)
        {
            InitializeComponent();
            fileName = filePath;
            ResetFields();
            if (imgPath != null)
            {
                InitializeView(imgPath);
            }
            
        }

        /* Returns the MainField enum value for the given TextBox */
        private MainField GetFieldEnum(TextBox box)
        {
            switch (box.Name)
            {
                case "title":
                    return MainField._title;
                case "fName":
                    return MainField._fName;
                case "lName":
                    return MainField._lName;
                case "addrLine1":
                    return MainField._addrLine1;
                case "addrLine2":
                    return MainField._addrLine2;
                case "city":
                    return MainField._city;
                case "zip":
                    return MainField._zip;
                case "email":
                    return MainField._email;
                case "phone":
                    return MainField._phone;
                default:
                    throw new ArgumentException("Unrecognized textbox name.");
            }
        }

        private void ResetFields()
        {
            SetDefaultText(title);
            SetDefaultText(fName);
            SetDefaultText(lName);
            SetDefaultText(addrLine1);
            SetDefaultText(addrLine2);
            SetDefaultText(city);
            defaultState.IsSelected = true;
            SetDefaultText(zip);
            SetDefaultText(email);
            SetDefaultText(phone);
        }

        /* Sets the default text for a text field.*/
        private void SetDefaultText(TextBox box)
        {
            int fieldIdx = (int)GetFieldEnum(box);
            box.Foreground = defaultColor;
            box.Text = defaultTxt[fieldIdx];
            textInput[fieldIdx] = false;
        }

        
        /* Sets the background image */
        private async void InitializeView(string imgPath)
        {
            char[] quote = new char[] { '"' };
            string[] separator = new string[] { "\" \"" };
            bgImagePaths = imgPath.Trim(quote).Split(separator, StringSplitOptions.RemoveEmptyEntries);
            BitmapImage bmImg = new BitmapImage();
            bmImg.BeginInit();
            bmImg.UriSource = new Uri(bgImagePaths[0]);
            bmImg.EndInit();
            bgImage.ImageSource = bmImg;

            /* If more than one background image was selected, start the thread that handles background image cycling. */
            if (bgImagePaths.Length > 1)
            {
                TimeSpan interval = new TimeSpan(1, 0, 0);
                CancellationToken tok = new CancellationToken();
                int i = 1;
                while (!shuttingDown)
                {
                    bmImg = new BitmapImage();
                    bmImg.BeginInit();
                    bmImg.UriSource = new Uri(bgImagePaths[i]);
                    bmImg.EndInit();
                    bgImage.ImageSource = bmImg;
                    i = (i + 1) % bgImagePaths.Length;
                    await Task.Delay(interval, tok);
                }
            } 
        }

        
        /* Saves the entered information to the excel file FILENAME*/
        private void Submit_Click(object sender, RoutedEventArgs e)
        {
            if (IsInputValid())
            {
                Excel.Application excelApp = null;
                try
                {
                    excelApp = new Excel.Application();
                    Excel.Workbook wkbk = excelApp.Workbooks.Open(fileName);
                    if (wkbk.ReadOnly)
                    {
                        throw new Exception("Contact list file may be in use by another application.");
                    }
                    Excel.Worksheet sheet = excelApp.ActiveSheet;

                    int i = 1;
                    bool blank_row;
                    do
                    {
                        i++;
                        blank_row = true;
                        string A = "A" + i;
                        string J = "K" + i;
                        Excel.Range row = sheet.Range[A, J];
                        foreach (Excel.Range r in row)
                        {
                            if (r.Value2 != null)
                            {
                                blank_row = false;
                                break;
                            }
                        }
                    }
                    while (!blank_row);

                    if (textInput[(int)MainField._title])
                        sheet.Cells[i, "A"] = title.Text;
                    sheet.Cells[i, "B"] = fName.Text;
                    sheet.Cells[i, "C"] = lName.Text;
                    if (textInput[(int)MainField._addrLine1])
                    {
                        sheet.Cells[i, "D"] = addrLine1.Text;
                        if (textInput[(int)MainField._addrLine2])
                            sheet.Cells[i, "E"] = addrLine2.Text;
                        sheet.Cells[i, "F"] = city.Text;
                        sheet.Cells[i, "G"] = state.Text;
                        sheet.Cells[i, "H"] = zip.Text;
                    }
                    if (textInput[(int)MainField._email])
                        sheet.Cells[i, "I"] = email.Text;
                    if (textInput[(int)MainField._phone])
                        sheet.Cells[i, "J"] = phone.Text;
                    sheet.Cells[i, "K"] = DateTime.Today.ToShortDateString();
                    wkbk.Save();

                    ResetFields();
                }
                catch (Exception err)
                {
                    Forms.MessageBox.Show(err.Message, "Error: Unable to save contact information.");
                }
                finally
                {
                    if (excelApp != null)
                    {
                        excelApp.Quit();
                        excelApp = null;
                    }
                }
            }
        }

        private bool IsInputValid()
        {
            bool result = false;
            // if first and last name are blank
            if (!(textInput[(int)MainField._fName] && textInput[(int)MainField._lName]))
            {
                Forms.MessageBox.Show("Please enter a first and last name.");
            }
            // if address, email, and phone are blank
            else if (!(textInput[(int)MainField._addrLine1] || textInput[(int)MainField._addrLine2] || textInput[(int)MainField._city]
                || textInput[(int)MainField._zip] || textInput[(int)MainField._email] || textInput[(int)MainField._phone]))
            {
                Forms.MessageBox.Show("Please enter either a mailing address, email address, or phone number.");
            }
            else if (!isValidAddress())
            {

            }
            else if (textInput[(int)MainField._email] && !Regex.IsMatch(email.Text, "\\A.+@.+\\..+\\z"))
            {
                Forms.MessageBox.Show("Please enter a valid email address.");
            }
            else if (!IsValidPhone())
            {
                
            }
            else
            {
                result = true;
            }

            if (result == true)
            {
                string nameStr = null;
                string addrStr = null;
                string emailStr = null;
                string phoneStr = null;

                nameStr = "";
                if (textInput[(int)MainField._title])
                    nameStr += title.Text + " ";
                nameStr += fName.Text + " " + lName.Text;
                
                if (textInput[(int)MainField._addrLine1])
                {
                    addrStr = addrLine1.Text + "\n";
                    if (textInput[(int)MainField._addrLine2])
                        addrStr += addrLine2.Text + "\n";
                    addrStr += city.Text + ", " + state.Text + " " + zip.Text;
                }

                if (textInput[(int)MainField._email])
                    emailStr = email.Text;

                if (textInput[(int)MainField._phone])
                    phoneStr = phone.Text;

                ConfirmationWindow win = new ConfirmationWindow(nameStr, addrStr, emailStr, phoneStr);
                win.ShowDialog();
                result = win.Confirmed;
            }

            return result;
        }

        /* Returns true if address fields are valid or are all blank. */
        private bool isValidAddress()
        {
            bool result = true;
            bool line1Bool = textInput[(int)MainField._addrLine1];
            bool line2Bool = textInput[(int)MainField._addrLine2];
            bool cityBool = textInput[(int)MainField._city];
            bool zipBool = textInput[(int)MainField._zip];

            if (line1Bool || line2Bool || cityBool || zipBool)
            {
                if (!line1Bool || !cityBool || !zipBool)
                {
                    Forms.MessageBox.Show("Please enter a street address, city and zip code.");
                    result = false;
                } else if (!Regex.IsMatch(zip.Text, "\\A\\d{5}(-\\d{4})?\\z"))
                {
                    Forms.MessageBox.Show("Please enter a valid zip code.");
                    result = false;
                }
            }

            return result;
        }

        /* Returns true if phone number is valid or blank */
        private bool IsValidPhone()
        {
            bool result = true;

            // phone is blank or matches XXX-XXX-XXXX
            if (!textInput[(int)MainField._phone] || Regex.IsMatch(phone.Text, "\\A\\d{3}-\\d{3}-\\d{4}\\z"))
            {
                
            }
            else if (Regex.IsMatch(phone.Text, "(\\d+[()\\- ]+)+\\d+"))
            {
                Regex digits = new Regex("\\d+");
                Match m = digits.Match(phone.Text);
                string num = m.Value;
                m = m.NextMatch();
                while (m.Success)
                {
                    num += "-" + m.Value;
                    m = m.NextMatch();
                }
                phone.Text = num;
            }
            else if (Regex.IsMatch(phone.Text, "\\A\\d{10}\\z"))
            {
                string ph = phone.Text;
                phone.Text = ph.Substring(0, 3) + "-" + ph.Substring(3, 3) + "-" + ph.Substring(6);
            }
            else
            {
                Forms.MessageBox.Show("Please enter a phone number in the format, XXX-XXX-XXXX");
                result = false;
            }

            return result;
        }

        

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            shuttingDown = true;
        }

        private void FieldGotFocus(object sender, RoutedEventArgs e)
        {
            TextBox box = (TextBox)sender;
            int fieldIdx = (int)GetFieldEnum(box);
            if (!textInput[fieldIdx])
            {
                box.Text = "";
                box.Foreground = inputColor;
            }
        }

        private void FieldLostFocus(object sender, RoutedEventArgs e)
        {
            TextBox box = (TextBox)sender;
            MainField fieldEnumVal = GetFieldEnum(box);
            if (box.Text == "")
                SetDefaultText(box);
            else
                textInput[(int)fieldEnumVal] = true;
        }
    }
}
