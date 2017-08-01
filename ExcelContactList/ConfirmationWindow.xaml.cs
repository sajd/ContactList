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
using System.Windows.Shapes;

namespace ExcelContactList
{
    /// <summary>
    /// Interaction logic for ConfirmationWindow.xaml
    /// </summary>
    public partial class ConfirmationWindow : Window
    {
        internal bool Confirmed { get; set; }

        public ConfirmationWindow(string nm, string ad, string em, string ph)
        {
            GridLength blank = new GridLength(0);

            InitializeComponent();
            if (nm == null)
                row1.Height = blank;
            else
                nameTxt.Text = nm;
            if (ad == null)
                row2.Height = blank;
            else
                addrTxt.Text = ad;
            if (em == null)
                row3.Height = blank;
            else
                emailTxt.Text = em;
            if (ph == null)
                row4.Height = blank;
            else
                phoneTxt.Text = ph;
            Confirmed = false;
        }

        private void OK_Click(object sender, RoutedEventArgs e)
        {
            Confirmed = true;
            this.Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
        
    }
}
