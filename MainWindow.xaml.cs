using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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

namespace QRGenerator
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

        private async void Button_Click(object sender, RoutedEventArgs e)
        {
            ProcessRange();   
        }

        private void NumberValidationTextBox(object sender, TextCompositionEventArgs e)
        {
            Regex regex = new Regex("[^0-9]+");
            e.Handled = regex.IsMatch(e.Text);
        }


        // This is a thread-safe method. You can run it in any thread
        internal async Task ProcessRange()
        {
            List<QRModel> listQr = new List<QRModel>();
            int count = int.Parse(finalRange.Text.ToString()) - int.Parse(initialRange.Text.ToString());
         
            foreach (int item in Enumerable.Range(int.Parse(initialRange.Text.ToString()), count+1))
            {
                string phrase = textPhrase.Text + " - " + item.ToString();
                
                listQr.Add(new QRModel
                {
                    Phrase = phrase,
                    Range = item
                });

                CodeGenerator.GenerateQR(phrase, item.ToString());
            }
                OdtMaker.GenerateOdt(listQr);
        }
    }
}
