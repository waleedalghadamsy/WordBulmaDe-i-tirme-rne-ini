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

namespace WordBulmaDeğiştirmeÖrneğini
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

        private void BulDeğiştir(object sender, RoutedEventArgs e)
        {
            try
            {
                Microsoft.Office.Interop.Word.Application wordApp = 
                                    new Microsoft.Office.Interop.Word.Application();
                object dosyaYolu = @"C:\Users\pc\Documents\Bulma Değiştirme Örneğini.docx";
                Microsoft.Office.Interop.Word.Document document = wordApp.Documents.Open(ref dosyaYolu);
                Microsoft.Office.Interop.Word.Range kelime;

                for (int i = 1; i < document.Words.Count; i++)
                {
                    kelime = document.Words[i];

                    if (kelime.Text.Equals(bulTextBox.Text))
                        kelime.Text = ileDeğiştirinTextBox.Text;
                }

                sonuçTextBlock.Text = "Değiştirme yaptı"; sonuçTextBlock.Foreground = Brushes.Green;

                document.SaveAs2(@$"C:\Users\pc\Documents\{değiştirilmişDosyaTextBox.Text}");

                document.Close();

                wordApp.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Hata", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
