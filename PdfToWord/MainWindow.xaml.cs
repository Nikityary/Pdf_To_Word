using Aspose.Pdf;
using GroupDocs.Conversion.Options.Convert;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
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

namespace PdfToWord
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string mydoс = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
                OpenFileDialog ofd = new OpenFileDialog();
                ofd.Filter = "PDF файлы (.pdf)|*.pdf|All Files (*.*)|*.*";
                label.Content = "Операция выполняется...";
                ofd.ShowDialog();
                string path = ofd.FileName;
                var converterInstance = new GroupDocs.Conversion.Converter(path);
                var optionsWordFile = new WordProcessingConvertOptions();
                converterInstance.Convert(mydoс + "ConvertedDoc.docx", optionsWordFile);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            label.Content = "Документ сохранён на основном диске";
        }
    }
}
