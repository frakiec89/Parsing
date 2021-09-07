using Microsoft.Win32;
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

namespace Parsing
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

        private void buttonSelect_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Только  ворд файлы (*.doc)|*.doc|Все файлы (*.*)|*.*";

            if (openFileDialog.ShowDialog()== true)
            {
                ParsingWord parsingWord = new ParsingWord();
                try
                {
                    content.Text = parsingWord.Metod(openFileDialog.FileName);
                }
                catch  (Exception ex)
                {
                    content.Text = ex.Message;
                }
            }
        }

        private void btREg_Click(object sender, RoutedEventArgs e)
        {
            ParsingWord parsingWord = new ParsingWord();
            content.Text = parsingWord.Reg(content.Text);

        }
    }
}
