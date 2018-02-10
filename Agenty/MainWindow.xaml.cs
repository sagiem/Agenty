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

namespace Agenty
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        string file;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openfile = new OpenFileDialog();
            openfile.Multiselect = false;
            openfile.DefaultExt = "*.xls;*.xlsx";
            openfile.Filter = "Microsoft Excel (*.xls*)|*.xls*";
            openfile.Title = "Выберите документ Excel";
            openfile.ShowDialog();
            if(openfile.FileName != null)
            {
                this.file = openfile.FileName;
            }
        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            Raschet raschet = new Raschet(file);
            //raschet.Exelreader();
            raschet.ExelAkt();
        }
    }
}
