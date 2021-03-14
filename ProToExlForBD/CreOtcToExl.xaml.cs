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

namespace ProToExlForBD
{
    /// <summary>
    /// Логика взаимодействия для CreOtcToExl.xaml
    /// </summary>
    public partial class CreOtcToExl : Window
    {
      
        public CreOtcToExl( )
        {
            InitializeComponent();
            DT.ItemsSource = QWER.WQER().Product.ToList();
          
        }

        private void EditBtn_Click(object sender, RoutedEventArgs e)
        {

        }

        private void EditBtn_Click_1(object sender, RoutedEventArgs e)
        {
            ToNexForExl WW = new ToNexForExl((sender as Button).DataContext as Product, 1);
            WW.Show();
        }
    }
}
