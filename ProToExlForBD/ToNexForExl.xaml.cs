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
using Excel = Microsoft.Office.Interop.Excel;

namespace ProToExlForBD
{
    /// <summary>
    /// Логика взаимодействия для ToNexForExl.xaml
    /// </summary>
    public partial class ToNexForExl : Window
    {
        private Product _currentTovar = new Product();
        private int num = 0;
        public ToNexForExl(Product PR, int Pro)
        {
            InitializeComponent();
            if (PR != null)
            { _currentTovar = PR; num = Pro; }
            DataContext = _currentTovar;
            

            Func();
        
            
        }

        private void Back_Click(object sender, RoutedEventArgs e)
        {
            CreOtcToExl WW = new CreOtcToExl();
            WW.Show();
            this.Close();
        }

        private void Ext_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        public void Func()
        {
            var App = new Excel.Application();
            App.SheetsInNewWorkbook = 1;
            int StartIndex = 1;
            Excel.Workbook workbook = App.Workbooks.Add();
            Excel.Worksheet worksheet = App.Worksheets.Item[StartIndex];
            worksheet.Name = " Накладная ";
            worksheet.Cells[1][StartIndex] = " Завод ";
            Excel.Range Head3 = worksheet.Range[worksheet.Cells[2][StartIndex], worksheet.Cells[5][StartIndex]];
            Head3.Merge();
            StartIndex += 1;

            worksheet.Cells[1][StartIndex] = " Наименование ";
            Excel.Range Head4 = worksheet.Range[worksheet.Cells[2][StartIndex], worksheet.Cells[5][StartIndex]];
            Head4.Merge();
            worksheet.Cells[2][StartIndex] = _currentTovar.Name;
            StartIndex += 1;
            worksheet.Cells[1][StartIndex] = " Поступило штук ";
            Excel.Range Head5 = worksheet.Range[worksheet.Cells[2][StartIndex], worksheet.Cells[5][StartIndex]];
            Head5.Merge();
            StartIndex += 1;
            worksheet.Cells[1][StartIndex] = " П№ чертежа ";
            Excel.Range Head6 = worksheet.Range[worksheet.Cells[2][StartIndex], worksheet.Cells[5][StartIndex]];
            Head6.Merge();
            worksheet.Cells[2][StartIndex] = _currentTovar.NomCher;
            StartIndex += 1;
            worksheet.Cells[1][StartIndex] = " ВидПроверки ";
            worksheet.Cells[2][StartIndex] = " Норма ";
            worksheet.Cells[3][StartIndex] = " Факт ";
            worksheet.Cells[4][StartIndex] = " Проверено, шт ";
            worksheet.Cells[5][StartIndex] = " Несоотв., шт ";
            StartIndex += 1;
            if (_currentTovar.PosOtv !=null)
            {
                worksheet.Cells[1][StartIndex] = " Ø посадочного отверстия, мм ";
                worksheet.Cells[2][StartIndex] = _currentTovar.PosOtv;

                StartIndex += 1;

            }
            worksheet.Cells[1][StartIndex] = " Масса,г ";
            worksheet.Cells[2][StartIndex] = _currentTovar.Mass;
            StartIndex += 1;
            if( _currentTovar.BurtNar !=null)
            {
                worksheet.Cells[1][StartIndex] = " Ø бурта наружный, мм";
                worksheet.Cells[2][StartIndex] = _currentTovar.BurtNar;
                StartIndex += 1;
            }
            if (_currentTovar.Hei != null)
            {
                worksheet.Cells[1][StartIndex] = "Высота, мм";
                worksheet.Cells[2][StartIndex] = _currentTovar.Hei;
                StartIndex += 1;
            }
            if (_currentTovar.DiaNar != null)
            {
                worksheet.Cells[1][StartIndex] = " Ø наружный, мм";
                worksheet.Cells[2][StartIndex] = _currentTovar.DiaNar;
                StartIndex += 1;
            }
            if (_currentTovar.ProxOtv != null)
            {
                worksheet.Cells[1][StartIndex] = " Ø Проходное отвестие, мм";
                worksheet.Cells[2][StartIndex] = _currentTovar.ProxOtv;
                StartIndex += 1;
            }
            if (_currentTovar.TrubRez != null)
            {
                worksheet.Cells[1][StartIndex] = " Трубная резьба, мм";
                worksheet.Cells[2][StartIndex] = _currentTovar.TrubRez;
                StartIndex += 1;
            }
            if (_currentTovar.MetrRez != null)
            {
                worksheet.Cells[1][StartIndex] = "Метрическая резьба, мм";
                worksheet.Cells[2][StartIndex] = _currentTovar.MetrRez;
                StartIndex += 1;
            }
            if (_currentTovar.VnutXvos != null)
            {
                worksheet.Cells[1][StartIndex] = " Внутренний Ø хвостовика, мм";
                worksheet.Cells[2][StartIndex] = _currentTovar.VnutXvos;
                StartIndex += 1;
            }
            if (_currentTovar.NarXvos != null)
            {
                worksheet.Cells[1][StartIndex] = " Наружный  Ø хвостовика, мм";
                worksheet.Cells[2][StartIndex] = _currentTovar.NarXvos;
                StartIndex += 1;
            }
            if (_currentTovar.ProSRezSoe != null)
            {
                worksheet.Cells[1][StartIndex] = " Прочность резьбового соединения";
                int Nwe = StartIndex;
                StartIndex += 1;
                worksheet.Cells[1][StartIndex] = "";
                StartIndex += 1;
                worksheet.Cells[1][StartIndex] = "";
                int cla = StartIndex;
                Excel.Range Head = worksheet.Range[worksheet.Cells[1][Nwe], worksheet.Cells[1][cla]];
                Head.Merge();
                Excel.Range Head2 = worksheet.Range[worksheet.Cells[2][Nwe], worksheet.Cells[2][cla]];
                Head2.Merge();
                worksheet.Cells[3][Nwe] = "1)                                   . " ;
                worksheet.Cells[3][Nwe+1] = "2)                                   . ";
                worksheet.Cells[3][Nwe+2] = "3)                                   . ";
                StartIndex += 1;
            }
            if (_currentTovar.ScruRezFit != null)
            {
                worksheet.Cells[1][StartIndex] = " Скручивание резьбы фитинга";
                int Nwe = StartIndex;
                StartIndex += 1;
                worksheet.Cells[1][StartIndex] = "";
                StartIndex += 1;
                worksheet.Cells[1][StartIndex] = "";
                int cla = StartIndex;
                Excel.Range Head = worksheet.Range[worksheet.Cells[1][Nwe], worksheet.Cells[1][cla]];
                Head.Merge();
                Excel.Range Head2 = worksheet.Range[worksheet.Cells[2][Nwe], worksheet.Cells[2][cla]];
                Head2.Merge();
                worksheet.Cells[3][Nwe] = "1)                                   . ";
                worksheet.Cells[3][Nwe + 1] = "2)                                   . ";
                worksheet.Cells[3][Nwe + 2] = "3)                                   . ";
                StartIndex += 1;
            }
            if (_currentTovar.NarOblKonCNakGai != null)
            {
                worksheet.Cells[1][StartIndex] = " Нар.Ø в обл.контакта с накид.гайкой";
                worksheet.Cells[2][StartIndex] = _currentTovar.NarOblKonCNakGai;
                StartIndex += 1;
            }
            if (_currentTovar.VnuPrisZakSize5 != null)
            {
                worksheet.Cells[1][StartIndex] = "Внутренний Ø в месте присоединения к закладной (SIZE 5)";
                worksheet.Cells[2][StartIndex] = _currentTovar.VnuPrisZakSize5;
                StartIndex += 1;
            }
            if (_currentTovar.PosPodProclSize6 != null)
            {
                worksheet.Cells[1][StartIndex] = "Посадочный Ø под прокладку (SIZE 6)";
                worksheet.Cells[2][StartIndex] = _currentTovar.PosPodProclSize6;
                StartIndex += 1;
            }
            if (_currentTovar.NarDiaBurt != null)
            {
                worksheet.Cells[1][StartIndex] = "Наружный диаметр бурта";
                worksheet.Cells[2][StartIndex] = _currentTovar.NarDiaBurt;
                StartIndex += 1;
            }
            if (_currentTovar.VnutKonCOblSet != null)
            {
                worksheet.Cells[1][StartIndex] = "Внутренний диаметр в области контакта с сеткой, мм";
                worksheet.Cells[2][StartIndex] = _currentTovar.VnutKonCOblSet;
                StartIndex += 1;
            }
            if (_currentTovar.PramPodRuch != null)
            {
                worksheet.Cells[1][StartIndex] = "Прямоугольник под ручку, мм";
                worksheet.Cells[2][StartIndex] = _currentTovar.PramPodRuch;
                StartIndex += 1;
            }
            if (_currentTovar.PramPodShar != null)
            {
                worksheet.Cells[1][StartIndex] = "Прямоугольник под шар, мм";
                worksheet.Cells[2][StartIndex] = _currentTovar.PramPodShar;
                StartIndex += 1;
            }
            if (_currentTovar.PramPodStok != null)
            {
                worksheet.Cells[1][StartIndex] = "Прямоугольник под шток, мм";
                worksheet.Cells[2][StartIndex] = _currentTovar.PramPodStok;
                StartIndex += 1;
            }
            if (_currentTovar.RazPodKlu != null)
            {
                worksheet.Cells[1][StartIndex] = "Размер под ключ";
                worksheet.Cells[2][StartIndex] = _currentTovar.RazPodKlu;
                StartIndex += 1;
            }
            if (_currentTovar.RazPodStok != null)
            {
                worksheet.Cells[1][StartIndex] = "Размер под шток";
                worksheet.Cells[2][StartIndex] = _currentTovar.RazPodStok;
                StartIndex += 1;
            }
            if (_currentTovar.ProxOtvPodBurt != null)
            {
                worksheet.Cells[1][StartIndex] = "Ø проходн. отверстия под бурт , мм";
                worksheet.Cells[2][StartIndex] = _currentTovar.ProxOtvPodBurt;
                StartIndex += 1;
            }
            if (_currentTovar.VnuPrisZakSize6 != null)
            {
                worksheet.Cells[1][StartIndex] = " Внутренний Ø к закладной (SIZE 5)";
                worksheet.Cells[2][StartIndex] = _currentTovar.VnuPrisZakSize6;
                StartIndex += 1;
            }
            if (_currentTovar.PosOtvSoStorXvos != null)
            {
                worksheet.Cells[1][StartIndex] = " Посадочное отверстие со стороны хвостика";
                worksheet.Cells[2][StartIndex] = _currentTovar.PosOtvSoStorXvos;
                StartIndex += 1;
            }
            if (_currentTovar.PosPodProclSize5 != null)
            {
                worksheet.Cells[1][StartIndex] = " Посадочный Ø под прокладку (SIZE 5)";
                worksheet.Cells[2][StartIndex] = _currentTovar.PosPodProclSize5;
                StartIndex += 1;
            }
            if (_currentTovar.VisVOtkSost != null)
            {
                worksheet.Cells[1][StartIndex] = "Высота в открытом состоянии";
                worksheet.Cells[2][StartIndex] = _currentTovar.VisVOtkSost;
                StartIndex += 1;
            }
            if (_currentTovar.VisVZakSost != null)
            {
                worksheet.Cells[1][StartIndex] = "Высота в закрытом состоянии";
                worksheet.Cells[2][StartIndex] = _currentTovar.VisVZakSost;
                StartIndex += 1;
            }
            if (_currentTovar.VnutKZaklSize5 != null)
            {
                worksheet.Cells[1][StartIndex] = "Внутренний Ø к закладной (SIZE 5)";
                worksheet.Cells[2][StartIndex] = _currentTovar.VnutKZaklSize5;
                StartIndex += 1;
            }
            if (_currentTovar.VnutKZaklSize5 != null)
            {
                worksheet.Cells[1][StartIndex] = "Внутренний Ø к закладной (SIZE 6)";
                worksheet.Cells[2][StartIndex] = _currentTovar.VnutKZaklSize6;
                StartIndex += 1;
            }

            Excel.Range RR1 = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[5][StartIndex-1]];
            RR1.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                RR1.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle =
                RR1.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                RR1.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle =
                RR1.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle =
                RR1.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;


            worksheet.Columns.AutoFit();
            App.Visible = true;
            if (App.Visible == true)
            { this.Close(); }
      
        
        
        }
    }
}
