using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.Remoting.Metadata.W3cXsd2001;
using System.Security.Cryptography.X509Certificates;
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
using WpfDemonstrated.Models;

namespace WpfDemonstrated
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private BindingList<data_grid_1> data_Grid_1s;
        private BindingList<data_grid_2> data_Grid_2s;
        private BindingList<main_grid> _main_grid;
        public MainWindow()
        {
            InitializeComponent();
        }
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
        }
        private string Path;


        private void Readfile()
        {
            var files = Directory.GetFiles(Path, "*.*", SearchOption.AllDirectories);
            string word;
            _main_grid = new BindingList<main_grid>();
            data_Grid_1s = new BindingList<data_grid_1>();
            data_Grid_2s = new BindingList<data_grid_2>();
            foreach (string file in files)
            {
                using (StreamReader sr = new StreamReader(file))
                {
                    word = sr.ReadToEnd();
                }
                Match chec_kpke_cxema = Regex.Match(word, "active_cxema=\"(.*?)\"", RegexOptions.Singleline);
                if (chec_kpke_cxema.Groups[1].Value == "1")
                {
                    sheme_check_1(word);
                }
                if (chec_kpke_cxema.Groups[1].Value == "2")
                    sheme_check_2(word);
            }
            Mai.ItemsSource = _main_grid;
            return;
        }
        private void sheme_check_2(string word)
        {
            Match chec_kpke_cxema = Regex.Match(word, "TimeStart=\"(.*?)\" TimeStop=\"(.*?)\"(.*?)averaging_interval_time=\"(.*?)\"(.*?)active_cxema=\"(.*?)\"(.*?)TimeTek=\"(.*?)\" UAB=\"(.*?)\" UBC=\"(.*?)\" UCA=\"(.*?)\" IAB=\"(.*?)\" IBC=\"(.*?)\" ICA=\"(.*?)\" IA=\"(.*?)\" IB=\"(.*?)\" IC=\"(.*?)\" PO=\"(.*?)\" PP=\"(.*?)\" QO=\"(.*?)\" QP=\"(.*?)\" SO=\"(.*?)\" SP=\"(.*?)\" UO=\"(.*?)\" UP=\"(.*?)\" IO=\"(.*?)\" IP=\"(.*?)\" KO=\"(.*?)\" Freq=\"(.*?)\" sigmaUy=\"(.*?)\" sigmaUyAB=\"(.*?)\" sigmaUyBC=\"(.*?)\" sigmaUyCA=\"(.*?)\"", RegexOptions.Singleline);
            _main_grid.Add(new main_grid()
            {
                NameObjekt = "Обьект 2",
                TimeStart = Perevod(chec_kpke_cxema.Groups[1].Value),
                TimeStop = Perevod(chec_kpke_cxema.Groups[2].Value),
                amountTime = Tominuteandsec(chec_kpke_cxema.Groups[4].Value),
                ShemeCheck = chec_kpke_cxema.Groups[6].Value
            });
            data_Grid_2s.Add(new data_grid_2()
            {
                CreationDates = Perevod(chec_kpke_cxema.Groups[8].Value),
                UAB = chec_kpke_cxema.Groups[9].Value,
                UBC = chec_kpke_cxema.Groups[10].Value,
                UCA = chec_kpke_cxema.Groups[11].Value,
                IAB = chec_kpke_cxema.Groups[12].Value,
                IBC = chec_kpke_cxema.Groups[13].Value,
                ICA = chec_kpke_cxema.Groups[14].Value,
                IA = chec_kpke_cxema.Groups[15].Value,
                IB = chec_kpke_cxema.Groups[16].Value,
                IC = chec_kpke_cxema.Groups[17].Value,
                P0 = chec_kpke_cxema.Groups[18].Value,
                PP = chec_kpke_cxema.Groups[19].Value,
                Q0 = chec_kpke_cxema.Groups[20].Value,
                QP = chec_kpke_cxema.Groups[21].Value,
                S0 = chec_kpke_cxema.Groups[22].Value,
                SP = chec_kpke_cxema.Groups[23].Value,
                U0 = chec_kpke_cxema.Groups[24].Value,
                UP = chec_kpke_cxema.Groups[25].Value,
                I0 = chec_kpke_cxema.Groups[26].Value,
                IP = chec_kpke_cxema.Groups[27].Value,
                K0 = chec_kpke_cxema.Groups[28].Value,
                Freq = chec_kpke_cxema.Groups[29].Value,
                sigmaUy = chec_kpke_cxema.Groups[30].Value,
                sigmaUyAB = chec_kpke_cxema.Groups[31].Value,
                sigmaUyBC = chec_kpke_cxema.Groups[32].Value,
                sigmaUyCA = chec_kpke_cxema.Groups[33].Value
            });
            return;
        }
        private void sheme_check_1(string word)
        {
            Match chec_kpke_cxema = Regex.Match(word, "TimeStart=\"(.*?)\" TimeStop=\"(.*?)\"(.*?)averaging_interval_time=\"(.*?)\"(.*?)active_cxema=\"(.*?)\"(.*?)TimeTek=\"(.*?)\" UA=\"(.*?)\" IA=\"(.*?)\" PA=\"(.*?)\" QA=\"(.*?)\" SA=\"(.*?)\" Freq=\"(.*?)\" sigmaUy=\"(.*?)\"", RegexOptions.Singleline);
            _main_grid.Add(new main_grid()
            {
                NameObjekt = "Обьект 1",
                TimeStart = Perevod(chec_kpke_cxema.Groups[1].Value),
                TimeStop = Perevod(chec_kpke_cxema.Groups[2].Value),
                amountTime = Tominuteandsec(chec_kpke_cxema.Groups[4].Value),
                ShemeCheck = chec_kpke_cxema.Groups[6].Value
            });
            data_Grid_1s.Add(new data_grid_1()
            {
                CreationDatef = Perevod(chec_kpke_cxema.Groups[8].Value),
                UA = chec_kpke_cxema.Groups[9].Value,
                IA = chec_kpke_cxema.Groups[10].Value,
                PA = chec_kpke_cxema.Groups[11].Value,
                QA = chec_kpke_cxema.Groups[12].Value,
                SA = chec_kpke_cxema.Groups[13].Value,
                Freq = chec_kpke_cxema.Groups[14].Value,
                sigmaUy = chec_kpke_cxema.Groups[15].Value
            });
            return;
        }

        private DateTime Perevod(string currenttime)
        {
            var tick = TimeSpan.FromMilliseconds(long.Parse(currenttime));
            var date = new DateTime(1970, 1, 1) + tick;
            return date;
        }
        private string Tominuteandsec(string currenttime)
        {
            int sec = int.Parse(currenttime) / 60;
            int minute = sec / 60;
            if (minute == 0)
                return sec.ToString() + " cек";
            else return minute.ToString() + " мин " + sec.ToString() + " сек";
        }

        private void TextBox_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                TextBox objTextBox = (TextBox)sender;
                Path = objTextBox.Text;
                Readfile();
            }
        }

        private void Mai_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            main_grid grid = Mai.SelectedItem as main_grid;
            if (grid != null)
            {
                if (grid.ShemeCheck == "2")
                    Second_Check.ItemsSource = data_Grid_2s;
                else if (grid.ShemeCheck == "1")
                    first_Check.ItemsSource = data_Grid_1s;
            }
            else
                return;
        }
        private void SaveData()
        {
           string dataDir = @"C:\Users\Artur\Desktop\GGGGWP";
            bool IsExists = System.IO.Directory.Exists(dataDir);
            if (!IsExists)
                System.IO.Directory.CreateDirectory(dataDir);
            Workbook wb = new Workbook();
            Worksheet sheet = wb.Worksheets[0];
            Cells cells = sheet.Cells;
            Cell cell = cells["A1"];
            cell.PutValue("главное окно");
            if (_main_grid != null)
            {
                int g = 2;
                foreach (var bllist in _main_grid)
                {
                    cell = cells["A"+g.ToString()];
                    cell.PutValue(bllist.NameObjekt.ToString());
                    cell = cells["B" + g.ToString()];
                    cell.PutValue(bllist.ShemeCheck.ToString());
                    cell = cells["C" + g.ToString()];
                    cell.PutValue(bllist.TimeStart.ToString());
                    cell = cells["D" + g.ToString()];
                    cell.PutValue(bllist.TimeStop.ToString());
                    g++;
                }
            }
            else
            {
                MessageBox.Show("Нету данных");
                return;
            }
            cell = cells["J1"];
            cell.PutValue("1_Обьект");
            int i = 2;
            foreach (var bllist in data_Grid_1s)
            {
                cell = cells["J" + i.ToString()];
                cell.PutValue(bllist.CreationDatef.ToString());
                cell = cells["K" + i.ToString()];
                cell.PutValue(bllist.UA.ToString());
                cell = cells["L" + i.ToString()];
                cell.PutValue(bllist.IA.ToString());
                cell = cells["M" + i.ToString()];
                cell.PutValue(bllist.PA.ToString());
                cell = cells["N" + i.ToString()];
                cell.PutValue(bllist.QA.ToString());
                cell = cells["O" + i.ToString()];
                cell.PutValue(bllist.Freq.ToString());
                cell = cells["P" + i.ToString()];
                cell.PutValue(bllist.sigmaUy.ToString());
                i++;
            }
            i = 2;
            cell = cells["Q1"];
            cell.PutValue("2_Обьект");
            foreach (var bllist in data_Grid_2s)
            {
                cell = cells["Q" + i.ToString()];
                cell.PutValue(bllist.CreationDates.ToString());
                cell = cells["R" + i.ToString()];
                cell.PutValue(bllist.UAB.ToString());
                cell = cells["S" + i.ToString()];
                cell.PutValue(bllist.UBC.ToString());
                cell = cells["T" + i.ToString()];
                cell.PutValue(bllist.UCA.ToString());
                cell = cells["U" + i.ToString()];
                cell.PutValue(bllist.IAB.ToString());
                cell = cells["V" + i.ToString()];
                cell.PutValue(bllist.IBC.ToString());
                cell = cells["W" + i.ToString()];
                cell.PutValue(bllist.ICA.ToString());
                cell = cells["X" + i.ToString()];
                cell.PutValue(bllist.IA.ToString());
                cell = cells["Y" + i.ToString()];
                cell.PutValue(bllist.IB.ToString());
                cell = cells["Z" + i.ToString()];
                cell.PutValue(bllist.IC.ToString());
                cell = cells["AA" + i.ToString()];
                cell.PutValue(bllist.P0.ToString());
                cell = cells["AB" + i.ToString()];
                cell.PutValue(bllist.PP.ToString());
                cell = cells["AC" + i.ToString()];
                cell.PutValue(bllist.Q0.ToString());
                cell = cells["AD" + i.ToString()];
                cell.PutValue(bllist.QP.ToString());
                 cell = cells["AE" + i.ToString()];
                cell.PutValue(bllist.S0.ToString());
                cell = cells["AF" + i.ToString()];
                cell.PutValue(bllist.SP.ToString());
                cell = cells["AG" + i.ToString()];
                cell.PutValue(bllist.U0.ToString());
                cell = cells["AH" + i.ToString()];
                cell.PutValue(bllist.UP.ToString());
                cell = cells["AI" + i.ToString()];
                cell.PutValue(bllist.I0.ToString());
                cell = cells["AJ" + i.ToString()];
                cell.PutValue(bllist.IP.ToString());
                cell = cells["AK" + i.ToString()];
                cell.PutValue(bllist.K0.ToString());
                cell = cells["AL" + i.ToString()];
                cell.PutValue(bllist.Freq.ToString());
                cell = cells["AM" + i.ToString()];
                cell.PutValue(bllist.sigmaUy.ToString());
                cell = cells["AN" + i.ToString()];
                cell.PutValue(bllist.sigmaUyAB.ToString());
                cell = cells["AO" + i.ToString()];
                cell.PutValue(bllist.sigmaUyBC.ToString());
                cell = cells["AP" + i.ToString()];
                cell.PutValue(bllist.sigmaUyCA.ToString());
                i++;
            }
            Aspose.Cells.Tables.ListObject listObject = sheet.ListObjects[sheet.ListObjects.Add("A1", "F2", true)];
            listObject.TableStyleType = Aspose.Cells.Tables.TableStyleType.TableStyleMedium10;
            wb.Save(dataDir + "output.xlsx");
            MessageBox.Show("Записано");
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            SaveData();
        }
    }
}
 