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
using System.IO;

using Excel = Microsoft.Office.Interop.Excel;

namespace CovidForm
{
    /// <summary>
    /// Логика взаимодействия для ContactWindow.xaml
    /// </summary>
    public partial class ContactWindow : Window
    {

        private string fileName = "";
        private List<ContactFace> items;

        public ContactWindow(string fileName, string fio_bolnogo)
        {
            InitializeComponent();
            items = new List<ContactFace>();
            lvUsers.ItemsSource = items;
            this.fileName = fileName;

            //ФИО больного
            Fio_bolnogo.Text = fio_bolnogo;
        }


        private void Button_Add_Contact(object sender, RoutedEventArgs e)
        {
            // Проверка чтобы хотя бы одно поле было заполнено
            if (isEmpty())
            {
                // Разница дат
                DateTime date_firtsvacc_contact_16_6month = date_firtsvacc_contact_16.SelectedDate ?? new DateTime(2020, 01, 01);
                DateTime date_secondvacc_contact_17_6month = date_secondvacc_contact_17.SelectedDate ?? new DateTime(2020, 01, 01);
                DateTime date_before_19_6month = date_before_19.SelectedDate ?? new DateTime(2020, 01, 01);
                
                //Вакцинация или переболел не позднее чем 6 месяцев назад
                if (DateTime.Compare(getAfterDate(date_firtsvacc_contact_16_6month, date_secondvacc_contact_17_6month, date_before_19_6month).AddMonths(6), DateTime.Today) > 0)
                {
                    MessageBox.Show("Да", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                else
                {
                    MessageBox.Show("Нет", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                }

                items.Add(new ContactFace()
                {
                    Id = items.Count + 1,
                    Notification = "да",
                    Name_contact_01 = name_contact_01.Text,
                    Floor_contact_02 = floor_contact_02.Text,
                    Date_birth_contact_03 = date_birth_contact_03.Text,
                    Address_contact_04 = address_contact_04.Text,
                    Place_work_contact_05 = place_work_contact_05.Text,
                    Contact_number_06 = contact_number_06.Text,
                    Date_sick_07 = date_sick_07.Text,
                    Date_end_isolation_08 = date_end_isolation_08.Text,
                    Name_sick_contact_09 = name_sick_contact_09.Text,
                    Num_decree_10 = num_decree_10.Text,
                    Date_decree_11 = date_decree_11.Text,
                    Sick_contact_12 = sick_contact_12.Text,
                    Self_observatory_13 = self_observatory_13.Text,
                    Med_organi_contact_14 = med_organi_contact_14.Text,
                    Vacc_name_contact_15 = vacc_name_contact_15.Text,
                    Date_firtsvacc_contact_16 = date_firtsvacc_contact_16.Text,
                    Date_secondvacc_contact_17 = date_secondvacc_contact_17.Text,
                    Revacc_contact_18 = revacc_contact_18.Text,
                    Date_before_19 = date_before_19.Text,
                });

                lvUsers.Items.Refresh();

                name_contact_01.Clear();
                floor_contact_02.SelectedItem = null;
                date_birth_contact_03.SelectedDate = null;
                address_contact_04.Clear();
                place_work_contact_05.Clear();
                contact_number_06.Clear();
                date_sick_07.SelectedDate = null;
                date_end_isolation_08.SelectedDate = null;
                name_sick_contact_09.Clear();
                num_decree_10.Clear();
                date_decree_11.SelectedDate = null;
                sick_contact_12.SelectedItem = null;
                self_observatory_13.SelectedItem = null;
                med_organi_contact_14.Clear();
                vacc_name_contact_15.Clear();
                date_firtsvacc_contact_16.SelectedDate = null;
                date_secondvacc_contact_17.SelectedDate = null;
                revacc_contact_18.Clear();
                date_before_19.SelectedDate = null;

                if (items.Count > 4)
                {
                    buttonAddContact.IsEnabled = false;
                }
            }
            else
            {
                MessageBox.Show("Заполните все обязательные * поля", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void Button_Export_Excel(object sender, RoutedEventArgs e)
        {
            var application = new Excel.Application(); //Объявляем переменную с приложением Excel
            Excel.Workbook workbook = application.Workbooks.Open(fileName);

            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets.get_Item(1);
            Excel.Worksheet worksheet_2 = (Excel.Worksheet)workbook.Worksheets.get_Item(3);

            var i = 3;
            foreach (var item in items)
            {
                worksheet_2.Cells[2][i] = item.Name_contact_01;
                worksheet_2.Cells[3][i] = item.Floor_contact_02;
                worksheet_2.Cells[4][i] = item.Date_birth_contact_03;
                worksheet_2.Cells[5][i] = item.Address_contact_04;
                worksheet_2.Cells[6][i] = item.Place_work_contact_05;
                worksheet_2.Cells[7][i] = item.Contact_number_06;
                worksheet_2.Cells[8][i] = item.Date_sick_07;
                worksheet_2.Cells[9][i] = item.Date_end_isolation_08;
                worksheet_2.Cells[10][i] = item.Name_sick_contact_09;
                worksheet_2.Cells[11][i] = item.Num_decree_10;
                worksheet_2.Cells[12][i] = item.Date_decree_11;
                worksheet_2.Cells[13][i] = item.Sick_contact_12;
                worksheet_2.Cells[14][i] = item.Self_observatory_13;
                worksheet_2.Cells[15][i] = item.Med_organi_contact_14;
                worksheet_2.Cells[16][i] = item.Vacc_name_contact_15;
                worksheet_2.Cells[17][i] = item.Date_firtsvacc_contact_16;
                worksheet_2.Cells[18][i] = item.Date_secondvacc_contact_17;
                worksheet_2.Cells[19][i] = item.Revacc_contact_18;
                worksheet_2.Cells[20][i] = item.Date_before_19;

                i++;
            }

            // Запись главного опрашиваемого в первую строку контактных для ориентировки
            worksheet_2.Cells[2][2] = worksheet.Cells[2][5];
            worksheet_2.Cells[3][2] = worksheet.Cells[2][12];
            worksheet_2.Cells[4][2] = worksheet.Cells[2][6];
            worksheet_2.Cells[5][2] = worksheet.Cells[2][9];
            worksheet_2.Cells[6][2] = worksheet.Cells[2][13];
            worksheet_2.Cells[7][2] = worksheet.Cells[2][10];
            worksheet_2.Cells[16][2] = worksheet.Cells[2][73];
            worksheet_2.Cells[17][2] = worksheet.Cells[2][74];
            worksheet_2.Cells[18][2] = worksheet.Cells[2][76];
            worksheet_2.Cells[19][2] = worksheet.Cells[2][80];

            try
            {
                workbook.Save();
                MessageBox.Show(workbook.Path, "Файл сохранен", MessageBoxButton.OK, MessageBoxImage.Information);
                System.Diagnostics.Process.Start("explorer.exe", workbook.Path); //Open folder in explorer
                workbook.Close(0);
                application.Quit();

                MainWindow mainWindow = new MainWindow();
                mainWindow.Show();
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка экспорта", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private bool isEmpty()
        {
            if (String.IsNullOrWhiteSpace(name_contact_01.Text) ||
                String.IsNullOrWhiteSpace(floor_contact_02.Text) ||
                String.IsNullOrWhiteSpace(date_birth_contact_03.Text) ||
                String.IsNullOrWhiteSpace(address_contact_04.Text) ||
                String.IsNullOrWhiteSpace(place_work_contact_05.Text) ||
                String.IsNullOrWhiteSpace(contact_number_06.Text) ||
                String.IsNullOrWhiteSpace(date_sick_07.Text) ||
                String.IsNullOrWhiteSpace(date_end_isolation_08.Text) ||
                String.IsNullOrWhiteSpace(name_sick_contact_09.Text) ||
                String.IsNullOrWhiteSpace(num_decree_10.Text) ||
                String.IsNullOrWhiteSpace(date_decree_11.Text) ||
                String.IsNullOrWhiteSpace(sick_contact_12.Text) ||
                String.IsNullOrWhiteSpace(self_observatory_13.Text)
                //String.IsNullOrWhiteSpace(med_organi_contact_14.Text) || // Не обязательное поле
                //String.IsNullOrWhiteSpace(vacc_name_contact_15.Text) || // Не обязательное поле
                //String.IsNullOrWhiteSpace(date_firtsvacc_contact_16.Text) || // Не обязательное поле
                //String.IsNullOrWhiteSpace(date_secondvacc_contact_17.Text) || // Не обязательное поле
                //String.IsNullOrWhiteSpace(revacc_contact_18.Text) || // Не обязательное поле
                //String.IsNullOrWhiteSpace(date_before_19.Text) // Не обязательное поле
                )
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private DateTime getAfterDate(DateTime t1, DateTime t2, DateTime t3)
        {
            DateTime result;
            if (DateTime.Compare(t1, t2) > 0)
            {
                if (DateTime.Compare(t1, t3) > 0)
                {
                    result = t1;
                }
                else
                {
                    result = t3;

                }
            }
            else if (DateTime.Compare(t2, t3) > 0)
            {
                result = t2;
            }
            else
            {
                result = t3;
            }
            return result;
        }
    } //end class

    public class ContactFace
    {
        public int Id { get; set; }
        public string Notification { get; set; }
        public string Name_contact_01 { get; set; }
        public string Floor_contact_02 { get; set; }
        public string Date_birth_contact_03 { get; set; }
        public string Address_contact_04 { get; set; }
        public string Place_work_contact_05 { get; set; }
        public string Contact_number_06 { get; set; }
        public string Date_sick_07 { get; set; }
        public string Date_end_isolation_08 { get; set; }
        public string Name_sick_contact_09 { get; set; }
        public string Num_decree_10 { get; set; }
        public string Date_decree_11 { get; set; }
        public string Sick_contact_12 { get; set; }
        public string Self_observatory_13 { get; set; }
        public string Med_organi_contact_14 { get; set; }
        public string Vacc_name_contact_15 { get; set; }
        public string Date_firtsvacc_contact_16 { get; set; }
        public string Date_secondvacc_contact_17 { get; set; }
        public string Revacc_contact_18 { get; set; }
        public string Date_before_19 { get; set; }
    }

}
