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

using Excel = Microsoft.Office.Interop.Excel;

namespace CovidForm
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
            if (!isNotEmpty())
            {
                MessageBox.Show("Заполните все поля", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            else
            {
                var application = new Excel.Application(); //Объявляем переменную с приложением Excel
                Excel.Workbook workbook = application.Workbooks.Open("oprosnik.xlsx");

                Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets.get_Item(1);

                //Запись в поля
                worksheet.Cells[2][3] = info_date_01.Text; //B3
                worksheet.Cells[2][4] = number_and_date_02.Text;
                worksheet.Cells[2][5] = fio_03.Text;
                worksheet.Cells[2][6] = birthday_date_04.Text;
                worksheet.Cells[2][7] = age_05.Text;
                worksheet.Cells[2][8] = age_group_06.Text;
                worksheet.Cells[2][9] = address_home_07.Text;
                worksheet.Cells[2][10] = mobile_phone_08.Text;
                worksheet.Cells[2][11] = citizen_09.Text;
                worksheet.Cells[2][12] = sex_10.Text;
                worksheet.Cells[2][13] = social_status_11.Text;
                worksheet.Cells[2][14] = infection_site_12.Text;
                worksheet.Cells[2][15] = diagnosis_13.Text;
                worksheet.Cells[2][16] = disease_severity_14.Text;
                worksheet.Cells[2][17] = illness_date_15.Text;
                worksheet.Cells[2][18] = application_date_16.Text;
                worksheet.Cells[2][19] = hospital_17.Text;
                worksheet.Cells[2][20] = hospital_date_18.Text;
                worksheet.Cells[2][21] = work_info_19.Text;
                worksheet.Cells[2][22] = date_last_visit_20.Text;
                worksheet.Cells[2][23] = epidem_history_21.Text;
                worksheet.Cells[2][24] = date_before_covid_22.Text;
                worksheet.Cells[2][25] = outside_vilage_23.Text;
                worksheet.Cells[2][26] = outside_subject_24.Text;
                worksheet.Cells[2][27] = outside_country_25.Text;
                worksheet.Cells[2][28] = date_return_26.Text;
                worksheet.Cells[2][29] = city_district_27.Text;
                worksheet.Cells[2][31] = risk_28.Text;
                worksheet.Cells[2][32] = mask_29.Text;
                worksheet.Cells[2][33] = hands_30.Text;
                worksheet.Cells[2][34] = antiseptic_31.Text;
                worksheet.Cells[2][35] = gloves_32.Text;
                worksheet.Cells[2][36] = distance_33.Text;

                try
                {
                    string folder = System.IO.Path.GetDirectoryName(System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName);
                    string path = folder + "\\"+ fio_03.Text.Replace(' ', '_') + System.DateTime.Now.ToString().Replace('.', '_').Replace(':', '_') + ".xlsx";
                    workbook.SaveCopyAs(path);
                    MessageBox.Show(path, "Файл сохранен", MessageBoxButton.OK, MessageBoxImage.Information);
                    System.Diagnostics.Process.Start("explorer.exe", @folder); //Open folder in explorer
                    workbook.Close(0);
                    application.Quit();
                }
                catch(Exception ex)
                {
                    MessageBox.Show(ex.Message, "Ошибка экспорта", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private bool isNotEmpty()
        {
            if (String.IsNullOrWhiteSpace(info_date_01.Text) ||
                String.IsNullOrWhiteSpace(number_and_date_02.Text) ||
                String.IsNullOrWhiteSpace(fio_03.Text) ||
                String.IsNullOrWhiteSpace(birthday_date_04.Text) ||
                String.IsNullOrWhiteSpace(age_05.Text) ||
                String.IsNullOrWhiteSpace(age_group_06.Text) ||
                String.IsNullOrWhiteSpace(address_home_07.Text) ||
                String.IsNullOrWhiteSpace(mobile_phone_08.Text) ||
                String.IsNullOrWhiteSpace(citizen_09.Text) ||
                String.IsNullOrWhiteSpace(sex_10.Text) ||
                String.IsNullOrWhiteSpace(social_status_11.Text) ||
                String.IsNullOrWhiteSpace(infection_site_12.Text) ||
                String.IsNullOrWhiteSpace(diagnosis_13.Text) ||
                String.IsNullOrWhiteSpace(disease_severity_14.Text) ||
                String.IsNullOrWhiteSpace(illness_date_15.Text) ||
                String.IsNullOrWhiteSpace(application_date_16.Text) ||
                String.IsNullOrWhiteSpace(hospital_17.Text) ||
                String.IsNullOrWhiteSpace(hospital_date_18.Text) ||
                String.IsNullOrWhiteSpace(work_info_19.Text) ||
                String.IsNullOrWhiteSpace(date_last_visit_20.Text) ||
                String.IsNullOrWhiteSpace(epidem_history_21.Text) ||
                String.IsNullOrWhiteSpace(date_before_covid_22.Text) ||
                String.IsNullOrWhiteSpace(outside_vilage_23.Text) || 
                String.IsNullOrWhiteSpace(outside_country_25.Text) || 
                String.IsNullOrWhiteSpace(date_return_26.Text) || 
                String.IsNullOrWhiteSpace(risk_28.Text) || 
                String.IsNullOrWhiteSpace(mask_29.Text) || 
                String.IsNullOrWhiteSpace(hands_30.Text) || 
                String.IsNullOrWhiteSpace(antiseptic_31.Text) || 
                String.IsNullOrWhiteSpace(gloves_32.Text) || 
                String.IsNullOrWhiteSpace(distance_33.Text)
                )
            {
                return true;
            }
            else
            {
                return false;
            }
        }
    }
}
