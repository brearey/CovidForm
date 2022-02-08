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
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        private string fileName = "";

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (false)
            {
                MessageBox.Show("Заполните все поля", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            else
            {
                var application = new Excel.Application(); //Объявляем переменную с приложением Excel
                Excel.Workbook workbook = application.Workbooks.Open("oprosnik.xlsx");
                //Excel.Workbook workbook = application.Workbooks.Open("C:/Users/admin/source/repos/CovidForm/123/CovidForm/oprosnik.xlsx");

                Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets.get_Item(1);
                Excel.Worksheet worksheet_2 = (Excel.Worksheet)workbook.Worksheets.get_Item(3);

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
                worksheet.Cells[2][32] = risk_28.Text;
                worksheet.Cells[2][33] = mask_29.Text;
                worksheet.Cells[2][34] = hands_30.Text;
                worksheet.Cells[2][35] = antiseptic_31.Text;
                worksheet.Cells[2][36] = gloves_32.Text;
                worksheet.Cells[2][38] = distance_33.Text;
                worksheet.Cells[2][39] = bus_34.Text;
                worksheet.Cells[2][40] = torg_35.Text;
                worksheet.Cells[2][41] = pitanie_36.Text;
                worksheet.Cells[2][42] = krasota_37.Text;
                worksheet.Cells[2][43] = med_38.Text;
                worksheet.Cells[2][44] = mass_naseleniya_39.Text;
                worksheet.Cells[2][45] = inpatient_care_40.Text;
                worksheet.Cells[2][46] = dormitory_41.Text;
                worksheet.Cells[2][47] = relatives_42.Text;
                worksheet.Cells[2][49] = private_43.Text;
                worksheet.Cells[2][50] = cause_44.Text;
                worksheet.Cells[2][51] = date_selection_45.Text;
                worksheet.Cells[2][52] = date_result_46.Text;
                worksheet.Cells[2][53] = date_confirmation_47.Text;
                worksheet.Cells[2][54] = organization_48.Text;
                worksheet.Cells[2][56] = other_infections_49.Text;
                worksheet.Cells[2][57] = symptoms_50.Text;
                worksheet.Cells[2][58] = temperature_51.Text;
                worksheet.Cells[2][59] = other_symptoms_52.Text;
                worksheet.Cells[2][60] = date_disease_53.Text;
                worksheet.Cells[2][61] = date_appeal_54.Text;
                worksheet.Cells[2][62] = place_appeal_55.Text;
                worksheet.Cells[2][63] = pleliminary_diagnosis_56.Text;
                worksheet.Cells[2][64] = place_estab_57.Text;
                worksheet.Cells[2][65] = hospitalzation_58.Text;
                worksheet.Cells[2][66] = date_hospitalzation_59.Text;
                worksheet.Cells[2][67] = place_hospitalzation_60.Text;
                worksheet.Cells[2][68] = diagnosis_61.Text;
                worksheet.Cells[2][69] = severity_62.Text;
                worksheet.Cells[2][70] = pathology_63.Text;
                worksheet.Cells[2][71] = pregnancy_64.Text;
                worksheet.Cells[2][72] = vaccinated_65.Text;
                worksheet.Cells[2][73] = name_vaccine_66.Text;
                worksheet.Cells[2][74] = date_onevacc_67.Text;
                worksheet.Cells[2][75] = day_illnessone_68.Text;
                worksheet.Cells[2][76] = date_twovacc_69.Text;
                worksheet.Cells[2][77] = day_illnesstwo_70.Text;
                worksheet.Cells[2][78] = flu_vaccinated_71.Text;
                worksheet.Cells[2][79] = revaccinated_72.Text;
                worksheet.Cells[2][80] = date_revacc_73.Text;
                worksheet.Cells[2][81] = final_diagnosis_74.Text;
                worksheet.Cells[2][82] = outcome_75.Text;
                worksheet.Cells[2][84] = date_recovery_76.Text;
                worksheet.Cells[2][85] = contact_77.Text;
                worksheet.Cells[2][86] = contact_category_78.Text;
                worksheet.Cells[2][87] = total_contact_79.Text;
                worksheet.Cells[2][88] = persons_covid_80.Text;
                worksheet.Cells[2][89] = household_81.Text;
                worksheet.Cells[2][90] = place_work_82.Text;
                worksheet.Cells[2][91] = place_education_83.Text;
                worksheet.Cells[2][92] = social_contact_84.Text;
                worksheet.Cells[2][93] = transport_contact_85.Text;
                worksheet.Cells[2][94] = other_organi_86.Text;
                worksheet.Cells[2][95] = medical_organi_87.Text;
                worksheet.Cells[2][96] = supervision_medical_88.Text;
                worksheet.Cells[2][97] = date_quarantine_89.Text;
                worksheet.Cells[2][98] = events_contact_90.Text;
                worksheet.Cells[2][99] = result_contact_91.Text;
                worksheet.Cells[2][100] = house_92.Text;
                worksheet.Cells[2][101] = entrance_93.Text;
                worksheet.Cells[2][102] = apartments_95.Text;
                try
                {
                    string folder = System.IO.Path.GetDirectoryName(System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName);
                    string my_folder = System.IO.Path.Combine(folder, "Excel");
                    System.IO.Directory.CreateDirectory(my_folder);
                    fileName = my_folder + "\\" + fio_03.Text.Replace(' ', '_') + System.DateTime.Now.ToString().Replace('.', '_').Replace(':', '_') + ".xlsx";
                    workbook.SaveCopyAs(fileName);
                    workbook.Close(0);
                    application.Quit();

                    //Open second window
                    ContactWindow contactWindow = new ContactWindow(fileName, fio_03.Text);
                    contactWindow.Show();
                    Close();
                }
                catch (Exception ex)
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
                String.IsNullOrWhiteSpace(distance_33.Text) ||
                String.IsNullOrWhiteSpace(bus_34.Text) ||
                String.IsNullOrWhiteSpace(bus_34.Text) ||
                String.IsNullOrWhiteSpace(torg_35.Text) ||
                String.IsNullOrWhiteSpace(pitanie_36.Text) ||
                String.IsNullOrWhiteSpace(krasota_37.Text) ||
                String.IsNullOrWhiteSpace(med_38.Text) ||
                String.IsNullOrWhiteSpace(mass_naseleniya_39.Text) ||
                String.IsNullOrWhiteSpace(inpatient_care_40.Text) ||
                String.IsNullOrWhiteSpace(dormitory_41.Text) ||
                String.IsNullOrWhiteSpace(relatives_42.Text) ||
                String.IsNullOrWhiteSpace(private_43.Text) ||
                String.IsNullOrWhiteSpace(cause_44.Text) ||
                String.IsNullOrWhiteSpace(date_selection_45.Text) ||
                String.IsNullOrWhiteSpace(date_result_46.Text) ||
                String.IsNullOrWhiteSpace(date_confirmation_47.Text) ||
                String.IsNullOrWhiteSpace(organization_48.Text) ||
                String.IsNullOrWhiteSpace(other_infections_49.Text) ||
                String.IsNullOrWhiteSpace(symptoms_50.Text) ||
                String.IsNullOrWhiteSpace(temperature_51.Text) ||
                String.IsNullOrWhiteSpace(other_symptoms_52.Text) ||
                String.IsNullOrWhiteSpace(date_disease_53.Text) ||
                String.IsNullOrWhiteSpace(date_appeal_54.Text) ||
                String.IsNullOrWhiteSpace(place_appeal_55.Text) ||
                String.IsNullOrWhiteSpace(pleliminary_diagnosis_56.Text) ||
                String.IsNullOrWhiteSpace(place_estab_57.Text) ||
                String.IsNullOrWhiteSpace(hospitalzation_58.Text) ||
                String.IsNullOrWhiteSpace(date_hospitalzation_59.Text) ||
                String.IsNullOrWhiteSpace(place_hospitalzation_60.Text) ||
                String.IsNullOrWhiteSpace(diagnosis_61.Text) ||
                String.IsNullOrWhiteSpace(severity_62.Text) ||
                String.IsNullOrWhiteSpace(pathology_63.Text) ||
                String.IsNullOrWhiteSpace(pregnancy_64.Text) ||
                String.IsNullOrWhiteSpace(vaccinated_65.Text) ||
                String.IsNullOrWhiteSpace(name_vaccine_66.Text) ||
                String.IsNullOrWhiteSpace(date_onevacc_67.Text) ||
                String.IsNullOrWhiteSpace(day_illnessone_68.Text) ||
                String.IsNullOrWhiteSpace(date_twovacc_69.Text) ||
                String.IsNullOrWhiteSpace(day_illnesstwo_70.Text) ||
                String.IsNullOrWhiteSpace(flu_vaccinated_71.Text) ||
                String.IsNullOrWhiteSpace(revaccinated_72.Text) ||
                String.IsNullOrWhiteSpace(date_revacc_73.Text) ||
                String.IsNullOrWhiteSpace(final_diagnosis_74.Text) ||
                String.IsNullOrWhiteSpace(outcome_75.Text) ||
                String.IsNullOrWhiteSpace(date_recovery_76.Text) ||
                String.IsNullOrWhiteSpace(contact_77.Text) ||
                String.IsNullOrWhiteSpace(contact_category_78.Text) ||
                String.IsNullOrWhiteSpace(total_contact_79.Text) ||
                String.IsNullOrWhiteSpace(persons_covid_80.Text) ||
                String.IsNullOrWhiteSpace(household_81.Text) ||
                String.IsNullOrWhiteSpace(place_work_82.Text) ||
                String.IsNullOrWhiteSpace(place_education_83.Text) ||
                String.IsNullOrWhiteSpace(social_contact_84.Text) ||
                String.IsNullOrWhiteSpace(transport_contact_85.Text) ||
                String.IsNullOrWhiteSpace(other_organi_86.Text) ||
                String.IsNullOrWhiteSpace(medical_organi_87.Text) ||
                String.IsNullOrWhiteSpace(supervision_medical_88.Text) ||
                String.IsNullOrWhiteSpace(date_quarantine_89.Text) ||
                String.IsNullOrWhiteSpace(events_contact_90.Text) ||
                String.IsNullOrWhiteSpace(result_contact_91.Text) ||
                String.IsNullOrWhiteSpace(house_92.Text) ||
                String.IsNullOrWhiteSpace(entrance_93.Text) ||
                String.IsNullOrWhiteSpace(floor_94.Text) ||
                String.IsNullOrWhiteSpace(apartments_95.Text)
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
