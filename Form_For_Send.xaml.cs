using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.IO;

using System.Data.OleDb;
using System.Data;

namespace IZI_Studio
{
    /// <summary>
    /// Логика взаимодействия для Form_For_Send.xaml
    /// </summary>
    /// 

    //Структура для заполнения данными письма с последующей отпрвкой в базу данных
    struct Letter
    {
        public string fio;
        public string theme;
        public string date;
        public string other;
        public string email;
        public string mailrussia;
        public string path_pdf;
        public string path_word;
        public string path_otchet;
        public string status_send;
        public string error;
    }
    public partial class Form_For_Send : Window
    {
        public static string connectString_FOR_CASE = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=DB_Case.accdb;";
        public static string connectString_FOR_LETTER = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=DB_Letter.accdb;";
        private OleDbConnection myConnection;

        MainFormMail form_mail;
        string theme_letter = "";
        string selected_delo = "";
        string send_email = "no";
        string send_mail = "no";
        string path_pdf = "";
        string path_word = "";
        string path_otchet = "";

        //Констурктор данной формы
        public Form_For_Send(MainFormMail mM, string delo)
        {
            form_mail = mM;
            selected_delo = delo;
            InitializeComponent();
            label_selected_delo.Content = selected_delo;
        }

        //Комбобокс с выбором темы письма
        private void combobox_theme_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ComboBox comboBox = (ComboBox)sender;
            TextBlock selectedItem = (TextBlock)comboBox.SelectedItem;
            theme_letter = selectedItem.Text;
        }

        //Закрытие окна
        private void Window_Closed(object sender, EventArgs e)
        {
            form_mail.Show();
        }

        //Выбор файла ПДФ или сжатой зип-папки
        private void button_pdf_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.OpenFileDialog();
            dialog.FileName = "Document"; // Default file name
            dialog.DefaultExt = ".pdf"; // Default file extension
            dialog.Filter = "Документ PDF (.pdf)|*.pdf|Архив ZIP (.zip)|*.zip|Архив RAR (.rar)|*.rar"; // Filter files by extension

            // Show open file dialog box
            bool? result = dialog.ShowDialog();

            // Process open file dialog box results
            if (result == true)
            {
                // Open document
                string filename = dialog.FileName;
                button_pdf.Content = filename;
                path_pdf = filename;
            }
        }

        //Выбор файла ворд
        private void button_word_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.OpenFileDialog();
            dialog.FileName = "Document"; // Default file name
            dialog.DefaultExt = ".docx"; // Default file extension
            dialog.Filter = "Документ WORD (.docx; .doc)|*.docx;*.doc"; // Filter files by extension

            // Show open file dialog box
            bool? result = dialog.ShowDialog();

            // Process open file dialog box results
            if (result == true)
            {
                // Open document
                string filename = dialog.FileName;
                button_word.Content = filename;
                path_word = filename;
            }
        }

        //Выбор отправки по эл. почте
        private void check_email_Checked(object sender, RoutedEventArgs e)
        {
            /*if (check_email.IsChecked == true)
            {
                send_email = "yes";
            }else
            {
                send_email = "no";
            }*/
        }

        //Без добавления ворда
        private void check_mail_russia_Checked(object sender, RoutedEventArgs e)
        {
           
        }

        //Событие при нажатии кнопки отправииь (окончательное)
        private void button_creat_Click(object sender, RoutedEventArgs e)
        {
            //copy_file_in_server();

            if (check_validation() == true)
            {
                copy_file_in_server();

                Letter mail_letter = new Letter();

                //установка отправителя
                mail_letter.fio = form_mail.User_Name;

                //установка темы
                mail_letter.theme = theme_letter;

                //установка даты
                string[] subs = DateTime.Today.ToString().Split(' ');
                mail_letter.date = subs[0];

                //установка примечания
                mail_letter.other = text_other.Text;

                //установка отправки по эл. почты
                mail_letter.email = send_email;

                //установка отправки по почте россии
                mail_letter.mailrussia = send_mail;

                //установка пути до пдф
                mail_letter.path_pdf = path_pdf;

                //установка пути до ворд
                mail_letter.path_word = path_word;

                //установка пути до отчета
                mail_letter.path_otchet = path_otchet;

                //установка статуса отправки
                mail_letter.status_send = "Не отправлено";
                //установка наличия ошибок
                mail_letter.error = "Нет";

                insert_to_table_FOR_LETTER(mail_letter);

                update_case_in_ALL_CASE();

                MessageBox.Show("Письмо добавлено!");
                form_mail.List_Letter_Loaded();
                this.Close();
            }
        }

        //копирование файла на сервер
        public void copy_file_in_server()
        {
            string type_doc = "";

            if (path_pdf.EndsWith(".pdf"))
                type_doc = ".pdf";
            else if (path_pdf.EndsWith(".zip"))
                type_doc = ".zip";
            else if (path_pdf.EndsWith(".rar"))
                type_doc = ".rar";

            string new_name = select_count_letter().ToString() + " " + selected_delo + " " + theme_letter + type_doc;

            new_name = new_name.Replace("/", " ");

            //  MessageBox.Show(new_name);
            string dir = @"\\negas\Obmen\IZI_STUDIO\MAIL\LETTER\";
            if (!Directory.Exists(@"\\negas\Obmen\IZI_STUDIO\MAIL\LETTER\")) // "!" забыл поставить
            {
                dir = @"\\NEGAS\Obmen\IZI_STUDIO\MAIL\LETTER\";
            }
            string new_path = dir + new_name;

           // MessageBox.Show(new_path);
            File.Copy(path_pdf, new_path, true);

            path_pdf = new_path;


            if (check_mail_russia.IsChecked == false)
            {
                string new_name_word = select_count_letter().ToString() + " " + selected_delo + " " + theme_letter + ".docx";

                new_name_word = new_name_word.Replace("/", " ");

                dir = @"\\negas\Obmen\IZI_STUDIO\MAIL\LETTER\";
                if (!Directory.Exists(@"\\negas\Obmen\IZI_STUDIO\MAIL\LETTER\")) // "!" забыл поставить
                {
                    dir = @"\\NEGAS\Obmen\IZI_STUDIO\MAIL\LETTER\";
                }

                string new_path_word = dir + new_name_word;

                File.Copy(path_word, new_path_word, true);

                path_word = new_path_word;
            }
        }

        //запрос на количество писем по данному делу
        public int select_count_letter()
        {
            int count = 0;
            myConnection = new OleDbConnection(connectString_FOR_LETTER);

            // открываем соединение с БД
            myConnection.Open();


            string query = "SELECT * FROM [" + selected_delo + "]";

          //  MessageBox.Show(query);

            OleDbCommand command = new OleDbCommand(query, myConnection);

            using (OleDbDataReader reader = command.ExecuteReader())
            {
                DataTable dt = new DataTable();
                dt.Load(reader);
                count = dt.Rows.Count;
            }
            myConnection.Close();
            return count;
        }

        //проверка на заполненность формы
        public bool check_validation()
        {
            if(theme_letter == null)
            {
                MessageBox.Show("Выберите тему!");
                return false;
            }
            if (path_pdf == "")
            {
                MessageBox.Show("Выберите отправляемый файл!");
                return false;
            }
            if (path_word == "" && check_mail_russia.IsChecked==false)
            {
                MessageBox.Show("Выберите отправляемый файл типа WORD!");
                return false;
            }
            return true;
        }

        //Обновление статуса отправки в письме
        private void update_case_in_ALL_CASE()
        {
            myConnection = new OleDbConnection(connectString_FOR_CASE);

            // открываем соединение с БД
            myConnection.Open();

            string query = "UPDATE AllCase SET [Статус_отправки] = 'Не отправлено' WHERE [Номер_дела] = '"+selected_delo+"'";

            // создаем объект OleDbCommand для выполнения запроса к БД MS Access
            OleDbCommand command = new OleDbCommand(query, myConnection);

            // выполняем запрос к MS Access
            command.ExecuteNonQuery();

            myConnection.Close();
        }

        //Добавление данных в таблицу для писем
        private void insert_to_table_FOR_LETTER(Letter letter_m)
        {
            myConnection = new OleDbConnection(connectString_FOR_LETTER);

            // открываем соединение с БД
            myConnection.Open();

            // ДОБАВЛЕНИЕ В ТАБЛИЦУ Пользователя
            string query = "INSERT INTO [" + selected_delo + "] ([ФИО],[Тема],[Дата],[Примечание],[Почта_РФ],[Путь_до_файла],[Путь_до_ворд],[Путь_до_отчета],[Статус_отправки],[Наличие_ошибок]) VALUES ('"+letter_m.fio+"','" + letter_m.theme + "','"+letter_m.date+"','"+letter_m.other+"','"+letter_m.mailrussia+"','"+letter_m.path_pdf+"','"+letter_m.path_word+"','"+letter_m.path_otchet+"','"+letter_m.status_send+"','"+letter_m.error+"')";

            // создаем объект OleDbCommand для выполнения запроса к БД MS Access
            OleDbCommand command = new OleDbCommand(query, myConnection);

            // выполняем запрос к MS Access
            command.ExecuteNonQuery();

            myConnection.Close();
        }

        //Изменение чекбокса и кнопки добавить
        private void check_mail_russia_Click(object sender, RoutedEventArgs e)
        {
            if (check_mail_russia.IsChecked == true)
            {
                send_mail = "yes";
                button_word.IsEnabled = false;
            }
            else
            {
                send_mail = "no";
                button_word.IsEnabled = true;
            }
        }
    }
}
