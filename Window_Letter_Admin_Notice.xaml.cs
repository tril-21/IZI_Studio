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
using System.Data.OleDb;
using System.IO;

namespace IZI_Studio
{
    /// <summary>
    /// Логика взаимодействия для Window_Letter_Admin_Notice.xaml
    /// </summary>
    /// 

    //Стурктура для отчета по отправлению
    public struct Notice_Letter
    {
        public string path_notice;
        public string status_send;
        public string erros;
    }
    public partial class Window_Letter_Admin_Notice : Window
    {
        public static string connectString_FOR_LETTER = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=DB_Letter.accdb;";
        public static string connectString_FOR_CASE = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=DB_Case.accdb;";
        public OleDbConnection myConnection;

        MainFormMail MainFM;
        Notice_Letter st_notice;

        //Конструктор данной формы
        public Window_Letter_Admin_Notice(MainFormMail mf)
        {
            MainFM = mf;
            InitializeComponent();
            text_notice.Text = "Не добавлено";
            text_errors.Text = "Нет";
            init_struct_Notice_Letter();
        }

        //Открытие окна для выбора файла отчета
        private void button_insert_notice_Click(object sender, RoutedEventArgs e)
        {
            text_notice.Text = "Не добавлено";
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
                text_notice.Text = "Добавлено";
                st_notice.path_notice = filename;
            }
        }

        //Изменение статуса об ошибках
        private void button_change_errors_Click(object sender, RoutedEventArgs e)
        {
            if (text_errors.Text == "Нет")
            {
                text_errors.Text = "Есть";
            }
            else
            {
                text_errors.Text = "Нет";
            }
            
        }

        //Копирование файла на сервер + обновление поля в Базе данных
        private void button_save_notice_Click(object sender, RoutedEventArgs e)
        {
            if (text_notice.Text == "Добавлено")
            {
                copy_file_in_server();
                st_notice.status_send = "Отправлено";
                st_notice.erros = text_errors.Text;
                myConnection = new OleDbConnection(connectString_FOR_LETTER);
                myConnection.Open();
                string query = "UPDATE [" + MainFM.selected_delo + "] SET [Путь_до_отчета]='" + st_notice.path_notice + "', [Статус_отправки]='" + st_notice.status_send + "', [Наличие_ошибок]='" + st_notice.erros + "' WHERE [Код]=" + MainFM.selected_letter;
                OleDbCommand command = new OleDbCommand(query, myConnection);
                command.ExecuteNonQuery();
                myConnection.Close();
                change_case_on_table_allcase();
                this.Close();
            }
            else
            {
                MessageBox.Show("Добавьте отчет!");
            }
        }

        //Событие при закрытии окна
        private void Window_Closed(object sender, EventArgs e)
        {
            if (check_send_letter())
            {
                MainFM.List_Letter_Loaded();
            }
            else
            {
                MainFM.update_for_button_up_case();
            }
            MainFM.Show();
        }

        //Изменение статуса дела в таблице всех дел, если  все письма по делу отправлены
        private void change_case_on_table_allcase()
        {
            if (!check_send_letter())
            {
                myConnection = new OleDbConnection(connectString_FOR_CASE);
                myConnection.Open();
                string query = "UPDATE [AllCase] SET [Статус_отправки]='Отправлено' WHERE [Номер_дела]='" + MainFM.selected_delo+"'";
                OleDbCommand command = new OleDbCommand(query, myConnection);
                command.ExecuteNonQuery();
                myConnection.Close();
            }
            
        }

        //Проверка на наличие не отправленных писем
        private bool check_send_letter()
        {
            myConnection = new OleDbConnection(connectString_FOR_LETTER);
            myConnection.Open();
            string query = "SELECT [Статус_отправки] FROM [" + MainFM.selected_delo + "] WHERE [Статус_отправки]='Не отправлено'";
            OleDbCommand command = new OleDbCommand(query, myConnection);
            OleDbDataReader reader = command.ExecuteReader();
            if(reader.HasRows)
            {
                myConnection.Close();
                return true;
            }
            else
            {
                myConnection.Close();
                return false;
            }
            
        }

        //Инициализация структуры
        private void init_struct_Notice_Letter()
        {
            st_notice.path_notice = "";
            st_notice.status_send = "";
            st_notice.erros = "";
        }

        //Функция для копирования файла на сервер
        public void copy_file_in_server()
        {
            string type_doc = "";
            if (st_notice.path_notice.EndsWith(".pdf"))
                type_doc = ".pdf";
            else if (st_notice.path_notice.EndsWith(".zip"))
                type_doc = ".zip";
            else if (st_notice.path_notice.EndsWith(".rar"))
                type_doc = ".rar";
            string new_name =MainFM.selected_letter + " " + MainFM.selected_delo + " " + "Notice" + type_doc;
            new_name = new_name.Replace("/", " ");
            string dir = @"\\negas\Obmen\IZI_STUDIO\MAIL\NOTICE\";
            if (!Directory.Exists(@"\\negas\Obmen\IZI_STUDIO\MAIL\NOTICE\")) // "!" забыл поставить
            {
                dir = @"\\NEGAS\Obmen\IZI_STUDIO\MAIL\NOTICE\";
            }
            string new_path = dir + new_name;
            File.Copy(st_notice.path_notice, new_path, true);
            st_notice.path_notice = new_path;
        }
    }
}
