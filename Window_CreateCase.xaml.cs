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
using System.Data;

namespace IZI_Studio
{
    /// <summary>
    /// Логика взаимодействия для Window_CreateCase.xaml
    /// </summary>
    public partial class Window_CreateCase : Window
    {
        public static string connectString_FOR_CASE = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=DB_Case.accdb;";
        public static string connectString_FOR_LETTER = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=DB_Letter.accdb;";
        private OleDbConnection myConnection;

        MainFormMail main_for_mail;
        string User_Name="";

        string index = "";
        string year = "";

        //Конструктор данной формы
        public Window_CreateCase(MainFormMail mfm, string un)
        {
            main_for_mail = mfm;
            User_Name = un;
            InitializeComponent();
        }

        //Событие кнопки на добавление дела
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (index == "")
            {
                MessageBox.Show("Выберите индекс коллегии!");
            }else if(number_case.Text.Length==0 && number_case.IsEnabled!=false){
                MessageBox.Show("Введите порядковый номер дела!");
            }else if (year == "")
            {
                MessageBox.Show("Выберите год!");
            }
            else
            {
                string true_number_case = "";
                if (index != "66" && index != "66а" && index != "55")
                {
                    true_number_case = index + "/" + year;
                }
                else
                {
                    true_number_case = index + "-" + number_case.Text + "/" + year;
                }
                
                if (check_case(true_number_case,"AllCase"))
                {
                    if (check_case(true_number_case, User_Name))
                    {
                        MessageBox.Show("Дело уже находится в вашем списке!");
                        this.Close();
                    }
                    else
                    {
                        insert_to_table_FOR_CASE_UserCase(true_number_case);
                        MessageBox.Show("Дело добавлено!");
                        this.Close();
                    }
                }
                else
                {
                    insert_to_table_FOR_CASE_AllCase(true_number_case);
                    insert_to_table_FOR_CASE_UserCase(true_number_case);
                    create_table_FOR_LETTER(true_number_case);
                    MessageBox.Show("Дело добавлено!");
                    this.Close();
                }
            }
        }

        //Событие закрытия формы
        private void Window_Closed(object sender, EventArgs e)
        {
            main_for_mail.Show();
        }

        //Комбобокс выбора индекса дела/отправления
        private void combobox_index_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ComboBox comboBox = (ComboBox)sender;
            TextBlock selectedItem = (TextBlock)comboBox.SelectedItem;
            index=selectedItem.Text;
            if (index != "66" && index != "66а" && index != "55")
            {
                number_case.IsEnabled = false;
            }
            else
            {
                number_case.IsEnabled = true;
            }
        }

        //Комбобокс выбора года дела
        private void combobox_year_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ComboBox comboBox = (ComboBox)sender;
            TextBlock selectedItem = (TextBlock)comboBox.SelectedItem;
            year=selectedItem.Text;
        }

        //Ограничение на ввод в поле номера дела
        private void number_case_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            int val;
            if (!Int32.TryParse(e.Text, out val) || e.Text == "/" || e.Text == ".")
            {
                e.Handled = true; // отклоняем ввод
            }
        }

        //Ограничение на пробел
        private void number_case_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Space)
            {
                e.Handled = true; // если пробел, отклоняем ввод
            }
        }

        //Добавление в таблицы в БД всех дел
        private void insert_to_table_FOR_CASE_AllCase(string numb_case)
        {
            myConnection = new OleDbConnection(connectString_FOR_CASE);

            // открываем соединение с БД
            myConnection.Open();

            // ДОБАВЛЕНИЕ В ТАБЛИЦУ ALLCASE
            string query = "INSERT INTO AllCase ([Номер_дела], [Статус_отправки]) VALUES ('" + numb_case + "','Отправлено')";

            // создаем объект OleDbCommand для выполнения запроса к БД MS Access
            OleDbCommand command = new OleDbCommand(query, myConnection);

            // выполняем запрос к MS Access
            command.ExecuteNonQuery();

            myConnection.Close();
        }

        //Добавление в таблицы в БД пользоателя
        private void insert_to_table_FOR_CASE_UserCase(string numb_case)
        {
            myConnection = new OleDbConnection(connectString_FOR_CASE);

            // открываем соединение с БД
            myConnection.Open();

            // ДОБАВЛЕНИЕ В ТАБЛИЦУ Пользователя
            string query = "INSERT INTO [" + User_Name + "] ([Номер_дела]) VALUES ('" + numb_case + "')";

            // создаем объект OleDbCommand для выполнения запроса к БД MS Access
            OleDbCommand command = new OleDbCommand(query, myConnection);

            // выполняем запрос к MS Access
            command.ExecuteNonQuery();

            myConnection.Close();
        }

        //Создание таблицы в БД писем
        private void create_table_FOR_LETTER(string numb_case)
        {
            myConnection = new OleDbConnection(connectString_FOR_LETTER);
            myConnection.Open();
            OleDbCommand myCommand = new OleDbCommand();
            myCommand.Connection = myConnection;
            myCommand.CommandText = "CREATE TABLE [" + numb_case + "] ([Код] AUTOINCREMENT PRIMARY KEY,[ФИО] text, [Тема] text, [Дата] text, [Примечание] text, [Почта_РФ] text, [Путь_до_файла] text, [Путь_до_ворд] text, [Путь_до_отчета] text, [Статус_отправки] text, [Наличие_ошибок] text)";
            myCommand.ExecuteNonQuery();
            myCommand.Connection.Close();
            myConnection.Close();
        }

        //Проверка на наличие дела
        private bool check_case(string numb_case, string table_name)
        {
            myConnection = new OleDbConnection(connectString_FOR_CASE);

            // открываем соединение с БД
            myConnection.Open();


            string query = "SELECT [Номер_дела] FROM [" + table_name + "] WHERE [Номер_дела]='" + numb_case + "'";

            OleDbCommand command = new OleDbCommand(query, myConnection);

            OleDbDataReader reader = command.ExecuteReader();

            if (reader.HasRows)
            {
                myConnection.Close();
                return true;
                
            } else
            {
                return false;
            }
        }

        //Закрытие формы
        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            main_for_mail.Show();
            main_for_mail.Update_Data_Case();
        }
    }
}
