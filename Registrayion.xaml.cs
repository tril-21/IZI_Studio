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
    /// Логика взаимодействия для Registrayion.xaml
    /// </summary>
    public partial class Registrayion : Window
    {
        //Настройки для подключения к БД
        public static string connectString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=DB_Users.accdb;";
        private OleDbConnection myConnection;

        MainWindow mainW;
        string NamePC;

        //Конструктор данной формы
        public Registrayion(MainWindow mW, string Name_PC)
        {
            NamePC = Name_PC;
            mainW = mW;
            InitializeComponent();
        }

        //Событие кнопки для регистрации пользователи в программе
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (textBox_FIO.Text.Length == 0)
            {
                MessageBox.Show("Пожалуйста, заполните поле!");
            }
            else if (textBox_Login.Text.Length<4 || textBox_Login.Text.Length>8)
            {
                MessageBox.Show("Длина логина должна быть от 4 до 8 символов. Введите заново.");
            } 
            else if (textBox_Password.Text.Length < 4 || textBox_Password.Text.Length > 8)
            {
                MessageBox.Show("Длина пароля должна быть от 4 до 8 символов. Введите заново.");
            }
            else
            {
                myConnection = new OleDbConnection(connectString);

                // открываем соединение с БД
                myConnection.Open();

                // текст запроса
                string query = "INSERT INTO [Users] (ФИО,Логин,Пароль,ИмяПК)" + "VALUES ('" + textBox_FIO.Text.ToString() + "','" + textBox_Login.Text.ToString() + "','" + textBox_Password.Text.ToString() + "','" + NamePC + "')";

                    // создаем объект OleDbCommand для выполнения запроса к БД MS Access
                OleDbCommand command = new OleDbCommand(query, myConnection);

                    // выполняем запрос к MS Access
                command.ExecuteNonQuery();
                  
                //  CreateAndSort();
                    

                myConnection.Close();

                create_Table_For_Mail(textBox_FIO.Text);
                textBox_FIO.Text = "";
                textBox_Login.Text = "";    
                textBox_Password.Text = "";

                MessageBox.Show("Пользователь зарегистрирован!");
                this.Close();
            }
        }

        //создает таблицу со списком дел для пользователя
        private void create_Table_For_Mail(string login)
        {
            string connectStringForMail = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=DB_Case.accdb;";
            myConnection = new OleDbConnection(connectStringForMail);
            myConnection.Open();
         //   string strTemp = "[Код] AUTOINCREMENT PRIMARY KEY, [Номердела] text";
            OleDbCommand myCommand = new OleDbCommand();
            myCommand.Connection = myConnection;
            myCommand.CommandText = "CREATE TABLE ["+login+"] ([Код] AUTOINCREMENT PRIMARY KEY, [Номер_дела] text)";
            //myCommand.CommandText = "CREATE TABLE [" + login + "] (" + strTemp + ")";
            myCommand.ExecuteNonQuery();
            myCommand.Connection.Close();
            myConnection.Close();
        }

        //Проверка на валидность введенных символов в Фамилия и инициалы
        private void TextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            int val;
            if (Int32.TryParse(e.Text, out val) || e.Text == "/"||e.Text==".")
            {
                e.Handled = true; // отклоняем ввод
            }
        }

        //Проверка на нажитие пробела
        private void TextBox_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Space)
            {
                e.Handled = true; // если пробел, отклоняем ввод
            }
        }

        //Выполняется при закрытии
        private void Closing_Closed(object sender, EventArgs e)
        {
            mainW.Show();
        }
    }
}
