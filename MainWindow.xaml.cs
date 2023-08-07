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
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Data.OleDb;

namespace IZI_Studio
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //Настройки для подключения к БД
        public static string connectString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=DB_Users.accdb;";
        private OleDbConnection myConnection;

        //Получаем имя пк
        string NamePC = System.Net.Dns.GetHostName();

        string Login = "";
        string Password = "";
        bool Check = false;
        string UserName = "";

        //Конструктор данной формы
        public MainWindow()
        {
            InitializeComponent();
            AutoInput();
        }

        //Автовход по имени компьютера
        public void AutoInput()
        {
            
            // создаем экземпляр класса OleDbConnection
            myConnection = new OleDbConnection(connectString);

            // открываем соединение с БД
            myConnection.Open();

            string query = "SELECT * FROM [Users] WHERE [ИмяПК]='" + NamePC + "'";

            OleDbCommand command = new OleDbCommand(query, myConnection);

            OleDbDataReader reader = command.ExecuteReader();

            reader.Read();
            if (reader.HasRows)
            {
                UserName = reader[1].ToString();
                if (reader[5].ToString() == "True")
                {
                    textBox_Login.Text = reader[2].ToString();
                    passwordBox_password.Password = reader[3].ToString();
                }
            }
            myConnection.Close();
        }

        //Открытие окна регистрации
        void ShowRegistration (string namepc)
        {
            Registrayion form_reg = new Registrayion(this, namepc);
            form_reg.Show();
            this.Hide();
        }

        //Открывает окно отправки писем
        void ShowMainFormMail(string UserName, bool admin)
        {
            /*Modules modul = new Modules(this, UserName, admin);
            modul.Show();
            this.Hide();*/
            
            MainFormMail form_mail = new MainFormMail(this, UserName, admin);
            form_mail.Show();
            this.Hide();
        }

        //Событие кнопки для открытие окна регистрации
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            ShowRegistration(NamePC);
        }

        //Кнопка для запоминания логина и пароля для автоввода
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            if (Check == false)
            {
                Check = true;
                Info_Data.Content = "Запоминание включено";
            }
            else
            {
                Check = false;
                Info_Data.Content = "Запоминание выключено";
            }
        }

        //Событие кнопки для проверки логина и пароля и дальнейшего входа
        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            Login = textBox_Login.Text;
            Password = passwordBox_password.Password.ToString();
            // создаем экземпляр класса OleDbConnection
            myConnection = new OleDbConnection(connectString);

            // открываем соединение с БД
            myConnection.Open();

            string query = "SELECT * FROM [Users] WHERE [Логин]='" + Login + "'";

            

            OleDbCommand command = new OleDbCommand(query, myConnection);

            OleDbDataReader reader = command.ExecuteReader();

            reader.Read();
            if (reader.HasRows)
            {

                
                if (reader[3].ToString() == Password && Check==true)
                {
                    ChangedFieldBD();
                    UserName = reader[1].ToString();
                    if (reader[6].ToString()=="True")
                        ShowMainFormMail(UserName, true);
                    else
                        ShowMainFormMail(UserName, false);
                } else if(reader[3].ToString() == Password && Check == false)
                {
                    UserName = reader[1].ToString();
                    if (reader[6].ToString() == "True")
                        ShowMainFormMail(UserName, true);
                    else
                        ShowMainFormMail(UserName, false);
                } else
                {
                    MessageBox.Show("Неверный Логин и/или Пароль.");
                }
            }
            myConnection.Close();
        }

        //Изменение данных в базе данных при нажатии кнопки Запомнить
        public void ChangedFieldBD()
        {
            myConnection = new OleDbConnection(connectString);

            myConnection.Open();

            string query = "UPDATE [Users] SET [Логин]='"+Login+"',[Пароль]='"+Password+ "',[ИмяПК]='" + NamePC + "',[Автоввод]='True' WHERE [Логин]='" + Login+"'";
            // создаем объект OleDbCommand для выполнения запроса к БД MS Access
            OleDbCommand command = new OleDbCommand(query, myConnection);

            // выполняем запрос к MS Access
            command.ExecuteNonQuery();

            myConnection.Close();
        }
    }
}
