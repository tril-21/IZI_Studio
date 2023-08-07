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

namespace IZI_Studio
{
    /// <summary>
    /// Логика взаимодействия для Modules.xaml
    /// </summary>
    public partial class Modules : Window
    {
        MainWindow MainWin;
        string UserName;
        bool Admin;
        bool theDeadWorf = false;
        public Modules(MainWindow mw, string un, bool ad)
        {
            MainWin = mw;
            UserName = un;
            Admin = ad;
            InitializeComponent();
            disabled_Button();

        }
        //Отключение не работающих кнопок
        private void disabled_Button()
        {
            open_Email.IsEnabled = true;
            open_MailRF.IsEnabled = false;
            open_Apel.IsEnabled = false;
        }

        //Открытие модуля электронной почты
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            MainFormMail form_mail = new MainFormMail(MainWin, UserName, Admin);
            form_mail.Show();
            theDeadWorf = true;
            this.Close();
        }

        //Открытие модуля Почты РФ
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {

        }

        //Открытие модуля Апелляционных определений
        private void Button_Click_2(object sender, RoutedEventArgs e)
        {

        }

        private void Window_Closed(object sender, EventArgs e)
        {
            if (theDeadWorf == false)
                MainWin.Close();
        }
    }
}
