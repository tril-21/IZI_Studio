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
    /// Логика взаимодействия для Window_Information_Dialog.xaml
    /// </summary>
    public partial class Window_Information_Dialog : Window
    {
        MainFormMail main_form_Mail;
        public Window_Information_Dialog(string text_info, MainFormMail mfm)
        {
            main_form_Mail = mfm;
            InitializeComponent();
            text_information.Content = text_info;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            main_form_Mail.isValid_WindowInfo = true;
            this.Close();
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            main_form_Mail.isValid_WindowInfo = false;
            this.Close();
        }
    }
}
