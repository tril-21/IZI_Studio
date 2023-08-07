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
using System.IO;

namespace IZI_Studio
{
    //Класс для отображения списка дел в ДатаГрид
    class Data_ListDelo
    {
        public Data_ListDelo(string NumberDelo)
        {
            this.Дело = NumberDelo;
        }
        public string Дело { get; set; }
    }

    //Класс для отображения списка писем в ДатаГрид
    class Data_ListLetter
    {
        public Data_ListLetter( string id, string fio, string theme, string data, string other, string status_send, string error)
        {
            this.Код = id;
            this.ФИО = fio;
            this.Тема = theme;
            this.Дата = data;
            this.Примечание = other;
            this.Статус = status_send;
            this.Ошибки = error;
        }
        public string Код { get; set; }
        public string ФИО { get; set; }
        public string Тема { get; set; }
        public string Дата { get; set; }
        public string Примечание { get; set; }
        public string Статус { get; set; }
        public string Ошибки { get; set; }
    }

    public partial class MainFormMail : Window
    {
        public static string connectString_FOR_CASE = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=DB_Case.accdb;";
        public static string connectString_FOR_LETTER = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=DB_Letter.accdb;";
        public OleDbConnection myConnection;
        public bool isValid_WindowInfo = false; //переменная для удаления дела из списка

        public string User_Name = "";
        bool is_Admin = false;
        DataGrid datagrid_Case = new DataGrid();

        public string selected_delo = "";
        public string selected_letter = "";

        MainWindow MainW;

        //Конструктор данной формы
        public MainFormMail(MainWindow form, string username, bool admin)
        {
            MainW = form;
            User_Name = username;
            is_Admin = admin;
            InitializeComponent();
            timer_Update();
            Settings_Tool_Button();
        }

        //Настройка расположения и активности кнопок на форме
        public void Settings_Tool_Button()
        {
            Grid.SetRow(update_letter, 0);
            Grid.SetColumn(update_letter, 19);
            if (is_Admin)
            {
                //Клавиша создания отправления
                user_create.IsEnabled = false;
                user_create.Visibility = Visibility.Hidden;
                //Клавиша открытия отправления
                Grid.SetRow(admin_open_letter, 0);
                Grid.SetColumn(admin_open_letter, 0);
                //Клавиша открытия файла
                Grid.SetRow(user_open, 0);
                Grid.SetColumn(user_open, 1);
                //Клавиша скачивания файла
                Grid.SetRow(user_download, 0);
                Grid.SetColumn(user_download, 2);
                //Клавиша открытия отчета
                Grid.SetRow(open_notice, 0);
                Grid.SetColumn(open_notice, 3);
                //Клавиша открытия формы добавления отчета
                Grid.SetRow(admin_open_notice, 0);
                Grid.SetColumn(admin_open_notice, 4);
                //Загрузка списка дел для отправки
                List_Delo_Loaded_Not_Send_Admin();
            }
            else
            {
                //Клавиша открытия отправления
                admin_open_letter.IsEnabled = false;
                admin_open_letter.Visibility = Visibility.Hidden;
                //Клавиша открытия формы добавления отчета
                admin_open_notice.IsEnabled = false;
                admin_open_notice.Visibility = Visibility.Hidden;
                //Клавиша создания отправления
                Grid.SetRow(user_create, 0);
                Grid.SetColumn(user_create, 0);
                //Клавиша открытия файла
                Grid.SetRow(user_open, 0);
                Grid.SetColumn(user_open, 1);
                //Клавиша скачивания файла
                Grid.SetRow(user_download, 0);
                Grid.SetColumn(user_download, 2);
                //Клавиша открытия отчета
                Grid.SetRow(open_notice, 0);
                Grid.SetColumn(open_notice, 3);
                //Загрузка списка дел для отправки
                List_Delo_Loaded();
            }
        }

        //Закрытие окна
        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            MainW.Close();
        }
        
        //Загрузка списка дел при первом открытии окна (для пользователей)
        private void List_Delo_Loaded()/*object sender, RoutedEventArgs e*/
        {
            myConnection = new OleDbConnection(connectString_FOR_CASE);
            myConnection.Open();
            string query = "SELECT [Номер_дела] FROM ["+User_Name+"]";
            OleDbCommand command = new OleDbCommand(query, myConnection);
            OleDbDataReader reader = command.ExecuteReader();
            List<Data_ListDelo> result = new List<Data_ListDelo>();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    result.Add(new Data_ListDelo(reader[0].ToString()));
                }
                List_Delo.ItemsSource = result;
                myConnection.Close();
            }else
            {
                //MessageBox.Show("В вашем личном кабинете нет добавленных дел.");
                myConnection.Close();
            }
        }

        //Загрузка списка дел с неотправленными письмами для администратора
        public void List_Delo_Loaded_Not_Send_Admin()
        {
            myConnection = new OleDbConnection(connectString_FOR_CASE);
            myConnection.Open();
            string query = "SELECT [Номер_дела] FROM [AllCase] WHERE [Статус_отправки]='Не отправлено'";
            OleDbCommand command = new OleDbCommand(query, myConnection);
            OleDbDataReader reader = command.ExecuteReader();
            List<Data_ListDelo> result = new List<Data_ListDelo>();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    result.Add(new Data_ListDelo(reader[0].ToString()));
                }
                List_Delo.ItemsSource = result;
                myConnection.Close();
            }else
            {
               // MessageBox.Show("В вашем личном кабинете нет добавленных дел.");
                if (List_Delo.ItemsSource != null)
                {
                    List_Delo.ItemsSource = null;
                    List_Letter.ItemsSource = null;
                }
                myConnection.Close();
            }
        }

        //создание датагрид (не используется)
        private void CreateDateGrid()
        {
            DataGrid datagrid_Case = new DataGrid();
            ForDataGrid.Children.Add(datagrid_Case);
            Grid.SetRow(datagrid_Case, 0);
            Grid.SetColumn(datagrid_Case, 1);
        }

        //ПОЛУЧЕНИЕ ЗНАЧЕНИЯ ИЗ ЯЧЕЙКИ
        private void List_Delo_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (List_Delo.SelectedItem == null)
            {
            }
            else
            {
                int selectedColumn = List_Delo.CurrentCell.Column.DisplayIndex;
                var selectedCell = List_Delo.SelectedCells[selectedColumn];
                var cellContent = selectedCell.Column.GetCellContent(selectedCell.Item);
                selected_delo = (cellContent as TextBlock).Text;
                info_delo_letter(selected_delo, "");

                List_Letter_Loaded();
            }
            
            List_Delo.SelectedItem =null;
        }

        //Получение писем после выбора дела
        public void List_Letter_Loaded()
        {
            string query = "";
            if (is_Admin)
            {
                query = "SELECT [Код],[ФИО],[Тема],[Дата],[Примечание],[Статус_отправки],[Наличие_ошибок] FROM [" + selected_delo + "]";
            }
            else
            {
                query = "SELECT [Код],[ФИО],[Тема],[Дата],[Примечание],[Статус_отправки],[Наличие_ошибок] FROM [" + selected_delo + "]";
            }
            myConnection = new OleDbConnection(connectString_FOR_LETTER);
            myConnection.Open();
            OleDbCommand command = new OleDbCommand(query, myConnection);
            OleDbDataReader reader = command.ExecuteReader();
            List<Data_ListLetter> result = new List<Data_ListLetter>();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    result.Add(new Data_ListLetter(reader[0].ToString(), reader[1].ToString(), reader[2].ToString(), reader[3].ToString(), reader[4].ToString(), reader[5].ToString(), reader[6].ToString()));
                }
                List_Letter.ItemsSource = result;
                myConnection.Close();
            }
            else
            {
                MessageBox.Show("По данному делу нет писем.");
                List_Letter.ItemsSource = null;
                myConnection.Close();
            }
        }

        //открытие окна создания дела
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Window_CreateCase win_create_case = new Window_CreateCase(this, User_Name);
            win_create_case.Show();
            this.Hide();
        }
        
        //обновление списка дел
        public void Update_Data_Case()
        {
            myConnection = new OleDbConnection(connectString_FOR_CASE);
            myConnection.Open();
            string query = "SELECT [Номер_дела] FROM [" + User_Name + "]";
            OleDbCommand command = new OleDbCommand(query, myConnection);
            OleDbDataReader reader = command.ExecuteReader();
            List<Data_ListDelo> result = new List<Data_ListDelo>();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    result.Add(new Data_ListDelo(reader[0].ToString()));
                }
                List_Delo.ItemsSource = result;
                myConnection.Close();
            }
            else
            {
                result = null;
                List_Delo.ItemsSource = result;
                myConnection.Close();
            }
        }

        //удаление дела из списка пользователя
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            if (selected_delo != "")
            {
                Window_Information_Dialog w_info = new Window_Information_Dialog("Вы действительно хотите удалить дело № " + selected_delo+" ?", this);
                w_info.ShowDialog();
                if (isValid_WindowInfo)
                {
                    myConnection = new OleDbConnection(connectString_FOR_CASE);
                    myConnection.Open();
                    string query = "DELETE FROM [" + User_Name + "] WHERE [Номер_дела] = '" + selected_delo + "'";
                    OleDbCommand command = new OleDbCommand(query, myConnection);
                    command.ExecuteNonQuery();
                    myConnection.Close();
                    Update_Data_Case();
                    info_delo_letter("", "");
                    List_Letter.ItemsSource = null;
                    isValid_WindowInfo = false;
                }
            } else
            {
                MessageBox.Show("Выберете дело!");
            }
        }

        //открытие окна создания письма
        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            if (selected_delo != "")
            {
                Form_For_Send create_Letter = new Form_For_Send(this, selected_delo);
                create_Letter.Show();
                this.Hide();
            } else
            {
                MessageBox.Show("Пожалуйста выберите дело!");
            }
        }

        //открытие файла
        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            if (selected_letter != "")
            {
                start_File_PDF_ZIP("");
            } else
            {
                MessageBox.Show("Выберите отправление!");
            }
        }

        //Запуск файла
        private void start_File_PDF_ZIP(string path)
        {
            myConnection = new OleDbConnection(connectString_FOR_LETTER);
            myConnection.Open();
            string query = "SELECT [Путь_до_файла] FROM [" + selected_delo + "] WHERE [Код]=" + Convert.ToInt32(selected_letter);
            OleDbCommand command = new OleDbCommand(query, myConnection);
            OleDbDataReader reader = command.ExecuteReader();
            if (path != "")
            {
                System.Diagnostics.Process p = new System.Diagnostics.Process();
                p.StartInfo.FileName = path;
                p.Start();
            } else
            {
                while (reader.Read())
                {
                    System.Diagnostics.Process p = new System.Diagnostics.Process();
                    p.StartInfo.FileName = reader[0].ToString();
                    p.Start();
                }
            }
            myConnection.Close();
        } 

        //получение имени файла для пдф
        private string get_filename(string path_query)
        {
            string file_name = "";
            myConnection = new OleDbConnection(connectString_FOR_LETTER);
            myConnection.Open();
            string query = "SELECT ["+path_query+"] FROM [" + selected_delo + "] WHERE [Код]=" + Convert.ToInt32(selected_letter);
            OleDbCommand command = new OleDbCommand(query, myConnection);
            OleDbDataReader reader = command.ExecuteReader();
            while (reader.Read())
            {
                file_name = reader[0].ToString();
            }
            myConnection.Close();
            return file_name;
        }

        //событие при нажитии на таблицу писем
        private void List_Letter_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (List_Letter.SelectedItem == null)
            {
            }
            else
            {
                int selectedColumn = 0;
                var selectedCell = List_Letter.SelectedCells[selectedColumn];
                var cellContent = selectedCell.Column.GetCellContent(selectedCell.Item);
                selected_letter = (cellContent as TextBlock).Text;
                info_delo_letter(selected_delo, selected_letter);

            }
            List_Letter.SelectedItem = null;
        }

        //скачивание с сервера
        private void Button_Click_4(object sender, RoutedEventArgs e)
        {
            if (selected_letter != "")
            {
                string file_name = get_filename("Путь_до_файла");
                string new_path_filename = "";
                string type_doc = "";
                string[] name = file_name.Split(new char[] { '.' });
                if (file_name.EndsWith(".pdf"))
                    type_doc = ".pdf";
                else if (file_name.EndsWith(".zip"))
                    type_doc = ".zip";
                else if (file_name.EndsWith(".rar"))
                    type_doc = ".rar";
                Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();
                dlg.FileName = name[0];
                dlg.DefaultExt = type_doc; 
                dlg.Filter = "AllFiles (.*)|*.*"; 
                Nullable<bool> result = dlg.ShowDialog();
                if (result == true)
                {
                    new_path_filename = dlg.FileName;
                }
                File.Copy(file_name, new_path_filename, true);
                System.Diagnostics.Process p = new System.Diagnostics.Process();
                p.StartInfo.FileName = new_path_filename;
                p.Start();
            }
            else
            {
                MessageBox.Show("Выберите отправление!");
            }
        }

        //открытие окна с информацией о письме необходимой для отправки по электронной почте
        private void admin_open_letter_Click(object sender, RoutedEventArgs e)
        {
            if (selected_letter != "")
            {
                Window_Letter_Admin winADmletter = new Window_Letter_Admin(this);
                winADmletter.Show();
                this.Hide();
            }
            else
            {
                MessageBox.Show("Выьерите отправления!");
            }
        }

        //Открытие формы добавления отчета
        public void admin_open_letter_for_notice()
        {
            if (selected_letter != "")
            {
                Window_Letter_Admin_Notice adm_notice = new Window_Letter_Admin_Notice(this);
                adm_notice.Show();
                this.Hide();
            }
            else
            {
                MessageBox.Show("Выберите отправление!");
            }
        }

        //Открытие формы добавления отчета по нажатию кнопки
        private void admin_open_notice_Click(object sender, RoutedEventArgs e)
        {
            admin_open_letter_for_notice();
        }

        //Отрыктие файла отправления
        public void start_File_PDF_ZIP_notice(string path)
        {
            myConnection = new OleDbConnection(connectString_FOR_LETTER);
            myConnection.Open();
            string query = "SELECT [Путь_до_отчета] FROM [" + selected_delo + "] WHERE [Код]=" + Convert.ToInt32(selected_letter);
            OleDbCommand command = new OleDbCommand(query, myConnection);
            OleDbDataReader reader = command.ExecuteReader();
            if (path != "")
            {
                System.Diagnostics.Process p = new System.Diagnostics.Process();
                p.StartInfo.FileName = path;
                p.Start();
            }
            else
            {
                while (reader.Read())
                {
                    if (reader[0].ToString() != "")
                    {
                        System.Diagnostics.Process p = new System.Diagnostics.Process();
                        p.StartInfo.FileName = reader[0].ToString();
                        p.Start();
                    }
                    else
                    {
                        MessageBox.Show("Отчет еще не загружен");
                    }
                    
                }
            }
            myConnection.Close();
        }

        //Открытие файла отчета
        private void open_notice_Click(object sender, RoutedEventArgs e)
        {
            if (selected_letter != "")
            {
                start_File_PDF_ZIP_notice("");
            }
            else
            {
                MessageBox.Show("Выберите отправление!");
            }
        }

        //Обновление списка дел
        private void Button_Click_5(object sender, RoutedEventArgs e)
        {
            update_for_button_up_case();
        }

        //обновления списка дел (логика вынесена из кнопки для ввызова из других скриптов)
        public void update_for_button_up_case() {
            if (is_Admin)
            {
                List_Delo_Loaded_Not_Send_Admin();
                List_Letter.ItemsSource = null;
                info_delo_letter("", "");
            }
            else
            {
                List_Delo_Loaded();
                List_Letter.ItemsSource = null;
                info_delo_letter("", "");

            }
        }

        //таймер для обновления списка дел
        private void timer_Update()
        {
            System.Windows.Threading.DispatcherTimer timer = new System.Windows.Threading.DispatcherTimer();

            timer.Tick += new EventHandler(timerTick);
            timer.Interval = new TimeSpan(0, 0, 300);
            timer.Start();
        }

        //функция запуска обновления списка дел
        private void timerTick(object sender, EventArgs e)
        {
            if (is_Admin)
            {
                List_Delo_Loaded_Not_Send_Admin();
                if (selected_delo != "")
                    List_Letter_Loaded();
            }
            else
            {
                List_Delo_Loaded();
                if (selected_delo != "")
                    List_Letter_Loaded();
            }
            
        }

        //Обновление списка писем
        private void update_letter_Click(object sender, RoutedEventArgs e)
        {
            if (selected_delo != "")
                List_Letter_Loaded();
            else
                MessageBox.Show("Выберите дело!");
        }

        //Показ выбранного дела и выбранного письма
        private void info_delo_letter(string delo, string letter)
        {
            selected_delo = delo;
            selected_letter = letter;

            Selected_Delo_Info.Content = selected_delo;
            Selected_Letter_Info.Content = selected_letter;
        }

    }
}
