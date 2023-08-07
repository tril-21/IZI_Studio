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
using System.Threading;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;


namespace IZI_Studio
{
    /// <summary>
    /// Логика взаимодействия для Window_Letter_Admin.xaml
    /// </summary>
    /// 

    //Структура для отправления админа
    public struct Letter_Admin
    {
        public string theme;
        public string e_mail;
        public string path_folder;
        public string path_file;
        public string other_letter;
    }

    //Структура для выбранного письма
    public struct Sellect_Letter
    {
        public string Theme;
        public string Path_PDF;
        public string Path_WORD;
        public string Other_Letter;
    }
    public partial class Window_Letter_Admin : Window
    {
        public static string connectString_FOR_LETTER = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=DB_Letter.accdb;";
        public OleDbConnection myConnection;

        MainFormMail MainFM;
        Letter_Admin st_letter;
        Sellect_Letter st_select_letter;

        bool isReady=false;

        //Конструктор данной формы
        public Window_Letter_Admin(MainFormMail mainFM)
        {
            MainFM = mainFM;
            InitializeComponent();
            init_struct_Letter_Admin();
            init_struct_Select_Letter();
            get_Letter_from_DataBase();
            start_filling();
            filling_email();
        }

        //копирует тему
        private void button_copy_theme_Click(object sender, RoutedEventArgs e)
        {
            Clipboard.Clear();
            Clipboard.SetText(text_theme.Text);
        }

        //копирует адреса эл. почты
        private void button_copy_email_Click(object sender, RoutedEventArgs e)
        {
            Clipboard.Clear();
            Clipboard.SetText(text_email.Text);
        }

        //копирует имя папки
        private void button_copy_name_folder_Click(object sender, RoutedEventArgs e)
        {
            Clipboard.Clear();
            Clipboard.SetText(text_name_folder.Text);
        }

        //копирует имя файла
        private void button_copy_name_file_Click(object sender, RoutedEventArgs e)
        {
            Clipboard.Clear();
            Clipboard.SetText(text_name_file.Text);
        }

        //копирует примечание
        private void button_copy_onther_Click(object sender, RoutedEventArgs e)
        {
            Clipboard.Clear();
            Clipboard.SetText(text_other.Text);
        }

        //запуск окна для отчета если нажата кнопка "готово"
        private void button_ready_Click(object sender, RoutedEventArgs e)
        {
            isReady = true;
            this.Close();
        }

        //закрытие окна, открытие основного
        private void Window_Closed(object sender, EventArgs e)
        {
            if (isReady)
            {
                MainFM.admin_open_letter_for_notice();
            }
            else
            {
                MainFM.Show();
            }
        }

        //инициализация элементов структуры, в которой элементы хранят данные для формы
        private void init_struct_Letter_Admin()
        {
            st_letter.theme = "";
            st_letter.e_mail = "";
            st_letter.path_folder = "";
            st_letter.path_file = "";
            st_letter.other_letter = "";
        }

        //инициализация элементов структуры, в которой элементы хранят данные из базы данных
        private void init_struct_Select_Letter()
        {
            st_select_letter.Theme = "";
            st_select_letter.Path_PDF = "";
            st_select_letter.Path_WORD = "";
            st_select_letter.Other_Letter = "";
        }

        //получение данных из базы для конкретного письма, заполнение структуры для базы данных
        private void get_Letter_from_DataBase()
        {
            //            string query = "SELECT [Код],[Тема],[Дата],[Примечание],[Статус_отправки],[Наличие_ошибок] FROM [" + MainFM.selected_delo + "]";

            myConnection = new OleDbConnection(connectString_FOR_LETTER);
            myConnection.Open();
            string query = "SELECT [Тема],[Путь_до_файла],[Путь_до_ворд],[Примечание] FROM [" + MainFM.selected_delo + "] WHERE [Код]=" + MainFM.selected_letter;
            OleDbCommand command = new OleDbCommand(query, myConnection);
            OleDbDataReader reader = command.ExecuteReader();
            reader.Read();
            st_select_letter.Theme = reader[0].ToString();
            st_select_letter.Path_PDF = reader[1].ToString();
            st_select_letter.Path_WORD = reader[2].ToString();
            st_select_letter.Other_Letter = reader[3].ToString();
            myConnection.Close();
        }

        //заполнение формы данными из структуры
        private void start_filling()
        {
            filling_theme();
            filling_name_folder_file();
            filling_other_letter();
            filling_email();

            text_theme.Text = st_letter.theme;
            text_email.Text = st_letter.e_mail;
            text_name_folder.Text = st_letter.path_folder;
            text_name_file.Text = st_letter.path_file;
            text_other.Text = st_letter.other_letter;
        }

        //формирование текста для темы
        private void filling_theme()
        {
            st_letter.theme = MainFM.selected_delo + " " + st_select_letter.Theme;
        }

        //формирование текста для эл. адресов
        private void filling_email()
        {
            if (st_select_letter.Path_WORD != "")
            {

                Word.Application app = new Word.Application();
                Object fileName = st_select_letter.Path_WORD;
                app.Documents.Open(ref fileName);
                Word.Document doc = app.ActiveDocument;
                // Нумерация параграфов начинается с одного
                for (int i = 1; i < doc.Paragraphs.Count; i++)
                {
                    string parText = doc.Paragraphs[i].Range.Text;

                    for (int j = 0; j < parText.Length; j++)
                    {
                        if (parText[j] == '@' && parText.Contains("5ap@sudrf.ru") != true)
                        {
                            parText = parText.TrimEnd('\r');
                            st_letter.e_mail = st_letter.e_mail + parText + ";";
                            //MessageBox.Show(parText+";");
                        }
                    }

                    // MessageBox.Show(parText);
                }

                app.Quit();

                /*
                WordprocessingDocument wordDoc = WordprocessingDocument.Open(st_select_letter.Path_WORD, true);
                int rowCount = 0;

                // Find the first table in the document.   
                DocumentFormat.OpenXml.Wordprocessing.Table table = wordDoc.MainDocumentPart.Document.Body.Elements<DocumentFormat.OpenXml.Wordprocessing.Table>().First();

                // To get all rows from table  
                IEnumerable<DocumentFormat.OpenXml.Wordprocessing.TableRow> rows = table.Elements<DocumentFormat.OpenXml.Wordprocessing.TableRow>();

                string strochka = "";

                // To read data from rows and to add records to the temporary table  
                foreach (DocumentFormat.OpenXml.Wordprocessing.TableRow row in rows)
                {
                    if (rowCount == 0)
                    {
                        foreach (DocumentFormat.OpenXml.Wordprocessing.TableCell cell in row.Descendants<DocumentFormat.OpenXml.Wordprocessing.TableCell>())
                        {
                            strochka = cell.InnerText;
                        }
                        rowCount += 1;
                    }
                }

                wordDoc.Close();

                //  strochka = Regex.Replace(strochka, @"[а-яА-ЯёЁ]", "");

                string[] words = strochka.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

                foreach (string s in words)
                {
                    for (int i = 0; i < s.Length; i++)
                    {
                        if (s[i] == '@')
                        {
                            st_letter.e_mail = st_letter.e_mail + s + ";";
                        }
                    }
                }
                */
            }
        }

        //формирование имени папки и имени файла
        private void filling_name_folder_file()
        {
            char ch = '\\';
            int index_char = st_select_letter.Path_PDF.LastIndexOf(ch);

            st_letter.path_file = st_select_letter.Path_PDF.Substring(index_char + 1);
            st_letter.path_folder = st_select_letter.Path_PDF.Substring(0, st_select_letter.Path_PDF.Length - (st_select_letter.Path_PDF.Length - index_char));
        }

        //формирование текста для примечания
        private void filling_other_letter()
        {
            st_letter.other_letter = st_select_letter.Other_Letter;
        }

    }
}
