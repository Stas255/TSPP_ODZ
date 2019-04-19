using System;
using System.Collections.Generic;
using System.Windows;
using MySql.Data.MySqlClient;
using System.IO;

namespace ODZ
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string connStr;
        public List<BooksClass> fList = new List<BooksClass>();
        int bookNum;
        public MainWindow()
        {
            InitializeComponent();
            OpenDbFile();
        }

        private void OpenDbFile()
        {
            try
            {
                connStr = "Server = 127.0.0.1; Database = liberybooks; Uid = root; Pwd = ;";
                MySqlConnection conn = new MySqlConnection(connStr);
                MySqlCommand command = new MySqlCommand();
                string commandString = "SELECT * FROM books;";
                command.CommandText = commandString;
                command.Connection = conn;
                MySqlDataReader reader;
                command.Connection.Open();
                reader = command.ExecuteReader();
                int i = 0;
                while (reader.Read())
                {
                    fList.Add(new BooksClass((string)reader["idbook"], (string)reader["nameauthor"],
                        (string)reader["title"], (int)reader["date"], (string)reader["placing"]));
                    i += 1;
                }
                reader.Close();
                BooksListDG.ItemsSource = fList;

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message + char.ConvertFromUtf32(13) +
                                char.ConvertFromUtf32(13) + "Для завантаження файлу " +
                                "виконайте команду Файл-Завантажити", "Помилка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BooksListDG_MouseUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            bookNum = BooksListDG.SelectedIndex;
            BooksClass editerBook = BooksListDG.SelectedItem as BooksClass;
            try
            {
                TextBoxIDBook.Text = editerBook.IDBook;
                TextBoxDate.Text = editerBook.Date.ToString();
                TextBoxNameAuthor.Text = editerBook.NameAuthor;
                TextBoxPlacing.Text = editerBook.Placing;
                TextBoxTitle.Text = editerBook.Title;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + char.ConvertFromUtf32(13) + char.ConvertFromUtf32(13), "",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadDataMenuItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                BooksListDG.ItemsSource = null;
                fList.Clear();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + char.ConvertFromUtf32(13),
                    "Помилка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            OpenDbFile();
        }
    }
}
