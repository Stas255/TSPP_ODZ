using System;
using System.Collections.Generic;
using System.Windows;
using MySql.Data.MySqlClient;
using System.IO;
using System.Windows.Controls;

namespace ODZ
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private bool isCreate = false;


        private string connStr;
        public List<BooksClass> fList = new List<BooksClass>();
        public List<BooksClass> selectedList = new List<BooksClass>();
        int bookNum;
        private int NumberBook;
        bool bookAdd = false;
        string filePath;//папка із виконуваним файлом програми
        Microsoft.Office.Interop.Word.Application wordApp;
        Microsoft.Office.Interop.Word.Document wordDoc;
        public List<BooksClass> selectedNameList = new List<BooksClass>();
        public MainWindow()
        {
            InitializeComponent();
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
                    fList.Add(new BooksClass((int)reader["id"],(string)reader["idbook"], (string)reader["nameauthor"],
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

        private void EditDataMenuItem_Click(object sender, RoutedEventArgs e)
        {
            GroupBoxEdit.Visibility = Visibility.Visible;
            this.Height = BooksListDG.Margin.Top + BooksListDG.RenderSize.Height + 60 +
                    GroupBoxEdit.RenderSize.Height;
            Button1.Content = "Редагувати";
        }

        private void inforBookForm_Loaded(object sender, RoutedEventArgs e)
        {
            OpenDbFile();
            if (BooksClass.logUser == 1)
            {
                Menu.Items.Remove(Menu.Items[1]);
                Menu.Width = 95;
            }
            GroupBoxEdit.Visibility = Visibility.Hidden;
            GroupBoxNameAndAuthor.Visibility = Visibility.Hidden;
            GroupBoxYear.Visibility = Visibility.Hidden;
            this.Width = BooksListDG.Margin.Left + BooksListDG.RenderSize.Width + 60;
            this.Height = BooksListDG.Margin.Top + BooksListDG.RenderSize.Height + 60;
        }

        private void AddDataMenuItem_Click(object sender, RoutedEventArgs e)
        {
            GroupBoxEdit.Visibility = Visibility.Visible;
            this.Height = BooksListDG.Margin.Top + BooksListDG.RenderSize.Height + 60 +
                          GroupBoxEdit.RenderSize.Height;
            Button1.Content = "Додати";
            
            bookNum = fList.Count;
        }

        private void ChangeFlightListData(int num)
        {
            TimeSpan depTime;
            if (bookAdd)
            {
                fList.Add(new BooksClass(fList.Count,"", "", "", 0, ""));
                num = fList.Count - 1;
            }

                fList[num].Title = TextBoxTitle.Text;
                fList[num].Date = Convert.ToInt16(TextBoxDate.Text);
                fList[num].IDBook = TextBoxIDBook.Text;
                fList[num].NameAuthor = TextBoxNameAuthor.Text;
                fList[num].Placing = TextBoxPlacing.Text;



            BooksListDG.ItemsSource = null;
            BooksListDG.ItemsSource = fList;
            if (bookAdd)
            {
                try
                {
                    using (MySqlConnection conn = new MySqlConnection(connStr))
                    using (MySqlCommand cmd =
                        new MySqlCommand(
                            "INSERT INTO books (IDBook, NameAuthor, Title, Date, Placing) VALUES (?,?,?,?,?)",
                            conn))
                    {
                        cmd.Parameters.Add("@idbook", MySqlDbType.VarChar, 5).Value = TextBoxIDBook.Text;
                        cmd.Parameters.Add("@nameauthor", MySqlDbType.VarChar, 25).Value = TextBoxNameAuthor.Text;
                        cmd.Parameters.Add("@title", MySqlDbType.VarChar, 25).Value = TextBoxTitle.Text;
                        cmd.Parameters.Add("@date", MySqlDbType.Int16, 4).Value =
                            Convert.ToInt16(TextBoxDate.Text);
                        cmd.Parameters.Add("@placing", MySqlDbType.VarChar,7).Value = TextBoxPlacing.Text;
                        cmd.Parameters.Add("@id", MySqlDbType.Int16, 11).Value = fList[num].id;
                        conn.Open();
                        cmd.ExecuteNonQuery();
                    }
                }
                catch (Exception ex)
                {
                    string errMsg = "";
                    if (ex.Message == "Unable to connect to any of the specified MySQL hosts.")
                    {
                        errMsg = "Підключення веб-сервер MySQL та завантажте дані командою Файл-Завантажити";
                    }
                    else
                    {
                        errMsg = "Для завантаження даних виконайте команду Файл-Завантаджити";
                    }

                    MessageBox.Show(ex.Message + char.ConvertFromUtf32(13) + char.ConvertFromUtf32(13) + errMsg,
                        "Помика",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            else
            {
                try
                {
                    using (MySqlConnection conn = new MySqlConnection(connStr))
                    using (MySqlCommand cmd =
                        new MySqlCommand(
                            "UPDATE books SET idbook = ?, nameauthor = ?, title = ?, date = ?, placing = ?  WHERE id = ?",
                            conn))
                    {
                        cmd.Parameters.Add("@idbook", MySqlDbType.VarChar, 5).Value = TextBoxIDBook.Text;
                        cmd.Parameters.Add("@nameauthor", MySqlDbType.VarChar, 25).Value = TextBoxNameAuthor.Text;
                        cmd.Parameters.Add("@title", MySqlDbType.VarChar, 25).Value = TextBoxTitle.Text;
                        cmd.Parameters.Add("@date", MySqlDbType.Int16, 4).Value =
                            Convert.ToInt16(TextBoxDate.Text);
                        cmd.Parameters.Add("@placing", MySqlDbType.VarChar, 7).Value = TextBoxPlacing.Text;
                        cmd.Parameters.Add("@id", MySqlDbType.Int16, 11).Value = fList[num].id;
                        conn.Open();
                        cmd.ExecuteNonQuery();
                    }
                }
                catch (Exception ex)
                {
                    string errMsg = "";
                    if (ex.Message == "Unable to connect to any of the specified MySQL hosts.")
                    {
                        errMsg = "Підключення веб-сервер MySQL та завантажте дані командою Файл-Завантажити";
                    }
                    else
                    {
                        errMsg = "Для завантаження даних виконайте команду Файл-Завантаджити";
                    }

                    MessageBox.Show(ex.Message + char.ConvertFromUtf32(13) + char.ConvertFromUtf32(13) + errMsg,
                        "Помика",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }

        }


        private void Button1_Click(object sender, RoutedEventArgs e)
        {
            if (Button1.Content == "Додати")
            {
                bookAdd = true;
                ChangeFlightListData(bookNum);
                LoadDataMenuItem_Click(sender, e);
            }
            else
            {
                bookAdd = false;
                ChangeFlightListData(bookNum);
            }
        }

        private List<BooksClass> Selected(string avthor = "", string title = "")
        {
            List<BooksClass> selectedList = new List<BooksClass>();
            ListBoxPlacing.Items.Clear();
            title = Convert.ToString(ComboBoxTitle.Items[ComboBoxTitle.SelectedIndex]);
            avthor = Convert.ToString(ComboBoxAutor.Items[ComboBoxAutor.SelectedIndex]);
            int j = 0;
            for (int i = 0; i < fList.Count; i++) //???
            {
                if (avthor == fList[i].NameAuthor && title == fList[i].Title)
                {
                    selectedList.Add(fList[i]);
                    j++;
                }
            }
            return selectedList;
        }
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            string selectedAutor = "";
            string selectedTitle = "";
            selectedAutor = Convert.ToString(ComboBoxAutor.Items[ComboBoxAutor.SelectedIndex]);
            selectedTitle = Convert.ToString(ComboBoxTitle.Items[ComboBoxTitle.SelectedIndex]);
            selectedList = Selected(selectedAutor, selectedTitle);//????
            selectedNameList.Clear();
            for (int i = 0; i < selectedList.Count; i++)
            {
                if (selectedList[i] != null)
                {
                    ListBoxPlacing.Items.Add(selectedList[i].Placing + " ");
                    selectedNameList.Add(selectedList[i]);
                }
            }
        }
        private void FillAutorList()
        {
            bool nameExist = false;
            ComboBoxAutor.Items.Add(fList[0].NameAuthor);

            for (int i = 1; i < fList.Count; i++) //ошибка
            {
                for (int j = 0; j < ComboBoxAutor.Items.Count; j++)
                {
                    if (ComboBoxAutor.Items[j].ToString() == fList[i].NameAuthor)
                    {
                        nameExist = true;
                    }
                }

                if (!nameExist)
                {
                    ComboBoxAutor.Items.Add(fList[i].NameAuthor);
                }

                nameExist = false;
            }
        }
        private void FillTitleist()
        {
            bool nameExist = false;
            ComboBoxTitle.Items.Clear();
            for (int i = 0; i < fList.Count; i++) //ошибка
            {
                for (int j = 0; j < ComboBoxTitle.Items.Count; j++)
                {
                    if (ComboBoxTitle.Items[j].ToString() == fList[i].Title)
                    {
                        nameExist = true;
                    }
                }

                if (!nameExist && fList[i].NameAuthor == ComboBoxAutor.SelectionBoxItem.ToString())
                {
                    ComboBoxTitle.Items.Add(fList[i].Title);
                }

                nameExist = false;
            }
        }
        private void SelectXXMenuItem_Click(object sender, RoutedEventArgs e)
        {
            if (GroupBoxYear.Visibility != Visibility.Visible)
            {
                GroupBoxYear.Visibility = Visibility.Visible;
                AddMargin(GroupBoxYear);
                this.Width += GroupBoxYear.Width + 20;
                ComboBoxAutor.Items.Clear();
            }
        }
        private void SelectXYMenuItem_Click(object sender, RoutedEventArgs e)
        {
            if (GroupBoxNameAndAuthor.Visibility != Visibility.Visible)
            {
                GroupBoxNameAndAuthor.Visibility = Visibility.Visible;
                AddMargin(GroupBoxNameAndAuthor);
                this.Width += GroupBoxNameAndAuthor.Width + 20;
                ComboBoxAutor.Items.Clear();
                FillAutorList();
            }
        }

        private void ComboBoxTitle_DropDownOpened(object sender, EventArgs e)
        {
            if (ComboBoxAutor.SelectionBoxItem.ToString() != string.Empty)
            {
                FillTitleist();
            }
            else
            {
                MessageBox.Show("Виберать дані автора" + char.ConvertFromUtf32(13),
                    "Помилка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }


        private void ComboBoxAutor_DropDownClosed(object sender, EventArgs e)
        {
            ComboBoxTitle.Items.Clear();
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            int timebook = Convert.ToInt32(TextBoxYear.Text);

            ListBoxNumber.Items.Clear();
            NumberBook = 0;

            for (int i = 0; i < fList.Count; i++)
            {
                if (fList[i].Date == timebook)
                {
                    NumberBook++;
                }
            }

            if (NumberBook != null)
                {
                    ListBoxNumber.Items.Add(" Кількість книг " + NumberBook);
                }
        }
        private void WriteData(List<BooksClass> selXYList)
        {
            try
            {
                filePath = Environment.CurrentDirectory.ToString();
                try
                {
                    wordApp = new Microsoft.Office.Interop.Word.Application();
                    wordDoc = wordApp.Documents.Add(filePath + "\\Шаблон_Пошуку_Книг.dotx");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message + char.ConvertFromUtf32(13) +
                                    "Недостатньо даних!" + char.ConvertFromUtf32(13) +
                                    "Помістіть файл Шаблон_Пошуку_книгів.dot" + char.ConvertFromUtf32(13) +
                                    "у каталог із exe-файлом програми і повторіть збереження", "Помилка",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }

                if (ComboBoxAutor != null)
                {
                    string selectedAutor = ComboBoxAutor.SelectedItem.ToString();
                    string selectedTitle = ComboBoxTitle.SelectedItem.ToString();
                    ReplaceText(selectedAutor, "[X]");
                    ReplaceText(selectedTitle, "[Y]");
                }

                string selectedYear = TextBoxYear.Text;
                string selectedQuantity = NumberBook.ToString();
                ReplaceText(selXYList, 1);
                ReplaceText(selectedYear, "[Z]");
                ReplaceText(selectedQuantity, "[W]");
                wordDoc.Save();
                if (wordDoc != null)
                {
                    wordDoc.Close();
                }

                if (wordApp != null)
                {
                    wordApp.Quit();
                }
            }
            catch (Exception e) { }
        }
        private void ReplaceText(List<BooksClass> selectedLixt, int numTable)
        {
            for (int i = 0; i < selectedLixt.Count; i++)
            {
                if (selectedLixt[i] != null)
                {
                    if (numTable == 1)
                    {
                        wordDoc.Tables[numTable].Rows.Add();
                        wordDoc.Tables[numTable].Cell(2 + i, 1).Range.Text =
                            selectedLixt[i].Title;
                        wordDoc.Tables[numTable].Cell(2 + i, 2).Range.Text =
                            selectedLixt[i].NameAuthor;
                        wordDoc.Tables[numTable].Cell(2 + i, 3).Range.Text =
                            selectedLixt[i].Placing;
                    }
                    else
                    {
                        wordDoc.Tables[numTable].Rows.Add();
                        wordDoc.Tables[numTable].Cell(1 + i, 1).Range.Text =
                            " Кількість книг " + NumberBook;
                    }
                }
            }

        }
        private void ReplaceText(string textToReplace, string replacedText)
        {
            Object missing = Type.Missing;

            Microsoft.Office.Interop.Word.Range selText;
            selText = wordDoc.Range(wordDoc.Content.Start, wordDoc.Content.End);

            Microsoft.Office.Interop.Word.Find find = wordApp.Selection.Find;
            find.Text = replacedText;
            find.Replacement.Text = textToReplace;
            Object wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue;
            Object replace = Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll;

            find.Execute(FindText: Type.Missing,
                MatchCase: false,
                MatchWholeWord: false,
                MatchWildcards: false,
                MatchSoundsLike: missing,
                MatchAllWordForms: false,
                Forward: true,
                Wrap: wrap,
                Format: false,
                ReplaceWith: missing, Replace: replace);
        }

        private void ButtonSave_Click(object sender, RoutedEventArgs e)
        {
            WriteData(selectedNameList);
        }

        void AddMargin(GroupBox groupBox)
        {
            if (!isCreate)
            {
                groupBox.Margin = new Thickness(385, 10, 0, 0);
                isCreate = true;
            }
            else
            {
                groupBox.Margin = new Thickness(600, 10, 0, 0);
            }
        }

        private void ComboBoxAutor_DropDownOpened(object sender, EventArgs e)
        {
            FillAutorList();
        }

        private void SaveDataMenuItem_Click(object sender, RoutedEventArgs e)
        {
            ButtonSave_Click(sender, e);
        }
    }
}
