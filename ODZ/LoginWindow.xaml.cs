using System;
using System.Windows;
using ODZ;

namespace Aviadispetcher
{
    /// <summary>
    /// Логика взаимодействия для LoginWindow.xaml
    /// </summary>
    public partial class LoginWindow : Window
    {
        public LoginWindow()
        {
            InitializeComponent();
        }

        private int LogCheck()
        {
            int logUser = 0;
            if ((logTextBox.Text == "Користувач") &&
                (passwordTextBox.Text == "111"))
            {
                logUser = 1;
            }
            else if ((logTextBox.Text == "Редактор") &&
                     (passwordTextBox.Text == "222"))
            {
                logUser = 2;
            }
            else
            {
                MessageBox.Show("Введіть правильний пароль!", "Помилка!");
            }
            return logUser;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            BooksClass.logUser = LogCheck();
            if ((BooksClass.logUser == 1) ||
                (BooksClass.logUser == 2))
            {
                MainWindow mainWindow = new MainWindow();
                this.Close();
                mainWindow.Show();
            }

            logTextBox.Text = String.Empty;
            passwordTextBox.Text = String.Empty;
        }


    }
}
