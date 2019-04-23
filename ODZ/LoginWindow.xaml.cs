﻿using System;
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
                mainWindow.Show();
            }
            this.Close();
        }


    }
}
