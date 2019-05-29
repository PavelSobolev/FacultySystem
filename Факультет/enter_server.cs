using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace FSystem
{
    public partial class enter_server : Form
    {
        public enter_server ( )
        {
            InitializeComponent ( );
            Text = txt;
        }

        public bool enter_server_result = false;
        public string txt = "";

        private void button1_Click ( object sender, EventArgs e )
        {
            //проверка
            if ( srv.Text == "" )
            {
                MessageBox.Show(
                    "Не введено имя сервера.\nОбязательно введите имя сервера и пароль.",
                    "Ошибка ввода данных", MessageBoxButtons.OK, MessageBoxIcon.Error);                
                return;
            }

            if ( pwd.Text == string.Empty )
            {
                MessageBox.Show(
                    "Не введён пароль доступа на сервер.\nОбязательно введите имя сервера и пароль.",
                    "Ошибка ввода данных", MessageBoxButtons.OK, MessageBoxIcon.Error);                
                return;
            }

            enter_server_result = true;
        }

        private void button2_Click ( object sender, EventArgs e )
        {
            enter_server_result = false;
        }

        private void enter_server_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                DialogResult = DialogResult.Cancel;
            }

            if (e.KeyCode == Keys.Return)
            {
                //button1.Focus();
                if (srv.Text == "")
                {
                    MessageBox.Show(
                        "Не введено имя сервера.\nОбязательно введите имя сервера и пароль.",
                        "Ошибка ввода данных", MessageBoxButtons.OK, MessageBoxIcon.Error);                    
                    return;
                }

                if (pwd.Text == string.Empty)
                {
                    MessageBox.Show(
                        "Не введён пароль доступа на сервер.\nОбязательно введите имя сервера и пароль.",
                        "Ошибка ввода данных", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                enter_server_result = true;
            }

        }

        private void enter_server_Load(object sender, EventArgs e)
        {
            label4.Text = InputLanguage.CurrentInputLanguage.Culture.IetfLanguageTag.Substring(0, 2);
            if (main.final)
            {                
                //srv.Text = "";
                //pwd.Text = "";
            }
        }

        private void enter_server_InputLanguageChanged(object sender, InputLanguageChangedEventArgs e)
        {
            label4.Text = InputLanguage.CurrentInputLanguage.Culture.IetfLanguageTag.Substring(0, 2);
        }

        private void label4_Click(object sender, EventArgs e)
        {
            // to be removed
        }
    }
}