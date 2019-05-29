using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace FSystem
{
    public partial class pass_edit : Form
    {
        public pass_edit()
        {
            InitializeComponent();
        }


        public string wrong = "'#?";
        public string pss = "";
        private int c = 1;

        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private string encript(string ps)
        {
            return ps;
        }

        private string decript(string ps)
        {
            return ps;
        }

        public bool correct_ps()
        {
            if (pss.Contains("'")) return false;
            if (pss.Contains("?")) return false;
            if (pss.Contains("#")) return false;
            if (pss.Contains(" ")) return false;

            return true;
        }

        public void getpass()
        {
            SqlCommand cmd = new SqlCommand("select pass from prepod where id = @ID",
                main.global_connection);
            cmd.Parameters.Add("@ID", SqlDbType.Int).Value = main.active_user_id;

            SqlDataReader rd = null;

            try
            {
                rd = cmd.ExecuteReader();
            }
            catch(Exception exx)
            {
                MessageBox.Show("Невозможно получить пароль для текущего пользователя.\n" + 
                    "Повторите операцию позднее.",
                    "Ошибка чтения данных", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                DialogResult = DialogResult.Cancel;
            }

            bool res = rd.Read();

            if (!res)
            {
                MessageBox.Show("Невозможно получить пароль для текущего пользователя.\n" +
                    "Повторите операцию позднее.",
                    "Ошибка чтения данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                DialogResult = DialogResult.Cancel;
            }

            pss = rd[0].ToString();
            rd.Close();
        }

        private void pass_edit_Load(object sender, EventArgs e)
        {
            label4.Text = InputLanguage.CurrentInputLanguage.Culture.IetfLanguageTag.Substring(0, 2);
            getpass();

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (c >= 3)
            {
                DialogResult = DialogResult.Cancel;                
            }

            if (textBox1.Text.Trim() == string.Empty && pss!=string.Empty)
            {
                MessageBox.Show("Не введен текущий пароль.\n" +
                    "Повторите ввод.", "Ошибка ввода",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                textBox1.Select();
                textBox1.Focus();
                return;
            }

            if (textBox2.Text.Trim() == string.Empty)
            {
                MessageBox.Show("Не введен новый пароль.\n" +
                    "Повторите ввод.", "Ошибка ввода",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                textBox2.Select();
                textBox2.Focus();
                return;
            }

            if (textBox1.Text.Trim() == string.Empty && pss != string.Empty)
            {
                MessageBox.Show("Не введен повторный пароль.\n" +
                    "Повторите ввод.", "Ошибка ввода",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                textBox3.Select();
                textBox3.Focus();
                return;
            }

            if (textBox1.Text.Trim() == string.Empty && pss != string.Empty)
            {
                MessageBox.Show("Не введен текущий пароль.\n" +
                    "Повторите ввод.", "Ошибка ввода",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                c++;
                textBox1.Select();
                textBox1.Focus();
                return;
            }

            string old = textBox1.Text.Trim();
            string new1 = textBox2.Text.Trim();
            string new2 = textBox3.Text.Trim();

            if (old != pss)
            {
                MessageBox.Show("Введен неправильный текущий пароль.\n" +
                    "Повторите ввод.", "Ошибка ввода",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                c++;
                textBox1.Select();
                textBox1.Focus();
                return;
            }


            String newpss1 = textBox2.Text;

            if (newpss1 != textBox3.Text)
            {
                MessageBox.Show("Повторение пароля в жёлтом поле не совпдает с основным паролем в розовом поле.\n" +
                    "Повторите ввод.",
                    "Ошибка ввода",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                textBox3.Clear();
                textBox3.Select();
                textBox3.Focus();
                return;
            }

            pss = newpss1;

            if (!correct_ps())
            {
                MessageBox.Show("Введенный пароль содержит недопустимые символы ( ' или  #  или  ?).\n" +
                    "Повторите ввод.",
                    "Ошибка ввода",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;                
            }

            save_pss();

            MessageBox.Show("Ваш пароль был изменен.", "Сообщение", 
                MessageBoxButtons.OK, MessageBoxIcon.Information);

            DialogResult = DialogResult.OK;
        }

        public void save_pss()
        {
            SqlCommand cmd = new SqlCommand("update prepod set pass = @PSS where id = @ID",
                   main.global_connection);
            cmd.Parameters.Add("@ID", SqlDbType.Int).Value = main.active_user_id;
            cmd.Parameters.Add("@PSS", SqlDbType.NVarChar).Value = pss;

            cmd.ExecuteNonQuery();

            try
            {
                ;
            }
            catch (Exception exx)
            {
                MessageBox.Show("Невозможно получить сохранить пароль для текущего пользователя.\n" +
                    "Повторите операцию позднее.",
                    "Ошибка чтения данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                DialogResult = DialogResult.Cancel;
            }            
   
        }

        private void pass_edit_InputLanguageChanged(object sender, InputLanguageChangedEventArgs e)
        {
            label4.Text = InputLanguage.CurrentInputLanguage.Culture.IetfLanguageTag.Substring(0, 2);
        }

    }
}