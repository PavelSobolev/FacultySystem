using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace FSystem
{
    public partial class grupa_edit : Form
    {
        public grupa_edit()
        {
            InitializeComponent();
        }

        /// <summary>
        /// название группы
        /// </summary>      
        public string name = "";

        /// <summary>
        /// ид специалности
        /// </summary>
        public int spec_id = 0;

        /// <summary>
        /// название выпускающей кафедры
        /// </summary>
        public string kafname = "";

        /// <summary>
        /// номер курса группы
        /// </summary>
        public int nomer_kurs = 0;

        /// <summary>
        /// ид выпускающей кафедры
        /// </summary>
        public int kaf_id = 0;

        /// <summary>
        /// сщуствует ли группа (поле actual)
        /// </summary>
        public bool gr_exists = false;

        /// <summary>
        /// группа выводится в сетке
        /// </summary>
        public bool show_in_grid = false;

        /// <summary>
        /// идентификатор группы
        /// </summary>
        public int grupa_id = 0;

        /// <summary>
        /// группа существует
        /// </summary>
        public bool gr_actual = false;

        public string last = "";

        /// <summary>
        /// таблица специальностей
        /// </summary>
        public DataTable spec_set = new DataTable();             

        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private void grupa_edit_Load(object sender, EventArgs e)
        {
            string cmd = "select specialnost.id, specialnost.name, kafedra_id, prefix,  " + 
                " kafedra.name, srok, kod from specialnost " +
                " join kafedra on kafedra.id = specialnost.kafedra_id " +
                " where kafedra.fakultet_id = " + main.fakultet_id.ToString();

            SqlDataAdapter sda = new SqlDataAdapter(cmd, main.global_connection);
            spec_set = new DataTable();
            sda.Fill(spec_set);
            
            foreach (DataRow dr in spec_set.Rows)
            {
                spec_list.Items.Add(dr[6].ToString() + " -- " + dr[1].ToString());
            }

            if (spec_id == 0)            
                spec_list.SelectedIndex = 0;            
            else            
                spec_list.SelectedIndex = GetPosById(spec_set, spec_id);

            if (nomer_kurs != 0)
            {
                kurs_list.Maximum = (int)spec_set.Rows[spec_list.SelectedIndex][5];
                kurs_list.Value = nomer_kurs;
                grupa_name_box.Text = name;
                if (name.Length>0)
                last = name.Substring(name.Length - 1, 1);
            }

        }

        private int GetPosById(DataTable box, int id)
        {
            int i = 0;

            foreach (DataRow item in box.Rows)
            {
                if ((int)item[0] == id)
                {
                    return i;
                }
                i++;
            }

            return -1;
        }

        private void spec_list_SelectedIndexChanged(object sender, EventArgs e)
        {
            int index = spec_list.SelectedIndex;            
            
            spec_id = (int)spec_set.Rows[index][0];

            if (last != "")
                grupa_name_box.Text = spec_set.Rows[index][3].ToString() +
                    "-" + kurs_list.Value.ToString() + last;
            else
                grupa_name_box.Text = spec_set.Rows[index][3].ToString() +
                "-" + kurs_list.Value.ToString();
            
            kaf_id = (int)spec_set.Rows[index][2];
            kaf_box.Text = spec_set.Rows[index][4].ToString();

            int k = (int)spec_set.Rows[index][5];
            kurs_list.Maximum = k;
            srok_label.Text = "Срок обуч.:" + k.ToString() + " " + word(k);
        }

        private void kurs_list_ValueChanged(object sender, EventArgs e)
        {
            int index = spec_list.SelectedIndex;

            if (last != "")
                grupa_name_box.Text = spec_set.Rows[index][3].ToString() +
                    "-" + kurs_list.Value.ToString() + last;
            else
                grupa_name_box.Text = spec_set.Rows[index][3].ToString() +
                "-" + kurs_list.Value.ToString();
            
            nomer_kurs = (int)kurs_list.Value;
        }

        public string NormalizeLetters(string str)
        {
            foreach (char c in str)
            {

                bool can = Char.IsLetter(c) || Char.IsDigit(c) ||
                    c == '-' || c == ' ';

                if (!can)
                {
                    int pos = str.IndexOf(c);
                    str = str.Remove(pos, 1);
                    if (str.Length == 0) break;
                }
            }

            return str;
        }

        private void button1_Click(object sender, EventArgs e)
        {

            grupa_name_box.Text = NormalizeLetters(grupa_name_box.Text);
            
            if (grupa_name_box.Text.Trim()==string.Empty)
            {
                MessageBox.Show("Не введёно имя группы.",
                    "Ошибка данных", MessageBoxButtons.OK, MessageBoxIcon.Error);
                DialogResult = DialogResult.Abort;
                return;
            }

            string txt = grupa_name_box.Text;
            Normalize(ref txt);
            grupa_name_box.Text = txt;

            SqlCommand scmd = new SqlCommand("select count(*) from grupa where " +
                " name like '" + grupa_name_box.Text.Trim() + "' and " + 
                " id <> " + grupa_id.ToString(),
                main.global_connection);
            
            int res = (int)scmd.ExecuteScalar();

            if (res > 0)
            {
                MessageBox.Show("Введённое имя группы уже используется.",
                    "Ошибка данных", MessageBoxButtons.OK, MessageBoxIcon.Error);
                DialogResult = DialogResult.Abort;
                return;
            }

            if (!show_box.Checked)
            {
                DialogResult res1 = MessageBox.Show("Данная группа не будет выводиться в " + 
                    " таблице расписания.\n\nПродолжить?",
                    "Запрос на продолжение", MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);
                if (res1 == DialogResult.No)
                {
                    DialogResult = DialogResult.Abort;
                    return;
                }
            }

            if (!exists_box.Checked)
            {
                DialogResult res2 = MessageBox.Show("Данная группа не будет использоваться в " +
                    " программе (в дальнейшем ее можно будет восстановить).\n\nПродолжить?",
                    "Запрос на продолжение", MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);
                if (res2 == DialogResult.No)
                {
                    DialogResult = DialogResult.Abort;
                    return;
                }
            }

            if (!exists_box.Checked)
                show_box.Checked = false;

            DialogResult = DialogResult.OK;

        }

        /// <summary>
        /// удалить из строки метасимволы запросов '
        /// </summary>
        /// <param name="str"></param>
        public void Normalize(ref string str)
        {
            while (str.Contains("'"))
            {
                int pos = str.IndexOf("'");
                str = str.Remove(pos, 1);
                if (str.Length == 0) break;
            }
        }

        private void exists_box_CheckedChanged(object sender, EventArgs e)
        {
            gr_exists = exists_box.Checked;
        }

        private void show_box_CheckedChanged(object sender, EventArgs e)
        {
            show_in_grid = show_box.Checked;
        }

        public string word(int num)
        {
            string res = "";
            switch (num)
            {
                case 1: res = "год"; break;
                case 2:
                case 3:
                case 4: res = "года"; break;
                case 5:
                case 6:
                case 7: res = "лет"; break;
            }

            return res;
        }
     
   
    }
}