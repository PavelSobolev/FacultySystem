using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace FSystem
{
    public partial class city_street_edit : Form
    {

        public DataTable streets;
        public string StreetID = string.Empty;
        public string StreetName = string.Empty;
        public string GorodID = string.Empty;
        public bool edit = false;
        /// <summary>
        /// создание окна редактирования улиц
        /// </summary>
        /// <param name="data">таблица со списком улиц</param>
        /// <param name="gor_id">ид населённого пункта</param>
        public city_street_edit(DataTable data, string gor_id)
        {
            InitializeComponent();
            streets = data;
            listBox1.Items.Clear();
            foreach (DataRow d in streets.Rows)
            {
                listBox1.Items.Add(d[1].ToString());
            }
            GorodID = gor_id;

            textBox2.AutoCompleteCustomSource.Clear();
            foreach (DataRow d in streets.Rows)
            {
                textBox2.AutoCompleteCustomSource.Add(d[1].ToString());
            }
        }

        public void FixResult()
        {
            int ind = listBox1.SelectedIndex;
            if (ind < 0)
            {
                StreetID = "";
                StreetName = "";
                DialogResult = DialogResult.Cancel;
            }
            else
            {
                StreetID = streets.Rows[ind][0].ToString();
                StreetName = streets.Rows[ind][1].ToString();
                DialogResult = DialogResult.OK;
            }
        }

        private void city_street_edit_Load(object sender, EventArgs e)
        {

        }

        /// <summary>
        /// выбор улицы при двойном щелчке
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void listBox1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            FixResult();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            button1.Enabled = true;
            button4.Enabled = true;
            Width = 401;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;
            button4.Enabled = false;
            edit = false;
            textBox2.Text = string.Empty;
            Width = 771;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (textBox2.Text.Trim().Length == 0) return;            

            if (!edit)
            {
                textBox2.Text = textBox2.Text.Trim();

                string ulname = textBox2.Text.Trim().Replace("'", "");
                textBox2.Text = ulname;

                DataRow[] drr = streets.Select("name='" + ulname + "'");

                if (drr.Length > 0)
                {
                    MessageBox.Show("Редактируемая улица уже имеется в списке.",
                        "Отказ операции", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    listBox1.Text = ulname;
                    return;
                }

                string sql = "insert into street (city_id, name) values(@citid,@nm)";
                SqlCommand cmd = new SqlCommand(sql, main.global_connection);
                cmd.Parameters.Add("@citid", SqlDbType.Int).Value = GorodID;
                cmd.Parameters.Add("@nm", SqlDbType.NVarChar).Value = ulname;
                cmd.ExecuteNonQuery();

                streets = new DataTable();
                sql = "select street.id, street.name from street " +
                    " where street.city_id = " + GorodID + " order by name";
                (new SqlDataAdapter(sql, main.global_connection)).Fill(streets);
                listBox1.Items.Clear();
                textBox2.AutoCompleteCustomSource.Clear();
                foreach (DataRow d in streets.Rows)
                {
                    listBox1.Items.Add(d[1].ToString());
                    textBox2.AutoCompleteCustomSource.Add(d[1].ToString());
                }

                listBox1.Text = ulname;
                Width = 401;
            }
            else
            {
                textBox2.Text = textBox2.Text.Trim();

                string ulname = textBox2.Text.Trim().Replace("'", "");
                textBox2.Text = ulname;

                DataRow[] drr = streets.Select("name='" + ulname + "'");

                if (drr.Length > 0)
                {
                    MessageBox.Show("Редактируемая улица уже имеется в списке.",
                        "Отказ операции", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    listBox1.Text = ulname;
                    return;
                }

                string sql = "update street set name=@nm where id = @ulid";
                SqlCommand cmd = new SqlCommand(sql, main.global_connection);
                cmd.Parameters.Add("@ulid", SqlDbType.Int).Value =
                    streets.Rows[listBox1.SelectedIndex][0].ToString();
                cmd.Parameters.Add("@nm", SqlDbType.NVarChar).Value = ulname;
                cmd.ExecuteNonQuery();

                Width = 401;

                streets = new DataTable();
                sql = "select street.id, street.name from street " +
                    " where street.city_id = " + GorodID + " order by name";
                (new SqlDataAdapter(sql, main.global_connection)).Fill(streets);
                listBox1.Items.Clear();
                textBox2.AutoCompleteCustomSource.Clear();
                foreach (DataRow d in streets.Rows)
                {
                    listBox1.Items.Add(d[1].ToString());
                    textBox2.AutoCompleteCustomSource.Add(d[1].ToString());
                }

                listBox1.Text = ulname;                
            }

            button1.Enabled = true;
            button4.Enabled = true;
        }

        /// <summary>
        /// сохранение результата при нажатии кнопки Принять
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button7_Click(object sender, EventArgs e)
        {
            FixResult();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }



        private void button4_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedIndex < 0)
            {
                MessageBox.Show("Выбирите название улицы для редактирования",
                    "Запрос", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            button1.Enabled = false;
            button4.Enabled = false;
            edit = true;
            textBox2.Text = listBox1.Text;
            Width = 771;            
        }
    }
}
