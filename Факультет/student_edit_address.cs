using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace FSystem
{
    public partial class student_edit_address : Form
    {
        public student_edit_address()
        {
            InitializeComponent();
        }

        public string stud_id = string.Empty;
        public string region_id = "1";
        public string gorod_id = "1";
        public string ul_id = string.Empty;
        public string dom = "1";
        public string kvart = "1";
        public bool edit = false;

        private void student_edit_address_Load(object sender, EventArgs e)
        {
            streets = new DataTable();
            string sql = "select street.id, street.name from street " +
                " where street.city_id = " + gorod_id + " order by name";
            (new SqlDataAdapter(sql, main.global_connection)).Fill(streets);
        }

        public DataTable regions;
        private void born_place__button_Click(object sender, EventArgs e)
        {
            regions = new DataTable();
            string sql = "select id, name from region";
            (new SqlDataAdapter(sql, main.global_connection)).Fill(regions);

            ListWindow lw = new ListWindow();
            lw.Text = "Выберите название региона из списка";
            lw.tbl = regions;
            DialogResult res = lw.ShowDialog();

            if (res == DialogResult.OK)
            {
                region_id = lw.resId.ToString();
                region_box.Text = lw.str_res;
            }

            lw.Dispose();
            GC.Collect();
        }

        public DataTable Cities;
        private void button1_Click(object sender, EventArgs e)
        {
            Cities = new DataTable();
            string sql = "select city.id, nm = naspunkt_type.name + ' ' + city.name from city " + 
                " join naspunkt_type on naspunkt_type.id = city.naspunkt_type_id " + 
                " where city.region_id = " + region_id + 
                " order by naspunkt_type.out_order";
            (new SqlDataAdapter(sql, main.global_connection)).Fill(Cities);

            ListWindow lw = new ListWindow();
            lw.Text = "Выберите нас. пункт";
            lw.tbl = Cities;
            DialogResult res = lw.ShowDialog();

            if (res == DialogResult.OK)
            {
                gorod_id = lw.resId.ToString();
                nas_punkt_box.Text = lw.str_res;
            }

            lw.Dispose();
            GC.Collect();

            streets = new DataTable();
            sql = "select street.id, street.name from street " +
                " where street.city_id = " + gorod_id + " order by name";
            (new SqlDataAdapter(sql, main.global_connection)).Fill(streets);

            street.Text = string.Empty;

        }

        public DataTable streets;
        private void button2_Click(object sender, EventArgs e)
        {
            streets = new DataTable();
            string sql = "select street.id, street.name from street " +
                " where street.city_id = " + gorod_id + " order by name";
            (new SqlDataAdapter(sql, main.global_connection)).Fill(streets);
            city_street_edit cse = new city_street_edit(streets, gorod_id);
            DialogResult res = cse.ShowDialog();

            if (res == DialogResult.OK)
            {
                ul_id = cse.StreetID;
                street.Text = cse.StreetName;
            }

            cse.Dispose();
            GC.Collect();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            
        }

        private void button7_Click_1(object sender, EventArgs e)
        {
            // вернуться в окно редактирования данных студента
            if (street.Text.Trim().Length == 0)
            {
                MessageBox.Show("Введите название улицы!", "Отказ операции",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (house.Text.Trim().Length == 0)
            {
                MessageBox.Show("Введите номер дома!", "Отказ операции",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            DialogResult = DialogResult.OK;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }
    }
}