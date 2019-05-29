using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace FSystem
{
    public partial class student_edit_born_place : Form
    {
        public student_edit_born_place()
        {
            InitializeComponent();
        }

        DataTable regions = null;
        DataTable city = null;
        int can = 0;
        public int region_Id = 0;
        public int city_id = 0;
        public int student_id = 0;

        private void student_edit_born_place_Load(object sender, EventArgs e)
        {
            string select = "select id, name from region order by id, name";
            main.global_adapter = new SqlDataAdapter(select, main.global_connection);
            regions = new DataTable();
            main.global_adapter.Fill(regions);

            foreach (DataRow dr in regions.Rows)
            {
                region_list.Items.Add(dr[1].ToString());
            }

            region_list.SelectedIndex = 0;
            can++;
        }

        private void region_list_SelectedIndexChanged(object sender, EventArgs e)
        {
            city_list.Items.Clear();
            
            //выбрать город
            string region_id = regions.Rows[region_list.SelectedIndex][0].ToString();

            string select = "select id, name, naspunkt_type_id, postindex from city where region_id = " + region_id +
                " order by naspunkt_type_id desc";         
          
            main.global_adapter = new SqlDataAdapter(select, main.global_connection);
            city = new DataTable();
            main.global_adapter.Fill(city);

            foreach (DataRow dr in city.Rows)
            {
                city_list.Items.Add(dr[1]);
            }

            if (city_list.Items.Count>0)
                city_list.SelectedIndex = 0;

            region_Id = (int)regions.Rows[region_list.SelectedIndex][0];
        }

        private void add_region_Click(object sender, EventArgs e)
        {
            inputbox ib = new inputbox("введите название субъекта российской федерации",
                "дополнение", "", "название субъекта РФ");
            ib.is_numeric = false;

            ib.ShowDialog();
        }

        //добавление нового города
        private void button4_Click(object sender, EventArgs e)
        {
            sprav_city sc = new sprav_city();
            sc.region_name = region_list.Text;
            sc.label2.Text = "Введите название населённого пункта";
            sc.city_name = "";//"введите новое название нас. пункта";
            sc.edit = false;
            if (city_list.Items.Count>0) sc.city_id = (int)city.Rows[city_list.SelectedIndex][0];
            sc.region_id = region_Id;
            if (city_list.Items.Count > 0) sc.vid_poselenie_id = (int)city.Rows[city_list.SelectedIndex][2];
            if (city_list.Items.Count > 0) sc.postindex = (int)city.Rows[city_list.SelectedIndex][3];
            sc.Height = 0;
            
            sc.Left = Left + button4.Left;
            sc.Top = Top + button4.Bottom;

            sc.ShowDialog();

            if (sc.city_id == -1)
            {
                sc.Dispose();
                return;
            }

            city_id = sc.city_id;

            sc.Dispose();

            //снова заполнить список

            city_list.Items.Clear();

            //выбрать город
            string region_id = regions.Rows[region_list.SelectedIndex][0].ToString();

            string select = "select id, name, naspunkt_type_id, postindex from city where region_id = " + region_id +
                " order by naspunkt_type_id desc";

            main.global_adapter = new SqlDataAdapter(select, main.global_connection);
            city = new DataTable();
            main.global_adapter.Fill(city);

            int ind = 0, i = 0;
            foreach (DataRow dr in city.Rows)
            {
                int cid = (int)dr[0];
                if (cid == city_id) ind = i;                    
                city_list.Items.Add(dr[1]);
                i++;
            }

            if (city_list.Items.Count > 0)
                city_list.SelectedIndex = ind;            

        }

        private void city_list_SelectedIndexChanged(object sender, EventArgs e)
        {
            city_id = (int)city.Rows[city_list.SelectedIndex][0];
            if (!textBox1.Focused) textBox1.Text = city_list.Items[city_list.SelectedIndex].ToString();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (city_list.Items.Count == 0) return;

            //сохранить данные о области и городе рождения студента
            string sql = "update student set born_region_id = @bri, " +
                "born_city_id = @bci where id = @sti";
            main.global_command = new SqlCommand(sql, main.global_connection);
            main.global_command.Parameters.Add("@bri", SqlDbType.Int).Value = region_Id;
            main.global_command.Parameters.Add("@bci", SqlDbType.Int).Value = city_id;
            main.global_command.Parameters.Add("@sti", SqlDbType.Int).Value = student_id;
            main.global_command.ExecuteNonQuery();
            DialogResult = DialogResult.OK;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Text = button2.Left.ToString() + " " + button2.Top.ToString();
            
            if (city_list.Items.Count == 0) return;
            
            sprav_city sc = new sprav_city();
            sc.label2.Text = "Измените название населённого пункта";
            sc.region_name = region_list.Text;
            sc.city_name = city_list.Text;
            sc.vid_poselenie_id = (int)city.Rows[city_list.SelectedIndex][2];
            sc.postindex = (int)city.Rows[city_list.SelectedIndex][3];

            sc.edit = true;
            sc.city_id = (int)city.Rows[city_list.SelectedIndex][0];
            sc.region_id = region_Id;
            sc.Left = Left + button2.Left;
            sc.Top = Top + button2.Bottom;
            sc.Height = 0;
           

            sc.ShowDialog();

            if (sc.city_id == -1)
            {
                sc.Dispose();
                return;
            }

            //снова заполнить список

            city_list.Items.Clear();

            //выбрать город
            string region_id = regions.Rows[region_list.SelectedIndex][0].ToString();

            string select = "select id, name, naspunkt_type_id, postindex from city where region_id = " + region_id +
                " order by naspunkt_type_id desc";

            main.global_adapter = new SqlDataAdapter(select, main.global_connection);
            city = new DataTable();
            main.global_adapter.Fill(city);

            int ind = 0, i = 0;
            foreach (DataRow dr in city.Rows)
            {
                int cid = (int)dr[0];
                if (cid == city_id) ind = i;
                city_list.Items.Add(dr[1]);
                i++;
            }

            if (city_list.Items.Count > 0)
                city_list.SelectedIndex = ind;

            sc.Dispose();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {            
            string txt = textBox1.Text.Trim();
            for (int i = 0; i < city_list.Items.Count; i++)
            {
                string lt = city_list.Items[i].ToString();
                if (lt.StartsWith(txt, StringComparison.OrdinalIgnoreCase))
                {
                    city_list.SelectedIndex = i;
                    break;
                }
            }

        }
    }
}