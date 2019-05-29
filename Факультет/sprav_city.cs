using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace FSystem
{
    public partial class sprav_city : Form
    {
        public sprav_city()
        {
            InitializeComponent();
        }

        public string region_name = "";
        public string city_name = "";
        public int city_id = 0;
        public bool edit = false;
        public int region_id = 0;
        public int vid_poselenie_id = 0;
        public int postindex = 0;
        DataTable vids = new DataTable();

        private void sprav_city_Load(object sender, EventArgs e)
        {

            timer1.Enabled = true;

            
            nas_punkt_box.Text = city_name;
            index.Text = postindex.ToString();
            Text = "Насел. пункты (" + region_name + ")";

            string sql = "select id, name from naspunkt_type";
            main.global_adapter = new SqlDataAdapter(sql, main.global_connection);
            main.global_adapter.Fill(vids);

            comboBox1.DataSource = vids;
            comboBox1.DisplayMember = "name";
            comboBox1.SelectedIndex = main.getIndex(vids,vid_poselenie_id);

            nas_punkt_box.Select();
        }



        private void button7_Click(object sender, EventArgs e)
        {
            if (nas_punkt_box.Text.Trim().Length == 0) return;

            string sql = "select name from city where name like '" + nas_punkt_box.Text.Trim() + "'" + 
                " and region_id = " + region_id.ToString();
            main.global_command = new SqlCommand(sql, main.global_connection);
            SqlDataReader dr = main.global_command.ExecuteReader();
            if (dr.Read())
            {
                MessageBox.Show(
                    "Населённый пункт с таким названием уже существует. Измените имя.",
                    "Ошибка ввода",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                dr.Close();
                return;
            }
            dr.Close();

            if (!edit)
            {
                sql =
                    "insert into city(region_id, name, naspunkt_type_id, postindex)" +
                    " values(@region_id, @name, @naspunkt_type_id, @postindex)";

                main.global_command = new SqlCommand(sql, main.global_connection);
                main.global_command.Parameters.Add("@region_id", SqlDbType.Int).Value = region_id;
                main.global_command.Parameters.Add("@name", SqlDbType.NVarChar).Value = nas_punkt_box.Text;
                main.global_command.Parameters.Add("@naspunkt_type_id", SqlDbType.NVarChar).Value = main.getID(vids, comboBox1.SelectedIndex);
                if (index.Text.Trim().Length == 0) index.Text = "0";
                main.global_command.Parameters.Add("@postindex", SqlDbType.Int).Value = Convert.ToInt32(index.Text);

                try
                {
                    main.global_command.ExecuteNonQuery();
                }
                catch (Exception exx)
                {
                    city_id = -1;
                    DialogResult = DialogResult.Cancel;
                    return;
                }

                sql = "select @@identity";
                main.global_command = new SqlCommand(sql, main.global_connection);
                city_id = Convert.ToInt32(main.global_command.ExecuteScalar());
                DialogResult = DialogResult.OK;
            }
            else
            {
                sql = " update city set name=@name, naspunkt_type_id=@naspunkt_type_id, postindex=@postindex " +
                    " where region_id = @region_id and id = " + city_id.ToString();

                main.global_command = new SqlCommand(sql, main.global_connection);
                main.global_command.Parameters.Add("@region_id", SqlDbType.Int).Value = region_id;
                main.global_command.Parameters.Add("@name", SqlDbType.NVarChar).Value = nas_punkt_box.Text;
                main.global_command.Parameters.Add("@naspunkt_type_id", SqlDbType.NVarChar).Value = main.getID(vids, comboBox1.SelectedIndex);
                if (index.Text.Length == 0) index.Text = "0";
                main.global_command.Parameters.Add("@postindex", SqlDbType.Int).Value = Convert.ToInt32(index.Text);

                try
                {
                    main.global_command.ExecuteNonQuery();
                }
                catch (Exception exx)
                {
                    city_id = -1;
                    DialogResult = DialogResult.Cancel;
                    return;
                }

                DialogResult = DialogResult.OK;
            }

        }

        private void krat_name_KeyPress(object sender, KeyPressEventArgs e)
        {            
            int[] nums = { 39, 33, 64, 36, 37, 94, 38, 42, 43, 47, 46, 40, 41, 61, 91, 93, 123, 125, 124, 92, 34, 
                    48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 8470, 35, 59, 58, 63, 95, 126, 96, 60, 62 };
            if (Array.IndexOf(nums, (int)e.KeyChar) != -1) e.KeyChar = '\0';   
        }

        private void index_KeyPress(object sender, KeyPressEventArgs e)
        {
            int[] nums = { 8, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57 };
            if (Array.IndexOf(nums, (int)e.KeyChar) == -1) e.KeyChar = '\0';
        }

        private void sprav_city_Move(object sender, EventArgs e)
        {
            Text = Left.ToString() + " " + Top.ToString();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {

            if (Height <= 160)
            {
                Height = Height + 10;
                return;
            }

            timer1.Enabled = false;
        }

        private void sprav_city_FormClosing(object sender, FormClosingEventArgs e)
        {
            vids.Dispose();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            //
        }
    }
}