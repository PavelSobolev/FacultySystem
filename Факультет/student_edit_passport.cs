using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

namespace FSystem
{
    public partial class student_edit_passport : Form
    {
        public student_edit_passport()
        {
            InitializeComponent();
        }

        public AutoCompleteStringCollection UVDS = new AutoCompleteStringCollection();
        public int stud_id = 0;

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            int[] nums = { 8, 32, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57 };
            if (Array.IndexOf(nums, (int)e.KeyChar)== -1) e.KeyChar = '\0';
        }

        private void nomer_KeyPress(object sender, KeyPressEventArgs e)
        {
            int[] nums = { 8, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57 };
            if (Array.IndexOf(nums, (int)e.KeyChar) == -1) e.KeyChar = '\0';   
        }

        private void student_edit_passport_Load(object sender, EventArgs e)
        {
            дата_выдачи.MaxDate = DateTime.Now.AddDays(-1);
            string sql = "select 'УВД г. ' + name from city where naspunkt_type_id=4";
            DataTable uvds = new DataTable();
            main.global_adapter = new System.Data.SqlClient.SqlDataAdapter(sql, main.global_connection);
            main.global_adapter.Fill(uvds);

            UVDS.Add(" УВД г. Южно-Сахалинска");
            foreach(DataRow dr in uvds.Rows) UVDS.Add(dr[0].ToString());
            vydano.AutoCompleteCustomSource = UVDS;                        
        }

        private void dateTimePicker1_Enter(object sender, EventArgs e)
        {
            // to be removed
        }

        private void textBox1_Enter(object sender, EventArgs e)
        {
            ((Control)sender).BackColor = Color.Yellow;
        }

        private void textBox1_Leave(object sender, EventArgs e)
        {
            ((Control)sender).BackColor = Color.White;
                //235; 240; 243
        }

        private void button7_Click(object sender, EventArgs e)
        {
            bool fail = false;
            string res = "";

            if (seria.Text.Trim().Length==0)
            {
                res = "Не указана серия паспорта\n\n";
                fail = true;
            }

            if (nomer.Text.Trim().Length == 0)
            {
                res = res + "Не указан номер паспорта\n\n";
                fail = true;
            }

            if (vydano.Text.Trim().Length == 0)
            {
                res = res + "Не указана организация, выдавшая паспорт.\n\n";
                fail = true;
            }

            if (fail)
            {
                MessageBox.Show(
                    res + "Исправьте указанные замечения и повторите ввод.", 
                    "Ошибка данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            string sql = "update student set passport_nomer=@pn , passport_seria=@ps , " + 
                "passport_date=@pd, passport_vydan=@pv where id = @sid";
            main.global_command = new System.Data.SqlClient.SqlCommand(sql, main.global_connection);
            main.global_command.Parameters.Add("@pn",SqlDbType.NVarChar).Value = nomer.Text.Trim();
            main.global_command.Parameters.Add("@ps",SqlDbType.NVarChar).Value = seria.Text;
            main.global_command.Parameters.Add("@pd",SqlDbType.DateTime).Value = дата_выдачи.Value;
            main.global_command.Parameters.Add("@pv",SqlDbType.NVarChar).Value = vydano.Text;
            main.global_command.Parameters.Add("@sid",SqlDbType.Int).Value = stud_id;
            main.global_command.ExecuteNonQuery();

            DialogResult = DialogResult.OK;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }
    }
}