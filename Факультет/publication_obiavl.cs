using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace FSystem
{
    public partial class publication_obiavl : Form
    {
        public publication_obiavl()
        {
            InitializeComponent();
        }

        DataTable tema_set;
        public int id = 0;
        public string title = "", textt = "";
        public DateTime dt;

        private void publication_obiavl_Load(object sender, EventArgs e)
        {
            tema_set = new DataTable();
            string q = "select id, text from tema_ob";

            main.global_adapter = new SqlDataAdapter(q, main.global_connection);
            main.global_adapter.Fill(tema_set);
            
            foreach (DataRow dr in tema_set.Rows)
            {
                tema_list.Items.Add(dr[1].ToString());
            }

            if (id == 0)
            {
                tema_list.SelectedIndex = 0;
                end_date.Value = end_date.Value.AddDays(10);
            }
            else
            {
                tema_list.SelectedIndex = find_tema(id);
                titletxt.Text = title;
                textob.Text = textt;
                end_date.Value = dt;
            }
        }

        int find_tema(int number)
        {
            int i = 0;
            foreach (DataRow dr in tema_set.Rows)
            {

                int idd = Convert.ToInt32(dr[0].ToString());
                if (number == idd) return i;
                else i++;
            }

            return 0;
        }

        private void tema_list_SelectedIndexChanged(object sender, EventArgs e)
        {
            id = Convert.ToInt32(tema_set.Rows[tema_list.SelectedIndex][0].ToString());
        }

        private void save_Click(object sender, EventArgs e)
        {
            if (end_date.Value <= main.server_date)
            {
                MessageBox.Show("Дата удаления объявления не может быть раньше текущей даты.",
                    "Ошибка выбора даты", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return;
            }

            if (titletxt.Text.Trim().Length == 0)
            {
                MessageBox.Show("Заголовок объявления не может быть пустым.",
                    "Ошибка ввода заголовка", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return;
            }

            if (textob.Text.Trim().Length == 0)
            {
                MessageBox.Show("Текст объявления не может быть пустым.",
                    "Ошибка ввода объявления", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return;
            }

            DialogResult = DialogResult.OK;
        }
    }
}