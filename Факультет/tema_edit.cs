using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace FSystem
{
    public partial class tema_edit : Form
    {
        public tema_edit()
        {
            InitializeComponent();
        }

        public int predmet_id = 0;
        public int vid_rab_id = 0,
            new_tema_id = 0;
        DataTable temas = new DataTable();
        public string new_tema_name = "";
        public DataRow[] filteredtema;
        public bool Filtered = false;


        /// <summary>
        /// выделить строку в таблице с темой по ее ид
        /// </summary>
        public void select_by_id(int id)
        {
            int i = 0;
            foreach (DataRow dr in temas.Rows)
            {
                int id_l = Convert.ToInt32(dr[2]);
                if (id_l == id)
                {
                    table.CurrentCell = table.Rows[i].Cells[0];
                }
                i++;
            }
        }

        /// <summary>
        /// удалить недопустимые символы
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static string Normalize(string str)
        {
            while (str.Contains("'"))
            {
                int pos = str.IndexOf("'");
                str = str.Remove(pos, 1);
                if (str.Length == 0) break;
            }

            return str;
        }

        public string filter = "";

        public void fill_temas()
        {
            //загрузить темы по данному предмету
            temas = new DataTable();

            string q = "";
            q = "select name, content, id from tema_rabota where " +
                //" predmet_id = " + predmet_id.ToString() + " and "
                " vid_rabota_id = " + vid_rab_id.ToString() +
                //" and id not in (select isnull(tema_id,-1) from student_rabota) " + 
                " order by name ";
            main.global_adapter = new SqlDataAdapter(q, main.global_connection);
            main.global_adapter.Fill(temas);            

            table.Rows.Clear();
            if (temas.Rows.Count == 0)
                return;

            foreach (DataRow dr in temas.Rows)
            {
                object[] vals = new object[2] { dr[0], dr[1] };
                table.Rows.Add(vals);
            }            
        }

        private void tema_edit_Load(object sender, EventArgs e)
        {
            fill_temas();
        }

        private void toolStripButton1_Click_1(object sender, EventArgs e)
        {
            //создать новую тему
            tema_add ta = new tema_add();

            DialogResult res = ta.ShowDialog();

            if (res == DialogResult.Cancel) return;

            string tm = "", cnt = "";

            tm = Normalize(ta.tema.Text.Trim());
            cnt = Normalize(ta.cont.Text.Trim());

            ta.Dispose();

            string q = string.Format(
                "insert into tema_rabota(name, content, predmet_id, vid_rabota_id) " +
                " values('{0}','{1}',{2},{3})",
                tm, cnt, predmet_id, vid_rab_id);
            main.global_command = new SqlCommand(q, main.global_connection);

            try
            {
                main.global_command.ExecuteNonQuery();
            }
            catch (Exception exx)
            {
                MessageBox.Show("Невозможно выполнить операцию вследствие сетевой ошибки. Повторите попытку позднее.",
                    "Ошибка операции", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            fill_temas();

            q = "select @@identity";
            main.global_command = new SqlCommand(q, main.global_connection);
            SqlDataReader r = main.global_command.ExecuteReader();

            r.Read();
            new_tema_id = Convert.ToInt32(r[0]);
            select_by_id(new_tema_id);
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {                    
            
            //редактирование темы
            int row = 0;
            if (table.CurrentCell != null)
                row = table.CurrentCell.RowIndex;
            else
                return;            

            
            string tm = "";
            string cnt = "";
            int id = 0;

            if (!Filtered)
            {
                tm = temas.Rows[row][0].ToString();
                cnt = temas.Rows[row][1].ToString();
                id = Convert.ToInt32(temas.Rows[row][2]);
            }
            else
            {
                tm = filteredtema[row][0].ToString();
                cnt = filteredtema[row][1].ToString();
                id = Convert.ToInt32(filteredtema[row][2]);
            }

            tema_add ta = new tema_add();
            ta.tema.Text = tm;
            ta.cont.Text = cnt;

            DialogResult res = ta.ShowDialog();
            if (res == DialogResult.Cancel) return;

            tm = Normalize(ta.tema.Text.Trim());
            cnt = Normalize(ta.cont.Text.Trim()); 

            string q = string.Format(
                "update tema_rabota set name = '{0}', content = '{1}', predmet_id = {2}, vid_rabota_id = {3} " +
                " where id = {4}", tm, cnt, predmet_id, vid_rab_id, id);
            main.global_command = new SqlCommand(q, main.global_connection);

            try
            {
                main.global_command.ExecuteNonQuery();
            }
            catch (Exception exx)
            {
                MessageBox.Show("Невозможно выполнить операцию вследствие сетевой ошибки. Повторите попытку позднее.",
                    "Ошибка операции", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            fill_temas();
            if (Filtered) toolStripTextBox1_KeyUp(sender, new KeyEventArgs(Keys.Enter));

            //select_by_id(id);
            ta.Dispose();
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            int row = 0;
            if (table.CurrentCell != null)
                row = table.CurrentCell.RowIndex;
            else
            {
                if (table.Rows.Count > 0)
                    MessageBox.Show("Выделите строку для удаления.", "Что удалять?",
                        MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                else
                    MessageBox.Show("Нет строк для удаления.", "Нечего удалять",
                        MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            DialogResult r = MessageBox.Show("Выделенная тема будет удалена. Продолжить?",
                "Запрос потдтверждения",
                MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);

            if (r == DialogResult.No) return;

            string q = "delete from tema_rabota " +
                " where id = " + temas.Rows[row][2].ToString();
            main.global_command = new SqlCommand(q, main.global_connection);

            try
            {
                main.global_command.ExecuteNonQuery();
            }
            catch (Exception exx)
            {
                MessageBox.Show("Произошел сбой при удалении данных. Повторите операцию позднее.",
                    "Сетевой сбой", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return;

            }

            if (table.Rows.Count > 0)
            {
                if (row >= 1)
                    table.CurrentCell = table.Rows[row - 1].Cells[0];
                else
                    table.CurrentCell = table.Rows[0].Cells[0];
            }

            fill_temas();
        }

        private void table_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            toolStripButton2_Click(sender, new EventArgs());
        }

        private void toolStripButton4_Click(object sender, EventArgs e)
        {

            if (table.Rows.Count == 0)
            {
                MessageBox.Show(
                    "Перечень тем пуст. Добавьте тему или освободите для использования одну из использованных ранее тем.",
                    "Ошибка операции", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            int row = table.CurrentCell.RowIndex;
            if (!Filtered)
            {
                new_tema_id = Convert.ToInt32(temas.Rows[row][2]);
                new_tema_name = temas.Rows[row][0].ToString();
            }
            else
            {
                new_tema_id = Convert.ToInt32(filteredtema[row][2]);
                new_tema_name = filteredtema[row][0].ToString();                    
            }

            DialogResult = DialogResult.OK;
        }

        private void toolStripButton5_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private void toolStripTextBox1_KeyUp(object sender, KeyEventArgs e)
        {
            if (toolStripTextBox1.Text.Trim().Length == 0)
                Filtered = false;
            else
                Filtered = true;

            filter = "name like '%" + toolStripTextBox1.Text + "%'";
            filteredtema = temas.Select(filter);

            table.Rows.Clear();            
            foreach (DataRow dr in filteredtema)
            {
                object[] vals = new object[2] { dr[0], dr[1] };
                table.Rows.Add(vals);
            }
        }
    }
}