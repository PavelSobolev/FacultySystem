using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Windows.Forms;

namespace FSystem
{
    public partial class sprav_ob : Form
    {
        public sprav_ob()
        {
            InitializeComponent();
        }

        DataTable ob_set;
        DataTable tema_set;


        void fill_tema()
        {
            tema_set = new DataTable();
            string q = "select id, text from tema_ob";

            main.global_adapter = new SqlDataAdapter(q, main.global_connection);
            main.global_adapter.Fill(tema_set);

            tema_list.Items.Add("Все категории");
            foreach (DataRow dr in tema_set.Rows)
            {
                tema_list.Items.Add(dr[1].ToString());
            }

            tema_list.SelectedIndex = 0;
        }

        void load_obs(bool old, int num)
        {
            ob_set = new DataTable();

            string q = "";
            
            if (!old)
                q = string.Format(
                " select obiavlenie.id, [Категория] = tema_ob.text, [Заголовок]=title, " +
                " [Текст] = ob_text," +
                " [Опубликовано] = start_date, [Дата удаления] = end_date, tid = tema_ob.id " + 
                " from obiavlenie " +
                " join tema_ob on tema_ob.id = obiavlenie.tema_id " +
                " where fakultet_id = {0} and obiavlenie.actual=1 " + 
                " and end_date > dbo.get_date({1},{2},{3})", 
                 main.fakultet_id, main.server_date.Year, main.server_date.Month, main.server_date.Day);
            else
                q = string.Format(
                " select obiavlenie.id, [Категория] = tema_ob.text, [Заголовок]=title, " +
                " [Текст] = ob_text, " +
                " [Опубликовано] = start_date, [Дата удаления] = end_date, tid = tema_ob.id from obiavlenie " +
                " join tema_ob on tema_ob.id = obiavlenie.tema_id " +
                " where fakultet_id = {0} and obiavlenie.actual=1",
                 main.fakultet_id);

            if (num > 0) //Фильтр по катеогрии
            {
                q = q + " and tema_ob.id = " + tema_set.Rows[num-1][0].ToString();
            }           

            main.global_adapter = new SqlDataAdapter(q, main.global_connection);
            main.global_adapter.Fill(ob_set);

            obs.Rows.Clear();

            foreach (DataRow dr in ob_set.Rows)
            {
                object[] pars = new object[5]{ dr[1], dr[2], dr[3], dr[4], dr[5] };
                obs.Rows.Add(pars);
            }

            obs.Select();

        }

        private void sprav_ob_Load(object sender, EventArgs e)
        {            
            fill_tema();
            load_obs(oldshow.Checked, tema_list.SelectedIndex);
        }

        private void oldshow_Click(object sender, EventArgs e)
        {
            oldshow.Checked = !oldshow.Checked;
            if (oldshow.Checked)
                oldshow.Font = new Font("tahoma", 8, FontStyle.Bold);
            else
                oldshow.Font = new Font("tahoma", 8, FontStyle.Regular);
            load_obs(oldshow.Checked, tema_list.SelectedIndex);
        }

        private void tema_list_SelectedIndexChanged(object sender, EventArgs e)
        {
            load_obs(oldshow.Checked, tema_list.SelectedIndex);
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {                        
            int row = 0;            
            if (obs.CurrentCell!=null)
                row = obs.CurrentCell.RowIndex;

            publication_obiavl po = new publication_obiavl();
            po.dt = main.server_date.AddDays(10.0);
            DialogResult dr = po.ShowDialog();            

            if (dr == DialogResult.Cancel) return;

            // --- сохранить ---
            string q = "insert into obiavlenie " +
                " (ob_text, fakultet_id, start_date, end_date, " +
                " actual, tema_id, title) " +
                "values " +
                " (@ob_text, @fakultet_id, @start_date, @end_date, " +
                " @actual, @tema_id, @title) ";

            main.global_command = new SqlCommand(q, main.global_connection);
            main.global_command.Parameters.Add("@ob_text", SqlDbType.NVarChar).Value = po.textob.Text;
            main.global_command.Parameters.Add("@fakultet_id", SqlDbType.Int).Value = main.fakultet_id;
            main.global_command.Parameters.Add("@start_date", SqlDbType.DateTime).Value = main.server_date;
            main.global_command.Parameters.Add("@end_date", SqlDbType.DateTime).Value = po.end_date.Value;
            main.global_command.Parameters.Add("@actual", SqlDbType.Bit).Value = 1;
            main.global_command.Parameters.Add("@tema_id", SqlDbType.Int).Value = po.id;
            main.global_command.Parameters.Add("@title", SqlDbType.NVarChar).Value = po.titletxt.Text;
            
            try
            {
                main.global_command.ExecuteNonQuery();
            }
            catch (Exception exx)
            {
                MessageBox.Show("Произошел сбой при сохранении данных. Повторите операцию позднее.",
                    "Сетевой сбой", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                po.Dispose();
                return;

            }

            po.Dispose();
            load_obs(oldshow.Checked, tema_list.SelectedIndex);

            obs.CurrentCell = obs.Rows[row].Cells[0];
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            int row = 0;            
            if (obs.CurrentCell != null)
                row = obs.CurrentCell.RowIndex;
            else
                return;
            
            publication_obiavl po = new publication_obiavl();
            
            po.id = Convert.ToInt32(ob_set.Rows[row][6].ToString());
            po.dt = Convert.ToDateTime(ob_set.Rows[row][5]);
            po.textt = ob_set.Rows[row][3].ToString();
            po.title = ob_set.Rows[row][2].ToString();
            DialogResult dr = po.ShowDialog();
            po.Dispose();

            if (dr == DialogResult.Cancel) return;

            // --- сохранить ---
            string q = "update obiavlenie set " +
                " ob_text = @ob_text, fakultet_id = @fakultet_id, start_date = @start_date, end_date = @end_date, " +
                " actual = @actual, tema_id = @tema_id, title = @title  " +
                " where id = " + ob_set.Rows[row][0].ToString();

            main.global_command = new SqlCommand(q, main.global_connection);
            main.global_command.Parameters.Add("@ob_text", SqlDbType.NVarChar).Value = po.textob.Text;
            main.global_command.Parameters.Add("@fakultet_id", SqlDbType.Int).Value = main.fakultet_id;
            main.global_command.Parameters.Add("@start_date", SqlDbType.DateTime).Value = main.server_date;
            main.global_command.Parameters.Add("@end_date", SqlDbType.DateTime).Value = po.end_date.Value;
            main.global_command.Parameters.Add("@actual", SqlDbType.Bit).Value = 1;
            main.global_command.Parameters.Add("@tema_id", SqlDbType.Int).Value = po.id;
            main.global_command.Parameters.Add("@title", SqlDbType.NVarChar).Value = po.titletxt.Text;

            try
            {
                main.global_command.ExecuteNonQuery();
            }
            catch (Exception exx)
            {
                MessageBox.Show("Произошел сбой при сохранении данных. Повторите операцию позднее.",
                    "Сетевой сбой", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                po.Dispose();
                return;

            }

            po.Dispose();
            load_obs(oldshow.Checked, tema_list.SelectedIndex);
            obs.CurrentCell = obs.Rows[row].Cells[0];

        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            int row = 0;
            if (obs.CurrentCell != null)
                row = obs.CurrentCell.RowIndex;
            else
            {
                if (obs.Rows.Count>0)
                    MessageBox.Show("Выделите строку для удаления.", "Что удалять?",
                        MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                else
                    MessageBox.Show("Нет строк для удаления.", "Нечего удалять",
                        MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            DialogResult r = MessageBox.Show("Выделенное объявление будет удалено. Продолжить?", 
                "Запрос потдтверждения",
                MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);

            if (r == DialogResult.No) return;

            string q = "delete from obiavlenie " +                
                " where id = " + ob_set.Rows[row][0].ToString();

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

            load_obs(oldshow.Checked, tema_list.SelectedIndex);

            if (obs.Rows.Count>1)
                obs.CurrentCell = obs.Rows[row-1].Cells[0];
        }
    }
}