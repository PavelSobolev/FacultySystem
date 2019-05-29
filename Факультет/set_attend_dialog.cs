using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Windows.Forms;

namespace FSystem
{
    public partial class set_attend_dialog : Form
    {
        public set_attend_dialog()
        {
            InitializeComponent();
        }

        //внешние данные
        public int zan_id = 0, grupa_id = 0, subgr = 0;
        public bool delenie = false;
        public string tema = "";
        public string prim = "";

        //для работы с sql
        private SqlCommand cmd = new SqlCommand();
        public DataTable student_set = null, attend_set = null;
        public DataGridViewComboBoxCell otm_cell; //ячейки для хранения перечня отметок
        public DataGridViewComboBoxCell zach_cell; //ячейки для хранения перечня отметок
        DataTable otmetki;
        DataTable zachet;

        List<int> ids = new List<int>();      

        public int c_pris = 0, c_ots = 0, c_neizv = 0;

        public void check_or_create_attend()
        {                       
            string q1 = "select count(*) from attend where " + 
                "zan_id = " + zan_id.ToString();

            cmd = new SqlCommand(q1, main.global_connection);
            int res = (int)cmd.ExecuteScalar();            

            if (res != 0) return;

            //создать записи для данной подгруппы
            if (delenie)
                q1 = string.Format("select id from student where " +
                    " gr_id = {0} and subgr_nomer = {1} and actual=1 and status_id=1",
                    grupa_id, subgr);
            else
                q1 = string.Format("select id from student where " +
                " gr_id = {0} and actual=1  and status_id=1",
                grupa_id);

            student_set = new DataTable();          
            
            SqlDataAdapter sda = new SqlDataAdapter(q1, main.global_connection);            
            sda.Fill(student_set);

            string insert_command = "";

            foreach (DataRow dr in student_set.Rows)
            {
                q1 = "insert into attend(stud_id, zan_id, attend_id, " +
                        " otmetka1,otmetka2,otmetka3, otmetka4, otmetka5) " + 
                        " values ( " + 
                        dr[0].ToString() + ", " + 
                        zan_id.ToString() + ", 2, 10, 10, 10, 10, 10)";
                cmd = new SqlCommand(q1, main.global_connection);
                cmd.ExecuteNonQuery();
            }

        }

        public void fill_attend_table()
        {
            if (zan_id == 0)
            {
                Dispose();
                return;
            }
            if (grupa_id == 0)
            {
                Dispose();
                return;
            }
            
            check_or_create_attend();
            
            try
            {
                ;
            }
            catch (Exception exx)
            {
                MessageBox.Show("Непредвиденная ошибка при получении данных. " +
                    "\nВозможно потреяно соединение с сервером данных. Повторите операцию позднее.",
                    "Операция отменена",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                Dispose();
                return;
            }

            textBox1.Text = tema;
            textBox2.Text = prim;

            string q = "select aid=attend.id, " +
                " vattid=vid_attend.id, " +
                " otmetka1,  otmetka2, otmetka3, " +
                " fio = fam + ' ' + im, " +
                " vid_attend.name " +
                " from student " +
                " join attend on student.id = attend.stud_id " +
                " join vid_attend on attend.attend_id=vid_attend.id " +
                " where student.gr_id = " + grupa_id.ToString() +
                " and attend.zan_id = " + zan_id.ToString() +
                " and fam<>'-' and student.actual=1 and len(fam)>0 and status_id=1 order by fio ";

            SqlDataAdapter sda = new SqlDataAdapter(q, main.global_connection);
            attend_set = new DataTable();
            sda.Fill(attend_set);

            if (attend_set.Rows.Count == 0)
            {
                if (!delenie)
                    MessageBox.Show("Нет данных о списочном составе данной группы.\nПроизведите заполнение группы и повторите операцию.",
                        "Операция отменена",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                else
                    MessageBox.Show("Нет данных о списочном составе данной подгруппы.\nПроизведите разделение группы на подгруппы и повторите операцию.",
                        "Операция отменена",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);

                Dispose();
                return;
            }

            int i = 0;

            attend_grid.Rows.Clear();

            do
            {
                attend_grid.Rows.Add();
                attend_grid[0, i].Value = (i + 1).ToString();
                attend_grid[1, i].Value = attend_set.Rows[i][5].ToString();

                string att_name = attend_set.Rows[i][6].ToString();

                attend_grid[2, i].Value = att_name;

                if (att_name.Contains("опред"))
                {
                    attend_grid[2, i].Style.ForeColor = Color.Gray;
                    c_neizv++;
                }

                if (att_name.Contains("отсут"))
                {
                    attend_grid[2, i].Style.ForeColor = Color.Red;
                    c_ots++;
                }

                if (att_name.Contains("присут"))
                {
                    attend_grid[2, i].Style.ForeColor = Color.Blue;
                    c_pris++;
                }

                i++;
            }
            while (i < attend_set.Rows.Count);

            pris_label.Text = "Присутствуют: " + c_pris.ToString() + "   ";
            ots_label.Text = "Отсутствуют: " + c_ots.ToString() + "   ";

            attend_grid.Columns[1].Width = 170;

        }

        private void set_attend_dialog_Load(object sender, EventArgs e)
        {
            fill_attend_table();

            //заполнить списики оценок и зачетов для отображения в сетке
            otm_cell = new DataGridViewComboBoxCell();
            zach_cell = new DataGridViewComboBoxCell();
            
            zachet = new DataTable();
            string query = "select vid_otmetka.id, vid_otmetka.str_name " +
                " from vid_otmetka " +
                " where vid_otmetka.str_name like '%зачтено%'";
            main.global_adapter = new SqlDataAdapter(query, main.global_connection);
            main.global_adapter.Fill(zachet);
            int i = 0;
            for (i = 0; i < zachet.Rows.Count; i++)
            {
                string nmz = zachet.Rows[i][1].ToString();
                zach_cell.Items.Add(nmz);
            }

            otmetki = new DataTable();
            query = "select vid_otmetka.id, vid_otmetka.str_name " +
                " from vid_otmetka " +
                " where vid_otmetka.str_name like '%удовлетворит%' or vid_otmetka.str_name like '%хорошо%' or vid_otmetka.str_name like '%отлично%'";
            main.global_adapter = new SqlDataAdapter(query, main.global_connection);
            main.global_adapter.Fill(otmetki);
            i = 0;
            for (i = 0; i < otmetki.Rows.Count; i++)
            {
                string nmo = otmetki.Rows[i][1].ToString();
                otm_cell.Items.Add(nmo);
            }


        }

        private void attend_grid_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex == -1 || e.ColumnIndex!=2) return;

            int i = e.RowIndex;
            DataRow r = attend_set.Rows[i];

            string att_name = attend_grid[2, e.RowIndex].Value.ToString();


            string cmd = "update attend set attend_id = @AID where id = @ID";

            int aid = 0;
            
            if (att_name.Contains("опред"))
            {
                aid = 2;
            }

            if (att_name.Contains("отсут"))
            {
                aid = 2;
            }

            if (att_name.Contains("присут"))
            {
                aid = 1;
            }


            SqlCommand scmd = new SqlCommand(cmd, main.global_connection);

            scmd.Parameters.Add("@AID", SqlDbType.Int).Value = aid;
            scmd.Parameters.Add("@ID", SqlDbType.Int).Value = r[0];

            try
            {
                scmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Непредвиденная ошибка при передаче данных. Повторите операцию позднее.\nПричина:\n" + 
                    ex.Message, "Ошибка при сохранении данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (att_name.Contains("опред"))
            {
                attend_grid[2, i].Style.ForeColor = Color.Blue;
                attend_grid[2, i].Value = "присутствует";
                c_pris++;
            }

            if (att_name.Contains("отсут"))
            {
                attend_grid[2, i].Style.ForeColor = Color.Blue;
                attend_grid[2, i].Value = "присутствует";
                c_ots--;
                c_pris++;
            }

            if (att_name.Contains("присут"))
            {
                attend_grid[2, i].Style.ForeColor = Color.Red;
                attend_grid[2, i].Value = "отсутствует";
                c_ots++;
                c_pris--;
            }

            pris_label.Text = "Присутствуют: " + c_pris.ToString() + "   ";
            ots_label.Text = "Отсутствуют: " + c_ots.ToString() + "   ";

        }

        private void set_attend_dialog_HelpRequested(object sender, HelpEventArgs hlpevent)
        {       
            MessageBox.Show("Для изменения статуса студента выполните двойной" + 
                "\nщелчок в столбце 'Присутствие' напротив нужной фамилии.",
                "Справка",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        public bool tema_changed = false;
        public bool prim_changed = false;

        private void set_attend_dialog_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (tema.ToLower().Trim() != textBox1.Text.ToLower().Trim())
            {
                string sql = "update rasp set tema = '" + Normalize(textBox1.Text) +
                    "', prim_text = '" + Normalize(textBox2.Text) + "' where id = " + zan_id.ToString();

                SqlCommand csql = new SqlCommand(sql, main.global_connection);
                try
                {
                    csql.ExecuteNonQuery();
                }
                catch (Exception exx)
                {
                    ;
                }

                tema_changed = true;
            }

            //if (prim.ToLower().Trim() != textBox2.Text.ToLower().Trim())
            //{
                string sql1 = "update rasp set prim_text = '" + Normalize(textBox2.Text) + "' where id = " + zan_id.ToString();

                SqlCommand  csql1= new SqlCommand(sql1, main.global_connection);
                try
                {
                    csql1.ExecuteNonQuery();
                }
                catch (Exception exx1)
                {
                    ;
                }

                prim_changed  = true;
            //}
        }


        public string Normalize(string str)
        {
            while (str.Contains("'"))
            {
                int pos = str.IndexOf("'");
                str = str.Remove(pos, 1);
                if (str.Length == 0) break;
            }

            return str;
        }

        
        int otm_count = 0;
        private void выставитьОценкуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (otm_count + zach_count == 5) return;

            otm_count++;
            int col = attend_grid.Columns.Count;            

            DataGridViewColumn coll = new DataGridViewColumn(otm_cell);
            coll.SortMode = DataGridViewColumnSortMode.NotSortable;
            coll.DefaultCellStyle.BackColor = Color.MintCream;
            coll.HeaderText = "Отметка " + otm_count.ToString();
            attend_grid.Columns.Add(coll);

            coll = new DataGridViewTextBoxColumn();
            coll.SortMode = DataGridViewColumnSortMode.NotSortable;
            coll.DefaultCellStyle.BackColor = Color.LightGray;
            coll.HeaderText = "Прим. к отм. " + otm_count.ToString();
            attend_grid.Columns.Add(coll);
        }
        
        int zach_count = 0;   
        //добавить столбец для выставления зачетов на занятии
        private void выставитьЗачётToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (otm_count + zach_count == 5) return;
            zach_count++;
            int col = attend_grid.Columns.Count;            

            DataGridViewColumn coll = new DataGridViewColumn(zach_cell);
            coll.SortMode = DataGridViewColumnSortMode.NotSortable;
            coll.DefaultCellStyle.BackColor = Color.Pink;
            coll.HeaderText = "Зачёт " + zach_count.ToString();
            attend_grid.Columns.Add(coll);

            coll = new DataGridViewTextBoxColumn();
            coll.SortMode = DataGridViewColumnSortMode.NotSortable;
            coll.DefaultCellStyle.BackColor = Color.LightGray;
            coll.HeaderText = "Прим. к зач. " + zach_count.ToString();
            attend_grid.Columns.Add(coll);
        }

        private void attend_grid_KeyPress(object sender, KeyPressEventArgs e)
        {
        }

        private void attend_grid_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode != Keys.Space)
            {
                return;
            }


            int row = attend_grid.CurrentCell.RowIndex;
            int col = attend_grid.CurrentCell.ColumnIndex;

            if (row == -1 || col != 2) return;

            int i = row;
            DataRow r = attend_set.Rows[i];

            string att_name = attend_grid[2, row].Value.ToString();


            string cmd = "update attend set attend_id = @AID where id = @ID";

            int aid = 0;

            if (att_name.Contains("опред"))
            {
                aid = 2;
            }

            if (att_name.Contains("отсут"))
            {
                aid = 2;
            }

            if (att_name.Contains("присут"))
            {
                aid = 1;
            }


            SqlCommand scmd = new SqlCommand(cmd, main.global_connection);

            scmd.Parameters.Add("@AID", SqlDbType.Int).Value = aid;
            scmd.Parameters.Add("@ID", SqlDbType.Int).Value = r[0];

            try
            {
                scmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Непредвиденная ошибка при передаче данных. Повторите операцию позднее.\nПричина:\n" +
                    ex.Message, "Ошибка при сохранении данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (att_name.Contains("опред"))
            {
                attend_grid[2, i].Style.ForeColor = Color.Blue;
                attend_grid[2, i].Value = "присутствует";
                c_pris++;
            }

            if (att_name.Contains("отсут"))
            {
                attend_grid[2, i].Style.ForeColor = Color.Blue;
                attend_grid[2, i].Value = "присутствует";
                c_ots--;
                c_pris++;
            }

            if (att_name.Contains("присут"))
            {
                attend_grid[2, i].Style.ForeColor = Color.Red;
                attend_grid[2, i].Value = "отсутствует";
                c_ots++;
                c_pris--;
            }

            pris_label.Text = "Присутствуют: " + c_pris.ToString() + "   ";
            ots_label.Text = "Отсутствуют: " + c_ots.ToString() + "   ";
        }
        //public DataTable alien_set = new DataTable();
    }
}