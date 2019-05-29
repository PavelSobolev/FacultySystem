using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace FSystem
{
    public partial class sprav_grupa : Form
    {
        public sprav_grupa()
        {
            InitializeComponent();
        }

        public DataTable grupa_set = null;
        public int num = 0;
        public int sg = 0;

        public void fill_grups()
        {
            string cmd = "select 'ID' = grupa.id, 'SID' = specialnost_id, " + 
                " 'Группа'= grupa.name, " + 
                " 'Специальность' = specialnost.kod + ' - ' + specialnost.name, " +
                " [Вып. кафедра] = kafedra.name," + 
                " 'Курс' = kurs_id, " + 
                " [Группа существует] =  grupa.actual, " + 
                " [Отображается в раписании] = show_in_grid, " + 
                " subgrups " +
                " from grupa" + 
                " join specialnost on specialnost.id = grupa.specialnost_id " + 
                " join kafedra on specialnost.kafedra_id = kafedra.id " + 
                " where kafedra.fakultet_id = " + main.fakultet_id.ToString() + 
                " order by specialnost_id, grupa.kurs_id ";

            SqlDataAdapter sda = new SqlDataAdapter(cmd, main.global_connection);
            grupa_set = new DataTable();
            sda.Fill(grupa_set);
            grupa_grid.DataSource = grupa_set;

            grupa_grid.Columns[2].SortMode = DataGridViewColumnSortMode.NotSortable;
            grupa_grid.Columns[2].Width = 70;
            
            grupa_grid.Columns[3].SortMode = DataGridViewColumnSortMode.NotSortable;
            grupa_grid.Columns[3].Width = 275;
            
            grupa_grid.Columns[4].SortMode = DataGridViewColumnSortMode.NotSortable;
            grupa_grid.Columns[4].Width = 250;
            
            grupa_grid.Columns[5].SortMode = DataGridViewColumnSortMode.NotSortable;
            grupa_grid.Columns[5].Width = 50;
            grupa_grid.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            grupa_grid.Columns[6].SortMode = DataGridViewColumnSortMode.NotSortable;
            grupa_grid.Columns[6].Width = 90;

            grupa_grid.Columns[7].SortMode = DataGridViewColumnSortMode.NotSortable;
            grupa_grid.Columns[7].Width = 90;

            grupa_grid.Columns[1].Visible = false;
            grupa_grid.Columns[0].Visible = false;
            grupa_grid.Columns[8].Visible = false;
            //grupa_grid.Columns[9].Visible = false;

            if (grupa_grid.Rows.Count>0)
            grupa_grid.Rows[0].Selected = true;
        }

        private void sprav_grupa_Load(object sender, EventArgs e)
        {
            fill_grups();
        }

        private void grupa_grid_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex == -1) return;

            if (grupa_grid.Rows.Count > 0)
                num = e.RowIndex;
            sg = (int)grupa_set.Rows[num][8];
        }

        //редактировать группу
        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            grupa_edit ge = new grupa_edit();

            ge.spec_id = (int)grupa_set.Rows[num][1];           
            ge.grupa_id = (int)grupa_set.Rows[num][0];
            ge.exists_box.Checked = (bool)grupa_set.Rows[num][6];
            ge.show_box.Checked = (bool)grupa_set.Rows[num][7];
            ge.nomer_kurs = (int)grupa_set.Rows[num][5];
            ge.name = grupa_set.Rows[num][2].ToString();

            DialogResult dres;

            string tmpname = "";

            do
            {
                if (tmpname.Trim().Length > 0)
                {
                    dres = ge.ShowDialog();
                    ge.name = tmpname;
                }
                else
                    dres = ge.ShowDialog();
                
                tmpname = ge.name;
                
                if (dres == DialogResult.Cancel)
                {
                    ge.Dispose();
                    return;
                }
            }
            while(dres!=DialogResult.OK);
                      
            //сохранить изменения
            string cmd = "update grupa set " +
                " name = @NM, specialnost_id = @SID, kafedra_id = @KID,  " +
                " kurs_id = @KURS, actual = @ACT, subgrups = @SG, " +
                " show_in_grid = @SIG where id = @GRID";
            SqlCommand command = new SqlCommand(cmd, main.global_connection);

            command.Parameters.Add("@NM", SqlDbType.NVarChar).Value = ge.grupa_name_box.Text.Trim();
            command.Parameters.Add("@SID", SqlDbType.Int).Value = ge.spec_id;
            command.Parameters.Add("@KID", SqlDbType.Int).Value = ge.kaf_id;
            command.Parameters.Add("@KURS", SqlDbType.Int).Value = ge.kurs_list.Value;
            command.Parameters.Add("@ACT", SqlDbType.Bit).Value = ge.exists_box.Checked;
            command.Parameters.Add("@SG", SqlDbType.Int).Value = grupa_set.Rows[num][8];
            command.Parameters.Add("@SIG", SqlDbType.Bit).Value = ge.show_box.Checked;
            command.Parameters.Add("@GRID", SqlDbType.Int).Value = grupa_set.Rows[num][0];

            try
            {
                command.ExecuteNonQuery();
            }
            catch (Exception exx)
            {
                MessageBox.Show("Ошибка при передаче данных. Повторите операцию позднее.",
                    "Ошибка данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                ge.Dispose();
                return;
            }

            ge.Dispose();
            fill_grups();
        }

        //добавбить новую группу
        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            grupa_edit ge = new grupa_edit();

            DialogResult dres;

            do
            {
                dres = ge.ShowDialog();
                if (dres == DialogResult.Cancel)
                {
                    ge.Dispose();
                    return;
                }
            }
            while (dres != DialogResult.OK);

            //сохранить изменения
            string cmd = 
            " insert into grupa (name,specialnost_id,kafedra_id,kurs_id,actual,subgrups,show_in_grid, fakultet_id) " +
            " values ( @NM, @SID, @KID, @KURS, @ACT, @SG, @SIG, @FID )";

            SqlCommand command = new SqlCommand(cmd, main.global_connection);

            command.Parameters.Add("@NM", SqlDbType.NVarChar).Value = ge.grupa_name_box.Text.Trim();
            command.Parameters.Add("@SID", SqlDbType.Int).Value = ge.spec_id;
            command.Parameters.Add("@KID", SqlDbType.Int).Value = ge.kaf_id;
            command.Parameters.Add("@KURS", SqlDbType.Int).Value = ge.kurs_list.Value;
            command.Parameters.Add("@ACT", SqlDbType.Bit).Value = ge.exists_box.Checked;
            command.Parameters.Add("@SG", SqlDbType.Int).Value = grupa_set.Rows[num][8];
            command.Parameters.Add("@SIG", SqlDbType.Bit).Value = ge.show_box.Checked;
            command.Parameters.Add("@FID", SqlDbType.Int).Value = main.fakultet_id;


            command.ExecuteNonQuery();

            try
            {
                ;//command.ExecuteNonQuery();
            }
            catch (Exception exx)
            {
                MessageBox.Show("Ошибка при передаче данных. Повторите операцию позднее.",
                    "Ошибка данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                ge.Dispose();
                return;
            }

            ge.Dispose();
            fill_grups();
        }

        private void grupa_grid_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            toolStripButton2_Click(sender, new EventArgs());
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            if (grupa_grid.CurrentCell.RowIndex < 0) return;
            
            sprav_student ss = new sprav_student();
            ss.gr_id = (int)grupa_set.Rows[grupa_grid.CurrentCell.RowIndex][0];
            ss.ShowDialog();
            ss.Dispose();
        }
    }
}