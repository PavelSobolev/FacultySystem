using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;


namespace FSystem
{
    public partial class sprav_predmets : Form
    {
        public sprav_predmets()
        {
            InitializeComponent();
        }


        public string kaf_text = "",
            prepod_filter = "",
            grupa_filter = "",
            predmet_filter = "",
            actual_filter = " predmet.actual=1 and ",
            kaf_filter = "",
            semestr_filter = "";

        public bool from_clear = false;

        DataTable kaf_set, predmet_set;
        int predm_id = 0, row = 0;
        DataRow crow = null;

        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            actual.Checked = !actual.Checked;
            if (actual.Checked)
                actual_filter = " predmet.actual=1 and ";
            else
                actual_filter = "";
            fill_table();
        }

        public void fill_kafdera()
        {
            //загрузить кафедры, получить активную
            string selcom = "select id, name, kurs_id from grupa " +
                "where actual=1 and fakultet_id = " + 
                main.fakultet_id.ToString();

            main.global_adapter = new SqlDataAdapter(selcom,
                main.global_connection);

            kaf_set = new DataTable();

            main.global_adapter.Fill(kaf_set);

            foreach (DataRow dr in kaf_set.Rows)
            {
                kaf_list.Items.Add(dr[1]);
            }

            kaf_list.SelectedIndex = 0;

            stat_text.Text = "Выбрана группа: " + 
                kaf_set.Rows[kaf_list.SelectedIndex][1].ToString();            
            kaf_filter = " grupa.id = " + 
                kaf_set.Rows[kaf_list.SelectedIndex][0].ToString();
                
        }
        
        public AutoCompleteStringCollection acsc1 = 
            new AutoCompleteStringCollection();
        public AutoCompleteStringCollection acsc2 = 
            new AutoCompleteStringCollection();

        private void sprav_predmets_Load(object sender, EventArgs e)
        {
            fill_kafdera();
            kaf_filter = " grupa.id = " + 
                kaf_set.Rows[kaf_list.SelectedIndex][0].ToString();

            DataTable dt = new DataTable();

            try
            {
                SqlDataAdapter cmd = new SqlDataAdapter(
                    "select name, name_krat from predmet order by name", 
                    main.global_connection);
                cmd.Fill(dt);
            }
            catch (Exception exx)
            {
                MessageBox.Show("Ошибка при загрузке данных.  " + 
                    " Повторите операцию позднее.");
                return;
            }

            int i = 0;
            foreach (DataRow dr in dt.Rows)
            {             
                acsc1.Add(dr[0].ToString());
                acsc2.Add(dr[1].ToString());
                i++;
            }
            
            dt.Dispose();
        }

        private void kaf_list_SelectedIndexChanged(object sender, EventArgs e)
        {
            stat_text.Text = "Выбрана группа: " + 
                kaf_set.Rows[kaf_list.SelectedIndex][1].ToString();            
            kaf_filter = " grupa.id = " + 
                kaf_set.Rows[kaf_list.SelectedIndex][0].ToString();
            kaf_list.ToolTipText = 
                kaf_set.Rows[kaf_list.SelectedIndex][1].ToString();

            semestr_list.Items.Clear();

            int k = 2*(int)kaf_set.Rows[kaf_list.SelectedIndex][2];

            semestr_list.Items.Add("-");
            semestr_list.Items.Add((k-1).ToString());
            semestr_list.Items.Add(k.ToString());

            fill_table();

            semestr_list.SelectedIndex = 0;
        }

        private void toolStripComboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (from_clear) return;

            if (semestr_list.SelectedIndex == 0)
                semestr_filter = "";
            else
                semestr_filter = " and semestr = " + semestr_list.Text;

            EnDis_DropFilter();
            fill_table();
        }

        private void prepod_TextChanged(object sender, EventArgs e)
        {
            if (from_clear) return;
            //фильтр для препоавателя

            string txt = prepod.Text.Trim();
            Normalize(ref txt);

            if (txt.Length == 0)
                prepod_filter = "";
            else
            {              
                prepod_filter = " and prepod.fam like '" +
                    txt + "%' ";
            }

            EnDis_DropFilter();
            fill_table();
        }

        private void predmet_TextChanged_1(object sender, EventArgs e)
        {
            if (from_clear) return;

            //фильтр для предмета
            string txt = predmet.Text.Trim();
            Normalize(ref txt);
            
            if (txt.Length == 0)
                predmet_filter = "";
            else
            {                
                predmet_filter = " and predmet.name like '" +
                    txt + "%' ";
            }

            EnDis_DropFilter();
            fill_table();
        }

        private void grupa_TextChanged(object sender, EventArgs e)
        {
            if (from_clear) return;

            string txt = grupa.Text.Trim();
            Normalize(ref txt);
            
            //фильтр для предмета
            if (txt.Length == 0)
                grupa_filter = "";
            else
            {                
                grupa_filter = " and kafedra.name_krat like '" +
                    txt + "%' ";
            }

            EnDis_DropFilter();
            fill_table();
        }

        public void EnDis_DropFilter()
        {
            if (predmet.Text.Trim().Length == 0 
                && prepod.Text.Trim().Length == 0 &&
                grupa.Text.Trim().Length == 0 
                && semestr_list.Text=="-")
                drop_filter.Enabled = false;
            else
                drop_filter.Enabled = true;
        }


        /// <summary>
        /// удалить из строки метасимволы запросов '
        /// </summary>
        /// <param name="str"></param>
        public void Normalize(ref string str)
        {            
            while(str.Contains("'"))
            {
                int pos = str.IndexOf("'");
                str = str.Remove(pos, 1);
                if (str.Length == 0) break;               
            }           
        }

        public void fill_table()
        {
            Application.DoEvents();

            string q = "select " + 
                " predmet.id, predmet.name, predmet.name_krat, " + //0,1,2
                " 'prepod' = prepod.fam  + ' ' + left(prepod.im,1)  +  " + 
                " '. ' + left(prepod.ot,1) + '.', " + //3
                " grupa.id, grupa.name, " + //4,5
                " kurs.nomer, semestr, " + //6,7
                " predmet.kafedra_id, predmet.prepod_id, " + //8,9
                " predmet.actual, " + //10
                " case  " +   //11
                " when predmet.delenie=1 then 'есть' " +
                " when predmet.delenie=0 then 'нет' " +
                " end, predmet.delenie,  " + 
                " predmet.fakultet_id, predmet.type_id, " + //12, 13, 14
                " kafedra.name_krat, " + //15
                " kredit " + // 16
                " from predmet " + 
                " join prepod on prepod.id=predmet.prepod_id " +
                " join kafedra on kafedra.id=predmet.kafedra_id " +
                " join kurs on kurs.id=predmet.kurs_id " +
                " join grupa on grupa.id=predmet.grupa_id " +
                " where " + 
                actual_filter + kaf_filter + prepod_filter + predmet_filter 
                + grupa_filter + semestr_filter +
                " order by predmet.name, grupa.name, prepod, kurs.nomer, semestr ";

            SqlDataAdapter sda = new SqlDataAdapter(q, main.global_connection);
            predmet_set = new DataTable();
            sda.Fill(predmet_set);            

            table.Rows.Clear();
            table.Columns.Clear();

            table.Columns.Add("number", "№"); //0
            table.Columns[0].Width = 30;          
            table.Columns.Add("predm", "Предмет"); //1
            table.Columns[1].Width = 350;
            table.Columns.Add("prepod", "Преподаватель");//2
            table.Columns.Add("grupa", "Кафедра"); //3
            table.Columns.Add("sem", "Семестр"); //4
            table.Columns.Add("del", "Кредиты"); //5            

            int i = 1;
            
            foreach (DataRow row in predmet_set.Rows)
            {
                object[] par = new object[6] { i, row[1], row[3], row[15], row[7], 
                    row[16]};

                table.Rows.Add(par);

                i++;
            }

            if (predmet_set.Rows.Count > 0)
                table.Rows[0].Selected = true;

        }

        /// <summary>
        /// сбросить фильтр
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            from_clear = true;

            predmet_filter = "";
            grupa_filter = "";
            prepod_filter = "";
            semestr_filter = "";

            prepod.Clear();
            grupa.Clear();
            predmet.Clear();
            semestr_list.SelectedIndex = 0;

            fill_table();

            from_clear = false;

            drop_filter.Enabled = false;
        }

        private void table_CellMouseClick(object sender, 
            DataGridViewCellMouseEventArgs e)
        {
            crow = null;
            if (table.Rows.Count == 0) return;
           
            //CurrentRow

            DataGridViewCell cell = table.SelectedCells[0];
            row = cell.RowIndex;            
            crow = predmet_set.Rows[row];

            predm_id = (int)crow[0];

        }

        private void table_CellMouseDoubleClick(object sender, 
            DataGridViewCellMouseEventArgs e)
        {
            crow = null;
            if (table.Rows.Count == 0) return;

            DataGridViewCell cell = table.SelectedCells[0];
            row = cell.RowIndex;
            crow = predmet_set.Rows[row];

            predm_id = (int)crow[0];
            toolStripButton2_Click(sender, new EventArgs());
        }

        private void скопироватьВБуферОбменаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Clipboard.Clear();
            try
            {
                main.CopyGridToClipBoard(table);
                //Clipboard.SetDataObject(table.GetClipboardContent());
            }
            catch (Exception ex)
            {
                ;
            }
        }


        //добабвление нового предмета
        private void toolStripButton1_Click(object sender, EventArgs e)
        {            
            predmet_edit pe = new predmet_edit();
            
            pe.kaf_id = (int)kaf_set.Rows[kaf_list.SelectedIndex][0];            
            pe.full_name.AutoCompleteCustomSource = acsc1;
            pe.krat_name.AutoCompleteCustomSource = acsc2;
            pe.grup_id = (int)kaf_set.Rows[kaf_list.SelectedIndex][0];
            pe.kurs_id = Convert.ToInt32(kaf_set.Rows[kaf_list.SelectedIndex][2]);
            pe.KreditUpDown.Value = 1;
            pe.edit = false;

            DialogResult pe_res = pe.ShowDialog();           

            if (pe_res != DialogResult.OK) return;

            /*2+ - 33,3 
            3 от 50%, 
            4 - 67,7
            5 - 83,3 */

            //1. добавить запись в таблицу predmet и получить id
            string q = "insert into predmet (" + 
                " prepod_id,name,      fakultet_id, kurs_id, grupa_id," + 
                " semestr,  name_krat, kafedra_id,  actual,  delenie, type_id, kredit ) " +
                " values ( " + 
                " @PRID, " +
                " @NAME, " +
                " @FID, " +
                " @KID, " +
                " @GID, " +
                " @SID, " +
                " @NK, " +
                " @KFID, " +
                " @ACT, " +
                " @DEL, " +
                " @TYPE, " + 
                " @KREDIT)";

            SqlCommand cmd = new SqlCommand(q, main.global_connection);
            cmd.Parameters.Add("@PRID", SqlDbType.Int).Value = pe.prepod_id;
            cmd.Parameters.Add("@NAME", SqlDbType.NVarChar).Value = pe.full_name.Text;
            cmd.Parameters.Add("@FID", SqlDbType.Int).Value = pe.fakultet_id;
            cmd.Parameters.Add("@KID", SqlDbType.Int).Value = pe.kurs_id;
            cmd.Parameters.Add("@GID", SqlDbType.Int).Value = pe.grup_id;
            cmd.Parameters.Add("@SID", SqlDbType.Int).Value = pe.semestr.Value;
            cmd.Parameters.Add("@NK", SqlDbType.NVarChar).Value = pe.krat_name.Text;
            cmd.Parameters.Add("@KFID", SqlDbType.Int).Value = pe.kaf_id;
            cmd.Parameters.Add("@ACT", SqlDbType.Bit).Value = pe.checkBox1.Checked;
            cmd.Parameters.Add("@DEL", SqlDbType.Bit).Value = pe.delenie;
            cmd.Parameters.Add("@TYPE", SqlDbType.Int).Value = pe.type_id;
            cmd.Parameters.Add("@KREDIT", SqlDbType.Int).Value = pe.KreditUpDown.Value;
                                 
            try
            {
                cmd.ExecuteNonQuery();    
            }
            catch (Exception exx)
            {
                MessageBox.Show("Ошибка записи данных. Повторите операцию позднее.",
                    "Ошибка передачи данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            //получить ID новой записи
            cmd = new SqlCommand("select @@identity", main.global_connection);
            int newid = 0;

            

            try
            {
                newid = Convert.ToInt32(cmd.ExecuteScalar());
            }
            catch (Exception exx)
            {
                MessageBox.Show("Ошибка записи данных. Повторите операцию позднее.",
                    "Ошибка передачи данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            
            //2. удалить все записи о предмете из таблицы vidzan_predmet
            q = "delete from vidzan_predmet where predmet_id = @ID";

            cmd = new SqlCommand(q, main.global_connection);
            cmd.Parameters.Add("@ID", SqlDbType.Int).Value = newid;            

            try
            {
                cmd.ExecuteNonQuery();
            }
            catch (Exception exx)
            {
                MessageBox.Show("Ошибка записи данных. Повторите операцию позднее.",
                    "Ошибка передачи данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }


            //3. добавить записи в таблицу vidzan_predmet
            foreach (ListViewItem lv in pe.vid_view.Items)
            {
                int vid = Convert.ToInt32(lv.Tag);
                double ch = Convert.ToDouble(lv.SubItems[1].Text);


                q = "insert into vidzan_predmet (vidzan_id, predmet_id, kol_chas) " +
                    " values (" +
                    " @VID, @PID, @KOL )";

                cmd = new SqlCommand(q, main.global_connection);
                cmd.Parameters.Add("@PID", SqlDbType.Int).Value = newid;
                cmd.Parameters.Add("@VID", SqlDbType.Int).Value = vid;
                cmd.Parameters.Add("@KOL", SqlDbType.Float).Value = ch;                                
                
                try
                {
                    cmd.ExecuteNonQuery();
                }
                catch (Exception exx)
                {
                    MessageBox.Show("Ошибка записи данных. Повторите операцию позднее.",
                        "Ошибка передачи данных",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            fill_table();
        }

        //обновление данных
        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            if (table.Rows.Count==0) return;

            if (crow == null) return;
            
            predmet_edit pe = new predmet_edit();

            pe.full_name.AutoCompleteCustomSource = acsc1;
            pe.krat_name.AutoCompleteCustomSource = acsc2;

            pe.Text = "Сведения о предмете: " + crow[1].ToString();
            pe.pred_id = (int)crow[0];
            pe.full_name.Text = crow[1].ToString();
            pe.krat_name.Text = crow[2].ToString();
            pe.grup_id = (int)crow[4];
            pe.prepod_id = (int)crow[9];
            pe.kaf_id = (int)crow[8];
            pe.kurs_id = (int)crow[6];
            pe.semestr_id = (int)crow[7];
            pe.semestr.Value = pe.semestr_id;
            pe.semestr.Maximum = pe.kurs_id * 2;
            pe.semestr.Minimum = pe.kurs_id * 2 - 1;                        
            pe.delenie = (bool)crow[12];
            pe.delenie_list.SelectedIndex = (!pe.delenie) ? 0 : 1;
            pe.type_id = (int)crow[14];
            pe.checkBox1.Checked = (bool)crow[10];
            pe.fakultet_id = (int)crow[13];
            pe.edit = true;
            int kred = Convert.ToInt32(crow[16]);
            if (kred==0) kred = 1;            
            pe.KreditUpDown.Value = kred;

            DialogResult pe_res = pe.ShowDialog();

            if (pe_res != DialogResult.OK) return;

            //1. обновить запись  в таблице predmet
            string q = "update predmet set " +
                " prepod_id = @PRID, " + 
                " name = @NAME, " + 
                " fakultet_id = @FID, " + 
                " kurs_id = @KID, " + 
                " grupa_id = @GID, " + 
                " semestr = @SID, " + 
                " name_krat = @NK, " + 
                " kafedra_id = @KFID, " +
                " actual = @ACT, " + 
                " delenie = @DEL, " + 
                " type_id = @TYPE, " +
                " kredit = @KRED " + 
                " where id = @ID";

            SqlCommand cmd = new SqlCommand(q, main.global_connection);
            cmd.Parameters.Add("@PRID", SqlDbType.Int).Value = pe.prepod_id;
            cmd.Parameters.Add("@NAME", SqlDbType.NVarChar).Value = pe.full_name.Text;
            cmd.Parameters.Add("@FID", SqlDbType.Int).Value = pe.fakultet_id;
            cmd.Parameters.Add("@KID", SqlDbType.Int).Value = pe.kurs_id;
            cmd.Parameters.Add("@GID", SqlDbType.Int).Value = pe.grup_id;
            cmd.Parameters.Add("@SID", SqlDbType.Int).Value = pe.semestr.Value;
            cmd.Parameters.Add("@NK", SqlDbType.NVarChar).Value = pe.krat_name.Text;
            cmd.Parameters.Add("@KFID", SqlDbType.Int).Value = pe.kaf_id;
            cmd.Parameters.Add("@ACT", SqlDbType.Bit).Value = pe.checkBox1.Checked;
            cmd.Parameters.Add("@DEL", SqlDbType.Bit).Value = pe.delenie;
            cmd.Parameters.Add("@TYPE", SqlDbType.Int).Value = pe.type_id;
            cmd.Parameters.Add("@KRED", SqlDbType.Int).Value = pe.KreditUpDown.Value;
            cmd.Parameters.Add("@ID", SqlDbType.Int).Value = pe.pred_id;
            
            try
            {
                cmd.ExecuteNonQuery();
            }
            catch (Exception exx)
            {
                MessageBox.Show("Ошибка записи данных. Повторите операцию позднее.",
                    "Ошибка передачи данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }


            //2. удалить все записи о предмете из таблицы vidzan_predmet
            q = "delete from vidzan_predmet where predmet_id = @ID";

            cmd = new SqlCommand(q, main.global_connection);            
            cmd.Parameters.Add("@ID", SqlDbType.Int).Value = pe.pred_id;

            try
            {
                cmd.ExecuteNonQuery();
            }
            catch (Exception exx)
            {
                MessageBox.Show("Ошибка записи данных. Повторите операцию позднее.",
                    "Ошибка передачи данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }


            //3. добавить записи в таблицу vidzan_predmet
            foreach (ListViewItem lv in pe.vid_view.Items)
            {
                int vid = Convert.ToInt32(lv.Tag);
                double ch = Convert.ToDouble(lv.SubItems[1].Text);


                q = "insert into vidzan_predmet (vidzan_id, predmet_id, kol_chas) " +
                    " values (" +
                    " @VID, @PID, @KOL )";

                cmd = new SqlCommand(q, main.global_connection);                
                cmd.Parameters.Add("@PID", SqlDbType.Int).Value = pe.pred_id;
                cmd.Parameters.Add("@VID", SqlDbType.Int).Value = vid;
                cmd.Parameters.Add("@KOL", SqlDbType.Float).Value = ch;

                try
                {
                    cmd.ExecuteNonQuery();
                }
                catch (Exception exx)
                {
                    MessageBox.Show("Ошибка записи данных. Повторите операцию позднее.",
                        "Ошибка передачи данных",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            fill_table();
        }

    }
}