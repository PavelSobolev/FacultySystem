using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace FSystem
{
    public partial class sprav_student : Form
    {
        public sprav_student()
        {
            InitializeComponent();
        }

        /// <summary>
        /// список групп
        /// </summary>
        public DataTable grupa_set = null;

        /// <summary>
        /// студенты
        /// </summary>
        public DataTable student_set = null;

        /// <summary>
        /// идентификатор активной группы
        /// </summary>
        public int gr_id = 0;

        /// <summary>
        /// фильтр группы
        /// </summary>
        public string grupa_filter = "";

        public DataTable RatingTable;

        /// <summary>
        /// заполнить таблицу студентов
        /// </summary>
        public void fill_students()
        {
            student_grid.Rows.Clear();
            pictureBox1.Image = null;

            string q = "select student.id, fam, im, ott = isnull(ot,''), student.actual,  " +    // 0 1 2 3 4 
                " zach_kn_number, subgr_nomer, " +  // 5 6
                " case when student.actual=1 then 'учится' else 'отчислен' end, " + // 7
                " phone, cell_phone, sex, isnull(born_date,getdate()), " +  // 8 9 10 11
                " work_place, graduated_from, isnull(graduated_date,getdate()-365) as draddate, " + //12 13 14 
                " mother_info, father_info, " +  // 15 16
                " prik = isnull(prikaz_nom_zach,'-'), isnull(military_id,1) as army, " + // 17 18
                " gkurs = grupa.kurs_id, " +  // 19
                " isdatepr = cast(isnull(start_date,-1) as int), start_date, end_date, " + // 20 21 22
                " statid = student_status.id, stattxt = student_status.text, " + // 23 24
                " fio = dbo.GetStudentFIOByID(student.id), photo " +   //25 26
                " from student "  +
                " join grupa on grupa.id = student.gr_id " + 
                " join student_status on student_status.id = student.status_id " + 
                " where fam<>'-'  " + 
                //" and student.actual = 1 " +
                " and status_id =  " + StatusTable.Rows[StatusCombo.SelectedIndex][0].ToString() + "  " + 
                grupa_filter +
                " order by fam, im, ot ";
            
            SqlCommand cmd = new SqlCommand(q, main.global_connection);
            SqlDataAdapter sda = new SqlDataAdapter(cmd);
            student_set = new DataTable();
            sda.Fill(student_set);

            if (mrs == 1)
            {
                RatingTable = new DataTable();
                q = "select * from dbo.TGetGrupRating(" +
                    grupa_set.Rows[grupa_list.SelectedIndex][0].ToString() + ")";
                (new SqlDataAdapter(q, main.global_connection)).Fill(RatingTable);
            }

            // проверка статуса студента
            if (student_set.Rows.Count == 0)
            {
                //StatusCombo.Enabled = false;
                //StatusCombo.SelectedIndex = 0;
                return;
            }
            else
                StatusCombo.Enabled = true;

            if (mrs != 0)
            {
                student_grid.Columns[5].Visible = true;
            }
            else
            {
                student_grid.Columns[5].Visible = false;
            }

            int i = 0;
            foreach (DataRow dr in student_set.Rows)
            {
                student_grid.Rows.Add();
                student_grid.Rows[i].Tag = dr[0]; // ид студента

                student_grid[0, i].Value = i + 1; // номер

                student_grid[1, i].Value = dr[1] + " " + dr[2] + " " + dr[3]; // фио

                student_grid[2, i].Value = dr["stattxt"]; // статус
                student_grid[2, i].Tag = dr["statid"];
                
                student_grid[3, i].Value = dr[5];  //номер зк

                if (dr[8].ToString().Length != 0)
                    student_grid[4, i].Value = dr[8]; // телефон
                else
                    student_grid[4, i].Value = "не задан";

                if (mrs == 1)
                {
                    DataRow[] rating = RatingTable.Select("stid=" + dr[0].ToString());
                    if (rating.Length > 0)
                        student_grid[5, i].Value = rating[0][2].ToString() + " (" + rating[0][1].ToString() + " б.)";  // рейтинг
                    else
                        student_grid[5, i].Value = "нет данных";
                }

                if ((int)dr[6] == 0)
                    student_grid[6, i].Value = "не выбрана";
                else
                    student_grid[6, i].Value = (int)dr[6];

                //student_grid[7, i].Value = dr["photo"];

                i++;
            }
        }

        public int mrs = 0;

        /// <summary>
        /// загрузить список групп
        /// </summary>
        public void fill_grupa()
        {
            //загрузить группы, получить активную
            string selcom = "select grupa.id,  " + 
                " grupa.name, kurs_id, specialnost.srok, mrs " +
                " from grupa " +
                " join specialnost on  " + 
                " grupa.specialnost_id = specialnost.id " +                 
                " where actual=1 and fakultet_id = " +
                main.fakultet_id.ToString();

            main.global_adapter = new SqlDataAdapter(selcom,
                main.global_connection);

            grupa_set = new DataTable();

            main.global_adapter.Fill(grupa_set);


            grupa_list.Items.Clear();
            foreach (DataRow dr in grupa_set.Rows)
            {
                grupa_list.Items.Add(dr[1]);
            }

            if (gr_id == 0)
                grupa_list.SelectedIndex = 0;
            else
                grupa_list.SelectedIndex = GetPosById(grupa_set, gr_id);            
        }

        private int GetPosById(DataTable box, int id)
        {
            int i = 0;

            foreach (DataRow item in box.Rows)
            {
                if ((int)item[0] == id)
                {
                    return i;
                }
                i++;
            }

            return -1;
        }

        DataRow c_gruppa = null;

        private void grupa_list_SelectedIndexChanged(object sender, EventArgs e)
        {
            int i = grupa_list.SelectedIndex;

            mrs = Convert.ToInt32(grupa_set.Rows[i][4]);

            c_gruppa = grupa_set.Rows[i];

            gr_id = (int)c_gruppa[0];
            grupa_filter = "and gr_id = " + gr_id.ToString();
            stat_text.Text = "Выбрана группа: " + c_gruppa[1].ToString();

            // ---------------------------------------------------
            for (i = 1; i <= 300; i++)
            {
                start_number.Items.Add(i);
            }
            start_number.SelectedIndex = 0;
            fakult_prefix.Text = main.fakultet_prfix;


            //номер курса
            int srok = (int)c_gruppa[3];
            int kurs = (int)c_gruppa[2];
            for (i = 1; i <= srok; i++)
            {
                kurs_list.Items.Add(i);
            }
            kurs_list.SelectedIndex = kurs - 1;

            int current_year = main.starts[0].Year; //год начала учебного года
            int y = current_year - kurs; //номер года организации группы
            int y1 = current_year - 10; //номер минимального года списка
            int y2 = main.ends[main.starts.Count - 1].Year;

            for (i = y1; i <= y2; i++)
            {
                year.Items.Add(i);
            }

            year.SelectedIndex = y - y1 + 1;            

            fill_students();
        }

        /// <summary>
        /// заполнить список гупп
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void sprav_student_Load(object sender, EventArgs e)
        {
            //заполнить статусы
            FillStatus();

            //заполнить группы
            fill_grupa();
        }

        /// <summary>
        /// таблица с перечнем статусов студентов
        /// </summary>
        public DataTable StatusTable;

        /// <summary>
        /// заполнить список статусов студентов
        /// </summary>
        public void FillStatus()
        {
            string sql = string.Format("select id, text, code from student_status order by id");
            StatusTable = new DataTable();
            (new SqlDataAdapter(sql, main.global_connection)).Fill(StatusTable);
            StatusCombo.Items.Clear();
            foreach (DataRow dr in StatusTable.Rows)
            {
                StatusCombo.Items.Add(dr[1].ToString());
            }
            StatusCombo.SelectedIndex = 0;
        }


        private void sablon_button_Click_1(object sender, EventArgs e)
        {
            if (sablon_button.Checked)
            {
                start_number.Enabled = false;
                fakult_prefix.Enabled = false;
                kurs_list.Enabled = false;
                year.Enabled = false;
                apply_sablon.Enabled = false;
            }
            else
            {
                start_number.Enabled = true;
                fakult_prefix.Enabled = true;
                kurs_list.Enabled = true;
                year.Enabled = true;
                apply_sablon.Enabled = true;
            }

            sablon_button.Checked = !sablon_button.Checked;
        }        

        private void apply_sablon_Click(object sender, EventArgs e)
        {
            int start = start_number.SelectedIndex + 1;
            string yeaR = year.Text.Substring(2, 2);
            string sql = "";

            for (int i = 0; i < student_grid.Rows.Count; i++)
            {
                student_grid[5, i].Value = start.ToString() + "-" +
                    fakult_prefix.Text + "-" + kurs_list.Text + "-" + 
                    yeaR;

                sql += "update student set zach_kn_number = '" + 
                    student_grid[5, i].Value.ToString() + "'" + 
                    " where id = " + student_set.Rows[i][0].ToString() + "; ";

                start++;
            }

            SqlCommand cmd = new SqlCommand(sql, main.global_connection);
           
            try
            {
                cmd.ExecuteNonQuery();
            }
            catch(Exception exx)
            {
                return;
            }

            fill_grupa();
        }

        private void student_grid_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex != 3) return;

            string newval = string.Empty;

            if (student_grid.CurrentRow.Cells[3].Value != null)
                newval = student_grid.CurrentRow.Cells[3].Value.ToString();
            else
            {
                student_grid.CurrentRow.Cells[3].Value = curval;
                return;
            }


            if (newval.Trim().Length == 0)
            {
                student_grid.CurrentRow.Cells[3].Value = curval;
                return;
            }

            if (newval == curval) return;

            bool res = save_field("zach_kn_number", newval, true);

            if (!res)
            {
                MessageBox.Show("Введённый Вами номер зачётной книжки уже используется. " + 
                    "Сохранение невозможно.",
                    "Отказ операции", MessageBoxButtons.OK, MessageBoxIcon.Error);
                student_grid.CurrentRow.Cells[3].Value = curval;
            }
        }

        private void student_grid_CellMouseUp(object sender, DataGridViewCellMouseEventArgs e)
        {            
            if (e.ColumnIndex == 6 && e.RowIndex >= 0)
            {
                string c_str = student_grid[6, e.RowIndex].Value.ToString();

                if (c_str.Contains("выбран"))
                    student_grid[6, e.RowIndex].Value = "1";
                
                if (c_str=="1")
                    student_grid[6, e.RowIndex].Value = "2";

                if (c_str == "2")
                    student_grid[6, e.RowIndex].Value = "1";

                string sql = "update student set subgr_nomer = @SUBGR " +
                    " where id = @STUDID";
                SqlCommand cmd = new SqlCommand(sql, main.global_connection);
                cmd.Parameters.Add("@SUBGR", SqlDbType.Int).Value = 
                    student_grid[6, e.RowIndex].Value;
                cmd.Parameters.Add("@STUDID", SqlDbType.Int).Value = 
                    student_set.Rows[e.RowIndex][0];

                try
                {
                    cmd.ExecuteNonQuery();
                }
                catch (Exception exx)
                {
                    ;
                    return;
                }

                //fill_grupa();
                cmd.Dispose();
            }


            /*if (e.ColumnIndex == 4 && e.RowIndex >= 0)
            {
                string c_str = student_grid[4, e.RowIndex].Value.ToString();
                bool actual = false;

                if (c_str.Contains("учится"))
                {
                    student_grid[4, e.RowIndex].Value = "отчислен";
                    actual = false;
                }
                else
                {
                    student_grid[4, e.RowIndex].Value = "учится";
                    actual = true;
                }
                                

                //отправить изменения в БД
                string sql = "update student set actual = @ACT " +
                    " where id = @STUDID";

                SqlCommand cmd = new SqlCommand(sql, main.global_connection);
                cmd.Parameters.Add("@ACT", SqlDbType.Int).Value = actual;
                cmd.Parameters.Add("@STUDID", SqlDbType.Int).Value =
                    student_set.Rows[e.RowIndex][0];

                try
                {
                    cmd.ExecuteNonQuery();
                }
                catch (Exception exx)
                {
                    ;
                    return;
                }

                //fill_grupa();
                cmd.Dispose();
            }*/
        }

        //редактирование данных студента
        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            if (student_grid.Rows.Count == 0) return;

            int row = student_grid.CurrentCell.RowIndex;
            if (row < 0) return;

            DataRow rr = student_set.Rows[row];
            
            student_edit se = new student_edit();

            se.stud_id = Convert.ToInt32(rr[0]);
            se.grupa_id = Convert.ToInt32(grupa_set.Rows[grupa_list.SelectedIndex][0]);
            se.newstud = false;

            se.fam.Text = rr[1].ToString();
            se.im.Text = rr[2].ToString();
            se.ot.Text = rr[3].ToString();
            se.zach.Text = rr[5].ToString();
            se.fakult_str = fakult_prefix.Text;
            se.phone.Text = rr[8].ToString();
            se.email.Text = rr[9].ToString();
            se.work_place_box.Text = rr["work_place"].ToString();
            se.graduated_from_box.Text = rr["graduated_from"].ToString();            
            
            se.dateTimePicker1.MaxDate = main.starts[0];
            DateTime val = Convert.ToDateTime(rr["draddate"]);
            if (val >= main.starts[0])
            {
                val = val.AddMonths(-3);
            }
            se.dateTimePicker1.Value = val;
            
            if (se.zach.Text.Trim().Length == 0)
            {
                int start = start_number.SelectedIndex + 1;
                string yeaR = year.Text.Substring(2, 2);

                se.zach.Text = start.ToString() + "-" +
                fakult_prefix.Text + "-" + kurs_list.Text + "-" + yeaR;               
            }

            se.mother_box.Text = rr["mother_info"].ToString();
            se.father_box.Text = rr["father_info"].ToString();
            se.prikaz_box.Text = rr["prik"].ToString();

            se.status_id = Convert.ToInt32(rr["statid"]);
            
            if (Convert.ToInt32(rr["army"]) == 1)
                se.radioButton4.Checked = true;
            else
                se.radioButton3.Checked = true;

            if (Convert.ToInt32(rr["isdatepr"]) == -1)
            {
                int gkurs = Convert.ToInt32(rr["gkurs"]);
                DateTime strt = new DateTime(DateTime.Now.Year - gkurs, 9, 1, 12, 0, 0);
                DateTime endd = strt.AddYears(5);

                se.enter_date_box.Value = strt;
                se.end_date_box.Value = endd;
            }
            else
            {
                se.enter_date_box.Value = Convert.ToDateTime(rr["start_date"]);
                se.end_date_box.Value = Convert.ToDateTime(rr["end_date"]);
            }

            bool sex = (bool)rr[10];
            if (sex)
                se.male.Checked = true;
            else
                se.female.Checked = true;

            int kurs = (int)grupa_set.Rows[grupa_list.SelectedIndex][2];
     
            DateTime dt = new DateTime(DateTime.Now.Year, 1,1);
            DateTime db = Convert.ToDateTime(rr[11]);

            if (db.Date == main.server_date.Date)
                se.born_date.Value = dt.AddYears(-(18 + kurs));
            else
                se.born_date.Value = db;

            DialogResult rs = se.ShowDialog();
            fill_grupa();
            fill_students();
            //cmd.Dispose();
            se.Dispose();

            if (leave_rownum < student_grid.Rows.Count)
                student_grid.Rows[leave_rownum].Cells[0].Selected = true;
        }

        //добавить студента
        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            string sql = "INSERT INTO student " +
                      " (fam, im, ot, gr_id, sex) " +
                      " VALUES ('','','',@gr_id,1)";
            main.global_command = new SqlCommand(sql, main.global_connection);
            main.global_command.Parameters.Add("@gr_id", SqlDbType.Int).Value = 
                grupa_set.Rows[grupa_list.SelectedIndex][0];
            main.global_command.ExecuteNonQuery();

            sql = "select @@identity";
            main.global_command = new SqlCommand(sql, main.global_connection);            
            SqlDataReader id = main.global_command.ExecuteReader();
            id.Read();
            int st_id = Convert.ToInt32(id[0]);
            id.Close();

            student_edit se = new student_edit();
            se.grupa_id = Convert.ToInt32(grupa_set.Rows[grupa_list.SelectedIndex][0]);
            se.fakult_str = fakult_prefix.Text;
            se.stud_id = st_id;
            se.newstud = true;

            int kurs = (int)grupa_set.Rows[grupa_list.SelectedIndex][2];

            DateTime dt = new DateTime(DateTime.Now.Year, 1, 1);
            se.born_date.Value = dt.AddYears(-(18 + kurs));

            DialogResult rs = se.ShowDialog();

            if (rs != DialogResult.OK)
            {
                fill_students();
                se.Dispose();
                return;
            }

            //fill_grupa();
            fill_students();
            se.Dispose();
        }

        /// <summary>
        /// удалить из строки метасимволы запросов '
        /// </summary>
        /// <param name="str"></param>
        public string Normalize(string str)
        {
            while (str.Contains("'"))
            {
                int pos = str.IndexOf("'");
                str = str.Remove(pos, 1);
                if (str.Length == 0) break;
            }

            while (str.Contains(" "))
            {
                int pos = str.IndexOf(" ");
                str = str.Remove(pos, 1);
                if (str.Length == 0) break;
            }

            return str;
        }

        public string NormalizeLetters(string str)
        {
            foreach (char c in str)
            {
                if (!Char.IsLetter(c))
                {
                    if (c != '-')
                    {
                        int pos = str.IndexOf(c);
                        str = str.Remove(pos, 1);
                        if (str.Length == 0) break;
                    }
                }
            }

            return str;
        }


        int leave_rownum;
        private void student_grid_CellMouseDoubleClick(object sender, 
            DataGridViewCellMouseEventArgs e)
        {
            leave_rownum = student_grid.CurrentRow.Index;
            toolStripButton2_Click(sender, new EventArgs());
        }

        private void редактироватьДанныеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            toolStripButton2_Click(sender, new EventArgs());
        }

        private void перевестиВДругуюГруппуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            toolStripButton1_Click(sender, new EventArgs());
        }

        private void закрытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        string curval = "";
        private void student_grid_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            if (e.ColumnIndex == 3)
                curval = student_grid[e.ColumnIndex, e.RowIndex].Value.ToString();
        }

        /// <summary>
        /// сохранить знаачение поля в БД (для текущей строки таблицы студентов)
        /// </summary>
        /// <param name="fname">имя поля (должно иметь тип СТРОКА!)</param>
        /// <param name="fvalue">значение поля</param>
        /// <param name="unique">должно ли быть значение поля уникальным</param>
        /// <returns></returns>
        public bool save_field(string fname, string fvalue, bool unique)
        {
            string stud_id = student_grid.CurrentRow.Tag.ToString();
            string sql = string.Empty;

            if (unique) // проверить значени на уникальность
            {
                sql = "select id from student where " + fname + "=@val";
                main.global_command = new SqlCommand(sql, main.global_connection);
                main.global_command.Parameters.Add("@val", SqlDbType.NVarChar).Value = fvalue;
                DataTable t = new DataTable();
                (new SqlDataAdapter(main.global_command)).Fill(t);
                if (t.Rows.Count > 0) return false;
            }
            
            sql = "update student set " + fname + " = @val where id = @st_id";
            main.global_command = new SqlCommand(sql, main.global_connection);
            main.global_command.Parameters.Add("@val", SqlDbType.NVarChar).Value = fvalue;
            main.global_command.Parameters.Add("@st_id", SqlDbType.Int).Value = stud_id;

            bool res = true;

            try
            {
                main.global_command.ExecuteNonQuery();
            }
            catch (Exception exx)
            {
                res = false;
            }

            main.global_command.Dispose();
            return res;
        }

        //выбор фотографии и вывод успеваеомсти студента (только задолженности)
        private void student_grid_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            listBox1.Items.Clear();
            pictureBox1.Image = null;

            if (student_grid.CurrentRow.Tag == null) return;

            // фото
            object pic = student_set.Rows[student_grid.CurrentRow.Index]["photo"];
            try
            {
                pictureBox1.Image = new Bitmap(new MemoryStream((byte[])(pic)));
            }
            catch (Exception exx)
            {
                pictureBox1.Image = null;
            }

            //задолженность            
            string sql = "Select " +
                " inf = predmet.name_krat + ' - ' + vid_zan.krat_name + ' - ' + cast(predmet.semestr as nvarchar(2)) + ' сем. (оценка - ' + vid_otmetka.str_alias + ')' " +
                " from session  " +
                " join student on student.id = session.student_id   " +
                " join vid_otmetka on vid_otmetka.id = session.otmetka_id  " +
                " join predmet on predmet.id = session.predmet_id  " +
                " join vid_zan on vid_zan.id = session.vid_zan_id  " +
                " where student.id = " + student_grid.CurrentRow.Tag.ToString() +
                " and (vid_otmetka.id = 2 or vid_otmetka.id=6)   " +
                " and (vid_zan.is_kontrol=1) and vid_zan.id<=16 ";
            DataTable t = new DataTable();
            (new SqlDataAdapter(sql, main.global_connection)).Fill(t);
            if (t.Rows.Count == 0) return;

            listBox1.Items.Add("Имеются академ. задолженности");
            listBox1.Items.Add("---------------------------");
            foreach (DataRow d in t.Rows)
            {
                listBox1.Items.Add(d[0].ToString());
            }
        }

        // вывод студентов по статусу
        private void StatusCombo_SelectedIndexChanged(object sender, EventArgs e)
        {
            fill_students();
        }

        private void FilterGRUPAButton_Click(object sender, EventArgs e)
        {
            FilterGRUPAButton.Checked = !FilterGRUPAButton.Checked;

            if (FilterGRUPAButton.Checked)
            {
                FilterGRUPAButton.Text = "Учитывать группу при выборке";
                FilterGRUPAButton.Image = Properties.Resources.filter_data_16;
                grupa_filter = "and gr_id = " + grupa_set.Rows[grupa_list.SelectedIndex][0].ToString();
                grupa_list.Enabled = true;
                fill_students();
            }
            else
            {
                FilterGRUPAButton.Text = "НЕ учитывать группу при выборке";
                FilterGRUPAButton.Image = Properties.Resources.cancel;
                grupa_filter = string.Empty;
                grupa_list.Enabled = false;
                fill_students();
            }
        }
    }
}