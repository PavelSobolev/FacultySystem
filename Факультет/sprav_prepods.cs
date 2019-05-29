using System;
using System.Data;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace FSystem
{
    public partial class sprav_prepods : Form
    {
        public sprav_prepods()
        {
            InitializeComponent();
        }

        private void закрытьОкноToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
        }

        public static DataTable kaf_set = null;
        public DataTable prep_set = null;
        public static string DefaultPath = Environment.GetFolderPath(
            Environment.SpecialFolder.MyPictures) + "\\emptyface.jpg"; 
        
        private void sprav_prepods_Load(object sender, EventArgs e)
        {
            //загрузить кафедры, получить активную
            string selcom = "select id, name_krat,  " + 
                " name, zav_kaf_id from kafedra " + 
                "where actual=1 " + 
                "order by priority";

            main.global_adapter = new SqlDataAdapter(selcom,
                main.global_connection);

            kaf_set = new DataTable();

            main.global_adapter.Fill(kaf_set);

            foreach (DataRow dr in kaf_set.Rows)
            {
                kaf_list.Items.Add(dr[1]);
            }

            kaf_list.SelectedIndex = 0;

            stat_text.Text = "Выбрана кафедра: " + 
                kaf_set.Rows[kaf_list.SelectedIndex][2].ToString();
            FilePhoto = DefaultPath;
        }

        public static int kafedra_sel = 0;

        private void kaf_list_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            int sel = kaf_list.SelectedIndex;
            kafedra_sel = sel;

            stat_text.Text = "Выбрана кафедра: " + kaf_set.Rows[kaf_list.SelectedIndex][2].ToString();
           
            string status = (toolStripButton3.Checked)?"and prepod.actual = 1":string.Empty;

            string selcom = " select " +
                " [ФИО] = fam + ' ' + im + ' '  + ot, " + //0
                " fam, " +  //1
                " im, " +  //2
                " ot," +   //3
                " dolznost_id," +  //4
                " zvanie_id," +  //5
                " stepen_id, " + //6
                " [Должность] = dolznost.name," +  //7
                " di = dolznost.id," + //8
                " [Звание] = zvanie.name," +  //9
                " zi = zvanie.id, " + //10
                " [Степень] = stepen.name," + //11
                " si = stepen.id, " + //12
                " prepod.id," + //13
                " pract = prepod.actual, " + //14
                " case " +
                "  when prepod.actual=1  THEN 'работает' " +
                " else " +
                "  'уволен' " +
                " end as [Статус], " + //15
                " [Телефон]=phone, " + //16
                " sex, " + //17
                " address " + //18
                " from prepod " +
                " join kafedra on kafedra.id=prepod.kafedra_id " +
                " join dolznost on dolznost.id=prepod.dolznost_id " +
                " join zvanie on zvanie.id=prepod.zvanie_id " +
                " join stepen on stepen.id=prepod.stepen_id " +
                " where kafedra_id = " + kaf_set.Rows[sel][0].ToString() + status +
                " order by fam";

            main.global_adapter = new SqlDataAdapter(selcom,
                main.global_connection);

            prep_set = new DataTable();

            main.global_adapter.Fill(prep_set);


            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();

            
            dataGridView1.Columns.Add("ФИО", "ФИО");
            dataGridView1.Columns[0].Width = 350;
            dataGridView1.Columns[0].SortMode = 
                DataGridViewColumnSortMode.NotSortable;
            
            dataGridView1.Columns.Add("Должность", "Должность");
            dataGridView1.Columns[1].SortMode = 
                DataGridViewColumnSortMode.NotSortable;
            
            dataGridView1.Columns.Add("Звание", "Звание");
            dataGridView1.Columns[2].SortMode = 
                DataGridViewColumnSortMode.NotSortable;
            
            dataGridView1.Columns.Add("Степень", "Степень");
            dataGridView1.Columns[3].SortMode = 
                DataGridViewColumnSortMode.NotSortable;
            
            dataGridView1.Columns.Add("Статус", "Статус");
            dataGridView1.Columns[4].SortMode = 
                DataGridViewColumnSortMode.NotSortable;
            
            dataGridView1.Columns.Add("Телефон", "Телефон");
            dataGridView1.Columns[5].SortMode = 
                DataGridViewColumnSortMode.NotSortable;

            int rownum = 0;
            foreach(DataRow dr in prep_set.Rows)
            {
                object[] par = new object[6] { dr[0], dr[7], dr[9], 
                    dr[11], dr[15], dr[16] };               

                dataGridView1.Rows.Add(par);

                if (dr[15].ToString() == "уволен")
                    dataGridView1[4, rownum].Style.ForeColor = 
                        Color.Green;

                if (dr[11].ToString().ToLower().Contains("кандидат") ||
                    dr[11].ToString().ToLower().Contains("доктор"))
                {
                    dataGridView1[3, rownum].Style.ForeColor = Color.Red;
                }

                if (dr[16].ToString() == string.Empty)
                    dataGridView1[5, rownum].Value = "не задан";

                rownum++;

            }
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            toolStripButton3.Checked = !toolStripButton3.Checked;
            kaf_list_SelectedIndexChanged_1(sender, new EventArgs());
        }

        /// <summary>
        /// добавление нового преподавателя
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            prepod_edit pe = new prepod_edit();

            pe.pictureBox1.Image = GetPhotoFromBD("prepod", 69);
            pe.status_box.Checked = true;
            pe.first = false;
            pe.zavkaf_id = 
                (int)kaf_set.Rows[sprav_prepods.kafedra_sel][3];


            FilePhoto = DefaultPath;                      

            pe.pictureBox1.Image.Save(FilePhoto, 
                System.Drawing.Imaging.ImageFormat.Jpeg);
            pe.FilePhoto = FilePhoto;

            DialogResult peres = pe.ShowDialog();

            if (peres == DialogResult.Cancel) return;
            
            string famn = pe.fam.Text.Trim();
            string imn = pe.im.Text.Trim();
            string otn = pe.ot.Text.Trim();
            int kafidn = pe.kaf_id;
            int dolzn = pe.dolz_id;
            int stepidn = pe.uch_id;
            int zvann = pe.zvan_id;
            bool actualn = pe.status_box.Checked;
            string addrn = pe.email.Text.Trim(); ;
            string phonen = pe.phone.Text.Trim();
            bool sexn = pe.male.Checked;

            bool newzav = 
                (pe.dolz_list.Text.ToLower().Trim().Contains(
                "заведующий"));
            
            //фото сохраняется отдельно

            SqlCommand cmd = new SqlCommand();
            cmd.Connection = main.global_connection;
            cmd.CommandText = "select prepod.actual,  " + 
                " kafedra.name from prepod " +
                " join kafedra on prepod.kafedra_id = kafedra.id " + 
                " where fam like @FAM and " +
                " im like @IM and " +
                " ot like @OT and " +
                " sex = @SEX ";
            cmd.Parameters.Add("@FAM", SqlDbType.NVarChar).Value = famn;
            cmd.Parameters.Add("@IM", SqlDbType.NVarChar).Value = imn;
            cmd.Parameters.Add("@OT", SqlDbType.NVarChar).Value = otn;
            cmd.Parameters.Add("@SEX", SqlDbType.Bit).Value = sexn;

            SqlDataAdapter sda = new SqlDataAdapter(cmd);

            DataTable oldprepods = new DataTable();
            sda.Fill(oldprepods);

            string message = "";

            if (oldprepods.Rows.Count > 0)
            {
                if (oldprepods.Rows.Count == 1)
                    message = "Обнаружен преподаватель с такой же " + 
                        " фамилией, именем и отчеством:\n\n";
                else
                    message = "Обнаружены преподаватели с аналогичными  " + 
                        " фамилией, именем и отчеством:\n\n";

                int i = 1;
                foreach (DataRow dr in oldprepods.Rows)
                {
                    string stat = ((bool)dr[0])?"работает":"уволен(а)";

                    message += i.ToString() + ". " +
                        famn + " " + imn + " " + otn + " [кафедра " +
                        dr[1].ToString() + ", статус - " + stat + "];\n";
                }

                message += "\n\nВыполнить сохранение введенной Вами информации?";                

                DialogResult dres = MessageBox.Show(message, 
                    "Выбор действия",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (dres==DialogResult.No)
                {
                    return;
                }                
            }        

            //добавить новую запись в БД
            cmd = new SqlCommand();
            cmd.Connection = main.global_connection;

            cmd.CommandText = "insert into prepod " +
                "( fam, im, ot, kafedra_id, dolznost_id, " +
                "  stepen_id, zvanie_id, actual, address, phone, sex )" +
                "  values " +
                "( @FAM, @IM, @OT, @KAFEDRA_ID, @DOLZNOST_ID, " +
                "  @STEPEN_ID, @ZVANIE_ID, @ACTUAL, @ADDRESS, @PHONE, @SEX )";

            cmd.Parameters.Add("@FAM", SqlDbType.NVarChar).Value = famn;
            cmd.Parameters.Add("@IM", SqlDbType.NVarChar).Value = imn;
            cmd.Parameters.Add("@OT", SqlDbType.NVarChar).Value = otn;

            cmd.Parameters.Add("@KAFEDRA_ID", SqlDbType.Int).Value = kafidn;
            cmd.Parameters.Add("@DOLZNOST_ID", SqlDbType.Int).Value = dolzn;
            cmd.Parameters.Add("@STEPEN_ID", SqlDbType.Int).Value = stepidn;
            cmd.Parameters.Add("@ZVANIE_ID", SqlDbType.Int).Value = zvann;
            cmd.Parameters.Add("@ACTUAL", SqlDbType.Bit).Value = actualn;
            cmd.Parameters.Add("@ADDRESS", SqlDbType.NVarChar).Value = addrn;
            cmd.Parameters.Add("@PHONE", SqlDbType.NVarChar).Value = phonen;            
            cmd.Parameters.Add("@SEX", SqlDbType.Bit).Value = sexn;            

            try
            {
                cmd.ExecuteNonQuery();
            }
            catch (Exception exx)
            {
                MessageBox.Show("Неожиданный сбой при передаче данных.  " + 
                    " Повтоите операцию ввода еще раз.",
                    "Ошибка передачи данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            cmd = new SqlCommand("select @@Identity", main.global_connection);
            int prepidn = Convert.ToInt32(cmd.ExecuteScalar().ToString());

            if (pe.deny_photo==false)
            FilePhoto = pe.FilePhoto;
            save_prepod_photo_by_id(prepidn);

            //сохранить данные нового зав каф
            if (newzav)
            {
                //дать запрос и поставить в поле завкаф нулевого препода
                cmd = new SqlCommand("update kafedra set zav_kaf_id=@ZAV where " + 
                    "id = " + kaf_set.Rows[kaf_list.SelectedIndex][0].ToString(), 
                    main.global_connection);
                cmd.Parameters.Add("@ZAV", SqlDbType.Int).Value = prepidn;
                cmd.ExecuteNonQuery();

                int Sel = kaf_list.SelectedIndex;

                //загрузить кафедры, получить активную
                string selcom = "select id, name_krat,  " + 
                    " name, zav_kaf_id from kafedra " +
                    "where actual=1 " +
                    "order by priority";

                main.global_adapter = new SqlDataAdapter(selcom,
                    main.global_connection);

                kaf_set = new DataTable();

                main.global_adapter.Fill(kaf_set);

                foreach (DataRow dr in kaf_set.Rows)
                {
                    kaf_list.Items.Add(dr[1]);
                }

                if (Sel<=kaf_list.Items.Count)
                    kaf_list.SelectedIndex = Sel;
            }

            pe.Dispose();

            //сделать повторную загрузку
            kaf_list_SelectedIndexChanged_1(sender, new EventArgs());
        }


        /// <summary>
        /// редактирование преподавателя      
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            DataGridViewCell cell = dataGridView1.SelectedCells[0];
            int row = cell.RowIndex;
            
            int prep_id = (int)prep_set.Rows[row][13];           

            prepod_edit pe = new prepod_edit();

            pe.prep_id = prep_id;

            pe.dolz_id = (int)prep_set.Rows[row][8];
            pe.zvan_id = (int)prep_set.Rows[row][10];
            pe.uch_id = (int)prep_set.Rows[row][12];
            pe.kaf_id = (int)kaf_set.Rows[kaf_list.SelectedIndex][0];            
            pe.pictureBox1.Image = GetPhotoFromBD("prepod", prep_id);
            pe.deny_photo = true;
            pe.zavkaf_id = (int)kaf_set.Rows[sprav_prepods.kafedra_sel][3];


            pe.status_box.Checked = (bool)prep_set.Rows[row][14];
            if (pe.status_box.Checked)
                pe.status_box.Text = "Статус: работает";
            else
                pe.status_box.Text = "Статус: уволен";

            bool sex = (bool)prep_set.Rows[row][17];

            if (sex == false)
            {
                pe.female.Checked = true;
            }
          
            pe.fam.Text = prep_set.Rows[row][1].ToString();
            pe.im.Text = prep_set.Rows[row][2].ToString();
            pe.ot.Text = prep_set.Rows[row][3].ToString();
            pe.phone.Text = prep_set.Rows[row][16].ToString();
            pe.email.Text = prep_set.Rows[row][18].ToString();
            
            DialogResult pres = pe.ShowDialog(); // -----------

            if (pres == DialogResult.Cancel) return;

            string famn = pe.fam.Text.Trim();
            string imn = pe.im.Text.Trim();
            string otn = pe.ot.Text.Trim();
            int kafidn = pe.kaf_id;
            int dolzn = pe.dolz_id;
            int stepidn = pe.uch_id;
            int zvann = pe.zvan_id;
            bool actualn = pe.status_box.Checked;
            string addrn = pe.email.Text.Trim(); ;
            string phonen = pe.phone.Text.Trim();
            bool sexn = pe.male.Checked;

            //фото сохраняется отдельно

            SqlCommand cmd = new SqlCommand();
            cmd.Connection = main.global_connection;
            cmd.CommandText = "select prepod.actual, " + 
                "  kafedra.name from prepod " +
                " join kafedra on prepod.kafedra_id = kafedra.id " +
                " where fam like @FAM and " +
                " im like @IM and " +
                " ot like @OT and " +
                " sex = @SEX ";
            cmd.Parameters.Add("@FAM", SqlDbType.NVarChar).Value = famn;
            cmd.Parameters.Add("@IM", SqlDbType.NVarChar).Value = imn;
            cmd.Parameters.Add("@OT", SqlDbType.NVarChar).Value = otn;
            cmd.Parameters.Add("@SEX", SqlDbType.Bit).Value = sexn;

            SqlDataAdapter sda = new SqlDataAdapter(cmd);

            DataTable oldprepods = new DataTable();
            sda.Fill(oldprepods);

            //добавить новую запись в БД
            cmd = new SqlCommand();
            cmd.Connection = main.global_connection;

            cmd.CommandText = "update prepod set " +
                "  fam = @FAM, im = @IM, ot = @OT,  " + 
                " kafedra_id = @KAFEDRA_ID, dolznost_id = @DOLZNOST_ID, " +
                "  stepen_id = @STEPEN_ID, zvanie_id = @ZVANIE_ID,  " + 
                " actual = @ACTUAL, address = @ADDRESS, " + 
                "  phone = @PHONE, sex = @SEX " +
                "  where id =  @PREPPID";

            cmd.Parameters.Add("@FAM", SqlDbType.NVarChar).Value = famn;
            cmd.Parameters.Add("@IM", SqlDbType.NVarChar).Value = imn;
            cmd.Parameters.Add("@OT", SqlDbType.NVarChar).Value = otn;
            cmd.Parameters.Add("@KAFEDRA_ID", SqlDbType.Int).Value = kafidn;
            cmd.Parameters.Add("@DOLZNOST_ID", SqlDbType.Int).Value = dolzn;
            cmd.Parameters.Add("@STEPEN_ID", SqlDbType.Int).Value = stepidn;
            cmd.Parameters.Add("@ZVANIE_ID", SqlDbType.Int).Value = zvann;
            cmd.Parameters.Add("@ACTUAL", SqlDbType.Bit).Value = actualn;
            cmd.Parameters.Add("@ADDRESS", SqlDbType.NVarChar).Value = addrn;
            cmd.Parameters.Add("@PHONE", SqlDbType.NVarChar).Value = phonen;
            cmd.Parameters.Add("@SEX", SqlDbType.Bit).Value = sexn;
            cmd.Parameters.Add("@PREPPID", SqlDbType.Int).Value = prep_id;

            try
            {
                cmd.ExecuteNonQuery();
            }
            catch (Exception exx)
            {
                MessageBox.Show("Неожиданный сбой при передаче данных.  " + 
                    " Повтоите операцию ввода еще раз.",
                    "Ошибка передачи данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (pe.photochanged)
            {
                FilePhoto = pe.FilePhoto;
                save_prepod_photo_by_id(prep_id);
            }


            //проверить зав кафедрой
            if (pe.zav_changed)
            {
                //дать запрос и поставить в поле завкаф нулевого препода
                cmd = new SqlCommand("update kafedra set zav_kaf_id=69 where " + 
                    "id = " + kaf_set.Rows[kaf_list.SelectedIndex][0].ToString(), 
                    main.global_connection);
                cmd.ExecuteNonQuery();

                int Sel = kaf_list.SelectedIndex;

                //загрузить кафедры, получить активную
                string selcom = "select id, name_krat, name,  " + 
                    " zav_kaf_id from kafedra " +
                    "where actual=1 " +
                    "order by priority";

                main.global_adapter = new SqlDataAdapter(selcom,
                    main.global_connection);

                kaf_set = new DataTable();

                main.global_adapter.Fill(kaf_set);

                foreach (DataRow dr in kaf_set.Rows)
                {
                    kaf_list.Items.Add(dr[1]);
                }

                if (Sel<=kaf_list.Items.Count)
                    kaf_list.SelectedIndex = Sel;
                
            }

            bool newzav = (pe.dolz_list.Text.ToLower().Trim().Contains("заведующий"));

            //сохранить данные нового зав каф
            if (newzav)
            {
                //дать запрос и поставить в поле завкаф нулевого препода
                cmd = new SqlCommand("update kafedra set zav_kaf_id=@ZAV where " +
                    "id = " + kaf_set.Rows[kaf_list.SelectedIndex][0].ToString(),
                    main.global_connection);
                cmd.Parameters.Add("@ZAV", SqlDbType.Int).Value = prep_id;
                cmd.ExecuteNonQuery();

                int Sel = kaf_list.SelectedIndex;

                //загрузить кафедры, получить активную
                string selcom = "select id, name_krat, name, zav_kaf_id from kafedra " +
                    "where actual=1 " +
                    "order by priority";

                main.global_adapter = new SqlDataAdapter(selcom,
                    main.global_connection);

                kaf_set = new DataTable();

                main.global_adapter.Fill(kaf_set);

                foreach (DataRow dr in kaf_set.Rows)
                {
                    kaf_list.Items.Add(dr[1]);
                }

                if (Sel <= kaf_list.Items.Count)
                    kaf_list.SelectedIndex = Sel;
            }


            pe.Dispose();

            //сделать повторную загрузку
            kaf_list_SelectedIndexChanged_1(sender, new EventArgs());
        }

        public Bitmap GetPhotoFromBD(string tablename, int id)
        {

            main.global_command = new SqlCommand(
                "SELECT photo FROM " + tablename + " where id = " + id.ToString(),
                main.global_connection);

            SqlDataReader reader = 
                main.global_command.ExecuteReader(CommandBehavior.SequentialAccess);

            SqlBytes bytes = null;
            Bitmap image = null;

            while (reader.Read())
            {

                bytes = reader.GetSqlBytes(0);
                image = new Bitmap(bytes.Stream);
            }

            reader.Close();

            return image;
        }

        public static byte[] GetPhotoFromFile(string filePath)
        {
            FileStream stream = new FileStream(
                filePath, FileMode.Open, FileAccess.Read);
            BinaryReader reader = new BinaryReader(stream);

            byte[] photo = reader.ReadBytes((int)stream.Length);

            reader.Close();
            stream.Close();

            return photo;
        }

        string FilePhoto = "";

        /// <summary>
        /// сохранить фото в БД
        /// </summary>
        /// <param name="id"></param>
        private void save_prepod_photo_by_id(int id)
        {
            //DataGridViewRow cr = dataGridView1.CurrentRow;
            //int id = (int)cr.Cells["id"].Value;

            byte[] photo = GetPhotoFromFile(FilePhoto);

            main.global_command = new SqlCommand();
            main.global_command.CommandText = "update prepod set " +
                " photo = @p where id = @id";
            main.global_command.Connection = main.global_connection;

            main.global_command.Parameters.Add("@p", 
                SqlDbType.Image, photo.Length).Value = photo;
            main.global_command.Parameters.Add("@id", 
                SqlDbType.Int).Value = id;

            try
            {
                main.global_command.ExecuteNonQuery();
            }
            catch (Exception exx)
            {
                MessageBox.Show(exx.Message);
            }

        }

        private void dataGridView1_CellDoubleClick(object sender, 
            DataGridViewCellEventArgs e)
        {
            toolStripButton2_Click(sender, new EventArgs());
        }

        private void kaf_list_Click(object sender, EventArgs e)
        {

        }

        private void копироватьВБуферToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Clipboard.Clear();
            try
            {
                Clipboard.SetDataObject(dataGridView1.GetClipboardContent());
            }
            catch(Exception ex)
            {
                ;
            }
        }
    }
}