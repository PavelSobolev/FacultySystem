using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace FSystem
{
    public partial class predmet_edit : Form
    {
        public predmet_edit()
        {
            InitializeComponent();
        }

        public string NormalizeLetters(string str)
        {
            foreach (char c in str)
            {

                bool can = Char.IsLetter(c) || Char.IsDigit(c) ||
                    c=='-' || c=='/' || c=='.' || c=='\"' || c==',' || c==' ';

                if (!can)
                {
                       int pos = str.IndexOf(c);
                        str = str.Remove(pos, 1);
                        if (str.Length == 0) break;                    
                }
            }

            return str;
        }

        public bool edit = false;
        private void button1_Click(object sender, EventArgs e)
        {
            full_name.Text = NormalizeLetters(full_name.Text);
            krat_name.Text = NormalizeLetters(krat_name.Text);
            
            string error_string = "Обнаружены следующие ошибки ввода:\n\n";
            bool can_continue = true;

            // проверить есть ли такой предмет в этой группе
            string sql = "select id from predmet " +
                " where kurs_id = @kid and semestr = @sem and name like @nm and grupa_id = @gr";
            SqlCommand cmd = new SqlCommand(sql, main.global_connection);
            cmd.Parameters.Add("@kid", SqlDbType.Int).Value = kurs_id;
            cmd.Parameters.Add("@sem", SqlDbType.Int).Value = semestr.Value;
            cmd.Parameters.Add("@nm", SqlDbType.NVarChar).Value = NormalizeLetters(full_name.Text);
            cmd.Parameters.Add("@gr", SqlDbType.Int).Value = grup_id;
            DataTable t = new DataTable();
            (new SqlDataAdapter(cmd)).Fill(t);

            if (t.Rows.Count > 0)
            {
                if (!edit)
                {
                    error_string += " - ВНИМАНИЕ! ПРЕДМЕТ С ТАКИМ НАЗВАНИЕМ УЖЕ СУЩЕСТВУЕТ В ЭТОЙ ГРУППЕ!\n" +
                        "ВНИМАТЕЛЬНО ИЗУЧИТЕ ПЕРЕЧЕНЬ ИМЕЮЩИХСЯ ПРЕДМЕТОВ! НЕЛЬЗЯ ДОБАВЛЯТЬ ПРЕДМЕТЫ С ПОВТОРЯЮЩИМИСЯ НАЗВАНИЯМИ!\n";
                    can_continue = false;
                }
                
            }

            if (full_name.Text.Trim().Length == 0)
            {
                error_string += " - не введено название предмета [не может быть пустым]\n";
                can_continue = false;
                full_name.Select();
            }

            if (krat_name.Text.Trim().Length == 0)
            {
                error_string += " - не введено краткое название предмета [не может быть пустым]\n";
                krat_name.Select();
                can_continue = false;
            }
            else
            {
                if (krat_name.Text.Trim().Length >= 22)
                {
                    error_string += " - краткое название предмета содержит слишком много символов\n" + 
                        "[в кратком названии должно быть не более 22 символов]\n";
                    krat_name.Select();
                    can_continue = false;
                }
            }

            if (vid_view.Items.Count == 0)
            {
                error_string += " - не указаны виды занятий для предмета [выберите, по крайней мере, один]\n";
                include_button.Focus();
                can_continue = false;
            }

            foreach (ListViewItem lv in vid_view.Items)
            {
                double ch = Convert.ToDouble(lv.SubItems[1].Text);
                if (ch > 200.0 || ch < 1.0)
                {
                    error_string += " - недопустимое число часов по виду занятия '" +
                        lv.SubItems[0].Text + "' [допустимое значение от 1 до 200]\n";
                    can_continue = false;
                }
            }

            if (!can_continue)
            {
                MessageBox.Show(error_string, "Ошибка ввода",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return;
            }                        

            string ntxt = full_name.Text.Trim();
            Normalize(ref ntxt);
            full_name.Text = ntxt;

            ntxt = krat_name.Text.Trim();
            Normalize(ref ntxt);
            krat_name.Text = ntxt;            
            
            DialogResult = DialogResult.OK;
        }

        /// <summary>
        /// удалить из строки метасимволы запросов '
        /// </summary>
        /// <param name="str"></param>
        public void Normalize(ref string str)
        {
            while (str.Contains("'"))
            {
                int pos = str.IndexOf("'");
                str = str.Remove(pos, 1);
                if (str.Length == 0) break;
            }
        }

        public DataTable vid_set, vid_predmet_set, prepod_set, grupa_set, type_set, kaf_set;
        public string q = ""; //текст запроса

        public int pred_id = 0;
        public int grup_id = 0;
        public int prepod_id = 0;
        public int predmet_type_id = 0;
        public int kaf_id = 0;
        public int semestr_id = 0;
        public int kurs_id = 0;
        public bool delenie = false;
        public int fakultet_id = 0;
        public int type_id = 0;

        public bool is_id_intable(int id)
        {
            for (int i = 0; i < vid_view.Items.Count; i++)
            {
                if ((int)vid_view.Items[i].Tag == id)
                    return true;
            }

            return false;
        }

        DataRow current_vid = null; 
        ListViewItem current_lvi = null; //текущий выбранный пункт списка

        //найти текущйи выделенные пункт в списке
        private void vid_view_MouseDown(object sender, MouseEventArgs e)
        {
            ListViewItem lvi = vid_view.GetItemAt(e.X, e.Y);            
            current_lvi = lvi;

            if (lvi == null) return;

            int id = (int)lvi.Tag;

            foreach (DataRow dr in vid_set.Rows)
            {
                if (id == (int)dr[0])
                {
                    current_vid = dr;
                    return;
                }
            }
        }

        /// <summary>
        /// добавить новый вид занятий в предмет
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void include_button_Click(object sender, EventArgs e)
        {           
            if (vid_list.SelectedIndex < 0)
            {
                if (vid_list.Items.Count > 0)
                    vid_list.SelectedIndex = 0;
            }

            DataRow dr = vid_set.Rows[vid_list.SelectedIndex];
            int id = (int)dr[0];

            if (is_id_intable(id))
            {
                MessageBox.Show("Данный вид занятия уже добавлен.",
                    "Ошибка редактирования",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                vid_list.Focus();
                vid_list.Select();
                if (vid_list.SelectedIndex < vid_list.Items.Count - 1) vid_list.SelectedIndex++;
                return;
            }

            int num = 0;

            ListViewItem lvi = new ListViewItem();
            
            lvi.Text = dr[1].ToString();            
            lvi.Tag = dr[0];

            double ch = Convert.ToDouble(dr[2]);
            bool recount = Convert.ToBoolean(dr[3]);

            if (recount)
                ch = ch * (int)grupa_set.Rows[grupa_list.SelectedIndex][2];
            else
                ch = 0.0;

            lvi.SubItems.Add(string.Format("{0:F2}", ch));
            vid_view.Items.Add(lvi);


            if (vid_list.SelectedIndex < vid_list.Items.Count - 1) vid_list.SelectedIndex++;
        }

        /// <summary>
        /// удалить вид занятия из предмета
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void exclude_button_Click(object sender, EventArgs e)
        {
            if (vid_view.Items.Count == 0)
                return;

            if (vid_view.Items.Count == 1)
            {
                MessageBox.Show("Удаление невозможно. Список видов занятия не может быть пустым.",
                    "Запрос отклонен",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            vid_view.Items.Remove(current_lvi);


        }


        private void predmet_edit_Load(object sender, EventArgs e)
        {
            //заполнить перечень видов занятий
            q = "select id, name, koef, recount " +
                " from vid_zan " +
                " where (show_in_grid = 1 or (kod = 'кнр' or kod = 'ср'))" + 
                " order by name ";
            SqlDataAdapter sda = new SqlDataAdapter(q, main.global_connection);
            vid_set = new DataTable();
            sda.Fill(vid_set);

            foreach(DataRow r in vid_set.Rows)
                vid_list.Items.Add(r[1].ToString());

            ///вывести виды занятий в таблицу           
            q = "select " +
                " vid_zan.koef,  " + //0
                " vid_zan.name,  " + //1
                " vid_zan.id,  " +  //2                               
                " vidzan_predmet.kol_chas, vid_zan.kod " + //3, 4                                    
                " from vidzan_predmet " +
                " join vid_zan on vid_zan.id = vidzan_predmet.vidzan_id " +                
                " where vidzan_predmet.predmet_id =  " + pred_id.ToString() +
                " and (show_in_grid=1 or (kod = 'кнр' or kod = 'ср')) " +
                " order by vid_zan.id ";

            sda = new SqlDataAdapter(q, main.global_connection);
            vid_predmet_set = new DataTable();
            sda.Fill(vid_predmet_set);

            foreach (DataRow r in vid_predmet_set.Rows)
            {                
                ListViewItem lvi = new ListViewItem();
                lvi.Text = r[1].ToString();
                lvi.Tag = r[2];

                double ch = Convert.ToDouble(r[3]);                
                lvi.SubItems.Add(string.Format("{0:F2}", ch));

                vid_view.Items.Add(lvi);
            }
            

            //заполнить преподов по алфавиту
            q = "select id, " +
                " 'prepod' = prepod.fam  + ' ' + left(prepod.im,1)  + '. ' + left(prepod.ot,1) + '.', " +
                " dolznost_id, stepen_id, zvanie_id, kafedra_id, phone, sex, address, actual, fam, im, ot " +                 
                " from prepod " +
                " where fam <> '0' " +
                " order by fam, im, ot ";
            sda = new SqlDataAdapter(q, main.global_connection);
            prepod_set = new DataTable();
            sda.Fill(prepod_set);

            foreach (DataRow rr in prepod_set.Rows)
                prepod_list.Items.Add(rr[1].ToString());

            if (pred_id == 0)
                prepod_list.SelectedIndex = 0;
            else
                prepod_list.SelectedIndex = GetPosById(prepod_set, prepod_id);

            //заполнить группы по специальности
            q = "select distinct grupa.id, name, count(student.id), kurs_id, fakultet_id from grupa " +
                " left outer join student on student.gr_id = grupa.id " +
                " where fakultet_id = " + main.fakultet_id.ToString() +  
                " group  by specialnost_id, grupa.id, name, kurs_id, fakultet_id ";
            sda = new SqlDataAdapter(q, main.global_connection);
            grupa_set = new DataTable();
            sda.Fill(grupa_set);

            foreach (DataRow rrr in grupa_set.Rows)
                grupa_list.Items.Add(rrr[1].ToString());

            grupa_list.SelectedIndex = GetPosById(grupa_set, grup_id);
    

            //заполнить кафедры
            q = "select id, name, name_krat, zav_kaf_id from kafedra " + 
                " where actual = 1 order by priority ";

            sda = new SqlDataAdapter(q, main.global_connection);
            kaf_set = new DataTable();
            sda.Fill(kaf_set);

            foreach (DataRow rrrr in kaf_set.Rows)
                kaf_list.Items.Add(rrrr[2].ToString());

            if (pred_id != 0)
                kaf_list.SelectedIndex = GetPosById(kaf_set, kaf_id);
            else
                kaf_list.SelectedIndex = 0;
           
            //заполнить типы занятий
            //заполнить кафедры
            q = "select id, name from predmet_type ";

            sda = new SqlDataAdapter(q, main.global_connection);
            type_set = new DataTable();
            sda.Fill(type_set);

            foreach (DataRow r in type_set.Rows)
                type_predmet_list.Items.Add(r[1].ToString());

            if (pred_id != 0)
                type_predmet_list.SelectedIndex = GetPosById(type_set, type_id);
            else
                type_predmet_list.SelectedIndex = 0;
            
            vid_list.SelectedIndex = 0;
            delenie_list.SelectedIndex = 0;
            if (pred_id==0) checkBox1.Checked = true;
        }

        private int GetPosById(DataTable box, int id)
        {
            //взять номер списка по его ИД
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

        private void prepod_list_SelectedIndexChanged(object sender, EventArgs e)
        {
            //выставитиь ид препода
            if (prepod_list.Items.Count >= 0)
                prepod_id = (int)prepod_set.Rows[prepod_list.SelectedIndex][0];
            else
                prepod_id = 0;
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            //показать полную инфу про препода
            DataRow row = null;
            if (prepod_list.SelectedIndex >= 0)
                row = prepod_set.Rows[prepod_list.SelectedIndex];
            else
                return;

            prepod_edit pe = new prepod_edit();

            pe.prep_id = prepod_id;

            pe.dolz_id = (int)row[2];
            pe.zvan_id = (int)row[4];
            pe.uch_id = (int)row[3];
            pe.kaf_id = (int)row[5];
            pe.pictureBox1.Image = main.GetPhotoFromBD("prepod", prepod_id);
            pe.deny_photo = true;

            pe.status_box.Checked = (bool)row[9];
            if (pe.status_box.Checked)
                pe.status_box.Text = "Статус: работает";
            else
                pe.status_box.Text = "Статус: уволен";

            bool sex = (bool)row[7];

            if (sex == false)
            {
                pe.female.Checked = true;
            }

            pe.fam.Text = row[10].ToString();
            pe.im.Text = row[11].ToString();
            pe.ot.Text = row[12].ToString();
            pe.phone.Text = row[6].ToString();
            pe.email.Text = row[8].ToString();

            pe.button8.Left = pe.button7.Left;
            pe.button8.Text = "Закрыть";
            pe.button7.Visible = false;
            pe.button4.Visible = false;
            pe.button3.Visible = false;
            pe.button2.Visible = false;
            pe.button1.Visible = false;

            DialogResult pres = pe.ShowDialog();
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (grupa_list.SelectedIndex < 0) return;
            
            //вывести инфу про группу
            int cont = (int)grupa_set.Rows[grupa_list.SelectedIndex][2];
            string scont = "Группа " + grupa_list.Text + "\n\n";

            if (cont == 0)
            {
                scont += "Контингент группы не введен в базу.\nВыполните заполнение группы через раздел " + 
                    "\"Справочники/Учебные группы\".";
            }
            else
            {
                scont += "Контингент на " + DateTime.Now.ToShortDateString() + 
                    ": в группе состоит " + cont.ToString() + " чел.";
            }

            MessageBox.Show(scont, "Красткая информация о группе.",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public bool first = true; 

        private void grupa_list_SelectedIndexChanged(object sender, EventArgs e)
        {
            //выставить группу, FSystem, курс, семестр
            grup_id = (int)grupa_set.Rows[grupa_list.SelectedIndex][0];
            fakultet_id = (int)grupa_set.Rows[grupa_list.SelectedIndex][4];
            kurs_id = (int)grupa_set.Rows[grupa_list.SelectedIndex][3];
            semestr.Minimum = kurs_id * 2 - 1;
            semestr.Maximum = kurs_id * 2 ;
            
            if (!first)
                semestr.Value = kurs_id * 2 - 1;

            first = false;
        }

        private void kaf_list_SelectedIndexChanged(object sender, EventArgs e)
        {
            //выставить кафедру
            kaf_id = (int)kaf_set.Rows[kaf_list.SelectedIndex][0];
            toolTip1.SetToolTip(kaf_list, 
                kaf_set.Rows[kaf_list.SelectedIndex][1].ToString());
        }

        private void type_predmet_list_SelectedIndexChanged(object sender, EventArgs e)
        {
            //выставить тип
            type_id = (int)type_set.Rows[type_predmet_list.SelectedIndex][0];
        }

        private void delenie_list_SelectedIndexChanged(object sender, EventArgs e)
        {
            //выставить деление
            delenie = (delenie_list.SelectedIndex == 0) ? false : true;
        }

        private void vid_view_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            //задать количество часов для предмета
            if (current_lvi == null) return;

            inputbox ib = new inputbox(
                "Введите в окно редактирования количество часов\n" + 
                "по указанному виду занятия.\n\nЦелая часть числа отделяется от дробной знаком 'запятая' (,).",
                current_lvi.Text, 
                current_lvi.SubItems[1].Text,
                "Кол-во часов:");
            ib.is_numeric = true;
            
            DialogResult res;

            do
            {
                res = ib.ShowDialog();
                if (res == DialogResult.Cancel) break;
            }
            while(res!=DialogResult.OK);

            if (res == DialogResult.OK)
            {
                double ch = Convert.ToDouble(ib.textBox1.Value);

                if (ch <= 200.0 && ch >= 1.00)
                    current_lvi.SubItems[1].Text = string.Format("{0:F2}", ch);
                else
                    MessageBox.Show("Введенно недопустимое количество часов [допустимое значение: от 1 до 200]",
                        "Ошибка ввода",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
            }

            ib.Dispose();
        }

        private void задатьКоличествоЧасовEnterToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //задать количество часов для предмета
            if (current_lvi == null) return;

            inputbox ib = new inputbox(
                "Введите в окно редактирования количество часов\n" +
                "по указанному виду занятия.\n\nЦелая часть числа отделяется от дробной знаком 'запятая' (,).",
                current_lvi.Text,
                current_lvi.SubItems[1].Text,
                "Кол-во часов:");

            ib.is_numeric = true;

            DialogResult res;

            do
            {
                res = ib.ShowDialog();
                if (res == DialogResult.Cancel) break;
            }
            while (res != DialogResult.OK);

            if (res == DialogResult.OK)
            {
                double ch = Convert.ToDouble(ib.textBox1.Text);

                if (ch <= 200.0 && ch >= 1.00)
                    current_lvi.SubItems[1].Text = string.Format("{0:F2}", ch);
                else
                    MessageBox.Show("Введенно недопустимое количество часов [допустимое значение: от 1 до 200]",
                        "Ошибка ввода",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
            }

            ib.Dispose();
        }

        private void vid_view_KeyDown(object sender, KeyEventArgs e)
        {
            //задать количество часов для предмета
            if (e.KeyCode == Keys.Return)
            {
                if (current_lvi == null) return;

                inputbox ib = new inputbox(
                    "Введите в окно редактирования количество часов\n" +
                    "по указанному виду занятия.\n\nЦелая часть числа отделяется от дробной знаком 'запятая' (,).",
                    current_lvi.Text,
                    current_lvi.SubItems[1].Text,
                    "Кол-во часов:");

                ib.is_numeric = true;

                DialogResult res;

                do
                {
                    res = ib.ShowDialog();
                    if (res == DialogResult.Cancel) break;
                }
                while (res != DialogResult.OK);

                if (res == DialogResult.OK)
                {
                    double ch = Convert.ToDouble(ib.textBox1.Text);

                    if (ch <= 200.0 && ch >= 1.00)
                        current_lvi.SubItems[1].Text = string.Format("{0:F2}", ch);
                    else
                        MessageBox.Show("Введенно недопустимое количество часов [допустимое значение: от 1 до 200]",
                            "Ошибка ввода",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
                }

                ib.Dispose();
            }

            //удалиить вид занятия из предмета
            if (e.KeyCode == Keys.Delete)
            {
                exclude_button_Click(sender, new EventArgs());
            }
        }


        private void vid_view_SelectedIndexChanged(object sender, EventArgs e)
        {
           //получить текущий выеделнный пункт в списке видов занятий предмета
            ListView.SelectedIndexCollection inds = vid_view.SelectedIndices;

            ListViewItem lvi = null;

            foreach (int i in inds)
            {
                lvi = vid_view.Items[i];
            }

            current_lvi = lvi;

            if (lvi == null) return;

            int id = (int)lvi.Tag;

            foreach (DataRow dr in vid_set.Rows)
            {
                if (id == (int)dr[0])
                {
                    current_vid = dr;
                    return;
                }
            }
        }


        private void vid_list_KeyDown(object sender, KeyEventArgs e)
        {
            ///удалить вид занятия
            if (e.KeyCode == Keys.Right)
            {
                include_button_Click(sender, new EventArgs());
            }
        }

        private void выполнитьПересчетЧасовПоКонтингентуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //сделать пересчет по факту

            int count_student = (int)grupa_set.Rows[grupa_list.SelectedIndex][2];

            if (count_student == 0) return;

            foreach(ListViewItem lvi in vid_view.Items )
            {
                int id = (int)lvi.Tag;
                int i = GetPosById(vid_set, id);

                DataRow dr = vid_set.Rows[i];

                double ch = Convert.ToDouble(dr[2])*count_student;
                bool recount = Convert.ToBoolean(dr[3]);

                if (recount)
                lvi.SubItems[1].Text = string.Format("{0:F2}", ch);

            }
        }


        public static string DefaultPath = Environment.GetFolderPath(
            Environment.SpecialFolder.MyPictures) + "\\emptyface.jpg";
        public string FilePhoto = "";

        private void linkLabel5_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            //добавить нового препода
            prepod_edit pe = new prepod_edit();

            pe.pictureBox1.Image = main.GetPhotoFromBD("prepod", 69);
            pe.status_box.Checked = true;
            pe.first = false;
            pe.zavkaf_id = (int)kaf_set.Rows[kaf_list.SelectedIndex][3];

            FilePhoto = DefaultPath;

            pe.pictureBox1.Image.Save(FilePhoto, System.Drawing.Imaging.ImageFormat.Jpeg);
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

            bool newzav = (pe.dolz_list.Text.ToLower().Trim().Contains("заведующий"));

            //фото сохраняется отдельно

            SqlCommand cmd = new SqlCommand();
            cmd.Connection = main.global_connection;
            cmd.CommandText = "select prepod.actual, kafedra.name from prepod " +
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
                    message = "Обнаружен преподаватель с такой же фамилией, именем и отчеством:\n\n";
                else
                    message = "Обнаружены преподаватели с аналогичными фамилией, именем и отчеством:\n\n";

                int i = 1;
                foreach (DataRow dr in oldprepods.Rows)
                {
                    string stat = ((bool)dr[0]) ? "работает" : "уволен(а)";

                    message += i.ToString() + ". " +
                        famn + " " + imn + " " + otn + " [кафедра " +
                        dr[1].ToString() + ", статус - " + stat + "];\n";
                }

                message += "\n\nВыполнить сохранение введенной Вами информации?";

                DialogResult dres = MessageBox.Show(message,
                    "Выбор действия",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (dres == DialogResult.No)
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
                MessageBox.Show("Неожиданный сбой при передаче данных. Повтоите операцию ввода еще раз.",
                    "Ошибка передачи данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            cmd = new SqlCommand("select @@Identity", main.global_connection);
            int prepidn = Convert.ToInt32(cmd.ExecuteScalar().ToString());

            if (pe.deny_photo == false)
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

            //сделать повторную загрузку в список групп
            //заполнить преподов по алфавиту
            string q = "select id, " +
                " 'prepod' = prepod.fam  + ' ' + left(prepod.im,1)  + '. ' + left(prepod.ot,1) + '.', " +
                " dolznost_id, stepen_id, zvanie_id, kafedra_id, phone, sex, address, actual, fam, im, ot " +
                " from prepod " +
                " where fam <> '0' " +
                " order by fam, im, ot ";
            sda = new SqlDataAdapter(q, main.global_connection);
            prepod_set = new DataTable();
            sda.Fill(prepod_set);

            prepod_list.Items.Clear();

            foreach (DataRow rr in prepod_set.Rows)
                prepod_list.Items.Add(rr[1].ToString());

            prepod_list.SelectedIndex = GetPosById(prepod_set, prepidn);
        }

        /// <summary>
        /// сохранить фото в БД
        /// </summary>
        /// <param name="id"></param>
        private void save_prepod_photo_by_id(int id)
        {
            //DataGridViewRow cr = dataGridView1.CurrentRow;
            //int id = (int)cr.Cells["id"].Value;

            byte[] photo = main.GetPhotoFromFile(FilePhoto);

            main.global_command = new SqlCommand();
            main.global_command.CommandText = "update prepod set " +
                " photo = @p where id = @id";
            main.global_command.Connection = main.global_connection;

            main.global_command.Parameters.Add("@p", SqlDbType.Image, photo.Length).Value = photo;
            main.global_command.Parameters.Add("@id", SqlDbType.Int).Value = id;

            try
            {
                main.global_command.ExecuteNonQuery();
            }
            catch (Exception exx)
            {
                MessageBox.Show(exx.Message);
            }

        }

        private void full_name_TextChanged(object sender, EventArgs e)
        {
            //
        }

    }
}