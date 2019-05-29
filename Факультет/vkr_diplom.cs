using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace FSystem
{
    public partial class vkr_diplom : Form
    {
        public vkr_diplom()
        {
            InitializeComponent();
        }


        public int kurs = 0, grupa_id = 0, student_id = 0, rabota_id = 0, kol_chas = 10,
            student_rabota_id = 0, ruk_id = 0;

        public DataTable groups = new DataTable(),
            students = new DataTable(),
            works = new DataTable(),
            otmetki = new DataTable(),
            details = new DataTable();
        private string q = "";


        private void create_rabota()
        {
            string vid = (kurs == 4) ? "вкр" : "др";
            kol_chas = (kurs == 4) ? 10 : 15;
            
            //получить ид ВКР или Дипломной работы
            q = "select id, name from vid_rab where kod='" + vid + "'";
            main.global_adapter = new SqlDataAdapter(q, main.global_connection);
            DataTable vid_rab = new DataTable();
            main.global_adapter.Fill(vid_rab);
            int vid_rabota_id = Convert.ToInt32(vid_rab.Rows[0][0]);
            string vid_rabota_name = vid_rab.Rows[0][1].ToString();
            vid_rab.Dispose();
            
            //проверить, существует ли работа и создать её если нет
            q = string.Format("select rabota.id from rabota " +
                " join vid_rab on vid_rab.id = rabota.vid_rab_id " +
                " where y = {0} and vid_rab.id={1} and grupa_id = {2}",
                main.ends[main.ends.Count - 1].Year, vid_rabota_id, grupa_id);

            main.global_command = new SqlCommand(q, main.global_connection);
            DataTable rab_id = new DataTable();
            main.global_adapter = new SqlDataAdapter(main.global_command);
            main.global_adapter.Fill(rab_id);

            if (rab_id.Rows.Count == 0)
            {
                q = "insert into rabota(vid_rab_id, name, status, y, poruch, grupa_id, kol_chas) " +
                    " values (@vid_rab_id, @name, @status, @y, @poruch, @grupa_id, @kol_chas)";
                main.global_command = new SqlCommand(q, main.global_connection);

                main.global_command.Parameters.Add("@vid_rab_id", SqlDbType.Int).Value = vid_rabota_id;
                main.global_command.Parameters.Add("@name", SqlDbType.NVarChar).Value = vid_rabota_name + " (группа " +
                    grupa_list.Text + ")";
                main.global_command.Parameters.Add("@status", SqlDbType.Bit).Value = 0;
                main.global_command.Parameters.Add("@y", SqlDbType.Int).Value = main.ends[main.ends.Count - 1].Year;
                main.global_command.Parameters.Add("@poruch", SqlDbType.Bit).Value = true;
                main.global_command.Parameters.Add("@grupa_id", SqlDbType.Int).Value = grupa_id;
                main.global_command.Parameters.Add("@kol_chas", SqlDbType.Int).Value = kol_chas;
                
                try
                {
                    main.global_command.ExecuteNonQuery();
                }
                catch (Exception exx)
                {
                    MessageBox.Show("Невозможно выполнить операцию вследствие сетевой ошибки (" + vid_rabota_name + " не может быть создана)." + 
                        "Повторите попытку позднее.",
                        "Ошибка операции", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                q = "select @@identity";
                main.global_command = new SqlCommand(q, main.global_connection);
                SqlDataReader r = main.global_command.ExecuteReader();

                r.Read();
                rabota_id = Convert.ToInt32(r[0]);
                r.Close();
            }
            else
            {
                rabota_id = Convert.ToInt32(rab_id.Rows[0][0]);
            }
            rab_id.Dispose();
        }

        private void vkr_diplom_Load(object sender, EventArgs e)
        {
            fill_grups();            

            //заполнить список оценок
            q = "select id, str_name from vid_otmetka " +
                " where name = 2 or name = 3 or name = 4 or name = 5 " +
                " order by name";
            main.global_adapter = new SqlDataAdapter(q, main.global_connection);
            main.global_adapter.Fill(otmetki);
            otm_list.Items.Clear();
            foreach (DataRow dr in otmetki.Rows)
            {
                string nm = dr[1].ToString();
                otm_list.Items.Add(nm);
            }
            otm_list.Items.Add("не определена");
            otm_list.SelectedIndex = otm_list.Items.Count - 1;
        }

        /// <summary>
        /// получить список групп 4 курса данного FSystemа
        /// </summary>
        public void fill_grups()
        {
            groups = new DataTable();
            q = string.Format("select grupa.id, grupa.name from grupa " + 
                " join kurs on kurs.id=grupa.kurs_id " + 
                " where kurs.nomer = {0} and fakultet_id = {1} and grupa.actual = 1 " + 
                " order by specialnost_id",
                kurs, main.fakultet_id);
            main.global_adapter = new SqlDataAdapter(q, main.global_connection);
            main.global_adapter.Fill(groups);

            grupa_list.Items.Clear();
            foreach (DataRow dr in groups.Rows)
            {
                string nm = dr[1].ToString();
                grupa_list.Items.Add(nm);
            }

            if (grupa_list.Items.Count>0)
                grupa_list.SelectedIndex = 0;
            else
            {
                MessageBox.Show(string.Format("Не найдено групп {0} курса.", kurs),
                    "Ошибка данных", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }            
        }

        private void grupa_list_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (grupa_list.Items.Count == 0) return;

            int pos = grupa_list.SelectedIndex;
            grupa_id = Convert.ToInt32(groups.Rows[pos][0]);                       

            //перезаполнить список группы
            fill_student();
        }

        public void fill_student()
        {
            students = new DataTable();
            student_list.Items.Clear();
            student_list.Text = "";

            q = "select	student.id, " +
                " fio = student.fam + ' ' + student.im " + //7
                    " from student  " +
                    " join grupa on grupa.id = student.gr_id " +                    
                    " where student.actual=1 and fam<>'-' " + 
                    " and grupa.id = " + grupa_id.ToString() +
                    " order by fam, im, ot ";
            main.global_adapter = new SqlDataAdapter(q, main.global_connection);
            main.global_adapter.Fill(students);

            if (students.Rows.Count > 0) create_rabota();

            foreach (DataRow dr in students.Rows)
            {
                string nm = dr[1].ToString();
                student_list.Items.Add(nm);
            }

            student_list.Height = 500;

            if (student_list.Items.Count > 0)
                student_list.SelectedIndex = 0;
            else
            {
                MessageBox.Show("Не найдено студентов в списке группы.",
                    "Ошибка данных", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        private void student_list_SelectedIndexChanged(object sender, EventArgs e)
        {
            сохранено.Visible = false;
            if (student_list.Items.Count == 0) return;

            int pos = student_list.SelectedIndex;
            student_id = Convert.ToInt32(students.Rows[pos][0]);

            show_details();
        }

        /// <summary>
        /// вывести тему студента, научного руководителя, статус предзащиты, оценку и дату защиты
        /// </summary>
        public void show_details()
        {
            if (student_id == 0) return;
            details = new DataTable();
            ruk_id = -1;
            checkBox1.Checked = false;
            otm_list.SelectedIndex = otm_list.Items.Count - 1;
            
            q = string.Format(
                "select student_rabota.id, student_id, tema, " +  //0 1 2
                " isnull(otmetka_id,-1), pred_status, isnull(ruk_id,-1), " + //3 4 5
                " fio = prepod.fam + ' ' + prepod.im + ' ' + prepod.ot " + //6
                " from student_rabota " + 
                " left outer join prepod on prepod.id = student_rabota.ruk_id" +
                " where rabota_id = {0} and student_id = {1}", rabota_id, student_id);
            main.global_adapter = new SqlDataAdapter(q, main.global_connection);
            main.global_adapter.Fill(details);

            if (details.Rows.Count == 0)
            {
                //для данного студента не выбрана тема - задать              
                q = "insert into student_rabota(student_id, rabota_id, tema, pred_status) " +
                    " values (@student_id, @rabota_id, '', 0)";
                main.global_command = new SqlCommand(q, main.global_connection);

                main.global_command.Parameters.Add("@student_id", SqlDbType.Int).Value = student_id;
                main.global_command.Parameters.Add("@rabota_id", SqlDbType.Int).Value = rabota_id;

                try
                {
                    main.global_command.ExecuteNonQuery();
                }
                catch (Exception exx)
                {
                    MessageBox.Show("Невозможно выполнить операцию вследствие сетевой ошибки.\n" +
                        "Повторите попытку позднее.",
                        "Ошибка операции", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                q = "select @@identity";
                main.global_command = new SqlCommand(q, main.global_connection);
                SqlDataReader r = main.global_command.ExecuteReader();
                r.Read();
                student_rabota_id = Convert.ToInt32(r[0]);
                r.Close();

                tema.Text = "нет";
                ruk.Text = "нет";
            }
            else
            {
                //иформация есть - вывести                
                DataRow dr = details.Rows[0];

                if (dr[2].ToString().Trim().Length > 0)
                    tema.Text = dr[2].ToString();
                else
                    tema.Text = "нет";

                if (dr[6].ToString().Trim().Length > 0)
                    ruk.Text = dr[6].ToString();
                else
                    ruk.Text = "нет";

                bool predzach = Convert.ToBoolean(dr[4]);
                int otm_id = Convert.ToInt32(dr[3]);

                if (predzach)
                {
                    checkBox1.Checked = true;
                    if (otm_id != -1)
                        otm_list.SelectedIndex = get_selected_otm(otm_id);
                    else
                        otm_list.SelectedIndex = otm_list.Items.Count - 1;
                }

                ruk_id = Convert.ToInt32(dr[5]);
            }

        }

        public int get_selected_otm(int id)
        {
            int i = 0;
            foreach (DataRow dr in otmetki.Rows)
            {
                int o_id = Convert.ToInt32(dr[0]);
                if (o_id == id) break;
                i++;
            }

            return i;
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            otm_list.Enabled = checkBox1.Checked;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string new_tema = "";
            int new_ruk = -1;
            bool pred_stat = false;
            int new_otm = 0;

            if (tema.Text.Trim().Length == 0 || tema.Text.Trim()=="нет")
            {
                MessageBox.Show("Не задана тема работы","Некорректные данные",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                сохранено.Visible = false;
                return;
            }

            if (ruk_id == -1)
            {
                MessageBox.Show("Не выбран руководитель работы", "Некорректные данные",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                сохранено.Visible = false;
                return;
            }

            //сохранение

            new_tema = main.Normalize1(tema.Text.Trim());
            new_ruk = ruk_id;
            pred_stat = checkBox1.Checked;
            if (otm_list.SelectedIndex<otm_list.Items.Count-1)
                new_otm = Convert.ToInt32(otmetki.Rows[otm_list.SelectedIndex][0]);

            if (new_otm == 0)
            {
                q = "update student_rabota set " +
                    " tema = @tema, pred_status = @pred, ruk_id = @ruk_id " +
                    " where student_id = @student_id and rabota_id = @rabota_id";
                main.global_command = new SqlCommand(q, main.global_connection);
            }
            else
            {
                q = "update student_rabota set " +
                    " tema = @tema, pred_status = @pred, ruk_id = @ruk_id, " +
                    " otmetka_id = @otmetka_id " + 
                    " where student_id = @student_id and rabota_id = @rabota_id";                
                main.global_command = new SqlCommand(q, main.global_connection);
                main.global_command.Parameters.Add("@otmetka_id", SqlDbType.Int).Value = new_otm;
            }
            
            main.global_command.Parameters.Add("@student_id", SqlDbType.Int).Value = student_id;
            main.global_command.Parameters.Add("@rabota_id", SqlDbType.Int).Value = rabota_id;
            main.global_command.Parameters.Add("@ruk_id", SqlDbType.Int).Value = new_ruk;
            main.global_command.Parameters.Add("@tema", SqlDbType.NVarChar).Value = new_tema;
            main.global_command.Parameters.Add("@pred", SqlDbType.Bit).Value = pred_stat;

            main.global_command.ExecuteNonQuery();

            try
            {
                ;
            }
            catch (Exception exx)
            {
                MessageBox.Show("Невозможно выполнить операцию вследствие сетевой ошибки.\n" +
                    "Повторите попытку позднее.",
                    "Ошибка операции", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            сохранено.Visible = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DataTable preps = new DataTable();
            q = "select prepod.id, " + 
                    " fio = prepod.fam + ' ' + prepod.im + ' ' + prepod.ot " + 
                    " from prepod " +
                    " join kafedra on kafedra.id = prepod.kafedra_id " +                     
                    " where prepod.actual = 1 and fam<>'-' " + 
                    " order by priority, fam, im, ot";
            main.global_adapter = new SqlDataAdapter(q, main.global_connection);
            main.global_adapter.Fill(preps);

            ListWindow lw = new ListWindow();
            lw.Text = "Выбор руководителя";
            lw.tbl = preps;

            DialogResult res = lw.ShowDialog();

            if (res==DialogResult.Cancel) return;

            ruk_id = lw.resId;
            ruk.Text = lw.str_res;

            lw.Dispose();
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            MessageBox.Show("скоро будет :)");
        }

        private void toolStripContainer1_ContentPanel_Load(object sender, EventArgs e)
        {
            // to be removed
        }
    }
}