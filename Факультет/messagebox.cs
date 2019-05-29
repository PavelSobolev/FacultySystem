using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Windows.Forms;

namespace FSystem
{
    public partial class messagebox : Form
    {

        /// <summary>
        /// запрос к БД
        /// </summary>
        public string query = "";

        /// <summary>
        /// таблица групп
        /// </summary>
        public DataTable grups = new DataTable();

        /// <summary>
        /// таблицы расписания
        /// </summary>
        public DataTable local_rasp = new DataTable();

        /// <summary>
        /// массив идентификаторов занятий
        /// </summary>
        public List<int> ids = new List<int>();
        /// <summary>
        /// количество групп
        /// </summary>
        int N = 0;


        public bool changed = false;

        /// <summary>
        /// выставить расписание на день
        /// </summary>
        /// <param name="year">год</param>
        /// <param name="mon">месяц</param>
        /// <param name="day">день</param>
        /// <param name="prep_id">ид преподавателя</param>
        /// <param name="pr_name">имя преподавателя</param>
        
        int year, mon, day, prep_id; 
        string pr_name;

        public messagebox (int _year, int _mon, int _day, int _prep_id, string _pr_name)
        {
            InitializeComponent ( );

            year = _year;
            mon = _mon;
            day = _day;
            prep_id = _prep_id;
            pr_name = _pr_name;

            query = "select distinct grupa_id, grupa.name from rasp " +
                " join grupa on grupa.id = rasp.grupa_id " +
                " where y=@YEAR and m=@MON and d=@DAY and " +
                " rasp.prepod_id=@PREP_ID and rasp.fakultet_id=@F_ID";

            SqlCommand cmd = new SqlCommand(query, main.global_connection);
            cmd.Parameters.Add("@YEAR", SqlDbType.Int).Value = year;
            cmd.Parameters.Add("@MON", SqlDbType.Int).Value = mon;
            cmd.Parameters.Add("@DAY", SqlDbType.Int).Value = day;
            cmd.Parameters.Add("@PREP_ID", SqlDbType.Int).Value = prep_id;
            cmd.Parameters.Add("@F_ID", SqlDbType.Int).Value = main.fakultet_id;

            SqlDataAdapter sda = new SqlDataAdapter(cmd);
            sda.Fill(grups);

            query = "select rasp.grupa_id, grupa.name, nom_zan, rasp.kol_chas," +
                " vid_zan.krat_name, predmet.name_krat, rasp.id, vid_zan.name  " +
                " from rasp " +
                " join grupa on grupa.id = rasp.grupa_id " +
                " join vid_zan on vid_zan.id = rasp.vid_zan_id " +
                " join predmet on predmet.id = rasp.predmet_id " +
                " where y=@YEAR and m=@MON and d=@DAY and " +
                " rasp.prepod_id=@PREP_ID and rasp.fakultet_id=@F_ID";

            cmd = new SqlCommand(query, main.global_connection);
            cmd.Parameters.Add("@YEAR", SqlDbType.Int).Value = year;
            cmd.Parameters.Add("@MON", SqlDbType.Int).Value = mon;
            cmd.Parameters.Add("@DAY", SqlDbType.Int).Value = day;
            cmd.Parameters.Add("@PREP_ID", SqlDbType.Int).Value = prep_id;
            cmd.Parameters.Add("@F_ID", SqlDbType.Int).Value = main.fakultet_id;

            sda = new SqlDataAdapter(cmd);
            sda.Fill(local_rasp);


            if (local_rasp.Rows.Count == 0) return;

            DateTime d = new DateTime(year, mon, day);
            string dayname = "";

            switch (d.DayOfWeek)
            {
                case DayOfWeek.Sunday: dayname = "воскресенье"; break;
                case DayOfWeek.Monday: dayname = "понедельник"; break;
                case DayOfWeek.Tuesday: dayname = "вторник"; break;
                case DayOfWeek.Thursday: dayname = "четверг"; break;
                case DayOfWeek.Wednesday: dayname = "среда"; break;
                case DayOfWeek.Friday: dayname = "пятница"; break;
                case DayOfWeek.Saturday: dayname = "суббота"; break;
            }

            Text = "Распределение часов: " + d.ToLongDateString() + " (" + dayname + ")";
            prepod_name.Text = pr_name;

            N = grups.Rows.Count;
            rsp.Rows = 7;
            rsp.Cols = N + 1;

            rsp.ColumnCollection[0].Style.BackColor = Color.LightYellow;
            rsp.ColumnCollection[0].Style.ForeColor = Color.Blue;
            rsp.ColumnCollection[0].Style.Font = new Font("Tahoma", 11, FontStyle.Bold);
            rsp.set_RowHeight(0, 20);

            for (int k = 0; k < 6 * N; k++)
            {
                ids.Add(0);
            }

            for (int n = 1; n <= 6; n++)
            {
                rsp[n, 0] = n;
            }

            rsp[0, 0] = "№";

            int i = 1;
            foreach (DataRow gr in grups.Rows)
            {
                rsp[0, i] = gr[1].ToString();
                rsp.set_Cell(C1.Win.C1FlexGrid.Classic.CellPropertySettings.flexcpBackColor, 0, i,
                    Color.LightYellow);
                rsp.set_Cell(C1.Win.C1FlexGrid.Classic.CellPropertySettings.flexcpFont, 0, i,
                    new Font("Tahoma", 11, FontStyle.Bold));
                rsp.set_Cell(C1.Win.C1FlexGrid.Classic.CellPropertySettings.flexcpForeColor, 0, i,
                    Color.Blue);

                DataRow[] gr_row = local_rasp.Select(" grupa_id = " + gr[0].ToString());

                foreach (DataRow gr_rasp in gr_row)
                {
                    int num = (int)gr_rasp[2];
                    double chas = (double)gr_rasp[3];

                    rsp[num, i] = gr_rasp[5].ToString() + "\n" +
                         gr_rasp[4].ToString() + " | " + string.Format("{0:F2}", chas) + " ч.";

                    ids[(num - 1) * N + i - 1] = (int)gr_rasp[6];
                }

                i++;
            }
        }       

        //переместиь курсор при правом щелчке
        private void rsp_MouseDown(object sender, MouseEventArgs e)
        {
            rsp.Col = rsp.MouseCol;
            rsp.Row = rsp.MouseRow;
        }

        private void rsp_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            set_chas();
        }

        public void set_chas()
        {
            //поменять количество часов для данного занятия
            int c = rsp.Col;
            int r = rsp.Row;
            int num = ids[(r - 1) * N + c - 1];

            DataRow[] gr_row = local_rasp.Select(" id = " + num.ToString());


            if (c == 0 || r == 0) return;
            if (rsp[r, c] == null) return;

            inputbox ib = new inputbox(
                "Введите в окно редактирования количество часов\n" +
                "по указанному виду занятия.\n\nЦелая часть числа отделяется от дробной знаком 'запятая' (,).",
                gr_row[0][7].ToString(),
                gr_row[0][3].ToString(),
                "Кол-во часов:");
            ib.is_numeric = true;

            DialogResult res;
            double ch = 0.0;

            do
            {
                res = ib.ShowDialog();
                if (res == DialogResult.Cancel) return;
            }
            while (res != DialogResult.OK);

            if (res == DialogResult.OK)
            {
                ch = Convert.ToDouble(ib.textBox1.Text);

                if (!(ch <= 200.0 && ch >= 0.0))
                {
                    MessageBox.Show("Введенно недопустимое количество часов [допустимое значение: от 1 до 200]",
                        "Ошибка ввода",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    ib.Dispose();
                    return;
                }
            }

            string q = "update rasp set kol_chas = @CHAS where id = @ID";
            SqlCommand cmd = new SqlCommand(q, main.global_connection);
            cmd.Parameters.Add("@CHAS", SqlDbType.Float).Value = ch;
            cmd.Parameters.Add("@ID", SqlDbType.Int).Value = num;            

            try
            {
                cmd.ExecuteNonQuery();
            }
            catch(Exception exx)
            {
                MessageBox.Show("Ошибка при передаче данных. Повторите операцию позднее.",
                    "Ошибка данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            changed = true;

            rsp[r, c] = gr_row[0][5].ToString() + "\n" +
                gr_row[0][4].ToString() + " | " + string.Format("{0:F2}", ch) + " ч.";

            ib.Dispose();
        }

        private void rsp_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
                set_chas();

        }
    }
}