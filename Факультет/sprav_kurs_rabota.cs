using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Windows.Forms;

namespace FSystem
{
    public partial class sprav_kurs_rabota : Form
    {
        public sprav_kurs_rabota()
        {
            InitializeComponent();
        }

        /// <summary>
        /// идентификатор курсовой работы
        /// </summary>
        public string rab_id = "";

        /// <summary>
        /// ид предмета курсовой работы
        /// </summary>
        public string predm_id = "";

        /// <summary>
        /// запрос для работы с БД
        /// </summary>
        public string sql = "";
        

        /// <summary>
        /// массив текстовых элементов
        /// </summary>
        Control[] txtbox;

        /// <summary>
        /// массив именованных параметров запроса
        /// </summary>
        string[] sqlparams;

        public DataTable KursRabTable = null;

        private void button3_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Здесь следует ввести шаблонную фразу заключения о качестве курсовой работы (для каждой отметки)." +
                "\nДанный текст будет по умолчанию добавлен в автоматически создаваемую рецензию курсовой работы." +
                "\n" +
                "\nПример фразы вывода: Вместе с тем, имеются несущественные недостатки в работе программного обеспечения. Программа не в полной мере является защищённой от ввода некорректных данных и при некотором критическом (специально подобранном) наборе входной информации может аварийно завершить свою работу.  Несмотря на указанный недостаток, все остальные моменты работы реализованы на достойном уровне, защита проведена грамотно, автор уверенно ответил на дополнительные вопросы, Таким образом, данная работа заслуживает оценки «хорошо»." +
                "\n" +
                "\nСпецифические замечания по конкретной работе студента можно ввести через таблицу курсовых работ.",
                "Справка", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void richTextBox1_Enter(object sender, EventArgs e)
        {
            ((TextBox)sender).BackColor = Color.LightYellow;
            ((TextBox)sender).ForeColor = Color.DarkGreen;
        }

        private void richTextBox1_Leave(object sender, EventArgs e)
        {
            ((TextBox)sender).BackColor = Color.White;
            ((TextBox)sender).ForeColor = Color.Black;
        }

        private void richTextBox5_Enter(object sender, EventArgs e)
        {
            ((TextBox)sender).BackColor = Color.LightYellow;
            ((TextBox)sender).ForeColor = Color.Navy;
        }

        private void richTextBox5_Leave(object sender, EventArgs e)
        {
            ((TextBox)sender).BackColor = Color.White;
            ((TextBox)sender).ForeColor = Color.Black;
        }

        private void rabotaNametextBox_Enter(object sender, EventArgs e)
        {
            // to be removed
        }

        private void rabotaNametextBox_Leave(object sender, EventArgs e)
        {
            // to be removed
        }



        private void sprav_kurs_rabota_Load(object sender, EventArgs e)
        {
            timer1.Start();
            KursRabTable = new DataTable();
            sql = "SELECT name, opisanie, " + 
                "otzyv2, otzyv3, otzyv4, otzyv5, " + 
                " vivod2, vivod3, vivod4, vivod5 " +
                " FROM  rabota WHERE id = " + rab_id;
            (new SqlDataAdapter(sql, main.global_connection)).Fill(KursRabTable);            

            DataRow dr = KursRabTable.Rows[0];

            txtbox = new Control[] { kursRabNametextBox, kursRabOpisanieTextBox,
                               richTextBox1, richTextBox2, richTextBox3, richTextBox4,
                               richTextBox5, richTextBox6, richTextBox7, richTextBox8};

            sqlparams = new string[] { "@NM", 
                                "@OPIS", 
                                "@OT2",
                                "@OT3",
                                "@OT4",
                                "@OT5",
                                "@VIV2",
                                "@VIV3",
                                "@VIV4",
                                "@VIV5"};            

            ((TextBox)txtbox[0]).Text = dr[0].ToString();
            ((TextBox)txtbox[1]).Text = dr[1].ToString();

            for (int i = 2; i < txtbox.Length; i++)
            {
                ((TextBox)txtbox[i]).Text = dr[i].ToString();
            }          
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (kursRabNametextBox.Text.Trim().Length == 0)
            {
                MessageBox.Show("Введите название курсовой работы!", "Отказ операции", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                DialogResult = DialogResult.Abort;
                return;
            }
            
            sql = "update rabota set ";

            int i = 0;

            for (; i < sqlparams.Length; i++)
            {
                if (i != sqlparams.Length - 1)
                    sql = sql + KursRabTable.Columns[i].ColumnName + " = " + sqlparams[i] + ", ";
                else
                    sql = sql + KursRabTable.Columns[i].ColumnName + " = " + sqlparams[i] + " ";
            }

            sql = sql + " where id = " + rab_id;

            SqlCommand cmd = new SqlCommand(sql, main.global_connection);
            for (i = 0; i < sqlparams.Length; i++)
            {
                cmd.Parameters.Add(sqlparams[i], SqlDbType.NVarChar).Value = txtbox[i].Text;
            }

            //MessageBox.Show(sql);

            cmd.ExecuteNonQuery();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            MessageBox.Show("Здесь следует ввести шаблонную фразу заключения о качестве курсовой работы (для каждой отметки)." +
                "\nДанный текст будет по умолчанию добавлен в автоматически создаваемую рецензию курсовой работы." +
                "\n" +
                "\nПример фразы вывода: Вместе с тем, имеются несущественные недостатки в работе программного обеспечения. Программа не в полной мере является защищённой от ввода некорректных данных и при некотором критическом (специально подобранном) наборе входной информации может аварийно завершить свою работу.  Несмотря на указанный недостаток, все остальные моменты работы реализованы на достойном уровне, защита проведена грамотно, автор уверенно ответил на дополнительные вопросы, Таким образом, данная работа заслуживает оценки «хорошо»." +
                "\n" +
                "\nСпецифические замечания по конкретной работе студента можно ввести через таблицу курсовых работ.",
                "Справка", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            MessageBox.Show("Здесь следует ввести шаблонную фразу отзыва для каждой отметки за курсовую работу." +
                "\nДанный текст будет по умолчанию добавлен в автоматически создаваемую рецензию курсовой работы." +
                "\n" +
                "\nПример фразы отзыва: Представленная работа состоит из пояснительной записки и программного обеспечения, разработанного автором в рамках решения поставленной перед ним проблемы. В целом автор достаточно подробно рассматривает явление языка программирования С++, указанное в теме. Им изучаются библиотеки типов данных и функции, реализующие суть изучаемой проблемы. Созданная программа демонстрирует, что автор не только изучил теорию вопроса, но и может применить её на практике для разработки программ. Представленный демонстрационный проект можно считать вполне удовлетворительно показывающим возможности изучаемого в работе явления. " +
                "\n" +
                "\nСпецифические замечания по конкретной работе студента можно ввести через таблицу курсовых работ.",
                "Справка", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            string txt = "";

            if (kursRabNametextBox.Text.Trim().Length == 0)
            {
                txt = "название работы";
            }

            if (kursRabOpisanieTextBox.Text.Trim().Length == 0)
            {
                if (txt.Length > 0)
                    txt = txt + ", " + "описание работы";
                else
                    txt = "описание работы";
            }

            if (richTextBox1.Text.Trim().Length == 0)
            {
                if (txt.Length > 0)
                    txt = txt + ", " + "отзыв на 2";
                else
                    txt = "отзыв на 2";
            }

            if (richTextBox2.Text.Trim().Length == 0)
            {
                if (txt.Length > 0)
                    txt = txt + ", " + "отзыв на 3";
                else
                    txt = "отзыв на 3";
            }

            if (richTextBox3.Text.Trim().Length == 0)
            {
                if (txt.Length > 0)
                    txt = txt + ", " + "отзыв на 4";
                else
                    txt = "отзыв на 4";
            }

            if (richTextBox4.Text.Trim().Length == 0)
            {
                if (txt.Length > 0)
                    txt = txt + ", " + "отзыв на 5";
                else
                    txt = "отзыв на 5";
            }

            if (richTextBox5.Text.Trim().Length == 0)
            {
                if (txt.Length > 0)
                    txt = txt + ", " + "вывод на 2";
                else
                    txt = "вывод на 2";
            }

            if (richTextBox6.Text.Trim().Length == 0)
            {
                if (txt.Length > 0)
                    txt = txt + ", " + "вывод на 3";
                else
                    txt = "вывод на 3";
            }

            if (richTextBox7.Text.Trim().Length == 0)
            {
                if (txt.Length > 0)
                    txt = txt + ", " + "вывод на 4";
                else
                    txt = "вывод на 4";
            }

            if (richTextBox8.Text.Trim().Length == 0)
            {
                if (txt.Length > 0)
                    txt = txt + ", " + "вывод на 5";
                else
                    txt = "вывод на 5";
            }

            if (txt.Length > 0)
            {
                StatusLabel.Text = "Не заполнены поля: " + txt;
            }
            else
            {
                StatusLabel.Text = string.Empty;
            }

            if (StatusLabel.ForeColor == Color.Red)
            {
                StatusLabel.ForeColor = Color.Blue;
            }
            else
            {
                StatusLabel.ForeColor = Color.Red;
            }

        }

        /// <summary>
        /// скопировать поля в шаблон отзыва из предыдущей работы 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            //получить сведения по предыдущей КР 

            int predyear = main.ends[main.ends.Count - 1].Year - 1;

            sql = "select * from rabota where predmet_id = @PRID and Y = @Y and vid_rab_id = 2";
            main.global_command = new SqlCommand(sql, main.global_connection);
            main.global_command.Parameters.Add("@PRID", SqlDbType.Int).Value = predm_id;
            main.global_command.Parameters.Add("@Y", SqlDbType.Int).Value = predyear;
            DataTable t = new DataTable();
            (new SqlDataAdapter(main.global_command)).Fill(t);

            if (t.Rows.Count == 0)
            {
                MessageBox.Show("Не найдено сведений о предыдщей работе по данному предмету.",
                    "Отказ операции", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            DataRow r = t.Rows[0];
            string res = "Имеются сведения о работе: " + r["name"].ToString() + "\n\n" + 
                "Примеры формулировок отзывов и выводов по работе:\n\n";


            string otz2 = r["otzyv2"].ToString();
            string viv2 = r["vivod2"].ToString();

            string otz3 = r["otzyv3"].ToString();
            string viv3 = r["vivod3"].ToString();

            string otz4 = r["otzyv4"].ToString();
            string viv4 = r["vivod4"].ToString();

            string otz5 = r["otzyv5"].ToString();
            string viv5 = r["vivod5"].ToString();

            if (otz2.Length > 0)
                res += "Отзыв на 2: " + otz2.Substring(0, 40) + " ... \n";
            else
                res += "Отзыв на 2: отсутствует\n";

            if (viv2.Length > 0)
                res += "Вывод на 2: " + viv2.Substring(0, 40) + " ... \n\n";
            else
                res += "Вывод на 2: отсутствует\n\n";

            // ---------------- кон на 2

            if (otz3.Length > 0)
                res += "Отзыв на 3: " + otz3.Substring(0, 40) + " ... \n";
            else
                res += "Отзыв на 3: отсутствует\n";

            if (viv3.Length > 0)
                res += "Вывод на 3: " + viv3.Substring(0, 40) + " ... \n\n";
            else
                res += "Вывод на 3: отсутствует\n\n";

            // ---------------- кон на 3

            if (otz4.Length > 0)
                res += "Отзыв на 4: " + otz4.Substring(0, 40) + " ... \n";
            else
                res += "Отзыв на 4: отсутствует\n";

            if (viv4.Length > 0)
                res += "Вывод на 4: " + viv4.Substring(0, 40) + " ... \n\n";
            else
                res += "Вывод на 4: отсутствует\n\n";

            // ---------------- кон на 4

            if (otz5.Length > 0)
                res += "Отзыв на 5: " + otz5.Substring(0, 40) + " ... \n";
            else
                res += "Отзыв на 5: отсутствует\n";

            if (viv5.Length > 0)
                res += "Вывод на 5: " + viv5.Substring(0, 40) + " ... \n\n";
            else
                res += "Вывод на 5: отсутствует\n\n";

            // ---------------- кон на 5

            DialogResult dres = MessageBox.Show(res + "\n\nПринять и скопировать данные в этот шаблон?", 
                "Запрос операции", 
                MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (dres == DialogResult.No) return;

            txtbox[2].Text = otz2;
            txtbox[3].Text = otz3;
            txtbox[4].Text = otz4;
            txtbox[5].Text = otz5;

            txtbox[6].Text = viv2;
            txtbox[7].Text = viv3;
            txtbox[8].Text = viv4;
            txtbox[9].Text = viv5;
        }
    }
}
