using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace FSystem
{
    public partial class kursrab_archiv_edit : Form
    {
        public kursrab_archiv_edit()
        {
            InitializeComponent();
        }


        string sql = string.Empty;
        public string CurOtm = string.Empty;

        private void button1_Click(object sender, EventArgs e)
        {
            Random r = new Random();
            main.global_connection.Open();
            sql = sql = "select distinct gr = grupa.name + ' - ' + predmet.name, rid = rabota.id, " +
            " pname = predmet.name,otzyv2, otzyv3, otzyv4, otzyv5, opisanie, otzyv_title, vivod2, vivod3, vivod4, vivod5 " +
            " from grupa " +
            " join predmet on predmet.grupa_id = grupa.id " +
            " join tema_rabota on tema_rabota.predmet_id = predmet.id " +
            " join vid_rab on vid_rab.id = tema_rabota.vid_rabota_id " +
            " join rabota on rabota.predmet_id = predmet.id and rabota.vid_rab_id = vid_rab.id " +
            " where vid_rab.ID = 2 And Rabota.y = 2010 ";
            DataTable d = new DataTable();
            (new SqlDataAdapter(sql, main.global_connection)).Fill(d);

            Word.Application wa = new Word.Application();

            object template = Type.Missing;
            object newtemplate = false;
            object documentType = Word.WdNewDocumentType.wdNewBlankDocument;
            object visible = true;


            Word.Document doc = wa.Documents.Add(ref template,
                ref newtemplate,
                ref documentType,
                ref visible);

            object fileName = @"nod" + r.Next().ToString() + ".doc"; ;
            object fileFormat = Word.WdSaveFormat.wdFormatDocument;
            object lockComments = false;
            object password = "";
            object addToRecentFiles = false;
            object writePassword = "";
            object readOnlyRecommended = false;
            object embedTrueTypeFonts = false;
            object saveNativePictureFormat = false;
            object saveFormsData = false;
            object saveAsAOCELetter = Type.Missing;
            object encoding = Type.Missing;
            object insertLineBreaks = Type.Missing;
            object allowSubstitutions = Type.Missing;
            object lineEnding = Type.Missing;
            object addBiDiMarks = Type.Missing;

            Object begin = 0;
            Object end = 10;
            Word.Range wordrange = doc.Paragraphs[1].Range;
            wordrange.Select();

            wordrange.Font.Size = 11;
            wordrange.Font.Color = Word.WdColor.wdColorBlue;
            wordrange.Text = sql;

            int i = 0;
            int j = 2;
            for (; i < d.Rows.Count; i++)
            {
                doc.Paragraphs.Add(ref addBiDiMarks);
                wordrange = doc.Paragraphs[j].Range;
                wordrange.Select();
                wordrange.Text = d.Rows[i][0].ToString();
                j++;
            }


            doc.SaveAs(ref fileName,
                ref fileFormat,
                ref lockComments,
                ref password,
                ref addToRecentFiles,
                ref writePassword,
                ref readOnlyRecommended,
                ref embedTrueTypeFonts,
                ref saveNativePictureFormat,
                ref saveFormsData,
                ref saveAsAOCELetter,
                ref encoding,
                ref insertLineBreaks,
                ref allowSubstitutions,
                ref lineEnding,
                ref addBiDiMarks);

            object saveChages = false;
            object originalFormat = Word.WdOriginalFormat.wdWordDocument;
            object routeDocument = Type.Missing;
            wa.Quit(ref saveChages, ref originalFormat, ref routeDocument);
        }


        /// <summary>
        /// список учебных лет
        /// </summary>
        public DataTable UchGodList = null;

        /// <summary>
        /// заполнить список учебных лет
        /// </summary>
        public void FillUchGolList()
        {
            UchGodList = new DataTable();
            sql = string.Format("select * from uch_god where id < {0} order by start", main.uch_god);
            (new SqlDataAdapter(sql, main.global_connection)).Fill(UchGodList);
            int i = 0, k = 0;
            foreach (DataRow dr in UchGodList.Rows)
            {
                int y1 = Convert.ToDateTime(dr["start"]).Year;
                int y2 = Convert.ToDateTime(dr["finish"]).Year;
                string item = "Учебный год " +
                    y1.ToString() + " - " + y2.ToString();
                uchGodcomboBox.Items.Add(item);

                if (Convert.ToInt32(dr[0]) == main.uch_god)
                {
                    i = k;
                }
                k++;
            }

            uchGodcomboBox.SelectedIndex = i;
        }

        /// <summary>
        /// список курсовых работ
        /// </summary>
        DataTable KursRabotaList = null;

        /// <summary>
        /// заполнить список работ
        /// </summary>
        /// <param name="y">учебный год выполнения работы</param>
        public void FillKursRabotaList(int y)
        {
            KursRabotaList = new DataTable();
            sql = "select distinct gr = grupa.name + ' - ' + predmet.name, rid = rabota.id, " +
            " pname = predmet.name,otzyv2, otzyv3, otzyv4, otzyv5, opisanie, otzyv_title, vivod2, vivod3, vivod4, vivod5 " +
            " from grupa " +
            " join predmet on predmet.grupa_id = grupa.id " +
            " join tema_rabota on tema_rabota.predmet_id = predmet.id " +
            " join vid_rab on vid_rab.id = tema_rabota.vid_rabota_id " +
            " join rabota on rabota.predmet_id = predmet.id and rabota.vid_rab_id = vid_rab.id " +
            " where vid_rab.ID = 2 And Rabota.y = " + y.ToString();
            (new SqlDataAdapter(sql, main.global_connection)).Fill(KursRabotaList);

            kursRablistBox.Items.Clear();

            foreach (DataRow dr in KursRabotaList.Rows)
            {
                kursRablistBox.Items.Add(dr["gr"]);
            }

            if (kursRablistBox.Items.Count > 0)
                kursRablistBox.SelectedIndex = 0;
        }

        /// <summary>
        /// список студентов, тем и оценок
        /// </summary>
        DataTable StudRabotaList = null;

        /// <summary>
        /// получить список по данной работе
        /// </summary>
        /// <param name="RabotaID">ид работы, которую надо вывести</param>
        public void FillStudRabotaList(int RabotaID)
        {
            RabListdataGridView.Rows.Clear();

            sql = string.Format("select " +
                " student.id, isnull(vid_otmetka.id,-1), isnull(tema_rabota.name, ''),   " +  //0 1 2 
                " student_rabota.id, isnull(vid_otmetka.str_name,''), tema_id, isnull(tema_rabota.content, ''), " + //3 4 5 6
                " fio = student.fam + ' ' + left(student.im,1) + '. ' + left(student.ot,1) + '.', " + //7
                " ps = isnull(student_rabota.pred_status,0), rabota.id, isnull(vid_otmetka.str_alias,''),  " + //8  9 10
                " goal = isnull(tema_rabota.content,''), " +  // 11
                " grupa.name " + //12
                " from student_rabota " +
                "   join rabota on rabota.id = student_rabota.rabota_id " +
                "    join student on student.id = student_rabota.student_id " +
                "    join grupa on grupa.id = student.gr_id " +
                "    left outer join tema_rabota on tema_rabota.id = student_rabota.tema_id " +
                "    left outer join vid_otmetka on vid_otmetka.id = student_rabota.otmetka_id " +
                " where rabota.id = {0} and student.actual = 1 " +
                " order by student.fam, student.im, student.ot", RabotaID);

            StudRabotaList = new DataTable();
            (new SqlDataAdapter(sql, main.global_connection)).Fill(StudRabotaList);

            int i = 0;
            foreach (DataRow dr in StudRabotaList.Rows)
            {
                string otm = dr[1].ToString();
                if (otm == "-1") otm = "";
                RabListdataGridView.Rows.Add(new object[] { dr[7], dr[2], otm });
                RabListdataGridView.Rows[i].Tag = dr[3];
                i++;
            }

        }


        // --------- обработчики событий ---------------- 

        /// <summary>
        /// загрузка сведений о КР (начиная с учебного года)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void kursrab_archiv_edit_Load_1(object sender, EventArgs e)
        {
            FillUchGolList();
        }

        /// <summary>
        /// заполнение списка работ за указанный год при выборе учебного года
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void uchGodcomboBox_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            int y = Convert.ToDateTime(
                            UchGodList.Rows[uchGodcomboBox.SelectedIndex]["finish"]).Year;
            FillKursRabotaList(y);
        }

        private void kursRablistBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            int rab_id =
                Convert.ToInt32(KursRabotaList.Rows[kursRablistBox.SelectedIndex][1]);
            FillStudRabotaList(rab_id);
        }


        // запомнить тек значение
        private void RabListdataGridView_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            if (e.ColumnIndex == 2)
            {
                CurOtm = RabListdataGridView.Rows[e.RowIndex].Cells[2].Value.ToString();
            }
        }

        /// <summary>
        /// фиксация изменения оценки
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void RabListdataGridView_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex != 2)
            {
                return;
            }

            if (RabListdataGridView.Rows[e.RowIndex].Cells[1].Value.ToString().Trim().Length == 0)
            {
                MessageBox.Show("Нет темы. Редактирование невозможно.",
                    "Отказ операции", MessageBoxButtons.OK, MessageBoxIcon.Error);
                RabListdataGridView.Rows[e.RowIndex].Cells[2].Value = CurOtm;
                return;
            }

            string newotm = RabListdataGridView.Rows[e.RowIndex].Cells[2].Value.ToString();
            if (newotm.Length > 1)
            {
                RabListdataGridView.Rows[e.RowIndex].Cells[2].Value = CurOtm;
                return;
            }

            if (!Char.IsDigit(newotm[0]))
            {
                RabListdataGridView.Rows[e.RowIndex].Cells[2].Value = CurOtm;
                return;
            }
            else
            {
                int d = int.Parse(newotm);
                if (d < 2 || d > 5)
                {
                    RabListdataGridView.Rows[e.RowIndex].Cells[2].Value = CurOtm;
                    return;
                }
            }

            sql = "update student_rabota set otmetka_id = @OTMID where id = @ID";

            SqlCommand cmd = new SqlCommand(sql, main.global_connection);
            cmd.Parameters.Add("@OTMID", SqlDbType.Int).Value = newotm;
            cmd.Parameters.Add("@ID", SqlDbType.Int).Value =
                RabListdataGridView.Rows[e.RowIndex].Tag;
            cmd.ExecuteNonQuery();
        }

        private void RabListdataGridView_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (RabListdataGridView.Rows.Count == 0) return;

            //редактирование темы
            int row = 0;
            if (RabListdataGridView.CurrentCell != null)
                row = RabListdataGridView.CurrentCell.RowIndex;
            else
            {
                MessageBox.Show("Следует указать студента для выбора темы.",
                        "Ошибка операции", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            //вызвать окно редактирования темы
            tema_edit te = new tema_edit();
            te.predmet_id = main.id_predmet_in_tree;
            te.vid_rab_id = 2;
            te.Text = "Выбор темы КР по предмету: " + main.name_predmet_in_tree;

            DialogResult dr = te.ShowDialog();
            if (dr == DialogResult.Cancel) return;

            int tema_id = te.new_tema_id;

            //сохранить выбор темы для выбранного студента
            string q = "update student_rabota set tema_id = " +
                tema_id + " where id = " + RabListdataGridView.Rows[e.RowIndex].Tag.ToString();
            SqlCommand global_command = new SqlCommand(q, main.global_connection);
            global_command.ExecuteNonQuery();

            RabListdataGridView.Rows[row].Cells[1].Value = te.new_tema_name;
            te.Dispose();
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            int rows = RabListdataGridView.Rows.Count;

            if (rows == 0)
            {
                MessageBox.Show("Нет сведений о курсовых работах.",
                    "Откза операции", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (RabListdataGridView.SelectedCells.Count == 0)
            {
                MessageBox.Show("Выберите строку с работой, для которой нужно построить отзыв.",
                    "Откза операции", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            string rab_id = KursRabotaList.Rows[kursRablistBox.SelectedIndex][1].ToString();
            string otm_digit = RabListdataGridView.Rows[RabListdataGridView.SelectedCells[0].RowIndex].Cells[2].Value.ToString();
            string tema = RabListdataGridView.Rows[RabListdataGridView.SelectedCells[0].RowIndex].Cells[1].Value.ToString();
            string stud = RabListdataGridView.Rows[RabListdataGridView.SelectedCells[0].RowIndex].Cells[0].Value.ToString();
            string prname = KursRabotaList.Rows[kursRablistBox.SelectedIndex][2].ToString();
            string content = StudRabotaList.Rows[RabListdataGridView.SelectedCells[0].RowIndex][11].ToString();

            if (tema.Trim().Length == 0)
            {
                MessageBox.Show("Тема работы не задана. Невозможно построить отзыв.",
                    "Откза операции", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (otm_digit.Trim().Length == 0)
            {
                MessageBox.Show("Оценка работы не задана. Невозможно построить отзыв.",
                    "Откза операции", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            Word.Application wa = null;
            Word.Document Doc = main.CreateNewWordDoc(ref wa);

            string sql = "select * from rabota where id = " + rab_id;
            DataTable Rabota = new DataTable();
            (new SqlDataAdapter(sql, main.global_connection)).Fill(Rabota);

            saveExcel.Title = "Введите имя для файла рецензии.";
            saveExcel.Filter = "Файл рецензии в формате MS Word|*.doc";
            saveExcel.FileName = "Отзыв на курс.раб. " + stud + ".doc";

            if (saveExcel.ShowDialog() != DialogResult.OK) return;
            string FileName = saveExcel.FileName;

            Word.Range Range;
            object nullval = Type.Missing;

            Range = Doc.Paragraphs[1].Range;
            Range.Select();
            Range.Font.Size = 14;
            Range.Font.Bold = 1;
            Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
            Range.Text = "Отзыв на курсовую работу по дисциплине";

            Doc.Paragraphs.Add(ref nullval);
            Range = Doc.Paragraphs[2].Range;
            Range.Font.Bold = 1;
            Range.Text = "\"" + prname + "\"";

            Doc.Paragraphs.Add(ref nullval);
            Range = Doc.Paragraphs[3].Range;
            Range.Text = " ";

            Doc.Paragraphs.Add(ref nullval);
            Range = Doc.Paragraphs[4].Range;
            Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;
            Range.Font.Size = 12;
            Range.Text = "ФИО автора работы: " + stud;

            Doc.Paragraphs.Add(ref nullval);
            Range = Doc.Paragraphs[5].Range;
            Range.Text = " ";

            Doc.Paragraphs.Add(ref nullval);
            Range = Doc.Paragraphs[6].Range;
            Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;
            Range.Font.Size = 12;
            Range.Text = "Тема работы: " + tema;

            Doc.Paragraphs.Add(ref nullval);
            Range = Doc.Paragraphs[7].Range;
            Range.Text = " ";

            Doc.Paragraphs.Add(ref nullval);
            Range = Doc.Paragraphs[8].Range;
            Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
            Range.Font.Size = 12;
            Range.Font.Bold = 1;
            Range.Text = "Цели и задачи работы.";

            Doc.Paragraphs.Add(ref nullval);
            Range = Doc.Paragraphs[9].Range;
            Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
            Range.Font.Size = 12;
            Range.Font.Bold = 0;
            Range.Text = content;

            Doc.Paragraphs.Add(ref nullval);
            Range = Doc.Paragraphs[10].Range;
            Range.Text = " ";

            Doc.Paragraphs.Add(ref nullval);
            Range = Doc.Paragraphs[11].Range;
            Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
            Range.Font.Size = 12;
            Range.Font.Bold = 1;
            Range.Text = "Заключение о качестве выполнения работы.";

            Doc.Paragraphs.Add(ref nullval);
            Range = Doc.Paragraphs[12].Range;
            Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
            Range.Font.Size = 12;
            Range.Font.Bold = 0;
            Range.Text = Rabota.Rows[0]["otzyv" + otm_digit].ToString();

            Doc.Paragraphs.Add(ref nullval);
            Range = Doc.Paragraphs[13].Range;
            Range.Text = " ";

            Doc.Paragraphs.Add(ref nullval);
            Range = Doc.Paragraphs[14].Range;
            Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;
            Range.Font.Size = 12;
            Range.Font.Bold = 1;
            Range.Text = "Оценка работы.";

            Doc.Paragraphs.Add(ref nullval);
            Range = Doc.Paragraphs[15].Range;
            Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
            Range.Font.Size = 12;
            Range.Font.Bold = 0;
            Range.Text = Rabota.Rows[0]["vivod" + otm_digit].ToString();

            Doc.Paragraphs.Add(ref nullval);
            Range = Doc.Paragraphs[16].Range;
            Range.Text = " ";

            Doc.Paragraphs.Add(ref nullval);
            Range = Doc.Paragraphs[17].Range;
            Range.Text = " ";

            Doc.Paragraphs.Add(ref nullval);
            Range = Doc.Paragraphs[18].Range;
            Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;
            Range.Font.Size = 12;
            Range.Font.Bold = 1;

            string active_user_dolz = main.active_user_dolz.Substring(0, 1).ToUpper() + main.active_user_dolz.Substring(1);

            Range.Text = active_user_dolz + " кафедры " + main.active_user_kaf + "                             " + main.active_user_name;

            main.SaveWordDoc(FileName, ref Doc);
            main.WordQuit(wa);
            System.Diagnostics.Process.Start(FileName);
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            Word.Application wa = null;
            Word.Document doc = main.CreateNewWordDoc(ref wa);
            object nulval = Type.Missing;

            string prname = KursRabotaList.Rows[kursRablistBox.SelectedIndex][2].ToString();

            Word.Range Range = doc.Range(ref nulval, ref nulval);
            Range.Select();
            Range.ParagraphFormat.FirstLineIndent = 0.0f;

            Range = doc.Paragraphs[1].Range;
            Range.Select();
            Range.Text = "А К Т";
            Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            Range.Font.Bold = 1;
            Range.Font.Italic = 1;

            Range = main.AddWordDocParagraph(ref doc,
                "сдачи курсовых работ  в архив", Word.WdParagraphAlignment.wdAlignParagraphCenter);
            Range.Font.Bold = 0;
            Range.Font.Italic = 0;

            Range = main.AddWordDocParagraph(ref doc,
                " ", Word.WdParagraphAlignment.wdAlignParagraphCenter);

            Range = main.AddWordDocParagraph(ref doc,
                "Комиссия в составе:  архивариус Зуенок Е.В.",
                Word.WdParagraphAlignment.wdAlignParagraphLeft);

            Range = main.AddWordDocParagraph(ref doc,
                "\t\t\t" +
                " зам. декана ФИВТ     В.В. Семикина",
                Word.WdParagraphAlignment.wdAlignParagraphLeft);

            Range = main.AddWordDocParagraph(ref doc,
                "\t\t\t" +
                " зав. кафедры КТИС   И.К. Мазур",
                Word.WdParagraphAlignment.wdAlignParagraphLeft);

            Range = main.AddWordDocParagraph(ref doc,
                " ",
                Word.WdParagraphAlignment.wdAlignParagraphLeft);


            Range = main.AddWordDocParagraph(ref doc,
                "FSystem ИВТ, кафедра  “Компьютерные технологии и системы”",
                Word.WdParagraphAlignment.wdAlignParagraphLeft);


            Range = main.AddWordDocParagraph(ref doc,
                " ",
                Word.WdParagraphAlignment.wdAlignParagraphLeft);

            Range = main.AddWordDocParagraph(ref doc,
                "За " + main.year_start.Year.ToString() + "/" + main.year_end.Year.ToString() + " уч. год",
                Word.WdParagraphAlignment.wdAlignParagraphLeft);


            Range = main.AddWordDocParagraph(ref doc,
                " ",
                Word.WdParagraphAlignment.wdAlignParagraphLeft);

            Range = main.AddWordDocParagraph(ref doc,
                "Группа " + StudRabotaList.Rows[0][12].ToString().ToUpper(),
                Word.WdParagraphAlignment.wdAlignParagraphCenter);
            Range.Font.Bold = 1;

            Range = main.AddWordDocParagraph(ref doc,
                "Дисциплина “" + prname + "”",
                Word.WdParagraphAlignment.wdAlignParagraphCenter);
            Range.Font.Bold = 1;

            Range = main.AddWordDocParagraph(ref doc,
                " ",
                Word.WdParagraphAlignment.wdAlignParagraphLeft);
            Range.Font.Bold = 0;


            List<int> tabrows = new List<int>();
            foreach (DataGridViewCell c in RabListdataGridView.SelectedCells)
            {
                MessageBox.Show(c.Value.ToString());

                if (!tabrows.Contains(c.RowIndex))
                    tabrows.Add(c.RowIndex);
            }

            int k = 1;
            for (int i = 0; i < tabrows.Count; i++)
            {
                int ind = tabrows[i];
                if (RabListdataGridView.Rows[ind].Cells[2].Value.ToString() != "2")
                {
                    if (RabListdataGridView.Rows[ind].Cells[1].Value.ToString().Length != 0)
                    {
                        Range = main.AddWordDocParagraph(ref doc,
                            k.ToString() + ". " + RabListdataGridView.Rows[ind].Cells[0].Value.ToString() +
                            " (" + RabListdataGridView.Rows[ind].Cells[1].Value.ToString() + ")",
                            Word.WdParagraphAlignment.wdAlignParagraphLeft);
                        k++;
                    }
                }
            }

            if (k == 1)
            {
                MessageBox.Show("Акт не построен. В таблице нет работ, " + 
                    "которые могут быть актированы (либо у работ не указана тема,\nлибо все выбранные работы оценены на неудовлетворительно.",
                    "Отказ операции", MessageBoxButtons.OK, MessageBoxIcon.Error);
                main.WordQuit(wa);
                return;
            }

            saveExcel.Title = "Введите имя для файла акта курсовой работы.";
            saveExcel.Filter = "Файл акта КР в формате MS Word|*.doc";
            saveExcel.FileName = "Акт курс. работ по " + prname + ".doc";

            if (saveExcel.ShowDialog() != DialogResult.OK) return;
            string FileName = saveExcel.FileName;

            Range = main.AddWordDocParagraph(ref doc,
                " ",
                Word.WdParagraphAlignment.wdAlignParagraphLeft);

            Range = main.AddWordDocParagraph(ref doc,
                " ",
                Word.WdParagraphAlignment.wdAlignParagraphLeft);

            Range = main.AddWordDocParagraph(ref doc,
                "Итого сдано работ: " + (k - 1).ToString(),
                Word.WdParagraphAlignment.wdAlignParagraphLeft);
            Range.Font.Bold = 1;
            Range.Font.Underline = Word.WdUnderline.wdUnderlineSingle;

            Range = main.AddWordDocParagraph(ref doc,
                " ",
                Word.WdParagraphAlignment.wdAlignParagraphLeft);
            Range.ParagraphFormat.LineSpacing = wa.LinesToPoints(1.5f);
            Range.Font.Bold = 0;
            Range.Font.Underline = Word.WdUnderline.wdUnderlineNone;

            Range = main.AddWordDocParagraph(ref doc,
                "Преподаватель           " + main.active_user_name + "___________________",
                Word.WdParagraphAlignment.wdAlignParagraphLeft);

            Range = main.AddWordDocParagraph(ref doc,
                "Зав. кафедры КТИС   И.К. Мазур________________________",
                Word.WdParagraphAlignment.wdAlignParagraphLeft);

            Range = main.AddWordDocParagraph(ref doc,
                "Зам. декана ФИВТ     В.В. Семикина_____________________",
                Word.WdParagraphAlignment.wdAlignParagraphLeft);

            Range = main.AddWordDocParagraph(ref doc,
                "Архивариус           Зуенок Е.В._______________________",
                Word.WdParagraphAlignment.wdAlignParagraphLeft);

            Range = main.AddWordDocParagraph(ref doc,
                "Бухгалтер                   ________________________________",
                Word.WdParagraphAlignment.wdAlignParagraphLeft);

            Range = main.AddWordDocParagraph(ref doc,
                DateTime.Now.ToLongDateString(),
                Word.WdParagraphAlignment.wdAlignParagraphRight);

            main.SaveWordDoc(FileName, ref doc);
            wa.Visible = true;
            wa.Activate();
        }


    }
}
