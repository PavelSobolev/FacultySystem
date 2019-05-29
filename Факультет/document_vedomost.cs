using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Windows.Forms;
using GemBox.Spreadsheet;
using System.IO;
using System.Diagnostics;
using System.Threading;

namespace FSystem
{
    public partial class document_vedomost : Form
    {
        public document_vedomost()
        {
            InitializeComponent();
        }

        SqlDataAdapter sdap;
        SqlCommand cmd;
        DataTable grupa_set, predmet_set, vid_set;


        private void document_vedomost_Load(object sender, EventArgs e)
        {
            fill_gr();
            fill_predm();
            fill_vid();
        }

        void fill_gr()
        {
            //заполнить список групп
            grupa_list.Items.Clear();
            string sql = string.Format("select id, name,kurs_id  from grupa where actual = 1 " +
                " and fakultet_id = {0}", main.fakultet_id);
            sdap = new SqlDataAdapter(sql, main.global_connection);
            grupa_set = new DataTable();
            sdap.Fill(grupa_set);
            foreach (DataRow r in grupa_set.Rows) grupa_list.Items.Add(r[1].ToString());

            grupa_list.SelectedIndex = 0;
        }

        void fill_predm()
        {
            int sem = 0;

            //заполнить список предметов этих групп
            predmet_list.Items.Clear();
            string sql = string.Format("select distinct  predmet.id, name, name_krat, semestr, " +
                " fio = fam + ' ' + im + ' ' + ot " + 
                " from predmet " + 
                " join prepod on prepod.id = predmet.prepod_id " + 
                " where predmet.actual=1 and " +
                " grupa_id = {0} and name not like '%тест%' and name not like '%срез%' " +
                //" and semestr % 2 = " + sem.ToString() +   
                " order by semestr, name", grupa_set.Rows[grupa_list.SelectedIndex][0]);
            

            sdap = new SqlDataAdapter(sql, main.global_connection);
            predmet_set = new DataTable();
            sdap.Fill(predmet_set);
            foreach (DataRow r in predmet_set.Rows) predmet_list.Items.Add(r[1].ToString() +
                " [семестр № " + r[3].ToString() + " ]");

            if (predmet_list.Items.Count > 0)
                predmet_list.SelectedIndex = 0;
        }


        void fill_vid()
        {
            if (predmet_list.Items.Count == 0) return;

            //заполнить виды контроля
            vid_zan_list.Items.Clear();
            vid_zan_list.Items.Add("межсессионная аттестация");
            vid_zan_list.Items.Add("проверка остаточных знаний");
            vid_zan_list.Items.Add("производственная практика");
            string sql = string.Format("select vid_zan.id, name, title_ved, kod " +
                " from vid_zan " +
                " join vidzan_predmet on vidzan_predmet.vidzan_id = vid_zan.id " +
                " where is_kontrol=1 and vidzan_predmet.predmet_id = {0} " +
                " order by name", predmet_set.Rows[predmet_list.SelectedIndex][0]);

            sdap = new SqlDataAdapter(sql, main.global_connection);
            vid_set = new DataTable();
            sdap.Fill(vid_set);
            foreach (DataRow r in vid_set.Rows) vid_zan_list.Items.Add(r[1].ToString());

            vid_zan_list.SelectedIndex = 0;
        }

        private void grupa_list_SelectedIndexChanged(object sender, EventArgs e)
        {
            fill_predm();
            fill_vid();
        }

        private void predmet_list_SelectedIndexChanged(object sender, EventArgs e)
        {
            fill_vid();
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            predmet_list.Enabled = !free_form.Checked;
            vid_zan_list.Enabled = !free_form.Checked;
            if (free_form.Checked)
            {
                title.Text = "Ведомость ... ";
                title.BackColor = Color.Yellow;
            }
            else
            {
                vid_zan_list_SelectedIndexChanged(sender, new EventArgs());
                title.BackColor = Color.White;
            }
        }

        private void vid_zan_list_SelectedIndexChanged(object sender, EventArgs e)
        {
            int k = vid_zan_list.SelectedIndex;

            switch (k)
            {
                case 0: title.Text = "Ведомость межсессионной аттестации"; break;
                case 1: title.Text = "Ведомость проверки остаточных знаний"; break;
                case 2: title.Text = "Ведомость производственной практики"; break;
                default:
                    title.Text = vid_set.Rows[k-3][2].ToString(); break;
            }
        }


        string toRoman(int n)
        {
            string res = string.Empty;

            switch (n)
            {
                case 1: res = "I"; break;
                case 2: res = "II"; break;
                case 3: res = "III"; break;
                case 4: res = "IV"; break;
                case 5: res = "V"; break;
                case 6: res = "VI"; break;
                case 7: res = "VII"; break;
                case 8: res = "VIII"; break;
                case 9: res = "XI"; break;
                case 10: res = "X"; break;
                case 11: res = "XI"; break;
                case 12: res = "XII"; break;
                case 13: res = "XIII"; break;
            }

            return res;
        }


        string toRoman(string n)
        {
            string res = string.Empty;

            switch (n)
            {
                case "1": res = "I"; break;
                case "2": res = "II"; break;
                case "3": res = "III"; break;
                case "4": res = "IV"; break;
                case "5": res = "V"; break;
                case "6": res = "VI"; break;
                case "7": res = "VII"; break;
                case "8": res = "VIII"; break;
                case "9": res = "XI"; break;
                case "10": res = "X"; break;
                case "11": res = "XI"; break;
                case "12": res = "XII"; break;
                case "13": res = "XIII"; break;
            }

            return res;
        }

        public void SaveToExcel()
        {

            string FileName = ""; //имя файла для сохранения            

            CellRange cr;
            ExcelFile excel = new ExcelFile();

            ExcelWorksheet sheet = excel.Worksheets.Add("Ведомость");
            sheet.PrintOptions.HeaderMargin = 0.0;
            sheet.PrintOptions.FooterMargin = 0.0;

            //задать общие свойства свойства
            sheet.Columns[0].Width = 3 * 256;
            cr = sheet.Cells.GetSubrange("a1", "i48");
            cr.Merged = true;
            cr.Style.Font.Name = "Times New Roman";
            cr.Style.Font.Size = 11 * 20;           
            cr.Style.VerticalAlignment = VerticalAlignmentStyle.Center;
            cr.Style.HorizontalAlignment = HorizontalAlignmentStyle.Left;
            cr.Style.FillPattern.SetSolid(Color.White);
            cr.Merged = false;
            

            cr = sheet.Cells.GetSubrange("a1", "i1");
            cr.Merged = true;
            cr.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            cr.Style.Font.Weight = ExcelFont.BoldWeight;
            sheet.Cells["A1"].Value = "Южно-Сахалинский институт экономики, права и информатики";
            //cr.Merged = false;

            cr = sheet.Cells.GetSubrange("a2", "i2");
            cr.Merged = true;
            cr.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            cr.Style.VerticalAlignment = VerticalAlignmentStyle.Bottom;
            cr.Style.Font.Weight = ExcelFont.BoldWeight;
            sheet.Cells["A2"].Value = title.Text.ToUpper();
            sheet.Rows[1].Height = 30 * 20;
            //cr.Merged = false;

            cr = sheet.Cells.GetSubrange("a4", "i4");
            cr.Merged = true;
            cr.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            sheet.Cells["A4"].Value = "от " + ved_date.Value.ToLongDateString();
            //cr.Merged = false;


            string frm = "", frm2 = "", kod = "", chas_string = ", Количество часов _____";
            int num = vid_zan_list.SelectedIndex;
            
            if (num>2)
                kod = vid_set.Rows[num-3][3].ToString();
            else
            {
                switch(num)
                {
                    case 0: kod = "ма"; break;
                    case 1: kod = "пост"; break;
                    case 2: kod = "прпр"; break;
                }
            }

            switch (kod)
            {
                case "э": frm = "Начало экзамена:    "; frm2 = "Экзам. оценка"; break;
                case "з": frm = "Начало зачёта:    "; frm2 = "Отметка о сдаче зачёта"; break;
                case "ма": frm = ""; frm2 = "Аттест. оценка"; chas_string = ""; break;
                case "прпр": frm = "Дата начала:    "; frm2 = "Оценка за произв. практику"; chas_string = ""; break;
                case "пост": frm = "Начало:    "; frm2 = "Оценка"; chas_string = ""; break;
                case "зкр": frm = "Начало защиты:    "; frm2 = "Оценка защиты"; break;
                case "кнр": frm = "Начало контр. работы:    "; frm2 = "Оценка контр. работы"; chas_string = ""; break;
                case "дз": frm = "Начало зачёта:    "; frm2 = "Зачётная оценка"; break;
            }

            cr = sheet.Cells.GetSubrange("a6", "i6");
            cr.Merged = true;
            cr.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            sheet.Rows[5].Height = 17 * 20;

            if (kod != "ма" && kod!="пост")
            {
                if (kod != "прпр")
                    sheet.Cells["A6"].Value = frm + ved_time.Value.ToShortTimeString() + "            Окончание ________";
                else
                    sheet.Cells["A6"].Value = frm + "________            Окончание ________";
            }
            else
            {
                sheet.Rows[4].Height = 3 * 20;
                sheet.Rows[5].Height = 3 * 20;
            }

            cr = sheet.Cells.GetSubrange("a7", "i7");
            cr.Merged = true;
            cr.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            sheet.Rows[6].Height = 17 * 20;
            sheet.Cells["A7"].Value = "FSystem:    " + main.fakultet_name;


            cr = sheet.Cells.GetSubrange("a8", "i8");
            cr.Merged = true;
            cr.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            sheet.Rows[7].Height = 17 * 20;
            sheet.Cells["A8"].Value =
                "Курс:    " + toRoman(grupa_set.Rows[grupa_list.SelectedIndex][2].ToString()) + ",    " +
                "Группа:    " + grupa_set.Rows[grupa_list.SelectedIndex][1].ToString() + ",    " +
                "Семестр:    " + toRoman(predmet_set.Rows[predmet_list.SelectedIndex][3].ToString()) + ",    " +
                main.starts[0].Year.ToString() + "/" + main.ends[main.ends.Count - 1].Year.ToString() + " уч. г.";


            cr = sheet.Cells.GetSubrange("a9", "i9");
            cr.Merged = true;
            cr.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            sheet.Rows[8].Height = 17 * 20;
            
            if (!free_form.Checked)
            {
                sheet.Cells["A9"].Value = "Дисциплина:    " + predmet_set.Rows[predmet_list.SelectedIndex][1].ToString() +
                    chas_string;
            }
            else
            {
                sheet.Cells["A9"].Value = "Дисциплина: ____________________________________________________ " +
                    chas_string;
            }

            cr = sheet.Cells.GetSubrange("a10", "i10");
            cr.Merged = true;
            cr.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            sheet.Rows[9].Height = 17 * 20;
            sheet.Cells["A10"].Value = "Экзаменатор:    " + predmet_set.Rows[predmet_list.SelectedIndex][4].ToString();

            
            DataTable stud_set = new DataTable();
            string sql = "select id, fio = fam + ' ' + im + ' ' + isnull(ot,''), zach_kn_number " +
                " from student " +
                " where actual = 1 and fam<>'0' and gr_id = " + grupa_set.Rows[grupa_list.SelectedIndex][0].ToString() + 
                " order by fam, im, ot ";
            sdap = new SqlDataAdapter(sql, main.global_connection);
            sdap.Fill(stud_set);

            if (stud_set.Rows.Count == 0)
            {
                //нельзя
                DialogResult r = 
                MessageBox.Show("Нет студентов в списке группы. Заполните список группы в разделе:\n" + 
                    "\tСправочники/Группы/Списочный состав", 
                    "Нет данных", MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                return;
            }

            cr = sheet.Cells.GetSubrange("a12", "i" + (12 + stud_set.Rows.Count).ToString());
            cr.Merged = true;
            cr.Style.Font.Size = 10*20;
            cr.SetBorders(MultipleBorders.Horizontal | MultipleBorders.Vertical , Color.Black, LineStyle.Thin);
            cr.Merged = false;

            sheet.Cells["A12"].Value = "№п/п";
            sheet.Cells["a12"].Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            
            cr = sheet.Cells.GetSubrange("b12", "d12");            
            cr.Merged = true;
            cr.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            sheet.Cells["B12"].Value = "Фамилия, имя, отчество";
            
            sheet.Cells["e12"].Value = "№ зачётн.\nкнижки";
            sheet.Cells["e12"].Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;

            cr = sheet.Cells.GetSubrange("f12", "g12");            
            cr.Merged = true;
            cr.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            sheet.Cells["f12"].Value = frm2;

            cr = sheet.Cells.GetSubrange("h12", "i12");            
            cr.Merged = true;
            cr.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            sheet.Cells["h12"].Value = "Подпись экзаменатора";            

            
            int i = 0;
            for (i = 0; i < stud_set.Rows.Count; i++)
            {
                sheet.Cells["A" + (i + 13).ToString()].Value = (i + 1).ToString();

                cr = sheet.Cells.GetSubrange("b" + (i + 13).ToString(), "d" + (i + 13).ToString());
                cr.Merged = true;
                cr.Style.HorizontalAlignment = HorizontalAlignmentStyle.Left;
                sheet.Cells["B" + (i + 13).ToString()].Value = stud_set.Rows[i][1].ToString();

                sheet.Cells["E" + (i + 13).ToString()].Value = stud_set.Rows[i][2].ToString();

                cr = sheet.Cells.GetSubrange("f" + (i + 13).ToString(), "g" + (i + 13).ToString());
                cr.Merged = true;

                cr = sheet.Cells.GetSubrange("h" + (i + 13).ToString(), "i" + (i + 13).ToString());
                cr.Merged = true;
            }

            int bottom = i, bottom2 = i;

            if (radio_zachet.Checked)
            {
                sheet.Cells["A" + (i + 14).ToString()].Value = 
                "Зачтено__________________";
                i++;
                sheet.Cells["A" + (i + 14).ToString()].Value = 
                "Не зачтено_______________";
                i++;
            }

            if (radio_otmetka.Checked)
            {
                sheet.Cells["A" + (i + 14).ToString()].Value = 
                "Отлично________________";
                i++;
                sheet.Cells["A" + (i + 14).ToString()].Value = 
                "Хорошо_________________";
                i++;
                sheet.Cells["A" + (i + 14).ToString()].Value = 
                "Удовлетворительно________";
                i++;
                sheet.Cells["A" + (i + 14).ToString()].Value = 
                "Неудовлетворительно______";
                i++;                
            }

            if (checkBox1.Checked)
            {
                sheet.Cells["A" + (i + 14).ToString()].Value = 
                "Уровень обученности______";
                i++;
            }

            sheet.Cells["A" + (i + 14).ToString()].Value = 
            "Не явился_______________";
            
            sheet.Cells["f" + (bottom + 14).ToString()].Value =
            "Итого сдавали _______________________";
            bottom++;
            sheet.Cells["f" + (bottom + 14).ToString()].Value = 
            "Подпись секретаря деканата ____________";
            bottom++;
            sheet.Cells["f" + (bottom + 14).ToString()].Value = 
            "Подпись экзаменатора _________________";
            bottom++;
            sheet.Cells["f" + (bottom + 14).ToString()].Value = 
            "Подпись декана FSystemа _____________";

            FileName = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) +
                "\\ " + title.Text + " (" + predmet_list.Items[predmet_list.SelectedIndex].ToString() + ")" + ".xls";

            // --------------- сохранение и открытие --------------
            if (File.Exists(FileName))
            {
                try
                {
                    File.Delete(FileName);
                }
                catch
                {
                    MessageBox.Show("Невозможно сохранить файл на диск, так как файл уже существует и открыт в окне программы Excel.\n" +
                        "Повторите операцию и выберите другое имя файла",
                        "Ошибка создания",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            
            excel.SaveXls(FileName);

            Thread.Sleep(500);

            try
            {
                Process.Start(FileName);
            }
            catch (Exception exx)
            {
                //не получилось открыть, сообщить об ошибке открытия и преложить найти самостоятельно
                MessageBox.Show("Невозможно открыть файл:" + FileName + "\n" +
                        "Попробуйте открыть файл самостоятельно или повторите опеерацию конвертирования.",
                        "Ошибка создания",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SaveToExcel();
        }

    }
}