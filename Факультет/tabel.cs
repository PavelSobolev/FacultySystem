using GemBox.Spreadsheet;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace FSystem
{
    public partial class tabel : Form
    {
        public tabel()
        {
            InitializeComponent();
        }

        public DataTable prepod_set;
        public DataTable show_table; //набор данных для вывода в таблицу и в файл
        public DataTable rasp; //расписание за данный месяц

        public string month_filter = "";
        public string prepod_filter = "";
        public string year_filter = "";

        public int[] days = new int[] { 30, 31, 30, 31, 31, 28, 31, 30, 31, 30, 31, 31 };
        public int[] nums = new int[] { 9, 10, 11, 12, 1, 2, 3, 4, 5, 6, 7, 8 };
        public int[] num_index = new int[] { 4, 5, 6, 7, 8, 9, 10, 11, 0, 1, 2, 3 };
        public int[] sem = new int[] { 1, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1 };

        public int current_m = 0, current_s = 0, current_y = 0;

        int current_day = 0, current_id = 0, current_row = 0;


        private void tabel_Load(object sender, EventArgs e)
        {
            months_list.SelectedIndex = num_index[DateTime.Now.Month - 1];

            current_s = sem[DateTime.Now.Month - 1];
            current_m = nums[months_list.SelectedIndex];

            month_filter = " m = " + current_m.ToString();

            if (current_m >= 9 && current_m <= 12)
            {
                current_y = main.starts[0].Year;
            }
            else
            {
                current_y = main.ends[main.ends.Count - 1].Year;
            }

            year_filter = " y = " + current_y.ToString();

        }

        //изменить ттекущий месяц и семестр, заполнить список преподов
        private void months_list_SelectedIndexChanged(object sender, EventArgs e)
        {
            current_m = nums[months_list.SelectedIndex];
            current_s = sem[current_m - 1];
            month_filter = " m = " + current_m.ToString();


            if (current_m >= 9 && current_m <= 12)
            {
                current_y = main.starts[0].Year;
            }
            else
            {
                current_y = main.ends[main.ends.Count - 1].Year;
            }

            year_filter = " y = " + current_y.ToString();

            fill_prepods();            
        }

        //заполнить список преподавателей
        public void fill_prepods()
        {
            prepod_list.Items.Clear();

            string cmd = " select distinct prepod.id ,  fio = fam + ' '  + im + ' ' + ot, " +
                " fiokr = fam + ' '  + left(im,1) + '. ' + left(ot,1) + '.' " +
                " from prepod " +
                " join predmet on predmet.prepod_id = prepod.id " +
                " join rasp on prepod.id = rasp.prepod_id and " + 
                " rasp.predmet_id = predmet.id " + //left outer
                " where " +
                " predmet.fakultet_id  = " + main.fakultet_id.ToString() +
                " and prepod.actual = 1 and predmet.semestr%2 = " 
                + current_s.ToString() +
                " and y between " + main.year_start.Year.ToString() + 
                " and " + main.year_end.Year.ToString() + 
                " and m = " + nums[months_list.SelectedIndex].ToString() + 
                " and predmet.actual =1 order by fio ";

            SqlDataAdapter prep_adapter = new SqlDataAdapter(cmd, 
                main.global_connection);
            prepod_set = new DataTable();
            prep_adapter.Fill(prepod_set);

            prepod_list.Items.Add("Все");
            foreach (DataRow row in prepod_set.Rows)
                prepod_list.Items.Add(row[1]);

            prepod_list.SelectedIndex = 0;
            prepod_filter = "";

            //fill_tabel();
        }

        //заполнить сетку табеля
        public void fill_tabel()
        {
            Application.DoEvents();

            tabel_grid.Clear(C1.Win.C1FlexGrid.ClearFlags.Content);

            int finish = days[months_list.SelectedIndex];

            int i = 0;

            switch (finish)
            {
                case 28:
                    tabel_grid.ColumnCollection[30].Visible = false;
                    tabel_grid.ColumnCollection[31].Visible = false;
                    tabel_grid.ColumnCollection[32].Visible = false;
                    break;
                case 29:
                    tabel_grid.ColumnCollection[30].Visible = true;
                    tabel_grid.ColumnCollection[31].Visible = false;
                    tabel_grid.ColumnCollection[32].Visible = false;
                    break;
                case 30:
                    tabel_grid.ColumnCollection[30].Visible = true;
                    tabel_grid.ColumnCollection[31].Visible = true;
                    tabel_grid.ColumnCollection[32].Visible = false;
                    break;
                case 31:
                    tabel_grid.ColumnCollection[30].Visible = true;
                    tabel_grid.ColumnCollection[31].Visible = true;
                    tabel_grid.ColumnCollection[32].Visible = true;
                    break;
            }

            tabel_grid.ColumnCollection[0].Style.TextAlign = 
                C1.Win.C1FlexGrid.TextAlignEnum.LeftCenter;
            tabel_grid.ColumnCollection[0].Style.ForeColor = 
                Color.Navy;
            tabel_grid.ColumnCollection[0].Style.TextEffect = 
                C1.Win.C1FlexGrid.TextEffectEnum.Inset;

            tabel_grid.ColumnCollection[1].Style.TextAlign = 
                C1.Win.C1FlexGrid.TextAlignEnum.CenterCenter;
            tabel_grid.ColumnCollection[1].Style.ForeColor = 
                Color.Red;
            tabel_grid.ColumnCollection[1].Style.TextEffect = 
                C1.Win.C1FlexGrid.TextEffectEnum.Inset;


            tabel_grid[0, 0] = "ФИО преподавателя";
            tabel_grid[0, 1] = "ИТОГО";
           
            for (i = 1; i <= finish; i++)
            {
                tabel_grid[0, i + 1] = i.ToString();
                tabel_grid.ColumnCollection[i + 1].Style.TextAlign =
                    C1.Win.C1FlexGrid.TextAlignEnum.CenterCenter;

                DateTime d = new DateTime(current_y, current_m, i);

                if (d.DayOfWeek == DayOfWeek.Sunday || d.DayOfWeek == DayOfWeek.Saturday)
                    tabel_grid.ColumnCollection[i + 1].Style.BackColor = Color.AliceBlue;
                else
                    tabel_grid.ColumnCollection[i + 1].Style.BackColor = Color.White;
            }


            tabel_grid[0, 33] = "ВСЕГО";
            tabel_grid.ColumnCollection[32].Style.TextAlign = C1.Win.C1FlexGrid.TextAlignEnum.CenterCenter;
            tabel_grid.ColumnCollection[32].Style.ForeColor = Color.Red;
            tabel_grid.ColumnCollection[32].Style.TextEffect = C1.Win.C1FlexGrid.TextEffectEnum.Inset;
            tabel_grid[0, 34] = "Инд. раб.";
            tabel_grid[0, 35] = "Контр. раб.";
            tabel_grid[0, 36] = "Зачёты";
            tabel_grid[0, 37] = "Консульт.";
            tabel_grid[0, 38] = "ВКР";
            tabel_grid[0, 39] = "Гос. экз.";
            tabel_grid[0, 40] = "Обз. лекц.";
            tabel_grid[0, 41] = "Экзамены";
            tabel_grid[0, 42] = "Курс. раб.";
            tabel_grid[0, 43] = "ИТОГО";

            //получить часы
            string q = "select id, prepod_id, kol_chas, potok_id, d from rasp " +
                " where " +
                year_filter +
                " and   " + month_filter +
                " and   fakultet_id = " + main.fakultet_id.ToString() +
                prepod_filter;

            rasp = new DataTable();
            SqlDataAdapter sda = new SqlDataAdapter(q, main.global_connection);
            sda.Fill(rasp);


            //пройти по дням и поставить часы у каждого препоода (сумму)
            int rownum = 1, colnum = 2;

            if (prepod_list.SelectedIndex > 0)
            {
                tabel_grid.Rows = 2;
                tabel_grid[1, 0] = 
                    prepod_set.Rows[prepod_list.SelectedIndex - 1][2].ToString();

                double chas_summa = 0;

                for (i = 1; i <= finish; i++)
                {
                    DataRow[] prep = rasp.Select(" d = " + i.ToString());

                    double chas = 0.0;

                    if (prep.Length > 0)
                    {
                        foreach (DataRow cr in prep)
                        {
                            chas += (double)cr[2];
                        }
                    }

                    if (chas > 0)
                        tabel_grid[1, colnum] = string.Format("{0:F2}", chas);

                    chas_summa += chas;

                    colnum++;
                }

                tabel_grid[1, 1] = string.Format("{0:F2}", chas_summa);
            }
            else
            {
                tabel_grid.Rows = prepod_set.Rows.Count + 1;

                for (i = 0; i < prepod_set.Rows.Count; i++)
                {
                    tabel_grid[i + 1, 0] = prepod_set.Rows[i][2].ToString();
                }

                foreach (DataRow rr in prepod_set.Rows)
                {
                    string id = rr[0].ToString();
                    colnum = 2;

                    double chas_summa = 0;

                    for (i = 1; i <= finish; i++)
                    {
                        DataRow[] prep = rasp.Select(" d = " + i.ToString() +
                            " and prepod_id = " + id);

                        double chas = 0.0;

                        if (prep.Length > 0)
                        {
                            foreach (DataRow cr in prep)
                            {
                                chas += (double)cr[2];
                            }
                        }

                        if (chas > 0)
                            tabel_grid[rownum, colnum] = string.Format("{0:F2}", chas);

                        chas_summa += chas;

                        colnum++;
                    }

                    tabel_grid[rownum, 1] = string.Format("{0:F2}", chas_summa);
                    rownum++;
                }
            }

        }

        //выбор преподавателя
        private void prepod_list_SelectedIndexChanged(object sender, EventArgs e)
        {
            int index = prepod_list.SelectedIndex;

            if (index == 0)
            {
                prepod_filter = "";
            }
            else
            {
                prepod_filter = " and prepod_id = " +
                    prepod_set.Rows[index - 1][0].ToString() + " ";
            }

            fill_tabel();
        }

        //действия при закрытии окна
        private void tabel_FormClosed(object sender, FormClosedEventArgs e)
        {
            main.tabel_exists = false;
        }


        //копировать в буфер обмена
        private void копироватьВБуферОбменаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int c = tabel_grid.Col;
            int r = tabel_grid.Row;

            if (tabel_grid.Col <= 1 || tabel_grid.Row == 0) return;

            if (tabel_grid[tabel_grid.Row, tabel_grid.Col] == null) return;

            messagebox mb = null;

            if (prepod_list.SelectedIndex == 0)
            {
                mb = new messagebox(current_y, current_m, current_day,
                    (int)prepod_set.Rows[current_row][0],
                    tabel_grid[tabel_grid.Row, 0].ToString());
            }
            else
            {
                mb = new messagebox(current_y, current_m, current_day,
                    (int)prepod_set.Rows[prepod_list.SelectedIndex - 1][0],
                    tabel_grid[1, 0].ToString());
            }

            mb.ShowDialog();

            if (mb.changed)
            {
                fill_tabel();
            }

            mb.Dispose();
        }

        private void tabel_grid_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (tabel_grid.Col <= 1 || tabel_grid.Row == 0) return;

            if (tabel_grid[tabel_grid.Row, tabel_grid.Col] == null) return;

            messagebox mb = null;

            if (prepod_list.SelectedIndex == 0)
            {
                mb = new messagebox(current_y, current_m, current_day,
                    (int)prepod_set.Rows[current_row][0],
                    tabel_grid[tabel_grid.Row, 0].ToString());
            }
            else
            {
                mb = new messagebox(current_y, current_m, current_day,
                    (int)prepod_set.Rows[prepod_list.SelectedIndex - 1][0],
                    tabel_grid[1, 0].ToString());
            }

            mb.ShowDialog();

            if (mb.changed)
            {
                fill_tabel();
            }

            mb.Dispose();
        }

        private void tabel_grid_CellMouseDown(object sender, 
            DataGridViewCellMouseEventArgs e)
        {
            /*current_row = e.RowIndex;
            current_id = (int)prepod_set.Rows[current_row][0];

            if (e.ColumnIndex > 2)
                current_day = e.ColumnIndex - 1;*/
        }

        private void отсавитьТолькоЭтуСтрокуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            prepod_list.SelectedIndex = current_row + 1;
        }

        private void tabel_grid_MouseDown(object sender, MouseEventArgs e)
        {
            tabel_grid.Row = tabel_grid.MouseRow;
            tabel_grid.Col = tabel_grid.MouseCol;

            if (tabel_grid.Col > 1)
                current_day = tabel_grid.Col - 1;

            if (tabel_grid.Row > 0)
                current_row = tabel_grid.Row - 1;
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            if (prepod_list.SelectedIndex > 0)
                prepod_list.SelectedIndex--;
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            if (prepod_list.SelectedIndex < prepod_list.Items.Count - 1)
                prepod_list.SelectedIndex++;
        }

        private void tabel_grid_KeyDown(object sender, KeyEventArgs e)
        {
        }

        private void tabel_grid_KeyUp(object sender, KeyEventArgs e)
        {
        }

        public void CreateFolder(string FolderName)
        {
            string FinalPath = FolderName;

            if (!Directory.Exists(FinalPath))
            {
                Directory.CreateDirectory(FinalPath);
            }
        }

        /// <summary>
        /// вывести табель во внешний файл в формате Excel
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            bool ShowHeader = true; //показыать стандартную шапку
            string S = ""; //строковый буфер
            
            CellRange cr; //диапазон ячеек на рабочем листе книги
            string[] Letters = new string[]{
                "A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P",
                "Q","R","S","T","U","V","W","X","Y","Z",
                "AA","AB","AC","AD","AE","AF","AF","AG","AH","AI","AJ","AK","AL",
                "AM","AN","AO","AP","AQ","AR","AS","AT"};

            tabel_grid.ClipSeparators = "; ";
            tabel_grid.SaveGrid(@"c:\tabel.xls", 
                C1.Win.C1FlexGrid.FileFormatEnum.TextCustom, true, 
                System.Text.Encoding.UTF8);
              
            string root = Environment.GetFolderPath(
                Environment.SpecialFolder.MyDocuments) + "\\Табель";
            //имя файла для сохранения
            string FileName = string.Format("\\табель за {0} {1} года.xls", 
                months_list.Items[months_list.SelectedIndex],DateTime.Now.Year);

            ExcelFile excel = new ExcelFile();
            ExcelWorksheet sheet = excel.Worksheets.Add(string.Format("{0} {1} год", 
                months_list.Items[months_list.SelectedIndex],DateTime.Now.Year));                            

            //задать общие свойства свойства
            cr = sheet.Cells.GetSubrange("a1", "az50");
            cr.Merged = true;
            cr.Style.Font.Name = "Arial Narrow";
            cr.Style.Font.Size = 8 * 20;
            cr.Style.Font.Italic = false;
            cr.Style.VerticalAlignment = VerticalAlignmentStyle.Center;
            cr.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            cr.Style.FillPattern.SetSolid(Color.White);
            cr.Merged = false;


            // вывод шапки табеля 

            /* //сместить вниз под таблицу             
             * sheet.Cells.GetSubrange("A1", "D4").Merged = true;
                            sheet.Cells.GetSubrange(
             * "A1", "D4").Style.HorizontalAlignment = HorizontalAlignmentStyle.Left;

                S = "\"УТВЕРЖДАЮ\"\n" +
                    "Декан ____________ " + dekan_name + "\n" +
                    "\"__\" ________________  " + DateTime.Now.Year.ToString() + " г.";

                sheet.Cells["A1"].Value = S; */
            

            // ------------------   ===== взять шапку раписания из БД
            cr = sheet.Cells.GetSubrange("A1", Letters[days[months_list.SelectedIndex] + 13] + "3");
            cr.Merged = true;
            cr.Style.Font.Name = "Times New Roman";
            cr.Style.Font.Weight = ExcelFont.BoldWeight;
            cr.Style.Font.Italic = false;
            cr.Style.Font.Size = 14 * 20;

            S = string.Format(" Т А Б Е Л Ь\nучета работы преподавателей FSystemа {0}   за   {1}  {2} г.",                
                main.fakultet_name.ToUpper(), months_list.Items[months_list.SelectedIndex],DateTime.Now.Year);

            sheet.Cells["A1"].Value = S;

            /*
            //определить параметры вывдимого расписания

            //перечень вывдомых строк
            OutLines = empty_rows;

            int num = 1;
            int i = 0;
            for (i = 0; i < grupa_list.Items.Count; i++)
            {
                OutGroups.Add(num);
                num += 2;
            }


            //поставить решетку             

            cr = sheet.Cells.GetSubrange("A6", "C" + (OutLines[OutLines.Count - 1] + 5).ToString());
            cr.Merged = true;
            cr.Style.Borders.SetBorders(MultipleBorders.Outside, Color.Black,
                 LineStyle.Thin);
            cr.Merged = false;

            //sheet.Cells.GetSubrange("A6", "C6").SetBorders(MultipleBorders.Top, Color.Black, LineStyle.DoubleLine);
            cr = sheet.Cells.GetSubrange("C6", "C" + (OutLines[OutLines.Count - 1] + 5).ToString());
            cr.Merged = true;
            cr.Style.Borders[IndividualBorder.Right].LineStyle = LineStyle.Double;
            cr.Merged = false;

            cr = sheet.Cells.GetSubrange("A6", "A" + (OutLines[OutLines.Count - 1] + 5).ToString());
            cr.Merged = true;
            cr.Style.Borders[IndividualBorder.Left].LineStyle = LineStyle.Double;
            cr.Merged = false;

            cr = sheet.Cells.GetSubrange("A" + (OutLines[OutLines.Count - 1] + 5).ToString(),
                "C" + (OutLines[OutLines.Count - 1] + 5).ToString());
            cr.Merged = true;
            cr.Style.Borders[IndividualBorder.Bottom].LineStyle = LineStyle.Double;
            cr.Merged = false;

            cr = sheet.Cells.GetSubrange("A6", "C6");
            cr.Merged = true;
            cr.Style.Borders[IndividualBorder.Top].LineStyle = LineStyle.Double;
            cr.Merged = false;

            cr = sheet.Cells.GetSubrange("A6", "C6");
            cr.Merged = true;
            cr.Style.Borders[IndividualBorder.Bottom].LineStyle = LineStyle.Double;
            cr.Merged = false;

            //первый столбец - дни недели
            sheet.Cells["A6"].Value = "День\nнедели";
            sheet.Cells["B6"].Value = "№";
            sheet.Cells["C6"].Value = "Время";
            sheet.Rows[5].Height = 25 * 20;


            num = 6;
            string cell1 = "", cell2 = "";
            foreach (int j in OutLines)  //вывод дней недели, пар и дат
            {
                cell1 = string.Format("A{0}", num + 1);

                sheet.Cells[cell1].Value = table[j, 1].ToString().Substring(0, 5);

                cell1 = string.Format("A{0}", num + 2);
                cell2 = string.Format("A{0}", num + 6);
                cr = sheet.Cells.GetSubrange(cell1, cell2);
                cr.Merged = true;

                cr.Style.Rotation = 90;
                sheet.Cells[cell1].Value = table[j, 0].ToString();
                cr.Style.Font.Weight = ExcelFont.MaxWeight;
                cr.Style.Font.Size = 12 * 20;
                sheet.Columns[0].Width = 6 * 256;


                for (int k = 1; k <= 6; k++)
                {
                    sheet.Cells[num + k - 1, 1].Value = k.ToString();
                    sheet.Columns[1].Width = 3 * 256;
                    sheet.Cells[num + k - 1, 2].Value = table[j + k, 0].ToString();
                    sheet.Columns[2].Width = 11 * 256;
                    sheet.Rows[num + k - 1].Height = 25 * 20;
                }

                num += 6;
            }

            int cols = table.Cols - 1;
            int ColWidth = 10 * 256;
            int ColWidthWide = 20 * 256;
            int counter = 3;

            for (i = 1; i < grups_count * 2; i += 2) ///исправить цикл взять количество из Outgroups
            {
                sheet.Columns[i + 2].Width = ColWidth;
                sheet.Columns[i + 3].Width = ColWidth;

                //название группы
                cell1 = Letters[counter] + "6";
                cell2 = Letters[counter + 1] + "6";

                cr = sheet.Cells.GetSubrange(cell1, cell2);
                cr.Merged = true;
                sheet.Cells[cell1].Value = table[0, i].ToString();
                cr.Style.Font.Weight = ExcelFont.MaxWeight;
                cr.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
                cr.Style.VerticalAlignment = VerticalAlignmentStyle.Center;
                cr.Style.Font.Italic = true;
                cr.Style.Font.Name = "Times New Roman";
                cr.Style.Font.Size = 10 * 20;
                cr.Style.Borders.SetBorders(MultipleBorders.Outside, Color.Black, LineStyle.DoubleLine);

                //цикл вывода предметов
                num = 7;

                foreach (int j in OutLines)
                {
                    int ii = 0;
                    for (ii = 1; ii <= 6; ii++)
                    {
                        string text1 = get_cell_shorttext(j + ii, i);
                        string text2 = get_cell_shorttext(j + ii, i + 1);

                        cell1 = Letters[counter] + (num + ii - 1).ToString();
                        cell2 = Letters[counter + 1] + (num + ii - 1).ToString();


                        cr = sheet.Cells.GetSubrange(cell1, cell1);
                        cr.Merged = true;
                        cr.Style.Borders.SetBorders(MultipleBorders.Outside,
                                Color.Black, LineStyle.Thin);
                        cr.Merged = false;

                        cr = sheet.Cells.GetSubrange(cell2, cell2);
                        cr.Merged = true;
                        cr.Style.Borders.SetBorders(MultipleBorders.Outside,
                                Color.Black, LineStyle.Thin);
                        cr.Merged = false;

                        if (text1.ToLower().Trim() == text2.ToLower().Trim())
                        {
                            //нет деления - соединить
                            cr = sheet.Cells.GetSubrange(cell1, cell2);
                            cr.Merged = true;
                            cr.Value = text1;
                        }
                        else
                        {
                            cr = null;
                            sheet.Columns[Letters[counter]].Width = ColWidthWide;
                            sheet.Columns[Letters[counter + 1]].Width = ColWidthWide;
                            //есть деление - не соединять
                            sheet.Cells[cell1].Value = text1;
                            sheet.Cells[cell2].Value = text2;
                        }
                        sheet.Cells[cell2].SetBorders(MultipleBorders.Right,
                            Color.Black, LineStyle.DoubleLine);

                    }

                    //вывести горизонтальный разделитель
                    sheet.Cells[cell1].SetBorders(MultipleBorders.Bottom,
                        Color.Black, LineStyle.DoubleLine);
                    sheet.Cells[cell2].SetBorders(MultipleBorders.Bottom,
                        Color.Black, LineStyle.DoubleLine);


                    sheet.Cells["A" + (num + ii - 2).ToString()].Style.Borders[IndividualBorder.Bottom].LineStyle =
                        LineStyle.DoubleLine;

                    cr = sheet.Cells.GetSubrange("B" + (num + ii - 2).ToString(), "B" + (num + ii - 2).ToString());
                    cr.Merged = true;
                    cr.Style.Borders[IndividualBorder.Bottom].LineStyle = LineStyle.DoubleLine;
                    cr.Merged = false;

                    cr = sheet.Cells.GetSubrange("C" + (num + ii - 2).ToString(), "C" + (num + ii - 2).ToString());
                    cr.Merged = true;
                    cr.Style.Borders[IndividualBorder.Bottom].LineStyle = LineStyle.DoubleLine;
                    cr.Merged = false;


                    num += 6;
                }

                counter += 2;
            }


            sheet.PrintOptions.HeaderMargin = 0.0;
            sheet.PrintOptions.FooterMargin = 0.0;
            sheet.PrintOptions.Portrait = false;


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
            }*/

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
                        "Попробуйте открыть файл самостоятельно  " + 
                        " или повторите опеерацию конвертирования.",
                        "Ошибка создания",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
    }
}