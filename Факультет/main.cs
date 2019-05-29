using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Threading;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using System.IO;
using System.Globalization;
using GemBox.Spreadsheet;
using System.Diagnostics;
using DevExpress.XtraCharts;
using Word = Microsoft.Office.Interop.Word;

namespace FSystem
{
    public partial class main : Form
    {
        /// <summary>
        /// конструктор класса
        /// </summary>
        public main ( )
        {
            InitializeComponent ( );

            for ( int i = 1; i <=43; i += 7 )
                empty_rows.Add ( i );

            left_panel.Visible = false;
            info_panel.Visible = false;
            content.Visible = false;
            first_enter = true;
            fakultet_node = object_tree.Nodes["fakultet_node"];
            sprav_node = object_tree.Nodes["sprav_node"];            
        }

        public static bool final = true;

        //поля общего назначения
        static Mutex mutex;  //глобальный мьютекс
        private int stopweeks = 50; //количество недель, по прошествии которых нельзя редактировать расписание
        private string weekword = " недель";
        private bool movecell = true;
        TreeNode fakultet_node=null, sprav_node=null;
        public string status_text_up = "";
        Random rnd = new Random();
        
        public Type ExcelType = null;

        //поля для работы с базой данных
        public static SqlConnection global_connection;
        public static SqlCommand global_command;
        public static SqlDataAdapter global_adapter;
        public string global_query = "";
        public static string con_string, srv_name, srv_pwd;
        public bool success = false;
        public static DateTime server_date;

        /// <summary>
        /// массив для хранения требуемого количества часов
        /// </summary>
        public List<double> nado = new List<double>();
        /// <summary>
        /// массив для хранения количества выданного часов
        /// </summary>
        public List<double> fakt = new List<double>();


        /// <summary>
        /// окно ожидания
        /// </summary>
        wait w;

        //данные из БД "vkr"
        /// <summary>
        /// список групп FSystemа
        /// </summary>
        public DataTable grups_set;

        /// <summary>
        /// список преподавателей
        /// </summary>
        public DataTable prepod_set; 
        /// <summary>
        /// список предметов
        /// </summary>
        public DataTable predmet_set; 
        /// <summary>
        /// список видов занятий
        /// </summary>
        public DataTable vidzan_set; 
        /// <summary>
        /// список аудиторий
        /// </summary>
        public DataTable aud_set;
        public DataSet prepod_predmet;
        public DataTable statistica_set;

        //поля для работы с таблицей
        Data table_data;
        List<string> groups = new List<string>();

        // ------------  данные даты и времени --------------------------
        public static string[] DaysLong = new string[]{"Понедельник","Вторник",
				"Среда","Четверг","Пятница","Суббота","Воскресенье"};
        public static string[] DaysMed = new string[]{"Понед.","Вторник",
				"Среда","Четверг","Пятница","Суббота","Воскр."};
        public string[] DaysShort = new string[]{"пнд","втр",
				"срд","чтв","птн","сбт","вскр"};
        public string[] months = new string[]{
            "январь","февраль","март", "апрель","май","июнь","июль","август", "сентябрь","октябрь",
            "ноябрь","декабрь"};
        public static List<DateTime> starts, ends;
        public int semestr = 0;


        /// <summary>
        /// контекстное меню
        /// </summary>
        //table_context tc = new table_context();


        //значения, устнавливаемые диалоговыми окнами -----
        public static bool first_enter = true; //доступ к программе осущетв. впервые
        public bool use_week_list_change = false;
        //для функций слияния или разъединения ячеек
        public static string left = "", right = "", goaltext = "";

        // -----  данные  сетки ----
        public static string fakultet_name = "",
            fakultet_name_krat = "", active_user_name = "", active_user_kaf = "", active_user_dolz = "";
        public static int fakultet_id = 0, active_user_id = 0, active_user_dolz_id = 0;
        public static string fakultet_prfix = "";
        public static int peremena = 10, long_peremena = 20;
        public static int first_long_peremena = 2, second_long_peremena = 4;
        public static DateTime semestr2_start = new DateTime(2000, 1, 1, 9, 0, 0, 0);
        public static DateTime year_start = new DateTime(2000, 1, 1, 9, 0, 0, 0);
        public static DateTime year_end = new DateTime(2000, 1, 1, 9, 0, 0, 0);
        public static DateTime semestr1_end = new DateTime(2000, 1, 1, 9, 0, 0, 0);
        public static DateTime start_time = new DateTime ( );
        public static bool dekan_online = false;
        public static string dekan_name = "";
        public static int uch_god = 1;
        public static DateTime Att1_1 = new DateTime();
        public static DateTime Att1_2 = new DateTime();
        public static DateTime Att2_1 = new DateTime();
        public static DateTime Att2_2 = new DateTime();
        public static string user_role = String.Empty;
        public static string df = "";
        List<int> empty_rows = new List<int> ( );
        List<string> probel = new List<string> ( );
        public int grups_count = 11;

        //служебные
        bool from_change_user = false;
        bool predm_updated = false;
        
        /// <summary>
        /// ссылка на процесс приложения MS Ofiice Word
        /// </summary>
        Word.Application wa;

        /// <summary>
        /// проверка повтороного запуска программы
        /// </summary>
        /// <returns>true, если программа уже запущена</returns>
        static bool AlreadyRunning ( )
        {
            bool canOwn;
            mutex = new Mutex ( false, "tableISANIE", out canOwn );
            return !canOwn;
        }

        /// <summary>
        /// действия при загрузке формы в память
        /// </summary>
        private void main_Load ( object sender, EventArgs e )
        {            
            //проверить, работает ли уже данная программа
            if ( AlreadyRunning ( ) )
            {
                //если да, то выйти
                MessageBox.Show("\nПрограмма уже запущена.\nПерейдите в запущенный ранее экземпляр программы.",
                    "Повторный запуск запрещен",
                    MessageBoxButtons.OK, MessageBoxIcon.Stop);
                //msg.ShowDialog ( );
                //msg.Dispose ( );
                this.Dispose ( );
                Application.Exit ( );
                return;
            }

            //отобразить окно диалога для входа на сервер данных
            enter_server es = new enter_server ( );
            
            do
            {
                DialogResult es_res = es.ShowDialog ( );
                if ( es_res == DialogResult.Cancel )
                {
                    es.Dispose ( );
                    this.Dispose ( );
                    Application.Exit ( );
                    return;
                }
            }
            while(!es.enter_server_result);

            srv_name = es.srv.Text.Trim ( );
            srv_pwd = es.pwd.Text.Trim ( );

            es.Dispose ( );
            

            //установить соединение с сервером БД

            //показать окно ожидания            
            w = new wait ( );
            w.Show ( );

            //выполнить вход на сервер
            //передать управление программе для работы с БД
            con_string = "timeout = 30; Data Source=" + srv_name +
                "; Initial Catalog=VKR; " +
                "User ID=sa;" +
                "; Password=" + srv_pwd;
                //"; Network Library=dbnmpntw; ";

            global_connection = new SqlConnection ( con_string );

            Application.DoEvents ( );

            try
            {
                global_connection.Open ( );
            }
            catch ( Exception ex1 )
            {
                success = false;
            }


            if ( global_connection.State == ConnectionState.Open )
                success = true;
            else
                success = false;
        
            w.Dispose ( );
            
            if ( !success )
            {
                MessageBox.Show(
                    "К сожалению, в данный момент вход на сервер не возможен.\n"+
                    "Введены неверные данные для входа или сервер недоступен.\n"+
                    "Повторите попытку входа позднее.",
                    "Ошибка регистрации", MessageBoxButtons.OK, MessageBoxIcon.Error);               
                this.Dispose();
                Application.Exit();
                return;
            }

            //ввод учетных данных
            enter_fakultet_and_user efau = new enter_fakultet_and_user ( );

            try
            {
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(new CultureInfo("ru-RU"));
            }
            catch(Exception exx)
            {
                //
            }

            int i = 1;
            do
            {        
                DialogResult dres = efau.ShowDialog ( );
                if ( dres == DialogResult.Cancel || i == 3 )
                {
                    efau.Dispose ( );
                    this.Dispose ( );
                    Application.Exit ( );
                    return;
                }
                
                i++;
            }
            while(!efau.res);

            efau.Dispose ( );


            if (dekan_online)
            {
                table.ContextMenuStrip = contextMenuStrip1;                
            }
            else
            {
                table.ContextMenuStrip = teacher_menu;               
            }

            starts = new List<DateTime> ( );
            ends = new List<DateTime> ( );

            fill_tree ( );
            fill_grupa_list ( );
            fill_date_list ( );
            table.AllowDrop = true;
            init_table ( );

            // ------------------------------------------------------------------------

            left_panel.Visible = true;
            info_panel.Visible = true;
            content.Visible = true;

            // ----- получить дату и время сервера ----
            global_command = new SqlCommand("select getdate()", global_connection);
            server_date = (DateTime)global_command.ExecuteScalar();

            table.ContextMenuStrip = null;
                        
            if (stopweeks == 1) weekword = " недели";            

            if (dekan_online)
            {
                //table.ContextMenuStrip = contextMenuStrip1;
                statistica.Visible = true;
                //prepod_details.Visible = true;
                //prepod_details.Visible = true;                
            }
            else
            {
                //table.ContextMenuStrip = teacher_menu;
                statistica.Visible = false;
                //prepod_details.Visible = false;
                //prepod_details.Visible = false;
                content.SelectedIndex = 1;
                teacher_tab.SelectedIndex = 0;
                
            }            

            //работа с вкладками
            //content.TabPages.Remove(prepod_tab);
            teacher_tab.TabPages.Remove(teacher_predmet);
            teacher_tab.TabPages.Remove(teacher_tab_zurnal);
            //teacher_tab.TabPages.Remove(tecaher_tab_raspisanie);
            //teacher_tab.TabPages.Remove(teacher_tab_kontrlist);
            teacher_tab.TabPages.Remove(teacher_tab_grafik);
            teacher_tab.TabPages.Remove(teacher_tab_poruch);
            teacher_tab.TabPages.Remove(teacher_tab_vkr);
            teacher_tab.TabPages.Remove(teacher_tab_dipl);
            teacher_tab_predmet.TabPages.Remove(teacher_tab_predmet_MRS);

            if (active_user_dolz_id == 15)
            {
                content.TabPages.Remove(prepod_tab);
            }
            else
            {
                if (!dekan_online)
                {
                    content.TabPages.Remove(tabПосещаемость);
                    content.TabPages.Remove(tabPageAttest);
                }
            }

            begin.MaxDate = end.Value;
            end.MinDate = begin.Value;

            begin.Value = end.Value.AddDays(-1);

        }

        /// <summary>
        /// номер дня недели по дате
        /// </summary>
        /// <param name="x">дата для определения номера дня</param>
        /// <returns></returns>
        public static int daynumer ( DateTime x )
        {
            int day = 0;
            switch ( x.DayOfWeek )
            {
                case DayOfWeek.Sunday: day = 7; break;
                case DayOfWeek.Monday: day = 1; break;
                case DayOfWeek.Tuesday: day = 2; break;
                case DayOfWeek.Thursday: day = 4; break;
                case DayOfWeek.Wednesday: day = 3; break;
                case DayOfWeek.Friday: day = 5; break;
                case DayOfWeek.Saturday: day = 6; break;
            }

            return day;
        }

        /// <summary>
        /// заполнить список дат расписания
        /// по началу и концу учебной недели
        /// </summary>
        public void fill_date_list ( )
        {
            DateTime actual, start, end, current;

            starts.Clear ( );
            ends.Clear ( );
            
            actual = DateTime.Today;

            int m = actual.Month;
            int y = actual.Year;

            if ( m >= 1 && m <= 6 )
                start = new DateTime ( y - 1, 9, 1 );
            else
                start = new DateTime ( y, 9, 1 );

            if ( m <= 12 && m > 6 )
                end = new DateTime ( y + 1, 6, 20 );
            else
                end = new DateTime ( y, 6, 20 );

            TimeSpan ts = new TimeSpan ( 1 , 0, 0, 0, 0 );
            TimeSpan ts7 = new TimeSpan ( 7, 0, 0, 0, 0 );


            do
            {
                if ( start.DayOfWeek != DayOfWeek.Monday )
                    start = start.Subtract ( ts );
                else
                    break;
            }
            while(true);

            do
            {
                if ( end.DayOfWeek != DayOfWeek.Sunday )
                    end = end.AddDays ( 1.0 );
                else
                    break;
            }
            while ( true );

            for ( current = start; current < end; current = current.AddDays ( 7.0 ) )
            {
                starts.Add ( current );
                ends.Add ( current.AddDays ( 6.0 ) );

                string dateitem = "неделя c " + current.ToShortDateString ( ) + " по " +
                    current.AddDays ( 6 ).ToShortDateString ( );
                week_list.Items.Add ( dateitem );
            }

            week_list.SelectedIndex = get_current_date_listnumber ( );

        }

        /// <summary>
        /// получить номер текущей недели в списке выбора недель
        /// </summary>
        /// <returns></returns>
        public int get_current_date_listnumber ( )
        {
            DateTime now = DateTime.Now.Date;
            for ( int i = 0; i < starts.Count; i++ )
            {
                if ( now >= starts[i] && now <= ends[i] )
                {
                    return i;
                }
            }

            return 0;
        }

        /// <summary>
        /// заполнить таблицу раписания на текущую 
        /// </summary>
        public void fill_table ( )
        {
            if (table_data!=null) table_data.data.Clear();

            table_data = new Data(starts[week_list.SelectedIndex],
                empty_rows, groups, fakultet_id);

            int m = starts[week_list.SelectedIndex].Month;
            int d = starts[week_list.SelectedIndex].Day;

            if (m >= 9 && m <= 12)
                semestr = 1;
            else
            {
                if ((m == 1 && d >= 30) || (m > 1 && m <= 6))
                    semestr = 2;
                else
                    semestr = 1;
            }

            //заполнить списки преподов, предметов и видов занятий
            fill_prepods();
            fill_predemts();
            fill_vidzan();
            fill_auds();

            // поместить данные раписания в таблицу ------------------------

            //удалить текущие значения из таблицы
            for (int i = 2; i < table.Rows; i++)
            {
                if (empty_rows.Contains(i)) continue;
                for (int j = 1, k = 0; j < table.Cols; j += 2, k++)
                {
                    table.Cell[0, i, j] = probel[j];
                    table.Cell[0, i, j + 1] = probel[j];
                    table.set_MergeRow(i, true);
                }
                if (i > 0) table.set_RowHeight(i, 35);
            }

            int days = 0;
            for (int j = 1; j < table.Rows; j += 7)
            {

                for (int i = 1; i < table.Cols; i++)
                {
                    table.Select(j, i, false);
                    table.CellBackColor = System.Drawing.Color.LightSteelBlue;
                    table.Cell[0, j, i] = starts[week_list.SelectedIndex].AddDays(days).ToShortDateString();
                }
                days++;
            }

            Color coler1 = new Color();
            Color coler2 = new Color();

            foreach (DateTime dt in table_data.data.Keys)
            {
                foreach (string gr in table_data.data[dt].Keys)
                {
                    for (int para = 1; para <= 6; para++)
                    {
                        int r = table_data[dt, gr, para].row;
                        int c = table_data[dt, gr, para].col[0];

                        bool res1 = active_user_id == table_data[dt, gr, para].prepod_id[0];
                        bool res2 = active_user_id == table_data[dt, gr, para].prepod_id[1];

                        if (res1)
                            coler1 = Color.Red;
                        else
                            coler1 = Color.Black;

                        if (res2)
                            coler2 = Color.Red;
                        else
                            coler2 = Color.Black;

                             
                        //задать текст и цвет для ячейки
                        if (table_data[dt, gr, para].use_two_cells())
                        {
                            set_one_value(get_cell_text(dt, gr, para, 0), r, c, true);
                            set_one_style(r, c, true, C1.Win.C1FlexGrid.Classic.CellPropertySettings.flexcpForeColor,
                                coler1);                            
                        }
                        else
                        {
                            if ((int)table_data[dt, gr, para].subgr_nomer[0] != 0)
                            {
                                table.Cell[0, r, c] = get_cell_text(dt, gr, para, 0);
                                table.Cell[C1.Win.C1FlexGrid.Classic.CellPropertySettings.flexcpForeColor,
                                    r, c] = coler1;
                            }

                            if ((int)table_data[dt, gr, para].subgr_nomer[1] != 0)
                            {
                                table.Cell[0, r, c + 1] = get_cell_text(dt, gr, para, 1);
                                table.Cell[C1.Win.C1FlexGrid.Classic.CellPropertySettings.flexcpForeColor,
                                    r, c + 1] = coler2;
                            }
                        }
                    }
                }
            }

            for (int k = 1; k < table.ColumnCollection.Count-1; k++)
            {
                try
                {
                    table.ColumnCollection[k].Style.TextAlign = C1.Win.C1FlexGrid.TextAlignEnum.CenterCenter;
                }
                catch(Exception exx)
                {
                    ;
                }
            }
            
            table.Row = 1;
        }

        
        /// <summary>
        /// получить текст для отображения в ячейке
        /// </summary>
        /// <param name="d">дата</param>
        /// <param name="g">название группы</param>
        /// <param name="p">номер пары</param>
        /// <param name="sub">номер подгруппы</param>
        /// <returns>текст для отображения в ячейке расписания</returns>
        public string get_cell_text(DateTime d, string g, int p, int sub)
        {
            if (table_data[d, g, p].prepod_id[sub] == 0) return string.Empty;
            string predm = table_data[d, g, p].predmet_name[sub];
            string prep = "/" + table_data[d, g, p].prepod_name[sub] + "/";
            string aud = (table_data[d, g, p].aud_name[sub]=="--")?"":("," + table_data[d, g, p].aud_name[sub]);
            string vidzan = table_data[d, g, p].vid_zan_name[sub];

            string outstr = predm + ", " + vidzan + "\n" + prep + aud;

            return outstr;
                
        }


        //получить без вида занятия (убрать лекцию, практ, семинар, лаб)
        public string get_cell_shorttext(int r, int c)
        {

            //определить колонки подгрупп
            int first = 0, second = 0;
            if (c % 2 == 0)
            {
                first = c - 1;
                second = c;
            }
            else
            {
                first = c;
                second = c + 1;
            }

            int dest_col = first; //колонка назначения
            if (c == second) dest_col = second;
            if (c == first) dest_col = first;

            bool cell_divided = false;  //разделена ли исходная ячейка
            if (table.Cell[0, r, first].ToString() !=
                table.Cell[0, r, second].ToString())
                cell_divided = true;

            //определить параметры исходной ячейки
            DateTime d = table_data.RowDate(r);
            int p = table_data.RowPair(r);
            string g = table_data.ColumnGroup(dest_col);
            int sub = table_data.ColumnSubGroup(dest_col);

            bool empty = table_data.data[d][g][p].prepod_id[sub - 1] == 0;

            if (table_data[d, g, p].prepod_id[sub - 1] == 0) return string.Empty;

            string predm = table_data[d, g, p].predmet_name[sub-1];
            string prep = "/" + table_data[d, g, p].prepod_name[sub-1] + "/";
            string aud = (table_data[d, g, p].aud_name[sub-1] == "--") ? "" : ("," + table_data[d, g, p].aud_name[sub-1]);
            
            string vidzan = table_data[d, g, p].vid_zan_name[sub-1];
            string vidzanfull = table_data[d, g, p].vid_full_name[sub - 1];

            string outstr = "";

            if (vidzanfull.ToLower().Contains("лекция")||
                vidzanfull.ToLower().Contains("семинар") ||
                vidzanfull.ToLower().Contains("практич") ||
                vidzanfull.ToLower().Contains("лаборатор"))                           
                outstr = predm + "\n" + prep + aud;
            else
                outstr = predm + ", " + vidzan + "\n" + prep + aud;

            return outstr;

        }

        //получить без вида занятия (убрать лекцию, практ, семинар, лаб)
        //без разбиения на строки
        public string get_cell_shorttext__one_line(int r, int c, bool fullpredmetname)
        {

            //определить колонки подгрупп
            int first = 0, second = 0;
            if (c % 2 == 0)
            {
                first = c - 1;
                second = c;
            }
            else
            {
                first = c;
                second = c + 1;
            }

            int dest_col = first; //колонка назначения
            if (c == second) dest_col = second;
            if (c == first) dest_col = first;

            bool cell_divided = false;  //разделена ли исходная ячейка
            if (table.Cell[0, r, first].ToString() !=
                table.Cell[0, r, second].ToString())
                cell_divided = true;

            //определить параметры исходной ячейки
            DateTime d = table_data.RowDate(r);
            int p = table_data.RowPair(r);
            string g = table_data.ColumnGroup(dest_col);
            int sub = table_data.ColumnSubGroup(dest_col);

            bool empty = table_data.data[d][g][p].prepod_id[sub - 1] == 0;

            if (table_data[d, g, p].prepod_id[sub - 1] == 0) return string.Empty;

            string predm = table_data[d, g, p].predmet_name[sub - 1];
            if (fullpredmetname)
                predm = table_data[d, g, p].predmet_fullname[sub - 1];

            string prep = "/" + table_data[d, g, p].prepod_name[sub - 1] + "/";
            string aud = (table_data[d, g, p].aud_name[sub - 1] == "--") ? "" : (", " + table_data[d, g, p].aud_name[sub - 1]);

            string vidzan = table_data[d, g, p].vid_zan_name[sub - 1];
            string vidzanfull = table_data[d, g, p].vid_full_name[sub - 1];

            string outstr = "";

            if (vidzanfull.ToLower().Contains("лекция") ||
                vidzanfull.ToLower().Contains("семинар") ||
                vidzanfull.ToLower().Contains("практич") ||
                vidzanfull.ToLower().Contains("лаборатор"))
                outstr = predm + " " + prep + aud;
            else
                outstr = predm + ", " + vidzan + " " + prep + aud;

            return outstr;

        }



        /// <summary>
        /// заполнить дерево объектов
        /// </summary>
        public void fill_tree ( )
        {
            //работа с деревом объектов -------------------------------------

            Text = String.Format ( "FSystem \"{0}\" для пользователя {1} [статус: {2}]",
                fakultet_name, active_user_name, user_role );

            object_tree.Nodes[0].Text = active_user_name;           

            if ( !dekan_online )
            {
                object_tree.Nodes[0].ExpandAll ( );
                set_controls_status ( false );

            }
            else
            {
                if (object_tree.Nodes.Count == 1)
                {
                    object_tree.Nodes.Add(fakultet_node);
                    object_tree.Nodes.Add(sprav_node);
                }

                object_tree.Nodes[1].Text = "FSystem " + fakultet_name_krat;
                object_tree.Nodes[1].ExpandAll ( );
                object_tree.Nodes[2].ExpandAll ( );
            }

            //заполнить предметы данного преподавателя
            fill_tree_prepod_predmet ( );
        }

        /// <summary>
        /// установить доступность элементов управления
        /// </summary>
        /// <param name="status">доступность компонентов</param>
        public void set_controls_status ( bool status )
        {
            grupa_list.Enabled = status;
            predmet_list.Enabled = status;
            prepod_list.Enabled = status;
            vid_zan_list.Enabled = status;
            aud_list.Enabled = status;

            if (!status)
            {
                if (object_tree.Nodes.Count > 1)
                {
                    object_tree.Nodes.Remove(fakultet_node);
                    object_tree.Nodes.Remove(sprav_node);
                }
            }
            else
            {
                if (object_tree.Nodes.Count == 1)
                {
                    object_tree.Nodes.Add(fakultet_node);
                    object_tree.Nodes.Add(sprav_node);
                }
            }
            
        }


        /// <summary>
        /// получить список предметов преподавателя 
        /// и сохранить в prepod_predmet
        /// </summary>
        void fill_tree_prepod_predmet ( )
        {
            object_tree.Nodes[0].Text = active_user_name;
            object_tree.Nodes[0].Nodes[3].Nodes.Clear ( );

            int y = 0;

            if (DateTime.Today.Month >= 1 && DateTime.Today.Month <= 6)
                y = DateTime.Today.Year - 1;
            else
                y = DateTime.Today.Year;

            string query = string.Format("select " + 
	                "predmet.name, predmet.name_krat, grupa.name, predmet.id,  " +
	                "predmet_type.name, predmet.semestr " +
                "from predmet " +
	                "join prepod on prepod.id = predmet.prepod_id   " +
                 	"join predmet_type on predmet.type_id = predmet_type.id   " +
                 	"join grupa on grupa.id = predmet.grupa_id   " +
                 	"join fakultet on fakultet.id = grupa.fakultet_id   " +
                	" where grupa.show_in_grid = 1 and prepod.id = {0} and predmet.actual = 1   " +
	                " and predmet_type.kod not like '%срез%' and predmet_type.kod not like '%подготовка%' " +
                "order by fakultet.id, grupa.name", active_user_id);
                                
                /*string.Format(" select predmet.name, predmet.name_krat, grupa.name, predmet.id, " +
                " predmet_type.name, predmet.semestr  " +
                " from poruch " +
                " join prepod on prepod.id = poruch.prepod_id " +
                " join predmet on predmet.id = poruch.predmet_id " +
                " join predmet_type on predmet.type_id = predmet_type.id " +
                " join grupa on grupa.id = predmet.grupa_id " +
                " join fakultet on fakultet.id = grupa.fakultet_id " +
                " where grupa.show_in_grid = 1 and poruch.prepod_id = {0} and y = {1} and predmet.actual = 1 " +
                " order by fakultet.id, grupa.name ", active_user_id, y);*/


            prepod_predmet = new DataSet();
            global_adapter = new SqlDataAdapter(query, global_connection);
            global_adapter.Fill(prepod_predmet);

            int i = 0;
            foreach (DataRow row in prepod_predmet.Tables[0].Rows)
            {
                string node = row[1].ToString() + " [" + row[2].ToString() + ", " + row[5].ToString() + " сем.]";
                int id = (int)row[3];

                object_tree.Nodes[0].Nodes[3].Nodes.Add(node);
                
                object_tree.Nodes[0].Nodes[3].Nodes[i].ForeColor = Color.Red;
                object_tree.Nodes[0].Nodes[3].Nodes[i].Tag = row[3];
                object_tree.Nodes[0].Nodes[3].Nodes[i].Name = "pr_" + i.ToString();
                
                i++;
            }

            object_tree.Nodes[0].Nodes[3].Collapse ( true );
        }


        private void main_Shown ( object sender, EventArgs e )
        {
            //
        }

        private void выйтиИзПрограммыToolStripMenuItem_Click ( object sender, EventArgs e )
        {
            global_connection.Dispose ( );
            Application.Exit ( );
        }

        /// <summary>
        /// обработка разворачивания узлов дерева объектов
        /// </summary>        
        private void object_tree_AfterExpand ( object sender, TreeViewEventArgs e )
        {
            if (dekan_online)
            {
                set_controls_status(true);
            }
            
            object_tree_AfterSelect(sender, e);

        }

        /// <summary>
        /// задать FSystem и пользователя
        /// </summary>
        bool fakultet = true;
        private void set_fakultet ( )
        {
            enter_fakultet_and_user sf = new enter_fakultet_and_user ( );
         
            if ( from_change_user ) sf.txt = " Смена пользователя";
            if (sf.ShowDialog ( ) == DialogResult.Cancel || !sf.res)
                return;

            sf.Dispose ( );

            if (dekan_online)
            {
                table.ContextMenuStrip = contextMenuStrip1;
                statistica.Visible = true;
                //prepod_details.Visible = true;   
                if (!content.TabPages.Contains(tabПосещаемость))
                    content.TabPages.Add(tabПосещаемость);

                if (!content.TabPages.Contains(tabPageAttest))
                    content.TabPages.Add(tabPageAttest);

                if (active_user_dolz_id == 15)
                {
                    if (content.TabPages.Contains(prepod_tab))
                        content.TabPages.Remove(prepod_tab);
                }
            }
            else
            {
                table.ContextMenuStrip = teacher_menu;
                statistica.Visible = false;
                //prepod_details.Visible = false;
                if (!content.TabPages.Contains(prepod_tab))
                    content.TabPages.Add(prepod_tab);

                if (content.TabPages.Contains(tabПосещаемость))
                    content.TabPages.Remove(tabПосещаемость);

                if (content.TabPages.Contains(tabPageAttest))
                    content.TabPages.Remove(tabPageAttest);

            }

            fill_tree ( );
            fill_grupa_list ( );
            grupa_list.SelectedIndex = 0;
            init_table();
        }


        /// <summary>
        /// заполнить список предподавателей для текушей группы
        /// </summary>
        public void fill_prepods()
        {
            int kurs_id = 0;
            int item = (grupa_list.SelectedIndex>=0)?grupa_list.SelectedIndex:0;

            kurs_id = (int)grups_set.Rows[item][1];

            int x = (semestr % 2 == 0) ? 0 : 1;
            int number = kurs_id * (semestr + x) - x;

            global_query = string.Format("select distinct prepod.id, fio = fam + ' ' + im + ' ' + ot, sex, " +
                        " fiokr = fam + ' ' + left(im,1) + '.' + left(ot,1) + '.', " +                         
                        " fam, im, ot, " + 
                        " dolznost_id, stepen_id, zvanie_id, prepod.kafedra_id, prepod.actual, " + 
                        " address, phone " +
                        " from prepod " + 
                        " join predmet on predmet.prepod_id=prepod.id " +
                        " where prepod.actual=1 and predmet.actual=1 " +
                        " and predmet.grupa_id = {0}" +
                        " and predmet.semestr = {1}" +
                        " order by fio, prepod.id ", grups_set.Rows[item][2],
                      number);  
            
            
            prepod_list.Items.Clear();

            global_adapter = new SqlDataAdapter(global_query,global_connection);
            prepod_set = new DataTable();
            global_adapter.Fill(prepod_set);

            //MessageBox.Show(prepod_set.Rows.Count.ToString());
            for (int i = 0; i < prepod_set.Rows.Count; i++)
            {
                prepod_list.Items.Add(
                    prepod_set.Rows[i][1].ToString());
            }

            if (prepod_list.Items.Count > 0) prepod_list.SelectedIndex = 0;

            predm_updated = true;
        }

        /// <summary>
        /// заполнить список предметов
        /// </summary>
        public void fill_predemts()
        {

            if (prepod_list.Items.Count == 0)
            {
                predmet_list.Items.Clear();
                return;
            }

            int kurs_id = 0;
            int grup_item = (grupa_list.SelectedIndex >= 0) ? grupa_list.SelectedIndex : 0;
            int prepod_item = (prepod_list.SelectedIndex >= 0) ? prepod_list.SelectedIndex : 0;

            kurs_id = (int)grups_set.Rows[grup_item][1];

            int x = (semestr % 2 == 0) ? 0 : 1;
            int number = kurs_id * (semestr + x) - x;

            global_query = string.Format("select distinct predmet.id, name, name_krat, delenie, " +
                " zachet, ekzam, semestr " +
                " from predmet where actual=1 and " +
                " grupa_id = {0}" +
                " and semestr = {1}" +
                " and prepod_id = {2}" +
                " order by name", grups_set.Rows[grup_item][2],
                number, prepod_set.Rows[prepod_item][0]);

            /*global_query = string.Format("select distinct predmet.id, name, name_krat, delenie, " + 
                " zachet, " + 
                " ekzam, " + 
                " semestr " + 
                " from predmet " +
                " join vidzan_predmet on vidzan_predmet.predmet_id = predmet.id " +
                " join vid_zan on vid_zan.id = vidzan_predmet.vidzan_id " + 
                " where actual=1 and " +
                " grupa_id = {0}" +
                " and semestr = {1}" +
                " and prepod_id = {2}" +
                " order by name", grups_set.Rows[grup_item][2],
                number, prepod_set.Rows[prepod_item][0]);*/

            global_adapter = new SqlDataAdapter(global_query,global_connection);
            predmet_set = new DataTable();
            global_adapter.Fill(predmet_set);
   
            predmet_list.Items.Clear();
            for (int i = 0; i < predmet_set.Rows.Count; i++)
            {
                predmet_list.Items.Add(
                    predmet_set.Rows[i][1].ToString());
            }
            
            if (predmet_list.Items.Count > 0) predmet_list.SelectedIndex = 0; 
        }

        /// <summary>
        /// заполнить список видов занятий
        /// </summary>
        public void fill_vidzan()
        {
            if (predmet_list.Items.Count == 0)
            {
                vid_zan_list.Items.Clear();
                return;
            }            

            int kurs_id = 0;
            int grup_item = (grupa_list.SelectedIndex >= 0) ? grupa_list.SelectedIndex : 0;
            int prepod_item = (prepod_list.SelectedIndex >= 0) ? prepod_list.SelectedIndex : 0;
            int predm_item = (predmet_list.SelectedIndex >= 0) ? predmet_list.SelectedIndex : 0;

            global_query = "select " +
                " vid_zan.koef,  " + //0
                " vid_zan.name,  " + //1
                " vid_zan.id,  " +  //2
                " vid_zan.krat_name,  " +  //3
                " vid_zan.delenie, " +  //4
                " vidzan_predmet.kol_chas,  " + //5
                " vidspisan = vid_zan.spisanie,  " + //6
                " vid_zan.out_type,  " +                //7                 
                " predmetspisan = predmet_type.spisanie, " + //8
                " fakt_text, plan_text, vid_zan.name " +  //9, 10, 11
                " from vidzan_predmet " +
                " join vid_zan on vid_zan.id = vidzan_predmet.vidzan_id " +
                " join predmet on predmet.id=vidzan_predmet.predmet_id " +
                " join predmet_type on predmet_type.id = predmet.type_id " +
                " where vidzan_predmet.predmet_id =  " + predmet_set.Rows[predm_item][0].ToString() + 
                " and show_in_grid=1 " +
                " order by vid_zan.id";

            global_adapter = new SqlDataAdapter(global_query, global_connection);
            vidzan_set = new DataTable();
            global_adapter.Fill(vidzan_set);

            vid_zan_list.Items.Clear();

            for (int i = 0; i < vidzan_set.Rows.Count; i++)
			{
                vid_zan_list.Items.Add(
                    vidzan_set.Rows[i][1].ToString());
			}

            if (vid_zan_list.Items.Count > 0) vid_zan_list.SelectedIndex = 0;

            nado.Clear();
            fakt.Clear();

            string pr_id = predmet_set.Rows[predm_item][0].ToString();
            int predm_semestr = (int)predmet_set.Rows[predm_item][6];

            string date_filter = "";

            int sy = semestr2_start.Year,
                sd = semestr2_start.Day,
                sm = semestr2_start.Month;

            if (predm_semestr % 2 != 0)
            {
                date_filter = string.Format(" (y={0} or y={1}) ", starts[0].Year, ends[ends.Count-1].Year);
            }
            else
            {
                date_filter = string.Format(" y={0} ", ends[ends.Count - 1].Year);
            }

            //цикл получения количества часов
            for (int i = 0; i < vid_zan_list.Items.Count; i++)
            {

                DataRow dr = vidzan_set.Rows[i];
                string vid_id = dr[2].ToString();

                global_query = "select num = sum(rasp.kol_chas), c = count(rasp.kol_chas) from rasp " + 
                    " where rasp.predmet_id = " + pr_id +  
                    " and " + date_filter + 
                    " and " + 
                    " rasp.vid_zan_id =" + vid_id + 
                    " and prepod_id = " + prepod_set.Rows[prepod_list.SelectedIndex][0].ToString() +
                    " and rasp.uch_god_id = " + uch_god;
                
                //Text = pr_id.ToString() + "-" + vid_id.ToString();

                global_adapter = new SqlDataAdapter(global_query, global_connection);
                statistica_set = new DataTable();
                global_adapter.Fill(statistica_set);

                //вычисление факта                
                double fakt_chas = 0;
                if ((bool)dr[7] == true)
                {
                    if (!Convert.IsDBNull(statistica_set.Rows[0][0]))
                        fakt_chas = Convert.ToDouble(statistica_set.Rows[0][0]);// --- подумать о формуле !!!
                    else
                        fakt_chas = 0.0;
                }
                else
                {
                    if (!Convert.IsDBNull(statistica_set.Rows[0][1]))
                        fakt_chas = Convert.ToDouble(statistica_set.Rows[0][1]);
                    else
                        fakt_chas = 0.0;
                }
                fakt.Add(fakt_chas);

                //вычисление остатка
                double kol_chas = Convert.ToDouble(dr[5]);
                if ((bool)dr[8]==true)
                {
                    kol_chas = kol_chas * (1.0 - (int)dr[6] / 100.0);
                }                

                double ost = Convert.ToDouble(kol_chas - fakt_chas);
                nado.Add(ost);

            }
            
            statistica.Groups.Clear();
            statistica.Items.Clear();

            statistica.ShowItemToolTips = true;

            for (int i=0; i<vid_zan_list.Items.Count; i++)
            {
                ListViewGroup lvg = new ListViewGroup(vid_zan_list.Items[i].ToString(), vid_zan_list.Items[i].ToString());                
                statistica.Groups.Add(lvg);                
            }

            for (int i = 0; i < statistica.Groups.Count; i++)
            {
                
                string res = "";
                ListViewItem lvi = null;

                if ((bool)vidzan_set.Rows[i][7]==true)
                {
                    res = (fakt[i]>0)?fakt[i].ToString():"0";
                    lvi = new ListViewItem(res, statistica.Groups[i]);                 
                    res = (nado[i] > 0) ? nado[i].ToString("F0") : "0";
                    lvi.SubItems.Add(res);
                    lvi.ToolTipText = "Статистика по предмету: " + predmet_list.Text + "\n" +
                            "в группе " + grupa_list.Text;
                    statistica.Items.Add(lvi);                    
                }
                else
                {
                    if (fakt[i] == 0)
                    {
                        res = "еще не " + vidzan_set.Rows[i][9].ToString();
                        lvi = new ListViewItem(res, statistica.Groups[i]);
                        lvi.ToolTipText = "Статистика по предмету: " + predmet_list.Text + "\n" +
                            "в группе " + grupa_list.Text;
                        statistica.Items.Add(lvi); 
                    }
                    else
                    {
                        res = vidzan_set.Rows[i][9].ToString() + " или " + vidzan_set.Rows[i][10].ToString();
                        lvi = new ListViewItem(res, statistica.Groups[i]);
                        lvi.SubItems.Add(nado[i].ToString("F2"));
                        lvi.ToolTipText = "Статистика по предмету: " + predmet_list.Text + "\n" + 
                            "в группе " + grupa_list.Text;
                        statistica.Items.Add(lvi); 
                    }
                }                               
            }                       
        }

        /// <summary>
        /// заполнить список групп
        /// </summary>
        private void fill_grupa_list()
        {
            //получить группы данного FSystemа
            string q = "select grupa.name, grupa.kurs_id, grupa.id, zaoch from grupa  " +
                " join specialnost on specialnost.id = grupa.specialnost_id " +
                " where actual = 1 and show_in_grid = 1 and fakultet_id = " + fakultet_id.ToString ( ) +
                " order by outorder";
            SqlDataAdapter grups_adapter = new SqlDataAdapter ( q, global_connection );
            grups_set = new DataTable ( );

            grups_adapter.Fill ( grups_set );

            grups_count = grups_set.Rows.Count;

            probel.Clear ( );
            for ( int num = 0; num < grups_count*2+1; num++ )
            {
                probel.Add ( new String ( ' ', num + 1 ) );
            }

            //заполнить список групп
            grupa_list.Items.Clear ( );

            for ( int i = 0; i < grups_set.Rows.Count; i++ )
            {
                grupa_list.Items.Add ( grups_set.Rows[i][0].ToString ( ) );
            }
            
            if (grupa_list.Items.Count>0) grupa_list.SelectedIndex = 0;
            
        }

        
        /// <summary>
        /// заполнить список аудиторий
        /// </summary>
        public void fill_auds()
        {
            aud_list.Items.Clear();
            aud_set = new DataTable();
            

            if (predmet_list.Items.Count==0) return;
            if (vid_zan_list.Items.Count==0) return;

            if (predmet_list.Text.Trim().ToLower() != "физическая культура")
            {
                                //убрать те аудитории, которые уже заняты
                int rr = table.Row;
                int paranomer = (table_data.para_row.ContainsKey(rr)) ? table_data.RowPair(rr) : -1;
                DateTime datarasp = (table_data.date_row.ContainsKey(rr)) ? table_data.RowDate(rr) : new DateTime(2000, 1, 1);

                //if (paranomer == -1)
                    global_query = "select nomer, kabinet.id from kabinet " +
                        " join korpus on korpus.id=kabinet.korpus_id " +
                        " join fakultet on fakultet.korpus_id=korpus.id" +
                        " where fakultet.id = " + fakultet_id.ToString() +
                        " and not (nomer like '%стад%' or nomer like '%лыж%' or nomer like '%спорт%')" +
                        " order by prioritet";
                /*else
                    global_query = string.Format("select nomer, kabinet.id from kabinet " +
                        " join korpus on korpus.id=kabinet.korpus_id " +
                        " join fakultet on fakultet.korpus_id=korpus.id" +
                        " where fakultet.id = " + fakultet_id.ToString() +
                        " and nomer not in " +
                        "(select kabinet.nomer as kabname from rasp join kabinet on kabinet.id=rasp.kabinet_id where " +
                        " y={0} and m={1} and d={2} and nom_zan={3} and kabinet.nomer!='--')" +
                        " and not (nomer like '%стад%' or nomer like '%лыж%' or nomer like '%спорт%')" +
                        " order by prioritet", datarasp.Year, datarasp.Month, datarasp.Day, paranomer);*/
            }
            else
            {
                global_query = "select nomer, kabinet.id from kabinet " +
                    " join korpus on korpus.id=kabinet.korpus_id " +
                    " join fakultet on fakultet.korpus_id=korpus.id" +
                    " where fakultet.id = " + fakultet_id.ToString() +
                    " and (nomer like '%стад%' or nomer like '%лыж%' or nomer like '%спорт%')" +
                    " order by prioritet";
            }

            global_adapter = new SqlDataAdapter(global_query, global_connection);
            global_adapter.Fill(aud_set);

            foreach(DataRow r in aud_set.Rows)  aud_list.Items.Add(r[0]);

            aud_list.SelectedIndex = 0;
        }

        /// <summary>
        /// изменить теущего пользователя или FSystem
        /// </summary>      
        private void изменитьToolStripMenuItem_Click ( object sender, EventArgs e )
        {
            fakultet = true;
            dekan_online = false;
            from_change_user = false;
            set_fakultet ( );
            load_individ_rasp();
        }

        /// <summary>
        /// изменить текущего пользователя или FSystem
        /// </summary>       
        private void toolStripMenuItem1_Click ( object sender, EventArgs e )
        {
            fakultet = false;
            dekan_online = false;
            from_change_user = true;
            set_fakultet ( );
        }


        /// <summary>
        /// инициализация таблицы расписания
        /// </summary>
        private void init_table ( )
        {
            //заполнить таблицу расписания
            table.Cols = grups_count * 2 + 1;

            //вывести названия дней недели ------------
            int days = 0;
            for ( int j = 1; j < table.Rows; j += 7 )
            {
                
                for ( int i = 1; i < table.Cols; i++ )
                {
                    table.Select ( j, i, false );
                    table.CellBackColor = System.Drawing.Color.LightSteelBlue;
                    table.Cell[0, j, i] = starts[week_list.SelectedIndex].AddDays(days).ToShortDateString();
                }
                days++;
            }

            int num_d = 0;

            //объеденить строки, разделяющие дни недели
            for ( int i = 1; i < table.Rows; i += 7 )
            {

                table.Cell[0, i, 0] = DaysLong[num_d];
                table.set_MergeRow ( i, true );
                table.set_Cell ( C1.Win.C1FlexGrid.Classic.CellPropertySettings.flexcpFontName, i, 0,
                    (object) "georgia" );
                table.set_Cell ( C1.Win.C1FlexGrid.Classic.CellPropertySettings.flexcpFontBold,
                    i, 0, true );
                table.set_Cell ( C1.Win.C1FlexGrid.Classic.CellPropertySettings.flexcpForeColor,
                    i, 0, Color.Navy );
                Font f = new Font ( "georgia", 11 );
                table.set_RowHeight ( i, Convert.ToInt32 ( f.GetHeight ( ) ) + 10 );
                num_d++;
            }

            table.Col = 1;
            table.Row = 1;

            //соеднить колонки по 2
            string str = "";
            for ( int i = 0; i < table.Rows; i++ )
            {
                if ( empty_rows.Contains ( i ) ) continue;
                for ( int j = 1, k = 0; j < table.Cols; j += 2, k++ )
                {
                    table.Cell[0, i, j] = probel[j];
                    table.Cell[0, i, j + 1] = probel[j];
                    table.set_MergeRow ( i, true );
                }
                if ( i > 0 ) table.set_RowHeight ( i, 35 );
            }

            //установить названия групп
            groups.Clear();
            for ( int i = 1, j = 0; i < table.Cols; i += 2, j++ )
            {
                set_one_value ( grups_set.Rows[j][0].ToString ( ),
                    0, i + 1, false );
                groups.Add(grups_set.Rows[j][0].ToString());
                set_one_style ( 0, i + 1, false, 
                    C1.Win.C1FlexGrid.Classic.CellPropertySettings.flexcpFontBold,
                    true );
                set_one_style ( 0, i + 1, false, 
                    C1.Win.C1FlexGrid.Classic.CellPropertySettings.flexcpForeColor,
                    Color.Red );
                if ( (int)grups_set.Rows[j][1]%2==0)
                set_one_style ( 0, i + 1, false, 
                    C1.Win.C1FlexGrid.Classic.CellPropertySettings.flexcpBackColor,
                    Color.LightYellow);
                if ( (int)grups_set.Rows[j][1]%2!=0 )
                set_one_style ( 0, i + 1, false,
                    C1.Win.C1FlexGrid.Classic.CellPropertySettings.flexcpBackColor,
                    Color.Yellow);
                set_one_style(0, i + 1, false,
                    C1.Win.C1FlexGrid.Classic.CellPropertySettings.flexcpAlignment,
                    C1.Win.C1FlexGrid.Classic.AlignmentSettings.flexAlignCenterCenter);
            }          


            //выставить время занятий
            DateTime local_start = start_time, local_end = start_time;
            TimeSpan para = new TimeSpan(1, 20, 0);
            TimeSpan per1 = new TimeSpan(0, peremena, 0);
            TimeSpan per2 = new TimeSpan(0, long_peremena, 0);

            for (int i = 0; i < empty_rows.Count; i++)
            {
                local_start = start_time;
                int para_count = 1;
                for (int j = empty_rows[i] + 1; j < empty_rows[i] + 7; j++)
                {
                    local_end = local_start.Add(para);
                    table.Cell[0, j, 0] = local_start.ToShortTimeString() + " - " +
                        local_end.ToShortTimeString();
                    
                    if (para_count==first_long_peremena || para_count==second_long_peremena)
                        local_start = local_end.Add(per2); 
                    else
                        local_start = local_end.Add(per1); 

                    para_count++;
                }
            }

            fill_table();

        }

        /// <summary>
        /// установить одинаковый стиль для пары ячеек
        /// </summary>
        /// <param name="row">номер строки</param>
        /// <param name="col">номер столбца</param>
        /// <param name="check">признак провекри на допустимость присвания в эту ячейку</param>
        /// <param name="style">задаваемый стиль</param>
        /// <param name="newstyle">новое значение стиля</param>
        public void set_one_style(int row, int col, bool check, 
            C1.Win.C1FlexGrid.Classic.CellPropertySettings style, object newstyle)
        {

            if ( check )
            {
                if ( !is_correct_cell ( row, col ) ) return;
            }

            int start = 0, end = 0;

            if ( col % 2 == 0 )
            {
                start = col - 1;
                end = col;
            }
            else
            {
                start = col;
                end = col + 1;
            }
            
            table.set_Cell ( style, row, start, newstyle);
            table.set_Cell ( style, row, end, newstyle );
        }

        /// <summary>
        /// устновить одно значение в двух соседних ячеках одной группы
        /// добившись тем самым их объединения
        /// </summary>
        /// <param name="txt">текст для помещения в ячейку</param>
        /// <param name="row">номер строки</param>
        /// <param name="col">номер столбца</param>
        /// <param name="check">проверять ли корректность ячейки перед 
        /// присваиванием значения</param>
        public void set_one_value ( string outstr, int row, int col, bool check )
        {
            if ( check )
            {
                if ( !is_correct_cell ( row, col ) ) return;
            }

            int start = 0, end = 0;

            if ( col % 2 == 0 )
            {
                start = col - 1;
                end = col;
            }
            else
            {
                start = col;
                end = col + 1;
            }

            //проверить значения в соседних ячейках соседних групп для 
            //предотвращения их слияния в одну ячейку
            if (start > 2)
            {

                if (table.Cell[0, row, start - 2].ToString() == outstr)
                {
                    if ((start - 1) % 4 == 0)
                        outstr = " " + outstr;
                    else
                        outstr = outstr + " ";
                }
            }

            if (start < table.Cols - 2)
            {
                if (table.Cell[0, row, start + 2].ToString() == outstr)
                {
                    if ((start - 1) % 4 == 0)
                        outstr = " " + outstr;
                    else
                        outstr = outstr + " ";
                }
            }

            table.Cell[0, row, start] = outstr;
            table.Cell[0, row, end] = outstr;
        }

        // ----------   определить корректность строки, столбца, ячейки
        public bool is_correct_row ( int row )
        {
            if ( row == 0 ) return false;
            if ( empty_rows.Contains ( row ) ) return false;

            return true;
        }

        public bool is_correct_col ( int col )
        {
            if ( col == 0 ) return false;
            else
                return true;
        }

        public bool is_correct_cell ( int row, int col )
        {
            if ( is_correct_row ( row ) && is_correct_col ( col ) )
                return true;
            else
                return false;
        }

        // ----------------------------------------------------

        /// <summary>
        /// номер подгруппы, выбранной при слиянии ячеек
        /// </summary>
        public static int mergeresult = -1;


        /// <summary>
        /// произвести слияние со смежной ячейкой
        /// </summary>
        /// <param name="row">номер строки</param>
        /// <param name="col">номер столбца присоединяющей ячейки</param>
        private void merge ( int row, int col )
        {
            if (row <= 0 || col <= 0) return;
            
            if ( !is_correct_cell ( row, col ) ) return;

            int start = 0, end = 0;

            if ( col % 2 == 0 )
            {
                start = col - 1;
                end = col;
            }
            else
            {
                start = col;
                end = col + 1;
            }

            //простой случай - обе ячейки пустые
            if ( table.Cell[0, row, start].ToString ( ).Trim ( ) == "" &&
                table.Cell[0, row, end].ToString ( ).Trim ( ) == "" )
            {
                table.Cell[0, row, start] = probel[start];
                table.Cell[0, row, end] = probel[start];
                table.set_MergeRow ( start, true );
                return;
            }

            //Если ячейки уже соединены, выйти
            if ( table.Cell[0, row, start].ToString ( ) == table.Cell[0, row, end].ToString ( ) )
            {
                return;
            }

            left = table.Cell[0, row, start].ToString ( );
            right = table.Cell[0, row, end].ToString ( );

            Cell cell = table_data[cd, cgp, cp];    

            //получить выбор пользователя и поменять текст
            choose_merge cm = new choose_merge ( );
            cm.label_text = "В результате соединения  будет потеряно значение одной из ячеек.\n" +
             "Выберите значение, которое нужно оставить, нажав на соотвествующую кнопку.";
            DialogResult res = cm.ShowDialog ( );

            if (res != DialogResult.Cancel)
            {
                table.Cell[0, row, start] = goaltext;
                table.Cell[0, row, end] = goaltext;
            }
            else
            {
                cm.Dispose();
                return;
            }

            cm.Dispose();

            if (mergeresult == 2)
            {
                // --->>>>> удалить каждую запись если она непустая

                if (table_data[cd, cgp, cp].id[0] != 0)
                {
                    global_command = table_data[cd, cgp, cp].DeleteCommand(0);

                    global_command.ExecuteNonQuery();
                }

                if (table_data[cd, cgp, cp].id[1] != 0)
                {
                    global_command = table_data[cd, cgp, cp].DeleteCommand(1);
                    
                    global_command.ExecuteNonQuery();
                }

                table_data[cd, cgp, cp] = new Cell();
                return;
            }

            /// 1. пустая полуячейка распространилась на всю ячейку
            /// в первом случае для всех полей выставить значение по умолчанию (кроме даты и т.д.), 
            /// старая запись удаляется из БД

            if (table_data[cd, cgp, cp].prepod_id[mergeresult] == 0)
            {                
                // ---- >>> удалить перекрываемую запись из БД   <<<<< ----------- добавить сюда!
                global_command = table_data[cd, cgp, cp].DeleteCommand(1-mergeresult);
                global_command.ExecuteNonQuery();

                table_data[cd, cgp, cp] = new Cell();
                return;
            }
             
                        
            /// 2а. непустая полуячейка распространилась на пустую ячейку

            if (table_data[cd, cgp, cp].prepod_id[mergeresult] != 0 && table_data[cd, cgp, cp].prepod_id[1 - mergeresult] == 0)
            {                
                // скопировать все поля в элементы [1-mergeresult]
                table_data[cd, cgp, cp].copy_subgroups(mergeresult, 1 - mergeresult);
                table_data[cd, cgp, cp].subgr_nomer[0] = table_data[cd, cgp, cp].subgr_nomer[1] = 0;
                
                // ---- >>>> изменить данную запись в БД (ушло деление на подгруппы) <<<<< ----------- добавить сюда!
                global_command = table_data[cd, cgp, cp].UpdateCommand(0);
                global_command.ExecuteNonQuery();
                
                return;
            }
            
            /// 2б. непустая полуячейка распространилась на непустую ячейку
            /// меняется: Все + у исходной отменяется деление на подгруппы
            if (table_data[cd, cgp, cp].prepod_id[mergeresult] != 0 && table_data[cd, cgp, cp].prepod_id[1 - mergeresult] != 0)
            {
                //MessageBox.Show("непустая полуячейка распространилась на непустую ячейку");
                
                // ---- >>>> удалить перекрываемую запись в БД <<<<< ----------- добавить сюда!
                global_command = table_data[cd, cgp, cp].DeleteCommand(1 - mergeresult);
                global_command.ExecuteNonQuery();

                table_data[cd, cgp, cp].copy_subgroups(mergeresult, 1 - mergeresult);
                table_data[cd, cgp, cp].subgr_nomer[0] = table_data[cd, cgp, cp].subgr_nomer[1] = 0;

                // ---- >>>> изменить перекрывающую запись в БД (ушло деление на подгруппы) <<<<< ----------- добавить сюда!    
                global_command = table_data[cd, cgp, cp].UpdateCommand(mergeresult);
                global_command.ExecuteNonQuery();
            }

        }

        /// <summary>
        /// номер ячейки, выбранной при соединении двух подгрупп
        /// </summary>
        public static int unmerge_result = -1;

        /// <summary>
        /// произвести отмену слияния со смежной ячейкой
        /// </summary>
        /// <param name="row">номер строки</param>
        /// <param name="col">номер столбца присоединяющей ячейки</param>
        private void unmerge ( int row, int col )
        {
            if ( row == 0 || col == 0 ) return;
            if (!is_correct_cell(row, col)) return;

            int start = 0, end = 0;
            if ( col % 2 == 0 )
            {
                start = col - 1;
                end = col;
            }
            else
            {
                start = col;
                end = col + 1;
            }

            if ( table.Cell[0, row, start].ToString ( ).Trim ( ) !=         //ячейка уже разделена
                table.Cell[0, row, end].ToString ( ).Trim ( ) ) return;

            if ( table.Cell[0, row, start].ToString ( ).Trim ( ) == "" &&  //ячейка пустая
                table.Cell[0, row, end].ToString ( ).Trim ( ) == "" )
            {
                table.Cell[0, row, start] = probel[start];
                table.Cell[0, row, end] = probel[end];
                return;
            }


            if (!table_data.data[cd][cgp][cp].delenie[0])
            {
                MessageBox.Show("Деление на подгруппы для предмета \"" +
                    table_data[cd,cgp,cp].predmet_name[0] + "\"\n" +
                    "не предусмотрено учебным планом.\n",
                    "Ошибка редактирования",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else
            {
                /*bool alowed = table_data.data[cd][cgp][cp].vid_zan_name[0].ToLower().Contains("прак") ||
                    table_data[cd, cgp, cp].vid_zan_name[0].ToLower().Contains("сем") ||
                    table_data[cd, cgp, cp].vid_zan_name[0].ToLower().Contains("лаб");*/              

                if (!(bool)table_data[cd, cgp, cp].vid_delenie[0])
                {
                    MessageBox.Show("Деление на подгруппы для предмета \"" +
                        table_data[cd, cgp, cp].predmet_name[0] + "\"\n" +
                        "для данного вида занятий (" + table_data[cd, cgp, cp].vid_full_name[0] +  
                        ") не предусмотрено учебным планом.\n\n" + 
                        "Измените тип занятия на практическое, семинарское или лабораторное и повторите операцию.",
                        "Ошибка редактирования",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

            }

            left = "Поместить\nслева";
            right = "Поместить\nсправа";

            choose_merge cm = new choose_merge();
            cm.Text = "Куда поместить значение ячейки?";
            cm.label_text = "В результате разъдинения  будет создана новая ячейка.\n" +
             "В какую ячейку поместить существующее сейчас значение?";
            cm.button3.Visible = false;
            DialogResult res = cm.ShowDialog();

            if (res == DialogResult.Cancel)
            {
                cm.Dispose();
                return;
            }

            if (res == DialogResult.OK)
            {                
                table.Cell[0, row, end] = probel[end];
            }
            else
            {
                table.Cell[0, row, start] = probel[start];             
            }

            cm.Dispose();            

            if (unmerge_result == 1)
                table_data[cd, cgp, cp].copy_subgroups(0, 1);

            table_data[cd, cgp, cp].drop_subgroup(1 - unmerge_result);
            
            table_data[cd, cgp, cp].subgr_nomer[unmerge_result] = unmerge_result + 1;


            // ----->>>> обновить данную запись по полю подгруппа
            global_command = table_data[cd, cgp, cp].UpdateCommand(unmerge_result);  //кукарямба
            global_command.ExecuteNonQuery();

        }

        /// <summary>
        /// обработка нажатия на кнопку мыши в таблице
        /// </summary>    
        private void table_MouseDown ( object sender, MouseEventArgs e )
        {

            int c = table.MouseCol;
            int r = table.MouseRow;

            if (c < 0 || r < 0) return;

            movecell = false;

            if (dekan_online)
            {
                table.ContextMenuStrip = contextMenuStrip1;
            }
            else
            {
                table.ContextMenuStrip = teacher_menu;
            }

            for (int i = 0; i < contextMenuStrip1.Items.Count; i++)
            {
                contextMenuStrip1.Items[i].Enabled = false;
            }

            if (dekan_online)
            {
                соединитьЯчейкиToolStripMenuItem.Enabled = true;
                поменятьМестамиToolStripMenuItem.Enabled = true;
                разъеденитьЯчейкиToolStripMenuItem.Enabled = true;
            }

            for (int i = 0; i < 5; i++)
            {
                contextMenuStrip1.Items[i].Visible = false;
            }
            sep.Visible = false;

            for (int i = 0; i < teacher_menu.Items.Count; i++)
            {
                teacher_menu.Items[i].Enabled = false;
            }                   
            
            table.Focus();          

            if (!is_correct_cell(r, c)) return;
            
            if ( e.Button == MouseButtons.Right )
            {
                table.Col = c;
                table.Row = r;
            }

            //определить колонки подгрупп
            int first = 0, second = 0;
            if (c % 2 == 0)
            {
                first = c - 1;
                second = c;
            }
            else
            {
                first = c;
                second = c + 1;
            }

            grupa_list.SelectedIndex = (first - 1) / 2;

            int dest_col = first; //колонка назначения
            if (c == second) dest_col = second;
            if (c == first) dest_col = first;

            bool cell_divided = false;  //разделена ли исходная ячейка
            if (table.Cell[0, r, first].ToString() !=
                table.Cell[0, r, second].ToString())
                cell_divided = true;

            //определить параметры исходной ячейки
            DateTime d = table_data.RowDate(r);
            int p = table_data.RowPair(r);
            string g = table_data.ColumnGroup(dest_col);
            int sg = table_data.ColumnSubGroup(dest_col);
                       
            bool empty = table_data.data[d][g][p].prepod_id[sg - 1] == 0;

            if (empty)
            {
                set_teacher_attend.Enabled = false;
                set_teacher_tema.Enabled = false;

                return;
            }

            string vid_name = table_data.data[d][g][p].vid_full_name[sg - 1];

            if (dekan_online)
            {
                for (int i = 0; i < contextMenuStrip1.Items.Count; i++)
                {
                    contextMenuStrip1.Items[i].Enabled = true;
                }

                bool set_chas_allowed = vid_name.ToLower().Contains("зач") ||
                    vid_name.ToLower().Contains("курс") ||
                    (vid_name.ToLower().Contains("экзамен") && !vid_name.ToLower().Contains("конс"));
                if (set_chas_allowed)
                    set_chas.Enabled = false;
            }

            if (active_user_id == table_data.data[d][g][p].prepod_id[sg - 1])
            {
                sep.Visible = true;

                //показать пункт меню, соотвествующий типу занятия              
                if (vid_name.ToLower().Contains("зач"))
                {
                    set_zachet.Enabled = true;
                    set_zachet.Visible = true;
                    выставитьОтметкиОЗачёте.Enabled = true;                      
                    return;
                }

                if (vid_name.ToLower().Contains("курс"))
                {
                    set_kurs.Enabled = true;
                    set_kurs.Visible = true;
                    выставитьОтметкиКурсовойРаботы.Enabled = true;
                    return;
                }

                if (vid_name.ToLower().Contains("экзамен") && !vid_name.ToLower().Contains("конс"))
                {
                    set_exam.Enabled = true;
                    set_exam.Visible = true;
                    выставитьОтметкиОбЭкзамене.Enabled = true;
                    return;
                }

                set_teacher_attend.Enabled = true;
                set_teacher_tema.Enabled = true;
                
                set_attend.Enabled = true;                
                set_tema.Enabled = true;
                set_tema.Visible = true;
                set_attend.Visible = true;
                set_chas.Enabled = true;
            }
            else
            {
                set_teacher_attend.Enabled = false;
                set_teacher_tema.Enabled = false;
                выставитьОтметкиОЗачёте.Enabled = false;
                выставитьОтметкиКурсовойРаботы.Enabled = false;
                выставитьОтметкиОбЭкзамене.Enabled = false;
            }
        }


        private void choose_vid_zan (object sender, EventArgs e)
        {
            //MessageBox.Show("выбрано");
        }

        private void table_MouseUp(object sender, MouseEventArgs e)
        {
        }

        /// <summary>
        /// обработка выбора пунка соединения ячеек
        /// </summary>
        private void соединитьЯчейкиToolStripMenuItem_Click ( object sender, EventArgs e )
        {
            if (server_date.AddDays(-stopweeks * 7) > starts[week_list.SelectedIndex])
            {
                MessageBox.Show("Редактирование расписания запрещено, так как с момента его создания прошло более " +
                    stopweeks.ToString() + weekword,
                    "Отклонение запроса",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (dekan_online)
            merge(table.Row, table.Col);
        }

        /// <summary>
        /// обработка выбора пункта меню "разъеденить ячейки"
        /// </summary>
        private void разъеденитьЯчейкиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (server_date.AddDays(-stopweeks * 7) > starts[week_list.SelectedIndex])
            {
                MessageBox.Show("Редактирование расписания запрещено, так как с момента его создания прошло более " +
                    stopweeks.ToString() + weekword,
                    "Отклонение запроса",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (dekan_online)
                unmerge(table.Row, table.Col);
        }

        /// <summary>
        /// обработка выбора пункта меню "поменять местами"
        /// </summary>
        private void поменятьМестамиToolStripMenuItem_Click(object sender, EventArgs e)
        {

            if (server_date.AddDays(-stopweeks * 7) > starts[week_list.SelectedIndex])
            {
                MessageBox.Show("Редактирование расписания запрещено, так как с момента его создания прошло более " +
                    stopweeks.ToString() + weekword,
                    "Отклонение запроса",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            
            if (!dekan_online) return;

            int row = table.Row;
            int col = table.Col;

            if (row == 0 || col == 0) return;

            int start = 0, end = 0;
            if (col % 2 == 0)
            {
                start = col - 1;
                end = col;
            }
            else
            {
                start = col;
                end = col + 1;
            }

            if (table.Cell[0, row, start].ToString().Trim() ==
                table.Cell[0, row, end].ToString().Trim()) return;

            if (table.Cell[0, row, start].ToString().Trim() == "" &&
                table.Cell[0, row, end].ToString().Trim() == "")
                return;

            string tmp = table.Cell[0, row, start].ToString();
            table.Cell[0, row, start] = table.Cell[0, row, end];
            table.Cell[0, row, end] = tmp;

            if (table.Cell[0, row, end].ToString().Trim().Length == 0)
                table.Cell[0, row, end] = probel[end];

            if (table.Cell[0, row, start].ToString().Trim().Length == 0)
                table.Cell[0, row, start] = probel[start];


            table_data[cd, cgp, cp].swap_subgroups();

            if (table_data[cd, cgp, cp].prepod_id[0] != 0)
            {
                // ---->>>> обновить запиись в БД 
                global_command = table_data[cd, cgp, cp].UpdateCommand(0);
                global_command.ExecuteNonQuery();
            }

            if (table_data[cd, cgp, cp].prepod_id[1] != 0)
            {
                // ---->>>> обновить запиись в БД                
                global_command = table_data[cd, cgp, cp].UpdateCommand(1);
                global_command.ExecuteNonQuery();
            }

            
        }


        
        public int old_col=0, old_row=0;
        string oldtxt = "";
        /// <summary>
        /// дата текущей ячейки
        /// </summary>
        public DateTime cd;
        /// <summary>
        /// группа текущей ячейки
        /// </summary>
        public string cgp = "";
        /// <summary>
        /// подшруппа текущей ячейки
        /// </summary>
        public int csub = 0;
        /// <summary>
        /// номер пары текущей ячейки
        /// </summary>
        public int cp = 0;
        
        /// <summary>
        /// обработка перемещения мыши над таблицей
        /// </summary>
        private void table_MouseMove(object sender, MouseEventArgs e)
        {          
            int col = table.MouseCol;
            int row = table.MouseRow;

            if (col < 0 || row < 0) return;

            //table_data[table_data.RowDate(col), table_data.ColumnGroup(col), 1].lekt[0] = true;

            if (table_data==null) return;
            if (!is_correct_cell(row, col)) return;

            cd = table_data.RowDate(row);
            cgp = table_data.ColumnGroup(col);
            csub = table_data.ColumnSubGroup(col);
            cp = table_data.RowPair(row);

            string txt = "";
            if (table_data[cd, cgp, cp].tema[csub - 1].Trim().Length != 0)
                txt = String.Format("Дата:{0}, Группа:{1}, Пара: #{2}, Преподаватель:{3}, Предмет:{4}, Тема:{5}, Кол. часов:{6}",
                    cd.ToShortDateString(), cgp, cp,
                    table_data[cd, cgp, cp].prepod_name[csub - 1],
                    table_data[cd, cgp, cp].predmet_name[csub - 1],
                    table_data[cd, cgp, cp].tema[csub - 1],
                    table_data[cd, cgp, cp].col_chas[csub - 1],
                    table_data[cd, cgp, cp].id[csub - 1]);
            else
                txt = String.Format("Дата:{0}, Группа:{1}, Пара: #{2}, Преподаватель:{3}, Предмет:{4}, Кол. часов:{5}",
                cd.ToShortDateString(), cgp, cp,
                table_data[cd, cgp, cp].prepod_name[csub - 1],
                table_data[cd, cgp, cp].predmet_name[csub - 1],
                table_data[cd, cgp, cp].col_chas[csub - 1],
                table_data[cd, cgp, cp].id[csub - 1]);

            if (table_data[cd, cgp, cp].prepod_id[csub - 1] != 0)
            {
                stat_text.Text = txt;
                if (toolTip2.ToolTipTitle == "")
                {
                    
                    string shorttext = String.Format("Дата:{0}\nТема:{1}\nКол-во часов:{2}",
                    cd.ToShortDateString(),
                    (table_data[cd, cgp, cp].tema[csub - 1].Length == 0) ? "не задана" : table_data[cd, cgp, cp].tema[csub - 1],
                    table_data[cd, cgp, cp].col_chas[csub - 1]);
                    toolTip2.SetToolTip(table, shorttext);
                    toolTip2.InitialDelay = 1000;
                    toolTip2.ToolTipTitle = "Активная ячейка";
                }
            }
            else
            {
                stat_text.Text = "";
                if (toolTip2.ToolTipTitle != "")
                {
                    toolTip2.SetToolTip(table, "");
                    toolTip2.ToolTipTitle = "";
                }
            }                       

            old_col = col;
            old_row = row;            
        }

        private void week_list_Click(object sender, EventArgs e)
        {
        }

        /// <summary>
        /// обработка выбора новой учебной недели
        /// </summary>
        private void week_list_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (use_week_list_change) fill_table();
            use_week_list_change = true;
            table.Focus();
            table.Select();
        }

        //можно ли перемещаться в списке групп при перемещении курсора таблицы с по
        //мощью клавиатуры
        //bool can_move_left = true;
        //bool can_move_right = true;
        
        /// <summary>
        /// обработка выбора группы в списке
        /// </summary>
        private void grupa_list_SelectedIndexChanged(object sender, EventArgs e)
        {
            fill_prepods();
            fill_predemts();
            fill_vidzan();

            if (prepod_list.Items.Count>0) prepod_list.SelectedIndex = 0;
            if (predmet_list.Items.Count > 0) predmet_list.SelectedIndex = 0;
            if (vid_zan_list.Items.Count > 0) vid_zan_list.SelectedIndex = 0;          

            //перейти к выделенной группе
            table.Focus();

            if (empty_rows.Contains(table.Row)) table.Row++;

            if (movecell)
            {
                if (grupa_list.SelectedIndex == grupa_list.Items.Count - 1)
                    table.Col = grupa_list.SelectedIndex * 2 + 2;
                else
                    table.Col = grupa_list.SelectedIndex * 2 + 1;
            }

            movecell = true;            

            /*int c = table.Col;  // -- код для обработки клавиатурного интерфейса
            int r = table.Row;            

            if (!is_correct_cell(r, c))
            {
                can_move_left = true;
                can_move_right = true;
                return;
            }

            //определить колонки подгрупп
            int first = 0, second = 0;
            if (c % 2 == 0)
            {
                first = c - 1;
                second = c;
            }
            else
            {
                first = c;
                second = c + 1;
            }

            bool cell_divided = false;  //разделена ли целевая ячейка
            if (table.Cell[0, r, first].ToString() !=
                table.Cell[0, r, second].ToString())
                cell_divided = true;

            if (cell_divided)
            {
                if (c == first)
                {
                    can_move_left = true;
                    can_move_right = false;
                }

                if (c == second)
                {
                    can_move_left = false;
                    can_move_right = true;
                }

            }
            else
            {
                can_move_left = true;
                can_move_right = true;
            }*/

        }


        /// <summary>
        /// обработка щелчка по таблице
        /// </summary>
        private void table_Click(object sender, EventArgs e)
        {
            //if (table_data.date_row.ContainsKey(table.Row))
            //Text = table_data.RowDate(table.Row).ToShortDateString();
        }
        
        /// <summary>
        /// обработка выбора преподавателя в списке
        /// </summary>
        private void prepod_list_SelectedIndexChanged(object sender, EventArgs e)
        {
            fill_predemts();
            fill_vidzan();
         
            if (predmet_list.Items.Count > 0) predmet_list.SelectedIndex = 0;
            if (vid_zan_list.Items.Count > 0) vid_zan_list.SelectedIndex = 0;            
        }

        /// <summary>
        /// обработка выбора предмета в списке
        /// </summary>
        private void predmet_list_SelectedIndexChanged(object sender, EventArgs e)
        {
            fill_vidzan();
            fill_auds();
        }

        private void aud_list_MouseDown(object sender, MouseEventArgs e)
        {
            if (grupa_list.Items.Count == 0) return;
            if (prepod_list.Items.Count == 0) return;
            if (predmet_list.Items.Count == 0) return;
            if (vid_zan_list.Items.Count == 0) return;
            if (aud_list.Items.Count == 0) return;

            DragDropEffects dde = DoDragDrop("predmet", DragDropEffects.Copy | DragDropEffects.Scroll);
        }

        /// <summary>
        /// обработка события перемещения объекта над таблицей
        /// </summary>        
        private void table_DragOver(object sender, DragEventArgs e)
        {
            //if ((ListBox)sender != predmet_list) return;
            

            table.Col = table.MouseCol;
            table.Row = table.MouseRow;

            e.Effect = DragDropEffects.Copy | DragDropEffects.Scroll;
        }

        
        /// <summary>
        /// обработка события отпускания перетаскиваемого объекта
        /// </summary>
        /// <param name="sender">отправитель</param>
        /// <param name="e">параметры события</param>
        private void table_DragDrop(object sender, DragEventArgs e)
        {            
            string source = e.Data.GetData ( DataFormats.Text ).ToString ( );
            if (source != "predmet") return;

            set_cell();
        }

        /// <summary>
        /// выставить пункт расписания на тот же день недели на неделю вперёд
        /// </summary>
        public void push_week()
        {
            if (!dekan_online) return;

            // копирует занятие на неделю вперёд от текущей даты в данной ячейке
            int c = table.Col;
            int r = table.Row;


            if (!is_correct_cell(r, c))
            {
                MessageBox.Show("Выбранная ячейка не может быть использована.", "Ошибка редактирования",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            //определить колонки подгрупп
            int first = 0, second = 0;
            if (c % 2 == 0)
            {
                first = c - 1;
                second = c;
            }
            else
            {
                first = c;
                second = c + 1;
            }

            bool cell_divided = false;  //разделена ли целевая ячейка
            if (table.Cell[0, r, first].ToString() !=
                table.Cell[0, r, second].ToString())
                cell_divided = true;

            int dest_col = first; //колонка назначения
            if (c == second) dest_col = second;
            if (c == first) dest_col = first;

            //определить параметры целевой ячейки
            DateTime d = table_data.RowDate(r);
            int p = table_data.RowPair(r);
            string g = table_data.ColumnGroup(dest_col);
            int sg = table_data.ColumnSubGroup(dest_col);

            int colnum = (cell_divided) ? (sg - 1) : 0;

            DateTime newdate = d.AddDays(7);
            string dd = newdate.Day.ToString();
            string mm = newdate.Month.ToString();
            string yy = newdate.Year.ToString();
            
            string predm_id = table_data[d, g, p].predmet_id[colnum].ToString();
            string gr_id = table_data[d, g, p].grupa_id[colnum].ToString();
            string prepod_id = table_data[d, g, p].prepod_id[colnum].ToString();
            string fak_id = table_data[d, g, p].fakultet_id[colnum].ToString();
            string kurs_id = table_data[d, g, p].kurs_id[colnum].ToString();
            string para_id = table_data[d, g, p].nom_zan[colnum].ToString();
            string vid_zan = table_data[d, g, p].vid_zan_id[colnum].ToString();
            // контрольные виды занятий повторно ставить?? спросить!!
            string kabi_id = table_data[d, g, p].kabinet_id[colnum].ToString();
            string semestr_id = table_data[d, g, p].semestr_id[colnum].ToString();
            string potok_id = rnd.Next().ToString();
            string kol_chas = table_data[d, g, p].col_chas[colnum].ToString();
            string uch_god = table_data[d, g, p].uch_god[colnum].ToString();
            string sub_gr = table_data[d, g, p].subgr_nomer[colnum].ToString();

            DataTable selectTable;

            // свободен ли преподаватель
            string sql = "select grupa.name from rasp " + 
                " join grupa on grupa.id = rasp.grupa_id " +
                " where d=@d and m=@m and y=@y and rasp.nom_zan=@nomzan and rasp.prepod_id=@prepod_id";
            selectTable = new DataTable();
            global_command = new SqlCommand(sql, global_connection);
            global_command.Parameters.Add("@d", SqlDbType.Int).Value = dd;
            global_command.Parameters.Add("@m", SqlDbType.Int).Value = mm;
            global_command.Parameters.Add("@y", SqlDbType.Int).Value = yy;
            global_command.Parameters.Add("@nomzan", SqlDbType.Int).Value = para_id;
            global_command.Parameters.Add("@prepod_id", SqlDbType.Int).Value = prepod_id;
            (new SqlDataAdapter(global_command)).Fill(selectTable);

            if (selectTable.Rows.Count > 0)
            {
                MessageBox.Show("Преподаватель, ведущий данный предмет, на следующей неделе на данной паре уже задействован в группе " + 
                    selectTable.Rows[0][0].ToString() + "!\n" +
                    "Вероятно, следует освободить преподавателя от занятий в другой группе на этой паре на дату " + newdate.ToShortDateString(),
                    "Отказ операции",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }


            // свободна ли аудитория
            sql = "select kabinet.nomer, grupa.name from rasp " +
                " join kabinet on kabinet.id = rasp.kabinet_id " +
                " join grupa on grupa.id = rasp.grupa_id " +
                " where d=@d and m=@m and y=@y and rasp.fakultet_id=@fid and rasp.nom_zan=@nomzan";
            selectTable = new DataTable();
            global_command = new SqlCommand(sql, global_connection);
            global_command.Parameters.Add("@d", SqlDbType.Int).Value = dd;
            global_command.Parameters.Add("@m", SqlDbType.Int).Value = mm;
            global_command.Parameters.Add("@y", SqlDbType.Int).Value = yy;
            global_command.Parameters.Add("@fid", SqlDbType.Int).Value = fak_id;
            global_command.Parameters.Add("@nomzan", SqlDbType.Int).Value = para_id;            
            (new SqlDataAdapter(global_command)).Fill(selectTable);

            if (selectTable.Rows.Count > 0)
            {
                MessageBox.Show("Аудитория №" + selectTable.Rows[0][0].ToString() +
                    ", куда запланировано поставить занятие, на следующей неделе на данной паре уже занята в группе " +
                    selectTable.Rows[0][1].ToString() + "!\n\n" +
                    "Занятие всё-таки будет поставлено в сетку расписания без установленной аудитории.\nНомер аудитории следует выбрать отдельно.",
                    "Коррекция операции",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                kabi_id = "7";
            }


            sql = "select predmet.name, grupa.name, rasp.id from rasp " + 
                " join predmet on predmet.id = rasp.predmet_id " + 
                " join grupa on grupa.id = rasp.grupa_id " +
                " where d=@d and m=@m and y=@y and rasp.fakultet_id=@fid and rasp.nom_zan=@nomzan and rasp.grupa_id=@gr_id";
            global_command = new SqlCommand(sql, global_connection);
            global_command.Parameters.Add("@d", SqlDbType.Int).Value = dd;
            global_command.Parameters.Add("@m", SqlDbType.Int).Value = mm;
            global_command.Parameters.Add("@y", SqlDbType.Int).Value = yy;
            global_command.Parameters.Add("@fid", SqlDbType.Int).Value = fak_id;
            global_command.Parameters.Add("@nomzan", SqlDbType.Int).Value = para_id;
            global_command.Parameters.Add("@gr_id", SqlDbType.Int).Value = gr_id;         

            DataTable celltable = new DataTable();
            global_adapter = new SqlDataAdapter(global_command);
            global_adapter.Fill(celltable);

            if (celltable.Rows.Count > 0)
            {
                DialogResult dr =
                MessageBox.Show("На выбранные Вами номер занятия и дату " + newdate.ToShortDateString() +
                    " уже запланировано другое занятие.\n" +
                    "Группа: " + celltable.Rows[0][1].ToString() + Environment.NewLine +
                    "Предмет: " + celltable.Rows[0][0].ToString() + Environment.NewLine + Environment.NewLine +
                    "Вы желаете заменить указанное занятие на выбранное вами?\n\n",
                    "Запрос программы",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (dr == DialogResult.Yes)
                {
                    //выполнить команду обновления
                    sql = "UPDATE rasp " + 
                        " SET [d]=@d, [m]=@m, [y]=@y, [predmet_id]=@predmet_id, [grupa_id]=@grupa_id, " + 
                        "     [prepod_id]=@prepod_id, [fakultet_id]=@fakultet_id, [kurs_id]=@kurs_id, [nom_zan]=@nom_zan, " + 
                        "     [vid_zan_id]=@vid_zan_id, [kabinet_id]=@kabinet_id, " + 
                        "     [semestr_id]=@semestr_id, [potok_id]=@potok_id, [subgr_nomer]=@subgr_nomer, [kol_chas]=@kol_chas, " + 
                        "     [uch_god_id]=@uch_god_id  WHERE [id]= @id";
                    global_command = new SqlCommand(sql, global_connection);
                    global_command.Parameters.Add("@d", SqlDbType.Int).Value = dd;
                    global_command.Parameters.Add("@m", SqlDbType.Int).Value = mm;
                    global_command.Parameters.Add("@y", SqlDbType.Int).Value = yy;
                    global_command.Parameters.Add("@predmet_id", SqlDbType.Int).Value = predm_id;
                    global_command.Parameters.Add("@grupa_id", SqlDbType.Int).Value = gr_id;
                    global_command.Parameters.Add("@prepod_id", SqlDbType.Int).Value = prepod_id;
                    global_command.Parameters.Add("@fakultet_id", SqlDbType.Int).Value = fak_id;
                    global_command.Parameters.Add("@kurs_id", SqlDbType.Int).Value = kurs_id;
                    global_command.Parameters.Add("@nom_zan", SqlDbType.Int).Value = para_id;
                    global_command.Parameters.Add("@vid_zan_id", SqlDbType.Int).Value = vid_zan;
                    global_command.Parameters.Add("@kabinet_id", SqlDbType.Int).Value = kabi_id;
                    global_command.Parameters.Add("@semestr_id", SqlDbType.Int).Value = semestr_id;
                    global_command.Parameters.Add("@potok_id", SqlDbType.Int).Value = potok_id;
                    global_command.Parameters.Add("@subgr_nomer", SqlDbType.Int).Value = sub_gr;
                    global_command.Parameters.Add("@kol_chas", SqlDbType.Float).Value = kol_chas;
                    global_command.Parameters.Add("@uch_god_id", SqlDbType.Int).Value = uch_god;
                    global_command.Parameters.Add("@id", SqlDbType.Int).Value = celltable.Rows[0][2].ToString();

                    try
                    {
                        global_command.ExecuteNonQuery();
                    }
                    catch (Exception exx)
                    {
                        MessageBox.Show(
                            "В данный момент из-за сетевого сбоя выполнение операции невозможно. Попробуйте выполнить действие ещё раз.",
                            "Ошибка", 
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                        if (global_connection.State == ConnectionState.Broken)
                        {
                            global_connection.Close();
                            global_connection.Open();
                        }
                    }
                }
            }
            else
            {
                //Выпонлить вставку                
                sql = "INSERT INTO rasp(" +
                            "[d], [m], [y], [predmet_id], [grupa_id], [prepod_id], [fakultet_id], [kurs_id], [nom_zan], [vid_zan_id], [kabinet_id], [semestr_id], [potok_id], [subgr_nomer], [kol_chas], [uch_god_id])" +
                " VALUES    (@d,  @m,  @y,  @predmet_id,  @grupa_id,  @prepod_id,  @fakultet_id,  @kurs_id,  @nom_zan,  @vid_zan_id,  @kabinet_id,  @semestr_id,  @potok_id,  @subgr_nomer,  @kol_chas,  @uch_god_id)";
                global_command = new SqlCommand(sql, global_connection);
                global_command.Parameters.Add("@d", SqlDbType.Int).Value = dd;
                global_command.Parameters.Add("@m", SqlDbType.Int).Value = mm;
                global_command.Parameters.Add("@y", SqlDbType.Int).Value = yy;
                global_command.Parameters.Add("@predmet_id", SqlDbType.Int).Value = predm_id;
                global_command.Parameters.Add("@grupa_id", SqlDbType.Int).Value = gr_id;
                global_command.Parameters.Add("@prepod_id", SqlDbType.Int).Value = prepod_id;
                global_command.Parameters.Add("@fakultet_id", SqlDbType.Int).Value = fak_id;
                global_command.Parameters.Add("@kurs_id", SqlDbType.Int).Value = kurs_id;
                global_command.Parameters.Add("@nom_zan", SqlDbType.Int).Value = para_id;
                global_command.Parameters.Add("@vid_zan_id", SqlDbType.Int).Value = vid_zan;
                global_command.Parameters.Add("@kabinet_id", SqlDbType.Int).Value = kabi_id;
                global_command.Parameters.Add("@semestr_id", SqlDbType.Int).Value = semestr_id;
                global_command.Parameters.Add("@potok_id", SqlDbType.Int).Value = potok_id;
                global_command.Parameters.Add("@subgr_nomer", SqlDbType.Int).Value = sub_gr;
                global_command.Parameters.Add("@kol_chas", SqlDbType.Float).Value = kol_chas;
                global_command.Parameters.Add("@uch_god_id", SqlDbType.Int).Value = uch_god;
                
                try
                {
                    global_command.ExecuteNonQuery();
                }
                catch (Exception exx)
                {
                    MessageBox.Show(
                        "В данный момент из-за сетевого сбоя выполнение операции невозможно. Попробуйте выполнить действие ещё раз." + 
                        "\n" + exx.Message,
                        "Ошибка",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    if (global_connection.State == ConnectionState.Broken)
                    {
                        global_connection.Close();
                        global_connection.Open();
                    }
                }
            }
        }


        /// <summary>
        /// задать значение предмета и сопутствующих значений в текущей клетке
        /// </summary>
        public void set_cell()
        {
            //if (server_date <= starts[0]) return;
            int c = table.Col;
            int r = table.Row;

            if (!dekan_online) return;

            if (server_date.AddDays(-stopweeks * 7) > starts[week_list.SelectedIndex])
            {
                MessageBox.Show("Редактирование расписания запрещено, так как с момента его создания прошло более " +
                    stopweeks.ToString() + weekword,
                    "Отклонение запроса",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (prepod_list.Items.Count == 0) return;
            if (predmet_list.Items.Count == 0) return;
            if (vid_zan_list.Items.Count == 0) return;
            if (aud_list.Items.Count == 0) return;

            if (!is_correct_cell(r, c))
            {
                MessageBox.Show("Выбранная ячейка не может быть использована.", "Ошибка редактирования",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            //определить колонки подгрупп
            int first = 0, second = 0;
            if (c % 2 == 0)
            {
                first = c - 1;
                second = c;
            }
            else
            {
                first = c;
                second = c + 1;
            }

            bool cell_divided = false;  //разделена ли целевая ячейка
            if (table.Cell[0, r, first].ToString() !=
                table.Cell[0, r, second].ToString())
                cell_divided = true;

            int dest_col = first; //колонка назначения
            if (c == second) dest_col = second;
            if (c == first) dest_col = first;

            //определить параметры целевой ячейки
            DateTime d = table_data.RowDate(r);
            int p = table_data.RowPair(r);
            string g = table_data.ColumnGroup(dest_col);
            int sg = table_data.ColumnSubGroup(dest_col);

            //определить тип выполняемой операции - обнолвение или вставка
            bool update = ((table_data.data[d][g][p].prepod_id[sg - 1] != 0));// || (table_data.data[d][g][p].prepod_id[1] != 0));

            //если группа приемник не равна группе источнику
            if (grupa_list.Text != table.Cell[0, 0, dest_col].ToString())
            {
                MessageBox.Show("Группа назначения выбрана неправильно.",
                    "Ошибка редактирования",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (cell_divided) //проверить, возможно ли разделение на подгруппы
            {
                bool del = Convert.ToBoolean(predmet_set.Rows[predmet_list.SelectedIndex]["delenie"]);

                if (!del)
                {
                    MessageBox.Show("Выбранный вами предмет не предусматривает деление на подгруппы.\n" +
                        "Выберите другой предмет или отмените разделение целевой ячейки.",
                        "Ошибка редактирования",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                string vid = vid_zan_list.Items[vid_zan_list.SelectedIndex].ToString().ToLower();
                bool alowed = (bool)vidzan_set.Rows[vid_zan_list.SelectedIndex][4];

                if (!alowed)
                {
                    MessageBox.Show("Деление на подгруппы для предмета \"" +
                        predmet_list.Text + "\"\n" +
                        "для данного вида занятий (" + vid + ") не предусмотрено учебным планом.\n\n" +
                        "Измените тип занятия на практическое, семинарское или лабораторное и повторите операцию.",
                        "Ошибка редактирования",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            //проверка размещения данных в резделенных ячейках
            if (cell_divided)
            {
                if (table_data.data[d][g][p].prepod_id[1 - (sg - 1)] == Convert.ToInt32(prepod_set.Rows[prepod_list.SelectedIndex][0]))
                {
                    MessageBox.Show("Невозможно поместить предметы одного преподавателя в одну ячейку два раза.",
                        "Ошибка редактирования",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (table_data[d, g, p].kabinet_id[1 - (sg - 1)] == Convert.ToInt32(aud_set.Rows[aud_list.SelectedIndex][1]))
                {
                    MessageBox.Show("Невозможно проведение двух разных занятий в одной аудитории (№" +
                        aud_set.Rows[aud_list.SelectedIndex][0].ToString() + ") .",
                        "Ошибка редактирования",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            Cell cell_candidate = null; //ячейка-предендент на сохранение в БД

            cell_candidate = new Cell();

            if (cell_divided) // задать значение только для одной ячейки (есть деление)
            {
                if (table_data.data[d][g][p].prepod_id[0] != 0 || table_data.data[d][g][p].prepod_id[1] != 0)
                {                    
                    cell_candidate.copy_fields(table_data[d, g, p]);
                }

                if (update)
                    cell_candidate.id[sg - 1] = table_data[d, g, p].id[sg - 1];
                else
                    cell_candidate.id[sg - 1] = -1;

                cell_candidate.nom_zan[sg - 1] = p;

                cell_candidate.y[sg - 1] = d.Year;
                cell_candidate.m[sg - 1] = d.Month;
                cell_candidate.d[sg - 1] = d.Day;

                cell_candidate.fakultet_id[sg - 1] = fakultet_id;
                cell_candidate.grupa_id[sg - 1] = Convert.ToInt32(grups_set.Rows[grupa_list.SelectedIndex][2]);

                cell_candidate.predmet_id[sg - 1] = Convert.ToInt32(predmet_set.Rows[predmet_list.SelectedIndex][0]);
                cell_candidate.predmet_name[sg - 1] = predmet_set.Rows[predmet_list.SelectedIndex][2].ToString();
                cell_candidate.delenie[sg - 1] = (bool)predmet_set.Rows[predmet_list.SelectedIndex][3];
                cell_candidate.vid_delenie[sg - 1] = (bool)vidzan_set.Rows[vid_zan_list.SelectedIndex][4];

                cell_candidate.prepod_id[sg - 1] = Convert.ToInt32(prepod_set.Rows[prepod_list.SelectedIndex][0]);
                cell_candidate.prepod_name[sg - 1] = prepod_set.Rows[prepod_list.SelectedIndex][3].ToString();

                cell_candidate.subgr_nomer[sg - 1] = sg;

                cell_candidate.aud_name[sg - 1] = aud_set.Rows[aud_list.SelectedIndex][0].ToString();
                cell_candidate.kabinet_id[sg - 1] = Convert.ToInt32(aud_set.Rows[aud_list.SelectedIndex][1]);

                cell_candidate.kurs_id[sg - 1] = Convert.ToInt32(grups_set.Rows[grupa_list.SelectedIndex][1]);
                cell_candidate.potok_id[sg - 1] = -1; //?

                int x = (semestr % 2 == 0) ? 0 : 1;
                int number = cell_candidate.kurs_id[sg - 1] * (semestr + x) - x;
                cell_candidate.semestr_id[sg - 1] = number;

                cell_candidate.vid_zan_id[sg - 1] = (int)vidzan_set.Rows[vid_zan_list.SelectedIndex][2];
                cell_candidate.vid_zan_name[sg - 1] = vidzan_set.Rows[vid_zan_list.SelectedIndex][3].ToString();
                cell_candidate.vid_full_name[sg - 1] = vidzan_set.Rows[vid_zan_list.SelectedIndex][11].ToString();
                
                cell_candidate.col_chas[sg - 1] = 2.0;
                
                if (cell_candidate.vid_full_name[sg - 1].ToLower().Contains("зач"))
                {
                    cell_candidate.col_chas[sg - 1] = (double)nado[vid_zan_list.SelectedIndex];
                }

                if (cell_candidate.vid_full_name[sg - 1].ToLower().Contains("экзамен")
                    && (!cell_candidate.vid_full_name[sg - 1].ToLower().Contains("перед")))
                {
                    cell_candidate.col_chas[sg - 1] = (double)nado[vid_zan_list.SelectedIndex];
                }

                if (cell_candidate.vid_full_name[sg - 1].ToLower().Contains("курсовая"))
                {
                    cell_candidate.col_chas[sg - 1] = 0.0;
                }

                if (cell_candidate.vid_full_name[sg - 1].ToLower().Contains("контрольная"))
                {
                    cell_candidate.col_chas[sg - 1] = 0.0;
                }


            }
            else   // задать одинаковые значения для двух ячеек (нет деления)
            {
                if (update)
                {
                    int pid = table_data[d, g, p].id[0];

                    cell_candidate.id[0] = pid;
                    cell_candidate.id[1] = pid;
                }
                else
                {
                    cell_candidate.id[0] = -1; 
                    cell_candidate.id[1] = -1;
                }

                cell_candidate.nom_zan[0] = cell_candidate.nom_zan[1] = p;

                cell_candidate.y[0] = cell_candidate.y[1] = d.Year;
                cell_candidate.m[0] = cell_candidate.m[1] = d.Month;
                cell_candidate.d[0] = cell_candidate.d[1] = d.Day;

                cell_candidate.fakultet_id[0] = cell_candidate.fakultet_id[1] = fakultet_id;

                cell_candidate.grupa_id[0] = cell_candidate.grupa_id[1] = Convert.ToInt32(grups_set.Rows[grupa_list.SelectedIndex][2]);

                cell_candidate.predmet_id[0] = cell_candidate.predmet_id[1] = Convert.ToInt32(predmet_set.Rows[predmet_list.SelectedIndex][0]);
                cell_candidate.predmet_name[0] = cell_candidate.predmet_name[1] = predmet_set.Rows[predmet_list.SelectedIndex][2].ToString();
                
                cell_candidate.delenie[0] = cell_candidate.delenie[1] = (bool)predmet_set.Rows[predmet_list.SelectedIndex][3];
                cell_candidate.vid_delenie[0] = cell_candidate.vid_delenie[1] = (bool)vidzan_set.Rows[vid_zan_list.SelectedIndex][4];

                cell_candidate.prepod_id[0] = cell_candidate.prepod_id[1] = Convert.ToInt32(prepod_set.Rows[prepod_list.SelectedIndex][0]);
                cell_candidate.prepod_name[0] = cell_candidate.prepod_name[1] = prepod_set.Rows[prepod_list.SelectedIndex][3].ToString();

                cell_candidate.subgr_nomer[0] = cell_candidate.subgr_nomer[1] = 0;

                cell_candidate.aud_name[0] = cell_candidate.aud_name[1] = aud_set.Rows[aud_list.SelectedIndex][0].ToString();
                cell_candidate.kabinet_id[0] = cell_candidate.kabinet_id[1] = Convert.ToInt32(aud_set.Rows[aud_list.SelectedIndex][1]);

                cell_candidate.kurs_id[0] = cell_candidate.kurs_id[1] = Convert.ToInt32(grups_set.Rows[grupa_list.SelectedIndex][1]);
                cell_candidate.potok_id[0] = cell_candidate.potok_id[1] = -1; //?

                int x = (semestr % 2 == 0) ? 0 : 1;
                int number = cell_candidate.kurs_id[0] * (semestr + x) - x;
                cell_candidate.semestr_id[0] = cell_candidate.semestr_id[1] = number;

                cell_candidate.vid_zan_id[0] = cell_candidate.vid_zan_id[1] = (int)vidzan_set.Rows[vid_zan_list.SelectedIndex][2];
                cell_candidate.vid_zan_name[0] = cell_candidate.vid_zan_name[1] = vidzan_set.Rows[vid_zan_list.SelectedIndex][3].ToString();
                cell_candidate.vid_full_name[0] = cell_candidate.vid_full_name[1] = vidzan_set.Rows[vid_zan_list.SelectedIndex][11].ToString();

                cell_candidate.col_chas[0] = cell_candidate.col_chas[1] = 2.0;
                if (cell_candidate.vid_full_name[sg - 1].ToLower().Contains("зач"))
                {
                    cell_candidate.col_chas[0] = cell_candidate.col_chas[1] = 
                        (double)nado[vid_zan_list.SelectedIndex];
                }

                if (cell_candidate.vid_full_name[sg - 1].ToLower().Contains("экзамен")
                    && (!cell_candidate.vid_full_name[sg - 1].ToLower().Contains("перед")))
                {
                    cell_candidate.col_chas[0] = cell_candidate.col_chas[1] = Convert.ToDouble(
                        (double)nado[vid_zan_list.SelectedIndex]);
                }
                if (cell_candidate.vid_full_name[sg - 1].ToLower().Contains("курсовая"))
                {
                    cell_candidate.col_chas[0] = cell_candidate.col_chas[1] = 0.0;
                }
                if (cell_candidate.vid_full_name[sg - 1].ToLower().Contains("контрольная"))
                {
                    cell_candidate.col_chas[0] = cell_candidate.col_chas[1] = 0.0;
                }

            }


            string tmp = "";

            if (cell_candidate != table_data.data[d][g][p])
            {
                //занят ли преподаватель 
                if (table_data.IsPrepodBuisy(cell_candidate, sg, out tmp,
                    Convert.ToBoolean(prepod_set.Rows[prepod_list.SelectedIndex][2])))
                {
                    MessageBox.Show(tmp, "Ошибка редактирования",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                //занята ли аудитория
                if (table_data.IsRoomBuisy(cell_candidate, sg, out tmp))
                {
                    MessageBox.Show(tmp, "Ошибка редактирования",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }


            //вывод в клетку
            string predm = cell_candidate.predmet_name[sg - 1];
            string prep = "/" + cell_candidate.prepod_name[sg - 1] + "/";
            string aud = "";
            if (cell_candidate.aud_name[sg - 1] != "--") aud = ", " + cell_candidate.aud_name[sg - 1];
            string vidzan = cell_candidate.vid_zan_name[sg - 1];

            string outstr = predm + ", " + vidzan + "\n" + prep + aud;

            //проверить значения в соседних ячейках соседних групп для 
            //предотвращения их слияния в одну ячейку
            if (first > 2)
            {

                if (table.Cell[0, r, first - 2].ToString() == outstr)
                {
                    if ((first - 1) % 4 == 0)
                        outstr = " " + outstr;
                    else
                        outstr = outstr + " ";
                }
            }

            if (first < table.Cols - 2)
            {
                if (table.Cell[0, r, first + 2].ToString() == outstr)
                {
                    if ((first - 1) % 4 == 0)
                        outstr = " " + outstr;
                    else
                        outstr = outstr + " ";
                }
            }

            if (!cell_divided)
                set_one_value(outstr, r, c, false);
            else
                table.Cell[0, r, dest_col] = outstr;

             //сохраненние в базу данных либо обновить, либо вставить
            table_data.data[d][g][p] = cell_candidate;

            bool succesBD = false;

            if (update)
            {
                // обновление ---->>>>>>>>>
                global_command = table_data[d, g, p].UpdateCommand(sg - 1);
                int res = global_command.ExecuteNonQuery();
            }
            else
            {

                // вставка  ----->>>>>>>>>>                
                global_command = table_data[d, g, p].InsertCommand(sg - 1);
                int res = global_command.ExecuteNonQuery();

                if (res != 0)
                {
                    global_command.CommandText = "select @@identity";
                    int newid = Convert.ToInt32(global_command.ExecuteScalar());
                    table_data[d, g, p].id[sg - 1] = newid;                    
                }
            }

        }

        /// <summary>
        /// возвращает ячеку с указанными координатами
        /// </summary>
        /// <param name="r">номер строки</param>
        /// <param name="c">номер столбца</param>
        /// <returns></returns>
        public Cell cell(int r, int c)
        {
            return table_data[table_data.RowDate(r), table_data.ColumnGroup(c), 
                table_data.RowPair(r)];
        }

        
        /// <summary>
        /// Вызов окна настроек расписания
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void настройкиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //учесть привилегии пользователя
            if (!dekan_online)
            {
                MessageBox.Show("Сообщение для пользователя: " + active_user_name + "\n\n" +
                    "Извините, но у Вас нет прав на редактирование свойств таблицы расписания.",
                    "Отклонение запроса",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }


            rasp_nastroika rn = new rasp_nastroika();
            DialogResult dr = rn.ShowDialog();

            if (dr == DialogResult.Cancel)
            {
                rn.Dispose();
                return;
            }

            ///сохранить изменения,
            ///внесенные при работе в окне настроек
            
            //для групп
            if (rn.grups_changed)
            {
                string cmd = "";
                for (int i = 0; i < rn.listBox1.Items.Count; i++)
                {
                    string gr = rn.listBox1.Items[i].ToString();
                    int id = rn.gr_id[gr];
                    cmd = cmd +
                        string.Format("update grupa set outorder = {0}, show_in_grid = {1} " + 
                        "where id={2};", rn.order[id], Convert.ToInt32(rn.status[id]), id);
                }

                /*System.IO.StreamWriter sr = new System.IO.StreamWriter("c:\\r.txt");
                sr.Write(cmd);
                sr.Close();*/

                //MessageBox.Show(cmd);
                SqlCommand sqlcmd = new SqlCommand(cmd, global_connection);
                sqlcmd.ExecuteNonQuery();            
                
                //снова заполнить таблицу
                fill_tree();
                fill_grupa_list();
                grupa_list.SelectedIndex = 0;
                init_table();
            }

            rn.Dispose();
        }

        private void table_DoubleClick(object sender, EventArgs e)
        {          
            if (dekan_online) set_cell();
        }

        /// <summary>
        /// перечисление для указания направления копирования или перемещения
        /// </summary>
        public enum direction
        {
            /// <summary>
            /// копировать вверх
            /// </summary>
            с_upforward, 
            /// <summary>
            /// копировать вниз
            /// </summary>
            с_downforward, 
            /// <summary>
            /// копировать крест-накрест вверх (только при наличии деления)
            /// </summary>
            с_upcross, 
            /// <summary>
            /// копировать крест-накрест вниз (только при наличии деления)
            /// </summary>
            с_downcross,
            /// <summary>
            /// переместить вверх
            /// </summary>
            m_upforward, 
            /// <summary>
            /// переместить вниз
            /// </summary>
            m_downforward, 
            /// <summary>
            /// переместить крест-накрест вверх (только при наличии деления)
            /// </summary>
            m_upcross, 
            /// <summary>
            /// переместить крест-накрест вниз (только при наличии деления)
            /// </summary>
            m_downcross
        }

        /// <summary>
        /// произвести копирование или перемещение ячейки в указанном направлении
        /// </summary>
        /// <param name="dir">направление и тип операции</param>
        public void copy_move_cell(direction dir)
        {
            int c = table.Col;
            int r = table.Row;

            if (!is_correct_cell(r, c))
            {
                return;
            }

            if (server_date.AddDays(-stopweeks * 7) > starts[week_list.SelectedIndex])
            {
                MessageBox.Show("Редактирование расписания запрещено, так как с момента его создания прошло более " +
                    stopweeks.ToString() + weekword,
                    "Отклонение запроса",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            //определить колонки подгрупп
            int first = 0, second = 0;
            if (c % 2 == 0)
            {
                first = c - 1;
                second = c;
            }
            else
            {
                first = c;
                second = c + 1;
            }

            int dest_col = first; //колонка назначения
            if (c == second) dest_col = second;
            if (c == first) dest_col = first;

            bool cell_divided = false;  //разделена ли исходная ячейка
            if (table.Cell[0, r, first].ToString() !=
                table.Cell[0, r, second].ToString())
                cell_divided = true;

            //определить параметры исходной ячейки
            DateTime d = table_data.RowDate(r);
            int p = table_data.RowPair(r);
            string g = table_data.ColumnGroup(dest_col);
            int sg = table_data.ColumnSubGroup(dest_col);

            bool empty = (table_data.data[d][g][p].prepod_id[0] == 0 && table_data.data[d][g][p].prepod_id[1] == 0);

            if (empty) return;

            //проверить возможность диагональной операции
            if (!cell_divided)
            {
                if (dir == direction.с_upcross || dir == direction.с_downcross
                    || dir == direction.m_upcross || dir == direction.m_downcross)
                {
                    MessageBox.Show("Для выполнения этой операции исходная ячейка должна быть разделена.",
                        "Ошибка редактирования",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            int shift = 1;
            int sign = -1;

            if (dir == direction.m_upcross || dir == direction.m_upforward || dir == direction.с_upcross || dir == direction.с_upforward)
            {
                sign = -1;
            }
            else
            {
                sign = 1;
            }
            
            int s_shift = shift * sign;

            if (!is_correct_cell(r + s_shift, c))
            {
                MessageBox.Show("Ячейка назначения не может быть использована.", "Ошибка редактирования",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            Cell cell_candidate = table_data[d, g, p];
            //Cell cell_candidate = new Cell();
            //cell_candidate.copy_fields(table_data[d, g, p]);
            cell_candidate.nom_zan[0] = p + s_shift;
            cell_candidate.nom_zan[1] = p + s_shift;
            
            string tmp = "";

            for (int i = 0; i < 2; i++)
            {
                if (cell_candidate.prepod_id[i] != 0)
                {
                    //занят ли преподаватель
                    if (table_data.IsPrepodBuisy(cell_candidate, i + 1, out tmp, true))
                    {                        
                        MessageBox.Show(tmp, "Операция остановлена",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    //занята ли аудитория
                    if (table_data.IsRoomBuisy(cell_candidate, i + 1, out tmp))
                    {
                        MessageBox.Show(tmp, "Операция остановлена",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }

            int r0 = r; //запомнить номер исходной строки
            r = r + s_shift; //рассматриваем целевую ячейку

            //определить параметры целевой ячейки
            DateTime d1 = table_data.RowDate(r);
            int p1 = table_data.RowPair(r);
            string g1 = table_data.ColumnGroup(dest_col);
            int sg1 = table_data.ColumnSubGroup(dest_col);

            bool empty1 = (table_data.data[d1][g1][p1].prepod_id[0] == 0 && table_data.data[d1][g1][p1].prepod_id[1] == 0);
                                
            //если целевая ячейка не пуста, то удалить обе подгруппы из БД и из структ. данных
            if (!empty1)   ///---------------------------------??????????????????????????????????????????? 1??
            {
                if (table_data.data[d1][g1][p1].id[0] != 0)
                {
                    // --->>>> удалить для первой подгруппы      
                    global_command = table_data[d1, g1, p1].DeleteCommand(0);
                    global_command.ExecuteNonQuery();
                }

                if (table_data.data[d1][g1][p1].id[1] != 0)
                {
                    // --->>>> удалить для второй подгруппы
                    global_command = table_data[d1, g1, p1].DeleteCommand(1);
                    global_command.ExecuteNonQuery();
                }
            }

            //сбросить значение ячейки в структуре данных
            table_data.data[d1][g1][p1] = new Cell();

            if (cell_divided) //очистить целевые ячейки
            {
                table.Cell[0, r, second] = probel[second];
                table.Cell[0, r, first] = probel[first];
            }
            else
            {
                table.Cell[0, r, second] = probel[first];
                table.Cell[0, r, first] = probel[first];
            }            

            //копировать данные из исходной ячейки
            table_data[d1, g1, p1].copy_fields(table_data[d, g, p]);
            if (dir == direction.с_upcross || dir == direction.с_downcross
                    || dir == direction.m_upcross || dir == direction.m_downcross)
            {
                table_data[d1,g1,p1].swap_subgroups();
            }

            table_data[d, g, p].nom_zan[0] = table_data[d, g, p].nom_zan[1] = p; 
            table_data[d1, g1, p1].nom_zan[0] = table_data[d1, g1, p1].nom_zan[1] = p1;
            

            //вывод в клетку  ------------------ 
            string outstr1 = get_cell_text(d1, g1, p1, 0);
            string outstr2 = get_cell_text(d1, g1, p1, 1);


            if (!cell_divided)
            {
                //проверить значения в соседних ячейках соседних групп для 
                //предотвращения их слияния в одну ячейку
                if (first > 2)
                {

                    if (table.Cell[0, r, first - 2].ToString() == outstr1)
                    {
                        if ((first - 1) % 4 == 0)
                            outstr1 = " " + outstr1;
                        else
                            outstr1 = outstr1 + " ";
                    }
                }

                if (first < table.Cols - 2)
                {
                    if (table.Cell[0, r, first + 2].ToString() == outstr1)
                    {
                        if ((first - 1) % 4 == 0)
                            outstr1 = " " + outstr1;
                        else
                            outstr1 = outstr1 + " ";
                    }
                }

                set_one_value(outstr1, r, c, false);
            }
            else
            {
                if (outstr1.Trim().Length > 0)
                    table.Cell[0, r, first] = outstr1;
                else
                    table.Cell[0, r, first] = probel[first];

                if (outstr2.Trim().Length > 0)
                    table.Cell[0, r, second] = outstr2;
                else
                    table.Cell[0, r, second] = probel[second];

            }   // ----  конец вывода в клетку


            //  ---->>>>>>  вставить ячейки в БД (только непустые) 

            // если ячейка не разделена, то создавать только одну запись
            if (!cell_divided)
            {
                global_command = table_data[d1, g1, p1].InsertCommand(0);
                int res = global_command.ExecuteNonQuery();
                if (res != 0)
                {
                    global_command.CommandText = "select @@identity";
                    int newid = Convert.ToInt32(global_command.ExecuteScalar());
                    table_data[d1, g1, p1].id[0] = newid;
                    //MessageBox.Show("присваивание нового ИД");
                    table_data[d1, g1, p1].copy_subgroups(0, 1);
                }
            }
            else
            {
                if (table_data[d1, g1, p1].prepod_id[0] != 0)
                {
                    ////------>>>>>>>>>> вставка
                    global_command = table_data[d1, g1, p1].InsertCommand(0);
                    int res = global_command.ExecuteNonQuery();
                    if (res != 0)
                    {
                        global_command.CommandText = "select @@identity";
                        int newid = Convert.ToInt32(global_command.ExecuteScalar());
                        table_data[d1, g1, p1].id[0] = newid;
                        //MessageBox.Show("присваивание нового ИД");
                    }

                }

                if (table_data[d1, g1, p1].prepod_id[1] != 0)
                {
                    ////------>>>>>>>>>> вставка
                    global_command = table_data[d1, g1, p1].InsertCommand(1);
                    int res = global_command.ExecuteNonQuery();
                    if (res != 0)
                    {
                        global_command.CommandText = "select @@identity";
                        int newid = Convert.ToInt32(global_command.ExecuteScalar());
                        table_data[d1, g1, p1].id[1] = newid;
                        //MessageBox.Show("присваивание нового ИД");
                    }
                }
            }

            // если задана операция перемещения, удалить данные об исходных ячейках и очистить их
            // только для непустых ячеек
            if (dir == direction.m_upforward || dir == direction.m_downforward ||
                dir == direction.m_downcross || dir == direction.m_upcross)
            {
                if (cell_divided)
                {
                    if (table_data[d, g, p].prepod_id[0] != 0)
                    {
                        //-------->>>>> удаление
                        //MessageBox.Show("Первая удаляется:" + table_data[d, g, p].id[0].ToString());
                        global_command = table_data[d, g, p].DeleteCommand(0);
                        int res = global_command.ExecuteNonQuery();


                    }

                    if (table_data[d, g, p].prepod_id[1] != 0)
                    {
                        //-------->>>>> удаление
                        //MessageBox.Show("Первая удаляется:" + table_data[d, g, p].id[0].ToString());
                        global_command = table_data[d, g, p].DeleteCommand(1);
                        int res = global_command.ExecuteNonQuery();

                    }
                }
                else
                {
                    if (table_data[d, g, p].prepod_id[0] != 0)
                    {
                        //-------->>>>> удаление
                        //MessageBox.Show("Первая удаляется:" + table_data[d, g, p].id[0].ToString());
                        global_command = table_data[d, g, p].DeleteCommand(0);
                        int res = global_command.ExecuteNonQuery();


                    }
                }

                table_data[d, g, p] = new Cell();
                
                if (cell_divided) //очистить исходные ячейки
                {
                    table.Cell[0, r0, second] = probel[second];
                    table.Cell[0, r0, first] = probel[first];
                }
                else
                {
                    table.Cell[0, r0, second] = probel[first];
                    table.Cell[0, r0, first] = probel[first];
                } 
            }
        }

        public void delete_cell()
        {
            int c = table.Col;
            int r = table.Row;

            if (!is_correct_cell(r, c))
            {
                return;
            }

            //определить колонки подгрупп
            int first = 0, second = 0;
            if (c % 2 == 0)
            {
                first = c - 1;
                second = c;
            }
            else
            {
                first = c;
                second = c + 1;
            }

            int dest_col = first; //колонка назначения
            if (c == second) dest_col = second;
            if (c == first) dest_col = first;

            bool cell_divided = false;  //разделена ли исходная ячейка
            if (table.Cell[0, r, first].ToString() !=
                table.Cell[0, r, second].ToString())
                cell_divided = true;

            //определить параметры исходной ячейки
            DateTime d = table_data.RowDate(r);
            int p = table_data.RowPair(r);
            string g = table_data.ColumnGroup(dest_col);
            int sg = table_data.ColumnSubGroup(dest_col);

            bool empty = (table_data.data[d][g][p].prepod_id[0] == 0 && table_data.data[d][g][p].prepod_id[1] == 0);

            if (empty) return;

            //если целевая ячейка не пуста, то удалить обе подгруппы из БД и из структ. данных
            if (cell_divided)
            {
                if (table_data.data[d][g][p].prepod_id[sg - 1] != 0)
                {
                    // --->>>> удалить для целевой подгруппы     
                    global_command = table_data[d, g, p].DeleteCommand(sg - 1);
                    try
                    {
                        global_command.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {
                        //
                    }
                }

                table.Cell[0, r, dest_col] = probel[dest_col];
                table_data[d, g, p].drop_subgroup(sg - 1);
            }
            else
            {
                if (table_data.data[d][g][p].prepod_id[0] != 0)
                {
                    // --->>>> удалить для Первой подгруппы
                    global_command = table_data[d, g, p].DeleteCommand(0);
                    global_command.ExecuteNonQuery();
                }

                if (table_data.data[d][g][p].prepod_id[1] != 0)
                {
                    // --->>>> удалить для второй подгруппы
                    global_command = table_data[d, g, p].DeleteCommand(1);
                    global_command.ExecuteNonQuery();
                }

                table.Cell[0, r, second] = probel[first];
                table.Cell[0, r, first] = probel[first];

                table_data[d, g, p] = new Cell();
            }                            
        }

        public void dell_both_in_cell()
        {
            int c = table.Col;
            int r = table.Row;

            if (!is_correct_cell(r, c))
            {
                return;
            }

            //определить колонки подгрупп
            int first = 0, second = 0;
            if (c % 2 == 0)
            {
                first = c - 1;
                second = c;
            }
            else
            {
                first = c;
                second = c + 1;
            }

            int dest_col = first; //колонка назначения
            if (c == second) dest_col = second;
            if (c == first) dest_col = first;

            bool cell_divided = false;  //разделена ли исходная ячейка
            if (table.Cell[0, r, first].ToString() !=
                table.Cell[0, r, second].ToString())
                cell_divided = true;

            //определить параметры исходной ячейки
            DateTime d = table_data.RowDate(r);
            int p = table_data.RowPair(r);
            string g = table_data.ColumnGroup(dest_col);
            int sg = table_data.ColumnSubGroup(dest_col);

            bool empty = (table_data.data[d][g][p].prepod_id[0] == 0 && table_data.data[d][g][p].prepod_id[1] == 0);

            if (empty) return;
            if (!cell_divided) return;

            //если целевая ячейка не пуста, то удалить обе подгруппы из БД и из структ. данных
            if (table_data.data[d][g][p].prepod_id[0] != 0)
            {
                    // --->>>> удалить для Первой подгруппы
                global_command = table_data[d, g, p].DeleteCommand(0);
                global_command.ExecuteNonQuery();
            }

            if (table_data.data[d][g][p].prepod_id[1] != 0)
            {
                    // --->>>> удалить для второй подгруппы
                global_command = table_data[d, g, p].DeleteCommand(1);
                global_command.ExecuteNonQuery();
            }
            
            table.Cell[0, r, second] = probel[first];
            table.Cell[0, r, first] = probel[first];
            
            table_data[d, g, p] = new Cell();            
        }

        private void отправитьВExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveToExcel();
        }

        private void upcopy_Click(object sender, EventArgs e)
        {
            copy_move_cell(direction.с_upforward);
        }

        private void downcopy_Click(object sender, EventArgs e)
        {
            copy_move_cell(direction.с_downforward);
        }

        private void toolStripMenuItem5_Click(object sender, EventArgs e)
        {
            copy_move_cell(direction.с_upcross);
        }

        private void toolStripMenuItem6_Click(object sender, EventArgs e)
        {
            copy_move_cell(direction.с_downcross);
        }

        private void toolStripMenuItem7_Click(object sender, EventArgs e)
        {
            copy_move_cell(direction.m_upforward);
        }

        private void toolStripMenuItem8_Click(object sender, EventArgs e)
        {
            copy_move_cell(direction.m_downforward);
        }

        private void toolStripMenuItem9_Click(object sender, EventArgs e)
        {
            copy_move_cell(direction.m_upcross);
        }

        private void toolStripMenuItem10_Click(object sender, EventArgs e)
        {
            copy_move_cell(direction.m_downcross);
        }

        private void toolStripMenuItem11_Click(object sender, EventArgs e)
        {
            if (server_date.AddDays(-stopweeks * 7) > starts[week_list.SelectedIndex])
            {
                MessageBox.Show("Редактирование расписания запрещено, так как с момента его создания прошло более " +
                    stopweeks.ToString() + weekword,
                    "Отклонение запроса",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            delete_cell();
        }

        private void toolStripMenuItem12_Click(object sender, EventArgs e)
        {
            if (server_date.AddDays(-stopweeks * 7) > starts[week_list.SelectedIndex])
            {
                MessageBox.Show("Редактирование расписания запрещено, так как с момента его создания прошло более " +
                    stopweeks.ToString() + weekword,
                    "Отклонение запроса",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            dell_both_in_cell();
        }


        // bool rt = false, lt = false;
        private void table_KeyDown(object sender, KeyEventArgs e)
        {
            //поствить в текущую ячейку
            if (e.Shift) set_cell();
            if (e.KeyCode == Keys.Return) set_cell();

            if (!e.Control)  //перемещение по спискам
            {
                switch (e.KeyValue)
                {
                    case 90:
                    case 85: cycle_list(grupa_list); break;
                    case 88:
                    case 71: cycle_list(prepod_list); break;
                    case 67:
                    case 72: cycle_list(predmet_list); break;
                    case 86:
                    case 68: cycle_list(vid_zan_list); break;
                    case 66:
                    case 70: cycle_list(aud_list); break;
                }
            }

            if (e.Control)//перемещение по спискам
            {
                switch (e.KeyValue)
                {
                    case 90:
                    case 85: cycle_list_back(grupa_list); break;
                    case 88:
                    case 71: cycle_list_back(prepod_list); break;
                    case 67:
                    case 72: cycle_list_back(predmet_list); break;
                    case 86:
                    case 68: cycle_list_back(vid_zan_list); break;
                    case 66:
                    case 70: cycle_list_back(aud_list); break;
                }
            }

            table.Focus();
            table.Select();
        }

       /// <summary>
       /// циклическое перемещение по списку назад
       /// </summary>
       /// <param name="lb">список для выполнения операции</param>
        private void cycle_list_back(ListBox lb)
        {
            if (lb.SelectedIndex < 0) return;

            int c = lb.Items.Count - 1;
            int l = lb.SelectedIndex;

            if (l == 0)
            {
                lb.SelectedIndex = c;
            }
            else
                lb.SelectedIndex--;
        }

        /// <summary>
        /// циклическое перемещение по списку вперед
        /// </summary>
        /// <param name="lb">список для выполнения операции</param>
        public void cycle_list(ListBox lb)
        {
            if (lb.SelectedIndex < 0) return;

            int c = lb.Items.Count - 1;
            int l = lb.SelectedIndex;

            if (l == c)
            {
                lb.SelectedIndex = 0;
            }
            else
                lb.SelectedIndex++;
        }        

        /// <summary>
        /// получить данные из графического файла
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public string FilePhoto = "";
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

        /// <summary>
        /// сохранить фото в БД
        /// </summary>
        /// <param name="id"></param>
        private void save_prepod_photo_by_id(int id)
        {
            //DataGridViewRow cr = dataGridView1.CurrentRow;
            //int id = (int)cr.Cells["id"].Value;

            byte[] photo = GetPhotoFromFile(FilePhoto);
          
            global_command = new SqlCommand();
            global_command.CommandText = "update prepod set " + 
                " photo = @p where id = @id";
            global_command.Connection = global_connection;

            global_command.Parameters.Add("@p", SqlDbType.Image, photo.Length).Value = photo;
            global_command.Parameters.Add("@id", SqlDbType.Int).Value = id;

            try
            {
                global_command.ExecuteNonQuery();
            }
            catch (Exception exx)
            {
                MessageBox.Show(exx.Message);
            }

        }

        /// <summary>
        /// получить фото из БД
        /// </summary>
        /// <param name="tablename">имя таблицы из которой извлекается фото</param>
        /// <param name="id">идентификатор записи для которой извлекается фото</param>
        /// <returns></returns>
        public static Bitmap GetPhotoFromBD(string tablename, int id)
        {
            
            global_command = new SqlCommand("SELECT photo FROM " + tablename + " where id = " + id.ToString(),
                global_connection);                                             
            
            SqlDataReader reader = global_command.ExecuteReader(CommandBehavior.SequentialAccess);
            
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

        /// <summary>
        /// выставление информации о зачете
        /// </summary>
        public void goto_zachet(int predm_id, string predname)
        {
            teacher_predmet.Text = "Предмет " + predname;
            teacher_tab.SelectedIndex = 1;
            teacher_tab_predmet.SelectedIndex = 3;          
        }

        //работа с вкладками предметов -------------------------------------------- 

        // разрешить/запретить редактирование баллов
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (AttMarks == null) return;
            for (int i = 0; i < 15; i++)
            {
                AttMarks[i].Enabled = checkBox1.Checked;
            }
        }

        DataTable dtt = new DataTable(); //таблица видов занятий
        DataTable TemaTable = new DataTable(); //таблица тематических блоков
        DataTable Datatable = null; // таблица шкалы оценивания предмета
        
        // массив имен компонентов вывода значений баллов для аттестации
        NumericUpDown[] AttMarks = null;

        string[] MarksFieldsNames = new string[15] {
            "mrs_2", "mrs_2p", "mrs_3", "mrs_4", "mrs_5",
            "mrs1_2", "mrs1_2p", "mrs1_3", "mrs1_4", "mrs1_5",
            "mrs2_2", "mrs2_2p", "mrs2_3", "mrs2_4", "mrs2_5"};
        
        /// <summary>
        /// вывести сведения по МБРС на вкладку teacher_tab_predmet_MRS
        /// </summary>
        /// <param name="pr_id">системный ид предмета</param>
        public void PrintMBRS(int pr_id, string pr_name)
        {
            tema_list_listBox.Tag = pr_id;
            BallPlanLabel.Text = string.Empty;

            AttMarks = new NumericUpDown[15] { 
                numericUpDown1,  numericUpDown2,  numericUpDown3, 
                numericUpDown4,  numericUpDown5,  numericUpDown6, 
                numericUpDown7,  numericUpDown8,  numericUpDown9, 
                numericUpDown10, numericUpDown11, numericUpDown12, 
                numericUpDown13, numericUpDown14, numericUpDown15 };

            // вывод текущих значений
            Datatable = new DataTable();
            string SqlGetMarks = string.Format("select mrs_2, mrs_2p, mrs_3, mrs_4, mrs_5," +
                " mrs1_2, mrs1_2p, mrs1_3, mrs1_4, mrs1_5," +
                " mrs2_2, mrs2_2p, mrs2_3, mrs2_4, mrs2_5," +
                " mrs, kredit, grupa_id from predmet where id = {0}", pr_id);
            (new SqlDataAdapter(SqlGetMarks, global_connection)).Fill(Datatable);

            if (Datatable.Rows.Count > 0)
            {
                for (int i = 0; i < 15; i++)
                {
                    AttMarks[i].Maximum = 500;
                    AttMarks[i].Minimum = 0;
                    AttMarks[i].Value = Convert.ToDecimal(Datatable.Rows[0][i]);
                }
                label1.Text = string.Format("Количество кредитов по плану = {0}", Datatable.Rows[0][16]);
            }
            else
            {
                for (int i = 0; i < 14; i++)
                {
                    AttMarks[i].Value = 0;
                    AttMarks[i].Maximum = 500;
                    AttMarks[i].Minimum = 0;
                }
                label1.Text = string.Format("Нет сведений о кредитах по предмету (следует обратиться к секретарю для установки сведений)");
            }

            AttMarks[4].Maximum = Convert.ToDecimal(Datatable.Rows[0][15]);
            AttMarks[9].Maximum = Convert.ToDecimal(Datatable.Rows[0][15]);
            AttMarks[14].Maximum = Convert.ToDecimal(Datatable.Rows[0][15]);

            AttMarks[0].Maximum = AttMarks[1].Value - 1;
            AttMarks[1].Maximum = AttMarks[2].Value - 1;
            AttMarks[2].Maximum = AttMarks[3].Value - 1;
            AttMarks[3].Maximum = AttMarks[4].Value - 1;

            AttMarks[5].Maximum = AttMarks[6].Value - 1;
            AttMarks[6].Maximum = AttMarks[7].Value - 1;
            AttMarks[7].Maximum = AttMarks[8].Value - 1;
            AttMarks[8].Maximum = AttMarks[9].Value - 1;

            AttMarks[10].Maximum = AttMarks[11].Value - 1;
            AttMarks[11].Maximum = AttMarks[12].Value - 1;
            AttMarks[12].Maximum = AttMarks[13].Value - 1;
            AttMarks[13].Maximum = AttMarks[14].Value - 1;

            AttMarks[14].Minimum = AttMarks[13].Value + 1;
            AttMarks[13].Minimum = AttMarks[12].Value + 1;
            AttMarks[12].Minimum = AttMarks[11].Value + 1;
            AttMarks[11].Minimum = AttMarks[10].Value + 1;

            AttMarks[9].Minimum = AttMarks[8].Value + 1;
            AttMarks[8].Minimum = AttMarks[7].Value + 1;
            AttMarks[7].Minimum = AttMarks[6].Value + 1;
            AttMarks[6].Minimum = AttMarks[5].Value + 1;

            AttMarks[4].Minimum = AttMarks[3].Value + 1;
            AttMarks[3].Minimum = AttMarks[2].Value + 1;
            AttMarks[2].Minimum = AttMarks[1].Value + 1;
            AttMarks[1].Minimum = AttMarks[0].Value + 1;

            AttMarks[0].Focus();

            FillTemaBlocks(); //Вывести тематические блоки предмета

        }

        //заполнить список тематических блоков предмета
        void FillTemaBlocks()
        {
            //показать тематические блоки
            tema_list_listBox.Items.Clear();
            zadanie_list_listBox.Items.Clear();
            student_zadacha_dataGrid.Rows.Clear();
            tema_zadacha_textBox.Text = string.Empty;
            label18.Text = string.Empty;


            string Sql = "select id, tema, tematext from zanyatie where predmet_id = " + id_predmet_in_tree.ToString();
            TemaTable = new DataTable();
            (new SqlDataAdapter(Sql, global_connection)).Fill(TemaTable);

            if (TemaTable.Rows.Count > 0)
            {
                foreach (DataRow dr in TemaTable.Rows)
                {
                    tema_list_listBox.Items.Add(dr[1].ToString() + " - " + dr[2].ToString());
                }
            }

            if (tema_list_listBox.Items.Count > 0)
                tema_list_listBox.SelectedIndex = tema_list_listBox.Items.Count - 1;
        }

        //отображение/скрытие элементов ввода названия тематических блоков предмета
        void ShowHideTemaBlock(bool show)
        {
            label12.Visible = show;
            label19.Visible = show;
            textBox2.Visible = show;
            tema_block_textBox.Visible = show;
            button3.Visible = show;
            button5.Visible = show;
            button1.Visible = !show;
            button2.Visible = !show;
            button4.Visible = !show;
        }

        // ---- управление пределами в редакторах баллов по МРС

        void UpdateMRSBalls(int num)
        {

            string sql = string.Format("update predmet set {0} = {1} where id = {2} ",
                MarksFieldsNames[num], AttMarks[num].Value, id_predmet_in_tree);
            global_command = new SqlCommand(sql, global_connection);
            
            try
            {
                global_command.ExecuteNonQuery();
            }
            catch(Exception exx)
            {
                MessageBox.Show(exx.Message);
                return;
            }
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            numericUpDown2.Minimum = numericUpDown1.Value + 1;
            UpdateMRSBalls(0);
        }

        private void numericUpDown2_ValueChanged(object sender, EventArgs e)
        {
            numericUpDown1.Maximum = numericUpDown2.Value - 1;
            numericUpDown3.Minimum = numericUpDown2.Value + 1;
            UpdateMRSBalls(1);
        }

        private void numericUpDown3_ValueChanged(object sender, EventArgs e)
        {
            numericUpDown2.Maximum = numericUpDown3.Value - 1;
            numericUpDown4.Minimum = numericUpDown3.Value + 1;
            UpdateMRSBalls(2);
        }

        private void numericUpDown4_ValueChanged(object sender, EventArgs e)
        {
            numericUpDown3.Maximum = numericUpDown4.Value - 1;
            numericUpDown5.Minimum = numericUpDown4.Value + 1;
            UpdateMRSBalls(3);
        }

        private void numericUpDown5_ValueChanged(object sender, EventArgs e)
        {
            numericUpDown4.Maximum = numericUpDown5.Value - 1;
            UpdateMRSBalls(4);
        }

        private void numericUpDown6_ValueChanged(object sender, EventArgs e)
        {
            numericUpDown7.Minimum = numericUpDown6.Value + 1;
            UpdateMRSBalls(5);
        }

        private void numericUpDown7_ValueChanged(object sender, EventArgs e)
        {
            numericUpDown6.Maximum = numericUpDown7.Value - 1;
            numericUpDown8.Minimum = numericUpDown7.Value + 1;
            UpdateMRSBalls(6);
        }

        private void numericUpDown8_ValueChanged(object sender, EventArgs e)
        {
            numericUpDown7.Maximum = numericUpDown8.Value - 1;
            numericUpDown9.Minimum = numericUpDown8.Value + 1;
            UpdateMRSBalls(7);
        }

        private void numericUpDown9_ValueChanged(object sender, EventArgs e)
        {
            numericUpDown8.Maximum = numericUpDown9.Value - 1;
            numericUpDown10.Minimum = numericUpDown9.Value + 1;
            UpdateMRSBalls(8);
        }

        private void numericUpDown10_ValueChanged(object sender, EventArgs e)
        {
            numericUpDown9.Maximum = numericUpDown10.Value - 1;
            UpdateMRSBalls(9);
        }

        private void numericUpDown11_ValueChanged(object sender, EventArgs e)
        {
            numericUpDown12.Minimum = numericUpDown11.Value + 1;
            UpdateMRSBalls(10);
        }

        private void numericUpDown12_ValueChanged(object sender, EventArgs e)
        {
            numericUpDown11.Maximum = numericUpDown12.Value - 1;
            numericUpDown13.Minimum = numericUpDown12.Value + 1;
            UpdateMRSBalls(11);
        }

        private void numericUpDown13_ValueChanged(object sender, EventArgs e)
        {
            numericUpDown12.Maximum = numericUpDown13.Value - 1;
            numericUpDown14.Minimum = numericUpDown13.Value + 1;
            UpdateMRSBalls(12);
        }

        private void numericUpDown14_ValueChanged(object sender, EventArgs e)
        {
            numericUpDown13.Maximum = numericUpDown14.Value - 1;
            numericUpDown15.Minimum = numericUpDown14.Value + 1;
            UpdateMRSBalls(13);
        }

        private void numericUpDown15_ValueChanged(object sender, EventArgs e)
        {
            numericUpDown14.Maximum = numericUpDown15.Value - 1;
            UpdateMRSBalls(14);
        }

        // ----------------------

        private void teacher_tab_predmet_MRS_Click(object sender, EventArgs e)
        {

        }

        // добавление нового тематического блока в предмет
        private void button1_Click(object sender, EventArgs e)
        {
            //if (tema_list_listBox.SelectedIndex == -1) return;
            ShowHideTemaBlock(true);
            tema_block_textBox.Text = "";
            tema_block_textBox.Tag = 0;
            textBox2.Text = string.Empty;
            tema_block_textBox.Focus();

        }

        // редактирование тематического блока в предмете
        private void button4_Click(object sender, EventArgs e)
        {
            if (tema_list_listBox.SelectedIndex == -1) return;

            tema_block_textBox.Tag = 1;
            ShowHideTemaBlock(true);

            if (tema_list_listBox.Items.Count == 0) return;

            int index = tema_list_listBox.SelectedIndex;

            tema_block_textBox.Text = TemaTable.Rows[index][2].ToString();
                //tema_list_listBox.Items[index].ToString();
            textBox2.Text = TemaTable.Rows[index][1].ToString();
            tema_block_textBox.Focus();

        }

        //скрыть блоки редактирования тематических блоков
        private void button5_Click(object sender, EventArgs e)
        {
            ShowHideTemaBlock(false);
        }

        /// <summary>
        /// сохраняет изменения тематического блока предмета
        /// </summary>
        /// <param name="update">признак обновления=true (или вставки=false)</param>
        /// <returns>результат выполнения операции</returns>
        bool SaveTemaBlock(bool update)
        {
            bool res = true;
            string sql = string.Empty;

            if (tema_block_textBox.Text.Trim().Length == 0)
            {                
                return false;
            }

            //сформировать запрос
            if (update)
            {
                //обновление
                int ind = tema_list_listBox.SelectedIndex;
                sql = "update zanyatie set tema=@TEMA, tematext=@TEMATXT where id = @ID";
                global_command = new SqlCommand(sql, global_connection);
                global_command.Parameters.Add("@TEMA", SqlDbType.NVarChar).Value = textBox2.Text;
                global_command.Parameters.Add("@ID", SqlDbType.Int).Value = TemaTable.Rows[ind][0];
                global_command.Parameters.Add("@TEMATXT", SqlDbType.NVarChar).Value = tema_block_textBox.Text;
            }
            else
            {
                //вставка
                sql = "insert into zanyatie (tema, predmet_id, tematext) values (@TEMA, @ID, @TEMATXT)";
                global_command = new SqlCommand(sql, global_connection);
                global_command.Parameters.Add("@TEMA", SqlDbType.NVarChar).Value = textBox2.Text;
                global_command.Parameters.Add("@ID", SqlDbType.Int).Value = tema_list_listBox.Tag;
                global_command.Parameters.Add("@TEMATXT", SqlDbType.NVarChar).Value = tema_block_textBox.Text;
            }

            try
            {
                global_command.ExecuteNonQuery();
            }
            catch (Exception exx)
            {
                res = false;
                MessageBox.Show(exx.Message);
            }

            return res;
        }

        // кнопка сохранения изменения в тематических блоках предмета
        private void button3_Click(object sender, EventArgs e)
        {
            bool update = Convert.ToInt32(tema_block_textBox.Tag) == 1 ? true : false;

            if (!SaveTemaBlock(update))
            {
                if (tema_block_textBox.Text.Trim().Length == 0)
                {
                    MessageBox.Show("Не было введено название тематического блока. Повторите операцию снова.",
                        "Отказ операции", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                else
                    MessageBox.Show("Ошибка во время выполнения (сбой компьютерной сети). Повторите команду снова",
                        "Отказ операции", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            ShowHideTemaBlock(false);
            FillTemaBlocks();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (tema_list_listBox.Items.Count == 0) return;
            if (tema_list_listBox.SelectedIndex == -1) return;

            int ind = tema_list_listBox.SelectedIndex;
            string sql = string.Format("select id from zadanie where zanyatie_id = {0}",
                TemaTable.Rows[ind][0]);

            DataTable tmptable = new DataTable();
            (new SqlDataAdapter(sql, global_connection)).Fill(tmptable);

            if (tmptable.Rows.Count > 0)
            {
                MessageBox.Show("Удаление тематического блока невозможно, так как в нём содержатся задания.",
                    "Отказ операции", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else
            {
                sql = "delete from zanyatie where id = @ID";
                global_command = new SqlCommand(sql, global_connection);
                global_command.Parameters.Add("@ID", SqlDbType.Int).Value = TemaTable.Rows[ind][0];
                try
                {
                    global_command.ExecuteNonQuery();
                }
                catch (Exception exx)
                {
                    MessageBox.Show(exx.Message);
                }
            }

            FillTemaBlocks();
            tmptable.Dispose();
            GC.Collect();

        }

        //отображение/скрытие элементов ввода заданий в блоке
        void ShowHideZadanieEdit(bool show)
        {
            label14.Visible = show;
            label15.Visible = show;
            label16.Visible = show;
            label17.Visible = show;

            tema_textBox.Visible = show;
            ball_zadacha_numericUpDown.Visible = show;
            zad_type_comboBox.Visible = show;
            zadacha_textbox.Visible = show;            

            button9.Visible = show;
            button10.Visible = show;

            button6.Visible = !show;
            button7.Visible = !show;
            button8.Visible = !show;
            tema_zadacha_textBox.Visible = !show;
        }


        // заполнить список заданий в рамках тематического блока
        DataTable ZadachaTable;
        DataTable ZadachaVidTable;
        void FillZadanieList()
        {
            zadanie_list_listBox.Items.Clear();
            tema_zadacha_textBox.Text = string.Empty;
            student_zadacha_dataGrid.Rows.Clear();
            label18.Text = string.Empty;
            BallPlanLabel.Text = string.Empty;

            if (tema_list_listBox.Items.Count == 0) return;

            string sql = string.Empty;
            int ind = tema_list_listBox.SelectedIndex;

            if (tema_list_listBox.SelectedIndex == -1) ind = 0;

            // получить список задач и их свойства
            sql = "SELECT name, predmet_id, isnull(description,'нет описания'), " +  //0 1 2
                " isnull(text,'нет текста задания'), isnull(zanyatie_id,-1), isnull(ball,1), id, type_id" +   // 3 4 5 6 7
                " FROM zadanie WHERE zanyatie_id = " + TemaTable.Rows[ind][0].ToString();            

            ZadachaTable = new DataTable();
            global_adapter = new SqlDataAdapter(sql, global_connection);
            global_adapter.Fill(ZadachaTable);
            foreach (DataRow dr in ZadachaTable.Rows)
            {
                zadanie_list_listBox.Items.Add(
                    string.Format("{0} - [баллов:{1}]", dr[0], dr[5]));
            }
            if (zadanie_list_listBox.Items.Count > 0)
            {
                if (redakt <= 0)
                    zadanie_list_listBox.SelectedIndex = zadanie_list_listBox.Items.Count - 1;
                else
                    zadanie_list_listBox.SelectedIndex = redakt;
                tema_zadacha_textBox.Text = ZadachaTable.Rows[zadanie_list_listBox.Items.Count - 1][3].ToString();
            }

            //заполнить список видов детяльности
            zad_type_comboBox.Items.Clear();
            sql = "select id, name from zadanie_type order by id";
            ZadachaVidTable = new DataTable();
            global_adapter = new SqlDataAdapter(sql, global_connection);
            global_adapter.Fill(ZadachaVidTable);

            foreach (DataRow dr in ZadachaVidTable.Rows)
            {
                zad_type_comboBox.Items.Add(dr[1]);
            }
            if (zadanie_list_listBox.SelectedIndex != 1)
                zad_type_comboBox.SelectedIndex = Convert.ToInt32(ZadachaVidTable.Rows[0][0]) - 1;
            else
                zad_type_comboBox.SelectedIndex = 0;

            FillStudent_zadacha_dataGrid();
            
        }
        
        //построить список выполнения задания
        void FillStudent_zadacha_dataGrid()
        {
            student_zadacha_dataGrid.Rows.Clear();
            if (ZadachaTable.Rows.Count == 0) return; //Если нет заданий

            int ind = zadanie_list_listBox.SelectedIndex == -1 ? 0 : zadanie_list_listBox.SelectedIndex;
            if (zadanie_list_listBox.SelectedIndex == -1) zadanie_list_listBox.SelectedIndex = 0;

            string sql = string.Format("SELECT " +
                " student_zadanie.id, student.id, dbo.GetStudentFIOByID(student.id), " +  // 0 1 2
                " isnull(student_zadanie.vipolnenie,0), " +  // 3
                " isnull(student_zadanie.data_otm, getdate()) " +  // 4
                " FROM         student_zadanie " +
                " INNER JOIN zadanie ON student_zadanie.zadanie_id = zadanie.id " +
                " INNER JOIN student ON student_zadanie.student_id = student.id " +
                " WHERE     (student_zadanie.zadanie_id = {0}) Order by fam, im, ot", 
                ZadachaTable.Rows[ind][6]);

            DataTable dt = new DataTable();
            (new SqlDataAdapter(sql, global_connection)).Fill(dt);

            foreach (DataRow dr in dt.Rows)
            {
                student_zadacha_dataGrid.Rows.Add(new object[] { dr[2], dr[3], dr[4], dr[0] });
            }
            dt.Dispose();
        }

        /// <summary>
        /// сохранить задание в тематическом блоке
        /// </summary>
        /// <param name="update">признак обновления (true) или вставки (false)</param>
        /// <returns>Результат выполнения операции (0=норм, 1=ошибка сети, 2=ошибка данных)</returns>
        int SaveZadanie(bool update)
        {
            int res = 0;
            string sql = string.Empty;

            if (tema_textBox.Text.Trim().Length == 0)
            {
                MessageBox.Show("Введите название задания.", "Отказ операции",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 2;
            }

            if (ball_zadacha_numericUpDown.Value == 0)
            {
                MessageBox.Show("Введите ненулевое количество баллов за выполнение задания.",
                    "Отказ операции",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 2;
            }

            if (update) //обновление
            {
                int k = 0;
                int ind = zadanie_list_listBox.SelectedIndex;
                for (k = 0; k < zadanie_list_listBox.Items.Count; k++)
                {
                    if (k == ind) continue;

                    if (zadanie_list_listBox.Items[k].ToString().Trim() == tema_textBox.Text.Trim())
                    {
                        MessageBox.Show("Задание с таким названием уже имеется в списке заданий. \n" +
                            "Измените название и повторите операцию.",
                            "Отказ операции",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return 2;
                    }
                }

                sql = "UPDATE zadanie " +
                    " SET type_id = @TYPE, text = @TXT, ball = @BLL, name = @NM" +
                    " WHERE (id = @ID)";
                SqlCommand cmd = new SqlCommand(sql, global_connection);
                cmd.Parameters.Add("@TYPE", SqlDbType.Int).Value =
                    ZadachaVidTable.Rows[zad_type_comboBox.SelectedIndex][0];
                cmd.Parameters.Add("@TXT", SqlDbType.NVarChar).Value = zadacha_textbox.Text.Trim();
                cmd.Parameters.Add("@BLL", SqlDbType.Float).Value = ball_zadacha_numericUpDown.Value;
                cmd.Parameters.Add("@NM", SqlDbType.NVarChar).Value = tema_textBox.Text.Trim();
                cmd.Parameters.Add("@ID", SqlDbType.Int).Value =
                    ZadachaTable.Rows[zadanie_list_listBox.SelectedIndex][6];
                try
                {
                    cmd.ExecuteNonQuery();
                }
                catch (Exception rxx)
                {
                    MessageBox.Show("Ошибка при выполнении операции.\n" + rxx.Message,
                        "Невозможно закончить операцию",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return 1;
                }
                cmd.Dispose();
            }
            else //вставка ------------------------------------
            {
                foreach (DataRow dr in ZadachaTable.Rows)
                {
                    if (dr[0].ToString().Trim() == tema_textBox.Text.Trim())
                    {
                        MessageBox.Show("Задание с таким названием уже имеется в списке заданий. \n" +
                            "Измените название и повторите операцию.",
                            "Отказ операции",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return 2;
                    }
                }

                int ind = tema_list_listBox.SelectedIndex;
                sql = "insert into zadanie (type_id,text,ball,name, predmet_id, zanyatie_id ) " +
                    " values (@TYPE,@TXT,@BLL,@NM, @PRID, @ZID)";
                SqlCommand cmd = new SqlCommand(sql, global_connection);
                cmd.Parameters.Add("@TYPE", SqlDbType.Int).Value =
                    ZadachaVidTable.Rows[zad_type_comboBox.SelectedIndex][0];
                cmd.Parameters.Add("@TXT", SqlDbType.NVarChar).Value = zadacha_textbox.Text.Trim();
                cmd.Parameters.Add("@BLL", SqlDbType.Float).Value = ball_zadacha_numericUpDown.Value;
                cmd.Parameters.Add("@NM", SqlDbType.NVarChar).Value = tema_textBox.Text.Trim();
                cmd.Parameters.Add("@PRID", SqlDbType.Int).Value = id_predmet_in_tree;
                cmd.Parameters.Add("@ZID", SqlDbType.Int).Value = TemaTable.Rows[ind][0];

                try
                {
                    cmd.ExecuteNonQuery();
                }
                catch (Exception rxx)
                {
                    MessageBox.Show("Ошибка при выполнении операции.\n" + rxx.Message,
                        "Невозможно закончить операцию",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return 1;
                }
                cmd.Dispose();

                // ------- получить ид нового пунтка списка заданий ----
                sql = "select id from zadanie where name = @NM and zanyatie_id = @ZID";
                cmd = new SqlCommand(sql, global_connection);
                cmd.Parameters.Add("@NM", SqlDbType.NVarChar).Value = tema_textBox.Text.Trim();
                cmd.Parameters.Add("@ZID", SqlDbType.Int).Value = TemaTable.Rows[ind][0];
                DataTable tmp = new DataTable();
                (new SqlDataAdapter(cmd)).Fill(tmp);
                string new_zadan_id = tmp.Rows[0][0].ToString();
                tmp.Dispose(); // ---------------

                // создать записи в таблице студент-задание
                sql = "if ((select count(id) from student_zadanie where zadanie_id=@ZID)=0) " +
                        " Insert Into student_zadanie  " +
                        " (student_id , zadanie_id, prim, vipolnenie, data_otm ) " +
                        " Select student.id, @ZID, '', 0, getdate() " +
                        "	From student  " +
                        "		Join grupa On grupa.id = student.gr_id  " +
                        "	Where student.actual = 1 and student.status_id = 1  " +
                        "	and student.gr_id = @GRID  " +
                        "	Order By fam, im, ot";
                cmd = new SqlCommand(sql, global_connection);
                cmd.Parameters.Add("@GRID", SqlDbType.Int).Value = Datatable.Rows[0][17];
                cmd.Parameters.Add("@ZID", SqlDbType.Int).Value = new_zadan_id;
                cmd.ExecuteNonQuery();

                // вывести список студентов в таблицу выполнения
                FillStudent_zadacha_dataGrid();

            } // ------------------ конец блока вставки записи в таблицу заданий


            return res;
        }

        void PrintStudZadachaTable()
        {

        }

        // создание нового задания
        private void button8_Click(object sender, EventArgs e)
        {
            if (tema_list_listBox.Items.Count == 0) return;

            redakt = -1;

            ShowHideZadanieEdit(true);

            tema_textBox.Text = "";
            ball_zadacha_numericUpDown.Value = 0;
            zad_type_comboBox.SelectedIndex = 3;
            zadacha_textbox.Text = "";
            zadacha_textbox.Tag = 0;

            tema_textBox.Focus();
        }

        int redakt = 0; // номер пункта списка задач при его редактировании

        private void button6_Click(object sender, EventArgs e)
        {
            zadacha_textbox.Tag = 1;
            if (zadanie_list_listBox.SelectedIndex < 0)
            {
                redakt = -1;
                return;
            }

            int ind = zadanie_list_listBox.SelectedIndex;
            redakt = ind;

            ShowHideZadanieEdit(true);

            ball_zadacha_numericUpDown.Value = Convert.ToDecimal(ZadachaTable.Rows[ind][5]);

            for (int i = 0; i < ZadachaVidTable.Rows.Count; i++)
            {
                if (Convert.ToInt32(ZadachaVidTable.Rows[i][0]) == Convert.ToInt32(ZadachaTable.Rows[ind][7]))
                    zad_type_comboBox.SelectedIndex = i;
            }

            zadacha_textbox.Text = ZadachaTable.Rows[ind][3].ToString();
            tema_textBox.Text = ZadachaTable.Rows[ind][0].ToString();

        }

        private void button9_Click(object sender, EventArgs e)
        {
            ShowHideZadanieEdit(false);
        }

        public void SetPlanMRSBall()
        {
            global_query = string.Format("select SUM(zadanie.ball) from zadanie join zanyatie on zanyatie.id = zadanie.zanyatie_id " +
                "where zanyatie.predmet_id = {0}", id_predmet_in_tree);
            DataTable t = new DataTable();
            (new SqlDataAdapter(global_query, global_connection)).Fill(t);
            BallPlanLabel.Text = "Запланировано баллов: " + t.Rows[0][0].ToString();
            t.Dispose();
        }


        private void button10_Click(object sender, EventArgs e)
        {
            bool update = Convert.ToInt32(zadacha_textbox.Tag) == 1 ? true : false;

            //ShowHideZadanieEdit(false);

            int i = SaveZadanie(update);
            if (i != 2)
                ShowHideZadanieEdit(false);
            else
                ShowHideZadanieEdit(true);

            FillZadanieList();
            SetPlanMRSBall();
        }

        // заполнить список заданий в выбранном тематическом блоке
        private void tema_list_listBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillZadanieList();
            if (tema_list_listBox.SelectedIndex < 0) { label18.Text = string.Empty; return; }

            label18.Text = tema_list_listBox.Items[tema_list_listBox.SelectedIndex].ToString();
            SetPlanMRSBall();
        }

        // показать текст задания и построить список его выполнения
        private void zadanie_list_listBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (zadanie_list_listBox.SelectedIndex == -1) return;

            int ind = zadanie_list_listBox.SelectedIndex;
            tema_zadacha_textBox.Text = ZadachaTable.Rows[ind][3].ToString();
            FillStudent_zadacha_dataGrid();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (zadanie_list_listBox.Items.Count == 0) return;
            if (zadanie_list_listBox.SelectedIndex == -1) zadanie_list_listBox.SelectedIndex = 0;

            redakt = -1;

            if (student_zadacha_dataGrid.Rows.Count > 0)
            {
                DialogResult dr = MessageBox.Show("Внимание!\n\nБудет удалено задание и информация об отметках ег овыполнения студентами!!\nПродолжить?",
                    "Запрос на продолжение",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);
                if (dr == DialogResult.No) return;
            }

            // выполнить удаление из таблицы студент-задание            
            (new SqlCommand("delete from student_zadanie where zadanie_id = " +
                ZadachaTable.Rows[zadanie_list_listBox.SelectedIndex][6].ToString(),
                global_connection)).ExecuteNonQuery();
            student_zadacha_dataGrid.Rows.Clear();

            //выполнить удаление задания
            (new SqlCommand("delete from zadanie where id = " +
                ZadachaTable.Rows[zadanie_list_listBox.SelectedIndex][6].ToString(),
                global_connection)).ExecuteNonQuery();

            FillZadanieList();
        }

        /// <summary>
        /// текущее значение в ячейке перед ее редактированием
        /// </summary>
        public string currentval1 = string.Empty;

        /// <summary>
        /// начало редактирования балла в ячейке
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MRSGridView_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            if (e.RowIndex == 0) e.Cancel = true;
            string tmpstr = MRSGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
            currentval1 = tmpstr;
        }

        /// <summary>
        /// выставить оценку выполнения задания в общей таблице стдунтов-заданий
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MRSGridView_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            // ид_студента, ид_задания            
            if (e.ColumnIndex <= 1) return;
            if (e.RowIndex <= 0) return;

            string stud_id = mrsTableZan.Rows[e.RowIndex+2][4].ToString();
            string zadan_id = MRSGridView.Columns[e.ColumnIndex].Tag.ToString();
            string newvalue = "";

            if (MRSGridView.Rows[e.RowIndex].Cells[e.ColumnIndex] != null)
            {
                if (MRSGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].Value != null)
                    newvalue = MRSGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
                else
                    newvalue = "0,0";
            }
            else
            {
                newvalue = "0,0";
            }

            double d = 0.0;
            if (!double.TryParse(newvalue, out d))
            {
                MRSGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = currentval1;
                return;
            }

            newvalue = string.Format("{0:F1}", d);
            MRSGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = newvalue;

            string sql = "update student_zadanie set vipolnenie = @VPLN, data_otm = @DTOTM  where " + 
                " student_id = @STID and zadanie_id = @ZadID";
            SqlCommand cmd = new SqlCommand(sql, global_connection);
            cmd.Parameters.Add("@VPLN", SqlDbType.Float).Value = newvalue;
            cmd.Parameters.Add("@DTOTM", SqlDbType.DateTime).Value = DateTime.Now;
            cmd.Parameters.Add("@STID", SqlDbType.Int).Value = stud_id;
            cmd.Parameters.Add("@ZadID", SqlDbType.Int).Value = zadan_id;
            cmd.ExecuteNonQuery();

            if (d == 0)
            {
                MRSGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.FromArgb(240, 240, 240);
            }
            else
            {
                MRSGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.LightYellow;
            }

            //перечистать сумму баллов по строке
            double sum = 0.0;
            for (int i = 2; i < MRSGridView.Columns.Count; i++)
            {
                sum += Convert.ToDouble(MRSGridView.Rows[e.RowIndex].Cells[i].Value)*
                    Convert.ToDouble(MRSGridView.Rows[0].Cells[i].Value);
            }
            MRSGridView.Rows[e.RowIndex].Cells[1].Value = sum;            
        }

        // сохранить выполнение для студента по данному заданию
        string currentval = "";
        private void student_zadacha_dataGrid_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex != 1) return;
            
            string stud_zan_id = student_zadacha_dataGrid.Rows[e.RowIndex].Cells[3].Value.ToString();
            string newvalue = "";

            if (student_zadacha_dataGrid.Rows[e.RowIndex].Cells[1] != null)
            {
                if (student_zadacha_dataGrid.Rows[e.RowIndex].Cells[1].Value != null)
                    newvalue = student_zadacha_dataGrid.Rows[e.RowIndex].Cells[1].Value.ToString();
                else
                    newvalue = "0,0";
            }
            else
            {
                newvalue = "0,0";
            }

            double d = 0.0;
            if (!double.TryParse(newvalue, out d))
            {
                student_zadacha_dataGrid.Rows[e.RowIndex].Cells[1].Value = currentval;
                return;
            }

            newvalue = string.Format("{0:F1}", d);
            student_zadacha_dataGrid.Rows[e.RowIndex].Cells[1].Value = newvalue;
            student_zadacha_dataGrid.Rows[e.RowIndex].Cells[2].Value = DateTime.Now;

            string sql = "update student_zadanie set vipolnenie = @VPLN, data_otm = @DTOTM  where id = @ID";
            SqlCommand cmd = new SqlCommand(sql, global_connection);
            cmd.Parameters.Add("@VPLN", SqlDbType.Float).Value = newvalue;
            cmd.Parameters.Add("@DTOTM", SqlDbType.DateTime).Value = DateTime.Now;
            cmd.Parameters.Add("@ID", SqlDbType.Int).Value = stud_zan_id;
            cmd.ExecuteNonQuery();
        }

        private void student_zadacha_dataGrid_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            currentval = student_zadacha_dataGrid.Rows[e.RowIndex].Cells[1].Value.ToString();
        }

        /// <summary>
        /// список тем для вкладки итогов
        /// </summary>
        DataTable stripTemaList = null;

        /// <summary>
        /// таблица отчетности МРС по занятию
        /// </summary>
        DataTable mrsTableZan = null;

        /// <summary>
        /// построение таблицы для отображения выполения МРС студентами
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void teacher_tab_predmet_Selected(object sender, TabControlEventArgs e)
        {
            if (e.TabPage != teacher_tab_predmet_MRS_fakt) return;

            temaBlockStripComboBox.Items.Clear();
            MRSGridView.Rows.Clear();
            while (MRSGridView.Columns.Count > 2)
            {
                MRSGridView.Columns.RemoveAt(2);
            }

            // заполнить список занятий выбранного предмета
            string Sql = "select id, tema, tematext from zanyatie where predmet_id = " + id_predmet_in_tree.ToString();
            stripTemaList = new DataTable();
            (new SqlDataAdapter(Sql, global_connection)).Fill(stripTemaList);

            if (stripTemaList.Rows.Count > 0)
            {
                foreach (DataRow dr in TemaTable.Rows)
                {
                    temaBlockStripComboBox.Items.Add(dr[1].ToString() + " - " + dr[2].ToString());
                }
            }

            if (temaBlockStripComboBox.Items.Count > 0)
                temaBlockStripComboBox.SelectedIndex = 0;

        }

        /// <summary>
        /// построение таблицы для отображения выполения МРС студентами
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>        
        private void temaBlockStripComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            int ind = temaBlockStripComboBox.SelectedIndex;

            MRSGridView.Rows.Clear();
            while (MRSGridView.Columns.Count > 2)
            {
                MRSGridView.Columns.RemoveAt(2);
            }

            if (ind < 0) return;

            string zan_id = stripTemaList.Rows[ind][0].ToString();

            global_query = "select predmet.grupa_id from predmet where predmet.id = @PredmID";
            global_command = new SqlCommand(global_query, global_connection);
            global_command.Parameters.Add("@PredmID", SqlDbType.Int).Value = id_predmet_in_tree;
            mrsTableZan = new DataTable();
            (new SqlDataAdapter(global_command)).Fill(mrsTableZan);

            string gr_id = mrsTableZan.Rows[0][0].ToString();

            global_query = "select * from dbo.GetMRS(@ZanID, @GrID) order by id";
            global_command = new SqlCommand(global_query, global_connection);
            global_command.Parameters.Add("@ZanID", SqlDbType.Int).Value = zan_id;
            global_command.Parameters.Add("@GrID", SqlDbType.Int).Value = gr_id;
            mrsTableZan = new DataTable();
            (new SqlDataAdapter(global_command)).Fill(mrsTableZan);

            // получить количество заданий в занятии
            DataRow[] zadanCountRow = mrsTableZan.Select("id=1");
            int colCount = Convert.ToInt32(zadanCountRow[0][2]);

            string[] StrZadanNames = ParseStrForMRS(mrsTableZan.Rows[0][3].ToString());
            string[] StrZadanIds = ParseStrForMRS(mrsTableZan.Rows[2][3].ToString());
            string[] StrZadanBalls = ParseStrForMRS(mrsTableZan.Rows[1][3].ToString());

            MRSGridView.Rows.Add();

            for (int i = 1; i <= colCount; i++)
            {
                DataGridViewTextBoxColumn col = new DataGridViewTextBoxColumn();
                col.ValueType = Type.GetType("double");
                col.Tag = StrZadanIds[i - 1]; //записать сюда ид задания
                col.HeaderText = StrZadanNames[i - 1];
                col.SortMode = DataGridViewColumnSortMode.NotSortable;
                MRSGridView.Columns.Add(col);
                MRSGridView.Rows[0].Cells[i + 1].Value = StrZadanBalls[i - 1].Replace(".",",");
                MRSGridView.Rows[0].Cells[i + 1].Style.ForeColor = Color.Red;
                MRSGridView.Rows[0].Cells[i + 1].ReadOnly = false;
                MRSGridView.Rows[0].Cells[i + 1].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                MRSGridView.Rows[0].Cells[i + 1].Style.Font =
                    new Font("Tahoma", 9.5f, FontStyle.Bold);
            }


            MRSGridView.Rows[0].Cells[1].Value = string.Format("{0:F1}", mrsTableZan.Rows[1][2]);

            int k = 1;
            for (int j = 3; j < mrsTableZan.Rows.Count; j++)
            {
                MRSGridView.Rows.Add();
                MRSGridView.Rows[k].Cells[0].Value = mrsTableZan.Rows[j][1].ToString();
                MRSGridView.Rows[k].Cells[1].Value = mrsTableZan.Rows[j][2].ToString();

                string[] StrStudBalls = ParseStrForMRS(mrsTableZan.Rows[j][3].ToString());
                for (int l = 0; l < colCount; l++)
                {
                    StrStudBalls[l] = StrStudBalls[l].Replace('.', ',');
                    MRSGridView.Rows[k].Cells[l + 2].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    MRSGridView.Rows[k].Cells[l + 2].Style.Format = "N1";
                    if (double.Parse(StrStudBalls[l]) == 0)
                        MRSGridView.Rows[k].Cells[l + 2].Style.BackColor = Color.FromArgb(240, 240, 240);
                    MRSGridView.Rows[k].Cells[l + 2].Value = string.Format("{0:F1}", Convert.ToDouble(StrStudBalls[l]));

                }

                k++;
            }
        }

        string[] ParseStrForMRS(string InStr)
        {
            string[] res;
            char[] separ = { '|' };
            res = InStr.Split(separ);
            return res;
        }

        // ------------------ кон - отметки по МБРС --------------------------

        /// <summary>
        /// вывод выдачи по предмету в таблицу резельтата
        /// </summary>
        /// <param name="id">ид предмета</param>
        /// <param name="pr_name">название предмета</param>
        public void select_predmet_reaction(int id, string pr_name)
        {
            result_pane.Clear();
            result_pane.Cols.DefaultSize = 40;
            result_pane.Cols.Count = 2;
            result_pane.Styles.Fixed.TextAlign = C1.Win.C1FlexGrid.TextAlignEnum.LeftCenter;
            result_pane.Styles.Normal.TextAlign = C1.Win.C1FlexGrid.TextAlignEnum.CenterCenter;           
            
            result_pane[0, 0] = pr_name;
            teacher_tab_predmet_statistica.Text = "Статистика выдачи по: " + pr_name; 

            if (teacher_tab_predmet.TabPages.Contains(teacher_tab_predmet_kursrab))
                teacher_tab_predmet.TabPages.Remove(teacher_tab_predmet_kursrab);

            if (teacher_tab_predmet.TabPages.Contains(teacher_tab_predmet_zachet))
                teacher_tab_predmet.TabPages.Remove(teacher_tab_predmet_zachet);

            if (teacher_tab_predmet.TabPages.Contains(teacher_tab_predmet_exam))
                teacher_tab_predmet.TabPages.Remove(teacher_tab_predmet_exam);

            if (teacher_tab_predmet.TabPages.Contains(teacher_tab_predmet_kontrrab))
                teacher_tab_predmet.TabPages.Remove(teacher_tab_predmet_kontrrab);

            if (teacher_tab_predmet.TabPages.Contains(teacher_tab_predmet_MRS))
                teacher_tab_predmet.TabPages.Remove(teacher_tab_predmet_MRS);

            if (teacher_tab_predmet.TabPages.Contains(teacher_tab_predmet_MRS_fakt))
                teacher_tab_predmet.TabPages.Remove(teacher_tab_predmet_MRS_fakt);
           
            global_query = string.Format("select tree_name, vid_zan.id, vid_zan.kod, vid_zan.name  from vid_zan " +
                " join vidzan_predmet on vidzan_predmet.vidzan_id = vid_zan.id " +
                " where show_in_tree=1 and vidzan_predmet.predmet_id={0} " +
                " order by tree_name", id);

            dtt = new DataTable();
            global_adapter = new SqlDataAdapter(global_query, global_connection);
            global_adapter.Fill(dtt);

            foreach (DataRow dr in dtt.Rows)
            {
                string kod = dr[2].ToString();

                switch (kod)
                {
                    case "э":
                        teacher_tab_predmet.TabPages.Add(teacher_tab_predmet_exam); break;
                    case "з": 
                        teacher_tab_predmet.TabPages.Add(teacher_tab_predmet_zachet); break;
                    case "дз":
                        teacher_tab_predmet.TabPages.Add(teacher_tab_predmet_zachet); break;
                    case "зкр":
                        teacher_tab_predmet.TabPages.Add(teacher_tab_predmet_kursrab); break;
                    case "кнр":                        
                        teacher_tab_predmet.TabPages.Add(teacher_tab_predmet_kontrrab); break;
                }                                
            }


            global_query = "select " +
                " vid_zan.koef,  " + //0
                " vid_zan.name,  " + //1
                " vid_zan.id,  " +  //2
                " vid_zan.krat_name,  " +  //3
                " vid_zan.delenie, " +  //4
                " vidzan_predmet.kol_chas,  " + //5
                " vidspisan = vid_zan.spisanie,  " + //6
                " vid_zan.out_type,  " +                //7                 
                " predmetspisan = predmet_type.spisanie, " + //8
                " fakt_text, plan_text, vid_zan.name, " +  //9, 10, 11
                " predmet.semestr, vid_zan.kod, vid_zan.is_kontrol, predmet.mrs " + //12, 13, 14, 15
                " from vidzan_predmet " +
                " join vid_zan on vid_zan.id = vidzan_predmet.vidzan_id " +
                " join predmet on predmet.id=vidzan_predmet.predmet_id " +
                " join predmet_type on predmet_type.id = predmet.type_id " +
                " where vidzan_predmet.predmet_id =  " + id.ToString() +
                " and show_in_grid=1 " +
                " order by vid_zan.id";

            DataTable nado_set = new DataTable();
            global_adapter = new SqlDataAdapter(global_query, global_connection);
            global_adapter.Fill(nado_set);           
            result_pane.Rows.Count = nado_set.Rows.Count * 5 + 1;

            // показать вкладку МРС и вывести сведения по МБРС
            if (Convert.ToInt32(nado_set.Rows[0][15])>0)
            {
                teacher_tab_predmet.TabPages.Add(teacher_tab_predmet_MRS);
                teacher_tab_predmet.TabPages.Add(teacher_tab_predmet_MRS_fakt);                
                PrintMBRS(id, pr_name);
            }

            //получить даты и часы по выбранному типу предмета
            int predm_semestr = (int)nado_set.Rows[0][12];

            string date_filter = "";

            int sy = semestr2_start.Year,
                sd = semestr2_start.Day,
                sm = semestr2_start.Month;            

            if (predm_semestr % 2 != 0)
            {
                date_filter = string.Format(
                    " dbo.get_date(y,m,d)>=dbo.get_date({0},9,1)and dbo.get_date(y,m,d)<=dbo.get_date({1},{2},{3}) ", 
                    starts[0].Year, sy, sm, sd);
            }
            else
            {
                date_filter = string.Format(" y={0} ", ends[ends.Count - 1].Year);
            }

            //получить часы по всем видам занятий из расписания
            global_query = "select vid_zan_id, kol_chas, d, m, rasp.id as rid, vid_zan.kod as vzk, nom_zan, koef from rasp " +
                " join vid_zan on vid_zan.id=rasp.vid_zan_id " +
                " where predmet_id = " + id.ToString() +
                " and " + date_filter +                 
                " and prepod_id = " + active_user_id.ToString();
            DataTable chas_set = new DataTable();
            global_adapter = new SqlDataAdapter(global_query, global_connection);
            global_adapter.Fill(chas_set);

            /*string rr = "";
            foreach (DataRow ddd in chas_set.Rows)
            {
                rr += ddd["vzk"].ToString() + " - " + ddd["rid"].ToString() +  "\n";
            }
            MessageBox.Show(rr);*/
            
            // ---------------- цикл вывода информации о предмете и его часах ------------------
            result_pane.Cols[0].Width = 200;
            result_pane.Cols[1].Width = 60;

            //MessageBox.Show("select_predmet_reaction вызвана - получено - " + chas_set.Rows.Count.ToString());

            int i = 1;
            string zanid = "-1";
            double zankoef = 0.0;
            foreach (DataRow dr in nado_set.Rows)
            {
                string name = dr[1].ToString();
                
                result_pane.Rows[i].IsNode = true;
                result_pane.Rows[i].Node.Level = 2;
                result_pane.Rows[i].StyleNew.BackColor = Color.AliceBlue;
                result_pane.Rows[i].StyleNew.TextEffect = C1.Win.C1FlexGrid.TextEffectEnum.Raised;
                result_pane.Rows[i].StyleNew.Border.Color = Color.AliceBlue;                
                result_pane[i, 0] = name;

                int nado = 0; //рассчитать план по видам занятий
                bool надо_списывать = Convert.ToBoolean(dr[8]); //надо ли списывать часы с предмета (если не спец семинар)

                result_pane[++i, 0] = "По плану:";

                if (!надо_списывать) //если предмет не предусматривает списание (это спец семинар или спец курс)
                {
                    nado = Convert.ToInt32(dr[5]);
                    result_pane[i, 1] = nado.ToString();
                }
                else
                {
                    int tmp = Convert.ToInt32(dr[5]);
                    int процент_списания = Convert.ToInt32(dr[6]);

                    if (процент_списания>0)
                    {
                        nado = tmp - (tmp * процент_списания) / 100; ///расчет списания здесь <<<<<<<<---------------
                    }
                    else
                    {
                        nado = tmp;
                    }

                    result_pane[i, 1] = nado.ToString();

                }

                int faktrow = ++i;
                result_pane[faktrow, 0] = "Фактически:";

                int datesrow = ++i;
                result_pane[datesrow, 0] = "Даты проведения:";

                int chasrow = ++i;
                result_pane[chasrow, 0] = "Кол-во часов:";
                i++;

                //получить даты и часы по даному виду занятий 
                string kod = dr[13].ToString();
                string v_id = dr[2].ToString();
                bool контрольное_занятие = Convert.ToBoolean(dr[14]);
                DataRow[] chasrows;

                if (!контрольное_занятие)
                {
                    chasrows = chas_set.Select(" vid_zan_id = " + v_id);
                    int cols = chasrows.Length;
                    
                    if (result_pane.Cols.Count < cols + 1) result_pane.Cols.Count = cols + 1;

                    int jj = 1;
                    double sum = 0;
                    foreach (DataRow ddr in chasrows)
                    {                        
                        result_pane[datesrow, jj] = string.Format("{0:D2}.{1:D2}",ddr[2], ddr[3]);
                        result_pane[chasrow, jj] = ddr[1].ToString();
                        sum += Convert.ToDouble(ddr[1]);
                        jj++;
                    }

                    result_pane[faktrow, 1] = string.Format("{0:F2}",sum);
                    double percent = 0;

                    if (sum > 0)
                    {
                        percent = (sum * 100.0) / nado;
                    }                    
                    
                    //result_pane[faktrow - 2, 1] = string.Format("{0:F2}%", percent);
                    result_pane.Rows[faktrow - 2].StyleNew.ForeColor = Color.Navy;
                    result_pane.Rows[faktrow - 2].StyleNew.Font = new Font("tahoma", 9, FontStyle.Bold);
                                                            
                }
                else
                {                    
                    //отдельно получить по контрольным, курсовым, зачетам, экзаменам, 
                    chasrows = chas_set.Select(" vid_zan_id = " + v_id,"m, d, nom_zan");
                    int cols = chasrows.Length; 
                   
                     //получить id занятия на которое приходтися зачет экзамен или защита 
                    if (zanid == "-1")
                    {
                        if (chasrows.Length > 0) 
                        {
                            zanid = chasrows[0][4].ToString();
                            zankoef = Convert.ToDouble(chasrows[0][7]);
                        }
                    }

                    if (result_pane.Cols.Count < cols + 1) result_pane.Cols.Count = cols + 1;

                    int jj = 1;                    
                    foreach (DataRow ddr in chasrows)
                    {
                        if (jj > 1) break;
                        result_pane[datesrow, jj] = string.Format("{0:D2}.{1:D2}", ddr[2], ddr[3]);                        
                        jj++;
                    }

                    DataTable session = new DataTable();
                    global_query = string.Format(" select count(*) from session where rasp_id={0}", zanid);
                    global_adapter = new SqlDataAdapter(global_query, global_connection);
                    global_adapter.Fill(session);
                    string full_list = session.Rows[0][0].ToString();

                    if (chasrows.Length > 0)
                    {
                        switch (kod)
                        {
                            case "э":
                            case "дз":                                
                                global_query  = string.Format(" select count(*) from session " + 
                                " join vid_otmetka on vid_otmetka.id = session.otmetka_id " + 
                                " where rasp_id={0} and (vid_otmetka.name = '2' or vid_otmetka.name = '3' or vid_otmetka.name = '4' or vid_otmetka.name = '5')", zanid);
                                break;                            
                            case "з":
                                global_query = string.Format(" select count(*) from session " +
                                " join vid_otmetka on vid_otmetka.id = session.otmetka_id " +
                                " where rasp_id={0} and (vid_otmetka.name = '1')", zanid);
                                break;
                            case "зкр":                                
                                // получить ид курсовой работы
                                session = new DataTable();
                                global_query = string.Format("select id from rabota where predmet_id = {0} and y={1} and vid_rab_id=2", id, semestr2_start.Year);
                                global_adapter = new SqlDataAdapter(global_query, global_connection);
                                global_adapter.Fill(session);
                                string rabota_id = "0";
                                if (session.Rows.Count>0) rabota_id = session.Rows[0][0].ToString();                                

                                //всего курсовых работ
                                session = new DataTable();
                                global_query = string.Format(" select count(isnull(vid_otmetka.name,'')) " + 
                                    " from student_rabota " + 
                                    " join rabota on rabota.id = student_rabota.rabota_id " + 
                                    " join student on student.id = student_rabota.student_id " + 
                                    " left outer join tema_rabota on tema_rabota.id = student_rabota.tema_id " + 
                                    " left outer join vid_otmetka on vid_otmetka.id = student_rabota.otmetka_id " + 
                                    " where rabota.id = {0} and student.actual = 1 and student.status_id=1 ", rabota_id);
                                global_adapter = new SqlDataAdapter(global_query, global_connection);
                                global_adapter.Fill(session);
                                full_list = session.Rows[0][0].ToString();

                                //узнать сколько человек сдали курсовую работу
                                global_query = string.Format(" select count(isnull(vid_otmetka.name,'')) " +
                                    " from student_rabota " +
                                    " join rabota on rabota.id = student_rabota.rabota_id " +
                                    " join student on student.id = student_rabota.student_id " +
                                    " left outer join tema_rabota on tema_rabota.id = student_rabota.tema_id " +
                                    " left outer join vid_otmetka on vid_otmetka.id = student_rabota.otmetka_id " +
                                    " where rabota.id = {0} and student.actual = 1 and student.status_id=1 " +
                                    " and (vid_otmetka.name = '3' or vid_otmetka.name = '4' or vid_otmetka.name = '5')", 
                                    rabota_id);
                                break;
                        }

                        session = new DataTable();
                        global_adapter = new SqlDataAdapter(global_query, global_connection);
                        global_adapter.Fill(session);

                        result_pane[faktrow, 1] = (zankoef*Convert.ToInt32(session.Rows[0][0])).ToString(); //ddr[1].ToString();
                        result_pane[faktrow, 0] = string.Format("Сдало {0} чел. из {1}:", session.Rows[0][0], full_list);
                        if (kod=="зкр" && full_list == "0") 
                            result_pane[chasrow, 0] = "нет данных о курс. раб.";
                        else
                            result_pane[chasrow, 0] = "";
                        result_pane[datesrow, 0] = "Дата проведения:";
                        result_pane[faktrow - 1, 1] = (zankoef * int.Parse(full_list)).ToString();
                    }
                    else
                    {
                        result_pane[chasrow, 1] = "0";
                    }

                    chasrows = null;
                }
            } 
        }
        // -----------------------------------------------------------------------------------------------------------

        //действия при выборе элементов в дереве личных объектов преподавателя

        public static int id_predmet_in_tree = 0; //ид предмета выделенного в дереве
        public static string name_predmet_in_tree = "";
        public static string name_predmet = "";

        private void object_tree_AfterSelect(object sender, TreeViewEventArgs e)
        {
            //Если выбраны предметы
            if (e.Node.Name == "predmets")
            {
                if (e.Node.Nodes.Count == 0) return;

                if (!teacher_tab.TabPages.Contains(teacher_predmet))
                    teacher_tab.TabPages.Add(teacher_predmet);
                
                teacher_predmet.Text = "Предметы преподавателя: " + e.Node.Parent.Text;
                teacher_tab.SelectedIndex = 2;
                teacher_tab_predmet.SelectedIndex = 0;
                content.SelectedIndex = 1;
                
                result_pane.Rows[0].IsNode = true;
                result_pane.Rows[0].Node.Level = 1;

                id_predmet_in_tree = (int)e.Node.Nodes[0].Tag;
                name_predmet_in_tree = e.Node.Nodes[0].Text;
                select_predmet_reaction(id_predmet_in_tree, e.Node.Nodes[0].Text);
            }

            int prid = 0;

            //выбран конкретный предмет
            if (e.Node.Name.StartsWith("pr_"))
            {
                teacher_tab.SelectedIndex = 2;
                teacher_tab_predmet.SelectedIndex = 0;

                prid = (int)e.Node.Tag;
                
                result_pane.Rows[0].IsNode = true;
                result_pane.Rows[0].Node.Level = 1;

                id_predmet_in_tree = prid;
                name_predmet_in_tree = e.Node.Text;                
                select_predmet_reaction(prid,e.Node.Text);

                DataTable pr = new DataTable();
                (new SqlDataAdapter(
                    "select name from predmet where id = " + id_predmet_in_tree.ToString(), global_connection)).Fill(pr);
                name_predmet = pr.Rows[0][0].ToString();
                GC.Collect();
            }

            //Если выбрано расписание
            if (e.Node.Name == "k_list")
            {                
                teacher_tab.SelectedIndex = 1;
                content.SelectedIndex = 3;
            }

            if (e.Node.Name == "raspisan")
            {
                teacher_tab.SelectedIndex = 0;
            }

        }

        //вывести статистику выдачи по данному предмету
        public void fill_tab_statistica(int id)
        {

        }

        /// <summary>
        /// выделить препода, предмет, вид занятия и аудиторию выеделнной ячейки
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void select_in_choosebar_Click(object sender, EventArgs e)
        {
            int c = table.Col;
            int r = table.Row;

            if (!is_correct_cell(r, c)) return;

            //определить колонки подгрупп
            int first = 0, second = 0;
            if (c % 2 == 0)
            {
                first = c - 1;
                second = c;
            }
            else
            {
                first = c;
                second = c + 1;
            }            

            int dest_col = first; //колонка назначения
            if (c == second) dest_col = second;
            if (c == first) dest_col = first;

            //определить параметры исходной ячейки
            DateTime d = table_data.RowDate(r);
            int p = table_data.RowPair(r);
            string g = table_data.ColumnGroup(dest_col);
            int sg = table_data.ColumnSubGroup(dest_col);

            //найти номер выделенного препода и перейти в него

            if ((int)grups_set.Rows[grupa_list.SelectedIndex][2] != table_data[d, g, p].grupa_id[sg - 1])
            {
                return;
            }


            int prepod_index = -1;
            int prepod_id = (int)table_data[d, g, p].prepod_id[sg - 1];
            int i = 0;


            foreach(DataRow dr in prepod_set.Rows)
            {
                if ((int)dr[0] == prepod_id)
                {
                    prepod_index = i;
                    break;
                }
                i++;
            }

            if (prepod_index != -1)
                prepod_list.SelectedIndex = prepod_index;
            else
                return;


            //если более одного предмета в списке, то выбрать
            if (predmet_list.Items.Count > 1)
            {
                int predm_index = -1;
                int predm_id = (int)table_data[d, g, p].predmet_id[sg - 1];
                i = 0;

                foreach (DataRow dr in predmet_set.Rows)
                {
                    if ((int)dr[0] == predm_id)
                    {
                        predm_index = i;
                        break;
                    }
                    i++;
                }

                if (predm_index != -1)
                    predmet_list.SelectedIndex = predm_index;
                else
                    return;
            }

            //найти вид занятия
            int vid_index = -1;
            int vid_id = (int)table_data[d, g, p].vid_zan_id[sg - 1];
            i = 0;

            foreach (DataRow dr in vidzan_set.Rows)
            {
                if ((int)dr[2] == vid_id)
                {
                    vid_index = i;
                    break;
                }
                i++;
            }

            if (vid_index != -1)
                vid_zan_list.SelectedIndex = vid_index;
            else
                return;


            //найти аудиторию
            int aud_index = -1;
            int aud_id = (int)table_data[d, g, p].kabinet_id[sg - 1];
            i = 0;

            foreach (DataRow dr in aud_set.Rows)
            {
                if ((int)dr[1] == aud_id)
                {
                    aud_index = i;
                    break;
                }
                i++;
            }

            if (aud_index != -1)
                aud_list.SelectedIndex = aud_index;
            else
                return;

        }


        /// <summary>
        /// оптрвить текст выделенной ячейки в буфер обемена Windows
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripMenuItem1_Click_1(object sender, EventArgs e)
        {
            int c = table.Col;
            int r = table.Row;

            if (!is_correct_cell(r, c)) return;

            //определить колонки подгрупп
            int first = 0, second = 0;
            if (c % 2 == 0)
            {
                first = c - 1;
                second = c;
            }
            else
            {
                first = c;
                second = c + 1;
            }

            int dest_col = first; //колонка назначения
            if (c == second) dest_col = second;
            if (c == first) dest_col = first;

            //определить параметры исходной ячейки
            DateTime d = table_data.RowDate(r);
            int p = table_data.RowPair(r);
            string g = table_data.ColumnGroup(dest_col);
            int sg = table_data.ColumnSubGroup(dest_col);

            string predm = table_data[d, g, p].predmet_name[sg - 1];
            string prep = "/" + table_data[d, g, p].prepod_name[sg - 1] + "/";
            string aud = (table_data[d, g, p].aud_name[sg - 1] == "--") ? "" : ("," + table_data[d, g, p].aud_name[sg - 1]);
            string vidzan = table_data[d, g, p].vid_zan_name[sg - 1];

            string outstr = predm + ", " + vidzan + ", " + prep + aud;

            Clipboard.Clear();
            Clipboard.SetText(outstr, TextDataFormat.Text);
        }

        /// <summary>
        /// сохранить в формате Excel раcписание на неделю для выбранного преподавателя
        /// </summary>
        /// <param name="id">ид преподавателя в БД</param>
        /// <param name="nm">полное имя преподавателя</param>
        public void SaveToExcel(int id, string nm)
        {
            string root = GetMyDocs() + "\\Расписания";
            CreateFolder(root);
            CreateFolder(root + "\\Расписания преподавателей");
            string FileName = root + "\\Расписания преподавателей\\" + nm;
            CreateFolder(FileName);
            //имя файла для сохранения
            string nmshort = nm.Substring(0, nm.IndexOf(" ", 0));
            FileName = FileName + "\\" + nmshort + "_неделя_с_" + starts[week_list.SelectedIndex].ToShortDateString() + ".xls";

            ExcelFile excel = new ExcelFile();

            ExcelWorksheet sheet = excel.Worksheets.Add(
                "Расп. с " + starts[week_list.SelectedIndex].ToShortDateString());

            DateTime st=starts[week_list.SelectedIndex], en = ends[week_list.SelectedIndex];
            int num = 1;

            sheet.PrintOptions.HeaderMargin = 0.0;
            sheet.PrintOptions.FooterMargin = 0.0;
            sheet.PrintOptions.FitToPage = true;
            sheet.PrintOptions.Portrait = false;

            //задать общие свойства свойства
            CellRange cr;
            cr = sheet.Cells.GetSubrange("a1","h9");
            cr.Merged = true;
            cr.Style.Font.Name = "Tahoma";
            cr.Style.Font.Size = 10 * 20;
            //cr.Style.Font.Italic = true;
            cr.Style.Font.Weight = ExcelFont.BoldWeight;
            cr.Style.VerticalAlignment = VerticalAlignmentStyle.Center;
            cr.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            //cr.Style.FillPattern.SetSolid(Color.White);
            cr.Merged = false;

            cr = sheet.Cells.GetSubrange("a3", "h9");
            cr.Merged = true;
            cr.Style.Borders[IndividualBorder.Left].LineStyle =  GemBox.Spreadsheet.LineStyle.Thin;
            cr.Style.Borders[IndividualBorder.Right].LineStyle = GemBox.Spreadsheet.LineStyle.Thin;
            cr.Style.Borders[IndividualBorder.Top].LineStyle = GemBox.Spreadsheet.LineStyle.Thin;
            cr.Style.Borders[IndividualBorder.Bottom].LineStyle = GemBox.Spreadsheet.LineStyle.Thin;
            cr.Merged = false;


            cr = sheet.Cells.GetSubrange("a1", "h2");
            cr.Merged = true;
            cr.Style.Font.Name = "Times New Roman";
            cr.Style.Font.Size = 14 * 20;
            cr.Style.Font.Italic = true;
            cr.Style.Font.Weight = ExcelFont.BoldWeight;
            cr.Style.VerticalAlignment = VerticalAlignmentStyle.Center;
            cr.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            cr.Value = "Расписание c " + starts[week_list.SelectedIndex].ToShortDateString() +
                " по " + ends[week_list.SelectedIndex].ToShortDateString() + 
                " ( преподаватель: " + nm + ", FSystem: " + fakultet_name_krat + ")";

            //cr.Merged = false;

            sheet.DefaultColumnWidth = 16 * 256;            
            sheet.Columns[0].Width = 10 * 256;
            sheet.Rows[2].Height = 45 * 20;

            for (DateTime d = st; d <= en; d = d.AddDays(1))
            {
                sheet.Cells[2, num].Value = d.ToShortDateString() + "\n" + DaysMed[num - 1];
                num++;
            }

            for (int i = 1; i <= 6; i++)
            {
                sheet.Cells[i + 2, 0].Value = i.ToString() + " пара";
                sheet.Rows[i + 2].Height = 45 * 20;
            }


            foreach (DateTime dt in table_data.data.Keys)
            {
                foreach (string gr in table_data.data[dt].Keys)
                {
                    for (int para = 1; para <= 6; para++)
                    {

                        int prep_id0 = table_data[dt, gr, para].prepod_id[0];
                        int prep_id1 = table_data[dt, gr, para].prepod_id[1];

                        int ex_row = para + 2;
                        int ex_col = daynumer(dt);

                        if (!(prep_id0 == 0 && prep_id1 == 0))
                        {
                            if (prep_id0 == id)
                            {
                                sheet.Cells[ex_row, ex_col].Value = table_data[dt, gr, para].predmet_name[0] + "\n" +
                                    gr + ", " + table_data[dt, gr, para].vid_zan_name[0];
                            }
                            else
                            {
                                if (prep_id1 == id)
                                {
                                    sheet.Cells[ex_row, ex_col].Value = table_data[dt, gr, para].predmet_name[1] + "\n" +
                                        gr + ", " + table_data[dt, gr, para].vid_zan_name[1];
                                }
                            }
                        }
                    }
                }
            }


            // -----------------------------------------------------------------------
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

        /// <summary>
        /// создать расписание в формате Excel для всех групп таблицы раписания
        /// </summary>
        public void SaveToExcel()
        {                        
            bool killemptyrows = false; //отбрасывать пустые строки
            List<int> OutGroups = new List<int>(); //список выводимых групп
            List<int> OutLines = new List<int>(); //список номеров строк выводимых дней                      
            bool ShowHeader = true; //показыать стандартную шапку
            string S = ""; //строковый буфер
            
            
            CellRange cr; //диапазон ячеек на рабочем листе книги
            string[] Letters = new string[]{
                "A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z",
                "AA","AB","AC","AD","AE","AF","AG","AH","AI","AJ","AK","AL","AM","AN","AO","AP","AQ","AR","AS","AT","AU","AV","AW","AX","AY","AZ",
                "BA","BB","BC","BD","BE","BF","BG","BH","BI","BJ","BK","BL","BM","BN","BO","BP","BQ","BR","BS","BT","BU","BV","BW","BX","BY","BZ"};


            string root = GetMyDocs() + "\\Расписания";
            CreateFolder(root);
            CreateFolder(root + "\\Расписания по FSystemу");
            string FileName = root + "\\Расписания по FSystemу";
            CreateFolder(FileName);
            //имя файла для сохранения
            FileName = FileName + "\\раписание от " + starts[week_list.SelectedIndex].ToShortDateString() + ".xls";

            ExcelFile excel = new ExcelFile();

            ExcelWorksheet sheet = excel.Worksheets.Add(
                "Расп. с " + starts[week_list.SelectedIndex].ToShortDateString());

            
            //задать общие свойства свойства
            cr = sheet.Cells.GetSubrange("a1", "y50");
            cr.Merged = true;
            cr.Style.Font.Name = "Times New Roman";
            cr.Style.Font.Size = 8 * 20;
            cr.Style.Font.Italic = true;
            cr.Style.VerticalAlignment = VerticalAlignmentStyle.Center;
            cr.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            cr.Style.FillPattern.SetSolid(Color.White);            
            cr.Merged = false;

            
            // вывод шапки расписания
            if (dekan_online && ShowHeader)
            {
                sheet.Cells.GetSubrange("A1", "D4").Merged = true;
                sheet.Cells.GetSubrange("A1", "D4").Style.HorizontalAlignment = HorizontalAlignmentStyle.Left;

                S = "\"УТВЕРЖДАЮ\"\n" +
                    "Декан ____________ " + dekan_name + "\n" +
                    "\"__\" ________________  " + DateTime.Now.Year.ToString() + " г.";

                sheet.Cells["A1"].Value = S;
            }
            
            // ------------------   ===== взять шапку раписания из БД
            cr = sheet.Cells.GetSubrange("E1", Letters[grupa_list.Items.Count*2 - 1] + "4");
            cr.Merged = true;
            cr.Style.Font.Name = "Verdana";
            cr.Style.Font.Weight = ExcelFont.BoldWeight;
            cr.Style.Font.Italic = false;
            cr.Style.Font.Size = 14 * 20;

            S = "Расписание занятий на FSystemе  " + 
                "\"" + fakultet_name + "\"\n" +  
                "c " + starts[week_list.SelectedIndex].ToShortDateString() +
                " по " + ends[week_list.SelectedIndex].ToShortDateString();

            sheet.Cells["E1"].Value = S;


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
                 GemBox.Spreadsheet.LineStyle.Thin);
            cr.Merged = false;

            //sheet.Cells.GetSubrange("A6", "C6").SetBorders(MultipleBorders.Top, Color.Black, GemBox.Spreadsheet.LineStyle.DoubleLine);
            cr = sheet.Cells.GetSubrange("C6", "C" + (OutLines[OutLines.Count - 1] + 5).ToString());
            cr.Merged = true;
            cr.Style.Borders[IndividualBorder.Right].LineStyle = GemBox.Spreadsheet.LineStyle.Double;
            cr.Merged = false;

            cr = sheet.Cells.GetSubrange("A6", "A" + (OutLines[OutLines.Count - 1] + 5).ToString());
            cr.Merged = true;
            cr.Style.Borders[IndividualBorder.Left].LineStyle = GemBox.Spreadsheet.LineStyle.Double;
            cr.Merged = false;

            cr = sheet.Cells.GetSubrange("A" + (OutLines[OutLines.Count - 1] + 5).ToString(), 
                "C" + (OutLines[OutLines.Count - 1] + 5).ToString());
            cr.Merged = true;
            cr.Style.Borders[IndividualBorder.Bottom].LineStyle = GemBox.Spreadsheet.LineStyle.Double;
            cr.Merged = false;

            cr = sheet.Cells.GetSubrange("A6", "C6");
            cr.Merged = true;
            cr.Style.Borders[IndividualBorder.Top].LineStyle = GemBox.Spreadsheet.LineStyle.Double;
            cr.Merged = false;

            cr = sheet.Cells.GetSubrange("A6", "C6");
            cr.Merged = true;
            cr.Style.Borders[IndividualBorder.Bottom].LineStyle = GemBox.Spreadsheet.LineStyle.Double;
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
                cell1 = string.Format("A{0}", num+1);

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

            for (i = 1; i < grups_count*2; i += 2) ///исправить цикл взять количество из Outgroups
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
                cr.Style.Borders.SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.DoubleLine);

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
                                Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
                        cr.Merged = false;

                        cr = sheet.Cells.GetSubrange(cell2, cell2);
                        cr.Merged = true;
                        cr.Style.Borders.SetBorders(MultipleBorders.Outside,
                                Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
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
                            Color.Black, GemBox.Spreadsheet.LineStyle.DoubleLine);

                    }

                    //вывести горизонтальный разделитель
                    sheet.Cells[cell1].SetBorders(MultipleBorders.Bottom, 
                        Color.Black, GemBox.Spreadsheet.LineStyle.DoubleLine);
                    sheet.Cells[cell2].SetBorders(MultipleBorders.Bottom,
                        Color.Black, GemBox.Spreadsheet.LineStyle.DoubleLine);


                    sheet.Cells["A" + (num + ii - 2).ToString()].Style.Borders[IndividualBorder.Bottom].LineStyle =
                        GemBox.Spreadsheet.LineStyle.DoubleLine;
                    
                    cr = sheet.Cells.GetSubrange("B" + (num + ii - 2).ToString(),"B" + (num + ii - 2).ToString());
                    cr.Merged = true;
                    cr.Style.Borders[IndividualBorder.Bottom].LineStyle = GemBox.Spreadsheet.LineStyle.DoubleLine;
                    cr.Merged = false;

                    cr = sheet.Cells.GetSubrange("C" + (num + ii - 2).ToString(), "C" + (num + ii - 2).ToString());
                    cr.Merged = true;
                    cr.Style.Borders[IndividualBorder.Bottom].LineStyle = GemBox.Spreadsheet.LineStyle.DoubleLine;
                    cr.Merged = false;


                    num += 6;
                }

                counter+=2;
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
        
        /// <summary>
        /// создать расписание для курса с указанным номером
        /// </summary>
        /// <param name="kurs">номер курса</param>
        public void SaveToExcel(int kurs)
        {
            bool killemptyrows = false; //отбрасывать пустые строки
            List<int> OutGroups = new List<int>(); //список выводимых групп
            List<int> OutLines = new List<int>(); //список номеров строк выводимых дней                      
            bool ShowHeader = true; //показыать стандартную шапку
            string S = ""; //строковый буфер
            CellRange cr;
            string[] Letters = new string[]{"A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T",
                "U","V","W","X","Y","Z","AA","AB"};

            string root = GetMyDocs() + "\\Расписания";
            CreateFolder(root);
            CreateFolder(root + "\\Расписания по курсам");
            string FileName = root + "\\Расписания по курсам\\" + kurs.ToString() + " курс"; 
            CreateFolder(FileName);
            //имя файла для сохранения
            FileName = FileName + "\\" + kurs.ToString() + "_курс_" + starts[week_list.SelectedIndex].ToShortDateString() + ".xls";

            ExcelFile excel = new ExcelFile();

            ExcelWorksheet sheet = excel.Worksheets.Add(
                "Расп. с " + starts[week_list.SelectedIndex].ToShortDateString());


            sheet.PrintOptions.HeaderMargin = 0.0;
            sheet.PrintOptions.FooterMargin = 0.0;
            sheet.PrintOptions.FitToPage = true;
            sheet.PrintOptions.Portrait = false;

            //перечень вывдомых строк
            OutLines.Clear();
            foreach (int er in empty_rows)
            {
                OutLines.Add(er);
            }

            int num = 1;
            int i = 0;
            OutGroups.Clear();
            for (i = 0; i < grupa_list.Items.Count; i++)
            {
                int grupa_kurs = Convert.ToInt32(grups_set.Rows[i][1]);

                switch (kurs)
                {
                    case 0:
                        if (Convert.ToBoolean(grups_set.Rows[i][3]))
                            OutGroups.Add(num);
                        break;
                    default:
                        if (grupa_kurs == kurs)
                        {
                            if (!Convert.ToBoolean(grups_set.Rows[i][3]))
                                OutGroups.Add(num);
                        }
                        break;
                }

                num += 2;
            }

            if (OutGroups.Count == 0)
            {
                excel = null;
                sheet = null;
                return;
            }


            //проверить есть ли суббота и воскресенье
            bool showsubb = false, showvoskr = false;

            for (int c = OutGroups[0]; c <= OutGroups[OutGroups.Count - 1] + 1; c++)
            {
                for (int cc = 37; cc <= 42; cc++)
                {
                    if (table[cc, c].ToString().Trim().Length > 0)
                    {
                        showsubb = true;
                        break;
                    }
                }

                if (showsubb == true) break;
            }

            for (int c = OutGroups[0]; c <= OutGroups[OutGroups.Count - 1] + 1; c++)
            {
                for (int cc = 44; cc <= 49; cc++)
                {
                    if (table[cc, c].ToString().Trim().Length > 0)
                    {
                        showvoskr = true;
                        break;
                    }
                }

                if (showvoskr == true) break;
            }

            //поставить решетку               
            if (showsubb == false && showvoskr == false) OutLines.Remove(36);
            if (showvoskr == false) OutLines.Remove(43);


            //задать общие свойства свойства
            cr = sheet.Cells.GetSubrange("a1",
                Letters[OutGroups.Count * 2 + 1] + (OutLines[OutLines.Count - 1] + 4).ToString());
            cr.Merged = true;
            cr.Style.Font.Name = "Times New Roman";
            cr.Style.Font.Size = 8 * 20;
            //cr.Style.Font.Italic = true;
            cr.Style.Font.Weight = ExcelFont.BoldWeight;
            cr.Style.VerticalAlignment = VerticalAlignmentStyle.Center;
            cr.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            //cr.Style.FillPattern.SetSolid(Color.White);
            cr.Merged = false;

            // ------------------   ===== взять шапку расписания из БД
            cr = sheet.Cells.GetSubrange("A1", Letters[OutGroups.Count * 2 + 1] + "1");
            cr.Merged = true;
            cr.Style.Font.Name = "Verdana";
            cr.Style.Font.Weight = ExcelFont.BoldWeight;
            cr.Style.Font.Italic = false;
            cr.Style.Font.Size = 13 * 20;

            S = "Расписание занятий для " + kurs.ToString() + " курса " +  
                " FSystemа " + fakultet_name_krat + " \n" + 
                "c " + starts[week_list.SelectedIndex].ToShortDateString() +
                " по " + ends[week_list.SelectedIndex].ToShortDateString();

            sheet.Cells["A1"].Value = S;
            sheet.Rows[0].Height = 34 * 20;
            sheet.Rows[1].Height = 5 * 20;

            //определить параметры вывдимого расписания

            int span = 0;

            switch (OutLines.Count)
            {
                case 5: span = 4; break;
                case 6: span = 3; break;
                case 7: span = 2; break;
            }

            cr = sheet.Cells.GetSubrange("A3", "B" + (OutLines[OutLines.Count - 1] + span).ToString());
            cr.Merged = true;
            cr.Style.Borders.SetBorders(MultipleBorders.Outside, Color.Black,
                 GemBox.Spreadsheet.LineStyle.Thin);
            cr.Merged = false;

            //sheet.Cells.GetSubrange("A6", "C6").SetBorders(MultipleBorders.Top, Color.Black, GemBox.Spreadsheet.LineStyle.DoubleLine);
            cr = sheet.Cells.GetSubrange("B3", "B" + (OutLines[OutLines.Count - 1] + span).ToString());
            cr.Merged = true;
            cr.Style.Borders[IndividualBorder.Right].LineStyle = GemBox.Spreadsheet.LineStyle.DoubleLine;
            cr.Merged = false;

            cr = sheet.Cells.GetSubrange("A3", "A" + (OutLines[OutLines.Count - 1] + span).ToString());
            cr.Merged = true;
            cr.Style.Borders[IndividualBorder.Left].LineStyle = GemBox.Spreadsheet.LineStyle.DoubleLine;
            cr.Merged = false;

            cr = sheet.Cells.GetSubrange("A" + (OutLines[OutLines.Count - 1] + span).ToString(),
                "B" + (OutLines[OutLines.Count - 1] + span).ToString());
            cr.Merged = true;
            cr.Style.Borders[IndividualBorder.Bottom].LineStyle = GemBox.Spreadsheet.LineStyle.DoubleLine;
            cr.Merged = false;

            cr = sheet.Cells.GetSubrange("A3", "B3");
            cr.Merged = true;
            cr.Style.Borders[IndividualBorder.Top].LineStyle = GemBox.Spreadsheet.LineStyle.DoubleLine;
            cr.Merged = false;

            cr = sheet.Cells.GetSubrange("A3", "B3");
            cr.Merged = true;
            cr.Style.Borders[IndividualBorder.Bottom].LineStyle = GemBox.Spreadsheet.LineStyle.DoubleLine;
            cr.Merged = false;

            //первый столбец - дни недели
            sheet.Cells["A3"].Value = "День\nнедели";
            sheet.Cells["B3"].Value = "№";
            //sheet.Cells["C3"].Value = "Время";
            sheet.Rows[5].Height = 25 * 20;


            num = 3;
            string cell1 = "", cell2 = "";
            int days = 0;
            foreach (int j in OutLines)  //вывод дней недели, пар и дат
            {               
                
                cell1 = string.Format("A{0}", num + 1);

                sheet.Cells[cell1].Value = table[j, 1].ToString().Substring(0, 5);

                cell1 = string.Format("A{0}", num + 2);
                cell2 = string.Format("A{0}", num + 6);
                cr = sheet.Cells.GetSubrange(cell1, cell2);
                cr.Merged = true;

                cr.Style.Rotation = 90;
                sheet.Cells[cell1].Value = DaysMed[days]; // table[j, 0].ToString();
                cr.Style.Font.Weight = ExcelFont.MaxWeight;
                cr.Style.Font.Size = 12 * 20;
                sheet.Columns[0].Width = 6 * 256;

                for (int k = 1; k <= 6; k++)
                {
                    sheet.Cells[num + k - 1, 1].Value = k.ToString();

                    sheet.Columns[1].Width = 3 * 256;

                    //sheet.Cells[num + k - 1, 2].Value = table[j + k, 0].ToString();
                    //sheet.Columns[2].Width = 11 * 256;

                    sheet.Rows[num + k - 1].Height = 12 * 20;
                }

                num += 6;
                days++;
            }
            
            // задать глобальные параметры столбцов таблицы расписания
            int cols = table.Cols - 1;
            int ColWidth = 17 * 256;                                    
            int ColWidthWide = 25 * 256;
            bool fullpredmettext = false;

            switch (OutGroups.Count)
            {
                case 1:
                    ColWidth = 30 * 256;
                    ColWidthWide = 60 * 256;
                    fullpredmettext = true;
                    break;
                case 2:
                    ColWidth = 30 * 256;
                    ColWidthWide = 60 * 256;
                    fullpredmettext = true;
                    break;
            }

            sheet.DefaultColumnWidth = ColWidth;
            int counter = 2;


            int i1 = 0;
            foreach(int iii in OutGroups)
            //for (i = OutGroups[0]; i < OutGroups.Count*2; i += 2) ///исправить цикл взять количество из Outgroups
            {
                i = iii;
                sheet.Columns[i + 2].Width = ColWidth;
                sheet.Columns[i + 3].Width = ColWidth;

                //название группы
                cell1 = Letters[counter] + "3";
                cell2 = Letters[counter + 1] + "3";

                cr = sheet.Cells.GetSubrange(cell1, cell2);
                cr.Merged = true;
                sheet.Cells[cell1].Value = table[0, i].ToString();
                cr.Style.Font.Weight = ExcelFont.MaxWeight;
                cr.Style.Font.Size = 10 * 20;
                cr.Style.Borders.SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.DoubleLine);

                //цикл вывода предметов
                num = 4;
                

                foreach (int j in OutLines)
                {
                    int ii = 0;
                    for (ii = 1; ii <= 6; ii++)
                    {

                        string text1 = get_cell_shorttext__one_line(j + ii, i, fullpredmettext);
                        string text2 = get_cell_shorttext__one_line(j + ii, i + 1, fullpredmettext);

                        cell1 = Letters[counter] + (num + ii - 1).ToString();
                        cell2 = Letters[counter + 1] + (num + ii - 1).ToString();

                        cr = sheet.Cells.GetSubrange(cell1, cell1);
                        cr.Merged = true;
                        cr.Style.Borders.SetBorders(MultipleBorders.Outside,
                                Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
                        cr.Merged = false;

                        cr = sheet.Cells.GetSubrange(cell2, cell2);
                        cr.Merged = true;
                        cr.Style.Borders.SetBorders(MultipleBorders.Outside,
                                Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
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
                            Color.Black, GemBox.Spreadsheet.LineStyle.DoubleLine);

                    }

                    //вывести горизонтальный разделитель
                    sheet.Cells[cell1].SetBorders(MultipleBorders.Bottom,
                        Color.Black, GemBox.Spreadsheet.LineStyle.DoubleLine);
                    sheet.Cells[cell2].SetBorders(MultipleBorders.Bottom,
                        Color.Black, GemBox.Spreadsheet.LineStyle.DoubleLine);


                    sheet.Cells["A" + (num + ii - 2).ToString()].Style.Borders[IndividualBorder.Bottom].LineStyle =
                        GemBox.Spreadsheet.LineStyle.DoubleLine;

                    cr = sheet.Cells.GetSubrange("B" + (num + ii - 2).ToString(), "B" + (num + ii - 2).ToString());
                    cr.Merged = true;
                    cr.Style.Borders[IndividualBorder.Bottom].LineStyle = GemBox.Spreadsheet.LineStyle.DoubleLine;
                    cr.Merged = false;

                    /*cr = sheet.Cells.GetSubrange("C" + (num + ii - 2).ToString(), "C" + (num + ii - 2).ToString());
                    cr.Merged = true;
                    cr.Style.Borders[IndividualBorder.Bottom].LineStyle = GemBox.Spreadsheet.LineStyle.DoubleLine;
                    cr.Merged = false;*/

                    num += 6;
                }

                counter += 2;
            }
            
            sheet.PrintOptions.HeaderMargin = 0.0;
            sheet.PrintOptions.FooterMargin = 0.0;
            sheet.PrintOptions.FitToPage = true; 
            sheet.PrintOptions.Portrait = false;
            sheet.PrintOptions.BottomMargin = 1;
            sheet.PrintOptions.TopMargin = 1;
            sheet.PrintOptions.LeftMargin = 1;
            sheet.PrintOptions.RightMargin = 1;

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

        //создать папку если она не существует
        public void CreateFolder(string FolderName)
        {            
            string FinalPath = FolderName;

            if (!Directory.Exists(FinalPath))
            {
                Directory.CreateDirectory(FinalPath);
            }
        }

        /// <summary>
        /// получить путь к папке мои документы для текущего пользователя
        /// </summary>
        /// <returns></returns>
        public string GetMyDocs()
        {
            string MyDocs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            return MyDocs;
        }

        private void set_tema_Click(object sender, EventArgs e)
        {
            SaveTema();
        }

        private void set_teacher_tema_Click(object sender, EventArgs e)
        {
            SaveTema();
        }

        /// <summary>
        /// сохранить тему занятия
        /// </summary>
        public void SaveTema()
        {
            int tema_count = 1;

            //стоит ли одинаковое занятие снизу или сверху от текущего
            int c = table.Col;
            int r = table.Row;

            int sgcur = 0;
            Cell current = GetCell(r, c, out sgcur);
            if (current == null) return;

            //верхняя ячейка
            r--;
            int sgup = 0;
            Cell cellup = GetCell(r, c, out sgup);

            //нижняя ячейка
            r += 2;
            int sgdown = 0;
            Cell celldown = GetCell(r, c, out sgdown);

            //целевая ячейка номер два
            Cell second = new Cell();
            //Целевая подгруппа
            int sg_second = 0;

            bool useup = false, usedown = false;

            //проверить непосредственный верх

            if (cellup != null && current.predmet_id[sgcur] == cellup.predmet_id[sgup] &&
            current.prepod_id[sgcur] == cellup.prepod_id[sgup])
            {
                second = cellup;
                sg_second = sgup;
                tema_count = 2;
                useup = true;
            }
            else
            {
                //проверить верхнюю по диагонали
                if (cellup != null && current.predmet_id[sgcur] == cellup.predmet_id[1 - sgup] &&
                current.prepod_id[sgcur] == cellup.prepod_id[1 - sgup])
                {
                    second = cellup;
                    sgup = 1 - sgup;
                    tema_count = 2;
                    useup = true;
                }
                else
                {
                    //проверить нижнюю
                    if (celldown != null && current.predmet_id[sgcur] == celldown.predmet_id[sgdown] &&
                    current.prepod_id[sgcur] == celldown.prepod_id[sgdown])
                    {
                        second = celldown;
                        sg_second = sgdown;
                        tema_count = 2;
                        usedown = true;
                    }
                    else
                    {
                        //проверить нижнюю диагональную
                        if (celldown != null && current.predmet_id[sgcur] == celldown.predmet_id[1 - sgdown] &&
                        current.prepod_id[sgcur] == celldown.prepod_id[1 - sgdown])
                        {
                            second = cellup;
                            sgdown = 1 - sgdown;
                            tema_count = 2;
                            usedown = true;
                        }
                    }
                }
            }


            //подумать о потоке ----                    

            set_tema_dialog std = new set_tema_dialog();
            std.Tag = (object)tema_count;

            DateTime d = new DateTime(current.y[sgcur], current.m[sgcur], current.d[sgcur]);
            string sd = d.ToShortDateString();

            if (tema_count == 1)
            {
                std.Height = 190;
                std.choose_cynchro.Visible = false;
                std.choose_cynchro.Checked = false;
                std.tema1.BackColor = Color.LightGray;
                std.tema1.Text = current.tema[sgcur];
                std.tema1.Focus();

                if (current.subgr_nomer[sgcur] == 0)
                    std.label1.Text = sd + ", " +
                        current.predmet_name[sgcur] + ", " + current.vid_zan_name[sgcur] +
                        ", пара №" + current.nom_zan[sgcur].ToString();
                else
                    std.label1.Text = sd + ", " +
                        current.predmet_name[sgcur] + ", " + current.vid_zan_name[sgcur] + ", " +
                        "подгруппа №" + current.subgr_nomer[sgcur].ToString() +
                        ", пара №" + current.nom_zan[sgcur].ToString();
            }
            else
            {
                std.Height = 310;
                std.choose_cynchro.Visible = true;

                if (useup)
                {
                    std.tema2.BackColor = Color.LightGray;

                    std.tema2.Focus();
                    std.tema2.Select();
                    std.tema2.Text = current.tema[sgcur];
                    std.tema1.Text = cellup.tema[sgup];

                    if (std.tema2.Text.ToLower().Trim() == std.tema1.Text.ToLower().Trim())
                        std.choose_cynchro.Checked = true;
                    else
                        std.choose_cynchro.Checked = false;

                    if (current.subgr_nomer[sgcur] == 0)
                        std.label2.Text = sd + ", " +
                            current.predmet_name[sgcur] + ", " + current.vid_zan_name[sgcur] +
                            ", пара №" + current.nom_zan[sgcur].ToString();
                    else
                        std.label2.Text = sd + ", " +
                            current.predmet_name[sgcur] + ", " + current.vid_zan_name[sgcur] + ", " +
                            "подгруппа №" + current.subgr_nomer[sgcur].ToString() +
                            ", пара №" + current.nom_zan[sgcur].ToString();

                    if (cellup.subgr_nomer[sgup] == 0)
                        std.label1.Text = sd + ", " +
                            cellup.predmet_name[sgup] + ", " + cellup.vid_zan_name[sgup] +
                            ", пара №" + cellup.nom_zan[sgup].ToString();
                    else
                        std.label1.Text = sd + ", " +
                            cellup.predmet_name[sgup] + ", " + cellup.vid_zan_name[sgup] + ", " +
                            "подгруппа №" + cellup.subgr_nomer[sgup].ToString() +
                            ", пара №" + cellup.nom_zan[sgup].ToString();

                }

                if (usedown)
                {
                    std.tema1.BackColor = Color.LightGray;
                    std.tema1.Focus();
                    std.tema1.Select();
                    std.tema1.Text = current.tema[sgcur];
                    std.tema2.Text = celldown.tema[sgdown];

                    if (std.tema2.Text.ToLower().Trim() == std.tema1.Text.ToLower().Trim())
                        std.choose_cynchro.Checked = true;
                    else
                        std.choose_cynchro.Checked = false;

                    if (current.subgr_nomer[sgcur] == 0)
                        std.label1.Text = sd + ", " +
                            current.predmet_name[sgcur] + ", " + current.vid_zan_name[sgcur] +
                            ", пара №" + current.nom_zan[sgcur];
                    else
                        std.label1.Text = sd + ", " +
                            current.predmet_name[sgcur] + ", " + current.vid_zan_name[sgcur] + ", " +
                            "подгруппа №" + current.subgr_nomer[sgcur].ToString() +
                            ", пара №" + current.nom_zan[sgcur].ToString();

                    if (celldown.subgr_nomer[sgdown] == 0)
                        std.label2.Text = sd + ", " +
                            celldown.predmet_name[sgdown] + ", " + celldown.vid_zan_name[sgdown] +
                            ", пара №" + celldown.nom_zan[sgdown].ToString();
                    else
                        std.label2.Text = sd + ", " +
                            celldown.predmet_name[sgdown] + ", " + celldown.vid_zan_name[sgdown] + ", " +
                            "подгруппа №" + celldown.subgr_nomer[sgdown].ToString() +
                            ", пара №" + celldown.nom_zan[sgdown].ToString();
                }

            }

            DialogResult res = std.ShowDialog();


            if (res == DialogResult.Cancel)
            {
                return;
            }

            if (tema_count == 1)
            {
                //обновить current и если не резделена то пложить в обе половины
                //инче в одну 
                if (current.Divided())
                {
                    current.tema[0] = std.tema1.Text;
                    current.tema[1] = std.tema1.Text;
                }
                else
                {
                    current.tema[sgcur] = std.tema1.Text;
                }

                global_command = current.UpdateTemaCommand(sgcur);
                global_command.ExecuteNonQuery();
            }
            else
            {
                if (useup)
                {
                    if (current.Divided())
                    {
                        current.tema[0] = std.tema2.Text;
                        current.tema[1] = std.tema2.Text;
                    }
                    else
                    {
                        current.tema[sgcur] = std.tema1.Text;
                    }

                    global_command = current.UpdateTemaCommand(sgcur);
                    global_command.ExecuteNonQuery();

                    if (cellup.Divided())
                    {
                        cellup.tema[0] = std.tema1.Text;
                        cellup.tema[1] = std.tema1.Text;
                    }
                    else
                    {
                        cellup.tema[sgup] = std.tema1.Text;
                    }

                    global_command = cellup.UpdateTemaCommand(sgup);
                    global_command.ExecuteNonQuery();
                }

                if (usedown)
                {
                    if (current.Divided())
                    {
                        current.tema[0] = std.tema1.Text;
                        current.tema[1] = std.tema1.Text;
                    }
                    else
                    {
                        current.tema[sgcur] = std.tema1.Text;
                    }

                    global_command = current.UpdateTemaCommand(sgcur);
                    global_command.ExecuteNonQuery();

                    if (celldown.Divided())
                    {
                        celldown.tema[0] = std.tema2.Text;
                        celldown.tema[1] = std.tema2.Text;
                    }
                    else
                    {
                        celldown.tema[sgdown] = std.tema2.Text;
                    }

                    global_command = celldown.UpdateTemaCommand(sgdown);
                    global_command.ExecuteNonQuery();
                }

            }

            std.Dispose();       
        }

        /// <summary>
        /// получить ссылку на ячейку структуры данных, соотвествующую выбранной клетке таблицы
        /// </summary>
        /// <param name="r">номер строки</param>
        /// <param name="c">номер столбца</param>
        /// <param name="sg">номер подгруппы (выходной параметр)</param>
        /// <returns></returns>
        public Cell GetCell(int r, int c, out int sg)
        {
            if (!is_correct_cell(r, c))
            {
                sg = 0;
                return null;
            }

            //определить колонки подгрупп
            int first = 0, second = 0;

            if (c % 2 == 0)
            {
                first = c - 1;
                second = c;
            }
            else
            {
                first = c;
                second = c + 1;
            }            

            int dest_col = first; //колонка назначения
            if (c == second) dest_col = second;
            if (c == first) dest_col = first;            

            //определить параметры исходной ячейки
            DateTime d = table_data.RowDate(r);
            int p = table_data.RowPair(r);
            string g = table_data.ColumnGroup(dest_col);
            int psg = table_data.ColumnSubGroup(dest_col);

            sg = psg - 1;
            return table_data[d, g, p];
        }


        //задать количество часов
        private void запомнитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int sg = 0;
            Cell c = GetCell(table.Row, table.Col, out sg);

            //получить введенное число и проверить его
            string txt = chas_box.Text.Trim();           

            string res = "Обнаружена ошибка ввода данных.\n\n";

            if (!IsNumber(ref txt))
            {
                res += " - введено некорректное значение для количества часов [значение не было сохранено]";
                MessageBox.Show(res, "Ошибка ввода",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            double chas = Convert.ToDouble(txt);

            if (chas > 20)
            {
                res += " - введено слишком большое (больше 20) значение для количества часов [значение не было сохранено]";
                MessageBox.Show(res, "Ошибка ввода",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            int predm_id = c.id[sg];

            global_command = new SqlCommand();
            global_command.Connection = global_connection;
            global_command.CommandText = "update rasp set kol_chas = @CHAS where id = @ID";
            global_command.Parameters.Add("@CHAS", SqlDbType.Float).Value = chas;
            global_command.Parameters.Add("@ID", SqlDbType.Int).Value = predm_id;

            c.col_chas[sg] = chas;
          
            try
            {
                global_command.ExecuteNonQuery();
            }
            catch (Exception exx)
            {
            }                       
        }

        /// <summary>
        /// определяет является хранит ли строка вещественное число
        /// удаляет первую и последнюю запятую
        /// </summary>
        /// <param name="txt">текст для анализа</param>
        /// <returns>хранится ли в строке число</returns>
        public static bool IsNumber(ref string txt)
        {
            string stxt = "";
            
            if (txt.Length == 0) return false;

            bool correct = true;

            int commas = 0;
            foreach (char sign in txt)
            {
                if (sign == ',') commas++;
                
                if (!Char.IsDigit(sign) && (sign != ','))
                {
                    correct = false;
                    break;
                }                
            }

            if (commas > 1) correct = false;

            if (commas == 1)
            {
                if (txt.StartsWith(","))
                {                    
                    txt = txt.Remove(0, 1);
                }

                if (txt.EndsWith(","))
                {
                    txt = txt.Remove(txt.Length - 1, 1);
                }
            }

            return correct;
        }

        public static bool tabel_exists = false;
        public tabel t = null;

        /// <summary>
        /// действия при выборе узла дерева
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void object_tree_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.Node.Name == "sprav_prepods")
            {
                sprav_prepods sp = new sprav_prepods();
                sp.ShowDialog();
                sp.Dispose();
            }

            if (e.Node.Name == "predmets_node")
            {
                sprav_predmets spr = new sprav_predmets();
                spr.ShowDialog();
                spr.Dispose();
            }

            if (e.Node.Name == "sprav_group")
            {
                sprav_grupa sg = new sprav_grupa();
                sg.ShowDialog();
                sg.Dispose();
            }

            if (e.Node.Name == "tabel_node")
            {
                if (!tabel_exists)
                {
                    t = new tabel();
                    tabel_exists = true;
                    t.Show();
                }
                else
                    t.Activate();               
            }

            if (e.Node.Name == "lorap")
            {
                pass_edit pe = new pass_edit();
                DialogResult pe_res = pe.ShowDialog();
                pe.Dispose();
            }

            if (e.Node.Name == "sprav_student")
            {
                sprav_student ss = new sprav_student();
                ss.ShowDialog();
                ss.Dispose();
            }

            if (e.Node.Name == "vedomost")
            {
                document_vedomost dv = new document_vedomost();
                dv.ShowDialog();
                dv.Dispose();
            }

            if (e.Node.Name == "obiavlene")
            {
                sprav_ob so = new sprav_ob();
                so.ShowDialog();
                so.Dispose();
            }

            //задание ВКР
            if (e.Node.Name == "vkr")
            {
                vkr_diplom vd = new vkr_diplom();
                vd.toolStripStatusLabel1.Text = "Определение параметров выпускных квалилфикационных работ";
                vd.kurs = 4;
                vd.ShowDialog();
                vd.Dispose();
            }

            //задание ВКР
            if (e.Node.Name == "dip")
            {
                vkr_diplom vd = new vkr_diplom();
                vd.toolStripStatusLabel1.Text = "Определение параметров дипломных работ";
                vd.kurs = 5;
                vd.ShowDialog();
                vd.Dispose();
            }

            //перевод на след курс
            if (e.Node.Name == "perevod")
            {
                //
            }

            if (e.Node.Name == "Uch_god_node")
            {
                //редактировать учбеные года
                //sprav_uch_god sug = new sprav_uch_god();
                //sug.ShowDialog();
                //sug.Dispose();                
            }
        }

        private void object_tree_KeyDown(object sender, KeyEventArgs e)
        {
            TreeNode node = object_tree.SelectedNode;

            if (node == null) return;

            //Text = node.Name;

            if (e.KeyCode == Keys.Return)
            {
                if (node.Name == "sprav_prepods")
                {
                    sprav_prepods sp = new sprav_prepods();
                    sp.ShowDialog();
                    sp.Dispose();
                }

                if (node.Name == "predmets_node")
                {
                    sprav_predmets spr = new sprav_predmets();
                    spr.ShowDialog();
                    spr.Dispose();
                }

                if (node.Name == "sprav_group")
                {
                    sprav_grupa sg = new sprav_grupa();
                    sg.ShowDialog();
                    sg.Dispose();
                }

                if (node.Name == "tabel_node")
                {
                    if (!tabel_exists)
                    {
                        t = new tabel();
                        tabel_exists = true;
                        t.Show();
                    }
                    else
                        t.Activate();

                }

                if (node.Name == "lorap")
                {
                    pass_edit pe = new pass_edit();
                    DialogResult pe_res = pe.ShowDialog();
                    pe.Dispose();
                }

                if (node.Name == "sprav_student")
                {
                    sprav_student ss = new sprav_student();
                    ss.ShowDialog();
                    ss.Dispose();
                }
            }

            if (dekan_online) return;

            if (e.Alt)
            {
                if (e.Control)
                {
                    if (e.Shift)
                    {
                        if (e.KeyCode == Keys.Return)
                        {
                            enter_as_dekan de = new enter_as_dekan();
                            DialogResult de_res = de.ShowDialog();

                            if (de_res == DialogResult.Cancel || de.status == false)
                            {
                                dekan_online = false;
                                return;
                            }

                            de.Dispose();
                            dekan_online = true;
                            object_tree.Nodes.Add(fakultet_node);
                            object_tree.Nodes.Add(sprav_node);
                            fakultet_node.ExpandAll();
                            sprav_node.ExpandAll();
                            table.ContextMenuStrip = contextMenuStrip1;
                            statistica.Visible = true;
                            
                            if (!content.TabPages.Contains(tabПосещаемость))
                                content.TabPages.Add(tabПосещаемость);

                            if (!content.TabPages.Contains(tabPageAttest))
                                content.TabPages.Add(tabPageAttest);
                            //prepod_details.Visible = true;

                        }
                    }
                }
            }

        }
        // ------------------  конец функции --- 

        private void photo_item_menu_Click(object sender, EventArgs e)
        {
            //показать полную инфу про препода
            DataRow row = prepod_set.Rows[prepod_list.SelectedIndex];

            int prep_id = (int)row[0];

            prepod_edit pe = new prepod_edit();

            pe.prep_id = prep_id;

            pe.dolz_id = (int)row[7];
            pe.zvan_id = (int)row[9];
            pe.uch_id = (int)row[8];
            pe.kaf_id = (int)row[10];
            pe.pictureBox1.Image = GetPhotoFromBD("prepod", prep_id);
            pe.deny_photo = true;

            pe.status_box.Checked = (bool)row[11];
            if (pe.status_box.Checked)
                pe.status_box.Text = "Статус: работает";
            else
                pe.status_box.Text = "Статус: уволен";

            bool sex = (bool)row[2];

            if (sex == false)
            {
                pe.female.Checked = true;
            }

            pe.fam.Text = row[4].ToString();
            pe.im.Text = row[5].ToString();
            pe.ot.Text = row[6].ToString();
            pe.phone.Text = row[13].ToString();
            pe.email.Text = row[12].ToString();

            pe.button8.Left = pe.button7.Left;
            pe.button8.Text = "Закрыть";
            pe.button7.Visible = false;
            pe.button4.Visible = false;
            pe.button3.Visible = false;
            pe.button2.Visible = false;
            pe.button1.Visible = false;

            DialogResult pres = pe.ShowDialog();
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            fill_table();
        }

        private void set_teacher_attend_Click(object sender, EventArgs e)
        {

            int c = table.MouseCol;
            int r = table.MouseRow;

            if (c < 0 || r < 0) return;

            if (!is_correct_cell(r, c)) return;

            Cell cc = table_data[cd, cgp, cp];

            //отметить посещение
            set_attend_dialog sad = new set_attend_dialog();

            sad.grupa_id = cc.grupa_id[csub - 1];
            sad.zan_id = cc.id[csub - 1];
            sad.subgr = csub;          
            sad.Text = grupa_list.Text + "," + cd.ToLongDateString() + ", пара №" + 
                cp.ToString() + " | " + cc.predmet_name[csub - 1];
            
            sad.delenie = (cc.id[0]!=cc.id[1]);
            sad.tema = cc.tema[csub - 1];
            sad.prim = cc.str_prim[csub - 1];
            
            sad.ShowDialog();
            
            if (sad!=null) sad.Dispose();
        }

        private void set_attend_Click(object sender, EventArgs e)
        {
            set_teacher_attend_Click(sender, new EventArgs());
        }

        // -------------  нормализация строковых значений ----

        public static string NormalizeLetters(string str)
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

            int dash = 0;
            foreach (char s in str)
            {
                if (s == '-') dash++;
            }

            if (dash == str.Length) str = "";

            return str;
        }

        public static string Normalize(string str)
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


        public static string Normalize1(string str)
        {
            while (str.Contains("'"))
            {
                int pos = str.IndexOf("'");
                str = str.Remove(pos, 1);
                if (str.Length == 0) break;
            }
            return str;
        }

        private void копироватьВБуферToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Clipboard.Clear();
            try
            {
                CopyGridToClipBoard(prepod_table);
                //Clipboard.SetDataObject(prepod_table.GetClipboardContent());
            }
            catch (Exception ex)
            {
                ;
            }
        }


        //работа с индивидульным раписанием
        private void tecaher_tab_raspisanie_Enter(object sender, EventArgs e)
        {
            //load_individ_rasp();
        }


        public DataTable personal_tabble = null;

        /// <summary>
        /// загрузить индивидуальное расписание    
        /// </summary>
        public void load_individ_rasp()
        {            
            //загрузка расписания
            global_query =
            " select " +
            " [Дата]=cast(d as char(2)) + '/' + " +                //0
            " cast (m as char(2)) + '/' + cast(y as char(4)) + '-' + " +
            " case " +
            " when datepart(dw, dbo.get_date(y,m,d))=1 then 'вс' " +
            " when datepart(dw, dbo.get_date(y,m,d))=2 then 'пн'  " +
            " when datepart(dw, dbo.get_date(y,m,d))=3 then 'вт' " +
            " when datepart(dw, dbo.get_date(y,m,d))=4 then 'ср' " +
            " when datepart(dw, dbo.get_date(y,m,d))=5 then 'чт' " +
            " when datepart(dw, dbo.get_date(y,m,d))=6 then 'пт' " +
            " when datepart(dw, dbo.get_date(y,m,d))=7 then 'сб' " +
            " end, [Пара]=nom_zan, " +  //1
            " [Группа] = grupa.name, " + //2  
            " [Предмет]=predmet.name_krat, [Тема]=tema, " + //3 4
            " [Вид занятия]=vid_zan.name, vid_zan_id, " + //5  6 
            " rasp.id, rasp.subgr_nomer, rasp.grupa_id, rasp.kol_chas, predmet.name, predmet.id, prim = isnull(rasp.prim_text,'') " + //7 8 9 10 11 12 13
            " from rasp " +
            " join predmet on predmet.id = rasp.predmet_id " +
            " join grupa on grupa.id=rasp.grupa_id " +
            " join vid_zan on vid_zan.id =rasp.vid_zan_id " +
            " where rasp.prepod_id = @USERID and dbo.get_date(y,m,d)>=@D1-1 and dbo.get_date(y,m,d)<=@D2 " +  
            PredmetNameFilter + 
            " order by y, m, d, nom_zan ";

            global_command = new SqlCommand(global_query, global_connection);
            global_command.Parameters.Add("@USERID", SqlDbType.Int).Value = active_user_id;
            global_command.Parameters.Add("@D1", SqlDbType.DateTime).Value = begin.Value;
            global_command.Parameters.Add("@D2", SqlDbType.DateTime).Value = end.Value;

            personal_tabble = new DataTable();
            global_adapter = new SqlDataAdapter(global_command);

            global_adapter.Fill(personal_tabble);

            prepod_table.Rows.Clear();

            int i = 0;
            double SumChas = 0.0;
            foreach (DataRow dr in personal_tabble.Rows)
            {
                object[] pars = new object[7] { dr[0], dr[1], dr[2], dr[3], dr[5], dr[10], dr[4] };
                prepod_table.Rows.Add(pars);
                SumChas += Convert.ToDouble(dr[10]);
                i++;
            }  
          
            // вывести инф о периоде
            // кол-во занятий, кол-во часов
            indRaspInfoLabel.Text = string.Empty;
            if (personal_tabble.Rows.Count > 0)
            {
                indRaspInfoLabel.Text = string.Format(
                    "Занятий = {0}, кол-во часов = {1:F2}", personal_tabble.Rows.Count, SumChas);                
            }

        }

        /// <summary>
        /// выбрать вид занятия в таблице инд расписания
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void выбратьВидЗантToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (personal_tabble.Rows.Count == 0) return;

            int num = prepod_table.CurrentCell.RowIndex;

            if (num < 0) return;

            string sql = string.Format(
                "select vid_zan.id, vid_zan.name from vid_zan " + 
                " join vidzan_predmet on vidzan_predmet.vidzan_id = vid_zan.id " + 
                " where vidzan_predmet.predmet_id = {0} and vid_zan.is_kontrol=0",
                personal_tabble.Rows[num][12]);
            
            DataTable dt = new DataTable();
            (new SqlDataAdapter(sql, global_connection)).Fill(dt);

            ListWindow lw = new ListWindow();
            lw.tbl = dt;
            lw.Text = "Выберите вид занятия из списка";
            lw.Width = 350;

            DialogResult dres = lw.ShowDialog();
            if (dres != DialogResult.OK) return;

            sql = "update rasp set vid_zan_id = @VZID where id = @ID";
            SqlCommand cmd = new SqlCommand(sql, global_connection);
            cmd.Parameters.Add("@VZID", SqlDbType.NVarChar).Value = lw.resId;
            cmd.Parameters.Add("@ID", SqlDbType.Int).Value = personal_tabble.Rows[num][7];
            cmd.ExecuteNonQuery();

            prepod_table.Rows[num].Cells[4].Value = lw.str_res;
            lw.Dispose();
            GC.Collect();
            
        }

        //запись темы занятия и отметка посещения
        private void отметитьПосещениеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (personal_tabble.Rows.Count == 0) return;

            int num = prepod_table.CurrentCell.RowIndex;

            if (num < 0) return;

            DataRow dr = personal_tabble.Rows[num];

            //отметить посещение
            set_attend_dialog sad = new set_attend_dialog();

            sad.grupa_id = (int)dr[9];
            sad.zan_id = (int)dr[7];
            sad.subgr = (int)dr[8];
   
            sad.Text =   
                dr[2].ToString() + "," + prepod_table[0,num].Value.ToString() + ", пара №" + 
                dr[1].ToString() + " | " + dr[3].ToString();
            
            sad.delenie = (sad.subgr!=0);
            sad.tema = dr[4].ToString();
            sad.prim = dr[13].ToString();
            sad.textBox1.ReadOnly = false;
            sad.textBox2.ReadOnly = false;
            
            sad.ShowDialog();

            if (sad.tema_changed || sad.prim_changed) load_individ_rasp();

            prepod_table.CurrentCell = prepod_table.Rows[num].Cells[0];

            if (sad!=null) sad.Dispose();
        }

        //отметить посещение занятия
        private void prepod_table_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (personal_tabble.Rows.Count == 0) return;

            int num = prepod_table.CurrentCell.RowIndex;

            if (num < 0) return;

            DataRow dr = personal_tabble.Rows[num];

            //отметить посещение
            set_attend_dialog sad = new set_attend_dialog();

            sad.grupa_id = (int)dr[9];
            sad.zan_id = (int)dr[7];
            sad.subgr = (int)dr[8];

            sad.Text =
                dr[2].ToString() + "," + prepod_table[0, num].Value.ToString() + ", пара №" +
                dr[1].ToString() + " | " + dr[3].ToString();

            sad.delenie = (sad.subgr != 0);
            sad.tema = dr[4].ToString();
            sad.prim = dr[13].ToString();
            sad.textBox1.ReadOnly = false;
            sad.textBox2.ReadOnly = false;

            sad.ShowDialog();

            if (sad.tema_changed || sad.prim_changed) load_individ_rasp();

            if (sad != null) sad.Dispose();
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            load_individ_rasp();
        }

        //действия при изменении начальной даты на вкладке индивидульного расписания
        private void begin_ValueChanged(object sender, EventArgs e)
        {
            end.MinDate = begin.Value;
            if (!одновременное_смещение)
            load_individ_rasp();
        }

        //действия при изменении конечной даты на вкладке индивидульного расписания
        private void end_ValueChanged(object sender, EventArgs e)
        {
            begin.MaxDate = end.Value;
            if (!одновременное_смещение)
            load_individ_rasp();
        }

        private bool одновременное_смещение = false;

        //смещение даты на день назад
        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            одновременное_смещение = true;
            begin.Value = begin.Value.AddDays(-1);
            end.Value = end.Value.AddDays(-1);
            load_individ_rasp();
            одновременное_смещение = false;
        }

        //смещение даты на день вперед
        private void toolStripButton5_Click(object sender, EventArgs e)
        {
            одновременное_смещение = true;
            end.Value = end.Value.AddDays(1);
            begin.Value = begin.Value.AddDays(1);            
            load_individ_rasp();
            одновременное_смещение = false;
        }

        //смещение даты на неделю назад
        private void toolStripButton6_Click(object sender, EventArgs e)
        {
            одновременное_смещение = true;
            begin.Value = begin.Value.AddDays(-7);
            end.Value = end.Value.AddDays(-7);
            load_individ_rasp();
            одновременное_смещение = false;
        }

        //смещение даты на неделю вперед
        private void toolStripButton7_Click(object sender, EventArgs e)
        {
            одновременное_смещение = true;
            end.Value = end.Value.AddDays(7);
            begin.Value = begin.Value.AddDays(7);         
            load_individ_rasp();
            одновременное_смещение = false;
        }

        //смещение даты на месяц назад
        private void toolStripButton8_Click(object sender, EventArgs e)
        {
            одновременное_смещение = true;
            begin.Value = begin.Value.AddMonths(-1);
            end.Value = end.Value.AddMonths(-1);
            load_individ_rasp();
            одновременное_смещение = false;
        }

        //смещение даты на месяц вперед
        private void toolStripButton9_Click(object sender, EventArgs e)
        {
            одновременное_смещение = true;
            end.Value = end.Value.AddMonths(1);
            begin.Value = begin.Value.AddMonths(1);            
            load_individ_rasp();
            одновременное_смещение = false;
        }
        
        //поменять количество часов для данного занятия
        private void задатьЧасыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (personal_tabble.Rows.Count == 0) return;

            int num = prepod_table.CurrentCell.RowIndex;

            if (num < 0) return;

            DataRow dr = personal_tabble.Rows[num];
            int zan_id = (int)dr[7];
                        
            inputbox ib = new inputbox(
                "Введите в окно редактирования количество часов\n" +
                "по указанному виду занятия.\n\nЦелая часть числа отделяется от дробной знаком 'запятая' (,).",
                dr[5].ToString(),
                dr[10].ToString(),
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
            cmd.Parameters.Add("@ID", SqlDbType.Int).Value = zan_id;

            try
            {
                cmd.ExecuteNonQuery();
            }
            catch (Exception exx)
            {
                MessageBox.Show("Ошибка при передаче данных. Повторите операцию позднее.",
                    "Ошибка данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            ib.Dispose();

            //переместить фокус в исходную строку
            load_individ_rasp();
            prepod_table.CurrentCell = prepod_table.Rows[num].Cells[0];
        }

        public string PredmetNameFilter = string.Empty;
        /// <summary>
        /// сброс фильтра в таблице инд расписания
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripButton23_Click(object sender, EventArgs e)
        {
            //MessageBox.Show(PredmetNameFilter);
            PredmetNameFilter = string.Empty;
            load_individ_rasp();
        }

        /// <summary>
        /// применить фильтр в таблице инд расписания
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripButton24_Click(object sender, EventArgs e)
        {
            int[] filterBoxes = { grupFilterBox.Text.Trim().Length, 
                                  predemFilterBox.Text.Trim().Length, 
                                  vidzanFilterBox.Text.Trim().Length, 
                                  temaFilterBox.Text.Trim().Length 
                                };

            string[] fieldNames = { " grupa.name like '%" + Normalize1(grupFilterBox.Text.Trim()) + "%'", 
                                    " predmet.name_krat like '%" + Normalize1(predemFilterBox.Text.Trim()) + "%'", 
                                    " vid_zan.name like '%" + Normalize1(vidzanFilterBox.Text.Trim()) + "%'", 
                                    " rasp.tema like '%" + Normalize1(temaFilterBox.Text.Trim()) + "%'"};

            PredmetNameFilter = string.Empty;
            int i = 0;
            int notempty = -1;
            for (i = 0; i < filterBoxes.Length; i++)
            {
                if (filterBoxes[i] != 0)
                {
                    notempty = i;
                    break;
                }
            }

            if (notempty == -1) return;

            PredmetNameFilter = " and " + fieldNames[notempty];

            if (notempty + 1 < filterBoxes.Length)
            {
                for (int j = notempty + 1; j < filterBoxes.Length; j++)
                {
                    if (filterBoxes[j] != 0)
                    {
                        PredmetNameFilter = PredmetNameFilter + " and " + fieldNames[j];
                    }
                }
            }

            if (PredmetNameFilter.Trim().Length != 0)
                load_individ_rasp();
        }


        // ------------------------------------------------------

        bool zachet_filled = false;
        //вывести фамилии студентов в засетную таблицу и если есть результаты зачета,
        //то их тоже вывести

        DataTable stud_set; //список студентов для выствления оценки или зачета
        DataTable rsp; //расписание для опреления зачета и др. и т.п.
        DataGridViewComboBoxCell zach_dcell;
        DataTable выставленные_отметки = new DataTable();//массив ид отметок   
        DataTable otmetki = new DataTable();
        
        //струкрура доя ведения статистики отметок 
        public class видотметки_количество
        {
            public string OtmName;
            public int Kol;

            public видотметки_количество(string nm, int k)
            {
                OtmName = nm;
                Kol = k;
            }
            public void inc()
            {
                Kol++;
            }
        }

        /// <summary>
        /// массив видов отметок с указанием их количества
        /// </summary>
        public List<видотметки_количество> статистика_отметок = new List<видотметки_количество>(5);

        /// <summary>
        /// выставить зачёт
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void teacher_tab_predmet_zachet_Enter(object sender, EventArgs e)
        {
            //if (zachet_filled) return;

            zachet_filled = false;

            //получить всех студентов изучающих предмет
            string query = "select 	student.id, grupa.id " +                    
                    " from student  " +
                    " join grupa on grupa.id = student.gr_id " +
                    " join predmet on predmet.grupa_id = grupa.id  " +
                    " where student.actual=1 and status_id = 1 and fam<>'-' and predmet.id = " + id_predmet_in_tree.ToString() + 
                    " order by fam, im, ot ";

            stud_set = new DataTable();
            global_adapter = new SqlDataAdapter(query, main.global_connection);
            global_adapter.Fill(stud_set);


            if (stud_set.Rows.Count == 0)
            {
                MessageBox.Show("Список студентов этой группы пуст. Необходимо создать список группы.",
                    "Операция невозможна",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            zachet_table.Rows.Clear();

            int i = 0;
            //foreach(DataRow dr in stud_set.Rows)
            //{
            //    zachet_table.Rows.Add();
            //    zachet_table.Rows[i].Cells[0].Value = dr[1];
            //    i++;
            //}

            //определить, заплнирован ли зачет по этому предмету
            query = string.Format(
                " select dbo.get_date(y,m,d), rasp.id, semestr_id, vid_zan.kod from rasp " + 
	            " join vid_zan on vid_zan.id = rasp.vid_zan_id   " + 
                " where predmet_id = {0} and   " + 	          
                " rasp.uch_god_id = {1} and " + 
	            " (vid_zan.kod = 'з' or vid_zan.kod = 'дз')",           
                id_predmet_in_tree, uch_god);


            /*
             "select dbo.get_date(y,m,d), rasp.id, semestr_id from rasp " +
                " join vid_zan on vid_zan.id = rasp.vid_zan_id " +
                " where predmet_id = {0} and " +
                " y >= {1} and y<={2} " +
                " and (m>=9 or (m>=1 and m<=6))" + 
                " and (vid_zan.kod = 'з' or vid_zan.kod = 'дз') " +
                " and -semestr_id%2+2 = {3}",
             */

            rsp = new DataTable();
            global_adapter = new SqlDataAdapter(query, global_connection);
            global_adapter.Fill(rsp);

            //MessageBox.Show(query + "\n" /*+ rsp.Rows[0][2].ToString()*/);


            int res = rsp.Rows.Count;

            if (res == 0)
            {
                zachet_message.BackColor = Color.Red;
                zachet_message.ForeColor = Color.White;
                zachet_table.BackgroundColor = Color.LightGray;
                zachet_message.Text = "Зачёт ещё не запланирован по расписанию.";
                //MessageBox.Show("Зачёт ещё не запланирован по расписанию.");
                zachet_table.Visible = false;
                return;
            }
            
            //получить дату зачета
            DateTime zach_date = Convert.ToDateTime(rsp.Rows[0][0]);
            string rasp_id = rsp.Rows[0][1].ToString();

            zachet_message.BackColor = statusStrip2.BackColor;
            zachet_message.ForeColor = Color.Black;
            zachet_message.Text = "Зачёт по расписанию: " + zach_date.ToShortDateString();
            zachet_table.BackgroundColor = Color.White;
            zachet_table.Visible = true;


            //заполнить список отметок

            //получить вид занятия - зачет или диф зачет
            query = string.Format(
                "select kod, vid_zan.id from predmet " +
                " join vidzan_predmet on vidzan_predmet.predmet_id = predmet.id " +
                " join vid_zan on vid_zan.id = vidzan_predmet.vidzan_id " +
                " where predmet.id = {0} and (vid_zan.kod = 'з' or vid_zan.kod = 'дз')", 
                id_predmet_in_tree);

            rsp = new DataTable();
            global_adapter = new SqlDataAdapter(query, global_connection);
            global_adapter.Fill(rsp);

            string vid = rsp.Rows[0][0].ToString();            
            
            ///   ---- получить и вывести все виды отметок --------------
            otmetki = new DataTable();
            query = "select vid_otmetka.id, vid_otmetka.str_name " +
                " from vid_otmetka " +
                " join vid_zan_otmetka on vid_zan_otmetka.vid_otmetka_id = vid_otmetka.id " +
                " join vid_zan on vid_zan.id = vid_zan_otmetka.vid_zan_id " +
                " where vid_zan.id = " + rsp.Rows[0][1].ToString();
            global_adapter = new SqlDataAdapter(query, global_connection);
            global_adapter.Fill(otmetki);

            zach_dcell = new DataGridViewComboBoxCell();
            zachet_table.Columns[1].CellTemplate = zach_dcell;

            vid_zach_table.Columns.Clear(); //очистить сетку вывода статистки отметок
            vid_zach_table.Rows.Clear();            
            //очистить массив статистики отметок
            статистика_отметок.Clear();
            видотметки_количество vk=null;

            int i1 = 0;
            for (i1 = 0; i1<otmetki.Rows.Count; i1++)
            {
                string nm = otmetki.Rows[i1][1].ToString();
                zach_dcell.Items.Add(nm);
                vid_zach_table.Columns.Add("vd" + i1.ToString(), nm);
                vid_zach_table.Columns[i1].SortMode = DataGridViewColumnSortMode.NotSortable;
                vk = null;
                vk = new видотметки_количество(nm, 0);                
                статистика_отметок.Add(vk);
            }
            
            vid_zach_table.Columns.Add("итог", "Итого сдавали");
            vid_zach_table.Columns[i1].SortMode = DataGridViewColumnSortMode.NotSortable;
            vid_zach_table.Columns[i1].DefaultCellStyle.BackColor = Color.Pink;
            
            vid_zach_table.Rows.Add();    

            // ---------------------------------------------------

            query = string.Format("exec dbo.TGetSessionResult {0}, {1}, {2}",
                stud_set.Rows[0][1], id_predmet_in_tree, rsp.Rows[0][1]);
            //MessageBox.Show("Это зачёт - \n" + query);
            global_command = new SqlCommand(query, global_connection);
            global_command.ExecuteNonQuery();

            //если уже есть оценки то вывести их
            /*query = "select count(*) from session " +
                " join rasp on rasp.id = session.rasp_id " +
                " where rasp.predmet_id = " + id_predmet_in_tree.ToString() + 
                " and rasp.id = " + rasp_id +
                " and rasp.vid_zan_id = " + rsp.Rows[0][1].ToString();

            global_command = new SqlCommand(query, global_connection);
            res = Convert.ToInt32(global_command.ExecuteScalar());

            //zach_button.Text = res.ToString();

            if (res == 0)  //записей еще нет, создаем ... 
            {
                //MessageBox.Show("новая");
                //создать записи в таблицы session
                int r = 0;
                foreach (DataRow dr in stud_set.Rows)
                {
                    query = "insert into session (student_id, vid_zan_id, rasp_id, predmet_id, otmetka_id) " +
                        " values (@student_id, @vid_zan_id, @rasp_id, @predm_id, 13)";
                    global_command = new SqlCommand(query, global_connection);

                    global_command.Parameters.Add("@student_id",SqlDbType.Int).Value = dr[0];
                    global_command.Parameters.Add("@vid_zan_id", SqlDbType.Int).Value = rsp.Rows[0][1];
                    global_command.Parameters.Add("@rasp_id", SqlDbType.Int).Value = rasp_id;
                    global_command.Parameters.Add("@predm_id", SqlDbType.Int).Value = id_predmet_in_tree;
                    global_command.ExecuteNonQuery();
                    r++;
                }

                //MessageBox.Show(r.ToString());
            }*/

            //получить данные из таблцы session             
            выставленные_отметки = new DataTable();

            query = "select	student.id, isnull(vid_otmetka.id,-1), prim, " +  //0-2
                " session.id, otm = isnull(vid_otmetka.str_name,''), " +   //3-4
                " fio = student.fam + ' ' + left(student.im,1) + '. ' + left(student.ot,1), session.rasp_id " +  //5-6
                " from student  " +
                " join session on session.student_id = student.id " +
                " left outer join vid_otmetka on vid_otmetka.id = session.otmetka_id  " +
                " where " +
                " student.gr_id = " + stud_set.Rows[0][1].ToString() +
                " and session.predmet_id = " + id_predmet_in_tree.ToString() +
                " and isnull(session.sessiondate, getdate()) between " +
                string.Format("cast('{0}' as datetime) and cast('{1}' as datetime)",
                    starts[0].ToString("yyyyMMdd"), starts[starts.Count - 1].ToString("yyyyMMdd")) + 
                " and (session.vid_zan_id = 6 or session.vid_zan_id = 16) " +
                " order by fam, im, ot ";

            int predm = id_predmet_in_tree;
            //int grupa = 

            global_adapter = new SqlDataAdapter(query, global_connection);
            global_adapter.Fill(выставленные_отметки);

            for (int i2 = 0; i2 < выставленные_отметки.Rows.Count; i2++)
            {
                zachet_table.Rows.Add();
                zachet_table.Rows[i2].Cells[0].Value = выставленные_отметки.Rows[i2][5].ToString();
                int otz = Convert.ToInt32(выставленные_отметки.Rows[i2][1]);
                if (otz != -1)
                {
                    string nm = otmetka_name(otmetki, 0, otz);
                    zachet_table.Rows[i2].Cells[1].Value = nm;
                    внести_отметку_в_статистику(nm);                    
                }
                zachet_table.Rows[i2].Cells[2].Value = выставленные_отметки.Rows[i2][2];
            }

            //вывод значений в нижнюю таблицу статистики
            int vsego = 0;
            for (i = 0; i < статистика_отметок.Count; i++)
            {
                vid_zach_table.Rows[0].Cells[i].Value = статистика_отметок[i].Kol;
                if (статистика_отметок[i].OtmName == "неявка") vsego -= статистика_отметок[i].Kol;
                if (статистика_отметок[i].OtmName == "недопуск") vsego -= статистика_отметок[i].Kol;
                vsego += статистика_отметок[i].Kol;
            }
            vid_zach_table.Rows[0].Cells[i].Value = vsego.ToString();



            zachet_filled = true;
        }

        private void внести_отметку_в_статистику(string p)
        {
            for (int i = 0; i < статистика_отметок.Count; i++) 
            {
                if (статистика_отметок[i].OtmName == p)
                {                    
                    статистика_отметок[i].inc();                    
                    break;
                }
            }
        }

        /// <summary>
        /// получить название отметки
        /// </summary>
        /// <param name="d"></param>
        /// <param name="ind"></param>
        /// <param name="id"></param>
        /// <returns></returns>
        string otmetka_name(DataTable d, int ind, int id)
        {
            for (int i = 0; i < d.Rows.Count; i++)
            {
                int idr = Convert.ToInt32(d.Rows[i][ind]);

                if (idr == id) return d.Rows[i][1].ToString();
            }

            return "";
        }

        /// <summary>
        /// найти индекс отметки с указанным названием
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        string otmetka_index(string name)
        {
            for (int i = 0; i < otmetki.Rows.Count; i++)
            {
                string idr = otmetki.Rows[i][1].ToString();

                if (idr == name) return otmetki.Rows[i][0].ToString();
            }

            return "";
        }


        //выбрана зачётная оценка
        private void zachet_table_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {            
            if (zachet_table.Rows.Count == 0) return;

            if (!zachet_filled) return;

            //сохранить оценку
            if (e.ColumnIndex == 1)
            {
                string sess_id = выставленные_отметки.Rows[e.RowIndex][3].ToString(),
                    name = zachet_table.Rows[e.RowIndex].Cells[1].Value.ToString();

                string q = "update session set otmetka_id = " + 
                    otmetka_index(name) + " where id = " + sess_id;
                global_command = new SqlCommand(q, global_connection);
                global_command.ExecuteNonQuery();

                int i = 0, vsego = 0;
                for (i = 0; i < статистика_отметок.Count; i++)
                {
                    string n = статистика_отметок[i].OtmName;
                    статистика_отметок[i].Kol = 0;
                    for (int j = 0; j < zachet_table.Rows.Count; j++)
                    {
                        if (zachet_table.Rows[j].Cells[1].Value != null)
                        {
                            if (zachet_table.Rows[j].Cells[1].Value.ToString() == n)
                                статистика_отметок[i].inc();
                        }
                    }
                    
                    vid_zach_table.Rows[0].Cells[i].Value = статистика_отметок[i].Kol.ToString();
                    if (статистика_отметок[i].OtmName == "неявка") vsego -= статистика_отметок[i].Kol;
                    if (статистика_отметок[i].OtmName == "недопуск") vsego -= статистика_отметок[i].Kol;
                    vsego += статистика_отметок[i].Kol;
                }
                vid_zach_table.Rows[0].Cells[i].Value = vsego.ToString();

                return;
            }

            //сохранить примечания
            if (e.ColumnIndex == 2)
            {
                string sess_id = выставленные_отметки.Rows[e.RowIndex][3].ToString(),
                    name = "";
                if (zachet_table.Rows[e.RowIndex].Cells[2].Value!=null)
                    name = zachet_table.Rows[e.RowIndex].Cells[2].Value.ToString();                

                string q = "update session set prim = '" + name + "' where id = " + sess_id;
                global_command = new SqlCommand(q, global_connection);
                global_command.ExecuteNonQuery();                
                return;
            } 
        }

        /// <summary>
        /// выставить экзаменационные оценки
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void teacher_tab_predmet_exam_Enter(object sender, EventArgs e)
        {
            //if (zachet_filled) return;

            zachet_filled = false;

            //получить всех студентов изучающих предмет
            string query = "select 	student.id, grupa_id " +                    
                    " from student  " +
                    " join grupa on grupa.id = student.gr_id " +
                    " join predmet on predmet.grupa_id = grupa.id  " +
                    " where student.actual=1 and student.status_id=1 and fam<>'-' and predmet.id = " + id_predmet_in_tree.ToString() +
                    " order by fam, im, ot ";

            stud_set = new DataTable();
            global_adapter = new SqlDataAdapter(query, main.global_connection);
            global_adapter.Fill(stud_set);

            if (stud_set.Rows.Count == 0)
            {
                MessageBox.Show("Список студентов этой группы пуст. Необходимо создать список группы.", 
                    "Операция невозможна", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            exam_table.Rows.Clear();

            int i = 0;
            //foreach (DataRow dr in stud_set.Rows)
            //{
            //    exam_table.Rows.Add();
            //    exam_table.Rows[i].Cells[0].Value = dr[1];
            //    i++;
            //}

            //определить, заплнирован ли экзамен по этому предмету
            query = string.Format(
                " select dbo.get_date(y,m,d), rasp.id, semestr_id, vid_zan.kod from rasp " + 
	            " join vid_zan on vid_zan.id = rasp.vid_zan_id   " + 
                " where predmet_id = {0} and   " + 	          
                " rasp.uch_god_id = {1} and " + 
	            " (vid_zan.kod = 'э')",           
                id_predmet_in_tree, uch_god);

            /*
                "select dbo.get_date(y,m,d), rasp.id from rasp " +
                " join vid_zan on vid_zan.id = rasp.vid_zan_id " +
                " where predmet_id = {0} and " +
                " y >= {1} and y<={2} and (vid_zan.kod = 'э')",
             */


            rsp = new DataTable();
            global_adapter = new SqlDataAdapter(query, global_connection);
            global_adapter.Fill(rsp);

            int res = rsp.Rows.Count;

            if (res == 0)
            {
                ex_message.BackColor = Color.Red;
                ex_message.ForeColor = Color.White;
                exam_table.BackgroundColor = Color.LightGray;
                ex_message.Text = "Экзамен ещё не запланирован по расписанию.";
                exam_table.Visible = false;
                exam_result.Visible = false;
                //return;
            }
            else
            {

                //получить дату экзамена
                DateTime zach_date = Convert.ToDateTime(rsp.Rows[0][0]);
                string rasp_id = rsp.Rows[0][1].ToString();

                ex_message.BackColor = statusStrip2.BackColor;
                ex_message.ForeColor = Color.Black;
                ex_message.Text = "Экзамен по расписанию: " + zach_date.ToShortDateString();
            }

            exam_table.BackgroundColor = Color.White;
            exam_table.Visible = true;
            exam_result.Visible = true;

            //заполнить список отметок

            //получить вид занятия
            query = string.Format(
                "select kod, vid_zan.id from predmet " +
                " join vidzan_predmet on vidzan_predmet.predmet_id = predmet.id " +
                " join vid_zan on vid_zan.id = vidzan_predmet.vidzan_id " +
                " where predmet.id = {0} and (vid_zan.kod = 'э')",
                id_predmet_in_tree);

            rsp = new DataTable();
            global_adapter = new SqlDataAdapter(query, global_connection);
            global_adapter.Fill(rsp);

            string vid = rsp.Rows[0][0].ToString();

            ///   ---- получить и вывести все виды отметок --------------
            otmetki = new DataTable();
            query = "select vid_otmetka.id, vid_otmetka.str_name " +
                " from vid_otmetka " +
                " join vid_zan_otmetka on vid_zan_otmetka.vid_otmetka_id = vid_otmetka.id " +
                " join vid_zan on vid_zan.id = vid_zan_otmetka.vid_zan_id " +
                " where vid_zan.id = " + rsp.Rows[0][1].ToString();
            global_adapter = new SqlDataAdapter(query, global_connection);
            global_adapter.Fill(otmetki);

            zach_dcell = new DataGridViewComboBoxCell();
            exam_table.Columns[1].CellTemplate = zach_dcell;

            exam_result.Columns.Clear(); //очистить сетку вывода статистки отметок
            exam_result.Rows.Clear();
            //очистить массив статистики отметок
            статистика_отметок.Clear();
            видотметки_количество vk = null;

            int i1 = 0;
            for (i1 = 0; i1 < otmetki.Rows.Count; i1++)
            {
                string nm = otmetki.Rows[i1][1].ToString();
                zach_dcell.Items.Add(nm);
                exam_result.Columns.Add("vd" + i1.ToString(), nm);
                exam_result.Columns[i1].SortMode = DataGridViewColumnSortMode.NotSortable;
                vk = null;
                vk = new видотметки_количество(nm, 0);
                статистика_отметок.Add(vk);
            }

            exam_result.Columns.Add("Итог", "Итого сдавали");
            exam_result.Columns[i1].SortMode = DataGridViewColumnSortMode.NotSortable;
            exam_result.Columns[i1].DefaultCellStyle.BackColor = Color.Pink;

            exam_result.Rows.Add();

            // ---------------------------------------------------

            query = string.Format("exec dbo.TGetSessionResult {0}, {1}, {2}",
                stud_set.Rows[0][1], id_predmet_in_tree, rsp.Rows[0][1]);
            выставленные_отметки = new DataTable();
            (new SqlDataAdapter(query, global_connection)).Fill(выставленные_отметки);

            for (int i2 = 0; i2 < выставленные_отметки.Rows.Count; i2++)
            {
                exam_table.Rows.Add();
                exam_table.Rows[i2].Cells[0].Value = выставленные_отметки.Rows[i2][1].ToString();
                int otz = Convert.ToInt32(выставленные_отметки.Rows[i2][3]);
                if (otz != -1)
                {
                    string nm = otmetka_name(otmetki, 0, otz);
                    exam_table.Rows[i2].Cells[1].Value = nm;
                    внести_отметку_в_статистику(nm);
                }
                exam_table.Rows[i2].Cells[2].Value = выставленные_отметки.Rows[i2][2];
            }

            //вывод значений в нижнюю таблицу статистики
            int vsego = 0;
            for (i = 0; i < статистика_отметок.Count; i++)
            {
                exam_result.Rows[0].Cells[i].Value = статистика_отметок[i].Kol;
                if (статистика_отметок[i].OtmName == "неявка") vsego -= статистика_отметок[i].Kol;
                if (статистика_отметок[i].OtmName == "недопуск") vsego -= статистика_отметок[i].Kol;
                vsego += статистика_отметок[i].Kol;
            }
            exam_result.Rows[0].Cells[i].Value = vsego.ToString();

            zachet_filled = true;
        }

        /// <summary>
        /// реакция на изменение экзаменацонной оценки
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void exam_table_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (exam_table.Rows.Count == 0) return;

            if (!zachet_filled) return;

            //сохранить оценку
            if (e.ColumnIndex == 1)
            {
                string sess_id = выставленные_отметки.Rows[e.RowIndex][0].ToString(), name = "";

                if (exam_table.Rows[e.RowIndex].Cells[1].Value!=null)
                    name = exam_table.Rows[e.RowIndex].Cells[1].Value.ToString();
                
                string q = "update session set otmetka_id = " +
                    otmetka_index(name) + " where id = " + sess_id;
                global_command = new SqlCommand(q, global_connection);
                global_command.ExecuteNonQuery();

                //MessageBox.Show(q);

                int i = 0, vsego = 0;
                for (i = 0; i < статистика_отметок.Count; i++)
                {
                    string n = статистика_отметок[i].OtmName;
                    статистика_отметок[i].Kol = 0;
                    for (int j = 0; j < exam_table.Rows.Count; j++)
                    {
                        if (exam_table.Rows[j].Cells[1].Value != null)
                        {
                            if (exam_table.Rows[j].Cells[1].Value.ToString() == n)
                                статистика_отметок[i].inc();
                        }
                    }

                    exam_result.Rows[0].Cells[i].Value = статистика_отметок[i].Kol.ToString();
                    if (статистика_отметок[i].OtmName == "неявка") vsego -= статистика_отметок[i].Kol;
                    if (статистика_отметок[i].OtmName == "недопуск") vsego -= статистика_отметок[i].Kol;
                    if (статистика_отметок[i].OtmName == "нет") vsego -= статистика_отметок[i].Kol;
                    vsego += статистика_отметок[i].Kol;
                }
                exam_result.Rows[0].Cells[i].Value = vsego.ToString();

                return;
            }

            //сохранить примечания
            if (e.ColumnIndex == 2)
            {
                string sess_id = выставленные_отметки.Rows[e.RowIndex][0].ToString(),
                    name = exam_table.Rows[e.RowIndex].Cells[2].Value.ToString();

                string q = "update session set ball = " + name + " where id = " + sess_id;
                global_command = new SqlCommand(q, global_connection);
                global_command.ExecuteNonQuery();
                return;
            } 
        }

        /// <summary>
        /// выставить МСА
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void teacher_tab_predmet_msa_Enter(object sender, EventArgs e)
        {
        }

        /// <summary>
        /// определить и вывести количество всех отметок
        /// </summary>
        /// <returns></returns>
        void count_all(DataGridView goal)
        {
            int c = 0;
            for (int i = 0; i < otmetki.Rows.Count; i++)
            {
                string nm = otmetki.Rows[i][1].ToString();
                c = 0;
                for (int j = 0; j < выставленные_отметки.Rows.Count; j++)
                {
                    string nm2 = выставленные_отметки.Rows[j][3].ToString();
                    if (nm == nm2) c++;
                }

                goal[i, 1].Value = c;
            }
        }

        private void копироватьВБуферToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Clipboard.Clear();
            try
            {
                Clipboard.SetDataObject(zachet_table.GetClipboardContent());
            }
            catch (Exception ex)
            {
                ;
            }
        }

        private void дляПервогоКурсаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveToExcel(1);
        }

        private void дляВторогоКурсаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveToExcel(2);
        }

        private void дляТретьегоКурсаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveToExcel(3);
        }

        private void дляЧетвертогоКурсаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveToExcel(4);
        }

        private void дляПятогоКурсаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveToExcel(5);
        }

        private void toolStripButton11_Click(object sender, EventArgs e)
        {
            SaveToExcel(1);
        }

        private void toolStripButton12_Click(object sender, EventArgs e)
        {
            SaveToExcel(2);
        }

        private void toolStripButton13_Click(object sender, EventArgs e)
        {
            SaveToExcel(3);
        }

        private void toolStripButton14_Click(object sender, EventArgs e)
        {
            SaveToExcel(4);
        }

        private void toolStripButton15_Click(object sender, EventArgs e)
        {
            SaveToExcel(5);
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            SaveToExcel();
        }

        private void toolStripButton16_Click(object sender, EventArgs e)
        {
            DataRow row = prepod_set.Rows[prepod_list.SelectedIndex];
            int prep_id = (int)row[0];
            string nm = prepod_list.Items[prepod_list.SelectedIndex].ToString();
            
            SaveToExcel(prep_id, nm);
        }


        /// <summary>
        /// ид темы выбранной работы
        /// </summary>
        public int tema_id = 0;
        public int vid_rabota_id = 0;

        /// <summary>
        /// выбрать тему курсовой работы
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripButton17_Click(object sender, EventArgs e)
        {
            if (kurs_table.Rows.Count == 0) return;

            

            //редактирование темы
            int row = 0;
            if (kurs_table.CurrentCell != null)
                row = kurs_table.CurrentCell.RowIndex;
            else
            {
                MessageBox.Show("Следует указать студента для выбора темы.",
                        "Ошибка операции", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            //string old_tema_id = "-1";
            //old_tema_id = выставленные_отметки.Rows[row][5].ToString();
            
            DataGridViewCell c = kurs_table.CurrentCell;

            //вызвать окно редактирования темы
            tema_edit te = new tema_edit();
            te.predmet_id = id_predmet_in_tree;
            te.vid_rab_id = vid_rabota_id;            
            te.Text = "Выбор темы КР по предмету: " + name_predmet_in_tree + " [" + 
                kurs_table.Rows[row].Cells[0].Value.ToString() + "]";

            DialogResult dr = te.ShowDialog();
            if (dr == DialogResult.Cancel) return;

            tema_id = te.new_tema_id;          
            
            //сохранить выбор темы для выбранного студента
            string q = "update student_rabota set tema_id = " +
                tema_id + " where id = " + выставленные_отметки.Rows[row][3].ToString();
            global_command = new SqlCommand(q, global_connection);
            global_command.ExecuteNonQuery();

            kurs_table.Rows[row].Cells[1].Value = te.new_tema_name;
            te.Dispose();

            teacher_tab_predmet_kursrab_Enter(sender, new EventArgs());            

            GC.Collect();
        }

        public int rabota_id = 0;
        //вывести студентов и их темы курсовой работы
        private void teacher_tab_predmet_kursrab_Enter(object sender, EventArgs e)
        {
            zachet_filled = false;

            //получить ид курсовой работы
            global_query = "select id from vid_rab where kod='кс'";
            global_command = new SqlCommand(global_query, global_connection);
            global_adapter = new SqlDataAdapter(global_command);
            DataTable vid_rab = new DataTable();
            global_adapter.Fill(vid_rab);
            vid_rabota_id = Convert.ToInt32(vid_rab.Rows[0][0]);
            vid_rab.Dispose();

            //получить всех студентов изучающих предмет
            string query = "select 	student.id " +                    
                    " from student  " +
                    " join grupa on grupa.id = student.gr_id " +
                    " join predmet on predmet.grupa_id = grupa.id  " +
                    " where student.actual=1 and fam<>'-' and predmet.id = " + id_predmet_in_tree.ToString() +
                    " order by fam, im, ot ";

            stud_set = new DataTable();
            global_adapter = new SqlDataAdapter(query, main.global_connection);
            global_adapter.Fill(stud_set);

            if (stud_set.Rows.Count == 0)
            {
                MessageBox.Show("Список студентов этой группы пуст. Необходимо создать список группы.",
                    "Операция невозможна",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            kurs_table.Rows.Clear();

            int i = 0;
            //foreach (DataRow dr in stud_set.Rows)
            //{
            //    kurs_table.Rows.Add();
            //    kurs_table.Rows[i].Cells[0].Value = dr[1];
            //    i++;
            //}

            //определить, запланирована ли защита курсовой работы
            query = string.Format(
                                " select dbo.get_date(y,m,d), rasp.id, semestr_id, vid_zan.kod from rasp " +
                " join vid_zan on vid_zan.id = rasp.vid_zan_id   " +
                " where predmet_id = {0} and   " +
                " ((y = {1} and m>=9 and m<=12) or (y={2} and m>=1 and m<=6)) and  " +
                " (vid_zan.kod = 'зкр')  and  " +
                " -semestr_id%2+2 = {3} ",
                id_predmet_in_tree, starts[0].Year, ends[week_list.Items.Count - 1].Year, semestr);
            rsp = new DataTable();
            global_adapter = new SqlDataAdapter(query, global_connection);
            global_adapter.Fill(rsp);

            int res = rsp.Rows.Count;            

            if (res == 0)
            {
                zachet_message.BackColor = Color.Red;
                zachet_message.ForeColor = Color.White;
                kurs_table.BackgroundColor = Color.LightGray;
                kurs_message.Text = "Защита курсовой работы ещё не запланирована по расписанию.";             
            }
            else
            {
                //получить дату защиты
                DateTime zach_date = Convert.ToDateTime(rsp.Rows[0][0]);
                string rasp_id = rsp.Rows[0][1].ToString();

                zachet_message.BackColor = statusStrip2.BackColor;
                zachet_message.ForeColor = Color.Black;
                kurs_message.Text = "Защита по расписанию: " + zach_date.ToShortDateString();
                kurs_table.BackgroundColor = Color.White;                
            }


            //заполнить список отметок

            //получить вид занятия
            query = string.Format(
                "select kod, vid_zan.id from predmet " +
                " join vidzan_predmet on vidzan_predmet.predmet_id = predmet.id " +
                " join vid_zan on vid_zan.id = vidzan_predmet.vidzan_id " +
                " where predmet.id = {0} and (vid_zan.kod = 'зкр')",
                id_predmet_in_tree);

            rsp = new DataTable();
            global_adapter = new SqlDataAdapter(query, global_connection);
            global_adapter.Fill(rsp);

            string vid = rsp.Rows[0][0].ToString();            

            ///   ---- получить и вывести все виды отметок --------------
            otmetki = new DataTable();
            query = "select vid_otmetka.id, vid_otmetka.str_name " +
                " from vid_otmetka " +
                " join vid_zan_otmetka on vid_zan_otmetka.vid_otmetka_id = vid_otmetka.id " +
                " join vid_zan on vid_zan.id = vid_zan_otmetka.vid_zan_id " +
                " where vid_zan.id = " + rsp.Rows[0][1].ToString();
            global_adapter = new SqlDataAdapter(query, global_connection);
            global_adapter.Fill(otmetki);

            zach_dcell = new DataGridViewComboBoxCell();
            kurs_table.Columns[2].CellTemplate = zach_dcell;

            kurs_result.Columns.Clear(); //очистить сетку вывода статистки отметок
            kurs_result.Rows.Clear();

            //очистить массив статистики отметок
            статистика_отметок.Clear();
            видотметки_количество vk = null;

            int i1 = 0;
            for (i1 = 0; i1 < otmetki.Rows.Count; i1++)
            {
                string nm = otmetki.Rows[i1][1].ToString();
                zach_dcell.Items.Add(nm);
                kurs_result.Columns.Add("vd" + i1.ToString(), nm);
                kurs_result.Columns[i1].SortMode = DataGridViewColumnSortMode.NotSortable;
                vk = null;
                vk = new видотметки_количество(nm, 0);
                статистика_отметок.Add(vk);
            }

            kurs_result.Columns.Add("итог", "Итого сдавали");
            kurs_result.Columns[i1].SortMode = DataGridViewColumnSortMode.NotSortable;
            kurs_result.Columns[i1].DefaultCellStyle.BackColor = Color.Pink;

            kurs_result.Rows.Add();

            //проверить, существует ли курсовая работа и создать её если нет
            query = string.Format("select rabota.id from rabota " +
                " join vid_rab on vid_rab.id = rabota.vid_rab_id " +
                " where predmet_id = {0} and prepod_id = {1} " +
                " and y = {2} and vid_rab.kod='кс'", 
                id_predmet_in_tree, active_user_id, ends[ends.Count - 1].Year);

            global_command = new SqlCommand(query, global_connection);
            DataTable rab_id = new DataTable();
            global_adapter = new SqlDataAdapter(global_command);
            global_adapter.Fill(rab_id);            

            if (rab_id.Rows.Count == 0)
            {
                query = "insert into rabota(prepod_id, predmet_id, vid_rab_id, name, status, y, poruch) " +
                    " values (@prepod_id, @predmet_id, @vid_rab_id, @name, @status, @y, @poruch)";
                global_command = new SqlCommand(query, global_connection);

                global_command.Parameters.Add("@prepod_id", SqlDbType.Int).Value = active_user_id;
                global_command.Parameters.Add("@predmet_id", SqlDbType.Int).Value = id_predmet_in_tree;
                global_command.Parameters.Add("@vid_rab_id", SqlDbType.Int).Value = vid_rabota_id;
                global_command.Parameters.Add("@name", SqlDbType.NVarChar).Value = "Курсовая работа: " + name_predmet_in_tree;
                global_command.Parameters.Add("@status", SqlDbType.Bit).Value = 0;
                global_command.Parameters.Add("@y", SqlDbType.Int).Value = ends[ends.Count - 1].Year;
                global_command.Parameters.Add("@poruch", SqlDbType.Bit).Value = true;

                try
                {
                    global_command.ExecuteNonQuery();
                }
                catch (Exception exx)
                {
                    MessageBox.Show("Невозможно выполнить операцию вследствие сетевой ошибки. Повторите попытку позднее.",
                        "Ошибка операции", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                query = "select @@identity";
                global_command = new SqlCommand(query, global_connection);
                SqlDataReader r = global_command.ExecuteReader();

                r.Read();
                rabota_id = Convert.ToInt32(r[0]);
                r.Close();

                //теперь в таблице student_rabota создать записи для созданной курсовой работы
                foreach (DataRow dr in stud_set.Rows)
                {
                    query = "insert into student_rabota (student_id, rabota_id, tema_id, pred_status) " +
                        " values (@student_id, @rabota_id, @tema_id, 0)";
                    global_command = new SqlCommand(query, global_connection);

                    global_command.Parameters.Add("@student_id", SqlDbType.Int).Value = dr[0];
                    global_command.Parameters.Add("@rabota_id", SqlDbType.Int).Value = rabota_id;
                    global_command.Parameters.Add("@tema_id", SqlDbType.Int).Value = -1;
                    global_command.ExecuteNonQuery();
                }
            }
            else
            {
                rabota_id = Convert.ToInt32(rab_id.Rows[0][0]);
            }

            rab_id.Dispose();

            
            //получить список студентов, их темы, их оценки                       
            выставленные_отметки = new DataTable();
            query = string.Format("select " + 
	            " student.id, isnull(vid_otmetka.id,-1), isnull(tema_rabota.name, ''),   " +  //0 1 2 
                " student_rabota.id, isnull(vid_otmetka.str_name,''), tema_id, isnull(tema_rabota.content, ''), " + //3 4 5 6
                " fio = student.fam + ' ' + left(student.im,1) + '. ' + left(student.ot,1) + '.', " + //7
                " ps = isnull(student_rabota.pred_status,0), rabota.id, isnull(vid_otmetka.str_alias,''),  " + //8  9 10
                " goal = isnull(tema_rabota.content,''), isnull(otzyv,''), " +  // 11 12
                " case when student.sex=1 then 'студента' else 'студентки' end, " + //13
                " grupa.name " + //14
                " from student_rabota " +
	            "   join rabota on rabota.id = student_rabota.rabota_id " +
	            "    join student on student.id = student_rabota.student_id " +
                "    join grupa on grupa.id = student.gr_id " + 
	            "    left outer join tema_rabota on tema_rabota.id = student_rabota.tema_id " +
	            "    left outer join vid_otmetka on vid_otmetka.id = student_rabota.otmetka_id " +
                " where rabota.id = {0} and student.actual = 1 " +
                " order by student.fam, student.im, student.ot", rabota_id);                                               

            global_adapter = new SqlDataAdapter(query, global_connection);
            global_adapter.Fill(выставленные_отметки);

            //MessageBox.Show(rabota_id.ToString() + "\n" + выставленные_отметки.Rows.Count.ToString());
           

            kurs_table.Columns[0].Tag = rabota_id;
            for (int i2 = 0; i2 < выставленные_отметки.Rows.Count; i2++)
            {
                kurs_table.Rows.Add();
                kurs_table.Rows[i2].Cells[0].Value = выставленные_отметки.Rows[i2][7].ToString();
                int otz = Convert.ToInt32(выставленные_отметки.Rows[i2][1]);
                kurs_table.Rows[i2].Cells[1].Value = выставленные_отметки.Rows[i2][2];
                kurs_table.Rows[i2].Tag = выставленные_отметки.Rows[i2][3];

                if (otz != -1)
                {
                    string nm = otmetka_name(otmetki, 0, otz);
                    kurs_table.Rows[i2].Cells[2].Value = nm;
                    внести_отметку_в_статистику(nm);
                }

                bool stat = Convert.ToBoolean(выставленные_отметки.Rows[i2][8]);
                if (stat)
                    kurs_table.Rows[i2].Cells[3].Value = "сдана";
                else
                    kurs_table.Rows[i2].Cells[3].Value = "не сдана";
                
            }

            //вывод значений в нижнюю таблицу статистики
            int vsego = 0;
            for (i = 0; i < статистика_отметок.Count; i++)
            {
                kurs_result.Rows[0].Cells[i].Value = статистика_отметок[i].Kol;
                if (статистика_отметок[i].OtmName == "неявка") vsego -= статистика_отметок[i].Kol;
                if (статистика_отметок[i].OtmName == "недопуск") vsego -= статистика_отметок[i].Kol;
                vsego += статистика_отметок[i].Kol;
            }
            kurs_result.Rows[0].Cells[i].Value = vsego.ToString();
            

            zachet_filled = true;
        }

        /// <summary>
        /// вызвать окно архива курсовых работ 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripButton31_Click(object sender, EventArgs e)
        {
            kursrab_archiv_edit krae = new kursrab_archiv_edit();
            krae.ShowDialog();
        }

        /// <summary>
        /// редактирование отзыва на работу студента
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripButton32_Click(object sender, EventArgs e)
        {
            if (kurs_table.Rows.Count == 0) return;

            if (kurs_table.SelectedCells.Count == 0) return;

            int row = kurs_table.SelectedCells[0].RowIndex;            
            string otz_txt = выставленные_отметки.Rows[row][12].ToString();
            string title = "Курсовая работа " + выставленные_отметки.Rows[row][13].ToString() + 
                ": " + kurs_table.Rows[row].Cells[0].Value.ToString();            

            inputTextBox itb = new inputTextBox();
            itb.kursRabOtzyvtextBox.Text = otz_txt;
            itb.Text = title;

            DialogResult dres;
            do
            {
                dres = itb.ShowDialog();
                if (dres == DialogResult.Cancel)
                    return;
            }
            while (dres != DialogResult.OK);

            /*if (itb.kursRabOtzyvtextBox.Text.Trim().Length == 0)
            {
                MessageBox.Show("Для сохранения требуется непустой текст отзыва. Изменения не сохранены",
                    "Отказ операции", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }*/

            string cmd = "update student_rabota set otzyv = @OTZ where id = @ID";
            global_command = new SqlCommand(cmd, global_connection);
            global_command.Parameters.Add("@OTZ", SqlDbType.VarChar).Value = itb.kursRabOtzyvtextBox.Text.Trim();
            global_command.Parameters.Add("@ID", SqlDbType.Int).Value = kurs_table.Rows[row].Tag;
            global_command.ExecuteNonQuery();

            teacher_tab_predmet_kursrab_Enter(sender, e);
            kurs_table.CurrentCell = kurs_table.Rows[row].Cells[0];

            itb.Dispose();
            GC.Collect();
        }

       /// <summary>
        /// редактирование шаблона отзыва на курсовую работу
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripButton30_Click(object sender, EventArgs e)
        {
            sprav_kurs_rabota skr = new sprav_kurs_rabota();
            skr.rab_id = выставленные_отметки.Rows[0][9].ToString();
            skr.predm_id = id_predmet_in_tree.ToString();

            DialogResult res;
            do
            {
                res = skr.ShowDialog();
                if (res == DialogResult.Cancel) break;
            }
            while (res != DialogResult.OK);
            
            skr.Dispose();
            GC.Collect();
        }

        
        /// <summary>
        /// построить отзыв на курсовую работу
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripButton29_Click(object sender, EventArgs e)
        {
            //MessageBox.Show(kurs_table.Rows.Count.ToString());
            int rows = kurs_table.Rows.Count;

            if (rows == 0)
            {
                MessageBox.Show("Нет сведений о курсовых работах.",
                    "Отказ операции", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (kurs_table.SelectedCells.Count == 0)
            {
                MessageBox.Show("Выберите строку с работой, для которой нужно построить отзыв.",
                    "Откза операции", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            string rab_id = kurs_table.Columns[0].Tag.ToString();
            string otm_digit = выставленные_отметки.Rows[kurs_table.SelectedCells[0].RowIndex][10].ToString();
            string tema = выставленные_отметки.Rows[kurs_table.SelectedCells[0].RowIndex][2].ToString();
            string stud = kurs_table.Rows[kurs_table.SelectedCells[0].RowIndex].Cells[0].Value.ToString();
            string prname = name_predmet_in_tree;
            string content = выставленные_отметки.Rows[kurs_table.SelectedCells[0].RowIndex][11].ToString();
            string stud_rab_otz = выставленные_отметки.Rows[kurs_table.SelectedCells[0].RowIndex][12].ToString();           

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


            Word.Document Doc = CreateWordDoc();

            string sql = "select * from rabota where id = " + rab_id;
            DataTable Rabota = new DataTable();
            (new SqlDataAdapter(sql, global_connection)).Fill(Rabota);

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
            Range.Text = "\"" + name_predmet + "\"";

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
            if (stud_rab_otz.Length != 0)
                Range.Text = stud_rab_otz;
            else
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

            active_user_dolz = active_user_dolz.Substring(0, 1).ToUpper() + active_user_dolz.Substring(1);

            Range.Text = active_user_dolz + " кафедры " + active_user_kaf + "                             " + active_user_name;
            

            SaveWordDoc(FileName, ref Doc);
            WordQuit();

            try
            {
                Process.Start(FileName);
            }
            catch (Exception eee)
            {
                MessageBox.Show("Невозможно открыть файл. Причина:\n" + eee.Message,
                    "Устраните указанную причину и повторите операцию.",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            Rabota.Dispose();
            GC.Collect();
        }

        /// <summary>
        /// вывод акта сдачи в архив для выделенных работ
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripButton33_Click(object sender, EventArgs e)
        {
            Word.Application wa = null;
            Word.Document doc = main.CreateNewWordDoc(ref wa);
            object nulval = Type.Missing;

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
                "За " + year_start.Year.ToString() + "/" + year_end.Year.ToString() + " уч. год",
                Word.WdParagraphAlignment.wdAlignParagraphLeft);


            Range = main.AddWordDocParagraph(ref doc,
                " ",
                Word.WdParagraphAlignment.wdAlignParagraphLeft);

            Range = main.AddWordDocParagraph(ref doc,
                "Группа " + выставленные_отметки.Rows[0][14].ToString().ToUpper(),
                Word.WdParagraphAlignment.wdAlignParagraphCenter);
            Range.Font.Bold = 1;

            Range = main.AddWordDocParagraph(ref doc,
                "Дисциплина “" + name_predmet + "”",
                Word.WdParagraphAlignment.wdAlignParagraphCenter);
            Range.Font.Bold = 1;

            Range = main.AddWordDocParagraph(ref doc,
                " ",
                Word.WdParagraphAlignment.wdAlignParagraphLeft);
            Range.Font.Bold = 0;

            
            List<int> tabrows = new List<int>();
            foreach (DataGridViewCell c in kurs_table.SelectedCells)
            {
                if (!tabrows.Contains(c.RowIndex))
                    tabrows.Add(c.RowIndex);
            }

            int k = 1;
            for (int i = 0; i < tabrows.Count; i++)
            {
                int ind = tabrows[i];

                if (kurs_table.Rows[ind].Cells[2].Value == null) continue;

                if ((kurs_table.Rows[ind].Cells[2].Value.ToString() == "недопуск" ||
                        kurs_table.Rows[ind].Cells[2].Value.ToString() == "неявка" ||
                        kurs_table.Rows[ind].Cells[2].Value.ToString() == "неудовлетворительно" ||
                        kurs_table.Rows[ind].Cells[2].Value.ToString().Length == 0))
                {
                    continue;
                }

                if (kurs_table.Rows[ind].Cells[3].Value.ToString() == "сдана")
                {

                    if (kurs_table.Rows[ind].Cells[1].Value.ToString().Length != 0)
                    {
                        Range = main.AddWordDocParagraph(ref doc,
                            k.ToString() + ". " + kurs_table.Rows[ind].Cells[0].Value.ToString() +
                            " (" + kurs_table.Rows[ind].Cells[1].Value.ToString() + ")",
                            Word.WdParagraphAlignment.wdAlignParagraphLeft);
                        k++;
                    }
                }
            }

            if (k == 1)
            {
                MessageBox.Show("Акт не построен. Среди выбранных Вами работ нет таких, " +
                    "которые могут быть актированы\n(либо у работ не указана тема,\nлибо поясн. записки выбранных работ не были сданы).",
                    "Отказ операции", MessageBoxButtons.OK, MessageBoxIcon.Error);
                main.WordQuit(wa);
                return;
            }


            saveExcel.Title = "Введите имя для файла акта курсовой работы.";
            saveExcel.Filter = "Файл акта КР в формате MS Word|*.doc";
            saveExcel.FileName = "Акт курс. работ по " + name_predmet_in_tree + ".doc";

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
                "Преподаватель           " + active_user_name + "___________________",
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

        /// <summary>
        /// вывод акта сдачи в архив
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripButton28_Click(object sender, EventArgs e)
        {
            Word.Application wa = null;
            Word.Document doc = main.CreateNewWordDoc(ref wa);
            object nulval = Type.Missing;            

            Word.Range Range = doc.Range(ref nulval, ref nulval);
            Range.Select();
            Range.ParagraphFormat.FirstLineIndent = 0.0f;

            Range = doc.Paragraphs[1].Range;
            Range.Select();
            Range.Text = "А К Т";
            Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            Range.Font.Bold = 1;

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
                "За " + year_start.Year.ToString() + "/" + year_end.Year.ToString() + " уч. год",
                Word.WdParagraphAlignment.wdAlignParagraphLeft);


            Range = main.AddWordDocParagraph(ref doc,
                " ",
                Word.WdParagraphAlignment.wdAlignParagraphLeft);

            Range = main.AddWordDocParagraph(ref doc,
                "Группа " + выставленные_отметки.Rows[0][14].ToString().ToUpper(),
                Word.WdParagraphAlignment.wdAlignParagraphCenter);
            Range.Font.Bold = 1;

            Range = main.AddWordDocParagraph(ref doc,
                "Дисциплина “" + name_predmet + "”",
                Word.WdParagraphAlignment.wdAlignParagraphCenter);
            Range.Font.Bold = 1;

            Range = main.AddWordDocParagraph(ref doc,
                " ",
                Word.WdParagraphAlignment.wdAlignParagraphLeft);
            Range.Font.Bold = 0;

            int k = 1;
            for (int i = 0; i < kurs_table.Rows.Count; i++)
            {

                if (kurs_table.Rows[i].Cells[2].Value == null) continue;
                
                if ((kurs_table.Rows[i].Cells[2].Value.ToString() == "недопуск" ||
                    kurs_table.Rows[i].Cells[2].Value.ToString() == "неявка" ||
                    kurs_table.Rows[i].Cells[2].Value.ToString() == "неудовлетворительно" ||
                    kurs_table.Rows[i].Cells[2].Value.ToString().Length == 0))
                {
                    continue;
                }

                if (kurs_table.Rows[i].Cells[3].Value.ToString() == "сдана")
                {
                    if (kurs_table.Rows[i].Cells[1].Value.ToString().Length != 0)
                    {
                        Range = main.AddWordDocParagraph(ref doc,
                            k.ToString() + ". " + kurs_table.Rows[i].Cells[0].Value.ToString() +
                            " (" + kurs_table.Rows[i].Cells[1].Value.ToString() + ")",
                            Word.WdParagraphAlignment.wdAlignParagraphLeft);
                        k++;
                    }
                }
            }

            if (k == 1)
            {
                MessageBox.Show("Акт не построен. Среди выбранных Вами работ нет таких, " +
                    "которые могут быть актированы\n(либо у работ не указана тема,\nлибо поясн. записки выбранных работ не были сданы).",
                    "Отказ операции", MessageBoxButtons.OK, MessageBoxIcon.Error);
                main.WordQuit(wa);
                return;
            }


            saveExcel.Title = "Введите имя для файла акта курсовой работы.";
            saveExcel.Filter = "Файл акта КР в формате MS Word|*.doc";
            saveExcel.FileName = "Акт курс. работ по " + name_predmet_in_tree + ".doc";

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
                "Преподаватель           " + active_user_name + "___________________",
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
            //MessageBox.Show(kurs_table.Rows.Count.ToString());
        }

        /// <summary>
        /// вызвать окно отображения архива курсовых работ
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripButton27_Click(object sender, EventArgs e)
        {
            kursrab_archiv_edit krae = new kursrab_archiv_edit();
            krae.ShowDialog();
            krae.Dispose();
            GC.Collect();
        }

        //удаление студента из списка курсовых работа (в случае его выбывания)
        private void удалитьСтудентаИзСпискаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int row = kurs_table.SelectedCells[0].RowIndex;
            string stud_id = выставленные_отметки.Rows[row][0].ToString();
            string rabota_id = выставленные_отметки.Rows[row][9].ToString();


            if (MessageBox.Show("Данные студента \"" + kurs_table.Rows[row].Cells[0].Value.ToString() + "\"\n" +
                "будут удалены из списка курсовых работ. Продолжить?", "Предупреждение",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No) return;

            string cmd = "delete from student_rabota where student_id = @STUDID and rabota_id=@RABID";
            global_command = new SqlCommand(cmd, global_connection);

            global_command.Parameters.Add("@STUDID", SqlDbType.Int).Value = stud_id;
            global_command.Parameters.Add("@RABID", SqlDbType.Int).Value = rabota_id;

            global_command.ExecuteNonQuery();

            teacher_tab_predmet_kursrab_Enter(sender, e);
        }

        /// <summary>
        /// сохранить оценку за КР
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void kurs_table_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (kurs_table.Rows.Count == 0) return;

            if (!zachet_filled) return;

            //сохранить оценку
            if (e.ColumnIndex == 2)
            {
                string stud_rab_id = выставленные_отметки.Rows[e.RowIndex][3].ToString(),
                    name = kurs_table.Rows[e.RowIndex].Cells[2].Value.ToString(),
                    nametema = kurs_table.Rows[e.RowIndex].Cells[1].Value.ToString();

                if (nametema.Trim().Length == 0)
                {
                    MessageBox.Show("Следует указать тему перед выставлением оценки. Оценка не была сохранена.",
                        "Ошибка операции", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                string q = "update student_rabota set otmetka_id = " +
                    otmetka_index(name) + " where id = " + stud_rab_id;
                global_command = new SqlCommand(q, global_connection);
                global_command.ExecuteNonQuery();

                teacher_tab_predmet_kursrab_Enter(sender, e);

                int i = 0, vsego = 0;
                for (i = 0; i < статистика_отметок.Count; i++)
                {
                    string n = статистика_отметок[i].OtmName;
                    статистика_отметок[i].Kol = 0;
                    for (int j = 0; j < kurs_table.Rows.Count; j++)
                    {
                        if (kurs_table.Rows[j].Cells[2].Value != null)
                        {
                            if (kurs_table.Rows[j].Cells[2].Value.ToString() == n)
                                статистика_отметок[i].inc();
                        }
                    }

                    kurs_result.Rows[0].Cells[i].Value = статистика_отметок[i].Kol.ToString();
                    if (статистика_отметок[i].OtmName == "неявка") vsego -= статистика_отметок[i].Kol;
                    if (статистика_отметок[i].OtmName == "недопуск") vsego -= статистика_отметок[i].Kol;
                    vsego += статистика_отметок[i].Kol;
                }
                kurs_result.Rows[0].Cells[i].Value = vsego.ToString();

                return;
            }

            //сохранить примечания
            if (e.ColumnIndex == 1)
            {
                return;
            }           
        }

        /// <summary>
        /// редактирование темы курсовой работы
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripButton18_Click(object sender, EventArgs e)
        {
            //редактирование темы
            int row = 0;
            if (kurs_table.CurrentCell != null)
                row = kurs_table.CurrentCell.RowIndex;
            else
                return;

            if (kurs_table.Rows[row].Cells[1].Value.ToString().Length == 0) return;

            string tm = выставленные_отметки.Rows[row][2].ToString();
            string cnt = выставленные_отметки.Rows[row][6].ToString();
            int id = Convert.ToInt32(выставленные_отметки.Rows[row][5]);

            tema_add ta = new tema_add();
            ta.tema.Text = tm;
            ta.cont.Text = cnt;

            DialogResult res = ta.ShowDialog();
            if (res == DialogResult.Cancel) return;

            tm = Normalize1(ta.tema.Text.Trim());
            cnt = Normalize1(ta.cont.Text.Trim());

            string q = string.Format(
                "update tema_rabota set name = '{0}', content = '{1}', predmet_id = {2}, vid_rabota_id = {3} " +
                " where id = {4}", tm, cnt, id_predmet_in_tree, vid_rabota_id, id);
            main.global_command = new SqlCommand(q, main.global_connection);

            try
            {
                main.global_command.ExecuteNonQuery();
            }
            catch (Exception exx)
            {
                MessageBox.Show("Невозможно выполнить операцию вследствие сетевой ошибки. Повторите попытку позднее.",
                    "Ошибка операции", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            
            ta.Dispose();

            kurs_table.Rows[row].Cells[1].Value = tm;
        }

        private void копироватьВБкферОбменаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Clipboard.Clear();
            try
            {
                Clipboard.SetDataObject(kurs_table.GetClipboardContent());
            }
            catch (Exception ex)
            {
                ;
            }
        }

        private void задатьТемуДляЭтогоСтуднтаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            toolStripButton17_Click(sender, new EventArgs());
        }

        private void редактироватьТемуЭтогоСтудентаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            toolStripButton18_Click(sender, new EventArgs());
        }

        private void toolStripMenuItem16_Click(object sender, EventArgs e)
        {
            Clipboard.Clear();
            try
            {
                Clipboard.SetDataObject(exam_table.GetClipboardContent());
            }
            catch (Exception ex)
            {
                ;
            }
        }


        public int[] mes = new int[10] { 9, 10, 11, 12, 1, 2, 3, 4, 5, 6 };
        public int porog = 4; //номер последнего элемента, входящего в первое полугодие
        public string[] vids = new string[] { "'л'", "'пр'", "'с'", "'лр'", "'ксп'", "'кпэ'", "'э'", "'з'" };
        public DataTable personal_tabble1;
        /// <summary>
        /// заполнение контрольного листа
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void teacher_tab_kontrlist_Enter(object sender, EventArgs e)
        {
            kontrol_list_1.Rows.Clear();
            kontrol_list_1.Rows.Add("План на I п.");
            kontrol_list_1.Rows.Add("сентябрь");
            kontrol_list_1.Rows.Add("октрябрь");
            kontrol_list_1.Rows.Add("ноябрь");
            kontrol_list_1.Rows.Add("декабрь");
            kontrol_list_1.Rows.Add("январь");
            kontrol_list_1.Rows.Add("Всего за I п.");

            kontrol_list_2.Rows.Clear();
            kontrol_list_2.Rows.Add("План на II п.");
            kontrol_list_2.Rows.Add("февраль");
            kontrol_list_2.Rows.Add("март");
            kontrol_list_2.Rows.Add("апрель");
            kontrol_list_2.Rows.Add("май");
            kontrol_list_2.Rows.Add("июнь");
            kontrol_list_2.Rows.Add("Всего за II п.");

            kontrol_list_res.Rows.Clear();
            kontrol_list_res.Rows.Add("План на год");
            kontrol_list_res.Rows.Add("Всего за год");
            kontrol_list_res.Rows.Add("Разница");

            //---------------------------  получить расписание
            //загрузка расписания
            global_query =
            " select dbo.get_date(y,m,d), " + //0                        
            " vid_zan_id, rasp.id, rasp.kol_chas, vk = vid_zan.kod, m " + //1 2 3 4 
            " from rasp " +           
            " join vid_zan on vid_zan.id = rasp.vid_zan_id " +
            " where rasp.prepod_id = @USERID " +
            //-------- !!!! 
            //" and dbo.get_date(y,m,d)>=@D1 and dbo.get_date(y,m,d)<=@D2 " + 
            " and rasp.uch_god_id = @UGOD" + //uch_god +
            " order by y, m, d, nom_zan,vid_zan.kod ";

            global_command = new SqlCommand(global_query, global_connection);
            global_command.Parameters.Add("@USERID", SqlDbType.Int).Value = active_user_id;
            global_command.Parameters.Add("@D1", SqlDbType.DateTime).Value = starts[0].Date;
            global_command.Parameters.Add("@D2", SqlDbType.DateTime).Value = ends[ends.Count - 1].Date;
            global_command.Parameters.Add("@UGOD", SqlDbType.Int).Value = uch_god;

            personal_tabble1 = new DataTable();
            global_adapter = new SqlDataAdapter(global_command);
            global_adapter.Fill(personal_tabble1);

            fill_kontr_list(kontrol_list_1, 0, porog, 1);
            fill_kontr_list(kontrol_list_2, porog+1, 9, 0);

            //вывод плана    ----------------------------  диплом + ВКР --------------------         
            //по дипломным и ВКР
            //dipl_plan*dipl_chas + <<<<<---- это часть для получения суммы ВКР и дипломных работ
            global_query = "select vkr_plan*vkr_chas + dipl_plan*dipl_chas from prepod " +
                " where id = " + active_user_id.ToString();
            global_adapter = new SqlDataAdapter(global_query, global_connection);
            DataTable zach = new DataTable();
            global_adapter.Fill(zach);
            int s = Convert.ToInt32(zach.Rows[0][0]);
            kontrol_list_res[13, 0].Value = s;


            //получить Факт по дипл и ВКР работы
            zach = new DataTable();
            global_query = "select dipl_chas from prepod where id = " + active_user_id.ToString();
            global_adapter = new SqlDataAdapter(global_query, global_connection);
            global_adapter.Fill(zach);
            int dipchas = Convert.ToInt32(zach.Rows[0][0]);
            //MessageBox.Show(dipchas.ToString());

            zach = new DataTable();
            global_query = "select vkr_chas from prepod where id = " + active_user_id.ToString();
            global_adapter = new SqlDataAdapter(global_query, global_connection);
            global_adapter.Fill(zach);
            int vkrchas = Convert.ToInt32(zach.Rows[0][0]);

            zach = new DataTable();
            global_query = "select rabota.id from rabota " +
                " join vid_rab on vid_rab.id = rabota.vid_rab_id " +
                " join student_rabota on student_rabota.rabota_id = rabota.id " +
                " join vid_otmetka on vid_otmetka.id = student_rabota.otmetka_id " +
            " where vid_rab.kod = 'вкр' and student_rabota.ruk_id = @USERID and student_rabota.pred_status = 1 and  " +
            " (vid_otmetka.name = 2 or vid_otmetka.name = 3 or vid_otmetka.name = 4 or vid_otmetka.name = 5)  " +
            " and rabota.y = @D2";

            global_command = new SqlCommand(global_query, global_connection);
            global_command.Parameters.Add("@USERID", SqlDbType.Int).Value = active_user_id;
            global_command.Parameters.Add("@D2", SqlDbType.Int).Value = ends[ends.Count - 1].Year;
            global_adapter = new SqlDataAdapter(global_command);
            global_adapter.Fill(zach);

            int sum_vkr = zach.Rows.Count * vkrchas;

            zach = new DataTable();
            global_query = "select rabota.id from rabota " +
                " join vid_rab on vid_rab.id = rabota.vid_rab_id " +
                " join student_rabota on student_rabota.rabota_id = rabota.id " +
                " join vid_otmetka on vid_otmetka.id = student_rabota.otmetka_id " +
            " where vid_rab.kod = 'др' and student_rabota.ruk_id = @USERID and student_rabota.pred_status = 1 and  " +
            " (vid_otmetka.name = 2 or vid_otmetka.name = 3 or vid_otmetka.name = 4 or vid_otmetka.name = 5)  " +
            " and rabota.y = @D2";

            global_command = new SqlCommand(global_query, global_connection);
            global_command.Parameters.Add("@USERID", SqlDbType.Int).Value = active_user_id;
            global_command.Parameters.Add("@D2", SqlDbType.Int).Value = ends[ends.Count - 1].Year;
            global_adapter = new SqlDataAdapter(global_command);
            global_adapter.Fill(zach);

            int sum_dip = zach.Rows.Count * dipchas;
            
            //----------------------------  диплом + ВКР -------------------------------
            kontrol_list_res[13, 1].Value = (sum_vkr + sum_dip).ToString();
            

            //вывод итогов за оба полугодия
            for (int i = 1; i < kontrol_list_res.Columns.Count; i++)
            {
                double sum_itog = Convert.ToDouble(kontrol_list_1[i, 0].Value) + Convert.ToDouble(kontrol_list_2[i, 0].Value);
                double sum_fakt = Convert.ToDouble(kontrol_list_1[i, 6].Value) + Convert.ToDouble(kontrol_list_2[i, 6].Value);

                if (i != 13)  //пропустить ВКР
                {
                    kontrol_list_res[i, 0].Value = sum_itog;
                    kontrol_list_res[i, 1].Value = sum_fakt;
                }
            }

            kontrol_list_res[1, 0].Value = Convert.ToDouble(kontrol_list_res[1, 0].Value) +
                Convert.ToDouble(kontrol_list_res[13, 0].Value);

            kontrol_list_res[1, 1].Value = Convert.ToDouble(kontrol_list_res[1, 1].Value) +
                Convert.ToDouble(kontrol_list_res[13, 1].Value);

            for (int i = 1; i < kontrol_list_res.Columns.Count; i++)
            {
                double plan = Convert.ToDouble(kontrol_list_res[i, 0].Value);
                double fakt = Convert.ToDouble(kontrol_list_res[i, 1].Value);
                double diff = plan - fakt;

                kontrol_list_res[i, 2].Value = diff;
                if (diff >= -2 && diff <= 2) 
                    kontrol_list_res[i, 2].Style.Font = new Font("Microsoft Sans Serif", 8.0f, FontStyle.Bold);
                
                if (diff < -2)
                {
                    kontrol_list_res[i, 2].Style.BackColor = Color.LightGray;
                    kontrol_list_res[i, 2].Style.ForeColor = Color.Blue;
                }

                if (diff > 2)
                {
                    kontrol_list_res[i, 2].Style.BackColor = Color.Pink;
                    kontrol_list_res[i, 2].Style.ForeColor = Color.Red;
                }
            }
        }

        public void fill_kontr_list(DataGridView kl, int start, int fin, int sem)
        {           
            int row = 0;
            for (int i = 0; i < vids.Length-2; i++)
            {
                row = 1;
                for (int j = start; j <= fin; j++)
                {
                    string select = string.Format(" vk = {0} and m = {1} ", vids[i], mes[j]);

                    DataRow[] data = personal_tabble1.Select(select);
                    int summa = 0;
                    foreach (DataRow dr in data)
                    {
                        int chas = Convert.ToInt32(dr[3]);
                        summa+=chas;
                    }

                    kl.Rows[row].Cells[i + 2].Value = summa.ToString();
                    row++;
                }
            }

            //вычислить зачеты и экзамены
          
            DataTable zach = new DataTable();
            global_query = "select koef from vid_zan where kod = 'з'";
            global_adapter = new SqlDataAdapter(global_query, global_connection);
            global_adapter.Fill(zach);
            double koef = Convert.ToDouble(zach.Rows[0][0]);

            zach = new DataTable();
            global_query = "select session.id, rm = rasp.m from session " + 
	                " join vid_zan on vid_zan.id = session.vid_zan_id " + 
	                " join vid_otmetka on vid_otmetka.id = session.otmetka_id " + 
	                " join rasp on rasp.id = session.rasp_id " +
                " where vid_zan.kod = 'з' and (vid_otmetka.name = 1 or vid_otmetka.name = 0) " + 
                " and dbo.get_date(y,m,d)>=@D1 and dbo.get_date(y,m,d)<=@D2 " +
                " and rasp.prepod_id = @USERID";

            global_command = new SqlCommand(global_query, global_connection);
            global_command.Parameters.Add("@USERID", SqlDbType.Int).Value = active_user_id;
            global_command.Parameters.Add("@D1", SqlDbType.DateTime).Value = starts[0].Date;
            global_command.Parameters.Add("@D2", SqlDbType.DateTime).Value = ends[ends.Count - 1].Date;            
            global_adapter = new SqlDataAdapter(global_command);
            global_adapter.Fill(zach);

            row = 1;
            for (int j = start; j <= fin; j++)
            {
                string select = string.Format(" rm = {0} ", mes[j]);

                DataRow[] data = zach.Select(select);
                double summa = data.Length * koef;

                kl.Rows[row].Cells[9].Value = summa.ToString("F2");
                row++;
            }

            // получить экзамены 
            zach = new DataTable();
            global_query = "select koef from vid_zan where kod = 'э'";
            global_adapter = new SqlDataAdapter(global_query, global_connection);
            global_adapter.Fill(zach);
            koef = Convert.ToDouble(zach.Rows[0][0]);

            zach = new DataTable();
            global_query = "select session.id, rm = rasp.m from session " +
                    " join vid_zan on vid_zan.id = session.vid_zan_id " +
                    " join vid_otmetka on vid_otmetka.id = session.otmetka_id " +
                    " join rasp on rasp.id = session.rasp_id " +
                " where vid_zan.kod = 'э' and " +
                "(vid_otmetka.name = 2 or vid_otmetka.name = 3 or vid_otmetka.name = 4 " + 
                " or vid_otmetka.name = 5) " +
                " and dbo.get_date(y,m,d)>=@D1 and dbo.get_date(y,m,d)<=@D2 " +
                " and rasp.prepod_id = @USERID";

            global_command = new SqlCommand(global_query, global_connection);
            global_command.Parameters.Add("@USERID", SqlDbType.Int).Value = active_user_id;
            global_command.Parameters.Add("@D1", SqlDbType.DateTime).Value = starts[0].Date;
            global_command.Parameters.Add("@D2", SqlDbType.DateTime).Value = ends[ends.Count - 1].Date;
            global_adapter = new SqlDataAdapter(global_command);
            global_adapter.Fill(zach);

            row = 1;
            for (int j = start; j <= fin; j++)
            {
                string select = string.Format(" rm = {0} ", mes[j]);

                DataRow[] data = zach.Select(select);
                double summa = data.Length * koef;

                kl.Rows[row].Cells[8].Value = summa.ToString("F2");
                row++;
            }

            //получить курсовые работы
            zach = new DataTable();
            global_query = "select koef from vid_zan where kod = 'зкр'";
            global_adapter = new SqlDataAdapter(global_query, global_connection);
            global_adapter.Fill(zach);
            int ikoef = Convert.ToInt32(zach.Rows[0][0]);

            zach = new DataTable();
            global_query = "select rabota.id from rabota " +
                " join vid_rab on vid_rab.id = rabota.vid_rab_id " +
                " join student_rabota on student_rabota.rabota_id = rabota.id " +
                " join vid_otmetka on vid_otmetka.id = student_rabota.otmetka_id " +
                " join predmet on predmet.id = rabota.predmet_id" + 
            " where vid_rab.kod = 'кс' and rabota.prepod_id = @USERID and student_rabota.pred_status = 1 and  " +
            " (vid_otmetka.name = 2 or vid_otmetka.name = 3 or vid_otmetka.name = 4 or vid_otmetka.name = 5)  " +
            " and rabota.y = @D2 and (predmet.semestr+2) % 2 = @sem";

            //textBox1.Text = string.Format("select rabota.id from rabota " +
            //    " join vid_rab on vid_rab.id = rabota.vid_rab_id " +
            //    " join student_rabota on student_rabota.rabota_id = rabota.id " +
            //    " join vid_otmetka on vid_otmetka.id = student_rabota.otmetka_id " +
            //    " join predmet on predmet.id = rabota.predmet_id" +
            //" where vid_rab.kod = 'кс' and rabota.prepod_id = {0} and student_rabota.pred_status = 1 and  " +
            //" (vid_otmetka.name = 2 or vid_otmetka.name = 3 or vid_otmetka.name = 4 or vid_otmetka.name = 5)  " +
            //" and rabota.y = {1} and (predmet.semestr+2) % 2 = {2}", active_user_id, ends[ends.Count - 1].Year, sem);

            global_command = new SqlCommand(global_query, global_connection);
            global_command.Parameters.Add("@USERID", SqlDbType.Int).Value = active_user_id;
            global_command.Parameters.Add("@D2", SqlDbType.Int).Value = ends[ends.Count - 1].Year;
            global_command.Parameters.Add("@sem", SqlDbType.Int).Value = sem;
            global_adapter = new SqlDataAdapter(global_command);
            global_adapter.Fill(zach);
            
            int s = zach.Rows.Count * ikoef;
            kl.Rows[6].Cells[10].Value = s.ToString();
            //MessageBox.Show(s.ToString());

            //вывести план по курсовым
            global_query = "select kurs_plan" +  (-sem+2).ToString() + 
                " from prepod " +
                " where id = " + active_user_id.ToString();
            global_adapter = new SqlDataAdapter(global_query, global_connection);
            zach = new DataTable();
            global_adapter.Fill(zach);
            s = Convert.ToInt32(zach.Rows[0][0]);
            kl[10, 0].Value = s;


            //вывести план по предметам
            //получить все предметы данного препода за данный семестр
            zach = new DataTable();
            global_query = "select vid_zan.id, vk = vid_zan.kod, " + 
                     "chas = case   " + 
                     "                when  type_id = 1  then    " + 
                     "                case   " + 
                     "                     when is_kontrol = 0 then   " + 
                     "                     case                       " +            
                     "                         when vid_zan.kod = 'ксп' then round(kol_chas,2)    " + 
                     "                     else                 " +                
                     "                         round(kol_chas - kol_chas*vid_zan.spisanie/100, 2)   " + 
                     "                     end   " +
                     "                     when is_kontrol = 1 then round(kol_chas - kol_chas*vid_zan.spisanie/100, 2) " + 
                     "                end " + 
                     "                else " +   
                     "                   kol_chas   " + 
                     "            end " + 
                     " from predmet " + 
	                 "    join vidzan_predmet on  vidzan_predmet.predmet_id = predmet.id  " + 
	                 "    join vid_zan on vid_zan.id = vidzan_predmet.vidzan_id " + 
	                 "    join predmet_type on predmet_type.id = predmet.type_id " +
                     " where predmet_type.id<=3 and predmet.prepod_id = @prid and ((predmet.semestr+2)%2=@sem)  " + 
                     " and predmet.actual = 1";
       
            global_command = new SqlCommand(global_query, global_connection);
            global_command.Parameters.Add("@prid", SqlDbType.Int).Value = active_user_id;
            global_command.Parameters.Add("@sem", SqlDbType.Int).Value = sem;
            global_adapter = new SqlDataAdapter(global_command);
            global_adapter.Fill(zach);

            for (int j = 0; j <vids.Length; j++)
            {
                string select = string.Format(" vk = {0} ", vids[j]);

                DataRow[] data = zach.Select(select);
                double summa = 0;
                foreach (DataRow dr in data)
                {
                    double chas = Convert.ToDouble(dr[2]);
                    summa += chas;
                }
                
                kl.Rows[0].Cells[j+2].Value =  Math.Truncate(summa).ToString();                
            }

            //подвести итого
            for (int i = 2; i < kl.Columns.Count; i++)
            {
                double sum_res = 0;
                for (int j = 1; j <= 5; j++)
                {
                    double chas_res = Convert.ToDouble(kl[i, j].Value);
                    sum_res += chas_res;
                }

                kl[i, 6].Style.BackColor = Color.LightYellow;
                kl[i, 6].Style.ForeColor = Color.Red;
                if (i!=10)
                kl[i, 6].Value = sum_res.ToString();
            }

            for (int j = 0; j <= 6; j++)
            {
                double sum_res = 0;
                for (int i = 2; i < kl.Columns.Count; i++)                
                {
                    double chas_res = Convert.ToDouble(kl[i, j].Value);
                    sum_res += chas_res;
                }

                kl[1, j].Style.BackColor = Color.Yellow;
                kl[1, j].Style.ForeColor = Color.Red;
                kl[1, j].Style.Font = new Font("Microsoft Sans Serif", 8.0f, FontStyle.Bold);
                kl[1, j].Value = sum_res.ToString();
            }
        }

        private void kurs_table_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (kurs_table.Rows.Count == 0) return;

            int row = 0;
            if (kurs_table.CurrentCell != null)
                row = e.RowIndex;
            else
                return;  

            if (e.ColumnIndex == 3)
            {
                int new_stat = 0;
                string old_stat = kurs_table.Rows[row].Cells[3].Value.ToString(),
                    new_str_stat = "";

                if (old_stat == "сдана")
                {
                    new_str_stat = "не сдана";
                    new_stat = 0;
                }
                else
                {
                    new_str_stat = "сдана";
                    new_stat = 1;
                }

                kurs_table.Rows[row].Cells[3].Value = new_str_stat;
                string stud_rab_id = выставленные_отметки.Rows[row][3].ToString();

                string q = "update student_rabota set pred_status = " +
                    new_stat.ToString() + " where id = " + stud_rab_id;
                global_command = new SqlCommand(q, global_connection);
                global_command.ExecuteNonQuery();
            }

            if (e.ColumnIndex == 1)
            {
                if (kurs_table.SelectedCells.Count == 0) return;

                string otz_txt = выставленные_отметки.Rows[row][12].ToString();
                string title = "Курсовая работа " + выставленные_отметки.Rows[row][13].ToString() +
                    ": " + kurs_table.Rows[row].Cells[0].Value.ToString();

                inputTextBox itb = new inputTextBox();
                itb.kursRabOtzyvtextBox.Text = otz_txt;
                itb.Text = title;

                DialogResult dres;
                do
                {
                    dres = itb.ShowDialog();
                    if (dres == DialogResult.Cancel)
                        return;
                }
                while (dres != DialogResult.OK);
                
                string cmd = "update student_rabota set otzyv = @OTZ where id = @ID";
                global_command = new SqlCommand(cmd, global_connection);
                global_command.Parameters.Add("@OTZ", SqlDbType.VarChar).Value = itb.kursRabOtzyvtextBox.Text.Trim();
                global_command.Parameters.Add("@ID", SqlDbType.Int).Value = kurs_table.Rows[row].Tag;
                global_command.ExecuteNonQuery();

                teacher_tab_predmet_kursrab_Enter(sender, e);
                kurs_table.CurrentCell = kurs_table.Rows[row].Cells[0];
            }

        }

        private void задатьКоличествоЧасовПоПлануToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //план по курсовым для первого полугодия
            inputbox ib = new inputbox("Введите общий план по курсовым работам из вашего поручения",
                   "План по курс. раб.", "0", "Задайте число");
            DialogResult res = ib.ShowDialog();

            //сохранить в БД
            global_query =
                "update prepod set kurs_plan1 = @kplan where id = @prid";
            global_command = new SqlCommand(global_query, global_connection);
            global_command.Parameters.Add("@kplan", SqlDbType.Int).Value = ib.textBox1.Value;
            global_command.Parameters.Add("@prid", SqlDbType.Int).Value = active_user_id;
            global_command.ExecuteNonQuery();

            //вывести в таблицу
            kontrol_list_1.Rows[0].Cells[10].Value = ib.textBox1.Value.ToString();                       
            
            ib.Dispose();            
        }

        private void задатьКоличествоЧасовВКРИДипломныхРаботПоПлануToolStripMenuItem_Click(object sender, EventArgs e)
        {
            vkr_diplom_plan vdp = new vkr_diplom_plan();
            DialogResult res = vdp.ShowDialog();

            if (res == DialogResult.Cancel)
            {
                vdp.Dispose();
                return;
            }

            //сохранить в БД
            global_query =
                "update prepod set vkr_plan = @vstud, vkr_chas=@vchas, " +
                " dipl_plan = @dstud, dipl_chas = @dchas " +
                " where id = @prid";
            global_command = new SqlCommand(global_query, global_connection);
            global_command.Parameters.Add("@vstud", SqlDbType.Int).Value = vdp.vkr_student.Value;
            global_command.Parameters.Add("@vchas", SqlDbType.Int).Value = vdp.vkr_chas.Value;
            global_command.Parameters.Add("@dstud", SqlDbType.Int).Value = vdp.dip_student.Value;
            global_command.Parameters.Add("@dchas", SqlDbType.Int).Value = vdp.dip_chas.Value;
            global_command.Parameters.Add("@prid", SqlDbType.Int).Value = active_user_id;
            global_command.ExecuteNonQuery();


            //вывести в таблицу
            kontrol_list_res.Rows[0].Cells[13].Value =
                vdp.vkr_student.Value * vdp.vkr_chas.Value + vdp.dip_student.Value * vdp.dip_chas.Value;

            vdp.Dispose();

        }

        private void toolStripMenuItem17_Click(object sender, EventArgs e)
        {
            //план по курсовым для 2 полугодия
            inputbox ib = new inputbox("Введите общий план по курсовым работам из вашего поручения",
                   "План по курс. раб.", "0", "Задайте число");
            DialogResult res = ib.ShowDialog();

            //сохранить в БД
            global_query =
                "update prepod set kurs_plan2 = @kplan where id = @prid";
            global_command = new SqlCommand(global_query, global_connection);
            global_command.Parameters.Add("@kplan", SqlDbType.Int).Value = ib.textBox1.Value;
            global_command.Parameters.Add("@prid", SqlDbType.Int).Value = active_user_id;
            global_command.ExecuteNonQuery();

            //вывести в таблицу
            kontrol_list_2.Rows[0].Cells[10].Value = ib.textBox1.Value.ToString();

            ib.Dispose(); 
        }

        private void копироватьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Clipboard.Clear();
            try
            {
                CopyGridToClipBoard(kontrol_list_1);
                //Clipboard.SetDataObject(.GetClipboardContent());
            }
            catch (Exception ex)
            {
                ;
            }
        }

        private void копироватьToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Clipboard.Clear();
            try
            {
                CopyGridToClipBoard(kontrol_list_2);
                //Clipboard.SetDataObject(kontrol_list_2.GetClipboardContent());
            }
            catch (Exception ex)
            {
                ;
            }
        }

        private void копироватьToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            Clipboard.Clear();
            try
            {
                CopyGridToClipBoard(kontrol_list_res);
                Clipboard.SetDataObject(kontrol_list_res.GetClipboardContent());
            }
            catch (Exception ex)
            {
                ;
            }
        }

        //отправить контольный лист в Excel
        private void toolStripButton19_Click(object sender, EventArgs e)
        {
            CellRange cr;
            ExcelFile excel = new ExcelFile();

            ExcelWorksheet sheet = excel.Worksheets.Add("Ведомость");
            sheet.PrintOptions.HeaderMargin = 0.0;
            sheet.PrintOptions.FooterMargin = 0.0;

            //задать общие свойства свойства
            sheet.DefaultColumnWidth = 3 * 256;
            sheet.Columns[0].Width = 10 * 256;
            sheet.Columns["o"].Width = 7 * 256;
            cr = sheet.Cells.GetSubrange("a1", "o35");
            cr.Merged = true;
            cr.Style.Font.Name = "Times New Roman";
            cr.Style.Font.Size = 10 * 20;
            cr.Style.VerticalAlignment = VerticalAlignmentStyle.Center;
            cr.Style.HorizontalAlignment = HorizontalAlignmentStyle.Left;
            cr.Style.FillPattern.SetSolid(Color.White);
            cr.Merged = false;


            //поставить границы
            cr = sheet.Cells.GetSubrange("a5", "o12");
            cr.Merged = true;
            cr.SetBorders(MultipleBorders.Horizontal | MultipleBorders.Vertical, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
            cr.Merged = false;

            cr = sheet.Cells.GetSubrange("a16", "o25");
            cr.Merged = true;
            cr.SetBorders(MultipleBorders.Horizontal | MultipleBorders.Vertical, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
            cr.Merged = false;

            cr = sheet.Cells.GetSubrange("a2", "o2");
            cr.Merged = true;
            cr.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            cr.Style.Font.Weight = ExcelFont.BoldWeight;
            sheet.Cells["A2"].Value = "КОНТРОЛЬНЫЙ ЛИСТ";

            cr = sheet.Cells.GetSubrange("a3", "o3");
            cr.Merged = true;
            cr.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;            
            cr.Style.Font.Italic = true;
            sheet.Cells["A3"].Value = "Преподаватель:" + active_user_name;


            cr = sheet.Cells.GetSubrange("a14", "o14");
            cr.Merged = true;
            cr.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            cr.Style.Font.Weight = ExcelFont.BoldWeight;
            cr.Style.Font.Italic = true;
            sheet.Cells["A14"].Value = "Зав. кафедрой: ___________________   " +
                "Преподаватель:___________________";


            cr = sheet.Cells.GetSubrange("a27", "o27");
            cr.Merged = true;
            cr.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            cr.Style.Font.Weight = ExcelFont.BoldWeight;
            cr.Style.Font.Italic = true;
            sheet.Cells["A27"].Value = "Зав. кафедрой: ___________________   " +
                "Преподаватель:___________________";



            cr = sheet.Cells.GetSubrange("a16", "o16");
            cr.Merged = true;
            cr.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            cr.Style.Font.Weight = ExcelFont.BoldWeight;
            cr.Style.Font.Italic = false;
            cr.Style.IsTextVertical = true;
            cr.Style.Rotation = 90;
            cr.Style.VerticalAlignment = VerticalAlignmentStyle.Bottom;
            cr.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            cr.Merged = false;
            sheet.Rows[15].Height = 72 * 20;

            cr = sheet.Cells.GetSubrange("a5", "o5");
            cr.Merged = true;
            cr.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            cr.Style.Font.Weight = ExcelFont.BoldWeight;
            cr.Style.Font.Italic = false;
            cr.Style.IsTextVertical = true;
            cr.Style.Rotation = 90;
            cr.Style.VerticalAlignment = VerticalAlignmentStyle.Bottom;
            cr.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            cr.Merged = false;

            sheet.Rows[4].Height = 72 * 20;

            string res = "";
            double chh = 0.0;
            int i = 0, j = 0;

            //вывести первую таблицу
            for (i = 0; i < kontrol_list_1.Columns.Count; i++)
            {
                if (i == 1)
                    sheet.Cells[4, 14].Value = kontrol_list_1.Columns[i].HeaderText;
                else
                {
                    if (i==0)
                        sheet.Cells[4, i].Value = kontrol_list_1.Columns[i].HeaderText;
                    else
                        sheet.Cells[4, i-1].Value = kontrol_list_1.Columns[i].HeaderText;
                }

                for (j = 0; j < kontrol_list_1.Rows.Count; j++)
                {
                    res = "";                    
                    if (i == 0)
                    {
                        res = kontrol_list_1[i, j].Value.ToString();
                    }
                    else
                    {
                        chh = Convert.ToDouble(kontrol_list_1[i, j].Value);

                        if (chh == 0)
                            res = "";
                        else
                            res = kontrol_list_1[i, j].Value.ToString();
                    }

                    if (i == 1)
                    {
                        sheet.Cells[j + 5, 14].Value = res;                        
                    }
                    else
                    {
                        if (i == 0)
                            sheet.Cells[j + 5, i].Value = res;
                        else
                            sheet.Cells[j + 5, i - 1].Value = res;
                    }                    
                    
                    //sheet.Cells[j + 4, i].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
                }
            }

            for (i = 0; i < kontrol_list_2.Columns.Count; i++)
            {
                if (i == 1)
                    sheet.Cells[15, 14].Value = kontrol_list_2.Columns[i].HeaderText;
                else
                {
                    if (i==0)
                        sheet.Cells[15, i].Value = kontrol_list_2.Columns[i].HeaderText;
                    else
                        sheet.Cells[15, i-1].Value = kontrol_list_2.Columns[i].HeaderText;
                }

                for (j = 0; j < kontrol_list_2.Rows.Count; j++)
                {

                    res = "";
                    if (i == 0)
                    {
                        res = kontrol_list_2[i, j].Value.ToString();
                    }
                    else
                    {
                        chh = Convert.ToDouble(kontrol_list_2[i, j].Value);

                        if (chh == 0)
                            res = "";
                        else
                            res = kontrol_list_2[i, j].Value.ToString();
                    }

                    if (i == 1)
                    {
                        sheet.Cells[j + 16, 14].Value = res;
                    }
                    else
                    {
                        if (i == 0)
                            sheet.Cells[j + 16, i].Value = res;
                        else
                            sheet.Cells[j + 16, i - 1].Value = res;
                    }                    
                    
                    //sheet.Cells[j + 4, i].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
                }

                //вывести итог
                for (int ii = 0; ii < kontrol_list_res.Columns.Count; ii++)
                {
                    int jj = 0;
                    res = "";
                    jj = 0;
                    if (ii == 0)
                    {
                        res = kontrol_list_res[ii, jj].Value.ToString();
                    }
                    else
                    {
                        chh = Convert.ToDouble(kontrol_list_res[ii, jj].Value);

                        if (chh == 0)
                            res = "";
                        else
                            res = kontrol_list_res[ii, jj].Value.ToString();
                    }

                    if (ii == 1)
                    {
                        sheet.Cells[23, 14].Value = res;
                    }
                    else
                    {
                        if (ii == 0)
                            sheet.Cells[23, ii].Value = res;
                        else
                            sheet.Cells[23, ii - 1].Value = res;
                    }

                    //--------------------------------------------

                    res = "";
                    jj = 1;
                    if (ii == 0)
                    {
                        res = kontrol_list_res[ii, jj].Value.ToString();
                    }
                    else
                    {
                        chh = Convert.ToDouble(kontrol_list_res[ii, jj].Value);

                        if (chh == 0)
                            res = "";
                        else
                            res = kontrol_list_res[ii, jj].Value.ToString();
                    }

                    if (ii == 1)
                    {
                        sheet.Cells[24, 14].Value = res;
                    }
                    else
                    {
                        if (ii == 0)
                            sheet.Cells[24, ii].Value = res;
                        else
                            sheet.Cells[24, ii - 1].Value = res;
                    }

                    //sheet.Cells[j + 4, i].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
                }
            }


            // --------------- сохранение и открытие --------------
            string FileName = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) +
                "\\ Контрольный_лист_" + active_user_name + ".xls";

            
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

        //вывести информацию в Excel за указанный промежуток времени
        private void toolStripButton20_Click(object sender, EventArgs e)
        {

            CellRange cr;
            ExcelFile excel = new ExcelFile();

            ExcelWorksheet sheet = excel.Worksheets.Add("Ведомость");
            sheet.PrintOptions.HeaderMargin = 0.0;
            sheet.PrintOptions.FooterMargin = 0.0;
            sheet.PrintOptions.PrintCellNotes = false;
            sheet.PrintOptions.PrintNotesSheetEnd = false;
            sheet.PrintOptions.PrintPagesInRows = false;
            sheet.PrintOptions.Portrait = false;

            DataTable zurnal = new DataTable();
            //загрузка расписания
            global_query =
            " select " +
            " [Дата]=cast(d as char(2)) + '/' + " +                //0
            " cast (m as char(2)) + '/' + cast(y as char(4)) + '-' + " +
            " case " +
            " when datepart(dw, dbo.get_date(y,m,d))=1 then 'вс' " +
            " when datepart(dw, dbo.get_date(y,m,d))=2 then 'пн'  " +
            " when datepart(dw, dbo.get_date(y,m,d))=3 then 'вт' " +
            " when datepart(dw, dbo.get_date(y,m,d))=4 then 'ср' " +
            " when datepart(dw, dbo.get_date(y,m,d))=5 then 'чт' " +
            " when datepart(dw, dbo.get_date(y,m,d))=6 then 'пт' " +
            " when datepart(dw, dbo.get_date(y,m,d))=7 then 'сб' " +
            " end, [Пара]=nom_zan, " +
            " [Группа] = grupa.name, " +
            " [Предмет]=predmet.name_krat, [Тема]=tema, " +
            " [Вид занятия]=vid_zan.krat_name, " +
            " [Часы] = rasp.kol_chas, [Подпись] = '' " +
            " from rasp " +
            " join predmet on predmet.id = rasp.predmet_id " +
            " join grupa on grupa.id=rasp.grupa_id " +
            " join vid_zan on vid_zan.id =rasp.vid_zan_id " +
            " where rasp.prepod_id = @USERID and dbo.get_date(y,m,d)>=@D1 and dbo.get_date(y,m,d)<=@D2 " +
            " order by y, m, d, nom_zan ";

            global_command = new SqlCommand(global_query, global_connection);
            global_command.Parameters.Add("@USERID", SqlDbType.Int).Value = active_user_id;
            global_command.Parameters.Add("@D1", SqlDbType.DateTime).Value = begin.Value;
            global_command.Parameters.Add("@D2", SqlDbType.DateTime).Value = end.Value;
            global_adapter = new SqlDataAdapter(global_command);
            global_adapter.Fill(zurnal);


            string endr = (zurnal.Rows.Count+2).ToString();
            string endcell = "h" + endr;

            //задать общие свойства свойства
            sheet.Columns[0].Width = 12 * 256; //день недели
            sheet.Columns[1].Width = 5 * 256; //пара
            sheet.Columns[2].Width = 7 * 256; //группа
            sheet.Columns[3].Width = 16 * 256; //предмет
            sheet.Columns[4].Width = 60 * 256; //Тема	
            sheet.Columns[4].Style.WrapText = true;
            sheet.Columns[5].Width = 10 * 256; //Вид занятия	
            sheet.Columns[6].Width = 6 * 256; //Часы
            sheet.Columns[7].Width = 10 * 256; // подпись


            cr = sheet.Cells.GetSubrange("a1", endcell);
            cr.Merged = true;
            cr.Style.Font.Name = "Times New Roman";
            cr.Style.Font.Size = 10 * 20;
            cr.Style.VerticalAlignment = VerticalAlignmentStyle.Center;
            cr.Style.HorizontalAlignment = HorizontalAlignmentStyle.Left;
            cr.Style.FillPattern.SetSolid(Color.White);
            cr.Merged = false;


            //поставить границы
            cr = sheet.Cells.GetSubrange("a1", endcell);
            cr.Merged = true;
            cr.SetBorders(MultipleBorders.Horizontal | MultipleBorders.Vertical, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
            cr.Merged = false;
            sheet.InsertDataTable(zurnal, "a1", true);

            sheet.Cells["E" + endr].Value = "Итого";
            sheet.Cells["F" + endr].Value = zurnal.Compute("sum(Часы)", "");

            // --------------- сохранение и открытие --------------
            string FileName = "";/* Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) +
                 "\\ Журнал_" + active_user_name + "("
                 + months[begin.Value.Month + 1] + " " + begin.Value.Day.ToString() + "-е по "
                 + months[end.Value.Month + 1] + " " + end.Value.Day.ToString()
                 + "-е).xls";*/

            DialogResult res = saveExcel.ShowDialog();
            saveExcel.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);

            if (res == DialogResult.OK)
            {
                excel.SaveXls(saveExcel.FileName);
            }
            else
            {
                return;
            }
                   
            Thread.Sleep(500);

            try
            {
                Process.Start(saveExcel.FileName);
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

        private void toolStripButton21_Click(object sender, EventArgs e)
        {
            teacher_tab_kontrlist_Enter(sender, new EventArgs());
        }

        //отметить только тему занятия
        private void задатьТемуЗанятияToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (personal_tabble.Rows.Count == 0) return;
            int num = prepod_table.CurrentCell.RowIndex;
            if (num < 0) return;
           
            DataRow dr = personal_tabble.Rows[num];
         
            //отметить посещение
            set_tema_from_indrasp stfi = new set_tema_from_indrasp();
            //set_attend_dialog sad = new set_attend_dialog();

            stfi.zan_id = (int)dr[7];
            stfi.predm_id = (int)dr[12];
            stfi.tema = dr[4].ToString();
            stfi.enddate = end.Value;

            stfi.Text =
                dr[2].ToString() + "," + prepod_table[0, num].Value.ToString() + ", пара №" +
                dr[1].ToString() + " | " + dr[3].ToString();

            stfi.ShowDialog();            
            if (stfi != null) stfi.Dispose();
            load_individ_rasp();

            prepod_table.CurrentCell = prepod_table.Rows[num].Cells[0];
        }

        /// <summary>
        /// обработка события листа личного расписания при вставке темы занятия из буфера обмена
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void вставитьТемуИзБуфераToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (personal_tabble.Rows.Count == 0) return;
            int num = prepod_table.CurrentCell.RowIndex;
            if (num < 0) return;

            // запомнить текущую строку таблицы
            DataRow dr = personal_tabble.Rows[num];

            //получить тему из буфера, нормализовать текст и проверить, что текст не пустой
            string newTema = main.Normalize1(Clipboard.GetText());
            if (newTema.Trim().Length == 0)
            {
                return;
            }

            // отправить тему в таблицу БД
            string zapros = "update rasp set tema = '" + newTema + "'" +
                " where id = " + dr[7].ToString();
            main.global_command = new SqlCommand(zapros, main.global_connection);

            try
            {
                main.global_command.ExecuteNonQuery();
            }
            catch (Exception exx)
            {
                MessageBox.Show("Сбой при сохранении. Нет связи. Повторите сохранение или перезапустите программу.",
                    "Тема не сохранена",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            load_individ_rasp();
            prepod_table.CurrentCell = prepod_table.Rows[num].Cells[6];
        }

        // --------------- служебные глобальные ---      
        /// <summary>
        /// возвращает номер пункта списка DataTable в котром стоит запись с указанным ИД
        /// требование к DataTable - в наборе данных id должно стоять на пером месте
        /// </summary>        
        /// <param name="dt">список элементов из DataTable - источник данных</param>
        /// <param name="ID">номер ид записи</param>
        /// <returns>номер элемента в наборе данных</returns>
        public static int getIndex(DataTable dt, int ID)
        {
            int num = 0;
            foreach (DataRow dr in dt.Rows)
            {
                int current_id = Convert.ToInt32(dr[0]);
                if (ID == current_id) return num;
                num++;
            }

            return 0;
        }

        /// <summary>
        /// получить и вернуть целое, равное ид записи в источнике данных в строке с указанным номером
        /// </summary>
        /// <param name="dt">источник данных (первое поле - id!!)</param>
        /// <param name="i">номер выбранного элемента</param>
        /// <returns></returns>
        public static int getID(DataTable dt, int i)
        {
            int res = 0;
            res = Convert.ToInt32(dt.Rows[i][0]);
            return res;
        }

        /// <summary>
        /// вывести в Excel сетку расписания на неделю по преподавателям
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void prepods_rasp_Click(object sender, EventArgs e)
        {                                   
            string sql = string.Format(
                "select " +
                    " dbo.get_fio(fam, im, ot) as fio, " + //0
                    " dbo.get_date(y,m,d), " + //1
                    " nom_zan, " + //2
                    " rasp.grupa_id as gid, " + //3
                    " grupa.name as gname, " + //4
                    " predmet.id as prid, " + //5
                    " predmet.name_krat as prname, " + //6
                    " kabinet.nomer as kabname, " + //7
                    " prepod.id as pid, d, " + //8 9
                    "	case " + 
                    " when datepart(dw,dbo.get_date(y,m,d))>1 then datepart(dw,dbo.get_date(y,m,d))-1 " + 
                    " else 7 " + 
                    " end as weeday " + //10
                    " from rasp " +
                " join prepod on prepod.id=rasp.prepod_id " +
                " join grupa on grupa.id=rasp.grupa_id " +
                " join predmet on predmet.id=rasp.predmet_id " +
                " join kabinet on kabinet.id=rasp.kabinet_id " +
                " where  " +
                    //" prepod.actual=1   " +
                    " dbo.get_date(y,m,d)>=dbo.get_date({0},{1},{2}) " +
                    " and dbo.get_date(y,m,d)<=dbo.get_date({3},{4},{5}) " +
                    " and rasp.fakultet_id = 9 " +
                " order by dbo.get_fio(fam, im, ot), dbo.get_date(y,m,d), nom_zan",
                starts[week_list.SelectedIndex].Year, starts[week_list.SelectedIndex].Month, starts[week_list.SelectedIndex].Day,
                ends[week_list.SelectedIndex].Year, ends[week_list.SelectedIndex].Month, ends[week_list.SelectedIndex].Day);

            DataTable prepodsrasptable = new DataTable();
            new SqlDataAdapter(sql, global_connection).Fill(prepodsrasptable);

            //фамилии работников с сортировкой по количеству ваданных за неделю часов
            /*sql = string.Format("select distinct dbo.get_fio(fam, im, ot), count(*) from rasp " +
                " join prepod on prepod.id=rasp.prepod_id " +
                " where dbo.get_date(y,m,d)>=dbo.get_date({0},{1},{2}) and dbo.get_date(y,m,d)<=dbo.get_date({3},{4},{5}) " + 
                " group by dbo.get_fio(fam, im, ot) " +
                " order by count(*) desc ",
                starts[week_list.SelectedIndex].Year, starts[week_list.SelectedIndex].Month, starts[week_list.SelectedIndex].Day,
                ends[week_list.SelectedIndex].Year, ends[week_list.SelectedIndex].Month, ends[week_list.SelectedIndex].Day);
            DataTable whoworks = new DataTable();
            new SqlDataAdapter(sql, global_connection).Fill(whoworks);*/


            //список фио работающих преподов
            List<string> whoworks = new List<string>();

            foreach (DataRow dr in prepodsrasptable.Rows)
            {
                string fio = dr[0].ToString();
                if (!whoworks.Contains(fio))
                {
                    whoworks.Add(fio);
                }
            }
            
            List<int> OutGroups = new List<int>(); //список выводимых групп
            List<int> OutLines = new List<int>(); //список номеров строк выводимых дней                      
            bool ShowHeader = true; //показыать стандартную шапку
            string S = ""; //строковый буфер
            CellRange cr;
            string[] Letters = new string[]{"A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T",
                "U","V","W","X","Y","Z","AA","AB","AC","AD","AE","AF","AG","AH","AI","AJ","AK","AL","AM","AN","AO","AP","AQ","AR","AS","AT","AU","AV","AW","AX","AY","AZ"};

            string root = GetMyDocs() + "\\Расписания";
            CreateFolder(root);
            CreateFolder(root + "\\Расписания по преподавателям");
            string FileName = root + "\\Расписания по преподавателям";
            CreateFolder(FileName);
            //имя файла для сохранения
            FileName = FileName + "\\" + week_list.Items[week_list.SelectedIndex].ToString() + ".xls";            

            //перечень вывдомых строк
            OutLines.Clear();
            foreach (int er in empty_rows)
            {
                OutLines.Add(er);
            }

            int num = 1;
            int i = 0;

            //проверить есть ли суббота и воскресенье
            bool showsubb = false, showvoskr = false;
            DataRow[] voskr = prepodsrasptable.Select(" d= " + ends[week_list.SelectedIndex].Day.ToString());
            DataRow[] subb = prepodsrasptable.Select(" d= " + (ends[week_list.SelectedIndex].Day - 1).ToString());            
            if (subb.Length > 0) showsubb = true;
            if (voskr.Length > 0) showvoskr = true;            

            //поставить решетку               
            if (showsubb == false && showvoskr == false) OutLines.Remove(36);
            if (showvoskr == false) OutLines.Remove(43);

            int cols = table.Cols - 1;
            int ColWidth = 17 * 256;
            int ColWidthWide = 25 * 256;

            ExcelFile excel = new ExcelFile();
            excel.LimitNear += new LimitEventHandler(stopwritingerror);
            excel.LimitReached += new LimitEventHandler(stopwritingerror);
            ExcelWorksheet sheet;

            if (prepodsrasptable.Rows.Count == 0) //v
            {
                excel = null;
                sheet = null;
                return;
            }

            int КоличествоЛистовВкниге = 8;

            for (i = 0; i < 5; i++)
            {
                sheet = excel.Worksheets.Add("Часть " + (i+1).ToString());

                sheet.PrintOptions.HeaderMargin = 0.0;
                sheet.PrintOptions.FooterMargin = 0.0;
                sheet.PrintOptions.FitToPage = true;
                sheet.PrintOptions.Portrait = false;

                //задать общие свойства свойства
                cr = sheet.Cells.GetSubrange("a1",
                    Letters[OutGroups.Count * 2 + 1] + (OutLines[OutLines.Count - 1] + 4).ToString());
                cr.Merged = true;
                cr.Style.Font.Name = "Times New Roman";
                cr.Style.Font.Size = 8 * 20;
                cr.Style.Font.Italic = true;
                cr.Style.Font.Weight = ExcelFont.BoldWeight;
                cr.Style.VerticalAlignment = VerticalAlignmentStyle.Center;
                cr.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
                cr.Style.FillPattern.SetSolid(Color.White);
                cr.Merged = false;

                // ------------------   ===== взять шапку раписания из БД
                cr = sheet.Cells.GetSubrange("A1", Letters[КоличествоЛистовВкниге + 1] + "1");
                cr.Merged = true;
                cr.Style.Font.Name = "Times New Roman";
                cr.Style.Font.Weight = ExcelFont.BoldWeight;
                cr.Style.Font.Italic = false;
                cr.Style.Font.Size = 13 * 20;

                S = "Расписание преподавателей на FSystemе " + fakultet_name_krat + " " +
                    "c " + starts[week_list.SelectedIndex].ToShortDateString() +
                    " по " + ends[week_list.SelectedIndex].ToShortDateString();

                sheet.Cells["A1"].Value = S;

                //определить параметры вывдимого расписания

                int span = 0;

                switch (OutLines.Count)
                {
                    case 5: span = 4; break;
                    case 6: span = 3; break;
                    case 7: span = 2; break;
                }

                cr = sheet.Cells.GetSubrange("A3", "B" + (OutLines[OutLines.Count - 1] + span).ToString());
                cr.Merged = true;
                cr.Style.Borders.SetBorders(MultipleBorders.Outside, Color.Black,
                     GemBox.Spreadsheet.LineStyle.Thin);
                cr.Merged = false;

                //sheet.Cells.GetSubrange("A6", "C6").SetBorders(MultipleBorders.Top, Color.Black, GemBox.Spreadsheet.LineStyle.DoubleLine);
                cr = sheet.Cells.GetSubrange("B3", "B" + (OutLines[OutLines.Count - 1] + span).ToString());
                cr.Merged = true;
                cr.Style.Borders[IndividualBorder.Right].LineStyle = GemBox.Spreadsheet.LineStyle.Double;
                cr.Merged = false;

                cr = sheet.Cells.GetSubrange("A3", "A" + (OutLines[OutLines.Count - 1] + span).ToString());
                cr.Merged = true;
                cr.Style.Borders[IndividualBorder.Left].LineStyle = GemBox.Spreadsheet.LineStyle.Double;
                cr.Merged = false;

                cr = sheet.Cells.GetSubrange("A" + (OutLines[OutLines.Count - 1] + span).ToString(),
                    "B" + (OutLines[OutLines.Count - 1] + span).ToString());
                cr.Merged = true;
                cr.Style.Borders[IndividualBorder.Bottom].LineStyle = GemBox.Spreadsheet.LineStyle.Double;
                cr.Merged = false;

                cr = sheet.Cells.GetSubrange("A3", "B3");
                cr.Merged = true;
                cr.Style.Borders[IndividualBorder.Top].LineStyle = GemBox.Spreadsheet.LineStyle.Double;
                cr.Merged = false;

                cr = sheet.Cells.GetSubrange("A3", "B3");
                cr.Merged = true;
                cr.Style.Borders[IndividualBorder.Bottom].LineStyle = GemBox.Spreadsheet.LineStyle.Double;
                cr.Merged = false;

                //первый столбец - дни недели
                sheet.Cells["A3"].Value = "День\nнедели";
                sheet.Cells["B3"].Value = "№";
                //sheet.Cells["C3"].Value = "Время";
                sheet.Rows[5].Height = 25 * 20;

                num = 3;
                string cell1 = "", cell2 = "";
                int days = 0;
                foreach (int j in OutLines)  //вывод дней недели, пар и дат
                {

                    cell1 = string.Format("A{0}", num + 1);

                    sheet.Cells[cell1].Value = table[j, 1].ToString().Substring(0, 5);

                    cell1 = string.Format("A{0}", num + 2);
                    cell2 = string.Format("A{0}", num + 6);
                    cr = sheet.Cells.GetSubrange(cell1, cell2);
                    cr.Merged = true;

                    cr.Style.Rotation = 90;
                    sheet.Cells[cell1].Value = DaysMed[days]; // table[j, 0].ToString();
                    cr.Style.Font.Weight = ExcelFont.MaxWeight;
                    cr.Style.Font.Size = 12 * 20;
                    sheet.Columns[0].Width = 6 * 256;

                    for (int k = 1; k <= 6; k++)
                    {
                        sheet.Cells[num + k - 1, 1].Value = k.ToString();

                        sheet.Columns[1].Width = 3 * 256;

                        //sheet.Cells[num + k - 1, 2].Value = table[j + k, 0].ToString();
                        sheet.Columns[2].Width = 11 * 256;

                        sheet.Rows[num + k - 1].Height = 12 * 20;
                    }

                    num += 6;
                    days++;
                }
                
                sheet.DefaultColumnWidth = ColWidth;            
                
                //горизонтальные одинарные тонкие
                cr = sheet.Cells.GetSubrange("C2", "J33");
                if (showsubb) cr = sheet.Cells.GetSubrange("C3", "J39");
                if (showvoskr) cr = sheet.Cells.GetSubrange("C3", "J45");
                cr.Merged = true;
                cr.Style.Borders[IndividualBorder.Bottom].LineStyle = GemBox.Spreadsheet.LineStyle.Thin;
                cr.Merged = false;


                //правая двойная
                cr = sheet.Cells.GetSubrange("J2", "J33");
                if (showsubb) cr = sheet.Cells.GetSubrange("J3", "J39");
                if (showvoskr) cr = sheet.Cells.GetSubrange("J3", "J45");
                cr.Merged = true;
                cr.Style.Borders[IndividualBorder.Right].LineStyle = GemBox.Spreadsheet.LineStyle.Double;
                cr.Merged = false;

                //верхняя двойная рамка
                cr = sheet.Cells.GetSubrange("C3", "J3");
                cr.Merged = true;
                cr.Style.Borders[IndividualBorder.Top].LineStyle = GemBox.Spreadsheet.LineStyle.Double;
                cr.Merged = false;


                // промежуточные двойные горизонтальные
                int stopline = 33;
                if (showsubb) stopline = 39;
                if (showvoskr) stopline = 45;
                for (int doublelines = 3; doublelines <= stopline; doublelines += 6)
                {
                    cr = sheet.Cells.GetSubrange("C" + doublelines.ToString(), "J" + doublelines.ToString());
                    cr.Merged = true;
                    cr.Style.Borders[IndividualBorder.Bottom].LineStyle = GemBox.Spreadsheet.LineStyle.Double;
                    cr.Merged = false;
                }

                //вертикальные двойные
                /*for (int doublelines = 2; doublelines <= 9; doublelines++)
                {
                    cr = sheet.Cells.GetSubrange(Letters[doublelines] + "3", Letters[doublelines] + stopline.ToString());
                    cr.Merged = true;
                    cr.Style.Borders[IndividualBorder.Right].LineStyle = GemBox.Spreadsheet.LineStyle.Double;
                    cr.Merged = false;
                }*/
            } // --------- конец цикла формирования пустой таблицы расписания

            int counter = 2;
            int i1 = 0, i2 = 0;
            int spanrow = 4;
            if (showsubb) spanrow = 3;
            if (showvoskr) spanrow = 2;


            //foreach (DataRow who in whoworks.Rows)
            for (i = 0; i < whoworks.Count; i++)
            {
                if (i > 0 && i % КоличествоЛистовВкниге == 0)
                {
                    i1++;
                    i2 = 0;
                    counter = 2;
                }   //переходим на новый лист каждые КоличествоЛистовВкниге столбцов

                if (i1 > 4) break; //не болеее 5 листов в книге

                sheet = excel.Worksheets[i1];
                sheet.Columns[i2 + 2].Width = ColWidth;

                /*cr = sheet.Cells.GetSubrange(Letters[i2 + 2] + "4", Letters[i2 + 2] + (OutLines[OutLines.Count - 1] + spanrow).ToString());
                cr.Merged = true;
                cr.Style.Borders.SetBorders(MultipleBorders.Horizontal, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
                cr.Merged = false;*/

                //имя преподавателя
                string cell1 = Letters[counter] + "3";
                sheet.Cells[cell1].Value = whoworks[i];
                sheet.Cells[cell1].Style.Font.Weight = ExcelFont.MaxWeight;
                sheet.Cells[cell1].Style.Font.Size = 10 * 20;
                sheet.Cells[cell1].Style.Borders.SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.DoubleLine);
                sheet.Cells[cell1].Style.VerticalAlignment = VerticalAlignmentStyle.Center;
                sheet.Cells[cell1].Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
                sheet.Cells[cell1].Style.Font.Name = "Times New Roman";

                ////цикл вывода 
                num = 4;

                //получить предметы текущего препода
                DataRow[] prepodinfo = prepodsrasptable.Select(string.Format("fio='{0}'", whoworks[i]));

                // и вывести их
                foreach (DataRow j in prepodinfo)
                {
                    string celltext = j[6].ToString() + " | " + j[4].ToString();
                    if (j[7].ToString() != "--")
                        celltext = celltext + " | " + j[7].ToString();

                    int rownum = Convert.ToInt32(j[10]) * 6 - 2 + Convert.ToInt32(j[2]) - 2;

                    sheet.Cells[rownum, i2 + 2].Style.WrapText = true;
                    //if (rownum >= 9 && rownum % 3 == 0 && rownum % 2 != 0)
                    //    sheet.Cells[rownum, i2 + 2].Style.Borders.SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
                    //else
                    //    sheet.Cells[rownum, i2 + 2].Style.Borders.SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
                    sheet.Cells[rownum, i2 + 2].Style.Borders.SetBorders(MultipleBorders.Left, Color.Black, GemBox.Spreadsheet.LineStyle.Double);
                    sheet.Cells[rownum, i2 + 2].Style.Borders.SetBorders(MultipleBorders.Right, Color.Black, GemBox.Spreadsheet.LineStyle.Double);
                    sheet.Cells[rownum, i2 + 2].Style.VerticalAlignment = VerticalAlignmentStyle.Center;
                    sheet.Cells[rownum, i2 + 2].Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
                    sheet.Cells[rownum, i2 + 2].Style.Font.Name = "Arial Narrow";
                    sheet.Cells[rownum, i2 + 2].Style.Font.Size = 8 * 20;
                    sheet.Cells[rownum, i2 + 2].Value = celltext; // j[10].ToString() + " - " + j[2].ToString();
                }

                sheet.Columns[i2 + 2].AutoFit();
                sheet.Columns[i2 + 2].Width = 20 * 256;
                sheet.PrintOptions.HeaderMargin = 0.0;
                sheet.PrintOptions.FooterMargin = 0.0;
                sheet.PrintOptions.BottomMargin = 1.0;
                sheet.PrintOptions.LeftMargin = 1.0;
                sheet.PrintOptions.RightMargin = 1.0;
                sheet.PrintOptions.TopMargin = 1.0;
                sheet.PrintOptions.FitToPage = false;
                sheet.PrintOptions.Portrait = false;

                counter++;
                i2++;
            }

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


            //MessageBox.Show(starts[week_list.SelectedIndex].ToShortDateString() + "\n" + ends[week_list.SelectedIndex].ToShortDateString());

        }

        //отменить вывод сообщения об ошибке объектом класса Excel (не принимает более 5 листов в книге)
        public void stopwritingerror(object sender, LimitEventArgs lea)
        {
            lea.WriteWarningWorksheet = false;
        }

        private void toolStripButton22_Click(object sender, EventArgs e)
        {
            SaveToExcel(0);
        }

        private void калькуляторToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Process.Start("calc.exe");
        }

        private void prepod_table_KeyUp(object sender, KeyEventArgs e)
        {
            //if (e.Alt)
            {
                switch (e.KeyCode)
                {
                    case Keys.Left: 
                        toolStripButton4_Click(sender, e);
                        break;
                    case Keys.Right:
                        toolStripButton5_Click(sender, e);
                        break;
                    case Keys.Subtract:
                        begin.Value = begin.Value.AddDays(-1);
                        break;
                    case Keys.Multiply:
                        end.Value = end.Value.AddDays(1);
                        break;
                }
            }
        }

       
        private void удалитьСтудентаИзВедомостиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // производит удаление ФИО студента из ведомости по зачету, экзамену
            if (ActiveVedRow == -1) return;

            if (MessageBox.Show("Будут удалены данные студента:" +
                выставленные_отметки.Rows[ActiveVedRow][5].ToString(),
                "Вопрос", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
            {
                return;
            }

            string delsql = "delete from session where rasp_id = @SEID and " + 
                " student_id = @STID and (session.vid_zan_id = 6 or session.vid_zan_id = 16)";
            SqlCommand cmd = new SqlCommand(delsql, global_connection);
            cmd.Parameters.Add("@STID", SqlDbType.Int).Value = выставленные_отметки.Rows[ActiveVedRow][0];
            cmd.Parameters.Add("@SEID", SqlDbType.Int).Value = выставленные_отметки.Rows[ActiveVedRow][6];

            try
            {
                cmd.ExecuteNonQuery();
                teacher_tab_predmet_zachet_Enter(sender, e);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);             
            }
        }

        // -----------------------------------

        /// <summary>
        /// номер строки в сетке ведомости (зачета)
        /// </summary>
        public int ActiveVedRow = -1;
        /// щелчок по сетке зачетной ведомости
        private void zachet_table_CellMouseUp(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex == -1) ActiveVedRow = -1;
            ActiveVedRow = e.RowIndex;
        }

        // ------------------------------

        private void удалитьСтудентаИзВедомостиToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            // производит удаление ФИО студента из ведомости по экзамену
            if (ActiveЭкзамRow == -1) return;

            if (MessageBox.Show("Будут удалены данные студента: " +
                выставленные_отметки.Rows[ActiveЭкзамRow][1].ToString(),
                "Вопрос", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
            {
                return;
            }

            string delsql = "delete from session where id = @SEID ";
            SqlCommand cmd = new SqlCommand(delsql, global_connection);
            cmd.Parameters.Add("@SEID", SqlDbType.Int).Value = выставленные_отметки.Rows[ActiveЭкзамRow][0];
            //cmd.Parameters.Add("@SEID", SqlDbType.Int).Value = выставленные_отметки.Rows[ActiveЭкзамRow][6];

            try
            {
                cmd.ExecuteNonQuery();
                teacher_tab_predmet_exam_Enter(sender, e);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        // --------------------------------------

        public int ActiveЭкзамRow = -1;

        private void exam_table_CellMouseUp(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex == -1) ActiveЭкзамRow = -1;
            ActiveЭкзамRow = e.RowIndex;
        }

        // ---------------------------------------

        // вывести ведомость контрольной работы
        private void teacher_tab_predmet_kontrrab_Enter(object sender, EventArgs e)
        {
            zachet_filled = false;

            //получить всех студентов изучающих предмет
            string query = "select 	student.id, grupa.id " +
                    " from student  " +
                    " join grupa on grupa.id = student.gr_id " +
                    " join predmet on predmet.grupa_id = grupa.id  " +
                    " where student.actual=1 and student.status_id=1 and fam<>'-' and predmet.id = " + id_predmet_in_tree.ToString() +
                    " order by fam, im, ot ";

            stud_set = new DataTable();
            global_adapter = new SqlDataAdapter(query, main.global_connection);
            global_adapter.Fill(stud_set);

            if (stud_set.Rows.Count == 0)
            {
                MessageBox.Show("Список студентов этой группы пуст. Необходимо создать список группы.",
                    "Операция невозможна",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            kontr_table.Rows.Clear();

            int i = 0;

            //определить, заплнирован ли экзамен по этому предмету
            query = string.Format(
                " select dbo.get_date(y,m,d), rasp.id, semestr_id, vid_zan.kod from rasp " +
                " join vid_zan on vid_zan.id = rasp.vid_zan_id   " +
                " where predmet_id = {0} and   " +
                " rasp.uch_god_id = {1} and " +
                " (vid_zan.kod = 'э' or vid_zan.kod = 'з')",
                id_predmet_in_tree, uch_god);


            rsp = new DataTable();
            global_adapter = new SqlDataAdapter(query, global_connection);
            global_adapter.Fill(rsp);

            int res = rsp.Rows.Count;
            if (res == 0)
            {
                ex_message.BackColor = Color.Red;
                ex_message.ForeColor = Color.White;
                exam_table.BackgroundColor = Color.LightGray;
                ex_message.Text = "Экзамен или зачёт ещё не запланирован по расписанию.";
                exam_table.Visible = false;
                exam_result.Visible = false;
                return;
            }

            //получить дату экзамена
            DateTime zach_date = Convert.ToDateTime(rsp.Rows[0][0]);
            string rasp_id = rsp.Rows[0][1].ToString();

            ex_message.BackColor = statusStrip2.BackColor;
            ex_message.ForeColor = Color.Black;
            ex_message.Text = "Экзамен (зачёт) по расписанию: " + zach_date.ToShortDateString();
            kontr_table.BackgroundColor = Color.White;
            kontr_table.Visible = true;
            kontr_result.Visible = true;

            //заполнить список отметок

            //получить вид занятия
            query = string.Format(
                "select kod, vid_zan.id from predmet " +
                " join vidzan_predmet on vidzan_predmet.predmet_id = predmet.id " +
                " join vid_zan on vid_zan.id = vidzan_predmet.vidzan_id " +
                " where predmet.id = {0} and (vid_zan.kod = 'кнр')",
                id_predmet_in_tree);

            rsp = new DataTable();
            global_adapter = new SqlDataAdapter(query, global_connection);
            global_adapter.Fill(rsp);

            string vid = rsp.Rows[0][0].ToString();

            ///   ---- получить и вывести все виды отметок --------------
            otmetki = new DataTable();
            query = "select vid_otmetka.id, vid_otmetka.str_name " +
                " from vid_otmetka " +
                " join vid_zan_otmetka on vid_zan_otmetka.vid_otmetka_id = vid_otmetka.id " +
                " join vid_zan on vid_zan.id = vid_zan_otmetka.vid_zan_id " +
                " where vid_zan.id = " + rsp.Rows[0][1].ToString();
            global_adapter = new SqlDataAdapter(query, global_connection);
            global_adapter.Fill(otmetki);

            zach_dcell = new DataGridViewComboBoxCell();
            kontr_table.Columns[1].CellTemplate = zach_dcell;

            kontr_result.Columns.Clear(); //очистить сетку вывода статистки отметок
            kontr_result.Rows.Clear();
            //очистить массив статистики отметок
            статистика_отметок.Clear();
            видотметки_количество vk = null;

            int i1 = 0;
            for (i1 = 0; i1 < otmetki.Rows.Count; i1++)
            {
                string nm = otmetki.Rows[i1][1].ToString();
                zach_dcell.Items.Add(nm);
                kontr_result.Columns.Add("vd" + i1.ToString(), nm);
                kontr_result.Columns[i1].SortMode = DataGridViewColumnSortMode.NotSortable;
                vk = null;
                vk = new видотметки_количество(nm, 0);
                статистика_отметок.Add(vk);
            }

            kontr_result.Columns.Add("итог", "Итого сдавали");
            kontr_result.Columns[i1].SortMode = DataGridViewColumnSortMode.NotSortable;
            kontr_result.Columns[i1].DefaultCellStyle.BackColor = Color.Pink;

            kontr_result.Rows.Add();

            // ---------------------------------------------------
            query = string.Format("exec dbo.TGetSessionResult {0}, {1}, {2}",
                stud_set.Rows[0][1], id_predmet_in_tree, 15);
            //MessageBox.Show("Это кр - \n" + query);
            global_command = new SqlCommand(query, global_connection);
            global_command.ExecuteNonQuery();

            //если уже есть оценки то вывести их
            /*query = "select count(*) from session " +
                " join rasp on rasp.id = session.rasp_id " +
                " where rasp.predmet_id = " + id_predmet_in_tree.ToString() +
                " and session.rasp_id = " + rasp_id +
                " and session.vid_zan_id = " + rsp.Rows[0][1].ToString();

            global_command = new SqlCommand(query, global_connection);
            res = Convert.ToInt32(global_command.ExecuteScalar());

            //zach_button.Text = res.ToString();

            if (res == 0)  //записей еще нет, создаем ... // \/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
            {
                //MessageBox.Show("новая");
                //создать записи в таблицы session
                foreach (DataRow dr in stud_set.Rows)
                {
                    query = "insert into session (student_id, vid_zan_id, rasp_id, predmet_id, otmetka_id) " +
                        " values (@student_id, @vid_zan_id, @rasp_id, @predm_id, 13)";
                    global_command = new SqlCommand(query, global_connection);

                    global_command.Parameters.Add("@student_id", SqlDbType.Int).Value = dr[0];
                    global_command.Parameters.Add("@vid_zan_id", SqlDbType.Int).Value = 15;
                    global_command.Parameters.Add("@rasp_id", SqlDbType.Int).Value = rasp_id;
                    global_command.Parameters.Add("@predm_id", SqlDbType.Int).Value = id_predmet_in_tree;
                    global_command.ExecuteNonQuery();
                }
            }*/

            //получить данные из таблцы session             
            выставленные_отметки = new DataTable();
            query = "select	student.id, isnull(vid_otmetka.id,-1), prim, " +  // 0 1 2 
                " session.id, otm = isnull(vid_otmetka.str_name,''), " +  // 3 4
                " fio = student.fam + ' ' + left(student.im,1) + '. ' + left(student.ot,1), session.rasp_id, " + //5 6
                " grupa.name " +   //7
                " from student  " +
                " join grupa on grupa.id = student.gr_id " + 
                " join session on session.student_id = student.id " +
                " left outer join vid_otmetka on vid_otmetka.id = session.otmetka_id  " +
                " where " + 
                //" session.rasp_id = " + rasp_id + 
                " grupa.id = " + stud_set.Rows[0][1].ToString() + 
                " and session.vid_zan_id = 15 " + 
                " and session.predmet_id = " + id_predmet_in_tree.ToString() + 
                " and isnull(session.sessiondate, getdate()) between " +
                string.Format("cast('{0}' as datetime) and cast('{1}' as datetime)",
                    starts[0].ToString("yyyyMMdd"), starts[starts.Count - 1].ToString("yyyyMMdd")) + 
                " order by fam, im, ot ";

            global_adapter = new SqlDataAdapter(query, global_connection);
            global_adapter.Fill(выставленные_отметки);

            for (int i2 = 0; i2 < выставленные_отметки.Rows.Count; i2++)
            {
                kontr_table.Rows.Add();
                kontr_table.Rows[i2].Cells[0].Value = выставленные_отметки.Rows[i2][5].ToString();
                int otz = Convert.ToInt32(выставленные_отметки.Rows[i2][1]);
                if (otz != -1)
                {
                    string nm = otmetka_name(otmetki, 0, otz);
                    kontr_table.Rows[i2].Cells[1].Value = nm;
                    внести_отметку_в_статистику(nm);
                }
                kontr_table.Rows[i2].Cells[2].Value = выставленные_отметки.Rows[i2][2];
            }

            //вывод значений в нижнюю таблицу статистики
            int vsego = 0;
            for (i = 0; i < статистика_отметок.Count; i++)
            {
                kontr_result.Rows[0].Cells[i].Value = статистика_отметок[i].Kol;
                if (статистика_отметок[i].OtmName == "неявка") vsego -= статистика_отметок[i].Kol;
                if (статистика_отметок[i].OtmName == "недопуск") vsego -= статистика_отметок[i].Kol;
                vsego += статистика_отметок[i].Kol;
            }
            kontr_result.Rows[0].Cells[i].Value = vsego.ToString();

            zachet_filled = true;
        }

        public int ActiveKrRow = -1;

        private void вывестиАктСдачиКонтрольныхРаботВАрхивToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //teacher_tab_predmet_kontrrab_Enter(sender, e);

            Word.Application wa = null;
            Word.Document doc = main.CreateNewWordDoc(ref wa);
            object nulval = Type.Missing;

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
                "сдачи контрольных работ  в архив", Word.WdParagraphAlignment.wdAlignParagraphCenter);
            Range.Font.Bold = 0;
            Range.Font.Italic = 0;

            Range = main.AddWordDocParagraph(ref doc,
                " ", Word.WdParagraphAlignment.wdAlignParagraphCenter);

            Range = main.AddWordDocParagraph(ref doc,
                "Комиссия в составе:  ", //"главного бухгалтера Кан Н. Е.",
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
                "За " + year_start.Year.ToString() + "/" + year_end.Year.ToString() + " уч. год",
                Word.WdParagraphAlignment.wdAlignParagraphLeft);


            Range = main.AddWordDocParagraph(ref doc,
                " ",
                Word.WdParagraphAlignment.wdAlignParagraphLeft);

            Range = main.AddWordDocParagraph(ref doc,
                "Группа " + выставленные_отметки.Rows[0][7].ToString().ToUpper(),
                Word.WdParagraphAlignment.wdAlignParagraphCenter);
            Range.Font.Bold = 1;

            Range = main.AddWordDocParagraph(ref doc,
                "Дисциплина “" + name_predmet + "”",
                Word.WdParagraphAlignment.wdAlignParagraphCenter);
            Range.Font.Bold = 1;

            Range = main.AddWordDocParagraph(ref doc,
                " ",
                Word.WdParagraphAlignment.wdAlignParagraphLeft);
            Range.Font.Bold = 0;

            int k = 1;
            for (int i = 0; i < kontr_table.Rows.Count; i++)
            {

                if (kontr_table.Rows[i].Cells[1].Value == null) continue;

                if ((kontr_table.Rows[i].Cells[1].Value.ToString() == "недопуск" ||
                    kontr_table.Rows[i].Cells[1].Value.ToString() == "неявка" ||
                    kontr_table.Rows[i].Cells[1].Value.ToString() == "неудовлетворительно" ||
                    kontr_table.Rows[i].Cells[1].Value.ToString().Length == 0))
                {
                    continue;
                }


                Range = main.AddWordDocParagraph(ref doc,
                    k.ToString() + ". " + kontr_table.Rows[i].Cells[0].Value.ToString(),
                    Word.WdParagraphAlignment.wdAlignParagraphLeft);
                k++;
            }

            if (k == 1)
            {
                MessageBox.Show("Акт не построен. Среди выбранных Вами работ нет таких, " +
                    "которые могут быть актированы (нет работ с положительными оценками).",
                    "Отказ операции", MessageBoxButtons.OK, MessageBoxIcon.Error);
                main.WordQuit(wa);
                return;
            }


            saveExcel.Title = "Введите имя для файла акта контрольной работы.";
            saveExcel.Filter = "Файл акта КР в формате MS Word|*.doc";
            saveExcel.FileName = "Акт контр. работ по " + name_predmet_in_tree + ".doc";

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
                "Преподаватель           " + active_user_name + "___________________",
                Word.WdParagraphAlignment.wdAlignParagraphLeft);

            Range = main.AddWordDocParagraph(ref doc,
                "Зав. кафедры КТИС   И.К. Мазур________________________",
                Word.WdParagraphAlignment.wdAlignParagraphLeft);

            Range = main.AddWordDocParagraph(ref doc,
                "Зам. декана ФИВТ     В.В. Семикина_____________________",
                Word.WdParagraphAlignment.wdAlignParagraphLeft);

            /*Range = main.AddWordDocParagraph(ref doc,
                "Архивариус                _______________________________",
                Word.WdParagraphAlignment.wdAlignParagraphLeft);*/

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

        private void toolStripMenuItem25_Click(object sender, EventArgs e)
        {
            // производит удаление ФИО студента из ведомости по экзамену
            if (ActiveKrRow == -1) return;

            if (MessageBox.Show("Будут удалены данные студента:" +
                выставленные_отметки.Rows[ActiveKrRow][5].ToString(),
                "Вопрос", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
            {
                return;
            }

            string delsql = "delete from session where rasp_id = @SEID and student_id = @STID and session.vid_zan_id = 15 ";
            SqlCommand cmd = new SqlCommand(delsql, global_connection);
            cmd.Parameters.Add("@STID", SqlDbType.Int).Value = выставленные_отметки.Rows[ActiveKrRow][0];
            cmd.Parameters.Add("@SEID", SqlDbType.Int).Value = выставленные_отметки.Rows[ActiveKrRow][6];

            try
            {
                cmd.ExecuteNonQuery();
                teacher_tab_predmet_kontrrab_Enter(sender, e);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void kontr_table_CellMouseUp(object sender, DataGridViewCellMouseEventArgs e)
        {     
            ActiveKrRow = e.RowIndex;
        }

        private void kontr_table_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (kontr_table.Rows.Count == 0) return;

            if (!zachet_filled) return;

            //сохранить оценку
            if (e.ColumnIndex == 1)
            {
                string sess_id = выставленные_отметки.Rows[e.RowIndex][3].ToString(), name = "";

                if (kontr_table.Rows[e.RowIndex].Cells[1].Value != null)
                    name = kontr_table.Rows[e.RowIndex].Cells[1].Value.ToString();

                string q = "update session set otmetka_id = " +
                    otmetka_index(name) + " where id = " + sess_id;
                global_command = new SqlCommand(q, global_connection);
                global_command.ExecuteNonQuery();

                int i = 0, vsego = 0;
                for (i = 0; i < статистика_отметок.Count; i++)
                {
                    string n = статистика_отметок[i].OtmName;
                    статистика_отметок[i].Kol = 0;
                    for (int j = 0; j < kontr_table.Rows.Count; j++)
                    {
                        if (kontr_table.Rows[j].Cells[1].Value != null)
                        {
                            if (kontr_table.Rows[j].Cells[1].Value.ToString() == n)
                                статистика_отметок[i].inc();
                        }
                    }

                    kontr_result.Rows[0].Cells[i].Value = статистика_отметок[i].Kol.ToString();
                    if (статистика_отметок[i].OtmName == "неявка") vsego -= статистика_отметок[i].Kol;
                    if (статистика_отметок[i].OtmName == "недопуск") vsego -= статистика_отметок[i].Kol;
                    vsego += статистика_отметок[i].Kol;
                }
                kontr_result.Rows[0].Cells[i].Value = vsego.ToString();

                return;
            }

            //сохранить примечания
            if (e.ColumnIndex == 2)
            {
                string sess_id = выставленные_отметки.Rows[e.RowIndex][3].ToString(),
                    name = kontr_table.Rows[e.RowIndex].Cells[2].Value.ToString();

                string q = "update session set prim = '" + name + "' where id = " + sess_id;
                global_command = new SqlCommand(q, global_connection);
                global_command.ExecuteNonQuery();
                return;
            } 
        }


        // ----- посещаемость -------------- посещаемость ----------------------------

        public string posechGrupId = "";  //ид группы, для которой выставляется посещение
        public DateTime posechDate = DateTime.Today; //дата выставления посещения
        public DataTable posechRaspDayTable = null; //расписание на день
        public DataTable posechStatTable = null; //таблица посещений за день
        public string posechStudName = ""; //имя выделенного студента
        public DataTable posechGlobalStatTable = null; //таблица статистики посещений для вывода в панели посещений
        public string posechCurrentStudID = ""; //ид выбранного студента
        public string posechSelectedMonth = "-1"; //номер месяца выделенной даты

        /// <summary>
        /// заполнение окна посещаемости
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void tabПосещаемость_Enter(object sender, EventArgs e)
        {
            if (tab_grupBox.Items.Count > 0) return;

            foreach(DataRow dr in grups_set.Rows)
                tab_grupBox.Items.Add(dr[0].ToString());

            if (tab_grupBox.Items.Count>0)
            {
                tab_grupBox.SelectedIndex = 0;                
            }

            posechDate = posechDate = new DateTime(DateTime.Today.Year, DateTime.Today.Month, DateTime.Today.Day,
                0, 0, 0, 0);
            //posechCalendar.TodayDate = posechDate;

        }

        private void tab_grupBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            posechGrupId = grups_set.Rows[tab_grupBox.SelectedIndex][2].ToString();
            buildPosechTable();
            FillPosechStatGrid();
        }

        private void posechCalendar_DateSelected(object sender, DateRangeEventArgs e)
        {
            posechDate = new DateTime(e.Start.Year, e.Start.Month, e.Start.Day,
                0, 0, 0, 0);
            buildPosechTable();
            FillPosechStatGrid();//получить статистику                                  
        }

        /// <summary>
        /// построить таблицу посещения на выбранную дату для выбранной группы
        /// </summary>
        public void buildPosechTable()
        {
            if (posechGrid.Rows.Count > 0) posechGrid.Rows.Clear();
            for (int c = 1; c <= 6; c++) posechGrid.Columns[c].HeaderText = "";

            // вывод сведений о выбранной группе и дате
            infoLabel.ForeColor = Color.Navy;
            infoLabel.Tag = 0;
            blinkTimerinPosechTab.Stop();
            infoLabel.Text = tab_grupBox.Items[tab_grupBox.SelectedIndex].ToString() + " - " + 
                posechDate.ToLongDateString();

            // получить расписание группы на выбранный день
            // вызов: функция Select * From dbo.TGetGrupRaspTable(13,'2010-11-14 00:00:00.000')
            // 0 - rasp.nom_zan, 1 - rasp.id, 2 - predmet.name_krat, 
			//   3 - predmet.name, 4 - dbo.get_date(rasp.y, rasp.m, rasp.d), 5 - rasp.predmet_id

            string query = "Select * From dbo.TGetGrupRaspTable(@GrID,@ZanDate)";
            global_command = new SqlCommand(query, global_connection);
            global_command.Parameters.Add("@GrID", SqlDbType.Int).Value = posechGrupId;
            global_command.Parameters.Add("@ZanDate", SqlDbType.DateTime).Value = posechDate;

            global_adapter = new SqlDataAdapter(global_command);
            posechRaspDayTable = new DataTable();
            global_adapter.Fill(posechRaspDayTable);

            global_adapter.Dispose();
            global_command.Dispose();            
            GC.Collect();

            // вывести в шапку номера занятий и названия предметов за день
            int num = 1;
            foreach (DataRow r in posechRaspDayTable.Rows)
            {
                posechGrid.Columns[num++].HeaderText = string.Format("{0}.{1}", r[0], r[2]);
                
            }

            // построить список посещений на указанный день
            // вызов ХП: Создать_Посещение_За_День @GrID int, @Date DateTime, 
            //                                      @Res int output, @ResStr nvarchar(25)

            query = "EXEC Создать_Посещение_За_День @GrID, @ZanDate";
            global_command = new SqlCommand(query, global_connection);
            global_command.Parameters.Add("@GrID", SqlDbType.Int).Value = posechGrupId;
            global_command.Parameters.Add("@ZanDate", SqlDbType.DateTime).Value = posechDate;
            global_command.ExecuteNonQuery();

            // получить список посещений на указанный день
            // вызов: запрос SQL

            query = 
                " Select	fiostud = dbo.GetStudentFIOByID(attend.stud_id), rn = rasp.nom_zan, " +  // 0 1
		                " attend.id,  " +  //2 
		                " asid = attend.stud_id,  " +  //3 
		                " attend.zan_id,  " +  // 4
		                " attend.attend_id,  " +  //5
		                " vid_attend.name  " +  //6
                " From attend  " + 
	                " Join rasp on rasp.id = attend.zan_id  " + 
	                " Join vid_attend on vid_attend.id = attend.attend_id  " + 
	                " Where  " +
                        " dbo.get_date(rasp.y, rasp.m, rasp.d)= @ZanDate  " + 
			                " And   " +
                        " rasp.grupa_id = @GrID  " +
                " Order by rasp.nom_zan, fiostud  ";

            global_command = new SqlCommand(query, global_connection);
            global_command.Parameters.Add("@GrID", SqlDbType.Int).Value = posechGrupId;
            global_command.Parameters.Add("@ZanDate", SqlDbType.DateTime).Value = posechDate;


            global_adapter = new SqlDataAdapter(global_command);
            posechStatTable = new DataTable();
            global_adapter.Fill(posechStatTable);

            global_adapter.Dispose();
            global_command.Dispose();
            GC.Collect();

            // нет сведений о списке
            if (posechRaspDayTable.Rows.Count == 0)
            {
                posechCurrentStudID = "-";
                infoLabel.Text = "Нет занятий.";
                infoLabel.Tag = 1;
                infoLabel.ForeColor = Color.Red;
                blinkTimerinPosechTab.Start();
                return;
            }
            if (posechStatTable.Rows.Count == 0)
            {
                posechCurrentStudID = "-";
                infoLabel.Text = "Список пуст.";
                infoLabel.Tag = 2;
                infoLabel.ForeColor = Color.Red;
                blinkTimerinPosechTab.Start();
                return;
            }

                        
            //определить количество разных фамилий
            DataRow[] StudentList = posechStatTable.Select("rn=" + posechRaspDayTable.Rows[0][0].ToString());
            int FamCount = StudentList.Length;

            //цикл по парам
            int colnumber = 1;
            int rownumber = 0;
            int stud_plus = 0, gr_plus = 0;
            int stud_minus = 0, gr_minus = 0;

            foreach (DataRow studrow in StudentList) // перебор студентов
            {
                posechGrid.Rows.Add(new object[] { studrow[0].ToString(), Properties.Resources.empty, 
                    Properties.Resources.empty, Properties.Resources.empty, Properties.Resources.empty, 
                    Properties.Resources.empty, Properties.Resources.empty});
                
                string StudID = studrow[3].ToString();

                colnumber = 1;
                stud_plus = 0; stud_minus = 0; // перебор посещений студента
                foreach (DataRow posechRow in posechRaspDayTable.Rows)
                {
                    DataRow[] attendinfo = posechStatTable.Select("rn=" + posechRow[0].ToString() + " and " +
                        "asid=" + StudID);
                    string idposech = "3";
                    if (attendinfo.Length != 0)
                    {
                        idposech = attendinfo[0][5].ToString();
                    }

                    switch (idposech)
                    {
                        case "1": posechGrid.Rows[rownumber].Cells[colnumber].Value = Properties.Resources.небыл;
                            posechGrid.Rows[rownumber].Cells[colnumber].Tag = -1;
                            stud_minus++;
                            break;
                        case "2": posechGrid.Rows[rownumber].Cells[colnumber].Value = Properties.Resources.был;
                            posechGrid.Rows[rownumber].Cells[colnumber].Tag = 1;
                            stud_plus++;
                            break;
                        case "3": posechGrid.Rows[rownumber].Cells[colnumber].Value = Properties.Resources.неизв;
                            posechGrid.Rows[rownumber].Cells[colnumber].Tag = 0;
                            break;
                    }

                    colnumber++;
                    gr_minus += stud_minus;
                    gr_plus += stud_plus;

                    posechGrid.Rows[rownumber].Cells[0].Value = string.Format("{0} - [ нб={1}, б={2} ]",
                        studrow[0].ToString(), stud_minus, stud_plus);
                    posechGrid.Rows[rownumber].Cells[0].Tag = studrow[3].ToString(); //запомнить ид студента в строке
           
                }
                rownumber++;
                posechGrid.Columns[0].Width = 250;
            }

            // запомнить ид выбранного студента
            if (posechGrid.Rows.Count > 0)
            {
                posechCurrentStudID = posechGrid.Rows[0].Cells[0].Tag.ToString();
                posechStudName = posechGrid.Rows[0].Cells[0].Value.ToString().Substring(0, posechGrid.Rows[0].Cells[0].Value.ToString().IndexOf(" "));
                PosechStatGrid.Columns[1].HeaderText = posechStudName + " на выбр. дату";
                PosechStatGrid.Columns[2].HeaderText = posechStudName + " за неделю";
                PosechStatGrid.Columns[3].HeaderText = posechStudName + " за месяц";
                //Text = posechCurrentStudID;
            }
            else
            {
                PosechStatGrid.Columns[1].HeaderText = "Студент на выбр. дату";
                PosechStatGrid.Columns[2].HeaderText = "Студент за неделю";
                PosechStatGrid.Columns[3].HeaderText = "Студент за месяц";
            }

            //posechRaspDayTable.Rows[0].ItemArray[0] = "1";
            

        }

        //мигает надпись в панели инструментов если нет занятий или список группы пуст
        private void blinkTimerinPosechTab_Tick(object sender, EventArgs e)
        {
            if (infoLabel.ForeColor == Color.Red)
            {
                infoLabel.ForeColor = Color.Blue;
            }
            else
            {
                infoLabel.ForeColor = Color.Red;                
            }

        }

        private void infoLabel_Click(object sender, EventArgs e)
        {
            int tag = Convert.ToInt32(infoLabel.Tag);
            switch (tag)
            {
                case 1:
                    MessageBox.Show("В этот день для группы " + tab_grupBox.Items[tab_grupBox.SelectedIndex].ToString() + 
                        " нет занятий.", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    blinkTimerinPosechTab.Stop();
                    break;
                case 2:
                    MessageBox.Show("В группе " + tab_grupBox.Items[tab_grupBox.SelectedIndex].ToString() +
                        " не задан список студентов. Заполните список студентов и вернитесь к выполнению этой операции", 
                        "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    break;
                default:
                    break;
            }
        }

        private void posechGrid_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            posechCurrentStudID = "-";
            
            if (posechGrid.RowCount == 0) return;
            if (e.RowIndex < 0) return;

            // запомнить ид выбранного студента
            posechCurrentStudID = posechGrid.Rows[e.RowIndex].Cells[0].Tag.ToString();
            // Вывод студента
            if (PosechStatGrid.Rows.Count > 0)
            {
                posechStudName = posechGrid.Rows[e.RowIndex].Cells[0].Value.ToString().Substring(0, posechGrid.Rows[e.RowIndex].Cells[0].Value.ToString().IndexOf(" "));
                PosechStatGrid.Columns[1].HeaderText = posechStudName + " на выбр. дату";
                PosechStatGrid.Columns[2].HeaderText = posechStudName + " за неделю";
                PosechStatGrid.Columns[3].HeaderText = posechStudName + " за месяц";
            }
            else
            {
                PosechStatGrid.Columns[1].HeaderText = "Студент на выбр. дату";
                PosechStatGrid.Columns[2].HeaderText = "Студент за неделю";
                PosechStatGrid.Columns[3].HeaderText = "Студент за месяц";
            }

            if (e.ColumnIndex < 1) return;                        
            
            if (e.ColumnIndex > posechRaspDayTable.Rows.Count) return;

            if (posechGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value == Properties.Resources.empty)
                return;

            string otm = "";            
            int tag = Convert.ToInt32(posechGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Tag);

            switch (tag)
            {
                case -1:
                    otm = "не был на занятии";
                    break;
                case 0:
                    otm = "отметка о посещении не выставлялась и нуждается уточнении";
                    break;
                case 1:
                    otm = "присутствовал на занятии";
                    break;
            }

            //label2.Text = posechGrid.Rows[e.RowIndex].Cells[0].Value.ToString() + " - " +
             //   posechRaspDayTable.Rows[e.ColumnIndex - 1][3].ToString() + " - " + otm;

            // пересчет студента
            if (PosechStatGrid.Rows.Count > 0)
            {
                posechStudName = posechGrid.Rows[e.RowIndex].Cells[0].Value.ToString().Substring(0, posechGrid.Rows[e.RowIndex].Cells[0].Value.ToString().IndexOf(" "));
                PosechStatGrid.Columns[1].HeaderText = posechStudName + " на выбр. дату";
                PosechStatGrid.Columns[2].HeaderText = posechStudName + " за неделю";
                PosechStatGrid.Columns[3].HeaderText = posechStudName + " за месяц";
                UpdateStudentStatDay();
                FillPosechStatGrid();
                UpdateGroupStatDay();
            }
            else
            {
                PosechStatGrid.Columns[1].HeaderText = "Студент на выбр. дату";
                PosechStatGrid.Columns[2].HeaderText = "Студент за неделю";
                PosechStatGrid.Columns[3].HeaderText = "Студент за месяц";
            }

            //MessageBox.Show("click");

        }

        /// <summary>
        /// сохранить данные о посещении из полученной ячейки
        /// </summary>
        /// <param name="e">ячйека, в которой сохраняется отметка посещения</param>
        public void SaveCell(DataGridViewCell e)
        {
            string otm = "";
            int tag = Convert.ToInt32(posechGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Tag);

            // меняем отметку о посещении            
            
            switch (tag)
            {
                case -1:
                    otm = "присутствовал на занятии";
                    posechGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Tag = 1;
                    tag = 1;
                    posechGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = Properties.Resources.был;
                    //UpdateStudStatWeek(1);
                    //UpdateStudStatMonth(1);
                    break;
                case 0:
                    otm = "присутствовал на занятии";
                    posechGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Tag = 1;
                    tag = 1;
                    posechGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = Properties.Resources.был;
                    //UpdateStudStatWeek(1);
                    //UpdateStudStatMonth(1);
                    break;
                case 1:
                    otm = "не был на занятии";
                    posechGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Tag = -1;
                    tag = -1;
                    posechGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = Properties.Resources.небыл;
                    //UpdateStudStatWeek(0);
                    //UpdateStudStatMonth(0);
                    break;
            }

            // выполнить сохранение в БД
            string sqlcmd = "update attend set attend_id=@AttID where zan_id=@ZanID and stud_id=@StudID";
            global_command = new SqlCommand(sqlcmd, global_connection);
            global_command.Parameters.Add("@AttID", SqlDbType.Int).Value = (tag == 1) ? 2 : 1;
            global_command.Parameters.Add("@ZanID", SqlDbType.Int).Value = posechRaspDayTable.Rows[e.ColumnIndex - 1][1].ToString();
            global_command.Parameters.Add("@StudID", SqlDbType.Int).Value = posechGrid.Rows[e.RowIndex].Cells[0].Tag.ToString();

            try
            {
                if (global_connection.State != ConnectionState.Open) global_connection.Open();
                global_command.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при выполнении операции. Сетевой сбой. Перезапустите программу и повторите действие ещё раз.",
                    "Сбой программы", MessageBoxButtons.OK, MessageBoxIcon.Error);
                switch (tag)
                {
                    case -1:
                    case 0:
                        otm = "присутствовал на занятии";
                        posechGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Tag = 1;
                        tag = 1;
                        posechGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = Properties.Resources.был;
                        //UpdateStudStatWeek(1);
                        //UpdateStudStatMonth(1);
                        break;
                    case 1:
                        otm = "не был на занятии";
                        posechGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Tag = -1;
                        tag = -1;
                        posechGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = Properties.Resources.небыл;
                        //UpdateStudStatWeek(0);
                        //UpdateStudStatMonth(1);
                        break;

                }
                return;
            }

            //выполнить пересчёт по строке
            int plus = 0, minus = 0;
            for (int cellnum = 1; cellnum <= posechRaspDayTable.Rows.Count; cellnum++)
            {
                if ((int)posechGrid.Rows[e.RowIndex].Cells[cellnum].Tag == 1) plus++;
                if ((int)posechGrid.Rows[e.RowIndex].Cells[cellnum].Tag == -1) minus++;
            }

            string tmp = posechGrid.Rows[e.RowIndex].Cells[0].Value.ToString();
            posechStudName = tmp.Substring(0, tmp.IndexOf(" -"));
            posechGrid.Rows[e.RowIndex].Cells[0].Value = string.Format("{0} - [ нб={1}, б={2} ]",
                        posechStudName, minus, plus);

            FillPosechStatGrid();

            //MessageBox.Show("dbl_click");

            //label2.Text = posechGrid.Rows[e.RowIndex].Cells[0].Value.ToString() + " - " +
            //    posechRaspDayTable.Rows[e.ColumnIndex - 1][3].ToString() + " - " + otm;
        }

        /// <summary>
        /// выставить посещение "был" в ячейку
        /// </summary>
        /// <param name="e">ячейка для выставления посещения</param>
        public void SaveCellPlus(DataGridViewCell e) 
        {
            string otm = "";
            int tag = Convert.ToInt32(posechGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Tag);

            // меняем отметку о посещении            
            otm = "присутствовал на занятии";
            posechGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Tag = 1;
            tag = 1;
            posechGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = Properties.Resources.был;

            // выполнить сохранение в БД
            string sqlcmd = "update attend set attend_id=@AttID where zan_id=@ZanID and stud_id=@StudID";
            global_command = new SqlCommand(sqlcmd, global_connection);
            global_command.Parameters.Add("@AttID", SqlDbType.Int).Value = 2;
            global_command.Parameters.Add("@ZanID", SqlDbType.Int).Value = posechRaspDayTable.Rows[e.ColumnIndex - 1][1].ToString();
            global_command.Parameters.Add("@StudID", SqlDbType.Int).Value = posechGrid.Rows[e.RowIndex].Cells[0].Tag.ToString();

            try
            {
                if (global_connection.State != ConnectionState.Open) global_connection.Open();
                global_command.ExecuteNonQuery();
            }
            catch (Exception ex)
            {              
                return;
            }

            //выполнить пересчёт по строке
            int plus = 0, minus = 0;
            for (int cellnum = 1; cellnum <= posechRaspDayTable.Rows.Count; cellnum++)
            {
                if ((int)posechGrid.Rows[e.RowIndex].Cells[cellnum].Tag == 1) plus++;
                if ((int)posechGrid.Rows[e.RowIndex].Cells[cellnum].Tag == -1) minus++;
            }

            string tmp = posechGrid.Rows[e.RowIndex].Cells[0].Value.ToString();
            posechStudName = tmp.Substring(0, tmp.IndexOf(" -"));
            posechGrid.Rows[e.RowIndex].Cells[0].Value = string.Format("{0} - [ нб={1}, б={2} ]",
                        posechStudName, minus, plus);

            //label2.Text = posechGrid.Rows[e.RowIndex].Cells[0].Value.ToString() + " - " +
            //    posechRaspDayTable.Rows[e.ColumnIndex - 1][3].ToString() + " - " + otm;
        }

        /// <summary>
        /// выставить посещение "не был" в ячейку
        /// </summary>
        /// <param name="e">ячейка для выставления посещения</param>
        public void SaveCellMinus(DataGridViewCell e)
        {
            string otm = "";
            int tag = Convert.ToInt32(posechGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Tag);

            // меняем отметку о посещении            
            otm = "присутствовал на занятии";
            posechGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Tag = -1;
            tag = -1;
            posechGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = Properties.Resources.небыл;

            // выполнить сохранение в БД
            string sqlcmd = "update attend set attend_id=@AttID where zan_id=@ZanID and stud_id=@StudID";
            global_command = new SqlCommand(sqlcmd, global_connection);
            global_command.Parameters.Add("@AttID", SqlDbType.Int).Value = 1;
            global_command.Parameters.Add("@ZanID", SqlDbType.Int).Value = posechRaspDayTable.Rows[e.ColumnIndex - 1][1].ToString();
            global_command.Parameters.Add("@StudID", SqlDbType.Int).Value = posechGrid.Rows[e.RowIndex].Cells[0].Tag.ToString();

            try
            {
                if (global_connection.State != ConnectionState.Open) global_connection.Open();
                global_command.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                return;
            }

            //выполнить пересчёт по строке
            int plus = 0, minus = 0;
            for (int cellnum = 1; cellnum <= posechRaspDayTable.Rows.Count; cellnum++)
            {
                if ((int)posechGrid.Rows[e.RowIndex].Cells[cellnum].Tag == 1) plus++;
                if ((int)posechGrid.Rows[e.RowIndex].Cells[cellnum].Tag == -1) minus++;
            }

            string tmp = posechGrid.Rows[e.RowIndex].Cells[0].Value.ToString();
            posechStudName = tmp.Substring(0, tmp.IndexOf(" -"));
            posechGrid.Rows[e.RowIndex].Cells[0].Value = string.Format("{0} - [ нб={1}, б={2} ]",
                        posechStudName, minus, plus);

            //label2.Text = posechGrid.Rows[e.RowIndex].Cells[0].Value.ToString() + " - " +
            //    posechRaspDayTable.Rows[e.ColumnIndex - 1][3].ToString() + " - " + otm;
        }

        //выполнить отметку посещения
        private void posechGrid_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (posechGrid.RowCount == 0) return;
            if (e.ColumnIndex < 1) return;
            if (e.RowIndex < 0) return;
            if (e.ColumnIndex > posechRaspDayTable.Rows.Count) return;

            if (posechGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value == Properties.Resources.empty)
                return;

            SaveCell(posechGrid.Rows[e.RowIndex].Cells[e.ColumnIndex]);
            //получить студента за день
            UpdateStudentStatDay();
            UpdateGroupStatDay();
        }

        private void действиеДляСтрокиВыставитьбылДляВсейСтрокиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (posechGrid.Rows.Count == 0) return;
            int SelRow = posechGrid.SelectedCells[0].RowIndex;
            if (SelRow<0) return;            

            
            foreach (DataGridViewCell cell in posechGrid.Rows[SelRow].Cells)
            {
                if (cell.ColumnIndex == 0 || cell.ColumnIndex > posechRaspDayTable.Rows.Count) continue;
                SaveCellPlus(cell);
            }
            FillPosechStatGrid();
        }

        private void выставитьбылДляВсехВыделенныхToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (posechGrid.Rows.Count == 0) return;
            foreach (DataGridViewCell cell in posechGrid.SelectedCells)
            {
                if (cell.ColumnIndex == 0) continue;
                if (cell.ColumnIndex > posechRaspDayTable.Rows.Count) continue;
                SaveCellPlus(cell);
            }
            FillPosechStatGrid();
        }

        private void выставитьнеБылДляВсехВыделенныхToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (posechGrid.Rows.Count == 0) return;
            foreach (DataGridViewCell cell in posechGrid.SelectedCells)
            {
                if (cell.ColumnIndex == 0) continue;
                if (cell.ColumnIndex > posechRaspDayTable.Rows.Count) continue;
                SaveCellMinus(cell);
            }
            FillPosechStatGrid();
        }

        private void выставитьнеБылДляВсейСтрокиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (posechGrid.Rows.Count == 0) return;
            int SelRow = posechGrid.SelectedCells[0].RowIndex;
            if (SelRow < 0) return;


            foreach (DataGridViewCell cell in posechGrid.Rows[SelRow].Cells)
            {
                if (cell.ColumnIndex == 0 || cell.ColumnIndex > posechRaspDayTable.Rows.Count) continue;
                SaveCellMinus(cell);
            }
            FillPosechStatGrid();
        }

        private void выставитьбылДляВсегоСтолбцаСтрокиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (posechGrid.Rows.Count == 0) return;
            int SelCol = posechGrid.SelectedCells[0].ColumnIndex;
            if (SelCol <= 0) return;
            if (SelCol > posechRaspDayTable.Rows.Count) return;

            for (int i = 0; i < posechGrid.Rows.Count; i++)
            {
                SaveCellPlus(posechGrid.Rows[i].Cells[SelCol]);
            }
            FillPosechStatGrid();
        }

        private void выставитьбылДляВсехКромеВыделенныхToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (posechGrid.Rows.Count == 0) return;
            for (int i = 1; i <= posechRaspDayTable.Rows.Count; i++)
            {
                for (int j = 0; j < posechGrid.Rows.Count; j++)
                {
                    if (!posechGrid.Rows[j].Cells[i].Selected)
                        SaveCellPlus(posechGrid.Rows[j].Cells[i]);
                }
            }
            FillPosechStatGrid();
        }

        private void выставитьнеБылДляВсехКромеВыделенныхToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (posechGrid.Rows.Count == 0) return;
            for (int i = 1; i <= posechRaspDayTable.Rows.Count; i++)
            {
                for (int j = 0; j < posechGrid.Rows.Count; j++)
                {
                    if (!posechGrid.Rows[j].Cells[i].Selected)
                        SaveCellMinus(posechGrid.Rows[j].Cells[i]);
                }
            }
            FillPosechStatGrid();
        }

        private void выставитьнеБылДляВсегоСтолбцаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (posechGrid.Rows.Count == 0) return;
            int SelCol = posechGrid.SelectedCells[0].ColumnIndex;
            if (SelCol <= 0) return;
            if (SelCol > posechRaspDayTable.Rows.Count) return;

            for (int i = 0; i < posechGrid.Rows.Count; i++)
            {
                SaveCellMinus(posechGrid.Rows[i].Cells[SelCol]);
            }
        }

        //получить начало и конец недели, на которую выпала выбранная дата посещения
        //метод даёт номер в массиве Starts и Ends
        public int GetWeekNumByDate()
        {
            for (int i = 0; i < starts.Count; i++)
            {
                if (posechDate >= starts[i] && posechDate <= ends[i])
                    return i;
            }

            return 0;
        }

        //получить статистику по столбцам, по студенту и по группе выделенный день, за неделю и за месяц и выести в 
        // информационную панель внизу
        public void FillPosechStatGrid()
        {
            // получить данные
            string sql = "select atid = attend.attend_id, rid = rasp.grupa_id, " +
                    " dt = dbo.get_date(rasp.y, rasp.m, rasp.d), stid = attend.stud_id, mes = rasp.m " +
                    " from attend  " +
                        " join rasp on rasp.id = attend.zan_id  " +
                        " join grupa on grupa.id = rasp.grupa_id  " +
                    " where  " +
                        " rasp.uch_god_id = @UchGodID   " +
                            " and grupa.fakultet_id = @FakultID  " +
                            " and grupa.id = @GrID " +
                            " and (rasp.m between @Mes-1 and @Mes+1  )";
            global_command = new SqlCommand(sql, global_connection);
            global_command.Parameters.Add("@UchGodID", SqlDbType.Int).Value = uch_god;
            global_command.Parameters.Add("@FakultID", SqlDbType.Int).Value = fakultet_id;
            global_command.Parameters.Add("@GrID", SqlDbType.Int).Value = posechGrupId;
            global_command.Parameters.Add("@Mes", SqlDbType.Int).Value = posechDate.Month;
            posechGlobalStatTable = new DataTable();
            (new SqlDataAdapter(global_command)).Fill(posechGlobalStatTable);

            PosechStatGrid.Rows.Clear();
            PosechStatGrid.Rows.Add(new object[] { "Количество посещений", "", "", "", "", "", "" });
            PosechStatGrid.Rows.Add(new object[] { "Количество пропусков", "", "", "", "", "", "" });
            
            //получить студента за день
            UpdateStudentStatDay();  
          
            //получить студента за неделю
            GetStudStatWeek();

            //получить студента за месяц
            GetStudStatMonth();

            //вывести сведения по группе за день
            UpdateGroupStatDay();
            GetGroupStatWeek();
            GetGroupStatMonth();
        }

        // пересчитать и вывести статистику посещаемости на основе набора данных posechGlobalStatTable
        public enum КритерийПересчета { ByStudent, ByGroup, All };
        public void CalcAndPrintStatGrid(КритерийПересчета kp)
        {            
            switch (kp)
            {
                case КритерийПересчета.ByStudent: //вывод по студенту
                    break;
                case КритерийПересчета.ByGroup: //вывод по группе
                    break;
                case КритерийПересчета.All: //вывод по студенту и группе
                    break; 
            }
        }

        //обновить статистику студента за день
        public void UpdateStudentStatDay()
        {
            if (posechCurrentStudID != "-")
            {
                int plus = 0, minus = 0;
                int posechStudIndex = (posechGrid.SelectedCells.Count > 0) ? posechGrid.SelectedCells[0].RowIndex : 0;
                for (int i = 1; i <= posechRaspDayTable.Rows.Count; i++)
                {
                    if (Convert.ToInt32(posechGrid.Rows[posechStudIndex].Cells[i].Tag) == 1) plus++;
                    if (Convert.ToInt32(posechGrid.Rows[posechStudIndex].Cells[i].Tag) == -1) minus++;
                }

                PosechStatGrid.Rows[0].Cells[1].Value = plus.ToString();
                PosechStatGrid.Rows[1].Cells[1].Value = minus.ToString();
            }
            else
            {
                PosechStatGrid.Rows[0].Cells[1].Value = "-";
                PosechStatGrid.Rows[1].Cells[1].Value = "-";
                PosechStatGrid.Columns[1].HeaderText = "Студент на выбр. дату";
                PosechStatGrid.Columns[2].HeaderText = "Студент за неделю";
                PosechStatGrid.Columns[3].HeaderText = "Студент за месяц";
            }
        }

        //обновить данные посещаемости за неделю для студента
        public void UpdateStudStatWeek(int pr)
        {
            if (posechCurrentStudID != "-")
            {
                if (pr == 1)
                {
                    PosechStatGrid.Rows[0].Cells[2].Value = (Convert.ToInt32(PosechStatGrid.Rows[0].Cells[2].Value) + 1).ToString();
                    PosechStatGrid.Rows[1].Cells[2].Value = (Convert.ToInt32(PosechStatGrid.Rows[1].Cells[2].Value) - 1).ToString();
                }
                else
                {
                    PosechStatGrid.Rows[0].Cells[2].Value = (Convert.ToInt32(PosechStatGrid.Rows[0].Cells[2].Value) - 1).ToString();
                    PosechStatGrid.Rows[1].Cells[2].Value = (Convert.ToInt32(PosechStatGrid.Rows[1].Cells[2].Value) + 1).ToString();
                }
            }
        }

        //получить данные посещаемости студента за неделю
        public void GetStudStatWeek()
        {
            //получить статистику студента за неделю
            int numw = GetWeekNumByDate();
            if (posechCurrentStudID != "-")
            {
                DataRow[] studweek = posechGlobalStatTable.Select(" dt>= '" + starts[numw].ToShortDateString() + "' " +
                        " and dt <='" + ends[numw].ToShortDateString() + "' and stid = " + posechCurrentStudID);

                DataRow[] studweekplus = posechGlobalStatTable.Select(" dt>= '" + starts[numw].ToShortDateString() + "' " +
                        " and dt <='" + ends[numw].ToShortDateString() + "' and stid = " + posechCurrentStudID +
                        " and atid=2");

                PosechStatGrid.Rows[0].Cells[2].Value = studweekplus.Length.ToString();
                PosechStatGrid.Rows[1].Cells[2].Value = (studweek.Length - studweekplus.Length).ToString();
            }
            else
            {
                PosechStatGrid.Rows[0].Cells[2].Value = "-";
                PosechStatGrid.Rows[1].Cells[2].Value = "-";
                PosechStatGrid.Columns[1].HeaderText = "Студент на выбр. дату";
                PosechStatGrid.Columns[2].HeaderText = "Студент за неделю";
                PosechStatGrid.Columns[3].HeaderText = "Студент за месяц";
            }
        }

        //обновить данные посещаемости за месяц для студента
        public void UpdateStudStatMonth(int pr)
        {
            if (posechCurrentStudID != "-")
            {
                if (pr == 1)
                {
                    PosechStatGrid.Rows[0].Cells[3].Value = (Convert.ToInt32(PosechStatGrid.Rows[0].Cells[2].Value) + 1).ToString();
                    PosechStatGrid.Rows[1].Cells[3].Value = (Convert.ToInt32(PosechStatGrid.Rows[1].Cells[2].Value) - 1).ToString();
                }
                else
                {
                    PosechStatGrid.Rows[0].Cells[3].Value = (Convert.ToInt32(PosechStatGrid.Rows[0].Cells[2].Value) - 1).ToString();
                    PosechStatGrid.Rows[1].Cells[3].Value = (Convert.ToInt32(PosechStatGrid.Rows[1].Cells[2].Value) + 1).ToString();
                }
            }
        }

        //получить данные посещаемости студента за неделю
        public void GetStudStatMonth()
        {
            //получить статистику студента за неделю            
            if (posechCurrentStudID != "-")
            {
                DataRow[] studmon = posechGlobalStatTable.Select(" mes = " + posechDate.Month.ToString() + " and stid = " + posechCurrentStudID);

                DataRow[] studmonplus = posechGlobalStatTable.Select(" mes = " + posechDate.Month.ToString() +
                    " and atid=2 " + " and stid = " + posechCurrentStudID);

                PosechStatGrid.Rows[0].Cells[3].Value = studmonplus.Length.ToString();
                PosechStatGrid.Rows[1].Cells[3].Value = (studmon.Length - studmonplus.Length).ToString();
            }
            else
            {
                PosechStatGrid.Rows[0].Cells[3].Value = "-";
                PosechStatGrid.Rows[1].Cells[3].Value = "-";
                PosechStatGrid.Columns[1].HeaderText = "Студент на выбр. дату";
                PosechStatGrid.Columns[2].HeaderText = "Студент за неделю";
                PosechStatGrid.Columns[3].HeaderText = "Студент за месяц";
            }
        }

        //обновить статистику группы за день
        public void UpdateGroupStatDay()
        {
            if (posechRaspDayTable.Rows.Count > 0)
            {
                int plus = 0, minus = 0;                
                for (int i = 1; i <= posechRaspDayTable.Rows.Count; i++)
                {
                    for (int j = 0; j < posechGrid.Rows.Count; j++)
                    {
                        if (Convert.ToInt32(posechGrid.Rows[j].Cells[i].Tag) == 1) plus++;
                        if (Convert.ToInt32(posechGrid.Rows[j].Cells[i].Tag) == -1) minus++;
                    }
                }

                PosechStatGrid.Rows[0].Cells[4].Value = plus.ToString();
                PosechStatGrid.Rows[1].Cells[4].Value = minus.ToString();
            }
            else
            {
                PosechStatGrid.Rows[0].Cells[4].Value = "-";
                PosechStatGrid.Rows[1].Cells[4].Value = "-";
                PosechStatGrid.Columns[4].HeaderText = "Группа на выбр. дату";
                PosechStatGrid.Columns[5].HeaderText = "Группа за неделю";
                PosechStatGrid.Columns[6].HeaderText = "Группа за месяц";
            }
        }

        //получить данные посещаемости группы за неделю
        public void GetGroupStatWeek()
        {
            //получить статистику студента за неделю
            int numw = GetWeekNumByDate();
            if (posechGlobalStatTable.Rows.Count > 0)
            {
                DataRow[] studweek = posechGlobalStatTable.Select(" dt>= '" + starts[numw].ToShortDateString() + "' " +
                        " and dt <='" + ends[numw].ToShortDateString() + "'");

                DataRow[] studweekplus = posechGlobalStatTable.Select(" dt>= '" + starts[numw].ToShortDateString() + "' " +
                        " and dt <='" + ends[numw].ToShortDateString() + "' and atid = 2");

                PosechStatGrid.Rows[0].Cells[5].Value = studweekplus.Length.ToString();
                PosechStatGrid.Rows[1].Cells[5].Value = (studweek.Length - studweekplus.Length).ToString();
            }
            else
            {
                PosechStatGrid.Rows[0].Cells[5].Value = "-";
                PosechStatGrid.Rows[1].Cells[5].Value = "-";
            }
        }

        //получить данные посещаемости группы за месяц
        public void GetGroupStatMonth()
        {
            //получить статистику студента за неделю
            int numw = GetWeekNumByDate();
            if (posechGlobalStatTable.Rows.Count > 0)
            {
                DataRow[] studweek = posechGlobalStatTable.Select(" mes = " + posechDate.Month.ToString());

                DataRow[] studweekplus = posechGlobalStatTable.Select(" mes = " + posechDate.Month.ToString() + " and atid = 2");

                PosechStatGrid.Rows[0].Cells[6].Value = studweekplus.Length.ToString();
                PosechStatGrid.Rows[1].Cells[6].Value = (studweek.Length - studweekplus.Length).ToString();
            }
            else
            {
                PosechStatGrid.Rows[0].Cells[6].Value = "-";
                PosechStatGrid.Rows[1].Cells[6].Value = "-";
            }
        }


        /// <summary>
        /// внести тему из списка (из таблицы zanyatie)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void выбратьТемуИзСпискаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (personal_tabble.Rows.Count == 0) return;
            
            int num = prepod_table.CurrentCell.RowIndex;
            if (num < 0) return;

            DataRow dr = personal_tabble.Rows[num];
            string zan_id = dr[7].ToString();
            string pr_id = dr[12].ToString();

            DataTable dt = new DataTable();
            (new SqlDataAdapter("select id, tema + ' - ' + tematext as name from zanyatie " +
                " where predmet_id = " + pr_id, global_connection)).Fill(dt);

            if (dt.Rows.Count == 0)
            {
                MessageBox.Show("Нет тем в списке","Отказ операции", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            ListWindow lw = new ListWindow();
            lw.tbl = dt;
            lw.Text = "Выберите тему занятия из списка";
            lw.Width = 500;

            DialogResult dres = lw.ShowDialog();
            if (dres != DialogResult.OK) return;

            string sql = "update rasp set tema = @TEMA where id = @ID";
            SqlCommand cmd = new SqlCommand(sql, global_connection);
            cmd.Parameters.Add("@TEMA", SqlDbType.NVarChar).Value = lw.str_res;
            cmd.Parameters.Add("@ID", SqlDbType.Int).Value = zan_id;
            cmd.ExecuteNonQuery();

            prepod_table.Rows[num].Cells[6].Value = lw.str_res;
            lw.Dispose();
            GC.Collect();
        }


        /// <summary>
        /// пункт меню для вывода списка тем
        /// </summary>             
        ToolStripSeparator Separ = new ToolStripSeparator();
        string zan_id = string.Empty;
        string pr_id = string.Empty;
        private void prepod_table_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            // -- перестроить контекстное меню ячейки для отображения списка тем занятий
            // --  если они есть

            int i = 0;
            while (i < indrasp_context.Items.Count)
            {
                int tag = Convert.ToInt32(indrasp_context.Items[i].Tag);
                if (tag != -1)
                {
                    indrasp_context.Items.RemoveAt(i);
                }
                else
                    i++;
            }
            indrasp_context.Items.Remove(Separ);

            if (personal_tabble.Rows.Count == 0) return;

            int num = prepod_table.CurrentCell.RowIndex;
            if (num < 0) return;

            DataRow dr = personal_tabble.Rows[num];
            zan_id = dr[7].ToString();
            pr_id = dr[12].ToString();

            // удалить подпункт, связанный с отображением тем, если он есть

            DataTable dt = new DataTable();
            (new SqlDataAdapter("select id, tematext from zanyatie " +
                " where predmet_id = " + pr_id + " and tematext not like '%посещ%'", 
                global_connection)).Fill(dt);

            if (dt.Rows.Count == 0)
            {                                              
                return;
            }

            // создать подпункт тем и заполнить его           
            indrasp_context.Items.Add(Separ);

            foreach (DataRow drow in dt.Rows)
            {
                ToolStripItem it = new ToolStripMenuItem(drow[1].ToString());
                it.Tag = drow[0];
                it.Click += new EventHandler(выбратьТемуЗанятияВМеню);
                indrasp_context.Items.Add(it);
            }
        }

        /// <summary>
        /// действия при выборе темы занятия в меню
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void выбратьТемуЗанятияВМеню(object sender, EventArgs e)
        {
            int num = prepod_table.CurrentCell.RowIndex;
            if (num < 0) return;

            ToolStripItem it = ((ToolStripItem)sender);
            int tag = Convert.ToInt32(it.Tag);

            string sql = "update rasp set tema = @TEMA, razdel_tema_id = @TMID where id = @ID";
            SqlCommand cmd = new SqlCommand(sql, global_connection);
            cmd.Parameters.Add("@TEMA", SqlDbType.NVarChar).Value = it.Text;
            cmd.Parameters.Add("@ID", SqlDbType.Int).Value = zan_id;
            cmd.Parameters.Add("@TMID", SqlDbType.Int).Value = tag;
            cmd.ExecuteNonQuery();

            prepod_table.Rows[num].Cells[6].Value = it.Text;

        }

        /// <summary>
        /// расписание преподавателя на текующу неделю
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripButton25_Click(object sender, EventArgs e)
        {            
            SaveToExcel(active_user_id, active_user_name);
        }

        // загрузить итоговую таблицу по выбранному предмету
        private void tabControl1_Selected(object sender, TabControlEventArgs e)
        {
            if (e.TabPage != tabPage2) return;

            int sem = (DateTime.Now.Month >= 2 && DateTime.Now.Month <= 6) ? 2 : 1;                        

            string sql = "select * from dbo.GetResMRS(@PredmID, @DT1, @DT2) order by id";

            global_command = new SqlCommand(sql, global_connection);
            global_command.Parameters.Add("@PredmID", SqlDbType.Int).Value = id_predmet_in_tree;
            if (sem == 1)
            {
                global_command.Parameters.Add("@DT1", SqlDbType.DateTime).Value = Att1_1;
                global_command.Parameters.Add("@DT2", SqlDbType.DateTime).Value = Att1_2;
            }
            else
            {
                global_command.Parameters.Add("@DT1", SqlDbType.DateTime).Value = Att2_1;
                global_command.Parameters.Add("@DT2", SqlDbType.DateTime).Value = Att2_2;
            }

            DataTable MRSItog = new DataTable();
            (new SqlDataAdapter(global_command)).Fill(MRSItog);

            MRSItogGridView.Rows.Clear();
            while (MRSItogGridView.Columns.Count > 7)
                MRSItogGridView.Columns.RemoveAt(7);

            string[] ZanNames = ParseStrForMRS(MRSItog.Rows[0][8].ToString());
            string[] ZanIDs = ParseStrForMRS(MRSItog.Rows[1][8].ToString());
            string[] ZanBalls = null;

            int colCount = Convert.ToInt32(MRSItog.Rows[0][7].ToString());
            for (int i = 0; i < colCount; i++)
            {
                DataGridViewTextBoxColumn col = new DataGridViewTextBoxColumn();
                col.SortMode = DataGridViewColumnSortMode.Automatic;
                col.Tag = ZanIDs[i];
                col.HeaderText = ZanNames[i];
                col.AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;
                MRSItogGridView.Columns.Add(col);
            }

            int k = 0;
            for (int j = 2; j < MRSItog.Rows.Count; j++)
            {
                MRSItogGridView.Rows.Add();
                MRSItogGridView.Rows[k].Cells[0].Value = MRSItog.Rows[j][1].ToString();
                MRSItogGridView.Rows[k].Cells[1].Value = MRSItog.Rows[j][2].ToString();
                MRSItogGridView.Rows[k].Cells[2].Value = MRSItog.Rows[j][3].ToString();
                MRSItogGridView.Rows[k].Cells[3].Value = MRSItog.Rows[j][4].ToString();
                MRSItogGridView.Rows[k].Cells[4].Value = MRSItog.Rows[j][5].ToString();
                MRSItogGridView.Rows[k].Cells[5].Value = MRSItog.Rows[j][6].ToString();
                MRSItogGridView.Rows[k].Cells[6].Value = MRSItog.Rows[j][7].ToString();

                ZanBalls = ParseStrForMRS(MRSItog.Rows[j][8].ToString());
                for (int l = 0; l < colCount; l++)
                {
                    ZanBalls[l] = ZanBalls[l].Replace(".", ",");
                    MRSItogGridView.Rows[k].Cells[l + 7].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    MRSItogGridView.Rows[k].Cells[l + 7].Style.Format = "N1";
                    if (double.Parse(ZanBalls[l]) == 0)
                        MRSItogGridView.Rows[k].Cells[l + 7].Style.BackColor = Color.FromArgb(240, 240, 240);
                    MRSItogGridView.Rows[k].Cells[l + 7].Value =
                        string.Format("{0:F1}", Convert.ToDouble(ZanBalls[l]));
                }
                k++;                
            }
        }

        /// <summary>
        /// вывести описание темы занятия в статус вкладки итоговой таблицы МБРС предмета
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MRSItogGridView_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            toolStripStatusLabel2.Text = "";
            if (e.ColumnIndex <= 6) return;
            string sql = "select tematext from zanyatie where id = " +
                MRSItogGridView.Columns[e.ColumnIndex].Tag.ToString();
            DataTable dt = new DataTable();
            (new SqlDataAdapter(sql, global_connection)).Fill(dt);
            toolStripStatusLabel2.Text = dt.Rows[0][0].ToString();
        }

        private void MRSGridView_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            zadachaTextBox.Text = "";
            if (e.ColumnIndex <= 1) return;
            string sql = "select text from zadanie where id = " +
                MRSGridView.Columns[e.ColumnIndex].Tag.ToString();
            DataTable dt = new DataTable();
            (new SqlDataAdapter(sql, global_connection)).Fill(dt);
            zadachaTextBox.Text = dt.Rows[0][0].ToString();
        }


        // ----- отображение , ввод и редактирование сведений аттестаций и сессий ----


        public bool AttStarted = false;

        /// <summary>
        /// список аттестаций 
        /// </summary>
        public DataTable AttListTable = null;

        /// <summary>
        /// список предметов на закладке аттестаций
        /// </summary>
        public DataTable AttPredmListTable = null;

        /// <summary>
        /// список групп для вкладки аттестации и сессии
        /// </summary>
        public DataTable AttGrupList = null;

        /// <summary>
        /// список предметов для аттестации
        /// </summary>
        public DataTable AttPredmList = null;

        /// <summary>
        /// список студентов
        /// </summary>
        public DataTable AttStudentList = null;

        /// <summary>
        /// таблица для хранения оценок аттестации
        /// </summary>
        public DataTable AttItogTable = null;

        /// <summary>
        /// ячейка со списком оценок
        /// </summary>
        public DataGridViewComboBoxCell attcell = null;
        public DataGridViewTextBoxCell txtcell = new DataGridViewTextBoxCell();

        /// <summary>
        /// таблица с названиями отметок (для списка в таблице)
        /// </summary>
        public DataTable AttOtmTable = null;

        /// <summary>
        /// таблца для отображения внутри графика
        /// </summary>
        DataTable dt = null;

        /// <summary>
        /// выбор вкладок аттестаций, сессий, статистики
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void tabControl2_Selected(object sender, TabControlEventArgs e)
        {
            if (e.TabPage == SessiontabPage) SessionTabEnter();
            if (e.TabPage == ItogSessiontabPage) ItogSessionEnter();
        }

        /// <summary>
        /// выбор вкладки аттестаций
        /// </summary>
        public void AttTabEnter()
        {
            // заполнить список аттестаций
            AttListTable = new DataTable();
            attListComboBox.Items.Clear();

            // вывести список аттестаций
            (new SqlDataAdapter("select id, name, kod, tree_name from vid_zan where kod = '1а' or kod = '2а' or kod = '3а' or kod = '4а' ", 
                global_connection)).Fill(AttListTable);
            foreach (DataRow dr in AttListTable.Rows)
            {
                attListComboBox.Items.Add(dr["name"].ToString().ToUpper());
            }

            if (attListComboBox.Items.Count > 0)
                attListComboBox.SelectedIndex = 0;

        }

        /// <summary>
        /// заполнение списка групп на вкладке аттестаций
        /// </summary>
        public void grupAttListFill()
        {
            grupAttComboBox.Items.Clear();
            
            // вывести группы
            AttGrupList = new DataTable();
            if (attListComboBox.SelectedIndex <= 1)
                (new SqlDataAdapter("select grupa.id, grupa.name, grupa.kurs_id from grupa " +
                    " join specialnost on specialnost.id = grupa.specialnost_id " +
                    " where grupa.actual = 1 and " +
                    " fakultet_id = " + fakultet_id.ToString() +
                    " and specialnost.zaoch = 0 and grupa.kurs_id <=4 " +
                    "order by outorder",
                    global_connection)).Fill(AttGrupList);
            else
                (new SqlDataAdapter("select grupa.id, grupa.name, grupa.kurs_id from grupa " +
                    " join specialnost on specialnost.id = grupa.specialnost_id " +
                    " where grupa.actual = 1 and " +
                    " fakultet_id = " + fakultet_id.ToString() +
                    " and specialnost.zaoch = 0 and grupa.kurs_id <=3 " +
                    " order by outorder",
                global_connection)).Fill(AttGrupList);

            foreach (DataRow dr in AttGrupList.Rows)
            {
                grupAttComboBox.Items.Add(dr["name"].ToString().ToUpper());
            }

            if (grupAttComboBox.Items.Count > 0)
                grupAttComboBox.SelectedIndex = 0; 
        }

        private void grupAttComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            //заполнить сетку аттестации для группы
            FillAttTable();
        }

        private void attListComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            grupAttListFill();
        }


        /// <summary>
        /// заполнить сетку по выбранной аттестации для выбранной группы
        /// </summary>
        public void FillAttTable()
        {
            // получить возможные оценки аттестации
            attcell = new DataGridViewComboBoxCell();
            attcell.Items.Clear();
            AttOtmTable = new DataTable();
            (new SqlDataAdapter(
                "SELECT  vid_otmetka.id, vid_otmetka.str_name, vid_otmetka.str_alias " +
                    "  FROM vid_zan_otmetka INNER JOIN " +
                    "  vid_otmetka ON vid_zan_otmetka.vid_otmetka_id = vid_otmetka.id " +
                    "  WHERE vid_zan_otmetka.vid_zan_id = " + 
                    AttListTable.Rows[attListComboBox.SelectedIndex][0].ToString() +
                    "  ORDER BY vid_zan_otmetka.id",
                global_connection)).Fill(AttOtmTable);
            foreach (DataRow dr in AttOtmTable.Rows)
            {
                attcell.Items.Add(dr[2].ToString());
            }
            
                        
            // очистить сетку
            AttTableGridView.Rows.Clear();
            while (AttTableGridView.Columns.Count > 2)
                AttTableGridView.Columns.RemoveAt(2);

            AttTableGridView.Columns[0].CellTemplate = txtcell;
            AttTableGridView.Columns[1].CellTemplate = txtcell;

            AttItogTable = new DataTable();

            string AttId = AttListTable.Rows[attListComboBox.SelectedIndex][0].ToString();
            string GrID = AttGrupList.Rows[grupAttComboBox.SelectedIndex][0].ToString();
            string sql = string.Format("exec dbo.TGetAttestResult {0}, {1}, {2}", GrID, AttId, uch_god);           

            (new SqlDataAdapter(sql, global_connection)).Fill(AttItogTable);

            if (AttItogTable.Rows.Count == 0)
            {
                MessageBox.Show("Не удалось получить сведения об аттестации, попробуйте выполнить операцию позднее.",
                    "Ошибка выполнения операции.", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            //AttTableGridView.Rows.Add(); // первая строка
            int colCount = Convert.ToInt32(AttItogTable.Rows[0][1]);

            //названия предметов
            string[] PredmNames = ParseStrForMRS(AttItogTable.Rows[0][2].ToString());
            string[] PredmIds = ParseStrForMRS(AttItogTable.Rows[0][3].ToString());

            if (PredmIds.Length == 0)
            {
                MessageBox.Show("", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                AttTableGridView.Enabled = false;
                return;
            }
            else
            {
                AttTableGridView.Enabled = true;
            }

            DataGridViewCellStyle style1 = new DataGridViewCellStyle();
            style1.Alignment = DataGridViewContentAlignment.MiddleCenter;
            style1.BackColor = Color.FromArgb(240, 240, 240);

            DataGridViewCellStyle style2 = new DataGridViewCellStyle();
            style2.Alignment = DataGridViewContentAlignment.MiddleCenter;
            style2.BackColor = Color.LightYellow;

            for (int col = 0; col < colCount; col++)
            {
                DataGridViewComboBoxColumn cl = new DataGridViewComboBoxColumn();
                cl.CellTemplate = attcell;
                cl.SortMode = DataGridViewColumnSortMode.NotSortable;
                cl.Tag = PredmIds[col];
                cl.HeaderText = PredmNames[col];
                cl.AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;              
                cl.DefaultCellStyle = style1;
                AttTableGridView.Columns.Add(cl);


                DataGridViewTextBoxColumn cl1 = new DataGridViewTextBoxColumn();
                cl1.CellTemplate = txtcell;
                cl1.SortMode = DataGridViewColumnSortMode.NotSortable;
                cl1.HeaderText = "Баллы";
                cl1.AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;
                //cl1.ValueType = typeof(int);
                cl1.DefaultCellStyle = style2;
                
                AttTableGridView.Columns.Add(cl1);
                cl1 = null;
            }

            int rownum = 1;
            for (; rownum < AttItogTable.Rows.Count; rownum++)
            {
                //студент и его сумма
                AttTableGridView.Rows.Add();                
                
                DataRow drow = AttItogTable.Rows[rownum];
                
                AttTableGridView.Rows[rownum - 1].Cells[0].Value = drow[2];
                AttTableGridView.Rows[rownum - 1].Cells[1].Value = drow[6];

                string[] SessIds = ParseStrForMRS(drow[3].ToString());
                string[] Balls = ParseStrForMRS(drow[4].ToString());
                string[] OtmIds = ParseStrForMRS(drow[5].ToString());

                //Text = AttTableGridView.Columns.Count.ToString();

                // вывести оценки, баллы и задать Tag для ячеек
                int colnm = 2;
                for (int cnum = 0; cnum < colCount; cnum++)
                {
                    //вывод оценки
                    AttTableGridView.Rows[rownum - 1].Cells[colnm].Value = OtmValue(OtmIds[cnum]);
                    AttTableGridView.Rows[rownum - 1].Cells[colnm].Tag = SessIds[cnum];
                    colnm++;

                    //вывод баллов и задание Tag
                    AttTableGridView.Rows[rownum - 1].Cells[colnm].Tag = SessIds[cnum];
                    AttTableGridView.Rows[rownum - 1].Cells[colnm].Value = Balls[cnum].Replace(".", ",");
                    colnm++;
                }
            }

            // построение диагаммы
            dt = new DataTable();
            dt = AttItogTable;

            dt.DefaultView.Sort = "summa";
            
            dt.Rows.RemoveAt(0);
            //if (chartAtt.Series.Count > 0) chartAtt.Series.RemoveAt(0);
            Series ser = new Series("Att", ViewType.Bar);            
            ser.DataSource = dt;
            ser.ArgumentDataMember = "names";
            ser.ArgumentScaleType = ScaleType.Qualitative;
            ser.ValueScaleType = ScaleType.Numerical;
            ser.ValueDataMembers.AddRange(new string[] { "summa" });
            //chartAtt.Series.Add(ser);

            // подведение итогов  -------------------------------------                     
            DataGridViewRow dgvr = new DataGridViewRow();
            DataGridViewRow dgvr2 = new DataGridViewRow();
            DataGridViewRow dgvr2p = new DataGridViewRow();
            DataGridViewRow dgvr3 = new DataGridViewRow();
            DataGridViewRow dgvr4 = new DataGridViewRow();
            DataGridViewRow dgvr5 = new DataGridViewRow();
            
            int ii = 0;
            for (ii = 0; ii < AttTableGridView.Columns.Count; ii++)
            {                
                DataGridViewTextBoxCell c = new DataGridViewTextBoxCell();
                c.Style.BackColor = Color.Red;
                c.Style.ForeColor = Color.White;
                c.Style.Font = new Font("Tahoma", 12.0f, FontStyle.Regular);
                dgvr.Cells.Add(c);
                c.ReadOnly = true;

                DataGridViewTextBoxCell cc2 = new DataGridViewTextBoxCell();
                DataGridViewTextBoxCell cc2p = new DataGridViewTextBoxCell();
                DataGridViewTextBoxCell cc3 = new DataGridViewTextBoxCell();
                DataGridViewTextBoxCell cc4 = new DataGridViewTextBoxCell();
                DataGridViewTextBoxCell cc5 = new DataGridViewTextBoxCell();
                //c.Style.BackColor = Color.LightYellow;
                //c.Style.ForeColor = Color.Navy;
                dgvr2.Cells.Add(cc2);
                dgvr2p.Cells.Add(cc2p);
                dgvr3.Cells.Add(cc3);
                dgvr4.Cells.Add(cc4);
                dgvr5.Cells.Add(cc5);
                cc2.ReadOnly = true;
                cc3.ReadOnly = true;
                cc2p.ReadOnly = true;
                cc4.ReadOnly = true;
                cc5.ReadOnly = true;

                double avg = 0.0;
                int c2 = 0, c2p = 0, c3 = 0, c4 = 0, c5 = 0;
                if (ii > 0)
                {                    
                    for (int r = 0; r < AttTableGridView.Rows.Count; r++)
                    {
                        if (ii % 2 == 0) //сред. оценка
                        {
                            string otm = AttTableGridView.Rows[r].Cells[ii].Value.ToString();
                            switch (otm)
                            {
                                case "2": avg += 2;
                                    c2++;
                                    break;
                                case "2+": avg += 2;
                                    c2p++;
                                    break;
                                case "3": avg += 3;
                                    c3++;
                                    break;
                                case "3+": avg += 3;
                                    c3++;
                                    break;
                                case "4": avg += 4;
                                    c4++;
                                    break;
                                case "4+": avg += 4;
                                    c4++;
                                    break;
                                case "5": avg += 5;
                                    c5++;
                                    break;
                                case "5+": avg += 5;
                                    c5++;
                                    break;
                            }
                        }
                        else //сред. балл
                        {
                            double ball = Convert.ToDouble(AttTableGridView.Rows[r].Cells[ii].Value);
                            avg += ball;
                        }
                    }
                }

                c.Value = string.Format("{0:F1}", avg / AttTableGridView.Rows.Count);
                if(ii>0 && ii%2==0)
                {
                    dgvr2.Cells[ii].Value = c2;
                    dgvr2.Cells[ii].ToolTipText = "Количество двоек по предмету '" + AttTableGridView.Columns[ii].HeaderText + "'";

                    dgvr2p.Cells[ii].Value = c2p;
                    dgvr2p.Cells[ii].ToolTipText = "Количество двоек c плюсом по предмету '" + AttTableGridView.Columns[ii].HeaderText + "'";
                    
                    dgvr3.Cells[ii].Value = c3;
                    dgvr3.Cells[ii].ToolTipText = "Количество троек по предмету '" + AttTableGridView.Columns[ii].HeaderText + "'";
                    
                    dgvr4.Cells[ii].Value = c4;
                    dgvr4.Cells[ii].ToolTipText = "Количество четвёрок по предмету '" + AttTableGridView.Columns[ii].HeaderText + "'";

                    dgvr5.Cells[ii].Value = c5;
                    dgvr5.Cells[ii].ToolTipText = "Количество пятёрок по предмету '" + AttTableGridView.Columns[ii].HeaderText + "'";

                    dgvr2.Cells[1].Value = Convert.ToInt32(dgvr2.Cells[1].Value) + c2;
                    dgvr2p.Cells[1].Value = Convert.ToInt32(dgvr2p.Cells[1].Value) + c2p;
                    dgvr3.Cells[1].Value = Convert.ToInt32(dgvr3.Cells[1].Value) + c3;
                    dgvr4.Cells[1].Value = Convert.ToInt32(dgvr4.Cells[1].Value) + c4;
                    dgvr5.Cells[1].Value = Convert.ToInt32(dgvr5.Cells[1].Value) + c5;
                }
            }

            dgvr.Cells[0].Value = "Средние показатели";
            dgvr2.Cells[0].Value = "Количество оценки '2'";
            dgvr2p.Cells[0].Value = "Количество оценки '2+'";
            dgvr3.Cells[0].Value = "Количество оценки '3'";
            dgvr4.Cells[0].Value = "Количество оценки '4'";
            dgvr5.Cells[0].Value = "Количество оценки '5'";
            
            AttTableGridView.Rows.Add(dgvr);
            AttTableGridView.Rows.Add(dgvr2);
            AttTableGridView.Rows.Add(dgvr2p);
            AttTableGridView.Rows.Add(dgvr3);
            AttTableGridView.Rows.Add(dgvr4);
            AttTableGridView.Rows.Add(dgvr5);
        }

        /// <summary>
        /// получить название отметки по ее индексу
        /// </summary>
        /// <param name="ind"></param>
        /// <returns></returns>
        public string OtmValue(string ind)
        {
            string res = "";

            foreach (DataRow dr in AttOtmTable.Rows)
            {
                if (dr[0].ToString() == ind)
                    res = dr[2].ToString();
            }

            return res;
        }

        /// <summary>
        /// получить индекс отметки по ее названию
        /// </summary>
        /// <param name="val"></param>
        /// <returns></returns>
        public string OtmIndex(string val)
        {
            string res = "";

            foreach (DataRow dr in AttOtmTable.Rows)
            {
                if (dr[2].ToString() == val)
                    res = dr[0].ToString();
            }
            return res;
        }

        // произвести вывод списка предметов и студентов группы
        public void FillPredmAndStudentAttList()
        {
            /*uchGodAttComboBox.Items.Clear();
            studAttComboBox.Items.Clear();

            if (grupAttComboBox.Items.Count == 0) return;
            
            int ind = 0;
            if (grupAttComboBox.SelectedIndex>=0) ind = grupAttComboBox.SelectedIndex;

            // получить ид группы -------------- 
            string gr_id = AttGrupList.Rows[ind][0].ToString();

            string sem_number = " and " + 
                ((attListComboBox.SelectedIndex <= 1) ?
                    " (predmet.semestr = 1 or predmet.semestr = 3 or predmet.semestr = 5 or predmet.semestr = 7)" :
                    " (predmet.semestr = 2 or predmet.semestr = 4 or predmet.semestr = 6)");

            // вывести список предметов в данной аттестации для данной группы
            AttPredmList = new DataTable();
            (new SqlDataAdapter("select predmet.id, predmet.name_krat, predmet.name from predmet " +
                " where predmet.actual = 1 and " +
                " predmet.grupa_id = " + gr_id + sem_number + 
                " order by predmet.name",
                global_connection)).Fill(AttPredmList);

            foreach (DataRow dr in AttPredmList.Rows)
            {
                uchGodAttComboBox.Items.Add(dr[1].ToString().ToUpper());
            }

            if (uchGodAttComboBox.Items.Count > 0)
                uchGodAttComboBox.SelectedIndex = 0;

            // получить список студентов группы и показать его ============
            AttStudentList = new DataTable();
            (new SqlDataAdapter("select * from dbo.GetStudentList(" + gr_id + ")", global_connection)).Fill(AttStudentList);

            foreach (DataRow dr in AttStudentList.Rows)
            {
                studAttComboBox.Items.Add(dr[1].ToString());
            }

            if (studAttComboBox.Items.Count > 0)
                studAttComboBox.SelectedIndex = 0;*/
        }


        /// <summary>
        /// обработка входа на страницы аттестации, сесии, итого и статистки успеваемости
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void content_Selected(object sender, TabControlEventArgs e)
        {            
            if (e.TabPage == tabPageAttest)
                AttTabEnter();
            if (e.TabPage == SessiontabPage)
                SessionTabEnter();
            if (e.TabPage == ItogSessiontabPage)
                ItogSessionEnter();
            
        }

        /// <summary>
        /// отправка результатов аттестации во внешний файл с получением статистики
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripButton26_Click(object sender, EventArgs e)
        {                        
            /*Bitmap bm = new Bitmap(chartAtt.Width,chartAtt.Height);
            Graphics g = Graphics.FromImage(bm);            
            g.CopyFromScreen(chartAtt.Left, chartAtt.Top, chartAtt.Width+chartAtt.Left, chartAtt.Top+chartAtt.Height, 
                new Size(chartAtt.Width, chartAtt.Height), CopyPixelOperation.SourceCopy);
            Clipboard.SetImage(bm);*/
            //attPanel.Handle

            int i = 0;            

            if (AttTableGridView.Rows.Count == 0)
            {
                MessageBox.Show("Нет данных об аттестации.", "Отказ операции.",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            double summ = 0;
            for (i = 1; i < AttTableGridView.Rows.Count; i++)
            {
                summ += Convert.ToDouble(AttTableGridView.Rows[i].Cells[1].Value);
            }

            if (summ == 0)
            {
                MessageBox.Show("В таблице нет данных по баллам для построения отчёта.", "Отказ операции.",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            CellRange cr = null; //диапазон ячеек на рабочем листе книги
            string[] Letters = new string[]{"A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T",
                "U","V","W","X","Y","Z","AA","AB","AC","AD","AE","AF"};

            ExcelFile excel = new ExcelFile();

            ExcelWorksheet sheet1 = excel.Worksheets.Add("Таблица");
            ExcelWorksheet sheet2 = excel.Worksheets.Add("Итоги");
            ExcelWorksheet sheet3 = excel.Worksheets.Add("Рейтинг");         

            string attstr = "";
            if (attListComboBox.SelectedIndex % 2 == 1)
            {
                attstr = "Итоги второй аттестации";
            }
            else
            {
                attstr = "Итоги первой аттестации";
            }


            // -- запрос на сохранение

            saveExcel.Title = "Выберите или введите имя для файла отчёта";
            saveExcel.FileName = attstr.ToUpper() + " В ГРУППЕ " + grupAttComboBox.Text + ".xls";
            if (saveExcel.ShowDialog() != DialogResult.OK) return;
            string Path = saveExcel.FileName;

            // --- создание страниц

            sheet1.Cells[0, 1].Value = attstr.ToUpper() + " В ГРУППЕ " + grupAttComboBox.Text;
            sheet1.Cells[0, 1].Style.Font.Weight = ExcelFont.MaxWeight;

            sheet1.Cells[1, 0].Value = "Студент";
            sheet1.Cells[1, 0].Style.VerticalAlignment = VerticalAlignmentStyle.Center;
            sheet1.Cells[1, 0].Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            sheet1.Cells[1, 0].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
            sheet1.Cells[1, 1].Value = "Всего\nбаллов";
            sheet1.Cells[1, 1].Style.VerticalAlignment = VerticalAlignmentStyle.Center;
            sheet1.Cells[1, 1].Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            sheet1.Cells[1, 1].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);

            sheet1.Columns[1].Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;

            for (i = 0; i < AttTableGridView.Rows.Count - 6; i++)
            {
                sheet1.Cells[i + 2, 0].Value = AttTableGridView.Rows[i].Cells[0].Value;
                sheet1.Cells[i + 2, 0].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
                
                sheet1.Cells[i + 2, 1].Value = AttTableGridView.Rows[i].Cells[1].Value;
                sheet1.Cells[i + 2, 1].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
                if (AttTableGridView.Rows[i].Cells[1].Value.ToString() == "0")
                    sheet1.Cells[i + 2, 1].Style.Font.Weight = ExcelFont.BoldWeight;
                sheet1.Rows[i + 2].Height = 15 * 20;
            }

            int x = 0;
            for (i = 2; i < AttTableGridView.Columns.Count; i += 2)
            {
                cr = sheet1.Cells.GetSubrange(Letters[i] + 2.ToString(), Letters[i+1] + 2.ToString());
                //MessageBox.Show(Letters[i] + 2.ToString() + " : " + Letters[i+1] + 2.ToString());                 
                cr.Merged = true;
                cr.SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);              
                cr.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
                cr.Style.VerticalAlignment = VerticalAlignmentStyle.Center;                
                //cr.Merged = false;
                sheet1.Cells[1, i].Value = AttTableGridView.Columns[i].HeaderText.ToUpper();
                sheet1.Cells[1, i].Style.Rotation = 90;
                sheet1.Columns[i].AutoFit();

                sheet1.Columns[i].Width = 5 * 256;
                sheet1.Columns[i].Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
                sheet1.Columns[i + 1].Width = 5 * 256;
                sheet1.Columns[i + 1].Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;

                for (int j = 0; j < AttTableGridView.Rows.Count - 6; j++)
                {
                    sheet1.Cells[j+2, i].Value = AttTableGridView.Rows[j].Cells[i].Value;
                    sheet1.Cells[j+2, i].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);                    

                    sheet1.Cells[j+2, i + 1].Value = AttTableGridView.Rows[j].Cells[i + 1].Value;
                    sheet1.Cells[j+2, i + 1].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);

                    //sheet1.Cells[j + 2, i].SetBorders(MultipleBorders.Right, Color.Black, GemBox.Spreadsheet.LineStyle.Dotted);

                    if (AttTableGridView.Rows[j].Cells[i].Value.ToString() == "2" || AttTableGridView.Rows[j].Cells[i].Value.ToString() == "2+")
                    {
                        sheet1.Cells[j + 2, i].Style.Font.Weight = ExcelFont.BoldWeight;
                        sheet1.Cells[j + 2, i].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thick);
                    }
                }

                string sql = string.Format("select dbo.GetPrepodFIOByID(prepod.id) from prepod " +
                    " join predmet on predmet.prepod_id = prepod.id " +
                    " where predmet.id = {0}", AttTableGridView.Columns[i].Tag);
                DataTable prname = new DataTable();
                (new SqlDataAdapter(sql, global_connection)).Fill(prname);

                x = AttTableGridView.Rows.Count - 6 + 2;
                cr = sheet1.Cells.GetSubrange(Letters[i] + (x+1).ToString(), Letters[i + 1] + (x+1).ToString());
                //MessageBox.Show(Letters[i] + 2.ToString() + " : " + Letters[i+1] + 2.ToString());                 
                cr.Merged = true;
                cr.SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
                cr.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
                cr.Style.VerticalAlignment = VerticalAlignmentStyle.Center;
                //cr.Merged = false;
                sheet1.Cells[x, i].Value = prname.Rows[0][0].ToString();
                sheet1.Cells[x, i].Style.Rotation = 90;                
            }

            sheet1.Rows[1].Height = 135 * 20;
            sheet1.Rows[x].Height = 94 * 20;
            sheet1.Columns[0].Width = 18 * 256;
            sheet1.Columns[1].Width = 8 * 256;

            sheet1.PrintOptions.HeaderMargin = 0.0;
            sheet1.PrintOptions.FooterMargin = 0.0;
            sheet1.PrintOptions.Portrait = false;

            // ---- кон: первый лист


            // --- второй лист - статистика -------------------            

            sheet2.Cells[1, 0].Value = "";
            sheet2.Cells[1, 0].Style.VerticalAlignment = VerticalAlignmentStyle.Center;
            sheet2.Cells[1, 0].Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            sheet2.Cells[1, 0].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
            sheet2.Cells[1, 1].Value = "";
            sheet2.Cells[1, 1].Style.VerticalAlignment = VerticalAlignmentStyle.Center;
            sheet2.Cells[1, 1].Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            sheet2.Cells[1, 1].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);

            sheet2.Columns[1].Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            
            int r = 3;
            for (i = AttTableGridView.Rows.Count - 6; i < AttTableGridView.Rows.Count; i++)
            {
                sheet2.Cells[r, 0].Value = AttTableGridView.Rows[i].Cells[0].Value;
                sheet2.Cells[r, 0].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);

                sheet2.Cells[r, 1].Value = AttTableGridView.Rows[i].Cells[1].Value;
                sheet2.Cells[r, 1].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);                
                sheet2.Rows[i + 2].Height = 15 * 20;
                r++;
            }

            sheet2.Columns[0].Width = 256 * 25;
            sheet2.Cells[2, 0].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
            sheet2.Cells[2, 1].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
            sheet2.Cells[2, 1].Value = "Ср. балл";            
            sheet2.Columns[1].AutoFit();

            for (i = 2; i < AttTableGridView.Columns.Count; i += 2)
            {
                cr = sheet2.Cells.GetSubrange(Letters[i] + 2.ToString(), Letters[i + 1] + 2.ToString());                
                cr.Merged = true;
                cr.SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
                cr.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
                cr.Style.VerticalAlignment = VerticalAlignmentStyle.Center;
                sheet2.Cells[1, i].Value = AttTableGridView.Columns[i].HeaderText.ToUpper();
                sheet2.Cells[1, i].Style.Rotation = 90;
                sheet2.Columns[i].AutoFit();

                sheet2.Columns[i].Width = 5 * 256;
                sheet2.Columns[i].Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
                sheet2.Columns[i + 1].Width = 5 * 256;
                sheet2.Columns[i + 1].Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;

                sheet2.Cells[2, i].Value = "Ср.оц.";
                sheet2.Cells[2, i].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);

                sheet2.Cells[2, i + 1].Value = "Ср.балл";
                sheet2.Cells[2, i+1].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);

                sheet2.Columns[i].AutoFit();
                sheet2.Columns[i + 1].AutoFit();

                r = 3;
                for (int j = AttTableGridView.Rows.Count - 6; j < AttTableGridView.Rows.Count; j++)
                {
                    sheet2.Cells[r, i].Value = AttTableGridView.Rows[j].Cells[i].Value;
                    sheet2.Cells[r, i].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);

                    sheet2.Cells[r, i + 1].Value = AttTableGridView.Rows[j].Cells[i + 1].Value;
                    sheet2.Cells[r, i + 1].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);                    

                    r++;
                }

                string sql = string.Format("select dbo.GetPrepodFIOByID(prepod.id) from prepod " +
                    " join predmet on predmet.prepod_id = prepod.id " +
                    " where predmet.id = {0}", AttTableGridView.Columns[i].Tag);
                DataTable prname = new DataTable();
                (new SqlDataAdapter(sql, global_connection)).Fill(prname);

                x = 9;
                cr = sheet2.Cells.GetSubrange(Letters[i] + (x + 1).ToString(), Letters[i + 1] + (x + 1).ToString());                
                cr.Merged = true;
                cr.SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
                cr.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
                cr.Style.VerticalAlignment = VerticalAlignmentStyle.Center;
                //cr.Merged = false;
                sheet2.Cells[x, i].Value = prname.Rows[0][0].ToString();
                sheet2.Cells[x, i].Style.Rotation = 90;
            }

            sheet2.Rows[1].Height = 135 * 20;
            sheet2.Rows[x].Height = 94 * 20;
            sheet2.Columns[0].Width = 256 * 20;
            sheet2.Columns[1].Width = 8 * 256;


            sheet2.PrintOptions.HeaderMargin = 0.0;
            sheet2.PrintOptions.FooterMargin = 0.0;
            sheet2.PrintOptions.Portrait = false;

            sheet2.Cells[0, 3].Value = attstr.ToUpper() + " В ГРУППЕ " + grupAttComboBox.Text + " (статистика по предметам)";
            sheet2.Cells[0, 3].Style.Font.Weight = ExcelFont.MaxWeight;


            // --- лист 3 - рейтинг


            SortedDictionary<int, string> ball_student = new SortedDictionary<int, string>(new DescendingComparer<int>());
            int ii = 0, jj = 0;
            
            for (jj = 0; jj < dt.Rows.Count; jj++)
            {
                int ball = Convert.ToInt32(dt.Rows[jj][6]);
                string stud = dt.Rows[jj][2].ToString();
                ball_student[ball] = stud;
            }            

            jj = 2;
            ii = 1;

            sheet3.Cells[1, 1].Value = "№";
            sheet3.Cells[1, 1].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
            sheet3.Cells[1, 2].Value = "ФИО студента";
            sheet3.Cells[1, 2].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
            sheet3.Cells[1, 3].Value = "Сумма баллов";
            sheet3.Cells[1, 3].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);

            sheet3.Columns[1].Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            sheet3.Columns[2].Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            sheet3.Columns[3].Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;

            foreach (int key in ball_student.Keys)
            {
                sheet3.Cells[jj, 1].Value = ii;
                sheet3.Cells[jj, 1].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
                sheet3.Cells[jj, 2].Value = ball_student[key];
                sheet3.Cells[jj, 2].Style.HorizontalAlignment = HorizontalAlignmentStyle.Left;
                sheet3.Cells[jj, 2].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
                sheet3.Cells[jj, 3].Value = key;
                sheet3.Cells[jj, 3].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
                jj++;
                ii++;
            }

            sheet3.Columns[1].AutoFit();
            sheet3.Columns[2].AutoFit();
            sheet3.Columns[3].AutoFit();

            sheet3.Cells[0, 2].Value = "Рейтинг студентов (группа " + grupAttComboBox.Text + ")";
            sheet3.Cells[0, 3].Style.Font.Weight = ExcelFont.MaxWeight;

            excel.SaveXls(Path);
            Process.Start(Path);
        }


        /// <summary>
        /// количество выпадающих строк таблицы аттестации 
        /// </summary>
        int span = 6;

        /// <summary>
        /// начало редактирования ячейки
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>        
        private void AttTableGridView_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            if (e.RowIndex >= AttTableGridView.Rows.Count - span) return;
            currentval = AttTableGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
        }

        /// <summary>
        /// список разных оценок 
        /// </summary>
        List<string> DiffOtms = new List<string>();

        public void GetDiffOtms()
        {
            DiffOtms.Clear();

            string str = "";

            int i = 2;
            for (int k = 0; k < AttTableGridView.Rows.Count; k++)
            {
                for (i = 2; i < AttTableGridView.Columns.Count - 1; i += 2)
                {
                    string s = AttTableGridView.Rows[k].Cells[i].Value.ToString();
                    if (!DiffOtms.Contains(s))
                    {
                        DiffOtms.Add(s);
                        str += (s + " ");
                    }
                }
            }
        }

        /// <summary>
        /// фиксация значений для аттестации в БД
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void AttTableGridView_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex <= 1) return;
            if (e.RowIndex >= AttTableGridView.Rows.Count - span) return;

            if (e.ColumnIndex % 2 == 0) // редактирование оценки
            {
                string newval = OtmIndex(AttTableGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString());
                string sessid = AttTableGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].Tag.ToString();

                string sql = "update session set otmetka_id = @OTM where id = @SESSID";

                global_command = new SqlCommand(sql, global_connection);
                global_command.Parameters.Add("@OTM", SqlDbType.Int).Value = newval;
                global_command.Parameters.Add("@SESSID", SqlDbType.Int).Value = sessid;
                global_command.ExecuteNonQuery();

                // пересчет статистики оценок
                int c2 = 0, c2p = 0, c3 = 0, c4 = 0, c5 = 0;
                double avg = 0.0;
                for (int r = 0; r < AttTableGridView.Rows.Count - 6; r++)
                {
                    string otm = AttTableGridView.Rows[r].Cells[e.ColumnIndex].Value.ToString();
                    switch (otm)
                    {
                        case "2": avg += 2;
                            c2++;
                            break;
                        case "2+": avg += 2;
                            c2p++;
                            break;
                        case "3": avg += 3;
                            c3++;
                            break;
                        case "3+": avg += 3;
                            c3++;
                            break;
                        case "4": avg += 4;
                            c4++;
                            break;
                        case "4+": avg += 4;
                            c4++;
                            break;
                        case "5": avg += 5;
                            c5++;
                            break;
                        case "5+": avg += 5;
                            c5++;
                            break;
                    }
                }

                int row = AttTableGridView.Rows.Count;
                int col = e.ColumnIndex;
                AttTableGridView.Rows[row - 6].Cells[col].Value =
                    string.Format("{0:F1}", avg / (AttTableGridView.Rows.Count - 6));

                AttTableGridView.Rows[row - 5].Cells[col].Value = c2;
                AttTableGridView.Rows[row - 4].Cells[col].Value = c2p;
                AttTableGridView.Rows[row - 3].Cells[col].Value = c3;
                AttTableGridView.Rows[row - 2].Cells[col].Value = c4;
                AttTableGridView.Rows[row - 1].Cells[col].Value = c5;

                
                for (int rr = AttTableGridView.Rows.Count - 5; rr < AttTableGridView.Rows.Count; rr++)
                {
                    c2 = 0;
                    for (int cc = 2; cc < AttTableGridView.Columns.Count - 1; cc += 2)
                    {
                        c2 += Convert.ToInt32(AttTableGridView.Rows[rr].Cells[cc].Value);                        
                    }
                    AttTableGridView.Rows[rr].Cells[1].Value = c2;
                }

            }
            else //редактирование баллов
            {
                string newvalue = "";

                if (AttTableGridView.Rows[e.RowIndex].Cells[e.ColumnIndex] != null)
                {
                    if (AttTableGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].Value != null)
                        newvalue = AttTableGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
                    else
                        newvalue = "0";
                }
                else
                {
                    newvalue = "0";
                }

                int d = 0;
                if (!int.TryParse(newvalue, out d))
                {
                    AttTableGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = currentval;
                    return;
                }

                AttTableGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = newvalue;

                string sessid = AttTableGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].Tag.ToString();

                string sql = "update session set ball = @BALL where id = @SESSID";

                global_command = new SqlCommand(sql, global_connection);
                global_command.Parameters.Add("@BALL", SqlDbType.Int).Value = newvalue;
                global_command.Parameters.Add("@SESSID", SqlDbType.Int).Value = sessid;
                global_command.ExecuteNonQuery();

                //пересчет суммы и других показателей            
                int summ = 0;
                for (int i = 3; i <= AttTableGridView.Columns.Count; i += 2)
                {
                    summ += Convert.ToInt32(AttTableGridView.Rows[e.RowIndex].Cells[i].Value);
                }

                AttTableGridView.Rows[e.RowIndex].Cells[1].Value = summ;

                double sum = 0.0;
                for (int j = 0; j < AttTableGridView.Rows.Count - 6; j++)
                    sum += Convert.ToInt32(AttTableGridView.Rows[j].Cells[1].Value);

                AttTableGridView.Rows[AttTableGridView.Rows.Count - 6].Cells[1].Value =
                    string.Format("{0:F1}", sum / (AttTableGridView.Rows.Count - 6));

                sum = 0.0;
                for (int j = 0; j < AttTableGridView.Rows.Count - 6; j++)
                    sum += Convert.ToInt32(AttTableGridView.Rows[j].Cells[e.ColumnIndex].Value);

                AttTableGridView.Rows[AttTableGridView.Rows.Count - 6].Cells[e.ColumnIndex].Value =
                    string.Format("{0:F1}", sum / (AttTableGridView.Rows.Count - 6));

                //обновить график
                if (dt.IsInitialized)
                {
                    dt.Rows[e.RowIndex][6] = summ;                                                            
                    //chartAtt.Series[0].ValueDataMembers.RemoveAt(0);
                    //chartAtt.Series[0].ValueDataMembers.AddRange(new string[] { "summa" });
                }
            }
        }

        /// <summary>
        /// отправка результатов аттестации во внешний файл с получением статистики
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void saveExcel_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
        {
            
        }


        // управление программой MS Word -------------------------
        
        /// <summary>
        /// создать документ Word и получить ссылку на него
        /// </summary>
        /// <returns>возвращает ссылку на созданный экземпляр документа MS Word</returns>
        public Word.Document CreateWordDoc()
        {
            // создать новый экземпляр приложения
            wa = new Word.Application();


            // и добавить в него новый документ 
            object template = Type.Missing;
            object newtemplate = false;
            object documentType = Word.WdNewDocumentType.wdNewBlankDocument;
            object visible = true;

            Word.Document document = wa.Documents.Add(ref template,
                ref newtemplate,
                ref documentType,
                ref visible);

            return document;
        }

        /// <summary>
        ///  создать документ
        /// </summary>
        /// <param name="wa">ссылка на экземпляр приложения Word</param>
        /// <returns></returns>
        public static Word.Document CreateNewWordDoc(ref Word.Application wa)
        {
            // создать новый экземпляр приложения
            wa = new Word.Application();

            // и добавить в него новый документ 
            object template = Type.Missing;
            object newtemplate = false;
            object documentType = Word.WdNewDocumentType.wdNewBlankDocument;
            object visible = true;

            Word.Document document = wa.Documents.Add(ref template,
                ref newtemplate,
                ref documentType,
                ref visible);

            return document;
        }

        /// <summary>
        /// создать документ из файла
        /// </summary>
        /// <param name="wa">ссылка на экземпляр приложения Word</param>
        /// <param name="FName">путь к файлу</param>
        /// <returns></returns>
        public static Word.Document CreateNewWordDoc(ref Word.Application wa, string FName)
        {
            // создать новый экземпляр приложения
            wa = new Word.Application();

            // и добавить в него новый документ 
            object template = FName;
            object newtemplate = false;
            object documentType = Word.WdNewDocumentType.wdNewBlankDocument;
            object visible = true;

            Word.Document document = wa.Documents.Add(ref template,
                ref newtemplate,
                ref documentType,
                ref visible);

            return document;
        }

        /// <summary>
        /// сохранить документ Word с указанным именем
        /// </summary>
        /// <param name="FileName">имя файла для сохраняемого документа</param>
        /// <param name="doc">ссылка на сохраняемый документ</param>
        public static void SaveWordDoc(string FileName, ref Word.Document doc)
        {
            object fileName = FileName;
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
        }

        /// <summary>
        /// закрыть приложение Word
        /// </summary>
        public void WordQuit()
        {
            object saveChages = false;
            object originalFormat = Word.WdOriginalFormat.wdWordDocument;
            object routeDocument = Type.Missing;
            wa.Quit(ref saveChages, ref originalFormat, ref routeDocument);
        }

        /// <summary>
        /// закрыть приложение Word
        /// </summary>
        /// <param name="wa">ссылка на закрываемое приложение</param>
        public static void WordQuit(Word.Application wa)
        {
            object saveChages = false;
            object originalFormat = Word.WdOriginalFormat.wdWordDocument;
            object routeDocument = Type.Missing;
            wa.Quit(ref saveChages, ref originalFormat, ref routeDocument);
        }

        /// <summary>
        /// добавить новый параграф в документ
        /// </summary>
        /// <param name="doc">ссылка на документ</param>
        /// <param name="text">текст, вставляемый в параграф</param>
        /// <param name="text">выравнивание текста</param>
        /// <returns></returns>
        public static Word.Range AddWordDocParagraph(ref Word.Document doc, string text, Word.WdParagraphAlignment al)
        {
            int parcount = doc.Paragraphs.Count;
            object rng = Type.Missing;            

            doc.Paragraphs.Add(ref rng);
            Word.Range Range = doc.Paragraphs[parcount+1].Range;
            Range.Select();
            Range.Text = text;                        
            Range.ParagraphFormat.Alignment = al;

            return Range;
        }

        // ------------------------------------- кон: работа с документами Word    


        // --- --  работа с данными по сесиям ------------------


        /// <summary>
        /// список групп на вкладке сессий
        /// </summary>
        public DataTable SessGrupTable = null;

        /// <summary>
        /// список сессий
        /// </summary>
        public DataTable SessTable = null;

        /// <summary>
        /// список предметов и форм контроля по ним
        /// </summary>
        public DataTable SessPredmetTable = null;

        /// <summary>
        /// таблица результата сессии по предмету
        /// </summary>
        public DataTable SessionResultTable = null;

        /// <summary>
        /// шапка в ведомости
        /// </summary>
        public string SessVedTitle = string.Empty;
        public string SessVedKod = string.Empty;

        /// <summary>
        /// начальная загрузка данных в элементы вкладки "сессия"
        /// </summary>
        public void SessionTabEnter()
        {
            SessGruplistBox.Items.Clear();
            global_query = "select id, name, kurs_id, potok, mrs from grupa where " +
                " fakultet_id = " + fakultet_id.ToString() +
                " and actual = 1 order by outorder";
            SessGrupTable = new DataTable();
            (new SqlDataAdapter(global_query, global_connection)).Fill(SessGrupTable);

            foreach (DataRow r in SessGrupTable.Rows)
            {
                SessGruplistBox.Items.Add(r[1].ToString());
            }

            if (SessGruplistBox.Items.Count > 0)
                SessGruplistBox.SelectedIndex = 0;            
        }


        /// <summary>
        /// загрузка списка форм контроля по сессии в зависимости от курса
        /// </summary>
        public void FillSessControlList()
        {
            SesslistBox.Items.Clear();

            SessionGridView.Columns[3].Visible = false;
            int i = SessGruplistBox.SelectedIndex;
            if (i < 0) return;

            int kurs = Convert.ToInt32(SessGrupTable.Rows[i][2]);

            global_query = "select id, name, kod from vid_rab where ";

            if (kurs <= 3)
            {
                global_query += "kod='зс' or kod='лс'";
            }

            if (kurs == 4)
            {
                global_query += "kod='зс' or kod='пп'";
            }

            if (kurs == 5)
            {
                global_query += "kod='зс' or kod='здр' or kod='гэ'";
            }

            global_query += " order by outorder";

            SessTable = new DataTable();
            (new SqlDataAdapter(global_query, global_connection)).Fill(SessTable);
            foreach (DataRow r in SessTable.Rows)
            {
                SesslistBox.Items.Add(r[1].ToString());
            }
            if (SesslistBox.Items.Count > 0)
                SesslistBox.SelectedIndex = 0;
            
            //SesslistBox.DataSource = SessTable;
            //SesslistBox.DisplayMember = "name";
        }

        /// <summary>
        /// заполнить список предметов выбранной сессии
        /// </summary>
        public void FillSessPredmetList()
        {
            SessPredmetlistBox.Items.Clear();
            SessionGridView.Columns[3].Visible = false;

            int SessIndex = SesslistBox.SelectedIndex;
            int GrIndex = SessGruplistBox.SelectedIndex;

            string SelGrId = SessGrupTable.Rows[GrIndex][0].ToString();
            string kod = SessTable.Rows[SessIndex][2].ToString();

            if (kod == "зс" || kod == "лс")
            {
                string sem = (kod == "зс") ? "1" : "0";
                global_query = "select predmet.id, predmet.name, vid_zan.name, " +  // 0 1 2
                    " vid_zan.id, vid_zan.kod, prkontr = predmet.name + ' - ' + vid_zan.name, title_ved, " + // 3 4 5 6
                    " semestr, dbo.GetPrepodFIOByID(prepod.id)  " +   //7   8
                    " from predmet " +   
                    "	join vidzan_predmet on vidzan_predmet.predmet_id = predmet.id " +
                    "	join vid_zan on vid_zan.id = vidzan_predmet.vidzan_id " +
                    "   join prepod on prepod.id = predmet.prepod_id " + 
                    " where predmet.actual = 1 and vid_zan.is_kontrol = 1 " +
                    " and predmet.grupa_id = " + SelGrId +  
                    " and predmet.semestr%2 = " + sem +
                    " order by outorder, predmet.name ";

                SessPredmetTable = new DataTable();
                (new SqlDataAdapter(global_query, global_connection)).Fill(SessPredmetTable);
                foreach (DataRow r in SessPredmetTable.Rows)
                {
                    SessPredmetlistBox.Items.Add(r[1].ToString() + " [" + r[2].ToString() + "]");
                }
                if (SessPredmetlistBox.Items.Count > 0)
                    SessPredmetlistBox.SelectedIndex = 0;

                //SessPredmetlistBox.DataSource = SessPredmetTable;
                //SessPredmetlistBox.DisplayMember = "prkontr";
            }

            if (kod == "пп")
            {                
                global_query = "select id, title_ved " + // 0 1 
                    " from vid_zan " +
                    " where kod = 'п'";
                SessPredmetTable = new DataTable();
                (new SqlDataAdapter(global_query, global_connection)).Fill(SessPredmetTable);
                SessVedTitle = SessPredmetTable.Rows[0][1].ToString();
                SessVedKod = "прпр";
                FillSessionTable(SessPredmetTable.Rows[0][0].ToString());
            }

            if (kod == "здр")
            {
                //FillSessionTable("26");
                global_query = "select id, title_ved " + // 0  1
                    " from vid_zan " +
                    " where kod = 'здр'";
                SessPredmetTable = new DataTable();
                (new SqlDataAdapter(global_query, global_connection)).Fill(SessPredmetTable);
                SessVedTitle = SessPredmetTable.Rows[0][1].ToString();
                SessVedKod = "здр";
                FillSessionTable(SessPredmetTable.Rows[0][0].ToString());                
            }

            if (kod == "гэ")
            {
                global_query = "select id, title_ved " + // 0  1
                    " from vid_zan " +
                    " where kod = 'иак'";
                SessPredmetTable = new DataTable();
                (new SqlDataAdapter(global_query, global_connection)).Fill(SessPredmetTable);
                SessVedTitle = SessPredmetTable.Rows[0][1].ToString();
                SessVedKod = "гэ";
                FillSessionTable(SessPredmetTable.Rows[0][0].ToString());
            }

        }

        /// <summary>
        /// заполнить результаты произв практики или дипломной работы
        /// </summary>
        /// <param name="vid_rab_id">идентификатор вида занятия</param>
        public void FillSessionTable(string vid_rab_id)
        {
            SessionGridView.Rows.Clear();
            SessStatGridView.Rows.Clear();
            SessSredBall.Text = string.Empty;
            SessionGridView.Columns[3].Visible = false;

            zach_dcell = new DataGridViewComboBoxCell();
            SessionGridView.Columns[2].CellTemplate = zach_dcell;

            // заполнить список допустимых отметок
            SessStatGridView.Rows.Clear();
            otmetki = new DataTable();
            global_query = "select vid_otmetka.id, nm = vid_otmetka.str_name " +
                " from vid_otmetka " +
                " join vid_zan_otmetka on vid_zan_otmetka.vid_otmetka_id = vid_otmetka.id " +
                " join vid_zan on vid_zan.id = vid_zan_otmetka.vid_zan_id " +
                " where vid_zan.id = " + vid_rab_id;
            (new SqlDataAdapter(global_query, global_connection)).Fill(otmetki);
            int i = 0;
            foreach (DataRow d in otmetki.Rows)
            {
                zach_dcell.Items.Add(d[1]);
                SessStatGridView.Rows.Add(d[1], 0);
                i++;
            }

            SessionResultTable = new DataTable();
            global_query = string.Format("exec dbo.TGetDiplomPPResult {0}, {1}",
                SessGrupTable.Rows[SessGruplistBox.SelectedIndex][0],
                vid_rab_id);
            (new SqlDataAdapter(global_query, global_connection)).Fill(SessionResultTable);
            i = 0;
            foreach (DataRow dr in SessionResultTable.Rows)
            {
                SessionGridView.Rows.Add(dr[1], dr[2], dr[5], dr[8]);
                SessionGridView.Rows[i].Cells[1].Tag = dr[0];
                SessionGridView.Rows[i].Cells[2].Tag = dr[0];
                SessionGridView.Rows[i].Cells[0].Tag = dr[7];
                i++;
            }
            
            SessionDate.Value = Convert.ToDateTime(SessionResultTable.Rows[0][6]);

            FillSessStatGridView();
            BuildSessDiagram();

            bool showball = Convert.ToBoolean(SessGrupTable.Rows[SessGruplistBox.SelectedIndex][4]);
            //SessionChart.Visible = showball;
            SessionGridView.Columns[1].Visible = showball;
            BallOut.Visible = showball;
            BallOut.Checked = showball;
            SessSredBall.Visible = showball;

            if (SessVedKod == "здр") //показать темы дипломных работ
            {
                SessionGridView.Columns[3].Visible = true;
            }
        }


        /// <summary>
        /// построение диагаммы по предмету для выбранной сессии
        /// </summary>
        public void BuildSessDiagram()
        {            
            dt = new DataTable();
            dt = SessionResultTable;

            dt.DefaultView.Sort = "ball";

            //if (SessionChart.Series.Count > 0) SessionChart.Series.RemoveAt(0);
            Series ser = new Series("Sess", ViewType.Bar);
            ser.DataSource = dt;
            ser.ArgumentDataMember = "fio";
            ser.ArgumentScaleType = ScaleType.Qualitative;
            ser.ValueScaleType = ScaleType.Numerical;
            ser.ValueDataMembers.AddRange(new string[] { "ball" });
            //SessionChart.Series.Add(ser);
        }

        /// <summary>
        /// заполнить сетку с результатами сессии
        /// </summary>       
        public void FillSessionTable()
        {
            SessionGridView.Rows.Clear();
            SessStatGridView.Rows.Clear();
            SessSredBall.Text = string.Empty;
            SessionGridView.Columns[3].Visible = false;

            int PredmIndex = SessPredmetlistBox.SelectedIndex;
            if (PredmIndex < 0) return;
            if (SessPredmetlistBox.Items.Count == 0) return;

            zach_dcell = new DataGridViewComboBoxCell();
            SessionGridView.Columns[2].CellTemplate = zach_dcell;

            string kod = SessPredmetTable.Rows[SessPredmetlistBox.SelectedIndex][4].ToString();
            if (kod == "зкр") //показать темы курсовых работ
            {
                SessionGridView.Columns[3].Visible = true;
            }

            // заполнить список допустимых отметок
            SessStatGridView.Rows.Clear();
            otmetki = new DataTable();
            
            global_query = "select vid_otmetka.id, nm = vid_otmetka.str_name " +
                " from vid_otmetka " +
                " join vid_zan_otmetka on vid_zan_otmetka.vid_otmetka_id = vid_otmetka.id " +
                " join vid_zan on vid_zan.id = vid_zan_otmetka.vid_zan_id " +
                " where vid_zan.id = " + SessPredmetTable.Rows[PredmIndex][3].ToString();
            
            
            (new SqlDataAdapter(global_query, global_connection)).Fill(otmetki);
            int i = 0;
            foreach (DataRow d in otmetki.Rows)
            {
                zach_dcell.Items.Add(d[1]);
                SessStatGridView.Rows.Add(d[1], 0);                
                i++;
            }
                 
            SessionResultTable = new DataTable();
            global_query = string.Format("exec dbo.TGetSessionResult {0}, {1}, {2}",
                SessGrupTable.Rows[SessGruplistBox.SelectedIndex][0],
                SessPredmetTable.Rows[PredmIndex][0],
                SessPredmetTable.Rows[PredmIndex][3]);            

            (new SqlDataAdapter(global_query, global_connection)).Fill(SessionResultTable);

            //MessageBox.Show(SessionResultTable.Rows.Count.ToString());

            i = 0;
            foreach (DataRow dr in SessionResultTable.Rows)
            {
                SessionGridView.Rows.Add(dr[1].ToString(), dr[2], dr[5], dr[8]);
                SessionGridView.Rows[i].Cells[1].Tag = dr[0];
                SessionGridView.Rows[i].Cells[2].Tag = dr[0];                
                SessionGridView.Rows[i].Cells[0].Tag = dr[7];
                i++;
            }

            //SessionDate.Value = Convert.ToDateTime(SessionResultTable.Rows[0][6]);
            SessVedTitle = SessPredmetTable.Rows[PredmIndex][6].ToString();
            SessVedKod = SessPredmetTable.Rows[PredmIndex][4].ToString();

            FillSessStatGridView();
            BuildSessDiagram();

            bool showball = Convert.ToBoolean(SessGrupTable.Rows[SessGruplistBox.SelectedIndex][4]);
            //SessionChart.Visible = showball;
            SessionGridView.Columns[1].Visible = showball;
            BallOut.Visible = showball;
            BallOut.Checked = showball;
            SessSredBall.Visible = showball;


            //удалить повторы из БД и из таблицы
        }

        /// <summary>
        /// пересчет статистики по оценкам сессии за предмет
        /// </summary>
        public void FillSessStatGridView()
        {
            int sum = 0;
            double srball = 0;
            for (int i=0; i < SessStatGridView.Rows.Count; i++)
            {                
                string otm = SessStatGridView.Rows[i].Cells[0].Value.ToString();
                sum = 0;
                for (int j = 0; j < SessionGridView.Rows.Count; j++)
                {
                    if (SessionGridView.Rows[j].Cells[2].Value.ToString() == otm)
                        sum++;
                }
                SessStatGridView.Rows[i].Cells[1].Value = sum;
            }

            for (int j = 0; j < SessionGridView.Rows.Count; j++)
            {
                if (SessionGridView.Rows[j].Cells[1].Value != null)
                {
                    int res = 0;
                    string b = SessionGridView.Rows[j].Cells[1].Value.ToString();
                    if (int.TryParse(b,out res))
                        srball += int.Parse(b);
                }
            }

            srball = srball / SessionGridView.Rows.Count;
            SessSredBall.Text = string.Format("Средний балл: {0:F2}", srball);
        }

        private void SessGruplistBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillSessControlList();
        }

        private void SesslistBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillSessPredmetList();
        }

        private void SessPredmetlistBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillSessionTable();
        }


        private void SessionGridView_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            if (e.ColumnIndex > 0)
                currentval = SessionGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
        }


        /// <summary>
        /// получить индекс отметки по ее названию
        /// </summary>
        /// <param name="val"></param>
        /// <returns></returns>
        public string SessOtmId(string val)
        {
            string res = "";

            foreach (DataRow dr in otmetki.Rows)
            {
                if (dr[1].ToString() == val)
                    res = dr[0].ToString();
            }
            return res;
        }

        private void SessionGridView_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 0) return;

            if (e.ColumnIndex == 2) //редактирование оценки
            {                
                string newval = SessOtmId(SessionGridView.Rows[e.RowIndex].Cells[2].Value.ToString());
                string sessid = SessionGridView.Rows[e.RowIndex].Cells[2].Tag.ToString();

                string sql = "update session set otmetka_id = @OTM where id = @SESSID";

                global_command = new SqlCommand(sql, global_connection);
                global_command.Parameters.Add("@OTM", SqlDbType.Int).Value = newval;
                global_command.Parameters.Add("@SESSID", SqlDbType.Int).Value = sessid;
                global_command.ExecuteNonQuery();

                FillSessStatGridView();
            }
            else //редактирование баллов
            {
                string newvalue = "";

                if (SessionGridView.Rows[e.RowIndex].Cells[1] != null)
                {
                    if (SessionGridView.Rows[e.RowIndex].Cells[1].Value != null)
                        newvalue = SessionGridView.Rows[e.RowIndex].Cells[1].Value.ToString();
                    else
                        newvalue = "0";
                }
                else
                {
                    newvalue = "0";
                }

                int d = 0;
                if (!int.TryParse(newvalue, out d))
                {
                    SessionGridView.Rows[e.RowIndex].Cells[1].Value = currentval;
                    return;
                }

                if (d < 0) d = (-d);

                SessionGridView.Rows[e.RowIndex].Cells[1].Value = d;

                string sessid = SessionGridView.Rows[e.RowIndex].Cells[1].Tag.ToString();

                string sql = "update session set ball = @BALL where id = @SESSID";

                global_command = new SqlCommand(sql, global_connection);
                global_command.Parameters.Add("@BALL", SqlDbType.Int).Value = newvalue;
                global_command.Parameters.Add("@SESSID", SqlDbType.Int).Value = sessid;
                global_command.ExecuteNonQuery();

                dt.Rows[e.RowIndex][2] = newvalue;

                FillSessStatGridView();
                BuildSessDiagram();
            }

            if (e.ColumnIndex == 3) //редактирование темы
            {
                string newval = SessionGridView.Rows[e.RowIndex].Cells[3].Value.ToString();
                string sessid = SessionGridView.Rows[e.RowIndex].Cells[2].Tag.ToString();

                string sql = "update session set tema = @TM where id = @SESSID";

                global_command = new SqlCommand(sql, global_connection);
                global_command.Parameters.Add("@TM", SqlDbType.VarChar).Value = newval;
                global_command.Parameters.Add("@SESSID", SqlDbType.Int).Value = sessid;
                global_command.ExecuteNonQuery();
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public void SessionVedomost()
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
            sheet.Cells["A2"].Value = SessVedTitle.ToUpper();
            sheet.Rows[1].Height = 30 * 20;
            //cr.Merged = false;

            cr = sheet.Cells.GetSubrange("a4", "i4");
            cr.Merged = true;
            cr.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            sheet.Cells["A4"].Value = "от " + SessionDate.Value.ToShortDateString();
            //cr.Merged = false;


            string frm = "", frm2 = "", kod = "", chas_string = ", Количество часов _____";

            //MessageBox.Show(SessVedKod);

            switch (SessVedKod)
            {
                case "э": frm = "Начало экзамена:    "; frm2 = "Экзам. оценка"; break;
                case "з": frm = "Начало зачёта:    "; frm2 = "Отметка о сдаче зачёта"; break;
                case "ма": frm = ""; frm2 = "Аттест. оценка"; chas_string = ""; break;
                case "прпр": frm = "Дата начала:    "; frm2 = "Оценка за произв. практику"; chas_string = ""; break;
                case "пост": frm = "Начало:    "; frm2 = "Оценка"; chas_string = ""; break;
                case "зкр": frm = "Начало защиты:    "; frm2 = "Оценка защиты"; break;
                case "здр": frm = "Начало: "; frm2 = "Оценка защиты"; break;
                case "кнр": frm = "Начало контр. работы:    "; frm2 = "Оценка контр. работы"; chas_string = ""; break;
                case "дз": frm = "Начало зачёта:    "; frm2 = "Зачётная оценка"; break;
                case "гэ": frm = "Начало экзамена:    "; frm2 = "Оценка экзамена"; break;
            }

            cr = sheet.Cells.GetSubrange("a6", "i6");
            cr.Merged = true;
            cr.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            sheet.Rows[5].Height = 17 * 20;

            if (SessVedKod != "ма" && SessVedKod != "пост")
            {
                if (!(SessVedKod == "прпр" || SessVedKod == "здр" || SessVedKod == "гэ"))
                    sheet.Cells["A6"].Value = frm + " " + SessionDate.Value.ToShortTimeString() + "            Окончание ________";
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

            if (SessPredmetlistBox.SelectedIndex > 0)
            {
                sheet.Cells["A8"].Value =
                    "Курс:    " + toRoman(SessGrupTable.Rows[SessGruplistBox.SelectedIndex][2].ToString()) + ",    " +
                    "Группа:    " + SessGrupTable.Rows[SessGruplistBox.SelectedIndex][1].ToString() + ",    " +
                    "Семестр:    " + toRoman(SessPredmetTable.Rows[SessPredmetlistBox.SelectedIndex][7].ToString()) + ",    " +
                    main.starts[0].Year.ToString() + "/" + main.ends[main.ends.Count - 1].Year.ToString() + " уч. г.";
            }
            else
            {
                int kurs = int.Parse(SessGrupTable.Rows[SessGruplistBox.SelectedIndex][2].ToString()) * 2;
                if (SessGrupTable.Rows[SessGruplistBox.SelectedIndex][2].ToString().Contains("ПРО")) kurs--;

                sheet.Cells["A8"].Value =
                    "Курс:    " + toRoman(SessGrupTable.Rows[SessGruplistBox.SelectedIndex][2].ToString()) + ",    " +
                    "Группа:    " + SessGrupTable.Rows[SessGruplistBox.SelectedIndex][1].ToString() + ",    " +
                    "Семестр:    " + toRoman(kurs.ToString()) + ",    " +
                    main.starts[0].Year.ToString() + "/" + main.ends[main.ends.Count - 1].Year.ToString() + " уч. г.";
            }


            cr = sheet.Cells.GetSubrange("a9", "i9");
            cr.Merged = true;
            cr.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            sheet.Rows[8].Height = 17 * 20;

            if (SessPredmetlistBox.SelectedIndex > 0)
                sheet.Cells["A9"].Value = "Дисциплина:    " +
                    SessPredmetTable.Rows[SessPredmetlistBox.SelectedIndex][1].ToString() +////////////
                    chas_string;


            cr = sheet.Cells.GetSubrange("a10", "i10");
            cr.Merged = true;
            cr.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            sheet.Rows[9].Height = 17 * 20;            
            if (SessPredmetlistBox.SelectedIndex > 0)
                sheet.Cells["A10"].Value = "Экзаменатор:    " +
                    SessPredmetTable.Rows[SessPredmetlistBox.SelectedIndex][8].ToString();/////////////
            else
                sheet.Cells["A10"].Value = "Экзаменатор:_______________________________________";



            cr = sheet.Cells.GetSubrange("a12", "i" + (12 + SessionGridView.Rows.Count).ToString());
            cr.Merged = true;
            cr.Style.Font.Size = 10 * 20;
            cr.SetBorders(MultipleBorders.Horizontal | MultipleBorders.Vertical, Color.Black, 
                GemBox.Spreadsheet.LineStyle.Thin);
            cr.Merged = false;

            sheet.Cells["A12"].Value = "№п/п";
            sheet.Cells["a12"].Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;

            cr = sheet.Cells.GetSubrange("b12", "c12");
            cr.Merged = true;
            cr.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            sheet.Cells["B12"].Value = "Фамилия, имя, отчество";

            sheet.Cells["d12"].Value = "№ зачётн.\nкнижки";
            sheet.Cells["d12"].Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;

            if (!BallOut.Checked) // Если выбрано не выводить баллы
            {
                cr = sheet.Cells.GetSubrange("e12", "g12");
                cr.Merged = true;
                cr.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
                sheet.Cells["f12"].Value = frm2;
            }
            else
            {
                sheet.Cells["e12"].Value = "Баллы";
                sheet.Cells["e12"].Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
                cr = sheet.Cells.GetSubrange("f12", "g12");
                cr.Merged = true;
                cr.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
                cr.Value = frm2;
            }

            cr = sheet.Cells.GetSubrange("h12", "i12");
            cr.Merged = true;
            cr.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            sheet.Cells["h12"].Value = "Подпись экзаменатора";

            int i = 0;            
            for (i = 0; i < SessionGridView.Rows.Count; i++)
            {
                sheet.Cells["A" + (i + 13).ToString()].Value = (i + 1).ToString();

                // фио
                cr = sheet.Cells.GetSubrange("b" + (i + 13).ToString(), "c" + (i + 13).ToString());
                cr.Merged = true;
                cr.Style.HorizontalAlignment = HorizontalAlignmentStyle.Left;
                sheet.Cells["B" + (i + 13).ToString()].Value = SessionGridView.Rows[i].Cells[0].Value.ToString();

                // номер зач. книж
                sheet.Cells["D" + (i + 13).ToString()].Value = SessionGridView.Rows[i].Cells[0].Tag.ToString();

                // баллы и оценки
                if (!BallOut.Checked) // если выбрано не выводить баллы то соединить f и g
                {
                    cr = sheet.Cells.GetSubrange("e" + (i + 13).ToString(), "g" + (i + 13).ToString());
                    cr.Merged = true;
                    if (OtmOut.Checked)
                        cr.Value = SessionGridView.Rows[i].Cells[2].Value.ToString();
                    else
                        cr.Value = "";
                    
                }
                else //вывести баллы и оценки
                {
                    // балл
                    sheet.Cells["e" + (i + 13).ToString()].Value = SessionGridView.Rows[i].Cells[1].Value.ToString();
                    sheet.Cells["e" + (i + 13).ToString()].Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;

                    // оценка
                    cr = null;
                    cr = sheet.Cells.GetSubrange("f" + (i + 13).ToString(), "g" + (i + 13).ToString());
                    cr.Merged = true;
                    cr.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;

                    if (OtmOut.Checked)
                    {
                        if (SessionGridView.Rows[i].Cells[2].Value.ToString().ToLower() != "нет")
                            cr.Value = SessionGridView.Rows[i].Cells[2].Value.ToString();
                        else
                        {
                            cr.Value = string.Empty;
                            sheet.Cells["e" + (i + 13).ToString()].Value = string.Empty;
                        }
                    }
                    else
                    {
                        cr.Value = string.Empty;
                        sheet.Cells["e" + (i + 13).ToString()].Value = string.Empty;
                    }
                    
                }

                cr = sheet.Cells.GetSubrange("h" + (i + 13).ToString(), "i" + (i + 13).ToString());
                cr.Merged = true;
            }

            int bottom = i, bottom2 = i;
            string tmp = string.Empty;
            int itog = 0;

            if (OtmOut.Checked) //вывести факт по оценкам
            {
                for (int k = 0; k < SessStatGridView.Rows.Count; k++)
                {
                    if (SessStatGridView.Rows[k].Cells[0].Value.ToString().ToLower() != "неявка")
                        tmp = "Не явился";

                    if (!(SessStatGridView.Rows[k].Cells[0].Value.ToString().ToLower() == "недопуск" ||
                        SessStatGridView.Rows[k].Cells[0].Value.ToString().ToLower() == "нет"))
                    {
                        string tmps = SessStatGridView.Rows[k].Cells[1].Value.ToString();
                        if (tmps != "0")
                            sheet.Cells["B" + (i + 14).ToString()].Value =
                                FirstLetterToCapital(SessStatGridView.Rows[k].Cells[0].Value.ToString()) + ": " + tmps;
                        else
                            sheet.Cells["B" + (i + 14).ToString()].Value =
                            FirstLetterToCapital(SessStatGridView.Rows[k].Cells[0].Value.ToString()) + " ____ ";

                        if (SessStatGridView.Rows[k].Cells[0].Value.ToString().ToLower() != "неявка")
                            itog += Convert.ToInt32(SessStatGridView.Rows[k].Cells[1].Value);

                        //sheet.Cells["B" + (i + 14).ToString()].SetBorders(MultipleBorders.Bottom, Color.Black,
                        //GemBox.Spreadsheet.LineStyle.Thin);
                    }
                    i++;
                }
            }
            else
            {
                for (int k = 0; k < SessStatGridView.Rows.Count; k++)
                {
                    if (SessStatGridView.Rows[k].Cells[0].Value.ToString().ToLower() != "неявка")
                        tmp = "Не явился";

                    if (!(SessStatGridView.Rows[k].Cells[0].Value.ToString().ToLower() == "недопуск" ||
                        SessStatGridView.Rows[k].Cells[0].Value.ToString().ToLower() == "нет"))
                    {
                        sheet.Cells["B" + (i + 14).ToString()].Value =
                            FirstLetterToCapital(SessStatGridView.Rows[k].Cells[0].Value.ToString()) + ": _____";
                    }
                    i++;
                }
            }

            if (OtmOut.Checked)
            {
                if (itog != 0)
                    sheet.Cells["f" + (bottom + 14).ToString()].Value =
                        "Итого сдавали: " + itog.ToString() + "_____________________";
                else
                    sheet.Cells["f" + (bottom + 14).ToString()].Value =
                        "Итого сдавали _____________________";
                
                bottom++;
            }
            else
            {
                sheet.Cells["f" + (bottom + 14).ToString()].Value =
                "Итого сдавали _____________________";
                bottom++;
            }
            
            sheet.Cells["f" + (bottom + 14).ToString()].Value =
            "Подпись секретаря деканата ____________";
            bottom++;
            
            sheet.Cells["f" + (bottom + 14).ToString()].Value =
            "Подпись экзаменатора _________________";
            bottom++;
            
            sheet.Cells["f" + (bottom + 14).ToString()].Value =
            "Подпись декана FSystemа _____________";


            if (SessPredmetlistBox.SelectedIndex > 0)
                FileName = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) +
                    "\\ " + SessVedTitle + " (" + SessPredmetlistBox.Items[SessPredmetlistBox.SelectedIndex].ToString() + ")" + ".xls";
            else
                FileName = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) +
                "\\ " + SessVedTitle + SessGruplistBox.Items[SessGruplistBox.SelectedIndex].ToString() + ".xls";


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

        private void toolStripButton34_Click(object sender, EventArgs e)
        {
            SessionVedomost();
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

        /// <summary>
        /// переводит первую букву строки в верхний регистр
        /// </summary>
        /// <param name="s">строка для изменения</param>
        /// <returns></returns>
        public string FirstLetterToCapital(string s)
        {
            return
                s.Substring(0, 1).ToUpper() + s.Substring(1);
        }

        /// <summary>
        /// установить дату проведения экзамена, зачета и так далее
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SessionDate_ValueChanged(object sender, EventArgs e)
        {
            string sql = string.Empty;
            
            foreach (DataGridViewRow dr in SessionGridView.Rows)
            {
                string sessid = dr.Cells[2].Tag.ToString();
                sql = sql + " update session set sessiondate = @DT where id = " + sessid + " ";
            }

            global_command = new SqlCommand(sql, global_connection);
            global_command.Parameters.Add("@DT", SqlDbType.DateTime).Value = SessionDate.Value;
            
            global_command.ExecuteNonQuery();
        }

        /// <summary>
        /// установить или отменить вывод столбца "баллы" в ведомости сессии
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BallOut_Click(object sender, EventArgs e)
        {
            if (BallOut.Checked)
            {
                BallOut.Text = "НЕ выводить столбец \"баллы\"";
                BallOut.Checked = !BallOut.Checked;
            }
            else
            {
                BallOut.Text = "Выводить столбец \"баллы\"";
                BallOut.Checked = !BallOut.Checked;
            }
        }

        /// <summary>
        /// установить или отменить вывод оценок в ведомости сессии
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OtmOut_Click(object sender, EventArgs e)
        {
            //
            if (OtmOut.Checked)
            {
                OtmOut.Text = "НЕ выводить оценки";
                OtmOut.Checked = !OtmOut.Checked;
            }
            else
            {
                OtmOut.Text = "Выводить оценки";
                OtmOut.Checked = !OtmOut.Checked;
            }
        }


        /// <summary>
        /// построение сводки по всем предметам за сессию
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripButton36_Click(object sender, EventArgs e)
        {

        }


        // ---------------- вкладка Итогов и статистики сессии ---------------------------

        /// <summary>
        /// шаблон ячейки для оценок экзамена
        /// </summary>
        DataGridViewComboBoxCell exam_dcell = null;

        /// <summary>
        /// список оценко на зачет
        /// </summary>
        public DataTable zach_otm;

        /// <summary>
        /// список оценок на экзамен
        /// </summary>
        public DataTable exam_otm;

        /// <summary>
        /// список оценок за контр работу
        /// </summary>
        public DataTable kont_otm;

        /// <summary>
        /// заполнить список групп
        /// </summary>
        public void ItogSessionEnter()
        {
            SessItogGrupComboBox.Items.Clear();

            global_query = "select id, name, kurs_id, potok, mrs from grupa where " +
                " fakultet_id = " + fakultet_id.ToString() +
                " and actual = 1 order by outorder";
            SessGrupTable = new DataTable();
            (new SqlDataAdapter(global_query, global_connection)).Fill(SessGrupTable);
            foreach (DataRow gr_row in SessGrupTable.Rows)
            {
                SessItogGrupComboBox.Items.Add(gr_row[1].ToString());
            }

            global_query = "select idd=vid_otmetka.id, otm = vid_otmetka.str_name from vid_zan_otmetka " + 
                " join vid_otmetka on vid_otmetka.id = vid_zan_otmetka.vid_otmetka_id " + 
                " where vid_zan_otmetka.vid_zan_id = 6 ";
            zach_otm = new DataTable();
            (new SqlDataAdapter(global_query, global_connection)).Fill(zach_otm);
            /*foreach (DataRow rz in zach_otm.Rows)
            {
                zach_dcell.Items.Add(rz[1].ToString());
            } */

            global_query = "select idd=vid_otmetka.id, otm = vid_otmetka.str_name from vid_zan_otmetka " +
                " join vid_otmetka on vid_otmetka.id = vid_zan_otmetka.vid_otmetka_id " +
                " where vid_zan_otmetka.vid_zan_id = 7 ";
            exam_otm = new DataTable();
            (new SqlDataAdapter(global_query, global_connection)).Fill(exam_otm);

            global_query = "select idd=vid_otmetka.id, otm = vid_otmetka.str_name from vid_zan_otmetka " +
                " join vid_otmetka on vid_otmetka.id = vid_zan_otmetka.vid_otmetka_id " +
                " where vid_zan_otmetka.vid_zan_id = 15 ";
            kont_otm = new DataTable();
            (new SqlDataAdapter(global_query, global_connection)).Fill(kont_otm);

            if (SessItogGrupComboBox.Items.Count > 0)
                SessItogGrupComboBox.SelectedIndex = 0;

            toolStrip13.Refresh();            
        }

        /// <summary>
        /// таблица студентов на вкладке итогов сессии
        /// </summary>
        public DataTable SessItogStudentTable = null;

        /// <summary>
        /// заполнить список студентов
        /// </summary>
        public void FillSessItogStudentList()
        {
            SessItogStudentComboBox.Items.Clear();
            global_query = 
            " select student.id, dbo.GetStudentFIOByID(student.id) from student " +
            " where student.actual = 1 and student.status_id = 1 " + 
            " and student.gr_id = " +
                SessGrupTable.Rows[SessItogGrupComboBox.SelectedIndex][0].ToString() + 
            " order by fam, im, ot ";
            SessItogStudentTable = new DataTable();
            (new SqlDataAdapter(global_query, global_connection)).Fill(SessItogStudentTable);
            foreach (DataRow gr_row in SessItogStudentTable.Rows)
            {
                SessItogStudentComboBox.Items.Add(gr_row[1].ToString());
            }

            if (SessItogStudentComboBox.Items.Count > 0)
                SessItogStudentComboBox.SelectedIndex = 0;
        }

        /// <summary>
        /// заполнить список студентов
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SessItogGrupComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (SessItogGrupComboBox.Items.Count > 0)
                FillSessItogStudentList();
        }

        /// <summary>
        /// таблица оценок студента по предметам
        /// </summary>
        public DataTable SessItogTable_Student = null;

        /// <summary>
        /// заполнение списка таблицы оценок студента
        /// </summary>
        /// <param name="kurs">номер курса (при значении kurs=="0"
        /// метод дает полный список предметов за все курсы)</param>
        /// <param name="exact">false=>получить предметы до данного курса включительно, true=>получить
        /// только предметы указанного курса</param>
        public void FillSessItog_Student(string kurs, string exact)
        {                        
            string StId =
                SessItogStudentTable.Rows[SessItogStudentComboBox.SelectedIndex][0].ToString();
                        
            global_query = string.Format("exec dbo.TGetStudentSessionResult {0}, {1}, {2}", StId, kurs, exact);
            SessItogTable_Student = new DataTable();
            (new SqlDataAdapter(global_query, global_connection)).Fill(SessItogTable_Student);

            int i = 0; // вывод строк
            foreach (DataRow dr in SessItogTable_Student.Rows)
            {
                DataGridViewRow drow = new DataGridViewRow();

                DataGridViewTextBoxCell txtbx = new DataGridViewTextBoxCell();
                txtbx.Value = dr[2].ToString(); // название предмета
                txtbx.Tag = dr[1].ToString(); // ид предмета
                txtbx.ToolTipText = "Преподаватель:" + dr[11].ToString();
                drow.Cells.Add(txtbx);
                
                txtbx = new DataGridViewTextBoxCell();
                txtbx.ValueType = typeof(int);
                txtbx.Value = dr[6].ToString(); //баллы
                txtbx.Tag = dr[5];                
                drow.Cells.Add(txtbx);
                
                DataGridViewComboBoxCell cmbx = new DataGridViewComboBoxCell();
                int vidz = Convert.ToInt32(dr[7]);  // оценка
                switch (vidz)
                {
                    case 6:
                        cmbx = FillCell(1);
                        cmbx.Tag = 1;
                        cmbx.Value = OtmNameByID(1, dr[5].ToString());
                        break;
                    case 7: case 16: case 9:
                        cmbx = FillCell(0);
                        cmbx.Tag = 0;
                        cmbx.Value = OtmNameByID(0, dr[5].ToString());
                        break;
                    case 15:
                        cmbx = FillCell(2);
                        cmbx.Tag = 2;
                        cmbx.Value = OtmNameByID(2, dr[5].ToString());
                        break;
                }                                
                drow.Cells.Add(cmbx);

                txtbx = new DataGridViewTextBoxCell();
                txtbx.Value = dr[8].ToString(); //отчетность
                txtbx.Tag = dr[7];
                drow.Cells.Add(txtbx);
                
                txtbx = new DataGridViewTextBoxCell();
                txtbx.Value = dr[3].ToString(); //курс
                drow.Cells.Add(txtbx);

                txtbx = new DataGridViewTextBoxCell();
                txtbx.Value = dr[4].ToString(); //семестр
                drow.Cells.Add(txtbx);

                txtbx = new DataGridViewTextBoxCell();
                txtbx.Value = dr[10].ToString(); //тема
                drow.Cells.Add(txtbx);

                SessItogdataGrid.Rows.Add(drow);
                SessItogdataGrid.Rows[i].Tag = dr[0];

                if (vidz == 7 || vidz == 16)
                    SessItogdataGrid.Rows[i].DefaultCellStyle.BackColor = Color.LightYellow;
                if (vidz == 6)
                    SessItogdataGrid.Rows[i].DefaultCellStyle.BackColor = Color.LightBlue;
                if (vidz == 9)
                    SessItogdataGrid.Rows[i].DefaultCellStyle.BackColor = Color.LightPink;
                if (vidz == 15)
                    SessItogdataGrid.Rows[i].DefaultCellStyle.BackColor = Color.FromArgb(240, 240, 240);

                int otmid = Convert.ToInt32(dr[5]);
                if (otmid == 2 || otmid == 6)
                {
                    SessItogdataGrid.Rows[i].Cells[0].Style.BackColor = Color.Red;
                    SessItogdataGrid.Rows[i].Cells[0].Style.ForeColor = Color.White;
                    SessItogdataGrid.Rows[i].Cells[2].Style.BackColor = Color.Red;
                    SessItogdataGrid.Rows[i].Cells[2].Style.ForeColor = Color.White;
                }

                i++;
            } // --- кон цикла

            SessItogdataGrid.Columns[1].ValueType = typeof(int);
            bool showball = Convert.ToBoolean(SessGrupTable.Rows[SessItogGrupComboBox.SelectedIndex][4]);
            SessItogdataGrid.Columns[1].Visible = showball;

            //toolStripStatusLabel3.Text = "Всего предметов (с учётом видов отчёности): " + 
            //    SessItogTable_Student.Rows.Count.ToString();

            FillItogStatGrid(true);
        }

        /// <summary>
        /// заполнить ячейку оценки
        /// </summary>
        /// <param name="ot">вид оценки: 0 - экзамен, =1 - зачет, =2 - контр. работа </param>
        /// <returns>заполненная ячейка оценки</returns>
        public DataGridViewComboBoxCell FillCell(int ot)
        {
            DataGridViewComboBoxCell c = new DataGridViewComboBoxCell();

            switch (ot)
            {
                case 0:
                    foreach (DataRow d1 in exam_otm.Rows) c.Items.Add(d1[1]);
                    break;
                case 1:
                    foreach (DataRow d2 in zach_otm.Rows) c.Items.Add(d2[1]);
                    break;
                case 2:
                    foreach (DataRow d3 in kont_otm.Rows) c.Items.Add(d3[1]);
                    break;
            }            

            return c;
        }

        /// <summary>
        /// получить имя оценки по ее ид
        /// </summary>
        /// <param name="type">если 0 - искать в таблице оценок экзамена, 1 - искать в списке оценок зачета, 2 - контр раб</param>
        /// <param name="val">ид оценки</param>
        /// <returns>строковое имя оценки</returns>
        public string OtmNameByID(int type, string val)
        {
            string otm = "нет";
            
            DataTable otmt = null;
            
            switch (type)
            {
                case 0: otmt = exam_otm; break;
                case 1: otmt = zach_otm; break;
                case 2: otmt = kont_otm; break;
            }
            
            foreach (DataRow otmrow in otmt.Rows)
            {
                if (val == otmrow[0].ToString())
                {
                    otm = otmrow[1].ToString();
                }
            }

            return otm;
        }

        private void SessItogStudentComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            SessItogdataGrid.Rows.Clear();
            //toolStripStatusLabel3.Text = "";
            string Kurs = SessGrupTable.Rows[SessItogGrupComboBox.SelectedIndex][2].ToString();
            if (SessItogStudentComboBox.Items.Count > 0)
                FillSessItog_Student(Kurs, "0");
        }

        private void SessItogdataGrid_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            //e.Cancel = true;
        }

        private void toolStripButton39_Click(object sender, EventArgs e)
        {
            toolStripButton38.Checked = false;
            toolStripButton38.BackColor = Color.Transparent;
            
            toolStripButton39.Checked = true;
            toolStripButton39.BackColor = Color.Yellow;
            
            toolStripButton40.Checked = false;
            toolStripButton40.BackColor = Color.Transparent;

            toolStripButton41.Checked = false;
            toolStripButton41.BackColor = Color.Transparent;
            
            toolStripButton42.Checked = false;
            toolStripButton42.BackColor = Color.Transparent;
            
            toolStripButton43.Checked = false;
            toolStripButton43.BackColor = Color.Transparent;
            
            toolStripButton44.Checked = false;
            toolStripButton44.BackColor = Color.Transparent;
            
            SessItogdataGrid.Rows.Clear();
            //toolStripStatusLabel3.Text = "";
            string Kurs = SessGrupTable.Rows[SessItogGrupComboBox.SelectedIndex][2].ToString();
            if (SessItogStudentComboBox.Items.Count > 0)
                FillSessItog_Student(Kurs, "0");
        }

        private void toolStripButton38_Click(object sender, EventArgs e)
        {
            toolStripButton39.Checked = false;
            toolStripButton39.BackColor = Color.Transparent;

            toolStripButton38.Checked = true;
            toolStripButton38.BackColor = Color.Yellow;

            toolStripButton40.Checked = false;
            toolStripButton40.BackColor = Color.Transparent;

            toolStripButton41.Checked = false;
            toolStripButton41.BackColor = Color.Transparent;

            toolStripButton42.Checked = false;
            toolStripButton42.BackColor = Color.Transparent;

            toolStripButton43.Checked = false;
            toolStripButton43.BackColor = Color.Transparent;

            toolStripButton44.Checked = false;
            toolStripButton44.BackColor = Color.Transparent;

            SessItogdataGrid.Rows.Clear();
            //toolStripStatusLabel3.Text = "";
            if (SessItogStudentComboBox.Items.Count > 0)
                FillSessItog_Student("0", "0");
        }

        private void toolStripButton40_Click(object sender, EventArgs e)
        {
            toolStripButton38.Checked = false;
            toolStripButton38.BackColor = Color.Transparent;

            toolStripButton40.Checked = true;
            toolStripButton40.BackColor = Color.Yellow;

            toolStripButton39.Checked = false;
            toolStripButton39.BackColor = Color.Transparent;

            toolStripButton41.Checked = false;
            toolStripButton41.BackColor = Color.Transparent;

            toolStripButton42.Checked = false;
            toolStripButton42.BackColor = Color.Transparent;

            toolStripButton43.Checked = false;
            toolStripButton43.BackColor = Color.Transparent;

            toolStripButton44.Checked = false;
            toolStripButton44.BackColor = Color.Transparent;

            SessItogdataGrid.Rows.Clear();
            //toolStripStatusLabel3.Text = "";
            if (SessItogStudentComboBox.Items.Count > 0)
                FillSessItog_Student("1","1");
        }

        private void toolStripButton41_Click(object sender, EventArgs e)
        {
            toolStripButton38.Checked = false;
            toolStripButton38.BackColor = Color.Transparent;

            toolStripButton41.Checked = true;
            toolStripButton41.BackColor = Color.Yellow;

            toolStripButton40.Checked = false;
            toolStripButton40.BackColor = Color.Transparent;

            toolStripButton39.Checked = false;
            toolStripButton39.BackColor = Color.Transparent;

            toolStripButton42.Checked = false;
            toolStripButton42.BackColor = Color.Transparent;

            toolStripButton43.Checked = false;
            toolStripButton43.BackColor = Color.Transparent;

            toolStripButton44.Checked = false;
            toolStripButton44.BackColor = Color.Transparent;

            SessItogdataGrid.Rows.Clear();
            //toolStripStatusLabel3.Text = "";
            if (SessItogStudentComboBox.Items.Count > 0)
                FillSessItog_Student("2", "1");
        }

        private void toolStripButton44_Click(object sender, EventArgs e)
        {
            toolStripButton38.Checked = false;
            toolStripButton38.BackColor = Color.Transparent;

            toolStripButton44.Checked = true;
            toolStripButton44.BackColor = Color.Yellow;

            toolStripButton40.Checked = false;
            toolStripButton40.BackColor = Color.Transparent;

            toolStripButton41.Checked = false;
            toolStripButton41.BackColor = Color.Transparent;

            toolStripButton42.Checked = false;
            toolStripButton42.BackColor = Color.Transparent;

            toolStripButton43.Checked = false;
            toolStripButton43.BackColor = Color.Transparent;

            toolStripButton39.Checked = false;
            toolStripButton39.BackColor = Color.Transparent;

            SessItogdataGrid.Rows.Clear();
            //toolStripStatusLabel3.Text = "";
            if (SessItogStudentComboBox.Items.Count > 0)
                FillSessItog_Student("3", "1");
        }

        private void toolStripButton43_Click(object sender, EventArgs e)
        {
            toolStripButton38.Checked = false;
            toolStripButton38.BackColor = Color.Transparent;

            toolStripButton43.Checked = true;
            toolStripButton43.BackColor = Color.Yellow;

            toolStripButton40.Checked = false;
            toolStripButton40.BackColor = Color.Transparent;

            toolStripButton41.Checked = false;
            toolStripButton41.BackColor = Color.Transparent;

            toolStripButton42.Checked = false;
            toolStripButton42.BackColor = Color.Transparent;

            toolStripButton39.Checked = false;
            toolStripButton39.BackColor = Color.Transparent;

            toolStripButton44.Checked = false;
            toolStripButton44.BackColor = Color.Transparent;

            SessItogdataGrid.Rows.Clear();
            //toolStripStatusLabel3.Text = "";
            if (SessItogStudentComboBox.Items.Count > 0)
                FillSessItog_Student("4", "1");
        }

        private void toolStripButton42_Click(object sender, EventArgs e)
        {
            toolStripButton38.Checked = false;
            toolStripButton38.BackColor = Color.Transparent;

            toolStripButton42.Checked = true;
            toolStripButton42.BackColor = Color.Yellow;

            toolStripButton40.Checked = false;
            toolStripButton40.BackColor = Color.Transparent;

            toolStripButton41.Checked = false;
            toolStripButton41.BackColor = Color.Transparent;

            toolStripButton39.Checked = false;
            toolStripButton39.BackColor = Color.Transparent;

            toolStripButton43.Checked = false;
            toolStripButton43.BackColor = Color.Transparent;

            toolStripButton44.Checked = false;
            toolStripButton44.BackColor = Color.Transparent;

            SessItogdataGrid.Rows.Clear();
            //toolStripStatusLabel3.Text = "";
            if (SessItogStudentComboBox.Items.Count > 0)
                FillSessItog_Student("5", "1");
        }

        /// <summary>
        /// запомнить текущее значение перед редактированием
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SessItogdataGrid_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            currentval = SessItogdataGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
        }

        /// <summary>
        /// получить ид оценки по ее названию
        /// </summary>
        /// <param name="val">название оценки</param>
        /// <param name="otmvid">1=получить оценку зачета, 0=получить оценку экзамена, 2-контр раб</param>
        /// <returns></returns>
        public string SessItogOtmID(string val, int otmvid)
        {
            string res = "";

            DataTable restbl = null;

            switch (otmvid)
            {
                case 0: restbl = exam_otm; break;
                case 1: restbl = zach_otm; break;
                case 2: restbl = kont_otm; break;
            }


            foreach (DataRow dr in restbl.Rows)
            {
                if (dr[1].ToString() == val)
                    res = dr[0].ToString();
            }

            return res;
        }

        /// <summary>
        /// редактирование оценок и баллов 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SessItogdataGrid_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 0) return;

            if (e.ColumnIndex == 2) //редактирование оценки
            {
                int vidzan = Convert.ToInt32(SessItogdataGrid.Rows[e.RowIndex].Cells[2].Tag);
                string newval = SessItogOtmID(SessItogdataGrid.Rows[e.RowIndex].Cells[2].Value.ToString(), vidzan);
                string sessid = SessItogdataGrid.Rows[e.RowIndex].Tag.ToString();

                string sql = "update session set otmetka_id = @OTM where id = @SESSID";

                global_command = new SqlCommand(sql, global_connection);
                global_command.Parameters.Add("@OTM", SqlDbType.Int).Value = newval;
                global_command.Parameters.Add("@SESSID", SqlDbType.Int).Value = sessid;
                global_command.ExecuteNonQuery();

                SessItogdataGrid.Rows[e.RowIndex].Cells[1].Tag = newval;
                
                if (newval == "2" || newval == "6")
                {
                    SessItogdataGrid.Rows[e.RowIndex].Cells[0].Style.BackColor = Color.Red;
                    SessItogdataGrid.Rows[e.RowIndex].Cells[0].Style.ForeColor = Color.White;
                    SessItogdataGrid.Rows[e.RowIndex].Cells[2].Style.BackColor = Color.Red;
                    SessItogdataGrid.Rows[e.RowIndex].Cells[2].Style.ForeColor = Color.White;
                }
                else
                {
                    SessItogdataGrid.Rows[e.RowIndex].Cells[0].Style.BackColor =
                        SessItogdataGrid.Rows[e.RowIndex].Cells[1].Style.BackColor;
                    SessItogdataGrid.Rows[e.RowIndex].Cells[0].Style.ForeColor = Color.Black;
                    SessItogdataGrid.Rows[e.RowIndex].Cells[2].Style.BackColor =
                        SessItogdataGrid.Rows[e.RowIndex].Cells[1].Style.BackColor;
                    SessItogdataGrid.Rows[e.RowIndex].Cells[2].Style.ForeColor = Color.Black;
                }

                FillItogStatGrid(false);
            }
            else //редактирование баллов
            {
                string newvalue = "";

                if (SessItogdataGrid.Rows[e.RowIndex].Cells[1] != null)
                {
                    if (SessItogdataGrid.Rows[e.RowIndex].Cells[1].Value != null)
                        newvalue = SessItogdataGrid.Rows[e.RowIndex].Cells[1].Value.ToString();
                    else
                        newvalue = "0";
                }
                else
                {
                    newvalue = "0";
                }

                int d = 0;
                if (!int.TryParse(newvalue, out d))
                {
                    SessItogdataGrid.Rows[e.RowIndex].Cells[1].Value = currentval;
                    SessItogdataGrid.Rows[e.RowIndex].Cells[1].ValueType = typeof(int);
                    return;
                }

                if (d < 0) d = (-d);

                SessItogdataGrid.Rows[e.RowIndex].Cells[1].Value = d;
                SessItogdataGrid.Rows[e.RowIndex].Cells[1].ValueType = typeof(int);

                string sessid = SessItogdataGrid.Rows[e.RowIndex].Tag.ToString();

                string sql = "update session set ball = @BALL where id = @SESSID";

                global_command = new SqlCommand(sql, global_connection);
                global_command.Parameters.Add("@BALL", SqlDbType.Int).Value = newvalue;
                global_command.Parameters.Add("@SESSID", SqlDbType.Int).Value = sessid;
                global_command.ExecuteNonQuery();
                FillItogStatGrid(true);
            }            
        }

        //оставить в таблице только строки с задолженностями
        private void toolStripButton45_Click(object sender, EventArgs e)
        {
            bool show = true; // по умолчанию показывать всё
            if (toolStripButton45.Text == "Показать задолженности")
            {
                toolStripButton45.Text = "Показать все оценки";
                show = false;
            }
            else
            {
                toolStripButton45.Text = "Показать задолженности";
                show = true;
            }

            if (show) //показать всё
            {
                foreach (DataGridViewRow dr in SessItogdataGrid.Rows)
                {
                    dr.Visible = true;
                }
            }
            else  // показать только задолженности
            {
                foreach (DataGridViewRow dr in SessItogdataGrid.Rows)
                {
                    int sc = Convert.ToInt32(dr.Cells[1].Tag);
                    if (sc == 2 || sc == 6 || sc == 10 || sc == 11)
                        dr.Visible = true;
                    else
                        dr.Visible = false;
                }
            }
        }
         
        /// <summary>
        /// заполнить таблицу статистики по оценкам
        /// </summary>
        /// <param name="rating">обновлять ли рейтинг (true - обновлять)</param>
        public void FillItogStatGrid(bool rating)
        {
            if (SessItogStatGrid.Rows.Count == 0)
                SessItogStatGrid.Rows.Add(0, 0, 0, 0, 0, 0, 0, 0, 0);

            int ball = 0;
            int[] itog = new int[9] { 0, 0, 0, 0, 0, 0, 0, 0, 0 };

            foreach (DataGridViewRow dr in SessItogdataGrid.Rows)
            {
                int curball = 0;
                int vidz = Convert.ToInt32(dr.Cells[3].Tag);

                if (vidz == 6 || vidz == 7 || vidz == 10 || vidz == 12 || vidz == 16)
                    curball = Convert.ToInt32(dr.Cells[1].Value);
                
                int CurOtm = Convert.ToInt32(dr.Cells[1].Tag);

                ball += curball;

                switch (CurOtm)
                {
                    case 7: itog[0]++; break; //зачт - 0
                    case 6: itog[1]++; break; //не зачт - 1
                    case 2: itog[2]++; break; //2 - 2
                    case 3: itog[3]++; break; //3 - 3
                    case 4: itog[4]++; break; //4 - 4
                    case 5: itog[5]++; break; //5 - 5
                    case 10: itog[6]++; break; //неяв - 6
                    case 11: itog[7]++; break; // недоп - 7
                    case 13: itog[8]++; break; // нет - 8
                }
            }

            SessItogStatGrid.Rows[0].Cells[0].Value = ball;
            SessItogStatGrid.Rows[0].Cells[1].Value = 0;

            for (int i = 0; i < itog.Length; i++)
                SessItogStatGrid.Rows[0].Cells[i + 2].Value = itog[i];

            if (itog[1] > 0)
            {
                SessItogStatGrid.Rows[0].Cells[3].Style.BackColor = Color.Red;
                SessItogStatGrid.Rows[0].Cells[3].Style.ForeColor = Color.White;
            }
            else
            {
                SessItogStatGrid.Rows[0].Cells[3].Style.BackColor = Color.White;
                SessItogStatGrid.Rows[0].Cells[3].Style.ForeColor = Color.Black;
            }

            if (itog[2] > 0)
            {
                SessItogStatGrid.Rows[0].Cells[4].Style.BackColor = Color.Red;
                SessItogStatGrid.Rows[0].Cells[4].Style.ForeColor = Color.White;
            }
            else
            {
                SessItogStatGrid.Rows[0].Cells[4].Style.BackColor = Color.White;
                SessItogStatGrid.Rows[0].Cells[4].Style.ForeColor = Color.Black;
            }

            if (itog[8] > 0)
            {
                SessItogStatGrid.Rows[0].Cells[10].Style.BackColor = Color.Red;
                SessItogStatGrid.Rows[0].Cells[10].Style.ForeColor = Color.White;
            }
            else
            {
                SessItogStatGrid.Rows[0].Cells[10].Style.BackColor = Color.White;
                SessItogStatGrid.Rows[0].Cells[10].Style.ForeColor = Color.Black;
            }

            // получить рейтинг
            if (rating)
            {
                global_query = string.Format("select Place from dbo.TGetGrupRating({0}) where StID = {1}",
                        SessGrupTable.Rows[SessItogGrupComboBox.SelectedIndex][0],
                        SessItogStudentTable.Rows[SessItogStudentComboBox.SelectedIndex][0]);
                DataTable Rating = new DataTable();
                (new SqlDataAdapter(global_query, global_connection)).Fill(Rating);
                int res = Convert.ToInt32(Rating.Rows[0][0]);
                SessItogStatGrid.Rows[0].Cells[1].Value = res;
            }

            if (Convert.ToInt32(SessGrupTable.Rows[SessItogGrupComboBox.SelectedIndex][4]) == 0)
                SessItogStatGrid.Columns[1].Visible = false;
            else
                SessItogStatGrid.Columns[1].Visible = true;
        }

        /// <summary>
        /// сохранить отчет по успевамости студента
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripButton46_Click(object sender, EventArgs e)
        {
            string FileName = "";
            ExcelFile excel = new ExcelFile();

            ExcelWorksheet sheet = excel.Worksheets.Add("Оценки");

            sheet.Cells[0, 1].Value = "Сводка успеваемости. ФИО студента: " + 
                SessItogStudentComboBox.Items[SessItogStudentComboBox.SelectedIndex].ToString();
            
            int gridrow = 0, exrow = 2;
            for (gridrow = 0, exrow = 2; gridrow < SessItogdataGrid.Rows.Count; gridrow++, exrow++)
            {
                sheet.Cells[exrow, 1].Value = SessItogdataGrid.Rows[gridrow].Cells[0].Value; //предм
                sheet.Cells[exrow, 1].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);

                if (SessItogdataGrid.Columns[1].Visible)
                    sheet.Cells[exrow, 2].Value = SessItogdataGrid.Rows[gridrow].Cells[1].Value; // балл
                else
                    sheet.Cells[exrow, 2].Value = "-"; // балл
                sheet.Cells[exrow, 2].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);

                int otmid = int.Parse(SessItogdataGrid.Rows[gridrow].Cells[1].Tag.ToString());

                sheet.Cells[exrow, 3].Value = SessItogdataGrid.Rows[gridrow].Cells[2].Value.ToString() +
                    " (" + SessItogdataGrid.Rows[gridrow].Cells[3].Value.ToString() + ")";

                if (otmid == 6 || otmid == 2)
                    sheet.Cells[exrow, 3].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thick);
                else
                    sheet.Cells[exrow, 3].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);

                sheet.Cells[exrow, 4].Value = SessItogdataGrid.Rows[gridrow].Cells[4].Value; // kurs
                sheet.Cells[exrow, 4].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);

                sheet.Cells[exrow, 5].Value = SessItogdataGrid.Rows[gridrow].Cells[5].Value; // semestr
                sheet.Cells[exrow, 5].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);

            }

            sheet.Cells[exrow, 1].Value = "Итог";
            sheet.Cells[exrow, 1].Style.HorizontalAlignment = HorizontalAlignmentStyle.Right;
            sheet.Cells[exrow, 1].Style.Font.Weight = ExcelFont.BoldWeight;

            sheet.Cells[exrow, 2].Value = SessItogStatGrid.Rows[0].Cells[0].Value;
            sheet.Cells[exrow, 2].Style.Font.Weight = ExcelFont.BoldWeight;

            sheet.Cells[1, 1].Value = SessItogdataGrid.Columns[0].HeaderText; //предм
            sheet.Cells[1, 1].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
            sheet.Columns[1].AutoFit();

            sheet.Cells[1, 2].Value = SessItogdataGrid.Columns[1].HeaderText; //балл
            sheet.Cells[1, 2].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
            sheet.Columns[2].AutoFit();

            sheet.Cells[1, 3].Value = SessItogdataGrid.Columns[2].HeaderText; //оц+отч
            sheet.Cells[1, 3].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
            sheet.Columns[3].AutoFit();

            sheet.Cells[1, 4].Value = SessItogdataGrid.Columns[4].HeaderText; //курс
            sheet.Cells[1, 4].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
            sheet.Columns[4].AutoFit();

            sheet.Cells[1, 5].Value = SessItogdataGrid.Columns[5].HeaderText; //семестр
            sheet.Cells[1, 5].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
            sheet.Columns[5].AutoFit();

            sheet.Columns["A"].Width = 256;
            sheet.Columns["B"].Width = 47 * 256;
            sheet.Columns["D"].Width = 39 * 256;

            sheet.PrintOptions.HeaderMargin = 0.0;
            sheet.PrintOptions.FooterMargin = 0.0;
            sheet.PrintOptions.FitToPage = true;
            sheet.PrintOptions.Portrait = false;

            saveExcel.Title = "Введите имя файла для отчёта успеваемости";
            saveExcel.DefaultExt = "xls";
            saveExcel.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            saveExcel.FileName = "Успеваемость (" +
                SessItogStudentComboBox.Items[SessItogStudentComboBox.SelectedIndex].ToString() + " - "
                + SessItogGrupComboBox.Items[SessItogGrupComboBox.SelectedIndex].ToString() + ").xls";
            
            if (saveExcel.ShowDialog() != DialogResult.OK) return;

            try
            {
                excel.SaveXls(saveExcel.FileName);
            }
            catch(Exception exx)
            {
                MessageBox.Show("Невозможно сохранить отчёт под указанным именем." +
                    "Вероятно файл с таким именем уже открыт в Excel или папка для сохранения недоступна." + 
                    "\n\nПовторите операцию снова (поменяйте имя файла или закройте его в окне программы Excel).",
                    "Отказ операции", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            Process.Start(saveExcel.FileName);
        }

        private void toolStripButton47_Click(object sender, EventArgs e)
        {

            ExcelFile excel = new ExcelFile();
            ExcelWorksheet sheet = excel.Worksheets.Add("Оценки");

            sheet.Cells[0, 1].Value = "Сводка по задолженностям. ФИО студента: " +
                SessItogStudentComboBox.Items[SessItogStudentComboBox.SelectedIndex].ToString();

            int gridrow = 0, exrow = 2;
            for (gridrow = 0, exrow = 2; gridrow < SessItogdataGrid.Rows.Count; gridrow++)
            {
                int otmid = int.Parse(SessItogdataGrid.Rows[gridrow].Cells[1].Tag.ToString());

                if ((otmid == 2 || otmid == 6 || otmid == 10 || otmid == 11))
                {
                    sheet.Cells[exrow, 1].Value = SessItogdataGrid.Rows[gridrow].Cells[0].Value; //предм
                    sheet.Cells[exrow, 1].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);

                    if (SessItogdataGrid.Columns[1].Visible)
                        sheet.Cells[exrow, 2].Value = SessItogdataGrid.Rows[gridrow].Cells[1].Value; // балл
                    else
                        sheet.Cells[exrow, 2].Value = "-"; // балл
                    sheet.Cells[exrow, 2].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);

                    sheet.Cells[exrow, 3].Value = SessItogdataGrid.Rows[gridrow].Cells[2].Value.ToString() +
                        " (" + SessItogdataGrid.Rows[gridrow].Cells[3].Value.ToString() + ")";
                    sheet.Cells[exrow, 3].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thick);

                    sheet.Cells[exrow, 4].Value = SessItogdataGrid.Rows[gridrow].Cells[4].Value; // kurs
                    sheet.Cells[exrow, 4].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);

                    sheet.Cells[exrow, 5].Value = SessItogdataGrid.Rows[gridrow].Cells[5].Value; // semestr
                    sheet.Cells[exrow, 5].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
                    exrow++;
                }
            }

            if (exrow == 2)
            {
                MessageBox.Show("Нет задолженностей!", "Отказ операции", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;                
            }

            sheet.Cells[1, 1].Value = SessItogdataGrid.Columns[0].HeaderText; //предм
            sheet.Cells[1, 1].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
            sheet.Columns[1].AutoFit();

            sheet.Cells[1, 2].Value = SessItogdataGrid.Columns[1].HeaderText; //балл
            sheet.Cells[1, 2].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
            sheet.Columns[2].AutoFit();

            sheet.Cells[1, 3].Value = SessItogdataGrid.Columns[2].HeaderText; //оц+отч
            sheet.Cells[1, 3].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
            sheet.Columns[3].AutoFit();

            sheet.Cells[1, 4].Value = SessItogdataGrid.Columns[4].HeaderText; //курс
            sheet.Cells[1, 4].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
            sheet.Columns[4].AutoFit();

            sheet.Cells[1, 5].Value = SessItogdataGrid.Columns[5].HeaderText; //семестр
            sheet.Cells[1, 5].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
            sheet.Columns[5].AutoFit();

            sheet.Columns["A"].Width = 256;
            sheet.Columns["B"].Width = 47 * 256;
            sheet.Columns["D"].Width = 39 * 256;

            sheet.PrintOptions.HeaderMargin = 0.0;
            sheet.PrintOptions.FooterMargin = 0.0;
            sheet.PrintOptions.FitToPage = true;
            sheet.PrintOptions.Portrait = false;

            saveExcel.Title = "Введите имя файла для отчёта успеваемости";
            saveExcel.DefaultExt = "xls";
            saveExcel.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            saveExcel.FileName = "Задолженности (" +
                SessItogStudentComboBox.Items[SessItogStudentComboBox.SelectedIndex].ToString() + " - "
                + SessItogGrupComboBox.Items[SessItogGrupComboBox.SelectedIndex].ToString() + ").xls";

            if (saveExcel.ShowDialog() != DialogResult.OK) return;

            try
            {
                excel.SaveXls(saveExcel.FileName);
            }
            catch (Exception exx)
            {
                MessageBox.Show("Невозможно сохранить отчёт под указанным именем." +
                    "Вероятно файл с таким именем уже открыт в Excel или папка для сохранения недоступна." +
                    "\n\nПовторите операцию снова (поменяйте имя файла или закройте его в окне программы Excel).",
                    "Отказ операции", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            Process.Start(saveExcel.FileName);
        }

        /// ----- сводка по сессии для группы ----------------------------------------------------------


        // выбор вкладки в группе итогов сессии
        private void tabControl3_Selected(object sender, TabControlEventArgs e)
        {
            if (e.TabPage == SessStatGrupaTabPage)
                EnterSessItogGrupa();
        }

        DataGridViewComboBoxCell ex_cell = new DataGridViewComboBoxCell();

        /// <summary>
        /// заполнить список групп на вкладке итогов сессии группы
        /// </summary>
        public void EnterSessItogGrupa()
        {
            SessItogGrupGrid.Rows.Clear();
            while (SessItogGrupGrid.Columns.Count > 2) SessItogGrupGrid.Columns.RemoveAt(2);

            SessItogKursCombo.Items.Clear();
            SessItogGrpaCombo.Items.Clear();

            global_query = "select id, name, kurs_id, potok, mrs from grupa where " +
                " fakultet_id = " + fakultet_id.ToString() +
                " and actual = 1 order by outorder";
            SessGrupTable = new DataTable();
            (new SqlDataAdapter(global_query, global_connection)).Fill(SessGrupTable);
            foreach (DataRow gr_row in SessGrupTable.Rows)
            {
                SessItogGrpaCombo.Items.Add(gr_row[1].ToString());
            }
            
            global_query = "select idd=vid_otmetka.id, otm = vid_otmetka.str_name from vid_zan_otmetka " +
                " join vid_otmetka on vid_otmetka.id = vid_zan_otmetka.vid_otmetka_id " +
                " where vid_zan_otmetka.vid_zan_id = 6";
            zach_otm = new DataTable();
            zach_dcell = new DataGridViewComboBoxCell();
            (new SqlDataAdapter(global_query, global_connection)).Fill(zach_otm);
            foreach (DataRow d1 in zach_otm.Rows)
            {
                zach_dcell.Items.Add(d1[1].ToString());
            }

            global_query = "select idd=vid_otmetka.id, otm = vid_otmetka.str_name from vid_zan_otmetka " +
                " join vid_otmetka on vid_otmetka.id = vid_zan_otmetka.vid_otmetka_id " +
                " where vid_zan_otmetka.vid_zan_id = 7 ";
            exam_otm = new DataTable();
            ex_cell = new DataGridViewComboBoxCell();
            (new SqlDataAdapter(global_query, global_connection)).Fill(exam_otm);
            foreach (DataRow d2 in exam_otm.Rows)
            {
                ex_cell.Items.Add(d2[1].ToString());
            }

            SessItogGrpaCombo.SelectedIndex = 0;
            toolStripButton50.Text = "Все семестры";
            toolStripButton56.Text = "Скрыть оценки";
            toolStripButton56.Enabled = true;
            toolStripButton49.Text = "Скрыть баллы";
            toolStripButton49.Enabled = true;

        }

        /// <summary>
        /// вывести список курсов 
        /// </summary>
        public void FillSessItogKursCombo()
        {
            toolStripButton50.Text = "Все семестры";
            SessItogGrupGrid.Rows.Clear();
            while (SessItogGrupGrid.Columns.Count > 2) SessItogGrupGrid.Columns.RemoveAt(2);

            int kurs = Convert.ToInt32(SessGrupTable.Rows[SessItogGrpaCombo.SelectedIndex][2]);
            SessItogKursCombo.Items.Clear();

            //SessItogKursCombo.Items.Add("Все курсы");
            string[] kurstxt = new string[5] { "1 курс", "2 курс", "3 курс", "4 курс", "5 курс" };

            for (int i = 0; i < kurs; i++)
            {
                SessItogKursCombo.Items.Add(kurstxt[i]);
            }

            SessItogKursCombo.SelectedIndex = kurs - 1;            
        }

        private void SessItogGrpaCombo_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillSessItogKursCombo();
        }


        /// <summary>
        /// таблица с итогами сессии для группы
        /// </summary>
        public DataTable SessItogGrupaTable = null;

        /// <summary>
        /// заполнить таблицу итогов сессии для группы
        /// </summary>
        public void FillSessIotgGrupaGrid()
        {
            toolStripButton50.Text = "Все семестры";
            SessItogGrupGrid.Rows.Clear();
            while (SessItogGrupGrid.Columns.Count > 2) SessItogGrupGrid.Columns.RemoveAt(2);
                        
            int kurs = SessItogKursCombo.SelectedIndex + 1;
            string grid = SessGrupTable.Rows[SessItogGrpaCombo.SelectedIndex][0].ToString();

            SessItogGrupaTable = new DataTable();
            global_query = string.Format("exec dbo.TGetGrupSessionResult {0}, {1}", grid, kurs);
            Application.DoEvents();
            (new SqlDataAdapter(global_query, global_connection)).Fill(SessItogGrupaTable);
            Application.DoEvents();

            if (SessItogGrupaTable.Rows.Count==0)
            {
                MessageBox.Show("Нет данных по итогам сессии. Введите данные на дгугих вкладках и повторите операцию.",
                    "Сбой операции", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            SessItogGrupGrid.Enabled = false;
            toolStrip14.Enabled = false;

            int rowcount = Convert.ToInt32(SessItogGrupaTable.Rows[0][0]);
            int predmcount = SessItogGrupaTable.Rows.Count / rowcount;
            int mrs = Convert.ToInt32(SessGrupTable.Rows[SessItogGrpaCombo.SelectedIndex][4]);
            if (mrs == 0)
            {
                SessItogGrupGrid.Columns[1].Visible = false;
                toolStripButton56.Visible = false;
                toolStripButton49.Visible = false;
            }
            else
            {
                SessItogGrupGrid.Columns[1].Visible = true;
                toolStripButton56.Visible = true;
                toolStripButton56.Text = "Скрыть оценки";
                toolStripButton49.Visible = true;
                toolStripButton49.Text = "Скрыть баллы";
            }

            zach_dcell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            ex_cell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

            int col = 2; //добавить колонки
            for (int j = 0; j < SessItogGrupaTable.Rows.Count; j += rowcount, col += 2)
            {
                Application.DoEvents();
                DataRow dr = SessItogGrupaTable.Rows[j];
                int vidz = Convert.ToInt32(dr[6]);                
                               
                // колонка баллов
                DataGridViewTextBoxColumn bcol = new DataGridViewTextBoxColumn();                
                bcol.HeaderText = "Баллы\n[" + dr[2].ToString() + "]";
                bcol.Tag = vidz; // ид вида занятия в столбце балла               
                if (mrs == 0) bcol.Visible = false;
                SessItogGrupGrid.Columns.Add(bcol);
                bcol.SortMode = DataGridViewColumnSortMode.NotSortable;

                //колонка оценок
                DataGridViewComboBoxColumn gcol = new DataGridViewComboBoxColumn();                
                if (vidz == 6)
                    gcol.CellTemplate = zach_dcell;
                else
                    gcol.CellTemplate = ex_cell;
                
                gcol.HeaderText = dr[2].ToString() + "\n" + 
                    dr[3].ToString() + " курс\n" + 
                    dr[4].ToString() + " cем.\n" +
                    dr[5].ToString();
                gcol.Tag = dr[4]; //ид семестра
                gcol.ToolTipText = "Преподаватель - " + dr[7].ToString();
                gcol.SortMode = DataGridViewColumnSortMode.NotSortable;
                SessItogGrupGrid.Columns.Add(gcol);                
            }

            int rownum = 0;
            for (int i = 0; i < rowcount; i++)
            {
                Application.DoEvents();
                DataRow dr = SessItogGrupaTable.Rows[i];
                object[] ob = new object[2 + predmcount * 2];               
                
                SessItogGrupGrid.Rows.Add(dr[8],0);

                rownum = i;
                int summa = 0;
                for (int k = 2; k < SessItogGrupGrid.Columns.Count; k += 2)
                {
                    DataRow dro = SessItogGrupaTable.Rows[rownum];

                    int vidz = Convert.ToInt32(dro[6]);
                    int ball = Convert.ToInt32(dro[10]);
                    int semestr = Convert.ToInt32(dro[4]);

                    summa += ball;
                    string otm = dro[11].ToString();
                    string otm1 = string.Empty;

                    SessItogGrupGrid.Rows[i].Cells[k].Value = ball;
                    SessItogGrupGrid.Rows[i].Cells[k].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

                    SessItogGrupGrid.Rows[i].Cells[k].Tag = dro[9]; //ид пункта сессии в ячейке балла

                    if (ball == 0)
                        SessItogGrupGrid.Rows[i].Cells[k].Style.BackColor = Color.Pink;

                    if (vidz == 6)
                        otm1 = OtmNameByID(1, otm);
                    else
                        otm1 = OtmNameByID(0, otm);

                    SessItogGrupGrid.Rows[i].Cells[k + 1].Value = otm1;
                    SessItogGrupGrid.Rows[i].Cells[k + 1].Tag = dro[9]; //ид пункта сессии в ячейке отметки

                    if (otm == "2" || otm == "6")
                        SessItogGrupGrid.Rows[i].Cells[k + 1].Style.ForeColor = Color.Red;
                    
                    rownum += rowcount;
                }

                SessItogGrupGrid.Rows[i].Cells[1].Value = summa;
                if (summa == 0)
                    SessItogGrupGrid.Rows[i].Cells[1].Style.ForeColor = Color.Red;
                
            }

            /// -------------------------------
            // подведение итогов  -------------                     
            DataGridViewRow dgvr = new DataGridViewRow();
            DataGridViewRow dgvr2 = new DataGridViewRow();
            DataGridViewRow dgvr2p = new DataGridViewRow();
            DataGridViewRow dgvr3 = new DataGridViewRow();
            DataGridViewRow dgvr4 = new DataGridViewRow();
            DataGridViewRow dgvr5 = new DataGridViewRow();
            DataGridViewRow dgvrnz = new DataGridViewRow();
            DataGridViewRow dgvrz = new DataGridViewRow();

            int ii = 0; // вставка строк и столбцов статистики
            for (ii = 0; ii < SessItogGrupGrid.Columns.Count; ii++)
            {
                DataGridViewTextBoxCell c = new DataGridViewTextBoxCell();
                c.Style.BackColor = Color.Red;
                c.Style.ForeColor = Color.White;
                c.Style.Font = new Font("Tahoma", 11.0f, FontStyle.Regular);
                dgvr.Cells.Add(c);
                c.ReadOnly = true;

                /*DataGridViewTextBoxCell cb = new DataGridViewTextBoxCell();
                cb.Style.BackColor = Color.Red;
                cb.Style.ForeColor = Color.White;
                cb.Style.Font = new Font("Tahoma", 11.0f, FontStyle.Regular);
                dgvr.Cells.Add(cb);
                cb.ReadOnly = true;*/


                DataGridViewTextBoxCell cc2 = new DataGridViewTextBoxCell();
                if (ii > 0) cc2.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

                DataGridViewTextBoxCell cc2p = new DataGridViewTextBoxCell();
                if (ii > 0) cc2p.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

                DataGridViewTextBoxCell cc3 = new DataGridViewTextBoxCell();
                if (ii > 0) cc3.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

                DataGridViewTextBoxCell cc4 = new DataGridViewTextBoxCell();
                if (ii > 0) cc4.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

                DataGridViewTextBoxCell cc5 = new DataGridViewTextBoxCell();
                if (ii > 0) cc5.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

                DataGridViewTextBoxCell ccnz = new DataGridViewTextBoxCell();
                if (ii > 0) ccnz.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

                DataGridViewTextBoxCell ccz = new DataGridViewTextBoxCell();
                if (ii > 0) ccz.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

                dgvr2.Cells.Add(cc2);
                dgvr2p.Cells.Add(cc2p);
                dgvr3.Cells.Add(cc3);
                dgvr4.Cells.Add(cc4);
                dgvr5.Cells.Add(cc5);
                dgvrnz.Cells.Add(ccnz);
                dgvrz.Cells.Add(ccz);

                cc2.ReadOnly = true;
                cc3.ReadOnly = true;
                cc2p.ReadOnly = true;
                cc4.ReadOnly = true;
                cc5.ReadOnly = true;
                ccnz.ReadOnly = true;
                ccz.ReadOnly = true;
            }


            dgvr.Cells[0].Value = "Средние показатели";

            dgvr2.Cells[0].Value = "Количество оц. '2'";
            dgvr2.Cells[0].Style.BackColor = Color.FromArgb(240, 240, 240);

            dgvr2p.Cells[0].Value = "Количество оц. '2+'";

            dgvr3.Cells[0].Value = "Количество оц. '3'";
            dgvr3.Cells[0].Style.BackColor = Color.FromArgb(240, 240, 240);

            dgvr4.Cells[0].Value = "Количество оц. '4'";
            
            dgvr5.Cells[0].Value = "Количество оц. '5'";
            dgvr5.Cells[0].Style.BackColor = Color.FromArgb(240, 240, 240);
            
            dgvrnz.Cells[0].Value = "Количество оц. 'затчено'";
            
            dgvrz.Cells[0].Value = "Количество оц. 'не затчено'";
            dgvrz.Cells[0].Style.BackColor = Color.FromArgb(240, 240, 240);

            SessItogGrupGrid.Rows.Add(dgvr);
            SessItogGrupGrid.Rows.Add(dgvr2);
            SessItogGrupGrid.Rows.Add(dgvr2p);
            SessItogGrupGrid.Rows.Add(dgvr3);
            SessItogGrupGrid.Rows.Add(dgvr4);
            SessItogGrupGrid.Rows.Add(dgvr5);
            SessItogGrupGrid.Rows.Add(dgvrnz);
            SessItogGrupGrid.Rows.Add(dgvrz);

            /// -------------------------------
            FillItogGrupaStatGrid();
            //FillItogGrupaStudentStatGrid(0);


            SessItogGrupGrid.Enabled = true;
            toolStrip14.Enabled = true;
        }


        public void FrmtCell(DataGridViewCell c, Color back, Color fore)
        {
            //SessItogGrupGrid.Rows[cols - 7].Cells[ii]
            c.Style.BackColor = back;
            c.Style.ForeColor = fore;
        }

        private void SessItogKursCombo_SelectedIndexChanged(object sender, EventArgs e)
        {            
            FillSessIotgGrupaGrid();            
        }

        private void SessItogGrupGrid_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            currentval = SessItogGrupGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
        }

        // редактирование сведений в таблице итогов сессии по группе
        private void SessItogGrupGrid_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 0) return;

            if (e.ColumnIndex%2==1) //редактирование оценки
            {
                int col = e.ColumnIndex;
                int vidzan = Convert.ToInt32(SessItogGrupGrid.Columns[col-1].Tag);
                
                string newval = string.Empty;
                if (vidzan == 6)
                    newval = SessItogOtmID(SessItogGrupGrid.Rows[e.RowIndex].Cells[col].Value.ToString(), 1);
                else
                    newval = SessItogOtmID(SessItogGrupGrid.Rows[e.RowIndex].Cells[col].Value.ToString(), 0);
                string sessid = SessItogGrupGrid.Rows[e.RowIndex].Cells[col].Tag.ToString();

                string sql = "update session set otmetka_id = @OTM where id = @SESSID";

                global_command = new SqlCommand(sql, global_connection);
                global_command.Parameters.Add("@OTM", SqlDbType.Int).Value = newval;
                global_command.Parameters.Add("@SESSID", SqlDbType.Int).Value = sessid;
                global_command.ExecuteNonQuery();                

                if (newval == "2" || newval == "6")
                {
                    SessItogGrupGrid.Rows[e.RowIndex].Cells[col].Style.BackColor = Color.Red;
                    SessItogGrupGrid.Rows[e.RowIndex].Cells[col].Style.ForeColor = Color.White;
                }
                else
                {
                    SessItogGrupGrid.Rows[e.RowIndex].Cells[col].Style.BackColor = Color.White;
                    SessItogGrupGrid.Rows[e.RowIndex].Cells[col].Style.ForeColor = Color.Black;
                }

                //FillItogStatGrid(false);
            }
            else //редактирование баллов
            {
                string newvalue = "";
                int col = e.ColumnIndex;
                int oldsum = Convert.ToInt32(SessItogGrupGrid.Rows[e.RowIndex].Cells[1].Value);
                int newsum = oldsum - Convert.ToInt32(currentval);

                if (SessItogGrupGrid.Rows[e.RowIndex].Cells[col] != null)
                {
                    if (SessItogGrupGrid.Rows[e.RowIndex].Cells[col].Value != null)
                        newvalue = SessItogGrupGrid.Rows[e.RowIndex].Cells[col].Value.ToString();
                    else
                        newvalue = "0";
                }
                else
                {
                    newvalue = "0";
                }

                int d = 0;
                if (!int.TryParse(newvalue, out d))
                {
                    SessItogGrupGrid.Rows[e.RowIndex].Cells[col].Value = currentval;
                    return;
                }

                if (d < 0) d = (-d);

                SessItogGrupGrid.Rows[e.RowIndex].Cells[col].Value = d;

                if (d == 0)
                    SessItogGrupGrid.Rows[e.RowIndex].Cells[col].Style.BackColor = Color.Pink;
                else
                    SessItogGrupGrid.Rows[e.RowIndex].Cells[col].Style.BackColor = Color.White;

                string sessid = SessItogGrupGrid.Rows[e.RowIndex].Cells[col].Tag.ToString();

                string sql = "update session set ball = @BALL where id = @SESSID";

                global_command = new SqlCommand(sql, global_connection);
                global_command.Parameters.Add("@BALL", SqlDbType.Int).Value = newvalue;
                global_command.Parameters.Add("@SESSID", SqlDbType.Int).Value = sessid;
                global_command.ExecuteNonQuery();

                newsum += d;
                SessItogGrupGrid.Rows[e.RowIndex].Cells[1].Value = newsum;                
            }

            FillItogGrupaStatGrid();
            FillItogGrupaStudentStatGrid(e.RowIndex);
        }

        private void SessItogGrupGrid_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            e.Cancel = true;
            /*MessageBox.Show(e.ColumnIndex + ";" + e.RowIndex + "\n" + 
                SessItogGrupGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString()+
                "\n\n"+
                SessItogGrupGrid.Columns[e.ColumnIndex].HeaderText);*/
        }

        /// <summary>
        /// фильтрация колонок баллов таблицы итогов сессии группы        
        /// </summary>
        /// <param name="visible">показывать или скрывать столбцы баллов</param>
        public void FilterSessItogGrupGrid_ball(bool visible)
        {
            SessItogGrupGrid.Columns[1].Visible = visible;

            for (int i = 2; i < SessItogGrupGrid.Columns.Count; i+=2)
            {
                SessItogGrupGrid.Columns[i].Visible = visible;
            }
        }

        /// <summary>
        ///  фильтрация колонок оценок таблицы итогов сессии группы
        /// </summary>
        /// <param name="visible">показывать или скрывать столбцы оценок</param>
        public void FilterSessItogGrupGrid_otm(bool visible)
        {
            for (int i = 3; i < SessItogGrupGrid.Columns.Count; i += 2)
            {
                SessItogGrupGrid.Columns[i].Visible = visible;
            }
        }

        private void toolStripButton56_Click(object sender, EventArgs e)
        {
            if (toolStripButton56.Text == "Скрыть оценки")
            {
                toolStripButton56.Text = "Показать оценки";
                toolStripButton49.Enabled = false;
                FilterSessItogGrupGrid_otm(false);
            }
            else
            {
                toolStripButton56.Text = "Скрыть оценки";
                toolStripButton49.Enabled = true;
                FilterSessItogGrupGrid_otm(true);
            }
        }

        private void toolStripButton49_Click(object sender, EventArgs e)
        {
            if (toolStripButton49.Text == "Скрыть баллы")
            {
                toolStripButton49.Text = "Показать баллы";
                toolStripButton56.Enabled = false;
                FilterSessItogGrupGrid_ball(false);
            }
            else
            {
                toolStripButton49.Text = "Скрыть баллы";
                toolStripButton56.Enabled = true;
                FilterSessItogGrupGrid_ball(true);
            }

        }

        /// <summary>
        /// фильтрация столбцов семестра
        /// </summary>
        /// <param name="semestr">1-нечетные семестры, 0-чётные семестры, -1-все семестры</param>
        public void FilterSessItogGrupGrid_semestr(int semestr)
        {
            int mrs = Convert.ToInt32(SessGrupTable.Rows[SessItogGrpaCombo.SelectedIndex][4]);
            if (mrs != 0)
            {
                if (semestr != -1)
                {
                    toolStripButton56.Text = "Скрыть оценки";
                    FilterSessItogGrupGrid_otm(true);

                    toolStripButton49.Text = "Скрыть баллы";                    
                    FilterSessItogGrupGrid_ball(true);

                    toolStripButton56.Enabled = false;
                    toolStripButton49.Enabled = false;
                }
                else
                {
                    toolStripButton56.Text = "Скрыть оценки";
                    toolStripButton56.Enabled = true;
                    FilterSessItogGrupGrid_otm(true);

                    toolStripButton49.Text = "Скрыть баллы";
                    toolStripButton49.Enabled = true;
                    FilterSessItogGrupGrid_ball(true);
                }
            }
            
            for (int i = 2; i < SessItogGrupGrid.Columns.Count; i += 2)
            {
                int seme = Convert.ToInt32(SessItogGrupGrid.Columns[i + 1].Tag) % 2;

                if (semestr==0) //нечётные скрыть
                {
                    switch (seme)
                    {
                        case 0: 
                            SessItogGrupGrid.Columns[i + 1].Visible = true;
                            if (mrs != 0)
                            SessItogGrupGrid.Columns[i].Visible = true;
                            break;
                        case 1:
                            SessItogGrupGrid.Columns[i + 1].Visible = false;
                            if (mrs != 0)
                            SessItogGrupGrid.Columns[i].Visible = false;
                            break;
                    }
                }

                if (semestr==1) //начётные показать
                {
                    switch (seme)
                    {
                        case 1:
                            SessItogGrupGrid.Columns[i + 1].Visible = true;
                            if (mrs != 0)
                            SessItogGrupGrid.Columns[i].Visible = true;
                            break;
                        case 0:
                            SessItogGrupGrid.Columns[i + 1].Visible = false;
                            if (mrs != 0)
                            SessItogGrupGrid.Columns[i].Visible = false;
                            break;
                    }
                }

                if (semestr==-1) //все показать
                {
                    SessItogGrupGrid.Columns[i + 1].Visible = true;
                    if (mrs != 0)
                    SessItogGrupGrid.Columns[i].Visible = true;
                }
            }

            FillItogGrupaStatGrid();
        }

        private void toolStripButton50_Click(object sender, EventArgs e)
        {            
            string txt = toolStripButton50.Text;
            switch (txt)
            {
                case "Все семестры":
                    toolStripButton50.Text = "Нечётный семестр";
                    FilterSessItogGrupGrid_semestr(1);
                    break;
                case "Нечётный семестр":
                    toolStripButton50.Text = "Чётный семестр";
                    FilterSessItogGrupGrid_semestr(0);
                    break;
                case "Чётный семестр":
                    toolStripButton50.Text = "Все семестры";
                    FilterSessItogGrupGrid_semestr(-1);
                    break;
            }
        }

        /// <summary>
        /// отчет по успеваемости ------------ в формате приложения Exclel ---------------
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripButton57_Click(object sender, EventArgs e)
        {
            //return;

            /*int i = 0;

            if (SessItogGrupGrid.Rows.Count == 0)
            {
                MessageBox.Show("Нет данных об аттестации.", "Отказ операции.",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            double summ = 0;
            for (i = 1; i < SessItogGrupGrid.Rows.Count; i++)
            {
                summ += Convert.ToDouble(SessItogGrupGrid.Rows[i].Cells[1].Value);
            }

            if (summ == 0)
            {
                MessageBox.Show("В таблице нет данных по баллам для построения отчёта.", "Отказ операции.",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            CellRange cr = null; //диапазон ячеек на рабочем листе книги
            string[] Letters = new string[]{
                "A","B","C","D","E","F","G","H","I","J",
                "K","L","M","N","O","P","Q","R","S","T",
                "U","V","W","X","Y","Z",
                "AA","AB","AC","AD","AE","AF",
                "AG","AH","AI","AJ","AK","AL",
                "AM","AN","AO","AP","AQ","AR",
                "AS","AT","AU","AV","AW","AX","AY","AZ",
                "BA","BB","BC","BD","BE","BF",
                "BG","BH","BI","BJ","BK","BL",
                "BM","BN","BO","BP","BQ","BR",
                "BS","BT","BU","BV","BW","BX","BY","BZ"
            };

            int mrs = Convert.ToInt32(SessGrupTable.Rows[SessItogGrpaCombo.SelectedIndex][4]);
            int kurs = SessItogKursCombo.SelectedIndex + 1;

            string sh1name = string.Empty;
            string sh2name = string.Empty;
            string name = string.Empty;
            int first_sem = 0, second_sem = 0;

            switch (kurs)
            {
                case 1: sh1name = "1 семестр"; sh2name = "2 семестр";
                    first_sem = 1; second_sem = 2;
                    break;
                case 2: sh1name = "3 семестр"; sh2name = "4 семестр";
                    first_sem = 3; second_sem = 4;
                    break;
                case 3: sh1name = "5 семестр"; sh2name = "6 семестр";
                    first_sem = 5; second_sem = 6;
                    break;
                case 4: sh1name = "7 семестр"; sh2name = "8 семестр";
                    first_sem = 7; second_sem = 8;
                    break;
                case 5: sh1name = "9 семестр"; sh2name = "10 семестр";
                    first_sem = 9; second_sem = 10;
                    break;
            }


            if (toolStripButton50.Text == "Все семестры")
            {
                name = string.Format("{0}, {1} семестры", first_sem, second_sem);
            }

            if (toolStripButton50.Text == "Нечётный семестр")
            {
                name = string.Format("{0} семестр", first_sem);
            }

            if (toolStripButton50.Text == "Чётный семестр")
            {
                name = string.Format("{0} семестр", second_sem);
            }

            ExcelFile excel = new ExcelFile();
            ExcelWorksheet sheet1 = excel.Worksheets.Add(name);            

            string attstr = "Итоги сессии";
            
            // -- запрос на сохранение

            saveExcel.Title = "Выберите или введите имя для файла отчёта";
            saveExcel.FileName = attstr.ToUpper() + " В ГРУППЕ " + SessItogGrpaCombo.Text + ".xls";
            if (saveExcel.ShowDialog() != DialogResult.OK) return;
            string Path = saveExcel.FileName;

            // --- создание страниц

            sheet1.Cells[0, 1].Value = attstr.ToUpper() + " В ГРУППЕ " + SessItogGrpaCombo.Text + 
                " (" + name + ")";
            sheet1.Cells[0, 1].Style.Font.Weight = ExcelFont.MaxWeight;

            if (mrs == 1)
            {
                sheet1.Cells[1, 0].Value = "Студент";
                sheet1.Cells[1, 0].Style.VerticalAlignment = VerticalAlignmentStyle.Center;
                sheet1.Cells[1, 0].Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
                sheet1.Cells[1, 0].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
                sheet1.Cells[1, 1].Value = "Всего\nбаллов";
                sheet1.Cells[1, 1].Style.VerticalAlignment = VerticalAlignmentStyle.Center;
                sheet1.Cells[1, 1].Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
                sheet1.Cells[1, 1].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
            }
            else
            {
                cr = sheet1.Cells.GetSubrange("A2", "B2");
                cr.Merged = true;
                cr.Value = "Студент";
                cr.Style.VerticalAlignment = VerticalAlignmentStyle.Center;
                cr.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
                cr.SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
            }

            sheet1.Columns[1].Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;

            for (i = 0; i < SessItogGrupGrid.Rows.Count - 8; i++)   /////// -6
            {
                if (mrs == 1)
                {
                    sheet1.Cells[i + 2, 0].Value = SessItogGrupGrid.Rows[i].Cells[0].Value;
                    sheet1.Cells[i + 2, 0].SetBorders(MultipleBorders.Outside, Color.Black,
                        GemBox.Spreadsheet.LineStyle.Thin);

                    sheet1.Cells[i + 2, 1].Value = SessItogGrupGrid.Rows[i].Cells[1].Value;
                    sheet1.Cells[i + 2, 1].SetBorders(MultipleBorders.Outside, Color.Black,
                        GemBox.Spreadsheet.LineStyle.Thin);

                    if (SessItogGrupGrid.Rows[i].Cells[1].Value.ToString() == "0")
                        sheet1.Cells[i + 2, 1].Style.Font.Weight = ExcelFont.BoldWeight;
                    sheet1.Rows[i + 2].Height = 15 * 20;
                }
                else
                {
                    cr = sheet1.Cells.GetSubrange("A" + (i + 3).ToString(), "B" + (i + 3).ToString());
                    cr.Merged = true;
                    cr.SetBorders(MultipleBorders.Outside, Color.Black,
                        GemBox.Spreadsheet.LineStyle.Thin);
                    cr.Value = SessItogGrupGrid.Rows[i].Cells[0].Value;
                }                
            }
            
            int x = 0;
            int ki = 0;
            for (ki = 2, i = 2; i < SessItogGrupGrid.Columns.Count; i += 2)
            {
                //MessageBox.Show(string.Format("{0} \n {1} \n {2}",i,SessItogGrupGrid.Columns[i].HeaderText,
                //    SessItogGrupGrid.Columns[i].Visible));

                if (mrs == 1)
                {
                    if (!SessItogGrupGrid.Columns[i].Visible) continue;
                }
                else
                {
                    if (!SessItogGrupGrid.Columns[i + 1].Visible) continue;
                }

                cr = sheet1.Cells.GetSubrange(Letters[ki] + 2.ToString(), Letters[ki + 1] + 2.ToString());
                cr.Merged = true;
                cr.SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
                cr.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
                cr.Style.VerticalAlignment = VerticalAlignmentStyle.Center;

                sheet1.Cells[1, ki].Value = SessItogGrupGrid.Columns[i + 1].HeaderText.ToUpper();
                sheet1.Cells[1, ki].Style.Rotation = 90;
                sheet1.Columns[ki].AutoFit();

                sheet1.Columns[ki].Width = 5 * 256;
                sheet1.Columns[ki].Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
                sheet1.Columns[ki + 1].Width = 5 * 256;
                sheet1.Columns[ki + 1].Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;

                for (int j = 0; j < SessItogGrupGrid.Rows.Count - 8; j++)  /////// -8
                {
                    string otm = SessItogGrupGrid.Rows[j].Cells[i + 1].Value.ToString();
                    string otmtxt = string.Empty;

                    switch (otm)
                    {
                        case "неудовлетворительно": otmtxt = "2"; break;
                        case "неуд+": otmtxt = "2+"; break;
                        case "удовлетворительно": otmtxt = "3"; break;
                        case "хорошо": otmtxt = "4"; break;
                        case "отлично": otmtxt = "5"; break;
                        case "зачтено": otmtxt = "зач"; break;
                        case "не зачтено": otmtxt = "нзач"; break;
                        default: otmtxt = ""; break;
                    }

                    if (mrs == 1)
                    {
                        sheet1.Cells[j + 2, ki].Value = SessItogGrupGrid.Rows[j].Cells[i].Value;
                        sheet1.Cells[j + 2, ki].SetBorders(MultipleBorders.Outside,
                            Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
                        if (otmtxt == "2" || otmtxt == "2+" || otmtxt == "нзач")
                        {
                            if (toolStripButton52.Checked)
                                sheet1.Cells[j + 2, ki + 1].Value = otmtxt;
                            else
                                sheet1.Cells[j + 2, ki + 1].Value = string.Empty;
                        }
                        else
                            sheet1.Cells[j + 2, ki + 1].Value = otmtxt;


                        sheet1.Cells[j + 2, ki + 1].SetBorders(MultipleBorders.Outside,
                            Color.Black, GemBox.Spreadsheet.LineStyle.Thin);


                        if (otmtxt == "2" || otmtxt == "2+" || otmtxt == "нзач")
                        {

                            sheet1.Cells[j + 2, ki + 1].Style.Font.Weight = ExcelFont.BoldWeight;
                            sheet1.Cells[j + 2, ki + 1].SetBorders(MultipleBorders.Outside, Color.Black,
                                GemBox.Spreadsheet.LineStyle.Thick);
                        }
                    }
                    else
                    {
                        cr = 
                            sheet1.Cells.GetSubrange(Letters[ki] + (j + 3).ToString(), 
                            Letters[ki + 1] + (j + 3).ToString());
                        cr.Merged = true;
                        //cr.SetBorders(MultipleBorders.Outside,
                        //    Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
                        cr.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;

                        if (otmtxt == "2" || otmtxt == "2+" || otmtxt == "нзач")
                        {
                            if (toolStripButton52.Checked)
                                cr.Value = otmtxt;
                            else
                                cr.Value = string.Empty;
                        }
                        else
                            cr.Value = otmtxt;


                        cr.SetBorders(MultipleBorders.Outside,
                            Color.Black, GemBox.Spreadsheet.LineStyle.Thin);


                        if (otmtxt == "2" || otmtxt == "2+" || otmtxt == "нзач")
                        {

                            cr.Style.Font.Weight = ExcelFont.BoldWeight;
                            cr.SetBorders(MultipleBorders.Outside, Color.Black,
                                GemBox.Spreadsheet.LineStyle.Thick);
                        }
                    }

                }                                 
                ki += 2;
            }

            sheet1.Rows[1].Height = 135 * 20;
            //sheet1.Rows[x].Height = 94 * 20;
            sheet1.Columns[0].Width = 18 * 256;
            sheet1.Columns[1].Width = 8 * 256;

            sheet1.PrintOptions.HeaderMargin = 0.0;
            sheet1.PrintOptions.FooterMargin = 0.0;
            sheet1.PrintOptions.Portrait = false;

            // ---- кон: первый лист

            if (mrs != 0)
            {
                // --- лист 2 - рейтинг ------------
                global_query = 
                    string.Format("select * from dbo.TGetGrupRating({0})",
                            SessGrupTable.Rows[SessItogGrpaCombo.SelectedIndex][0]);
                DataTable Rating = new DataTable();
                (new SqlDataAdapter(global_query, global_connection)).Fill(Rating);
                
                ExcelWorksheet sheet3 = excel.Worksheets.Add("Рейтинг");
                
                sheet3.Cells[1, 1].Value = "Рейтинг";
                sheet3.Cells[1, 1].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
                sheet3.Cells[1, 2].Value = "ФИО студента";
                sheet3.Cells[1, 2].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
                sheet3.Cells[1, 3].Value = "Сумма баллов";
                sheet3.Cells[1, 3].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);

                sheet3.Columns[1].Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
                sheet3.Columns[2].Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
                sheet3.Columns[3].Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;

                int jj = 2;
                foreach (DataRow r in Rating.Rows)
                {
                    sheet3.Cells[jj, 1].Value = r[2].ToString();
                    sheet3.Cells[jj, 1].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);

                    sheet3.Cells[jj, 2].Value = r[3].ToString();
                    sheet3.Cells[jj, 2].Style.HorizontalAlignment = HorizontalAlignmentStyle.Left;
                    sheet3.Cells[jj, 2].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
                    
                    sheet3.Cells[jj, 3].Value = r[1].ToString();
                    sheet3.Cells[jj, 3].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
                    jj++;
                }

                sheet3.Columns[1].AutoFit();
                sheet3.Columns[2].AutoFit();
                sheet3.Columns[3].AutoFit();

                sheet3.Cells[0, 2].Value = "Сведения о рейтинге студентов группы " + SessItogGrpaCombo.Text +
                    " за весь период обучения";
                sheet3.Cells[0, 3].Style.Font.Weight = ExcelFont.MaxWeight;

            }

            excel.SaveXls(Path);
            Process.Start(Path);*/
        }

        /// <summary>
        /// заполнить статистику по группе
        /// </summary>
        public void FillItogGrupaStatGrid()
        {
            if (SessItogGrupStatGrid.Rows.Count == 0)
                SessItogGrupStatGrid.Rows.Add("", 0, 0, 0, 0, 0, 0, 0, 0, 0);

            SessItogGrupStatGrid.Rows[0].Cells[0].Value = 
                SessItogGrpaCombo.Items[SessItogGrpaCombo.SelectedIndex].ToString();

            int ball = 0;
            int[] itog = new int[9] { 0, 0, 0, 0, 0, 0, 0, 0, 0 };

            int studball = 0;
            for (int kk = 0; kk < SessItogGrupGrid.Rows.Count - 8; kk++)
            {
                studball = 0;
                DataGridViewRow dr = SessItogGrupGrid.Rows[kk];
                for (int k = 2; k < SessItogGrupGrid.Columns.Count; k += 2)
                {
                    int curball = 0;
                    int vidz = Convert.ToInt32(SessItogGrupGrid.Columns[k].Tag);

                    if (vidz == 6 || vidz == 7 || vidz == 10 || vidz == 12 || vidz == 16)
                        curball = Convert.ToInt32(dr.Cells[k].Value);

                    if (SessItogGrupGrid.Columns[k].Visible)
                    {
                        ball += curball;
                        studball += curball;
                    }

                    int CurOtm = 0;
                    if (vidz == 6)
                        CurOtm = Convert.ToInt32(SessItogOtmID(dr.Cells[k + 1].Value.ToString(), 1));
                    else
                        CurOtm = Convert.ToInt32(SessItogOtmID(dr.Cells[k + 1].Value.ToString(), 0));

                    switch (CurOtm)
                    {
                        case 7: itog[0]++; break; //зачт - 0
                        case 6: itog[1]++; break; //не зачт - 1
                        case 2: itog[2]++; break; //2 - 2
                        case 3: itog[3]++; break; //3 - 3
                        case 4: itog[4]++; break; //4 - 4
                        case 5: itog[5]++; break; //5 - 5
                        case 10: itog[6]++; break; //неяв - 6
                        case 11: itog[7]++; break; // недоп - 7
                        case 13: itog[8]++; break; // нет - 8
                    }
                }

                dr.Cells[1].Value = studball;
            }

            double count = SessItogGrupGrid.Rows.Count;
            SessItogGrupStatGrid.Rows[0].Cells[1].Value = ball / count;
            SessItogGrupStatGrid.Rows[0].Cells[2].Value = 0;

            for (int i = 0; i < itog.Length; i++)
                SessItogGrupStatGrid.Rows[0].Cells[i + 3].Value = itog[i];

            if (itog[8] > 0)
            {
                SessItogGrupStatGrid.Rows[0].Cells[11].Style.BackColor = Color.Red;
                SessItogGrupStatGrid.Rows[0].Cells[11].Style.ForeColor = Color.White;
            }
            else
            {
                SessItogGrupStatGrid.Rows[0].Cells[11].Style.BackColor = Color.White;
                SessItogGrupStatGrid.Rows[0].Cells[11].Style.ForeColor = Color.Black;
            }


            // расчет баллов и оценок
            int cols = SessItogGrupGrid.Rows.Count;

            SessItogGrupGrid.Rows[cols - 7].Cells[1].Value = 0;
            SessItogGrupGrid.Rows[cols - 6].Cells[1].Value = 0;
            SessItogGrupGrid.Rows[cols - 5].Cells[1].Value = 0;
            SessItogGrupGrid.Rows[cols - 4].Cells[1].Value = 0;
            SessItogGrupGrid.Rows[cols - 3].Cells[1].Value = 0;
            SessItogGrupGrid.Rows[cols - 2].Cells[1].Value = 0;
            SessItogGrupGrid.Rows[cols - 1].Cells[1].Value = 0;

            double sum = 0;
            for (int kk = 0; kk < SessItogGrupGrid.Rows.Count - 8; kk++)
            {
                sum += Convert.ToInt32(SessItogGrupGrid.Rows[kk].Cells[1].Value);
            }
            SessItogGrupGrid.Rows[cols-8].Cells[1].Value =
                string.Format("{0:F2}", sum / (SessItogGrupGrid.Rows.Count - 8));

            for (int ii = 2; ii < SessItogGrupGrid.Columns.Count; ii++)
            {
                double avg = 0.0;
                int c2 = 0, c2p = 0, c3 = 0, c4 = 0, c5 = 0, cz = 0, cnz = 0;
                int priznak_otm = 0;
                string vidz = "";

                for (int r = 0; r < SessItogGrupGrid.Rows.Count - 8; r++)
                {
                    if (ii % 2 != 0) //сред. оценка
                    {
                        vidz = SessItogGrupGrid.Columns[ii - 1].Tag.ToString();
                        priznak_otm = (vidz == "6") ? 1 : 0;
                        string otm = SessItogOtmID(SessItogGrupGrid.Rows[r].Cells[ii].Value.ToString(), priznak_otm);
                        switch (otm)
                        {
                            case "2": avg += 2;
                                c2++;
                                break;
                            case "12": avg += 2;
                                c2p++;
                                break;
                            case "3": avg += 3;
                                c3++;
                                break;
                            case "4": avg += 4;
                                c4++;
                                break;
                            case "5": avg += 5;
                                c5++;
                                break;
                            case "6": avg++; // зачтено
                                cz++;
                                break;
                            case "7": // не зачтено
                                cnz++;
                                break;
                        }
                    }
                    else //сред. балл
                    {
                        double balll = Convert.ToDouble(SessItogGrupGrid.Rows[r].Cells[ii].Value);
                        avg += balll;
                    }

                    SessItogGrupGrid.Rows[cols - 8].Cells[ii].Style.Alignment =
                        DataGridViewContentAlignment.MiddleCenter;
                    if (priznak_otm != 1) // Если зачет то средний балл не считать
                        SessItogGrupGrid.Rows[cols - 8].Cells[ii].Value =
                            string.Format("{0:F2}", avg / (SessItogGrupGrid.Rows.Count - 8));
                    else
                        SessItogGrupGrid.Rows[cols - 8].Cells[ii].Value = "-";
                }

                if (ii % 2 != 0)
                {
                    SessItogGrupGrid.Rows[cols - 7].Cells[ii].Value = c2;
                    if (c2 > 0)
                    {
                        FrmtCell(SessItogGrupGrid.Rows[cols - 7].Cells[ii], Color.LightYellow, Color.Red);
                    }
                    else
                    {
                        FrmtCell(SessItogGrupGrid.Rows[cols - 7].Cells[ii], Color.FromArgb(240, 240, 240),
                            Color.Black);
                    }
                    SessItogGrupGrid.Rows[cols - 7].Cells[ii - 1].Style.BackColor = Color.FromArgb(240, 240, 240);



                    SessItogGrupGrid.Rows[cols - 6].Cells[ii].Value = c2p;
                    if (c2p > 0)
                    {
                        FrmtCell(SessItogGrupGrid.Rows[cols - 6].Cells[ii], Color.LightYellow, Color.Red);
                    }
                    else
                    {
                        FrmtCell(SessItogGrupGrid.Rows[cols - 6].Cells[ii], Color.White,
                            Color.Black);
                    }




                    SessItogGrupGrid.Rows[cols - 5].Cells[ii].Value = c3;
                    if (c3 > 0)
                    {
                        FrmtCell(SessItogGrupGrid.Rows[cols - 5].Cells[ii], Color.LightYellow, Color.Red);
                    }
                    else
                    {
                        FrmtCell(SessItogGrupGrid.Rows[cols - 5].Cells[ii], Color.FromArgb(240, 240, 240),
                            Color.Black);
                    }
                    SessItogGrupGrid.Rows[cols - 5].Cells[ii - 1].Style.BackColor = Color.FromArgb(240, 240, 240);



                    SessItogGrupGrid.Rows[cols - 4].Cells[ii].Value = c4;
                    if (c4 > 0)
                    {
                        FrmtCell(SessItogGrupGrid.Rows[cols - 4].Cells[ii], Color.LightYellow, Color.Red);
                    }
                    else
                    {
                        FrmtCell(SessItogGrupGrid.Rows[cols - 4].Cells[ii], Color.White,
                            Color.Black);
                    }


                    SessItogGrupGrid.Rows[cols - 3].Cells[ii].Value = c5;
                    if (c5 > 0)
                    {
                        FrmtCell(SessItogGrupGrid.Rows[cols - 3].Cells[ii], Color.LightYellow, Color.Red);
                    }
                    else
                    {
                        FrmtCell(SessItogGrupGrid.Rows[cols - 3].Cells[ii], Color.FromArgb(240, 240, 240),
                            Color.Black);
                    }
                    SessItogGrupGrid.Rows[cols - 3].Cells[ii - 1].Style.BackColor = Color.FromArgb(240, 240, 240);




                    SessItogGrupGrid.Rows[cols - 2].Cells[ii].Value = cnz;
                    if (cnz > 0)
                    {
                        FrmtCell(SessItogGrupGrid.Rows[cols - 2].Cells[ii], Color.LightYellow, Color.Red);
                    }
                    else
                    {
                        FrmtCell(SessItogGrupGrid.Rows[cols - 2].Cells[ii], Color.White,
                            Color.Black);
                    }



                    SessItogGrupGrid.Rows[cols - 1].Cells[ii].Value = cz;
                    if (cz > 0)
                    {
                        FrmtCell(SessItogGrupGrid.Rows[cols - 1].Cells[ii], Color.LightYellow, Color.Red);
                    }
                    else
                    {
                        FrmtCell(SessItogGrupGrid.Rows[cols - 1].Cells[ii], Color.FromArgb(240, 240, 240),
                            Color.Black);
                    }
                    SessItogGrupGrid.Rows[cols - 1].Cells[ii - 1].Style.BackColor = Color.FromArgb(240, 240, 240);


                    // --- сумма
                    SessItogGrupGrid.Rows[cols - 7].Cells[1].Value =
                        Convert.ToInt32(SessItogGrupGrid.Rows[cols - 7].Cells[1].Value) + c2;
                    SessItogGrupGrid.Rows[cols - 7].Cells[1].Style.BackColor = Color.FromArgb(240, 240, 240);


                    SessItogGrupGrid.Rows[cols - 6].Cells[1].Value =
                        Convert.ToInt32(SessItogGrupGrid.Rows[cols - 6].Cells[1].Value) + c2p;

                    SessItogGrupGrid.Rows[cols - 5].Cells[1].Value =
                        Convert.ToInt32(SessItogGrupGrid.Rows[cols - 5].Cells[1].Value) + c3;
                    SessItogGrupGrid.Rows[cols - 5].Cells[1].Style.BackColor = Color.FromArgb(240, 240, 240);

                    SessItogGrupGrid.Rows[cols - 4].Cells[1].Value =
                        Convert.ToInt32(SessItogGrupGrid.Rows[cols - 4].Cells[1].Value) + c4;

                    SessItogGrupGrid.Rows[cols - 3].Cells[1].Value =
                        Convert.ToInt32(SessItogGrupGrid.Rows[cols - 3].Cells[1].Value) + c5;
                    SessItogGrupGrid.Rows[cols - 3].Cells[1].Style.BackColor = Color.FromArgb(240, 240, 240);

                    SessItogGrupGrid.Rows[cols - 2].Cells[1].Value =
                        Convert.ToInt32(SessItogGrupGrid.Rows[cols - 2].Cells[1].Value) + cnz;

                    SessItogGrupGrid.Rows[cols - 1].Cells[1].Value =
                        Convert.ToInt32(SessItogGrupGrid.Rows[cols - 1].Cells[1].Value) + cz;
                    SessItogGrupGrid.Rows[cols - 1].Cells[1].Style.BackColor = Color.FromArgb(240, 240, 240);

                }
            }
        }


        /// <summary>
        /// заполнить статистику по студенту
        /// </summary>
        /// <param name="row">номер обновляемой строки</param>
        public void FillItogGrupaStudentStatGrid(int row)
        {
            if (SessItogStudentStatGrid.Rows.Count == 0)
                SessItogStudentStatGrid.Rows.Add("", 0, 0, 0, 0, 0, 0, 0, 0, 0);

            SessItogStudentStatGrid.Rows[0].Cells[0].Value = SessItogGrupGrid.Rows[row].Cells[0].Value;

            int ball = 0;
            int[] itog = new int[9] { 0, 0, 0, 0, 0, 0, 0, 0, 0 };

            DataGridViewRow dr = SessItogGrupGrid.Rows[row];

            int sred = 0;
            for (int k = 2; k < SessItogGrupGrid.Columns.Count; k += 2)
            {
                sred++;
                int curball = 0;
                int vidz = Convert.ToInt32(SessItogGrupGrid.Columns[k].Tag);

                if (vidz == 6 || vidz == 7 || vidz == 10 || vidz == 12 || vidz == 16)
                    curball = Convert.ToInt32(dr.Cells[k].Value);

                ball += curball;

                int CurOtm = 0;
                if (dr.Cells[k + 1].Value != null)
                {
                    if (vidz == 6)
                        CurOtm = Convert.ToInt32(SessItogOtmID(dr.Cells[k + 1].Value.ToString(), 1));
                    else
                        CurOtm = Convert.ToInt32(SessItogOtmID(dr.Cells[k + 1].Value.ToString(), 0));
                }

                switch (CurOtm)
                {
                    case 7: itog[0]++; break; //зачт - 0
                    case 6: itog[1]++; break; //не зачт - 1
                    case 2: itog[2]++; break; //2 - 2
                    case 3: itog[3]++; break; //3 - 3
                    case 4: itog[4]++; break; //4 - 4
                    case 5: itog[5]++; break; //5 - 5
                    case 10: itog[6]++; break; //неяв - 6
                    case 11: itog[7]++; break; // недоп - 7
                    case 13: itog[8]++; break; // нет - 8
                }
            }

            int rowcount = Convert.ToInt32(SessItogGrupaTable.Rows[0][0]);
            double predmcount = SessItogGrupaTable.Rows.Count / rowcount;

            SessItogStudentStatGrid.Rows[0].Cells[1].Value = ball / (double)sred;
            SessItogStudentStatGrid.Rows[0].Cells[2].Value = 0;

            for (int i = 0; i < itog.Length; i++)
                SessItogStudentStatGrid.Rows[0].Cells[i + 3].Value = itog[i];

            if (itog[8] > 0)
            {
                SessItogStudentStatGrid.Rows[0].Cells[11].Style.BackColor = Color.Red;
                SessItogStudentStatGrid.Rows[0].Cells[11].Style.ForeColor = Color.White;
            }
            else
            {
                SessItogStudentStatGrid.Rows[0].Cells[11].Style.BackColor = Color.White;
                SessItogStudentStatGrid.Rows[0].Cells[11].Style.ForeColor = Color.Black;
            }

            if (itog[1] > 0)
            {
                SessItogStudentStatGrid.Rows[0].Cells[4].Style.BackColor = Color.Red;
                SessItogStudentStatGrid.Rows[0].Cells[4].Style.ForeColor = Color.White;
            }
            else
            {
                SessItogStudentStatGrid.Rows[0].Cells[4].Style.BackColor = Color.White;
                SessItogStudentStatGrid.Rows[0].Cells[4].Style.ForeColor = Color.Black;
            }

            if (itog[2] > 0)
            {
                SessItogStudentStatGrid.Rows[0].Cells[5].Style.BackColor = Color.Red;
                SessItogStudentStatGrid.Rows[0].Cells[5].Style.ForeColor = Color.White;
            }
            else
            {
                SessItogStudentStatGrid.Rows[0].Cells[5].Style.BackColor = Color.White;
                SessItogStudentStatGrid.Rows[0].Cells[5].Style.ForeColor = Color.Black;
            }

            // расчет баллов и оценок
            int cols = SessItogGrupGrid.Rows.Count;

            SessItogGrupGrid.Rows[cols - 7].Cells[1].Value = 0;
            SessItogGrupGrid.Rows[cols - 6].Cells[1].Value = 0;
            SessItogGrupGrid.Rows[cols - 5].Cells[1].Value = 0;
            SessItogGrupGrid.Rows[cols - 4].Cells[1].Value = 0;
            SessItogGrupGrid.Rows[cols - 3].Cells[1].Value = 0;
            SessItogGrupGrid.Rows[cols - 2].Cells[1].Value = 0;
            SessItogGrupGrid.Rows[cols - 1].Cells[1].Value = 0;

            double sum = 0;
            for (int kk = 0; kk < SessItogGrupGrid.Rows.Count - 8; kk++)
            {
                sum += Convert.ToInt32(SessItogGrupGrid.Rows[kk].Cells[1].Value);
            }
            SessItogGrupGrid.Rows[cols - 8].Cells[1].Value =
                string.Format("{0:F2}", sum / (SessItogGrupGrid.Rows.Count - 8));

            for (int ii = 2; ii < SessItogGrupGrid.Columns.Count; ii++)
            {
                double avg = 0.0;
                int c2 = 0, c2p = 0, c3 = 0, c4 = 0, c5 = 0, cz = 0, cnz = 0;
                int priznak_otm = 0;
                string vidz = "";

                for (int r = 0; r < SessItogGrupGrid.Rows.Count - 8; r++)
                {
                    if (ii % 2 != 0) //сред. оценка
                    {
                        vidz = SessItogGrupGrid.Columns[ii - 1].Tag.ToString();
                        priznak_otm = (vidz == "6") ? 1 : 0;
                        string otm = SessItogOtmID(SessItogGrupGrid.Rows[r].Cells[ii].Value.ToString(), priznak_otm);
                        switch (otm)
                        {
                            case "2": avg += 2;
                                c2++;
                                break;
                            case "12": avg += 2;
                                c2p++;
                                break;
                            case "3": avg += 3;
                                c3++;
                                break;
                            case "4": avg += 4;
                                c4++;
                                break;
                            case "5": avg += 5;
                                c5++;
                                break;
                            case "6": avg++; // зачтено
                                cz++;
                                break;
                            case "7": // не зачтено
                                cnz++;
                                break;
                        }
                    }
                    else //сред. балл
                    {                        
                        double balll = Convert.ToDouble(SessItogGrupGrid.Rows[r].Cells[ii].Value);
                        if (SessItogGrupGrid.Columns[ii].Visible)
                            avg += balll;
                    }

                    SessItogGrupGrid.Rows[cols - 8].Cells[ii].Style.Alignment =
                        DataGridViewContentAlignment.MiddleCenter;
                    if (priznak_otm != 1) // Если зачет то средний балл не считать
                        SessItogGrupGrid.Rows[cols - 8].Cells[ii].Value =
                            string.Format("{0:F2}", avg / (SessItogGrupGrid.Rows.Count - 8));
                    else
                        SessItogGrupGrid.Rows[cols - 8].Cells[ii].Value = "-";
                }

                if (ii % 2 != 0)
                {
                    SessItogGrupGrid.Rows[cols - 7].Cells[ii].Value = c2;
                    if (c2 > 0)
                    {
                        FrmtCell(SessItogGrupGrid.Rows[cols - 7].Cells[ii], Color.LightYellow, Color.Red);
                    }
                    else
                    {
                        FrmtCell(SessItogGrupGrid.Rows[cols - 7].Cells[ii], Color.FromArgb(240, 240, 240),
                            Color.Black);
                    }
                    SessItogGrupGrid.Rows[cols - 7].Cells[ii - 1].Style.BackColor = Color.FromArgb(240, 240, 240);



                    SessItogGrupGrid.Rows[cols - 6].Cells[ii].Value = c2p;
                    if (c2p > 0)
                    {
                        FrmtCell(SessItogGrupGrid.Rows[cols - 6].Cells[ii], Color.LightYellow, Color.Red);
                    }
                    else
                    {
                        FrmtCell(SessItogGrupGrid.Rows[cols - 6].Cells[ii], Color.White,
                            Color.Black);
                    }




                    SessItogGrupGrid.Rows[cols - 5].Cells[ii].Value = c3;
                    if (c3 > 0)
                    {
                        FrmtCell(SessItogGrupGrid.Rows[cols - 5].Cells[ii], Color.LightYellow, Color.Red);
                    }
                    else
                    {
                        FrmtCell(SessItogGrupGrid.Rows[cols - 5].Cells[ii], Color.FromArgb(240, 240, 240),
                            Color.Black);
                    }
                    SessItogGrupGrid.Rows[cols - 5].Cells[ii - 1].Style.BackColor = Color.FromArgb(240, 240, 240);



                    SessItogGrupGrid.Rows[cols - 4].Cells[ii].Value = c4;
                    if (c4 > 0)
                    {
                        FrmtCell(SessItogGrupGrid.Rows[cols - 4].Cells[ii], Color.LightYellow, Color.Red);
                    }
                    else
                    {
                        FrmtCell(SessItogGrupGrid.Rows[cols - 4].Cells[ii], Color.White,
                            Color.Black);
                    }



                    SessItogGrupGrid.Rows[cols - 3].Cells[ii].Value = c5;
                    if (c5 > 0)
                    {
                        FrmtCell(SessItogGrupGrid.Rows[cols - 3].Cells[ii], Color.LightYellow, Color.Red);
                    }
                    else
                    {
                        FrmtCell(SessItogGrupGrid.Rows[cols - 3].Cells[ii], Color.FromArgb(240, 240, 240),
                            Color.Black);
                    }
                    SessItogGrupGrid.Rows[cols - 3].Cells[ii - 1].Style.BackColor = Color.FromArgb(240, 240, 240);




                    SessItogGrupGrid.Rows[cols - 2].Cells[ii].Value = cnz;
                    if (cnz > 0)
                    {
                        FrmtCell(SessItogGrupGrid.Rows[cols - 2].Cells[ii], Color.LightYellow, Color.Red);
                    }
                    else
                    {
                        FrmtCell(SessItogGrupGrid.Rows[cols - 2].Cells[ii], Color.White,
                            Color.Black);
                    }



                    SessItogGrupGrid.Rows[cols - 1].Cells[ii].Value = cz;
                    if (cz > 0)
                    {
                        FrmtCell(SessItogGrupGrid.Rows[cols - 1].Cells[ii], Color.LightYellow, Color.Red);
                    }
                    else
                    {
                        FrmtCell(SessItogGrupGrid.Rows[cols - 1].Cells[ii], Color.FromArgb(240, 240, 240),
                            Color.Black);
                    }
                    SessItogGrupGrid.Rows[cols - 1].Cells[ii - 1].Style.BackColor = Color.FromArgb(240, 240, 240);



                    // --- сумма
                    SessItogGrupGrid.Rows[cols - 7].Cells[1].Value =
                        Convert.ToInt32(SessItogGrupGrid.Rows[cols - 7].Cells[1].Value) + c2;
                    SessItogGrupGrid.Rows[cols - 7].Cells[1].Style.BackColor = Color.FromArgb(240, 240, 240);


                    SessItogGrupGrid.Rows[cols - 6].Cells[1].Value =
                        Convert.ToInt32(SessItogGrupGrid.Rows[cols - 6].Cells[1].Value) + c2p;

                    SessItogGrupGrid.Rows[cols - 5].Cells[1].Value =
                        Convert.ToInt32(SessItogGrupGrid.Rows[cols - 5].Cells[1].Value) + c3;
                    SessItogGrupGrid.Rows[cols - 5].Cells[1].Style.BackColor = Color.FromArgb(240, 240, 240);

                    SessItogGrupGrid.Rows[cols - 4].Cells[1].Value =
                        Convert.ToInt32(SessItogGrupGrid.Rows[cols - 4].Cells[1].Value) + c4;

                    SessItogGrupGrid.Rows[cols - 3].Cells[1].Value =
                        Convert.ToInt32(SessItogGrupGrid.Rows[cols - 3].Cells[1].Value) + c5;
                    SessItogGrupGrid.Rows[cols - 3].Cells[1].Style.BackColor = Color.FromArgb(240, 240, 240);

                    SessItogGrupGrid.Rows[cols - 2].Cells[1].Value =
                        Convert.ToInt32(SessItogGrupGrid.Rows[cols - 2].Cells[1].Value) + cnz;

                    SessItogGrupGrid.Rows[cols - 1].Cells[1].Value =
                        Convert.ToInt32(SessItogGrupGrid.Rows[cols - 1].Cells[1].Value) + cz;
                    SessItogGrupGrid.Rows[cols - 1].Cells[1].Style.BackColor = Color.FromArgb(240, 240, 240);

                }
            }
        }

        private void SessItogGrupGrid_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < SessItogGrupGrid.Rows.Count - 8)
                FillItogGrupaStudentStatGrid(e.RowIndex);
        }

        private void toolStripButton52_Click(object sender, EventArgs e)
        {
            //Показывать отр. оценки
            if (toolStripButton52.Text == "Показывать отр. оценки")
            {
                toolStripButton52.Text = "Скрывать отр. оценки";
                toolStripButton52.Checked = false;
            }
            else
            {
                toolStripButton52.Text = "Показывать отр. оценки";
                toolStripButton52.Checked = true;
            }
        }

        private void datagrid_context_Opening(object sender, System.ComponentModel.CancelEventArgs e)
        {
            CopyGridToClipBoard(SessItogStudentStatGrid);
        }

        private void toolStripMenuItem24_Click(object sender, EventArgs e)
        {
            CopyGridToClipBoard(kontr_table);
        }

        private void копироватьВБуферToolStripMenuItem3_Click(object sender, EventArgs e)
        {
            CopyGridToClipBoard(exam_table);
        }

        /// <summary>
        /// скопировать содержимое из указанной таблицы в буфер обмена
        /// </summary>
        /// <param name="tabl">таблица, содержимое которой нужно скопировать</param>
        public static void CopyGridToClipBoard(DataGridView tabl)
        {
            try
            {
                Clipboard.SetText(
                       tabl.GetClipboardContent().GetText(TextDataFormat.Text),
                       TextDataFormat.Text);
            }
            catch (Exception exx)
            {
                //MessageBox.Show(exx.Message);
            }
        }

        private void копироватьВБуферToolStripMenuItem4_Click(object sender, EventArgs e)
        {
            CopyGridToClipBoard(kurs_table);
        }

        private void копироватьВБуферToolStripMenuItem5_Click(object sender, EventArgs e)
        {
            CopyGridToClipBoard(zachet_table);
        }

        private void копироватьВБуферToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            CopyGridToClipBoard(AttTableGridView);
        }

        private void toolStripMenuItem34_Click(object sender, EventArgs e)
        {
            CopyGridToClipBoard(SessItogGrupStatGrid);
        }

        private void toolStripMenuItem35_Click(object sender, EventArgs e)
        {
            CopyGridToClipBoard(SessItogStudentStatGrid);
        }

        private void toolStripMenuItem33_Click(object sender, EventArgs e)
        {
            CopyGridToClipBoard(SessItogGrupGrid);
        }

        private void toolStripMenuItem31_Click(object sender, EventArgs e)
        {
            CopyGridToClipBoard(SessItogdataGrid);
        }

        private void toolStripMenuItem32_Click(object sender, EventArgs e)
        {
            CopyGridToClipBoard(SessItogStatGrid);
        }

        private void toolStripMenuItem29_Click(object sender, EventArgs e)
        {
            CopyGridToClipBoard(SessionGridView);
        }

        private void toolStripMenuItem30_Click(object sender, EventArgs e)
        {
            CopyGridToClipBoard(SessStatGridView);
        }

        /// <summary>
        /// вывод рейтинга во внешний файл
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripButton53_Click(object sender, EventArgs e)
        {
            int i = 0;
                        
            int mrs = Convert.ToInt32(SessGrupTable.Rows[SessItogGrpaCombo.SelectedIndex][4]);
                       
            string attstr = "Итоги сессии";

            if (mrs != 0)
            {
                // -- запрос на сохранение
                saveExcel.Title = "Выберите или введите имя для файла отчёта";
                saveExcel.FileName = "Рейтинг " + " В ГРУППЕ " + SessItogGrpaCombo.Text + ".xls";

                if (saveExcel.ShowDialog() != DialogResult.OK) return;
                string Path = saveExcel.FileName;

                // --- лист - рейтинг ------------
                global_query =
                    string.Format("select * from dbo.TGetGrupRating({0})",
                            SessGrupTable.Rows[SessItogGrpaCombo.SelectedIndex][0]);
                DataTable Rating = new DataTable();
                (new SqlDataAdapter(global_query, global_connection)).Fill(Rating);

                ExcelFile excel = new ExcelFile();
                ExcelWorksheet sheet3 = excel.Worksheets.Add("Рейтинг " + SessItogGrpaCombo.Text);

                sheet3.Cells[1, 1].Value = "Рейтинг";
                sheet3.Cells[1, 1].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
                sheet3.Cells[1, 2].Value = "ФИО студента";
                sheet3.Cells[1, 2].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
                sheet3.Cells[1, 3].Value = "Сумма баллов";
                sheet3.Cells[1, 3].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);

                sheet3.Columns[1].Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
                sheet3.Columns[2].Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
                sheet3.Columns[3].Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;

                int jj = 2;
                foreach (DataRow r in Rating.Rows)
                {
                    sheet3.Cells[jj, 1].Value = r[2].ToString();
                    sheet3.Cells[jj, 1].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);

                    sheet3.Cells[jj, 2].Value = r[3].ToString();
                    sheet3.Cells[jj, 2].Style.HorizontalAlignment = HorizontalAlignmentStyle.Left;
                    sheet3.Cells[jj, 2].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);

                    sheet3.Cells[jj, 3].Value = r[1].ToString();
                    sheet3.Cells[jj, 3].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
                    jj++;
                }

                sheet3.Columns[1].AutoFit();
                sheet3.Columns[2].AutoFit();
                sheet3.Columns[3].AutoFit();

                sheet3.Cells[0, 2].Value = "Сведения о рейтинге студентов группы " + SessItogGrpaCombo.Text +
                    " за весь период обучения";
                sheet3.Cells[0, 3].Style.Font.Weight = ExcelFont.MaxWeight;

                excel.SaveXls(Path);
                Process.Start(Path);
            }
            else
            {
                MessageBox.Show("Группа " + SessItogGrpaCombo.Text + 
                    " не обучается по балльно-рейтинговой системе. Учёт рейтинга для неё не ведётся.",
                    "Отказ операции", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            
        }

        private void toolStripMenuItem36_Click(object sender, EventArgs e)
        {
            CopyGridToClipBoard(PosechStatGrid);
        }

        private void копироватьТаблицуВБуферToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CopyGridToClipBoard(prepod_table);
        }


        // удалить в архив предмет из указанного столбца в таблице результатов сессии
        private void toolStripMenuItem37_Click(object sender, EventArgs e)
        {            
            DataGridViewCell cell = SessItogGrupGrid.CurrentCell;

            if (cell == null) return;
            if (cell.ColumnIndex <= 1) return;

            int r = cell.RowIndex;
            int c = cell.ColumnIndex;

            if (c % 2 != 0) c--;

            string sess_id = string.Empty;
            string sql = "update session set isactual=0 where ";

            int rcount = SessItogGrupGrid.Rows.Count - 8;

            for (int i = 0; i < rcount; i++)
            {
                sess_id = SessItogGrupGrid.Rows[i].Cells[c].Tag.ToString();
                if (i < rcount - 1)
                    sql += " id = " + sess_id + " or ";
                else
                    sql += " id = " + sess_id;
            }

            //MessageBox.Show(sql);

            if (MessageBox.Show("Вы собираетесь переместить в архив информацию по отчетности:\n" +
                SessItogGrupGrid.Columns[c + 1].HeaderText +
                "\n\nЭта операция является отменяемой. Позднее предмет снова можно вернуть в сетку сессии.",
                "Сообщение", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) != DialogResult.Cancel)
            {
                SqlCommand cmd = new SqlCommand(sql, global_connection);
                cmd.ExecuteNonQuery();
                FillSessIotgGrupaGrid();
            }

        }

        private void toolStripButton51_Click(object sender, EventArgs e)
        {
            FillSessIotgGrupaGrid();
        }


        /// <summary>
        /// команда для восстановления предмета в сессии
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripMenuItem38_Click(object sender, EventArgs e)
        {
            string gr_id = SessGrupTable.Rows[SessItogGrpaCombo.SelectedIndex][0].ToString();
            string gr_name = SessItogGrpaCombo.Text;
            string kurs = (SessItogKursCombo.SelectedIndex + 1).ToString();

            edit_session_predmet_restotre espr = new edit_session_predmet_restotre(gr_id, gr_name, kurs, "");
            
            DialogResult d = espr.ShowDialog();

            if (d == DialogResult.OK)
            {
                FillSessIotgGrupaGrid();
            }
        }

        /// <summary>
        /// удалить предмет из перчня МСА
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripMenuItem39_Click(object sender, EventArgs e)
        {
            DataGridViewCell cell = AttTableGridView.CurrentCell;

            if (cell == null) return;
            if (cell.ColumnIndex <= 1) return;

            int r = cell.RowIndex;
            int c = cell.ColumnIndex;

            if (c % 2 != 0) c--;

            string sess_id = string.Empty;
            string sql = "update session set isactual=0 where ";

            int rcount = AttTableGridView.Rows.Count - 6;

            for (int i = 0; i < rcount; i++)
            {
                sess_id = AttTableGridView.Rows[i].Cells[c].Tag.ToString();
                if (i < rcount - 1)
                    sql += " id = " + sess_id + " or ";
                else
                    sql += " id = " + sess_id;
            }

            //MessageBox.Show(sql);

            if (MessageBox.Show("Вы собираетесь переместить в архив информацию по отчетности:\n" +
                AttTableGridView.Columns[c].HeaderText +
                "\n\nЭта операция является отменяемой. Позднее предмет снова можно вернуть в сетку аттестации.",
                "Сообщение", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) != DialogResult.Cancel)
            {
                SqlCommand cmd = new SqlCommand(sql, global_connection);
                cmd.ExecuteNonQuery();
                FillAttTable();
            }
        }

        /// <summary>
        /// вернуть предмет в перечень МСА
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripMenuItem40_Click(object sender, EventArgs e)
        {
            string gr_id = AttGrupList.Rows[grupAttComboBox.SelectedIndex][0].ToString();
            string gr_name = grupAttComboBox.Text;
            string kurs = AttGrupList.Rows[grupAttComboBox.SelectedIndex][2].ToString();

            edit_session_predmet_restotre espr = 
                new edit_session_predmet_restotre(gr_id, gr_name, kurs,
                    AttListTable.Rows[attListComboBox.SelectedIndex][0].ToString());

            DialogResult d = espr.ShowDialog();

            if (d == DialogResult.OK)
            {
                FillAttTable();
            }
        }

        private void текущаяГруппаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //return;

            int i = 0;

            if (SessItogGrupGrid.Rows.Count == 0)
            {
                MessageBox.Show("Нет данных об аттестации.", "Отказ операции.",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            double summ = 0;
            for (i = 1; i < SessItogGrupGrid.Rows.Count; i++)
            {
                summ += Convert.ToDouble(SessItogGrupGrid.Rows[i].Cells[1].Value);
            }

            if (summ == 0)
            {
                MessageBox.Show("В таблице нет данных по баллам для построения отчёта.", "Отказ операции.",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            CellRange cr = null; //диапазон ячеек на рабочем листе книги
            string[] Letters = new string[]{
                "A","B","C","D","E","F","G","H","I","J",
                "K","L","M","N","O","P","Q","R","S","T",
                "U","V","W","X","Y","Z",
                "AA","AB","AC","AD","AE","AF",
                "AG","AH","AI","AJ","AK","AL",
                "AM","AN","AO","AP","AQ","AR",
                "AS","AT","AU","AV","AW","AX","AY","AZ",
                "BA","BB","BC","BD","BE","BF",
                "BG","BH","BI","BJ","BK","BL",
                "BM","BN","BO","BP","BQ","BR",
                "BS","BT","BU","BV","BW","BX","BY","BZ"
            };

            int mrs = Convert.ToInt32(SessGrupTable.Rows[SessItogGrpaCombo.SelectedIndex][4]);
            int kurs = SessItogKursCombo.SelectedIndex + 1;

            string sh1name = string.Empty;
            string sh2name = string.Empty;
            string name = string.Empty;
            int first_sem = 0, second_sem = 0;

            switch (kurs)
            {
                case 1: sh1name = "1 семестр"; sh2name = "2 семестр";
                    first_sem = 1; second_sem = 2;
                    break;
                case 2: sh1name = "3 семестр"; sh2name = "4 семестр";
                    first_sem = 3; second_sem = 4;
                    break;
                case 3: sh1name = "5 семестр"; sh2name = "6 семестр";
                    first_sem = 5; second_sem = 6;
                    break;
                case 4: sh1name = "7 семестр"; sh2name = "8 семестр";
                    first_sem = 7; second_sem = 8;
                    break;
                case 5: sh1name = "9 семестр"; sh2name = "10 семестр";
                    first_sem = 9; second_sem = 10;
                    break;
            }


            if (toolStripButton50.Text == "Все семестры")
            {
                name = string.Format("{0}, {1} семестры", first_sem, second_sem);
            }

            if (toolStripButton50.Text == "Нечётный семестр")
            {
                name = string.Format("{0} семестр", first_sem);
            }

            if (toolStripButton50.Text == "Чётный семестр")
            {
                name = string.Format("{0} семестр", second_sem);
            }

            ExcelFile excel = new ExcelFile();
            ExcelWorksheet sheet1 = excel.Worksheets.Add(name);

            string attstr = "Итоги сессии";

            // -- запрос на сохранение

            saveExcel.Title = "Выберите или введите имя для файла отчёта";
            saveExcel.FileName = attstr.ToUpper() + " В ГРУППЕ " + SessItogGrpaCombo.Text + ".xls";
            if (saveExcel.ShowDialog() != DialogResult.OK) return;
            string Path = saveExcel.FileName;

            // --- создание страниц

            sheet1.Cells[0, 1].Value = attstr.ToUpper() + " В ГРУППЕ " + SessItogGrpaCombo.Text +
                " (" + name + ")";
            sheet1.Cells[0, 1].Style.Font.Weight = ExcelFont.MaxWeight;

            if (mrs == 1)
            {
                sheet1.Cells[1, 0].Value = "Студент";
                sheet1.Cells[1, 0].Style.VerticalAlignment = VerticalAlignmentStyle.Center;
                sheet1.Cells[1, 0].Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
                sheet1.Cells[1, 0].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
                sheet1.Cells[1, 1].Value = "Всего\nбаллов";
                sheet1.Cells[1, 1].Style.VerticalAlignment = VerticalAlignmentStyle.Center;
                sheet1.Cells[1, 1].Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
                sheet1.Cells[1, 1].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
            }
            else
            {
                cr = sheet1.Cells.GetSubrange("A2", "B2");
                cr.Merged = true;
                cr.Value = "Студент";
                cr.Style.VerticalAlignment = VerticalAlignmentStyle.Center;
                cr.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
                cr.SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
            }

            sheet1.Columns[1].Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;

            for (i = 0; i < SessItogGrupGrid.Rows.Count - 8; i++)   /////// -6
            {
                if (mrs == 1)
                {
                    sheet1.Cells[i + 2, 0].Value = SessItogGrupGrid.Rows[i].Cells[0].Value;
                    sheet1.Cells[i + 2, 0].SetBorders(MultipleBorders.Outside, Color.Black,
                        GemBox.Spreadsheet.LineStyle.Thin);

                    sheet1.Cells[i + 2, 1].Value = SessItogGrupGrid.Rows[i].Cells[1].Value;
                    sheet1.Cells[i + 2, 1].SetBorders(MultipleBorders.Outside, Color.Black,
                        GemBox.Spreadsheet.LineStyle.Thin);

                    if (SessItogGrupGrid.Rows[i].Cells[1].Value.ToString() == "0")
                        sheet1.Cells[i + 2, 1].Style.Font.Weight = ExcelFont.BoldWeight;
                    sheet1.Rows[i + 2].Height = 15 * 20;
                }
                else
                {
                    cr = sheet1.Cells.GetSubrange("A" + (i + 3).ToString(), "B" + (i + 3).ToString());
                    cr.Merged = true;
                    cr.SetBorders(MultipleBorders.Outside, Color.Black,
                        GemBox.Spreadsheet.LineStyle.Thin);
                    cr.Value = SessItogGrupGrid.Rows[i].Cells[0].Value;
                }
            }

            int x = 0;
            int ki = 0;
            for (ki = 2, i = 2; i < SessItogGrupGrid.Columns.Count; i += 2)
            {
                //MessageBox.Show(string.Format("{0} \n {1} \n {2}",i,SessItogGrupGrid.Columns[i].HeaderText,
                //    SessItogGrupGrid.Columns[i].Visible));

                if (mrs == 1)
                {
                    if (!SessItogGrupGrid.Columns[i].Visible) continue;
                }
                else
                {
                    if (!SessItogGrupGrid.Columns[i + 1].Visible) continue;
                }

                cr = sheet1.Cells.GetSubrange(Letters[ki] + 2.ToString(), Letters[ki + 1] + 2.ToString());
                cr.Merged = true;
                cr.SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
                cr.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
                cr.Style.VerticalAlignment = VerticalAlignmentStyle.Center;

                sheet1.Cells[1, ki].Value = SessItogGrupGrid.Columns[i + 1].HeaderText.ToUpper();
                sheet1.Cells[1, ki].Style.Rotation = 90;
                sheet1.Columns[ki].AutoFit();

                sheet1.Columns[ki].Width = 5 * 256;
                sheet1.Columns[ki].Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
                sheet1.Columns[ki + 1].Width = 5 * 256;
                sheet1.Columns[ki + 1].Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;

                for (int j = 0; j < SessItogGrupGrid.Rows.Count - 8; j++)  /////// -8
                {
                    string otm = SessItogGrupGrid.Rows[j].Cells[i + 1].Value.ToString();
                    string otmtxt = string.Empty;

                    switch (otm)
                    {
                        case "неудовлетворительно": otmtxt = "2"; break;
                        case "неуд+": otmtxt = "2+"; break;
                        case "удовлетворительно": otmtxt = "3"; break;
                        case "хорошо": otmtxt = "4"; break;
                        case "отлично": otmtxt = "5"; break;
                        case "зачтено": otmtxt = "зач"; break;
                        case "не зачтено": otmtxt = "нзач"; break;
                        default: otmtxt = ""; break;
                    }

                    if (mrs == 1)
                    {
                        sheet1.Cells[j + 2, ki].Value = SessItogGrupGrid.Rows[j].Cells[i].Value;
                        sheet1.Cells[j + 2, ki].SetBorders(MultipleBorders.Outside,
                            Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
                        if (otmtxt == "2" || otmtxt == "2+" || otmtxt == "нзач")
                        {
                            if (toolStripButton52.Checked)
                                sheet1.Cells[j + 2, ki + 1].Value = otmtxt;
                            else
                                sheet1.Cells[j + 2, ki + 1].Value = string.Empty;
                        }
                        else
                            sheet1.Cells[j + 2, ki + 1].Value = otmtxt;


                        sheet1.Cells[j + 2, ki + 1].SetBorders(MultipleBorders.Outside,
                            Color.Black, GemBox.Spreadsheet.LineStyle.Thin);


                        if (otmtxt == "2" || otmtxt == "2+" || otmtxt == "нзач")
                        {

                            sheet1.Cells[j + 2, ki + 1].Style.Font.Weight = ExcelFont.BoldWeight;
                            sheet1.Cells[j + 2, ki + 1].SetBorders(MultipleBorders.Outside, Color.Black,
                                GemBox.Spreadsheet.LineStyle.Thick);
                        }
                    }
                    else
                    {
                        cr =
                            sheet1.Cells.GetSubrange(Letters[ki] + (j + 3).ToString(),
                            Letters[ki + 1] + (j + 3).ToString());
                        cr.Merged = true;
                        //cr.SetBorders(MultipleBorders.Outside,
                        //    Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
                        cr.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;

                        if (otmtxt == "2" || otmtxt == "2+" || otmtxt == "нзач")
                        {
                            if (toolStripButton52.Checked)
                                cr.Value = otmtxt;
                            else
                                cr.Value = string.Empty;
                        }
                        else
                            cr.Value = otmtxt;


                        cr.SetBorders(MultipleBorders.Outside,
                            Color.Black, GemBox.Spreadsheet.LineStyle.Thin);


                        if (otmtxt == "2" || otmtxt == "2+" || otmtxt == "нзач")
                        {

                            cr.Style.Font.Weight = ExcelFont.BoldWeight;
                            cr.SetBorders(MultipleBorders.Outside, Color.Black,
                                GemBox.Spreadsheet.LineStyle.Thick);
                        }
                    }

                }
                ki += 2;
            }

            sheet1.Rows[1].Height = 135 * 20;
            //sheet1.Rows[x].Height = 94 * 20;
            sheet1.Columns[0].Width = 18 * 256;
            sheet1.Columns[1].Width = 8 * 256;

            sheet1.PrintOptions.HeaderMargin = 0.0;
            sheet1.PrintOptions.FooterMargin = 0.0;
            sheet1.PrintOptions.Portrait = false;

            // ---- кон: первый лист

            if (mrs != 0)
            {
                // --- лист 2 - рейтинг ------------
                global_query =
                    string.Format("select * from dbo.TGetGrupRating({0})",
                            SessGrupTable.Rows[SessItogGrpaCombo.SelectedIndex][0]);
                DataTable Rating = new DataTable();
                (new SqlDataAdapter(global_query, global_connection)).Fill(Rating);

                ExcelWorksheet sheet3 = excel.Worksheets.Add("Рейтинг");

                sheet3.Cells[1, 1].Value = "Рейтинг";
                sheet3.Cells[1, 1].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
                sheet3.Cells[1, 2].Value = "ФИО студента";
                sheet3.Cells[1, 2].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
                sheet3.Cells[1, 3].Value = "Сумма баллов";
                sheet3.Cells[1, 3].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);

                sheet3.Columns[1].Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
                sheet3.Columns[2].Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
                sheet3.Columns[3].Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;

                int jj = 2;
                foreach (DataRow r in Rating.Rows)
                {
                    sheet3.Cells[jj, 1].Value = r[2].ToString();
                    sheet3.Cells[jj, 1].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);

                    sheet3.Cells[jj, 2].Value = r[3].ToString();
                    sheet3.Cells[jj, 2].Style.HorizontalAlignment = HorizontalAlignmentStyle.Left;
                    sheet3.Cells[jj, 2].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);

                    sheet3.Cells[jj, 3].Value = r[1].ToString();
                    sheet3.Cells[jj, 3].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
                    jj++;
                }

                sheet3.Columns[1].AutoFit();
                sheet3.Columns[2].AutoFit();
                sheet3.Columns[3].AutoFit();

                sheet3.Cells[0, 2].Value = "Сведения о рейтинге студентов группы " + SessItogGrpaCombo.Text +
                    " за весь период обучения";
                sheet3.Cells[0, 3].Style.Font.Weight = ExcelFont.MaxWeight;

            }


            // --- третий лист - сводка
            ExcelWorksheet sheet4 = excel.Worksheets.Add("Сводка");

            int er = 2; //стартовая строка вывода в excel
            int cnum = 0;

            //предмет
            sheet4.Cells[1, 1].Value = "Предмет";
            sheet4.Cells[1, 1].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);

            // ср оценка
            sheet4.Cells[1, 2].Value = "Ср. оценка";
            sheet4.Cells[1, 2].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);

            //ср балл
            sheet4.Cells[1, 3].Value = "Ср. балл";
            sheet4.Cells[1, 3].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);

            //2
            sheet4.Cells[1, 4].Value = "Неуд.";
            sheet4.Cells[1, 4].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);

            //2+
            sheet4.Cells[1, 5].Value = "Неуд.+";
            sheet4.Cells[1, 5].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);


            //3
            sheet4.Cells[1, 6].Value = "Удовл.";
            sheet4.Cells[1, 6].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);


            //4
            sheet4.Cells[1, 7].Value = "Хорошо";
            sheet4.Cells[1, 7].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);

            //5
            sheet4.Cells[1, 8].Value = "Отлично";
            sheet4.Cells[1, 8].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);

            // зач
            sheet4.Cells[1, 9].Value = "Зачтено";
            sheet4.Cells[1, 9].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);


            //не зач
            sheet4.Cells[1, 10].Value = "Не зачтено";
            sheet4.Cells[1, 10].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);


            for (cnum = 3; cnum < SessItogGrupGrid.Columns.Count; cnum += 2, er++)
            {
                string predm = SessItogGrupGrid.Columns[cnum].HeaderText;
                predm = predm.Replace("\n", ", ");
                
                //предмет
                sheet4.Cells[er, 1].Value = predm;
                sheet4.Cells[er, 1].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);                

                // ср оценка
                SetCell(ref sheet4, SessItogGrupGrid.Rows[SessItogGrupGrid.Rows.Count - 8].Cells[cnum].Value.ToString(), er, 2, cnum);

                //ср балл
                SetCell(ref sheet4, SessItogGrupGrid.Rows[SessItogGrupGrid.Rows.Count - 8].Cells[cnum-1].Value.ToString(), er, 3, cnum-1);
                
                //2
                SetCell(ref sheet4, SessItogGrupGrid.Rows[SessItogGrupGrid.Rows.Count - 7].Cells[cnum].Value.ToString(), er, 4, cnum);

                //2+
                SetCell(ref sheet4, SessItogGrupGrid.Rows[SessItogGrupGrid.Rows.Count - 6].Cells[cnum].Value.ToString(), er, 5, cnum);

                //3
                SetCell(ref sheet4, SessItogGrupGrid.Rows[SessItogGrupGrid.Rows.Count - 5].Cells[cnum].Value.ToString(), er, 6, cnum);

                //4
                SetCell(ref sheet4, SessItogGrupGrid.Rows[SessItogGrupGrid.Rows.Count - 4].Cells[cnum].Value.ToString(), er, 7, cnum);

                //5
                SetCell(ref sheet4, SessItogGrupGrid.Rows[SessItogGrupGrid.Rows.Count - 3].Cells[cnum].Value.ToString(), er, 8, cnum);
                
                // зач
                SetCell(ref sheet4, SessItogGrupGrid.Rows[SessItogGrupGrid.Rows.Count - 2].Cells[cnum].Value.ToString(), er, 9, cnum);

                //не зач
                SetCell(ref sheet4, SessItogGrupGrid.Rows[SessItogGrupGrid.Rows.Count - 1].Cells[cnum].Value.ToString(), er, 10, cnum);
            }

            sheet4.Columns[1].AutoFit();

            excel.SaveXls(Path);
            Process.Start(Path);

        }

        void SetCell(ref ExcelWorksheet sh, string val, int row, int col, int cnum)
        {
            if (val != "-")
            {
                double dval = double.Parse(val);
                if (dval != 0)
                {
                    sh.Cells[row, col].Value = val;
                }
            }
            else
                sh.Cells[row, col].Value = val;

            sh.Cells[row, col].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
            sh.Cells[row, col].Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            sh.Cells[row, col].Style.VerticalAlignment = VerticalAlignmentStyle.Center;
        }

        private void задолженностиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int crow = SessItogGrupGrid.CurrentRow.Index;

            if (crow > SessItogGrupGrid.Rows.Count - 8)
            {
                MessageBox.Show("Выберите ФИО студента для получения списка задолженностей.", "Отказ операции.",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }


            int i = 0, j = 1;

            if (SessItogGrupGrid.Rows.Count == 0)
            {
                MessageBox.Show("Нет данных об аттестации.", "Отказ операции.",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            
            string sessid = SessItogGrupGrid.CurrentRow.Cells[2].Tag.ToString();
            string qu = "select 	'Предмет' = predmet.name, 'Оценка' = vid_otmetka.str_name, " +
                " 'Отчётность' = vid_zan.name, " +
                " 'Номер курса' = predmet.kurs_id,  " +
                " 'Семестр' = predmet.semestr, " +
                " 'Преподаватель' = dbo.GetPrepodFIOByID(prepod.id) " +
                "    from session " +
                "    join predmet on predmet.id = session.predmet_id " +
                "    join prepod on prepod.id = predmet.prepod_id " +
                "    join vid_zan on vid_zan.id = session.vid_zan_id " +
                "    join  vid_otmetka on  vid_otmetka.id = session.otmetka_id " +
                "    where session.student_id = (select student_id from session where session.id = " +  sessid +  ") " +
                "    and (session.otmetka_id = 2 or session.otmetka_id = 6  or session.otmetka_id = 12 or session.otmetka_id = 11 or session.otmetka_id = 10) and session.vid_zan_id <= 16 " +
                "    order by predmet.kurs_id, predmet.semestr ";

            DataTable DolgTable = new DataTable();
            (new SqlDataAdapter(qu, global_connection)).Fill(DolgTable);

            if (DolgTable.Rows.Count==0)
            {
                MessageBox.Show("Выбранный студент не имеет академических задолженностей.", 
                    "Отказ операции.",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            
            string attstr = "Задолженности студента -  " + SessItogGrupGrid.CurrentRow.Cells[0].Value;

            ExcelFile excel = new ExcelFile();
            ExcelWorksheet sheet1 = excel.Worksheets.Add(attstr);

            
            // -- запрос на сохранение

            saveExcel.Title = "Выберите или введите имя для файла отчёта";
            saveExcel.FileName = attstr.ToUpper() + ".xls";
            if (saveExcel.ShowDialog() != DialogResult.OK) return;

            string Path = saveExcel.FileName;

            // --- создание страниц

            sheet1.Cells[0, 1].Value = attstr.ToUpper();                
            sheet1.Cells[0, 1].Style.Font.Weight = ExcelFont.MaxWeight;

            i = 1;
            j = 1;

            foreach (DataColumn col in DolgTable.Columns)
            {
                sheet1.Cells[i, j].Value = col.ColumnName;                
                sheet1.Cells[i, j].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
                sheet1.Columns[j++].AutoFit();
            }

            i = 2;
            foreach (DataRow dr in DolgTable.Rows)
            {
                j = 1;
                foreach (object dc in dr.ItemArray)
                {
                    sheet1.Cells[i, j].Value = dc.ToString();
                    sheet1.Cells[i, j].SetBorders(MultipleBorders.Outside, Color.Black, GemBox.Spreadsheet.LineStyle.Thin);
                    sheet1.Columns[j++].AutoFit();
                }
                i++;
            }

            sheet1.Cells[i, 1].Value = "Итого задолженностей";
            sheet1.Cells[i, 1].Style.Font.Weight = ExcelFont.BoldWeight;

            sheet1.Cells[i, 2].Value = DolgTable.Rows.Count.ToString();
            sheet1.Cells[i, 2].Style.Font.Weight = ExcelFont.BoldWeight;

            excel.SaveXls(Path);
            Process.Start(Path);

        }


        // добавить студента в список курсовой работы
        private void добавитьСтудентаВСписокToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // проверить список студентов в курсовой работе, найти таких, которых нет в этом списке
            //DataTable SessionResultTable = new DataTable();
            //global_query = string.Format("exec dbo.TGetSessionResult {0}, {1}, {2}",
            //    SessGrupTable.Rows[SessGruplistBox.SelectedIndex][0],
            //    SessPredmetTable.Rows[PredmIndex][0],
            //    SessPredmetTable.Rows[PredmIndex][3]);
            //(new SqlDataAdapter(global_query, global_connection)).Fill(SessionResultTable);

            // вывести список

            //получить выбранного студента
        }

        private void SessionGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void повторитьЗанятиеЧерезНеделюToolStripMenuItem_Click(object sender, EventArgs e)
        {
            push_week();
        }
    }
}