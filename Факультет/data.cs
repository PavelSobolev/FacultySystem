using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Windows.Forms;

namespace FSystem
{
    /// <summary>
    /// данные для хранния таблицы "расписание"
    /// </summary>
    public class Data
    {
        /// <summary>
        /// словарь для доступа к ячейкам таблицы расписания
        /// </summary>
        public Dictionary<DateTime, Dictionary<string, Dictionary<int, Cell>>> data;

        /// <summary>
        /// источник данных для полей класса
        /// </summary>
        public System.Data.DataTable BD;
        /// <summary>
        /// количество столбцов (количество групп)
        /// </summary>
        public int cols;
        /// <summary>
        /// количество строк (количество пунктов расписания)
        /// </summary>
        public int rows;
        /// <summary>
        /// строковый фильтр по дате учебной недели
        /// </summary>
        private string query; //фильтр занятий на данную учебную неделю

        /// <summary>
        /// путь каталога для сохранения временных файлов программы
        /// </summary>
        public string path;

        public string Query
        {
            get
            {
                return query;
            }
        }


        /// <summary>
        /// таблица соответствия "дата-номер строки"
        /// </summary>
        public Dictionary<int, DateTime> date_row = new Dictionary<int, DateTime> ( );

        /// <summary>
        /// таблица соотвествия "группа - начальная колонка"
        /// (на каждую группу приходится по две колонки)
        /// </summary>
        public Dictionary<int, string> grupa_column = new Dictionary<int, string> ( );
        public Dictionary<string, int> column_grupa = new Dictionary<string, int> ( );

        /// <summary>
        /// таблица соотношения "пара - строка таблицы"
        /// </summary>
        public Dictionary<int, int> para_row = new Dictionary<int, int> ( );

        /// <summary>
        /// конструктор класса
        /// </summary>
        /// <param name="start_day">дата первого дня учебной недели</param>
        public Data ( DateTime start_day, List<int> empty_rows, List<string> grups, int fakultet )
        {
            rows = 42;
            cols = grups.Count;
            
            //задать соотвествие "дата-номер строки"
            DateTime curdate = start_day;

            for ( int i = 0; i < empty_rows.Count; i++ )
            {

                for ( int j = empty_rows[i] + 1; j < empty_rows[i] + 7; j++ )
                {
                    date_row.Add ( j, curdate );
                }
                curdate = curdate.AddDays ( 1 );
            }

            //задать соотвествие "группа-столбец"
            int kkk = 1;
            for ( int j = 0; j < grups.Count; j++ )
            {
                grupa_column.Add ( kkk, grups[j] );
                column_grupa.Add ( grups[j], kkk );
                kkk += 2;
            }

            //задать соотвествие "пара-номер_строки"
            for ( int i = 0; i < empty_rows.Count; i++ )
            {
                int para_count = 1;
                for ( int j = empty_rows[i] + 1; j < empty_rows[i] + 7; j++ )
                {
                    para_row.Add ( j, para_count );
                    para_count++;
                }
            }

            //создать фильтр для данной недели
            int ds = start_day.Day,
                ms = start_day.Month,
                ys = start_day.Year;

            int de = start_day.AddDays ( 6 ).Day,
                me = start_day.AddDays ( 6 ).Month,
                ye = start_day.AddDays ( 6 ).Year;


            string res = "";
            res = string.Format ( " and (y>={0} and y<={1}) ", ys, ye );

            if ( ms == 12 && me == 1 )
                res = res + " and (m=1 or m=12) ";
            else
                res = res + string.Format ( " and (m>={0} and m<={1})", ms, me );

            if ( de < ds )
                res = res + string.Format ( " and ((d>={0} and d<=31 and m={2}) or (d>=1 and d<={1} and m={3})) "
                                , ds, de, ms, me );
            else
                res = res + string.Format ( " and (d>={0} and d<={1}) ", ds, de );

            query = string.Format("select " +
                " gn=grupa.name, gi=grupa.id, " +
                " prepn = prepod.fam + ' ' + left(prepod.im,1) + '.' + left(prepod.ot,1) + '.'," +
                " predmn = predmet.name_krat, " +
                " audn = kabinet.nomer, " +
                " rasp.y, rasp.d, rasp.m, rasp.predmet_id, rasp.grupa_id, rasp.prepod_id, rasp.fakultet_id, " +
                " rasp.kurs_id, rasp.nom_zan, rasp.vid_zan_id, " +
                " rasp.kabinet_id, rasp.semestr_id, rasp.potok_id, " +
                " rasp.subgr_nomer, chas = rasp.kol_chas, " +
                " rasp.id, krat_name, predmet.delenie, tema, vid_del=vid_zan.delenie, vidname = vid_zan.name, prim = isnull(prim_text,'') " +  
                " , prfullname = predmet.name " + 
                " from rasp " +
                " join grupa on grupa.id=rasp.grupa_id " +
                " join prepod on prepod.id=rasp.prepod_id " +
                " join predmet on predmet.id=rasp.predmet_id " +
                " join kabinet on kabinet.id=rasp.kabinet_id " +
                " join vid_zan on vid_zan.id=rasp.vid_zan_id " +
                " where rasp.fakultet_id={0} and grupa.show_in_grid=1 ", fakultet)
                + res +
                " order by y, m, d, nom_zan, grupa.outorder";

            //получить данные раписания на указанную неделю
            SqlDataAdapter da = new SqlDataAdapter ( query, main.global_connection );
            BD = new DataTable ( "rasp" );
            da.Fill ( BD );

            //создать временную папку для размещения файлов данных
            path = Environment.GetEnvironmentVariable ( "TMP", EnvironmentVariableTarget.User ) +
                "\\Facultet";

            if ( !Directory.Exists ( path ) )
                Directory.CreateDirectory ( path );
            else
            {
                foreach ( string f in Directory.GetFiles ( path ) )
                {
                    try
                    {
                        File.Delete ( f );
                    }
                    catch ( Exception exx )
                    {
                        //
                    }
                }
            }

            if ( BD.Rows.Count > 0 ) BD.WriteXml ( path + "\\0.xml" );

            // -----------------------  сздать структуру данных для хранения информации в программе

            data =
            new Dictionary<DateTime, Dictionary<string, Dictionary<int, Cell>>> ( );

            //data[DateTime.Now]["П21"][6].egz[0] = true;
            //цикл по дням недели

            int paranumber = 0;
            for ( DateTime i = start_day; i <= start_day.AddDays ( 6 ); i = i.AddDays ( 1 ) )
            {
                Dictionary<string, Dictionary<int, Cell>> gruppa_cell =
                    new Dictionary<string, Dictionary<int, Cell>> ( );

                for ( int j = 0; j < grups.Count; j++ ) //цикл по группам
                {
                    string grname = grups[j];
                    Dictionary<int, Cell> para_gruppa = new Dictionary<int, Cell> ( );

                    for ( paranumber = 1; paranumber <= 6; paranumber++ ) //цикл по парам
                    {
                        Cell tmpcell = new Cell ( );
                        para_gruppa.Add ( paranumber, tmpcell );
                    }
                    gruppa_cell.Add ( grname, para_gruppa );
                }
                data.Add ( i, gruppa_cell );
            }

            // ----------------------------------   заполнить структуру data данными из базы
            foreach ( DataRow row in BD.Rows )
            {
                Cell c = new Cell ( );
                DateTime currdate = new DateTime ( (int) row["y"], (int) row["m"], (int) row["d"] );

                if ( this[currdate, row["gn"].ToString ( ), (int) row["nom_zan"]].id[0] > 0 ||
                    this[currdate, row["gn"].ToString ( ), (int) row["nom_zan"]].id[1] > 0 )
                {
                    c = this[currdate, row["gn"].ToString ( ), (int) row["nom_zan"]];
                }

                int sub = (int) row["subgr_nomer"];

                c.col[0] = column_grupa[row["gn"].ToString ( )];
                c.col[1] = c.col[0] + 1;

                c.row = 0;
                foreach ( int drow in date_row.Keys )
                {
                    if ( date_row[drow] == currdate )
                    {
                        c.row = drow + (int) row["nom_zan"] - 6;
                    }
                }

                if ( sub == 0 ) sub = 0; else sub--;

                c.subgr_nomer[sub] = (int) row["subgr_nomer"];
                
                c.id[sub] = (int) row["id"];//задать id зантяия
                
                c.y[sub] = (int) row["y"];
                c.m[sub] = (int) row["m"];
                c.d[sub] = (int) row["d"];
                c.y[sub] = (int) row["y"];

                c.predmet_id[sub] = (int) row["predmet_id"];
                c.predmet_name[sub] = row["predmn"].ToString();
                c.predmet_fullname[sub] = row["prfullname"].ToString();

                c.prepod_id[sub] = (int) row["prepod_id"];
                c.prepod_name[sub] = row["prepn"].ToString();
                
                c.grupa_id[sub] = (int) row["grupa_id"];
                
                c.fakultet_id[sub] = (int) row["fakultet_id"];
                
                c.kurs_id[sub] = (int) row["kurs_id"];
                
                c.nom_zan[sub] = (int) row["nom_zan"];
                
                c.vid_zan_id[sub] = (int) row["vid_zan_id"];
                c.vid_zan_name[sub] = row["krat_name"].ToString();
                c.vid_full_name[sub] = row["vidname"].ToString();
                c.vid_delenie[sub] = (bool)row["vid_del"];
                
                c.kabinet_id[sub] = (int) row["kabinet_id"];
                c.aud_name[sub] = row["audn"].ToString();                
                
                c.semestr_id[sub] = (int) row["semestr_id"];
                c.potok_id[sub] = (int) row["potok_id"];
                                
                c.delenie[sub] = (bool)row["delenie"];
                c.tema[sub] = row["tema"].ToString();
                c.str_prim[sub] = row["prim"].ToString();
                if (!row.IsNull("chas"))
                    c.col_chas[sub] = (double)row["chas"];
                else
                    c.col_chas[sub] = 0.0;


                if ((int)row["subgr_nomer"] == 0)
                {
                    c.copy_subgroups(sub, 1 - sub);
                }

                this[currdate, row["gn"].ToString ( ), (int) row["nom_zan"]] = c;
            }

        }

        /// <summary>
        /// получить дату для выбранной строки
        /// </summary>
        /// <param name="row">номер строки для получения даных</param>
        /// <returns>возвращает дату, на которую приходится данная строка</returns>
        public DateTime RowDate ( int row )
        {
            return date_row[row];
        }

        /// <summary>
        /// получить имя группы для данной колонки
        /// </summary>
        /// <param name="col"></param>
        /// <returns>возвоащает имя группы по номеру колонки</returns>
        public string ColumnGroup ( int col )
        {
            int _col = 0;

            if ( col % 2 == 0 )
                _col = col - 1;
            else
                _col = col;

            if ( !grupa_column.ContainsKey ( _col ) ) return "-";

            return grupa_column[_col];
        }


        /// <summary>
        /// возвращает номер подгруппы  для текущего столбца
        /// </summary>
        /// <param name="col">номер столбца</param>
        /// <returns>номер подгруппы</returns>
        public int ColumnSubGroup ( int col )
        {
            int _col = 0;

            if ( col % 2 == 0 )
                _col = 2;
            else
                _col = 1;

            //if (!grupa_column.ContainsKey(col)) return 0;

            return _col;
        }

        /// <summary>
        /// получить номер пары для данной ячейки
        /// </summary>
        /// <param name="row">номер строки в таблице расписания</param>
        /// <returns>возвращает номер пары, соотвествующей данной строке
        /// или -1, если не найдено соотвествия
        /// </returns>
        public int RowPair ( int row )
        {
            if ( !para_row.ContainsKey ( row ) ) return -1;

            return para_row[row];
        }


        /// <summary>
        /// возвращает ссылку на указанную ячейку таблицы расписания
        /// </summary>
        /// <param name="dt">дата ячейки</param>
        /// <param name="grupa">группа ячейки</param>
        /// <param name="para">пара ячейки</param>
        /// <returns>ссылка на ячейку</returns>
        public Cell this[DateTime dt, string grupa, int para]
        {
            get
            {
                return data[dt][grupa][para];
            }
            set
            {
                data[dt][grupa][para] = value;
                //вывести значение в ячейку таблицы для данного занятия
            }
        }

                
        /// <summary>
        /// определить, занят ли данный преподаватель в указанный 
        /// день на указанной на указанной паре
        /// </summary>
        /// <param name="c">ячейка</param>
        /// <param name="sg">подгруппа (1 или 2)</param>
        /// <param name="res">строковое сообщение</param>
        /// <param name="sex">пол преподавателя</param>
        /// <returns></returns>
        public bool IsPrepodBuisy ( Cell c, int sg, out string res, bool sex)
        {
            int prepod_sought = c.prepod_id[sg - 1];
            DateTime dt = new DateTime(c.y[sg - 1], c.m[sg - 1], c.d[sg - 1]);
            int para = c.nom_zan[sg - 1];
            Random rnd = new Random();

            res = "";
            string s = "";
            if (sex) s = "занят"; else s = "занята";

            foreach (string gr in column_grupa.Keys)
            {   
                Cell current = data[dt][gr][para];

                for (int i = 0; i < 2; i++)
                {
                    int current_p = current.prepod_id[i];
                    int current_g = current.grupa_id[i];

                    if (prepod_sought == current_p)
                    {
                        if (current_g != c.grupa_id[sg - 1])
                        {
                            res = "Преподаватель " + c.prepod_name[sg-1] + "\nуже " + s + "  в группе  " + gr + ".";
                            if (!CanFormStream(current, c))
                                return true;
                        }
                    }
                }
            }


            ///занят ли на других FSystemах --------------------------------------------
            
            string q = string.Format("select fakultet.name from rasp " +
                " join fakultet on fakultet.id=rasp.fakultet_id " +
                " where y={0} and m={1} and d={2} and nom_zan={3} and prepod_id={4} and fakultet.id<>{5}",
                dt.Year, dt.Month, dt.Day, para, prepod_sought, c.fakultet_id[sg - 1]);
            
            SqlDataAdapter sda = null;
            DataTable dtbl = null;
            
            try
            {
                sda = new SqlDataAdapter(q, main.global_connection);
                dtbl = new DataTable();
                sda.Fill(dtbl);
            }
            catch(Exception exx)
            {
                MessageBox.Show("Ошибка извлечения данных.\n" + exx.Message + "\n" + 
                    "Повторите данное действие через некоторое время.",
                    "Ошибка данных", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return true;
            }

            if (dtbl.Rows.Count > 0)
            {
                res = "Преподаватель " + c.prepod_name[sg - 1] + "\nуже " + s 
                    + " на FSystemе: \"" + dtbl.Rows[0][0].ToString() + "\"";
                return true;
            }

            c.potok_id[sg - 1] = rnd.Next();
         
            return false;
        }

        /// <summary>
        /// могут ли указанные ячейки образовать поток (занятие на несколько групп одновременно)
        /// </summary>
        /// <param name="c1">первая ячейка - ячейка, с которой производится сравнение</param>
        /// <param name="c2">вторая ячейка - проверяемая ячейка</param>
        /// <returns></returns>
        public bool CanFormStream ( Cell c1, Cell c2 )
        {
            //поток формируется только полными группами для лекционных занятий одного курса
            //или группами разных курсов для одного преподавателя и предмета c фикс. названием            

            bool res = ( ( c1.predmet_name[0].Trim ( ) == c2.predmet_name[0].Trim ( ) ) &&
                (c1.vid_zan_name[0].ToLower().Contains("лек") && c2.vid_zan_name[0].ToLower().Contains("лек")) && 
                /*( c1.id[0] == c1.id[1] && c2.id[0] == c2.id[1] ) &&*/
                ( c1.kabinet_id[0] == c2.kabinet_id[0]));

            //MessageBox.Show(res.ToString());

            if (res == true) //если поток возможен, то скопировать значение ид_потока из с1 в с2
            {
                int pid = 0;

                c2.potok_id[0] = 0;
                c2.potok_id[1] = 0;

                pid = c1.potok_id[0];

                c2.potok_id[0] = pid;
                c2.potok_id[1] = pid;

                // --- обнулить часы
                if (c1.col_chas[0]!=0)
                    c2.col_chas[0] = 0;

                if (c1.col_chas[1] != 0)
                    c2.col_chas[1] = 0;
            }            

            return res;
        }


        /// <summary>
        /// определить, занята ли на данной паре данная аудитория
        /// </summary>
        public bool IsRoomBuisy(Cell c, int sg, out string res)
        {

            res = ""; 
                          
            if (c.aud_name[sg-1] == "--") return false;
            
            int room_sought = 0;
            DateTime dt = DateTime.Now;
            int para = 0;

          
            room_sought = c.kabinet_id[sg - 1];
            dt = new DateTime(c.y[sg - 1], c.m[sg - 1], c.d[sg - 1]);
            para = c.nom_zan[sg - 1];

            if (c.predmet_name[sg - 1].ToLower().Contains("физ") &&
                c.predmet_name[sg - 1].ToLower().Contains("восп"))
                return false;

            foreach (string gr in column_grupa.Keys)
            {
                Cell current = data[dt][gr][para];

                for (int i = 0; i < 2; i++)
                {
                    int current_r = current.kabinet_id[i];
                    int current_g = current.grupa_id[i];

                    if (room_sought == current_r)
                    {
                        if (current_g != c.grupa_id[sg - 1])
                        {
                            res = "Аудитория " + c.aud_name[sg-1] + " уже занята группой " + gr + ".";
                            if (!CanFormStream(current, c))
                                return true;
                        }
                    }
                }
            }

            return false;
        }      

        // -------------------   конец определения класса Data -------------------
    }

    public class Cell
    {
        public int row = 0;      //1
        public int[] col = new int[2] { 0, 0 }; //2

        //дата и время проведения
        public int[] m = new int[2] { 0, 0 };      //1
        public int[] d = new int[2] { 0, 0 };      //2
        public int[] y = new int[2] { 0, 0 };      //3

        //строковые данные
        public string[] prepod_name = new string[2] { "", "" };  //4
        public string[] predmet_name = new string[2] { "", "" }; //5
        public string[] predmet_fullname = new string[2] { "", "" }; //27
        public string[] aud_name = new string[2] { "", "" }; //для отображения в таблице //6
        public string[] vid_zan_name = new string[2] { "", "" };  //7
        public string[] tema = new string[2] { string.Empty, string.Empty };
        public string[] str_prim = new string[2] { string.Empty, string.Empty };
        public string[] vid_full_name = new string[2] { string.Empty, string.Empty };

        //номер подгруппы
        public int[] subgr_nomer = new int[2] { 0, 0 };         //8
        
        //номер потока (одно значение для всех занятий, объединяемых в поток)
        public int[] potok_id = new int[2] { 0, 0 };            //9

        public int[] id = new int[2] { 0, 0 };                  //10
        public int[] fakultet_id = new int[2] { 0, 0 };         //11
        public int[] grupa_id = new int[2] { 0, 0 };            //12
        public int[] kabinet_id = new int[2] { 0, 0 };          //13
        
        public int[] kurs_id = new int[2] { 0, 0 };             //14

        public int[] nom_zan = new int[2] { 0, 0 };             //15
        public int[] predmet_id = new int[2] { 0, 0 };          //16
        public int[] prepod_id = new int[2] { 0, 0 };           //17
        public int[] semestr_id = new int[2] { 0, 0 };          //18
        public int[] vid_zan_id = new int[2] { 0, 0 };          //19
        public bool[] delenie = new bool[2] { false, false }; //20
        public bool[] vid_delenie = new bool[2] { false, false };
        public double[] col_chas = new double[2] {0.0, 0.0};
        public int[] uch_god = new int[2] { 1, 1 };
        
        
        /// <summary>
        /// определить, использует ли данная ячейка оба элемента массива для своих свойств
        /// </summary>
        /// <returns></returns>
        public bool use_two_cells ( )
        {
            if ( subgr_nomer[0] == 1 || subgr_nomer[1] == 2 )
                return false;
            else
                return true;
        }

        public bool Divided()
        {
            if (id[0] != id[1])
                return true;
            else
                return false;
        }

        /// <summary>
        /// вернуть дату подгруппы по ее номеру
        /// </summary>
        /// <param name="index"></param>
        /// <returns>дата подгруппы</returns>
        public DateTime this[int index]
        {
            get
            {
                return new DateTime ( y[index], m[index], d[index] );
            }            
        }

        /// <summary>
        /// скопировать данные о занятии из одной подгруппы в другую подгруппу
        /// </summary>
        /// <param name="from">номер копируемой подгруппы</param>
        /// <param name="to">номер целевой подгруппы</param>
        public void copy_subgroups(int from, int to)
        {
            id[to] = id[from];

            nom_zan[to] = nom_zan[from];

            y[to] = y[from];
            m[to] = m[from];
            d[to] = d[from];

            fakultet_id[to] = fakultet_id[from];

            grupa_id[to] = grupa_id[from];

            predmet_id[to] = predmet_id[from];
            predmet_name[to] = predmet_name[from];
            predmet_fullname[to] = predmet_fullname[from];

            delenie[to] = delenie[from];
            vid_delenie[to] = vid_delenie[from];

            prepod_id[to] = prepod_id[from];
            prepod_name[to] = prepod_name[from];

            subgr_nomer[to] = subgr_nomer[from];

            aud_name[to] = aud_name[from];
            kabinet_id[to] = kabinet_id[from];

            kurs_id[to] = kurs_id[from];
            potok_id[to] = potok_id[from];
          
            semestr_id[to] = semestr_id[from];

            vid_zan_id[to] = vid_zan_id[from];
            vid_zan_name[to] = vid_zan_name[from];
            vid_full_name[to] = vid_full_name[from];
            tema[to] = tema[from];
            str_prim[to] = str_prim[from];
            col_chas[to] = col_chas[from];
            uch_god[to] = uch_god[from];
        }

        /// <summary>
        /// скопировать поля одной ячейки Cell в другую ячейку Cell
        /// </summary>
        /// <param name="from">источник</param>
        public void copy_fields(Cell source)
        {
            id[0] = source.id[0]; id[1] = source.id[1];

            nom_zan[0] = source.nom_zan[0]; nom_zan[1] = source.nom_zan[1];

            y[0] = source.y[0]; y[1] = source.y[1];
            m[0] = source.m[0]; m[1] = source.m[1];
            d[0] = source.d[0]; d[1] = source.d[1];

            fakultet_id[0] = source.fakultet_id[0]; fakultet_id[1] = source.fakultet_id[1];

            grupa_id[0] = source.grupa_id[0]; grupa_id[1] = source.grupa_id[1];

            predmet_id[0] = source.predmet_id[0]; predmet_id[1] = source.predmet_id[1];
            predmet_name[0] = source.predmet_name[0]; predmet_name[1] = source.predmet_name[1];
            predmet_fullname[0] = source.predmet_fullname[0]; predmet_fullname[1] = source.predmet_fullname[1];
            delenie[0] = source.delenie[0]; delenie[1] = source.delenie[1];
            vid_delenie[0] = source.vid_delenie[0]; vid_delenie[1] = source.vid_delenie[1];

            prepod_id[0] = source.prepod_id[0]; prepod_id[1] = source.prepod_id[1];
            prepod_name[0] = source.prepod_name[0]; prepod_name[1] = source.prepod_name[1];

            subgr_nomer[0] = source.subgr_nomer[0]; subgr_nomer[1] = source.subgr_nomer[1];

            aud_name[0] = source.aud_name[0]; aud_name[1] = source.aud_name[1];
            kabinet_id[0] = source.kabinet_id[0]; kabinet_id[1] = source.kabinet_id[1];

            kurs_id[0] = source.kurs_id[0]; kurs_id[1] = source.kurs_id[1];
            potok_id[0] = source.potok_id[0]; potok_id[1] = source.potok_id[1];

            semestr_id[0] = source.semestr_id[0]; semestr_id[1] = source.semestr_id[1];

            vid_zan_id[0] = source.vid_zan_id[0]; vid_zan_id[1] = source.vid_zan_id[1];
            vid_zan_name[0] = source.vid_zan_name[0]; vid_zan_name[1] = source.vid_zan_name[1];
            vid_full_name[0] = source.vid_full_name[0]; vid_full_name[1] = source.vid_full_name[1];
            tema[0] = source.tema[0]; tema[1] = source.tema[1];
            str_prim[0] = source.str_prim[0]; str_prim[1] = source.str_prim[1];
            col_chas[0] = source.col_chas[0]; col_chas[1] = source.col_chas[1];
            uch_god[0] = source.uch_god[0]; uch_god[1] = source.uch_god[1];
        }

        /// <summary>
        /// задать нулевые значения для полей подгруппы
        /// </summary>
        /// <param name="number">номер подгруппы (0 или 1)</param>
        public void drop_subgroup(int number)
        {
            id[number] = 0;

            nom_zan[number] = 0;

            y[number] = 0;
            m[number] = 0;
            d[number] = 0;

            fakultet_id[number] = 0;

            grupa_id[number] = 0;

            predmet_id[number] = 0;
            predmet_name[number] = "";
            predmet_fullname[number] = "";
            delenie[number] = false;
            vid_delenie[number] = false;

            prepod_id[number] = 0;
            prepod_name[number] = "";

            subgr_nomer[number] = 0;

            aud_name[number] = "";
            kabinet_id[number] = 0;

            kurs_id[number] = 0;
            potok_id[number] = 0;

            semestr_id[number] = 0;

            vid_zan_id[number] = 0;
            vid_zan_name[number] = "";
            vid_full_name[number] = "";
            tema[number] = string.Empty;
            str_prim[number] = string.Empty;
            col_chas[number] = 0.0;
            uch_god[number] = 0;
        }

        /// <summary>
        /// поменять местами подгруппы
        /// </summary>        
        public void swap_subgroups()
        {
            int itmp = 0;
            string stmp = "";
            bool btmp = false;
            double dtmp = 0.0;

            itmp = id[0]; id[0] = id[1]; id[1] = itmp;

            itmp = nom_zan[0]; nom_zan[0] = nom_zan[1]; nom_zan[1] = itmp;

            itmp = y[0]; y[0] = y[1]; y[1] = itmp;            
            itmp = m[0]; m[0] = m[1]; m[1] = itmp;            
            itmp = d[0]; d[0] = d[1]; d[1] = itmp;

            itmp = fakultet_id[0]; fakultet_id[0] = fakultet_id[1]; fakultet_id[1] = itmp;

            itmp = grupa_id[0]; grupa_id[0] = grupa_id[1]; grupa_id[1] = itmp;

            itmp = predmet_id[0]; predmet_id[0] = predmet_id[1]; predmet_id[1] = itmp;
            stmp = predmet_name[0]; predmet_name[0] = predmet_name[1]; predmet_name[1] = stmp;
            stmp = predmet_fullname[0]; predmet_fullname[0] = predmet_fullname[1]; predmet_fullname[1] = stmp;

            btmp = delenie[0]; delenie[0] = delenie[1]; delenie[1] = btmp;
            btmp = vid_delenie[0]; vid_delenie[0] = vid_delenie[1]; vid_delenie[1] = btmp;

            itmp = prepod_id[0]; prepod_id[0] = prepod_id[1]; prepod_id[1] = itmp;
            stmp = prepod_name[0]; prepod_name[0] = prepod_name[1]; prepod_name[1] = stmp;

            itmp = subgr_nomer[0]; subgr_nomer[0] = subgr_nomer[1]; subgr_nomer[1] = itmp;

            subgr_nomer[0] = 1;
            subgr_nomer[1] = 2;

            stmp = aud_name[0]; aud_name[0] = aud_name[1]; aud_name[1] = stmp;
            itmp = kabinet_id[0]; kabinet_id[0] = kabinet_id[1]; kabinet_id[1] = itmp;

            itmp = kurs_id[0]; kurs_id[0] = kurs_id[1]; kurs_id[1] = itmp;
            itmp = potok_id[0]; potok_id[0] = potok_id[1]; potok_id[1] = itmp;

            itmp = semestr_id[0]; semestr_id[0] = semestr_id[1]; semestr_id[1] = itmp;

            itmp = vid_zan_id[0]; vid_zan_id[0] = vid_zan_id[1]; vid_zan_id[1] = itmp;
            stmp = vid_zan_name[0]; vid_zan_name[0] = vid_zan_name[1]; vid_zan_name[1] = stmp;
            stmp = vid_full_name[0]; vid_full_name[0] = vid_full_name[1]; vid_full_name[1] = stmp;

            stmp = tema[0]; tema[0] = tema[1]; tema[1] = stmp;
            stmp = str_prim[0]; str_prim[0] = str_prim[1]; str_prim[1] = stmp;

            dtmp = col_chas[0]; col_chas[0] = col_chas[1]; col_chas[1] = dtmp;
            
            //учебный год не меняет своего значения при обмене
        }


        /// <summary>
        /// формирует команду вставки новой записи
        /// </summary>
        /// <param name="sg">номер ячейки, значение которой сохраняется в БД (0 или 1)</param>
        /// <returns>команда типа данных SqlCommand</returns>
        public SqlCommand InsertCommand(int sg)
        {
            SqlCommand scmd = new SqlCommand();
            scmd.Connection = main.global_connection;
            
            string strquery = "insert into rasp " + 
                "(  d,       m,       y,          predmet_id, grupa_id,   prepod_id,   fakultet_id, " +
                "   kurs_id, nom_zan, vid_zan_id, kabinet_id, semestr_id, subgr_nomer, kol_chas,  " +
                //"   lekt,    prakt,   sem,        lab,        individ,    kons,        egz, " + 
                //"   zach,    kurs,    kontr,      " + 
                "   potok_id, uch_god_id " + 
                " ) "  +
                " values " +
                " ( @D,       @M,       @Y,          @PREDMET_ID, @GRUPA_ID,   @PREPOD_ID,   @FAKULTET_ID, " +
                "   @KURS_ID, @NOM_ZAN, @VID_ZAN_ID, @KABINET_ID, @SEMESTR_ID, @SUBGR_NOMER, @KOL_CHAS,  " +
                //"   @LEKT,    @PRAKT,   @SEM,        @LAB,        @INDIVID,    @KONS,        @EGZ, " + 
                //"   @ZACH,    @KURS,    @KONTR, " +      
                " @POTOK_ID, @UCH_G " +
                " )" ; //24

            //----------------------------------------------------------------------------------------
            //StreamWriter sw = new StreamWriter("c:\\d.txt", false, System.Text.Encoding.Default);
            string s = string.Format("insert into rasp " +
                "(  d,       m,       y,          predmet_id, grupa_id,   prepod_id,   fakultet_id, " +
                "   kurs_id, nom_zan, vid_zan_id, kabinet_id, semestr_id, subgr_nomer, kol_chas,  " +
                "   potok_id, uch_god_id " +
                " ) " +
                " values " +
                " ( {0},       {1},       {2},          {3}, {4},   {5},   {6}, " +
                "   {7}, {8}, {9}, {10}, {11}, {12}, (13),  " +
                " {14}, {15} " +
                " )",
                d[sg],
                m[sg],
                y[sg],
                predmet_id[sg],
                grupa_id[sg],
                prepod_id[sg],
                fakultet_id[sg],
                kurs_id[sg],
                nom_zan[sg],
                vid_zan_id[sg],
                kabinet_id[sg],
                semestr_id[sg],
                subgr_nomer[sg],                
                col_chas[sg],
                potok_id[sg],
                main.uch_god);

            //sw.WriteLine(s);
            //sw.Close();
            //-------------------------------------------------------------------------------------

            scmd.CommandText = strquery;

            scmd.Parameters.Add("@D", SqlDbType.Int).Value = d[sg];
            scmd.Parameters.Add("@M", SqlDbType.Int).Value = m[sg];
            scmd.Parameters.Add("@Y", SqlDbType.Int).Value = y[sg];
            scmd.Parameters.Add("@PREDMET_ID", SqlDbType.Int).Value = predmet_id[sg];
            scmd.Parameters.Add("@GRUPA_ID", SqlDbType.Int).Value = grupa_id[sg];
            scmd.Parameters.Add("@PREPOD_ID", SqlDbType.Int).Value = prepod_id[sg];
            scmd.Parameters.Add("@FAKULTET_ID", SqlDbType.Int).Value = fakultet_id[sg];
            scmd.Parameters.Add("@KURS_ID", SqlDbType.Int).Value = kurs_id[sg];
            scmd.Parameters.Add("@NOM_ZAN", SqlDbType.Int).Value = nom_zan[sg];
            scmd.Parameters.Add("@VID_ZAN_ID", SqlDbType.Int).Value = vid_zan_id[sg];
            scmd.Parameters.Add("@KABINET_ID", SqlDbType.Int).Value = kabinet_id[sg];
            scmd.Parameters.Add("@SEMESTR_ID", SqlDbType.Int).Value = semestr_id[sg];
            scmd.Parameters.Add("@SUBGR_NOMER", SqlDbType.Int).Value = subgr_nomer[sg];
            scmd.Parameters.Add("@KOL_CHAS", SqlDbType.Float).Value = col_chas[sg];
            scmd.Parameters.Add("@POTOK_ID", SqlDbType.Int).Value = potok_id[sg];
            scmd.Parameters.Add("@UCH_G", SqlDbType.Int).Value = main.uch_god;
                        
            return scmd;
        }

        /// <summary>
        /// формирует команду для обновления для данной ячейки
        /// </summary>
        /// <param name="sg">номер ячейки, значение которой сохраняется в БД (0 или 1)</param>
        /// <returns>команда типа данных SqlCommand</returns>
        public SqlCommand UpdateCommand(int sg)
        {
            SqlCommand scmd = new SqlCommand();
            scmd.Connection = main.global_connection;

            string strquery = "update rasp set " +
                " predmet_id=@PREDMET_ID, prepod_id=@PREPOD_ID, " +
                " vid_zan_id=@VID_ZAN_ID, kabinet_id=@KABINET_ID, subgr_nomer=@SUBGR_NOMER, kol_chas=@KOL_CHAS,  " +                
                " potok_id=@POTOK_ID " +
                " where id = @ID";

            scmd.CommandText = strquery;

            scmd.Parameters.Add("@PREDMET_ID", SqlDbType.Int).Value = predmet_id[sg];
            scmd.Parameters.Add("@PREPOD_ID", SqlDbType.Int).Value = prepod_id[sg];
            scmd.Parameters.Add("@VID_ZAN_ID", SqlDbType.Int).Value = vid_zan_id[sg];
            scmd.Parameters.Add("@KABINET_ID", SqlDbType.Int).Value = kabinet_id[sg];
            scmd.Parameters.Add("@SUBGR_NOMER", SqlDbType.Int).Value = subgr_nomer[sg];
            scmd.Parameters.Add("@KOL_CHAS", SqlDbType.Float).Value = col_chas[sg];
            scmd.Parameters.Add("@POTOK_ID", SqlDbType.Int).Value = potok_id[sg];
            scmd.Parameters.Add("@ID", SqlDbType.Int).Value = id[sg];


            return scmd;
        }

        /// <summary>
        /// формирует команду для удаления данной ячейки
        /// </summary>
        /// <param name="id">идентификатор записи</param>
        /// <param name="sg">номер ячейки (0 или 1)</param>
        /// <returns></returns>
        public SqlCommand DeleteCommand(int sg)
        {
            SqlCommand scmd = new SqlCommand();
            scmd.Connection = main.global_connection;

            string strquery = "delete from rasp where id = @ID";

            scmd.CommandText = strquery;

            scmd.Parameters.Add("@ID", SqlDbType.Int).Value = id[sg];

            return scmd;
        }

        public SqlCommand UpdateTemaCommand(int sg)
        {
            SqlCommand scmd = new SqlCommand();
            scmd.Connection = main.global_connection;

            string strquery = "update rasp set " +
                " tema=@TEMA where id = @ID";

            scmd.CommandText = strquery;

            scmd.Parameters.Add("@TEMA", SqlDbType.NVarChar).Value = tema[sg];            
            scmd.Parameters.Add("@ID", SqlDbType.Int).Value = id[sg];


            return scmd;
        }

    }

    /// <summary>
    /// сравнитель данных для построения списка, упорядоченного по возрастанию
    /// </summary>
    /// <typeparam name="T"></typeparam>
    class DescendingComparer<T> : IComparer<T> where T : IComparable<T> 
    { 
        public int Compare(T x, T y) 
        { 
            return y.CompareTo(x); 
        } 
    }

    // ----------   Конец определения пространтсва имен ---------------
}