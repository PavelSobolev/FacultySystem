using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace FSystem
{
    public partial class student_edit : Form
    {
        public student_edit()
        {
            InitializeComponent();
        }

        #region fams_ots_ims
        public string[] mailfam = new string[] {
            "Авдосьев",
            "Аврутин",
            "Азанов",
            "Акимов",
            "Александров",
            "Амосов",
            "Ан",
            "Андреев",
            "Андреенко",
            "Антонов",
            "Артемьев",
            "Асямов",
            "Афанасьев",
            "Барабанов",
            "Баранов",
            "Бедак",
            "Белокосов",
            "Бельтюков",
            "Берлин",
            "Берняцкий",
            "Бобров",
            "Богданов",
            "Богомолов",
            "Боку",
            "Борисов",
            "Бородей",
            "Брагужин",
            "Бурдин",
            "Бурмистров",
            "Вальтер",
            "Василюк",
            "Верминский",
            "Вилижанин",
            "Виноградов",
            "Воржов",
            "Гаврилов",
            "Гвон",
            "Генералов",
            "Гилязетдинов",
            "Глухов",
            "Гордеев",
            "Горло",
            "Григорьев",
            "Гризодуб",
            "Гришин",
            "Гударев",
            "Гусев",
            "Гуща",
            "Давыдов",
            "Диденко",
            "Дорничев",
            "Дорогов",
            "Дощинский",
            "Емельянченко",
            "Еремин",
            "Ефименко",
            "Ефимов",
            "Ефремов",
            "Жариков",
            "Жарков",
            "Жеребцов",
            "Жихарев",
            "Жук",
            "Жуков",
            "Забелин",
            "Заикин",
            "Залялетдинов",
            "Зарубин",
            "Захаренков",
            "Звягин",
            "Зенкин",
            "Злобин",
            "Зубак",
            "Зыков",
            "Зюльков",
            "И",
            "Иванов",
            "Игушкин",
            "Идрисов",
            "Ильев",
            "Кайстренко",
            "Калашников",
            "Камнев",
            "Кампан",
            "Кан",
            "Ким",
            "Ким",
            "Кирсанов",
            "Кияница",
            "Климов",
            "Кляцкий",
            "Ко",
            "Ковалев",
            "Ковалев",
            "Козорез",
            "Колесников",
            "Колтуновский",
            "Комаров",
            "Комков",
            "Комлев",
            "Кондратьев",
            "Конорюков",
            "Коробков",
            "Корсун",
            "Косоруков",
            "Котляров",
            "Котляров",
            "Кошович",
            "Крюков",
            "Кубеков",
            "Кударенко",
            "Кудин",
            "Кузин",
            "Кузнецов",
            "Кузнецов",
            "Кузьмин",
            "Куцов",
            "Лавренов",
            "Лапенко",
            "Лапп",
            "Ласточкин",
            "Лебедев",
            "Лепехин",
            "Лещук",
            "Ли",
            "Лилюев",
            "Лихач",
            "Лобанов",
            "Логинов",
            "Лось",
            "Лучинин",
            "Лучников",
            "Лягуцкий",
            "Маковецкий",
            "Максаев",
            "Максачук",
            "Максименко",
            "Макушев",
            "Малахов",
            "Маленков",
            "Марков",
            "Мартыненко",
            "Мартынюк",
            "Мартьянов",
            "Мащенко",
            "Мезин",
            "Мельниченко",
            "Мерзляков",
            "Мирзоев",
            "Мискинис",
            "Митрофанов",
            "Митрохин",
            "Михеев",
            "Моисеев",
            "Морозов",
            "Москвитин",
            "Моталыгин",
            "Мухорин",
            "Мушников",
            "Насенков",
            "Насибулин",
            "Несмеев",
            "Никитцов",
            "Николаев",
            "Новиков",
            "Оганесян",
            "Орлянский",
            "Остапенко",
            "Осыкин",
            "Пак",
            "Парахин",
            "Пахарь",
            "Пенский",
            "Переверзев",
            "Поздняков",
            "Портных",
            "Потапов",
            "Прасолов",
            "Прилепский",
            "Проскуряков",
            "Протещенко",
            "Проценко",
            "Пузаренко",
            "Раевский",
            "Ри",
            "Роев",
            "Рыбалко",
            "Рыжов",
            "Рябушкин",
            "Савченко",
            "Сайко",
            "Салихов",
            "Самарин",
            "Самарский",
            "Самочко",
            "Сапегин",
            "Саранчин",
            "Саратовцев",
            "Северов",
            "Седачев",
            "Сеначин",
            "Сергиенко",
            "Сидоров",
            "Сим",
            "Скворцов",
            "Скляр",
            "Скнарин",
            "Скоробогатый",
            "Скрябин",
            "Соболев",
            "Соколов",
            "Соловьев",
            "Спатарь",
            "Ставсков",
            "Степанов",
            "Степанский",
            "Столповский",
            "Стрепетов",
            "Сулейманов",
            "Тен",
            "Теплов",
            "Терещенко",
            "Тибайкин",
            "Тим",
            "Тимашов",
            "Тимофеев",
            "Титов",
            "Топырик",
            "Трахачев",
            "Трофименко",
            "Трофимов",
            "Тупиков",
            "Тян",
            "Унтевский",
            "Урес",
            "Утенков",
            "Ушаков",
            "Фатеев",
            "Федотов",
            "Федянин",
            "Фендриков",
            "Фещенко",
            "Филенко",
            "Фоканов",
            "Фоменко",
            "Французов",
            "Фугенфиров",
            "Хавроничев",
            "Харитонов",
            "Харченко",
            "Хе",
            "Хмара",
            "Хорошавин",
            "Цой",
            "Чайкин",
            "Чен",
            "Черняев",
            "Чертов",
            "Чесноков",
            "Чистяков",
            "Чмутов",
            "Чо",
            "Чужинов",
            "Чун",
            "Шаймарданов",
            "Шалыгин",
            "Шамараев",
            "Шарапов",
            "Шаройкин",
            "Шведов",
            "Шебаршов",
            "Шевелев",
            "Шевченко",
            "Шевченко",
            "Шелюто",
            "Шестибратов",
            "Шубин",
            "Щелчков",
            "Щукин",
            "Югай",
            "Юденок",
            "Юдин",
            "Юн",
            "Языков",
            "Яхонтов",
            "Ячменёв" };

        public string[] mailim = new string[] { 
            "Александр",
            "Алексей",
            "Алик",
            "Альберт",
            "Анатолий",
            "Андрей",
            "Анисим",
            "Антон",
            "Антонин",
            "Аристарх",
            "Аркадий",
            "Артём",
            "Артемий",
            "Артур",
            "Архипп",
            "Арчил",
            "Борис",
            "Боян",
            "Вадим",
            "Валентин",
            "Валериан",
            "Валерий",
            "Василий",
            "Вахтанг",
            "Виктор",
            "Виссарион",
            "Виталий",
            "Владимир",
            "Владислав",
            "Всеволод",
            "Вячеслав",
            "Гавриил",
            "Геннадий",
            "Георгий",
            "Герасим",
            "Герман",
            "Глеб",
            "Григорий",
            "Давид",
            "Дамир",
            "Дён",
            "Денис",
            "Дмитрий",
            "Евгений",
            "Евдокий",
            "Евдоким",
            "Егор",
            "Ен",
            "Ермолай",
            "Ефрем",
            "Зиновий",
            "Иван",
            "Игнатий",
            "Игорь",
            "Илья",
            "Карп",
            "Кирилл",
            "Константин",
            "Лев",
            "Леонард",
            "Леонид",
            "Леонтий",
            "Макар",
            "Максим",
            "Марк",
            "Матвей",
            "Мирон",
            "Митрофан",
            "Михаил",
            "Нестор",
            "Никита",
            "Николай",
            "Овсеп",
            "Олег",
            "Павел",
            "Пантолеон",
            "Петр",
            "Петроний",
            "Платон",
            "Ревкат",
            "Родион",
            "Роман",
            "Ромил",
            "Ростислав",
            "Руслан",
            "Рустам",
            "Савва",
            "Святослав",
            "Семен",
            "Сергей",
            "Станислав",
            "Тимофей",
            "Феликс",
            "Филимон",
            "Филипп",
            "Харитон",
            "Хиун",
            "Эдуард",
            "Эрнест",
            "Юрий",
            "Ян",
            "Ярослав" };

        public string[] mailot = new string[] { 
            "Александрович",
            "Алексеевич",
            "Амирханович",
            "Анатольевич",
            "Андреевич",
            "Бисланович",
            "Борисович",
            "Валентинович",
            "Валериевич",
            "Васильевич",
            "Викторович",
            "Владимирович",
            "Владленович",
            "Вонсуевич",
            "Вонтэенович",
            "Вячеславович",
            "Гвондяевич",
            "Геннадьевич",
            "Георгиевич",
            "Герасимович",
            "Гир",
            "Григорьевич",
            "Гю",
            "Дегынович",
            "Десенович",
            "Дмитриевич",
            "Евгеньевич",
            "Енкириевич",
            "Енманович",
            "Еннамович",
            "Иванович",
            "Игоревич",
            "Исонович",
            "Камильевич",
            "Кентыгиевич",
            "Кириллович",
            "Константинович",
            "Кузьмич",
            "Леонидович",
            "Львович",
            "Менкильевич",
            "Михайлович",
            "Назарович",
            "Намильевич",
            "Николаевич",
            "Оккунович",
            "Олегович",
            "Оливерович",
            "Охенович",
            "Павлович",
            "Пёнзович",
            "Петрович",
            "Понгильевич",
            "Ришатович",
            "Родионович",
            "Романович",
            "Самиулаевич",
            "Санчарович",
            "Семенович",
            "Сергеевич",
            "Сондоевич",
            "Сонненович",
            "Сонсуевич",
            "Станиславович",
            "Сундонович",
            "Теннамиевич",
            "Тесуевич",
            "Теханович",
            "Тонхоевич",
            "Унбонович",
            "Хесонович",
            "Чанбеевич",
            "Ченманович",
            "Чунгирович",
            "Эдуардович",
            "Эрихович",
            "Юрьевич"};

        public string[] femailfam = new string[] { 
            "Алексютина",
            "Артеменко",
            "Бак",
            "Барыбина",
            "Быкова",
            "Васильева",
            "Гадеудина",
            "Голышева",
            "Громыко",
            "Демидова",
            "Деникина",
            "Джумайло",
            "Дорофеева",
            "Дудник",
            "Дю",
            "Ерыгина",
            "Ефремова",
            "Жеребцова",
            "Иванова",
            "Кан",
            "Карбышева",
            "Ким",
            "Князева",
            "Козлова",
            "Колганова",
            "Кон",
            "Кошелева",
            "Краснобаева",
            "Левченко",
            "Маер",
            "Мазур",
            "Маловичко",
            "Мамчева",
            "Мартина",
            "Марченко",
            "Медведенко",
            "Мещерякова",
            "Можаева",
            "Моргун",
            "Николаева",
            "Олейник",
            "Олимпиева",
            "Пакеева",
            "Пантелеева",
            "Петряева",
            "Пешкова",
            "Плотникова",
            "Пономаренко",
            "Самутенко",
            "Семикина",
            "Сметанина",
            "Смольникова",
            "Сорока",
            "Сороко",
            "Супрунчук",
            "Тарасова",
            "Тен",
            "Тимошенко",
            "Трофименко",
            "Филенко",
            "Че",
            "Чевычелова",
            "Чернова",
            "Ческидова",
            "Чирко",
            "Чирскова",
            "Шалкус",
            "Шевченко",
            "Шевченко",
            "Эртен",
            "Ясенева" };

        public string[] femailim = new string[] {
            "Алевтина",
            "Александра",
            "Алла",
            "Анастасия",
            "Анна",
            "Антонина",
            "Анфиса",
            "Валентина",
            "Валерия",
            "Варвара",
            "Василиса",
            "Васса",
            "Вера",
            "Вероника",
            "Виктория",
            "Галина",
            "Дарья",
            "Ева",
            "Евгения",
            "Евлампия",
            "Екатерина",
            "Елена",
            "Зинаида",
            "Зоя",
            "Ирина",
            "Кира",
            "Клавдия",
            "Ксения",
            "Лариса",
            "Лидия",
            "Любовь",
            "Людмила",
            "Маргарита",
            "Марианна",
            "Марина",
            "Мария",
            "Надежда",
            "Нана",
            "Наталия",
            "Наталья",
            "Нина",
            "Нонна",
            "Оксана",
            "Олеся",
            "Ольга",
            "Раиса ",
            "Сара",
            "Светлана",
            "София",
            "Софья",
            "Таисия",
            "Тамара",
            "Татьяна",
            "Фекла",
            "Юлия" };

        public string[] femailot = new string[] {
            "Александровна",
            "Алексеевна",
            "Анатольевна",
            "Андреевна",
            "Аркадьевна",
            "Ашотовна",
            "Борисовна",
            "Вадимовна",
            "Васильевна",
            "Викторовна",
            "Витальевна",
            "Владимировна",
            "Вячеславовна",
            "Геннадьевна",
            "Георгиевна",
            "Григорьевна",
            "Дмитриевна",
            "Евгеньевна",
            "Енбоковна",
            "Енсековна",
            "Ивановна",
            "Игоревна",
            "Ильинична",
            "Константиновна",
            "Леонидовна",
            "Михайловна",
            "Николаевна",
            "Николаевна",
            "Петровна",
            "Рафизовна",
            "Салиховна",
            "Сергеевна",
            "Федоровна",
            "Эдуардовна",
            "Юрьевна" };

        #endregion

        // ----- --------------------------------------------------------------------

        public int dolz = 0, zvan = 0, uch = 0, kaf = 0;
        public int stud_id = 0, grupa_id = 0, status_id;
        public string fakult_str = "", sql = "";
        DataTable grupa_set, status_set;                
        public bool first = true, photochanged = false;
        bool famsaved = false, imsaved = false, otsaved = false;



        private int GetPosById(DataTable box, int id)
        {
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

        public void fill_status()
        {
            //загрузить статус и получить текущий
            status_set = new DataTable();
            string selcom = "select id, text from student_status ";                
            main.global_adapter = new SqlDataAdapter(selcom,main.global_connection);            
            main.global_adapter.Fill(status_set);

            int i = 0;
            int num = 0;
            foreach (DataRow dr in status_set.Rows)
            {
                if (status_id == Convert.ToInt32(dr[0])) num = i;

                status_list.Items.Add(dr[1].ToString());

                i++;
            }

            status_list.SelectedIndex = num;
        }

        public void fill_grupa()
        {
            //загрузить группы, получить активную
            string selcom = "select grupa.id, grupa.name, kurs_id, specialnost.srok " +
                " from grupa " +
                " join specialnost on grupa.specialnost_id = specialnost.id " +
                " where actual=1 and fakultet_id = " + main.fakultet_id.ToString();

            main.global_adapter = new SqlDataAdapter(selcom,
                main.global_connection);

            grupa_set = new DataTable();

            main.global_adapter.Fill(grupa_set);


            grupa_list.Items.Clear();
            foreach (DataRow dr in grupa_set.Rows)
            {
                grupa_list.Items.Add(dr[1]);
            }

            if (grupa_id == 0)
                grupa_list.SelectedIndex = 0;
            else
                grupa_list.SelectedIndex = GetPosById(grupa_set, grupa_id);
        }


        private void student_edit_Load(object sender, EventArgs e)
        {
            if (male.Checked)
            {
                fam.AutoCompleteCustomSource.AddRange(mailfam);
                im.AutoCompleteCustomSource.AddRange(mailim);
                ot.AutoCompleteCustomSource.AddRange(mailot);
            }

            if (female.Checked)
            {
                fam.AutoCompleteCustomSource.AddRange(femailfam);
                im.AutoCompleteCustomSource.AddRange(femailim);
                ot.AutoCompleteCustomSource.AddRange(femailot);
            }

            fill_grupa();
            fill_status();
            fill_address();
            

            //создание студента
            if (newstud)
            {
                byte[] photo = (byte[])(new SqlCommand("SELECT photo FROM student where id = 997",
                            main.global_connection).ExecuteScalar());
                pictureBox1.Image = new Bitmap(new MemoryStream(photo));

                main.global_command = new SqlCommand();
                main.global_command.CommandText = "update student set " +
                    " photo = @p where id = @id";
                main.global_command.Connection = main.global_connection;
                main.global_command.Parameters.Add("@p", SqlDbType.Image, photo.Length).Value = photo;
                main.global_command.Parameters.Add("@id", SqlDbType.Int).Value = stud_id;
                main.global_command.ExecuteNonQuery();

            }
            else //редактирование студента
            {
                update_born_place();
                update_passport();
                update_sem_pol();
                update_status();
                pictureBox1.Image = new Bitmap(new MemoryStream((byte[])(new SqlCommand(
                            "SELECT photo FROM student where id = " + stud_id,
                            main.global_connection)).ExecuteScalar()));
            }
            
            born_date.MinDate = DateTime.Now.AddYears(-50);
            born_date.MaxDate = DateTime.Now.AddYears(-17);            
        }


        public Bitmap GetPhotoFromBD()
        {
            return new Bitmap(new MemoryStream((byte[])(new SqlCommand(
                "SELECT photo FROM student where id = " + stud_id,
                main.global_connection)).ExecuteScalar()));
        }

        DataRow addr_row;
        private void fill_address()
        {
            sql = "SELECT student.id, region.name, city.name, street.name," + // 0 1 2 3
                " isnull(student.live_house_corpus, '-') as dom, " +  // 4
                " isnull(student.live_kv,0) as kva, " +   // 5
                " region.id, city.id, street.id " +   // 6 7 8
                " FROM student " + 
                " JOIN region ON student.live_region_id = region.id " + 
                " JOIN city ON student.live_city_id = city.id " + 
                " JOIN street ON student.live_street_id = street.id " + 
                " where  student.id = " + stud_id.ToString();

            DataTable t = new DataTable();
            (new SqlDataAdapter(sql, main.global_connection)).Fill(t);

            if (t.Rows.Count == 0)
            {
                addres_box.Text = "";
                GC.Collect();
                return;
            }

            addr_row = t.Rows[0];

            string address = addr_row[1].ToString() + ", " +
                addr_row[2].ToString() + ", ул. " +
                addr_row[3].ToString() + ", д." +
                addr_row[4].ToString();
            if (addr_row[5].ToString() != "0")
                address += " кв." + addr_row[5].ToString();

            addres_box.Text = address;  
        }

        private void male_CheckedChanged(object sender, EventArgs e)
        {
            fam.AutoCompleteCustomSource.Clear();
            im.AutoCompleteCustomSource.Clear();
            ot.AutoCompleteCustomSource.Clear();
            
            fam.AutoCompleteCustomSource.AddRange(mailfam);
            im.AutoCompleteCustomSource.AddRange(mailim);
            ot.AutoCompleteCustomSource.AddRange(mailot);

            sql = "update student set sex = 1 where id=@id";
            main.global_command = new SqlCommand(sql, main.global_connection);
            main.global_command.Parameters.Add("@id", SqlDbType.Int).Value = stud_id;
            main.global_command.ExecuteNonQuery();

            groupBox3.Visible = true;
        }

        private void female_CheckedChanged(object sender, EventArgs e)
        {
            fam.AutoCompleteCustomSource.Clear();
            im.AutoCompleteCustomSource.Clear();
            ot.AutoCompleteCustomSource.Clear();
            
            fam.AutoCompleteCustomSource.AddRange(femailfam);
            im.AutoCompleteCustomSource.AddRange(femailim);
            ot.AutoCompleteCustomSource.AddRange(femailot);

            sql = "update student set sex = 0 where id=@id";
            main.global_command = new SqlCommand(sql, main.global_connection);
            main.global_command.Parameters.Add("@id", SqlDbType.Int).Value = stud_id;
            main.global_command.ExecuteNonQuery();

            groupBox3.Visible = false;
        }

        //выбрать фаилию
        private void button1_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("id");
            dt.Columns.Add("name");

            string[] fams = new string[] { "" };

            if (male.Checked) fams = mailfam;
            if (female.Checked) fams = femailfam;

            int i = 0;
            foreach (string nm in fams)
            {
                object[] parms = new object[] { i, nm };
                dt.Rows.Add(parms);
                i++;
            }

            ListWindow lw = new ListWindow();
            lw.tbl = dt;
            lw.Text = "Выбор фамилии";

            DialogResult res = lw.ShowDialog();

            if (res == DialogResult.Cancel) return;

            fam.Text = dt.Rows[lw.resId][1].ToString();
            string famss = NormalizeLetters(fam.Text.Trim());
            fam.Text = famss;
            save_field("fam", "'" + famss + "'");

            lw.Dispose();
        }

        //задать имя
        private void button2_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("id");
            dt.Columns.Add("name");

            string[] fams = new string[] { "" };

            if (male.Checked) fams = mailim;
            if (female.Checked) fams = femailim;

            int i = 0;
            foreach (string nm in fams)
            {
                object[] parms = new object[] { i, nm };
                dt.Rows.Add(parms);
                i++;
            }

            ListWindow lw = new ListWindow();
            lw.tbl = dt;
            lw.Text = "Выбор имени";

            DialogResult res = lw.ShowDialog();

            if (res == DialogResult.Cancel) return;

            im.Text = dt.Rows[lw.resId][1].ToString();
            string ims = NormalizeLetters(im.Text.Trim());
            im.Text = ims;
            save_field("im", "'" + ims + "'");

            lw.Dispose();
        }

        //задать отчество
        private void button3_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("id");
            dt.Columns.Add("name");

            string[] fams = new string[] { "" };

            if (male.Checked) fams = mailot;
            if (female.Checked) fams = femailot;

            int i = 0;
            foreach (string nm in fams)
            {
                object[] parms = new object[] { i, nm };
                dt.Rows.Add(parms);
                i++;
            }

            ListWindow lw = new ListWindow();
            lw.tbl = dt;
            lw.Text = "Выбор отчества";

            DialogResult res = lw.ShowDialog();

            if (res == DialogResult.Cancel) return;

            ot.Text = dt.Rows[lw.resId][1].ToString();
            string ots = NormalizeLetters(ot.Text.Trim());
            ot.Text = ots;
            save_field("ot", "'" + ots + "'");

            lw.Dispose();
        }

        private void grupa_list_SelectedIndexChanged(object sender, EventArgs e)
        {
            grupa_id = (int)grupa_set.Rows[grupa_list.SelectedIndex][0];

            int kurs = (int)grupa_set.Rows[grupa_list.SelectedIndex][2];

            int current_year = main.starts[0].Year; //год начала учебного года
            int y = current_year - kurs + 1; //номер года организации группы

            if (stud_id==0)
            zach.Text = "*-" + fakult_str + "-" +
                  "1" + "-" + y.ToString().Substring(2, 2);

            save_field("gr_id", grupa_id.ToString());
        }

        private void button7_Click(object sender, EventArgs e)
        {
            fam.Text = NormalizeLetters(fam.Text);
            im.Text = NormalizeLetters(im.Text);
            ot.Text = NormalizeLetters(ot.Text);
            
            if (fam.Text.Trim().Length == 0)
            {
                MessageBox.Show("Введите фамилию студента",
                    "Недостаточно данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                fam.Focus();
                fam.Select();                
                return;
            }

            if (im.Text.Trim().Length == 0)
            {
                MessageBox.Show("Введите имя студента",
                    "Недостаточно данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                im.Focus();
                im.Select();               
                return;
            }

            if (zach.Text.Trim().Length == 0)
            {
                MessageBox.Show("Введите номер зачётной книжки студента",
                    "Недостаточно данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                zach.Select();
                zach.Focus();
                return;
            }

            string emailtemplate = @"^[0-9]+([0-9]|\-{1})*[0-9]+$";

            MatchCollection searchres;

            if (phone.Text.Trim().Length > 0)
            {
                searchres = Regex.Matches(phone.Text.Trim(), emailtemplate,
                    RegexOptions.IgnoreCase);
                if (searchres.Count == 0)
                {
                    MessageBox.Show("Вы ввели недопустимы номер телефона. Повторите ввод.",
                        "Недостаточно данных",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    phone.Select();
                    phone.Focus();
                    phone.SelectAll();
                    return;
                }
            }

            if (email.Text.Trim().Length > 0)
            {
                searchres = Regex.Matches(email.Text.Trim(), emailtemplate, RegexOptions.IgnoreCase);
                if (searchres.Count == 0)
                {
                    MessageBox.Show("Вы ввели недопустимый номер сотового телефона. Повторите ввод.",
                        "Недостаточно данных",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    email.Select();
                    email.Focus();
                    email.SelectAll();
                    return;
                }
            }
       

            DialogResult = DialogResult.OK;
        }

        public string NormalizeLetters(string str)
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

        //задать параметры паспорта
        private void button5_Click(object sender, EventArgs e)
        {         
            student_edit_passport sep = new student_edit_passport();

            if (pass_row != null)
            {
                sep.nomer.Text = pass_row[0].ToString();
                sep.seria.Text = pass_row[1].ToString();
                sep.vydano.Text = pass_row[3].ToString();
                sep.дата_выдачи.Value = Convert.ToDateTime(pass_row[2]);
            }
            
            sep.stud_id = stud_id;

            sep.ShowDialog();

            if (sep.DialogResult == DialogResult.Cancel)
            {
                sep.Dispose();
                return;
            }

            string res = "";
            res = "серия: " + sep.seria.Text + ", номер: " + sep.nomer.Text.Trim() +
                ", выдан: " + sep.vydano.Text + " " +
                sep.дата_выдачи.Value.ToShortDateString();
            
            passport_box.Text = res;
            sep.Dispose();
        }

        //задать место рождения
        private void born_place__button_Click(object sender, EventArgs e)
        {
            student_edit_born_place sebp = new student_edit_born_place();
            sebp.student_id = stud_id;
            sebp.ShowDialog();
            update_born_place();
            sebp.Dispose();
        }

        //вывести название региона и города
        private void update_born_place()
        {
            string sql = "select br = isnull(region.name,'-'), " +
	            " ctype = dbo.naspunkt_type.name , " +
	            " bc = isnull(city.name,'-')" +
                " from student " + 
                " join region on region.id = student.born_region_id " + 
                " join city on city.id = student.born_city_id and city.region_id = region.id " + 
                " join naspunkt_type on naspunkt_type.id = city.naspunkt_type_id " + 
                " where student.id = @sid";

            main.global_command = new SqlCommand(sql, main.global_connection);
            main.global_command.Parameters.Add("@sid", SqlDbType.Int).Value = stud_id;

            main.global_adapter = new SqlDataAdapter(main.global_command);
            DataTable dtb = new DataTable();
            main.global_adapter.Fill(dtb);

            if (dtb.Rows.Count == 0)
            {
                born_place_box.Text = "место рождения не указано";
                return;
            }

            string res = string.Empty;

            if (dtb.Rows[0][0].ToString() == "-") res = "место рождения не указано";

            res = dtb.Rows[0][0].ToString() + ", " + dtb.Rows[0][1].ToString() + " " +
                dtb.Rows[0][2].ToString();
            born_place_box.Text = res;
            toolTip1.SetToolTip(born_place_box, res);

            main.global_adapter.Dispose();
            main.global_command.Dispose();
        }

        //вывести паспортные данные
        public DataRow pass_row = null;
        private void update_passport()
        {
            if (stud_id == 0)
            {
                passport_box.Text = "наспортные данные не указаны";
            }
            string sql = "select isnull(passport_nomer,''), isnull(passport_seria,''), " + 
                " isnull(passport_date,''), isnull(passport_vydan,getdate()-1) " + 
                " from student " +                 
                " where student.id = " + stud_id.ToString();

            main.global_adapter = new SqlDataAdapter(sql, main.global_connection);
            DataTable dtb = new DataTable();
            main.global_adapter.Fill(dtb);

            if (dtb.Rows.Count == 0)
            {
                passport_box.Text = "наспортные данные не указаны";
                return;
            }

            if (dtb.Rows[0][0].ToString().Length == 0)
            {
                passport_box.Text = "наспортные данные не указаны";
                return;
            }

            string res = string.Empty;

            pass_row = dtb.Rows[0];

            res = "серия: " + pass_row[1].ToString() + ", номер: " + pass_row[0].ToString() +
                ", выдан: " + pass_row[3].ToString() + " " +
                Convert.ToDateTime(pass_row[2]).ToShortDateString();

            passport_box.Text = res;
            toolTip1.SetToolTip(passport_box, res);
        }

        public void update_status()
        {
            if (stud_id == 0)
            {
                return;
            }

            string sql = "select status_id from student where student.id = " + stud_id.ToString();            
            main.global_adapter = new SqlDataAdapter(sql, main.global_connection);
            DataTable dtb = new DataTable();
            main.global_adapter.Fill(dtb);

            if (dtb.Rows.Count == 0)
            {
                status_id = 1;
                return;
            }
            else
                status_id = Convert.ToInt32(dtb.Rows[0][0]);

            for (int i = 0; i < status_set.Rows.Count; i++)
            {
                if (status_set.Rows[i][0].ToString() == status_id.ToString())
                {
                    status_list.SelectedIndex = i;                    
                    break;
                }
            }                        
        }

        /// <summary>
        /// сохранить значени в указанное поле (поле имеет строковый тип)
        /// </summary>
        /// <param name="fname">имя поля</param>
        /// <param name="fvalue">новое значение поля</param>
        public bool save_field_str(string fname, string fvalue)
        {
            string sql = "update student set " + fname + " = @val where id = @st";
            
            main.global_command = new SqlCommand(sql, main.global_connection);
            main.global_command.Parameters.Add("@val", SqlDbType.NVarChar).Value = fvalue;
            main.global_command.Parameters.Add("@st", SqlDbType.NVarChar).Value = stud_id.ToString();

            bool res = true;

            try
            {
                main.global_command.ExecuteNonQuery();
            }
            catch (Exception exx)
            {
                res = false;
            }

            main.global_command.Dispose();
            return res;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            //задать адрес студента
            student_edit_address sea = new student_edit_address();

            if (addres_box.Text.Trim().Length > 0)
            {
                sea.region_id = "1";

                sea.gorod_id = addr_row[7].ToString();
                sea.nas_punkt_box.Text = addr_row[2].ToString();

                sea.ul_id = addr_row[8].ToString();
                sea.street.Text = addr_row[3].ToString();

                sea.house.Text = addr_row[4].ToString();
                
                sea.Kvartira.Value = Convert.ToInt32(addr_row[5]);
            }

            if (sea.ShowDialog() == DialogResult.Cancel) return;

            save_field("live_street_id", sea.ul_id);
            save_field("live_region_id", "1");
            save_field("live_city_id", sea.gorod_id);
            save_field("live_kv", sea.Kvartira.Value.ToString());
            save_field_str("live_house_corpus", sea.house.Text);

            fill_address();
        }

        private void fam_Enter(object sender, EventArgs e)
        {
            ((Control)sender).BackColor = Color.Yellow;
        }

        private void fam_Leave(object sender, EventArgs e)
        {
            ((Control)sender).BackColor = Color.White;
            string fams = NormalizeLetters(fam.Text.Trim());
            fam.Text = fams;            
            save_field("fam", "'" + fams + "'");
        }

        private void im_Leave(object sender, EventArgs e)
        {
            ((Control)sender).BackColor = Color.White;
            string ims = NormalizeLetters(im.Text.Trim());
            im.Text = ims;
            save_field("im", "'" + ims + "'");
        }

        private void ot_Leave(object sender, EventArgs e)
        {
            ((Control)sender).BackColor = Color.White;
            string ots = NormalizeLetters(ot.Text.Trim());
            ot.Text = ots;
            save_field("ot", "'" + ots + "'");
        }

        private void born_place_box_Enter(object sender, EventArgs e)
        {
            ((Control)sender).BackColor = Color.LightYellow;
        }

        private void born_place_box_Leave(object sender, EventArgs e)
        {
            ((Control)sender).BackColor = Color.White;
        }

        private void zach_kn_group_box_Enter(object sender, EventArgs e)
        {
            //
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            //
        }

        private void update_sem_pol()
        {
            string sql = "select isnull(sem_pologen,0) from student where id = " + stud_id.ToString();
            DataTable sem = new DataTable();
            main.global_adapter = new SqlDataAdapter(sql,main.global_connection);
            main.global_adapter.Fill(sem);

            if (sem.Rows.Count == 0) return;

            bool sempol = Convert.ToBoolean(sem.Rows[0][0]);

            if (!sempol) radioButton1.Checked = true; else radioButton2.Checked = true;
        }

        private void radioButton2_Click(object sender, EventArgs e)
        {
            bool res = false;
            if (radioButton1.Checked)
            {
                res = false;
            }
            else
            {
                res = true;
            }

            string sql = "update student set sem_pologen = @sp where id = @sid";
            main.global_command = new SqlCommand(sql,main.global_connection);
            main.global_command.Parameters.Add("@sp", SqlDbType.Bit).Value = res;
            main.global_command.Parameters.Add("@sid", SqlDbType.Int).Value = stud_id;
            main.global_command.ExecuteNonQuery();
        }
        
        /// <summary>
        /// сохранить изменения в таблице студент для указанного поля (для строковых значений апострофы передавать)
        /// </summary>
        /// <param name="fname">название поля, которое надо изменить</param>
        /// <param name="fvalue">новое значение поля</param>
        /// <returns></returns>
        public bool save_field(string fname, string fvalue)
        {
            string sql = "update student set " + fname + " = " + fvalue + " where id = " + stud_id.ToString();
            main.global_command = new SqlCommand(sql, main.global_connection);

            bool res = true;

            try
            {
                main.global_command.ExecuteNonQuery();
            }
            catch (Exception exx)
            {
                res = false;
            }

            main.global_command.Dispose();
            return res;
        }

        public bool save_fields(string[] fnames, string[] fvalues)
        {
            string sql = "update student set ";

            for (int i = 0; i < fnames.Length; i++)
            {
                sql = sql + fnames[i] + " = " + fvalues[i] + ", ";
            }

            sql = sql + " where id = " + stud_id.ToString();
            main.global_command = new SqlCommand(sql, main.global_connection);

            bool res = true;

            try
            {
                main.global_command.ExecuteNonQuery();
            }
            catch (Exception exx)
            {
                res = false;
            }

            main.global_command.Dispose();
            return res;
        }

        private void button7_Click_1(object sender, EventArgs e)
        {
            delete = false;
            //пол меняется автоматически

            //проверить, есть ли студент с такими ФИО в этой или другой группе и показать его статус
            //сделать тоже самое при добавлении групп, преподавателей, предметов


            //сохранить фам им от
            //проверить что ввели фам им от
            Control[] check = new Control[] { fam, im, ot };
            Label[] labs = new Label[] { famlabel, im_label, ot_label };
            string ress = "";
            int i = 0;

            for(i=0; i<check.Length; i++)            
            {
                if (check[i].Text.Trim().Length == 0)
                {
                    if (ress.Length == 0)
                    {
                        ress = "Сохранение невозможно! Задайте значения указанных полей!\n\n";
                    }

                    ress += " -" + labs[i].Text + "\n";
                }
            }

            string sql = string.Format("select grupa.name, student_status.text, born_date " +
                " from student " +
                " join grupa on grupa.id = student.gr_id " +
                " join student_status on student_status.id = student.status_id " +
                " where grupa.fakultet_id = {0} and " +
                " fam like '{2}' and im like '{3}' and ot like '{4}'",
                main.fakultet_id, grupa_id, fam.Text.Trim(), im.Text.Trim(), ot.Text.Trim());
            DataTable studentcheck = new DataTable();
            (new SqlDataAdapter(sql,main.global_connection)).Fill(studentcheck);

            string ress2 = string.Empty;

            for (i = 0; i < studentcheck.Rows.Count; i++)
            {
                if (ress2 == string.Empty)
                    ress2 = "Внимание!!\n Найдены студенты с такими же фамилией, именем и отчеством.\n" +
                        "Возможно не следует добавлять нового студента в эту группу?\nВ этом случае следует изменить группу или статутс одного " +
                        " из показанных далее студентов (для этого нажмите кнопку отмена и откройте окно редактирования соотвествующего студента).\n";

                ress2 += string.Format("{0}) {1} {2} {3} в группе {4} (статус - {5}).", i + 1, 
                    fam.Text.Trim(), im.Text.Trim(), ot.Text.Trim(),
                    studentcheck.Rows[i][0], studentcheck.Rows[i][0]) + "\n\n";

            }

            if (ress2 != string.Empty)
                ress2 += "\n\nЕсли Вы уверены, что нужно добавить этого студента, то нажмите кнопку Ок.";

            if (ress.Length == 0)
            {
                save_field("start_date", enter_date_box.Value);
                save_field("end_date", end_date_box.Value);
                DialogResult = DialogResult.OK;
            }
            else
            {
                MessageBox.Show(ress,
                    "Ошибка редактирования данных", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        public bool save_field(string fname, DateTime fvalue)
        {
            string sql = "update student set " + fname + " = @dt where id = @st_id";
            main.global_command = new SqlCommand(sql, main.global_connection);
            main.global_command.Parameters.Add("@dt", SqlDbType.DateTime).Value = fvalue;
            main.global_command.Parameters.Add("@st_id", SqlDbType.Int).Value = stud_id;

            bool res = true;

            try
            {
                main.global_command.ExecuteNonQuery();
            }
            catch (Exception exx)
            {
                res = false;
            }

            main.global_command.Dispose();
            return res;
        }

        private void born_date_ValueChanged(object sender, EventArgs e)
        {
            //if (born_date.Value < DateTime.Now.AddYears(-17))
            save_field("born_date", born_date.Value);
            /*else
                MessageBox.Show("Выбрана недопустимая дата рождения","Ошибка ввода",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);*/
        }


        bool delete = true;
        private void student_edit_FormClosing(object sender, FormClosingEventArgs e)
        {

            if (delete)
            {
                // удалить только что добавленного студента
                sql = "delete from student where id = " + stud_id;
                SqlCommand cmd = new SqlCommand(sql, main.global_connection);
                cmd.ExecuteNonQuery();
            }
            else
            {
                //DialogResult = DialogResult.OK;
            }
        }

        private void status_list_Enter(object sender, EventArgs e)
        {
            ((Control)sender).BackColor = Color.LightYellow;
        }

        private void status_list_Leave(object sender, EventArgs e)
        {
            ((Control)sender).BackColor = Color.White;
        }

        private void status_list_SelectedIndexChanged(object sender, EventArgs e)
        {            
            save_field("status_id", status_set.Rows[status_list.SelectedIndex][0].ToString());

            if (status_set.Rows[status_list.SelectedIndex][0].ToString() == "1")
            {
                save_field("actual", "1");
            }
            else
            {
                save_field("actual", "0");
            }
        }

        public bool newstud = false;
        private void button8_Click(object sender, EventArgs e)
        {
            delete = false;
            if (newstud)
            {
                // удалить только что добавленного студента
                sql = "delete from student where id = " + stud_id;
                SqlCommand cmd = new SqlCommand(sql, main.global_connection);
                cmd.ExecuteNonQuery();
                newstud = false;
            }

            DialogResult = DialogResult.Cancel;
        }

        //сохранить место работы
        private void work_place_box_Leave(object sender, EventArgs e)
        {
            string res = work_place_box.Text.Trim().Replace("'","").Trim();
            if (res.Length == 0) return;
            save_field_str("work_place", res);
            work_place_box.BackColor = Color.White;
        }

        private void graduated_from_box_Leave(object sender, EventArgs e)
        {
            string res = graduated_from_box.Text.Trim().Replace("'", "").Trim();
            if (res.Length == 0) return;
            save_field_str("graduated_from", res);
            graduated_from_box.BackColor = Color.White;
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            DateTime val = dateTimePicker1.Value;
            save_field("graduated_date", val);
        }

        private void mother_box_Leave(object sender, EventArgs e)
        {
            string res = mother_box.Text.Trim().Replace("'", "").Trim();
            if (res.Length == 0) return;
            save_field_str("mother_info", res);
            mother_box.BackColor = Color.White;                       
        }

        private void father_box_Leave(object sender, EventArgs e)
        {

            string res = father_box.Text.Trim().Replace("'", "").Trim();
            if (res.Length == 0) return;
            save_field_str("father_info", res);
            father_box.BackColor = Color.White;
        }

        private void phone_Leave(object sender, EventArgs e)
        {
            string res = phone.Text.Trim().Replace("'", "").Trim();
            if (res.Length == 0)
            {
                phone.BackColor = Color.White;
                return;
            }
            
            string emailtemplate = @"^[0-9]+([0-9]|\-{1})*[0-9]+$";

            MatchCollection searchres;
            searchres = Regex.Matches(res, emailtemplate, RegexOptions.IgnoreCase);
            if (searchres.Count == 0)
            {
                MessageBox.Show("Вы ввели недопустимый текст  (не номер телефона). Повторите ввод.",
                    "Недостаточно данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                phone.Select();
                phone.Focus();
                phone.SelectAll();
                return;
            }

            save_field_str("phone", res);
            phone.BackColor = Color.White;
        }

        private void email_Leave(object sender, EventArgs e)
        {

            string res = email.Text.Trim().Replace("'", "").Trim();
            if (res.Length == 0)
            {
                email.BackColor = Color.White;
                return;
            }

            string emailtemplate = @"^[0-9]+([0-9]|\-{1})*[0-9]+$";

            MatchCollection searchres;
            searchres = Regex.Matches(res, emailtemplate, RegexOptions.IgnoreCase);
            if (searchres.Count == 0)
            {
                MessageBox.Show("Вы ввели недопустимый текст (это не номер телефона). Повторите ввод.",
                    "Недостаточно данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                email.Select();
                email.Focus();
                email.SelectAll();
                return;
            }

            save_field_str("cell_phone", res);
            email.BackColor = Color.White;
        }

        private void zach_Leave(object sender, EventArgs e)
        {            
            string res = zach.Text.Trim().Replace("'", "").Trim();
            if (res.Length == 0)
            {
                zach.BackColor = Color.White;
                return;
            }

            save_field_str("zach_kn_number", res);
            zach.BackColor = Color.White;
        }

        private void prikaz_box_Leave(object sender, EventArgs e)
        {
            //номер приказа о зачислении
            string res = prikaz_box.Text.Trim().Replace("'", "").Trim();
            if (res.Length == 0)
            {
                prikaz_box.BackColor = Color.White;
                return;
            }

            bool ress = save_field_str("prikaz_nom_zach", res);
            prikaz_box.BackColor = Color.White;
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            
            bool rrr;
            if (radioButton4.Checked)
                rrr = save_field("military_id", "1");
            else
                rrr = save_field("military_id", "0");
            //MessageBox.Show(rrr.ToString());
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            //if (radioButton3.Checked)
              //  save_field_str("military_id", "0");
        }

        /// <summary>
        /// получить массив пикселей фотографии
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static byte[] GetPhotoFromFile(string filePath)
        {
            FileStream stream = new FileStream(
                filePath, FileMode.Open, FileAccess.Read);
            BinaryReader reader = new BinaryReader(stream);

            byte[] photo = reader.ReadBytes((int)stream.Length);

            reader.Close();           
            stream.Close();
            stream.Dispose();

            return photo;
        }

        public string FilePhoto = "";

        //сохранить фото в БД
        private void button4_Click(object sender, EventArgs e)
        {
            DialogResult dr = openFileDialog1.ShowDialog();
            if (dr == DialogResult.Cancel) return;

            FilePhoto = openFileDialog1.FileName;
            Bitmap bmp = new Bitmap(FilePhoto);                      

            if (bmp.Width > 250 || bmp.Height > 250)
            {
                MessageBox.Show("Загружаемое изображение слишком велико " + 
                    "(допустимый размер не более 200 пикселей по каждому измерению).",
                    "Отказ операции", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            pictureBox1.Image = bmp;            
            byte[] photo = GetPhotoFromFile(FilePhoto);

            main.global_command = new SqlCommand();
            main.global_command.CommandText = "update student set " +
                " photo = @p where id = @id";
            main.global_command.Connection = main.global_connection;

            main.global_command.Parameters.Add("@p",SqlDbType.Image, photo.Length).Value = photo;
            main.global_command.Parameters.Add("@id", SqlDbType.Int).Value = stud_id;

            try
            {
                main.global_command.ExecuteNonQuery();
            }
            catch (Exception exx)
            {
                MessageBox.Show(exx.Message);
            }
        }

        private void textBox1_Leave(object sender, EventArgs e)
        {
            //адрес электронной почты

        }

        private void textBox2_Leave(object sender, EventArgs e)
        {
            //номер личного дела
        }

        private void enter_date_box_Leave(object sender, EventArgs e)
        {
            save_field("start_date", enter_date_box.Value);            
        }

        private void end_date_box_Leave(object sender, EventArgs e)
        {
            save_field("end_date", end_date_box.Value);
        }

        private void копироватьИзображениеВБуферToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (pictureBox1.Image != null)
            {
                Clipboard.SetImage(pictureBox1.Image);
            }
        }
    }
}