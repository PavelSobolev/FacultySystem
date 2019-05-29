using System;
using System.Data;
using System.Data.SqlClient;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace FSystem
{
    public partial class prepod_edit : Form
    {
        public prepod_edit()
        {
            InitializeComponent();
        }

        #region fams_ims_ots
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

        public int dolz = 0, zvan = 0, uch = 0, kaf = 0;
        public int dolz_id = 0, zvan_id = 0, uch_id = 0, kaf_id = 0, zavkaf_id = 0;
        DataTable dolz_table, zvan_table, uch_table, kaf_table;
        public int prep_id = 0;
        public bool first = true, photochanged = false, is_zav = false, zav_changed = false;

        private void button1_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("id");
            dt.Columns.Add("name");

            string[] fams = new string[] { "" };

            if (male.Checked)   fams = mailfam;
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

            lw.Dispose();
        }

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

            lw.Dispose();
        }

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

            lw.Dispose();
        }

        private void mail_CheckedChanged(object sender, EventArgs e)
        {
            if (prep_id == 0)
            {
                fam.Clear();
                im.Clear();
                ot.Clear();
            }
            
            fam.AutoCompleteCustomSource.Clear();
            im.AutoCompleteCustomSource.Clear();
            ot.AutoCompleteCustomSource.Clear();

            fam.AutoCompleteCustomSource.AddRange(mailfam);
            im.AutoCompleteCustomSource.AddRange(mailim);
            ot.AutoCompleteCustomSource.AddRange(mailot);
        }

        private void femail_CheckedChanged(object sender, EventArgs e)
        {
            if (prep_id == 0)
            {
                fam.Clear();
                im.Clear();
                ot.Clear();
            }

            fam.AutoCompleteCustomSource.Clear();
            im.AutoCompleteCustomSource.Clear();
            ot.AutoCompleteCustomSource.Clear();

            fam.AutoCompleteCustomSource.AddRange(femailfam);
            im.AutoCompleteCustomSource.AddRange(femailim);
            ot.AutoCompleteCustomSource.AddRange(femailot);
        }

        private void prepod_edit_Load(object sender, EventArgs e)
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

            //получить должности
            string cmd = "select id, name from dolznost order by name";
            main.global_adapter = new SqlDataAdapter(cmd,
                main.global_connection);
            dolz_table = new DataTable();
            main.global_adapter.Fill(dolz_table);
            foreach (DataRow dr in dolz_table.Rows)
                dolz_list.Items.Add(dr[1].ToString());

            int itempos = 0;

            dolz_list.SelectedIndex = 0;

            if (prep_id != 0)                           
            {
                itempos = GetPosById(dolz_table, dolz_id);
                if (itempos>=0)
                    dolz_list.SelectedIndex = itempos;
            }

            //получить звания
            cmd = "select id, name from zvanie  order by name";
            main.global_adapter = new SqlDataAdapter(cmd,
                main.global_connection);
            zvan_table = new DataTable();
            main.global_adapter.Fill(zvan_table);
            foreach (DataRow dr in zvan_table.Rows)
                zvan_list.Items.Add(dr[1].ToString());

            zvan_list.SelectedIndex = 0;
            if (prep_id != 0)
            {
                itempos = GetPosById(zvan_table, zvan_id);
                if (itempos >= 0)
                    zvan_list.SelectedIndex = itempos;
            }

            //получить степени
            cmd = "select id, name from stepen  order by name";
            main.global_adapter = new SqlDataAdapter(cmd,
                main.global_connection);
            uch_table = new DataTable();
            main.global_adapter.Fill(uch_table);
            foreach (DataRow dr in uch_table.Rows)
                uch_step_list.Items.Add(dr[1].ToString());

            uch_step_list.SelectedIndex = 0;
            if (prep_id != 0)
            {
                itempos = GetPosById(uch_table, uch_id);
                if (itempos >= 0)
                    uch_step_list.SelectedIndex = itempos;
            }

            //получить степени
            cmd = "select id, name from kafedra " + 
                " where actual=1 order by priority";
            main.global_adapter = new SqlDataAdapter(cmd,
                main.global_connection);
            kaf_table = new DataTable();
            main.global_adapter.Fill(kaf_table);
            foreach (DataRow dr in kaf_table.Rows)
                kaf_list.Items.Add(dr[1].ToString());

            kaf_list.SelectedIndex = 0;
            if (prep_id != 0)
            {
                itempos = GetPosById(kaf_table, kaf_id);
                if (itempos >= 0)
                    kaf_list.SelectedIndex = itempos;
            }

            first = false;

            if (dolz_list.Text.ToLower().Trim().Contains("заведующий"))
            {
                is_zav = true;
            }

        }

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

        public bool deny_photo = false;
        public static string DefaultPath = Environment.GetFolderPath(
            Environment.SpecialFolder.MyPictures) +
            "\\emptyface.jpg";

        private void button7_Click(object sender, EventArgs e)
        {

            fam.Text = main.NormalizeLetters(fam.Text);
            im.Text = main.NormalizeLetters(im.Text);
            ot.Text = main.NormalizeLetters(ot.Text);
            
            
            bool fe = true, ie = true;
            bool cont = true;            

            if (fam.Text.Trim() != string.Empty) fe = false;
            if (im.Text.Trim() != string.Empty) ie = false;
            

            string res = "Обнаружены следующие ошибки данных:\n\n";
            if (fe) res += " - не введена фамилия [ввод обязателен]\n";
            if (ie) res += " - не введено имя [ввод обязателен]\n";

            if (fe || ie)
            {
                MessageBox.Show(res + "\nИсправьте ошибки и повторите операцию.", "Отклонение запроса",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (fe) fam.Select();
                if (ie) im.Select();
                return;
            }

            //проверить, не добавлчется ли еще один зав каф или замваз каф
            //int zavid = (int)sprav_prepods.;

            if (zavkaf_id != 69 && dolz_list.Text.ToLower().Trim().Contains("заведующий"))
            {
                if (zavkaf_id != prep_id) //разрешить поменять данные у текущего зав каФ
                {
                    MessageBox.Show("Вы пытаетесь присвоить данному преподавателю должность " +
                        "\"заведующий кафедрой\", в то время как данная должность уже занята.",
                        "Запрос отколнён",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            //проверка размеров фото
            if (pictureBox1.Image.Width > 300 && pictureBox1.Image.Height > 300)
            {
                MessageBox.Show("Загруженная Вами фотография слишком велика (допустимый размер не более 300*300 пикселей.)\n" +
                    "Изображение не будет сохранено в базу данных.", "ОШибка данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                deny_photo = true;
                pictureBox1.ImageLocation = DefaultPath;
                FilePhoto = DefaultPath;
                return;
            }
            else
                deny_photo = false;

            //проверить совпадение пола и имени
            bool male_enterd = fam.Text.EndsWith("ов") || fam.Text.EndsWith("ев") ||
                fam.Text.EndsWith("ив") ||
                fam.Text.EndsWith("ин") || fam.Text.EndsWith("ий") ||
                fam.Text.EndsWith("ой") || ot.Text.EndsWith("ич") ||
                (Array.IndexOf(mailfam, fam.Text)!=-1) ||
                (Array.IndexOf(mailim, im.Text)!=-1) ||
                (Array.IndexOf(mailot, ot.Text)!=-1);

            bool  female_enterd = fam.Text.EndsWith("ва") ||
                fam.Text.EndsWith("ина") || fam.Text.EndsWith("ая") ||
                ot.Text.EndsWith("на") ||
                (im.Text[im.Text.Length - 1] == 'a' && im.Text.Trim().ToLower()!="никита")  || 
                im.Text[im.Text.Length - 1] == 'я' ||
                (Array.IndexOf(femailfam, fam.Text)!=-1) ||
                (Array.IndexOf(femailim, im.Text)!=-1) ||
                (Array.IndexOf(femailot, ot.Text)!=-1);

            if ((male_enterd && female.Checked)||(female_enterd && male.Checked))
            {                
                res = " - вероятно введенное имя, фамилия или отчество" + 
                    " не соотвествуют\nвыбранному полу [данное замечание можно проигнорировать]\n\n";
                cont = false;
            }


            bool tel = false, em = false;

            if (phone.Text.Trim() == string.Empty)
            {
                res += " - не введен номер телефона [данное замечание можно проигнорировать]\n\n";
                cont = false;
            }
            else
            {
                string emailtemplate = @"^[0-9]+([0-9]|\-{1})*[0-9]+$";
                MatchCollection searchres = Regex.Matches(phone.Text.Trim(), emailtemplate,
                    RegexOptions.IgnoreCase);
                if (searchres.Count == 0)
                {
                    res += " - введен некорректный номер телефона [исправьте]\n\n";
                    tel = true;
                    cont = false;
                }
            }

            if (email.Text.Trim() == string.Empty)
            {
                res += " - не введен электронный адрес [данное замечание можно проигнорировать]\n\n";
                cont = false;
            }
            else
            {
                string emailtemplate = @".+@.+\..{2}";
                MatchCollection searchres = Regex.Matches(email.Text.Trim(), emailtemplate,
                    RegexOptions.IgnoreCase);
                if (searchres.Count == 0)
                {
                    res += " - введен некорректный электронный адрес [исправьте]\n\n";
                    em = true;
                    cont = false;
                }
            }

            if (em || tel)
            {
                MessageBox.Show(res + "\nИсправьте ошибки и повторите операцию.", "Отклонение запроса",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (tel) phone.Select();
                if (em) email.Select();
                return;
            }

            DialogResult lres = DialogResult.OK;

            if (!dolz_list.Text.ToLower().Trim().Contains("заведующий"))
            {
                if (is_zav)
                zav_changed = true;
            }


            if (cont==false)
            {
                lres = MessageBox.Show(res + "\nПродолжить операцию?", "Выбор действия",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (lres==DialogResult.Yes)
                    DialogResult = DialogResult.OK;
            }
            else
                DialogResult = DialogResult.OK;
        }

        private void status_box_CheckedChanged(object sender, EventArgs e)
        {
            if (status_box.Checked)
                status_box.Text = "Статус: работает";
            else
                status_box.Text = "Статус: уволен"; 
        }

        private void button4_Click(object sender, EventArgs e)
        {
            DialogResult dr = openFileDialog1.ShowDialog();
            if (dr == DialogResult.Cancel) return;

            FilePhoto = openFileDialog1.FileName;
            pictureBox1.ImageLocation = FilePhoto;
            photochanged = true;
        }

        public string FilePhoto = "";

        private void kaf_list_SelectedIndexChanged(object sender, EventArgs e)
        {            
            int pos = kaf_list.SelectedIndex;
            if (!first)
            kaf_id = (int)kaf_table.Rows[pos][0];
        }

        private void dolz_list_SelectedIndexChanged(object sender, EventArgs e)
        {
            int pos = dolz_list.SelectedIndex;
            if (!first)
            dolz_id = (int)dolz_table.Rows[pos][0];             
        }

        private void uch_step_list_SelectedIndexChanged(object sender, EventArgs e)
        {
            int pos = uch_step_list.SelectedIndex;
            if (!first)
            uch_id = (int)uch_table.Rows[pos][0];
        }

        private void zvan_list_SelectedIndexChanged(object sender, EventArgs e)
        {
            int pos = zvan_list.SelectedIndex;
            if (!first)
            zvan_id = (int)zvan_table.Rows[pos][0];
        }

        private void sex_Enter(object sender, EventArgs e)
        {

        }

    }
}