using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace FSystem
{
    public partial class set_tema_from_indrasp : Form
    {
        public set_tema_from_indrasp()
        {
            InitializeComponent();
        }

        //внешние данные
        public int zan_id = 0, predm_id = 0;
        public string tema = "";
        public DateTime enddate = DateTime.Now;

        //для работы с sql
        private SqlCommand cmd = new SqlCommand();
        public DataTable tema_set = null, attend_set = null;

        private void set_tema_from_indrasp_Load(object sender, EventArgs e)
        {
            //действия при загрузке и отображении окна

            //заполнить список тем на текущую дату по этому предмету
            string zapros = string.Format(
                "select дата = dbo.get_date(y,m,d), [вид занятия] = vid_zan.name, " + 
                " тема = case " + 
	            " when isnull(tema,'-')='-' then 'не задана' " +
	            " when len(tema)=0 then 'не задана' " +
	            " else tema " +
                " end " +
                " from rasp " +
                " join vid_zan on vid_zan.id = rasp.vid_zan_id " +
                " where dbo.get_date(y,m,d)>=dbo.get_date({3},{4},{5}) and dbo.get_date(y,m,d)<=dbo.get_date({0},{1},{2}) " +
	            "    and predmet_id = " + predm_id.ToString() + 
                " order by дата desc", enddate.Year, enddate.Month, enddate.Day, 
                    main.year_start.Year, main.year_start.Month, main.year_start.Day);

            main.global_adapter = new System.Data.SqlClient.SqlDataAdapter(zapros,
                    main.global_connection);

            tema_set = new DataTable();
            main.global_adapter.Fill(tema_set);
            tema_table.DataSource = tema_set;
            tema_table.Columns[0].Width = 100;          
            tema_table.Columns[1].Width = 120;            
            tema_table.Columns[2].Width = 330;            

            //поставить тему в текстБокс если она есть
            temaBox.Text = tema;

            //выделить в решетке то занятие которое было выделено в решетке инд расписания
            // ---------- возможна реализация в будущем ----------------------------------

        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (button3.Text.Contains("показать"))
            {
                button3.Text = "скрыть темы";
                Height = 398;
            }
            else
            {
                button3.Text = "показать все темы";
                Height = 209;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //сохранить тему
            if (temaBox.Text.Trim().Length == 0)
            {
                MessageBox.Show("Введите, пожалуйста, тему!","Тема не введена",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            temaBox.Text = main.Normalize1(temaBox.Text);

            string zapros = "update rasp set tema = '" + temaBox.Text + "'" +
                " where id = " + zan_id.ToString();
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

            //повторно заполнить сетку с темами
            //заполнить список тем на текущую дату по этому предмету
            zapros = string.Format(
                "select дата = dbo.get_date(y,m,d), [вид занятия] = vid_zan.name, " +
                " тема = case " +
                " when isnull(tema,'-')='-' then 'не задана' " +
                " when len(tema)=0 then 'не задана' " +
                " else tema " +
                " end " +
                " from rasp " +
                " join vid_zan on vid_zan.id = rasp.vid_zan_id " +
                " where dbo.get_date(y,m,d)>=dbo.get_date({3},{4},{5}) and dbo.get_date(y,m,d)<=dbo.get_date({0},{1},{2}) " +
                "    and predmet_id = " + predm_id.ToString() +
                " order by дата desc", enddate.Year, enddate.Month, enddate.Day,
                    main.year_start.Year, main.year_start.Month, main.year_start.Day);

            main.global_adapter = new System.Data.SqlClient.SqlDataAdapter(zapros,
                    main.global_connection);

            tema_set = new DataTable();
            main.global_adapter.Fill(tema_set);
            tema_table.DataSource = tema_set;
            tema_table.Columns[0].Width = 100;
            tema_table.Columns[1].Width = 120;
            tema_table.Columns[2].Width = 330;
            DialogResult = DialogResult.OK;
        }
        
    }
}