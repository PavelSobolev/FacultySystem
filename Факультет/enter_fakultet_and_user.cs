using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace FSystem
{
    public partial class enter_fakultet_and_user : Form
    {
        public enter_fakultet_and_user ( )
        {
            InitializeComponent ( );        
            Text = txt;
        }

        public SqlDataAdapter fakult_adapter, prep_adapter;
        public DataSet fakult_dataset, prep_dataset;
        public DataTable uch_god_tbl;
        public bool res = false;
        public string txt = "";        

        /// <summary>
        /// //заполнить списки FSystemов
        /// </summary>
        public void fiil_fakult ( )
        {
            string cmd = "select fakultet.id, fakultet.name, name_krat, nach_zan, korpus_id, " +
                " dekan_id, zam_dekan_id, peremena, long_peremena, semestr2_start, " +
                " first_long_peremena, second_long_peremena," +
                " fio = prepod.fam + ' ' + left(prepod.im,1) + '.' + left(prepod.ot,1) + '.', semestr1_end, " +
                " prefix " +  //14
                " from fakultet " +
                " join prepod on prepod.id = fakultet.dekan_id " +
                " where fakultet.actual = 1 and (prepod.kafedra_id between 10 and 14) order by priority";
            fakult_adapter = new SqlDataAdapter ( cmd, main.global_connection );

            fakult_dataset = new DataSet ( );
            fakult_adapter.Fill ( fakult_dataset );

            cmd = "select * from uch_god where @d>=start-1 and @d<=finish+1";
            SqlCommand Cmd = new SqlCommand(cmd, main.global_connection);
            Cmd.Parameters.Add("@d", SqlDbType.DateTime).Value = DateTime.Today;
            uch_god_tbl = new DataTable();
            fakult_adapter = new SqlDataAdapter(Cmd);
            fakult_adapter.Fill(uch_god_tbl);
            //MessageBox.Show(uch_god_tbl.Rows[0][3].ToString()); //*************

            for (int i = 0; i < fakult_dataset.Tables[0].Rows.Count; i++)
                fakultet_list.Items.Add(fakult_dataset.Tables[0].Rows[i][1]);

            main.uch_god = Convert.ToInt32(uch_god_tbl.Rows[0][0]);
            main.Att1_1 = Convert.ToDateTime(uch_god_tbl.Rows[0][5]);
            main.Att1_2 = Convert.ToDateTime(uch_god_tbl.Rows[0][6]);
            main.Att2_1 = Convert.ToDateTime(uch_god_tbl.Rows[0][7]);
            main.Att2_2 = Convert.ToDateTime(uch_god_tbl.Rows[0][8]);
            //MessageBox.Show(main.uch_god.ToString());
        }

        /// <summary>
        /// заполнить списки преподавателей
        /// </summary>
        public void fill_prepod ( )
        {
            prepod_list.Items.Clear ( );
            int sel = fakultet_list.SelectedIndex;

            string cmd = " select distinct prepod.id , " +  //0
                " fio = fam + ' '  + left(im, 1) + '. ' + left(ot, 1) + '.' " + // 1 
                " , pass, dolznost.name, kafedra.name_krat, dolznost.id  " +  // 2 3 4 5
                " from prepod " +
                " join dolznost on dolznost.id = prepod.dolznost_id " +
                " join kafedra on kafedra.id = prepod.kafedra_id " +
                " join predmet on predmet.prepod_id = prepod.id " +
                " where predmet.fakultet_id = " +
                fakult_dataset.Tables[0].Rows[sel][0].ToString ( ) +
                " and prepod.actual=1 and kafedra.id>=10 and kafedra.id<=13 " +
                " order by fio ";
            prep_adapter = new SqlDataAdapter ( cmd, main.global_connection );
            prep_dataset = new DataSet ( );
            prep_adapter.Fill ( prep_dataset );

            foreach ( DataRow row in prep_dataset.Tables[0].Rows )
                prepod_list.Items.Add ( row[1] );

            prepod_list.SelectedIndex = 0;
        }

        private void button1_Click ( object sender, EventArgs e )
        {
            //обработка ввода
            int selp = prepod_list.SelectedIndex;
            int self = fakultet_list.SelectedIndex;

            if ( prep_dataset.Tables[0].Rows[selp][2].ToString ( ) == pwd.Text )
            {
                res = true;
                main.fakultet_id = (int) fakult_dataset.Tables[0].Rows[self][0];
                main.fakultet_name = fakult_dataset.Tables[0].Rows[self][1].ToString();
                main.fakultet_name_krat = fakult_dataset.Tables[0].Rows[self][2].ToString();

                //MessageBox.Show(uch_god_tbl.Rows[0][1].ToString());

                main.year_start = Convert.ToDateTime(uch_god_tbl.Rows[0][1]);
                main.year_end = Convert.ToDateTime(uch_god_tbl.Rows[0][2]);
                main.semestr2_start = Convert.ToDateTime(uch_god_tbl.Rows[0][4]); //13 fakult_dataset.Tables[0].Rows[self][9]
                main.semestr1_end = Convert.ToDateTime(uch_god_tbl.Rows[0][3]); //fakult_dataset.Tables[0].Rows[self][13]);
                
                //MessageBox.Show(main.semestr1_end.ToLongDateString());

                main.active_user_id = Convert.ToInt32(prep_dataset.Tables[0].Rows[selp][0]);
                main.active_user_name = prep_dataset.Tables[0].Rows[selp][1].ToString();
                main.active_user_kaf = prep_dataset.Tables[0].Rows[selp][4].ToString();
                main.active_user_dolz = prep_dataset.Tables[0].Rows[selp][3].ToString();
                main.start_time = (DateTime) fakult_dataset.Tables[0].Rows[self][3];
                //main.semestr2_start = (DateTime)fakult_dataset.Tables[0].Rows[self][9];
                main.peremena = (int)fakult_dataset.Tables[0].Rows[self][7];
                main.long_peremena = (int)fakult_dataset.Tables[0].Rows[self][8];
                main.first_long_peremena = (int)fakult_dataset.Tables[0].Rows[self][10];
                main.second_long_peremena = (int)fakult_dataset.Tables[0].Rows[self][11];
                main.fakultet_prfix = fakult_dataset.Tables[0].Rows[self][14].ToString();

                int dek_id = (int) fakult_dataset.Tables[0].Rows[self][5];
                int zam_dek_id = (int) fakult_dataset.Tables[0].Rows[self][6];
                main.dekan_name = fakult_dataset.Tables[0].Rows[self][12].ToString();

                int dol_id = Convert.ToInt32( prep_dataset.Tables[0].Rows[selp][5]);
                main.active_user_dolz_id = dol_id;

                if ( main.active_user_id == dek_id ||
                    main.active_user_id == zam_dek_id || dol_id==15)
                {
                    main.df = prep_dataset.Tables[0].Rows[selp][2].ToString ( );
                    main.dekan_online = true;
                    if (main.active_user_id == dek_id)
                        main.user_role = "декан";
                    else
                    {
                        if (main.active_user_id == zam_dek_id)
                            main.user_role = "заместитель декана";
                        else
                            main.user_role = "секретарь деканата";
                    }
                }
                else
                {
                    main.user_role = "преподаватель";
                    string q = String.Format ( "select prepod.pass from fakultet " +
                        " join prepod on prepod.id=fakultet.zam_dekan_id " +
                        " where fakultet.actual=1 and fakultet.id={0}",
                        main.fakultet_id );
                    prep_adapter = new SqlDataAdapter ( q, main.global_connection );
                    DataSet tmp = new DataSet ( );
                    prep_adapter.Fill ( tmp );
                    main.df = (string) tmp.Tables[0].Rows[0][0].ToString ( );
                    
                }
            }
            else
            {
                res = false;
            }
        }

        private void enter_fakultet_and_user_Load ( object sender, EventArgs e )
        {
            //if ( !enter_server.re ) return;

            if ( main.first_enter )
                pictureBox1.Image = Properties.Resources.enter_fak_start;
            else
                pictureBox1.Image = Properties.Resources.enter_fak_second;

            fiil_fakult ( );
            fakultet_list.SelectedIndex = 0;
            main.first_enter = false;

            label4.Text = InputLanguage.CurrentInputLanguage.Culture.IetfLanguageTag.Substring(0, 2);
        }

        /// <summary>
        /// заново заполнить список преподавателей
        /// при изменении FSystemа
        /// </summary>
        private void fakultet_list_SelectedIndexChanged ( object sender, EventArgs e )
        {
            fill_prepod ( );
        }

        private void enter_fakultet_and_user_FormClosed ( object sender, FormClosedEventArgs e )
        {
            fakult_adapter.Dispose ( );
            prep_adapter.Dispose ( );
            fakult_dataset.Dispose ( );
            prep_dataset.Dispose ( );
        }

        private void linkLabel1_LinkClicked ( object sender, LinkLabelLinkClickedEventArgs e )
        {
            MessageBox.Show ( "В списке отображены фамилии всех преподавателей," +
                        "\nведущих дисциплины на выбранном FSystemе.", "Справка",
                        MessageBoxButtons.OK, MessageBoxIcon.Information );
        }

        private void enter_fakultet_and_user_InputLanguageChanged(object sender, InputLanguageChangedEventArgs e)
        {
            label4.Text = InputLanguage
                .CurrentInputLanguage
                .Culture.IetfLanguageTag.Substring(0,2);
        }
    }
}