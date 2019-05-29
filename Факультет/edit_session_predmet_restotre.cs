using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace FSystem
{
    public partial class edit_session_predmet_restotre : Form
    {
        /// <summary>
        /// ид группы
        /// </summary>
        string GrupaID = string.Empty;
        string GrupaName = string.Empty;
        string Kurs = string.Empty;
        string PredmID = string.Empty;
        string VidZanID = string.Empty;

        /// <summary>
        /// конструктор окна 
        /// </summary>
        /// <param name="GrId">иднетификатор группы для которой выполняется восстановление данных</param>
        /// <param name="GrName">название группы для которой выполняется восстановление данных</param>
        /// <param name="Krs">курс которого выполняется восстановление данных</param>
        /// <param name="Vzan">ид вида занятия, для котрого делается восстановление</param>
        public edit_session_predmet_restotre(
            string GrId, string GrName, string Krs, string Vzan)
        {
            InitializeComponent();
            GrupaID = GrId;
            GrupaName = GrName;
            Kurs = Krs;
            VidZanID = Vzan;
            label1.Text = 
                string.Format("Список ранее удалённых предметов сессии для группы {0} (номер курса сессии: {1})",
                GrName, Kurs);
            FIllPredmetGrid();
            FillREsult();
        }

        /// <summary>
        /// таблица отменённых предметов
        /// </summary>
        DataTable OtmPredmets = null;

        void FIllPredmetGrid()
        {
            PredmetGrid.Rows.Clear();

            string vidzanfilter = " and vid_zan.id<=16  and vid_zan.id<>15 ";

            if (VidZanID.Length != 0)
                vidzanfilter = " and vid_zan.id = " + VidZanID;

            string sql = "select distinct predmet.id, predmet.name, vid_zan.name, predmet.semestr " +
                " from predmet " +
                " join grupa on grupa.id = predmet.grupa_id " +
                " join session on session.predmet_id = predmet.id " +
                " join student on student.id = session.student_id " +
                " join vid_zan on vid_zan.id = session.vid_zan_id " +
                " where " +
                " student.gr_id = " + GrupaID + 
                " and predmet.kurs_id =  " + Kurs +
                vidzanfilter +
                " and session.isactual=0 " +
                "order by predmet.name,  predmet.semestr";
            OtmPredmets = new DataTable();
            (new SqlDataAdapter(sql, main.global_connection)).Fill(OtmPredmets);

            if (OtmPredmets.Rows.Count == 0)
            {
                label1.Text = "Нет удаленных предметов для данной группы в сессии за " + Kurs
                    + " курс";
                label1.BackColor = Color.Red;
                button1.Enabled = false;
                PredmetGrid.Enabled = false;
            }

            int i = 0;
            foreach (DataRow dr in OtmPredmets.Rows)
            {
                PredmetGrid.Rows.Add(dr[1], dr[2], dr[3]);
                PredmetGrid.Rows[i].Tag = dr[0];
                i++;
            }

            //textBox1.Text = sql;
        }

        /// <summary>
        /// таблица с результатами сессии по отменённому предмету
        /// </summary>
        DataTable OtmResTable = null;

        void FillREsult()
        {
            if (PredmetGrid.Rows.Count == 0) return;
            if (PredmetGrid.CurrentCell == null) return;

            ResultGrid.Rows.Clear();

            string vidzanfilter = " and vid_zan.id<=16  and vid_zan.id<>15 ";

            if (VidZanID.Length != 0)
                vidzanfilter = " and vid_zan.id = " + VidZanID;

            string sql = "select dbo.GetStudentFIOByID(session.student_id), vid_otmetka.str_name, " +
                " session.id " +
                " from predmet " +
                " join grupa on grupa.id = predmet.grupa_id " +
                " join session on session.predmet_id = predmet.id " +
                " join student on student.id = session.student_id " +
                " join vid_zan on vid_zan.id = session.vid_zan_id " +
                " join vid_otmetka on vid_otmetka.id = session.otmetka_id " +
                " where  " +
                " student.gr_id = " + GrupaID +
                " and predmet.kurs_id = " + Kurs +
                vidzanfilter +
                " and session.isactual=0 " +
                " and predmet.id = " + PredmetGrid.CurrentRow.Tag.ToString() + 
                " order by dbo.GetStudentFIOByID(session.student_id)";
            OtmResTable = new DataTable();
            (new SqlDataAdapter(sql, main.global_connection)).Fill(OtmResTable);
            int i = 0;
            foreach (DataRow dr in OtmResTable.Rows)
            {
                ResultGrid.Rows.Add(dr[0], dr[1]);
                ResultGrid.Rows[i].Tag = dr[2];
                i++;
            }

            label2.Visible = true;
            label2.Text = "Выбран для восстановления предмет:" + 
                PredmetGrid.CurrentRow.Cells[0].Value.ToString();
        }

        private void PredmetGrid_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            FillREsult();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string sql = "update session set isactual = 1 where ";

            int rcount = ResultGrid.Rows.Count;

            if (rcount > 0)
            {
                for (int i = 0; i < rcount; i++)
                {
                    string sess_id = ResultGrid.Rows[i].Tag.ToString();
                    if (i < rcount - 1)
                        sql += " id = " + sess_id + " or ";
                    else
                        sql += " id = " + sess_id;
                }
                
                SqlCommand cmd = new SqlCommand(sql, main.global_connection);
                cmd.ExecuteNonQuery();

                MessageBox.Show("Предмет " + PredmetGrid.CurrentRow.Cells[0].Value.ToString() +
                    " был восстановлен в сетке сессии группы " + GrupaName,
                    "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);

                DialogResult = DialogResult.OK;
            }
        }

        private void edit_session_predmet_restotre_Load(object sender, EventArgs e)
        {
            // to be removed
        }
    }
}
