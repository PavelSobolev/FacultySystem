using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace FSystem
{
    public partial class rasp_nastroika : Form
    {
        public rasp_nastroika()
        {
            InitializeComponent();
        }

        public SqlCommand com = new SqlCommand();
        public DataSet ds = new DataSet();
        SqlDataAdapter sda;

        
        public bool grups_changed = false;        // данные для групп
        public Dictionary<int, bool> status = new Dictionary<int, bool>();
        public Dictionary<int, int> order = new Dictionary<int, int>();
        public Dictionary<string, int> gr_id = new Dictionary<string, int>();


        private void rasp_nastroika_Load(object sender, EventArgs e)
        {
            
            com.Connection = main.global_connection;
            
            //заполнить первую закладку (группы)
            string command1 = "select id, name from grupa where " +
                " actual = 1 and fakultet_id = " + main.fakultet_id.ToString() +
                " order by outorder ";

            string command2 = "select id, name from grupa where " +
                " actual = 1 and fakultet_id = " + main.fakultet_id.ToString() +
                " and show_in_grid = 1 order by outorder ";

            try
            {

                sda = new SqlDataAdapter(command1 + ";" + command2,
                    main.global_connection);
                sda.Fill(ds);
                int k = 1;

                for (int i = 0; i < 2; i++)
                {
                    for (int j = 0; j < ds.Tables[i].Rows.Count; j++)
                    {
                        if (i == 0)
                        {
                            listBox1.Items.Add(ds.Tables[i].Rows[j][1]);
                            status.Add((int)ds.Tables[i].Rows[j][0], false);
                            order.Add((int)ds.Tables[i].Rows[j][0], 0);
                            gr_id.Add(ds.Tables[i].Rows[j][1].ToString(),
                                (int)ds.Tables[i].Rows[j][0]);
                        }
                        else
                        {
                            listBox2.Items.Add(ds.Tables[i].Rows[j][1]);
                            status[(int)ds.Tables[i].Rows[j][0]] = true;
                            order[(int)ds.Tables[i].Rows[j][0]] = k;
                            k++;
                        }
                    }
                }
            }
            catch (Exception exx)
            {
                //
                MessageBox.Show("При извлечении данных произошла ошибка." + 
                    "\nПопробуйте вызвать данную команду позднее.","Ошибка работы с данными",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                DialogResult = DialogResult.Cancel;
            }

            listBox1.SelectedIndex = 0;
            listBox2.SelectedIndex = 0;
            listBox1.Focus();

            // -----  конец работы с вкладкой групп ----- 
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //добавление в список группы
            string gr = listBox1.Text;

            if (gr.Trim().Length == 0) return;

            if (listBox2.Items.Contains((object)gr))
            {
                MessageBox.Show("Данная группа уже добавлена." ,"Ошибка работы с данными",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            listBox2.Items.Insert(listBox2.SelectedIndex + 1, listBox1.Text);

            status[gr_id[gr]] = true; //группа будет выводиться
            order[gr_id[gr]] = listBox2.SelectedIndex + 2;

            if (listBox1.SelectedIndex != listBox1.Items.Count - 1)
                listBox1.SelectedIndex++;

            grups_changed = true;

        }

        private void button5_Click(object sender, EventArgs e)
        {
            listBox2.Focus();
            
            //перемесить группу вверх
            int num = listBox2.SelectedIndex;

            if (num == -1) return;

            string txt = (string)listBox2.Items[num];

            if (num == 0) return;

            listBox2.Items.RemoveAt(num);
            listBox2.Items.Insert(num - 1, txt);

            order[gr_id[txt]] = num;            

            listBox2.SelectedIndex = num - 1;

            grups_changed = true;

        }

        private void button6_Click(object sender, EventArgs e)
        {
            listBox2.Focus();

            //перемесить группу вниз
            int num = listBox2.SelectedIndex;

            if (num == -1) return;

            string txt = (string)listBox2.Items[num];

            if (num == listBox2.Items.Count - 1) return;

            listBox2.Items.RemoveAt(num);
            listBox2.Items.Insert(num + 1, txt);

            order[gr_id[txt]] = num + 2;

            listBox2.SelectedIndex = num + 1;

            grups_changed = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //удалить группу из списка
            
            string gr = listBox2.Text;

            if (gr.Trim().Length == 0) return;

            if (listBox2.Items.Count == 1)
            {
                MessageBox.Show("В раписании должна быть, по крайней мере, одна группа.", 
                    "Невозможно выполнить команду",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            status[gr_id[gr]] = false; //группа не будет выводиться
            order[gr_id[gr]] = 0;

            int del = listBox2.SelectedIndex;

            listBox2.Items.RemoveAt(del);

            if (del < listBox2.Items.Count - 1) listBox2.SelectedIndex = del;

            if (del == listBox2.Items.Count - 1)
                listBox2.SelectedIndex = 0;         

            listBox2.Focus();

            grups_changed = true;

        }

        private void button3_Click(object sender, EventArgs e)
        {
            // to be removed
        }
    }
}