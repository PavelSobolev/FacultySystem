using System;
using System.Data;
using System.Windows.Forms;

namespace FSystem
{
    public partial class ListWindow : Form
    {
        /// <summary>
        /// выбранный ИД пункта списка
        /// </summary>
        public int resId = 0; 
        /// <summary>
        /// загружаемый список (должен содержать только два поля - ИД  и строковое значение)
        /// </summary>
        public DataTable tbl = new DataTable();

        /// <summary>
        /// строковый результат
        /// </summary>
        public string str_res = "";

        public ListWindow()
        {
            InitializeComponent();
        }

        private void ListWindow_Load(object sender, EventArgs e)
        {
            if (tbl.Rows.Count == 0) return;

            foreach (DataRow dr in tbl.Rows)
            {
                listBox1.Items.Add(dr[1].ToString());
            }

            listBox1.SelectedIndex = 0;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            resId = Convert.ToInt32(tbl.Rows[listBox1.SelectedIndex][0].ToString());
            str_res = listBox1.Text;
            DialogResult = DialogResult.OK;
        }

        private void listBox1_DoubleClick(object sender, EventArgs e)
        {            
            resId = Convert.ToInt32(tbl.Rows[listBox1.SelectedIndex][0].ToString());
            str_res = listBox1.Text;
            DialogResult = DialogResult.OK;
        }


        private void listBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
            {
                resId = Convert.ToInt32(tbl.Rows[listBox1.SelectedIndex][0].ToString());
                str_res = listBox1.Text;
                DialogResult = DialogResult.OK;
            }
        }
    }
}