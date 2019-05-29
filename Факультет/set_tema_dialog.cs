using System;
using System.Drawing;
using System.Windows.Forms;

namespace FSystem
{
    public partial class set_tema_dialog : Form
    {

        public int id_1 = 0, id_2 = 0;
        
        public set_tema_dialog()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
        }

        private void tema1_MouseDown(object sender, MouseEventArgs e)
        {
            tema1.BackColor = Color.LightGray;
            tema2.BackColor = Color.White;
        }

        private void tema2_MouseDown(object sender, MouseEventArgs e)
        {
            tema2.BackColor = Color.LightGray;
            tema1.BackColor = Color.White;
        }

        private void tema1_TextChanged(object sender, EventArgs e)
        {
            if (choose_cynchro.Checked)
            {
                string txt = tema1.Text;
                int sel = tema1.SelectionStart;
                tema2.Text = txt;
                tema1.SelectionStart = sel;// + 1;
            }
        }

        private void tema2_TextChanged(object sender, EventArgs e)
        {
            if (choose_cynchro.Checked)
            {
                string txt = tema2.Text;
                int sel = tema2.SelectionStart;
                tema1.Text = txt;
                tema2.SelectionStart = sel;// +1;
            }
        }

        private void set_tema_dialog_Load(object sender, EventArgs e)
        {
            tema1.SelectionStart = tema1.Text.Length;
            tema2.SelectionStart = tema2.Text.Length;
        }
    }
}