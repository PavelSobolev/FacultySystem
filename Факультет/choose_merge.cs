using System;
using System.Windows.Forms;

namespace FSystem
{
    public partial class choose_merge : Form
    {
        public choose_merge()
        {
            InitializeComponent();
        }

        public string label_text = "";

        private void choose_merge_Load(object sender, EventArgs e)
        {
            label1.Text = label_text;

            button1.Enabled = true;
            button2.Enabled = true;

            button1.Text = main.left;
            button2.Text = main.right;

            if (button1.Text.Length == 0) button1.Enabled = false;
            if (button2.Text.Length == 0) button2.Enabled = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            main.goaltext = button1.Text;
            main.mergeresult = 0;
            main.unmerge_result = 0;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            main.goaltext = button2.Text;
            main.mergeresult = 1;
            main.unmerge_result = 1;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            main.goaltext = " ";
            main.mergeresult = 2;        
        }

        private void otmena_Click(object sender, EventArgs e)
        {
            main.mergeresult = -1;
            main.unmerge_result = -1;
        }
    }
}