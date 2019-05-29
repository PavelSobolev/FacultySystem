using System;
using System.Drawing;
using System.Windows.Forms;

namespace FSystem
{
    public partial class inputbox : Form
    {

        public string explanation_text;
        public bool is_numeric = false; 

        /// <summary>
        /// вывод окна ввода текста или числа (с калькулятором)
        /// </summary>
        /// <param name="explan">справка по вводу (объяснение что надо вводить)</param>
        /// <param name="caption">заголовок окна</param>
        /// <param name="txt">текст в поле ввода</param>
        /// <param name="labele">надпись над строкой ввода</param>
        public inputbox ( string explan, string caption, string txt, string labele)
        {
            InitializeComponent ( );
            Text = caption;
            explanation_text = explan;
            textBox1.Value = (object)txt;
                        
            label1.Text = labele;

            textBox1.Focus();
            textBox1.Select();
        }

        private void explanation_Click ( object sender, EventArgs e )
        {
            MessageBox.Show ( explanation_text, "Пояснение", 
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (is_numeric)
            {
                string txt = textBox1.Text;

                if (!main.IsNumber(ref txt))
                {
                    MessageBox.Show("Введено нечисловое значание.",
                        "Ошибка ввода",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    
                    textBox1.Value = txt;
                    textBox1.ForeColor = Color.Red;
                    DialogResult = DialogResult.Abort;
                }
                else
                {
                    textBox1.Value = txt;
                    textBox1.ForeColor = Color.Black;
                    DialogResult = DialogResult.OK;
                }
            }
            else
                DialogResult = DialogResult.OK;
        }

        private void inputbox_Load(object sender, EventArgs e)
        {
            // to be removed
        }
    }
}