using System;
using System.Windows.Forms;

namespace FSystem
{
    public partial class enter_as_dekan : Form
    {
        public enter_as_dekan ( )
        {
            InitializeComponent ( );
        }

        public bool status = false;

        private void button1_Click ( object sender, EventArgs e )
        {
            if ( textBox2.Text.Trim ( ) == main.df )
            {
                status = true;
            }
            else
            {
                status = false;
            }
        }

        private void button2_Click ( object sender, EventArgs e )
        {
            status = false;
        }

        private void textBox2_KeyDown ( object sender, KeyEventArgs e )
        {
            if ( textBox2.Text.Trim ( ).Length > 0 )
            {
                if ( e.KeyCode == Keys.Return )
                {
                    if ( textBox2.Text.Trim ( ) == main.df )
                    {
                        status = true;
                    }
                    else
                    {
                        status = false;
                    }

                    DialogResult = DialogResult.OK;
                }
            }
        }

        private void enter_as_dekan_KeyDown ( object sender, KeyEventArgs e )
        {
            if ( e.KeyCode == Keys.Escape )
            {
                status = false;
                DialogResult = DialogResult.Cancel;
            }
        }

        private void enter_as_dekan_Load(object sender, EventArgs e)
        {
            label4.Text = InputLanguage.CurrentInputLanguage.Culture.IetfLanguageTag.Substring(0, 2);
        }

        private void enter_as_dekan_InputLanguageChanged(object sender, InputLanguageChangedEventArgs e)
        {
            label4.Text = InputLanguage.CurrentInputLanguage.Culture.IetfLanguageTag.Substring(0, 2);
        }


    }
}