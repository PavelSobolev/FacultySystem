using System;
using System.Windows.Forms;

namespace FSystem
{
    public partial class wait : Form
    {
        public wait ( )
        {
            InitializeComponent ( );
        }

        private void wait_Load ( object sender, EventArgs e )
        {
            Application.DoEvents();
        }  
    }
}