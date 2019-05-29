using System;
using System.Windows.Forms;

namespace FSystem
{
    public partial class table_context : Form
    {
        public table_context()
        {
            InitializeComponent();
        }

        private int posx=0, posy=0;

        private void table_context_Deactivate(object sender, EventArgs e)
        {
            // to be removed
        }

        private void table_context_MouseDown(object sender, MouseEventArgs e)
        {
            posx = e.X;
            posy = e.Y;
        }

        private void table_context_MouseMove(object sender, MouseEventArgs e)
        {
              int dx = posx - e.X;
              int dy = posy - e.Y;
              if (e.Button==MouseButtons.Left)
              {
                 Left -= dx;
                 Top -= dy;
              }
        }
    }
}