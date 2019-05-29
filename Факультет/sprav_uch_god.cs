using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace Факультет
{
    public partial class sprav_uch_god : Form
    {
        public sprav_uch_god()
        {
            InitializeComponent();
        }

        private void uch_godBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.uch_godBindingSource.EndEdit();
            this.uch_godTableAdapter.Update(this.vKRDataSet.uch_god);

        }

        private void sprav_uch_god_Load(object sender, EventArgs e)
        {
            //global::Факультет.Properties.Settings.Default.VKRConnectionString = main.global_connection.ConnectionString;
            // TODO: This line of code loads data into the 'vKRDataSet.uch_god' table. You can move, or remove it, as needed.
            this.uch_godTableAdapter.Fill(this.vKRDataSet.uch_god);

        }
    }
}