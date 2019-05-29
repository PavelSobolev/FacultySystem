namespace FSystem
{
    partial class set_attend_dialog
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(set_attend_dialog));
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.pris_label = new System.Windows.Forms.ToolStripStatusLabel();
            this.ots_label = new System.Windows.Forms.ToolStripStatusLabel();
            this.attend_grid = new System.Windows.Forms.DataGridView();
            this.number = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.fio = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.attend_name = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.заётыИОценкиToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.выставитьОценкуToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.выставитьЗачётToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.выйтиToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.statusStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.attend_grid)).BeginInit();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // statusStrip1
            // 
            this.statusStrip1.BackColor = System.Drawing.Color.Transparent;
            this.statusStrip1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.pris_label,
            this.ots_label});
            this.statusStrip1.Location = new System.Drawing.Point(0, 659);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Padding = new System.Windows.Forms.Padding(1, 0, 16, 0);
            this.statusStrip1.Size = new System.Drawing.Size(688, 22);
            this.statusStrip1.SizingGrip = false;
            this.statusStrip1.TabIndex = 0;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // pris_label
            // 
            this.pris_label.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.pris_label.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.pris_label.Image = global::FSystem.Properties.Resources.smallsuccess;
            this.pris_label.ImageTransparentColor = System.Drawing.Color.White;
            this.pris_label.Name = "pris_label";
            this.pris_label.Size = new System.Drawing.Size(146, 17);
            this.pris_label.Text = "toolStripStatusLabel1";
            // 
            // ots_label
            // 
            this.ots_label.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.ots_label.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.ots_label.Image = global::FSystem.Properties.Resources.smallfail;
            this.ots_label.ImageTransparentColor = System.Drawing.Color.White;
            this.ots_label.Name = "ots_label";
            this.ots_label.Size = new System.Drawing.Size(146, 17);
            this.ots_label.Text = "toolStripStatusLabel2";
            // 
            // attend_grid
            // 
            this.attend_grid.AllowUserToAddRows = false;
            this.attend_grid.AllowUserToDeleteRows = false;
            this.attend_grid.AllowUserToOrderColumns = true;
            this.attend_grid.AllowUserToResizeColumns = false;
            this.attend_grid.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.attend_grid.BackgroundColor = System.Drawing.Color.White;
            this.attend_grid.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.attend_grid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.attend_grid.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.number,
            this.fio,
            this.attend_name});
            this.attend_grid.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter;
            this.attend_grid.Location = new System.Drawing.Point(10, 114);
            this.attend_grid.Name = "attend_grid";
            this.attend_grid.Size = new System.Drawing.Size(666, 530);
            this.attend_grid.TabIndex = 1;
            this.attend_grid.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.attend_grid_CellDoubleClick);
            this.attend_grid.KeyDown += new System.Windows.Forms.KeyEventHandler(this.attend_grid_KeyDown);
            this.attend_grid.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.attend_grid_KeyPress);
            // 
            // number
            // 
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            this.number.DefaultCellStyle = dataGridViewCellStyle3;
            this.number.HeaderText = "№";
            this.number.Name = "number";
            this.number.ReadOnly = true;
            this.number.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.number.Width = 30;
            // 
            // fio
            // 
            this.fio.HeaderText = "ФИО студента";
            this.fio.Name = "fio";
            this.fio.ReadOnly = true;
            this.fio.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.fio.Width = 270;
            // 
            // attend_name
            // 
            dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            this.attend_name.DefaultCellStyle = dataGridViewCellStyle4;
            this.attend_name.HeaderText = "Присутствие";
            this.attend_name.Name = "attend_name";
            this.attend_name.ReadOnly = true;
            this.attend_name.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.attend_name.Width = 110;
            // 
            // textBox1
            // 
            this.textBox1.BackColor = System.Drawing.Color.WhiteSmoke;
            this.textBox1.Location = new System.Drawing.Point(10, 55);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.ReadOnly = true;
            this.textBox1.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.textBox1.Size = new System.Drawing.Size(451, 53);
            this.textBox1.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.Color.Transparent;
            this.label1.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.Location = new System.Drawing.Point(7, 38);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(90, 14);
            this.label1.TabIndex = 4;
            this.label1.Text = "Тема занятия";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BackColor = System.Drawing.Color.Transparent;
            this.label2.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label2.Location = new System.Drawing.Point(464, 38);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(83, 14);
            this.label2.TabIndex = 6;
            this.label2.Text = "Примечания";
            // 
            // textBox2
            // 
            this.textBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox2.BackColor = System.Drawing.Color.WhiteSmoke;
            this.textBox2.Location = new System.Drawing.Point(467, 55);
            this.textBox2.Multiline = true;
            this.textBox2.Name = "textBox2";
            this.textBox2.ReadOnly = true;
            this.textBox2.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.textBox2.Size = new System.Drawing.Size(209, 53);
            this.textBox2.TabIndex = 5;
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.заётыИОценкиToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1045, 24);
            this.menuStrip1.TabIndex = 7;
            this.menuStrip1.Text = "menuStrip1";
            this.menuStrip1.Visible = false;
            // 
            // заётыИОценкиToolStripMenuItem
            // 
            this.заётыИОценкиToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.выставитьОценкуToolStripMenuItem,
            this.выставитьЗачётToolStripMenuItem,
            this.toolStripSeparator1,
            this.выйтиToolStripMenuItem});
            this.заётыИОценкиToolStripMenuItem.Image = global::FSystem.Properties.Resources.fo;
            this.заётыИОценкиToolStripMenuItem.Name = "заётыИОценкиToolStripMenuItem";
            this.заётыИОценкиToolStripMenuItem.Size = new System.Drawing.Size(115, 20);
            this.заётыИОценкиToolStripMenuItem.Text = "Заёты и оценки";
            // 
            // выставитьОценкуToolStripMenuItem
            // 
            this.выставитьОценкуToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("выставитьОценкуToolStripMenuItem.Image")));
            this.выставитьОценкуToolStripMenuItem.Name = "выставитьОценкуToolStripMenuItem";
            this.выставитьОценкуToolStripMenuItem.Size = new System.Drawing.Size(179, 22);
            this.выставитьОценкуToolStripMenuItem.Text = "Выставить оценку";
            this.выставитьОценкуToolStripMenuItem.Click += new System.EventHandler(this.выставитьОценкуToolStripMenuItem_Click);
            // 
            // выставитьЗачётToolStripMenuItem
            // 
            this.выставитьЗачётToolStripMenuItem.Image = global::FSystem.Properties.Resources.ok;
            this.выставитьЗачётToolStripMenuItem.Name = "выставитьЗачётToolStripMenuItem";
            this.выставитьЗачётToolStripMenuItem.Size = new System.Drawing.Size(179, 22);
            this.выставитьЗачётToolStripMenuItem.Text = "Выставить зачёт";
            this.выставитьЗачётToolStripMenuItem.Click += new System.EventHandler(this.выставитьЗачётToolStripMenuItem_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(176, 6);
            // 
            // выйтиToolStripMenuItem
            // 
            this.выйтиToolStripMenuItem.Name = "выйтиToolStripMenuItem";
            this.выйтиToolStripMenuItem.Size = new System.Drawing.Size(179, 22);
            this.выйтиToolStripMenuItem.Text = "Выйти";
            // 
            // set_attend_dialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 14F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = global::FSystem.Properties.Resources.tab_fon;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.ClientSize = new System.Drawing.Size(688, 681);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.attend_grid);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.menuStrip1);
            this.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.HelpButton = true;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "set_attend_dialog";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "отметка посещаемости";
            this.Load += new System.EventHandler(this.set_attend_dialog_Load);
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.set_attend_dialog_FormClosing);
            this.HelpRequested += new System.Windows.Forms.HelpEventHandler(this.set_attend_dialog_HelpRequested);
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.attend_grid)).EndInit();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel pris_label;
        private System.Windows.Forms.DataGridView attend_grid;
        private System.Windows.Forms.DataGridViewTextBoxColumn number;
        private System.Windows.Forms.DataGridViewTextBoxColumn fio;
        private System.Windows.Forms.DataGridViewTextBoxColumn attend_name;
        private System.Windows.Forms.ToolStripStatusLabel ots_label;
        public System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        public System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem заётыИОценкиToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem выставитьОценкуToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem выставитьЗачётToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripMenuItem выйтиToolStripMenuItem;
    }
}