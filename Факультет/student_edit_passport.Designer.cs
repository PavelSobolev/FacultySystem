namespace FSystem
{
    partial class student_edit_passport
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(student_edit_passport));
            this.дата_выдачи = new System.Windows.Forms.DateTimePicker();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.nomer = new System.Windows.Forms.TextBox();
            this.button8 = new System.Windows.Forms.Button();
            this.button7 = new System.Windows.Forms.Button();
            this.vydano = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.seria = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // дата_выдачи
            // 
            this.дата_выдачи.CalendarMonthBackground = System.Drawing.SystemColors.Info;
            this.дата_выдачи.Location = new System.Drawing.Point(12, 21);
            this.дата_выдачи.Name = "дата_выдачи";
            this.дата_выдачи.Size = new System.Drawing.Size(200, 20);
            this.дата_выдачи.TabIndex = 0;
            this.дата_выдачи.Value = new System.DateTime(2003, 10, 1, 0, 0, 0, 0);
            this.дата_выдачи.Enter += new System.EventHandler(this.dateTimePicker1_Enter);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.Color.Transparent;
            this.label1.Location = new System.Drawing.Point(9, 5);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(123, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Дата выдачи паспорта";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BackColor = System.Drawing.Color.Transparent;
            this.label2.Location = new System.Drawing.Point(12, 48);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(38, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "Серия";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.BackColor = System.Drawing.Color.Transparent;
            this.label3.Location = new System.Drawing.Point(215, 48);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(41, 13);
            this.label3.TabIndex = 3;
            this.label3.Text = "Номер";
            // 
            // nomer
            // 
            this.nomer.Location = new System.Drawing.Point(218, 64);
            this.nomer.MaxLength = 6;
            this.nomer.Name = "nomer";
            this.nomer.Size = new System.Drawing.Size(200, 20);
            this.nomer.TabIndex = 5;
            this.nomer.Leave += new System.EventHandler(this.textBox1_Leave);
            this.nomer.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.nomer_KeyPress);
            this.nomer.Enter += new System.EventHandler(this.textBox1_Enter);
            // 
            // button8
            // 
            this.button8.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button8.Image = global::FSystem.Properties.Resources.delete_x16_h;
            this.button8.Location = new System.Drawing.Point(120, 136);
            this.button8.Name = "button8";
            this.button8.Size = new System.Drawing.Size(99, 29);
            this.button8.TabIndex = 20;
            this.button8.Text = "отмена";
            this.button8.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.button8.UseVisualStyleBackColor = true;
            this.button8.Click += new System.EventHandler(this.button8_Click);
            // 
            // button7
            // 
            this.button7.Image = global::FSystem.Properties.Resources.ok;
            this.button7.Location = new System.Drawing.Point(12, 136);
            this.button7.Name = "button7";
            this.button7.Size = new System.Drawing.Size(99, 29);
            this.button7.TabIndex = 19;
            this.button7.Text = "принять";
            this.button7.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.button7.UseVisualStyleBackColor = true;
            this.button7.Click += new System.EventHandler(this.button7_Click);
            // 
            // vydano
            // 
            this.vydano.AutoCompleteCustomSource.AddRange(new string[] {
            "УВД г. Южно-Сахалинска",
            "УВД г. Холмска",
            "УВД г. Охи",
            "УВД г. Невельска",
            "УВД г. Поронайска",
            "УВД г. Корсакова",
            "УВД г. Анивы",
            "УВД г. Александровска-Сахалинского"});
            this.vydano.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.vydano.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.CustomSource;
            this.vydano.Location = new System.Drawing.Point(12, 108);
            this.vydano.Name = "vydano";
            this.vydano.Size = new System.Drawing.Size(406, 20);
            this.vydano.TabIndex = 11;
            this.vydano.Text = "УВД г. Южно-Сахалинска";
            this.vydano.Leave += new System.EventHandler(this.textBox1_Leave);
            this.vydano.Enter += new System.EventHandler(this.textBox1_Enter);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.BackColor = System.Drawing.Color.Transparent;
            this.label4.Location = new System.Drawing.Point(12, 92);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(63, 13);
            this.label4.TabIndex = 22;
            this.label4.Text = "Кем выдан";
            // 
            // seria
            // 
            this.seria.Location = new System.Drawing.Point(12, 64);
            this.seria.MaxLength = 5;
            this.seria.Name = "seria";
            this.seria.Size = new System.Drawing.Size(200, 20);
            this.seria.TabIndex = 3;
            this.seria.Leave += new System.EventHandler(this.textBox1_Leave);
            this.seria.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBox1_KeyPress);
            this.seria.Enter += new System.EventHandler(this.textBox1_Enter);
            // 
            // student_edit_passport
            // 
            this.AcceptButton = this.button7;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = global::FSystem.Properties.Resources.tab_fon;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.CancelButton = this.button8;
            this.ClientSize = new System.Drawing.Size(427, 175);
            this.Controls.Add(this.seria);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.vydano);
            this.Controls.Add(this.button8);
            this.Controls.Add(this.button7);
            this.Controls.Add(this.nomer);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.дата_выдачи);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "student_edit_passport";
            this.Opacity = 0.9;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "паспортные данные студента";
            this.Load += new System.EventHandler(this.student_edit_passport_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        public System.Windows.Forms.Button button8;
        public System.Windows.Forms.Button button7;
        public System.Windows.Forms.DateTimePicker дата_выдачи;
        public System.Windows.Forms.Label label1;
        public System.Windows.Forms.Label label2;
        public System.Windows.Forms.Label label3;
        public System.Windows.Forms.TextBox nomer;
        public System.Windows.Forms.TextBox vydano;
        public System.Windows.Forms.Label label4;
        public System.Windows.Forms.TextBox seria;
    }
}