namespace FSystem
{
    partial class grupa_edit
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(grupa_edit));
            this.label1 = new System.Windows.Forms.Label();
            this.grupa_name_box = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.spec_list = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.kaf_box = new System.Windows.Forms.TextBox();
            this.exists_box = new System.Windows.Forms.CheckBox();
            this.show_box = new System.Windows.Forms.CheckBox();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.kurs_list = new System.Windows.Forms.NumericUpDown();
            this.label4 = new System.Windows.Forms.Label();
            this.srok_label = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.kurs_list)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.Color.Transparent;
            this.label1.Location = new System.Drawing.Point(9, 96);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(94, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "название группы";
            // 
            // grupa_name_box
            // 
            this.grupa_name_box.Location = new System.Drawing.Point(12, 112);
            this.grupa_name_box.Name = "grupa_name_box";
            this.grupa_name_box.Size = new System.Drawing.Size(227, 20);
            this.grupa_name_box.TabIndex = 1;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BackColor = System.Drawing.Color.Transparent;
            this.label2.Location = new System.Drawing.Point(12, 8);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(84, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "специальность";
            // 
            // spec_list
            // 
            this.spec_list.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.spec_list.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.spec_list.FormattingEnabled = true;
            this.spec_list.Location = new System.Drawing.Point(12, 24);
            this.spec_list.Name = "spec_list";
            this.spec_list.Size = new System.Drawing.Size(347, 21);
            this.spec_list.TabIndex = 2;
            this.spec_list.SelectedIndexChanged += new System.EventHandler(this.spec_list_SelectedIndexChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.BackColor = System.Drawing.Color.Transparent;
            this.label3.Location = new System.Drawing.Point(12, 137);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(126, 13);
            this.label3.TabIndex = 4;
            this.label3.Text = "выпускающая кафедра";
            // 
            // kaf_box
            // 
            this.kaf_box.BackColor = System.Drawing.Color.Azure;
            this.kaf_box.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.kaf_box.Location = new System.Drawing.Point(12, 153);
            this.kaf_box.Name = "kaf_box";
            this.kaf_box.ReadOnly = true;
            this.kaf_box.Size = new System.Drawing.Size(344, 20);
            this.kaf_box.TabIndex = 15;
            // 
            // exists_box
            // 
            this.exists_box.AutoSize = true;
            this.exists_box.BackColor = System.Drawing.Color.Transparent;
            this.exists_box.Location = new System.Drawing.Point(12, 183);
            this.exists_box.Name = "exists_box";
            this.exists_box.Size = new System.Drawing.Size(122, 17);
            this.exists_box.TabIndex = 3;
            this.exists_box.Text = "группа существует";
            this.exists_box.UseVisualStyleBackColor = false;
            this.exists_box.CheckedChanged += new System.EventHandler(this.exists_box_CheckedChanged);
            // 
            // show_box
            // 
            this.show_box.AutoSize = true;
            this.show_box.BackColor = System.Drawing.Color.Transparent;
            this.show_box.Location = new System.Drawing.Point(12, 206);
            this.show_box.Name = "show_box";
            this.show_box.Size = new System.Drawing.Size(234, 17);
            this.show_box.TabIndex = 4;
            this.show_box.Text = "группа выводится в таблице расписания";
            this.show_box.UseVisualStyleBackColor = false;
            this.show_box.CheckedChanged += new System.EventHandler(this.show_box_CheckedChanged);
            // 
            // button1
            // 
            this.button1.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button1.Image = global::FSystem.Properties.Resources.ok;
            this.button1.Location = new System.Drawing.Point(255, 70);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(104, 23);
            this.button1.TabIndex = 5;
            this.button1.Text = "принять";
            this.button1.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button2.Image = global::FSystem.Properties.Resources.delete_x16_h;
            this.button2.Location = new System.Drawing.Point(255, 110);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(104, 23);
            this.button2.TabIndex = 6;
            this.button2.Text = "отменить";
            this.button2.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // kurs_list
            // 
            this.kurs_list.BackColor = System.Drawing.Color.White;
            this.kurs_list.Location = new System.Drawing.Point(12, 70);
            this.kurs_list.Maximum = new decimal(new int[] {
            6,
            0,
            0,
            0});
            this.kurs_list.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.kurs_list.Name = "kurs_list";
            this.kurs_list.ReadOnly = true;
            this.kurs_list.Size = new System.Drawing.Size(227, 20);
            this.kurs_list.TabIndex = 0;
            this.kurs_list.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.kurs_list.ValueChanged += new System.EventHandler(this.kurs_list_ValueChanged);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.BackColor = System.Drawing.Color.Transparent;
            this.label4.Location = new System.Drawing.Point(12, 54);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(71, 13);
            this.label4.TabIndex = 12;
            this.label4.Text = "номер курса";
            // 
            // srok_label
            // 
            this.srok_label.AutoSize = true;
            this.srok_label.BackColor = System.Drawing.Color.Transparent;
            this.srok_label.Location = new System.Drawing.Point(255, 8);
            this.srok_label.Name = "srok_label";
            this.srok_label.Size = new System.Drawing.Size(84, 13);
            this.srok_label.TabIndex = 14;
            this.srok_label.Text = "специальность";
            // 
            // grupa_edit
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = global::FSystem.Properties.Resources.tab_fon;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.CancelButton = this.button2;
            this.ClientSize = new System.Drawing.Size(368, 234);
            this.Controls.Add(this.srok_label);
            this.Controls.Add(this.kurs_list);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.show_box);
            this.Controls.Add(this.exists_box);
            this.Controls.Add(this.kaf_box);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.spec_list);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.grupa_name_box);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "grupa_edit";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Данные группы";
            this.Load += new System.EventHandler(this.grupa_edit_Load);
            ((System.ComponentModel.ISupportInitialize)(this.kurs_list)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        public System.Windows.Forms.Label label1;
        public System.Windows.Forms.TextBox grupa_name_box;
        public System.Windows.Forms.Label label2;
        public System.Windows.Forms.ComboBox spec_list;
        public System.Windows.Forms.Label label3;
        public System.Windows.Forms.TextBox kaf_box;
        public System.Windows.Forms.CheckBox exists_box;
        public System.Windows.Forms.CheckBox show_box;
        public System.Windows.Forms.Button button1;
        public System.Windows.Forms.Button button2;
        public System.Windows.Forms.NumericUpDown kurs_list;
        public System.Windows.Forms.Label label4;
        public System.Windows.Forms.Label srok_label;

    }
}