namespace FSystem
{
    partial class student_edit_address
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(student_edit_address));
            this.button8 = new System.Windows.Forms.Button();
            this.button7 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.region_box = new System.Windows.Forms.TextBox();
            this.nas_punkt_box = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.street = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.house = new System.Windows.Forms.TextBox();
            this.Дом = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.born_place__button = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.Kvartira = new System.Windows.Forms.NumericUpDown();
            ((System.ComponentModel.ISupportInitialize)(this.Kvartira)).BeginInit();
            this.SuspendLayout();
            // 
            // button8
            // 
            this.button8.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button8.Image = global::FSystem.Properties.Resources.delete_x16_h;
            this.button8.Location = new System.Drawing.Point(119, 185);
            this.button8.Name = "button8";
            this.button8.Size = new System.Drawing.Size(99, 29);
            this.button8.TabIndex = 36;
            this.button8.Text = "отмена";
            this.button8.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.button8.UseVisualStyleBackColor = true;
            this.button8.Click += new System.EventHandler(this.button8_Click);
            // 
            // button7
            // 
            this.button7.Image = global::FSystem.Properties.Resources.ok;
            this.button7.Location = new System.Drawing.Point(11, 185);
            this.button7.Name = "button7";
            this.button7.Size = new System.Drawing.Size(99, 29);
            this.button7.TabIndex = 35;
            this.button7.Text = "принять";
            this.button7.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.button7.UseVisualStyleBackColor = true;
            this.button7.Click += new System.EventHandler(this.button7_Click_1);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.Color.Transparent;
            this.label1.Location = new System.Drawing.Point(9, 2);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(148, 13);
            this.label1.TabIndex = 25;
            this.label1.Text = "Область (край, республика)";
            // 
            // region_box
            // 
            this.region_box.BackColor = System.Drawing.Color.White;
            this.region_box.Location = new System.Drawing.Point(12, 18);
            this.region_box.Name = "region_box";
            this.region_box.ReadOnly = true;
            this.region_box.Size = new System.Drawing.Size(289, 20);
            this.region_box.TabIndex = 26;
            this.region_box.Text = "Сахалинская область";
            // 
            // nas_punkt_box
            // 
            this.nas_punkt_box.BackColor = System.Drawing.Color.White;
            this.nas_punkt_box.Location = new System.Drawing.Point(12, 60);
            this.nas_punkt_box.Name = "nas_punkt_box";
            this.nas_punkt_box.ReadOnly = true;
            this.nas_punkt_box.Size = new System.Drawing.Size(289, 20);
            this.nas_punkt_box.TabIndex = 28;
            this.nas_punkt_box.Text = "город Южно-Сахалинск";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BackColor = System.Drawing.Color.Transparent;
            this.label2.Location = new System.Drawing.Point(9, 44);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(102, 13);
            this.label2.TabIndex = 27;
            this.label2.Text = "Населённый пункт";
            // 
            // street
            // 
            this.street.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.street.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.CustomSource;
            this.street.BackColor = System.Drawing.Color.White;
            this.street.Location = new System.Drawing.Point(12, 103);
            this.street.Name = "street";
            this.street.ReadOnly = true;
            this.street.Size = new System.Drawing.Size(289, 20);
            this.street.TabIndex = 30;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.BackColor = System.Drawing.Color.Transparent;
            this.label3.Location = new System.Drawing.Point(9, 87);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(39, 13);
            this.label3.TabIndex = 29;
            this.label3.Text = "Улица";
            // 
            // house
            // 
            this.house.Location = new System.Drawing.Point(11, 150);
            this.house.MaxLength = 5;
            this.house.Name = "house";
            this.house.Size = new System.Drawing.Size(184, 20);
            this.house.TabIndex = 32;
            this.house.Text = "1";
            // 
            // Дом
            // 
            this.Дом.AutoSize = true;
            this.Дом.BackColor = System.Drawing.Color.Transparent;
            this.Дом.Location = new System.Drawing.Point(9, 134);
            this.Дом.Name = "Дом";
            this.Дом.Size = new System.Drawing.Size(173, 13);
            this.Дом.TabIndex = 31;
            this.Дом.Text = "Номер дома (строения, корпуса)";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.BackColor = System.Drawing.Color.Transparent;
            this.label4.Location = new System.Drawing.Point(198, 134);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(55, 13);
            this.label4.TabIndex = 33;
            this.label4.Text = "Квартира";
            // 
            // born_place__button
            // 
            this.born_place__button.BackColor = System.Drawing.Color.Transparent;
            this.born_place__button.FlatAppearance.MouseDownBackColor = System.Drawing.Color.CornflowerBlue;
            this.born_place__button.FlatAppearance.MouseOverBackColor = System.Drawing.Color.LightBlue;
            this.born_place__button.Font = new System.Drawing.Font("Wingdings", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(2)));
            this.born_place__button.Image = global::FSystem.Properties.Resources.CRDFLE07;
            this.born_place__button.Location = new System.Drawing.Point(307, 13);
            this.born_place__button.Name = "born_place__button";
            this.born_place__button.Size = new System.Drawing.Size(34, 27);
            this.born_place__button.TabIndex = 27;
            this.born_place__button.UseVisualStyleBackColor = false;
            this.born_place__button.Visible = false;
            this.born_place__button.Click += new System.EventHandler(this.born_place__button_Click);
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.Transparent;
            this.button1.FlatAppearance.MouseDownBackColor = System.Drawing.Color.CornflowerBlue;
            this.button1.FlatAppearance.MouseOverBackColor = System.Drawing.Color.LightBlue;
            this.button1.Font = new System.Drawing.Font("Wingdings", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(2)));
            this.button1.Image = global::FSystem.Properties.Resources.CRDFLE07;
            this.button1.Location = new System.Drawing.Point(307, 57);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(34, 27);
            this.button1.TabIndex = 29;
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.BackColor = System.Drawing.Color.Transparent;
            this.button2.FlatAppearance.MouseDownBackColor = System.Drawing.Color.CornflowerBlue;
            this.button2.FlatAppearance.MouseOverBackColor = System.Drawing.Color.LightBlue;
            this.button2.Font = new System.Drawing.Font("Wingdings", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(2)));
            this.button2.Image = global::FSystem.Properties.Resources.CRDFLE07;
            this.button2.Location = new System.Drawing.Point(307, 99);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(34, 27);
            this.button2.TabIndex = 31;
            this.toolTip1.SetToolTip(this.button2, "Выберите название улицы");
            this.button2.UseVisualStyleBackColor = false;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // Kvartira
            // 
            this.Kvartira.Location = new System.Drawing.Point(201, 151);
            this.Kvartira.Name = "Kvartira";
            this.Kvartira.Size = new System.Drawing.Size(120, 20);
            this.Kvartira.TabIndex = 37;
            this.Kvartira.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // student_edit_address
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = global::FSystem.Properties.Resources.tab_fon;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.CancelButton = this.button8;
            this.ClientSize = new System.Drawing.Size(348, 223);
            this.Controls.Add(this.Kvartira);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.born_place__button);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.house);
            this.Controls.Add(this.Дом);
            this.Controls.Add(this.street);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.nas_punkt_box);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.region_box);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button8);
            this.Controls.Add(this.button7);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "student_edit_address";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Адрес прописки студента";
            this.Load += new System.EventHandler(this.student_edit_address_Load);
            ((System.ComponentModel.ISupportInitialize)(this.Kvartira)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        public System.Windows.Forms.Button button8;
        public System.Windows.Forms.Button button7;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label Дом;
        private System.Windows.Forms.Label label4;
        public System.Windows.Forms.Button born_place__button;
        public System.Windows.Forms.Button button1;
        public System.Windows.Forms.Button button2;
        private System.Windows.Forms.ToolTip toolTip1;
        public System.Windows.Forms.TextBox street;
        public System.Windows.Forms.TextBox region_box;
        public System.Windows.Forms.TextBox nas_punkt_box;
        public System.Windows.Forms.TextBox house;
        public System.Windows.Forms.NumericUpDown Kvartira;
    }
}