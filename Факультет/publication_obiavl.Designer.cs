namespace FSystem
{
    partial class publication_obiavl
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(publication_obiavl));
            this.tema_list = new System.Windows.Forms.ComboBox();
            this.textob = new System.Windows.Forms.TextBox();
            this.titletxt = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.end_date = new System.Windows.Forms.DateTimePicker();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.save = new System.Windows.Forms.Button();
            this.cancell = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // tema_list
            // 
            this.tema_list.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.tema_list.FormattingEnabled = true;
            this.tema_list.Location = new System.Drawing.Point(16, 29);
            this.tema_list.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.tema_list.Name = "tema_list";
            this.tema_list.Size = new System.Drawing.Size(308, 24);
            this.tema_list.TabIndex = 0;
            this.tema_list.SelectedIndexChanged += new System.EventHandler(this.tema_list_SelectedIndexChanged);
            // 
            // textob
            // 
            this.textob.Location = new System.Drawing.Point(15, 143);
            this.textob.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.textob.Multiline = true;
            this.textob.Name = "textob";
            this.textob.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.textob.Size = new System.Drawing.Size(568, 288);
            this.textob.TabIndex = 0;
            // 
            // titletxt
            // 
            this.titletxt.Location = new System.Drawing.Point(16, 87);
            this.titletxt.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.titletxt.Name = "titletxt";
            this.titletxt.Size = new System.Drawing.Size(567, 23);
            this.titletxt.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.Color.Transparent;
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(142, 16);
            this.label1.TabIndex = 5;
            this.label1.Text = "Категория объявления";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BackColor = System.Drawing.Color.Transparent;
            this.label2.Location = new System.Drawing.Point(12, 66);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(142, 16);
            this.label2.TabIndex = 6;
            this.label2.Text = "Заголовок объявления";
            // 
            // end_date
            // 
            this.end_date.Location = new System.Drawing.Point(352, 29);
            this.end_date.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.end_date.Name = "end_date";
            this.end_date.Size = new System.Drawing.Size(231, 23);
            this.end_date.TabIndex = 7;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.BackColor = System.Drawing.Color.Transparent;
            this.label3.Location = new System.Drawing.Point(348, 9);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(175, 16);
            this.label3.TabIndex = 8;
            this.label3.Text = "Дата окончания публикации";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.BackColor = System.Drawing.Color.Transparent;
            this.label4.Location = new System.Drawing.Point(13, 123);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(115, 16);
            this.label4.TabIndex = 9;
            this.label4.Text = "Текст объявления";
            // 
            // save
            // 
            this.save.Image = global::FSystem.Properties.Resources.save_green16;
            this.save.Location = new System.Drawing.Point(16, 449);
            this.save.Name = "save";
            this.save.Size = new System.Drawing.Size(112, 28);
            this.save.TabIndex = 10;
            this.save.Text = "Сохранить";
            this.save.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.save.UseVisualStyleBackColor = true;
            this.save.Click += new System.EventHandler(this.save_Click);
            // 
            // cancell
            // 
            this.cancell.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.cancell.Image = global::FSystem.Properties.Resources.delete_x16_h;
            this.cancell.Location = new System.Drawing.Point(148, 449);
            this.cancell.Name = "cancell";
            this.cancell.Size = new System.Drawing.Size(112, 28);
            this.cancell.TabIndex = 11;
            this.cancell.Text = "Отмена";
            this.cancell.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.cancell.UseVisualStyleBackColor = true;
            // 
            // publication_obiavl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = global::FSystem.Properties.Resources.tab_fon;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.CancelButton = this.cancell;
            this.ClientSize = new System.Drawing.Size(591, 486);
            this.Controls.Add(this.cancell);
            this.Controls.Add(this.save);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.textob);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.end_date);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.titletxt);
            this.Controls.Add(this.tema_list);
            this.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.MaximizeBox = false;
            this.Name = "publication_obiavl";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Создание объявления на сайт";
            this.Load += new System.EventHandler(this.publication_obiavl_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        public System.Windows.Forms.ComboBox tema_list;
        public System.Windows.Forms.TextBox textob;
        public System.Windows.Forms.TextBox titletxt;
        public System.Windows.Forms.Label label1;
        public System.Windows.Forms.Label label2;
        public System.Windows.Forms.DateTimePicker end_date;
        public System.Windows.Forms.Label label3;
        public System.Windows.Forms.Label label4;
        public System.Windows.Forms.Button save;
        public System.Windows.Forms.Button cancell;

    }
}