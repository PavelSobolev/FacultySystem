namespace FSystem
{
    partial class set_tema_dialog
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
            this.tema1 = new System.Windows.Forms.RichTextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.tema2 = new System.Windows.Forms.RichTextBox();
            this.choose_cynchro = new System.Windows.Forms.CheckBox();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // tema1
            // 
            this.tema1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.tema1.DetectUrls = false;
            this.tema1.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.tema1.ForeColor = System.Drawing.Color.Navy;
            this.tema1.Location = new System.Drawing.Point(12, 29);
            this.tema1.Name = "tema1";
            this.tema1.Size = new System.Drawing.Size(418, 88);
            this.tema1.TabIndex = 0;
            this.tema1.Text = "";
            this.tema1.TextChanged += new System.EventHandler(this.tema1_TextChanged);
            this.tema1.MouseDown += new System.Windows.Forms.MouseEventHandler(this.tema1_MouseDown);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.Color.Transparent;
            this.label1.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.Location = new System.Drawing.Point(12, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(98, 16);
            this.label1.TabIndex = 1;
            this.label1.Text = "Тема занятия";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BackColor = System.Drawing.Color.Transparent;
            this.label2.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label2.Location = new System.Drawing.Point(12, 169);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(98, 16);
            this.label2.TabIndex = 3;
            this.label2.Text = "Тема занятия";
            // 
            // tema2
            // 
            this.tema2.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.tema2.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.tema2.ForeColor = System.Drawing.Color.Navy;
            this.tema2.Location = new System.Drawing.Point(12, 185);
            this.tema2.Name = "tema2";
            this.tema2.Size = new System.Drawing.Size(418, 88);
            this.tema2.TabIndex = 2;
            this.tema2.Text = "";
            this.tema2.TextChanged += new System.EventHandler(this.tema2_TextChanged);
            this.tema2.MouseDown += new System.Windows.Forms.MouseEventHandler(this.tema2_MouseDown);
            // 
            // choose_cynchro
            // 
            this.choose_cynchro.AutoSize = true;
            this.choose_cynchro.BackColor = System.Drawing.Color.Transparent;
            this.choose_cynchro.Checked = true;
            this.choose_cynchro.CheckState = System.Windows.Forms.CheckState.Checked;
            this.choose_cynchro.Location = new System.Drawing.Point(12, 135);
            this.choose_cynchro.Name = "choose_cynchro";
            this.choose_cynchro.Size = new System.Drawing.Size(160, 17);
            this.choose_cynchro.TabIndex = 4;
            this.choose_cynchro.Text = "Одна тема на оба занятия";
            this.choose_cynchro.UseVisualStyleBackColor = false;
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.Transparent;
            this.button1.Image = global::FSystem.Properties.Resources.ok;
            this.button1.Location = new System.Drawing.Point(207, 129);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(104, 27);
            this.button1.TabIndex = 5;
            this.button1.Text = "принять";
            this.button1.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.BackColor = System.Drawing.Color.Transparent;
            this.button2.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button2.Image = global::FSystem.Properties.Resources.delete_x16_h;
            this.button2.Location = new System.Drawing.Point(326, 129);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(104, 27);
            this.button2.TabIndex = 6;
            this.button2.Text = "отменить";
            this.button2.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.button2.UseVisualStyleBackColor = false;
            // 
            // set_tema_dialog
            // 
            this.AcceptButton = this.button1;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = global::FSystem.Properties.Resources.tab_fon;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.CancelButton = this.button2;
            this.ClientSize = new System.Drawing.Size(445, 286);
            this.ControlBox = false;
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.choose_cynchro);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.tema2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.tema1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.KeyPreview = true;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "set_tema_dialog";
            this.Opacity = 0.95;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Tag = "1";
            this.Text = "   Тема занятия";
            this.Load += new System.EventHandler(this.set_tema_dialog_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        public System.Windows.Forms.Label label1;
        public System.Windows.Forms.RichTextBox tema1;
        public System.Windows.Forms.Label label2;
        public System.Windows.Forms.RichTextBox tema2;
        public System.Windows.Forms.CheckBox choose_cynchro;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
    }
}