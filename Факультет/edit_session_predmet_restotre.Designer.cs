namespace FSystem
{
    partial class edit_session_predmet_restotre
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(edit_session_predmet_restotre));
            this.PredmetGrid = new System.Windows.Forms.DataGridView();
            this.Пр = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.фк = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Семе = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.label1 = new System.Windows.Forms.Label();
            this.ResultGrid = new System.Windows.Forms.DataGridView();
            this.dataGridViewTextBoxColumn1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.label2 = new System.Windows.Forms.Label();
            this.button2 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.PredmetGrid)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.ResultGrid)).BeginInit();
            this.SuspendLayout();
            // 
            // PredmetGrid
            // 
            this.PredmetGrid.AllowUserToAddRows = false;
            this.PredmetGrid.AllowUserToDeleteRows = false;
            this.PredmetGrid.BackgroundColor = System.Drawing.Color.White;
            this.PredmetGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.PredmetGrid.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Пр,
            this.фк,
            this.Семе});
            this.PredmetGrid.Location = new System.Drawing.Point(12, 35);
            this.PredmetGrid.Margin = new System.Windows.Forms.Padding(4);
            this.PredmetGrid.Name = "PredmetGrid";
            this.PredmetGrid.Size = new System.Drawing.Size(586, 357);
            this.PredmetGrid.TabIndex = 0;
            this.PredmetGrid.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.PredmetGrid_CellClick);
            // 
            // Пр
            // 
            this.Пр.Frozen = true;
            this.Пр.HeaderText = "Предмет";
            this.Пр.Name = "Пр";
            this.Пр.ReadOnly = true;
            this.Пр.Width = 250;
            // 
            // фк
            // 
            this.фк.Frozen = true;
            this.фк.HeaderText = "Форма контроля";
            this.фк.Name = "фк";
            this.фк.ReadOnly = true;
            this.фк.Width = 200;
            // 
            // Семе
            // 
            this.Семе.Frozen = true;
            this.Семе.HeaderText = "Семестр";
            this.Семе.Name = "Семе";
            this.Семе.ReadOnly = true;
            this.Семе.Width = 80;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.Color.Transparent;
            this.label1.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.Location = new System.Drawing.Point(11, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(436, 18);
            this.label1.TabIndex = 1;
            this.label1.Text = "Список ранее удалённых предметов сесси для группы ";
            // 
            // ResultGrid
            // 
            this.ResultGrid.AllowUserToAddRows = false;
            this.ResultGrid.AllowUserToDeleteRows = false;
            this.ResultGrid.BackgroundColor = System.Drawing.Color.White;
            this.ResultGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.ResultGrid.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn1,
            this.dataGridViewTextBoxColumn2});
            this.ResultGrid.Location = new System.Drawing.Point(606, 35);
            this.ResultGrid.Margin = new System.Windows.Forms.Padding(4);
            this.ResultGrid.Name = "ResultGrid";
            this.ResultGrid.Size = new System.Drawing.Size(448, 357);
            this.ResultGrid.TabIndex = 2;
            // 
            // dataGridViewTextBoxColumn1
            // 
            this.dataGridViewTextBoxColumn1.Frozen = true;
            this.dataGridViewTextBoxColumn1.HeaderText = "Студент";
            this.dataGridViewTextBoxColumn1.Name = "dataGridViewTextBoxColumn1";
            this.dataGridViewTextBoxColumn1.ReadOnly = true;
            this.dataGridViewTextBoxColumn1.Width = 250;
            // 
            // dataGridViewTextBoxColumn2
            // 
            this.dataGridViewTextBoxColumn2.Frozen = true;
            this.dataGridViewTextBoxColumn2.HeaderText = "Оценка";
            this.dataGridViewTextBoxColumn2.Name = "dataGridViewTextBoxColumn2";
            this.dataGridViewTextBoxColumn2.ReadOnly = true;
            this.dataGridViewTextBoxColumn2.Width = 150;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BackColor = System.Drawing.Color.Transparent;
            this.label2.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label2.Location = new System.Drawing.Point(12, 405);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(310, 18);
            this.label2.TabIndex = 3;
            this.label2.Text = "Выбран для восстановления предмет: ";
            this.label2.Visible = false;
            // 
            // button2
            // 
            this.button2.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button2.Image = global::FSystem.Properties.Resources.delete_x16_h;
            this.button2.Location = new System.Drawing.Point(949, 402);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(104, 35);
            this.button2.TabIndex = 8;
            this.button2.Text = "отменить";
            this.button2.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.button2.UseVisualStyleBackColor = true;
            // 
            // button1
            // 
            this.button1.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button1.Image = global::FSystem.Properties.Resources.ok;
            this.button1.Location = new System.Drawing.Point(839, 402);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(104, 35);
            this.button1.TabIndex = 7;
            this.button1.Text = "принять";
            this.button1.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(65, 462);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(792, 257);
            this.textBox1.TabIndex = 9;
            // 
            // edit_session_predmet_restotre
            // 
            this.AcceptButton = this.button1;
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = global::FSystem.Properties.Resources.fon_бел_син_бледн;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.ClientSize = new System.Drawing.Size(1061, 444);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.ResultGrid);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.PredmetGrid);
            this.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.MaximizeBox = false;
            this.Name = "edit_session_predmet_restotre";
            this.ShowInTaskbar = false;
            this.Text = "Восстановление предмета в перечене сессии";
            this.Load += new System.EventHandler(this.edit_session_predmet_restotre_Load);
            ((System.ComponentModel.ISupportInitialize)(this.PredmetGrid)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.ResultGrid)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView PredmetGrid;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DataGridViewTextBoxColumn Пр;
        private System.Windows.Forms.DataGridViewTextBoxColumn фк;
        private System.Windows.Forms.DataGridViewTextBoxColumn Семе;
        private System.Windows.Forms.DataGridView ResultGrid;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn2;
        private System.Windows.Forms.Label label2;
        public System.Windows.Forms.Button button2;
        public System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox textBox1;
    }
}