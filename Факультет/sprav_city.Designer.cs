namespace FSystem
{
    partial class sprav_city
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(sprav_city));
            this.nas_punkt_box = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.button8 = new System.Windows.Forms.Button();
            this.button7 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.index = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.SuspendLayout();
            // 
            // nas_punkt_box
            // 
            this.nas_punkt_box.Location = new System.Drawing.Point(12, 32);
            this.nas_punkt_box.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.nas_punkt_box.MaxLength = 50;
            this.nas_punkt_box.Name = "nas_punkt_box";
            this.nas_punkt_box.Size = new System.Drawing.Size(300, 23);
            this.nas_punkt_box.TabIndex = 0;
            this.nas_punkt_box.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.krat_name_KeyPress);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BackColor = System.Drawing.Color.Transparent;
            this.label2.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label2.Location = new System.Drawing.Point(12, 12);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(134, 16);
            this.label2.TabIndex = 29;
            this.label2.Text = "Населённый пункт";
            // 
            // button8
            // 
            this.button8.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button8.Image = global::FSystem.Properties.Resources.delete_x16_h;
            this.button8.Location = new System.Drawing.Point(140, 117);
            this.button8.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.button8.Name = "button8";
            this.button8.Size = new System.Drawing.Size(115, 29);
            this.button8.TabIndex = 5;
            this.button8.Text = "отмена";
            this.button8.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.button8.UseVisualStyleBackColor = true;
            this.button8.Click += new System.EventHandler(this.button8_Click);
            // 
            // button7
            // 
            this.button7.Image = global::FSystem.Properties.Resources.ok;
            this.button7.Location = new System.Drawing.Point(14, 117);
            this.button7.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.button7.Name = "button7";
            this.button7.Size = new System.Drawing.Size(115, 29);
            this.button7.TabIndex = 4;
            this.button7.Text = "принять";
            this.button7.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.button7.UseVisualStyleBackColor = true;
            this.button7.Click += new System.EventHandler(this.button7_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.Color.Transparent;
            this.label1.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.Location = new System.Drawing.Point(10, 62);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(112, 16);
            this.label1.TabIndex = 39;
            this.label1.Text = "Тип нас. пункта";
            // 
            // index
            // 
            this.index.Location = new System.Drawing.Point(217, 81);
            this.index.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.index.MaxLength = 6;
            this.index.Name = "index";
            this.index.Size = new System.Drawing.Size(95, 23);
            this.index.TabIndex = 3;
            this.index.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.index_KeyPress);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.BackColor = System.Drawing.Color.Transparent;
            this.label3.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label3.Location = new System.Drawing.Point(214, 61);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(56, 16);
            this.label3.TabIndex = 41;
            this.label3.Text = "Индекс";
            // 
            // comboBox1
            // 
            this.comboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(14, 80);
            this.comboBox1.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(197, 24);
            this.comboBox1.TabIndex = 2;
            // 
            // timer1
            // 
            this.timer1.Interval = 7;
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // sprav_city
            // 
            this.AcceptButton = this.button7;
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = global::FSystem.Properties.Resources.fon_green_small;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.CancelButton = this.button8;
            this.ClientSize = new System.Drawing.Size(326, 160);
            this.Controls.Add(this.comboBox1);
            this.Controls.Add(this.index);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button8);
            this.Controls.Add(this.button7);
            this.Controls.Add(this.nas_punkt_box);
            this.Controls.Add(this.label2);
            this.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "sprav_city";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.Text = "Города ...";
            this.Move += new System.EventHandler(this.sprav_city_Move);
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.sprav_city_FormClosing);
            this.Load += new System.EventHandler(this.sprav_city_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        public System.Windows.Forms.Button button8;
        public System.Windows.Forms.Button button7;
        private System.Windows.Forms.Label label1;
        public System.Windows.Forms.TextBox nas_punkt_box;
        public System.Windows.Forms.TextBox index;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.Timer timer1;
        public System.Windows.Forms.Label label2;
        public System.Windows.Forms.Label label3;
    }
}