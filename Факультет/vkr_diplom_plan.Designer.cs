namespace FSystem
{
    partial class vkr_diplom_plan
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(vkr_diplom_plan));
            this.vkr = new System.Windows.Forms.GroupBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.vkr_chas = new System.Windows.Forms.NumericUpDown();
            this.vkr_student = new System.Windows.Forms.NumericUpDown();
            this.button2 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.dip_chas = new System.Windows.Forms.NumericUpDown();
            this.dip_student = new System.Windows.Forms.NumericUpDown();
            this.vkr.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.vkr_chas)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.vkr_student)).BeginInit();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dip_chas)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dip_student)).BeginInit();
            this.SuspendLayout();
            // 
            // vkr
            // 
            this.vkr.BackColor = System.Drawing.Color.Transparent;
            this.vkr.Controls.Add(this.label2);
            this.vkr.Controls.Add(this.label1);
            this.vkr.Controls.Add(this.vkr_chas);
            this.vkr.Controls.Add(this.vkr_student);
            this.vkr.Location = new System.Drawing.Point(13, 12);
            this.vkr.Name = "vkr";
            this.vkr.Size = new System.Drawing.Size(302, 108);
            this.vkr.TabIndex = 0;
            this.vkr.TabStop = false;
            this.vkr.Text = "  ВКР  ";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(6, 74);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(189, 14);
            this.label2.TabIndex = 3;
            this.label2.Text = "Кол-во часов на каждую работу";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 28);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(193, 14);
            this.label1.TabIndex = 2;
            this.label1.Text = "Кол-во студентов по поручению";
            // 
            // vkr_chas
            // 
            this.vkr_chas.Location = new System.Drawing.Point(205, 72);
            this.vkr_chas.Maximum = new decimal(new int[] {
            25,
            0,
            0,
            0});
            this.vkr_chas.Minimum = new decimal(new int[] {
            5,
            0,
            0,
            0});
            this.vkr_chas.Name = "vkr_chas";
            this.vkr_chas.Size = new System.Drawing.Size(82, 22);
            this.vkr_chas.TabIndex = 1;
            this.vkr_chas.Value = new decimal(new int[] {
            10,
            0,
            0,
            0});
            // 
            // vkr_student
            // 
            this.vkr_student.Location = new System.Drawing.Point(205, 26);
            this.vkr_student.Maximum = new decimal(new int[] {
            20,
            0,
            0,
            0});
            this.vkr_student.Name = "vkr_student";
            this.vkr_student.Size = new System.Drawing.Size(82, 22);
            this.vkr_student.TabIndex = 0;
            // 
            // button2
            // 
            this.button2.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button2.Image = global::FSystem.Properties.Resources.delete_x16_h;
            this.button2.Location = new System.Drawing.Point(121, 242);
            this.button2.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(87, 28);
            this.button2.TabIndex = 8;
            this.button2.Text = "отменить";
            this.button2.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.button2.UseVisualStyleBackColor = true;
            // 
            // button1
            // 
            this.button1.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.button1.Image = global::FSystem.Properties.Resources.ok;
            this.button1.Location = new System.Drawing.Point(12, 242);
            this.button1.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(94, 28);
            this.button1.TabIndex = 7;
            this.button1.Text = "запомнить";
            this.button1.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.button1.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            this.groupBox1.BackColor = System.Drawing.Color.Transparent;
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.dip_chas);
            this.groupBox1.Controls.Add(this.dip_student);
            this.groupBox1.Location = new System.Drawing.Point(12, 126);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(302, 108);
            this.groupBox1.TabIndex = 9;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "  Дипломные  ";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(6, 74);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(189, 14);
            this.label3.TabIndex = 3;
            this.label3.Text = "Кол-во часов на каждую работу";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(6, 28);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(193, 14);
            this.label4.TabIndex = 2;
            this.label4.Text = "Кол-во студентов по поручению";
            // 
            // dip_chas
            // 
            this.dip_chas.Location = new System.Drawing.Point(205, 72);
            this.dip_chas.Maximum = new decimal(new int[] {
            25,
            0,
            0,
            0});
            this.dip_chas.Minimum = new decimal(new int[] {
            5,
            0,
            0,
            0});
            this.dip_chas.Name = "dip_chas";
            this.dip_chas.Size = new System.Drawing.Size(82, 22);
            this.dip_chas.TabIndex = 3;
            this.dip_chas.Value = new decimal(new int[] {
            15,
            0,
            0,
            0});
            // 
            // dip_student
            // 
            this.dip_student.Location = new System.Drawing.Point(205, 26);
            this.dip_student.Maximum = new decimal(new int[] {
            20,
            0,
            0,
            0});
            this.dip_student.Name = "dip_student";
            this.dip_student.Size = new System.Drawing.Size(82, 22);
            this.dip_student.TabIndex = 2;
            // 
            // vkr_diplom_plan
            // 
            this.AcceptButton = this.button1;
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 14F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = global::FSystem.Properties.Resources.tab_fon;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.CancelButton = this.button2;
            this.ClientSize = new System.Drawing.Size(324, 279);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.vkr);
            this.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "vkr_diplom_plan";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Задание плана по ВКР и дипломным работам";
            this.vkr.ResumeLayout(false);
            this.vkr.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.vkr_chas)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.vkr_student)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dip_chas)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dip_student)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        public System.Windows.Forms.GroupBox vkr;
        public System.Windows.Forms.Label label2;
        public System.Windows.Forms.Label label1;
        public System.Windows.Forms.NumericUpDown vkr_chas;
        public System.Windows.Forms.NumericUpDown vkr_student;
        public System.Windows.Forms.Button button2;
        public System.Windows.Forms.Button button1;
        public System.Windows.Forms.GroupBox groupBox1;
        public System.Windows.Forms.Label label3;
        public System.Windows.Forms.Label label4;
        public System.Windows.Forms.NumericUpDown dip_chas;
        public System.Windows.Forms.NumericUpDown dip_student;

    }
}