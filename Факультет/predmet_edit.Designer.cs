namespace FSystem
{
    partial class predmet_edit
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(predmet_edit));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.linkLabel5 = new System.Windows.Forms.LinkLabel();
            this.semestr = new System.Windows.Forms.NumericUpDown();
            this.label4 = new System.Windows.Forms.Label();
            this.delenie_list = new System.Windows.Forms.ComboBox();
            this.linkLabel4 = new System.Windows.Forms.LinkLabel();
            this.type_predmet_list = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.linkLabel3 = new System.Windows.Forms.LinkLabel();
            this.grupa_list = new System.Windows.Forms.ComboBox();
            this.linkLabel2 = new System.Windows.Forms.LinkLabel();
            this.kaf_list = new System.Windows.Forms.ComboBox();
            this.linkLabel1 = new System.Windows.Forms.LinkLabel();
            this.prepod_list = new System.Windows.Forms.ComboBox();
            this.krat_name = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.full_name = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.vid_view = new System.Windows.Forms.ListView();
            this.vid = new System.Windows.Forms.ColumnHeader();
            this.chass = new System.Windows.Forms.ColumnHeader();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.задатьКоличествоЧасовEnterToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.выполнитьПересчетЧасовПоКонтингентуToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.exclude_button = new System.Windows.Forms.Button();
            this.include_button = new System.Windows.Forms.Button();
            this.vid_list = new System.Windows.Forms.ListBox();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.KreditUpDown = new System.Windows.Forms.NumericUpDown();
            this.label5 = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.semestr)).BeginInit();
            this.groupBox2.SuspendLayout();
            this.contextMenuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.KreditUpDown)).BeginInit();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.BackColor = System.Drawing.Color.Transparent;
            this.groupBox1.Controls.Add(this.KreditUpDown);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.linkLabel5);
            this.groupBox1.Controls.Add(this.semestr);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.delenie_list);
            this.groupBox1.Controls.Add(this.linkLabel4);
            this.groupBox1.Controls.Add(this.type_predmet_list);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.linkLabel3);
            this.groupBox1.Controls.Add(this.grupa_list);
            this.groupBox1.Controls.Add(this.linkLabel2);
            this.groupBox1.Controls.Add(this.kaf_list);
            this.groupBox1.Controls.Add(this.linkLabel1);
            this.groupBox1.Controls.Add(this.prepod_list);
            this.groupBox1.Controls.Add(this.krat_name);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.full_name);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Location = new System.Drawing.Point(12, 13);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.groupBox1.Size = new System.Drawing.Size(808, 186);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "   название, группа, преподаватель ... ";
            // 
            // linkLabel5
            // 
            this.linkLabel5.AutoSize = true;
            this.linkLabel5.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.linkLabel5.LinkColor = System.Drawing.Color.Green;
            this.linkLabel5.Location = new System.Drawing.Point(182, 74);
            this.linkLabel5.Name = "linkLabel5";
            this.linkLabel5.Size = new System.Drawing.Size(70, 16);
            this.linkLabel5.TabIndex = 101;
            this.linkLabel5.TabStop = true;
            this.linkLabel5.Text = "добавить";
            this.linkLabel5.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel5_LinkClicked);
            // 
            // semestr
            // 
            this.semestr.BackColor = System.Drawing.Color.White;
            this.semestr.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.semestr.Location = new System.Drawing.Point(270, 143);
            this.semestr.Name = "semestr";
            this.semestr.ReadOnly = true;
            this.semestr.Size = new System.Drawing.Size(169, 23);
            this.semestr.TabIndex = 7;
            this.semestr.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(453, 123);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(141, 16);
            this.label4.TabIndex = 100;
            this.label4.Text = "деление на подгруппы";
            // 
            // delenie_list
            // 
            this.delenie_list.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.delenie_list.FormattingEnabled = true;
            this.delenie_list.Items.AddRange(new object[] {
            "деления нет",
            "деление есть"});
            this.delenie_list.Location = new System.Drawing.Point(456, 142);
            this.delenie_list.Name = "delenie_list";
            this.delenie_list.Size = new System.Drawing.Size(202, 24);
            this.delenie_list.TabIndex = 8;
            this.delenie_list.SelectedIndexChanged += new System.EventHandler(this.delenie_list_SelectedIndexChanged);
            // 
            // linkLabel4
            // 
            this.linkLabel4.AutoSize = true;
            this.linkLabel4.Location = new System.Drawing.Point(453, 73);
            this.linkLabel4.Name = "linkLabel4";
            this.linkLabel4.Size = new System.Drawing.Size(89, 16);
            this.linkLabel4.TabIndex = 18;
            this.linkLabel4.TabStop = true;
            this.linkLabel4.Text = "вид предмета";
            // 
            // type_predmet_list
            // 
            this.type_predmet_list.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.type_predmet_list.FormattingEnabled = true;
            this.type_predmet_list.Location = new System.Drawing.Point(456, 92);
            this.type_predmet_list.Name = "type_predmet_list";
            this.type_predmet_list.Size = new System.Drawing.Size(202, 24);
            this.type_predmet_list.TabIndex = 5;
            this.type_predmet_list.SelectedIndexChanged += new System.EventHandler(this.type_predmet_list_SelectedIndexChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(267, 124);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(55, 16);
            this.label3.TabIndex = 100;
            this.label3.Text = "семестр";
            // 
            // linkLabel3
            // 
            this.linkLabel3.AutoSize = true;
            this.linkLabel3.Location = new System.Drawing.Point(3, 124);
            this.linkLabel3.Name = "linkLabel3";
            this.linkLabel3.Size = new System.Drawing.Size(59, 16);
            this.linkLabel3.TabIndex = 19;
            this.linkLabel3.TabStop = true;
            this.linkLabel3.Text = "кафедра";
            // 
            // grupa_list
            // 
            this.grupa_list.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.grupa_list.FormattingEnabled = true;
            this.grupa_list.Location = new System.Drawing.Point(270, 93);
            this.grupa_list.Name = "grupa_list";
            this.grupa_list.Size = new System.Drawing.Size(169, 24);
            this.grupa_list.TabIndex = 4;
            this.grupa_list.SelectedIndexChanged += new System.EventHandler(this.grupa_list_SelectedIndexChanged);
            // 
            // linkLabel2
            // 
            this.linkLabel2.AutoSize = true;
            this.linkLabel2.Location = new System.Drawing.Point(267, 74);
            this.linkLabel2.Name = "linkLabel2";
            this.linkLabel2.Size = new System.Drawing.Size(47, 16);
            this.linkLabel2.TabIndex = 17;
            this.linkLabel2.TabStop = true;
            this.linkLabel2.Text = "группа";
            this.linkLabel2.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel2_LinkClicked);
            // 
            // kaf_list
            // 
            this.kaf_list.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.kaf_list.FormattingEnabled = true;
            this.kaf_list.Location = new System.Drawing.Point(6, 143);
            this.kaf_list.Name = "kaf_list";
            this.kaf_list.Size = new System.Drawing.Size(247, 24);
            this.kaf_list.TabIndex = 6;
            this.kaf_list.SelectedIndexChanged += new System.EventHandler(this.kaf_list_SelectedIndexChanged);
            // 
            // linkLabel1
            // 
            this.linkLabel1.AutoSize = true;
            this.linkLabel1.Location = new System.Drawing.Point(6, 74);
            this.linkLabel1.Name = "linkLabel1";
            this.linkLabel1.Size = new System.Drawing.Size(98, 16);
            this.linkLabel1.TabIndex = 16;
            this.linkLabel1.TabStop = true;
            this.linkLabel1.Text = "преподаватель";
            this.linkLabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel1_LinkClicked);
            // 
            // prepod_list
            // 
            this.prepod_list.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.prepod_list.FormattingEnabled = true;
            this.prepod_list.Location = new System.Drawing.Point(6, 93);
            this.prepod_list.Name = "prepod_list";
            this.prepod_list.Size = new System.Drawing.Size(247, 24);
            this.prepod_list.TabIndex = 3;
            this.prepod_list.SelectedIndexChanged += new System.EventHandler(this.prepod_list_SelectedIndexChanged);
            // 
            // krat_name
            // 
            this.krat_name.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.krat_name.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.CustomSource;
            this.krat_name.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.krat_name.ForeColor = System.Drawing.Color.Navy;
            this.krat_name.Location = new System.Drawing.Point(594, 41);
            this.krat_name.MaxLength = 22;
            this.krat_name.Name = "krat_name";
            this.krat_name.Size = new System.Drawing.Size(208, 23);
            this.krat_name.TabIndex = 2;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label2.Location = new System.Drawing.Point(591, 22);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(181, 16);
            this.label2.TabIndex = 2;
            this.label2.Text = "Название для расписания";
            // 
            // full_name
            // 
            this.full_name.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.full_name.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.CustomSource;
            this.full_name.Location = new System.Drawing.Point(6, 41);
            this.full_name.Name = "full_name";
            this.full_name.Size = new System.Drawing.Size(582, 23);
            this.full_name.TabIndex = 1;
            this.full_name.TextChanged += new System.EventHandler(this.full_name_TextChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.Location = new System.Drawing.Point(3, 22);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(195, 16);
            this.label1.TabIndex = 0;
            this.label1.Text = "Полное название предмета";
            // 
            // groupBox2
            // 
            this.groupBox2.BackColor = System.Drawing.Color.Transparent;
            this.groupBox2.Controls.Add(this.vid_view);
            this.groupBox2.Controls.Add(this.exclude_button);
            this.groupBox2.Controls.Add(this.include_button);
            this.groupBox2.Controls.Add(this.vid_list);
            this.groupBox2.Location = new System.Drawing.Point(12, 207);
            this.groupBox2.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Padding = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.groupBox2.Size = new System.Drawing.Size(808, 231);
            this.groupBox2.TabIndex = 3;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "  виды занятий и распределение часов  ";
            // 
            // vid_view
            // 
            this.vid_view.AutoArrange = false;
            this.vid_view.BackgroundImageTiled = true;
            this.vid_view.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.vid,
            this.chass});
            this.vid_view.ContextMenuStrip = this.contextMenuStrip1;
            this.vid_view.FullRowSelect = true;
            this.vid_view.GridLines = true;
            this.vid_view.Location = new System.Drawing.Point(351, 29);
            this.vid_view.MultiSelect = false;
            this.vid_view.Name = "vid_view";
            this.vid_view.Size = new System.Drawing.Size(446, 196);
            this.vid_view.TabIndex = 11;
            this.toolTip1.SetToolTip(this.vid_view, "используйте двойной щелчок или Enter \r\nдля задания количества часов");
            this.vid_view.UseCompatibleStateImageBehavior = false;
            this.vid_view.View = System.Windows.Forms.View.Details;
            this.vid_view.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.vid_view_MouseDoubleClick);
            this.vid_view.SelectedIndexChanged += new System.EventHandler(this.vid_view_SelectedIndexChanged);
            this.vid_view.MouseDown += new System.Windows.Forms.MouseEventHandler(this.vid_view_MouseDown);
            this.vid_view.KeyDown += new System.Windows.Forms.KeyEventHandler(this.vid_view_KeyDown);
            // 
            // vid
            // 
            this.vid.Text = "Вид занятия";
            this.vid.Width = 270;
            // 
            // chass
            // 
            this.chass.Text = "Количество часов";
            this.chass.Width = 150;
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.задатьКоличествоЧасовEnterToolStripMenuItem,
            this.выполнитьПересчетЧасовПоКонтингентуToolStripMenuItem});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(325, 48);
            // 
            // задатьКоличествоЧасовEnterToolStripMenuItem
            // 
            this.задатьКоличествоЧасовEnterToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("задатьКоличествоЧасовEnterToolStripMenuItem.Image")));
            this.задатьКоличествоЧасовEnterToolStripMenuItem.Name = "задатьКоличествоЧасовEnterToolStripMenuItem";
            this.задатьКоличествоЧасовEnterToolStripMenuItem.Size = new System.Drawing.Size(324, 22);
            this.задатьКоличествоЧасовEnterToolStripMenuItem.Text = "задать количество часов (Enter)";
            this.задатьКоличествоЧасовEnterToolStripMenuItem.Click += new System.EventHandler(this.задатьКоличествоЧасовEnterToolStripMenuItem_Click);
            // 
            // выполнитьПересчетЧасовПоКонтингентуToolStripMenuItem
            // 
            this.выполнитьПересчетЧасовПоКонтингентуToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("выполнитьПересчетЧасовПоКонтингентуToolStripMenuItem.Image")));
            this.выполнитьПересчетЧасовПоКонтингентуToolStripMenuItem.Name = "выполнитьПересчетЧасовПоКонтингентуToolStripMenuItem";
            this.выполнитьПересчетЧасовПоКонтингентуToolStripMenuItem.Size = new System.Drawing.Size(324, 22);
            this.выполнитьПересчетЧасовПоКонтингентуToolStripMenuItem.Text = "Выполнить пересчет часов по контингенту ...";
            this.выполнитьПересчетЧасовПоКонтингентуToolStripMenuItem.Click += new System.EventHandler(this.выполнитьПересчетЧасовПоКонтингентуToolStripMenuItem_Click);
            // 
            // exclude_button
            // 
            this.exclude_button.Font = new System.Drawing.Font("Wingdings", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(2)));
            this.exclude_button.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.exclude_button.Location = new System.Drawing.Point(270, 134);
            this.exclude_button.Name = "exclude_button";
            this.exclude_button.Size = new System.Drawing.Size(67, 31);
            this.exclude_button.TabIndex = 12;
            this.exclude_button.Text = "";
            this.toolTip1.SetToolTip(this.exclude_button, "удалить из предмета");
            this.exclude_button.UseVisualStyleBackColor = true;
            this.exclude_button.Click += new System.EventHandler(this.exclude_button_Click);
            // 
            // include_button
            // 
            this.include_button.Font = new System.Drawing.Font("Wingdings", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(2)));
            this.include_button.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.include_button.Location = new System.Drawing.Point(270, 80);
            this.include_button.Name = "include_button";
            this.include_button.Size = new System.Drawing.Size(67, 31);
            this.include_button.TabIndex = 10;
            this.include_button.Text = "";
            this.toolTip1.SetToolTip(this.include_button, "добавить в предмет");
            this.include_button.UseVisualStyleBackColor = true;
            this.include_button.Click += new System.EventHandler(this.include_button_Click);
            // 
            // vid_list
            // 
            this.vid_list.FormattingEnabled = true;
            this.vid_list.ItemHeight = 16;
            this.vid_list.Location = new System.Drawing.Point(11, 29);
            this.vid_list.Name = "vid_list";
            this.vid_list.Size = new System.Drawing.Size(242, 196);
            this.vid_list.TabIndex = 9;
            this.toolTip1.SetToolTip(this.vid_list, "перечень доступных видов\r\nзанятий для данного предмета");
            this.vid_list.KeyDown += new System.Windows.Forms.KeyEventHandler(this.vid_list_KeyDown);
            // 
            // button1
            // 
            this.button1.Image = global::FSystem.Properties.Resources.ok;
            this.button1.Location = new System.Drawing.Point(590, 455);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(107, 29);
            this.button1.TabIndex = 14;
            this.button1.Text = "Принять";
            this.button1.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button2.Image = global::FSystem.Properties.Resources.delete_x16_h;
            this.button2.Location = new System.Drawing.Point(702, 455);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(107, 29);
            this.button2.TabIndex = 15;
            this.button2.Text = "Отменить";
            this.button2.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.button2.UseVisualStyleBackColor = true;
            // 
            // toolTip1
            // 
            this.toolTip1.BackColor = System.Drawing.Color.AliceBlue;
            this.toolTip1.IsBalloon = true;
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.BackColor = System.Drawing.Color.Transparent;
            this.checkBox1.Font = new System.Drawing.Font("Tahoma", 9.75F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.checkBox1.ForeColor = System.Drawing.Color.Red;
            this.checkBox1.Location = new System.Drawing.Point(23, 455);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(342, 20);
            this.checkBox1.TabIndex = 13;
            this.checkBox1.Text = "предмет находится в текущем учебном плане";
            this.toolTip1.SetToolTip(this.checkBox1, "Если предмет нужно удалить из расписания и всех видов \r\nотчетности, то снимите вы" +
                    "деление этого пункта!");
            this.checkBox1.UseVisualStyleBackColor = false;
            // 
            // KreditUpDown
            // 
            this.KreditUpDown.BackColor = System.Drawing.Color.White;
            this.KreditUpDown.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.KreditUpDown.Location = new System.Drawing.Point(668, 94);
            this.KreditUpDown.Maximum = new decimal(new int[] {
            10,
            0,
            0,
            0});
            this.KreditUpDown.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.KreditUpDown.Name = "KreditUpDown";
            this.KreditUpDown.ReadOnly = true;
            this.KreditUpDown.Size = new System.Drawing.Size(108, 23);
            this.KreditUpDown.TabIndex = 102;
            this.KreditUpDown.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.KreditUpDown.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label5.ForeColor = System.Drawing.Color.Blue;
            this.label5.Location = new System.Drawing.Point(665, 75);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(115, 16);
            this.label5.TabIndex = 103;
            this.label5.Text = "Число кредитов";
            // 
            // predmet_edit
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = global::FSystem.Properties.Resources.tab_fon;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.CancelButton = this.button2;
            this.ClientSize = new System.Drawing.Size(831, 496);
            this.Controls.Add(this.checkBox1);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.MaximizeBox = false;
            this.Name = "predmet_edit";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "   Сведения об учебной дисциплине";
            this.Load += new System.EventHandler(this.predmet_edit_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.semestr)).EndInit();
            this.groupBox2.ResumeLayout(false);
            this.contextMenuStrip1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.KreditUpDown)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        public System.Windows.Forms.GroupBox groupBox1;
        public System.Windows.Forms.GroupBox groupBox2;
        public System.Windows.Forms.Button button1;
        public System.Windows.Forms.Button button2;
        public System.Windows.Forms.ListBox vid_list;
        public System.Windows.Forms.Button exclude_button;
        public System.Windows.Forms.Button include_button;
        public System.Windows.Forms.ToolTip toolTip1;
        public System.Windows.Forms.TextBox krat_name;
        public System.Windows.Forms.Label label2;
        public System.Windows.Forms.TextBox full_name;
        public System.Windows.Forms.Label label1;
        public System.Windows.Forms.ComboBox prepod_list;
        public System.Windows.Forms.LinkLabel linkLabel2;
        public System.Windows.Forms.ComboBox kaf_list;
        public System.Windows.Forms.LinkLabel linkLabel1;
        public System.Windows.Forms.Label label3;
        public System.Windows.Forms.LinkLabel linkLabel3;
        public System.Windows.Forms.ComboBox grupa_list;
        public System.Windows.Forms.LinkLabel linkLabel4;
        public System.Windows.Forms.ComboBox type_predmet_list;
        public System.Windows.Forms.Label label4;
        public System.Windows.Forms.CheckBox checkBox1;
        public System.Windows.Forms.ComboBox delenie_list;
        public System.Windows.Forms.NumericUpDown semestr;
        public System.Windows.Forms.ColumnHeader vid;
        public System.Windows.Forms.ListView vid_view;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem задатьКоличествоЧасовEnterToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem выполнитьПересчетЧасовПоКонтингентуToolStripMenuItem;
        private System.Windows.Forms.LinkLabel linkLabel5;
        public System.Windows.Forms.ColumnHeader chass;
        public System.Windows.Forms.NumericUpDown KreditUpDown;
        public System.Windows.Forms.Label label5;


    }
}