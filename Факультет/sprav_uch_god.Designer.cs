namespace Факультет
{
    partial class sprav_uch_god
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(sprav_uch_god));
            System.Windows.Forms.Label startLabel;
            System.Windows.Forms.Label finishLabel;
            System.Windows.Forms.Label semestr1_finishLabel;
            System.Windows.Forms.Label semestr2_startLabel;
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle5 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle6 = new System.Windows.Forms.DataGridViewCellStyle();
            this.uch_godBindingNavigator = new System.Windows.Forms.BindingNavigator(this.components);
            this.bindingNavigatorMoveFirstItem = new System.Windows.Forms.ToolStripButton();
            this.bindingNavigatorMovePreviousItem = new System.Windows.Forms.ToolStripButton();
            this.bindingNavigatorSeparator = new System.Windows.Forms.ToolStripSeparator();
            this.bindingNavigatorPositionItem = new System.Windows.Forms.ToolStripTextBox();
            this.bindingNavigatorCountItem = new System.Windows.Forms.ToolStripLabel();
            this.bindingNavigatorSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.bindingNavigatorMoveNextItem = new System.Windows.Forms.ToolStripButton();
            this.bindingNavigatorMoveLastItem = new System.Windows.Forms.ToolStripButton();
            this.bindingNavigatorSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.bindingNavigatorAddNewItem = new System.Windows.Forms.ToolStripButton();
            this.uch_godBindingNavigatorSaveItem = new System.Windows.Forms.ToolStripButton();
            this.uch_godDataGridView = new System.Windows.Forms.DataGridView();
            this.startDateTimePicker = new System.Windows.Forms.DateTimePicker();
            this.finishDateTimePicker = new System.Windows.Forms.DateTimePicker();
            this.semestr1_finishDateTimePicker = new System.Windows.Forms.DateTimePicker();
            this.semestr2_startDateTimePicker = new System.Windows.Forms.DateTimePicker();
            this.uch_godBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.vKRDataSet = new Факультет.VKRDataSet();
            this.dataGridViewTextBoxColumn2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn4 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn5 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.uch_godTableAdapter = new Факультет.VKRDataSetTableAdapters.uch_godTableAdapter();
            startLabel = new System.Windows.Forms.Label();
            finishLabel = new System.Windows.Forms.Label();
            semestr1_finishLabel = new System.Windows.Forms.Label();
            semestr2_startLabel = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.uch_godBindingNavigator)).BeginInit();
            this.uch_godBindingNavigator.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.uch_godDataGridView)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.uch_godBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.vKRDataSet)).BeginInit();
            this.SuspendLayout();
            // 
            // uch_godBindingNavigator
            // 
            this.uch_godBindingNavigator.AddNewItem = this.bindingNavigatorAddNewItem;
            this.uch_godBindingNavigator.BindingSource = this.uch_godBindingSource;
            this.uch_godBindingNavigator.CountItem = this.bindingNavigatorCountItem;
            this.uch_godBindingNavigator.CountItemFormat = "запись из {0}";
            this.uch_godBindingNavigator.DeleteItem = null;
            this.uch_godBindingNavigator.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.bindingNavigatorMoveFirstItem,
            this.bindingNavigatorMovePreviousItem,
            this.bindingNavigatorSeparator,
            this.bindingNavigatorPositionItem,
            this.bindingNavigatorCountItem,
            this.bindingNavigatorSeparator1,
            this.bindingNavigatorMoveNextItem,
            this.bindingNavigatorMoveLastItem,
            this.bindingNavigatorSeparator2,
            this.bindingNavigatorAddNewItem,
            this.uch_godBindingNavigatorSaveItem});
            this.uch_godBindingNavigator.Location = new System.Drawing.Point(0, 0);
            this.uch_godBindingNavigator.MoveFirstItem = this.bindingNavigatorMoveFirstItem;
            this.uch_godBindingNavigator.MoveLastItem = this.bindingNavigatorMoveLastItem;
            this.uch_godBindingNavigator.MoveNextItem = this.bindingNavigatorMoveNextItem;
            this.uch_godBindingNavigator.MovePreviousItem = this.bindingNavigatorMovePreviousItem;
            this.uch_godBindingNavigator.Name = "uch_godBindingNavigator";
            this.uch_godBindingNavigator.PositionItem = this.bindingNavigatorPositionItem;
            this.uch_godBindingNavigator.Size = new System.Drawing.Size(665, 25);
            this.uch_godBindingNavigator.TabIndex = 0;
            this.uch_godBindingNavigator.Text = "bindingNavigator1";
            // 
            // bindingNavigatorMoveFirstItem
            // 
            this.bindingNavigatorMoveFirstItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.bindingNavigatorMoveFirstItem.Image = ((System.Drawing.Image)(resources.GetObject("bindingNavigatorMoveFirstItem.Image")));
            this.bindingNavigatorMoveFirstItem.Name = "bindingNavigatorMoveFirstItem";
            this.bindingNavigatorMoveFirstItem.RightToLeftAutoMirrorImage = true;
            this.bindingNavigatorMoveFirstItem.Size = new System.Drawing.Size(23, 22);
            this.bindingNavigatorMoveFirstItem.Text = "В начало";
            // 
            // bindingNavigatorMovePreviousItem
            // 
            this.bindingNavigatorMovePreviousItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.bindingNavigatorMovePreviousItem.Image = ((System.Drawing.Image)(resources.GetObject("bindingNavigatorMovePreviousItem.Image")));
            this.bindingNavigatorMovePreviousItem.Name = "bindingNavigatorMovePreviousItem";
            this.bindingNavigatorMovePreviousItem.RightToLeftAutoMirrorImage = true;
            this.bindingNavigatorMovePreviousItem.Size = new System.Drawing.Size(23, 22);
            this.bindingNavigatorMovePreviousItem.Text = "Предыдущий учебный год";
            // 
            // bindingNavigatorSeparator
            // 
            this.bindingNavigatorSeparator.Name = "bindingNavigatorSeparator";
            this.bindingNavigatorSeparator.Size = new System.Drawing.Size(6, 25);
            // 
            // bindingNavigatorPositionItem
            // 
            this.bindingNavigatorPositionItem.AccessibleName = "Position";
            this.bindingNavigatorPositionItem.AutoSize = false;
            this.bindingNavigatorPositionItem.Name = "bindingNavigatorPositionItem";
            this.bindingNavigatorPositionItem.Size = new System.Drawing.Size(50, 21);
            this.bindingNavigatorPositionItem.Text = "0";
            this.bindingNavigatorPositionItem.ToolTipText = "Current position";
            // 
            // bindingNavigatorCountItem
            // 
            this.bindingNavigatorCountItem.Name = "bindingNavigatorCountItem";
            this.bindingNavigatorCountItem.Size = new System.Drawing.Size(74, 22);
            this.bindingNavigatorCountItem.Text = "запись из {0}";
            this.bindingNavigatorCountItem.ToolTipText = "Total number of items";
            // 
            // bindingNavigatorSeparator1
            // 
            this.bindingNavigatorSeparator1.Name = "bindingNavigatorSeparator";
            this.bindingNavigatorSeparator1.Size = new System.Drawing.Size(6, 25);
            // 
            // bindingNavigatorMoveNextItem
            // 
            this.bindingNavigatorMoveNextItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.bindingNavigatorMoveNextItem.Image = ((System.Drawing.Image)(resources.GetObject("bindingNavigatorMoveNextItem.Image")));
            this.bindingNavigatorMoveNextItem.Name = "bindingNavigatorMoveNextItem";
            this.bindingNavigatorMoveNextItem.RightToLeftAutoMirrorImage = true;
            this.bindingNavigatorMoveNextItem.Size = new System.Drawing.Size(23, 22);
            this.bindingNavigatorMoveNextItem.Text = "Следующий учебный год";
            // 
            // bindingNavigatorMoveLastItem
            // 
            this.bindingNavigatorMoveLastItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.bindingNavigatorMoveLastItem.Image = ((System.Drawing.Image)(resources.GetObject("bindingNavigatorMoveLastItem.Image")));
            this.bindingNavigatorMoveLastItem.Name = "bindingNavigatorMoveLastItem";
            this.bindingNavigatorMoveLastItem.RightToLeftAutoMirrorImage = true;
            this.bindingNavigatorMoveLastItem.Size = new System.Drawing.Size(23, 22);
            this.bindingNavigatorMoveLastItem.Text = "В конец";
            // 
            // bindingNavigatorSeparator2
            // 
            this.bindingNavigatorSeparator2.Name = "bindingNavigatorSeparator";
            this.bindingNavigatorSeparator2.Size = new System.Drawing.Size(6, 25);
            // 
            // bindingNavigatorAddNewItem
            // 
            this.bindingNavigatorAddNewItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.bindingNavigatorAddNewItem.Image = ((System.Drawing.Image)(resources.GetObject("bindingNavigatorAddNewItem.Image")));
            this.bindingNavigatorAddNewItem.Name = "bindingNavigatorAddNewItem";
            this.bindingNavigatorAddNewItem.RightToLeftAutoMirrorImage = true;
            this.bindingNavigatorAddNewItem.Size = new System.Drawing.Size(23, 22);
            this.bindingNavigatorAddNewItem.Text = "Добавить новый учебный год";
            // 
            // uch_godBindingNavigatorSaveItem
            // 
            this.uch_godBindingNavigatorSaveItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.uch_godBindingNavigatorSaveItem.Image = ((System.Drawing.Image)(resources.GetObject("uch_godBindingNavigatorSaveItem.Image")));
            this.uch_godBindingNavigatorSaveItem.Name = "uch_godBindingNavigatorSaveItem";
            this.uch_godBindingNavigatorSaveItem.Size = new System.Drawing.Size(23, 22);
            this.uch_godBindingNavigatorSaveItem.Text = "Сохоанить все изменения";
            this.uch_godBindingNavigatorSaveItem.Click += new System.EventHandler(this.uch_godBindingNavigatorSaveItem_Click);
            // 
            // uch_godDataGridView
            // 
            this.uch_godDataGridView.AllowUserToAddRows = false;
            this.uch_godDataGridView.AllowUserToDeleteRows = false;
            this.uch_godDataGridView.AutoGenerateColumns = false;
            this.uch_godDataGridView.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
            this.uch_godDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.uch_godDataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn2,
            this.dataGridViewTextBoxColumn3,
            this.dataGridViewTextBoxColumn4,
            this.dataGridViewTextBoxColumn5});
            this.uch_godDataGridView.DataSource = this.uch_godBindingSource;
            this.uch_godDataGridView.Dock = System.Windows.Forms.DockStyle.Right;
            this.uch_godDataGridView.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.uch_godDataGridView.Location = new System.Drawing.Point(221, 25);
            this.uch_godDataGridView.Name = "uch_godDataGridView";
            this.uch_godDataGridView.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.uch_godDataGridView.Size = new System.Drawing.Size(444, 241);
            this.uch_godDataGridView.TabIndex = 1;
            // 
            // startLabel
            // 
            startLabel.AutoSize = true;
            startLabel.BackColor = System.Drawing.Color.Transparent;
            startLabel.Location = new System.Drawing.Point(12, 42);
            startLabel.Name = "startLabel";
            startLabel.Size = new System.Drawing.Size(145, 13);
            startLabel.TabIndex = 2;
            startLabel.Text = "Дата начала учебного года";
            // 
            // startDateTimePicker
            // 
            this.startDateTimePicker.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.uch_godBindingSource, "start", true));
            this.startDateTimePicker.Location = new System.Drawing.Point(15, 58);
            this.startDateTimePicker.Name = "startDateTimePicker";
            this.startDateTimePicker.Size = new System.Drawing.Size(200, 20);
            this.startDateTimePicker.TabIndex = 3;
            // 
            // finishLabel
            // 
            finishLabel.AutoSize = true;
            finishLabel.BackColor = System.Drawing.Color.Transparent;
            finishLabel.Location = new System.Drawing.Point(12, 98);
            finishLabel.Name = "finishLabel";
            finishLabel.Size = new System.Drawing.Size(163, 13);
            finishLabel.TabIndex = 4;
            finishLabel.Text = "Дата окончания учебного года";
            // 
            // finishDateTimePicker
            // 
            this.finishDateTimePicker.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.uch_godBindingSource, "finish", true));
            this.finishDateTimePicker.Location = new System.Drawing.Point(15, 114);
            this.finishDateTimePicker.Name = "finishDateTimePicker";
            this.finishDateTimePicker.Size = new System.Drawing.Size(200, 20);
            this.finishDateTimePicker.TabIndex = 5;
            // 
            // semestr1_finishLabel
            // 
            semestr1_finishLabel.AutoSize = true;
            semestr1_finishLabel.BackColor = System.Drawing.Color.Transparent;
            semestr1_finishLabel.Location = new System.Drawing.Point(12, 154);
            semestr1_finishLabel.Name = "semestr1_finishLabel";
            semestr1_finishLabel.Size = new System.Drawing.Size(185, 13);
            semestr1_finishLabel.TabIndex = 6;
            semestr1_finishLabel.Text = "Дата окончания первого семестра";
            // 
            // semestr1_finishDateTimePicker
            // 
            this.semestr1_finishDateTimePicker.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.uch_godBindingSource, "semestr1_finish", true));
            this.semestr1_finishDateTimePicker.Location = new System.Drawing.Point(15, 170);
            this.semestr1_finishDateTimePicker.Name = "semestr1_finishDateTimePicker";
            this.semestr1_finishDateTimePicker.Size = new System.Drawing.Size(200, 20);
            this.semestr1_finishDateTimePicker.TabIndex = 7;
            // 
            // semestr2_startLabel
            // 
            semestr2_startLabel.AutoSize = true;
            semestr2_startLabel.BackColor = System.Drawing.Color.Transparent;
            semestr2_startLabel.Location = new System.Drawing.Point(12, 214);
            semestr2_startLabel.Name = "semestr2_startLabel";
            semestr2_startLabel.Size = new System.Drawing.Size(166, 13);
            semestr2_startLabel.TabIndex = 8;
            semestr2_startLabel.Text = "Дата начала второго семестра";
            // 
            // semestr2_startDateTimePicker
            // 
            this.semestr2_startDateTimePicker.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.uch_godBindingSource, "semestr2_start", true));
            this.semestr2_startDateTimePicker.Location = new System.Drawing.Point(15, 230);
            this.semestr2_startDateTimePicker.Name = "semestr2_startDateTimePicker";
            this.semestr2_startDateTimePicker.Size = new System.Drawing.Size(200, 20);
            this.semestr2_startDateTimePicker.TabIndex = 9;
            // 
            // uch_godBindingSource
            // 
            this.uch_godBindingSource.DataMember = "uch_god";
            this.uch_godBindingSource.DataSource = this.vKRDataSet;
            // 
            // vKRDataSet
            // 
            this.vKRDataSet.DataSetName = "VKRDataSet";
            this.vKRDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // dataGridViewTextBoxColumn2
            // 
            this.dataGridViewTextBoxColumn2.DataPropertyName = "start";
            this.dataGridViewTextBoxColumn2.HeaderText = "Начало у.г.";
            this.dataGridViewTextBoxColumn2.Name = "dataGridViewTextBoxColumn2";
            // 
            // dataGridViewTextBoxColumn3
            // 
            this.dataGridViewTextBoxColumn3.DataPropertyName = "finish";
            this.dataGridViewTextBoxColumn3.HeaderText = "Конец у.г.";
            this.dataGridViewTextBoxColumn3.Name = "dataGridViewTextBoxColumn3";
            // 
            // dataGridViewTextBoxColumn4
            // 
            this.dataGridViewTextBoxColumn4.DataPropertyName = "semestr1_finish";
            dataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridViewTextBoxColumn4.DefaultCellStyle = dataGridViewCellStyle5;
            this.dataGridViewTextBoxColumn4.HeaderText = "Семестр 1 (конец)";
            this.dataGridViewTextBoxColumn4.Name = "dataGridViewTextBoxColumn4";
            // 
            // dataGridViewTextBoxColumn5
            // 
            this.dataGridViewTextBoxColumn5.DataPropertyName = "semestr2_start";
            dataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridViewTextBoxColumn5.DefaultCellStyle = dataGridViewCellStyle6;
            this.dataGridViewTextBoxColumn5.HeaderText = "Семестр 2 (начало)";
            this.dataGridViewTextBoxColumn5.Name = "dataGridViewTextBoxColumn5";
            // 
            // uch_godTableAdapter
            // 
            this.uch_godTableAdapter.ClearBeforeFill = true;
            // 
            // sprav_uch_god
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = global::Факультет.Properties.Resources.fon_бел_син_бледн;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.ClientSize = new System.Drawing.Size(665, 266);
            this.Controls.Add(semestr2_startLabel);
            this.Controls.Add(this.semestr2_startDateTimePicker);
            this.Controls.Add(semestr1_finishLabel);
            this.Controls.Add(this.semestr1_finishDateTimePicker);
            this.Controls.Add(finishLabel);
            this.Controls.Add(this.finishDateTimePicker);
            this.Controls.Add(startLabel);
            this.Controls.Add(this.startDateTimePicker);
            this.Controls.Add(this.uch_godDataGridView);
            this.Controls.Add(this.uch_godBindingNavigator);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "sprav_uch_god";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Редактировать учебный год";
            this.Load += new System.EventHandler(this.sprav_uch_god_Load);
            ((System.ComponentModel.ISupportInitialize)(this.uch_godBindingNavigator)).EndInit();
            this.uch_godBindingNavigator.ResumeLayout(false);
            this.uch_godBindingNavigator.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.uch_godDataGridView)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.uch_godBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.vKRDataSet)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private VKRDataSet vKRDataSet;
        private System.Windows.Forms.BindingSource uch_godBindingSource;
        private Факультет.VKRDataSetTableAdapters.uch_godTableAdapter uch_godTableAdapter;
        private System.Windows.Forms.BindingNavigator uch_godBindingNavigator;
        private System.Windows.Forms.ToolStripButton bindingNavigatorAddNewItem;
        private System.Windows.Forms.ToolStripLabel bindingNavigatorCountItem;
        private System.Windows.Forms.ToolStripButton bindingNavigatorMoveFirstItem;
        private System.Windows.Forms.ToolStripButton bindingNavigatorMovePreviousItem;
        private System.Windows.Forms.ToolStripSeparator bindingNavigatorSeparator;
        private System.Windows.Forms.ToolStripTextBox bindingNavigatorPositionItem;
        private System.Windows.Forms.ToolStripSeparator bindingNavigatorSeparator1;
        private System.Windows.Forms.ToolStripButton bindingNavigatorMoveNextItem;
        private System.Windows.Forms.ToolStripButton bindingNavigatorMoveLastItem;
        private System.Windows.Forms.ToolStripSeparator bindingNavigatorSeparator2;
        private System.Windows.Forms.ToolStripButton uch_godBindingNavigatorSaveItem;
        private System.Windows.Forms.DataGridView uch_godDataGridView;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn2;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn3;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn4;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn5;
        private System.Windows.Forms.DateTimePicker startDateTimePicker;
        private System.Windows.Forms.DateTimePicker finishDateTimePicker;
        private System.Windows.Forms.DateTimePicker semestr1_finishDateTimePicker;
        private System.Windows.Forms.DateTimePicker semestr2_startDateTimePicker;
    }
}