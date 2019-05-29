namespace FSystem
{
    partial class messagebox
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose ( bool disposing )
        {
            if ( disposing && ( components != null ) )
            {
                components.Dispose ( );
            }
            base.Dispose ( disposing );
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent ( )
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(messagebox));
            this.rsp = new C1.Win.C1FlexGrid.Classic.C1FlexGridClassic();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.prepod_name = new System.Windows.Forms.ToolStripStatusLabel();
            ((System.ComponentModel.ISupportInitialize)(this.rsp)).BeginInit();
            this.statusStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // rsp
            // 
            this.rsp.AllowUserResizing = C1.Win.C1FlexGrid.Classic.AllowUserResizeSettings.flexResizeColumns;
            this.rsp.BackColor = System.Drawing.Color.Transparent;
            this.rsp.BackColorAlternate = System.Drawing.Color.AliceBlue;
            this.rsp.BackColorBkg = System.Drawing.SystemColors.Control;
            this.rsp.BackColorFixed = System.Drawing.SystemColors.Control;
            this.rsp.BackColorSel = System.Drawing.SystemColors.Highlight;
            this.rsp.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.rsp.ColumnInfo = resources.GetString("rsp.ColumnInfo");
            this.rsp.Dock = System.Windows.Forms.DockStyle.Fill;
            this.rsp.FixedCols = 0;
            this.rsp.FixedRows = 0;
            this.rsp.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.rsp.ForeColor = System.Drawing.SystemColors.WindowText;
            this.rsp.ForeColorFixed = System.Drawing.SystemColors.ControlText;
            this.rsp.ForeColorSel = System.Drawing.SystemColors.HighlightText;
            this.rsp.FrozenCols = 1;
            this.rsp.FrozenRows = 1;
            this.rsp.GridColor = System.Drawing.Color.LightSteelBlue;
            this.rsp.GridColorFixed = System.Drawing.SystemColors.ControlDark;
            this.rsp.Location = new System.Drawing.Point(0, 0);
            this.rsp.Name = "rsp";
            this.rsp.NodeClosedPicture = null;
            this.rsp.NodeOpenPicture = null;
            this.rsp.OutlineBar = C1.Win.C1FlexGrid.Classic.OutlineBarSettings.flexOutlineBarSimple;
            this.rsp.OutlineCol = -1;
            this.rsp.RowHeightMax = 40;
            this.rsp.RowHeightMin = 40;
            this.rsp.Rows = 7;
            this.rsp.SheetBorder = System.Drawing.SystemColors.ControlDarkDark;
            this.rsp.Size = new System.Drawing.Size(630, 315);
            this.rsp.TabIndex = 0;
            this.rsp.TreeColor = System.Drawing.Color.DarkGray;
            this.rsp.WallPaper = global::FSystem.Properties.Resources.tab_fon;
            this.rsp.WordWrap = true;
            this.rsp.MouseDown += new System.Windows.Forms.MouseEventHandler(this.rsp_MouseDown);
            this.rsp.KeyDown += new System.Windows.Forms.KeyEventHandler(this.rsp_KeyDown);
            this.rsp.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.rsp_MouseDoubleClick);
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.prepod_name});
            this.statusStrip1.Location = new System.Drawing.Point(0, 293);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(630, 22);
            this.statusStrip1.TabIndex = 1;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // prepod_name
            // 
            this.prepod_name.Image = global::FSystem.Properties.Resources.user_male24;
            this.prepod_name.Name = "prepod_name";
            this.prepod_name.Size = new System.Drawing.Size(89, 17);
            this.prepod_name.Text = "prepod_name";
            // 
            // messagebox
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.ClientSize = new System.Drawing.Size(630, 315);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.rsp);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "messagebox";
            this.Opacity = 0.95;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "   Распределение часов на ";
            ((System.ComponentModel.ISupportInitialize)(this.rsp)).EndInit();
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private C1.Win.C1FlexGrid.Classic.C1FlexGridClassic rsp;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel prepod_name;

    }
}