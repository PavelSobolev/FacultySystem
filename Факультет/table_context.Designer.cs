namespace FSystem
{
    partial class table_context
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
            this.prepod = new System.Windows.Forms.ComboBox();
            this.predmet = new System.Windows.Forms.ComboBox();
            this.vid_zan = new System.Windows.Forms.ComboBox();
            this.aud = new System.Windows.Forms.ComboBox();
            this.SuspendLayout();
            // 
            // prepod
            // 
            this.prepod.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.prepod.FormattingEnabled = true;
            this.prepod.Location = new System.Drawing.Point(12, 12);
            this.prepod.Name = "prepod";
            this.prepod.Size = new System.Drawing.Size(146, 21);
            this.prepod.TabIndex = 0;
            // 
            // predmet
            // 
            this.predmet.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.predmet.FormattingEnabled = true;
            this.predmet.Location = new System.Drawing.Point(164, 12);
            this.predmet.Name = "predmet";
            this.predmet.Size = new System.Drawing.Size(127, 21);
            this.predmet.TabIndex = 1;
            // 
            // vid_zan
            // 
            this.vid_zan.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.vid_zan.FormattingEnabled = true;
            this.vid_zan.Location = new System.Drawing.Point(297, 12);
            this.vid_zan.Name = "vid_zan";
            this.vid_zan.Size = new System.Drawing.Size(84, 21);
            this.vid_zan.TabIndex = 2;
            // 
            // aud
            // 
            this.aud.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.aud.FormattingEnabled = true;
            this.aud.Location = new System.Drawing.Point(387, 12);
            this.aud.Name = "aud";
            this.aud.Size = new System.Drawing.Size(57, 21);
            this.aud.TabIndex = 3;
            // 
            // table_context
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.ClientSize = new System.Drawing.Size(454, 46);
            this.Controls.Add(this.aud);
            this.Controls.Add(this.vid_zan);
            this.Controls.Add(this.predmet);
            this.Controls.Add(this.prepod);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "table_context";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.Text = "table_context";
            this.Deactivate += new System.EventHandler(this.table_context_Deactivate);
            this.MouseMove += new System.Windows.Forms.MouseEventHandler(this.table_context_MouseMove);
            this.MouseDown += new System.Windows.Forms.MouseEventHandler(this.table_context_MouseDown);
            this.ResumeLayout(false);

        }

        #endregion

        public System.Windows.Forms.ComboBox prepod;
        public System.Windows.Forms.ComboBox predmet;
        public System.Windows.Forms.ComboBox vid_zan;
        public System.Windows.Forms.ComboBox aud;





    }
}