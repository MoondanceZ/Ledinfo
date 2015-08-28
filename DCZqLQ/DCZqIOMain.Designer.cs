namespace KY.Fi.DCZqLQ
{
    partial class DCZqIOMain
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(DCZqIOMain));
            this.notifyIcon = new System.Windows.Forms.NotifyIcon(this.components);
            this.notifyiconMnu = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.sysSet = new System.Windows.Forms.ToolStripMenuItem();
            this.exi = new System.Windows.Forms.ToolStripMenuItem();
            this.notifyiconMnu.SuspendLayout();
            this.SuspendLayout();
            // 
            // notifyIcon
            // 
            this.notifyIcon.ContextMenuStrip = this.notifyiconMnu;
            this.notifyIcon.Icon = ((System.Drawing.Icon)(resources.GetObject("notifyIcon.Icon")));
            this.notifyIcon.Text = "财务接口系统";
            this.notifyIcon.Visible = true;
            this.notifyIcon.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.notifyIcon_MouseDoubleClick);
            // 
            // notifyiconMnu
            // 
            this.notifyiconMnu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.sysSet,
            this.exi});
            this.notifyiconMnu.Name = "contextMenuStrip1";
            this.notifyiconMnu.Size = new System.Drawing.Size(153, 70);
            // 
            // sysSet
            // 
            this.sysSet.Name = "sysSet";
            this.sysSet.Size = new System.Drawing.Size(152, 22);
            this.sysSet.Text = "系统设置";
            this.sysSet.Click += new System.EventHandler(this.sysSet_Click);
            // 
            // exi
            // 
            this.exi.Name = "exi";
            this.exi.Size = new System.Drawing.Size(152, 22);
            this.exi.Text = "退出";
            this.exi.Click += new System.EventHandler(this.exi_Click);
            // 
            // DCZqIOMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(243, 166);
            this.Name = "DCZqIOMain";
            this.Text = "FinaceIOMain";
            this.Load += new System.EventHandler(this.DCZqIOMain_Load);
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.DCZqIOMain_FormClosing);
            this.notifyiconMnu.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        public System.Windows.Forms.NotifyIcon notifyIcon;
        private System.Windows.Forms.ContextMenuStrip notifyiconMnu;
        private System.Windows.Forms.ToolStripMenuItem sysSet;
        private System.Windows.Forms.ToolStripMenuItem exi;
    }
}