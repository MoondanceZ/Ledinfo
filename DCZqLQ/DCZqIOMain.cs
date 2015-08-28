using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using FI.Public;
namespace KY.Fi.DCZqLQ
{
    public partial class DCZqIOMain : Form
    {
        public int ExitFlg = 0;
        public DCZqIOMain()
        {
            InitializeComponent();
            setDataBaseConnect();
        }
        private void setDataBaseConnect()
        {
            string connstr = DBConn.GetConnStr(Const.DBConnFile);
            if (connstr != "" && DBConn.TestConnection(connstr))
            {
                DBConn.SetSqlConn(Const.DBConnFile);
            }

        }

        #region 菜单按钮事件
        private void notifyIcon_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                if (this.Visible)
                    this.Activate();
                this.Visible = true;
            }
        }
        //显示,隐藏
        private void SysTrayHideShow()
        {
            Visible = !Visible;
            if (Visible)
                Activate();
        }

        private void DCZqIOMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                this.Visible = false;
            }
        }
 
        private void exi_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("确定退出系统?", "提示", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                ExitFlg = 1;
                //隐藏托盘程序中的图标
                this.notifyIcon.Visible = false;
                //关闭系统
                Application.Exit();


            }
        }

        private void sysSet_Click(object sender, EventArgs e)
        {
            new SysSet().Show();
        }

 
        #endregion

        private void DCZqIOMain_Load(object sender, EventArgs e)
        {

        }



   

    }
}
