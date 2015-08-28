using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.IO;
using System.Text.RegularExpressions;
using FI.Public;
using FI.DataAccess;

namespace KY.Fi.DCZqLQ
{
    public partial class SysSet : Form
    {
        private StartupManager startupManager = new StartupManager();
        public SysSet()
        {
            InitializeComponent();
        }
        private void SysSet_Load(object sender, System.EventArgs e)
        {            
            ReadSetXml(Const.DBConnFile);            
            this.RunAsStart.Checked = startupManager.Startup;
        }

        /// <summary>
        /// 读取配置文件
        /// </summary>
        /// <param name="file"></param>
        public void ReadSetXml(string file)//读取配置文件
        {
            try
            {
                FileInfo fileinfo = new FileInfo(file);
                XmlDocument myXml = new XmlDocument();
                myXml.Load(file);//读取指定的XML文档 
                XmlNode Mess = myXml.DocumentElement;//读取XML的根节点
                foreach (XmlNode node in Mess.ChildNodes)//对子节点进行循环 
                {
                    //将每个节点的内容显示出来 
                    switch (node.Name)
                    {
                        //设置
                        case "Server":
                            txtServer.Text = node.InnerText.ToString();
                            break;
                        case "DataBase":
                            this.txtDatabase.Text = node.InnerText.ToString();
                            break;
                        case "User":
                            this.txtUser.Text = node.InnerText.ToString();
                            break;
                        case "PassWord":
                            this.txtPsw.Text = node.InnerText.ToString();
                            break;
                        case "model":
                            cModel.Text = node.InnerText.ToString();
                            break;




                        //自动运行同步设置
                        case "jTimes":
                            jTimes.Text = node.InnerText.ToString();
                            break;
                        case "Func":
                            DataTable dtAllFunc = LqImportDac.GetLedInfoByConn();
                            int count = 0;
                            //DataTable dtSelectFunc = LqImportDac.GetLedInfo(node.InnerText.ToString());
                            if (dtAllFunc.Rows.Count > 0)
                            {
                                cmbFunc.DataSource = dtAllFunc;
                                cmbFunc.DisplayMember = "txt";
                                cmbFunc.ValueMember = "sht";
                                for (int i = 0; i < dtAllFunc.Rows.Count; i++)
                                {                                    
                                    if(dtAllFunc.Rows[i]["sht"].ToString()==node.InnerText.ToString())
                                    {
                                        cmbFunc.SelectedValue = dtAllFunc.Rows[i]["sht"].ToString();
                                        cmbFunc.SelectedText = dtAllFunc.Rows[i]["txt"].ToString();
                                        break;
                                    }
                                    else
                                        count++;
                                }           
                                if(count==dtAllFunc.Rows.Count)
                                {
                                    cmbFunc.SelectedIndex = -1;
                                }
                            }                            
                            break;
                    }
                }

            }
            catch { }
        }
        /// <summary>
        /// 创建打印XML文件
        /// </summary>
        /// <param name="file"></param>
        /// <param name="dtDetailMess"></param>
        public static void CreateXML(string file, string server, string database, string user, string pass, bool isRunAsStart, string model, string jTimes, string Func)
        {
            try
            {
                XmlTextWriter myWriter = new XmlTextWriter(file, Encoding.UTF8);
                myWriter.Formatting = Formatting.Indented;
                myWriter.WriteStartDocument(); //start Document
                myWriter.WriteStartElement("DataSet"); //start writer first Element
                //源数据库设置
                myWriter.WriteStartElement("Server");
                myWriter.WriteString(server);
                myWriter.WriteEndElement();

                myWriter.WriteStartElement("DataBase");
                myWriter.WriteString(database);
                myWriter.WriteEndElement();

                myWriter.WriteStartElement("User");
                myWriter.WriteString(user);
                myWriter.WriteEndElement();

                myWriter.WriteStartElement("PassWord");
                myWriter.WriteString(pass);
                myWriter.WriteEndElement();

                myWriter.WriteStartElement("model");
                myWriter.WriteString(model);
                myWriter.WriteEndElement();



                myWriter.WriteStartElement("IsRunAsStart");
                if (isRunAsStart)
                {
                    myWriter.WriteString("1");
                }
                else
                {
                    myWriter.WriteString("0");
                }
                myWriter.WriteEndElement();




                //导EXCEL设置
                myWriter.WriteStartElement("jTimes");
                myWriter.WriteString(jTimes);
                myWriter.WriteEndElement();

                myWriter.WriteStartElement("Func");
                myWriter.WriteString(Func);
                myWriter.WriteEndElement();

                //myWriter.WriteStartElement("excelPath");
                //myWriter.WriteString(excelPath);
                //myWriter.WriteEndElement();

                myWriter.WriteEndElement(); //end writer first Element

                myWriter.WriteEndDocument(); //end Document
                myWriter.Flush();
                myWriter.Close();
            }
            catch { }
        }
        private void btnCancle_Click(object sender, System.EventArgs e)
        {
            this.Close();
        }

        private void btnOk_Click(object sender, System.EventArgs e)
        {


            FileInfo fileinfo = new FileInfo(Const.DBConnFile);
            if (fileinfo.Exists)
            {
                if (fileinfo.Attributes.ToString().IndexOf("ReadOnly") != -1)
                    fileinfo.Attributes = FileAttributes.Normal;
                File.Delete(Const.DBConnFile);
            }



            CreateXML(Const.DBConnFile, txtServer.Text, txtDatabase.Text, txtUser.Text, txtPsw.Text,
                this.RunAsStart.Checked, cModel.Text, jTimes.Text.Trim(), cmbFunc.SelectedValue.ToString());

            startupManager.Startup = this.RunAsStart.Checked;
            MessageBox.Show("保存成功!");

            System.Threading.Thread.Sleep(200);
            Application.Exit();
            System.Diagnostics.Process.Start(Application.ExecutablePath);
            //System.Environment.Exit(0);
        }

        private void BindDataCmbFunc()
        {
            DataTable dtSelectFunc = LqImportDac.GetLedInfoByConn();
            cmbFunc.DataSource = dtSelectFunc;
            cmbFunc.DisplayMember = "txt";
            cmbFunc.ValueMember = "sht";
            cmbFunc.SelectedIndex = -1;
        }

        private void testConnect_Click(object sender, System.EventArgs e)
        {
            string connstr = String.Format("Data Source={0};Initial Catalog={1};User ID={2};Password={3}", txtServer.Text.Trim(), txtDatabase.Text.Trim(), txtUser.Text.Trim(), this.txtPsw.Text.Trim());
            //string connstr = String.Format("Data Source={0};User ID={1};Password={2}", txtDatabase.Text.Trim(), txtUser.Text.Trim(), this.txtPsw.Text.Trim());            
            if (connstr != "" && DBConn.TestConnection(connstr))
            {
                DBConn.SetSqlConn(Const.DBConnFile);
                if (cmbFunc.Items.Count == 0)
                {
                    DBConn.SetConnStr(connstr);
                    BindDataCmbFunc();
                }
                MessageBox.Show("连接成功!");
            }
            else
            {
                MessageBox.Show("连接不上服务器，请重新配置!");
            }
        }


        private void jTimes_TextChanged(object sender, System.EventArgs e)
        {
            if ((!Regex.IsMatch(((TextBox)sender).Text, "^[0-9]\\d*$")) && ((TextBox)sender).Text != "")
            {
                MessageBox.Show("请输入正整数!");
                ((TextBox)sender).Text = "";
            }
            else
            {
                if (Convert.ToInt32(((TextBox)sender).Text.Trim()) > 23)
                {
                    MessageBox.Show("分钟不对");
                    ((TextBox)sender).Text = "0";
                }
            }
        }

        private void btnFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Multiselect = true;
            fileDialog.Title = "请选择文件";
            fileDialog.Filter = "所以文件(*.*)|*.*";
            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                //excelPath.Text = fileDialog.FileName;
                string[] pathArray = fileDialog.FileNames;
                if (pathArray.Length > 1)
                {
                    foreach (var path in pathArray)
                    {
                        excelPath.Text += path + ";";
                    }
                }
                else
                    excelPath.Text = pathArray[0];
            }
        }









    }
}
