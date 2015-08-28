using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Xml;
using System.Data;
using KY.Data.DB;
using KY.AP.DataServices;
namespace FI.Public
{
    public class DBConn
    {
        /// <summary>
        /// 获取连接字符串
        /// </summary>
        /// <returns></returns>
        public static string GetConnStr(string file)
        {
            try
            {
                XmlDocument xd;
                XmlNode nd;
                string db = "", server = "", user = "", psw = "";
                if (File.Exists(file))
                {
                    xd = new XmlDocument();
                    xd.Load(file);

                    nd = xd["DataSet"];
                    db = nd["DataBase"].InnerText.ToString();
                    server = nd["Server"].InnerText.ToString();
                    user = nd["User"].InnerText.ToString();
                    psw = nd["PassWord"].InnerText.ToString();

                    if (db == "" || server == "" || user == "")
                    {
                        return "";
                    }
                }
                else
                {
                    return "";
                }
                string connstr = "Data Source=" + server + ";Initial Catalog=" + db + ";User ID=" + user + ";Password=" + psw;
               // string connstr = "Data Source=" + db + ";User ID=" + user + ";Password=" + psw;
                Const.SqlConnentionString = connstr;
                Const.DataBaseName = db;
                return connstr;
            }
            catch { return ""; }
        }
        /// <summary>
        /// 同步数据时，源数据库连接字符串
        /// </summary>
        /// <returns></returns>
        public static string GetConnStr1(string file)
        {
            try
            {
                XmlDocument xd;
                XmlNode nd;
                string db = "", server = "", user = "", psw = "";
                if (File.Exists(file))
                {
                    xd = new XmlDocument();
                    xd.Load(file);

                    nd = xd["DataSet"];
                    db = nd["DataBase1"].InnerText.ToString();
                    server = nd["Server1"].InnerText.ToString();
                    user = nd["User1"].InnerText.ToString();
                    psw = nd["PassWord1"].InnerText.ToString();

                    if (db == "" || server == "" || user == "")
                    {
                        return "";
                    }
                }
                else
                {
                    return "";
                }
                string connstr = "ConnectType=SqlClient;Data Source=" + server + ";Initial Catalog=" + db + ";User ID=" + user + ";Password=" + psw;
                //string connstr = "ConnectType=SqlClient;Data Source=" + db + ";User ID=" + user + ";Password=" + psw;
                Const.SqlConnentionString1 = connstr;
                return connstr;
            }
            catch { return ""; }
        }

        public static void SetSqlConn(string file)
        {
            string errReturn = "d";
            string connstr = GetConnStr(file);
            if (connstr != "")
            {
                errReturn = KY.Data.DB.KYDbFactory.TestConnection(connstr);

                if (errReturn == "")
                {
                    connstr = "ConnectType=SqlClient;" + connstr;
                }
            }
            KY.AP.DataServices.ConnectionInfoService.SetSessionConnectString(connstr);
        }
        /// <summary>
        /// 测试连接
        /// </summary>
        /// <param name="connstr"></param>
        /// <returns></returns>
        public static bool TestConnection(string connstr)
        {
            string errReturn = "d";
            errReturn = KY.Data.DB.KYDbFactory.TestConnection(connstr);
            if (errReturn == "")
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        public static void SetSessionConnStr(string connstr)
        {
            connstr = "ConnectType=SqlClient;" + connstr;
            Const.SqlConnentionString = connstr;
            KY.AP.DataServices.ConnectionInfoService.SetSessionConnectString(connstr);
        }

        public static void SetConnStr(string connstr)
        {
            connstr = "ConnectType=SqlClient;" + connstr;
            Const.SqlConnentionString = connstr;
            KY.AP.DataServices.ConnectionInfoService.SetSessionConnectString(connstr);
        }
        /// <summary>
        /// DbHelper.Open
        /// </summary>
        public static void DbHelperOpen()
        {
            try
            {
                if (String.IsNullOrEmpty(KY.AP.DataServices.ConnectionInfoService.GetConnectString()))
                {
                    SetSessionConnStr(Const.SqlConnentionString);
                }
                if (TestConnection(Const.SqlConnentionString))
                {
                    DbHelper.Open();
                    Const.NoInterNet = false;
                }
                else
                {
                    Const.NoInterNet = true;
                }
            }
            catch
            {
                if (TestConnection(Const.SqlConnentionString))
                {
                    if (String.IsNullOrEmpty(KY.AP.DataServices.ConnectionInfoService.GetSessionConnectString()))
                    {
                        SetSessionConnStr(Const.SqlConnentionString);
                    }
                    DbHelper.Open();
                    Const.NoInterNet = false;
                }
                else
                {
                    Const.NoInterNet = true;
                }
            }
        }
        /// <summary>
        /// 获取结果集
        /// </summary>
        /// <param name="sql"></param>
        /// <returns></returns>
        public static DataTable GetSqlData(string sql)
        {
            DataTable dt = new DataTable();
            try
            {
                DbHelperOpen();
                if (!Const.NoInterNet)
                {
                    DbHelper.Fill(dt, sql);
                    DbHelper.Close();
                }
                return dt;
            }
            catch (Exception ex)
            {
                //CBase.AddErroLog("DBConn.GetSqlData:(" + DateTime.Now.ToString() + ")" + ex.ToString()); 
                return dt;
            }
        }
        /// <summary>
        /// 获取结果集
        /// </summary>
        /// <param name="sql"></param>
        /// <returns></returns>
        public static DataTable GetSqlData(string sql, KYDataParameter[] paramerers)
        {
            DataTable dt = new DataTable();
            try
            {
                DbHelperOpen();
                if (!Const.NoInterNet)
                {
                    DbHelper.Fill(dt, sql, paramerers);
                    DbHelper.Close();
                }
                return dt;
            }
            catch { return dt; }
        }
    }
}
