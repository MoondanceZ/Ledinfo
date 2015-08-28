using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.IO;


using KY.AP.DataServices;
using KY.Data.DB;
namespace FI.Public
{
  public class CBase
    {

        #region AddErrorLog 写错误日志
        /// <summary>
        /// 写报错日志
        /// </summary>
        /// <param name="txt"></param>
        public static void AddErroLog(string txt)
        {
            try
            {
                string sFileName = Const.StartupPath+ "\\log\\errolog.txt";
                if (!Directory.Exists(Const.StartupPath + "\\log"))
                {
                    Directory.CreateDirectory(Const.StartupPath + "\\log");
                }
                if (!System.IO.File.Exists(sFileName))
                {
                    System.IO.FileStream stream = System.IO.File.Create(sFileName);
                    stream.Close();
                }
                else  //lzl 如果文件超过10M，然后删除此文件，并重新创建
                {
                    FileInfo finfo = new FileInfo(sFileName);
                    if (finfo.Exists && finfo.Length > 10240000)
                    {
                        finfo.Delete();
                    }
                    if (!System.IO.File.Exists(sFileName))
                    {
                        System.IO.FileStream stream = System.IO.File.Create(sFileName);
                        stream.Close();
                    }
                }

                try
                {
                    using (FileStream fs = new FileStream(sFileName, FileMode.Append, FileAccess.Write))
                    {
                        StreamWriter sw = new StreamWriter(fs);
                        sw.WriteLine(txt);
                        sw.Close();
                    }
                }
                catch (Exception ex)
                {

                }
            }
            catch(Exception ex)
            {
                string error = ex.ToString();
            }
        }
        #endregion

        #region AddFile 写文件
        /// <summary>
        /// 写文件
        /// </summary>
        /// <param name="txt"></param>
        public static void AddFile(string filePath,string fileName,string txt,int typ)
        {
            try
            {
                int firstFlg = 0;//是否是刚创建文件
                string sFileName = filePath + "\\"+fileName;
                if (!Directory.Exists(filePath))
                {
                    Directory.CreateDirectory(filePath);
                }
                if (!System.IO.File.Exists(sFileName))
                {
                    System.IO.FileStream stream = System.IO.File.Create(sFileName);
                    stream.Close();
                    firstFlg = 1;

                }
                else  //lzl 如果文件超过100M，然后删除此文件，并重新创建
                {
                    FileInfo finfo = new FileInfo(sFileName);
                    if (finfo.Exists && finfo.Length > 102400000)
                    {
                        finfo.Delete();
                    }
                    if (!System.IO.File.Exists(sFileName))
                    {
                        System.IO.FileStream stream = System.IO.File.Create(sFileName);
                        stream.Close();
                        firstFlg = 1;
                    }
                }

                try
                {
                    using (FileStream fs = new FileStream(sFileName, FileMode.Append, FileAccess.Write))
                    {
                        StreamWriter sw = new StreamWriter(fs);
                        //如果是首次创建文件且是下载图片到考勤机时候，需要写文件第一行为表头
                        if (firstFlg == 1 && typ == 1)
                        {
                            sw.WriteLine("#WDDA$,xh,bh-NO,xm-Name,mm,gl,sr,zp,kh,zw,mj-1");
                        }
                        sw.WriteLine(txt);
                        sw.Close();
                    }
                }
                catch (Exception ex)
                {

                }
            }
            catch (Exception ex)
            {
                string error = ex.ToString();
            }
        }
        #endregion

        #region GetShgCod 获取系统编号
        /// <summary>
        /// 获取编号
        /// </summary>
        /// <param name="name">名称</param>
        /// <param name="add">递增标志(0为不递增,1为递增)</param>
        /// <param name="trans">0开启事务,1不开启事务</param>
        /// <returns></returns>
        public static string GetShgCod(string name, int add, int trans)
        {

                string procedure = "call usp_ep_GetShgCod";
                KYDataParameter[] parameter=new KYDataParameter [4];
                parameter[0]=new KYDataParameter("name", KYDbType.VarChar, 50);
                parameter[0].Value = name;

                parameter[1] = new KYDataParameter("billCode",KYDbType.VarChar, 50,System.Data.ParameterDirection .Output);

                parameter[2] = new KYDataParameter("add", KYDbType.SmallInt );
                parameter[2].Value = add;

                parameter[3] = new KYDataParameter("trans", KYDbType.SmallInt);
                parameter[3].Value = trans;

                DbHelper.ExecuteNonQuery(procedure, parameter);
                string billCode = parameter[1].Value.ToString();

                return billCode;
            
        }
        #endregion
        /// <summary>
        /// 获取全局参数
        /// </summary>
        /// <returns></returns>
        public static void GetSysParam(string file)
        {
            try
            {
                XmlDocument xd;
                XmlNode nd;
                if (File.Exists(file))
                {
                    xd = new XmlDocument();
                    xd.Load(file);

                    nd = xd["DataSet"];
                    //自动运行设置
                    Const.IsRunAsStart = nd["IsRunAsStart"].InnerText.ToString();
                    Const.model = nd["model"].InnerText.ToString();

                    if (nd["sminute"] != null)
                    {
                        Const.sminute = Convert.ToInt32(nd["sminute"].InnerText.ToString());
                    }
                    if (nd["RSTime"] != null)
                    {
                        Const.RSTime = nd["RSTime"].InnerText.ToString();
                    }
                    if (nd["stime"] != null)
                    {
                        Const.stime = Convert.ToInt32(nd["stime"].InnerText.ToString());
                    }

                    //开源财务软件科目设置
                    if (nd["Ycl"] != null)
                    {
                        Const.Ycl = nd["Ycl"].InnerText.ToString();
                    }
                    if (nd["Yfzk"] != null)
                    {
                        Const.Yfzk = nd["Yfzk"].InnerText.ToString();
                    }
                    if (nd["YfzkHs"] != null)
                    {
                        Const.YfzkHs = nd["YfzkHs"].InnerText.ToString();
                    }
                    if (nd["Yjsj"] != null)
                    {
                        Const.Yjsj = nd["Yjsj"].InnerText.ToString();
                    }
                    if (nd["YjsjHs"] != null)
                    {
                        Const.YjsjHs = nd["YjsjHs"].InnerText.ToString();
                    }
                    //乐其自动运行同步设置

                    if (nd["LMdateLQ"] != null)
                    {
                        Const.LMdateLQ = nd["LMdateLQ"].InnerText.ToString();
                    }
                    if (nd["IsMonthLQ"] != null)
                    {
                        Const.IsMonthLQ = nd["IsMonthLQ"].InnerText.ToString();
                    }
                    if (nd["jTimes"] != null)
                    {
                        Const.jTimes =  Convert.ToInt32(nd["jTimes"].InnerText.ToString());
                    }                    
                    if (nd["Func"] != null)
                    {
                        Const.Func = nd["Func"].InnerText.ToString();
                    }
                }
              }
            catch { }
        } 

    }
}
 
