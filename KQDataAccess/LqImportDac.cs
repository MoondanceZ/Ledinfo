using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Data;
using System.IO;

using FI.Public;
using KY.AP.DataServices;
using KY.Data.DB;
using System.Data.SqlClient;

namespace FI.DataAccess
{
    public class LqImportDac
    {

        /// <summary>
        /// 获取领料出库单
        /// </summary>
        /// <returns></returns>
        public static DataTable GetOutBill(string year, string month, string prjSht, string fdate)
        {
            DBConn.DbHelperOpen();
            DataTable dt = new DataTable();
            StringBuilder sqlBuilder = new StringBuilder(1000);

            try
            {
                sqlBuilder.Append("SELECT ICStockBill.FInterID, ICStockBill.Fdate, ICStockBill.Fcheckdate,t001.FNumber as  FDeptIDNumber, t001.FName as  FDeptIDName,ICStockBill.FBillNo,");
                sqlBuilder.Append("ICStockBill.FHeadSelfB0433, ICStockBill.FHeadSelfB0434, ICStockBill.FHeadSelfB0435,t032.FNumber as  FHeadSelfB0435Number, t032.FName as  FHeadSelfB0435Name,");
                sqlBuilder.Append("ICStockBill.FHeadSelfB0436,t033.FNumber as  FHeadSelfB0436Number, t033.FName as  FHeadSelfB0436Name,ICStockBill.FHeadSelfB0437,");
                sqlBuilder.Append("t034.FNumber as  FHeadSelfB0437Number, t034.FName as  FHeadSelfB0437Name, ICStockBill.FHeadSelfB0438,ICStockBill.FTranType,");
                sqlBuilder.Append("ICStockBill.FCancellation, ICStockBill.FStatus, ICStockBill.FUpStockWhenSave,ICStockBill.FROB, ICStockBill.FHookStatus");
                sqlBuilder.Append(" FROM ICStockBill left join t_Department  t001 on ICStockBill.FDeptID= t001.FItemID  AND t001.FItemID<>0");
                sqlBuilder.Append(" left join t_Item  t032 on ICStockBill.FHeadSelfB0435= t032.FItemID  AND t032.FItemID<>0 ");
                sqlBuilder.Append(" left join t_Item  t033 on ICStockBill.FHeadSelfB0436= t033.FItemID  AND t033.FItemID<>0");
                sqlBuilder.Append(" left join t_Emp  t034 on ICStockBill.FHeadSelfB0437= t034.FItemID  AND t034.FItemID<>0");
                sqlBuilder.Append(" WHERE icstockbill.fstatus=1 and icstockbill.FHeadSelfB0438<>'' and icstockbill.FHeadSelfB0438 is not null and  icstockbill.ftrantype=24");
                if (year != "" && month != "")
                {
                    sqlBuilder.AppendFormat("  and year(icstockbill.fdate)={0} and month(icstockbill.fdate)={1}", year, month);
                }
                if (prjSht != "")
                {
                    sqlBuilder.AppendFormat(" and ICStockBill.FHeadSelfB0438 like '%{0}%'", prjSht);
                }
                if (fdate != "")
                {
                    sqlBuilder.AppendFormat("  and icstockbill.fdate='{0}'", fdate);
                }
                sqlBuilder.Append(" order by ICStockBill.FInterID");

                //sqlBuilder.Append("SELECT ICStockBill.FInterID, ICStockBill.Fdate,ICStockBill.Fdate fcheckdate, ICStockBill.FDeptID, t001.FNumber as  FDeptIDNumber, t001.FName as  FDeptIDName,");
                //sqlBuilder.Append("ICStockBill.FBillNo, ICStockBill.FHeadSelfB0432 FHeadSelfB0433,'mcnsht'  FHeadSelfB0435Number, 'mcntxt'  FHeadSelfB0435Name,'bantxt' as  FHeadSelfB0436Name,");
                //sqlBuilder.Append("'optsht' as  FHeadSelfB0437Number, 'opttxt'  FHeadSelfB0437Name, ICStockBill.FHeadSelfB0433 FHeadSelfB0438, ICStockBill.FBrNo,");
                //sqlBuilder.Append("ICStockBill.FTranType, ICStockBill.FCancellation, ICStockBill.FStatus, ICStockBill.FUpStockWhenSave ");
                //sqlBuilder.Append("  FROM ICStockBill left join t_Department  t001 on ICStockBill.FDeptID= t001.FItemID  AND t001.FItemID<>0 ");
                //sqlBuilder.Append("  WHERE  icstockbill.fstatus=1 and icstockbill.ftrantype=24");
                //if (year != "" && month != "")
                //{
                //    sqlBuilder.AppendFormat("   and year(icstockbill.fdate)={0} and month(icstockbill.fdate)={1}", year, month);
                //}
                //if (prjSht != "")
                //{
                //    sqlBuilder.AppendFormat(" and ICStockBill.FHeadSelfB0433  like '%{0}%'", prjSht);
                //}
                //if (fdate != "")
                //{
                //    sqlBuilder.AppendFormat("  and icstockbill.fdate='{0}'", fdate);
                //}
                //sqlBuilder.Append(" order by ICStockBill.FInterID");

                dt = DBConn.GetSqlData(sqlBuilder.ToString());
            }
            catch (System.Exception ex)
            {
                CBase.AddErroLog(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  导入异常:" + ex.ToString());
            }
            return dt;
        }

        //判断此色号是否在在线采集中存在
        private static string IsExistPrj(string prjsht, string clr)
        {
            DBConn.GetConnStr1(Const.DBConnFile);
            string connstr = Const.SqlConnentionString1;
            string prj = "";

            StringBuilder sqlBuilder = new StringBuilder(200);
            sqlBuilder.AppendFormat("select prjsht from sd_prj where sht='{0}' and clr='{1}'", prjsht.Trim(), clr.Trim());
            object o = DbHelper.ExecuteScalar(connstr, sqlBuilder.ToString());
            if (o != null)
            {
                prj = o.ToString();
            }
            return prj;
        }

        /// <summary>
        /// 导数据到开源pp_mus
        /// </summary>
        /// <returns></returns>
        public static bool ImportData(string finterid, string fbillno, string prjsht, string clr, out string ret)
        {
            ret = "";
            string kyPrjSht = IsExistPrj(prjsht, clr);
            if (string.IsNullOrEmpty(kyPrjSht))
            {
                ret = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  导入异常,生产单:" + prjsht + " 色号:" + clr + " 在在线采集系统中不存在";
                CBase.AddErroLog(ret);
                return false;
            }

            StringBuilder sqlBuilder = new StringBuilder(1000);
            sqlBuilder.Append(" SELECT  0 id,isb.fbillno bll,t002.FNumber as  mtr,t002.FName+'('+t002.Fmodel+')' as mtr_txt,t002.FNumber as  mtr_sht, isbe.FBatchNo as bnm,");
            sqlBuilder.Append(" '' pss,'' flg,'import' exa,t033.FName as banTxt, t034.FName as  opt_txt,isb.Fdate  dat,isb.Fcheckdate  dt1,1000*isbe.FQty qty, isbe.Famount prs,isbe.Fauxprice ppr,");
            sqlBuilder.Append(" isb.Fdate pprDat,t001.FNumber as  dpt, t001.FName as  dpt_txt,'' wkpsht,'' wkp_txt,0 wkpseq,t032.FNumber as mcnsht,'' potsht,isb.FInterID mov,");
            sqlBuilder.Append(" isb.FHeadSelfB0438 codsht,'' prjsht,isb.FHeadSelfB0433 prj_clr,'Rl' coTypSht,'02' typ,0 spt,0 cls,0 zflg");
            sqlBuilder.Append(" FROM ICStockBillEntry  isbe join t_ICItem  t002 on t002.FItemID = isbe.FItemID AND t002.FItemID <>0");
            sqlBuilder.Append(" inner join  ICStockBill isb  on isbe.finterid=isb.finterid");
            sqlBuilder.Append(" left join t_Department  t001 on isb.FDeptID= t001.FItemID  AND t001.FItemID<>0");
            sqlBuilder.Append(" left join t_Item  t032 on isb.FHeadSelfB0435= t032.FItemID  AND t032.FItemID<>0");
            sqlBuilder.Append(" left join t_Item  t033 on isb.FHeadSelfB0436= t033.FItemID  AND t033.FItemID<>0");
            sqlBuilder.Append(" left join t_Emp  t034 on isb.FHeadSelfB0437= t034.FItemID  AND t034.FItemID<>0");
            sqlBuilder.AppendFormat(" WHERE isbe.FInterID={0}", finterid);
            sqlBuilder.Append(" ORDER BY isbe.FEntryID");

            //sqlBuilder.Append(" SELECT  0 id,isb.fbillno bll,t002.FNumber as  mtr,t002.FName+'('+t002.Fmodel+')' as mtr_txt,t002.FNumber as  mtr_sht, isbe.FBatchNo as bnm,");
            //sqlBuilder.Append(" '' pss,'' flg,'import' exa,'' as banTxt, '' as  opt_txt,isb.Fdate  dat,isb.Fdate  dt1,1000*isbe.FQty qty, isbe.Famount prs,isbe.Fauxprice ppr,");
            //sqlBuilder.Append(" isb.Fdate pprDat,t001.FNumber as  dpt, t001.FName as  dpt_txt,'' wkpsht,'' wkp_txt,0 wkpseq,'' as mcnsht,'' potsht,isb.FInterID mov,");
            //sqlBuilder.Append(" isb.FHeadSelfB0433 codsht,'' prjsht,isb.FHeadSelfB0432 prj_clr,'Rl' coTypSht,'02' typ,0 spt,0 cls,0 zflg");
            //sqlBuilder.Append(" FROM ICStockBillEntry  isbe join t_ICItem  t002 on t002.FItemID = isbe.FItemID AND t002.FItemID <>0");
            //sqlBuilder.Append(" inner join  ICStockBill isb  on isbe.finterid=isb.finterid");
            //sqlBuilder.Append(" left join t_Department  t001 on isb.FDeptID= t001.FItemID  AND t001.FItemID<>0");
            //sqlBuilder.AppendFormat(" WHERE isbe.FInterID={0}", finterid);
            //sqlBuilder.Append(" ORDER BY isbe.FEntryID");

            DataTable dt = new DataTable();
            try
            {
                dt = DbHelper.GetDataTable(sqlBuilder.ToString());
                int rows = dt.Rows.Count;
                for (int i = 0; i < rows; i++)
                {
                    dt.Rows[i]["prjsht"] = kyPrjSht;
                    dt.Rows[i].AcceptChanges();
                    dt.Rows[i].SetAdded();
                }
            }
            catch (System.Exception ex)
            {
                CBase.AddErroLog(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  导入异常:" + ex.ToString());
            }
            DBConn.GetConnStr1(Const.DBConnFile);
            string connstr = Const.SqlConnentionString1;
            try
            {
                DbHelper.Open(connstr);
                DbHelper.BeginTran(connstr);
                sqlBuilder.Remove(0, sqlBuilder.Length);
                sqlBuilder.AppendFormat("delete from pp_mus where  exa='import' and  bll='{0}'  and mov='{1}';", fbillno, finterid);
                DbHelper.ExecuteNonQuery(connstr, sqlBuilder.ToString());
                DbHelper.Update(connstr, dt, "select * from pp_mus where 1<>1");
                DbHelper.CommitTran(connstr);
                return true;
            }
            catch (System.Exception ex)
            {
                DbHelper.RollbackTran(connstr);
                ret = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  导入异常:" + ex.ToString();
                CBase.AddErroLog(ret);
                return false;
            }
            finally
            {
                DbHelper.Close(connstr);
            }
        }
        /// <summary>
        /// 导数据到开源pp_mus
        /// </summary>
        /// <returns></returns>
        public static void ImData(string year, string month, string sTime, string fdate)
        {
            StringBuilder sqlBuilder = new StringBuilder(1000);
            sqlBuilder.Append(" SELECT  0 id,isb.fbillno bll,t002.FNumber as  mtr,t002.FName+'('+t002.Fmodel+')' as mtr_txt,t002.FNumber as  mtr_sht, isbe.FBatchNo as bnm,");
            sqlBuilder.Append(" '' pss,'' flg,'import' exa,t033.FName as banTxt, t034.FName as  opt_txt,isb.Fdate  dat,isb.Fcheckdate  dt1,1000*isbe.FQty qty, isbe.Famount prs,isbe.Fauxprice ppr,");
            sqlBuilder.Append(" isb.Fdate pprDat,t001.FNumber as  dpt, t001.FName as  dpt_txt,'' wkpsht,'' wkp_txt,0 wkpseq,t032.FNumber as mcnsht,'' potsht,isb.FInterID mov,");
            sqlBuilder.Append(" isb.FHeadSelfB0438 codsht,'' prjsht,isb.FHeadSelfB0433 prj_clr,'Rl' coTypSht,'02' typ,0 spt,0 cls,0 zflg");
            sqlBuilder.Append(" FROM ICStockBillEntry  isbe join t_ICItem  t002 on t002.FItemID = isbe.FItemID AND t002.FItemID <>0");
            sqlBuilder.Append(" inner join  ICStockBill isb  on isbe.finterid=isb.finterid");
            sqlBuilder.Append(" left join t_Department  t001 on isb.FDeptID= t001.FItemID  AND t001.FItemID<>0");
            sqlBuilder.Append(" left join t_Item  t032 on isb.FHeadSelfB0435= t032.FItemID  AND t032.FItemID<>0");
            sqlBuilder.Append(" left join t_Item  t033 on isb.FHeadSelfB0436= t033.FItemID  AND t033.FItemID<>0");
            sqlBuilder.Append(" left join t_Emp  t034 on isb.FHeadSelfB0437= t034.FItemID  AND t034.FItemID<>0");
            sqlBuilder.Append(" where isb.fstatus=1 and isb.FHeadSelfB0438<>'' and isb.FHeadSelfB0438 is not null");

            //sqlBuilder.Append(" SELECT  0 id,isb.fbillno bll,t002.FNumber as  mtr,t002.FName+'('+t002.Fmodel+')' as mtr_txt,t002.FNumber as  mtr_sht, isbe.FBatchNo as bnm,");
            //sqlBuilder.Append(" '' pss,'' flg,'import' exa,'' as banTxt, '' as  opt_txt,isb.Fdate  dat,isb.Fdate  dt1,1000*isbe.FQty qty, isbe.Famount prs,isbe.Fauxprice ppr,");
            //sqlBuilder.Append(" isb.Fdate pprDat,t001.FNumber as  dpt, t001.FName as  dpt_txt,'' wkpsht,'' wkp_txt,0 wkpseq,'' as mcnsht,'' potsht,isb.FInterID mov,");
            //sqlBuilder.Append(" isb.FHeadSelfB0433 codsht,'' prjsht,isb.FHeadSelfB0432 prj_clr,'Rl' coTypSht,'02' typ,0 spt,0 cls,0 zflg");
            //sqlBuilder.Append(" FROM ICStockBillEntry  isbe join t_ICItem  t002 on t002.FItemID = isbe.FItemID AND t002.FItemID <>0");
            //sqlBuilder.Append(" inner join  ICStockBill isb  on isbe.finterid=isb.finterid");
            //sqlBuilder.Append(" left join t_Department  t001 on isb.FDeptID= t001.FItemID  AND t001.FItemID<>0");
            //sqlBuilder.Append(" where isb.FHeadSelfB0433<>'' and isb.FHeadSelfB0433 is not null");

            if (!string.IsNullOrEmpty(year))
            {
                sqlBuilder.AppendFormat(" and year(isb.fdate)={0}", year);
            }
            if (!string.IsNullOrEmpty(month))
            {
                sqlBuilder.AppendFormat(" and month(isb.fdate)={0}", month);
            }
            if (!string.IsNullOrEmpty(sTime))
            {
                sqlBuilder.AppendFormat(" and  isb.fdate>='{0}'", sTime);
            }
            if (!string.IsNullOrEmpty(fdate))
            {
                sqlBuilder.AppendFormat(" and  isb.fcheckdate='{0}'", fdate);
            }
            sqlBuilder.Append(" ORDER BY isb.fbillno,isbe.FEntryID");

            List<string> retList = new List<string>();
            DataTable dt = new DataTable();
            try
            {
                CBase.AddErroLog(sqlBuilder.ToString());
                dt = DbHelper.GetDataTable(sqlBuilder.ToString());


                //判断金蝶的生产单在开源是否存在
                string kyPrjSht = "";
                string prjsht = "";
                string clr = "";
                string ofbillno = "";
                string fbillno = "";//单据编号
                string fdate1 = ""; //单据日期
                string dpt = ""; //车间
                string mcn = "";//机台


                bool isSame = false;

                int rows = dt.Rows.Count;
                for (int i = 0; i < rows; i++)
                {
                    prjsht = dt.Rows[i]["codsht"].ToString();
                    clr = dt.Rows[i]["prj_clr"].ToString();
                    fbillno = dt.Rows[i]["bll"].ToString();
                    fdate1 = dt.Rows[i]["dt1"].ToString();
                    dpt = dt.Rows[i]["dpt_txt"].ToString();
                    mcn = dt.Rows[i]["mcnsht"].ToString();
                    if (fbillno != ofbillno)
                    {
                        ofbillno = fbillno;
                        isSame = true;
                    }
                    else
                    {
                        isSame = false;
                    }


                    kyPrjSht = IsExistPrj(prjsht, clr);
                    if (string.IsNullOrEmpty(kyPrjSht))
                    {
                        string retSt = "  导入异常,生产单:" + prjsht + " 色号:" + clr + " 在在线采集系统中不存在";
                        CBase.AddErroLog(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + retSt);
                        if (isSame)
                        {
                            LqImportDac.WriteXml(fdate1, fbillno, prjsht, clr, dpt, mcn, "导入失败 " + retSt, 0, 1);
                        }
                    }
                    else
                    {
                        dt.Rows[i]["prjsht"] = kyPrjSht;
                        dt.Rows[i].AcceptChanges();
                        dt.Rows[i].SetAdded();
                        //正常的单号用数组记录下来，供导入趁工后写日志
                        if (isSame)
                        {
                            retList.Add(fbillno + "," + fdate1 + "," + prjsht + "," + clr + "," + dpt + "," + mcn);
                        }

                    }

                }
            }
            catch (System.Exception ex)
            {
                CBase.AddErroLog(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  导入异常:" + ex.ToString());
            }
            //


            DBConn.GetConnStr1(Const.DBConnFile);
            string connstr = Const.SqlConnentionString1;
            DbHelper.Open(connstr);
            try
            {
                sqlBuilder.Remove(0, sqlBuilder.Length);
                sqlBuilder.AppendFormat("delete from pp_mus where exa='import' and year(dat)={0} and month(dat)={1}", year, month);
                if (!string.IsNullOrEmpty(fdate))
                {
                    sqlBuilder.AppendFormat(" and  dt1='{0}'", fdate);
                }

                DbHelper.BeginTran(connstr);
                DbHelper.ExecuteNonQuery(connstr, sqlBuilder.ToString());

                DbHelper.Update(connstr, dt, "select * from pp_mus where 1<>1");
                DbHelper.CommitTran(connstr);

                //将导成功的单写入到XML
                if (retList != null && retList.Count > 0)
                {
                    for (int i = 0; i < retList.Count; i++)
                    {

                        LqImportDac.WriteXml(retList[i].Split(',')[1], retList[i].Split(',')[0], retList[i].Split(',')[2], retList[i].Split(',')[3], retList[i].Split(',')[4], retList[i].Split(',')[5], "导入成功 ", 0, 0);
                    }

                }
            }
            catch (System.Exception ex)
            {
                DbHelper.RollbackTran(connstr);
                CBase.AddErroLog(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  导入异常:" + ex.ToString());
            }
            finally
            {
                DbHelper.Close(connstr);
            }

        }
        public static int ReadXml(string ldat, string logTyp, ref DataTable dt)
        {
            string sFileName = Const.StartupPath + "\\log\\import\\" + ldat + ".xml";
            DataSet logDs = new DataSet();
            if (System.IO.File.Exists(sFileName))
            {
                string sFilter = "flg=" + logTyp;
                logDs.ReadXml(sFileName);


                dt = logDs.Tables[0].Clone();

                foreach (DataRow dr in logDs.Tables[0].Select(sFilter))
                {

                    dt.ImportRow(dr);
                }
                return 1;
            }
            else
            {
                return 0;
            }

        }

        public static void WriteXml(string bdat, string fbillno, string prjsht, string clr, string dpt, string mcn, string info, int typ, int flg)
        {

            try
            {
                string sFileName = Const.StartupPath + "\\log\\import\\" + DateTime.Now.ToString("yyyy-MM-dd") + ".xml";
                if (!Directory.Exists(Const.StartupPath + "\\log\\import"))
                {
                    Directory.CreateDirectory(Const.StartupPath + "\\log\\import");
                }


                XmlDocument xmldoc;
                xmldoc = new XmlDocument();
                XmlNode xmlnode;
                XmlElement xmlelem;
                if (System.IO.File.Exists(sFileName))
                {
                    xmldoc.Load(sFileName);
                }
                else
                {

                    //加入XML的声明段落
                    xmlnode = xmldoc.CreateNode(XmlNodeType.XmlDeclaration, "", "");
                    xmldoc.AppendChild(xmlnode);
                    //加入一个根元素
                    xmlelem = xmldoc.CreateElement("", "Message", "");
                    xmldoc.AppendChild(xmlelem);
                }




                XmlNode root = xmldoc.SelectSingleNode("Message");//查找<Message> 
                XmlElement xe1 = xmldoc.CreateElement("Node");//创建一个<Node>节点 

                XmlElement xesub1 = xmldoc.CreateElement("fbillno"); //领料单号 
                xesub1.InnerText = fbillno;//设置文本节点 
                xe1.AppendChild(xesub1);//添加到<Node>节点中 

                XmlElement xesub2 = xmldoc.CreateElement("bdat"); //领料单日期
                xesub2.InnerText = Convert.ToDateTime(bdat).ToString("yyyy-MM-dd");
                xe1.AppendChild(xesub2);

                XmlElement xesub3 = xmldoc.CreateElement("typ");
                xesub3.InnerText = typ.ToString();   //类型：0:自动；1：手动
                xe1.AppendChild(xesub3);

                XmlElement xesub13 = xmldoc.CreateElement("prjsht");
                xesub13.InnerText = prjsht;   //生产单
                xe1.AppendChild(xesub13);

                XmlElement xesub14 = xmldoc.CreateElement("clr");
                xesub14.InnerText = clr;   //颜色
                xe1.AppendChild(xesub14);

                XmlElement xesub15 = xmldoc.CreateElement("dpt");
                xesub15.InnerText = dpt;   //部门
                xe1.AppendChild(xesub15);

                XmlElement xesub16 = xmldoc.CreateElement("mcn");
                xesub16.InnerText = mcn;   //机台
                xe1.AppendChild(xesub16);


                XmlElement xesub4 = xmldoc.CreateElement("info"); //消息
                xesub4.InnerText = info;
                xe1.AppendChild(xesub4);

                XmlElement xesub5 = xmldoc.CreateElement("flg");
                xesub5.InnerText = flg.ToString();   //类型：0:正常；1：异常
                xe1.AppendChild(xesub5);

                root.AppendChild(xe1);//添加到<Message>节点中 
                //保存创建好的XML文档
                xmldoc.Save(sFileName);
            }
            catch (System.Exception ex)
            {
                CBase.AddErroLog(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  写文件异常:" + ex.ToString());
            }

        }

        /// <summary>
        /// 机台能源日结
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        public static DataTable GetMcnDataGatherAllByDay(string date)
        {
            string sql = @"SELECT sht,txt,ban,CONVERT(DECIMAL(18,0),oprSumQty) AS oprSumQty,CONVERT(DECIMAL(18,1),watQty) AS watQty,CONVERT(DECIMAL(18,1),elecQty) AS elecQty,CONVERT(DECIMAL(18,1),vapQty) AS vapQty,CONVERT(DECIMAL(18,1),jnQty) AS jnQty,CONVERT(DATE,ISNULL(dc_McnDataGatherAll.cdat,GETDATE()))  AS cdat FROM dc_McnDataGatherAll JOIN PP_MCN ON sht=mcnsht  WHERE cdat=@date ORDER BY sht,ban ";

            IDataParameter[] param = new SqlParameter[1];
            param[0] = new SqlParameter("@date", date);
            //param[1] = new SqlParameter("@edt", strEdt);
            DataTable dt = new DataTable();
            DbHelper.Fill(dt, sql, param);
            return dt;
        }

        /// <summary>
        /// 机台能源月结
        /// </summary>
        /// <param name="year"></param>
        /// <param name="mon"></param>
        /// <returns></returns>
        public static DataTable GetMcnDataGatherAllByMon(string year, string mon)
        {                      
            string sql = @"
 SELECT  sht,txt,ban,CONVERT(DECIMAL(18,0),SUM(oprSumQty)) AS oprSumQty,CONVERT(DECIMAL(18,1),SUM(watQty)) AS watQty,CONVERT(DECIMAL(18,1),SUM(elecQty)) AS elecQty,CONVERT(DECIMAL(18,1),SUM(vapQty)) AS vapQty,
 CONVERT(DECIMAL(18,1),SUM(jnQty)) AS jnQty
 FROM dc_McnDataGatherAll JOIN PP_MCN ON sht=mcnsht  WHERE yer=@year AND mon=@mon GROUP BY sht,txt,ban ORDER BY sht,ban ";

            IDataParameter[] param = new SqlParameter[2];
            param[0] = new SqlParameter("@year", year);
            param[1] = new SqlParameter("@mon", mon);
            //param[2] = new SqlParameter("@rowNum", rowNum);
            DataTable dt = new DataTable();
            DbHelper.Fill(dt, sql, param);
            return dt;
        }

        public static DataTable GetMcnDataGatherAllByDayAndRownum(string date, int rowNum)
        {
            string sql = @"SELECT * FROM (SELECT dc_McnDataGatherAll.id, sht,txt,ban,oprSumQty,watQty,elecQty,vapQty,jnQty,
cdat, ROW_NUMBER() OVER (ORDER BY dbo.dc_McnDataGatherAll.id) AS rowNum 
FROM dc_McnDataGatherAll JOIN PP_MCN ON sht=mcnsht AND cdat=@date) TEMP
WHERE rowNum BETWEEN @rowNum AND @rowNumAdd6  ";

            IDataParameter[] param = new SqlParameter[3];
            param[0] = new SqlParameter("@date", date);
            param[1] = new SqlParameter("@rowNum", rowNum);
            param[2] = new SqlParameter("@rowNumAdd6", (rowNum + 5));
            //param[1] = new SqlParameter("@edt", strEdt);
            DataTable dt = new DataTable();
            DbHelper.Fill(dt, sql, param);
            return dt;
        }

        public static DataTable GetMcnDataGatherAllByMonAndRownum(string year, string mon)
        {
            string sql = @"SELECT TOP 6 MCN.sht,MCN.txt,MCN.BAN,ISNULL(SUM(oprSumQty),0) oprSumQty,ISNULL(SUM(watQty),0) watQty,ISNULL(SUM(elecQty),0) elecQty,
ISNULL(SUM(jnQty),0) vapQty,ISNULL(SUM(jnQty),0) jnQty,CONVERT(DATE,ISNULL(dc_McnDataGatherAll.cdat,GETDATE()))  AS cdat FROM
(
SELECT PP_MCN.TXT,PP_MCN.SHT,PP_ENV.TXT AS BAN FROM PP_MCN JOIN PP_ENV ON PP_ENV.typ=20 AND PP_ENV.del=0 
) MCN LEFT JOIN dc_McnDataGatherAll ON MCN.BAN = dc_McnDataGatherAll.ban AND MCN.SHT=dc_McnDataGatherAll.mcnsht
AND dc_McnDataGatherAll.yer=@year AND mon=@mon GROUP BY MCN.sht,MCN.txt,MCN.BAN,MCN.BAN,dc_McnDataGatherAll.CDAT
ORDER BY MCN.sht,BAN";

            IDataParameter[] param = new SqlParameter[2];
            param[0] = new SqlParameter("@year", year);
            param[1] = new SqlParameter("@mon", mon);
            DataTable dt = new DataTable();
            DbHelper.Fill(dt, sql, param);
            return dt;
        }

        /// <summary>
        /// 返回屏幕信息
        /// </summary>
        /// <param name="sht"></param>
        /// <returns></returns>
        public static DataTable GetLedInfo(string sht)
        {
            DataTable dt = new DataTable();
            string sql = @"select * from ledInfo where sht='" +  sht + "' ";
            DbHelper.Fill(dt, sql);
            return dt;
        }

        /// <summary>
        /// 获取功能点
        /// </summary>
        /// <returns></returns>
        public static DataTable GetLedInfoByConn()
        {            
            DBConn.DbHelperOpen();
            DataTable dt = new DataTable();
            string sql = @"select * from ledInfo where 1=1 ";
            dt = DbHelper.GetDataTable(sql);            
            return dt;
        }

        /// <summary>
        /// 化验室
        /// </summary>
        /// <returns></returns>
        public static DataTable GetHys_View()
        {
            DataTable dt = new DataTable();
            string sql = @"select * from hysled_view order by 打样完成时间,状态";
            DbHelper.Fill(dt, sql);
            return dt;
        }

        public static DataTable GetTjs_View()
        {
            DataTable dt = new DataTable();
            string sql = @"select * from tjsLED_view ";
            DbHelper.Fill(dt, sql);
            return dt;
        }

        public static DataTable GetZws_View()
        {
            DataTable dt = new DataTable();
            string sql = @"select * from zwsLED_view ";
            DbHelper.Fill(dt, sql);
            return dt;
        }

        public static DataTable GetMgs_View()
        {
            DataTable dt = new DataTable();
            string sql = @"select * from mgsLED_view";
            DbHelper.Fill(dt, sql);
            return dt;
        }

        public static DataTable GetZzs_View()
        {
            DataTable dt = new DataTable();
            string sql = @"select * from zzsLED_view";
            DbHelper.Fill(dt, sql);
            return dt;
        }

        public static DataTable GetZjs_View()
        {
            DataTable dt = new DataTable();
            string sql = @"select * from zjsLED_view";
            DbHelper.Fill(dt, sql);
            return dt;
        }

        public static DataTable GetZjs1_View()
        {
            DataTable dt = new DataTable();
            string sql = @"select * from zjs1LED_view";
            DbHelper.Fill(dt, sql);
            return dt;
        }
    }
}
