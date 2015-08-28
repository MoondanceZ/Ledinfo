using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using  FI.Public;
using KY.AP.DataServices;
using KY.Data.DB;

namespace FI.DataAccess
{
    public class CPageDac
    {

        /// <summary>
        /// 获取项目
        /// </summary>
        /// <returns></returns>
        public static DataTable GetProject()
        {
            DataTable dt = new DataTable();
            StringBuilder sqlBuilder = new StringBuilder(100);
            sqlBuilder.Append("select pno,pname from fproject order by pno");
            dt = DBConn.GetSqlData(sqlBuilder.ToString());
            return dt;
        }
        /// <summary>
        /// 获取当前项目
        /// </summary>
        /// <returns></returns>
        public static DataTable GetCurProj()
        {
            DataTable dt = new DataTable();
            StringBuilder sqlBuilder = new StringBuilder(100);
            sqlBuilder.Append("select fkeyName,fvalue from  sys_systemprofile  where fkey='CompanyName'");
            dt = DBConn.GetSqlData(sqlBuilder.ToString());
            return dt;
        }
        /// <summary>
        /// 更新当前项目
        /// </summary>
        /// <returns></returns>
        public static bool UpdateCurProj(string curProj)
        {
            DataTable dt = new DataTable();
            DbHelper.Open();
            try
            {
                DbHelper.BeginTran();
                StringBuilder sqlBuilder = new StringBuilder(100);
                sqlBuilder.AppendFormat("update  sys_systemprofile  set  fkeyName='{0}'  where fkey='CompanyName'", curProj);
                DbHelper.ExecuteNonQuery(sqlBuilder.ToString());
                DbHelper.CommitTran();
                return true;
            }
            catch
            {
                DbHelper.RollbackTran();
                return false;
            }
            finally
            {
                DbHelper.Close();
            }

        }
        /// <summary>
        /// 获取当前测试项目
        /// </summary>
        /// <returns></returns>
        public static DataTable GetCurTestProj()
        {
            DataTable dt = new DataTable();
            StringBuilder sqlBuilder = new StringBuilder(100);
            sqlBuilder.Append("select fkeyName,fvalue from  sys_systemprofile  where fkey='TestProject'");
            dt = DBConn.GetSqlData(sqlBuilder.ToString());
            return dt;
        }
        /// <summary>
        /// 更新当前测试项目
        /// </summary>
        /// <returns></returns>
        public static bool UpdateCurTestProj(string curProj)
        {
            DataTable dt = new DataTable();
            DbHelper.Open();
            try
            {
                DbHelper.BeginTran();
                StringBuilder sqlBuilder = new StringBuilder(100);
                sqlBuilder.AppendFormat("update  sys_systemprofile  set  fvalue='{0}'  where fkey='TestProject'", curProj);
                DbHelper.ExecuteNonQuery(sqlBuilder.ToString());

                DbHelper.CommitTran();
                return true;
            }
            catch
            {
                DbHelper.RollbackTran();
                return false;
            }
            finally
            {
                DbHelper.Close();
            }

        }
        /// <summary>
        /// 获取功能点使用项目
        /// </summary>
        /// <returns></returns>
        public static DataTable GetFunProject(string elemNo)
        {
            DataTable dt = new DataTable();
            StringBuilder sqlBuilder = new StringBuilder(100);
            sqlBuilder.AppendFormat("select fp.pno,fp.pname from fpagehis inner join  fproject fp on fpagehis.pno=fp.pno where id={0} order by pno",elemNo);
            dt = DBConn.GetSqlData(sqlBuilder.ToString());
            return dt;
        }
        /// <summary>
        /// 获取子系统
        /// </summary>
        /// <returns></returns>
        public static DataTable GetSub()
        {
            DataTable dt = new DataTable();
            StringBuilder sqlBuilder = new StringBuilder(100);
            sqlBuilder.Append("select elemno,title from fpage where supelem='00' order by serial ");
            dt = DBConn.GetSqlData(sqlBuilder.ToString());
            return dt;
        }
        /// <summary>
        /// 获取模块
        /// </summary>
        /// <returns></returns>
        public static DataTable GetMdl(string subNo)
        {
            DataTable dt = new DataTable();
            StringBuilder sqlBuilder = new StringBuilder(200);
            sqlBuilder.AppendFormat("select elemno,title from fpage where supelem='{0}'  order by serial ",subNo);
            dt = DBConn.GetSqlData(sqlBuilder.ToString());
            return dt;
        }
        /// <summary>
        /// 获取系统功能点
        /// </summary>
        /// <returns></returns>
        public static DataTable GetFun(string pNo,string subNo,string mdlNo,string md1no,string funno)
        {
            DataTable dt = new DataTable();
            StringBuilder sqlBuilder = new StringBuilder(1000);
            //sqlBuilder.Append(" select page3.id,page.title subTxt,page2.title mdlTxt,page3.title elemTxt,page.elemno subno,");
            //sqlBuilder.Append(" page2.elemno mdlNo,page3.elemno,page.verno,page3.fgp,page3.path,page3 .pagename,page3.serial,page3.show,");
            //sqlBuilder.Append(" (case when fpagehis.id is null then  0 else 1 end) isSelected");
            //sqlBuilder.Append(" from fpage  page join fpage as page2 on page.id = page2.fgp join fpage as page3 on page2.id = page3.fgp"); 
            //sqlBuilder.Append(" left join fpageHis on page3.id=fpagehis.id");
            //if(!string .IsNullOrEmpty (pNo))
            //{
            //    sqlBuilder.AppendFormat(" and fpageHis.pno='{0}'",pNo);
            //}
            //sqlBuilder.Append(" where page.fgp=1 and page2.rgtTyp=1 and page3.rgtTyp=1 and page.rgttyp=1");

            //if (!string.IsNullOrEmpty(subNo))
            //{
            //    sqlBuilder.AppendFormat(" and page.elemno='{0}'", subNo);
            //}
            //if (!string.IsNullOrEmpty(mdlNo))
            //{
            //    sqlBuilder.AppendFormat(" and page2.elemno='{0}'", mdlNo);
            //}
            
            sqlBuilder.Append(" select page3.id,page.title subTxt,page2.title mdlTxt,'' md1,'' fun,page3.title elemTxt,page.elemno subno, "); 
            sqlBuilder.Append(" page2.elemno mdlNo,page3.elemno,page3.verno,page3.fgp,page3.path,page3 .pagename,page3.serial,page3.show"); 
            sqlBuilder.Append(" ,(case when fpagehis.id is null then  0 else 1 end) isSelected,page.serial,page2.serial,page3.serial"); 
            sqlBuilder.Append(" from fpage  page join fpage as page2 on page.id = page2.fgp join fpage as page3 on page2.id = page3.fgp"); 
            sqlBuilder.Append(" left join fpageHis on page3.id=fpagehis.id ");
            if (!string.IsNullOrEmpty(pNo))
            {
                sqlBuilder.AppendFormat(" and fpageHis.pno='{0}'", pNo);
            }
            sqlBuilder.Append(" where isnull(page3.isfuntion,0)=0  and  page.fgp=1 and page2.rgtTyp=1 and page3.rgtTyp=1 and page.rgttyp=1");
            if (!string.IsNullOrEmpty(subNo))
            {
                sqlBuilder.AppendFormat(" and page.elemno='{0}'", subNo);
            }
            if (!string.IsNullOrEmpty(mdlNo))
            {
                sqlBuilder.AppendFormat(" and page2.elemno='{0}'", mdlNo);
            }
            sqlBuilder.Append(" union");
            sqlBuilder.Append(" ("); 
            sqlBuilder.Append(" select page3.id,page0.title subTxt,page1.title mdlTxt,page.title md1,page2.title fun,page3.title elemTxt,page.elemno subno,");  
            sqlBuilder.Append(" page2.elemno mdlNo,page3.elemno,page3.verno,page3.fgp,page3.path,page3 .pagename,page3.serial,page3.show,"); 
            sqlBuilder.Append("   (case when fpagehis.id is null then  0 else 1 end) isSelected,page.serial,page2.serial,page3.serial"); 
            sqlBuilder.Append(" from fpage  page join fpage as page2 on page.id = page2.fgp join fpage as page3 on page2.id = page3.fgp"); 
            sqlBuilder.Append(" join fpage as page1 on page1.id = page.fgp join fpage as page0 on page0.id = page1.fgp"); 
            sqlBuilder.Append(" left join fpageHis on page3.id=fpagehis.id ");
            if (!string.IsNullOrEmpty(pNo))
            {
                sqlBuilder.AppendFormat(" and fpageHis.pno='{0}'", pNo);
            }
            sqlBuilder.Append(" where isnull(page.isfuntion,0)=1");   
            sqlBuilder.Append(" and  page0.fgp=1 and page2.rgtTyp=1 and page3.rgtTyp=1 and page.rgttyp=1");
            if (!string.IsNullOrEmpty(subNo))
            {
                sqlBuilder.AppendFormat(" and page0.elemno='{0}'", subNo);
            }
            if (!string.IsNullOrEmpty(mdlNo))
            {
                sqlBuilder.AppendFormat(" and page1.elemno='{0}'", mdlNo);
            }
            if (!string.IsNullOrEmpty(md1no))
            {
                sqlBuilder.AppendFormat(" and page.elemno='{0}'", md1no);
            }
            if (!string.IsNullOrEmpty(funno))
            {
                sqlBuilder.AppendFormat(" and page2.elemno='{0}'", funno);
            }
            sqlBuilder.Append(" )"); 
            sqlBuilder.Append(" order by page.serial,page2.serial,page3.serial ");
            dt = DBConn.GetSqlData(sqlBuilder.ToString());
            return dt;
        }

        /// <summary>
        /// 插入PAGE到fpagehis
        /// </summary>
        /// <returns></returns>
        public static bool InsertPage(string apid,string pid,string pno,string pdat)
        {
            DbHelper.Open();
            try
            {
                DbHelper.BeginTran();
                StringBuilder sqlBuilder = new StringBuilder(8000);
                if (apid != "")
                {
                    // sqlBuilder.AppendFormat("delete from  fpageHis where pno='{0}' and pdat<'{1}';delete from fpageHis where pno='{0}' and pdat='{1}' and id in ({2});",pno,pdat,apid);
                    sqlBuilder.AppendFormat("delete from fpageHis where pno='{0}'  and id in ({1});", pno, apid);
                    DbHelper.ExecuteNonQuery(sqlBuilder.ToString());
                }
                if (pid != "")
                {
                    sqlBuilder.Remove(0, sqlBuilder.Length);
                    sqlBuilder.Append("insert into fpageHis");
                    sqlBuilder.Append(" select id,elemno,supelem,title,serial,fgp,path,pagename,help,param,show,child,");
                    sqlBuilder.AppendFormat(" pot,type,ispage,base,englishtitle,mdlno,rgttyp,ord,verno,isfuntion,funtionline ,'{0}' pno,1 sort,'{1}' pdat", pno, pdat);
                    sqlBuilder.AppendFormat(" from fpage where id in({0})", pid);
                    DbHelper.ExecuteNonQuery(sqlBuilder.ToString());
                }
                DbHelper.CommitTran();
                return true;
            }
            catch
            {
                DbHelper.RollbackTran();
                return false;
            }
            finally
            {
                DbHelper.Close();
            }
        }


        /// <summary>
        /// 导出PAGE到目标数据
        /// </summary>
        /// <returns></returns>
        public static bool ExPage(string pno, string db,string pageName)
        {
            DbHelper.Open();
            try
            {
                DbHelper.BeginTran();
                StringBuilder sqlBuilder = new StringBuilder(1000);
                try
                {
                    sqlBuilder.AppendFormat("drop table {0}.dbo.{1} ", db, pageName);
                    DbHelper.ExecuteNonQuery(sqlBuilder.ToString());
                }
                catch
                { 
                
                }
                sqlBuilder.Remove(0, sqlBuilder.Length);
               // sqlBuilder.AppendFormat("insert into {0}.db.page", db);
               // sqlBuilder.Append(" select id,elemno,supelem,title,serial,fgp,path,pagename,help,param,show,child,");
               // sqlBuilder.AppendFormat(" pot,type,ispage,base,englishtitle,mdlno,rgttyp,ord,verno into {0}.dbo.{1}", db,pageName);
               // sqlBuilder.AppendFormat(" from fpageHis where pno='{0}'", pno);

                sqlBuilder.AppendFormat("select * into {0}.dbo.{1}  from (", db, pageName);
                sqlBuilder.Append(" select id,elemno,supelem,title,serial,fgp,path,pagename,help,param,show,child,pot,type,ispage,base,englishtitle,mdlno,rgttyp,ord,verno,isfuntion,funtionline  from fpage where fpage.id =1");
                sqlBuilder.Append(" union");
                sqlBuilder.Append(" select id,elemno,supelem,title,serial,fgp,path,pagename,help,param,show,child,pot,type,ispage,base,englishtitle,mdlno,rgttyp,ord,verno,isfuntion,funtionline  ");
                sqlBuilder.AppendFormat(" from fpageHis where pno='{0}'",pno);
                sqlBuilder.Append(" union");
                sqlBuilder.Append(" select id,elemno,supelem,title,serial,fgp,path,pagename,help,param,show,child,pot,type,ispage,base,englishtitle,mdlno,rgttyp,ord,verno,isfuntion,funtionline  ");
                sqlBuilder.AppendFormat(" from fpage page where page.id in(select  distinct fgp from fpagehis where pno='{0}' )",pno);
                sqlBuilder.Append(" union");
                sqlBuilder.Append(" select id,elemno,supelem,title,serial,fgp,path,pagename,help,param,show,child,pot,type,ispage,base,englishtitle,mdlno,rgttyp,ord,verno,isfuntion,funtionline  ");
                sqlBuilder.AppendFormat(" from fpage page  where page.id in (select fgp from fpage  where  id in(select  distinct fgp from fpagehis where pno='{0}' ))",pno);

                sqlBuilder.Append(" union");
                sqlBuilder.Append(" select id,elemno,supelem,title,serial,fgp,path,pagename,help,param,show,child,pot,type,ispage,base,englishtitle,mdlno,rgttyp,ord,verno,isfuntion,funtionline   from  fpage page ");
                sqlBuilder.AppendFormat(" where page.id in (select fgp from fpage  where  id in (select fgp from fpage   where  id in(select distinct fgp from fpagehis where pno='{0}' )))",pno);
                sqlBuilder.Append(" union ");
                sqlBuilder.Append(" select id,elemno,supelem,title,serial,fgp,path,pagename,help,param,show,child,pot,type,ispage,base,englishtitle,mdlno,rgttyp,ord,verno,isfuntion,funtionline   from  fpage page ");
                sqlBuilder.AppendFormat(" where page.id in ( select fgp from fpage where  id in(select fgp from fpage   where  id in (select fgp from fpage   where  id in(select  distinct fgp from fpagehis where pno='{0}' ))))",pno);

                sqlBuilder.Append(" ) t order by t.id ");

                DbHelper.ExecuteNonQuery(sqlBuilder.ToString());
                DbHelper.CommitTran();
                return true;
            }
            catch
            {
                DbHelper.RollbackTran();
                return false;
            }
            finally
            {
                DbHelper.Close();
            }
        }
    }
}
