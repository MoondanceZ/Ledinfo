using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using  FI.Public;
using KY.AP.DataServices;
using KY.Data.DB;

namespace FI.DataAccess
{
    public class FIDac
    {
        /// <summary>
        /// 获取采购入库单
        /// </summary>
        /// <returns></returns>
        public static DataTable GetInBill(string year,string month,string pvd)
        {
            DataTable dt = new DataTable();
            StringBuilder sqlBuilder = new StringBuilder(1000);
            sqlBuilder.Append("select pvd.sht pvdsht,pvd.txt pvdtxt,bll.sht bsht,bll.id bll,bll.dat bdat,sum(act.prs) prs,sum(act.tax) tax,(sum(act.prs)+sum(act.tax)) tprs ");
            sqlBuilder.Append(" from bll inner join pvd on bll.mkr=pvd.id");
            sqlBuilder.Append(" inner join act on bll.id=act.bll");
            sqlBuilder.AppendFormat(" where (cflg is null or cflg=0) and bll.btp=4 AND bll.ext like '收料' and bll.rjt=0  and year(bll.dat)={0} and month(bll.dat)={1}", year, month);
            if (!string.IsNullOrEmpty(pvd))
            {
                sqlBuilder.AppendFormat("  and (pvd.sht like '%{0}%' or pvd.txt like '%{0}%') ",pvd);
            
            }
            sqlBuilder.Append(" group by pvd.sht,pvd.txt,pvd.id,bll.sht,bll.id,bll.dat");
            sqlBuilder.Append(" order by bll.dat desc");
            dt=DBConn.GetSqlData(sqlBuilder.ToString());
            return dt;
        }
        /// <summary>
        /// 获取采购入库单
        /// </summary>
        /// <returns></returns>
        public static DataTable GetInBill(string time)
        {
            DataTable dt = new DataTable();
            StringBuilder sqlBuilder = new StringBuilder(1000);
            sqlBuilder.Append("select pvd.sht pvdsht,pvd.txt pvdtxt,bll.sht bsht,bll.id bll,bll.dat bdat,sum(act.prs) prs,sum(act.tax) tax,(sum(act.prs)+sum(act.tax)) tprs ");
            sqlBuilder.Append(" from bll inner join pvd on bll.mkr=pvd.id");
            sqlBuilder.Append(" inner join act on bll.id=act.bll");
            sqlBuilder.AppendFormat(" where (cflg is null or cflg=0) and bll.btp=4 AND bll.ext like '收料' and bll.rjt=0");
            if (!string.IsNullOrEmpty(time))
            {
                sqlBuilder.AppendFormat(" and bll.dat>='{0}'", time);

            }
            sqlBuilder.Append(" group by pvd.sht,pvd.txt,pvd.id,bll.sht,bll.id,bll.dat");
            sqlBuilder.Append(" order by bll.dat desc");
            dt = DBConn.GetSqlData(sqlBuilder.ToString());
            return dt;
        }
        /// <summary>
        /// 根据入库生产单号,取入库物料明细前两个物料
        /// </summary>
        /// <returns></returns>
        public static string  GetAct(string bll)
        {
            DataTable dt = new DataTable();
            string sRet = "";
            StringBuilder sqlBuilder = new StringBuilder(1000);
            sqlBuilder.Append("select top 2 mtr.txt from act inner join mtr on act.mtr=mtr.id");
            sqlBuilder.AppendFormat(" where act.bll={0}",bll);
            sqlBuilder.Append(" order by act.prs desc");
            dt = DBConn.GetSqlData(sqlBuilder.ToString());

            if (dt != null && dt.Rows.Count > 0)
            {
                sRet += dt.Rows[0]["txt"].ToString();
                if (dt.Rows.Count > 1)
                {
                    sRet +=","+dt.Rows[1]["txt"].ToString();
                }
            }
            return sRet;
        }
        /// <summary>
        /// 更新已经生成过凭证的单据标志为1
        /// </summary>
        /// <returns></returns>
        public static void updateBll(string bll)
        {
            DbHelper.Open();
            try
            {
                DbHelper.BeginTran();
                StringBuilder sqlBuilder = new StringBuilder(1000);
                sqlBuilder.AppendFormat("update bll set cflg=1 where id={0}", bll);
                DbHelper.ExecuteNonQuery(sqlBuilder.ToString());
                DbHelper.CommitTran();
            }
            catch
            {
                DbHelper.RollbackTran();
            }
            finally
            {
                DbHelper.Close();
            }
        }

        /// <summary>
        /// 生成凭证
        /// </summary>
        /// <returns></returns>
        public static bool  CreatePZ(string pvdsht,string pvdtxt,string mtrTxt,string bsht,string  bll,string bdat,decimal prs,decimal tax,decimal tprs)
        {
            int fyear = Convert.ToDateTime(bdat).Year, fperiod =Convert.ToDateTime(bdat).Month;
            int yclAcc_id = 0, yjsjAcc_id = 0, yfzkAcc_id = 0;
            string yclAcc_number = "123", yjsjAcc_number = "221.001.01.01", yfzkAcc_number = "203";
            int fgroupid=0, fnumber=0, fvoucherid=0, fSerialNum=0, fEntryCount=0;
            int fdetailid = 0,fodetailid = 0,fentryid=0;
            int pvdid=0;
            string fexp = pvdtxt + ";" + GetAct(bll) + ";" + bsht; //会计摘要

            DBConn.GetConnStr1(Const.DBConnFile);
            string connstr = Const.SqlConnentionString1;

            //科目赋值
            yclAcc_number  = Const.Ycl;
            yjsjAcc_number = Const.Yjsj;
            yfzkAcc_number = Const.Yfzk;




            //取会计科目对应K3的科目ID,供应商ID
            StringBuilder sqlBuilder = new StringBuilder(1000);
            sqlBuilder.Append("select max(yclAcc_id) yclAcc_id ,max(yjsjAcc_id) yjsjAcc_id,max(yfzkAcc_id) yfzkAcc_id,max(pvdid) pvdid  from (");
            sqlBuilder.AppendFormat("select faccountid yclAcc_id,0 yjsjAcc_id,0 yfzkAcc_id,0 pvdid from t_Account where Fnumber ='{0}'", yclAcc_number);
            sqlBuilder.Append(" union all(");
            sqlBuilder.AppendFormat("select 0 yclAcc_id,faccountid yjsjAcc_id,0 yfzkAcc_id,0 pvdid from t_Account where Fnumber ='{0}'", yjsjAcc_number);
            sqlBuilder.Append(" ) union all (");
            sqlBuilder.AppendFormat("select 0 yclAcc_id,0 yjsjAcc_id,faccountid yfzkAcc_id,0 pvdid from t_Account where Fnumber ='{0}'", yfzkAcc_number);
            sqlBuilder.Append(" ) union all (");
    //根据业务系统的供应商编号匹配
   //         sqlBuilder.AppendFormat("select 0 yclAcc_id,0 yjsjAcc_id,0 yfzkAcc_id,fitemid pvdid from t_item where fitemclassid=8 and Fnumber ='{0}'", pvdsht);
    //根据业务系统的供应商名称匹配
            sqlBuilder.AppendFormat("select 0 yclAcc_id,0 yjsjAcc_id,0 yfzkAcc_id,fitemid pvdid from t_item where fitemclassid=8 and Fname ='{0}'", pvdtxt);
            sqlBuilder.Append(" )) t");
 
            DataTable dt = new DataTable();
            dt = DbHelper.GetDataTable(connstr, sqlBuilder.ToString());
            if (dt != null && dt.Rows.Count > 0)
            {
                yclAcc_id  = Convert.ToInt32(dt.Rows[0]["yclAcc_id"]);
                yjsjAcc_id = Convert.ToInt32(dt.Rows[0]["yjsjAcc_id"]);
                yfzkAcc_id = Convert.ToInt32(dt.Rows[0]["yfzkAcc_id"]);
                pvdid = Convert.ToInt32(dt.Rows[0]["pvdid"]);
            }
            //入库金蝶中没有此供应商，弹出错误，终止运行
            if (pvdid == 0)
            {

                return false ;
            }




            //取凭证的ID,当前序号，凭证号，项目核算ID
            sqlBuilder.Remove(0, sqlBuilder.Length);
            sqlBuilder.Append("select max(fvoucherid) fvoucherid,max(FSerialNum) FSerialNum,max(fnumber) fnumber,max(fdetailid)  fdetailid,max(fodetailid)  fodetailid from (");
            sqlBuilder.AppendFormat("select max(fvoucherid) fvoucherid,max(FSerialNum) FSerialNum,0 fnumber,0 fdetailid,0 fodetailid  from t_voucher");
            sqlBuilder.Append(" union all(");
            sqlBuilder.AppendFormat("select 0 fvoucherid,0 FSerialNum,max(fnumber) fnumber,0 fdetailid,0 fodetailid  from t_voucher where fyear={0} and fperiod={1} and FGroupID=1",fyear ,fperiod);
            sqlBuilder.Append(" ) union all (");
            sqlBuilder.AppendFormat("select 0 fvoucherid,0 FSerialNum,0 fnumber,max(fdetailid)  fdetailid,0 fodetailid from   t_itemdetail");
            sqlBuilder.Append(" ) union all (");
            sqlBuilder.AppendFormat("select 0 fvoucherid,0 FSerialNum,0 fnumber,0 fdetailid,max(fdetailid)  fodetailid from   t_itemdetail where FDetailCount=1 and f8={0}",pvdid);
            sqlBuilder.Append(" )) t");
            dt = new DataTable();
            dt = DbHelper.GetDataTable(connstr, sqlBuilder.ToString());
            if (dt != null && dt.Rows.Count > 0)
            {
                fvoucherid = Convert.ToInt32(dt.Rows[0]["fvoucherid"]);
                fSerialNum = Convert.ToInt32(dt.Rows[0]["FSerialNum"]);
                fnumber =    Convert.ToInt32(dt.Rows[0]["fnumber"]);
                fdetailid =  Convert.ToInt32(dt.Rows[0]["fdetailid"]);
                fodetailid =  Convert.ToInt32(dt.Rows[0]["fodetailid"]);
            }
			fvoucherid=fvoucherid+1;
			fSerialNum=fSerialNum+1;
            fnumber   = fnumber + 1;
            if(fodetailid ==0)
            {
              fdetailid=fdetailid+1 ;
            }
            else 
            {
              fdetailid=fodetailid;
            }



            DbHelper.Open(connstr);
            try
            {
                DbHelper.BeginTran(connstr);
                //如果K3没有此供应商，插入新供应商
                if (pvdid == 0)
                {
                    sqlBuilder.Remove(0, sqlBuilder.Length);
                    sqlBuilder.AppendFormat("insert Into t_Item(FItemClassID,FParentID,FLevel,FName,FNumber,FShortNumber,FFullNumber,FDetail) values(8,0,1,'{0}','{1}','{1}','{1}',1);", pvdtxt,pvdsht);
                    DbHelper.ExecuteNonQuery(connstr, sqlBuilder.ToString());
                    sqlBuilder.Remove(0, sqlBuilder.Length);
                    sqlBuilder.AppendFormat("SELECT FItemID FROM t_Item WHERE FItemClassID=8 AND FNumber='{0}';",pvdsht);
                    dt = new DataTable();
                    dt = DbHelper.GetDataTable(connstr, sqlBuilder.ToString());
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        pvdid = Convert.ToInt32(dt.Rows[0]["FItemID"]);
                    }

                    sqlBuilder.Remove(0, sqlBuilder.Length);
                    //商业
                   // sqlBuilder.AppendFormat("INSERT INTO com_Supplier(FFax,FAddress,FPhone,Fcorperate,Fdepartment,FBr,FTypeID,FCountry,FShortName,FCyID,FSetID,FBank,FContact,FRegionID,FProvince,FStockIDAssignee,FTaxNum,FCreditDays,FTrade,FCertify,FPermit,FAccount,FAPAccountID,FPayTaxAcctID,FLicence,FfavorPolicy,FPostalCode,FEmail,FPreAcctID,FRegmark,Femployee,FStatus,FminForeReceiveRate,FItemID)");
                   // sqlBuilder.AppendFormat("VALUES (NULL, NULL, NULL, NULL, 0, 0, 0, NULL, NULL, 0, 0, NULL, NULL, 0, NULL, 0, NULL, 0, 0, NULL, NULL, NULL, 0, 0, NULL, NULL, NULL, NULL, 0, NULL, 0, 0, 1.000000000000000e+000, {0})",pvdid);
                   //工业
                     sqlBuilder.AppendFormat("INSERT INTO t_Supplier (FFax,FAddress,FPhone,Fcorperate,Fdepartment,FCountry,FShortName,FCyID,FSetID,FBank,FContact,FRegionID,FProvince,FTaxNum,FCreditDays,FTrade,FAccount,FAPAccountID,FPayTaxAcctID,FfavorPolicy,FPostalCode,FEmail,FPreAcctID,FValueAddRate,Femployee,FStatus,FminForeReceiveRate,FNumber,FName,FShortNumber,FParentID,FItemID)");
                      sqlBuilder.AppendFormat("VALUES (NULL, NULL, NULL, NULL, 0, NULL, NULL, 0, 0, NULL, NULL, 0, NULL, NULL, 0, 0, NULL, 0, 0, NULL, NULL, NULL, 0, 1.700000000000000e+001, 0, 1072, 1.000000000000000e+000, '{0}', '{1}', '{0}', 0, {2})", pvdsht, pvdtxt, pvdid);
                  
                    
                    DbHelper.ExecuteNonQuery(connstr, sqlBuilder.ToString());
                
                }
                //插入项目核算	
                if (fodetailid == 0&&Const .YfzkHs=="1")
                {
                    sqlBuilder.Remove(0, sqlBuilder.Length);
                    sqlBuilder.AppendFormat("insert Into t_ItemDetail(fdetailid,FDetailCount,F8) values({0},1,{1});", fdetailid, pvdid);
                    sqlBuilder.AppendFormat("Insert Into t_ItemDetailV(FDetailID,FItemClassID,FItemID) values({0},8,{1});", fdetailid, pvdid);
                    DbHelper.ExecuteNonQuery(connstr, sqlBuilder.ToString());
                }
                sqlBuilder.Remove(0, sqlBuilder.Length);
                // 凭证原材料明细
                sqlBuilder.AppendFormat("INSERT INTO t_VoucherEntry (FVoucherID,FEntryID,FExplanation,FAccountID,FCurrencyID,FExchangeRate,FDC,FAmountFor,FAmount,FQuantity,FMeasureUnitID,FUnitPrice,FInternalInd,FAccountID2,FSettleTypeID,FSettleNo,FCashFlowItem,FTaskID,FResourceID,FTransNo,FDetailID)");
                sqlBuilder.AppendFormat(" select {0},{1},'{4}',{2},1,1.0,1,{3},{3},0,0,0,NULL, 0, 0, NULL, 0, 0, 0, NULL,0;", fvoucherid, fentryid, yclAcc_id, prs,fexp);
                fentryid = fentryid + 1;
                //凭证应交税金明细
                if (tax != 0)
                {
                    sqlBuilder.AppendFormat("INSERT INTO t_VoucherEntry (FVoucherID,FEntryID,FExplanation,FAccountID,FCurrencyID,FExchangeRate,FDC,FAmountFor,FAmount,FQuantity,FMeasureUnitID,FUnitPrice,FInternalInd,FAccountID2,FSettleTypeID,FSettleNo,FCashFlowItem,FTaskID,FResourceID,FTransNo,FDetailID)");
                    sqlBuilder.AppendFormat(" select {0},{1},'{5}',{2},1,1.0,1,{3},{3},0,0,0,NULL, 0, 0, NULL, 0, 0, 0, NULL,{4};", fvoucherid, fentryid, yjsjAcc_id, tax,fdetailid ,fexp);
                    fentryid = fentryid + 1;
                }
                //凭证贷方明细，应付账款
                sqlBuilder.AppendFormat("INSERT INTO t_VoucherEntry (FVoucherID,FEntryID,FExplanation,FAccountID,FCurrencyID,FExchangeRate,FDC,FAmountFor,FAmount,FQuantity,FMeasureUnitID,FUnitPrice,FInternalInd,FAccountID2,FSettleTypeID,FSettleNo,FCashFlowItem,FTaskID,FResourceID,FTransNo,FDetailID)");
                sqlBuilder.AppendFormat(" select {0},{1},'{5}',{2},1,1.0,0,{3},{3},0,0,0,NULL, 0, 0, NULL, 0, 0, 0, NULL,{4};", fvoucherid, fentryid, yfzkAcc_id, tprs, fdetailid,fexp);
                fentryid = fentryid + 1;
                //凭证主表
                sqlBuilder.AppendFormat("INSERT INTO t_Voucher (fvoucherid,FSerialNum,FDate,FTransDate,FYear,FPeriod,FGroupID,FNumber,FReference,");
                sqlBuilder.AppendFormat("FExplanation,FAttachments,FEntryCount,FDebitTotal,FCreditTotal,FInternalInd,FChecked,");
                sqlBuilder.AppendFormat("FPosted,FPreparerID,FCheckerID,FPosterID,FCashierID,FHandler,FObjectName,FParameter,");
                sqlBuilder.AppendFormat("FTranType,FOwnerGroupID) ");
                  // select :fvoucherid,:FSerialNum,:fdate,:fdate,:fyear,:fperiod,5,:fnumber,NULL,:fexp,0,:fentryid,:s_fgzje,:s_fgzje,NULL,0,0,16394,-1,-1,-1,NULL,NULL,NULL,0,1
                sqlBuilder.AppendFormat(" select {0},{1},'{2}','{2}',{3},{4},1,{5},NULL,'{8}',0,{6},{7},{7},NULL,0,0,16394,-1,-1,-1,NULL,NULL,NULL,0,1;", fvoucherid, fSerialNum, bdat, fyear, fperiod, fnumber, fentryid, tprs,fexp);
               
                DbHelper.ExecuteNonQuery(connstr, sqlBuilder.ToString());
                DbHelper.CommitTran(connstr);
                return true;
            }
            catch
            {
                DbHelper.RollbackTran(connstr);
                return false;
                
            }
            finally
            {
                DbHelper.Close(connstr);
            }
        }

    }
}
