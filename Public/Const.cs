using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;


namespace FI.Public
{
    //放置所以产品统一的常量，比如登录信息，系统相关信息等。具体的产品常量可以从此类继承
    public class Const
    {
        /// <summary>
        /// 系统应用路径
        /// </summary>
        public static string StartupPath ="";
        /// <summary>
        /// 连接数据库配置文件
        /// </summary>
        public static string DBConnFile = StartupPath + "\\Resources\\DataFile.xml";
        /// <summary>
        /// 默认数据库
        /// </summary>
        public static string DataBaseName = "";
         /// <summary>
        /// 服务器数据库连接串
        /// </summary>
        public static string SqlConnentionString = "";

        /// <summary>
        /// 断网标志  true为断网
        /// </summary>
        public static bool NoInterNet = false;

        /// <summary>
        /// 是否随机启动  1:是，0;不是
        /// </summary>
        public static string  IsRunAsStart = "1";

        /// <summary>
        ///读取业务单据开始日期
        /// </summary>
        public static string  RSTime = ""; 

        /// <summary>
        ///开始运行生成凭证开始时间（小时）
        /// </summary>
        public static int stime = 0;
        /// <summary>
        ///开始运行生成凭证开始时间（分）
        /// </summary>
        public static int sminute = 0;

  
        /// <summary>
        ///程序运行模式
        /// </summary>
        public static string model = "自动运行";
 
        /// <summary>
        /// 同步其他数据库数据， 数据库连接串
        /// </summary>
        public static string SqlConnentionString1 = "";

//开源导数据使用
        /// <summary>
        /// 原材料科目
        /// </summary>
        public static string Ycl = "";
        /// <summary>
        /// 应交税金
        /// </summary>
        public static string Yjsj = "";
        /// <summary>
        /// 应交税金是否供应商核算  1:是，0;不是
        /// </summary>
        public static string YjsjHs = "1";
        /// <summary>
        /// 应付账款
        /// </summary>
        public static string Yfzk = "";
        /// <summary>
        /// 应付账款是否供应商核算  1:是，0;不是
        /// </summary>
        public static string YfzkHs = "1";


 
       //乐其使用
        /// <summary>
        /// 乐其是否同步上月的化料数据  1:是，0;不是
        /// </summary>
        public static string IsMonthLQ = "1"; 
        /// <summary>
        /// 重新同步上月单据的日期
        /// </summary>
        public static string LMdateLQ = "";

        /// <summary>
        ///  蒸汽写EXCEL间隔时间 分
        /// </summary>
        public static int  jTimes = 5;

        /// <summary>
        ///  功能选择
        /// </summary>
        public static string Func = "";

        /// <summary>
        ///  蒸汽写的EXCEL文件路径
        /// </summary>
        public static string excelPath= "";
        
    }
}
