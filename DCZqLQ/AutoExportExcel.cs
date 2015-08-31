using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Diagnostics;
using System.Text;
using System.Windows.Forms;
using System.IO;

using KY.AP.DataServices;
using KY.Data.DB;
using FI.Public;
using FI.DataAccess;
using System.Runtime.InteropServices;
using EQ2008_DataStruct;
using Aspose.Cells.Rendering;
using Aspose.Cells;
using System.Drawing.Imaging;

namespace KY.Fi.DCZqLQ
{
    public class AutoExportExcel
    {
        OleDbConnection cn = new OleDbConnection();
        private string ImgFilePath = @".\Resources\IMAGE\";
        private string IniFilePath = @".\EQ2008_Dll_Set.ini";
        #region EQ DLL INFO
        //添加文本区
        [DllImport("EQ2008_Dll.dll", CharSet = CharSet.Ansi)]
        public static extern int User_AddText(int CardNum, ref User_Text pText, int iProgramIndex);

        //添加节目
        [DllImport("EQ2008_Dll.dll", CharSet = CharSet.Ansi)]
        public static extern int User_AddProgram(int CardNum, Boolean bWaitToEnd, int iPlayTime);

        //删除所有节目
        [DllImport("EQ2008_Dll.dll", CharSet = CharSet.Ansi, EntryPoint = "User_DelAllProgram")]
        public static extern Boolean User_DelAllProgram(int CardNum);

        //添加时间区
        [DllImport("EQ2008_Dll.dll", CharSet = CharSet.Ansi)]
        public static extern int User_AddTime(int CardNum, ref User_DateTime pdateTime, int iProgramIndex);

        //发送数据
        [DllImport("EQ2008_Dll.dll", CharSet = CharSet.Ansi)]
        public static extern Boolean User_SendToScreen(int CardNum);

        //添加RTF区
        [DllImport("EQ2008_Dll.dll", CharSet = CharSet.Ansi)]
        public static extern int User_AddRTF(int CardNum, ref User_RTF pRTF, int iProgramIndex);

        //添加单行文本区
        [DllImport("EQ2008_Dll.dll", CharSet = CharSet.Ansi)]
        public static extern int User_AddSingleText(int CardNum, ref User_SingleText pSingleText, int iProgramIndex);

        //添加图文区
        [DllImport("EQ2008_Dll.dll", CharSet = CharSet.Ansi)]
        public static extern int User_AddBmpZone(int CardNum, ref User_Bmp pBmp, int iProgramIndex);

        //指定图像句柄添加图片
        [DllImport("EQ2008_Dll.dll", CharSet = CharSet.Ansi)]
        public static extern bool User_AddBmp(int CardNum, int iBmpPartNum, IntPtr hBitmap, ref User_MoveSet pMoveSet, int iProgramIndex);

        //指定图像路径添加图片
        [DllImport("EQ2008_Dll.dll", CharSet = CharSet.Ansi)]
        public static extern bool User_AddBmpFile(int CardNum, int iBmpPartNum, string strFileName, ref User_MoveSet pMoveSet, int iProgramIndex);

        //控制卡地址
        //public static int g_iCardNum = 1;    
        public static int g_iGreen = 0xFF00; //绿色
        public static int g_iYellow = 0xFFFF; //黄色
        public static int g_iRed = 0x00FF; //红色

        //颜色常量
        public static int RED = 0x0000FF;
        public static int GREEN = 0x00FF00;
        public static int YELLOW = 0x00FFFF;
        //返回值常量
        public static int EQ_FALSE = 0;
        public static int EQ_TRUE = 1;

        [System.Runtime.InteropServices.DllImport("gdi32.dll")]
        public static extern bool DeleteObject(IntPtr hObject);
        //节目索引
        //public static int g_iProgramIndex = 0;
        //public static int g_iProgramCount = 0; 
        #endregion

        public void ExportExcel(object sender, DoWorkEventArgs e)
        {
            int times = Const.jTimes * 1000 * 60;
            while (true)
            {
                try
                {
                    CBase.AddErroLog("导出数据开始时间_" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                    setDataBaseConnect();
                    string[] funcTxt = Const.Func.Split(',');
                    for (int i = 0; i < funcTxt.Length; i++)
                    {
                        switch (funcTxt[i])
                        {
                            case "cjxc":
                                CBase.AddErroLog("车间现场数据发送时间_" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                                ShowWorkShop();
                                break;
                            case "hys":
                                CBase.AddErroLog("化验室数据发送时间_" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                                ShowHYS();
                                break;
                            case "tjs":
                                CBase.AddErroLog("调浆室数据发送时间_" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                                ShowTJS();
                                break;
                            case "zjs":
                                CBase.AddErroLog("助剂室数据发送时间_" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                                ShowZJS();
                                break;
                            case "zws":
                                CBase.AddErroLog("制网室数据发送时间_" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                                ShowZWS();
                                break;
                            case "mgs":
                                CBase.AddErroLog("描稿室数据发送时间_" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                                ShowMGS();
                                break;
                            case "zzs":
                                CBase.AddErroLog("整装室数据发送时间_" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                                ShowZZS();
                                break;
                            case "hzl":
                                CBase.AddErroLog("后整理数据发送时间_" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                                ShowHZL();
                                break;
                        }
                    }

                }
                catch (System.Exception ex)
                {
                    CBase.AddErroLog(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  导出数据异常:" + ex.ToString());
                }
                System.Threading.Thread.Sleep(times);
            }
        }

        //设置链接
        private bool SetOledbConn(string path)
        {
            string connectString2007Foramt = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source='{0}';Extended Properties='Excel 12.0 Xml;IMEX=2;HDR=No'";
            string connectString2003Foramt = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source='{0}';Extended Properties= 'Excel 8.0;IMEX=2;HDR=No'";

            if (path.EndsWith(".xls", StringComparison.InvariantCultureIgnoreCase))
            {
                cn = new OleDbConnection(String.Format(connectString2003Foramt, path));
                return true;
            }
            else if (path.EndsWith(".xlsx", StringComparison.InvariantCultureIgnoreCase))
            {
                cn = new OleDbConnection(String.Format(connectString2007Foramt, path));
                return true;
            }
            else
            {
                CBase.AddErroLog(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  导出数据异常: 导出文件不是EXCEL文件");
                return false;
            }
        }

        //车间现场
        private void ShowWorkShop()
        {
            LedInfo LedWkp = LedInfo("cjxc");
            int RowNum = 0;
            int g_iProgramIndex_wkp = 0;
            WriteFileIni(LedWkp, IniFilePath);
            User_Bmp BmpZone = new User_Bmp();
            User_MoveSet MoveSet = new User_MoveSet();
            int iBMPZoneNum = 0;
            string workShopPath = ".\\Resources\\EXCEL\\车间现场.xls";
            string HeadImgPath = LedWkp.filePath + "WHead.bmp";
            //添加节目
            addPro(ref g_iProgramIndex_wkp, LedWkp.cardNum);
            addTxt(g_iProgramIndex_wkp, LedWkp.cardNum, LedWkp);
            addTime(g_iProgramIndex_wkp, LedWkp.cardNum, LedWkp);
            addImgZoneHead(g_iProgramIndex_wkp, LedWkp.cardNum, LedWkp, HeadImgPath);
            addImgZoneRoll(ref BmpZone, ref MoveSet, ref iBMPZoneNum, g_iProgramIndex_wkp, LedWkp);
            //生成图片                                 
            System.Data.DataTable dtMcnDataGatherByDay = LqImportDac.GetMcnDataGatherAllByDay(DateTime.Today.ToString("yyyy-MM-dd"));
            for (int i = 0; i < dtMcnDataGatherByDay.Rows.Count; i++)
            {
                ToExcelOfWorkShop(workShopPath, RowNum);
                Workbook wb = new Workbook(workShopPath);
                Worksheet ws = wb.Worksheets[0];
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();
                imgOptions.ImageFormat = System.Drawing.Imaging.ImageFormat.Jpeg;
                imgOptions.OnePagePerSheet = true;
                SheetRender sr = new SheetRender(ws, imgOptions);
                Bitmap bitmap = sr.ToImage(0);
                bitmap.Save(LedWkp.filePath + "W.jpg");
                CloneImage(LedWkp.cX, LedWkp.cY, LedWkp.cWidth, LedWkp.cHeight, LedWkp.filePath + "W.jpg", LedWkp.filePath + "W" + (i + 1) + ".gif");
                addImgbmp(ref iBMPZoneNum, ref MoveSet, g_iProgramIndex_wkp, LedWkp.filePath + "W" + (i + 1) + ".gif", LedWkp);
                RowNum = (i + 1) * 6;
                if (RowNum >= dtMcnDataGatherByDay.Rows.Count)
                    break;
            }
            //发送节目

            sendPro(LedWkp.cardNum);
            //删除节目
            deletePro(LedWkp.cardNum, ref g_iProgramIndex_wkp);
        }

        //化验室
        private void ShowHYS()
        {
            LedInfo LedHys = LedInfo("hys");
            int RowNum = 0;
            int g_iProgramIndex_hys = 0;
            WriteFileIni(LedHys, IniFilePath);
            User_Bmp BmpZone = new User_Bmp();
            User_MoveSet MoveSet = new User_MoveSet();
            int iBMPZoneNum = 0;
            string HysPath = ".\\Resources\\EXCEL\\化验室.xls";
            string HeadImgPath = LedHys.filePath + "HHead.gif";
            //添加节目
            addPro(ref g_iProgramIndex_hys, LedHys.cardNum);
            addTxt(g_iProgramIndex_hys, LedHys.cardNum, LedHys);
            addTime(g_iProgramIndex_hys, LedHys.cardNum, LedHys);
            addImgZoneHead(g_iProgramIndex_hys, LedHys.cardNum, LedHys, HeadImgPath);
            addImgZoneRoll(ref BmpZone, ref MoveSet, ref iBMPZoneNum, g_iProgramIndex_hys, LedHys);
            //生成图片                                 
            System.Data.DataTable dtHys = LqImportDac.GetHys_View();
            for (int i = 0; i < dtHys.Rows.Count; i++)
            {
                toExcelOfHYS(HysPath, RowNum);
                Workbook wb = new Workbook(HysPath);
                Worksheet ws = wb.Worksheets[0];
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();
                imgOptions.ImageFormat = System.Drawing.Imaging.ImageFormat.Jpeg;
                imgOptions.OnePagePerSheet = true;
                SheetRender sr = new SheetRender(ws, imgOptions);
                Bitmap bitmap = sr.ToImage(0);
                bitmap.Save(LedHys.filePath + "H.jpg");
                CloneImage(LedHys.cX, LedHys.cY, LedHys.cWidth, LedHys.cHeight, LedHys.filePath + "H.jpg", LedHys.filePath + "H" + (i + 1) + ".gif");
                addImgbmp(ref iBMPZoneNum, ref MoveSet, g_iProgramIndex_hys, LedHys.filePath + "H" + (i + 1) + ".gif", LedHys);
                RowNum = (i + 1) * 3;
                if (RowNum >= dtHys.Rows.Count)
                    break;
            }
            //发送节目
            sendPro(LedHys.cardNum);
            //删除节目
            deletePro(LedHys.cardNum, ref g_iProgramIndex_hys);
        }

        //调浆室1
        private void ShowTJS()
        {
            LedInfo LedTjs = LedInfo("tjs");
            int RowNum = 0;
            int g_iProgramIndex_tjs = 0;
            WriteFileIni(LedTjs, IniFilePath);
            User_Bmp BmpZone = new User_Bmp();
            User_MoveSet MoveSet = new User_MoveSet();
            int iBMPZoneNum = 0;
            string TjsPath = ".\\Resources\\EXCEL\\调浆室.xls";
            string HeadImgPath = LedTjs.filePath + "THead.gif";
            //添加节目
            addPro(ref g_iProgramIndex_tjs, LedTjs.cardNum);
            addTxt(g_iProgramIndex_tjs, LedTjs.cardNum, LedTjs);
            addTime(g_iProgramIndex_tjs, LedTjs.cardNum, LedTjs);
            addImgZoneHead(g_iProgramIndex_tjs, LedTjs.cardNum, LedTjs, HeadImgPath);
            addImgZoneRoll(ref BmpZone, ref MoveSet, ref iBMPZoneNum, g_iProgramIndex_tjs, LedTjs);
            //生成图片                                 
            System.Data.DataTable dtTjs = LqImportDac.GetTjs_View();
            for (int i = 0; i < dtTjs.Rows.Count; i++)
            {
                toExcelOfTJS(TjsPath, RowNum);
                Workbook wb = new Workbook(TjsPath);
                Worksheet ws = wb.Worksheets[0];
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();
                imgOptions.ImageFormat = System.Drawing.Imaging.ImageFormat.Jpeg;
                imgOptions.OnePagePerSheet = true;
                SheetRender sr = new SheetRender(ws, imgOptions);
                Bitmap bitmap = sr.ToImage(0);
                bitmap.Save(LedTjs.filePath + "T.jpg");
                CloneImage(LedTjs.cX, LedTjs.cY, LedTjs.cWidth, LedTjs.cHeight, LedTjs.filePath + "T.jpg", LedTjs.filePath + "T" + (i + 1) + ".gif");
                addImgbmp(ref iBMPZoneNum, ref MoveSet, g_iProgramIndex_tjs, LedTjs.filePath + "T" + (i + 1) + ".gif", LedTjs);
                RowNum = (i + 1) * 3;
                if (RowNum >= dtTjs.Rows.Count)
                    break;
            }
            //发送节目
            sendPro(LedTjs.cardNum);
            //删除节目
            deletePro(LedTjs.cardNum, ref g_iProgramIndex_tjs);
            FileINIOpr fi = new FileINIOpr();
            CBase.AddErroLog(fi.GetIniKeyValue("地址：0", "IpAddress3", IniFilePath));
        }

        //调浆室2楼助剂
        private void ShowZJS()
        {
            LedInfo LedZjs = LedInfo("zjs");
            int RowNum = 0;
            int g_iProgramIndex_zjs = 0;
            WriteFileIni(LedZjs, IniFilePath);
            User_Bmp BmpZone = new User_Bmp();
            User_MoveSet MoveSet = new User_MoveSet();
            int iBMPZoneNum = 0;
            string ZjsPath = ".\\Resources\\EXCEL\\助剂室.xls";
            string HeadImgPath = LedZjs.filePath + "ZJHead.gif";
            //添加节目
            addPro(ref g_iProgramIndex_zjs, LedZjs.cardNum);
            addTxt(g_iProgramIndex_zjs, LedZjs.cardNum, LedZjs);
            addTime(g_iProgramIndex_zjs, LedZjs.cardNum, LedZjs);
            addImgZoneHead(g_iProgramIndex_zjs, LedZjs.cardNum, LedZjs, HeadImgPath);
            addImgZoneRoll(ref BmpZone, ref MoveSet, ref iBMPZoneNum, g_iProgramIndex_zjs, LedZjs);
            //生成图片                                 
            System.Data.DataTable dtZjs = LqImportDac.GetZjs_View();
            for (int i = 0; i < dtZjs.Rows.Count; i++)
            {
                toExcelOfZJS(ZjsPath, RowNum);
                Workbook wb = new Workbook(ZjsPath);
                Worksheet ws = wb.Worksheets[0];
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();
                imgOptions.ImageFormat = System.Drawing.Imaging.ImageFormat.Jpeg;
                imgOptions.OnePagePerSheet = true;
                SheetRender sr = new SheetRender(ws, imgOptions);
                Bitmap bitmap = sr.ToImage(0);
                bitmap.Save(LedZjs.filePath + "ZJ.jpg");
                CloneImage(LedZjs.cX, LedZjs.cY, LedZjs.cWidth, LedZjs.cHeight, LedZjs.filePath + "ZJ.jpg", LedZjs.filePath + "ZJ" + (i + 1) + ".gif");
                addImgbmp(ref iBMPZoneNum, ref MoveSet, g_iProgramIndex_zjs, LedZjs.filePath + "ZJ" + (i + 1) + ".gif", LedZjs);
                RowNum = (i + 1) * 3;
                if (RowNum >= dtZjs.Rows.Count)
                    break;
            }
            //发送节目
            sendPro(LedZjs.cardNum);
            //删除节目
            deletePro(LedZjs.cardNum, ref g_iProgramIndex_zjs);
        }

        //制网室
        private void ShowZWS()
        {
            LedInfo LedZws = LedInfo("zws");
            int RowNum = 0;
            int g_iProgramIndex_zws = 0;
            WriteFileIni(LedZws, IniFilePath);
            User_Bmp BmpZone = new User_Bmp();
            User_MoveSet MoveSet = new User_MoveSet();
            int iBMPZoneNum = 0;
            string ZwsPath = ".\\Resources\\EXCEL\\制网室.xls";
            string HeadImgPath = LedZws.filePath + "ZWHead.gif";
            //添加节目
            addPro(ref g_iProgramIndex_zws, LedZws.cardNum);
            addTxt(g_iProgramIndex_zws, LedZws.cardNum, LedZws);
            addTime(g_iProgramIndex_zws, LedZws.cardNum, LedZws);
            addImgZoneHead(g_iProgramIndex_zws, LedZws.cardNum, LedZws, HeadImgPath);
            addImgZoneRoll(ref BmpZone, ref MoveSet, ref iBMPZoneNum, g_iProgramIndex_zws, LedZws);
            //生成图片                                 
            System.Data.DataTable dtZws = LqImportDac.GetZws_View();
            for (int i = 0; i < dtZws.Rows.Count; i++)
            {
                toExcelOfZWS(ZwsPath, RowNum);
                Workbook wb = new Workbook(ZwsPath);
                Worksheet ws = wb.Worksheets[0];
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();
                imgOptions.ImageFormat = System.Drawing.Imaging.ImageFormat.Jpeg;
                imgOptions.OnePagePerSheet = true;
                SheetRender sr = new SheetRender(ws, imgOptions);
                Bitmap bitmap = sr.ToImage(0);
                bitmap.Save(LedZws.filePath + "ZW.jpg");
                CloneImage(LedZws.cX, LedZws.cY, LedZws.cWidth, LedZws.cHeight, LedZws.filePath + "ZW.jpg", LedZws.filePath + "ZW" + (i + 1) + ".gif");
                addImgbmp(ref iBMPZoneNum, ref MoveSet, g_iProgramIndex_zws, LedZws.filePath + "ZW" + (i + 1) + ".gif", LedZws);
                RowNum = (i + 1) * 3;
                if (RowNum >= dtZws.Rows.Count)
                    break;
            }
            //发送节目
            sendPro(LedZws.cardNum);
            //删除节目
            deletePro(LedZws.cardNum, ref g_iProgramIndex_zws);                        
        }

        //描稿室
        private void ShowMGS()
        {
            LedInfo LedMgs = LedInfo("mgs");
            int RowNum = 0;
            int g_iProgramIndex_mgs = 0;
            WriteFileIni(LedMgs, IniFilePath);
            User_Bmp BmpZone = new User_Bmp();
            User_MoveSet MoveSet = new User_MoveSet();
            int iBMPZoneNum = 0;
            string MgsPath = ".\\Resources\\EXCEL\\描稿室.xls";
            string HeadImgPath = LedMgs.filePath + "MHead.gif";
            //添加节目
            addPro(ref g_iProgramIndex_mgs, LedMgs.cardNum);
            addTxt(g_iProgramIndex_mgs, LedMgs.cardNum, LedMgs);
            addTime(g_iProgramIndex_mgs, LedMgs.cardNum, LedMgs);
            addImgZoneHead(g_iProgramIndex_mgs, LedMgs.cardNum, LedMgs, HeadImgPath);
            addImgZoneRoll(ref BmpZone, ref MoveSet, ref iBMPZoneNum, g_iProgramIndex_mgs, LedMgs);
            //生成图片                                 
            System.Data.DataTable dtMgs = LqImportDac.GetMgs_View();
            for (int i = 0; i < dtMgs.Rows.Count; i++)
            {
                toExcelOfMGS(MgsPath, RowNum);
                Workbook wb = new Workbook(MgsPath);
                Worksheet ws = wb.Worksheets[0];
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();
                imgOptions.ImageFormat = System.Drawing.Imaging.ImageFormat.Jpeg;
                imgOptions.OnePagePerSheet = true;
                SheetRender sr = new SheetRender(ws, imgOptions);
                Bitmap bitmap = sr.ToImage(0);
                bitmap.Save(LedMgs.filePath + "M.jpg");
                CloneImage(LedMgs.cX, LedMgs.cY, LedMgs.cWidth, LedMgs.cHeight, LedMgs.filePath + "M.jpg", LedMgs.filePath + "M" + (i + 1) + ".gif");
                addImgbmp(ref iBMPZoneNum, ref MoveSet, g_iProgramIndex_mgs, LedMgs.filePath + "M" + (i + 1) + ".gif", LedMgs);
                RowNum = (i + 1) * 3;
                if (RowNum >= dtMgs.Rows.Count)
                    break;
            }
            //发送节目
            sendPro(LedMgs.cardNum);
            //删除节目
            deletePro(LedMgs.cardNum, ref g_iProgramIndex_mgs);
        }

        //整装室
        private void ShowZZS()
        {
            LedInfo LedZzs = LedInfo("zzs");
            int RowNum = 0;
            int g_iProgramIndex_zzs = 0;
            WriteFileIni(LedZzs, IniFilePath);
            User_Bmp BmpZone = new User_Bmp();
            User_MoveSet MoveSet = new User_MoveSet();
            int iBMPZoneNum = 0;
            string ZzsPath = ".\\Resources\\EXCEL\\整装室.xls";
            string HeadImgPath = LedZzs.filePath + "ZZHead.gif";
            //添加节目
            addPro(ref g_iProgramIndex_zzs, LedZzs.cardNum);
            addTxt(g_iProgramIndex_zzs, LedZzs.cardNum, LedZzs);
            addTime(g_iProgramIndex_zzs, LedZzs.cardNum, LedZzs);
            addImgZoneHead(g_iProgramIndex_zzs, LedZzs.cardNum, LedZzs, HeadImgPath);
            addImgZoneRoll(ref BmpZone, ref MoveSet, ref iBMPZoneNum, g_iProgramIndex_zzs, LedZzs);
            //生成图片                                 
            System.Data.DataTable dtZzs = LqImportDac.GetZzs_View();
            for (int i = 0; i < dtZzs.Rows.Count; i++)
            {
                toExcelOfZZS(ZzsPath, RowNum);
                Workbook wb = new Workbook(ZzsPath);
                Worksheet ws = wb.Worksheets[0];
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();
                imgOptions.ImageFormat = System.Drawing.Imaging.ImageFormat.Jpeg;
                imgOptions.OnePagePerSheet = true;
                SheetRender sr = new SheetRender(ws, imgOptions);
                Bitmap bitmap = sr.ToImage(0);
                bitmap.Save(LedZzs.filePath + "ZZ.jpg");
                CloneImage(LedZzs.cX, LedZzs.cY, LedZzs.cWidth, LedZzs.cHeight, LedZzs.filePath + "ZZ.jpg", LedZzs.filePath + "ZZ" + (i + 1) + ".gif");
                addImgbmp(ref iBMPZoneNum, ref MoveSet, g_iProgramIndex_zzs, LedZzs.filePath + "ZZ" + (i + 1) + ".gif", LedZzs);
                RowNum = (i + 1) * 3;
                if (RowNum >= dtZzs.Rows.Count)
                    break;
            }
            //发送节目
            sendPro(LedZzs.cardNum);
            //删除节目
            deletePro(LedZzs.cardNum, ref g_iProgramIndex_zzs);
        }

        //后整理助剂
        private void ShowHZL()
        {
            LedInfo LedHzl = LedInfo("hzl");
            int RowNum = 0;
            int g_iProgramIndex_hzl = 0;
            WriteFileIni(LedHzl, IniFilePath);
            User_Bmp BmpZone = new User_Bmp();
            User_MoveSet MoveSet = new User_MoveSet();
            int iBMPZoneNum = 0;
            string HzlPath = ".\\Resources\\EXCEL\\后整理助剂.xls";
            string HeadImgPath = LedHzl.filePath + "HzHead.gif";
            //添加节目
            addPro(ref g_iProgramIndex_hzl, LedHzl.cardNum);
            addTxt(g_iProgramIndex_hzl, LedHzl.cardNum, LedHzl);
            addTime(g_iProgramIndex_hzl, LedHzl.cardNum, LedHzl);
            addImgZoneHead(g_iProgramIndex_hzl, LedHzl.cardNum, LedHzl, HeadImgPath);
            addImgZoneRoll(ref BmpZone, ref MoveSet, ref iBMPZoneNum, g_iProgramIndex_hzl, LedHzl);
            //生成图片                                 
            System.Data.DataTable dtHzl = LqImportDac.GetHzl_View();
            for (int i = 0; i < dtHzl.Rows.Count; i++)
            {
                toExcelOfHZL(HzlPath, RowNum);
                Workbook wb = new Workbook(HzlPath);
                Worksheet ws = wb.Worksheets[0];
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();
                imgOptions.ImageFormat = System.Drawing.Imaging.ImageFormat.Jpeg;
                imgOptions.OnePagePerSheet = true;
                SheetRender sr = new SheetRender(ws, imgOptions);
                Bitmap bitmap = sr.ToImage(0);
                bitmap.Save(LedHzl.filePath + "Hz.jpg");
                CloneImage(LedHzl.cX, LedHzl.cY, LedHzl.cWidth, LedHzl.cHeight, LedHzl.filePath + "Hz.jpg", LedHzl.filePath + "Hz" + (i + 1) + ".gif");
                addImgbmp(ref iBMPZoneNum, ref MoveSet, g_iProgramIndex_hzl, LedHzl.filePath + "Hz" + (i + 1) + ".gif", LedHzl);
                RowNum = (i + 1) * 3;
                if (RowNum >= dtHzl.Rows.Count)
                    break;
            }
            //发送节目
            sendPro(LedHzl.cardNum);
            //删除节目
            deletePro(LedHzl.cardNum, ref g_iProgramIndex_hzl);
        }

        //车间现场Excel
        private void ToExcelOfWorkShop(string workShopPath, int rowNum)
        {
            //System.Data.DataTable dtMcnDataGatherByMon = LqImportDac.GetMcnDataGatherAllByMon(DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), rowNum * count);            
            //rowNum = rowNum + 1;
            if (!string.IsNullOrEmpty(workShopPath))
            {
                if (!SetOledbConn(workShopPath))
                    return;

                //System.Data.DataTable dtMcnDataGatherByDay = LqImportDac.GetMcnDataGatherAllByDayAndRownum(DateTime.Today.ToString("yyyy-MM-dd"), rowNum);
                DataTable dtMcnDataGatherByDay = LqImportDac.GetMcnDataGatherAllByDay(DateTime.Today.ToString("yyyy-MM-dd"));
                DataTable dtMcnDataGatherByMon = LqImportDac.GetMcnDataGatherAllByMon(DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString());
                //rowNum = rowNum == 1 ? 7 : rowNum - 1;
                //System.Data.DataTable dtMcnDataGatherByMon = LqImportDac.GetMcnDataGatherAllByMon(DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), (rowNum - 1) * 2);

                OleDbCommand cmd = new OleDbCommand("SELECT * FROM [Sheet1$] WHERE 1<>1 ", cn);
                cn.Open();
                if (dtMcnDataGatherByDay != null || dtMcnDataGatherByDay.Rows.Count > 0)
                {
                    int count = 0;
                    for (int i = rowNum; i < rowNum + 6; i++)
                    {
                        if (i == dtMcnDataGatherByDay.Rows.Count)
                            break;
                        DataRow dr = dtMcnDataGatherByDay.Rows[i];
                        string index = (count + 1).ToString();
                        if (i == 0 || i % 3 == 0)
                        {
                            string sql = "UPDATE [Sheet1$" + "A" + index + ":A" + index + "] SET F1= '" + dtMcnDataGatherByDay.Rows[i]["txt"].ToString() + "' ";
                            cmd.CommandText = sql;
                            cmd.ExecuteNonQuery();
                        }

                        StringBuilder sb = new StringBuilder(200);
                        sb.AppendFormat("UPDATE [Sheet1$" + "B" + index + ":L" + index + "] SET F1='{0}',F2='{1}',F4='{2}',F6='{3}',F8='{4}',F10='{5}'",
                            dr["ban"].ToString(), dr["oprSumQty"].ToString(), dr["watQty"].ToString(), dr["elecQty"].ToString(), dr["vapQty"].ToString(),
                            dr["jnQty"].ToString());
                        cmd.CommandText = sb.ToString();
                        cmd.ExecuteNonQuery();
                        count = count + 1;
                    }
                }

                if (dtMcnDataGatherByMon != null || dtMcnDataGatherByMon.Rows.Count > 0)
                {
                    int count = 0;
                    for (int i = rowNum; i < rowNum + 6; i++)
                    {
                        if (i == dtMcnDataGatherByMon.Rows.Count)
                            break;
                        DataRow dr = dtMcnDataGatherByMon.Rows[i];
                        string index = (1 + count).ToString();
                        StringBuilder sb = new StringBuilder(200);
                        sb.AppendFormat("UPDATE [Sheet1$D{5}:L{5}] SET F1='{0}',F3='{1}',F5='{2}',F7='{3}',F9='{4}'", dr["oprSumQty"].ToString(),
                            dr["watQty"].ToString(), dr["elecQty"].ToString(), dr["vapQty"].ToString(), dr["jnQty"].ToString(), index);

                        cmd.CommandText = sb.ToString();
                        cmd.ExecuteNonQuery();
                        count = count + 1;
                    }
                }
                cn.Close();
                KillProcess("excel", workShopPath);
            }
        }

        //化验室Excel
        private void toExcelOfHYS(string HYSPath, int rowNum)
        {
            if (!string.IsNullOrEmpty(HYSPath))
            {
                if (!SetOledbConn(HYSPath))
                    return;

                DataTable dtHys = LqImportDac.GetHys_View();
                OleDbCommand cmd = new OleDbCommand("SELECT * FROM [Sheet1$] WHERE 1<>1 ", cn);
                cn.Open();

                if (dtHys != null || dtHys.Rows.Count > 0)
                {
                    int count = 0;
                    for (int i = rowNum; i < rowNum + 3; i++)
                    {
                        if (i == dtHys.Rows.Count)
                            break;
                        DataRow dr = dtHys.Rows[i];
                        string index = (1 + count).ToString();
                        StringBuilder sb = new StringBuilder(200);
                        sb.AppendFormat("UPDATE [Sheet1$A{5}:F{5}] SET F1='{0}',F2='{1}',F3='{2}',F4='{3}',F5='{4}',F6='{6}'", dr["订单"].ToString(),
                            dr["花型"].ToString(), dr["打样员"].ToString(), dr["计划生产时间"].ToString(), dr["打样完成时间"].ToString(), index, dr["状态"].ToString());

                        cmd.CommandText = sb.ToString();
                        cmd.ExecuteNonQuery();
                        count = count + 1;
                    }
                }
                cn.Close();
                KillProcess("excel", HYSPath);
            }
        }

        //调浆室Excel
        private void toExcelOfTJS(string TJSPath, int rowNum)
        {
            if (!string.IsNullOrEmpty(TJSPath))
            {
                if (!SetOledbConn(TJSPath))
                    return;

                DataTable dtTjs = LqImportDac.GetTjs_View();
                OleDbCommand cmd = new OleDbCommand("SELECT * FROM [Sheet1$] WHERE 1<>1 ", cn);
                cn.Open();

                if (dtTjs != null || dtTjs.Rows.Count > 0)
                {
                    int count = 0;
                    for (int i = rowNum; i < rowNum + 3; i++)
                    {
                        if (i == dtTjs.Rows.Count)
                            break;
                        DataRow dr = dtTjs.Rows[i];
                        string index = (1 + count).ToString();
                        StringBuilder sb = new StringBuilder(200);
                        sb.AppendFormat("UPDATE [Sheet1$A{5}:F{5}] SET F1='{0}',F2='{1}',F3='{2}',F4='{3}',F5='{4}',F6='{6}'", dr["订单"].ToString(),
                            dr["花型"].ToString(), dr["计划用料时间"].ToString(), dr["计划产量"].ToString(), dr["计划机台"].ToString(), index, dr["状态"].ToString());

                        cmd.CommandText = sb.ToString();
                        cmd.ExecuteNonQuery();
                        count = count + 1;
                    }
                }
                cn.Close();
                KillProcess("excel", TJSPath);
            }
        }

        //助剂室Excel
        private void toExcelOfZJS(string ZJPath, int rowNum)
        {
            if (!string.IsNullOrEmpty(ZJPath))
            {
                if (!SetOledbConn(ZJPath))
                    return;

                DataTable dtTjs = LqImportDac.GetZjs_View();
                OleDbCommand cmd = new OleDbCommand("SELECT * FROM [Sheet1$] WHERE 1<>1 ", cn);
                cn.Open();

                if (dtTjs != null || dtTjs.Rows.Count > 0)
                {
                    int count = 0;
                    for (int i = rowNum; i < rowNum + 3; i++)
                    {
                        if (i == dtTjs.Rows.Count)
                            break;
                        DataRow dr = dtTjs.Rows[i];
                        string index = (1 + count).ToString();
                        StringBuilder sb = new StringBuilder(200);
                        sb.AppendFormat("UPDATE [Sheet1$A{5}:F{5}] SET F1='{0}',F2='{1}',F3='{2}',F4='{3}',F5='{4}',F6='{6}'", dr["订单"].ToString(),
                            dr["花型"].ToString(), dr["计划产量"].ToString(), dr["计划用料时间"].ToString(), dr["状态"].ToString(), index, dr["计划机台"].ToString());

                        cmd.CommandText = sb.ToString();
                        cmd.ExecuteNonQuery();
                        count = count + 1;
                    }
                }
                cn.Close();
                KillProcess("excel", ZJPath);
            }
        }

        //制网室Excel
        private void toExcelOfZWS(string ZWSPath, int rowNum)
        {
            if (!string.IsNullOrEmpty(ZWSPath))
            {
                if (!SetOledbConn(ZWSPath))
                    return;

                DataTable dtZws = LqImportDac.GetZws_View();
                OleDbCommand cmd = new OleDbCommand("SELECT * FROM [Sheet1$] WHERE 1<>1 ", cn);
                cn.Open();

                if (dtZws != null || dtZws.Rows.Count > 0)
                {
                    int count = 0;
                    for (int i = rowNum; i < rowNum + 3; i++)
                    {
                        if (i == dtZws.Rows.Count)
                            break;
                        DataRow dr = dtZws.Rows[i];
                        string index = (1 + count).ToString();
                        StringBuilder sb = new StringBuilder(200);
                        sb.AppendFormat("UPDATE [Sheet1$A{5}:F{5}] SET F1='{0}',F2='{1}',F3='{2}',F4='{3}',F5='{4}',F6='{6}'", dr["订单"].ToString(),
                            dr["花型"].ToString(), dr["计划完成时间"].ToString(), dr["制网人员"].ToString(), dr["计划机台"].ToString(), index, dr["状态"].ToString());

                        cmd.CommandText = sb.ToString();
                        cmd.ExecuteNonQuery();
                        count = count + 1;
                    }
                }
                cn.Close();
                KillProcess("excel", ZWSPath);
            }
        }

        //描稿室Excel
        private void toExcelOfMGS(string MGSPath, int rowNum)
        {
            if (!string.IsNullOrEmpty(MGSPath))
            {
                if (!SetOledbConn(MGSPath))
                    return;

                DataTable dtMgs = LqImportDac.GetMgs_View();
                OleDbCommand cmd = new OleDbCommand("SELECT * FROM [Sheet1$] WHERE 1<>1 ", cn);
                cn.Open();

                if (dtMgs != null || dtMgs.Rows.Count > 0)
                {
                    int count = 0;
                    for (int i = rowNum; i < rowNum + 3; i++)
                    {
                        if (i == dtMgs.Rows.Count)
                            break;
                        DataRow dr = dtMgs.Rows[i];
                        string index = (1 + count).ToString();
                        StringBuilder sb = new StringBuilder(200);
                        sb.AppendFormat("UPDATE [Sheet1$A{5}:F{5}] SET F1='{0}',F2='{1}',F3='{2}',F4='{3}',F5='{4}',F6='{6}'", dr["订单"].ToString(),
                            dr["花型"].ToString(), dr["描稿人员"].ToString(), dr["计划完成时间"].ToString(), dr["实际完成时间"].ToString(), index, dr["状态"].ToString());

                        cmd.CommandText = sb.ToString();
                        cmd.ExecuteNonQuery();
                        count = count + 1;
                    }
                }
                cn.Close();
                KillProcess("excel", MGSPath);
            }
        }

        //整装室Excel
        private void toExcelOfZZS(string ZZSPath, int rowNum)
        {
            if (!string.IsNullOrEmpty(ZZSPath))
            {
                if (!SetOledbConn(ZZSPath))
                    return;

                DataTable dtZzs = LqImportDac.GetZzs_View();
                OleDbCommand cmd = new OleDbCommand("SELECT * FROM [Sheet1$] WHERE 1<>1 ", cn);
                cn.Open();

                if (dtZzs != null || dtZzs.Rows.Count > 0)
                {
                    int count = 0;
                    for (int i = rowNum; i < rowNum + 3; i++)
                    {
                        if (i == dtZzs.Rows.Count)
                            break;
                        DataRow dr = dtZzs.Rows[i];
                        string index = (1 + count).ToString();
                        StringBuilder sb = new StringBuilder(200);
                        sb.AppendFormat("UPDATE [Sheet1$A{5}:F{5}] SET F1='{0}',F2='{1}',F3='{2}',F4='{3}',F5='{4}',F6='{6}'", dr["订单"].ToString(),
                            dr["花型"].ToString(), dr["整装人员"].ToString(), dr["整装完成时间"].ToString(), dr["实际完成时间"].ToString(), index, dr["状态"].ToString());

                        cmd.CommandText = sb.ToString();
                        cmd.ExecuteNonQuery();
                        count = count + 1;
                    }
                }
                cn.Close();
                KillProcess("excel", ZZSPath);
            }
        }

        //后整理Excel
        private void toExcelOfHZL(string HZLPath, int rowNum)
        {

        }

        private void ToExcel()
        {
            if (!string.IsNullOrEmpty(Const.excelPath))
            {
                string connectString2007Foramt = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source='{0}';Extended Properties='Excel 12.0 Xml;IMEX=2;HDR=No'";
                string connectString2003Foramt = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source='{0}';Extended Properties= 'Excel 8.0;IMEX=2;HDR=No'";

                OleDbConnection cn;
                if (Const.excelPath.EndsWith(".xls", StringComparison.InvariantCultureIgnoreCase))
                {
                    cn = new OleDbConnection(String.Format(connectString2003Foramt, Const.excelPath));
                }
                else if (Const.excelPath.EndsWith(".xlsx", StringComparison.InvariantCultureIgnoreCase))
                {
                    cn = new OleDbConnection(String.Format(connectString2007Foramt, Const.excelPath));
                }
                else
                {
                    CBase.AddErroLog(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  导出数据异常: 到出文件不是EXCEL文件");
                    return;
                }


                OleDbCommand objCmdSelect = new OleDbCommand("SELECT f6,f8,f9 ,f7 FROM [Sheet1$]", cn);
                cn.Open();
                OleDbDataReader dr = objCmdSelect.ExecuteReader();
                string mName = "";
                string tName = "";
                string fNames = "";
                string xhao = "";
                StringBuilder sb = new StringBuilder(100);
                string reg1 = "";//流量
                string reg5 = "";//温度
                string reg4 = "";//压力
                string reg13 = "";//累计蒸汽


                string reg90 = "";


                System.Data.DataTable dt;

                while (dr.Read())
                {

                    try
                    {
                        if (dr.GetValue(1) == null || dr.GetValue(2) == null || dr.GetValue(3) == null)
                        {
                            continue;
                        }
                        // mName  = dr.GetString(0);  //模块

                        fNames = dr.GetValue(2).ToString(); //取值字段名
                        xhao = dr.GetValue(3).ToString(); //取序号 
                        tName = dr.GetValue(1).ToString(); //TT表名
                        // if (mName != "模块")
                        if (tName != "机台编号" && fNames != "取数据字段")
                        {
                            try
                            {
                                if (string.IsNullOrEmpty(tName) || string.IsNullOrEmpty(fNames))
                                {
                                    continue;
                                }

                                reg1 = "";
                                reg4 = "";
                                reg5 = "";
                                reg13 = "";
                                reg90 = "";


                                sb.Remove(0, sb.Length);
                                //sb.AppendFormat("select  top 1 reg1,reg5,reg4,reg13 from m_{0} order by Time desc", mName);
                                sb.AppendFormat("select  top 1 {0} from {1} order by dat desc", fNames, tName);
                                // CBase.AddErroLog(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "   正常1:" + sb.ToString() );
                                dt = new System.Data.DataTable();
                                dt = DBConn.GetSqlData(sb.ToString());
                                if (dt != null && dt.Rows.Count > 0)
                                {
                                    //reg1  = dt.Rows[0]["reg1"].ToString();
                                    //reg5  = dt.Rows[0]["reg5"].ToString();
                                    //reg4  = dt.Rows[0]["reg4"].ToString();
                                    //reg13 = dt.Rows[0]["reg13"].ToString();
                                    int cNum = dt.Columns.Count;

                                    sb.Remove(0, sb.Length);
                                    //sb.AppendFormat("UPDATE [Sheet1$] set f2={0},f3={1},f4={2},f5={3}  where f6='{4}'", reg1, reg5, reg4, reg13, mName);

                                    int dn = 0;
                                    sb.Append("UPDATE [Sheet1$] set ");
                                    //蒸汽流量
                                    if (cNum > 0 && dt.Rows[0][0] != null)
                                    {
                                        reg1 = dt.Rows[0][0].ToString();
                                        if (!string.IsNullOrEmpty(reg1))
                                        {
                                            sb.AppendFormat(" f2={0},", reg1);
                                            dn = dn + 1;
                                        }
                                    }
                                    //温度
                                    if (cNum > 1 && dt.Rows[0][1] != null)
                                    {
                                        reg5 = dt.Rows[0][1].ToString();
                                        if (!string.IsNullOrEmpty(reg5))
                                        {
                                            sb.AppendFormat(" f3={0},", reg5);
                                            dn = dn + 1;
                                        }
                                    }
                                    //压力
                                    if (cNum > 2 && dt.Rows[0][2] != null)
                                    {
                                        reg4 = dt.Rows[0][2].ToString();
                                        if (!string.IsNullOrEmpty(reg4))
                                        {
                                            sb.AppendFormat(" f4={0},", reg4);
                                            dn = dn + 1;
                                        }
                                    }
                                    //累计蒸汽
                                    if (cNum > 3 && dt.Rows[0][3] != null)
                                    {
                                        reg13 = dt.Rows[0][3].ToString();
                                        if (!string.IsNullOrEmpty(reg13))
                                        {
                                            sb.AppendFormat(" f5={0},", reg13);
                                            dn = dn + 1;
                                        }
                                    }
                                    //第5个数据
                                    if (cNum > 4 && dt.Rows[0][4] != null)
                                    {
                                        reg90 = dt.Rows[0][4].ToString();
                                        if (!string.IsNullOrEmpty(reg90))
                                        {
                                            sb.AppendFormat(" f6={0},", reg90);
                                            dn = dn + 1;
                                        }
                                    }

                                    if (dn > 0)
                                    {

                                        sb.AppendFormat("  f7=f7 where f7={0}", xhao);
                                        // CBase.AddErroLog(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "   正常2:" + sb.ToString());
                                        OleDbCommand cmd = new OleDbCommand(sb.ToString(), cn);
                                        cmd.ExecuteNonQuery();
                                    }
                                }
                            }
                            catch (System.Exception ex)
                            {
                                CBase.AddErroLog(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "   导出数据异常1:" + sb.ToString() + "___" + ex.ToString());
                            }
                        }
                    }
                    catch (System.Exception ex)
                    {
                        CBase.AddErroLog(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "   导出数据异常2:" + ex.ToString());
                    }
                }
                dr.Close();
                cn.Close();

                ////  string sqlCreate = "CREATE TABLE TestSheet ([ID] INTEGER,[Username] VarChar,[UserPwd] VarChar)"; 
                //OleDbCommand cmd = new OleDbCommand("select * from [Sheet1$] where 1<>1", cn);
                ////创建Excel文件：C:/test.xls
                //cn.Open();
                ////创建TestSheet工作表
                ////  cmd.ExecuteNonQuery();
                ////添加数据
                //cmd.CommandText = "INSERT INTO [Sheet1$](f1,f2,f3,f4) VALUES(1,2,3,'" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "')";
                //cmd.ExecuteNonQuery();
                ////关闭连接
                //cn.Close();


                //System.Diagnostics.Process.Start(Const.excelPath);
                Process[] curProcesses = Process.GetProcessesByName("excel");
                if (curProcesses.Length > 0)
                {
                    foreach (Process process1 in curProcesses)
                    {

                        int len = Const.excelPath.LastIndexOf("\\") + 1;
                        string filename = Const.excelPath.Substring(len, Const.excelPath.Length - len);
                        if (process1.MainWindowTitle.Contains(filename))
                        {
                            process1.Kill();
                        }

                    }
                }

                //ProcessStartInfo startInfo = new ProcessStartInfo("excel.exe");
                //startInfo.WindowStyle = ProcessWindowStyle.Minimized;
                //startInfo.UseShellExecute = true;
                ////startInfo.Arguments = Const.excelPath;
                //startInfo.FileName = Const.excelPath;
                //Process.Start(startInfo);


                Process excelProcess = new Process();
                excelProcess.StartInfo.FileName = "excel.exe";
                excelProcess.StartInfo.WindowStyle = ProcessWindowStyle.Minimized;
                excelProcess.StartInfo.Arguments = Const.excelPath;
                excelProcess.StartInfo.UseShellExecute = true;
                excelProcess.Start();



            }
        }

        //克隆剪切图片
        private void CloneImage(float x, float y, float width, float height, string jpgPath, string bmpPath)
        {
            //获取图像
            Bitmap myBitmap = new Bitmap(jpgPath);
            //设定图像剪切区域
            RectangleF cloneRect = new RectangleF(x, y, width, height);
            PixelFormat format = myBitmap.PixelFormat;
            Bitmap cloneBitmap = myBitmap.Clone(cloneRect, format);
            cloneBitmap.Save(bmpPath, System.Drawing.Imaging.ImageFormat.Gif);
            myBitmap.Dispose();
        }

        //屏幕信息
        private LedInfo LedInfo(string sht)
        {
            LedInfo led = new EQ2008_DataStruct.LedInfo();
            DataTable dtLed = LqImportDac.GetLedInfo(sht);
            if (dtLed.Rows.Count > 0)
            {
                DataRow dr = dtLed.Rows[0];
                led.sht = dr["sht"].ToString();
                led.txt = dr["txt"].ToString();
                led.cardNum = Convert.ToInt32(dr["cardNum"].ToString());
                led.sWidth = Convert.ToInt32(dr["sWidth"].ToString());
                led.sHeight = Convert.ToInt32(dr["sHeight"].ToString());
                led.tX = Convert.ToInt32(dr["tX"].ToString());
                led.tY = Convert.ToInt32(dr["tY"].ToString());
                led.tHeight = Convert.ToInt32(dr["tHeight"].ToString());
                led.tWidth = Convert.ToInt32(dr["tWidth"].ToString());
                led.tFont = Convert.ToInt32(dr["tFont"].ToString());
                led.dX = Convert.ToInt32(dr["dX"].ToString());
                led.dY = Convert.ToInt32(dr["dY"].ToString());
                led.dHeight = Convert.ToInt32(dr["dHeight"].ToString());
                led.dWidth = Convert.ToInt32(dr["dWidth"].ToString());
                led.dFont = Convert.ToInt32(dr["dFont"].ToString());
                led.hX = Convert.ToInt32(dr["hX"].ToString());
                led.hY = Convert.ToInt32(dr["hY"].ToString());
                led.hHeight = Convert.ToInt32(dr["hHeight"].ToString());
                led.hWidth = Convert.ToInt32(dr["hWidth"].ToString());
                led.rX = Convert.ToInt32(dr["rX"].ToString());
                led.rY = Convert.ToInt32(dr["rY"].ToString());
                led.rHeight = Convert.ToInt32(dr["rHeight"].ToString());
                led.rWidth = Convert.ToInt32(dr["rWidth"].ToString());
                led.cX = Convert.ToInt32(dr["cX"].ToString());
                led.cY = Convert.ToInt32(dr["cY"].ToString());
                led.cHeight = Convert.ToInt32(dr["cHeight"].ToString());
                led.cWidth = Convert.ToInt32(dr["cWidth"].ToString());
                led.filePath = ImgFilePath;
                led.CardType = Convert.ToInt32(dr["CardType"].ToString());
                led.CommunicationMode = Convert.ToInt32(dr["CommunicationMode"].ToString());
                led.IpAddress = dr["IpAddress"].ToString();
                led.ColorStyle = Convert.ToInt32(dr["ColorStyle"].ToString());
            }
            return led;

        }

        //添加节目
        private void addPro(ref int g_iProgramIndex, int g_iCardNum)
        {
            g_iProgramIndex = User_AddProgram(g_iCardNum, false, 10);

            //提示信息
            //string str = "当前节目号：";
            //g_iProgramIndex++;
            //string strProgramCount = Convert.ToString(g_iProgramIndex);
            //str = str + strProgramCount;
            //this.label1.Text = str;
        }

        //添加标题
        private void addTxt(int g_iProgramIndex, int g_iCardNum, LedInfo led)
        {
            User_Text Text = new User_Text();

            Text.BkColor = 0;
            Text.chContent = led.txt;

            Text.PartInfo.FrameColor = 0;
            Text.PartInfo.iFrameMode = 0;
            Text.PartInfo.iHeight = led.tHeight;
            Text.PartInfo.iWidth = led.tWidth;
            Text.PartInfo.iX = led.tX;
            Text.PartInfo.iY = led.tY;

            Text.FontInfo.bFontBold = false;
            Text.FontInfo.bFontItaic = false;
            Text.FontInfo.bFontUnderline = false;
            Text.FontInfo.colorFont = YELLOW;
            Text.FontInfo.iFontSize = led.tFont;
            Text.FontInfo.strFontName = "宋体";
            Text.FontInfo.iAlignStyle = 1;
            Text.FontInfo.iVAlignerStyle = 1;
            Text.FontInfo.iRowSpace = 0;

            Text.MoveSet.bClear = false;
            Text.MoveSet.iActionSpeed = 5;
            Text.MoveSet.iActionType = 1;
            Text.MoveSet.iHoldTime = 20;
            Text.MoveSet.iClearActionType = 1;
            Text.MoveSet.iClearSpeed = 4;
            Text.MoveSet.iFrameTime = 20;

            User_AddText(g_iCardNum, ref Text, g_iProgramIndex);

            //if (-1 == User_AddText(g_iCardNum, ref Text, g_iProgramIndex))
            //{
            //    MessageBox.Show("添加文本失败！");
            //}
            //else
            //{
            //    MessageBox.Show("添加文本成功！");
            //}
        }

        //添加时间
        private void addTime(int g_iProgramIndex, int g_iCardNum, LedInfo led)
        {
            User_DateTime DateTime = new User_DateTime();

            DateTime.bDay = true;
            DateTime.bHour = false;
            DateTime.BkColor = 0;
            DateTime.bMin = false;
            DateTime.bMouth = true;
            DateTime.bMulOrSingleLine = false;
            DateTime.bSec = false;
            DateTime.bWeek = false;
            DateTime.bYear = true;
            DateTime.bYearDisType = false;
            DateTime.chTitle = "";

            DateTime.PartInfo.iFrameMode = 0;
            DateTime.iDisplayType = 1;

            DateTime.PartInfo.FrameColor = 0xFFFF;
            DateTime.PartInfo.iHeight = led.dHeight;
            DateTime.PartInfo.iWidth = led.dWidth;
            DateTime.PartInfo.iX = led.dX;
            DateTime.PartInfo.iY = led.dY;

            DateTime.FontInfo.bFontBold = false;
            DateTime.FontInfo.bFontItaic = false;
            DateTime.FontInfo.bFontUnderline = false;
            DateTime.FontInfo.colorFont = 0xFF;
            DateTime.FontInfo.iAlignStyle = 1;
            DateTime.FontInfo.iFontSize = led.dFont;
            DateTime.FontInfo.strFontName = "宋体";

            User_AddTime(g_iCardNum, ref DateTime, g_iProgramIndex);
            //if (-1 == User_AddTime(g_iCardNum, ref DateTime, g_iProgramIndex))
            //{
            //    MessageBox.Show("添加时间失败！");
            //}
            //else
            //{
            //    MessageBox.Show("添加时间成功！");
            //}
        }

        //添加表头
        private void addImgZoneHead(int g_iProgramIndex, int g_iCardNum, LedInfo led, string HeadImgPath)
        {
            User_Bmp BmpZone = new User_Bmp();
            User_MoveSet MoveSet = new User_MoveSet();
            int iBMPZoneNum = 0;

            BmpZone.PartInfo.iX = led.hX;
            BmpZone.PartInfo.iY = led.hY;
            BmpZone.PartInfo.iHeight = led.hHeight;
            BmpZone.PartInfo.iWidth = led.hWidth;
            BmpZone.PartInfo.FrameColor = 0xFF00;
            BmpZone.PartInfo.iFrameMode = 0;

            MoveSet.bClear = true;
            MoveSet.iActionSpeed = 4;
            MoveSet.iActionType = 1;
            MoveSet.iHoldTime = 500000;
            MoveSet.iClearActionType = 1;
            MoveSet.iClearSpeed = 20;
            MoveSet.iFrameTime = 200000;

            iBMPZoneNum = User_AddBmpZone(g_iCardNum, ref BmpZone, g_iProgramIndex);

            string strBmpFile = HeadImgPath;
            User_AddBmpFile(g_iCardNum, iBMPZoneNum, strBmpFile, ref MoveSet, g_iProgramIndex);

            //if (false == User_AddBmpFile(g_iCardNum, iBMPZoneNum, strBmpFile, ref MoveSet, g_iProgramIndex))
            //{
            //    MessageBox.Show("添加表头图片失败！");
            //}
            //else
            //{
            //    MessageBox.Show("添加表头图片成功！");
            //}
        }

        //添加滚动内容
        private void addImgZoneRoll(ref User_Bmp BmpZone, ref User_MoveSet MoveSet, ref int iBMPZoneNum, int g_iProgramIndex, LedInfo led)
        {
            BmpZone.PartInfo.iX = led.rX;
            BmpZone.PartInfo.iY = led.rY;
            BmpZone.PartInfo.iHeight = led.rHeight;
            BmpZone.PartInfo.iWidth = led.rWidth;
            BmpZone.PartInfo.FrameColor = 0xFF00;
            BmpZone.PartInfo.iFrameMode = 0;

            MoveSet.bClear = true;
            MoveSet.iActionSpeed = 4;
            MoveSet.iActionType = 6;
            MoveSet.iHoldTime = 50;
            MoveSet.iClearActionType = 1;
            MoveSet.iClearSpeed = 4;
            MoveSet.iFrameTime = 20;

            iBMPZoneNum = User_AddBmpZone(led.cardNum, ref BmpZone, g_iProgramIndex);
        }

        //添加图片到图文区
        private void addImg(ref int iBMPZoneNum, ref User_MoveSet MoveSet, int g_iProgramIndex, string path, LedInfo led)
        {
            string strBmpFile = path;
            //User_AddBmpFile(led.cardNum, iBMPZoneNum, strBmpFile, ref MoveSet, g_iProgramIndex);

            if (false == User_AddBmpFile(led.cardNum, iBMPZoneNum, strBmpFile, ref MoveSet, g_iProgramIndex))
            {
                MessageBox.Show("添加滚动图片失败！");
            }
            else
            {
                MessageBox.Show("添加滚动图片成功！");
            }
        }

        //添加图片句柄
        private void addImgbmp(ref int iBMPZoneNum, ref User_MoveSet MoveSet, int g_iProgramIndex, string filePath, LedInfo LedWkp)
        {
            Bitmap btm = new Bitmap(filePath);
            IntPtr hBitmap = btm.GetHbitmap();

            User_AddBmp(LedWkp.cardNum, iBMPZoneNum, hBitmap, ref MoveSet, g_iProgramIndex);
            //if (false == User_AddBmp(LedWkp.cardNum, iBMPZoneNum, hBitmap, ref MoveSet, g_iProgramIndex))
            //{
            //    MessageBox.Show("添加图片句柄失败！");
            //}
            //else
            //{
            //    MessageBox.Show("添加图片句柄成功！");

            //}
            DeleteObject(hBitmap);
            btm.Dispose();
        }

        //发送数据
        private void sendPro(int g_iCardNum)
        {
            //User_SendToScreen(g_iCardNum);            
            if (User_SendToScreen(g_iCardNum) == false)
            {
                CBase.AddErroLog("  发送数据失败_" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            }
            else
            {
                //MessageBox.Show("发送节目成功！");
                CBase.AddErroLog("  发送数据成功_" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            }
        }

        //删除节目
        private void deletePro(int g_iCardNum, ref int g_iProgramIndex)
        {
            User_DelAllProgram(g_iCardNum);
            g_iProgramIndex = 0;
        }

        //杀Excel进程
        private void KillProcess(string type, string path)
        {
            Process[] curProcesses = Process.GetProcessesByName(type);
            if (curProcesses.Length > 0)
            {
                foreach (Process process1 in curProcesses)
                {
                    //if (process1.ProcessName == type)
                    //{
                    //    process1.Kill();
                    //}
                    int len = path.LastIndexOf("\\") + 1;
                    string filename = path.Substring(len, path.Length - len);
                    if (process1.MainWindowTitle.Contains(filename))
                    {
                        process1.Kill();
                    }

                }
            }
        }

        //写入配置文件
        private void WriteFileIni(LedInfo ledinfo, string path)
        {
            if (File.Exists(path))
            {
                try
                {
                    string[] IP = ledinfo.IpAddress.Split('.');
                    FileINIOpr fileIni = new FileINIOpr();
                    fileIni.SetIniKeyValue("地址：0", "CardType", ledinfo.CardType.ToString(), path);
                    fileIni.SetIniKeyValue("地址：0", "CardAddress", (ledinfo.cardNum - 1).ToString(), path);
                    fileIni.SetIniKeyValue("地址：0", "CommunicationMode", ledinfo.CommunicationMode.ToString(), path);
                    fileIni.SetIniKeyValue("地址：0", "ScreenHeight", ledinfo.sHeight.ToString(), path);
                    fileIni.SetIniKeyValue("地址：0", "ScreenWidth", ledinfo.sWidth.ToString(), path);
                    fileIni.SetIniKeyValue("地址：0", "IpAddress0", IP[0], path);
                    fileIni.SetIniKeyValue("地址：0", "IpAddress1", IP[1], path);
                    fileIni.SetIniKeyValue("地址：0", "IpAddress2", IP[2], path);
                    fileIni.SetIniKeyValue("地址：0", "IpAddress3", IP[3], path);
                    fileIni.SetIniKeyValue("地址：0", "ColorStyle", ledinfo.ColorStyle.ToString(), path);
                }
                catch (Exception ex)
                {
                    CBase.AddErroLog(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  写入配置文件异常:" + ex.ToString());
                }
            }
        }

        private Decimal NullToZero(object obj)
        {
            if (Convert.ToString(obj) == "")
            {
                return 0;
            }
            else
            {
                return Convert.ToDecimal(obj);
            }
        }

        #region 设置数据库连接
        public void setDataBaseConnect()
        {
            string connstr = DBConn.GetConnStr(Const.DBConnFile);
            if (connstr != "" && DBConn.TestConnection(connstr))
            {
                DBConn.SetSqlConn(Const.DBConnFile);
            }

        }
        #endregion

    }
}
