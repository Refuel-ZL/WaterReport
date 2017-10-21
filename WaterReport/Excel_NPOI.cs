using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;
using NPOI.SS.Util;
using System.Configuration;
using System.Data;

namespace WaterReport
{
    class Excel_NPOI
    {
       static String filepath=System.IO.Directory.GetCurrentDirectory()+"\\work\\";
        public static HSSFWorkbook CreateHead(DataSet data, Dictionary<int, string> varlist, Dictionary<int, string> waterlist, Dictionary<int, string>  waterflowlist, DateTime starttime, DateTime endtime,String when) {
            string[] headName = new string[] { "水泵", "开启时间", "开启时南水位", "开启时北水位", "关闭时间", "关闭时南水位", "关闭时北水位", "时间段内运行时间", "时间段内单程流量", "时间段内总运行时长", "时间段内总流量" };
            try
            {
                string val = ConfigurationManager.AppSettings["headnames"];
                string[] namelist = val.Split(new char[] { ',' });
                if (namelist.Length==11)
                {
                    headName = (String[])namelist.Clone();
                }
            }
            catch (Exception e)
            {
               
            }
               
            
            int[] columnWidth = { 10, 30, 15, 15, 30, 15, 15,20, 20, 20, 15 };
            string title = ConfigurationManager.AppSettings["title"];
            if (title == null || title == "") title = "{0}~{1}时间段内统计";
            String headTitle =string.Format(title, starttime, endtime, when);
            HSSFWorkbook workbook2003 = new HSSFWorkbook(); //新建工作簿 
            ISheet sheet = workbook2003.CreateSheet("Sheet0"); //创建Sheet页
           

            #region  如果为第一行
            IRow IRow = sheet.CreateRow(0);

            IRow.Height = 30 * 20;
            for (int h = 0; h < headName.Length; h++)
            {
                ICell Icell = IRow.CreateCell(h);
                Icell.SetCellValue(headTitle);
                ICellStyle style = workbook2003.CreateCellStyle();
                //设置单元格的样式：水平对齐居中

                style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;//文字水平对齐方式
                style.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;//文字垂直对齐方式
                style.WrapText = true;//自动换行
                //新建一个字体样式对象
                IFont font = workbook2003.CreateFont();
                font.FontName = "宋体";
                font.FontHeightInPoints = 18;
                //设置字体加粗样式
                font.Boldweight = (short)FontBoldWeight.Bold;
                //使用SetFont方法将字体样式添加到单元格样式中 
                style.SetFont(font);
                //将新的样式赋给单元格
                Icell.CellStyle = style;
                //合并单元格
            }
            sheet.AddMergedRegion(new CellRangeAddress(0, 0, 0, headName.Length-1));
            #endregion

            #region   表头
            IRow Irows2 = sheet.CreateRow(1);
            for (int j = 0; j < headName.Length; j++)
            {
                ICell Icell2 = Irows2.CreateCell(j);
                ICellStyle Istyle2 = workbook2003.CreateCellStyle();
                //设置边框
                Istyle2.BorderTop = BorderStyle.Thin;
                Istyle2.BorderBottom = BorderStyle.Thin;
                Istyle2.BorderLeft = BorderStyle.Thin;
                Istyle2.BorderRight = BorderStyle.Thin;
                //设置单元格的样式：水平对齐居中
                Istyle2.Alignment = HorizontalAlignment.Center;
                Istyle2.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;//文字垂直对齐方式
                //新建一个字体样式对象
                IFont Ifont2 = workbook2003.CreateFont();
                Ifont2.FontName = "宋体";
                Ifont2.FontHeightInPoints = 11;
                //设置字体加粗样式
                Ifont2.Boldweight = (short)FontBoldWeight.Bold;
               
                //使用SetFont方法将字体样式添加到单元格样式中 
                Istyle2.SetFont(Ifont2);
                //将新的样式赋给单元格
                Icell2.CellStyle = Istyle2;
                Icell2.SetCellValue(headName[j]);
                sheet.SetColumnWidth(j, 256 * columnWidth[j]);
            }
            #endregion
            #region   内容
            int m = 1;
            for (int i = 0; i < data.Tables.Count; i++) {
                int d = 0;
                d = m+1;
                Double T = 0;
                Double F = 0;
                for (int r = 0; r < data.Tables[i].Rows.Count; r = r + 2)
                {
                    DataRow mDr = data.Tables[i].Rows[r];
                    DataRow mDr_ = data.Tables[i].Rows[r + 1];
                    if (mDr["recordTime"].ToString() != "")
                    {
                        DateTime t1 = (DateTime)mDr["recordTime"];
                        if (DateTime.Compare(t1, starttime) < 0)
                        {
                            t1 = starttime;
                        }
                       
                        if (mDr_["recordTime"].ToString() == "") {
                            mDr_["recordTime"] = endtime;
                        }
                        DateTime t2 = (DateTime)mDr_["recordTime"];
                        TimeSpan ts = t2 - t1;
                        Math.Round(ts.TotalSeconds, MidpointRounding.AwayFromZero);
                        double w = Math.Round(ts.TotalSeconds, MidpointRounding.AwayFromZero);
                        T += w;
                        string pa = waterflowlist[int.Parse(mDr["varid"].ToString())];
                        if (w != 0 && pa != null && pa != "")
                        {
                            F += Math.Round(w / 3600.0 * Convert.ToDouble(pa), 2);
                        }
                    }
                    

                }
                for (int r=0; r < data.Tables[i].Rows.Count; r=r+2 )
                {
                    m++;
                    DataRow mDr = data.Tables[i].Rows[r];
                    DataRow mDr_ = data.Tables[i].Rows[r+1];
                   // Console.WriteLine(string.Format("{0},{1},{2},{3},{4}", mDr["varValue"].ToString(), mDr["recordTime"].ToString(), mDr_["varValue"].ToString(), mDr_["recordTime"].ToString(), mDr["varid"].ToString()));
                    IRow Irows3 = sheet.CreateRow(m);
                    Double w = 0;
                    for (int j = 0; j < headName.Length; j++)
                    {
                        ICell Icell3 = Irows3.CreateCell(j);

                        ICellStyle Istyle3 = workbook2003.CreateCellStyle();
                        //设置单元格的样式：水平对齐居中
                        Istyle3.Alignment = HorizontalAlignment.Center;
                        Istyle3.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;//文字垂直对齐方式

                        //设置边框
                        Istyle3.BorderTop = BorderStyle.Thin;
                        Istyle3.BorderBottom = BorderStyle.Thin;
                        Istyle3.BorderLeft = BorderStyle.Thin;
                        Istyle3.BorderRight = BorderStyle.Thin;
                        //新建一个字体样式对象
                        IFont Ifont3 = workbook2003.CreateFont();
                        Ifont3.FontName = "宋体";
                        Ifont3.FontHeightInPoints = 11;

                        //使用SetFont方法将字体样式添加到单元格样式中 
                        Istyle3.SetFont(Ifont3);
                        //将新的样式赋给单元格
                        Icell3.CellStyle = Istyle3;
                        if (j == 0) Icell3.SetCellValue(varlist[int.Parse(mDr["varid"].ToString())] );
                        if (j == 1) Icell3.SetCellValue(mDr["recordTime"].ToString());
                        if (j == 2) Icell3.SetCellValue(mDr["13"].ToString());
                        if (j == 3) Icell3.SetCellValue(mDr["14"].ToString());
                        if (j == 4) Icell3.SetCellValue(mDr_["recordTime"].ToString());
                        if (j == 5) Icell3.SetCellValue(mDr_["13"].ToString());
                        if (j == 6) Icell3.SetCellValue(mDr_["14"].ToString());
                        
                        if (j == 7) {
                            string val = "";
                            if (mDr["recordTime"].ToString()!="") {
                                DateTime t1 =(DateTime) mDr["recordTime"];
                                if (DateTime.Compare(t1,starttime)<0)
                                {
                                    t1 = starttime;
                                }
                                if (mDr_["recordTime"].ToString() == "")
                                {
                                    mDr_["recordTime"] = endtime;
                                }
                                DateTime t2 = (DateTime)mDr_["recordTime"];
                                TimeSpan ts = t2 - t1;
                                w = Math.Round(ts.TotalSeconds, MidpointRounding.AwayFromZero);
                                val = formatTs(w);
                            }
                            Icell3.SetCellValue(val);
                        }
                        if (j == 8) {
                            double val = 0;
                            string pa = waterflowlist[int.Parse(mDr["varid"].ToString())];
                            if (w!=0&& pa != null&& pa != "") {
                                val = Math.Round(w / 3600.0 * Convert.ToDouble(pa), 2) ;
                            }
                            Icell3.SetCellValue(val+"吨");
                        }
                        if (j == 9) {
                            Icell3.SetCellValue(formatTs(T));
                        }
                        if (j == 10) {
                            Icell3.SetCellValue(F+ "吨");
                        }
                    }
                }
                sheet.AddMergedRegion(new CellRangeAddress(d, m, 0, 0));
                sheet.AddMergedRegion(new CellRangeAddress(d, m, 9, 9));
                sheet.AddMergedRegion(new CellRangeAddress(d, m, 10, 10));
                m++;
               
                    for (int j = 0; j < headName.Length; j++)
                    {

                        IRow Irows4 = sheet.CreateRow(m);
                        ICell Icell4 = Irows4.CreateCell(j);
                        ICellStyle Istyle4 = workbook2003.CreateCellStyle();
                        Istyle4.BorderTop = BorderStyle.Thin;
                        Istyle4.BorderBottom = BorderStyle.Thin;
                        Istyle4.BorderLeft = BorderStyle.Thin;
                        Istyle4.BorderRight = BorderStyle.Thin;
                        Icell4.CellStyle = Istyle4;
                    }
                
                sheet.AddMergedRegion(new CellRangeAddress(m, m, 0, headName.Length - 1));

            }
            #endregion
            string Path = filepath;
            if (!System.IO.Directory.Exists(Path))
                System.IO.Directory.CreateDirectory(Path);
            string res = "天";
            switch (when)
            {
                case "Y":
                    res = "年";
                    break;
                case "M":
                    res = "月";
                    break;
                case "D":
                    res = "天";
                    break;
                case "W":
                    res = "周";
                    break;
            }
            string fileName = string.Format("{0} {1}内统计.xls", starttime.ToString("yyyy-MM-dd"), res);
            using (FileStream file = new FileStream(Path + "\\" + fileName, FileMode.Create))
            {
                workbook2003.Write(file);　　//创建test.xls文件。
                file.Close();
                workbook2003.Close();
            }
            return workbook2003;
        }

        public static string formatTs(Double duration) {
            TimeSpan ts = new TimeSpan(0, 0, Convert.ToInt32(duration));
            string str = "";
            if (ts.Hours > 0)
            {
                str = ts.Hours.ToString() + "小时 " + ts.Minutes.ToString() + "分钟 " + ts.Seconds + "秒";
            }
            if (ts.Hours == 0 && ts.Minutes > 0)
            {
                str = ts.Minutes.ToString() + "分钟 " + ts.Seconds + "秒";
            }
            if (ts.Hours == 0 && ts.Minutes == 0)
            {
                str = ts.Seconds + "秒";
            }
            return str;
        }
    }
}
