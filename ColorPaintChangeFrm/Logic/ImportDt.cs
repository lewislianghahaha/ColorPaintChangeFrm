﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using ColorPaintChangeFrm.DB;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ColorPaintChangeFrm.Logic
{
    public class ImportDt
    {
        DtList dtList = new DtList();

        /// <summary>
        /// 打开及导入至DT
        /// </summary>
        /// <param name="fileAddress"></param>
        /// <param name="typeid">导入类型=>1:常规纵向EXCEL导入 2:带空白纵向EXCEL导入</param>
        /// <returns></returns>
        public DataTable OpenExcelImporttoDt(string fileAddress, int typeid)
        {
            var dt = new DataTable();
            var resultdt=new DataTable();

            try
            {
                //使用NPOI技术进行导入EXCEL至DATATABLE
                var importExcelDt = OpenExcelToDataTable(fileAddress,typeid);
                //将从EXCEL过来的记录集为空的行清除
                dt = RemoveEmptyRows(importExcelDt);
                //change date:20200619 某些场景需要
                //resultdt = DeleteInclue2Layer(dt);
            }
            catch (Exception ex)
            {
                //var a = ex.Message;
                dt.Rows.Clear();
                dt.Columns.Clear();
            }
            return  dt;//resultdt;
        }

        private DataTable OpenExcelToDataTable(string fileAddress,int typeid)
        {
            IWorkbook wk;

            #region 变量定义
            //定义列ID
            var colid = 0;
            //定义制造商
            var company = string.Empty;
            //定义车型
            var car = string.Empty;
            //定义涂层
            var tulayer = string.Empty;
            //定义颜色描述
            var colordescript = string.Empty;
            //定义内部色号
            var colorcode = string.Empty;
            //定义主配方(差异色)
            var colorcha = string.Empty;
            //定义颜色组别
            var colorgroup = string.Empty;
            //定义标准色号
            var standurd = string.Empty;
            //定义RGBValue
            var rgb = string.Empty;
            //定义版本日期
            var fordt = string.Empty;
            //定义层
            var layer = string.Empty;
            //定义制作人
            var user = string.Empty;
            //定义二维码编号
            var code = string.Empty;
            #endregion

            //创建表标题
            var dt = dtList.Get_Importdt();

            using (var fsRead = File.OpenRead(fileAddress))
            {
                wk = new XSSFWorkbook(fsRead);
                //获取第一个sheet
                var sheet = wk.GetSheetAt(0);
                //获取第一行
                //var hearRow = sheet.GetRow(0);

                //创建完标题后,开始从第二行起读取对应列的值
                for (var r = 1; r <= sheet.LastRowNum; r++)
                {
                    var result = false;
                    var dr = dt.NewRow();
                    //获取当前行(注:只能获取行中有值的项,为空的项不能获取;即row.Cells.Count得出的总列数就只会汇总"有值的列"之和)
                    var row = sheet.GetRow(r);
                    if (row == null) continue;

                    //定义总列数
                    colid = 16;

                    //判断若从EXCEL获取的‘内部色号’不为空,且与对应变量不相同时,就将变量初始化:目的:用于区分一个配方,且令下一个配方记录能正确插入数据
                   // var excelfordt = Convert.ToString(row.GetCell(9));
                    var excelcolorcode = Convert.ToString(row.GetCell(4));

                    if (dt.Rows.Count > 0)
                    {
                        if (!string.IsNullOrEmpty(excelcolorcode) && excelcolorcode != colorcode)
                        {
                            //if (!string.IsNullOrEmpty(excelfordt) && excelfordt != fordt)
                            //{
                                #region 初始化各变量
                                //定义制造商
                                company = "";
                                //定义车型
                                car = "";
                                //定义涂层
                                tulayer = "";
                                //定义颜色描述
                                colordescript = "";
                                //定义内部色号
                                colorcode = "";
                                //定义主配方(差异色)
                                colorcha = "";
                                //定义颜色组别
                                colorgroup = "";
                                //定义标准色号
                                standurd = "";
                                //定义RGBValue
                                rgb = "";
                                //定义版本日期
                                fordt = "";
                                //定义层
                                layer = "";
                                //定义制作人
                                user = "";
                                //定义二维码编号
                                code = "";
                                #endregion
                            //}
                        }
                    }
 

                    for (var j = 0; j < colid/*row.Cells.Count*/; j++)
                    {
                        //循环获取行中的单元格
                        var cell = row.GetCell(j);
                        var cellValue = GetCellValue(cell);
                        
                        if (cellValue == string.Empty)
                        {
                            #region 若为空,将对应变量赋值给指定的dr[j]内=>(注:typeid=2时才执行)
                            if (typeid == 2)
                            {
                                switch (j)
                                {
                                    case 0:
                                        dr[j] = company;
                                        break;
                                    case 1:
                                        dr[j] = car;
                                        break;
                                    case 2:
                                        dr[j] = tulayer;
                                        break;
                                    case 3:
                                        dr[j] = colordescript;
                                        break;
                                    case 4:
                                        dr[j] = colorcode;
                                        break;
                                    case 5:
                                        dr[j] = colorcha;
                                        break;
                                    case 6:
                                        dr[j] = colorgroup;
                                        break;
                                    case 7:
                                        dr[j] = standurd;
                                        break;
                                    case 8:
                                        dr[j] = rgb;
                                        break;
                                    case 9:
                                        dr[j] = fordt;
                                        break;
                                    case 10:
                                        dr[j] = layer;
                                        break;
                                    case 11:
                                        dr[j] = user;
                                        break;
                                    case 12:
                                        dr[j] = code;
                                        break;
                                }
                            }
                            #endregion
                            else
                            {
                                continue;
                            }
                        }
                        else
                        {
                            dr[j] = cellValue;

                            #region 若不为空,将对应的J赋值给对应的变量内=>(注:typeid=2时才执行)
                            if (typeid == 2)
                            {
                                switch (j)
                                {
                                    case 0:
                                        company = Convert.ToString(dr[j]);
                                        break;
                                    case 1:
                                        car = Convert.ToString(dr[j]);
                                        break;
                                    case 2:
                                        tulayer = Convert.ToString(dr[j]);
                                        break;
                                    case 3:
                                        colordescript = Convert.ToString(dr[j]);
                                        break;
                                    case 4:
                                        colorcode = Convert.ToString(dr[j]);
                                        break;
                                    case 5:
                                        colorcha = Convert.ToString(dr[j]);
                                        break;
                                    case 6:
                                        colorgroup = Convert.ToString(dr[j]);
                                        break;
                                    case 7:
                                        standurd = Convert.ToString(dr[j]);
                                        break;
                                    case 8:
                                        rgb = Convert.ToString(dr[j]);
                                        break;
                                    case 9:
                                        fordt = Convert.ToString(dr[j]);
                                        break;
                                    case 10:
                                        layer = Convert.ToString(dr[j]);
                                        break;
                                    case 11:
                                        user = Convert.ToString(dr[j]);
                                        break;
                                    case 12:
                                        code = Convert.ToString(dr[j]);
                                        break;
                                }
                            }
                            #endregion
                        }

                        //全为空就不取
                        if (dr[j].ToString() != "")
                        {
                            result = true;
                        }
                    }

                    if (result == true)
                    {
                        //把每行增加到DataTable
                        dt.Rows.Add(dr);
                    }
                }
            }
            return dt;
        }

        /// <summary>
        /// 检查单元格的数据类型并获其中的值
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        private static string GetCellValue(ICell cell)
        {
            if (cell == null)
                return string.Empty;
            switch (cell.CellType)
            {
                case CellType.Blank: //空数据类型 这里类型注意一下，不同版本NPOI大小写可能不一样,有的版本是Blank（首字母大写)
                    return string.Empty;
                case CellType.Boolean: //bool类型
                    return cell.BooleanCellValue.ToString();
                case CellType.Error:
                    return cell.ErrorCellValue.ToString();
                case CellType.Numeric: //数字类型
                    if (DateUtil.IsCellDateFormatted(cell))//日期类型
                    {
                        return cell.DateCellValue.ToString();
                    }
                    else //其它数字
                    {
                        return cell.NumericCellValue.ToString();

                    }

                case CellType.Unknown: //无法识别类型
                default: //默认类型                    
                    return cell.ToString();
                case CellType.String: //string 类型
                    return cell.StringCellValue;
                case CellType.Formula: //带公式类型
                    try
                    {
                        var e = new XSSFFormulaEvaluator(cell.Sheet.Workbook);
                        e.EvaluateInCell(cell);
                        return cell.ToString();
                    }
                    catch
                    {
                        return cell.NumericCellValue.ToString();
                    }
            }
        }

        /// <summary>
        ///  将从EXCEL导入的DATATABLE的空白行清空
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        protected DataTable RemoveEmptyRows(DataTable dt)
        {
            var removeList = new List<DataRow>();
            for (var i = 0; i < dt.Rows.Count; i++)
            {
                var isNull = true;
                for (var j = 0; j < dt.Columns.Count; j++)
                {
                    //将不为空的行标记为False
                    if (!string.IsNullOrEmpty(dt.Rows[i][j].ToString().Trim()))
                    {
                        isNull = false;
                    }
                }
                //将整行都为空白的记录进行记录
                if (isNull)
                {
                    removeList.Add(dt.Rows[i]);
                }
            }

            //将整理出来的所有空白行通过循环进行删除
            for (var i = 0; i < removeList.Count; i++)
            {
                dt.Rows.Remove(removeList[i]);
            }
            return dt;
        }

        /// <summary>
        /// 将sourcedt内包含层为2的记录清除 change date:20200619
        /// </summary>
        /// <param name="sourcedt"></param>
        /// <returns></returns>
        private DataTable DeleteInclue2Layer(DataTable sourcedt)
        {
            var dt = sourcedt.Clone();
            //获取不重复的内容(循环使用)
            var diffColorDt = GetDiffColorDt(sourcedt);

            //循环选取好的内容
            foreach (DataRow rows in diffColorDt.Rows)
            {
                //若rows[2]在diffColorDt内两行(表示有两层) 或只有一行但‘层’为2,即排除
                if(Check2(Convert.ToString(rows[0]),Convert.ToString(rows[1]),Convert.ToString(rows[2]),diffColorDt)) continue;

                var dtlrows = sourcedt.Select("制造商='" + rows[0] + "' and 版本日期='" + rows[1] + "' and 内部色号='" + rows[2] + "' and 层='" + rows[3] + "' and 涂层='" + rows[4] + "'");
                //循环将适合的记录插入至最终DT
                for (int i = 0; i < dtlrows.Length; i++)
                {
                    dt.Merge(ImportImportdt(dt,dtlrows[i]));
                }
            }

            #region Hide
            //for (int i = 0; i < sourcedt.Rows.Count; i++)
            //{
            //涂层为:pearl 3C  && 层:2的不要
            //注:若同一个配方包含两层的就不要
            //if (/*Convert.ToString(sourcedt.Rows[i][2]) != "pearl 3C" &&*/ Convert.ToInt32(sourcedt.Rows[i][10]) != 2)
            //{
            //    var newrow = dt.NewRow();
            //    for (int j = 0; j < sourcedt.Columns.Count; j++)
            //    {
            //        newrow[j] = sourcedt.Rows[i][j];
            //    }
            //    dt.Rows.Add(newrow);
            //}
            //}
            #endregion

            return dt;
        }

        /// <summary>
        /// 将适合的记录插入至最终输出的DT内
        /// </summary>
        /// <param name="resultdt"></param>
        /// <param name="rows"></param>
        /// <returns></returns>
        private DataTable ImportImportdt(DataTable resultdt, DataRow rows)
        {
            var newrow = resultdt.NewRow();
            for (var j = 0; j < resultdt.Columns.Count; j++)
            {
                newrow[j] = rows[j];
            }
            resultdt.Rows.Add(newrow);
            return resultdt;
        }

        /// <summary>
        /// 在diffColorDt内两行(表示有两层) 或只有一行但‘层’为2
        /// </summary>
        /// <param name="factory"></param>
        /// <param name="dt"></param>
        /// <param name="code"></param>
        /// <param name="diffdt"></param>
        /// <returns></returns>
        private bool Check2(string factory,string dt,string code,DataTable diffdt)
        {
            var result = false;

            var dtlrows = diffdt.Select("制造商='"+factory+"' and 版本日期='"+dt+"' and 内部色号='"+code+"'");
            //判断为TRUE条件:1)若dtlrows=1,且对应的‘层’为2 2)dtlrows有两行
            if (dtlrows.Length == 1 && Convert.ToString(dtlrows[0][3]) == "2")
            {
                result = true;
            }
            else if (dtlrows.Length > 1)
            {
                result = true;
            }
            return result;
        }

        /// <summary>
        /// 整理不相同的指定配方记录
        /// </summary>
        /// <param name="sourcedt"></param>
        /// <returns></returns>
        private DataTable GetDiffColorDt(DataTable sourcedt)
        {
            var resultdt = dtList.Get_ColorcodeDt();

            //定义‘制造商’变量
            var factory = string.Empty;
            //定义‘版本日期’变量
            var comdt = string.Empty;
            //定义‘内部色号’变量
            var colorcode = string.Empty;
            //定义‘层’变量
            var layer = string.Empty;
            //定义‘涂层’变量
            var tulayer = string.Empty;

            foreach (DataRow rows in sourcedt.Rows)
            {
                var newrow = resultdt.NewRow();
                if (colorcode == "")
                {
                    newrow[0] = rows[0];      //制造商
                    newrow[1] = rows[9];      //版本日期
                    newrow[2] = rows[4];      //内部色号
                    newrow[3] = rows[10];     //层
                    newrow[4] = rows[2];      //涂层

                    //对变量赋值
                    factory = Convert.ToString(rows[0]);
                    comdt = Convert.ToString(rows[9]);
                    colorcode = Convert.ToString(rows[4]);
                    layer = Convert.ToString(rows[10]);
                    tulayer = Convert.ToString(rows[2]);
                }
                else
                {
                    if (factory == Convert.ToString(rows[0]) && comdt == Convert.ToString(rows[9]) && colorcode == Convert.ToString(rows[4])
                        && layer == Convert.ToString(rows[10]) && tulayer == Convert.ToString(rows[2])) continue;

                    //当不相同时才添加
                    newrow[0] = rows[0];      //制造商
                    newrow[1] = rows[9];      //版本日期
                    newrow[2] = rows[4];      //内部色号 
                    newrow[3] = rows[10];     //层
                    newrow[4] = rows[2];      //涂层

                    //对变量赋值
                    factory = Convert.ToString(rows[0]);
                    comdt = Convert.ToString(rows[9]);
                    colorcode = Convert.ToString(rows[4]);
                    layer = Convert.ToString(rows[10]);
                    tulayer = Convert.ToString(rows[2]);
                }
                resultdt.Rows.Add(newrow);
            }
            return resultdt;
        }

    }
}
