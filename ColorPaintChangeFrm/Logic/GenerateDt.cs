using System;
using System.Data;
using ColorPaintChangeFrm.DB;
using NPOI.OpenXmlFormats.Dml;

namespace ColorPaintChangeFrm.Logic
{
    public class GenerateDt
    {
        DtList dtList=new DtList();

        private DataTable _tempdt;

        /// <summary>
        /// 运算
        /// </summary>
        /// <param name="selectid">1:纵向 2:横向</param>
        /// <param name="sourcedtdt"></param>
        /// <returns></returns>
        public DataTable GenerateExcelSourceDt(int selectid, DataTable sourcedtdt)
        {
            //定义总色母量(中间值)
            decimal sumtemp = 0;

            //从sourcedt内找出不相同的内部色号等记录
            GetColorCodeDt(dtList.Get_ColorcodeDt(), sourcedtdt);
            //创建导出临时表
            var resultdt = dtList.Get_ExportHdt();

            //根据内部色号 层 版本日期 涂层在sourcedt内找到相关记录,并计算其色母量之和
            foreach (DataRow rows in _tempdt.Rows)
            {
                var dtlrows = sourcedtdt.Select("内部色号='"+rows[0]+ "' and 层='"+rows[1]+ "' and 版本日期='"+rows[2]+ "' and 涂层='"+rows[3]+"'");
                //计算出色母量之和
                sumtemp = GenerateSumQty(dtlrows);

                for (var i = 0; i < dtlrows.Length; i++)
                {
                    //循环插入至resultdt临时表 色母量公式:Round(单个色母量/色母量之和*1000,2)
                    var newrow = resultdt.NewRow();
                    for (var j = 0; j < resultdt.Columns.Count; j++)
                    {
                        //计算色母量
                        if (j == 15)
                        {
                            newrow[j] = Math.Round(Convert.ToDecimal(dtlrows[i][j]) / sumtemp * 1000, 2);
                        }
                        else
                        {
                            //其它列操作
                            newrow[j] = dtlrows[i][j];
                        }
                    }
                    resultdt.Rows.Add(newrow);
                }
            }
            //根据下拉列表所选择的导出模式,进行改变其导出效果
            return MakeExportMode(selectid, resultdt);
        }

        /// <summary>
        /// 根据不同模式转换输出效果
        /// </summary>
        /// <param name="selectid">1:纵向 2:横向</param>
        /// <param name="sourcedt"></param>
        /// <returns></returns>
        private DataTable MakeExportMode(int selectid,DataTable sourcedt)
        {
            var resultdt=new DataTable();

            //纵向输出
            if (selectid == 1)
            {
                //获取纵向输出模板
                resultdt = sourcedt.Clone();
                //
                foreach (DataRow rows in _tempdt.Rows)
                {
                    var dtrows = sourcedt.Select("内部色号='" + rows[0] + "' and 层='" + rows[1] + "' and 版本日期='" + rows[2] + "' and 涂层='" + rows[3] + "'");

                    for (var i = 0; i < dtrows.Length; i++)
                    {
                        var newrow = resultdt.NewRow();

                        for (var j = 0; j < resultdt.Columns.Count; j++)
                        {
                            newrow[j] = i == 0 ? dtrows[i][j] : DBNull.Value;

                            if (j == 13 || j == 14 || j == 15)
                            {
                                newrow[j] = dtrows[i][j];
                            }
                        }
                        resultdt.Rows.Add(newrow);
                    }
                }
            }
            //横向输出
            else
            {
                //获取横向输出模板
                resultdt = dtList.Get_ExportVdt();

            }
            return resultdt;
        }

        /// <summary>
        /// 计算色母量之和
        /// </summary>
        /// <param name="rows"></param>
        /// <returns></returns>
        private decimal GenerateSumQty(DataRow[] rows)
        {
            decimal result = 0;

            for (var i = 0; i < rows.Length; i++)
            {
                result += Convert.ToDecimal(rows[i][15]);
            }
            return result;
        }

        /// <summary>
        /// 从sourcedt内找出不相同的内部色号等记录
        /// </summary>
        /// <param name="tempdt"></param>
        /// <param name="sourcedt"></param>
        /// <returns></returns>
        private void GetColorCodeDt(DataTable tempdt,DataTable sourcedt)
        {
            //定义‘内部色号’变量
            var colorcode = string.Empty;
            //定义‘层’变量
            var layer = string.Empty;
            //定义‘版本日期’变量
            var comdt = string.Empty;
            //定义‘涂层’变量
            var tulayer = string.Empty;

            //获取内部色号记录
            foreach (DataRow rows in sourcedt.Rows)
            {
                var newrow = tempdt.NewRow();
                if (colorcode == "")
                {
                    newrow[0] = rows[4];      //内部色号
                    newrow[1] = rows[10];     //层
                    newrow[2] = rows[9];      //版本日期
                    newrow[3] = rows[2];      //涂层

                    //对变量赋值
                    colorcode = Convert.ToString(rows[4]);
                    layer = Convert.ToString(rows[10]);
                    comdt = Convert.ToString(rows[9]);
                    tulayer = Convert.ToString(rows[2]);
                }
                else
                {
                    if (colorcode == Convert.ToString(rows[4]) && layer == Convert.ToString(rows[10]) && comdt == Convert.ToString(rows[9])
                        && tulayer == Convert.ToString(rows[2])) continue;
                    //当不相同时才添加
                    newrow[0] = rows[4];      //内部色号
                    newrow[1] = rows[10];     //层
                    newrow[2] = rows[9];      //版本日期
                    newrow[3] = rows[2];      //涂层

                    //对变量赋值
                    colorcode = Convert.ToString(rows[4]);
                    layer = Convert.ToString(rows[10]);
                    comdt = Convert.ToString(rows[9]);
                    tulayer = Convert.ToString(rows[2]);
                }
                tempdt.Rows.Add(newrow);
            }
            _tempdt = tempdt;
        }

    }
}
