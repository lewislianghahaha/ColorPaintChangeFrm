using System;
using System.Data;
using ColorPaintChangeFrm.DB;

namespace ColorPaintChangeFrm.Logic
{
    public class GenerateDt
    {
        DtList dtList=new DtList();

        private DataTable _tempdt;

        /// <summary>
        /// 运算
        /// </summary>
        /// <param name="genid">1:按KG进行计算色母量 2:按L进行计算色母量 0:不需计算色母量</param>
        /// <param name="selectid">1:纵向 2:横向</param>
        /// <param name="sortid">筛选条件(1:不筛选 2:获取1个色母数的配方 3:获取2个色母数的配方 4:获取3个色母数并包含PC-60的配方 5:获取包含PC-60并占比>=20%的配方)</param>
        /// <param name="sourcedtdt"></param>
        /// <returns></returns>
        public DataTable GenerateExcelSourceDt(int genid,int selectid, int sortid, DataTable sourcedtdt)
        {
            //定义总色母量(中间值)
            decimal sumtemp = 0;

            //从sourcedt内找出不相同的内部色号等记录
            GetColorCodeDt(dtList.Get_ColorcodeDt(), sourcedtdt);
            //创建导出临时表
            var resultdt = dtList.Get_ExportHdt();

            //根据 制造商 版本日期 内部色号 层 涂层 在sourcedt内找到相关记录,并计算其色母量之和
            foreach (DataRow rows in _tempdt.Rows)
            {
                //排序方式改为:制造商 版本日期 内部色号 层 涂层 
                var dtlrows = sourcedtdt.Select("制造商='"+rows[0]+ "' and 版本日期='"+rows[1]+ "' and 内部色号='"+rows[2]+ "' and 层='"+rows[3]+"' and 涂层='"+rows[4]+"'");

                //计算色母量之和
                sumtemp = GenerateSumQty(dtlrows);

                //筛选条件=>作用:根据所选择的筛选条件ID,来判断此dtlrows内的记录能否插入至resultdt内
                var checkmark = CheckMaterial(sortid,dtlrows,sumtemp);

                //计算出色母量之和(genid=>0:不需计算色母量 1:按KG进行计算色母量 2:按L进行计算色母量)
                //不需计算色母量(即直接将数据插入至resultdt内)
                if (genid == 0)
                {
                    for (var i = 0; i < dtlrows.Length; i++)
                    {
                        //若不满足条件(返回值为FALSE)就跳出循环
                        if (!checkmark) break;
                        resultdt.Merge(Generatedt(resultdt,dtlrows[i]));
                    }
                }
                //按KG进行计算色母量
                else if (genid == 1)
                {
                    for (var i = 0; i < dtlrows.Length; i++)
                    {
                        //若不满足条件(返回值为FALSE)就跳出循环
                        if (!checkmark) break;
                        //执行运算及整理至resultdt内
                        resultdt.Merge(GenerateKGdt(resultdt,dtlrows[i],sumtemp));                           
                    }
                }
                //按L进行计算色母量
                else
                {
                    for (var i = 0; i < dtlrows.Length; i++)
                    {
                        //若不满足条件(返回值为FALSE)就跳出循环
                        if (!checkmark) break;
                        //执行运算及整理至resultdt内
                        resultdt.Merge(GenerateLdt(resultdt,dtlrows[i]));
                    }
                }
            }
            //根据下拉列表所选择的导出模式,进行改变其导出效果
            return MakeExportMode(selectid, sortid,resultdt);
        }

        /// <summary>
        /// 
        /// 循环判断dtlrows内的增白剂 用量 是否适合导出条件
        /// </summary>
        /// <param name="sortid">筛选条件(1:不筛选 2:获取1个色母数的配方 3:获取2个色母数的配方 4:获取3个色母数并包含PC-60的配方 5:获取包含PC-60并占比>=20%的配方)</param>
        /// <param name="dtlrows"></param>
        /// <param name="sumtemp"></param>
        /// <returns></returns>
        private bool CheckMaterial(int sortid,DataRow [] dtlrows,decimal sumtemp)
        {
            var result = false;

            if (sortid == 1)
            {
                result = true;
            }
            else if (sortid == 2)
            {
                if (dtlrows.Length == 1)
                    result = true;
            }
            else if (sortid == 3)
            {
                if (dtlrows.Length == 2)
                    result = true;
            }
            else if (sortid == 4)
            {
                for (var i = 0; i < dtlrows.Length; i++)
                {
                    if (Convert.ToString(dtlrows[i][13]) != "PC-60") continue;
                    //若检测到配方总行数为3时才继续
                    if (dtlrows.Length==3)
                        result = true;
                        break;
                }
            }
            else
            {
                for (var i = 0; i < dtlrows.Length; i++)
                {
                    if (Convert.ToString(dtlrows[i][13]) != "PC-60") continue;
                    //使用“PC-60”对应的色母量/总色母量,若>=0.2才继续
                    if (Convert.ToDecimal(Convert.ToDecimal(dtlrows[i][15]) / sumtemp) >= Convert.ToDecimal(0.2))
                        result = true;
                        break;
                }
            }
            return result;
        }

        /// <summary>
        /// 将结果整理到resultdt内(以公斤会式来运算)
        /// </summary>
        /// <param name="resultdt"></param>
        /// <param name="rows"></param>
        /// <param name="sumtemp"></param>
        /// <returns></returns>
        private DataTable GenerateKGdt(DataTable resultdt,DataRow rows,decimal sumtemp)
        {
            var newrow = resultdt.NewRow();
            //循环插入至resultdt临时表 色母量公式(KG):公式=Round(单个色母量/色母量之和*1000,2)
            //change date:2020-04-08 公式=Round(单个色母量/色母量之和*100,2)
            for (var j = 0; j < resultdt.Columns.Count; j++)
            {
                //表示到最后一列=>将内部色号+"&"+版本日期
                if (j == resultdt.Columns.Count - 1)
                {
                    newrow[j] = Convert.ToString(rows[4] + "&" + Convert.ToString(rows[9]));
                }
                //计算色母量
                else if (j == 15)
                {
                    newrow[j] = Math.Round(Convert.ToDecimal(rows[j]) / sumtemp * 100, 2);
                }
                else
                {
                    //其它列操作
                    newrow[j] = rows[j];
                }
            }
            resultdt.Rows.Add(newrow);
            return resultdt;
        }

        /// <summary>
        /// 将结果整理到resultdt内(以L会式来运算) todo:需修改公式 
        /// </summary>
        /// <param name="resultdt"></param>
        /// <param name="rows"></param>
        /// <returns></returns>
        private DataTable GenerateLdt(DataTable resultdt,DataRow rows)
        {
            var newrow = resultdt.NewRow();
            //循环插入至resultdt临时表 色母量(L):公式=Round(单个色母量*0.1,2)
            for (var j = 0; j < resultdt.Columns.Count; j++)
            {
                //表示到最后一列=>将内部色号+"&"+版本日期
                if (j == resultdt.Columns.Count - 1)
                {
                    newrow[j] = Convert.ToString(rows[4] + "&" + Convert.ToString(rows[9]));
                }
                //计算色母量
                else if (j == 15)
                {
                    newrow[j] = Math.Round(Convert.ToDecimal(rows[j]) * Convert.ToDecimal(0.1), 3);
                }
                else
                {
                    //其它列操作
                    newrow[j] = rows[j];
                }
            }
            resultdt.Rows.Add(newrow);
            return resultdt;
        }

        /// <summary>
        /// 不需计算色母量使用
        /// </summary>
        /// <param name="resultdt"></param>
        /// <param name="rows"></param>
        /// <returns></returns>
        private DataTable Generatedt(DataTable resultdt,DataRow rows)
        {
            var newrow = resultdt.NewRow();
            for (var j = 0; j < resultdt.Columns.Count; j++)
            {
                //表示到最后一列=>将内部色号+"&"+版本日期
                if (j == resultdt.Columns.Count - 1)
                {
                    newrow[j] = Convert.ToString(rows[4] + "&" + Convert.ToString(rows[9]));
                }
                else
                {
                    newrow[j] = rows[j];
                }
            }
            resultdt.Rows.Add(newrow);
            return resultdt;    
        }

        /// <summary>
        /// 根据不同模式转换输出效果
        /// </summary>
        /// <param name="selectid">1:纵向 2:横向</param>
        /// <param name="sortid">筛选条件(1:不筛选 2:获取1个色母数的配方 3:获取2个色母数的配方 4:获取3个色母数并包含PC-60的配方 5:获取包含PC-60并占比>=20%的配方)</param>
        /// <param name="sourcedt">数据源(以纵向方式)</param>
        /// <returns></returns>
        private DataTable MakeExportMode(int selectid,int sortid,DataTable sourcedt)
        {
            DataTable resultdt;
            DataTable tempdt;
            //除筛选条件为‘不筛选’外,其它都要进行判断获取表头信息
            tempdt = sortid == 1 ? _tempdt : GenerateNewDt(sourcedt);

            //纵向输出
            if (selectid == 1)
            {
                //获取纵向输出模板
                resultdt = sourcedt.Clone();

                foreach (DataRow rows in tempdt.Rows)
                {
                    var dtrows = sourcedt.Select("制造商='" + rows[0] + "' and 版本日期='" + rows[1] + "' and 内部色号='" + rows[2] + "' and 层='" + rows[3] + "' and 涂层='" + rows[4] + "'");

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
                    //在结束一个配方的插入后再插入一行空白行
                    resultdt.Merge(InsertNullRow(resultdt));
                }
            }
            //横向输出
            else
            {
                //获取横向输出模板
                resultdt = dtList.Get_ExportVdt();

                foreach (DataRow rows in tempdt.Rows)
                {
                    //一开始只获取查找到的明细内容的第一行(除色母编码明细外)
                    var dtrows = sourcedt.Select("制造商='" + rows[0] + "' and 版本日期='" + rows[1] + "' and 内部色号='" + rows[2] + "' and 层='" + rows[3] + "' and 涂层='" + rows[4] + "'");

                    var newrow = resultdt.NewRow();
                    newrow[0] = dtrows[0][0];      //制造商
                    newrow[1] = dtrows[0][1];      //车型
                    newrow[2] = dtrows[0][2];      //涂层
                    newrow[3] = dtrows[0][3];      //颜色描述
                    newrow[4] = dtrows[0][4];      //内部色号
                    newrow[5] = dtrows[0][5];      //主配方色号（差异色)
                    newrow[6] = dtrows[0][6];      //颜色组别
                    newrow[7] = dtrows[0][7];      //标准色号
                    newrow[8] = dtrows[0][8];      //RGBValue
                    newrow[9] = dtrows[0][9];      //版本日期
                    newrow[10] = dtrows[0][10];    //层
                    newrow[11] = dtrows[0][11];    //制作人
                    newrow[12] = dtrows[0][12];    //二维码编码

                    //将‘色母’相关信息,插入至对应的项内
                    var rowsdtl= sourcedt.Select("制造商='" + rows[0] + "' and 版本日期='" + rows[1] + "' and 内部色号='" + rows[2] + "' and 层='" + rows[3] + "' and 涂层='" + rows[4] + "'");

                    for (var i = 0; i < rowsdtl.Length; i++)
                    {
                        newrow[13 + i + i] = rowsdtl[i][13];      //色母编码
                        newrow[13 + i + i + 1] = rowsdtl[i][15];  //色母量
                    }
                    //生成内部色号&版本日期
                    newrow[73] = Convert.ToString(dtrows[0][4])+"&"+Convert.ToString(dtrows[0][9]);
                    resultdt.Rows.Add(newrow);
                }
            }
            return resultdt;
        }

        /// <summary>
        /// 插入空白行(纵向导出模式使用)
        /// </summary>
        /// <returns></returns>
        private DataTable InsertNullRow(DataTable sourcedt)
        {
            var newrow = sourcedt.NewRow();

            for (var i = 0; i < sourcedt.Columns.Count; i++)
            {
                newrow[i] = DBNull.Value;
            }

            sourcedt.Rows.Add(newrow);
            return sourcedt;
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
            //排序方式改为:制造商 版本日期 内部色号 层 涂层 

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

            //获取内部色号记录
            foreach (DataRow rows in sourcedt.Rows)
            {
                var newrow = tempdt.NewRow();
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
                tempdt.Rows.Add(newrow);
            }
            _tempdt = tempdt;
        }

        /// <summary>
        /// 若选择了'导入增白(控色剂)EXCEL'按钮时执行-
        /// 使用整理出来的sourcedt与_tempdt进行比较,若_tempdt在sourcedt内存在,才将记录插入至resultdt内
        /// </summary>
        /// <returns></returns>
        private DataTable GenerateNewDt(DataTable sourcedt)
        {
            var resultdt = _tempdt.Clone();

            foreach (DataRow rows in _tempdt.Rows)
            {
                var dtlrow= sourcedt.Select("制造商='" + rows[0] + "' and 版本日期='" + rows[1] + "' and 内部色号='" + rows[2] + "' and 层='" + rows[3] + "' and 涂层='" + rows[4] + "'");
                //若存在,才将记录插入至resultdt内
                if (dtlrow.Length <= 0) continue;
                var newrow = resultdt.NewRow();
                //将_tempdt的内容插入至resultdt内
                for (var j = 0; j < resultdt.Columns.Count; j++)
                {
                    newrow[j] = rows[j];
                }               
                resultdt.Rows.Add(newrow);
            }
            return resultdt;
        }
    }
}
