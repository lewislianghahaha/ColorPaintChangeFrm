using System;
using System.Data;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ColorPaintChangeFrm.Logic
{
    public class ExportDt
    {
        /// <summary>
        /// 导出
        /// </summary>
        /// <param name="fileAddress">导出地址</param>
        /// <param name="sourcedt"></param>
        /// <param name="comselectid">1:纵向  2:横向 3:占比率使用</param>
        public bool ExportDtToExcel(string fileAddress, DataTable sourcedt,int comselectid)
        {
            var result = true;
            var sheetcount = 0;  //记录所需的sheet页总数
            var rownum = 1;

            try
            {
                //声明一个WorkBook
                var xssfWorkbook = new XSSFWorkbook();

                //执行sheet页(注:1)先列表temp行数判断需拆分多少个sheet表进行填充; 以一个sheet表有100W行记录填充为基准)
                sheetcount = sourcedt.Rows.Count % 1000000 == 0 ? sourcedt.Rows.Count / 1000000 : sourcedt.Rows.Count / 1000000 + 1;
                //i为EXCEL的Sheet页数ID
                for (var i = 1; i <= sheetcount; i++)
                {
                    //创建sheet页
                    var sheet = xssfWorkbook.CreateSheet("Sheet" + i);
                    //创建"标题行"
                    var row = sheet.CreateRow(0);
                    //创建sheet页各列标题
                    for (var j = 0; j < sourcedt.Columns.Count; j++)
                    {
                        //设置列宽度
                        sheet.SetColumnWidth(j, (int)((20 + 0.72) * 256));
                        //创建标题
                        //纵向使用
                        if (comselectid == 1)
                        {
                            switch (j)
                            {
                                #region SetCellValue 纵向使用
                                case 0:
                                    row.CreateCell(j).SetCellValue("制造商");
                                    break;
                                case 1:
                                    row.CreateCell(j).SetCellValue("车型");
                                    break;
                                case 2:
                                    row.CreateCell(j).SetCellValue("涂层");
                                    break;
                                case 3:
                                    row.CreateCell(j).SetCellValue("颜色描述");
                                    break;
                                case 4:
                                    row.CreateCell(j).SetCellValue("内部色号");
                                    break;
                                case 5:
                                    row.CreateCell(j).SetCellValue("主配方(差异色)");
                                    break;
                                case 6:
                                    row.CreateCell(j).SetCellValue("颜色组别");
                                    break;
                                case 7:
                                    row.CreateCell(j).SetCellValue("标准色号");
                                    break;
                                case 8:
                                    row.CreateCell(j).SetCellValue("RGBValue");
                                    break;
                                case 9:
                                    row.CreateCell(j).SetCellValue("版本日期");
                                    break;
                                case 10:
                                    row.CreateCell(j).SetCellValue("层");
                                    break;
                                case 11:
                                    row.CreateCell(j).SetCellValue("制作人");
                                    break;
                                case 12:
                                    row.CreateCell(j).SetCellValue("二维码编号");
                                    break;
                                case 13:
                                    row.CreateCell(j).SetCellValue("色母编码");
                                    break;
                                case 14:
                                    row.CreateCell(j).SetCellValue("色母名称");
                                    break;
                                case 15:
                                    row.CreateCell(j).SetCellValue("色母量(KG)");
                                    break;
                                case 16:
                                    row.CreateCell(j).SetCellValue("内部色号&版本日期");
                                    break;
                                    #endregion
                            }
                        }
                        //横向使用
                        else if (comselectid == 2)
                        {
                            switch (j)
                            {
                                #region SetCellValue 横向使用
                                case 0:
                                    row.CreateCell(j).SetCellValue("制造商");
                                    break;
                                case 1:
                                    row.CreateCell(j).SetCellValue("车型");
                                    break;
                                case 2:
                                    row.CreateCell(j).SetCellValue("涂层");
                                    break;
                                case 3:
                                    row.CreateCell(j).SetCellValue("颜色描述");
                                    break;
                                case 4:
                                    row.CreateCell(j).SetCellValue("内部色号");
                                    break;
                                case 5:
                                    row.CreateCell(j).SetCellValue("主配方色号(差异色)");
                                    break;
                                case 6:
                                    row.CreateCell(j).SetCellValue("颜色组别");
                                    break;
                                case 7:
                                    row.CreateCell(j).SetCellValue("标准色号");
                                    break;
                                case 8:
                                    row.CreateCell(j).SetCellValue("RGBValue");
                                    break;
                                case 9:
                                    row.CreateCell(j).SetCellValue("版本日期");
                                    break;
                                case 10:
                                    row.CreateCell(j).SetCellValue("层");
                                    break;
                                case 11:
                                    row.CreateCell(j).SetCellValue("制作人");
                                    break;
                                case 12:
                                    row.CreateCell(j).SetCellValue("二维码编号");
                                    break;


                                case 13:
                                    row.CreateCell(j).SetCellValue("色母1");
                                    break;
                                case 14:
                                    row.CreateCell(j).SetCellValue("色母量1");
                                    break;
                                case 15:
                                    row.CreateCell(j).SetCellValue("色母2");
                                    break;
                                case 16:
                                    row.CreateCell(j).SetCellValue("色母量2");
                                    break;
                                case 17:
                                    row.CreateCell(j).SetCellValue("色母3");
                                    break;
                                case 18:
                                    row.CreateCell(j).SetCellValue("色母量3");
                                    break;
                                case 19:
                                    row.CreateCell(j).SetCellValue("色母4");
                                    break;
                                case 20:
                                    row.CreateCell(j).SetCellValue("色母量4");
                                    break;
                                case 21:
                                    row.CreateCell(j).SetCellValue("色母5");
                                    break;
                                case 22:
                                    row.CreateCell(j).SetCellValue("色母量5");
                                    break;
                                case 23:
                                    row.CreateCell(j).SetCellValue("色母6");
                                    break;
                                case 24:
                                    row.CreateCell(j).SetCellValue("色母量6");
                                    break;
                                case 25:
                                    row.CreateCell(j).SetCellValue("色母7");
                                    break;
                                case 26:
                                    row.CreateCell(j).SetCellValue("色母量7");
                                    break;
                                case 27:
                                    row.CreateCell(j).SetCellValue("色母8");
                                    break;
                                case 28:
                                    row.CreateCell(j).SetCellValue("色母量8");
                                    break;
                                case 29:
                                    row.CreateCell(j).SetCellValue("色母9");
                                    break;
                                case 30:
                                    row.CreateCell(j).SetCellValue("色母量9");
                                    break;
                                case 31:
                                    row.CreateCell(j).SetCellValue("色母10");
                                    break;
                                case 32:
                                    row.CreateCell(j).SetCellValue("色母量10");
                                    break;
                                case 33:
                                    row.CreateCell(j).SetCellValue("色母11");
                                    break;
                                case 34:
                                    row.CreateCell(j).SetCellValue("色母量11");
                                    break;
                                case 73:
                                    row.CreateCell(j).SetCellValue("内部色号&版本日期");
                                    break;
                                    #endregion
                            }
                        }
                        //占比率使用
                        else if (comselectid==3)
                        {
                            switch (j)
                            {
                                #region 占比率使用
                                case 0:
                                    row.CreateCell(j).SetCellValue("内部色号");
                                    break;
                                case 1:
                                    row.CreateCell(j).SetCellValue("版本日期");
                                    break;
                                case 2:
                                    row.CreateCell(j).SetCellValue("占比率");
                                    break;
                                    #endregion
                            }
                        }
                    }

                    //计算进行循环的起始行
                    var startrow = (i - 1) * 1000000;
                    //计算进行循环的结束行
                    var endrow = i == sheetcount ? sourcedt.Rows.Count : i * 1000000;

                    //每一个sheet表显示100000行  
                    for (var j = startrow; j < endrow; j++)
                    {
                        //创建行
                        row = sheet.CreateRow(rownum);
                        //循环获取DT内的列值记录
                        for (var k = 0; k < sourcedt.Columns.Count; k++)
                        {
                            if (Convert.ToString(sourcedt.Rows[j][k]) == "") continue;
                            else
                            {
                                //(注:要注意值小数位数保留两位;当超出三位小数的时候,会出现OutofMemory异常.)
                                if (comselectid == 1)
                                {
                                    if (k == 15)
                                    {
                                        row.CreateCell(k, CellType.Numeric).SetCellValue(Convert.ToDouble(sourcedt.Rows[j][k]));
                                    }
                                    else
                                    {
                                        //除‘色母量’以及‘累积量’外的值的转换赋值 或 横向导出时
                                        row.CreateCell(k, CellType.String).SetCellValue(Convert.ToString(sourcedt.Rows[j][k]));
                                    }
                                }
                                //横向使用
                                else if (comselectid == 2)
                                {
                                    row.CreateCell(k, CellType.String).SetCellValue(Convert.ToString(sourcedt.Rows[j][k]));
                                }
                                //占比率输出使用
                                else
                                {
                                    if (k == 2)
                                    {
                                        row.CreateCell(k, CellType.Numeric).SetCellValue(Convert.ToDouble(sourcedt.Rows[j][k]));
                                    }
                                    else
                                    {
                                        //除‘占比率’时
                                        row.CreateCell(k, CellType.String).SetCellValue(Convert.ToString(sourcedt.Rows[j][k]));
                                    }
                                }
                            }
                        }
                        rownum++;
                    }
                    //当一个SHEET页填充完毕后,需将变量初始化
                    rownum = 1;
                }

                //写入数据
                var file = new FileStream(fileAddress, FileMode.Create);
                xssfWorkbook.Write(file);
                file.Close();           //关闭文件流
                xssfWorkbook.Close();   //关闭工作簿
                file.Dispose();         //释放文件流
            }
            catch (Exception)
            {
                result = false;
            }
            return result;
        }


    }
}
