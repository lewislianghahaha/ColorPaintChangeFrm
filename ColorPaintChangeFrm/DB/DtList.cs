﻿using System;
using System.Data;

namespace ColorPaintChangeFrm.DB
{
    public class DtList
    {
        /// <summary>
        /// 导入模板
        /// </summary>
        /// <returns></returns>
        public DataTable Get_Importdt()
        {
            var dt = new DataTable();
            for (var i = 0; i < 16; i++)
            {
                var dc = new DataColumn();

                switch (i)
                {
                    case 0:
                        dc.ColumnName = "制造商";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 1:
                        dc.ColumnName = "车型";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 2:
                        dc.ColumnName = "涂层";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 3:
                        dc.ColumnName = "颜色描述";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 4:
                        dc.ColumnName = "内部色号";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 5:
                        dc.ColumnName = "主配方(差异色)";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 6:
                        dc.ColumnName = "颜色组别";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 7:
                        dc.ColumnName = "标准色号";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 8:
                        dc.ColumnName = "RGBValue";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 9:
                        dc.ColumnName = "版本日期";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 10:
                        dc.ColumnName = "层";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 11:
                        dc.ColumnName = "制作人";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 12:
                        dc.ColumnName = "二维码编号";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 13:
                        dc.ColumnName = "色母编码";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 14:
                        dc.ColumnName = "色母名称";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 15:
                        dc.ColumnName = "色母量";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                }
                dt.Columns.Add(dc);
            }
            return dt;
        }

        /// <summary>
        /// 导出模板(以纵向模式)
        /// </summary>
        /// <returns></returns>
        public DataTable Get_ExportHdt()
        {
            var dt = new DataTable();
            for (var i = 0; i < 17; i++)
            {
                var dc = new DataColumn();

                switch (i)
                {
                    case 0:
                        dc.ColumnName = "制造商";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 1:
                        dc.ColumnName = "车型";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 2:
                        dc.ColumnName = "涂层";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 3:
                        dc.ColumnName = "颜色描述";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 4:
                        dc.ColumnName = "内部色号";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 5:
                        dc.ColumnName = "主配方(差异色)";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 6:
                        dc.ColumnName = "颜色组别";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 7:
                        dc.ColumnName = "标准色号";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 8:
                        dc.ColumnName = "RGBValue";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 9:
                        dc.ColumnName = "版本日期";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 10:
                        dc.ColumnName = "层";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 11:
                        dc.ColumnName = "制作人";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 12:
                        dc.ColumnName = "二维码编号";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 13:
                        dc.ColumnName = "色母编码";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 14:
                        dc.ColumnName = "色母名称";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 15:
                        dc.ColumnName = "色母量(KG)";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                    case 16:
                        dc.ColumnName = "内部色号&版本日期";
                        dc.DataType = Type.GetType("System.String");
                        break;
                }
                dt.Columns.Add(dc);
            }
            return dt;
        }

        /// <summary>
        /// 导出模板(以横向模式)
        /// </summary>
        /// <returns></returns>
        public DataTable Get_ExportVdt()
        {
            var dt = new DataTable();
            for (var i = 0; i < 74; i++)
            {
                var dc = new DataColumn();

                switch (i)
                {
                    case 0:
                        dc.ColumnName = "制造商";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 1:
                        dc.ColumnName = "车型";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 2:
                        dc.ColumnName = "涂层";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 3:
                        dc.ColumnName = "颜色描述";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 4:
                        dc.ColumnName = "内部色号";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 5:
                        dc.ColumnName = "主配方色号（差异色)";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 6:
                        dc.ColumnName = "颜色组别";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 7:
                        dc.ColumnName = "标准色号";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 8:
                        dc.ColumnName = "RGBValue";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 9:
                        dc.ColumnName = "版本日期";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 10:
                        dc.ColumnName = "层";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 11:
                        dc.ColumnName = "制作人";
                        dc.DataType = Type.GetType("System.String"); 
                        break;
                    case 12:
                        dc.ColumnName = "二维码编号";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    
                    case 13:
                        dc.ColumnName = "色母1";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 14:
                        dc.ColumnName = "色母量1";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 15:
                        dc.ColumnName = "色母2";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 16:
                        dc.ColumnName = "色母量2";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 17:
                        dc.ColumnName = "色母3";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 18:
                        dc.ColumnName = "色母量3";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 19:
                        dc.ColumnName = "色母4";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 20:
                        dc.ColumnName = "色母量4";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 21:
                        dc.ColumnName = "色母5";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 22:
                        dc.ColumnName = "色母量5";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 23:
                        dc.ColumnName = "色母6";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 24:
                        dc.ColumnName = "色母量6";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 25:
                        dc.ColumnName = "色母7";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 26:
                        dc.ColumnName = "色母量7";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 27:
                        dc.ColumnName = "色母8";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 28:
                        dc.ColumnName = "色母量8";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 29:
                        dc.ColumnName = "色母9";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 30:
                        dc.ColumnName = "色母量9";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 31:
                        dc.ColumnName = "色母10";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 32:
                        dc.ColumnName = "色母量10";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 33:
                        dc.ColumnName = "色母11";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 34:
                        dc.ColumnName = "色母量11";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 35:
                        dc.ColumnName = "色母12";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 36:
                        dc.ColumnName = "色母量12";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 37:
                        dc.ColumnName = "色母13";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 38:
                        dc.ColumnName = "色母量13";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 39:
                        dc.ColumnName = "色母14";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 40:
                        dc.ColumnName = "色母量14";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 41:
                        dc.ColumnName = "色母15";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 42:
                        dc.ColumnName = "色母量15";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 43:
                        dc.ColumnName = "色母16";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 44:
                        dc.ColumnName = "色母量16";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 45:
                        dc.ColumnName = "色母17";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 46:
                        dc.ColumnName = "色母量17";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 47:
                        dc.ColumnName = "色母18";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 48:
                        dc.ColumnName = "色母量18";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 49:
                        dc.ColumnName = "色母19";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 50:
                        dc.ColumnName = "色母量19";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 51:
                        dc.ColumnName = "色母20";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 52:
                        dc.ColumnName = "色母量20";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 53:
                        dc.ColumnName = "色母21";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 54:
                        dc.ColumnName = "色母量21";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 55:
                        dc.ColumnName = "色母22";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 56:
                        dc.ColumnName = "色母量22";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 57:
                        dc.ColumnName = "色母23";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 58:
                        dc.ColumnName = "色母量23";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 59:
                        dc.ColumnName = "色母24";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 60:
                        dc.ColumnName = "色母量24";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 61:
                        dc.ColumnName = "色母25";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 62:
                        dc.ColumnName = "色母量25";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 63:
                        dc.ColumnName = "色母26";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 64:
                        dc.ColumnName = "色母量26";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 65:
                        dc.ColumnName = "色母27";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 66:
                        dc.ColumnName = "色母量27";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 67:
                        dc.ColumnName = "色母28";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 68:
                        dc.ColumnName = "色母量28";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 69:
                        dc.ColumnName = "色母29";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 70:
                        dc.ColumnName = "色母量29";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 71:
                        dc.ColumnName = "色母30";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 72:
                        dc.ColumnName = "色母量30";
                        dc.DataType = Type.GetType("System.Decimal");
                        break;
                    case 73:
                        dc.ColumnName = "内部色号&版本日期";
                        dc.DataType = Type.GetType("System.String");
                        break;
                }
                dt.Columns.Add(dc);
            }
            return dt;
        }

        /// <summary>
        /// 内部色号临时表-运算时使用
        /// </summary>
        /// <returns></returns>
        public DataTable Get_ColorcodeDt()
        {
            var dt = new DataTable();
            for (var i = 0; i < 5; i++)
            {
                var dc = new DataColumn();

                switch (i)
                {
                    case 0:
                        dc.ColumnName = "制造商";
                        dc.DataType = Type.GetType("System.String");
                    break;
                    case 1:
                        dc.ColumnName = "版本日期";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 2:
                        dc.ColumnName = "内部色号";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 3:
                        dc.ColumnName = "层";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 4:
                        dc.ColumnName = "涂层";
                        dc.DataType = Type.GetType("System.String");
                        break;
                }
                dt.Columns.Add(dc);
            }
            return dt;
        }

        /// <summary>
        /// 导出模板(色母占比率使用)
        /// </summary>
        /// <returns></returns>
        public DataTable Get_ExportPerDt()
        {
            var dt = new DataTable();
            for (var i = 0; i < 3; i++)
            {
                var dc = new DataColumn();

                switch (i)
                {
                    case 0:
                        dc.ColumnName = "内部色号";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 1:
                        dc.ColumnName = "版本日期";
                        dc.DataType = Type.GetType("System.String");
                        break;
                    case 2:
                        dc.ColumnName = "占比率";
                        dc.DataType = Type.GetType("System.Double");
                        break;
                }
                dt.Columns.Add(dc);
            }
            return dt;
        }
    }
}
